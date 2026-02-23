from dataclasses import dataclass
from typing import Dict, List, Optional
import ctypes
import json
import os
import subprocess
import sys
import tempfile
import time
import winreg

try:
    from PyQt6.QtCore import QEasingCurve, QPropertyAnimation, QRectF, Qt, QTimer, pyqtProperty
    from PyQt6.QtGui import QColor, QFont, QLinearGradient, QPainter
    from PyQt6.QtWidgets import (
        QApplication,
        QComboBox,
        QFrame,
        QGraphicsDropShadowEffect,
        QHBoxLayout,
        QLabel,
        QLineEdit,
        QMainWindow,
        QMessageBox,
        QPushButton,
        QScrollArea,
        QSizePolicy,
        QVBoxLayout,
        QWidget,
    )

    PYQT_VER = 6
except ImportError:
    from PyQt5.QtCore import QEasingCurve, QPropertyAnimation, QRectF, Qt, QTimer, pyqtProperty
    from PyQt5.QtGui import QColor, QFont, QLinearGradient, QPainter
    from PyQt5.QtWidgets import (
        QApplication,
        QComboBox,
        QFrame,
        QGraphicsDropShadowEffect,
        QHBoxLayout,
        QLabel,
        QLineEdit,
        QMainWindow,
        QMessageBox,
        QPushButton,
        QScrollArea,
        QSizePolicy,
        QVBoxLayout,
        QWidget,
    )

    PYQT_VER = 5

USB_ENUM_ROOT = r"SYSTEM\CurrentControlSet\Enum\USB"
REFRESH_INTERVAL_MS = 3000
PENDING_TIMEOUT_SEC = 75


def align_left_vcenter():
    if PYQT_VER == 6:
        return Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
    return Qt.AlignLeft | Qt.AlignVCenter


def align_right_vcenter():
    if PYQT_VER == 6:
        return Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
    return Qt.AlignRight | Qt.AlignVCenter


def is_access_denied(exc: OSError) -> bool:
    return getattr(exc, "winerror", None) == 5


def read_reg_value(path: str, name: str):
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path, 0, winreg.KEY_READ) as key:
            value, _ = winreg.QueryValueEx(key, name)
            return value
    except OSError:
        return None


def clean_registry_text(value) -> str:
    if not isinstance(value, str):
        return ""
    text = value.strip()
    if not text:
        return ""
    if ";" in text:
        tail = text.split(";", 1)[1].strip()
        if tail:
            return tail
    return text.lstrip("@").strip()


def select_display_name(parent_path: str) -> str:
    friendly = clean_registry_text(read_reg_value(parent_path, "FriendlyName"))
    if friendly:
        return friendly

    bus_desc = clean_registry_text(read_reg_value(parent_path, "BusReportedDeviceDesc"))
    if bus_desc:
        return bus_desc

    desc = clean_registry_text(read_reg_value(parent_path, "DeviceDesc"))
    if desc:
        return desc

    return "Unknown USB device"


def select_device_type(parent_path: str, key_path: str) -> str:
    class_name = clean_registry_text(read_reg_value(parent_path, "Class"))
    if class_name:
        return class_name

    service = clean_registry_text(read_reg_value(parent_path, "Service"))
    if service:
        return service

    upper_key = key_path.upper()
    if "HID" in upper_key:
        return "HID"
    if "VID_" in upper_key:
        return "USB"
    return "Other"


def enumerate_device_parameter_paths() -> List[str]:
    paths: List[str] = []

    def walk(key_handle, key_path: str):
        index = 0
        while True:
            try:
                child_name = winreg.EnumKey(key_handle, index)
            except OSError:
                break
            index += 1
            child_path = f"{key_path}\\{child_name}"

            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, child_path, 0, winreg.KEY_READ) as child_key:
                    if child_name.lower() == "device parameters":
                        paths.append(child_path)
                    walk(child_key, child_path)
            except PermissionError:
                continue
            except OSError:
                continue

    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, USB_ENUM_ROOT, 0, winreg.KEY_READ) as root_key:
        walk(root_key, USB_ENUM_ROOT)

    return paths


def set_epm_value(key_path: str, value: int):
    with winreg.OpenKey(
        winreg.HKEY_LOCAL_MACHINE,
        key_path,
        0,
        winreg.KEY_SET_VALUE | winreg.KEY_QUERY_VALUE,
    ) as key:
        winreg.SetValueEx(key, "EnhancedPowerManagementEnabled", 0, winreg.REG_DWORD, value)


def disable_epm_for_all() -> int:
    failures = 0
    for key_path in enumerate_device_parameter_paths():
        try:
            set_epm_value(key_path, 0)
        except OSError:
            failures += 1
    return failures


def shell_quote_ps(value: str) -> str:
    return value.replace("'", "''")


def launch_elevated_powershell(command: str) -> bool:
    params = subprocess.list2cmdline(
        ["-NoProfile", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden", "-Command", command]
    )
    result = ctypes.windll.shell32.ShellExecuteW(None, "runas", "powershell.exe", params, None, 0)
    return result > 32


@dataclass
class USBDevice:
    key_path: str
    parent_path: str
    device_desc: str
    manufacturer: str
    device_type: str
    epm_value: Optional[int]

    @property
    def sleep_disabled(self) -> Optional[bool]:
        if self.epm_value is None:
            return None
        return self.epm_value == 0


class FlowBackgroundWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._phase = 0.0
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._tick)
        self._timer.start(33)

    def _tick(self):
        self._phase += 0.006
        if self._phase > 1.0:
            self._phase = 0.0
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)

        w = max(self.width(), 1)
        h = max(self.height(), 1)

        base = QLinearGradient(0, 0, w, h)
        base.setColorAt(0.0, QColor("#09131f"))
        base.setColorAt(0.45, QColor("#0e2537"))
        base.setColorAt(1.0, QColor("#132b40"))
        painter.fillRect(self.rect(), base)

        shift = self._phase
        flow1 = QLinearGradient(0, 0, w, h)
        flow1.setColorAt(max(0.0, shift - 0.25), QColor(37, 165, 255, 0))
        flow1.setColorAt(shift, QColor(37, 165, 255, 75))
        flow1.setColorAt(min(1.0, shift + 0.25), QColor(37, 165, 255, 0))
        painter.fillRect(self.rect(), flow1)

        flow2 = QLinearGradient(w, 0, 0, h)
        shift2 = 1.0 - self._phase
        flow2.setColorAt(max(0.0, shift2 - 0.22), QColor(84, 232, 209, 0))
        flow2.setColorAt(shift2, QColor(84, 232, 209, 48))
        flow2.setColorAt(min(1.0, shift2 + 0.22), QColor(84, 232, 209, 0))
        painter.fillRect(self.rect(), flow2)

        super().paintEvent(event)


class ToggleSwitch(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(72, 34)
        self._checked = False
        self._enabled = True
        self._offset = 3.0
        self._anim = QPropertyAnimation(self, b"offset", self)
        self._anim.setDuration(170)
        self._anim.setEasingCurve(QEasingCurve.Type.OutCubic if PYQT_VER == 6 else QEasingCurve.OutCubic)
        self.on_toggled = None

    def mousePressEvent(self, event):
        if self._enabled:
            self.setChecked(not self._checked, animated=True)
            if callable(self.on_toggled):
                self.on_toggled(self._checked)
        super().mousePressEvent(event)

    def setChecked(self, checked: bool, animated: bool = True):
        self._checked = bool(checked)
        target = float(self.width() - self.height() + 3) if self._checked else 3.0
        self._anim.stop()
        if animated:
            self._anim.setStartValue(self._offset)
            self._anim.setEndValue(target)
            self._anim.start()
        else:
            self._offset = target
            self.update()

    def setEnabledState(self, enabled: bool):
        self._enabled = bool(enabled)
        self.update()

    def get_offset(self):
        return self._offset

    def set_offset(self, value):
        self._offset = float(value)
        self.update()

    offset = pyqtProperty(float, fget=get_offset, fset=set_offset)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)

        track_rect = QRectF(1, 1, self.width() - 2, self.height() - 2)
        radius = track_rect.height() / 2

        if not self._enabled:
            track_color = QColor("#475569")
            knob_color = QColor("#cbd5e1")
        elif self._checked:
            track_color = QColor("#16a34a")
            knob_color = QColor("#ecfdf5")
        else:
            track_color = QColor("#334155")
            knob_color = QColor("#f1f5f9")

        painter.setPen(Qt.PenStyle.NoPen if PYQT_VER == 6 else Qt.NoPen)
        painter.setBrush(track_color)
        painter.drawRoundedRect(track_rect, radius, radius)

        knob_d = self.height() - 6
        knob_rect = QRectF(self._offset, 3, knob_d, knob_d)
        painter.setBrush(knob_color)
        painter.drawEllipse(knob_rect)


class DeviceCard(QFrame):
    def __init__(self, device: USBDevice, toggle_callback, parent=None):
        super().__init__(parent)
        self.device = device
        self.toggle_callback = toggle_callback
        self._updating = False

        self.setObjectName("deviceCard")
        self.setFrameShape(QFrame.Shape.NoFrame if PYQT_VER == 6 else QFrame.NoFrame)

        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(20)
        shadow.setOffset(0, 4)
        shadow.setColor(QColor(0, 0, 0, 90))
        self.setGraphicsEffect(shadow)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(14, 12, 14, 12)
        outer.setSpacing(8)

        top = QHBoxLayout()
        top.setSpacing(12)

        self.title_label = QLabel()
        self.title_label.setObjectName("deviceTitle")
        self.title_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        self.title_label.setAlignment(align_left_vcenter())

        self.type_label = QLabel()
        self.type_label.setObjectName("typeTag")

        self.status_label = QLabel()
        self.status_label.setObjectName("statusLabel")

        self.switch = ToggleSwitch()
        self.switch.on_toggled = self._on_switch

        top.addWidget(self.title_label, 1)
        top.addWidget(self.type_label)
        top.addWidget(self.status_label)
        top.addWidget(self.switch)

        self.path_label = QLabel()
        self.path_label.setObjectName("pathLabel")
        self.path_label.setWordWrap(True)

        outer.addLayout(top)
        outer.addWidget(self.path_label)

        self.update_from_device(device)

    def _on_switch(self, checked: bool):
        if self._updating:
            return
        self.toggle_callback(self.device.key_path, checked)

    def update_from_device(self, device: USBDevice):
        self._updating = True
        self.device = device

        subtitle = device.manufacturer if device.manufacturer else "Unknown manufacturer"
        title = device.device_desc if device.device_desc else "Unknown USB device"
        self.title_label.setText(f"{title}  |  {subtitle}")
        self.type_label.setText(device.device_type)
        self.path_label.setText(device.key_path)

        if device.epm_value is None:
            self.status_label.setText("Sleep: Unavailable")
            self.status_label.setProperty("state", "na")
            self.switch.setEnabledState(False)
            self.switch.setChecked(False, animated=False)
        elif device.sleep_disabled:
            self.status_label.setText("Sleep: Disabled")
            self.status_label.setProperty("state", "off")
            self.switch.setEnabledState(True)
            self.switch.setChecked(True, animated=False)
        else:
            self.status_label.setText("Sleep: Enabled")
            self.status_label.setProperty("state", "on")
            self.switch.setEnabledState(True)
            self.switch.setChecked(False, animated=False)

        self.status_label.style().unpolish(self.status_label)
        self.status_label.style().polish(self.status_label)
        self._updating = False


class USBPowerMainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("USB Power Flow")
        self.resize(1260, 790)

        self.latest_devices: Dict[str, USBDevice] = {}
        self.cards: Dict[str, DeviceCard] = {}
        self.pending_operation = None

        self.bg = FlowBackgroundWidget()
        self.setCentralWidget(self.bg)

        root = QVBoxLayout(self.bg)
        root.setContentsMargins(18, 18, 18, 18)
        root.setSpacing(14)

        header = QFrame()
        header.setObjectName("headerFrame")
        header_layout = QVBoxLayout(header)
        header_layout.setContentsMargins(16, 14, 16, 14)
        header_layout.setSpacing(10)

        row1 = QHBoxLayout()
        title_box = QVBoxLayout()
        title = QLabel("USB Power Flow")
        title.setObjectName("appTitle")
        subtitle = QLabel("Live USB sleep-state control")
        subtitle.setObjectName("appSubtitle")
        title_box.addWidget(title)
        title_box.addWidget(subtitle)

        actions = QHBoxLayout()
        actions.setSpacing(10)
        self.refresh_btn = QPushButton("Refresh")
        self.refresh_btn.setObjectName("secondaryBtn")
        self.refresh_btn.clicked.connect(self.refresh_devices)

        self.disable_all_btn = QPushButton("Disable USB Sleep On All")
        self.disable_all_btn.setObjectName("primaryBtn")
        self.disable_all_btn.clicked.connect(self.disable_sleep_all)

        actions.addWidget(self.refresh_btn)
        actions.addWidget(self.disable_all_btn)

        self.status_label = QLabel("Ready")
        self.status_label.setObjectName("globalStatus")
        self.status_label.setAlignment(align_right_vcenter())

        row1.addLayout(title_box, 1)
        row1.addLayout(actions)
        row1.addWidget(self.status_label)

        row2 = QHBoxLayout()
        row2.setSpacing(10)

        self.search_input = QLineEdit()
        self.search_input.setObjectName("searchInput")
        self.search_input.setPlaceholderText("Search by name, manufacturer, or registry path")
        self.search_input.textChanged.connect(self.apply_view_filters)

        self.type_filter = QComboBox()
        self.type_filter.setObjectName("filterCombo")
        self.type_filter.currentIndexChanged.connect(self.apply_view_filters)

        self.sort_combo = QComboBox()
        self.sort_combo.setObjectName("filterCombo")
        self.sort_combo.addItems(["Name A-Z", "Name Z-A", "State", "Type", "Manufacturer"])
        self.sort_combo.currentIndexChanged.connect(self.apply_view_filters)

        row2.addWidget(self.search_input, 1)
        row2.addWidget(self.type_filter)
        row2.addWidget(self.sort_combo)

        header_layout.addLayout(row1)
        header_layout.addLayout(row2)

        root.addWidget(header)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setObjectName("deviceScroll")

        self.scroll_body = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_body)
        self.scroll_layout.setContentsMargins(2, 2, 2, 2)
        self.scroll_layout.setSpacing(10)
        self.scroll_layout.addStretch(1)

        self.scroll.setWidget(self.scroll_body)
        root.addWidget(self.scroll, 1)

        self.pending_timer = QTimer(self)
        self.pending_timer.timeout.connect(self.poll_pending_operation)

        self.apply_styles()
        self.refresh_devices()

        self.refresh_timer = QTimer(self)
        self.refresh_timer.timeout.connect(lambda: self.refresh_devices(silent=True))
        self.refresh_timer.start(REFRESH_INTERVAL_MS)

    def apply_styles(self):
        app_font = QFont("Segoe UI", 10)
        QApplication.instance().setFont(app_font)

        self.setStyleSheet(
            """
            QMainWindow { background: transparent; }
            QFrame#headerFrame {
                background-color: rgba(10, 19, 33, 190);
                border: 1px solid rgba(118, 191, 255, 70);
                border-radius: 14px;
            }
            QLabel#appTitle {
                font-size: 24px;
                font-weight: 700;
                color: #e2f3ff;
                letter-spacing: 0.4px;
            }
            QLabel#appSubtitle { font-size: 12px; color: #9fc8df; }
            QLabel#globalStatus { font-size: 12px; color: #b5d8eb; min-width: 300px; }
            QPushButton {
                border-radius: 10px;
                padding: 9px 14px;
                font-weight: 600;
            }
            QPushButton#primaryBtn {
                color: #e9fff8;
                border: 1px solid #2dd4bf;
                background-color: rgba(13, 148, 136, 140);
            }
            QPushButton#primaryBtn:hover { background-color: rgba(15, 170, 155, 190); }
            QPushButton#secondaryBtn {
                color: #e6f2ff;
                border: 1px solid #60a5fa;
                background-color: rgba(37, 99, 235, 115);
            }
            QPushButton#secondaryBtn:hover { background-color: rgba(37, 99, 235, 165); }
            QLineEdit#searchInput, QComboBox#filterCombo {
                background-color: rgba(15, 27, 43, 210);
                color: #d8ebf8;
                border: 1px solid rgba(129, 182, 220, 80);
                border-radius: 9px;
                padding: 8px;
                min-height: 16px;
            }
            QScrollArea#deviceScroll { border: none; background: transparent; }
            QFrame#deviceCard {
                background-color: rgba(15, 27, 43, 198);
                border: 1px solid rgba(129, 182, 220, 65);
                border-radius: 12px;
            }
            QLabel#deviceTitle { font-size: 14px; font-weight: 650; color: #d4ecff; }
            QLabel#typeTag {
                font-size: 11px;
                font-weight: 700;
                color: #cffafe;
                background-color: rgba(8, 145, 178, 145);
                border-radius: 7px;
                padding: 3px 8px;
            }
            QLabel#pathLabel { font-family: Consolas; font-size: 11px; color: #90aec4; }
            QLabel#statusLabel {
                font-size: 12px;
                font-weight: 700;
                padding: 4px 8px;
                border-radius: 7px;
                color: #d8eaff;
                background-color: rgba(71, 85, 105, 160);
            }
            QLabel#statusLabel[state="off"] { color: #d1fae5; background-color: rgba(22, 163, 74, 140); }
            QLabel#statusLabel[state="on"] { color: #e2e8f0; background-color: rgba(71, 85, 105, 160); }
            QLabel#statusLabel[state="na"] { color: #f8d9b8; background-color: rgba(194, 110, 24, 150); }
            """
        )

    def refresh_devices(self, silent: bool = False):
        try:
            devices = self.scan_usb_devices()
        except PermissionError:
            self.status_label.setText("Read failed: run as Administrator")
            return
        except OSError as exc:
            self.status_label.setText(f"Read failed: {exc}")
            return

        self.latest_devices = {d.key_path: d for d in devices}
        self.refresh_type_filter_items()
        self.apply_view_filters()

        if not silent:
            self.status_label.setText(f"Loaded {len(devices)} USB device parameter entries")

    def scan_usb_devices(self) -> List[USBDevice]:
        devices: List[USBDevice] = []
        for child_path in enumerate_device_parameter_paths():
            parent_path = child_path.rsplit("\\", 1)[0]
            desc = select_display_name(parent_path)
            mfg = clean_registry_text(read_reg_value(parent_path, "Mfg")) or ""
            dtype = select_device_type(parent_path, child_path)
            epm = read_reg_value(child_path, "EnhancedPowerManagementEnabled")
            epm_value = epm if isinstance(epm, int) else None
            devices.append(
                USBDevice(
                    key_path=child_path,
                    parent_path=parent_path,
                    device_desc=desc,
                    manufacturer=str(mfg),
                    device_type=dtype,
                    epm_value=epm_value,
                )
            )

        devices.sort(key=lambda d: (d.device_desc.lower(), d.key_path.lower()))
        return devices

    def refresh_type_filter_items(self):
        current = self.type_filter.currentText()
        types = sorted({d.device_type for d in self.latest_devices.values()}, key=lambda x: x.lower())
        items = ["All Types"] + types

        self.type_filter.blockSignals(True)
        self.type_filter.clear()
        self.type_filter.addItems(items)

        idx = self.type_filter.findText(current)
        if idx >= 0:
            self.type_filter.setCurrentIndex(idx)
        else:
            self.type_filter.setCurrentIndex(0)
        self.type_filter.blockSignals(False)

    def filtered_sorted_devices(self) -> List[USBDevice]:
        devices = list(self.latest_devices.values())

        query = self.search_input.text().strip().lower()
        selected_type = self.type_filter.currentText()

        if selected_type and selected_type != "All Types":
            devices = [d for d in devices if d.device_type == selected_type]

        if query:
            devices = [
                d
                for d in devices
                if query in d.device_desc.lower()
                or query in d.manufacturer.lower()
                or query in d.key_path.lower()
                or query in d.device_type.lower()
            ]

        sort_mode = self.sort_combo.currentText()
        if sort_mode == "Name Z-A":
            devices.sort(key=lambda d: (d.device_desc.lower(), d.key_path.lower()), reverse=True)
        elif sort_mode == "State":
            state_rank = {True: 0, False: 1, None: 2}
            devices.sort(key=lambda d: (state_rank[d.sleep_disabled], d.device_desc.lower()))
        elif sort_mode == "Type":
            devices.sort(key=lambda d: (d.device_type.lower(), d.device_desc.lower()))
        elif sort_mode == "Manufacturer":
            devices.sort(key=lambda d: (d.manufacturer.lower(), d.device_desc.lower()))
        else:
            devices.sort(key=lambda d: (d.device_desc.lower(), d.key_path.lower()))

        return devices

    def apply_view_filters(self):
        ordered = self.filtered_sorted_devices()
        visible_paths = {d.key_path for d in ordered}

        stale = [path for path in self.cards if path not in visible_paths]
        for path in stale:
            card = self.cards.pop(path)
            self.scroll_layout.removeWidget(card)
            card.setParent(None)
            card.deleteLater()

        for idx, device in enumerate(ordered):
            card = self.cards.get(device.key_path)
            if card is None:
                card = DeviceCard(device, self.set_device_sleep_state)
                self.cards[device.key_path] = card
            else:
                card.update_from_device(device)

            self.scroll_layout.removeWidget(card)
            self.scroll_layout.insertWidget(idx, card)

        self.status_label.setText(
            f"Showing {len(ordered)} of {len(self.latest_devices)} USB device parameter entries"
        )

    def set_device_sleep_state(self, key_path: str, disable_sleep: bool):
        target_value = 0 if disable_sleep else 1
        try:
            set_epm_value(key_path, target_value)
            self.status_label.setText(f"Updated: {key_path}")
            self.refresh_devices(silent=True)
            return
        except PermissionError:
            pass
        except OSError as exc:
            if not is_access_denied(exc):
                QMessageBox.critical(self, "Registry error", f"Failed to update {key_path}\n\n{exc}")
                self.refresh_devices(silent=True)
                return

        approved = self.ask_yes_no(
            "Administrator privileges required",
            "This change needs elevation. Prompt for Administrator permission now?",
        )
        if not approved:
            self.status_label.setText("Change canceled (elevation declined)")
            self.refresh_devices(silent=True)
            return

        result_path = self.make_result_file()
        if not result_path:
            QMessageBox.critical(self, "Error", "Could not allocate temporary result file for elevated action.")
            self.refresh_devices(silent=True)
            return

        key_escaped = shell_quote_ps(key_path)
        path = f"Registry::HKEY_LOCAL_MACHINE\\{key_escaped}"
        ps = (
            "$ErrorActionPreference='Stop';"
            "$r=@{success=$false;message=''};"
            f"try{{Set-ItemProperty -Path '{path}' -Name 'EnhancedPowerManagementEnabled' -Type DWord -Value {target_value};"
            "$r.success=$true;$r.message='Write completed';}"
            "catch{$r.message=$_.Exception.Message;}"
            f"$r|ConvertTo-Json -Compress|Set-Content -Path '{shell_quote_ps(result_path)}' -Encoding UTF8"
        )

        self.start_pending_operation(
            command=ps,
            result_path=result_path,
            user_status="Waiting for UAC approval...",
            success_status="Elevated write completed",
        )

    def disable_sleep_all(self):
        if not self.latest_devices:
            self.status_label.setText("No USB device entries found")
            return

        try:
            failures = disable_epm_for_all()
            if failures:
                self.status_label.setText(f"Disable all completed with {failures} failures")
            else:
                self.status_label.setText("USB sleep disabled for all listed devices")
            self.refresh_devices(silent=True)
            return
        except PermissionError:
            pass
        except OSError as exc:
            if not is_access_denied(exc):
                QMessageBox.critical(self, "Registry error", f"Disable-all failed.\n\n{exc}")
                self.refresh_devices(silent=True)
                return

        approved = self.ask_yes_no(
            "Administrator privileges required",
            "Disable-all needs elevation. Prompt for Administrator permission now?",
        )
        if not approved:
            self.status_label.setText("Disable-all canceled (elevation declined)")
            self.refresh_devices(silent=True)
            return

        result_path = self.make_result_file()
        if not result_path:
            QMessageBox.critical(self, "Error", "Could not allocate temporary result file for elevated action.")
            self.refresh_devices(silent=True)
            return

        ps = (
            "$ErrorActionPreference='Continue';"
            "$fails=0;$total=0;"
            "Get-ChildItem -Path 'HKLM:\\SYSTEM\\CurrentControlSet\\Enum\\USB' -Recurse -ErrorAction SilentlyContinue |"
            "Where-Object {$_.PSChildName -eq 'Device Parameters'} |"
            "ForEach-Object {$total++;try{Set-ItemProperty -Path $_.PSPath -Name 'EnhancedPowerManagementEnabled' -Type DWord -Value 0 -ErrorAction Stop}catch{$fails++}};"
            "$r=@{success=($fails -eq 0);message=(\"Processed $total entries; failures $fails\")};"
            f"$r|ConvertTo-Json -Compress|Set-Content -Path '{shell_quote_ps(result_path)}' -Encoding UTF8"
        )

        self.start_pending_operation(
            command=ps,
            result_path=result_path,
            user_status="Waiting for UAC approval for disable-all...",
            success_status="Elevated disable-all completed",
        )

    def ask_yes_no(self, title: str, text: str) -> bool:
        if PYQT_VER == 6:
            choice = QMessageBox.question(
                self,
                title,
                text,
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes,
            )
            return choice == QMessageBox.StandardButton.Yes

        choice = QMessageBox.question(
            self,
            title,
            text,
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes,
        )
        return choice == QMessageBox.Yes

    def make_result_file(self) -> Optional[str]:
        try:
            fd, path = tempfile.mkstemp(prefix="usb_power_flow_", suffix=".json")
            os.close(fd)
            os.unlink(path)
            return path
        except OSError:
            return None

    def start_pending_operation(self, command: str, result_path: str, user_status: str, success_status: str):
        if self.pending_operation is not None:
            QMessageBox.warning(self, "Operation in progress", "Wait for the current elevated operation to finish.")
            return

        launched = launch_elevated_powershell(command)
        if not launched:
            QMessageBox.critical(self, "Elevation failed", "Could not launch elevated helper process.")
            return

        self.pending_operation = {
            "result_path": result_path,
            "success_status": success_status,
            "started_at": time.time(),
        }
        self.status_label.setText(user_status)
        self.pending_timer.start(400)

    def poll_pending_operation(self):
        if not self.pending_operation:
            self.pending_timer.stop()
            return

        elapsed = time.time() - self.pending_operation["started_at"]
        result_path = self.pending_operation["result_path"]

        if elapsed > PENDING_TIMEOUT_SEC:
            self.pending_timer.stop()
            self.pending_operation = None
            self.status_label.setText("Elevated operation timed out")
            self.refresh_devices(silent=True)
            return

        if not os.path.exists(result_path):
            return

        try:
            with open(result_path, "r", encoding="utf-8-sig") as f:
                raw = f.read().strip()
            data = json.loads(raw) if raw else {}
        except Exception as exc:
            self.pending_timer.stop()
            self.pending_operation = None
            self.status_label.setText("Elevated operation returned unreadable output")
            QMessageBox.critical(self, "Elevation result error", str(exc))
            self.refresh_devices(silent=True)
            return
        finally:
            try:
                os.remove(result_path)
            except OSError:
                pass

        self.pending_timer.stop()
        pending = self.pending_operation
        self.pending_operation = None

        success = bool(data.get("success"))
        message = str(data.get("message", ""))

        if success:
            label = pending["success_status"]
            if message:
                label = f"{label} ({message})"
            self.status_label.setText(label)
        else:
            err = message or "Elevated action failed"
            self.status_label.setText("Elevated action failed")
            QMessageBox.critical(self, "Elevated action failed", err)

        self.refresh_devices(silent=True)


def main():
    app = QApplication(sys.argv)
    window = USBPowerMainWindow()
    window.show()
    sys.exit(app.exec() if PYQT_VER == 6 else app.exec_())


if __name__ == "__main__":
    main()
