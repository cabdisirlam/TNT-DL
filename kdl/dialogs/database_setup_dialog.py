"""
Database setup dialog.
Stores Oracle profile details for future integrations while keeping
current UI automation mode as the default loading path.
"""

from typing import Dict, List

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, QPushButton,
    QLineEdit, QCheckBox, QMessageBox, QDialogButtonBox, QGridLayout,
    QScrollArea, QWidget,
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QGuiApplication
from kdl.styles import dialog_qss


DB_MODES = [
    ("UI Automation (Recommended)", "ui_automation"),
]


class DatabaseSetupDialog(QDialog):
    """Manage database mode and stored connection profiles."""

    def __init__(self, current_settings: Dict, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Setup Databases")
        self.setMinimumWidth(460)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.setWindowFlag(Qt.WindowCloseButtonHint, True)

        db = current_settings.get("database", {}) if isinstance(current_settings, dict) else {}
        self._profiles: List[Dict] = list(db.get("profiles", []))
        self._active_profile = db.get("active_profile", "")
        self._mode = "ui_automation"

        from kdl.config_store import get_dark_mode
        self.setStyleSheet(dialog_qss(dark=get_dark_mode()))
        self._build_ui()
        self._load_profiles_to_combo()
        self._load_mode()
        self._select_initial_profile()
        self._fit_to_screen()

    def _build_ui(self):
        dialog_layout = QVBoxLayout(self)
        dialog_layout.setContentsMargins(0, 0, 0, 0)
        dialog_layout.setSpacing(0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.NoFrame)
        dialog_layout.addWidget(scroll)

        scroll_widget = QWidget()
        scroll.setWidget(scroll_widget)
        layout = QVBoxLayout(scroll_widget)
        layout.setSpacing(8)

        note = QLabel(
            "NT_DL currently loads data through IFMIS/Oracle front-end automation.\n"
            "Direct Oracle DB loading is disabled in this build for safety and consistency."
        )
        note.setWordWrap(True)
        layout.addWidget(note)

        mode_row = QHBoxLayout()
        mode_row.addWidget(QLabel("Mode:"))
        self.mode_combo = QComboBox()
        for label, key in DB_MODES:
            self.mode_combo.addItem(label, userData=key)
        mode_row.addWidget(self.mode_combo, 1)
        layout.addLayout(mode_row)

        profile_row = QHBoxLayout()
        profile_row.addWidget(QLabel("Profile:"))
        self.profile_combo = QComboBox()
        self.profile_combo.currentIndexChanged.connect(self._on_profile_selected)
        profile_row.addWidget(self.profile_combo, 1)

        new_btn = QPushButton("New")
        new_btn.clicked.connect(self._new_profile)
        profile_row.addWidget(new_btn)

        delete_btn = QPushButton("Delete")
        delete_btn.clicked.connect(self._delete_profile)
        profile_row.addWidget(delete_btn)
        layout.addLayout(profile_row)

        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(6)

        grid.addWidget(QLabel("Profile Name:"), 0, 0)
        self.name_edit = QLineEdit()
        grid.addWidget(self.name_edit, 0, 1)

        grid.addWidget(QLabel("Host:"), 1, 0)
        self.host_edit = QLineEdit()
        grid.addWidget(self.host_edit, 1, 1)

        grid.addWidget(QLabel("Port:"), 2, 0)
        self.port_edit = QLineEdit("1521")
        grid.addWidget(self.port_edit, 2, 1)

        grid.addWidget(QLabel("Service/SID:"), 3, 0)
        self.service_edit = QLineEdit()
        grid.addWidget(self.service_edit, 3, 1)

        grid.addWidget(QLabel("Username:"), 4, 0)
        self.username_edit = QLineEdit()
        grid.addWidget(self.username_edit, 4, 1)

        grid.addWidget(QLabel("Password:"), 5, 0)
        self.password_edit = QLineEdit()
        self.password_edit.setEchoMode(QLineEdit.Password)
        grid.addWidget(self.password_edit, 5, 1)

        self.store_password_check = QCheckBox("Store password in local settings")
        self.store_password_check.setChecked(False)
        self.store_password_check.setEnabled(False)
        self.store_password_check.setToolTip("Password storage is disabled in this build.")
        grid.addWidget(self.store_password_check, 6, 1)

        layout.addLayout(grid)

        save_btn_row = QHBoxLayout()
        save_btn_row.addStretch()
        save_profile_btn = QPushButton("Save Profile")
        save_profile_btn.clicked.connect(self._save_profile)
        save_btn_row.addWidget(save_profile_btn)
        layout.addLayout(save_btn_row)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        cancel_button = buttons.button(QDialogButtonBox.Cancel)
        if cancel_button is not None:
            cancel_button.setText("Close")
        buttons.accepted.connect(self._accept_dialog)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _fit_to_screen(self):
        screen = self.screen() or QGuiApplication.primaryScreen()
        if not screen:
            return
        geo = screen.availableGeometry()
        max_w = max(360, geo.width() - 24)
        max_h = max(320, geo.height() - 24)
        self.setMaximumSize(max_w, max_h)
        hint = self.sizeHint()
        target_w = min(max(self.minimumWidth(), hint.width()), max_w)
        target_h = min(max(360, hint.height()), max_h)
        self.resize(target_w, target_h)

    def _load_mode(self):
        idx = 0
        for i in range(self.mode_combo.count()):
            if self.mode_combo.itemData(i) == self._mode:
                idx = i
                break
        self.mode_combo.setCurrentIndex(idx)

    def _load_profiles_to_combo(self):
        self.profile_combo.blockSignals(True)
        self.profile_combo.clear()
        self.profile_combo.addItem("(None)", userData="")
        for profile in self._profiles:
            name = profile.get("name", "").strip()
            if name:
                self.profile_combo.addItem(name, userData=name)
        self.profile_combo.blockSignals(False)

    def _select_initial_profile(self):
        if not self._active_profile:
            self.profile_combo.setCurrentIndex(0)
            self._clear_fields()
            return

        idx = self.profile_combo.findData(self._active_profile)
        if idx < 0:
            idx = self.profile_combo.findText(self._active_profile)
        self.profile_combo.setCurrentIndex(max(0, idx))

    def _new_profile(self):
        self.profile_combo.setCurrentIndex(0)
        self._clear_fields()
        self.name_edit.setFocus()

    def _delete_profile(self):
        name = self.profile_combo.currentData() or self.profile_combo.currentText()
        if not name or name == "(None)":
            return

        reply = QMessageBox.question(
            self,
            "Delete Profile",
            f"Delete profile '{name}'?",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        self._profiles = [p for p in self._profiles if p.get("name") != name]
        if self._active_profile == name:
            self._active_profile = ""
        self._load_profiles_to_combo()
        self.profile_combo.setCurrentIndex(0)
        self._clear_fields()

    def _on_profile_selected(self):
        name = self.profile_combo.currentData()
        if not name:
            self._clear_fields()
            return

        profile = self._find_profile(name)
        if not profile:
            self._clear_fields()
            return

        self.name_edit.setText(profile.get("name", ""))
        self.host_edit.setText(profile.get("host", ""))
        self.port_edit.setText(str(profile.get("port", "1521")))
        self.service_edit.setText(profile.get("service", ""))
        self.username_edit.setText(profile.get("username", ""))
        self.password_edit.clear()
        self.store_password_check.setChecked(False)

    def _clear_fields(self):
        self.name_edit.clear()
        self.host_edit.clear()
        self.port_edit.setText("1521")
        self.service_edit.clear()
        self.username_edit.clear()
        self.password_edit.clear()
        self.store_password_check.setChecked(False)

    def _find_profile(self, name: str):
        for profile in self._profiles:
            if profile.get("name") == name:
                return profile
        return None

    def _save_profile(self):
        name = self.name_edit.text().strip()
        if not name:
            QMessageBox.warning(self, "Missing Name", "Profile Name is required.")
            return

        port_text = self.port_edit.text().strip() or "1521"
        try:
            port = int(port_text)
        except ValueError:
            QMessageBox.warning(self, "Invalid Port", "Port must be a number.")
            return

        existing = self._find_profile(name)
        preserved_password = ""
        if existing and isinstance(existing, dict):
            preserved_password = str(existing.get("password", "") or "")

        profile = {
            "name": name,
            "host": self.host_edit.text().strip(),
            "port": port,
            "service": self.service_edit.text().strip(),
            "username": self.username_edit.text().strip(),
            "password": preserved_password,
        }

        if existing is None:
            self._profiles.append(profile)
        else:
            existing.update(profile)

        self._active_profile = name
        self._load_profiles_to_combo()
        idx = self.profile_combo.findData(name)
        if idx >= 0:
            self.profile_combo.setCurrentIndex(idx)

    def _accept_dialog(self):
        self._mode = "ui_automation"
        current = self.profile_combo.currentData()
        self._active_profile = current if current else ""
        self.accept()

    def get_settings(self) -> Dict:
        persisted_profiles = []
        for profile in self._profiles:
            if not isinstance(profile, dict):
                continue
            clean = dict(profile)
            persisted_profiles.append(clean)

        return {
            "database": {
                "mode": "ui_automation",
                "active_profile": self._active_profile,
                "profiles": persisted_profiles,
            }
        }
