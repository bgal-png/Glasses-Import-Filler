# -*- coding: utf-8 -*-
"""⚙️ Settings — data source, admin unlock, optional AI key, cache."""
from __future__ import annotations

from PySide6.QtWidgets import (
    QDialog, QDialogButtonBox, QFormLayout, QGroupBox, QHBoxLayout, QLabel,
    QLineEdit, QMessageBox, QPushButton, QVBoxLayout,
)

import data_source


class SettingsDialog(QDialog):
    def __init__(self, settings, parent=None):
        super().__init__(parent)
        self.settings = settings
        self.setWindowTitle("Settings")
        self.setMinimumWidth(620)

        v = QVBoxLayout(self)

        # --- catalogue snapshot (read-only) ---
        snap = QGroupBox("☁️ Catalogue data (read-only snapshot)")
        sf = QFormLayout(snap)
        self.repo = QLineEdit(settings.snapshot_repo)
        self.repo.setPlaceholderText("owner/repo  (the private data repo)")
        self.token = QLineEdit(settings.snapshot_token)
        self.token.setPlaceholderText("GitHub token with read access to that repo")
        self.token.setEchoMode(QLineEdit.Password)
        self.branch = QLineEdit(settings.snapshot_branch)
        sf.addRow("Snapshot repo", self.repo)
        sf.addRow("Snapshot token", self.token)
        sf.addRow("Branch", self.branch)
        hint = QLabel(
            "The snapshot is published by the <i>publish-snapshot</i> GitHub Action "
            "after every catalogue ingest. Read-only — it cannot change the database."
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("color: #888; font-size: 11px;")
        sf.addRow(hint)
        v.addWidget(snap)

        # --- admin ---
        admin = QGroupBox("🔐 Admin access (optional)")
        af = QFormLayout(admin)
        self.db_url = QLineEdit(settings.db_url)
        self.db_url.setPlaceholderText("postgresql://…  (leave empty on colleagues' machines)")
        self.db_url.setEchoMode(QLineEdit.Password)
        af.addRow("Database URL", self.db_url)
        admin_hint = QLabel(
            "Filling this in unlocks the catalogue-upload and editing tabs and is "
            "stored only on this machine. It grants <b>write</b> access to the live "
            "database — never share a build with it pre-filled."
        )
        admin_hint.setWordWrap(True)
        admin_hint.setStyleSheet("color: #888; font-size: 11px;")
        af.addRow(admin_hint)
        v.addWidget(admin)

        # --- AI ---
        ai = QGroupBox("👓 AI shape recognition (optional)")
        aif = QFormLayout(ai)
        self.api_key = QLineEdit(settings.anthropic_key)
        self.api_key.setPlaceholderText("sk-ant-…")
        self.api_key.setEchoMode(QLineEdit.Password)
        aif.addRow("Anthropic API key", self.api_key)
        v.addWidget(ai)

        # --- cache ---
        cache = QGroupBox("🗄️ Local cache")
        cv = QHBoxLayout(cache)
        self.cache_label = QLabel(self._cache_text())
        cv.addWidget(self.cache_label, 1)
        clear = QPushButton("Clear cache")
        clear.clicked.connect(self._clear_cache)
        cv.addWidget(clear)
        v.addWidget(cache)

        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self._save)
        buttons.rejected.connect(self.reject)
        v.addWidget(buttons)

    def _cache_text(self) -> str:
        try:
            mb = data_source.cache_size_bytes() / (1024 * 1024)
        except Exception:
            mb = 0.0
        return f"Cached catalogue: {mb:,.1f} MB"

    def _clear_cache(self) -> None:
        try:
            data_source.clear_cache()
        except Exception as e:
            QMessageBox.warning(self, "Could not clear cache", str(e))
            return
        self.cache_label.setText(self._cache_text())
        QMessageBox.information(
            self, "Cache cleared",
            "The catalogue will be downloaded again on the next ☁️ Refresh data.",
        )

    def _save(self) -> None:
        was_admin = self.settings.admin_enabled
        self.settings.snapshot_repo = self.repo.text()
        self.settings.snapshot_token = self.token.text()
        self.settings.set("snapshot_branch", self.branch.text().strip() or "main")
        self.settings.db_url = self.db_url.text()
        self.settings.anthropic_key = self.api_key.text()

        if self.settings.admin_enabled != was_admin:
            QMessageBox.information(
                self, "Restart needed",
                "Admin tabs are added or removed when the app starts.\n\n"
                "Close and reopen the app to apply that change.",
            )
        self.accept()
