# -*- coding: utf-8 -*-
"""🏭 Catalogue — ingest a manufacturer file, and the danger-zone delete."""
from __future__ import annotations

import os

from PySide6.QtWidgets import (
    QComboBox, QFileDialog, QGroupBox, QLabel, QLineEdit, QMessageBox,
    QPushButton, QTreeWidget, QTreeWidgetItem, QVBoxLayout, QWidget,
)

import admin_core
import theme
from dictionaries import MANUFACTURER_CONFIG
from tabs.base import BaseTab
from widgets import MetricRow
from workers import Worker

METRICS = ["Rows", "Unique barcodes", "Unmapped values"]


class CatalogueTab(BaseTab):
    TITLE = "🏭 Catalogue"
    NEEDS_ADMIN = True
    SUPPORTS = {"open", "run"}

    def __init__(self, settings, parent=None):
        super().__init__(settings, parent)
        self.file_path = ""
        self._worker = None
        self._panel = None

        lay = QVBoxLayout(self)
        lay.addWidget(QLabel(
            "Upload a raw manufacturer catalogue. It is translated into our global "
            "categories and merged into the master catalogue by barcode."
        ))

        self.file_label = QLabel("No file selected.")
        self.file_label.setWordWrap(True)
        lay.addWidget(self.file_label)

        self.metrics = MetricRow(METRICS)
        lay.addWidget(self.metrics)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Result", "Detail"])
        self.tree.setColumnWidth(0, 460)
        lay.addWidget(self.tree, 1)

    # ------------------------------------------------------------------ dock
    def control_panel(self) -> QWidget:
        if self._panel is not None:
            return self._panel
        panel = QWidget()
        v = QVBoxLayout(panel)

        box = QGroupBox("Ingest")
        bv = QVBoxLayout(box)
        bv.addWidget(QLabel("Manufacturer"))
        self.mfg = QComboBox()
        self.mfg.addItems(sorted(MANUFACTURER_CONFIG.keys()))
        saved = str(self.settings.get("catalogue/mfg", "") or "")
        if saved and saved in MANUFACTURER_CONFIG:
            self.mfg.setCurrentText(saved)
        self.mfg.currentTextChanged.connect(
            lambda t: self.settings.set("catalogue/mfg", t)
        )
        bv.addWidget(self.mfg)

        b_open = QPushButton("📂 Choose catalogue file…")
        b_open.clicked.connect(self.open_file)
        bv.addWidget(b_open)

        b_run = QPushButton("🚀 Process & merge")
        b_run.clicked.connect(self.run_action)
        bv.addWidget(b_run)

        note = QLabel(
            "The format is auto-detected, so Safilo daily CSVs, Safilo catalogue "
            "exports and the Tom Ford file all go under their own manufacturer."
        )
        note.setWordWrap(True)
        note.setStyleSheet("color: #888; font-size: 11px;")
        bv.addWidget(note)
        v.addWidget(box)

        danger = QGroupBox("⚠️ Danger zone")
        dv = QVBoxLayout(danger)
        dv.addWidget(QLabel("Delete every row of one manufacturer"))
        self.del_mfg = QComboBox()
        self.del_mfg.addItems(sorted(MANUFACTURER_CONFIG.keys()))
        dv.addWidget(self.del_mfg)
        self.confirm = QLineEdit()
        self.confirm.setPlaceholderText("Type DELETE <NAME> to confirm")
        dv.addWidget(self.confirm)
        b_del = QPushButton("🗑 Delete rows")
        b_del.clicked.connect(self._delete)
        dv.addWidget(b_del)
        v.addWidget(danger)

        v.addStretch(1)
        self._panel = panel
        return panel

    # ------------------------------------------------------------------ open
    def open_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Choose the manufacturer catalogue",
            self.settings.last_dir("catalogue"),
            "Catalogues (*.xlsx *.csv *.zip);;All files (*.*)",
        )
        if not path:
            return
        self.settings.set_last_dir("catalogue", os.path.dirname(path))
        self.file_path = path
        self.file_label.setText(f"📂 {path}")
        self.tree.clear()
        for m in METRICS:
            self.metrics.set(m, "—", None)

    # ------------------------------------------------------------------- run
    def run_action(self) -> None:
        if not self.file_path:
            QMessageBox.information(self, "No file", "Choose a catalogue file first (📂).")
            return
        mfg = self.mfg.currentText()
        if QMessageBox.question(
            self, "Process catalogue",
            f"Process this file as “{mfg}” and merge it into the master catalogue?\n\n"
            f"{os.path.basename(self.file_path)}",
        ) != QMessageBox.Yes:
            return

        try:
            engine = admin_core.get_engine(self.settings.db_url)
        except ValueError as e:
            QMessageBox.warning(self, "Admin not configured", str(e))
            return

        path = self.file_path
        self.busy.emit(True)
        self._worker = Worker(
            admin_core.process_catalogue, engine, mfg, path, pass_progress=True
        )
        self._worker.progress.connect(lambda f, t: self.status_message.emit(t))
        self._worker.done.connect(self._on_done)
        self._worker.failed.connect(self._on_failed)
        self._worker.start()

    def _on_done(self, result: dict) -> None:
        self.busy.emit(False)
        self.metrics.set("Rows", f"{result['rows']:,}")
        self.metrics.set("Unique barcodes", f"{result['unique']:,}", theme.STATUS_READY)
        n_unmapped = len(result["unmapped"])
        self.metrics.set(
            "Unmapped values", n_unmapped,
            theme.STATUS_LOADING if n_unmapped else theme.STATUS_READY,
        )

        self.tree.clear()
        self.tree.addTopLevelItem(QTreeWidgetItem(["✅ " + result["message"], ""]))
        if result["unmapped"]:
            root = QTreeWidgetItem([
                f"⚠️ Unmapped values ({n_unmapped})",
                "These kept their original text and need a mapping added.",
            ])
            for v in result["unmapped"]:
                root.addChild(QTreeWidgetItem(["", v]))
            self.tree.addTopLevelItem(root)
        if result["skipped"]:
            root = QTreeWidgetItem([f"ℹ️ Skipped 'NOT MAPPED' ({len(result['skipped'])})", ""])
            for v in result["skipped"]:
                root.addChild(QTreeWidgetItem(["", v]))
            self.tree.addTopLevelItem(root)
        self.tree.expandToDepth(0)
        self.status_message.emit(result["message"])
        note = self.snapshot_note()
        if note:
            QMessageBox.information(self, "Merged", result["message"] + note)

    def _on_failed(self, message: str) -> None:
        self.busy.emit(False)
        box = QMessageBox(QMessageBox.Critical, "Ingest failed", message, parent=self)
        detail = getattr(self._worker, "error_detail", "")
        if detail:
            box.setDetailedText(detail)
        box.exec()

    # ---------------------------------------------------------------- delete
    def _delete(self) -> None:
        mfg = self.del_mfg.currentText()
        expected = f"DELETE {mfg.upper()}"
        if self.confirm.text().strip() != expected:
            QMessageBox.warning(
                self, "Confirmation doesn't match",
                f"Type exactly “{expected}” to confirm. Nothing was deleted.",
            )
            return
        try:
            engine = admin_core.get_engine(self.settings.db_url)
        except ValueError as e:
            QMessageBox.warning(self, "Admin not configured", str(e))
            return

        self.busy.emit(True)
        self._worker = Worker(admin_core.delete_manufacturer, engine, mfg, pass_progress=True)
        self._worker.progress.connect(lambda f, t: self.status_message.emit(t))

        def done(res):
            self.busy.emit(False)
            self.confirm.clear()
            msg = f"Deleted {res['deleted']:,} {mfg.title()} row(s). {res['remaining']:,} remain."
            self.status_message.emit(msg)
            QMessageBox.information(self, "Deleted", msg + self.snapshot_note())

        self._worker.done.connect(done)
        self._worker.failed.connect(self._on_failed)
        self._worker.start()
