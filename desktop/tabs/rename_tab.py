# -*- coding: utf-8 -*-
"""✏️ Rename — bulk-set product names from a barcode → name list."""
from __future__ import annotations

import os

import pandas as pd
from PySide6.QtWidgets import (
    QComboBox, QFileDialog, QGroupBox, QLabel, QMessageBox, QPushButton,
    QVBoxLayout, QWidget,
)

import admin_core
import theme
from tabs.base import BaseTab
from widgets import DataFrameTable, MetricRow
from workers import Worker

METRICS = ["Rows in file", "Renamed", "Not found"]


class RenameTab(BaseTab):
    TITLE = "✏️ Rename"
    NEEDS_ADMIN = True
    SUPPORTS = {"open", "run"}

    def __init__(self, settings, parent=None):
        super().__init__(settings, parent)
        self.df = pd.DataFrame()
        self._worker = None
        self._panel = None

        lay = QVBoxLayout(self)
        lay.addWidget(QLabel(
            "Open a file with a barcode column and a name column. Every barcode "
            "found in the catalogue gets its product name replaced."
        ))
        self.metrics = MetricRow(METRICS)
        lay.addWidget(self.metrics)
        self.table = DataFrameTable()
        lay.addWidget(self.table, 1)
        self.note = QLabel("")
        self.note.setStyleSheet("color: #888;")
        lay.addWidget(self.note)

    def control_panel(self) -> QWidget:
        if self._panel is not None:
            return self._panel
        panel = QWidget()
        v = QVBoxLayout(panel)

        box = QGroupBox("Columns")
        bv = QVBoxLayout(box)
        bv.addWidget(QLabel("Barcode column"))
        self.bc_col = QComboBox()
        bv.addWidget(self.bc_col)
        bv.addWidget(QLabel("Name column"))
        self.name_col = QComboBox()
        bv.addWidget(self.name_col)
        v.addWidget(box)

        b_open = QPushButton("📂 Open name list…")
        b_open.clicked.connect(self.open_file)
        v.addWidget(b_open)

        b_run = QPushButton("✏️ Apply renames")
        b_run.clicked.connect(self.run_action)
        v.addWidget(b_run)

        v.addStretch(1)
        self._panel = panel
        return panel

    # ------------------------------------------------------------------ open
    def open_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Open the barcode → name list", self.settings.last_dir("rename"),
            "Files (*.xlsx *.csv)",
        )
        if not path:
            return
        self.settings.set_last_dir("rename", os.path.dirname(path))
        try:
            if path.lower().endswith(".csv"):
                df = pd.read_csv(path, dtype=str, sep=None, engine="python", on_bad_lines="skip")
            else:
                df = pd.read_excel(path, dtype=str, engine="openpyxl")
        except Exception as e:
            QMessageBox.critical(self, "Could not open file", str(e))
            return

        df.columns = df.columns.astype(str).str.replace(r"\s+", " ", regex=True).str.strip()
        self.df = df
        cols = [str(c) for c in df.columns]

        self.bc_col.clear()
        self.bc_col.addItems(cols)
        self.name_col.clear()
        self.name_col.addItems(cols)

        def guess(combo, candidates):
            lower = {c.lower(): c for c in cols}
            for cand in candidates:
                if cand in lower:
                    combo.setCurrentText(lower[cand])
                    return
        guess(self.bc_col, ["barcode", "ean", "upc", "ean/upc"])
        guess(self.name_col, ["glasses name", "name", "assembled_name", "xml description"])

        self.metrics.set("Rows in file", len(df))
        for m in ("Renamed", "Not found"):
            self.metrics.set(m, "—", None)
        self.table.set_dataframe(df)
        self.note.setText(f"Loaded {len(df)} row(s) from {os.path.basename(path)}.")

    # ------------------------------------------------------------------- run
    def run_action(self) -> None:
        if self.df.empty:
            QMessageBox.information(self, "No file", "Open a name list first (📂).")
            return
        bc = self.bc_col.currentText()
        nm = self.name_col.currentText()
        if not bc or not nm:
            QMessageBox.information(self, "Pick the columns", "Choose the barcode and name columns.")
            return
        if bc == nm:
            QMessageBox.warning(self, "Same column twice",
                                "The barcode and name columns must be different.")
            return

        mapping = {}
        for _i, row in self.df.iterrows():
            mapping[row[bc]] = row[nm]

        if QMessageBox.question(
            self, "Apply renames",
            f"Rename up to {len(mapping)} product(s) in the catalogue?",
        ) != QMessageBox.Yes:
            return

        try:
            engine = admin_core.get_engine(self.settings.db_url)
        except ValueError as e:
            QMessageBox.warning(self, "Admin not configured", str(e))
            return

        self.busy.emit(True)
        self._worker = Worker(admin_core.apply_renames, engine, mapping, pass_progress=True)
        self._worker.progress.connect(lambda f, t: self.status_message.emit(t))
        self._worker.done.connect(self._on_done)
        self._worker.failed.connect(self._on_failed)
        self._worker.start()

    def _on_done(self, res: dict) -> None:
        self.busy.emit(False)
        self.metrics.set("Renamed", res["updated"], theme.STATUS_READY)
        n_missing = len(res["not_found"])
        self.metrics.set("Not found", n_missing, theme.STATUS_ERROR if n_missing else None)
        if res["not_found"]:
            self.table.set_dataframe(pd.DataFrame({"Barcode not in catalogue": res["not_found"]}))
            self.note.setText(
                f"{n_missing} barcode(s) were not in the catalogue — shown above."
            )
        else:
            self.note.setText("Every barcode was found.")
        msg = f"Renamed {res['updated']} row(s)."
        self.status_message.emit(msg)
        QMessageBox.information(self, "Done", msg + self.snapshot_note())

    def _on_failed(self, message: str) -> None:
        self.busy.emit(False)
        box = QMessageBox(QMessageBox.Critical, "Rename failed", message, parent=self)
        detail = getattr(self._worker, "error_detail", "")
        if detail:
            box.setDetailedText(detail)
        box.exec()
