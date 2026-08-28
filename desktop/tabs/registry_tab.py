# -*- coding: utf-8 -*-
"""📒 Registry — remember which products have already been created."""
from __future__ import annotations

import os

import pandas as pd
from PySide6.QtWidgets import (
    QFileDialog, QGroupBox, QLabel, QMessageBox, QPushButton, QVBoxLayout, QWidget,
)

import admin_core
import theme
from tabs.base import BaseTab
from widgets import DataFrameTable, MetricRow
from workers import Worker

METRICS = ["Checked", "Already created", "New"]


class RegistryTab(BaseTab):
    TITLE = "📒 Registry"
    NEEDS_ADMIN = True
    SUPPORTS = {"open"}

    def __init__(self, settings, parent=None):
        super().__init__(settings, parent)
        self.result = pd.DataFrame()
        self._worker = None
        self._panel = None

        lay = QVBoxLayout(self)
        lay.addWidget(QLabel(
            "Store past filled files (name + barcode + size), then check a barcode "
            "list against them to see what you've already created."
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

        store = QGroupBox("➕ Store created items")
        sv = QVBoxLayout(store)
        sv.addWidget(QLabel("Add one or more past filled files."))
        b_store = QPushButton("📂 Select filled file(s)…")
        b_store.clicked.connect(self._store)
        sv.addWidget(b_store)
        v.addWidget(store)

        check = QGroupBox("🔎 Check a barcode list")
        cv = QVBoxLayout(check)
        b_check = QPushButton("📂 Open barcode list…")
        b_check.clicked.connect(self.open_file)
        cv.addWidget(b_check)
        b_export = QPushButton("💾 Export result (CSV)…")
        b_export.clicked.connect(self._export)
        cv.addWidget(b_export)
        v.addWidget(check)

        v.addStretch(1)
        self._panel = panel
        return panel

    def _engine(self):
        try:
            return admin_core.get_engine(self.settings.db_url)
        except ValueError as e:
            QMessageBox.warning(self, "Admin not configured", str(e))
            return None

    @staticmethod
    def _read(path: str) -> pd.DataFrame:
        if path.lower().endswith(".csv"):
            df = pd.read_csv(path, dtype=str, sep=None, engine="python", on_bad_lines="skip")
        else:
            df = pd.read_excel(path, dtype=str, engine="openpyxl")
        df.columns = df.columns.astype(str).str.replace(r"\s+", " ", regex=True).str.strip()
        return df

    # ----------------------------------------------------------------- store
    def _store(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Select past filled file(s)", self.settings.last_dir("registry"),
            "Files (*.xlsx *.csv)",
        )
        if not paths:
            return
        self.settings.set_last_dir("registry", os.path.dirname(paths[0]))
        engine = self._engine()
        if engine is None:
            return

        frames, problems = [], []
        for path in paths:
            try:
                frames.append(admin_core.records_from_filled_file(self._read(path)))
            except Exception as e:
                problems.append(f"{os.path.basename(path)}: {e}")

        if problems:
            QMessageBox.warning(self, "Some files were skipped", "\n".join(problems))
        if not frames:
            return

        records = pd.concat(frames, ignore_index=True)
        self.busy.emit(True)
        self._worker = Worker(admin_core.store_created_items, engine, records, pass_progress=True)
        self._worker.progress.connect(lambda f, t: self.status_message.emit(t))

        def done(res):
            self.busy.emit(False)
            msg = (f"Stored {res['incoming']} row(s) from {len(frames)} file(s). "
                   f"Registry holds {res['total']:,} unique items ({res['added']:,} new).")
            self.note.setText(msg)
            self.status_message.emit(msg)
            QMessageBox.information(self, "Stored", msg)

        self._worker.done.connect(done)
        self._worker.failed.connect(self._on_failed)
        self._worker.start()

    # ----------------------------------------------------------------- check
    def open_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Open a barcode list", self.settings.last_dir("registry"),
            "Files (*.xlsx *.csv)",
        )
        if not path:
            return
        self.settings.set_last_dir("registry", os.path.dirname(path))
        engine = self._engine()
        if engine is None:
            return
        try:
            df = self._read(path)
        except Exception as e:
            QMessageBox.critical(self, "Could not open file", str(e))
            return

        lower = {str(c).lower(): c for c in df.columns}
        col = next((lower[c] for c in ("barcode", "ean", "upc", "ean/upc") if c in lower), None)
        if col is None:
            col = next((c for c in df.columns
                        if any(k in str(c).lower() for k in ("barcode", "ean", "upc"))), None)
        if col is None:
            QMessageBox.critical(
                self, "No barcode column",
                f"Couldn't find one.\n\nColumns: {', '.join(map(str, df.columns[:12]))}",
            )
            return

        barcodes = list(df[col])
        self.busy.emit(True)
        self._worker = Worker(admin_core.check_created_items, engine, barcodes)

        def done(res: pd.DataFrame):
            self.busy.emit(False)
            self.result = res
            already = int((res["Status"] == "Already created").sum())
            new = len(res) - already
            self.metrics.set("Checked", len(res))
            self.metrics.set("Already created", already, theme.STATUS_LOADING if already else None)
            self.metrics.set("New", new, theme.STATUS_READY if new else None)
            highlight = {
                (i, "Status"): ("warning" if s == "Already created" else "ok")
                for i, s in enumerate(res["Status"].head(self.table.MAX_ROWS))
            }
            self.table.set_dataframe(res, highlight=highlight)
            self.note.setText(f"{already} already created, {new} new.")
            self.status_message.emit(f"Checked {len(res)}: {already} already created, {new} new.")

        self._worker.done.connect(done)
        self._worker.failed.connect(self._on_failed)
        self._worker.start()

    def _export(self) -> None:
        if self.result.empty:
            QMessageBox.information(self, "Nothing to export", "Run a check first.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Export result",
            os.path.join(self.settings.last_dir("output") or "", "registry_check.csv"),
            "CSV (*.csv)",
        )
        if not path:
            return
        self.settings.set_last_dir("output", os.path.dirname(path))
        try:
            self.result.to_csv(path, index=False, encoding="utf-8-sig")
        except Exception as e:
            QMessageBox.critical(self, "Could not export", str(e))
            return
        self.status_message.emit(f"Exported → {path}")

    def _on_failed(self, message: str) -> None:
        self.busy.emit(False)
        box = QMessageBox(QMessageBox.Critical, "Failed", message, parent=self)
        detail = getattr(self._worker, "error_detail", "")
        if detail:
            box.setDetailedText(detail)
        box.exec()
