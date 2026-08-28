# -*- coding: utf-8 -*-
"""🔍 Barcode Checker — which barcodes are already in the catalogue?

Desktop equivalent of the standalone barcode_checker.py web app, plus a
paste-a-list box (no file needed) and a “last catalogue update per producer”
panel in the control dock.
"""
from __future__ import annotations

import os
import re
from datetime import datetime

import pandas as pd
from PySide6.QtWidgets import (
    QFileDialog, QGroupBox, QHBoxLayout, QLabel, QMessageBox, QPlainTextEdit,
    QPushButton, QTreeWidget, QTreeWidgetItem, QVBoxLayout, QWidget,
)

import theme
from tabs.base import BaseTab
from widgets import DataFrameTable, MetricRow

METRICS = ["Checked", "In database", "Not in database"]


def _clean_bc(x) -> str:
    return re.sub(r"\.0$", "", str(x).strip()).lstrip("0")


def _find_barcode_col(df: pd.DataFrame):
    lower = {str(c).lower(): c for c in df.columns}
    for cand in ("barcode", "ean", "upc", "ean/upc", "ean code", "* ean code"):
        if cand in lower:
            return lower[cand]
    for c in df.columns:
        if any(k in str(c).lower() for k in ("barcode", "ean", "upc")):
            return c
    return None


class BarcodeTab(BaseTab):
    TITLE = "🔍 Barcode Checker"
    SUPPORTS = {"open"}

    def __init__(self, settings, parent=None):
        super().__init__(settings, parent)
        self.result_df = pd.DataFrame()
        self._panel = None

        lay = QVBoxLayout(self)
        lay.addWidget(QLabel(
            "Open a file with a <b>Barcode</b> column, or paste barcodes in the "
            "control panel, to see which are already in the catalogue."
        ))

        self.metrics = MetricRow(METRICS)
        lay.addWidget(self.metrics)

        self.table = DataFrameTable()
        lay.addWidget(self.table, 1)

        self.note = QLabel("")
        self.note.setStyleSheet("color: #888;")
        lay.addWidget(self.note)

    # ------------------------------------------------------------------ dock
    def control_panel(self) -> QWidget:
        if self._panel is not None:
            self._refresh_updates_tree()
            return self._panel

        panel = QWidget()
        v = QVBoxLayout(panel)

        paste_box = QGroupBox("Paste barcodes")
        pv = QVBoxLayout(paste_box)
        self.paste_edit = QPlainTextEdit()
        self.paste_edit.setPlaceholderText("One barcode per line…")
        self.paste_edit.setMaximumHeight(160)
        pv.addWidget(self.paste_edit)
        b_check = QPushButton("🔍 Check pasted list")
        b_check.clicked.connect(self._check_pasted)
        pv.addWidget(b_check)
        v.addWidget(paste_box)

        b_open = QPushButton("📂 Open a file instead…")
        b_open.clicked.connect(self.open_file)
        v.addWidget(b_open)

        b_export = QPushButton("💾 Export result (CSV)…")
        b_export.clicked.connect(self._export)
        v.addWidget(b_export)

        upd_box = QGroupBox("🗓️ Last catalogue update")
        uv = QVBoxLayout(upd_box)
        self.updates_tree = QTreeWidget()
        self.updates_tree.setHeaderLabels(["Producer", "Last update"])
        self.updates_tree.setRootIsDecorated(False)
        self.updates_tree.setColumnWidth(0, 120)
        uv.addWidget(self.updates_tree)
        v.addWidget(upd_box)

        v.addStretch(1)
        self._panel = panel
        self._refresh_updates_tree()
        return panel

    def _refresh_updates_tree(self) -> None:
        if not hasattr(self, "updates_tree"):
            return
        self.updates_tree.clear()
        log = getattr(self.catalogue, "ingest_log", None) if self.catalogue else None

        rows = {}
        if log is not None and not log.empty and "manufacturer" in log.columns:
            for _, r in log.iterrows():
                rows[str(r["manufacturer"]).strip().lower()] = str(r.get("last_updated", ""))

        try:
            from dictionaries import MANUFACTURER_CONFIG
            producers = list(MANUFACTURER_CONFIG.keys())
        except Exception:
            producers = []
        for extra in rows:
            if extra not in producers:
                producers.append(extra)

        def fmt(iso: str) -> str:
            if not iso or iso.lower() in ("none", "nan"):
                return "never"
            try:
                dt = datetime.fromisoformat(iso.replace("Z", ""))
                return f"{dt.day}.{dt.month}.{dt.year}"
            except Exception:
                return iso

        for p in producers:
            self.updates_tree.addTopLevelItem(
                QTreeWidgetItem([p.title(), fmt(rows.get(p.lower(), ""))])
            )

    def on_catalogue(self, data) -> None:
        self._refresh_updates_tree()

    # ----------------------------------------------------------------- check
    def _known_keys(self):
        if self.catalogue is None or self.catalogue.is_empty:
            return None, None
        master = self.catalogue.master_db
        if "join_key" in master.columns:
            keys = master["join_key"].astype(str).str.strip()
            names = master["Assembled_Name"] if "Assembled_Name" in master.columns else None
        elif master.index.name == "join_key":
            keys = pd.Series(master.index.astype(str).str.strip())
            names = master["Assembled_Name"].reset_index(drop=True) if "Assembled_Name" in master.columns else None
        else:
            return None, None
        lookup = dict(zip(keys, names.astype(str).fillna(""))) if names is not None else {k: "" for k in keys}
        return set(lookup), lookup

    def _check(self, barcodes: list) -> None:
        known, lookup = self._known_keys()
        if known is None:
            QMessageBox.warning(
                self, "No catalogue data",
                "The catalogue hasn't loaded yet. Use ☁️ Refresh data, or check ⚙️ Settings.",
            )
            return

        rows = []
        for raw in barcodes:
            raw_s = str(raw).strip()
            if not raw_s or raw_s.lower() == "nan":
                continue
            key = _clean_bc(raw_s)
            in_db = key in known
            rows.append({
                "Barcode": raw_s,
                "Status": "✅ In database" if in_db else "❌ Not in database",
                "Product": (lookup.get(key, "") if in_db else ""),
            })

        if not rows:
            self.note.setText("No barcodes found in that input.")
            return

        self.result_df = pd.DataFrame(rows)
        n_in = int((self.result_df["Status"] == "✅ In database").sum())
        n_out = len(self.result_df) - n_in
        self.metrics.set("Checked", len(self.result_df))
        self.metrics.set("In database", n_in, theme.STATUS_READY)
        self.metrics.set("Not in database", n_out, theme.STATUS_ERROR if n_out else None)

        highlight = {
            (i, "Status"): ("ok" if s == "✅ In database" else "error")
            for i, s in enumerate(self.result_df["Status"].head(self.table.MAX_ROWS))
        }
        self.table.set_dataframe(self.result_df, highlight=highlight)
        hidden = self.table.truncated()
        self.note.setText(f"Showing the first {self.table.MAX_ROWS} of {len(self.result_df)} results."
                          if hidden else "")
        self.status_message.emit(f"Checked {len(self.result_df)}: {n_in} in database, {n_out} not.")

    def _check_pasted(self) -> None:
        text = self.paste_edit.toPlainText()
        parts = [p.strip() for p in re.split(r"[\s,;]+", text) if p.strip()]
        if not parts:
            QMessageBox.information(self, "Nothing to check", "Paste some barcodes first.")
            return
        self._check(parts)

    def open_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Open a file with barcodes", self.settings.last_dir("barcodes"),
            "Import files (*.xlsx *.csv);;Excel (*.xlsx);;CSV (*.csv)",
        )
        if not path:
            return
        self.settings.set_last_dir("barcodes", os.path.dirname(path))
        try:
            if path.lower().endswith(".csv"):
                df = pd.read_csv(path, dtype=str, sep=None, engine="python", on_bad_lines="skip")
            else:
                df = pd.read_excel(path, dtype=str, engine="openpyxl")
        except Exception as e:
            QMessageBox.critical(self, "Could not open file", str(e))
            return

        df.columns = df.columns.astype(str).str.replace(r"\s+", " ", regex=True).str.strip()
        col = _find_barcode_col(df)
        if col is None:
            QMessageBox.critical(
                self, "No barcode column",
                "Couldn't find a Barcode column in that file.\n\n"
                f"Columns found: {', '.join(map(str, df.columns[:12]))}",
            )
            return
        if str(col).lower() != "barcode":
            self.status_message.emit(f"Using column “{col}” as the barcode column.")
        self._check(list(df[col]))

    def _export(self) -> None:
        if self.result_df.empty:
            QMessageBox.information(self, "Nothing to export", "Run a check first.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Export result", os.path.join(
                self.settings.last_dir("output") or "", "barcode_check_result.csv"
            ), "CSV (*.csv)",
        )
        if not path:
            return
        self.settings.set_last_dir("output", os.path.dirname(path))
        try:
            self.result_df.to_csv(path, index=False, encoding="utf-8-sig")
        except Exception as e:
            QMessageBox.critical(self, "Could not export", str(e))
            return
        self.status_message.emit(f"Exported → {path}")
