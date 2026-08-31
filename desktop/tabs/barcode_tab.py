# -*- coding: utf-8 -*-
"""🔍 Barcode Checker — two jobs on one tab.

  📋 List check      which of these barcodes are in the catalogue?
  🔎 Single barcode  look one up, see every field, and (admin only) edit it

The lookup half is deliberately available without a DB_URL — colleagues can
inspect a product from the snapshot. Only editing needs admin, because it
writes to the live database.
"""
from __future__ import annotations

import os
import re
from datetime import datetime

import pandas as pd
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFileDialog, QGroupBox, QHBoxLayout, QLabel, QLineEdit, QMessageBox,
    QPlainTextEdit, QPushButton, QTableWidget, QTableWidgetItem, QTabWidget,
    QTreeWidget, QTreeWidgetItem, QVBoxLayout, QWidget,
)

import admin_core
import theme
from tabs.base import BaseTab
from widgets import DataFrameTable, MetricRow
from workers import Worker

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


def _clean_display(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip()
    return "" if s.lower() == "nan" else s


class BarcodeTab(BaseTab):
    TITLE = "🔍 Barcode Checker"
    SUPPORTS = {"open", "save"}

    def __init__(self, settings, parent=None):
        super().__init__(settings, parent)
        self.result_df = pd.DataFrame()
        self._panel = None
        self._row_barcode = ""
        self._original = {}          # field -> value as loaded
        self._worker = None

        lay = QVBoxLayout(self)
        self.inner = QTabWidget()
        self.inner.addTab(self._build_list_page(), "📋 List check")
        self.inner.addTab(self._build_single_page(), "🔎 Single barcode")
        lay.addWidget(self.inner)

    # ------------------------------------------------------------- list page
    def _build_list_page(self) -> QWidget:
        page = QWidget()
        v = QVBoxLayout(page)
        v.addWidget(QLabel(
            "Open a file with a <b>Barcode</b> column, or paste barcodes in the "
            "control panel, to see which are already in the catalogue."
        ))
        self.metrics = MetricRow(METRICS)
        v.addWidget(self.metrics)
        self.table = DataFrameTable()
        v.addWidget(self.table, 1)
        self.note = QLabel("")
        self.note.setStyleSheet("color: #888;")
        v.addWidget(self.note)
        return page

    # ----------------------------------------------------------- single page
    def _build_single_page(self) -> QWidget:
        page = QWidget()
        v = QVBoxLayout(page)

        row = QHBoxLayout()
        row.addWidget(QLabel("Barcode:"))
        self.single_input = QLineEdit()
        self.single_input.setPlaceholderText("e.g. 8056597123456")
        self.single_input.returnPressed.connect(self.lookup_one)
        row.addWidget(self.single_input, 1)
        b_go = QPushButton("🔎 Look up")
        b_go.clicked.connect(self.lookup_one)
        row.addWidget(b_go)
        v.addLayout(row)

        self.single_status = QLabel("Enter a barcode to see every stored field.")
        self.single_status.setWordWrap(True)
        v.addWidget(self.single_status)

        self.detail = QTableWidget(0, 2)
        self.detail.setHorizontalHeaderLabels(["Field", "Value"])
        self.detail.horizontalHeader().setStretchLastSection(True)
        self.detail.setColumnWidth(0, 300)
        self.detail.setAlternatingRowColors(True)
        self.detail.setEditTriggers(QTableWidget.NoEditTriggers)
        self.detail.itemChanged.connect(self._on_cell_edited)
        v.addWidget(self.detail, 1)

        bar = QHBoxLayout()
        self.changed_label = QLabel("")
        self.changed_label.setStyleSheet("color: #888;")
        bar.addWidget(self.changed_label, 1)
        self.b_save_row = QPushButton("💾 Save changes")
        self.b_save_row.clicked.connect(self.save_action)
        self.b_save_row.setEnabled(False)
        bar.addWidget(self.b_save_row)
        v.addLayout(bar)

        self.edit_hint = QLabel("")
        self.edit_hint.setWordWrap(True)
        self.edit_hint.setStyleSheet("color: #888; font-size: 11px;")
        v.addWidget(self.edit_hint)
        return page

    # ------------------------------------------------------------------ dock
    def control_panel(self) -> QWidget:
        if self._panel is not None:
            self._refresh_updates_tree()
            return self._panel

        panel = QWidget()
        v = QVBoxLayout(panel)

        look = QGroupBox("Look up one barcode")
        lv = QVBoxLayout(look)
        self.panel_input = QLineEdit()
        self.panel_input.setPlaceholderText("Barcode…")
        self.panel_input.returnPressed.connect(self._lookup_from_panel)
        lv.addWidget(self.panel_input)
        b_look = QPushButton("🔎 Look up")
        b_look.clicked.connect(self._lookup_from_panel)
        lv.addWidget(b_look)
        v.addWidget(look)

        paste_box = QGroupBox("Check a pasted list")
        pv = QVBoxLayout(paste_box)
        self.paste_edit = QPlainTextEdit()
        self.paste_edit.setPlaceholderText("One barcode per line…")
        self.paste_edit.setMaximumHeight(140)
        pv.addWidget(self.paste_edit)
        b_check = QPushButton("📋 Check pasted list")
        b_check.clicked.connect(self._check_pasted)
        pv.addWidget(b_check)
        v.addWidget(paste_box)

        b_open = QPushButton("📂 Check a file instead…")
        b_open.clicked.connect(self.open_file)
        v.addWidget(b_open)

        b_export = QPushButton("💾 Export list result (CSV)…")
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

    def _lookup_from_panel(self) -> None:
        self.single_input.setText(self.panel_input.text())
        self.inner.setCurrentIndex(1)
        self.lookup_one()

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

    # -------------------------------------------------- single barcode lookup
    def _snapshot_row(self, key: str):
        """Find one row in the in-memory catalogue (no DB needed)."""
        if self.catalogue is None or self.catalogue.is_empty:
            return None
        master = self.catalogue.master_db
        if master.index.name == "join_key":
            if key in master.index:
                row = master.loc[key]
                return row.iloc[0] if isinstance(row, pd.DataFrame) else row
            return None
        if "join_key" in master.columns:
            hit = master[master["join_key"].astype(str).str.strip() == key]
            return None if hit.empty else hit.iloc[0]
        return None

    def lookup_one(self) -> None:
        raw = self.single_input.text().strip()
        if not raw:
            QMessageBox.information(self, "No barcode", "Type a barcode first.")
            return
        key = _clean_bc(raw)
        admin = bool(self.settings.db_url)

        row = None
        source = ""
        if admin:
            # Read live from the database so edits are based on current values,
            # not a snapshot that may be a few hours behind.
            try:
                row = admin_core.fetch_row(admin_core.get_engine(self.settings.db_url), raw)
                source = "live database"
            except Exception as e:
                self.single_status.setText(f"Database lookup failed ({e}); trying the snapshot…")
        if row is None:
            row = self._snapshot_row(key)
            source = source or "snapshot"

        self.detail.blockSignals(True)
        self.detail.setRowCount(0)
        self.detail.blockSignals(False)
        self._original = {}
        self._row_barcode = ""
        self.b_save_row.setEnabled(False)
        self.changed_label.setText("")

        if row is None:
            self.single_status.setText(
                f"❌ Barcode <b>{raw}</b> (normalised {key}) is not in the catalogue."
            )
            self.single_status.setStyleSheet(f"color: {theme.STATUS_ERROR};")
            self.edit_hint.setText("")
            self.status_message.emit(f"{raw} not found.")
            return

        self._row_barcode = raw
        fields = [(str(k), _clean_display(v)) for k, v in row.items()]
        # Filled fields first, then the empty ones, each alphabetically.
        fields.sort(key=lambda kv: (kv[1] == "", kv[0].lower()))

        self.detail.blockSignals(True)
        self.detail.setRowCount(len(fields))
        for i, (name, value) in enumerate(fields):
            f_item = QTableWidgetItem(name)
            f_item.setFlags(f_item.flags() & ~Qt.ItemIsEditable)
            self.detail.setItem(i, 0, f_item)
            v_item = QTableWidgetItem(value)
            if name == "join_key":
                v_item.setFlags(v_item.flags() & ~Qt.ItemIsEditable)
            self.detail.setItem(i, 1, v_item)
            self._original[name] = value
        self.detail.blockSignals(False)

        self.detail.setEditTriggers(
            (QTableWidget.DoubleClicked | QTableWidget.EditKeyPressed)
            if admin else QTableWidget.NoEditTriggers
        )
        name = _clean_display(row.get("Assembled_Name", "")) or "(no name)"
        filled = sum(1 for _f, v in fields if v)
        self.single_status.setText(
            f"✅ <b>{name}</b> — {filled} of {len(fields)} fields filled, "
            f"read from the {source}."
        )
        self.single_status.setStyleSheet(f"color: {theme.STATUS_READY};")
        self.edit_hint.setText(
            "Double-click a value to edit it, then 💾 Save changes (or Ctrl+S). "
            "Only the fields you actually changed get written."
            if admin else
            "Read-only. Add a database URL in ⚙️ Settings to edit fields here."
        )
        self.status_message.emit(f"Found {name} — {filled}/{len(fields)} fields filled.")

    def _changed_fields(self) -> dict:
        changes = {}
        for i in range(self.detail.rowCount()):
            f = self.detail.item(i, 0)
            v = self.detail.item(i, 1)
            if f is None or v is None:
                continue
            name = f.text()
            if name == "join_key":
                continue
            if v.text() != self._original.get(name, ""):
                changes[name] = v.text()
        return changes

    def _on_cell_edited(self, _item) -> None:
        changes = self._changed_fields()
        self.b_save_row.setEnabled(bool(changes) and bool(self.settings.db_url))
        self.changed_label.setText(f"{len(changes)} field(s) changed" if changes else "")

    def save_action(self) -> None:
        if not self._row_barcode:
            QMessageBox.information(self, "Nothing to save", "Look up a barcode first (🔎).")
            return
        changes = self._changed_fields()
        if not changes:
            QMessageBox.information(self, "Nothing to save", "No fields were changed.")
            return
        if not self.settings.db_url:
            QMessageBox.warning(
                self, "Admin not configured",
                "Editing needs a database URL — add one in ⚙️ Settings.",
            )
            return

        preview = "\n".join(
            f"  - {k}:  {self._original.get(k, '') or '(empty)'}   ->   {v or '(empty)'}"
            for k, v in sorted(changes.items())
        )
        if QMessageBox.question(
            self, "Save changes",
            f"Update {len(changes)} field(s) on barcode {self._row_barcode}?\n\n{preview}",
        ) != QMessageBox.Yes:
            return

        engine = admin_core.get_engine(self.settings.db_url)
        self.busy.emit(True)
        self._worker = Worker(
            admin_core.update_row, engine, self._row_barcode, changes, pass_progress=True
        )
        self._worker.progress.connect(lambda f, t: self.status_message.emit(t))

        def done(res):
            self.busy.emit(False)
            for k, v in changes.items():
                self._original[k] = v
            self.b_save_row.setEnabled(False)
            self.changed_label.setText("")
            msg = f"Saved {res['updated']} field(s): {', '.join(res['columns'])}."
            self.status_message.emit(msg)
            QMessageBox.information(self, "Saved", msg + self.snapshot_note())

        def failed(message):
            self.busy.emit(False)
            box = QMessageBox(QMessageBox.Critical, "Could not save", message, parent=self)
            detail = getattr(self._worker, "error_detail", "")
            if detail:
                box.setDetailedText(detail)
            box.exec()

        self._worker.done.connect(done)
        self._worker.failed.connect(failed)
        self._worker.start()

    def has_unsaved_changes(self) -> bool:
        return bool(self._row_barcode) and bool(self._changed_fields())

    # ------------------------------------------------------------- list check
    def _known_keys(self):
        if self.catalogue is None or self.catalogue.is_empty:
            return None, None
        master = self.catalogue.master_db
        if "join_key" in master.columns:
            keys = master["join_key"].astype(str).str.strip()
            names = master["Assembled_Name"] if "Assembled_Name" in master.columns else None
        elif master.index.name == "join_key":
            keys = pd.Series(master.index.astype(str).str.strip())
            names = (master["Assembled_Name"].reset_index(drop=True)
                     if "Assembled_Name" in master.columns else None)
        else:
            return None, None
        lookup = (dict(zip(keys, names.astype(str).fillna("")))
                  if names is not None else {k: "" for k in keys})
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

        self.inner.setCurrentIndex(0)
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
        self.note.setText(
            f"Showing the first {self.table.MAX_ROWS} of {len(self.result_df)} results."
            if hidden else ""
        )
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
            QMessageBox.information(self, "Nothing to export", "Run a list check first.")
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
