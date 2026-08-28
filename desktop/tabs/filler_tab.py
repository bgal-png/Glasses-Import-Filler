# -*- coding: utf-8 -*-
"""🪄 Auto-Filler — open a target import file, fill it from the catalogue, save."""
from __future__ import annotations

import os

import pandas as pd
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox, QFileDialog, QFormLayout, QGroupBox, QHBoxLayout, QLabel,
    QLineEdit, QMessageBox, QPushButton, QSplitter, QTreeWidget,
    QTreeWidgetItem, QVBoxLayout, QWidget,
)

import theme
from filler_core import (
    FillOptions, changed_columns, fill_target, load_images_from_folder,
    read_target_file, run_ai_vision, target_barcode_column,
    write_filled_excel, write_onto_source_workbook,
)
from tabs.base import BaseTab
from widgets import DataFrameTable, MetricRow
from workers import Worker

METRICS = ["Rows", "Matched", "Not matched", "Issues"]


class FillerTab(BaseTab):
    TITLE = "🪄 Auto-Filler"
    SUPPORTS = {"open", "run", "save"}

    def __init__(self, settings, parent=None):
        super().__init__(settings, parent)
        self.source_path = ""
        self.original_df = pd.DataFrame()
        self.filled_df = pd.DataFrame()
        self.report = None
        self.image_folder = ""
        self._image_files: list = []
        self._dirty = False
        self._worker = None
        self._panel = None

        lay = QVBoxLayout(self)

        self.file_label = QLabel("No file opened. Use 📂 Open target file.")
        self.file_label.setWordWrap(True)
        lay.addWidget(self.file_label)

        self.metrics = MetricRow(METRICS)
        lay.addWidget(self.metrics)

        self.only_changed = QCheckBox("Show only the columns the filler changed")
        self.only_changed.setChecked(True)
        self.only_changed.toggled.connect(self._refresh_table)
        lay.addWidget(self.only_changed)

        splitter = QSplitter(Qt.Vertical)
        self.table = DataFrameTable()
        splitter.addWidget(self.table)

        self.report_tree = QTreeWidget()
        self.report_tree.setHeaderLabels(["Validation report", "Detail"])
        self.report_tree.setColumnWidth(0, 420)
        splitter.addWidget(self.report_tree)
        splitter.setSizes([520, 220])
        lay.addWidget(splitter, 1)

        self.truncation_note = QLabel("")
        self.truncation_note.setStyleSheet("color: #888;")
        lay.addWidget(self.truncation_note)

    # ------------------------------------------------------------------ dock
    def control_panel(self) -> QWidget:
        if self._panel is not None:
            return self._panel

        panel = QWidget()
        v = QVBoxLayout(panel)

        priv_box = QGroupBox("🏷️ Private name numbers")
        form = QFormLayout(priv_box)
        self.priv_inputs = {}
        for key, label, placeholder in (
            ("priv_sun", "Sunglasses", "e.g. 1001"),
            ("priv_eye", "Eyeglasses (Frames)", "e.g. 2001"),
            ("priv_pc", "PC Glasses", "e.g. 3001"),
            ("priv_sport", "Sport Glasses", "e.g. 4001"),
            ("priv_drive", "Driving Glasses", "e.g. 5001"),
        ):
            edit = QLineEdit()
            edit.setPlaceholderText(placeholder)
            edit.setText(str(self.settings.get(f"filler/{key}", "") or ""))
            edit.textChanged.connect(lambda t, k=key: self.settings.set(f"filler/{k}", t))
            self.priv_inputs[key] = edit
            form.addRow(label, edit)
        v.addWidget(priv_box)

        img_box = QGroupBox("👓 Shape recognition (optional)")
        iv = QVBoxLayout(img_box)
        self.image_label = QLabel("No images selected.")
        self.image_label.setWordWrap(True)
        iv.addWidget(self.image_label)
        row = QHBoxLayout()
        b_folder = QPushButton("Select folder…")
        b_folder.clicked.connect(self._pick_image_folder)
        b_files = QPushButton("Select files…")
        b_files.clicked.connect(self._pick_image_files)
        row.addWidget(b_folder)
        row.addWidget(b_files)
        iv.addLayout(row)
        b_clear = QPushButton("Clear images")
        b_clear.clicked.connect(self._clear_images)
        iv.addWidget(b_clear)
        self.ai_hint = QLabel("")
        self.ai_hint.setWordWrap(True)
        self.ai_hint.setStyleSheet("color: #888; font-size: 11px;")
        iv.addWidget(self.ai_hint)
        v.addWidget(img_box)

        run_btn = QPushButton("🪄 Run auto-filler")
        run_btn.clicked.connect(self.run_action)
        v.addWidget(run_btn)

        save_btn = QPushButton("💾 Save filled file…")
        save_btn.clicked.connect(self.save_action)
        v.addWidget(save_btn)

        v.addStretch(1)
        self._panel = panel
        self._update_ai_hint()
        return panel

    def _update_ai_hint(self) -> None:
        if not hasattr(self, "ai_hint"):
            return
        if self.settings.anthropic_key:
            self.ai_hint.setText(
                "Filenames must match the “Glasses name” column exactly. "
                "Only rows whose name has a matching image are classified."
            )
        else:
            self.ai_hint.setText(
                "Add an Anthropic API key in ⚙️ Settings to enable AI shape recognition."
            )

    # ---------------------------------------------------------------- images
    def _pick_image_folder(self) -> None:
        # NOTE: no ShowDirsOnly — Windows otherwise hides the files and a full
        # folder looks empty.
        folder = QFileDialog.getExistingDirectory(
            self, "Select the folder with product images",
            self.settings.last_dir("images"), QFileDialog.Option(0),
        )
        if not folder:
            return
        self.settings.set_last_dir("images", folder)
        self.image_folder = folder
        self._image_files = []
        images = load_images_from_folder(folder)
        if images:
            self.image_label.setText(f"📸 {len(images)} image(s) in {os.path.basename(folder)}")
        else:
            # "Empty results must explain themselves"
            try:
                entries = os.listdir(folder)
            except Exception:
                entries = []
            exts = {}
            for e in entries:
                ext = os.path.splitext(e)[1].lower() or "(no extension)"
                exts[ext] = exts.get(ext, 0) + 1
            found = ", ".join(f"{k} ×{v}" for k, v in sorted(exts.items(), key=lambda x: -x[1])[:6])
            self.image_label.setText(
                f"No supported images among {len(entries)} file(s)."
                + (f" Found: {found}" if found else "")
                + " Supported: .jpg .jpeg .png"
            )

    def _pick_image_files(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select product images", self.settings.last_dir("images"),
            "Images (*.jpg *.jpeg *.png)",
        )
        if not files:
            return
        self.settings.set_last_dir("images", os.path.dirname(files[0]))
        self.image_folder = ""
        self._image_files = files
        self.image_label.setText(f"📸 {len(files)} image file(s) selected")

    def _clear_images(self) -> None:
        self.image_folder = ""
        self._image_files = []
        self.image_label.setText("No images selected.")

    def _collect_images(self) -> dict:
        if self._image_files:
            out = {}
            for path in self._image_files:
                try:
                    with open(path, "rb") as fh:
                        out[os.path.splitext(os.path.basename(path))[0]] = fh.read()
                except Exception:
                    pass
            return out
        if self.image_folder:
            return load_images_from_folder(self.image_folder)
        return {}

    # ------------------------------------------------------------------ open
    def open_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Open the target import file", self.settings.last_dir("target"),
            "Import files (*.xlsx *.csv);;Excel (*.xlsx);;CSV (*.csv)",
        )
        if not path:
            return
        self.settings.set_last_dir("target", os.path.dirname(path))
        try:
            df = read_target_file(path)
        except Exception as e:
            QMessageBox.critical(self, "Could not open file", str(e))
            return

        if target_barcode_column(df) is None:
            QMessageBox.critical(
                self, "Missing barcode column",
                "This file has no “Barcode” column, so nothing can be matched.\n\n"
                f"Columns found: {', '.join(map(str, df.columns[:12]))}"
                + (" …" if len(df.columns) > 12 else ""),
            )
            return

        self.source_path = path
        self.original_df = df
        self.filled_df = df.copy()
        self.report = None
        self._dirty = False
        self.file_label.setText(f"📂 {path}  —  {len(df)} rows × {len(df.columns)} columns")
        self.metrics.set("Rows", len(df))
        for m in ("Matched", "Not matched", "Issues"):
            self.metrics.set(m, "—", None)
        self.report_tree.clear()
        self._refresh_table()
        self.status_message.emit(f"Opened {os.path.basename(path)} ({len(df)} rows).")

    # ------------------------------------------------------------------- run
    def run_action(self) -> None:
        if self.original_df.empty:
            QMessageBox.information(self, "Nothing to fill", "Open a target file first (📂).")
            return
        if self.catalogue is None or self.catalogue.is_empty:
            QMessageBox.warning(
                self, "No catalogue data",
                "The catalogue hasn't loaded yet. Use ☁️ Refresh data, or check ⚙️ Settings.",
            )
            return

        options = FillOptions(**{
            k: (self.priv_inputs[k].text().strip() if hasattr(self, "priv_inputs") else "")
            for k in ("priv_sun", "priv_eye", "priv_pc", "priv_sport", "priv_drive")
        })
        images = self._collect_images()
        api_key = self.settings.anthropic_key

        base_df = self.original_df.copy()
        cat = self.catalogue

        def job(progress=None):
            filled, report = fill_target(
                base_df, cat.master_db, cat.package_df, cat.origin_df,
                options=options, progress=progress,
            )
            ai = None
            if images and api_key:
                ai = run_ai_vision(filled, images, api_key, progress=progress)
            return filled, report, ai

        self.busy.emit(True)
        self.status_message.emit("Filling…")
        self._worker = Worker(job, pass_progress=True)
        self._worker.progress.connect(lambda f, t: self.status_message.emit(t))
        self._worker.done.connect(self._on_filled)
        self._worker.failed.connect(self._on_failed)
        self._worker.start()

    def _on_filled(self, result) -> None:
        filled, report, ai = result
        self.busy.emit(False)
        self.filled_df = filled
        self.report = report
        self._dirty = True

        self.metrics.set("Rows", report.total_rows)
        self.metrics.set("Matched", report.match_count, theme.STATUS_READY)
        self.metrics.set(
            "Not matched", report.unmatched_count,
            theme.STATUS_ERROR if report.unmatched_count else None,
        )
        self.metrics.set(
            "Issues", report.total_issues,
            theme.STATUS_LOADING if report.total_issues else None,
        )

        self._build_report_tree(report, ai)
        self._refresh_table()

        msg = f"Filled {report.match_count} of {report.total_rows} products."
        if ai:
            msg += f" AI shapes: {ai.shape_count}, sport: {ai.sport_count}."
        self.status_message.emit(msg)

    def _on_failed(self, message: str) -> None:
        self.busy.emit(False)
        self.status_message.emit("Fill failed.")
        detail = getattr(self._worker, "error_detail", "")
        box = QMessageBox(QMessageBox.Critical, "Fill failed", message, parent=self)
        if detail:
            box.setDetailedText(detail)
        box.exec()

    def _build_report_tree(self, report, ai) -> None:
        self.report_tree.clear()

        if report.found_sport_glasses:
            item = QTreeWidgetItem(["⚠️ Sport glasses labelled as “Ski goggles”",
                                    "Check the Meta description on those rows."])
            item.setForeground(0, Qt.GlobalColor.darkYellow)
            self.report_tree.addTopLevelItem(item)

        if report.found_polarized_clip_on:
            self.report_tree.addTopLevelItem(QTreeWidgetItem([
                "⚠️ Polarized clip-on alert",
                "A polarized clip-on got a standard clip-on value — verify the ' p' suffix.",
            ]))

        if report.unmapped:
            root = QTreeWidgetItem([f"🔴 Unmapped values ({len(report.unmapped)} columns)", ""])
            for col, vals in sorted(report.unmapped.items()):
                node = QTreeWidgetItem([col, f"{len(vals)} unmapped value(s)"])
                for v in sorted(vals):
                    node.addChild(QTreeWidgetItem(["", str(v)]))
                root.addChild(node)
            self.report_tree.addTopLevelItem(root)

        if report.missing:
            root = QTreeWidgetItem([f"🟡 Missing from source ({len(report.missing)} columns)", ""])
            for col, count in sorted(report.missing.items(), key=lambda x: -x[1]):
                root.addChild(QTreeWidgetItem([col, f"{count} row(s) with no data"]))
            self.report_tree.addTopLevelItem(root)

        if ai:
            self.report_tree.addTopLevelItem(QTreeWidgetItem([
                "👓 AI vision",
                f"{ai.shape_count} shape(s), {ai.sport_count} sport flag(s) "
                f"from {ai.image_count} image(s)",
            ]))

        if self.report_tree.topLevelItemCount() == 0:
            self.report_tree.addTopLevelItem(QTreeWidgetItem(["✅ No issues reported", ""]))
        self.report_tree.expandToDepth(0)

    # ----------------------------------------------------------------- table
    def _refresh_table(self) -> None:
        df = self.filled_df
        if df is None or df.empty:
            self.table.set_dataframe(pd.DataFrame())
            self.truncation_note.setText("")
            return

        if self.only_changed.isChecked() and not self.original_df.empty:
            cols = changed_columns(self.original_df, df)
            if cols:
                id_col = "Glasses name" if "Glasses name" in df.columns else df.columns[0]
                ordered = [id_col] + [c for c in cols if c != id_col]
                df = df[[c for c in ordered if c in df.columns]]

        self.table.set_dataframe(df)
        hidden = self.table.truncated()
        self.truncation_note.setText(
            f"Showing the first {self.table.MAX_ROWS} rows — {hidden} more are in the file "
            f"and will be saved." if hidden else ""
        )

    # ------------------------------------------------------------------ save
    def save_action(self) -> None:
        if self.filled_df is None or self.filled_df.empty:
            QMessageBox.information(self, "Nothing to save", "Fill a file first (🪄).")
            return

        default = ""
        if self.source_path:
            base, ext = os.path.splitext(os.path.basename(self.source_path))
            default = os.path.join(
                self.settings.last_dir("output") or os.path.dirname(self.source_path),
                f"{base} filled.xlsx",
            )
        path, _ = QFileDialog.getSaveFileName(
            self, "Save the filled file", default, "Excel (*.xlsx)"
        )
        if not path:
            return
        self.settings.set_last_dir("output", os.path.dirname(path))

        try:
            if self.source_path.lower().endswith(".xlsx"):
                # Fill a copy of the user's own workbook so the header row keeps
                # its colours/fonts and barcodes stay text-formatted.
                info = write_onto_source_workbook(self.source_path, self.filled_df, path)
                extra = (f" Added {len(info['new_columns'])} new column(s)."
                         if info["new_columns"] else "")
                note = f"Saved onto a copy of your workbook.{extra}"
            else:
                write_filled_excel(self.filled_df, path)
                note = "Saved as a new workbook (source was a CSV)."
        except Exception as e:
            QMessageBox.critical(self, "Could not save", str(e))
            return

        self._dirty = False
        self.status_message.emit(f"{note} → {path}")
        QMessageBox.information(self, "Saved", f"{note}\n\n{path}")

    def has_unsaved_changes(self) -> bool:
        return self._dirty

    def on_catalogue(self, data) -> None:
        self._update_ai_hint()
