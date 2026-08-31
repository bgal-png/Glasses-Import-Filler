# -*- coding: utf-8 -*-
"""🎨 Colours — fill missing colours by looking at the product photos.

The desktop version is the one that really wants to exist: a scrollable grid of
real thumbnails with combo boxes, no page reloads, no ZIP upload — point it at
the folder the photos already live in.
"""
from __future__ import annotations

import os

from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import (
    QCheckBox, QComboBox, QDialog, QDialogButtonBox, QFileDialog, QFrame,
    QGridLayout, QGroupBox, QLabel, QListWidget, QMessageBox, QPlainTextEdit,
    QPushButton, QScrollArea, QVBoxLayout, QWidget,
)

import admin_core
import theme
from dictionaries import _FRAME_TEMPLE_KEYWORDS, _LENS_KEYWORDS, MANUFACTURER_CONFIG
from tabs.base import BaseTab
from widgets import MetricRow
from workers import Worker

FRAME_COLOURS = list(dict.fromkeys(v for _k, v in _FRAME_TEMPLE_KEYWORDS))
LENS_COLOURS = list(dict.fromkeys(v for _k, v in _LENS_KEYWORDS))
LABELS = {c: l for c, l, _p, _cd in admin_core.COLOUR_FIELDS}
PALETTES = {c: (FRAME_COLOURS if p == "frame" else LENS_COLOURS)
            for c, _l, p, _cd in admin_core.COLOUR_FIELDS}

METRICS = ["Need colours", "Photo matched", "Assigned"]
THUMB = 190


class PhotoCard(QFrame):
    """One product: thumbnail, identity, and a combo per missing colour."""

    def __init__(self, item: dict, on_change, parent=None):
        super().__init__(parent)
        self.item = item
        self._on_change = on_change
        self.setFrameShape(QFrame.StyledPanel)
        v = QVBoxLayout(self)

        pic = QLabel()
        pic.setAlignment(Qt.AlignCenter)
        pic.setMinimumHeight(THUMB)
        pix = QPixmap(item["photo"])
        if pix.isNull():
            pic.setText("(cannot read image)")
        else:
            pic.setPixmap(pix.scaled(THUMB, THUMB, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        v.addWidget(pic)

        title = QLabel(f"<b>{item['brand']}</b> {item['model']} {item['colour_code']}")
        title.setWordWrap(True)
        v.addWidget(title)
        sub = QLabel(f"{item['type']} · {len(item['barcodes'])} barcode(s)")
        sub.setStyleSheet("color: #888; font-size: 11px;")
        v.addWidget(sub)

        self.combos = {}
        for field in item["missing"]:
            v.addWidget(QLabel(LABELS.get(field, field)))
            combo = QComboBox()
            combo.addItem("—")
            combo.addItems(PALETTES.get(field, FRAME_COLOURS))
            combo.currentTextChanged.connect(lambda _t: self._changed())
            self.combos[field] = combo
            v.addWidget(combo)

        self.gradient = None
        if item.get("can_gradient"):
            self.gradient = QCheckBox("Gradient lens")
            self.gradient.toggled.connect(lambda _b: self._changed())
            v.addWidget(self.gradient)

        v.addStretch(1)

    def _changed(self) -> None:
        self._on_change()

    def assignment(self) -> dict:
        out = {}
        for field, combo in self.combos.items():
            value = combo.currentText()
            if value and value != "—":
                out[field] = value
        if self.gradient is not None and self.gradient.isChecked():
            out[admin_core.GRADIENT_MARKER] = True
        return out


class ColoursTab(BaseTab):
    TITLE = "🎨 Colours"
    NEEDS_ADMIN = True
    SUPPORTS = {"save"}

    PAGE_SIZE = 24

    def __init__(self, settings, parent=None):
        super().__init__(settings, parent)
        self.worklist = []
        self.matched = []
        self.unmatched = []
        self.assignments = {}
        self.page = 0
        self.photo_dir = ""
        self.photo_names = {}
        self._cards = []
        self._worker = None
        self._panel = None

        lay = QVBoxLayout(self)
        lay.addWidget(QLabel(
            "Point this at the folder with the product photos. Filenames only need "
            "to contain the model and colour code — no renaming to barcodes."
        ))
        self.metrics = MetricRow(METRICS)
        lay.addWidget(self.metrics)

        self.page_label = QLabel("")
        lay.addWidget(self.page_label)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.grid_host = QWidget()
        self.grid = QGridLayout(self.grid_host)
        self.scroll.setWidget(self.grid_host)
        lay.addWidget(self.scroll, 1)

    # ------------------------------------------------------------------ dock
    def control_panel(self) -> QWidget:
        if self._panel is not None:
            return self._panel
        panel = QWidget()
        v = QVBoxLayout(panel)

        box = QGroupBox("1 · Photos")
        bv = QVBoxLayout(box)
        self.dir_label = QLabel("No folder selected.")
        self.dir_label.setWordWrap(True)
        bv.addWidget(self.dir_label)
        b_dir = QPushButton("📂 Select photo folder…")
        b_dir.clicked.connect(self._pick_folder)
        bv.addWidget(b_dir)
        v.addWidget(box)

        scope_box = QGroupBox("2 · Limit to producers (optional)")
        sv = QVBoxLayout(scope_box)
        self.scope = QListWidget()
        self.scope.setSelectionMode(QListWidget.MultiSelection)
        for name in sorted(MANUFACTURER_CONFIG.keys()):
            self.scope.addItem(name.title())
        self.scope.setMaximumHeight(120)
        sv.addWidget(self.scope)
        sv.addWidget(QLabel("Nothing selected = all producers."))
        v.addWidget(scope_box)

        b_build = QPushButton("🔍 Build worklist")
        b_build.clicked.connect(self._build)
        v.addWidget(b_build)

        nav = QGroupBox("3 · Review")
        nv = QVBoxLayout(nav)
        b_prev = QPushButton("⬅️ Previous page")
        b_prev.clicked.connect(lambda: self._go(-1))
        nv.addWidget(b_prev)
        b_next = QPushButton("➡️ Next page")
        b_next.clicked.connect(lambda: self._go(1))
        nv.addWidget(b_next)
        self.unmatched_label = QLabel("")
        self.unmatched_label.setWordWrap(True)
        self.unmatched_label.setStyleSheet("color: #888; font-size: 11px;")
        nv.addWidget(self.unmatched_label)
        b_why = QPushButton("🔎 Why didn't these match?")
        b_why.clicked.connect(self._show_unmatched)
        nv.addWidget(b_why)
        v.addWidget(nav)

        b_save = QPushButton("💾 Save colours to database")
        b_save.clicked.connect(self.save_action)
        v.addWidget(b_save)

        v.addStretch(1)
        self._panel = panel
        return panel

    def _pick_folder(self) -> None:
        # No ShowDirsOnly — Windows hides files otherwise and folders look empty.
        folder = QFileDialog.getExistingDirectory(
            self, "Select the folder with product photos",
            self.settings.last_dir("colour_photos"), QFileDialog.Option(0),
        )
        if not folder:
            return
        self.settings.set_last_dir("colour_photos", folder)
        self.photo_dir = folder
        names = self._photo_names()
        if names:
            self.dir_label.setText(f"📸 {len(names)} photo(s) in {os.path.basename(folder)}")
        else:
            try:
                entries = os.listdir(folder)
            except Exception:
                entries = []
            exts = {}
            for e in entries:
                ext = os.path.splitext(e)[1].lower() or "(no extension)"
                exts[ext] = exts.get(ext, 0) + 1
            found = ", ".join(f"{k} ×{v}" for k, v in sorted(exts.items(), key=lambda x: -x[1])[:6])
            self.dir_label.setText(
                f"No supported images among {len(entries)} file(s)."
                + (f" Found: {found}" if found else "")
                + " Supported: .jpg .jpeg .png .webp"
            )

    def _photo_names(self) -> dict:
        """{full path: basename without extension}"""
        out = {}
        if not self.photo_dir or not os.path.isdir(self.photo_dir):
            return out
        for entry in os.listdir(self.photo_dir):
            path = os.path.join(self.photo_dir, entry)
            if os.path.isfile(path) and entry.lower().endswith(
                (".jpg", ".jpeg", ".png", ".webp")
            ):
                out[path] = os.path.splitext(entry)[0]
        return out

    # ----------------------------------------------------------------- build
    def _build(self) -> None:
        if self.catalogue is None or self.catalogue.is_empty:
            QMessageBox.warning(self, "No catalogue data", "Use ☁️ Refresh data first.")
            return
        photos = self._photo_names()
        if not photos:
            QMessageBox.information(self, "No photos", "Select a folder with product photos first.")
            return

        self.photo_names = photos
        producers = [self.scope.item(i).text().lower()
                     for i in range(self.scope.count())
                     if self.scope.item(i).isSelected()] or None
        master = self.catalogue.master_db

        def job(progress=None):
            work = admin_core.build_colour_worklist(master, producers)
            matched, unmatched = admin_core.match_photos(work, photos)
            return work, matched, unmatched

        self.busy.emit(True)
        self.status_message.emit("Scanning the catalogue for missing colours…")
        self._worker = Worker(job, pass_progress=True)
        self._worker.done.connect(self._on_built)
        self._worker.failed.connect(self._on_failed)
        self._worker.start()

    def _on_built(self, result) -> None:
        self.busy.emit(False)
        self.worklist, self.matched, self.unmatched = result
        self.assignments = {}
        self.page = 0

        self.metrics.set("Need colours", len(self.worklist))
        self.metrics.set(
            "Photo matched", len(self.matched),
            theme.STATUS_READY if self.matched else theme.STATUS_ERROR,
        )
        self.metrics.set("Assigned", 0)
        self.unmatched_label.setText(
            f"{len(self.unmatched)} group(s) had no matching photo — their filenames "
            "probably don't contain the model + colour code."
            if self.unmatched else "Every group matched a photo."
        )
        self._render_page()
        self.status_message.emit(
            f"{len(self.worklist)} group(s) need colours; matched a photo for {len(self.matched)}."
        )

    def _on_failed(self, message: str) -> None:
        self.busy.emit(False)
        box = QMessageBox(QMessageBox.Critical, "Failed", message, parent=self)
        detail = getattr(self._worker, "error_detail", "")
        if detail:
            box.setDetailedText(detail)
        box.exec()

    # ------------------------------------------------------------------ grid
    def _render_page(self) -> None:
        while self.grid.count():
            item = self.grid.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()
        self._cards = []

        if not self.matched:
            self.page_label.setText("Nothing to review yet — build a worklist.")
            return

        pages = max(1, (len(self.matched) + self.PAGE_SIZE - 1) // self.PAGE_SIZE)
        self.page = max(0, min(self.page, pages - 1))
        start = self.page * self.PAGE_SIZE
        chunk = self.matched[start:start + self.PAGE_SIZE]
        self.page_label.setText(
            f"Page {self.page + 1} of {pages} — items {start + 1}–{start + len(chunk)} "
            f"of {len(self.matched)}. Choices are kept while you page around; "
            f"press 💾 when you're done."
        )

        cols = 4
        for i, item in enumerate(chunk):
            card = PhotoCard(item, self._collect, self)
            saved = self.assignments.get(item["key"], {})
            for field, combo in card.combos.items():
                if field in saved:
                    combo.setCurrentText(saved[field])
            if card.gradient is not None and saved.get(admin_core.GRADIENT_MARKER):
                card.gradient.setChecked(True)
            self.grid.addWidget(card, i // cols, i % cols)
            self._cards.append(card)

    def _show_unmatched(self) -> None:
        """Explain the misses instead of just counting them: show what we looked
        for next to what the filenames actually are, so the pattern is obvious."""
        if not self.worklist:
            QMessageBox.information(self, "Nothing to explain", "Build a worklist first.")
            return
        if not self.unmatched:
            QMessageBox.information(
                self, "Everything matched",
                f"All {len(self.matched)} group(s) found a photo.",
            )
            return

        lines = [
            f"{len(self.unmatched)} of {len(self.worklist)} group(s) found no photo, "
            f"out of {len(self.photo_names)} image(s) in the folder.",
            "",
            "A filename matches when it contains the model code AND either the "
            "colour code or size+colour (ignoring spaces, dashes, underscores "
            "and letter case).",
            "",
            "LOOKED FOR (first 25 unmatched groups)",
            f"{'model':<20} {'colour':<10} {'or size+colour':<16} brand",
        ]
        for g in self.unmatched[:25]:
            lines.append(
                f"{g['_model_n']:<20} {g['_colour_n']:<10} "
                f"{admin_core.norm_key(g.get('size', '')) + g['_colour_n']:<16} {g['brand']}"
            )
        if len(self.unmatched) > 25:
            lines.append(f"… and {len(self.unmatched) - 25} more")

        lines += ["", "ACTUAL FILENAMES (first 25 in the folder)"]
        for base in sorted(self.photo_names.values())[:25]:
            lines.append(f"{base}    ->  normalised: {admin_core.norm_key(base)}")
        if len(self.photo_names) > 25:
            lines.append(f"… and {len(self.photo_names) - 25} more")

        dlg = QDialog(self)
        dlg.setWindowTitle("Why photos didn't match")
        dlg.resize(920, 620)
        v = QVBoxLayout(dlg)
        box = QPlainTextEdit(chr(10).join(lines))
        box.setReadOnly(True)
        f = box.font()
        f.setFamily("Consolas")
        box.setFont(f)
        v.addWidget(box)
        v.addWidget(QLabel(
            "Copy this and send it over — it's enough to widen the matcher."
        ))
        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.rejected.connect(dlg.reject)
        v.addWidget(buttons)
        dlg.exec()

    def _collect(self) -> None:
        for card in self._cards:
            values = card.assignment()
            if values:
                self.assignments[card.item["key"]] = values
            else:
                self.assignments.pop(card.item["key"], None)
        self.metrics.set(
            "Assigned", len(self.assignments),
            theme.STATUS_READY if self.assignments else None,
        )

    def _go(self, delta: int) -> None:
        self._collect()
        self.page += delta
        self._render_page()

    # ------------------------------------------------------------------ save
    def save_action(self) -> None:
        self._collect()
        if not self.assignments:
            QMessageBox.information(self, "Nothing to save", "Assign some colours first.")
            return
        cells = sum(
            len(self.assignments[g["key"]]) * len(g["barcodes"])
            for g in self.matched if g["key"] in self.assignments
        )
        if QMessageBox.question(
            self, "Save colours",
            f"Write {len(self.assignments)} group(s) to the database?\n\n"
            f"That fills about {cells} cell(s) across every barcode sharing each "
            f"model + colour.",
        ) != QMessageBox.Yes:
            return

        try:
            engine = admin_core.get_engine(self.settings.db_url)
        except ValueError as e:
            QMessageBox.warning(self, "Admin not configured", str(e))
            return

        barcodes_by_group = {g["key"]: g["barcodes"] for g in self.matched}
        assignments = dict(self.assignments)

        self.busy.emit(True)
        self._worker = Worker(
            admin_core.save_colours, engine, assignments, barcodes_by_group,
            pass_progress=True,
        )
        self._worker.progress.connect(lambda f, t: self.status_message.emit(t))

        def done(res):
            self.busy.emit(False)
            msg = f"Saved {res['groups']} group(s), {res['cells']} cell(s)."
            self.assignments = {}
            self.metrics.set("Assigned", 0)
            self._render_page()
            self.status_message.emit(msg + " Use ☁️ Refresh data to see it in the filler.")
            QMessageBox.information(self, "Saved", msg)

        self._worker.done.connect(done)
        self._worker.failed.connect(self._on_failed)
        self._worker.start()

    def has_unsaved_changes(self) -> bool:
        return bool(self.assignments)
