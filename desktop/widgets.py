# -*- coding: utf-8 -*-
"""Small reusable widgets shared by the tabs."""
from __future__ import annotations

import pandas as pd
from PySide6.QtCore import Qt
from PySide6.QtGui import QAction, QColor, QGuiApplication
from PySide6.QtWidgets import (
    QGroupBox, QHBoxLayout, QLabel, QMenu, QTableWidget, QTableWidgetItem,
    QVBoxLayout, QWidget,
)

import theme


class MetricCard(QGroupBox):
    """A summary count: small caption, big bold number."""

    def __init__(self, title: str, value: str = "—", parent=None):
        super().__init__(title, parent)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(10, 6, 10, 8)
        self._label = QLabel(value)
        f = self._label.font()
        f.setPointSize(max(16, f.pointSize() + 8))
        f.setBold(True)
        self._label.setFont(f)
        self._label.setAlignment(Qt.AlignCenter)
        lay.addWidget(self._label)

    def set_value(self, value) -> None:
        self._label.setText(str(value))

    def set_colour(self, colour: str | None) -> None:
        self._label.setStyleSheet(f"color: {colour};" if colour else "")


class MetricRow(QWidget):
    """A row of MetricCards, addressed by key."""

    def __init__(self, titles: list, parent=None):
        super().__init__(parent)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        self.cards = {}
        for t in titles:
            card = MetricCard(t)
            self.cards[t] = card
            lay.addWidget(card)

    def set(self, title: str, value, colour: str | None = None) -> None:
        if title in self.cards:
            self.cards[title].set_value(value)
            self.cards[title].set_colour(colour)


class DataFrameTable(QTableWidget):
    """Read-only DataFrame view.

    Qt header labels are not selectable text, which makes column names
    impossible to copy — so clicking a header copies it, and the right-click
    menu offers copy-one / copy-all-headers / copy-whole-matrix-as-TSV (paste
    straight into Excel).
    """

    MAX_ROWS = 500  # keep the UI responsive; the file itself is unaffected

    def __init__(self, parent=None):
        super().__init__(parent)
        self._df = pd.DataFrame()
        self.setEditTriggers(QTableWidget.NoEditTriggers)
        self.setSelectionBehavior(QTableWidget.SelectItems)
        self.setAlternatingRowColors(True)
        self.setSortingEnabled(False)
        self.horizontalHeader().setStretchLastSection(True)
        self.horizontalHeader().setSectionsClickable(True)
        self.horizontalHeader().sectionClicked.connect(self._copy_header)
        self.horizontalHeader().setContextMenuPolicy(Qt.CustomContextMenu)
        self.horizontalHeader().customContextMenuRequested.connect(self._header_menu)
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self._cell_menu)

    # --- data ---
    def set_dataframe(self, df: pd.DataFrame, highlight: dict | None = None) -> None:
        """highlight: {(row_pos, col_name): 'error'|'warning'|'ok'}"""
        self._df = pd.DataFrame() if df is None else df
        highlight = highlight or {}
        shown = self._df.head(self.MAX_ROWS)

        self.clear()
        self.setRowCount(len(shown))
        self.setColumnCount(len(shown.columns))
        self.setHorizontalHeaderLabels([str(c) for c in shown.columns])

        colours = {
            "error": QColor(theme.COLOR_ERROR),
            "warning": QColor(theme.COLOR_WARNING),
            "ok": QColor(theme.COLOR_OK),
        }
        forced_text = QColor(theme.COLOR_FORCED_TEXT)

        for r in range(len(shown)):
            for c, col in enumerate(shown.columns):
                val = shown.iloc[r, c]
                text = "" if (val is None or (isinstance(val, float) and pd.isna(val))) else str(val)
                if text == "nan":
                    text = ""
                item = QTableWidgetItem(text)
                key = (r, str(col))
                if key in highlight and highlight[key] in colours:
                    item.setBackground(colours[highlight[key]])
                    # Force dark text so highlights stay readable in dark mode
                    item.setForeground(forced_text)
                self.setItem(r, c, item)

        self.resizeColumnsToContents()
        for c in range(self.columnCount()):
            if self.columnWidth(c) > 320:
                self.setColumnWidth(c, 320)

    def truncated(self) -> int:
        """How many rows were not displayed."""
        return max(0, len(self._df) - self.MAX_ROWS)

    # --- clipboard helpers ---
    @staticmethod
    def _to_clipboard(text: str) -> None:
        QGuiApplication.clipboard().setText(text)

    def _copy_header(self, index: int) -> None:
        if 0 <= index < self.columnCount():
            self._to_clipboard(self.horizontalHeaderItem(index).text())

    def _header_menu(self, pos) -> None:
        index = self.horizontalHeader().logicalIndexAt(pos)
        menu = QMenu(self)
        if index >= 0:
            a = QAction(f"Copy header “{self.horizontalHeaderItem(index).text()}”", self)
            a.triggered.connect(lambda: self._copy_header(index))
            menu.addAction(a)
        a_all = QAction("Copy all headers", self)
        a_all.triggered.connect(lambda: self._to_clipboard(
            "\t".join(self.horizontalHeaderItem(c).text() for c in range(self.columnCount()))
        ))
        menu.addAction(a_all)
        a_matrix = QAction("Copy whole table as TSV (for Excel)", self)
        a_matrix.triggered.connect(self._copy_matrix)
        menu.addAction(a_matrix)
        menu.exec(self.horizontalHeader().mapToGlobal(pos))

    def _cell_menu(self, pos) -> None:
        menu = QMenu(self)
        a_sel = QAction("Copy selection", self)
        a_sel.triggered.connect(self._copy_selection)
        menu.addAction(a_sel)
        a_matrix = QAction("Copy whole table as TSV (for Excel)", self)
        a_matrix.triggered.connect(self._copy_matrix)
        menu.addAction(a_matrix)
        menu.exec(self.viewport().mapToGlobal(pos))

    def _copy_selection(self) -> None:
        ranges = self.selectedRanges()
        if not ranges:
            return
        r0 = ranges[0]
        lines = []
        for r in range(r0.topRow(), r0.bottomRow() + 1):
            cells = []
            for c in range(r0.leftColumn(), r0.rightColumn() + 1):
                item = self.item(r, c)
                cells.append(item.text() if item else "")
            lines.append("\t".join(cells))
        self._to_clipboard("\n".join(lines))

    def _copy_matrix(self) -> None:
        if self._df.empty:
            return
        self._to_clipboard(self._df.to_csv(sep="\t", index=False))
