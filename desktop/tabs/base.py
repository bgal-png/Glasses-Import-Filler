# -*- coding: utf-8 -*-
"""Common contract for every tab.

Tabs are plain QWidgets. The main window owns the single toolbar and the right
dock; it asks the current tab what it supports and routes actions to it, and
swaps the dock's contents (hiding the dock entirely on tabs that don't use it).
"""
from __future__ import annotations

from PySide6.QtCore import Signal
from PySide6.QtWidgets import QWidget

NEWLINES = chr(10) * 2


class BaseTab(QWidget):
    TITLE = "Tab"
    NEEDS_ADMIN = False          # True = hidden unless a DB_URL is configured
    SUPPORTS: set = set()        # any of {"open", "run", "save"}

    # → status bar / progress bar
    status_message = Signal(str)
    busy = Signal(bool)

    def __init__(self, settings, parent=None):
        super().__init__(parent)
        self.settings = settings
        self.catalogue = None

    # --- catalogue plumbing ---
    def set_catalogue(self, data) -> None:
        self.catalogue = data
        self.on_catalogue(data)

    def on_catalogue(self, data) -> None:
        """Called whenever fresh catalogue data arrives. Override as needed."""

    # --- right dock ---
    def control_panel(self) -> QWidget | None:
        """Return the widget for the right-hand 'Control panel' dock, or None
        to hide the dock on this tab."""
        return None

    # --- toolbar actions (only called if listed in SUPPORTS) ---
    def open_file(self) -> None:
        pass

    def run_action(self) -> None:
        pass

    def save_action(self) -> None:
        pass

    # --- snapshot staleness ---
    def snapshot_note(self) -> str:
        """A DB write is invisible to the filler until the snapshot is
        republished, so say so. Empty when reading the database directly."""
        if getattr(self.settings, "snapshot_repo", ""):
            return (
                NEWLINES + "Heads up: the filler reads the published snapshot, not the "
                "database. Run the “Publish catalogue snapshot” Action on "
                "GitHub (or wait for the daily 06:45 UTC run) for this change to "
                "reach the filler and your colleagues."
            )
        return ""

    # --- unsaved-changes tracking ---
    def has_unsaved_changes(self) -> bool:
        return False
