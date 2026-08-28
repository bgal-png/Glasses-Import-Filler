# -*- coding: utf-8 -*-
"""Tabs package. Import order here defines the order they appear in the UI."""

from tabs.base import BaseTab
from tabs.filler_tab import FillerTab
from tabs.barcode_tab import BarcodeTab

# Admin tabs are appended in a later phase; they self-declare NEEDS_ADMIN so the
# main window can hide them when no DB_URL is configured.
ALL_TABS = [FillerTab, BarcodeTab]

__all__ = ["BaseTab", "FillerTab", "BarcodeTab", "ALL_TABS"]
