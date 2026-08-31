# -*- coding: utf-8 -*-
"""Tabs package. The order here is the order they appear in the UI.

Tabs with NEEDS_ADMIN = True are only added when a DB_URL is configured, so a
single build serves colleagues (read-only) and the admin (everything).
"""

from tabs.base import BaseTab
from tabs.filler_tab import FillerTab
from tabs.barcode_tab import BarcodeTab
from tabs.catalogue_tab import CatalogueTab
from tabs.rename_tab import RenameTab
from tabs.registry_tab import RegistryTab

ALL_TABS = [
    FillerTab,      # 🪄 read-only
    BarcodeTab,     # 🔍 read-only
    CatalogueTab,   # 🏭 admin
    RenameTab,      # ✏️ admin
    RegistryTab,    # 📒 admin
]

__all__ = [
    "BaseTab", "FillerTab", "BarcodeTab", "CatalogueTab",
    "RenameTab", "RegistryTab", "ALL_TABS",
]
