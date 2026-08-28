# -*- coding: utf-8 -*-
"""Persisted user settings (QSettings → Windows registry, per user).

IMPORTANT — credentials policy:
  * `snapshot_repo` / `snapshot_token` are read-only access to the catalogue
    snapshot. These MAY be shipped as build-time defaults (see defaults.py),
    because the worst case is "can read the product snapshot".
  * `db_url` is NEVER shipped. It grants read/write on the live database, so it
    lives only in the settings of whoever pastes it in. Admin tabs stay hidden
    until it is present — that is the gate that lets one build serve both
    colleagues (filler + barcode checker) and the admin (everything).
  * `anthropic_key` is per-user too; AI shape recognition is off without it.
"""
from __future__ import annotations

from PySide6.QtCore import QSettings

from version import APP_NAME, ORG_NAME

try:  # optional build-time defaults, not committed with real values
    import defaults  # type: ignore
except Exception:  # pragma: no cover
    defaults = None


def _default(name: str, fallback: str = "") -> str:
    return str(getattr(defaults, name, fallback) or "") if defaults else fallback


class Settings:
    def __init__(self) -> None:
        self._s = QSettings(ORG_NAME, APP_NAME)

    # --- generic ---
    def get(self, key: str, fallback=None):
        return self._s.value(key, fallback)

    def set(self, key: str, value) -> None:
        self._s.setValue(key, value)

    # --- appearance ---
    @property
    def dark_mode(self) -> bool:
        return str(self._s.value("dark_mode", "false")).lower() in ("1", "true", "yes")

    @dark_mode.setter
    def dark_mode(self, value: bool) -> None:
        self._s.setValue("dark_mode", bool(value))

    # --- snapshot (read-only catalogue data) ---
    @property
    def snapshot_repo(self) -> str:
        return str(self._s.value("snapshot_repo", _default("SNAPSHOT_REPO"))).strip()

    @snapshot_repo.setter
    def snapshot_repo(self, value: str) -> None:
        self._s.setValue("snapshot_repo", (value or "").strip())

    @property
    def snapshot_token(self) -> str:
        return str(self._s.value("snapshot_token", _default("SNAPSHOT_TOKEN"))).strip()

    @snapshot_token.setter
    def snapshot_token(self, value: str) -> None:
        self._s.setValue("snapshot_token", (value or "").strip())

    @property
    def snapshot_branch(self) -> str:
        return str(self._s.value("snapshot_branch", _default("SNAPSHOT_BRANCH", "main"))).strip() or "main"

    # --- admin (write access) ---
    @property
    def db_url(self) -> str:
        """Never has a shipped default — see the policy note above."""
        return str(self._s.value("db_url", "")).strip()

    @db_url.setter
    def db_url(self, value: str) -> None:
        self._s.setValue("db_url", (value or "").strip())

    @property
    def admin_enabled(self) -> bool:
        return bool(self.db_url)

    # --- optional AI ---
    @property
    def anthropic_key(self) -> str:
        return str(self._s.value("anthropic_key", "")).strip()

    @anthropic_key.setter
    def anthropic_key(self, value: str) -> None:
        self._s.setValue("anthropic_key", (value or "").strip())

    # --- last-used folders, so pickers open where the user left off ---
    def last_dir(self, key: str) -> str:
        return str(self._s.value(f"last_dir/{key}", "")).strip()

    def set_last_dir(self, key: str, path: str) -> None:
        self._s.setValue(f"last_dir/{key}", path or "")
