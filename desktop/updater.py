# -*- coding: utf-8 -*-
"""Self-update from GitHub Releases.

Compares version.__version__ with the newest `desktop-v*` release tag. If the
user accepts, downloads the .exe asset next to the running one and swaps it via
a small .bat that waits for this process to exit, replaces the file and
relaunches — the running .exe can't overwrite itself on Windows.
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import urllib.request

from app_paths import is_frozen
from version import RELEASE_REPO, RELEASE_TAG_PREFIX, __version__

_TIMEOUT = 30


def _version_tuple(text: str) -> tuple:
    parts = []
    for chunk in str(text).strip().lstrip("vV").split("."):
        digits = "".join(ch for ch in chunk if ch.isdigit())
        parts.append(int(digits) if digits else 0)
    while len(parts) < 3:
        parts.append(0)
    return tuple(parts[:3])


def latest_release() -> dict | None:
    """Newest desktop release as {'version','tag','url','name'} or None."""
    url = f"https://api.github.com/repos/{RELEASE_REPO}/releases"
    req = urllib.request.Request(url)
    req.add_header("User-Agent", "GlassesFiller-Desktop")
    req.add_header("Accept", "application/vnd.github+json")
    with urllib.request.urlopen(req, timeout=_TIMEOUT) as resp:
        releases = json.loads(resp.read().decode("utf-8"))

    best = None
    for rel in releases:
        tag = str(rel.get("tag_name", ""))
        if not tag.startswith(RELEASE_TAG_PREFIX) or rel.get("draft"):
            continue
        version = tag[len(RELEASE_TAG_PREFIX):]
        asset = next(
            (a for a in rel.get("assets", []) if str(a.get("name", "")).lower().endswith(".exe")),
            None,
        )
        if not asset:
            continue
        candidate = {
            "version": version,
            "tag": tag,
            "url": asset["browser_download_url"],
            "name": asset["name"],
        }
        if best is None or _version_tuple(version) > _version_tuple(best["version"]):
            best = candidate
    return best


def update_available() -> dict | None:
    """The newer release, or None if we're current / can't tell."""
    try:
        rel = latest_release()
    except Exception:
        return None
    if rel and _version_tuple(rel["version"]) > _version_tuple(__version__):
        return rel
    return None


def download_and_swap(release: dict, progress=None) -> str:
    """Download the new .exe and hand over to a swap script. Returns the path
    of the downloaded file. Only meaningful when running frozen."""
    target_dir = os.path.dirname(sys.executable if is_frozen() else os.path.abspath(__file__))
    new_path = os.path.join(target_dir, f"_update_{release['name']}")

    req = urllib.request.Request(release["url"])
    req.add_header("User-Agent", "GlassesFiller-Desktop")
    with urllib.request.urlopen(req, timeout=300) as resp:
        total = int(resp.headers.get("Content-Length") or 0)
        read = 0
        with open(new_path, "wb") as fh:
            while True:
                chunk = resp.read(256 * 1024)
                if not chunk:
                    break
                fh.write(chunk)
                read += len(chunk)
                if progress and total:
                    progress(read / total, f"Downloading update… {read // 1048576} MB")

    if not is_frozen():
        return new_path  # nothing to swap when running from source

    current = sys.executable
    bat = os.path.join(target_dir, "_update.bat")
    with open(bat, "w", encoding="utf-8") as fh:
        fh.write(
            "@echo off\r\n"
            "echo Updating Glasses Filler...\r\n"
            ":wait\r\n"
            f'tasklist /fi "PID eq {os.getpid()}" | find "{os.getpid()}" >nul\r\n'
            "if not errorlevel 1 (\r\n"
            "  timeout /t 1 /nobreak >nul\r\n"
            "  goto wait\r\n"
            ")\r\n"
            f'move /y "{new_path}" "{current}" >nul\r\n'
            f'start "" "{current}"\r\n'
            'del "%~f0"\r\n'
        )
    subprocess.Popen(["cmd", "/c", bat], creationflags=0x00000008)  # DETACHED_PROCESS
    return new_path
