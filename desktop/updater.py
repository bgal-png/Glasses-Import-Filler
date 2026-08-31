# -*- coding: utf-8 -*-
"""Self-update from GitHub Releases.

Compares version.__version__ with the newest `desktop-v*` release tag. If the
user accepts, downloads the .exe asset next to the running one and swaps it via
a detached PowerShell command that waits for this process to exit, replaces the
file and relaunches it — a running .exe can't overwrite itself on Windows.
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


SWAP_LOG_NAME = "glassesfiller_update.log"


def swap_log_path() -> str:
    import tempfile
    return os.path.join(tempfile.gettempdir(), SWAP_LOG_NAME)


def _ps_quote(path: str) -> str:
    """Single-quote a path for PowerShell (doubling any embedded quote)."""
    return "'" + str(path).replace("'", "''") + "'"


def build_swap_script(pid: int, src: str, dst: str, log: str) -> str:
    """The PowerShell one-liner that replaces the running .exe and relaunches it.

    PowerShell rather than a .bat: the first version used `timeout` and `start`
    inside a DETACHED_PROCESS cmd, and neither works without a console —
    `timeout` fails with "input redirection is not supported", so the wait loop
    and the relaunch both collapsed and the app closed without reopening.
    `Wait-Process` needs no console and is the right primitive.

    The move is retried, because Windows can hold the old file briefly after the
    process object is gone, and everything is logged so a failure is diagnosable.
    """
    q_src, q_dst, q_log = _ps_quote(src), _ps_quote(dst), _ps_quote(log)
    return (
        "$ErrorActionPreference='Continue'; "
        f"function L($m){{ \"$(Get-Date -Format o) $m\" | Out-File -FilePath {q_log} -Append -Encoding utf8 }}; "
        f"L 'waiting for pid {pid}'; "
        f"Wait-Process -Id {pid} -Timeout 180 -ErrorAction SilentlyContinue; "
        "$moved=$false; "
        "for($i=0;$i -lt 40 -and -not $moved;$i++){ "
        f"  try{{ Move-Item -LiteralPath {q_src} -Destination {q_dst} -Force -ErrorAction Stop; $moved=$true }}"
        "  catch{ Start-Sleep -Milliseconds 500 } }; "
        "if($moved){ L 'replaced the exe' } else { L 'MOVE FAILED - the old version is still in place' }; "
        f"try{{ Start-Process -FilePath {q_dst}; L 'relaunched' }}catch{{ L \"relaunch failed: $_\" }}"
    )


def download_and_swap(release: dict, progress=None) -> str:
    """Download the new .exe and hand over to the swap script. Returns the path
    of the downloaded file. The swap only happens when running frozen."""
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

    if total and read < total:
        os.remove(new_path)
        raise IOError(f"Download incomplete ({read} of {total} bytes) — update aborted.")

    if not is_frozen():
        return new_path  # nothing to swap when running from source

    script = build_swap_script(os.getpid(), new_path, sys.executable, swap_log_path())
    subprocess.Popen(
        ["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command", script],
        creationflags=0x08000000,  # CREATE_NO_WINDOW
    )
    return new_path
