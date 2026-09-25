# -*- coding: utf-8 -*-
"""One place that decides which PostgreSQL driver a connection string means.

SQLAlchemy 2.1 (released 2026-09-24 20:36 UTC) changed the default DBAPI for
`postgresql://` from psycopg2 to psycopg 3. Every requirements file here
installs psycopg2-binary and `SQLAlchemy` is deliberately unpinned, so the first
CI run after that release died with

    ModuleNotFoundError: No module named 'psycopg'

in under a second, before touching the network - and every ingest and snapshot
publish stayed broken until the driver was named explicitly. Supabase's Connect
panel hands out `postgresql+psycopg://` as well, which fails the same way.

So: always name psycopg2 in the URL. An explicit driver cannot be changed by an
upstream default, which is why this is better than pinning SQLAlchemy.

Separately, the direct `db.<ref>.supabase.co` host is IPv6-only while GitHub
runners are IPv4-only, so that host can never work from Actions. `warnings()`
says so rather than leaving it to be rediscovered.
"""
from __future__ import annotations

import re

_PSYCOPG3 = re.compile(r"^postgresql\+psycopg(?!2)", re.IGNORECASE)
_BARE = re.compile(r"^postgres(?:ql)?://", re.IGNORECASE)
_DIRECT_HOST = re.compile(r"@db\.[a-z0-9]+\.supabase\.co\b", re.IGNORECASE)

DRIVER = "postgresql+psycopg2"


def normalise(db_url: str) -> str:
    """Return db_url with the driver pinned to psycopg2.

    Everything after the scheme is left exactly as-is - the password may contain
    anything, and rewriting it is not our business.
    """
    url = (db_url or "").strip()
    if not url:
        return url
    url = _PSYCOPG3.sub(DRIVER, url)
    # A bare scheme means "whatever the installed SQLAlchemy defaults to", which
    # is precisely what changed underneath us. Pin it to the driver we install.
    return _BARE.sub(f"{DRIVER}://", url, count=1)


def warnings(db_url: str) -> list:
    """Human-readable problems worth printing. Never includes the password."""
    out = []
    url = (db_url or "").strip()
    if _PSYCOPG3.match(url):
        out.append(
            "connection string asked for psycopg3 (postgresql+psycopg://); "
            "using psycopg2 instead - the driver this project installs"
        )
    if _DIRECT_HOST.search(url):
        out.append(
            "connection string uses the direct db.<ref>.supabase.co host, which "
            "is IPv6-only - GitHub Actions runners are IPv4-only and cannot "
            "reach it. Use the Session pooler host (port 5432)"
        )
    return out
