# -*- coding: utf-8 -*-
"""One place that decides what a connection string means.

Supabase's "Connect" panel hands out several forms of the same URL, and two of
them break us:

  postgresql+psycopg://...    the ORM tab's default — that's psycopg **3**, and
                              every requirements file here installs psycopg2, so
                              SQLAlchemy dies with
                              "ModuleNotFoundError: No module named 'psycopg'"
                              in under a second, before touching the network.

  db.<ref>.supabase.co        the direct host, which is IPv6-only. GitHub's
                              runners are IPv4-only, so the Actions cannot reach
                              it at all. Use a pooler host instead.

Both cost a morning of broken ingests once. Normalising here means it cannot
happen again from a copy-and-paste, on any machine or any runner.
"""
from __future__ import annotations

import re

_PSYCOPG3 = re.compile(r"^postgresql\+psycopg(?!2)", re.IGNORECASE)
_DIRECT_HOST = re.compile(r"@db\.[a-z0-9]+\.supabase\.co\b", re.IGNORECASE)


def normalise(db_url: str) -> str:
    """Return db_url with the psycopg3 scheme rewritten to plain postgresql://.

    Everything after the scheme is left exactly as-is — the password may contain
    anything, and rewriting it is not our business.
    """
    url = (db_url or "").strip()
    if not url:
        return url
    return _PSYCOPG3.sub("postgresql", url)


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
