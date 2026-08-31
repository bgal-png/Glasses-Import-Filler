# -*- coding: utf-8 -*-
"""Tests for the pure admin logic — no DB, no Qt, no network.

Run:  python test_admin_core.py
"""
from __future__ import annotations

import sys

import pandas as pd

import admin_core

_ok = True


def check(label, cond, detail=""):
    global _ok
    print(f"  {'PASS' if cond else 'FAIL'}  {label}"
          f"{(' — ' + str(detail)) if detail and not cond else ''}")
    _ok = _ok and bool(cond)


def main() -> int:
    print("1. clean_barcode — must match how master_catalog stores join_key")
    check("leading zeros stripped", admin_core.clean_barcode("0716736139197") == "716736139197")
    check("trailing .0 stripped", admin_core.clean_barcode("889214112798.0") == "889214112798")
    check("whitespace stripped", admin_core.clean_barcode("  716736139197 ") == "716736139197")
    check("plain value untouched", admin_core.clean_barcode("197737248260") == "197737248260")

    print("2. records_from_filled_file")
    filled = pd.DataFrame({
        "Barcode": ["0716736139197", "889214112798.0", ""],
        "Glasses name": ["Marc Jacobs MARC 400 ISK", "Tom Ford FT0613 01D", "x"],
        "Combination (size on glasses)": ["54", "52", "50"],
    })
    recs = admin_core.records_from_filled_file(filled)
    check("blank barcode dropped", len(recs) == 2, len(recs))
    check("join_key normalised", recs["join_key"].tolist() == ["716736139197", "889214112798"],
          recs["join_key"].tolist())
    check("raw barcode preserved", recs["barcode"].iloc[0] == "0716736139197")
    check("name and size carried",
          recs["name"].iloc[1] == "Tom Ford FT0613 01D" and recs["size"].iloc[1] == "52")

    # A file with only barcodes must still work — name/size just come back empty.
    bare = admin_core.records_from_filled_file(pd.DataFrame({"Barcode": ["123"]}))
    check("barcode-only file accepted", len(bare) == 1 and bare["name"].iloc[0] == "",
          bare.to_dict("records"))

    try:
        admin_core.records_from_filled_file(pd.DataFrame({"Nope": [1]}))
        check("missing barcode column raises", False, "no exception")
    except ValueError:
        check("missing barcode column raises", True)

    print("3. get_engine refuses to run without a URL")
    try:
        admin_core.get_engine("")
        check("empty db_url raises", False, "no exception")
    except ValueError as e:
        check("empty db_url raises with guidance", "Settings" in str(e), str(e))

    print()
    print("ALL PASS" if _ok else "FAILURES ABOVE")
    return 0 if _ok else 1


if __name__ == "__main__":
    sys.exit(main())
