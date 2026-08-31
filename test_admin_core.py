# -*- coding: utf-8 -*-
"""Tests for the pure admin logic — no Qt, no network, no Supabase.

The single-row lookup/edit path is exercised against a real in-memory SQLite
database, so the SQL itself is tested rather than assumed. That matters: it
writes to the live catalogue in production.

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


def _sqlite_engine():
    """A throwaway master_catalog with two rows, shaped like the real one."""
    from sqlalchemy import create_engine
    engine = create_engine("sqlite://")   # in-memory, single connection pool
    pd.DataFrame([
        {"join_key": "716736139197", "Barcode": "716736139197",
         "Assembled_Name": "Marc Jacobs MARC 400 ISK", "Brand": "Marc Jacobs",
         "Frame_Colour": "Black", "Glasses_type": "Frames",
         "Producing_company": "Safilo"},
        {"join_key": "889214112798", "Barcode": "889214112798",
         "Assembled_Name": "Tom Ford FT0613 01D", "Brand": "Tom Ford",
         "Frame_Colour": "", "Glasses_type": "Sunglasses",
         "Producing_company": "Marcolin"},
    ]).to_sql("master_catalog", con=engine, index=False, if_exists="replace")
    return engine


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
    bare = admin_core.records_from_filled_file(pd.DataFrame({"Barcode": ["123"]}))
    check("barcode-only file accepted", len(bare) == 1 and bare["name"].iloc[0] == "")
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

    print("4. fetch_row — targeted lookup against a real SQL engine")
    engine = _sqlite_engine()
    row = admin_core.fetch_row(engine, "716736139197")
    check("row found", row is not None)
    check("right product", row is not None and row["Assembled_Name"] == "Marc Jacobs MARC 400 ISK",
          None if row is None else row.get("Assembled_Name"))
    check("normalises a messy barcode",
          admin_core.fetch_row(engine, " 0716736139197 ") is not None)
    check("unknown barcode -> None", admin_core.fetch_row(engine, "999999999999") is None)
    check("empty input -> None", admin_core.fetch_row(engine, "") is None)

    print("5. update_row — writes only what changed")
    res = admin_core.update_row(engine, "716736139197",
                                {"Frame_Colour": "Havana", "Brand": "Marc Jacobs"})
    check("reports 2 columns", res["updated"] == 2, res)
    after = admin_core.fetch_row(engine, "716736139197")
    check("value actually changed in the DB", after["Frame_Colour"] == "Havana",
          after["Frame_Colour"])
    check("untouched field intact", after["Assembled_Name"] == "Marc Jacobs MARC 400 ISK")

    other = admin_core.fetch_row(engine, "889214112798")
    check("the OTHER row was not touched",
          other["Assembled_Name"] == "Tom Ford FT0613 01D" and other["Frame_Colour"] == "",
          other.to_dict())

    check("no changes -> no write", admin_core.update_row(engine, "716736139197", {})["updated"] == 0)
    check("writing an empty string is allowed",
          admin_core.update_row(engine, "716736139197", {"Frame_Colour": ""})["updated"] == 1)
    check("empty landed", admin_core.fetch_row(engine, "716736139197")["Frame_Colour"] == "")

    print("6. update_row — refuses anything it shouldn't write")
    try:
        admin_core.update_row(engine, "716736139197", {"Frame_Colour = 'x'; DROP TABLE": "y"})
        check("unknown/injected column rejected", False, "no exception")
    except ValueError as e:
        check("unknown/injected column rejected", "Unknown column" in str(e), str(e))
    # the table must still be there after that attempt
    check("table survived the injection attempt",
          admin_core.fetch_row(engine, "716736139197") is not None)

    check("join_key can't be edited away",
          admin_core.update_row(engine, "716736139197", {"join_key": "hacked"})["updated"] == 0)
    check("join_key intact", admin_core.fetch_row(engine, "716736139197") is not None)

    try:
        admin_core.update_row(engine, "999999999999", {"Brand": "x"})
        check("missing row raises", False, "no exception")
    except ValueError as e:
        check("missing row raises", "not in the catalogue" in str(e), str(e))

    print()
    print("ALL PASS" if _ok else "FAILURES ABOVE")
    return 0 if _ok else 1


if __name__ == "__main__":
    sys.exit(main())
