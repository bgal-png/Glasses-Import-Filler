# -*- coding: utf-8 -*-
"""Tests for `fill_only_if_empty` - the rule that stops a lower-fidelity
catalogue from overwriting richer values already held.

The Marcolin "MATERIAL INFO" export carries one combined COLOR DESCRIPTION,
while the older Marcolin master had separate FRONT and TEMPLE colours that were
stored as "Blue|Havana". Ingesting the newer file must not flatten those to
"Blue" - but it must still fill the colour for products we have never seen.

Run against a real in-memory SQLite database, because the behaviour depends on
what DataFrame.update does with NaN rather than on anything we can assert by
reading the code.

Run:  python test_soft_columns.py
"""
from __future__ import annotations

import sys

import pandas as pd
from sqlalchemy import create_engine

from ingest import perform_upsert

_ok = True


def check(label, cond, detail=""):
    global _ok
    print(f"  {'PASS' if cond else 'FAIL'}  {label}"
          f"{(' - ' + str(detail)) if detail and not cond else ''}")
    _ok = _ok and bool(cond)


def _engine_with(rows):
    engine = create_engine("sqlite://")
    pd.DataFrame(rows).to_sql("master_catalog", con=engine, index=False, if_exists="replace")
    return engine


def _read(engine):
    return pd.read_sql_table("master_catalog", con=engine).set_index("join_key")


def main() -> int:
    print("1. a soft column never overwrites a stored value")
    engine = _engine_with([
        {"join_key": "111", "Frame_Colour": "Blue|Havana", "Glasses_shape": "Oval"},
        {"join_key": "222", "Frame_Colour": "", "Glasses_shape": "Oval"},
        {"join_key": "333", "Frame_Colour": None, "Glasses_shape": "Oval"},
    ])
    incoming = pd.DataFrame([
        {"join_key": "111", "Frame_Colour": "Blue", "Glasses_shape": "Round"},
        {"join_key": "222", "Frame_Colour": "Green", "Glasses_shape": "Round"},
        {"join_key": "333", "Frame_Colour": "Pink", "Glasses_shape": "Round"},
        {"join_key": "444", "Frame_Colour": "Grey", "Glasses_shape": "Round"},
    ])
    incoming.attrs["fill_only_if_empty"] = ("Frame_Colour",)
    perform_upsert(incoming, engine)
    after = _read(engine)

    check("two-tone value kept", after.loc["111", "Frame_Colour"] == "Blue|Havana",
          after.loc["111", "Frame_Colour"])
    check("empty string gets filled", after.loc["222", "Frame_Colour"] == "Green",
          after.loc["222", "Frame_Colour"])
    check("NULL gets filled", after.loc["333", "Frame_Colour"] == "Pink",
          after.loc["333", "Frame_Colour"])
    check("brand-new product keeps its colour", after.loc["444", "Frame_Colour"] == "Grey",
          after.loc["444", "Frame_Colour"])
    check("NON-soft columns still update normally",
          after.loc["111", "Glasses_shape"] == "Round", after.loc["111", "Glasses_shape"])
    check("row count right", len(after) == 4, len(after))

    print("2. without the flag, everything overwrites as before")
    engine = _engine_with([{"join_key": "111", "Frame_Colour": "Blue|Havana"}])
    plain = pd.DataFrame([{"join_key": "111", "Frame_Colour": "Blue"}])
    perform_upsert(plain, engine)
    check("no flag -> normal overwrite", _read(engine).loc["111", "Frame_Colour"] == "Blue",
          _read(engine).loc["111", "Frame_Colour"])

    print("3. a flag naming a column nobody has is harmless")
    engine = _engine_with([{"join_key": "111", "Frame_Colour": "Blue|Havana"}])
    odd = pd.DataFrame([{"join_key": "111", "Frame_Colour": "Blue"}])
    odd.attrs["fill_only_if_empty"] = ("Nope", "Frame_Colour")
    perform_upsert(odd, engine)
    check("unknown column ignored, known one honoured",
          _read(engine).loc["111", "Frame_Colour"] == "Blue|Havana")

    print("4. the real handler declares the flag")
    import ingest
    tiny = pd.DataFrame([{
        "EAN/UPC CODE": "889214614209", "BRAND": "GU", "CODE SUN/OPT": "01030",
        "SIZE": "57", "NOSE-BRIDGE SIZE": "16", "TEMPLE LENGHT": "140",
        "B Measurement": "28", "DESCRIPTION FRONT": "INJECTED",
        "FORM DESCRIPTION": "CAT", "RIM DESCRIPTION": "FULL RIM", "FLEX": "NO",
        "GENDER": "F", "COLOR DESCRIPTION": "shiny black / smoke",
        "LENSES CATEGORY": "3", "LENSES DESCRIPTION": "NORMAL", "REXABLE": "SI",
        "ORIGIN": "CN", "NET WEIGHT": "0.038", "MODEL": "GU00001", "SKU": "5701A",
        "CLIPIN": "No ClipIn", "CLIPON": "No ClipOn",
    }])
    out, _unmapped, _skipped = ingest._load_marcolin_mi(tiny)
    check("handler flags Frame_Colour",
          out.attrs.get("fill_only_if_empty") == ("Frame_Colour",),
          out.attrs)
    check("and Temple_Colour is absent entirely", "Temple_Colour" not in out.columns)

    print()
    print("ALL PASS" if _ok else "FAILURES ABOVE")
    return 0 if _ok else 1


if __name__ == "__main__":
    sys.exit(main())
