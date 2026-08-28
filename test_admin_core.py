# -*- coding: utf-8 -*-
"""Tests for the pure admin logic — no DB, no Qt, no network.

Covers the two pieces most likely to be wrong in a way nobody notices:
the colour worklist (which fields count as "missing") and the photo matcher
(filenames vs model + colour code).

Run:  python test_admin_core.py
"""
from __future__ import annotations

import sys

import pandas as pd

import admin_core
from dictionaries import MANUFACTURER_CONFIG
from ingest import load_single_catalog

CATALOGUE = r"C:\Users\blank\Downloads\DERIGO_Master data 06.2026.xlsx"

_ok = True


def check(label, cond, detail=""):
    global _ok
    print(f"  {'PASS' if cond else 'FAIL'}  {label}"
          f"{(' — ' + str(detail)) if detail and not cond else ''}")
    _ok = _ok and bool(cond)


def main() -> int:
    # ---------------- worklist on synthetic rows (exact control) -------------
    print("1. build_colour_worklist — which fields count as missing")
    rows = pd.DataFrame([
        # sunglasses, everything missing, no clip -> frame+temple+lens, gradient offered
        dict(join_key="1", Extracted_Model="M1", Glasses_color_code="01A", Combination="52",
             Glasses_type="Sunglasses", Brand="Police", Assembled_Name="Police M1 01A",
             Frame_Colour="", Temple_Colour="", Glasses_lens_Colour="",
             Clip_on_lens_colour="", Extracted_Clip_on="", Glasses_lens_effect="",
             Producing_company="Derigo"),
        # same model+colour, different size -> must MERGE into one group
        dict(join_key="2", Extracted_Model="M1", Glasses_color_code="01A", Combination="54",
             Glasses_type="Sunglasses", Brand="Police", Assembled_Name="Police M1 01A",
             Frame_Colour="", Temple_Colour="", Glasses_lens_Colour="",
             Clip_on_lens_colour="", Extracted_Clip_on="", Glasses_lens_effect="",
             Producing_company="Derigo"),
        # optical frame: lens colour legitimately empty -> must NOT be asked for
        dict(join_key="3", Extracted_Model="M2", Glasses_color_code="02B", Combination="50",
             Glasses_type="Frames", Brand="Furla", Assembled_Name="Furla M2 02B",
             Frame_Colour="", Temple_Colour="Black", Glasses_lens_Colour="",
             Clip_on_lens_colour="", Extracted_Clip_on="", Glasses_lens_effect="",
             Producing_company="Derigo"),
        # clip-on item -> clip lens colour IS asked for
        dict(join_key="4", Extracted_Model="M3", Glasses_color_code="03C", Combination="53",
             Glasses_type="Frames", Brand="Police", Assembled_Name="Police M3 03C",
             Frame_Colour="Black", Temple_Colour="Black", Glasses_lens_Colour="",
             Clip_on_lens_colour="", Extracted_Clip_on="Magnetic sun clip-on p",
             Glasses_lens_effect="Polarized", Producing_company="Derigo"),
        # nothing missing and gradient already present -> excluded entirely
        dict(join_key="5", Extracted_Model="M4", Glasses_color_code="04D", Combination="55",
             Glasses_type="Sunglasses", Brand="Police", Assembled_Name="Police M4 04D",
             Frame_Colour="Black", Temple_Colour="Black", Glasses_lens_Colour="Grey",
             Clip_on_lens_colour="", Extracted_Clip_on="",
             Glasses_lens_effect="Gradient|Polarized", Producing_company="Derigo"),
        # different producer -> filtered out when scoping
        dict(join_key="6", Extracted_Model="M9", Glasses_color_code="09Z", Combination="52",
             Glasses_type="Sunglasses", Brand="Carrera", Assembled_Name="Carrera M9 09Z",
             Frame_Colour="", Temple_Colour="", Glasses_lens_Colour="",
             Clip_on_lens_colour="", Extracted_Clip_on="", Glasses_lens_effect="",
             Producing_company="Safilo"),
    ])

    work = admin_core.build_colour_worklist(rows)
    by_model = {g["model"]: g for g in work}
    check("M4 excluded (nothing missing, already gradient)", "M4" not in by_model,
          sorted(by_model))
    check("M1 present", "M1" in by_model)
    if "M1" in by_model:
        g = by_model["M1"]
        check("M1 merged both sizes into one group", len(g["barcodes"]) == 2, g["barcodes"])
        check("M1 asks frame+temple+lens",
              g["missing"] == ["Frame_Colour", "Temple_Colour", "Glasses_lens_Colour"],
              g["missing"])
        check("M1 offers gradient", g["can_gradient"] is True)
    if "M2" in by_model:
        g = by_model["M2"]
        check("M2 (optical frame) does NOT ask for lens colour",
              "Glasses_lens_Colour" not in g["missing"], g["missing"])
        check("M2 does not ask for the colour it already has",
              "Temple_Colour" not in g["missing"], g["missing"])
        check("M2 offers no gradient (not sunglasses)", g["can_gradient"] is False)
    if "M3" in by_model:
        g = by_model["M3"]
        check("M3 (has clip) asks for clip lens colour",
              "Clip_on_lens_colour" in g["missing"], g["missing"])

    scoped = admin_core.build_colour_worklist(rows, producers=["derigo"])
    check("producer scope excludes Safilo row",
          all(g["model"] != "M9" for g in scoped), [g["model"] for g in scoped])

    # ---------------- photo matching ----------------
    print("2. match_photos — realistic filename styles")
    work_m1 = [g for g in work if g["model"] == "M1"]
    variants = {
        r"C:\p\M1 5201A.jpg": "M1 5201A",          # model + size+colour
        r"C:\p\M1_01A_P00.png": "M1_01A_P00",      # model + stripped colour
        r"C:\p\m1-5201a-front.jpg": "m1-5201a-front",  # lowercase, dashes
        r"C:\p\M15201A.JPG": "M15201A",            # concatenated
    }
    for path, base in variants.items():
        matched, unmatched = admin_core.match_photos(work_m1, {path: base})
        check(f"matches {base!r}", len(matched) == 1 and not unmatched)

    matched, unmatched = admin_core.match_photos(
        work_m1, {r"C:\p\ZZ9 999.jpg": "ZZ9 999"}
    )
    check("rejects an unrelated filename", not matched and len(unmatched) == 1)

    # ---------------- records_from_filled_file ----------------
    print("3. records_from_filled_file")
    filled = pd.DataFrame({
        "Barcode": ["0716736139197", "889214112798.0", ""],
        "Glasses name": ["Marc Jacobs MARC 400 ISK", "Tom Ford FT0613 01D", "x"],
        "Combination (size on glasses)": ["54", "52", "50"],
    })
    recs = admin_core.records_from_filled_file(filled)
    check("blank barcode dropped", len(recs) == 2, len(recs))
    check("leading zero stripped for join_key",
          recs["join_key"].iloc[0] == "716736139197", recs["join_key"].iloc[0])
    check("trailing .0 stripped", recs["join_key"].iloc[1] == "889214112798",
          recs["join_key"].iloc[1])
    check("raw barcode preserved", recs["barcode"].iloc[0] == "0716736139197")
    check("name and size carried", recs["name"].iloc[1] == "Tom Ford FT0613 01D"
          and recs["size"].iloc[1] == "52")
    try:
        admin_core.records_from_filled_file(pd.DataFrame({"Nope": [1]}))
        check("missing barcode column raises", False, "no exception")
    except ValueError:
        check("missing barcode column raises", True)

    # ---------------- against a real catalogue ----------------
    print("4. Real catalogue (De Rigo) — worklist is sane")
    try:
        mdf, _u, _s = load_single_catalog("derigo", MANUFACTURER_CONFIG["derigo"], CATALOGUE)
    except Exception as e:
        print(f"  SKIP  catalogue not available ({e})")
        mdf = None
    if mdf is not None:
        real = admin_core.build_colour_worklist(mdf)
        print(f"        {len(mdf)} rows -> {len(real)} group(s) needing attention")
        frames_asking_lens = [
            g for g in real
            if "Sunglasses" not in g["type"] and "Glasses_lens_Colour" in g["missing"]
        ]
        check("no optical frame is asked for a lens colour",
              not frames_asking_lens, len(frames_asking_lens))
        check("every group has at least one barcode",
              all(g["barcodes"] for g in real))
        check("every group has something to do",
              all(g["missing"] or g["can_gradient"] for g in real))
        clip_groups = [g for g in real if "Clip_on_lens_colour" in g["missing"]]
        print(f"        {len(clip_groups)} clip-on group(s) want a clip lens colour")

    print()
    print("ALL PASS" if _ok else "FAILURES ABOVE")
    return 0 if _ok else 1


if __name__ == "__main__":
    sys.exit(main())
