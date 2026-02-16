import streamlit as st
import pandas as pd
import os
import io
import re

# IMPORT THE MAPPINGS
try:
    from mappings import (
        BRAND_TO_COMPANY_MAP, 
        BRAND_TO_USABLE_MAP, 
        BRAND_TO_COLLECTION_MAP,
        FACE_SHAPE_MAP
    )
except ImportError:
    st.error("❌ Critical Error: 'mappings.py' file is missing. Please ensure it's in the same folder.")
    st.stop()

# 1. Page Configuration
st.set_page_config(page_title="Excel Auto-Filler Master", layout="wide")
st.title("⚡ Glasses Data Automation: Full 12-Rule Suite")

# Persistent Session State
if 'filled_df' not in st.session_state:
    st.session_state.filled_df = None
if 'report_df' not in st.session_state:
    st.session_state.report_df = None

# ==========================================
# 🔒 DATA LOADERS & HELPERS
# ==========================================
@st.cache_data
def load_master():
    current_dir = os.getcwd()
    candidates = [f for f in os.listdir(current_dir) if "master_clean" in f and not f.startswith('~$')]
    if not candidates:
        st.error("❌ 'master_clean.xlsx' not found in repository.")
        st.stop()
    file_path = candidates[0]
    df = pd.read_excel(file_path, dtype=str, engine='openpyxl')
    df.columns = df.columns.astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
    return df

def get_col_by_id(df, target_id):
    """Finds a column name in the user's file that contains a specific ID number."""
    for col in df.columns:
        if re.search(f"ID[:\s]+{target_id}\\b", col):
            return col
    return None

def check_type(value, target):
    """Smart Check for multi-value types (Split by |)"""
    value = str(value).strip()
    parts = [p.strip() for p in value.split('|')]
    return target in parts

def to_float(val):
    """Safely convert string to float for measurements"""
    try:
        return float(str(val).replace(',', '.').strip())
    except:
        return 0.0

def apply_safe_fill(df, target_col, calculated_results):
    """SAFETY CHECK: Only fills if cell is currently empty or NaN."""
    current_vals = df[target_col].astype(str).fillna("")
    final_vals = []
    report_reasons = []
    
    for curr, (new_val, new_reason) in zip(current_vals, calculated_results):
        curr_clean = curr.strip()
        if curr_clean == "" or curr_clean.lower() == "nan":
            final_vals.append(new_val)
            report_reasons.append(new_reason)
        else:
            final_vals.append(curr)
            report_reasons.append("") 
            
    return final_vals, report_reasons

# ==========================================
# 🧠 THE BRAIN: RULE LOGIC
# ==========================================

# Pre-flattened lookups for mapping efficiency
FLAT_BRAND_LOOKUP = {b.lower().strip(): c for c, brands in BRAND_TO_COMPANY_MAP.items() for b in brands}
FLAT_USABLE_LOOKUP = {b.lower().strip(): cat for cat, brands in BRAND_TO_USABLE_MAP.items() for b in brands}
FLAT_COLLECTION_LOOKUP = {b.lower().strip(): coll for coll, brands in BRAND_TO_COLLECTION_MAP.items() for b in brands}
FLAT_FACE_LOOKUP = {s.lower().strip(): face for face, sources in FACE_SHAPE_MAP.items() for s in sources}

# Rule 1: HS Code
def apply_hs_code(row, type_col, mat_col, sport_col):
    g_type = str(row.get(type_col, '')).strip()
    material = str(row.get(mat_col, '')).strip().lower()
    sport_val = str(row.get(sport_col, '')).strip().lower()
    if check_type(g_type, "Sunglasses"): return "90041091", "Sunglasses"
    if check_type(g_type, "Sport glasses"):
        if any(x in sport_val for x in ["swimm", "swim", "ski", "snowboard"]): return "90049090", "Specialty Sport"
        return "90041091", "Protection Sport"
    if "plastic" in material: return "90031100", "Eyewear + Plastic"
    if "metal" in material: return "90031900", "Eyewear + Metal"
    return "", ""

# Rule 2: Item Description
def apply_item_description(row, type_col, mat_col):
    g_type = str(row.get(type_col, '')).strip()
    material = str(row.get(mat_col, '')).strip().lower()
    if any(check_type(g_type, t) for t in ["Frames", "PC Glasses without power", "Driving Glasses without power", "Reading glasses"]):
        return "Eyeglasses", "Eyewear Group"
    if check_type(g_type, "Sunglasses"):
        return ("Sunglasses, plastic frame" if "plastic" in material else "Sunglasses, metal frame"), "Sunglasses Material Match"
    return "Sport glasses", "Sport glasses match"

# Rule 4: Glasses Usable
def apply_glasses_usable(row, brand_col, type_col, effect_col, sport_col):
    brand = str(row.get(brand_col, '')).strip().lower()
    g_type = str(row.get(type_col, ''))
    effect = str(row.get(effect_col, '')).strip().lower()
    sport = str(row.get(sport_col, '')).strip().lower()
    res = [FLAT_USABLE_LOOKUP[brand]] if brand in FLAT_USABLE_LOOKUP else []
    if check_type(g_type, "Sunglasses"):
        if "polarized" in effect: res.append("Driving glasses")
        elif not any(x in sport for x in ["swim", "ski", "snowboard"]): res.append("Common use")
    return ("|".join(res), "Usable Match") if res else ("", "")

# Rule 7: Gender & Child Sizing
def apply_glasses_gender(row, gen_col, color_col, shape_col, tl_col, lw_col):
    curr = str(row.get(gen_col, '')).strip()
    color, shape = str(row.get(color_col, '')).strip().lower(), str(row.get(shape_col, '')).strip().lower()
    is_fem = "cat eye" in shape or any(x in color for x in ["pink", "purple", "cat eye"])
    if "child" in curr.lower():
        tl, lw = to_float(row.get(tl_col, 0)), to_float(row.get(lw_col, 0))
        genders = ["Child-Girls"] if is_fem else ["Child-Boys", "Child-Girls"]
        sizes = []
        if 0 < tl <= 125 and 0 < lw <= 40: sizes.append("Toddlers")
        if 115 <= tl <= 135 and 35 <= lw <= 48: sizes.append("PreschoolKids")
        if 130 <= tl <= 145 and 45 <= lw <= 99: sizes.append("SchoolKids")
        final = ["Child"] + genders + [f"{g}-{s}" for g in genders for s in sizes]
        return "|".join(final), "Expanded Child Sizes"
    if curr in ["", "nan"] and is_fem: return "Woman", "Inferred Woman"
    return curr, ""

# Rule 9: Lenses No Order
def apply_lenses_no_order(row, ft_col, con_col):
    ft, con = str(row.get(ft_col, '')), str(row.get(con_col, '')).lower()
    res = []
    if check_type(ft, "Half rim"): res += ["CoatingPolarized", "Glasses index 1.5"]
    if check_type(ft, "Rimless"): res += ["CoatingPolarized", "Glasses index 1.5", "Glasses index 1.74"]
    if "clip" in con: res.append("Glasses index 1.5")
    return ("|".join(list(dict.fromkeys(res))), "Restriction Applied") if res else ("", "")

# Rule 10: Other Features
def apply_other_features(row, type_col, con_col, rx_col):
    g_type, con, rx = str(row.get(type_col, '')), str(row.get(con_col, '')), str(row.get(rx_col, '')).lower()
    features = []
    has_clip = False
    if check_type(g_type, "Sport glasses") and "clip" in con.lower(): features.append("Sport glasses with diopter clip")
    clips = ["Magnetic sun clip-on p", "Magnetic sun clip-on", "Sun clip-on p", "Sun clip-on"]
    for c in clips:
        if check_type(con, c): features.append(c); has_clip = True
    if check_type(g_type, "Frames") and has_clip: features.insert(0, "Glasses with sun clip-on")
    if rx == "yes": features.append("Prescription sunglasses")
    return ("|".join(list(dict.fromkeys(features))), "Features Added") if features else ("", "")

# Rules 11 & 12: Model & Color Extraction
def apply_model_and_code(row, name_col, brand_col, mode="model"):
    n, b = str(row.get(name_col, '')).strip(), str(row.get(brand_col, '')).strip()
    if not n or " " not in n: return ("", "")
    
    # Extract Color (Last part)
    code = n.rsplit(" ", 1)[1].strip()
    if mode == "code": return (code, "Extracted Code")
    
    # Extract Model (Middle part)
    model_part = n.rsplit(" ", 1)[0].strip()
    if b and model_part.lower().startswith(b.lower()):
        model_part = model_part[len(b):].strip()
    return (model_part, "Extracted Model")

# ==========================================
# 🚀 MAIN PROCESSING ENGINE
# ==========================================
def run_auto_fill(df):
    # ID Mapping Dictionary
    ids = {
        'name': df.columns[0],
        'brand': get_col_by_id(df, "11"),
        'model': get_col_by_id(df, "12"),
        'type': get_col_by_id(df, "13"),
        'gender': get_col_by_id(df, "22"),
        'shape': get_col_by_id(df, "25"),
        'color': get_col_by_id(df, "26"),
        'coll': get_col_by_id(df, "33"),
        'effect': get_col_by_id(df, "37"),
        'ft': get_col_by_id(df, "50"),
        'usable': get_col_by_id(df, "51"),
        'material': get_col_by_id(df, "53"),
        'uv': get_col_by_id(df, "60"),
        'tl': get_col_by_id(df, "70"),
        'lw': get_col_by_id(df, "72"),
        'con': get_col_by_id(df, "84"),
        'sport': get_col_by_id(df, "89"),
        'face': get_col_by_id(df, "94"),
        'no_order': get_col_by_id(df, "103"),
        'other': get_col_by_id(df, "104"),
        'code': get_col_by_id(df, "107"),
        'rx': get_col_by_id(df, "108"),
        'hs': get_col_by_id(df, "AO") or "HS Code",
        'desc': get_col_by_id(df, "AP") or "Item description",
        'prod': get_col_by_id(df, "146") or "Producing company"
    }

    # Apply Rules
    df[ids['hs']], _ = apply_safe_fill(df, ids['hs'], df.apply(lambda r: apply_hs_code(r, ids['type'], ids['material'], ids['sport']), axis=1))
    df[ids['desc']], _ = apply_safe_fill(df, ids['desc'], df.apply(lambda r: apply_item_description(r, ids['type'], ids['material']), axis=1))
    df[ids['prod']], _ = apply_safe_fill(df, ids['prod'], df.apply(lambda r: (FLAT_BRAND_LOOKUP.get(str(r.get(ids['brand'], '')).lower(), ""), "Match") if ids['brand'] else ("", ""), axis=1))
    df[ids['usable']], _ = apply_safe_fill(df, ids['usable'], df.apply(lambda r: apply_glasses_usable(r, ids['brand'], ids['type'], ids['effect'], ids['sport']), axis=1))
    df[ids['coll']], _ = apply_safe_fill(df, ids['coll'], df.apply(lambda r: (FLAT_COLLECTION_LOOKUP.get(str(r.get(ids['brand'], '')).lower(), ""), "Match") if ids['brand'] else ("", ""), axis=1))
    df[ids['uv']], _ = apply_safe_fill(df, ids['uv'], df.apply(lambda r: (("400", "Sun") if check_type(r.get(ids['type'], ''), "Sunglasses") else ("", "")), axis=1))
    
    # Gender Logic (Rule 7)
    gen_res = df.apply(lambda r: apply_glasses_gender(r, ids['gender'], ids['color'], ids['shape'], ids['tl'], ids['lw']), axis=1)
    df[ids['gender']] = [x[0] for x in gen_res]
    
    # Face, No-Order, Features (Rule 8, 9, 10)
    df[ids['face']], face_log = apply_safe_fill(df, ids['face'], df.apply(lambda r: (FLAT_FACE_LOOKUP.get(str(r.get(ids['shape'], '')).lower(), ""), "Match") if ids['shape'] else ("", ""), axis=1))
    df[ids['no_order']], no_log = apply_safe_fill(df, ids['no_order'], df.apply(lambda r: apply_lenses_no_order(r, ids['ft'], ids['con']), axis=1))
    df[ids['other']], feat_log = apply_safe_fill(df, ids['other'], df.apply(lambda r: apply_other_features(r, ids['type'], ids['con'], ids['rx']), axis=1))
    
    # Model & Color (Rule 11, 12)
    df[ids['model']], mod_log = apply_safe_fill(df, ids['model'], df.apply(lambda r: apply_model_and_code(r, ids['name'], ids['brand'], "model"), axis=1))
    df[ids['code']], code_log = apply_safe_fill(df, ids['code'], df.apply(lambda r: apply_model_and_code(r, ids['name'], ids['brand'], "code"), axis=1))

    # STRICT GUILLOTINE: Cut after ID: 103
    stop_col = get_col_by_id(df, "103")
    if stop_col:
        idx = df.columns.get_loc(stop_col)
        df = df.iloc[:, :idx + 1]

    # Report Preparation
    report = pd.DataFrame({'Model Logic': mod_log, 'Code Logic': code_log, 'Feature Logic': feat_log}).replace("", pd.NA).dropna(how='all')
    return df, report

# ==========================================
# 📤 USER INTERFACE
# ==========================================
file = st.file_uploader("Upload File", type=['xlsx'])
if file:
    u_df = pd.read_excel(file, dtype=str)
    if st.button("✨ Execute All 12 Rules"):
        st.session_state.filled_df, st.session_state.report_df = run_auto_fill(u_df.copy())
        st.success("✅ Rules Applied. Extras clipped at ID: 103.")

    if st.session_state.filled_df is not None:
        st.dataframe(st.session_state.report_df, use_container_width=True)
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
            st.session_state.filled_df.to_excel(writer, index=False, sheet_name='Sheet1')
            workbook, worksheet = writer.book, writer.sheets['Sheet1']
            # TEXT FORMAT FOR ID: 107
            fmt = workbook.add_format({'num_format': '@'})
            c_col = get_col_by_id(st.session_state.filled_df, "107")
            if c_col:
                idx = st.session_state.filled_df.columns.get_loc(c_col)
                worksheet.set_column(idx, idx, 15, fmt)
        st.download_button("📥 Download Result", data=buf.getvalue(), file_name="output.xlsx")