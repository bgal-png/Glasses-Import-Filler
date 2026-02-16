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
    st.error("❌ Critical Error: 'mappings.py' file is missing. Ensure it is in the same directory.")
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
    """SAFETY CHECK: Only fills if cell is currently empty."""
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

# Pre-flattened lookups for Rule 3, 4, 5, 8
FLAT_BRAND_LOOKUP = {b.lower().strip(): c for c, brands in BRAND_TO_COMPANY_MAP.items() for b in brands}
FLAT_USABLE_LOOKUP = {b.lower().strip(): cat for cat, brands in BRAND_TO_USABLE_MAP.items() for b in brands}
FLAT_COLLECTION_LOOKUP = {b.lower().strip(): coll for coll, brands in BRAND_TO_COLLECTION_MAP.items() for b in brands}
FLAT_FACE_LOOKUP = {s.lower().strip(): face for face, sources in FACE_SHAPE_MAP.items() for s in sources}

def apply_hs_code(row, type_col, mat_col, sport_col):
    g_type = str(row.get(type_col, '')).strip()
    material = str(row.get(mat_col, '')).strip().lower()
    sport_val = str(row.get(sport_col, '')).strip().lower()
    if check_type(g_type, "Sunglasses"): return "90041091", "Sunglasses"
    if check_type(g_type, "Sport glasses"):
        if any(x in sport_val for x in ["swimm", "swim", "ski", "snowboard"]): return "90049090", "Sport Specialty"
        return "90041091", "Sport Protection"
    if "plastic" in material: return "90031100", "Eyewear + Plastic"
    if "metal" in material: return "90031900", "Eyewear + Metal"
    return "", ""

def apply_item_description(row, type_col, mat_col):
    g_type = str(row.get(type_col, '')).strip()
    material = str(row.get(mat_col, '')).strip().lower()
    if any(check_type(g_type, t) for t in ["Frames", "PC Glasses", "Reading glasses"]): return "Eyeglasses", "Eyewear"
    if check_type(g_type, "Sunglasses"):
        return ("Sunglasses, plastic frame" if "plastic" in material else "Sunglasses, metal frame"), "Sunglasses"
    return "Sport glasses", "Sport"

def apply_glasses_usable(row, brand_col, type_col, effect_col, sport_col):
    brand = str(row.get(brand_col, '')).strip().lower()
    g_type = str(row.get(type_col, ''))
    effect = str(row.get(effect_col, '')).strip().lower()
    sport = str(row.get(sport_col, '')).strip().lower()
    res = [FLAT_USABLE_LOOKUP[brand]] if brand in FLAT_USABLE_LOOKUP else []
    if check_type(g_type, "Sunglasses"):
        if "polarized" in effect: res.append("Driving glasses")
        elif not any(x in sport for x in ["swim", "ski"]): res.append("Common use")
    return ("|".join(res), "Matched") if res else ("", "")

def apply_glasses_gender(row, gen_col, color_col, shape_col, tl_col, lw_col):
    curr = str(row.get(gen_col, '')).strip()
    color, shape = str(row.get(color_col, '')).strip().lower(), str(row.get(shape_col, '')).strip().lower()
    is_fem = "cat eye" in shape or any(x in color for x in ["pink", "purple"])
    if "child" in curr.lower():
        tl, lw = to_float(row.get(tl_col, 0)), to_float(row.get(lw_col, 0))
        genders = ["Child-Girls"] if is_fem else ["Child-Boys", "Child-Girls"]
        sizes = []
        if 0 < tl <= 125 and 0 < lw <= 40: sizes.append("Toddlers")
        if 115 <= tl <= 135 and 35 <= lw <= 48: sizes.append("PreschoolKids")
        if 130 <= tl <= 145 and 45 <= lw <= 99: sizes.append("SchoolKids")
        final = ["Child"] + genders + [f"{g}-{s}" for g in genders for s in sizes]
        return "|".join(final), "Expanded Child"
    if curr in ["", "nan"] and is_fem: return "Woman", "Inferred Woman"
    return curr, ""

def apply_lenses_no_order(row, ft_col, con_col):
    ft, con = str(row.get(ft_col, '')), str(row.get(con_col, '')).lower()
    res = []
    if check_type(ft, "Half rim"): res += ["CoatingPolarized", "Glasses index 1.5"]
    if check_type(ft, "Rimless"): res += ["CoatingPolarized", "Glasses index 1.5", "Glasses index 1.74"]
    if "clip" in con: res.append("Glasses index 1.5")
    return ("|".join(list(dict.fromkeys(res))), "Restriction") if res else ("", "")

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
    return ("|".join(list(dict.fromkeys(features))), "Logic Applied") if features else ("", "")

def apply_glasses_model(row, name_col, brand_col):
    n, b = str(row.get(name_col, '')).strip(), str(row.get(brand_col, '')).strip()
    if b and n.lower().startswith(b.lower()): n = n[len(b):].strip()
    if " " in n: n = n.rsplit(" ", 1)[0]
    return n.strip(), "Extracted"

def apply_color_code(row, name_col):
    n = str(row.get(name_col, '')).strip()
    if " " in n: return n.rsplit(" ", 1)[1].strip(), "Extracted"
    return "", ""

# ==========================================
# 🚀 MAIN RUNNER
# ==========================================
def run_auto_fill(df):
    # Locate all IDs
    cols = {
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

    # Initialize target columns if missing
    for k in ['hs', 'desc', 'prod', 'usable', 'coll', 'uv', 'gender', 'face', 'no_order', 'other', 'model', 'code']:
        if cols[k] and cols[k] not in df.columns: df[cols[k]] = ""

    # Execute Rules
    df[cols['hs']], _ = apply_safe_fill(df, cols['hs'], df.apply(lambda r: apply_hs_code(r, cols['type'], cols['material'], cols['sport']), axis=1))
    df[cols['desc']], _ = apply_safe_fill(df, cols['desc'], df.apply(lambda r: apply_item_description(r, cols['type'], cols['material']), axis=1))
    df[cols['prod']], _ = apply_safe_fill(df, cols['prod'], df.apply(lambda r: apply_producing_company(r, cols['brand']), axis=1))
    df[cols['usable']], _ = apply_safe_fill(df, cols['usable'], df.apply(lambda r: apply_glasses_usable(r, cols['brand'], cols['type'], cols['effect'], cols['sport']), axis=1))
    df[cols['coll']], _ = apply_safe_fill(df, cols['coll'], df.apply(lambda r: apply_glasses_collection(r, cols['brand']), axis=1))
    df[cols['uv']], _ = apply_safe_fill(df, cols['uv'], df.apply(lambda r: apply_uv_filter(r, cols['type']), axis=1))
    
    gen_res = df.apply(lambda r: apply_glasses_gender(r, cols['gender'], cols['color'], cols['shape'], cols['tl'], cols['lw']), axis=1)
    df[cols['gender']] = [x[0] for x in gen_res]
    
    df[cols['face']], face_log = apply_safe_fill(df, cols['face'], df.apply(lambda r: apply_face_shape(r, cols['shape']), axis=1))
    df[cols['no_order']], no_log = apply_safe_fill(df, cols['no_order'], df.apply(lambda r: apply_lenses_no_order(r, cols['ft'], cols['con']), axis=1))
    df[cols['other']], feat_log = apply_safe_fill(df, cols['other'], df.apply(lambda r: apply_other_features(r, cols['type'], cols['con'], cols['rx']), axis=1))
    df[cols['model']], mod_log = apply_safe_fill(df, cols['model'], df.apply(lambda r: apply_glasses_model(r, cols['name'], cols['brand']), axis=1))
    df[cols['code']], code_log = apply_safe_fill(df, cols['code'], df.apply(lambda r: apply_color_code(r, cols['name']), axis=1))

    # STRICT CLIP: Find ID: 103 and drop everything after it
    guillotine_col = get_col_by_id(df, "103")
    if guillotine_col:
        idx = df.columns.get_loc(guillotine_col)
        df = df.iloc[:, :idx + 1]

    # Persistent Report
    report = pd.DataFrame({
        'Model Change': mod_log, 
        'Code Change': code_log, 
        'Features': feat_log,
        'Face': face_log
    }).replace("", pd.NA).dropna(how='all')
    
    return df, report

# ==========================================
# 📤 STREAMLIT UI
# ==========================================
uploaded_file = st.file_uploader("Upload Glasses Data", type=['xlsx'])

if uploaded_file:
    user_df = pd.read_excel(uploaded_file, dtype=str)
    if st.button("✨ Run Automation Suite"):
        with st.spinner("Processing all 12 rules..."):
            st.session_state.filled_df, st.session_state.report_df = run_auto_fill(user_df.copy())
            st.success("✅ Complete. Extra columns clipped.")

    if st.session_state.filled_df is not None:
        st.subheader("Processing Summary")
        st.dataframe(st.session_state.report_df, use_container_width=True)
        
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            st.session_state.filled_df.to_excel(writer, index=False, sheet_name='Final_Export')
            workbook, worksheet = writer.book, writer.sheets['Final_Export']
            
            # FORCE TEXT FORMAT FOR ID: 107
            text_format = workbook.add_format({'num_format': '@'})
            code_col = get_col_by_id(st.session_state.filled_df, "107")
            if code_col:
                idx = st.session_state.filled_df.columns.get_loc(code_col)
                worksheet.set_column(idx, idx, 15, text_format)
                
        st.download_button("📥 Download Finalized File", data=buffer.getvalue(), file_name="filled_glasses_final.xlsx")