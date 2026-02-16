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
    st.error("❌ Critical Error: 'mappings.py' file is missing or missing dictionaries.")
    st.stop()

# 1. Page Configuration
st.set_page_config(page_title="Excel Auto-Filler", layout="wide")
st.title("⚡ Excel Data Filler: Glasses Edition")

# ==========================================
# 🔒 INDESTRUCTIBLE LOADER
# ==========================================
@st.cache_data
def load_master():
    current_dir = os.getcwd()
    candidates = [f for f in os.listdir(current_dir) if (f.endswith('.xlsx') or f.endswith('.csv')) and "master_clean" in f and not f.startswith('~$')]
    if not candidates:
        st.error("❌ 'master_clean.xlsx' not found in repository."); st.stop()
    file_path = candidates[0]
    df = None
    try:
        df = pd.read_excel(file_path, dtype=str, engine='openpyxl')
    except Exception:
        strategies = [{'sep': None, 'engine': 'python'}, {'sep': ',', 'engine': 'c'}, {'sep': ';', 'engine': 'c'}, {'sep': '\t', 'engine': 'c'}]
        for enc in ['utf-8', 'cp1252', 'latin1']:
            for strat in strategies:
                try:
                    df = pd.read_csv(file_path, dtype=str, encoding=enc, on_bad_lines='skip', **strat)
                    break
                except: continue
            if df is not None: break
    if df is None:
        st.error(f"❌ Could not read '{file_path}'."); st.stop()
    df.columns = df.columns.astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
    return df

# Load Master Data
raw_master_df = load_master()
target_col = next((c for c in raw_master_df.columns if "items type" in c.lower()), None)
if target_col:
    master_df = raw_master_df[raw_master_df[target_col].str.lower().str.strip() == "glasses"]
    st.success(f"✅ Brain Loaded: {len(master_df)} valid glasses rows.")
else:
    st.error("❌ 'Items type' column missing."); st.stop()

# ==========================================
# 🧠 THE BRAIN: FILLING LOGIC
# ==========================================

# --- PREPARE LOOKUPS ---
FLAT_BRAND_LOOKUP = {}
for company, brands in BRAND_TO_COMPANY_MAP.items():
    for brand in brands:
        FLAT_BRAND_LOOKUP[brand.lower().strip()] = company

FLAT_USABLE_LOOKUP = {}
for category, brands in BRAND_TO_USABLE_MAP.items():
    for brand in brands:
        FLAT_USABLE_LOOKUP[brand.lower().strip()] = category

FLAT_COLLECTION_LOOKUP = {}
for collection, brands in BRAND_TO_COLLECTION_MAP.items():
    for brand in brands:
        FLAT_COLLECTION_LOOKUP[brand.lower().strip()] = collection

FLAT_FACE_LOOKUP = {}
for face_shape, sources in FACE_SHAPE_MAP.items():
    for source in sources:
        FLAT_FACE_LOOKUP[source.lower().strip()] = face_shape

def get_col_by_id(df, target_id):
    for col in df.columns:
        if re.search(f"ID[:\s]+{target_id}\\b", col):
            return col
    return None

def check_type(value, target):
    value = str(value).strip()
    parts = [p.strip() for p in value.split('|')]
    return target in parts

def to_float(val):
    try:
        return float(str(val).replace(',', '.').strip())
    except:
        return 0.0

def apply_safe_fill(df, target_col, calculated_results):
    """
    CRITICAL SAFETY:
    - If cell is EMPTY -> Fill it (Reason included).
    - If cell has DATA -> Keep it (Reason is empty string).
    """
    current_vals = df[target_col].astype(str).fillna("")
    final_vals = []
    report_reasons = []
    
    for curr, (new_val, new_reason) in zip(current_vals, calculated_results):
        curr_clean = curr.strip()
        # Checks for true emptiness or the word 'nan'
        if curr_clean == "" or curr_clean.lower() == "nan":
            final_vals.append(new_val)
            report_reasons.append(new_reason)
        else:
            final_vals.append(curr)
            report_reasons.append("") # Empty reason = "I touched nothing"
            
    return final_vals, report_reasons

# --- RULE FUNCTIONS ---

def apply_hs_code(row, type_col, mat_col, sport_col):
    g_type = str(row.get(type_col, '')).strip()
    material = str(row.get(mat_col, '')).strip().lower()
    sport_val = str(row.get(sport_col, '')).strip().lower()

    if check_type(g_type, "Sunglasses"):
        return "90041091", "Group: Sunglasses"
    if check_type(g_type, "Sport glasses"):
        if any(x in sport_val for x in ["swimm", "swim", "ski", "snowboard"]):
            return "90049090", "Sport Specialty (Swim/Ski)"
        return "90041091", "Group: Sport (Protection)"
    eyewear_targets = ["Frames", "Reading glasses", "Driving Glasses without power", "PC Glasses without power"]
    if any(check_type(g_type, t) for t in eyewear_targets):
        if "plastic" in material: return "90031100", "Group: Eyewear + Plastic"
        if "metal" in material: return "90031900", "Group: Eyewear + Metal"
        return "", "Group: Eyewear - Missing Material"
    return "", "No Match"

def apply_item_description(row, type_col, mat_col):
    g_type = str(row.get(type_col, '')).strip()
    material = str(row.get(mat_col, '')).strip().lower()
    eyewear_targets = ["Frames", "PC Glasses without power", "Driving Glasses without power", "Reading glasses"]
    if any(check_type(g_type, t) for t in eyewear_targets):
        return "Eyeglasses", "Match: Eyewear Group"
    if check_type(g_type, "Sunglasses"):
        if "plastic" in material: return "Sunglasses, plastic frame", "Sunglasses + Plastic"
        if "metal" in material: return "Sunglasses, metal frame", "Sunglasses + Metal"
        return "Sunglasses", "Sunglasses (Unknown Material)"
    if check_type(g_type, "Sport glasses"):
        return "Sport glasses", "Exact match: Sport glasses"
    return "", "No Match"

def apply_producing_company(row, brand_col):
    brand_val = str(row.get(brand_col, '')).strip().lower()
    if brand_val in FLAT_BRAND_LOOKUP:
        return FLAT_BRAND_LOOKUP[brand_val], f"Matched Brand: {brand_val}"
    return "", "Unknown Brand"

def apply_glasses_usable(row, brand_col, type_col, effect_col, sport_col):
    brand_val = str(row.get(brand_col, '')).strip().lower()
    g_type = str(row.get(type_col, '')).strip()
    effect_val = str(row.get(effect_col, '')).strip().lower()
    sport_val = str(row.get(sport_col, '')).strip().lower()
    results = []
    if brand_val in FLAT_USABLE_LOOKUP:
        results.append(FLAT_USABLE_LOOKUP[brand_val])
    if check_type(g_type, "Sunglasses"):
        is_polarized = "polarized" in effect_val
        is_ski_swim = any(x in sport_val for x in ["swimm", "swim", "ski", "snowboard"])
        if is_polarized: results.append("Driving glasses")
        elif not is_ski_swim: results.append("Common use")
    if not results: return "", "No Match"
    return "|".join(results), f"Combined: {results}"

def apply_glasses_collection(row, brand_col):
    brand_val = str(row.get(brand_col, '')).strip().lower()
    if brand_val in FLAT_COLLECTION_LOOKUP:
        return FLAT_COLLECTION_LOOKUP[brand_val], f"Collection Match: {brand_val}"
    return "", ""

def apply_uv_filter(row, type_col):
    g_type = str(row.get(type_col, '')).strip()
    if check_type(g_type, "Sunglasses"):
        return "400", "Type is Sunglasses"
    return "", ""

def apply_glasses_gender(row, gender_col, color_col, shape_col, tl_col, lw_col):
    # This rule is SPECIAL: It allows editing of "Child" but respects other values.
    current_val = str(row.get(gender_col, '')).strip()
    current_lower = current_val.lower()
    color = str(row.get(color_col, '')).strip().lower()
    shape = str(row.get(shape_col, '')).strip().lower()
    is_female_specific = ("cat eye" in shape or "pink" in color or "purple" in color or "cat eye" in color)
    if "child" in current_lower:
        tl = to_float(row.get(tl_col, 0))
        lw = to_float(row.get(lw_col, 0))
        genders = ["Child-Girls"] if is_female_specific else ["Child-Boys", "Child-Girls"]
        size_cats = []
        if (0 < tl <= 125) and (0 < lw <= 40): size_cats.append("Toddlers")
        if (115 <= tl <= 135) and (35 <= lw <= 48): size_cats.append("PreschoolKids")
        if (130 <= tl <= 145) and (45 <= lw <= 99): size_cats.append("SchoolKids")
        final_parts = ["Child"]
        final_parts.extend(genders)
        for gender in genders:
            for size in size_cats:
                final_parts.append(f"{gender}-{size}")
        return "|".join(final_parts), f"Expanded Child"
    if current_val == "" or current_lower == "nan":
        if is_female_specific: return "Woman", "Inferred Woman"
    return current_val, ""

def apply_face_shape(row, shape_col):
    shape_val = str(row.get(shape_col, '')).strip().lower()
    if shape_val in FLAT_FACE_LOOKUP:
        return FLAT_FACE_LOOKUP[shape_val], f"Match: {shape_val}"
    return "", ""

def apply_lenses_no_order(row, frame_type_col, contain_col):
    frame_val = str(row.get(frame_type_col, '')).strip()
    contain_val = str(row.get(contain_col, '')).strip()
    restrictions = []
    if check_type(frame_val, "Half rim"): restrictions.extend(["CoatingPolarized", "Glasses index 1.5"])
    if check_type(frame_val, "Rimless"): restrictions.extend(["CoatingPolarized", "Glasses index 1.5", "Glasses index 1.74"])
    if "clip" in contain_val.lower(): restrictions.append("Glasses index 1.5")
    unique_restrictions = list(dict.fromkeys(restrictions))
    if not unique_restrictions: return "", ""
    return "|".join(unique_restrictions), f"Restrictions: {unique_restrictions}"

def apply_other_features(row, type_col, contain_col, rx_col):
    g_type = str(row.get(type_col, '')).strip()
    contain_val = str(row.get(contain_col, '')).strip()
    rx_val = str(row.get(rx_col, '')).strip().lower()
    features = []
    has_sun_clip = False
    
    if check_type(g_type, "Sport glasses") and "clip" in contain_val.lower():
        features.append("Sport glasses with diopter clip")
    
    target_clips = ["Magnetic sun clip-on p", "Magnetic sun clip-on", "Sun clip-on p", "Sun clip-on"]
    for clip in target_clips:
        if check_type(contain_val, clip):
            features.append(clip)
            has_sun_clip = True
            
    if check_type(g_type, "Frames") and has_sun_clip:
        features.insert(0, "Glasses with sun clip-on")
        
    if rx_val == "yes":
        features.append("Prescription sunglasses")
        
    unique_features = list(dict.fromkeys(features))
    if not unique_features: return "", ""
    return "|".join(unique_features), f"Features: {unique_features}"

def apply_glasses_model(row, name_col, brand_col):
    """Rule 11: Extract Model Name"""
    full_name = str(row.get(name_col, '')).strip()
    brand = str(row.get(brand_col, '')).strip()
    if not full_name: return "", ""
    # 1. CUT START: Remove Brand
    if brand and full_name.lower().startswith(brand.lower()):
        full_name = full_name[len(brand):].strip()
    # 2. CUT END: Remove everything after the last space
    if " " in full_name:
        full_name = full_name.rsplit(" ", 1)[0]
    # 3. Trim
    cleaned_model = full_name.strip()
    if not cleaned_model: return "", "Result Empty"
    return cleaned_model, "Extracted"


# --- MAIN EXECUTION ---

def run_auto_fill(user_df):
    # Identify Columns
    name_col = user_df.columns[0]
    
    type_col = get_col_by_id(user_df, "13")      
    material_col = get_col_by_id(user_df, "53")  
    sport_col = get_col_by_id(user_df, "89")
    brand_col = get_col_by_id(user_df, "11") 
    effect_col = get_col_by_id(user_df, "37") 
    color_col = get_col_by_id(user_df, "26") 
    shape_col = get_col_by_id(user_df, "25") 
    tl_col = get_col_by_id(user_df, "70")    
    lw_col = get_col_by_id(user_df, "72")    
    frame_type_col = get_col_by_id(user_df, "50")
    contain_col = get_col_by_id(user_df, "84")
    rx_col = get_col_by_id(user_df, "108") or "SunGlasses RX lenses"
    
    hs_col = get_col_by_id(user_df, "AO") or "HS Code"
    desc_col = get_col_by_id(user_df, "AP") or "Item description"
    prod_col = get_col_by_id(user_df, "146") or "Producing company"
    usable_col = get_col_by_id(user_df, "51") or "Glasses usable"
    coll_col = get_col_by_id(user_df, "33") or "Glasses collection"
    uv_col = get_col_by_id(user_df, "60") or "UV filter"
    gender_col = get_col_by_id(user_df, "22") or "Glasses gendre"
    face_col = get_col_by_id(user_df, "94") or "Glasses for your face shape"
    no_order_col = get_col_by_id(user_df, "103") or "Glasses lenses no-orders"
    features_col = get_col_by_id(user_df, "104") or "Glasses other features"
    model_col = get_col_by_id(user_df, "12") or "Model" 

    target_cols = [hs_col, desc_col, prod_col, usable_col, coll_col, uv_col, gender_col, face_col, no_order_col, features_col, model_col]
    for col in target_cols:
        if col not in user_df.columns: user_df[col] = ""

    # Apply Rules
    calc_hs = user_df.apply(lambda row: apply_hs_code(row, type_col, material_col, sport_col), axis=1)
    user_df[hs_col], _ = apply_safe_fill(user_df, hs_col, calc_hs)

    calc_desc = user_df.apply(lambda row: apply_item_description(row, type_col, material_col), axis=1)
    user_df[desc_col], _ = apply_safe_fill(user_df, desc_col, calc_desc)
    
    calc_prod = user_df.apply(lambda row: apply_producing_company(row, brand_col), axis=1)
    user_df[prod_col], _ = apply_safe_fill(user_df, prod_col, calc_prod)
    
    calc_usable = user_df.apply(lambda row: apply_glasses_usable(row, brand_col, type_col, effect_col, sport_col), axis=1)
    user_df[usable_col], _ = apply_safe_fill(user_df, usable_col, calc_usable)

    calc_coll = user_df.apply(lambda row: apply_glasses_collection(row, brand_col), axis=1)
    user_df[coll_col], _ = apply_safe_fill(user_df, coll_col, calc_coll)

    calc_uv = user_df.apply(lambda row: apply_uv_filter(row, type_col), axis=1)
    user_df[uv_col], _ = apply_safe_fill(user_df, uv_col, calc_uv)
    
    # Rule 7 (Gender) - Special Case, bypasses safe_fill
    gender_results = user_df.apply(lambda row: apply_glasses_gender(row, gender_col, color_col, shape_col, tl_col, lw_col), axis=1)
    user_df[gender_col] = [r[0] for r in gender_results]
    gender_reasons = [r[1] for r in gender_results]
    
    # Rule 8 (Face) - Uses safe_fill
    calc_face = user_df.apply(lambda row: apply_face_shape(row, shape_col), axis=1)
    user_df[face_col], face_reasons = apply_safe_fill(user_df, face_col, calc_face)
    
    # Rule 9 (No Order) - Uses safe_fill
    calc_no_order = user_df.apply(lambda row: apply_lenses_no_order(row, frame_type_col, contain_col), axis=1)
    user_df[no_order_col], no_order_reasons = apply_safe_fill(user_df, no_order_col, calc_no_order)
    
    # Rule 10 (Features) - Uses safe_fill
    calc_features = user_df.apply(lambda row: apply_other_features(row, type_col, contain_col, rx_col), axis=1)
    user_df[features_col], features_reasons = apply_safe_fill(user_df, features_col, calc_features)

    # Rule 11 (Model) - Uses safe_fill
    calc_model = user_df.apply(lambda row: apply_glasses_model(row, name_col, brand_col), axis=1)
    user_df[model_col], model_reasons = apply_safe_fill(user_df, model_col, calc_model)

    # Report Generation (Includes ALL logic columns now)
    report_df = pd.DataFrame({
        'Gender (Y)': user_df[gender_col],
        'Gender Logic': gender_reasons,
        'Face (AX)': user_df[face_col],
        'Face Logic': face_reasons,
        'Features (AD)': user_df[features_col],
        'Features Logic': features_reasons,
        'Model Name': user_df[model_col],
        'Model Logic': model_reasons
    })
    
    # Filter to show rows where ANY rule applied a change
    modified_rows = report_df[
        (report_df['Gender Logic'] != "") |
        (report_df['Face Logic'] != "") |
        (report_df['Features Logic'] != "") |
        (report_df['Model Logic'] != "")
    ]
    return user_df, modified_rows

# ==========================================
# 📤 USER INTERFACE
# ==========================================
st.divider()
st.subheader("1. Upload Partial Data")
uploaded_file = st.file_uploader("Choose Excel File", type=['xlsx'])

if uploaded_file:
    user_df = pd.read_excel(uploaded_file, dtype=str)
    st.write(f"Loaded {len(user_df)} rows.")

    st.divider()
    st.subheader("2. Run Auto-Fill")
    
    # BUTTON ACTION
    if st.button("✨ Auto-Fill Data", type="primary"):
        with st.spinner("Applying Rules 1-11..."):
            working_df = user_df.copy()
            filled_df, report = run_auto_fill(working_df)
            st.success(f"✅ Rules Applied!")

            # SHOW REPORT
            with st.expander("📊 View Processing Report", expanded=True):
                if not report.empty:
                    st.dataframe(report, use_container_width=True)
                else:
                    st.info("No rows matched the current rules.")
            
            # DOWNLOAD
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                filled_df.to_excel(writer, index=False)
            buffer.seek(0)
            
            st.download_button(
                label="📥 Download Updated Excel",
                data=buffer,
                file_name="filled_glasses_data.xlsx",
                mime="application/vnd.ms-excel"
            )