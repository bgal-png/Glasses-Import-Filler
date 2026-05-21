import streamlit as st
import pandas as pd
import re
import base64
import zipfile
import os
from io import BytesIO
from sqlalchemy import create_engine
from dictionaries import (
    TARGET_MAPPING,
    FACE_SHAPE_MAP,
    BRAND_USABLE_MAP,
    PREMIUM_KERING_BRANDS,
    estimate_filter_category,
)

# ==========================================
# 👓 SHAPE RECOGNITION VIA CLAUDE VISION
# ==========================================
SHAPE_CATEGORIES = [
    "Panthos / Tea cup", "Browline", "Cat Eye", "Oval / Elipse",
    "Butterfly", "Extravagant", "Single lens", "Square",
    "Oversize", "Hexagonal", "Pilot", "Rectangular", "Round",
]

def classify_glasses(image_bytes: bytes, api_key: str) -> dict:
    """Classify shape and sport type from a single image. Returns dict with 'shape' and 'is_sport'."""
    try:
        import anthropic
        client = anthropic.Anthropic(api_key=api_key)

        img_b64 = base64.b64encode(image_bytes).decode("utf-8")
        ext_check = image_bytes[:8]
        media_type = "image/png" if ext_check[:4] == b'\x89PNG' else "image/jpeg"

        response = client.messages.create(
            model="claude-haiku-4-20250414",
            max_tokens=80,
            messages=[{
                "role": "user",
                "content": [
                    {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": img_b64}},
                    {"type": "text", "text": (
                        "Analyze this eyewear image and answer TWO questions:\n\n"
                        "1. SHAPE: Classify into exactly ONE of these categories:\n"
                        + ", ".join(SHAPE_CATEGORIES) + "\n\n"
                        "2. SPORT: Are these sport/performance glasses (wrap-around, shield lens, "
                        "rubber grips, aerodynamic design, cycling/running/ski goggles)? Answer YES or NO.\n\n"
                        "Respond in exactly this format:\nSHAPE: <category>\nSPORT: <YES or NO>"
                    )},
                ],
            }],
        )
        result = response.content[0].text.strip()

        shape = ""
        is_sport = False
        for line in result.split("\n"):
            line = line.strip()
            if line.upper().startswith("SHAPE:"):
                raw_shape = line.split(":", 1)[1].strip()
                for cat in SHAPE_CATEGORIES:
                    if cat.lower() == raw_shape.lower():
                        shape = cat
                        break
                if not shape:
                    shape = raw_shape
            elif line.upper().startswith("SPORT:"):
                is_sport = "yes" in line.lower()

        return {"shape": shape, "is_sport": is_sport}
    except Exception:
        return {"shape": "", "is_sport": False}

def extract_images_from_zip(zip_file) -> dict:
    images = {}
    with zipfile.ZipFile(zip_file, "r") as z:
        for name in z.namelist():
            lower = name.lower()
            if lower.endswith((".jpg", ".jpeg", ".png")) and not name.startswith("__MACOSX"):
                basename = os.path.splitext(os.path.basename(name))[0]
                if basename:
                    images[basename] = z.read(name)
    return images

# ==========================================
# 🛑 VERSION CHECK & CONFIG
# ==========================================
st.set_page_config(page_title="Glasses Import Filler", layout="wide", page_icon="🕶️")
APP_VERSION = "v.Cloud.1.0"

st.markdown("""
<style>
    /* Header styling */
    .main-header {
        background: linear-gradient(135deg, #1A1F2E 0%, #0E1117 100%);
        padding: 1.5rem 2rem;
        border-radius: 12px;
        border: 1px solid #2A3040;
        margin-bottom: 1rem;
    }
    .main-header h1 {
        color: #4A9EFF;
        font-size: 1.8rem;
        margin: 0;
    }
    .main-header p {
        color: #8892A0;
        font-size: 0.85rem;
        margin: 0.3rem 0 0 0;
    }
    /* Card styling for sections */
    .stExpander {
        border: 1px solid #2A3040 !important;
        border-radius: 8px !important;
    }
    /* Button styling */
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #4A9EFF, #2563EB);
        border: none;
        border-radius: 8px;
        font-weight: 600;
    }
    /* Sidebar styling */
    section[data-testid="stSidebar"] {
        border-right: 1px solid #2A3040;
    }
    /* Divider */
    hr {
        border-color: #2A3040 !important;
    }
    /* Metric cards */
    [data-testid="stMetric"] {
        background: #1A1F2E;
        padding: 1rem;
        border-radius: 8px;
        border: 1px solid #2A3040;
    }
    /* Download button */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #10B981, #059669) !important;
        border: none !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
    }
    /* Tabs */
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px 8px 0 0;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
    <h1>Glasses Import Filler</h1>
    <p>Cloud Edition &bull; {}</p>
</div>
""".format(APP_VERSION), unsafe_allow_html=True)

# ==========================================
# 🔐 SECURE CLOUD CONNECTION
# ==========================================
try:
    DB_URL = st.secrets["DB_URL"]
except KeyError:
    st.error("❌ CRITICAL: 'DB_URL' secret is missing. Please add it to your Streamlit Cloud secrets or local .streamlit/secrets.toml file.")
    st.stop()

# ==========================================
# ☁️ CLOUD DATABASE CONNECTION
# ==========================================
@st.cache_data(show_spinner=False, ttl=3600)
def load_cloud_data():
    master_db = pd.DataFrame()
    package_df = pd.DataFrame()
    historical_df = pd.DataFrame()
    origin_df = pd.DataFrame()

    try:
        engine = create_engine(DB_URL, pool_pre_ping=True, pool_recycle=300)

        # 1. Fetch Master Catalog
        master_db = pd.read_sql_table('master_catalog', con=engine)
        if 'join_key' in master_db.columns:
            master_db.set_index('join_key', inplace=True)

        # 2. Fetch Package Data
        try:
            package_df = pd.read_sql_table('package_data', con=engine)
        except:
            pass

        # 3. Fetch Global Categories (master_clean)
        try:
            historical_df = pd.read_sql_table('historical_data', con=engine)
        except:
            pass

        # 4. Fetch Item Origin
        try:
            origin_df = pd.read_sql_table('origin_data', con=engine)
        except:
            pass

        return master_db, package_df, historical_df, origin_df
    except Exception as e:
        st.error(f"❌ Failed to connect to Cloud Database: {e}")
        return master_db, package_df, historical_df, origin_df

with st.spinner("☁️ Fetching live data from Supabase Vault..."):
    master_db, package_df, master_clean_df, origin_df = load_cloud_data()

if master_db.empty:
    st.warning("⚠️ Database is empty. Please run your 'admin_updater.py' script first.")
    st.stop()

# --- 📊 DATA STATUS METRICS ---
st.divider()

col_a, col_b, col_c, col_d = st.columns(4)
with col_a:
    st.metric("📦 Package Data", f"{len(package_df)} items" if not package_df.empty else "Not loaded")
with col_b:
    st.metric("📜 Global Categories", f"{len(master_clean_df)} glasses" if not master_clean_df.empty else "Not loaded")
with col_c:
    st.metric("🗄️ Master Catalog", f"{len(master_db)} products")
with col_d:
    st.metric("🌍 Item Origin", f"{len(origin_df)} items" if not origin_df.empty else "Not loaded")

# ==========================================
# 🚀 APP UI & CONTROL PANEL
# ==========================================
st.sidebar.markdown("### ⚙️ Control Panel")

if st.sidebar.button("🔄 Sync Fresh Data", type="primary", use_container_width=True):
    st.cache_data.clear()
    st.rerun()

st.sidebar.divider()
st.sidebar.markdown("### 🏷️ Private Name Numbers")
priv_sun = st.sidebar.text_input("Sunglasses", placeholder="e.g. 1001")
priv_eye = st.sidebar.text_input("Eyeglasses (Frames)", placeholder="e.g. 2001")
priv_pc = st.sidebar.text_input("PC Glasses", placeholder="e.g. 3001")
priv_sport = st.sidebar.text_input("Sport Glasses", placeholder="e.g. 4001")
priv_drive = st.sidebar.text_input("Driving Glasses", placeholder="e.g. 5001")

# ==========================================
# 📥 THE AUTO-FILLER ENGINE
# ==========================================
st.divider()
st.markdown("### 📥 Step 1: Upload Your Target File")

uploaded_file = st.file_uploader("Upload your Target Excel or CSV file", type=["xlsx", "csv"])

if uploaded_file:
    try:
        if uploaded_file.name.endswith(".csv"):
            target_df = pd.read_csv(uploaded_file, dtype=str)
        else:
            target_df = pd.read_excel(uploaded_file, dtype=str, engine="openpyxl")

        target_df.columns = (
            target_df.columns.astype(str)
            .str.replace("\n", " ", regex=False)
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
        )

        # Save a copy of the original data for before/after comparison
        original_df = target_df.copy()

    except Exception as e:
        st.error(f"Could not read your uploaded file: {e}")
        st.stop()

    target_barcode_col = TARGET_MAPPING.get("Barcode", "Barcode")
    if target_barcode_col not in target_df.columns:
        st.error(f"❌ Could not find the Barcode column '{target_barcode_col}' in your file.")
        st.stop()

    st.success(f"File uploaded! Contains {len(target_df)} rows. Click below to start matching.")

    # ==========================================
    # 👓 STEP 2: OPTIONAL IMAGE UPLOAD FOR SHAPES
    # ==========================================
    image_dict = {}
    has_api_key = False
    try:
        ANTHROPIC_API_KEY = st.secrets["ANTHROPIC_API_KEY"]
        has_api_key = True
    except KeyError:
        pass

    if has_api_key:
        st.divider()
        with st.expander("👓 Step 2: Upload Product Images for Shape Recognition (Optional)", expanded=False):
            st.caption("Upload a ZIP file with product images. Filenames must match the 'Glasses name' column exactly (e.g. `Ray-Ban RB3025 001/58 62.zip` containing `Ray-Ban RB3025 001/58 62.jpg`).")

            uploaded_zip = st.file_uploader("Upload ZIP with product images", type=["zip"], key="shape_images")
            if uploaded_zip:
                try:
                    image_dict = extract_images_from_zip(uploaded_zip)
                    st.success(f"📸 Extracted {len(image_dict)} images from ZIP.")
                except Exception as e:
                    st.error(f"Failed to read ZIP file: {e}")

    if st.button("🚀 Run Auto-Filler", type="primary"):
        with st.spinner("Matching barcodes and pouring data from the Cloud..."):

            for global_col, target_col in TARGET_MAPPING.items():
                if isinstance(target_col, list):
                    for tc in target_col:
                        if tc not in target_df.columns: target_df[tc] = ""
                else:
                    if target_col not in target_df.columns: target_df[target_col] = ""

            match_count = 0
            found_sport_glasses = False
            found_polarized_clip_on = False

            # --- VALIDATION TRACKING ---
            unmapped_tracker = {}   # col -> set of unmapped source values
            missing_tracker = {}    # col -> count of rows with no source data

            # --- CACHES FOR MAJORITY ENGINES ---
            brand_majority_cache = {}
            brand_contain_cache = {}
            brand_origin_cache = {}

            # Accept both "Glasses contain ID:84" (no space) and "Glasses contain ID: 84"
            # (with space) in templates — older templates used the space form, newer ones
            # don't. Whichever exists in the user's template is what we write to.
            if "Glasses contain ID:84" in target_df.columns:
                CONTAIN_COL = "Glasses contain ID:84"
            elif "Glasses contain ID: 84" in target_df.columns:
                CONTAIN_COL = "Glasses contain ID: 84"
            else:
                CONTAIN_COL = "Glasses contain ID:84"  # default for brand-new templates

            for c in ["Case length (mm)", "Case height (mm)", "Case width (mm)", "Case weight (g)", CONTAIN_COL]:
                if c not in target_df.columns: target_df[c] = ""

            for index, row in target_df.iterrows():
                raw_barcode = str(row[target_barcode_col]).strip()
                clean_barcode = re.sub(r"\.0$", "", raw_barcode).lstrip("0")

                if clean_barcode in master_db.index:
                    match_count += 1
                    master_row = master_db.loc[clean_barcode]

                    is_frames = str(master_row.get("Glasses_type", "")).strip() == "Frames"
                    lens_cols_to_skip = [
                        "Glasses_lens_Colour", "Glasses_lens_material",
                        "Sunglasses_filter", "Glasses_lens_effect",
                        "SunGlasses_RX_lenses",  # not applicable to optical frames
                    ]
                    
                    target_df.at[index, "Items type ID: 20"] = "Glasses"
                    target_df.at[index, "Items packing ID: 21"] = "Basic"

                    g_type = str(master_row.get("Glasses_type", "")).strip()
                    private_name = ""

                    if "Sunglasses" in g_type and priv_sun: private_name = f"(Sunglasses {priv_sun})"
                    elif "Sport glasses" in g_type and priv_sport: private_name = f"(Sports glasses {priv_sport})"
                    elif "Driving glasses" in g_type and priv_drive: private_name = f"(Eyeglasses driving {priv_drive})"
                    elif "PC Glasses without power" in g_type and priv_pc: private_name = f"(Eyeglasses PC {priv_pc})"
                    elif "Frames" in g_type and priv_eye: private_name = f"(Eyeglasses {priv_eye})"

                    if private_name: target_df.at[index, "Name private"] = private_name.strip()

                    assembled_name = str(master_row.get("Assembled_Name", "")).strip()
                    meta_desc = ""

                    if "Sunglasses" in g_type: meta_desc = f"Sunglasses {assembled_name}"
                    elif "Sport glasses" in g_type: 
                        meta_desc = f"Ski goggles {assembled_name}"
                        found_sport_glasses = True
                    elif "Driving glasses" in g_type: meta_desc = f"Driving glasses {assembled_name}"
                    elif "PC Glasses without power" in g_type: meta_desc = f"Computer glasses {assembled_name}"
                    elif "Frames" in g_type: meta_desc = f"Eyeglasses {assembled_name}"

                    if meta_desc: target_df.at[index, "Meta description"] = meta_desc.strip()

                    for global_col, target_col in TARGET_MAPPING.items():
                        if global_col == "Barcode": continue
                        if is_frames and global_col in lens_cols_to_skip: continue

                        if global_col in master_db.columns:
                            val = master_row[global_col]
                            t_col_name = target_col[0] if isinstance(target_col, list) else target_col
                            if pd.notna(val) and str(val).strip() != "":
                                val_str = str(val).strip()
                                # Filter out NOT MAPPED values (handles pipe-separated)
                                unmapped_parts = []
                                if "|" in val_str:
                                    all_parts = [p.strip() for p in val_str.split("|") if p.strip()]
                                    clean_parts = [p for p in all_parts if p != "NOT MAPPED"]
                                    unmapped_parts = [p for p in all_parts if p == "NOT MAPPED"]
                                    val_str = "|".join(clean_parts)
                                elif val_str == "NOT MAPPED":
                                    unmapped_parts = [val_str]
                                    val_str = ""
                                if unmapped_parts:
                                    if t_col_name not in unmapped_tracker: unmapped_tracker[t_col_name] = set()
                                    raw_val = str(val).strip()
                                    unmapped_tracker[t_col_name].add(raw_val)
                                if val_str:
                                    if isinstance(target_col, list):
                                        for tc in target_col: target_df.at[index, tc] = val_str
                                    else: target_df.at[index, target_col] = val_str
                                elif not val_str:
                                    missing_tracker[t_col_name] = missing_tracker.get(t_col_name, 0) + 1
                            else:
                                missing_tracker[t_col_name] = missing_tracker.get(t_col_name, 0) + 1

                    g_shape_raw = str(master_row.get("Glasses_shape", "")).strip()
                    if g_shape_raw and g_shape_raw.lower() not in ["nan", ""]:
                        shapes = [s.strip() for s in g_shape_raw.split("|")]
                        recommended_faces = set()
                        for s in shapes:
                            for shape_key, face_val in FACE_SHAPE_MAP.items():
                                if shape_key.lower() == s.lower():
                                    for face in face_val.split("|"): recommended_faces.add(face)
                        if recommended_faces:
                            target_df.at[index, "Glasses for your face shape ID:94"] = "|".join(sorted(list(recommended_faces)))

                    if "Sunglasses" in g_type: target_df.at[index, "UV filter ID: 60"] = "400"

                    # --- 🕶️ SUNGLASSES FILTER ESTIMATION (from lens color) ---
                    filter_col = "Sunglasses filter ID: 77"
                    if filter_col in target_df.columns and "Sunglasses" in g_type:
                        current_filter = str(target_df.at[index, filter_col]).strip()
                        if not current_filter or current_filter.lower() in ["nan", ""]:
                            raw_lens = str(master_row.get("Glasses_lens_Colour", "")).strip()
                            estimated = estimate_filter_category(raw_lens)
                            if estimated:
                                target_df.at[index, filter_col] = estimated

                    usable_tags = set()
                    raw_brand = str(master_row.get("Brand", "")).strip().lower()
                    if raw_brand in BRAND_USABLE_MAP: usable_tags.add(BRAND_USABLE_MAP[raw_brand])
                    lens_effect = str(master_row.get("Glasses_lens_effect", "")).strip()

                    if "Sunglasses" in g_type:
                        if "Polarized" in lens_effect: usable_tags.add("Driving glasses")
                        else: usable_tags.add("Common use")

                    if usable_tags: target_df.at[index, "Glasses usable ID: 51"] = "|".join(sorted(list(usable_tags)))

                    if raw_brand in PREMIUM_KERING_BRANDS: target_df.at[index, "Glasses collection ID: 33"] = "Prémiové brýle - Kering"

                    raw_material = str(master_row.get("Glasses_main_material", "")).strip().lower()
                    if "Sunglasses" in g_type: target_df.at[index, "HS Code"] = "90041091"
                    elif "Sport glasses" in g_type: target_df.at[index, "HS Code"] = "90049090"
                    elif "Frames" in g_type:
                        if "plastic" in raw_material: target_df.at[index, "HS Code"] = "90031100"
                        elif "metal" in raw_material: target_df.at[index, "HS Code"] = "90031900"

                    if "Frames" in g_type: target_df.at[index, "Item description"] = "Eyeglasses"
                    elif "PC Glasses without power" in g_type: target_df.at[index, "Item description"] = "PC Glasses without power"
                    elif "Driving glasses" in g_type: target_df.at[index, "Item description"] = "Driving glasses"
                    elif "Sunglasses" in g_type:
                        has_plastic = "plastic" in raw_material
                        has_metal = "metal" in raw_material
                        if has_plastic and has_metal: target_df.at[index, "Item description"] = "Sunglasses, mixed plastic and metal frame"
                        elif has_plastic: target_df.at[index, "Item description"] = "Sunglasses, plastic frame"
                        elif has_metal: target_df.at[index, "Item description"] = "Sunglasses, metal frame"

                    # --- 🧳 CASE DIMENSIONS MAJORITY ENGINE ---
                    if not package_df.empty and raw_brand and raw_brand != "nan":
                        if raw_brand not in brand_majority_cache:
                            mask = package_df['item_name'].astype(str).str.contains(rf'\b{re.escape(raw_brand)}\b', case=False, na=False)
                            brand_matches = package_df[mask]
                            
                            if not brand_matches.empty:
                                def get_mode(col_name):
                                    if col_name in brand_matches.columns:
                                        modes = brand_matches[col_name].dropna().mode()
                                        if not modes.empty: return re.sub(r'\.0$', '', str(modes.iloc[0]).strip())
                                    return ""
                                
                                brand_majority_cache[raw_brand] = {
                                    "Case length (mm)": get_mode("case_length"),
                                    "Case height (mm)": get_mode("case_height"),
                                    "Case width (mm)": get_mode("case_width"),
                                    "Case weight (g)": get_mode("case_weight"),
                                    "Glasses weight (g)": get_mode("item_weight"),
                                }
                            else: brand_majority_cache[raw_brand] = None
                                
                        cached_data = brand_majority_cache.get(raw_brand)
                        if cached_data:
                            target_df.at[index, "Case length (mm)"] = cached_data["Case length (mm)"]
                            target_df.at[index, "Case height (mm)"] = cached_data["Case height (mm)"]
                            target_df.at[index, "Case width (mm)"] = cached_data["Case width (mm)"]
                            target_df.at[index, "Case weight (g)"] = cached_data["Case weight (g)"]
                            if cached_data.get("Glasses weight (g)"):
                                target_df.at[index, "Glasses weight (g)"] = cached_data["Glasses weight (g)"]

                    # --- 🌍 ORIGIN COUNTRY MAJORITY ENGINE ---
                    if not origin_df.empty and raw_brand and raw_brand != "nan":
                        if raw_brand not in brand_origin_cache:
                            if "item_name" in origin_df.columns and "country_master" in origin_df.columns:
                                mask = origin_df['item_name'].astype(str).str.contains(rf'\b{re.escape(raw_brand)}\b', case=False, na=False)
                                brand_matches = origin_df[mask]
                                if not brand_matches.empty:
                                    modes = brand_matches["country_master"].dropna().mode()
                                    brand_origin_cache[raw_brand] = str(modes.iloc[0]).strip() if not modes.empty else ""
                                else:
                                    brand_origin_cache[raw_brand] = ""
                            else:
                                brand_origin_cache[raw_brand] = ""

                        cached_origin = brand_origin_cache.get(raw_brand, "")
                        if cached_origin and "Item origin country" in target_df.columns:
                            target_df.at[index, "Item origin country"] = cached_origin

                    # --- 🎁 GLASSES CONTAIN MAJORITY ENGINE (WITH RAW CLIP-ONS) ---
                    if not master_clean_df.empty and raw_brand and raw_brand != "nan":
                        if raw_brand not in brand_contain_cache:
                            if "Brand" in master_clean_df.columns and "Glasses contain" in master_clean_df.columns:
                                brand_mask = master_clean_df['Brand'].astype(str).str.strip().str.lower() == raw_brand
                                brand_matches = master_clean_df[brand_mask]
                                
                                if not brand_matches.empty:
                                    modes = brand_matches["Glasses contain"].dropna().mode()
                                    if not modes.empty:
                                        raw_contain = str(modes.iloc[0]).strip()
                                        parts = [p.strip() for p in raw_contain.split(',') if p.strip() and p.strip().lower() != 'nan']
                                        allowed_historical = {"original glasses case", "cleaning cloth"}
                                        filtered_parts = [p for p in parts if p.lower() in allowed_historical]
                                        brand_contain_cache[raw_brand] = "|".join(filtered_parts)
                                    else: brand_contain_cache[raw_brand] = ""
                                else: brand_contain_cache[raw_brand] = ""
                            else: brand_contain_cache[raw_brand] = ""
                                
                        cached_contain = brand_contain_cache.get(raw_brand, "")
                        
                        clip_on_val = str(master_row.get("Extracted_Clip_on", "")).strip()
                        needs_alert = master_row.get("Clip_on_Alert", False)

                        if needs_alert: found_polarized_clip_on = True

                        # Clip-on lens colour
                        clip_lens_col = "Glasses clip-on lens colour ID:112"
                        if clip_lens_col in target_df.columns:
                            clip_lens_val = str(master_row.get("Clip_on_lens_colour", "")).strip()
                            if clip_lens_val and clip_lens_val.lower() not in ["nan", ""]:
                                target_df.at[index, clip_lens_col] = clip_lens_val
                            
                        final_contain = []
                        if cached_contain: final_contain.extend(cached_contain.split("|"))
                        if clip_on_val and clip_on_val.lower() not in ["nan", ""]: final_contain.append(clip_on_val)
                            
                        if final_contain:
                            unique_contain_dict = {item.strip().lower(): item.strip() for item in final_contain if item.strip()}
                            ordered_items = []
                            if "original glasses case" in unique_contain_dict:
                                ordered_items.append("Original glasses case")
                                del unique_contain_dict["original glasses case"]
                            if "cleaning cloth" in unique_contain_dict:
                                ordered_items.append("Cleaning cloth")
                                del unique_contain_dict["cleaning cloth"]
                            
                            remaining_items = sorted(list(unique_contain_dict.values()))
                            ordered_items.extend(remaining_items)
                            target_df.at[index, CONTAIN_COL] = "|".join(ordered_items)

                    # --- 🌟 OTHER FEATURES ENGINE ---
                    other_features = set()
                    if "Glasses other features ID:99" in target_df.columns:
                        existing_features = str(target_df.at[index, "Glasses other features ID:99"]).strip()
                        if existing_features and existing_features.lower() not in ["nan", ""]:
                            for e in existing_features.split("|"): other_features.add(e.strip())

                    # "Prescription sunglasses" only makes sense for actual sunglasses —
                    # Safilo flags RXable=Y on optical frames too (any frame can be glazed),
                    # so without this check the feature gets added to every row.
                    if "Sunglasses" in g_type and "SunGlasses RX lenses ID:108" in target_df.columns:
                        if str(target_df.at[index, "SunGlasses RX lenses ID:108"]).strip().lower() == "yes":
                            other_features.add("Prescription sunglasses")

                    if CONTAIN_COL in target_df.columns:
                        contain_val = str(target_df.at[index, CONTAIN_COL]).strip().lower()
                        contain_items = [item.strip() for item in re.split(r"[,|]", contain_val) if item.strip()]

                        clip_on_found = False
                        if "sun clip-on" in contain_items: other_features.add("Sun clip-on"); clip_on_found = True
                        if "sun clip-on p" in contain_items: other_features.add("Sun clip-on p"); clip_on_found = True
                        if "magnetic sun clip-on" in contain_items: other_features.add("Magnetic sun clip-on"); clip_on_found = True
                        if "magnetic sun clip-on p" in contain_items: other_features.add("Magnetic sun clip-on p"); clip_on_found = True
                        if clip_on_found: other_features.add("Glasses with sun clip-on")

                    if other_features:
                        target_df.at[index, "Glasses other features ID:99"] = "|".join(sorted(list(other_features)))

                    # --- 🚫 LENSES NO-ORDERS ENGINE ---
                    no_orders = set()
                    frame_type = str(target_df.at[index, "Glasses frame type ID: 50"]).strip().lower() if "Glasses frame type ID: 50" in target_df.columns else ""
                    other_feat = str(target_df.at[index, "Glasses other features ID:99"]).strip().lower() if "Glasses other features ID:99" in target_df.columns else ""

                    if frame_type == "half rim":
                        no_orders.add("CoatingPolarized")
                        no_orders.add("Glasses index 1.5")
                    elif frame_type == "rimless":
                        no_orders.add("CoatingPolarized")
                        no_orders.add("Glasses index 1.5")
                        no_orders.add("Glasses index 1.74")

                    if "clip" in other_feat:
                        no_orders.add("Glasses index 1.5")

                    if no_orders and "Glasses lenses no-orders ID:103" in target_df.columns:
                        target_df.at[index, "Glasses lenses no-orders ID:103"] = "|".join(sorted(no_orders))

                    # --- Fill "None" for empty lens effect ---
                    effect_col = "Glasses lens effect ID: 37"
                    if effect_col in target_df.columns and not is_frames:
                        current_effect = str(target_df.at[index, effect_col]).strip()
                        if not current_effect or current_effect.lower() in ["nan", ""]:
                            target_df.at[index, effect_col] = "None"

            st.success(f"✅ Match Complete! Successfully filled {match_count} out of {len(target_df)} products.")

            # --- 👓 AI VISION ENGINE (Shape + Sport Detection) ---
            if image_dict and has_api_key:
                shape_col = "Glasses shape ID: 25"
                face_col = "Glasses for your face shape ID:94"
                sport_col = "Sports Glasses ID: 89"
                source_col = "Shape source"
                if shape_col not in target_df.columns: target_df[shape_col] = ""
                if face_col not in target_df.columns: target_df[face_col] = ""
                if sport_col not in target_df.columns: target_df[sport_col] = ""
                target_df[source_col] = ""

                # Mark existing shapes from database
                for idx, row in target_df.iterrows():
                    if str(row.get(shape_col, "")).strip() not in ["", "nan"]:
                        target_df.at[idx, source_col] = "Database"

                name_col = "Glasses name"
                if name_col not in target_df.columns:
                    for c in target_df.columns:
                        if "name" in c.lower() and "private" not in c.lower():
                            name_col = c
                            break

                shape_count = 0
                sport_count = 0
                shape_bar = st.progress(0, text="🔍 Classifying with AI vision...")
                total_rows = len(target_df)

                for idx, row in target_df.iterrows():
                    glasses_name = str(row.get(name_col, "")).strip()
                    if glasses_name and glasses_name in image_dict:
                        result = classify_glasses(image_dict[glasses_name], ANTHROPIC_API_KEY)

                        if result["shape"]:
                            target_df.at[idx, shape_col] = result["shape"]
                            target_df.at[idx, source_col] = "AI"
                            shape_count += 1

                            # Update face shape recommendation
                            recommended_faces = set()
                            for shape_key, face_val in FACE_SHAPE_MAP.items():
                                if shape_key.lower() == result["shape"].lower():
                                    for face in face_val.split("|"):
                                        recommended_faces.add(face)
                            if recommended_faces:
                                target_df.at[idx, face_col] = "|".join(sorted(recommended_faces))

                        if result["is_sport"]:
                            target_df.at[idx, sport_col] = "Yes"
                            sport_count += 1

                    progress = (list(target_df.index).index(idx) + 1) / total_rows
                    shape_bar.progress(progress, text=f"🔍 Classifying... ({list(target_df.index).index(idx) + 1}/{total_rows})")

                shape_bar.empty()
                st.success(f"👓 AI Vision Complete! Shapes: {shape_count}, Sport glasses: {sport_count} (out of {len(image_dict)} images)")

            # --- VALIDATION REPORT ---
            if unmapped_tracker or missing_tracker:
                total_issues = sum(missing_tracker.values()) + sum(len(v) for v in unmapped_tracker.values())
                st.warning(f"⚠️ **Validation Report:** {total_issues} potential issues found")
                if unmapped_tracker:
                    with st.expander(f"🔴 Unmapped values ({len(unmapped_tracker)} columns)"):
                        for col, vals in sorted(unmapped_tracker.items()):
                            st.write(f"**{col}:** {len(vals)} unmapped value(s)")
                            for v in sorted(vals):
                                st.caption(f"  → `{v}`")
                if missing_tracker:
                    with st.expander(f"🟡 Missing from source ({len(missing_tracker)} columns)"):
                        for col, count in sorted(missing_tracker.items(), key=lambda x: -x[1]):
                            st.write(f"**{col}:** {count} row(s) with no data")

            if found_sport_glasses:
                st.warning("⚠️ **Heads Up:** We found 'Sport glasses' in this batch and labeled them as 'Ski goggles' in the Meta Description. Double check them!")
                
            if found_polarized_clip_on:
                st.warning("⚠️ **Polarized Clip-On Alert:** We found a Marcolin/Kering clip-on that is marked as polarized, but it was assigned standard 'Sun clip-on' or 'Magnetic sun clip-on'. Verify if it needs the ' p' suffix!")

            st.write("### 📊 Before / After Comparison")
            preview_tab1, preview_tab2, preview_tab3 = st.tabs(["🔄 Changes Only", "📥 Original", "📤 Filled"])
            with preview_tab1:
                # Show only columns that changed
                changed_cols = []
                for col in target_df.columns:
                    if col in original_df.columns:
                        if not target_df[col].equals(original_df[col]):
                            changed_cols.append(col)
                    else:
                        changed_cols.append(col)
                if changed_cols:
                    id_col = "Glasses name" if "Glasses name" in target_df.columns else target_df.columns[0]
                    display_cols = [id_col] + [c for c in changed_cols if c != id_col]
                    st.caption(f"**{len(changed_cols)} columns** were modified by the filler")
                    st.dataframe(target_df[display_cols].head(20), use_container_width=True)
                else:
                    st.info("No changes were made.")
            with preview_tab2:
                st.dataframe(original_df.head(20), use_container_width=True)
            with preview_tab3:
                st.dataframe(target_df.head(20), use_container_width=True)

            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                target_df.to_excel(writer, index=False, sheet_name="Filled_Data")
            processed_data = output.getvalue()

            st.download_button(
                label="📥 Download Filled Excel File",
                data=processed_data,
                file_name="Master_Filled_Glasses.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )