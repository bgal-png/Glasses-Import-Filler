import streamlit as st
import pandas as pd
import re
from io import BytesIO
from sqlalchemy import create_engine
from dictionaries import (
    TARGET_MAPPING,
    FACE_SHAPE_MAP,
    BRAND_USABLE_MAP,
    PREMIUM_KERING_BRANDS,
)

# ==========================================
# 🛑 VERSION CHECK & CONFIG
# ==========================================
st.set_page_config(page_title="Manufacturer Data Linker", layout="wide")
APP_VERSION = "v.Cloud.1.0"

st.title(f"🏭 Manufacturer Data Filler (Cloud Edition)")
st.caption(f"🚀 Running Code Version: **{APP_VERSION}**")

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
    
    try:
        # 🔥 THE BULLETPROOF ENGINE 🔥
        # pool_pre_ping=True checks if the database is awake before pulling data!
        engine = create_engine(DB_URL, pool_pre_ping=True, pool_recycle=300)
        
        
        # 1. Fetch Master Catalog
        master_db = pd.read_sql_table('master_catalog', con=engine)
        if 'join_key' in master_db.columns:
            master_db.set_index('join_key', inplace=True)
            
        # 2. Fetch Package Data
        try:
            package_df = pd.read_sql_table('package_data', con=engine)
        except:
            pass # Table might not exist if they skipped it
            
        # 3. Fetch Historical Data (master_clean)
        try:
            historical_df = pd.read_sql_table('historical_data', con=engine)
        except:
            pass

        return master_db, package_df, historical_df
    except Exception as e:
        st.error(f"❌ Failed to connect to Cloud Database: {e}")
        return master_db, package_df, historical_df

with st.spinner("☁️ Fetching live data from Supabase Vault..."):
    master_db, package_df, master_clean_df = load_cloud_data()

if master_db.empty:
    st.warning("⚠️ Database is empty. Please run your 'admin_updater.py' script first.")
    st.stop()

# --- 📦 COMPACT BACKGROUND DATA STATUS ---
st.divider()
pkg_status = f"✅ **Package Data:** Loaded ({len(package_df)} items)" if not package_df.empty else "⚠️ **Package Data:** Not found in cloud."
mc_status = f"✅ **Historical Data:** Loaded ({len(master_clean_df)} glasses)" if not master_clean_df.empty else "⚠️ **Historical Data:** Not found in cloud."

st.markdown("### 📂 Cloud Data Status")
col_a, col_b = st.columns(2)
with col_a:
    st.markdown(pkg_status)
with col_b:
    st.markdown(mc_status)

# ==========================================
# 🚀 APP UI & CONTROL PANEL
# ==========================================
st.sidebar.header("Control Panel")

if st.sidebar.button("🗑️ Sync Fresh Data from Cloud", type="primary"):
    st.cache_data.clear()
    st.rerun()

st.sidebar.divider()
st.sidebar.subheader("🏷️ Private Name Numbers")
priv_sun = st.sidebar.text_input("Sunglasses", placeholder="e.g. 1001")
priv_eye = st.sidebar.text_input("Eyeglasses (Frames)", placeholder="e.g. 2001")
priv_pc = st.sidebar.text_input("PC Glasses", placeholder="e.g. 3001")
priv_sport = st.sidebar.text_input("Sport Glasses", placeholder="e.g. 4001")
priv_drive = st.sidebar.text_input("Driving Glasses", placeholder="e.g. 5001")

# ==========================================
# 🔍 QUICK EAN LOOKUP UTILITY
# ==========================================
st.divider()
with st.expander("🔍 Quick EAN / Barcode Lookup", expanded=False):
    col1, col2 = st.columns([3, 1])

    with col1:
        search_ean = st.text_input(
            "Enter EAN to search in vault:", placeholder="e.g. 8056597123456"
        )
    with col2:
        st.write("") 
        st.write("")
        search_btn = st.button("Search Database", use_container_width=True)

    if search_btn and search_ean:
        clean_search = re.sub(r"\.0$", "", str(search_ean).strip()).lstrip("0")

        if clean_search in master_db.index:
            st.success(f"✅ EAN '{search_ean}' found in the database!")
            found_data = master_db.loc[[clean_search]]
            display_cols = [c for c in ["Producing_company", "Brand", "Glasses_type", "Glasses_shape"] if c in found_data.columns]
            st.dataframe(found_data[display_cols], use_container_width=True)
        else:
            st.error(f"❌ EAN '{search_ean}' (Cleaned: {clean_search}) was NOT found.")

# ==========================================
# 📥 THE AUTO-FILLER ENGINE
# ==========================================
st.divider()
st.subheader("📥 Step 1: Upload Your Target File")

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

        mapped_targets = set()
        for val in TARGET_MAPPING.values():
            if isinstance(val, list):
                for item in val: mapped_targets.add(item)
            else:
                mapped_targets.add(val)

        custom_targets = {
            "Items type ID: 20", "Items packing ID: 21", "Name private", "Meta description",
            "Glasses for your face shape ID:94", "UV filter ID: 60", "Glasses usable ID: 51",
            "Glasses collection ID: 33", "HS Code", "Item description", "Glasses other features ID:99",
            "Case length (mm)", "Case height (mm)", "Case width (mm)", "Case weight (g)", "Glasses contain ID: 84"
        }

        all_completed_targets = mapped_targets.union(custom_targets)

        status_list = [f"✅ {col}" if col in all_completed_targets else f"⏳ {col}" for col in target_df.columns]

        with st.expander("🔍 PROGRESS TRACKER: Exact Bucket Names"):
            st.markdown("**✅ = Rule Applied | ⏳ = Still Needs Logic/Mapping**")
            st.write(status_list)

    except Exception as e:
        st.error(f"Could not read your uploaded file: {e}")
        st.stop()

    target_barcode_col = TARGET_MAPPING.get("Barcode", "Barcode")
    if target_barcode_col not in target_df.columns:
        st.error(f"❌ Could not find the Barcode column '{target_barcode_col}' in your file.")
        st.stop()

    st.success(f"File uploaded! Contains {len(target_df)} rows. Click below to start matching.")

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
            
            # --- CACHES FOR MAJORITY ENGINES ---
            brand_majority_cache = {}
            brand_contain_cache = {}
            
            for c in ["Case length (mm)", "Case height (mm)", "Case width (mm)", "Case weight (g)", "Glasses contain ID: 84"]:
                if c not in target_df.columns: target_df[c] = ""

            for index, row in target_df.iterrows():
                raw_barcode = str(row[target_barcode_col]).strip()
                clean_barcode = re.sub(r"\.0$", "", raw_barcode).lstrip("0")

                if clean_barcode in master_db.index:
                    match_count += 1
                    master_row = master_db.loc[clean_barcode]

                    is_frames = str(master_row.get("Glasses_type", "")).strip() == "Frames"
                    lens_cols_to_skip = ["Glasses_lens_Colour", "Glasses_lens_material", "Sunglasses_filter", "Glasses_lens_effect"]
                    
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
                            if pd.notna(val) and str(val).strip() != "":
                                if isinstance(target_col, list):
                                    for tc in target_col: target_df.at[index, tc] = val
                                else: target_df.at[index, target_col] = val

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
                                    "Case weight (g)": get_mode("case_weight")
                                }
                            else: brand_majority_cache[raw_brand] = None
                                
                        cached_data = brand_majority_cache.get(raw_brand)
                        if cached_data:
                            target_df.at[index, "Case length (mm)"] = cached_data["Case length (mm)"]
                            target_df.at[index, "Case height (mm)"] = cached_data["Case height (mm)"]
                            target_df.at[index, "Case width (mm)"] = cached_data["Case width (mm)"]
                            target_df.at[index, "Case weight (g)"] = cached_data["Case weight (g)"]

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
                            target_df.at[index, "Glasses contain ID: 84"] = "|".join(ordered_items)

                    # --- 🌟 OTHER FEATURES ENGINE ---
                    other_features = set()
                    if "Glasses other features ID:99" in target_df.columns:
                        existing_features = str(target_df.at[index, "Glasses other features ID:99"]).strip()
                        if existing_features and existing_features.lower() not in ["nan", ""]:
                            for e in existing_features.split("|"): other_features.add(e.strip())

                    if "SunGlasses RX lenses ID:108" in target_df.columns:
                        if str(target_df.at[index, "SunGlasses RX lenses ID:108"]).strip().lower() == "yes":
                            other_features.add("Prescription sunglasses")

                    if "Glasses contain ID: 84" in target_df.columns:
                        contain_val = str(target_df.at[index, "Glasses contain ID: 84"]).strip().lower()
                        contain_items = [item.strip() for item in re.split(r"[,|]", contain_val) if item.strip()]

                        clip_on_found = False
                        if "sun clip-on" in contain_items: other_features.add("Sun clip-on"); clip_on_found = True
                        if "sun clip-on p" in contain_items: other_features.add("Sun clip-on p"); clip_on_found = True
                        if "magnetic sun clip-on" in contain_items: other_features.add("Magnetic sun clip-on"); clip_on_found = True
                        if "magnetic sun clip-on p" in contain_items: other_features.add("Magnetic sun clip-on p"); clip_on_found = True
                        if clip_on_found: other_features.add("Glasses with sun clip-on")

                    if other_features:
                        target_df.at[index, "Glasses other features ID:99"] = "|".join(sorted(list(other_features)))

            st.success(f"✅ Match Complete! Successfully filled {match_count} out of {len(target_df)} products.")

            if found_sport_glasses:
                st.warning("⚠️ **Heads Up:** We found 'Sport glasses' in this batch and labeled them as 'Ski goggles' in the Meta Description. Double check them!")
                
            if found_polarized_clip_on:
                st.warning("⚠️ **Polarized Clip-On Alert:** We found a Marcolin/Kering clip-on that is marked as polarized, but it was assigned standard 'Sun clip-on' or 'Magnetic sun clip-on'. Verify if it needs the ' p' suffix!")

            st.write("### Preview of Filled Data:")
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