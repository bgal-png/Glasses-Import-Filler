import streamlit as st
import pandas as pd
import os
import re
from io import BytesIO
from sqlalchemy import create_engine
from dictionaries import TARGET_MAPPING, FACE_SHAPE_MAP, BRAND_USABLE_MAP, PREMIUM_KERING_BRANDS

# ==========================================
# 🛑 VERSION CHECK & CONFIG
# ==========================================
st.set_page_config(page_title="Manufacturer Data Linker", layout="wide")
APP_VERSION = "v.3000 (Cloud Edition ☁️)"

st.title(f"🏭 Manufacturer Data Filler")
st.caption(f"🚀 Running Code Version: **{APP_VERSION}**")

# ⚠️ REMEMBER: Change your password in Supabase later!
DB_URL = "postgresql://postgres.nxlwkzgfcmzsbogcenyi:lzAxicGTtsq1iEqB@aws-1-eu-central-1.pooler.supabase.com:5432/postgres?sslmode=require"

# ==========================================
# ☁️ CLOUD DATABASE CONNECTION
# ==========================================
@st.cache_data(show_spinner=False, ttl=3600) # Caches the data for 1 hour so it stays lightning fast
def load_cloud_database():
    try:
        engine = create_engine(DB_URL)
        # Pull the perfectly clean Gold Layer from your Vault
        df = pd.read_sql_table('master_catalog', con=engine)
        if 'join_key' in df.columns:
            df.set_index('join_key', inplace=True)
        return df
    except Exception as e:
        st.error(f"❌ Failed to connect to Cloud Database: {e}")
        return pd.DataFrame()

with st.spinner("☁️ Fetching Master Catalog from Vault..."):
    master_db = load_cloud_database()

if master_db.empty:
    st.warning("⚠️ Database is empty or unreachable. Please check your connection or run your build_master.py script.")
    st.stop()

# ==========================================
# 🎛️ SIDEBAR CONTROLS
# ==========================================
st.sidebar.header("Control Panel")

if st.sidebar.button("🔄 Refresh Cloud Data", type="primary"):
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
st.subheader("🔍 Quick EAN / Barcode Lookup")
col1, col2 = st.columns([3, 1])

with col1:
    search_ean = st.text_input("Enter EAN to search in Vault:", placeholder="e.g. 8056597123456")
with col2:
    st.write("") 
    st.write("")
    search_btn = st.button("Search Database", use_container_width=True)

if search_btn and search_ean:
    clean_search = re.sub(r'\.0$', '', str(search_ean).strip()).lstrip('0')
    
    if clean_search in master_db.index:
        st.success(f"✅ EAN '{search_ean}' found in the Vault!")
        found_data = master_db.loc[[clean_search]]
        
        display_cols = [c for c in ["Producing_company", "Brand", "Glasses_type", "Glasses_shape"] if c in found_data.columns]
        st.dataframe(found_data[display_cols], use_container_width=True)
    else:
        st.error(f"❌ EAN '{search_ean}' (Cleaned: {clean_search}) was NOT found in the cloud database.")

# ==========================================
# 📦 PACKAGE DATA LOADER
# ==========================================
st.divider()
st.markdown("### 📦 Step 1: Loading Package Data...")
package_df = pd.DataFrame() 
package_file_path = "package_data.xlsx"

if os.path.exists(package_file_path):
    try:
        package_df = pd.read_excel(package_file_path)
        st.success(f"✅ Package Data loaded locally! ({len(package_df)} items ready)")
        
        with st.expander("👀 Preview Package Data"):
            st.dataframe(package_df.head())
    except Exception as e:
        st.error(f"⚠️ Error reading local package_data.xlsx: {e}")
else:
    st.info("ℹ️ Local 'package_data.xlsx' not found. Weights/Dimensions will not be filled.")

# ==========================================
# 📥 TARGET FILE UPLOADER
# ==========================================
st.divider()
st.subheader("📥 Step 2: Upload Your File to Fill")

uploaded_file = st.file_uploader("Upload your Target Excel or CSV file", type=["xlsx", "csv"])

if uploaded_file:
    try:
        if uploaded_file.name.endswith('.csv'):
            target_df = pd.read_csv(uploaded_file, dtype=str)
        else:
            target_df = pd.read_excel(uploaded_file, dtype=str, engine='openpyxl')
            
        target_df.columns = target_df.columns.astype(str).str.replace('\n', ' ', regex=False).str.replace(r'\s+', ' ', regex=True).str.strip()
        
        # 🕵️‍♂️ PROGRESS TRACKER
        mapped_targets = set()
        for val in TARGET_MAPPING.values():
            if isinstance(val, list):
                for item in val:
                    mapped_targets.add(item)
            else:
                mapped_targets.add(val)
        
        custom_targets = {
            "Items type ID: 20", "Items packing ID: 21", "Name private",
            "Meta description", "Glasses for your face shape ID:94",
            "UV filter ID: 60", "Glasses usable ID: 51", "Glasses collection ID: 33",
            "HS Code", "Item description", "Glasses other features ID:99"
        }
        
        all_completed_targets = mapped_targets.union(custom_targets)
        status_list = []
        for col in target_df.columns:
            if col in all_completed_targets:
                status_list.append(f"✅ {col}")
            else:
                status_list.append(f"⏳ {col}")
                
        with st.expander("🔍 PROGRESS TRACKER: Exact Bucket Names"):
            st.markdown("**✅ = Rule Applied | ⏳ = Still Needs Logic/Mapping**")
            st.write(status_list)
        
    except Exception as e:
        st.error(f"Could not read your uploaded file: {e}")
        st.stop()

    target_barcode_col = TARGET_MAPPING.get("Barcode", "Barcode")
    if target_barcode_col not in target_df.columns:
        st.error(f"❌ Could not find the Barcode column '{target_barcode_col}' in your file. Found columns: {list(target_df.columns)}")
        st.stop()

    st.success(f"File uploaded! Contains {len(target_df)} rows. Click below to start matching.")

    # ==========================================
    # 🚀 THE AUTO-FILLER ENGINE
    # ==========================================
    if st.button("🚀 Run Auto-Filler", type="primary"):
        with st.spinner("Matching barcodes and pouring data..."):
            
            for global_col, target_col in TARGET_MAPPING.items():
                if isinstance(target_col, list):
                    for tc in target_col:
                        if tc not in target_df.columns:
                            target_df[tc] = ""
                else:
                    if target_col not in target_df.columns:
                        target_df[target_col] = "" 

            match_count = 0
            found_sport_glasses = False  
            
            for index, row in target_df.iterrows():
                raw_barcode = str(row[target_barcode_col]).strip()
                clean_barcode = re.sub(r'\.0$', '', raw_barcode).lstrip('0')
                
                if clean_barcode in master_db.index:
                    match_count += 1
                    master_row = master_db.loc[clean_barcode]
                    
                    is_frames = str(master_row.get("Glasses_type", "")).strip() == "Frames"
                    lens_cols_to_skip = ["Glasses_lens_Colour", "Glasses_lens_material", "Sunglasses_filter", "Glasses_lens_effect"]
                    
                    target_df.at[index, "Items type ID: 20"] = "Glasses"
                    target_df.at[index, "Items packing ID: 21"] = "Basic"

                    g_type = str(master_row.get("Glasses_type", "")).strip()
                    private_name = ""
                    
                    if "Sunglasses" in g_type:
                        if priv_sun: private_name = f"(Sunglasses {priv_sun})"
                    elif "Sport glasses" in g_type:
                        if priv_sport: private_name = f"(Sports glasses {priv_sport})"
                    elif "Driving glasses" in g_type:
                        if priv_drive: private_name = f"(Eyeglasses driving {priv_drive})"
                    elif "PC Glasses without power" in g_type:
                        if priv_pc: private_name = f"(Eyeglasses PC {priv_pc})"
                    elif "Frames" in g_type:
                        if priv_eye: private_name = f"(Eyeglasses {priv_eye})"
                        
                    if private_name:
                        target_df.at[index, "Name private"] = private_name.strip()
                    
                    assembled_name = str(master_row.get("Assembled_Name", "")).strip()
                    meta_desc = ""
                    
                    if "Sunglasses" in g_type:
                        meta_desc = f"Sunglasses {assembled_name}"
                    elif "Sport glasses" in g_type:
                        meta_desc = f"Ski goggles {assembled_name}"
                        found_sport_glasses = True  
                    elif "Driving glasses" in g_type:
                        meta_desc = f"Driving glasses {assembled_name}"
                    elif "PC Glasses without power" in g_type:
                        meta_desc = f"Computer glasses {assembled_name}"
                    elif "Frames" in g_type:
                        meta_desc = f"Eyeglasses {assembled_name}"
                        
                    if meta_desc:
                        target_df.at[index, "Meta description"] = meta_desc.strip()

                    for global_col, target_col in TARGET_MAPPING.items():
                        if global_col == "Barcode": continue
                        
                        if is_frames and global_col in lens_cols_to_skip:
                            continue
                        
                        if global_col in master_db.columns:
                            val = master_row[global_col]
                            if pd.notna(val) and str(val).strip() != "":
                                if isinstance(target_col, list):
                                    for tc in target_col:
                                        target_df.at[index, tc] = val
                                else:
                                    target_df.at[index, target_col] = val

                    g_shape_raw = str(master_row.get("Glasses_shape", "")).strip()
                    
                    if g_shape_raw and g_shape_raw.lower() not in ["nan", ""]:
                        shapes = [s.strip() for s in g_shape_raw.split(",")]
                        recommended_faces = set()
                        
                        for s in shapes:
                            for shape_key, face_val in FACE_SHAPE_MAP.items():
                                if shape_key.lower() == s.lower():
                                    for face in face_val.split("|"):
                                        recommended_faces.add(face)
                        
                        if recommended_faces:
                            target_df.at[index, "Glasses for your face shape ID:94"] = "|".join(sorted(list(recommended_faces)))
                    
                    if "Sunglasses" in g_type:
                        target_df.at[index, "UV filter ID: 60"] = "400"
                    
                    usable_tags = set()
                    raw_brand = str(master_row.get("Brand", "")).strip().lower()
                    if raw_brand in BRAND_USABLE_MAP:
                        usable_tags.add(BRAND_USABLE_MAP[raw_brand])
                        
                    lens_effect = str(master_row.get("Glasses_lens_effect", "")).strip()
                    
                    if "Sunglasses" in g_type:
                        if "Polarized" in lens_effect:
                            usable_tags.add("Driving glasses")
                        else:
                            usable_tags.add("Common use")
                            
                    if usable_tags:
                        target_df.at[index, "Glasses usable ID: 51"] = "|".join(sorted(list(usable_tags)))

                    if raw_brand in PREMIUM_KERING_BRANDS:
                        target_df.at[index, "Glasses collection ID: 33"] = "Prémiové brýle - Kering"

                    raw_material = str(master_row.get("Glasses_main_material", "")).strip().lower()
                    
                    if "Sunglasses" in g_type:
                        target_df.at[index, "HS Code"] = "90041091"
                    elif "Sport glasses" in g_type:
                        target_df.at[index, "HS Code"] = "90049090"
                    elif "Frames" in g_type:
                        if "plastic" in raw_material:
                            target_df.at[index, "HS Code"] = "90031100"
                        elif "metal" in raw_material:
                            target_df.at[index, "HS Code"] = "90031900"
                    
                    if "Frames" in g_type:
                        target_df.at[index, "Item description"] = "Eyeglasses"
                    elif "PC Glasses without power" in g_type:
                        target_df.at[index, "Item description"] = "PC Glasses without power"
                    elif "Driving glasses" in g_type:
                        target_df.at[index, "Item description"] = "Driving glasses"
                    elif "Sunglasses" in g_type:
                        has_plastic = "plastic" in raw_material
                        has_metal = "metal" in raw_material
                        
                        if has_plastic and has_metal:
                            target_df.at[index, "Item description"] = "Sunglasses, mixed plastic and metal frame"
                        elif has_plastic:
                            target_df.at[index, "Item description"] = "Sunglasses, plastic frame"
                        elif has_metal:
                            target_df.at[index, "Item description"] = "Sunglasses, metal frame"

                    other_features = set()
                    
                    if "Glasses other features ID:99" in target_df.columns:
                        existing_features = str(target_df.at[index, "Glasses other features ID:99"]).strip()
                        if existing_features and existing_features.lower() not in ["nan", ""]:
                            for e in existing_features.split("|"):
                                other_features.add(e.strip())
                    
                    if "SunGlasses RX lenses ID:108" in target_df.columns:
                        rx_val = str(target_df.at[index, "SunGlasses RX lenses ID:108"]).strip().lower()
                        if rx_val == "yes":
                            other_features.add("Prescription sunglasses")
                            
                    if "Glasses contain ID: 84" in target_df.columns:
                        contain_val = str(target_df.at[index, "Glasses contain ID: 84"]).strip().lower()
                        contain_items = [item.strip() for item in re.split(r'[,|]', contain_val) if item.strip()]
                        
                        clip_on_found = False
                        if "sun clip-on" in contain_items:
                            other_features.add("Sun clip-on")
                            clip_on_found = True
                        if "sun clip-on p" in contain_items:
                            other_features.add("Sun clip-on p")
                            clip_on_found = True
                        if "magnetic sun clip-on" in contain_items:
                            other_features.add("Magnetic sun clip-on")
                            clip_on_found = True
                        if "magnetic sun clip-on p" in contain_items:
                            other_features.add("Magnetic sun clip-on p")
                            clip_on_found = True
                            
                        if clip_on_found:
                            other_features.add("Glasses with sun clip-on")
                            
                    if other_features:
                        target_df.at[index, "Glasses other features ID:99"] = "|".join(sorted(list(other_features)))

            st.success(f"✅ Match Complete! Successfully filled {match_count} out of {len(target_df)} products.")

            if found_sport_glasses:
                st.warning("⚠️ **Heads Up:** We found 'Sport glasses' in this batch and labeled them as 'Ski goggles' in the Meta Description. Please double-check the final file to ensure they aren't cycling or swimming glasses!")
            
            st.write("### Preview of Filled Data:")
            st.dataframe(target_df.head(20), use_container_width=True)

            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                target_df.to_excel(writer, index=False, sheet_name='Filled_Data')
            processed_data = output.getvalue()

            st.download_button(
                label="📥 Download Filled Excel File",
                data=processed_data,
                file_name="Master_Filled_Glasses.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )