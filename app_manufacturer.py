import streamlit as st
import pandas as pd
import os
import re
from io import BytesIO
from dictionaries import (
    TARGET_MAPPING,
    VALUE_TRANSLATOR,
    MANUFACTURER_CONFIG,
    FACE_SHAPE_MAP,
    BRAND_USABLE_MAP,
    PREMIUM_KERING_BRANDS,
)

# ==========================================
# 🛑 VERSION CHECK
# ==========================================
st.set_page_config(page_title="Manufacturer Data Linker", layout="wide")
APP_VERSION = "v.260223"

st.title(f"🏭 Manufacturer Data Filler")
st.caption(f"🚀 Running Code Version: **{APP_VERSION}**")

from sqlalchemy import create_engine

# ⚠️ Use whichever connection string worked for you (Port 5432, 6543, or the direct string)
DB_URL = "postgresql://postgres.nxlwkzgfcmzsbogcenyi:YQe2oULo6y6WXOZN@aws-1-eu-central-1.pooler.supabase.com:5432/postgres"

# ==========================================
# ☁️ CLOUD DATABASE CONNECTION
# ==========================================
@st.cache_data(show_spinner=False, ttl=3600) # Caches for 1 hour so the app stays lightning fast
def load_cloud_database():
    try:
        engine = create_engine(DB_URL)
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
    st.warning("⚠️ Database is empty or unreachable. Please run your build_master.py script first.")
    st.stop()

# ==========================================
# 🚀 APP EXECUTION & UI
# ==========================================

st.sidebar.header("Control Panel")

if st.sidebar.button("🗑️ Clear Memory & Reload Data", type="primary"):
    st.cache_data.clear()
    st.session_state.clear()
    st.rerun()

if not catalog:
    st.warning("No manufacturer catalogs loaded. Fix errors before proceeding.")
    st.stop()

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
    search_ean = st.text_input(
        "Enter EAN to search in loaded catalogs:", placeholder="e.g. 8056597123456"
    )
with col2:
    st.write("")  # Spacing to align the button
    st.write("")
    search_btn = st.button("Search Database", use_container_width=True)

if search_btn and search_ean:
    # Clean the input exactly like the engine cleans the source barcodes
    clean_search = re.sub(r"\.0$", "", str(search_ean).strip()).lstrip("0")

    if clean_search in master_db.index:
        st.success(f"✅ EAN '{search_ean}' found in the database!")
        # Fetch the row(s) and display a few key columns so you know exactly what it is
        found_data = master_db.loc[[clean_search]]

        # We will show Manufacturer, Brand, and whatever else is available
        display_cols = [
            c
            for c in ["Producing_company", "Brand", "Glasses_type", "Glasses_shape"]
            if c in found_data.columns
        ]
        st.dataframe(found_data[display_cols], use_container_width=True)
    else:
        st.error(
            f"❌ EAN '{search_ean}' (Cleaned: {clean_search}) was NOT found in any loaded manufacturer catalog."
        )
# --- 📦 AUTOMATIC PACKAGE DATA LOADER (LOCAL REPO) ---
st.divider()
package_df = pd.DataFrame()
package_file_path = "package_data.xlsx"

if os.path.exists(package_file_path):
    try:
        # Read the file directly from your GitHub repo folder
        package_df = pd.read_excel(package_file_path)
        st.success(
            f"✅ Package Data loaded from repository! ({len(package_df)} items ready)"
        )

        with st.expander("👀 Preview Package Data"):
            st.dataframe(package_df.head())
    except Exception as e:
        st.error(f"⚠️ Error reading local package_data.xlsx: {e}")
else:
    st.info(
        "ℹ️ Local 'package_data.xlsx' not found in root directory. Weights will not be filled."
    )

st.divider()
st.subheader("📥 Step 1: Upload Your File to Fill")

uploaded_file = st.file_uploader(
    "Upload your Target Excel or CSV file", type=["xlsx", "csv"]
)

if uploaded_file:
    try:
        if uploaded_file.name.endswith(".csv"):
            target_df = pd.read_csv(uploaded_file, dtype=str)
        else:
            target_df = pd.read_excel(uploaded_file, dtype=str, engine="openpyxl")

        # This strips line breaks, tabs, and squishes multiple spaces into a single space!
        target_df.columns = (
            target_df.columns.astype(str)
            .str.replace("\n", " ", regex=False)
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
        )

        # This strips line breaks, tabs, and squishes multiple spaces into a single space!
        target_df.columns = (
            target_df.columns.astype(str)
            .str.replace("\n", " ", regex=False)
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
        )

        # 🕵️‍♂️ THE CHEAT CODE: Progress Tracker Edition
        # 1. Grab all the target columns safely (flattens any lists it finds!)
        mapped_targets = set()
        for val in TARGET_MAPPING.values():
            if isinstance(val, list):
                for item in val:
                    mapped_targets.add(item)
            else:
                mapped_targets.add(val)

        # 2. List the columns we built custom engines for
        custom_targets = {
            "Items type ID: 20",
            "Items packing ID: 21",
            "Name private",
            "Meta description",
            "Glasses for your face shape ID:94",
            "UV filter ID: 60",
            "Glasses usable ID: 51",
            "Glasses collection ID: 33",
            "HS Code",
            "Item description",
            "Glasses other features ID:99",
        }

        # Combine them into one master list of "Finished" columns
        all_completed_targets = mapped_targets.union(custom_targets)

        # 3. Build the visual list
        status_list = []
        for col in target_df.columns:
            if col in all_completed_targets:
                status_list.append(f"✅ {col}")
            else:
                status_list.append(f"⏳ {col}")

        # 4. Display in a clean, click-to-open box
        with st.expander("🔍 PROGRESS TRACKER: Exact Bucket Names"):
            st.markdown("**✅ = Rule Applied | ⏳ = Still Needs Logic/Mapping**")
            st.write(status_list)

    except Exception as e:
        st.error(f"Could not read your uploaded file: {e}")
        st.stop()

    target_barcode_col = TARGET_MAPPING.get("Barcode", "Barcode")
    if target_barcode_col not in target_df.columns:
        st.error(
            f"❌ Could not find the Barcode column '{target_barcode_col}' in your file. Found columns: {list(target_df.columns)}"
        )
        st.stop()

    st.success(
        f"File uploaded! Contains {len(target_df)} rows. Click below to start matching."
    )

    if st.button("🚀 Run Auto-Filler", type="primary"):
        with st.spinner("Matching barcodes and pouring data..."):

            # 1. Safely create missing columns (Handles both strings and lists)
            for global_col, target_col in TARGET_MAPPING.items():
                if isinstance(target_col, list):
                    for tc in target_col:
                        if tc not in target_df.columns:
                            target_df[tc] = ""
                else:
                    if target_col not in target_df.columns:
                        target_df[target_col] = ""

            match_count = 0
            found_sport_glasses = False  # 🚨 Our new tripwire!

            for index, row in target_df.iterrows():
                raw_barcode = str(row[target_barcode_col]).strip()

                # 🔥 ZERO-STRIPPER
                clean_barcode = re.sub(r"\.0$", "", raw_barcode).lstrip("0")

                if clean_barcode in master_db.index:
                    match_count += 1
                    master_row = master_db.loc[clean_barcode]

                    # --- 🛑 FRAMES BYPASS RULE ---
                    # Check if the mapped type is exactly "Frames"
                    is_frames = (
                        str(master_row.get("Glasses_type", "")).strip() == "Frames"
                    )
                    lens_cols_to_skip = [
                        "Glasses_lens_Colour",
                        "Glasses_lens_material",
                        "Sunglasses_filter",
                        "Glasses_lens_effect",
                    ]
                    # --- 📦 STATIC BUCKET FILLS ---
                    target_df.at[index, "Items type ID: 20"] = "Glasses"
                    target_df.at[index, "Items packing ID: 21"] = "Basic"

                    # --- 🕵️ PRIVATE NAME ENGINE ---
                    g_type = str(master_row.get("Glasses_type", "")).strip()
                    private_name = ""

                    # The order here acts as a strict priority hierarchy!
                    if "Sunglasses" in g_type:
                        if priv_sun:
                            private_name = f"(Sunglasses {priv_sun})"
                    elif "Sport glasses" in g_type:
                        if priv_sport:
                            private_name = f"(Sports glasses {priv_sport})"
                    elif "Driving glasses" in g_type:
                        if priv_drive:
                            private_name = f"(Eyeglasses driving {priv_drive})"
                    elif "PC Glasses without power" in g_type:
                        if priv_pc:
                            private_name = f"(Eyeglasses PC {priv_pc})"
                    elif "Frames" in g_type:
                        if priv_eye:
                            private_name = f"(Eyeglasses {priv_eye})"

                    if private_name:
                        target_df.at[index, "Name private"] = private_name.strip()

                    # --- 📝 META DESCRIPTION ENGINE ---
                    assembled_name = str(master_row.get("Assembled_Name", "")).strip()
                    meta_desc = ""

                    if "Sunglasses" in g_type:
                        meta_desc = f"Sunglasses {assembled_name}"
                    elif "Sport glasses" in g_type:
                        meta_desc = f"Ski goggles {assembled_name}"
                        found_sport_glasses = True  # Trip the wire!
                    elif "Driving glasses" in g_type:
                        meta_desc = f"Driving glasses {assembled_name}"
                    elif "PC Glasses without power" in g_type:
                        meta_desc = f"Computer glasses {assembled_name}"
                    elif "Frames" in g_type:
                        meta_desc = f"Eyeglasses {assembled_name}"

                    if meta_desc:
                        target_df.at[index, "Meta description"] = meta_desc.strip()

                    # 2. Safely pour data (Handles both strings and lists)
                    for global_col, target_col in TARGET_MAPPING.items():
                        if global_col == "Barcode":
                            continue

                        # If it's a frame, skip the lens columns entirely
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
                    # --- 👤 FACE SHAPE ENGINE ---
                    g_shape_raw = str(master_row.get("Glasses_shape", "")).strip()

                    if g_shape_raw and g_shape_raw.lower() not in ["nan", ""]:
                        # Handle potential multiple shapes separated by commas
                        shapes = [s.strip() for s in g_shape_raw.split(",")]
                        recommended_faces = set()

                        for s in shapes:
                            for shape_key, face_val in FACE_SHAPE_MAP.items():
                                if shape_key.lower() == s.lower():
                                    # Split the mapped string by '|' and add to our set to remove duplicates
                                    for face in face_val.split("|"):
                                        recommended_faces.add(face)

                        if recommended_faces:
                            # Join the unique face shapes back together with the pipe separator
                            # (We sort them so they always appear in a clean, alphabetical order!)
                            target_df.at[index, "Glasses for your face shape ID:94"] = (
                                "|".join(sorted(list(recommended_faces)))
                            )

                    # --- ☀️ UV FILTER ENGINE ---
                    if "Sunglasses" in g_type:
                        target_df.at[index, "UV filter ID: 60"] = "400"

                    # --- 🎯 GLASSES USABLE ENGINE ---
                    usable_tags = set()

                    # 1. Brand Logic
                    raw_brand = str(master_row.get("Brand", "")).strip().lower()
                    if raw_brand in BRAND_USABLE_MAP:
                        usable_tags.add(BRAND_USABLE_MAP[raw_brand])

                    # 2. Polarized / Sunglasses Logic
                    lens_effect = str(master_row.get("Glasses_lens_effect", "")).strip()

                    if "Sunglasses" in g_type:
                        if "Polarized" in lens_effect:
                            usable_tags.add("Driving glasses")
                        else:
                            usable_tags.add("Common use")

                    # 3. Join and Pour
                    if usable_tags:
                        # Sorting them ensures it always looks clean, e.g., "Common use|Luxury glasses"
                        target_df.at[index, "Glasses usable ID: 51"] = "|".join(
                            sorted(list(usable_tags))
                        )

                    # --- 💎 PREMIUM COLLECTION ENGINE ---
                    # We reuse the 'raw_brand' variable from the Usable Engine above!
                    if raw_brand in PREMIUM_KERING_BRANDS:
                        target_df.at[index, "Glasses collection ID: 33"] = (
                            "Prémiové brýle - Kering"
                        )

                        # --- 🌍 HS CODE ENGINE ---
                    # We reuse 'g_type' from the very beginning of the loop!
                    raw_material = (
                        str(master_row.get("Glasses_main_material", "")).strip().lower()
                    )

                    if "Sunglasses" in g_type:
                        target_df.at[index, "HS Code"] = "90041091"
                    elif "Sport glasses" in g_type:
                        target_df.at[index, "HS Code"] = "90049090"
                    elif "Frames" in g_type:
                        if "plastic" in raw_material:
                            target_df.at[index, "HS Code"] = "90031100"
                        elif "metal" in raw_material:
                            target_df.at[index, "HS Code"] = "90031900"

                    # --- 📝 ITEM DESCRIPTION ENGINE ---
                    # We continue to reuse 'g_type' and 'raw_material' from above!
                    if "Frames" in g_type:
                        target_df.at[index, "Item description"] = "Eyeglasses"
                    elif "PC Glasses without power" in g_type:
                        target_df.at[index, "Item description"] = (
                            "PC Glasses without power"
                        )
                    elif "Driving glasses" in g_type:
                        target_df.at[index, "Item description"] = "Driving glasses"
                    elif "Sunglasses" in g_type:
                        has_plastic = "plastic" in raw_material
                        has_metal = "metal" in raw_material

                        if has_plastic and has_metal:
                            target_df.at[index, "Item description"] = (
                                "Sunglasses, mixed plastic and metal frame"
                            )
                        elif has_plastic:
                            target_df.at[index, "Item description"] = (
                                "Sunglasses, plastic frame"
                            )
                        elif has_metal:
                            target_df.at[index, "Item description"] = (
                                "Sunglasses, metal frame"
                            )
                            # ---------------------------------------------------------
                    # THE FOLLOWING RULES MUST RUN *AFTER* THE GENERIC POUR
                    # ---------------------------------------------------------

                    # --- 🌟 OTHER FEATURES ENGINE ---
                    other_features = set()

                    # If the column already received data from the generic pour, grab it so we don't overwrite it!
                    if "Glasses other features ID:99" in target_df.columns:
                        existing_features = str(
                            target_df.at[index, "Glasses other features ID:99"]
                        ).strip()
                        if existing_features and existing_features.lower() not in [
                            "nan",
                            "",
                        ]:
                            for e in existing_features.split("|"):
                                other_features.add(e.strip())

                    # 1. Check RX Lenses
                    if "SunGlasses RX lenses ID:108" in target_df.columns:
                        rx_val = (
                            str(target_df.at[index, "SunGlasses RX lenses ID:108"])
                            .strip()
                            .lower()
                        )
                        if rx_val == "yes":
                            other_features.add("Prescription sunglasses")

                    # 2. Check Clip-ons
                    if "Glasses contain ID: 84" in target_df.columns:
                        contain_val = (
                            str(target_df.at[index, "Glasses contain ID: 84"])
                            .strip()
                            .lower()
                        )

                        # Split by comma or pipe to isolate the exact phrases
                        # (Prevents "magnetic sun clip-on" from accidentally triggering the basic "sun clip-on" rule!)
                        contain_items = [
                            item.strip()
                            for item in re.split(r"[,|]", contain_val)
                            if item.strip()
                        ]

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

                        # 3. Add the universal clip-on tag
                        if clip_on_found:
                            other_features.add("Glasses with sun clip-on")

                    # 4. Pour into target
                    if other_features:
                        target_df.at[index, "Glasses other features ID:99"] = "|".join(
                            sorted(list(other_features))
                        )

            st.success(
                f"✅ Match Complete! Successfully filled {match_count} out of {len(target_df)} products."
            )

            # 🚨 Trigger the Sport Glasses Warning if the tripwire was crossed
            if found_sport_glasses:
                st.warning(
                    "⚠️ **Heads Up:** We found 'Sport glasses' in this batch and labeled them as 'Ski goggles' in the Meta Description. Please double-check the final file to ensure they aren't cycling or swimming glasses!"
                )

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
