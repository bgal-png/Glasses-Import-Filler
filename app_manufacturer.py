import streamlit as st
import pandas as pd
from io import BytesIO
from sqlalchemy import create_engine

from dictionaries import TARGET_MAPPING
from filler_core import (
    FillOptions,
    changed_columns,
    extract_images_from_zip,
    fill_target,
    read_target_file,
    run_ai_vision,
    write_filled_excel,
)

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
        except Exception:
            pass

        # 3. Fetch Global Categories (kept for the status metric only — the
        #    filler now uses the static BRAND_GLASSES_CONTAIN table instead)
        try:
            historical_df = pd.read_sql_table('historical_data', con=engine)
        except Exception:
            pass

        # 4. Fetch Item Origin
        try:
            origin_df = pd.read_sql_table('origin_data', con=engine)
        except Exception:
            pass

        return master_db, package_df, historical_df, origin_df
    except Exception as e:
        st.error(f"❌ Failed to connect to Cloud Database: {e}")
        return master_db, package_df, historical_df, origin_df

with st.spinner("☁️ Fetching live data from Supabase Vault..."):
    master_db, package_df, master_clean_df, origin_df = load_cloud_data()

if master_db.empty:
    st.warning("⚠️ Database is empty. Please upload a catalogue through the admin panel first.")
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
        target_df = read_target_file(uploaded_file)
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
            options = FillOptions(
                priv_sun=priv_sun,
                priv_eye=priv_eye,
                priv_pc=priv_pc,
                priv_sport=priv_sport,
                priv_drive=priv_drive,
            )
            try:
                target_df, report = fill_target(
                    target_df, master_db, package_df, origin_df, options=options
                )
            except ValueError as e:
                st.error(f"❌ {e}")
                st.stop()

            st.success(
                f"✅ Match Complete! Successfully filled {report.match_count} "
                f"out of {report.total_rows} products."
            )

            # --- 👓 AI VISION ENGINE (Shape + Sport Detection) ---
            if image_dict and has_api_key:
                shape_bar = st.progress(0, text="🔍 Classifying with AI vision...")

                def _ai_progress(frac, text):
                    shape_bar.progress(frac, text=f"🔍 {text}")

                ai = run_ai_vision(target_df, image_dict, ANTHROPIC_API_KEY, progress=_ai_progress)
                shape_bar.empty()
                st.success(
                    f"👓 AI Vision Complete! Shapes: {ai.shape_count}, "
                    f"Sport glasses: {ai.sport_count} (out of {ai.image_count} images)"
                )

            # --- VALIDATION REPORT ---
            if report.unmapped or report.missing:
                st.warning(f"⚠️ **Validation Report:** {report.total_issues} potential issues found")
                if report.unmapped:
                    with st.expander(f"🔴 Unmapped values ({len(report.unmapped)} columns)"):
                        for col, vals in sorted(report.unmapped.items()):
                            st.write(f"**{col}:** {len(vals)} unmapped value(s)")
                            for v in sorted(vals):
                                st.caption(f"  → `{v}`")
                if report.missing:
                    with st.expander(f"🟡 Missing from source ({len(report.missing)} columns)"):
                        for col, count in sorted(report.missing.items(), key=lambda x: -x[1]):
                            st.write(f"**{col}:** {count} row(s) with no data")

            if report.found_sport_glasses:
                st.warning("⚠️ **Heads Up:** We found 'Sport glasses' in this batch and labeled them as 'Ski goggles' in the Meta Description. Double check them!")

            if report.found_polarized_clip_on:
                st.warning("⚠️ **Polarized Clip-On Alert:** We found a Marcolin/Kering clip-on that is marked as polarized, but it was assigned standard 'Sun clip-on' or 'Magnetic sun clip-on'. Verify if it needs the ' p' suffix!")

            st.write("### 📊 Before / After Comparison")
            preview_tab1, preview_tab2, preview_tab3 = st.tabs(["🔄 Changes Only", "📥 Original", "📤 Filled"])
            with preview_tab1:
                changed_cols = changed_columns(original_df, target_df)
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
            write_filled_excel(target_df, output)
            processed_data = output.getvalue()

            st.download_button(
                label="📥 Download Filled Excel File",
                data=processed_data,
                file_name="Master_Filled_Glasses.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
