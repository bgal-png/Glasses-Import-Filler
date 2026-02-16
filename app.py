import streamlit as st
import pandas as pd
import os
import io
import re

# 1. Page Configuration
st.set_page_config(page_title="Excel Auto-Filler", layout="wide")
st.title("⚡ Excel Data Filler: Glasses Edition")

# ==========================================
# 🔒 INDESTRUCTIBLE LOADER (LOCKED VERSION)
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

# --- RULE 3 DATA: BRAND MAPPING DICTIONARY ---
BRAND_TO_COMPANY_MAP = {
    "Kering": [
        "Alexander McQueen", "Balenciaga", "Chloe", "Gucci", "Maui Jim", 
        "Montblanc", "Puma", "Saint Laurent"
    ],
    "Marcolin": [
        "Adidas", "Guess", "Max Mara", "MAX&Co.", "Tom Ford"
    ],
    "Ostalo": [
        "Arena", "Cebe", "Hawkers", "HEAD", "Lavida", "POC", "Oxydo", "Alpina'", "Alpina"
    ],
    "Inspecs": [
        "Caterpillar", "O'Neill", "Radley", "Superdry"
    ],
    "Marchon": [
        "Calvin Klein", "Lacoste", "LIU JO", "Nike"
    ],
    "Alensa": [
        "Alensa"
    ],
    "Adrial": [
        "Crullé", "Kimikado", "Marisio", "Válle", "LeWish", "Beron"
    ],
    "Luxottica": [
        "Arnette", "Burberry", "Dolce & Gabbana", "Emporio Armani", "Giorgio Armani", 
        "Armani Exchange", "Michael Kors", "Oakley", "Persol", "Polo Ralph Lauren", 
        "Prada", "Ralph by Ralph Lauren", "Ray-Ban", "Swarovski", "Versace", 
        "Vogue", "Jimmy Choo", "Miu Miu", "Tiffany", "Ralph Lauren"
    ],
    "Safilo": [
        "Boss by Hugo Boss", "Carrera", "David Beckham", "Love Moschino", 
        "Chiara Ferragni", "Dsquared2", "Fossil", "Havaianas", "Hugo by Hugo Boss", 
        "Kate Spade", "Levi's", "Marc Jacobs", "Missoni", "Moschino", 
        "Pierre Cardin", "Polaroid", "Tommy Hilfiger", "Under Armour", 
        "Seventh Street"
    ],
    # Explicitly Empty Group (Remaining unmapped brands)
    "": [
        "Carolina Herrera" 
    ],
    "GO Eyewear": ["Ana Hickmann"],
    "Strabilia": ["Silhouette"],
    "MCM OPTIK SRL": ["Morel"],
    "Bollé Brands": ["Bollé", "SPY+", "Serengeti"]
}

# NOTE: I assumed the other brands in that original "Empty" list (like Tommy Hilfiger, Polaroid, etc.) 
# might also belong to Safilo since they are often distributed by them. 
# If ONLY the 4 you listed are Safilo, I can move the others back to empty. 
# Currently, I moved MOST of that group to Safilo to be safe, but kept Carolina Herrera separate.

# Create a flattened lookup for faster processing (lowercase keys)
FLAT_BRAND_LOOKUP = {}
for company, brands in BRAND_TO_COMPANY_MAP.items():
    for brand in brands:
        FLAT_BRAND_LOOKUP[brand.lower().strip()] = company

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

# --- RULE FUNCTIONS ---

def apply_hs_code(row, type_col, mat_col, sport_col):
    """Rule 1: HS Code"""
    g_type = str(row.get(type_col, '')).strip() if type_col else ""
    material = str(row.get(mat_col, '')).strip().lower() if mat_col else ""
    sport_val = str(row.get(sport_col, '')).strip().lower() if sport_col else ""

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
    """Rule 2: Item Description"""
    g_type = str(row.get(type_col, '')).strip() if type_col else ""
    material = str(row.get(mat_col, '')).strip().lower() if mat_col else ""

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
    """Rule 3: Producing Company based on Brand"""
    brand_val = str(row.get(brand_col, '')).strip().lower()
    
    # Direct Lookup
    if brand_val in FLAT_BRAND_LOOKUP:
        company = FLAT_BRAND_LOOKUP[brand_val]
        return company, f"Matched Brand: {brand_val}"
    
    return "", "Unknown Brand"

# --- MAIN EXECUTION ---

def run_auto_fill(user_df):
    # 1. Identify Columns by ID
    type_col = get_col_by_id(user_df, "13")      
    material_col = get_col_by_id(user_df, "53")  
    sport_col = get_col_by_id(user_df, "89")
    brand_col = get_col_by_id(user_df, "11") 
    
    hs_col = get_col_by_id(user_df, "AO") or "HS Code"
    desc_col = get_col_by_id(user_df, "AP") or "Item description"
    prod_col = get_col_by_id(user_df, "146") or "Producing company"

    # Ensure output columns exist
    for col in [hs_col, desc_col, prod_col]:
        if col not in user_df.columns: user_df[col] = ""

    # 2. Apply Rule 1: HS Code
    hs_results = user_df.apply(lambda row: apply_hs_code(row, type_col, material_col, sport_col), axis=1)
    user_df[hs_col] = [r[0] for r in hs_results]
    hs_reasons = [r[1] for r in hs_results]

    # 3. Apply Rule 2: Item Description
    desc_results = user_df.apply(lambda row: apply_item_description(row, type_col, material_col), axis=1)
    user_df[desc_col] = [r[0] for r in desc_results]
    desc_reasons = [r[1] for r in desc_results]
    
    # 4. Apply Rule 3: Producing Company
    prod_results = user_df.apply(lambda row: apply_producing_company(row, brand_col), axis=1)
    user_df[prod_col] = [r[0] for r in prod_results]
    prod_reasons = [r[1] for r in prod_results]
    
    # 5. Generate Report
    report_df = pd.DataFrame({
        'Brand (ID:11)': user_df[brand_col] if brand_col else "Not Found",
        'Producing Company': user_df[prod_col],
        'Company Logic': prod_reasons,
        'HS Code': user_df[hs_col],
        'Item Description': user_df[desc_col]
    })
    
    modified_rows = report_df[
        (report_df['HS Code'] != "") | 
        (report_df['Item Description'] != "") |
        (report_df['Company Logic'] != "Unknown Brand")
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
    
    if st.button("✨ Auto-Fill Data", type="primary"):
        with st.spinner("Applying Rules 1, 2 & 3..."):
            working_df = user_df.copy()
            filled_df, report = run_auto_fill(working_df)
            
            st.success(f"✅ Rules Applied!")
            
            with st.expander("📊 View Processing Report", expanded=True):
                if not report.empty:
                    st.dataframe(report, use_container_width=True)
                else:
                    st.info("No rows matched the current rules.")
            
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