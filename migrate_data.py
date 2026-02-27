import pandas as pd
from sqlalchemy import create_engine
import os

# ⚠️ Notice I changed 5432 to 6543, and added ?sslmode=require at the end!
DB_URL = "postgresql://postgres.pffajpiosudvcnlatoeh:Md6rl6P0hZtUeAxY@aws-1-eu-central-1.pooler.supabase.com:6543/postgres?sslmode=require"

def upload_reference_files():
    print("🚀 Connecting to Supabase Vault...")
    try:
        engine = create_engine(DB_URL)
        # Test connection
        with engine.connect() as conn:
            pass
        print("✅ Connection successful!\n")
    except Exception as e:
        print(f"❌ Connection failed: {e}")
        return

    # --- 1. UPLOAD PACKAGE DATA ---
    if os.path.exists("package_data.xlsx"):
        print("📦 Reading package_data.xlsx...")
        df_pkg = pd.read_excel("package_data.xlsx")
        
        # Clean column names for the database
        df_pkg.columns = df_pkg.columns.astype(str).str.strip()
        
        print("⬆️ Uploading to 'package_data' table...")
        df_pkg.to_sql('package_data', engine, if_exists='replace', index=False)
        print("✅ Package Data safely in the cloud!\n")
    else:
        print("⚠️ package_data.xlsx not found locally.\n")

    # --- 2. UPLOAD HISTORICAL DATA (MASTER CLEAN) ---
    if os.path.exists("master_clean.xlsx"):
        print("📜 Reading master_clean.xlsx...")
        df_hist = pd.read_excel("master_clean.xlsx", dtype=str, engine="openpyxl")
        
        if "Items type" in df_hist.columns:
            # Filter for Glasses only, just like we did in the app!
            df_glasses = df_hist[df_hist["Items type"].astype(str).str.strip().str.lower() == "glasses"]
            df_glasses.columns = df_glasses.columns.astype(str).str.strip()
            
            print(f"⬆️ Uploading {len(df_glasses)} rows to 'historical_data' table...")
            df_glasses.to_sql('historical_data', engine, if_exists='replace', index=False)
            print("✅ Historical Data safely in the cloud!\n")
        else:
            print("⚠️ 'Items type' column missing. Could not filter/upload historical data.\n")
    else:
        print("⚠️ master_clean.xlsx not found locally.\n")

    print("🎉 MIGRATION COMPLETE! You can now permanently delete these files from GitHub.")

if __name__ == "__main__":
    upload_reference_files()