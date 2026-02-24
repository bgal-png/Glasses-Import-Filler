import pandas as pd
import os
from sqlalchemy import create_engine
from dictionaries import MANUFACTURER_CONFIG

# ⚠️ REMEMBER: Change your password in Supabase after we are done today!
DB_URL = "postgresql://postgres.nxlwkzgfcmzsbogcenyi:YQe2oULo6y6WXOZN@aws-1-eu-central-1.pooler.supabase.com:5432/postgres"

def upload_raw_catalogs():
    print("🚀 Connecting to Supabase Vault...")
    try:
        engine = create_engine(DB_URL)
        print("✅ Connected!\n")
    except Exception as e:
        print(f"❌ Connection failed: {e}")
        return

    current_dir = os.getcwd()

    # Loop through your existing dictionary config!
    for mfg_name, settings in MANUFACTURER_CONFIG.items():
        file_name = settings["file"]
        file_path = os.path.join(current_dir, file_name)
        
        table_name = f"raw_{mfg_name.lower()}"
        
        if not os.path.exists(file_path):
            print(f"⚠️ Skipping {mfg_name.title()}: File '{file_name}' not found.")
            continue
            
        print(f"📦 Reading {mfg_name.title()} raw file ({file_name})...")
        
        try:
            # Handle both CSV and Excel safely
            if file_name.endswith('.csv'):
                try:
                    df = pd.read_csv(file_path, dtype=str, on_bad_lines='skip', sep=',')
                except:
                    df = pd.read_csv(file_path, dtype=str, on_bad_lines='skip', sep=';')
            else:
                df = pd.read_excel(file_path, dtype=str, engine='openpyxl')
            
            # 🧹 Crucial SQL Prep: Clean column names so the database doesn't crash
            df.columns = df.columns.astype(str).str.strip()
            new_cols = []
            seen = {}
            for c in df.columns:
                if c in seen:
                    seen[c] += 1
                    new_cols.append(f"{c}_{seen[c]}") # SQL prefers underscores over dots
                else:
                    seen[c] = 0
                    new_cols.append(c)
            df.columns = new_cols
            
            print(f"⬆️ Uploading {len(df)} rows to table: '{table_name}'...")
            
            # Push directly to PostgreSQL!
            df.to_sql(name=table_name, con=engine, if_exists='replace', index=False)
            print(f"✅ {mfg_name.title()} successfully stored in '{table_name}'!\n")
            
        except Exception as e:
            print(f"❌ Error uploading {mfg_name}: {e}\n")

if __name__ == "__main__":
    upload_raw_catalogs()