from sqlalchemy import create_engine, text

# ⚠️ Paste your Supabase Connection String here
# Make sure it starts with 'postgresql://' (NOT postgresql+psycopg2:// yet)
DB_URL = "postgresql://postgres.nxlwkzgfcmzsbogcenyi:lzAxicGTtsq1iEqB@aws-1-eu-central-1.pooler.supabase.com:5432/postgres"

def test_connection():
    print("⏳ Attempting to connect to the Vault...")
    try:
        # Create the engine that drives data to the database
        engine = create_engine(DB_URL)
        
        # Open a connection and ask the database what version it is running
        with engine.connect() as connection:
            result = connection.execute(text("SELECT version();"))
            db_version = result.fetchone()[0]
            
            print("\n✅ CONNECTION SUCCESSFUL!")
            print(f"🧠 Database Engine: {db_version}")
            
    except Exception as e:
        print("\n❌ CONNECTION FAILED. Here is the error:")
        print(e)

if __name__ == "__main__":
    test_connection()