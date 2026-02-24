import psycopg2

# The exact pooler credentials from your Supabase dashboard
USER = "postgres.nxlwkzgfcmzsbogcenyi"
PASSWORD = "YQe2oULo6y6WXOZN"  # <-- Put your real password here!
HOST = "aws-1-eu-central-1.pooler.supabase.com"
PORT = "5432"
DBNAME = "postgres"

print("⏳ Testing raw connection to Supabase pooler...")

try:
    # Connect using raw parameters instead of a URL string
    connection = psycopg2.connect(
        user=USER,
        password=PASSWORD,
        host=HOST,
        port=PORT,
        dbname=DBNAME,
        sslmode="require"  # Supabase requires SSL
    )
    print("✅ Connection successful! The circuit breaker is CLOSED.")
    
    cursor = connection.cursor()
    cursor.execute("SELECT NOW();")
    print("Current Database Time:", cursor.fetchone()[0])
    
    cursor.close()
    connection.close()

except Exception as e:
    print(f"❌ Failed to connect:\n{e}")