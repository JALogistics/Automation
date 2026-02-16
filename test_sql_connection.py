import pyodbc

# SQL Server connection string using ODBC Driver 18
CONNECTION_STRING = (
    "Driver={ODBC Driver 18 for SQL Server};"
    "Server=gpsmx7yaheuenp4qzxuy66nrwm-rflbe3kwkdpuhonbclc3si4xtq.database.fabric.microsoft.com,1433;"
    "Database=Daily Reporting-dc3e16eb-30ce-458c-8716-b0861ce67918;"
    "Encrypt=yes;"
    "TrustServerCertificate=no;"
    "Authentication=ActiveDirectoryInteractive;"
)

print("Testing SQL Server connection...")
print("=" * 60)
print(f"Available drivers: {', '.join(pyodbc.drivers())}\n")

try:
    print("Connecting with ODBC Driver 18 for SQL Server...")
    conn = pyodbc.connect(CONNECTION_STRING)
    print("\n[SUCCESS] Connection established!")
    
    cursor = conn.cursor()
    
    # Test query to verify connection
    cursor.execute("SELECT @@VERSION")
    version = cursor.fetchone()
    print(f"\nSQL Server Version:")
    print(f"  {version[0][:100]}...")
    
    # Check if tables exist
    cursor.execute("""
        SELECT TABLE_NAME 
        FROM INFORMATION_SCHEMA.TABLES 
        WHERE TABLE_TYPE = 'BASE TABLE'
        ORDER BY TABLE_NAME
    """)
    
    tables = cursor.fetchall()
    print(f"\nAvailable Tables ({len(tables)}):")
    for table in tables:
        print(f"  - {table[0]}")
    
    cursor.close()
    conn.close()
    
    print("\n" + "=" * 60)
    print("Connection test completed successfully!")
    
except Exception as e:
    print("\n[FAILED] Connection failed!")
    print(f"\nError: {e}")
    print("\n" + "=" * 60)
    print("Possible solutions:")
    print("  1. Install ODBC Driver 17 or 18 for SQL Server from:")
    print("     https://go.microsoft.com/fwlink/?linkid=2249004")
    print("  2. Check if you have the correct Azure AD permissions")
    print("  3. Verify the connection string is correct")
    print("  4. Run: python check_odbc_drivers.py to see available drivers")
