# SQL Server Connection Setup Guide

## Current Issue
Your system has the old "SQL Server" ODBC driver, which doesn't support:
- Active Directory Interactive authentication
- Modern encryption features
- Microsoft Fabric SQL databases

## Solution: Install Modern ODBC Driver

### Step 1: Download and Install ODBC Driver

**Quick Download Link:**
https://go.microsoft.com/fwlink/?linkid=2249004

This will download "ODBC Driver 18 for SQL Server"

**Installation Steps:**
1. Click the download link above
2. Run the downloaded installer
3. Follow the installation wizard (accept defaults)
4. Click "Finish" when complete

**Alternative Versions:**
- ODBC Driver 17: https://go.microsoft.com/fwlink/?linkid=2187214
- Both versions work with Azure AD Interactive authentication

### Step 2: Verify Installation

After installation, run this command to verify:

```bash
python check_odbc_drivers.py
```

**Expected Output:**
You should see one of these in the list:
- `ODBC Driver 18 for SQL Server`
- `ODBC Driver 17 for SQL Server`

### Step 3: Test Connection

Run the connection test:

```bash
python test_sql_connection.py
```

**What to Expect:**
1. A browser window will open for Azure AD login
2. Sign in with your Microsoft account
3. The script will confirm successful connection

### Step 4: Run Your Scripts

Once the connection test passes, you can run your ETL and data transfer scripts:

```bash
# Run ETL process
python tests/ETL.py

# Run daily data transfer
python tests/daily-data-transfer.py
```

## Troubleshooting

### Issue: "No suitable ODBC driver found"
**Solution:** Install the ODBC driver from the link above

### Issue: "Authentication failed"
**Solution:** 
- Ensure you're logged into Azure AD with the correct account
- Check that your account has permissions to the database
- Try running the test script again

### Issue: "Connection timeout"
**Solution:**
- Check your internet connection
- Verify the server address is correct
- Check if your firewall is blocking port 1433

## What Changed in Your Code

### Before (Supabase):
```python
from supabase import create_client, Client
supabase = create_client(url, key)
response = supabase.table("daily_report").select("*").execute()
```

### After (SQL Server):
```python
import pyodbc
conn = get_sql_connection()  # Tries Driver 18, then Driver 17
df = pd.read_sql("SELECT * FROM daily_report", conn)
```

## Files Updated

1. `tests/ETL.py` - Migrated to SQL Server
2. `tests/daily-data-transfer.py` - Migrated to SQL Server
3. `requirements.txt` - Updated dependencies

## Support

If you continue to have issues:
1. Check the error message in the terminal
2. Verify ODBC driver is installed: `python check_odbc_drivers.py`
3. Test basic connection: `python test_sql_connection.py`
4. Check Azure AD permissions with your database administrator
