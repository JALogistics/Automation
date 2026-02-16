# Install ODBC Driver for SQL Server

## Issue
The older "SQL Server" driver doesn't support:
- Active Directory Interactive authentication
- Modern encryption (Encrypt=yes)
- Microsoft Fabric SQL databases

## Solution: Install ODBC Driver 18 for SQL Server

### Option 1: Download and Install Manually
1. Download from Microsoft: https://go.microsoft.com/fwlink/?linkid=2249004
2. Run the installer
3. Follow the installation wizard
4. Restart your terminal/IDE after installation

### Option 2: Install via Winget (if available)
```powershell
winget install Microsoft.ODBCDriver.18
```

### Option 3: Install via Chocolatey (if available)
```powershell
choco install sqlserver-odbcdriver
```

## After Installation
1. Close all Python terminals
2. Restart Cursor IDE
3. Run the test script again:
   ```
   python test_sql_connection.py
   ```

## Verify Installation
Run this command to check if the driver is installed:
```
python check_odbc_drivers.py
```

You should see "ODBC Driver 18 for SQL Server" or "ODBC Driver 17 for SQL Server" in the list.
