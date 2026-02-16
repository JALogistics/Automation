"""
SQL Server connection helper that tries multiple ODBC drivers.
"""
import pyodbc

def get_sql_connection():
    """
    Try to connect to SQL Server using available ODBC drivers.
    Tries in order: Driver 18 -> Driver 17 -> fallback message
    
    Returns:
        pyodbc.Connection: Database connection object
        
    Raises:
        Exception: If no suitable driver is found or connection fails
    """
    server = "gpsmx7yaheuenp4qzxuy66nrwm-rflbe3kwkdpuhonbclc3si4xtq.database.fabric.microsoft.com,1433"
    database = "Daily Reporting-dc3e16eb-30ce-458c-8716-b0861ce67918"
    
    # List of drivers to try in order
    drivers_to_try = [
        "ODBC Driver 18 for SQL Server",
        "ODBC Driver 17 for SQL Server",
    ]
    
    # Get available drivers
    available_drivers = pyodbc.drivers()
    print(f"Available ODBC drivers: {', '.join(available_drivers)}")
    
    last_error = None
    
    for driver in drivers_to_try:
        if driver in available_drivers:
            connection_string = (
                f"Driver={{{driver}}};"
                f"Server={server};"
                f"Database={database};"
                "Encrypt=yes;"
                "TrustServerCertificate=no;"
                "Connection Timeout=30;"
                "Authentication=ActiveDirectoryInteractive;"
            )
            
            try:
                print(f"Trying to connect with {driver}...")
                conn = pyodbc.connect(connection_string)
                print(f"Successfully connected using {driver}")
                return conn
            except Exception as e:
                last_error = e
                print(f"Failed with {driver}: {e}")
                continue
    
    # If we get here, no driver worked
    error_msg = (
        "\n" + "="*80 + "\n"
        "ERROR: Unable to connect to SQL Server\n"
        "="*80 + "\n"
        "No suitable ODBC driver found for SQL Server.\n\n"
        "Available drivers on your system:\n"
    )
    for driver in available_drivers:
        error_msg += f"  - {driver}\n"
    
    error_msg += (
        "\nRequired: ODBC Driver 17 or 18 for SQL Server\n\n"
        "To fix this, install the ODBC driver:\n"
        "1. Download from: https://go.microsoft.com/fwlink/?linkid=2249004\n"
        "2. Run the installer\n"
        "3. Restart your IDE\n"
        "="*80 + "\n"
    )
    
    if last_error:
        error_msg += f"\nLast error: {last_error}\n"
    
    raise Exception(error_msg)


if __name__ == "__main__":
    # Test the connection
    try:
        conn = get_sql_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT @@VERSION")
        version = cursor.fetchone()
        print(f"\nConnected! SQL Server version: {version[0][:80]}...")
        cursor.close()
        conn.close()
    except Exception as e:
        print(str(e))
