import os
import re
import sys
from datetime import datetime
from pathlib import Path
import pandas as pd
import pyodbc
from dotenv import load_dotenv

# SQL Server connection string using ODBC Driver 18
CONNECTION_STRING = (
    "Driver={ODBC Driver 18 for SQL Server};"
    "Server=gpsmx7yaheuenp4qzxuy66nrwm-rflbe3kwkdpuhonbclc3si4xtq.database.fabric.microsoft.com,1433;"
    "Database=Daily Reporting-dc3e16eb-30ce-458c-8716-b0861ce67918;"
    "Encrypt=yes;"
    "TrustServerCertificate=no;"
    "Authentication=ActiveDirectoryInteractive;"
)


def get_sql_connection():
    """
    Get SQL Server connection using ODBC Driver 18.
    
    Returns:
        pyodbc.Connection: Database connection object
        
    Raises:
        Exception: If connection fails
    """
    try:
        print("Connecting to SQL Server with ODBC Driver 18...")
        conn = pyodbc.connect(CONNECTION_STRING)
        print("Connected successfully!")
        return conn
    except Exception as e:
        error_msg = (
            f"Failed to connect to SQL Server: {e}\n\n"
            "Please ensure:\n"
            "1. ODBC Driver 18 for SQL Server is installed\n"
            "   Download: https://go.microsoft.com/fwlink/?linkid=2249004\n"
            "2. You have proper Azure AD permissions\n"
            "3. Your network allows connection to the database\n"
        )
        print(error_msg)
        raise Exception(error_msg)


def run_daily_data_transfer():
    """
    Main function to handle the daily data transfer process from daily_report to current_report.
    Includes data fetching, transformation, and insertion into SQL Server.
    """
    # Load environment variables from default .env location if present
    env_path = Path(__file__).parents[1] / ".env"
    if env_path.exists():
        load_dotenv(env_path)
    else:
        load_dotenv()

    # Initialize SQL Server connection
    try:
        conn = get_sql_connection()
        cursor = conn.cursor()
        print("Connected to SQL Server successfully.")
    except Exception as e:
        print(f"Failed to connect to SQL Server: {e}")
        return False

    print("Starting daily data transfer process.")

    # Helper functions for data transformation
    def remove_special_chars(val):
        if pd.isnull(val):
            return val
        return re.sub(r'[^A-Za-z0-9 ]+', '', str(val)).strip()

    def get_status(row):
        if pd.isna(row.get('Outbound date')):
            if pd.isna(row.get('Inbound date')):
                return "On Sea"
            else:
                return "In-Stock"
        else:
            try:
                outbound_date = pd.to_datetime(row.get('Outbound date')).date() if pd.notna(row.get('Outbound date')) else None
                if outbound_date and outbound_date > datetime.now().date():
                    return "In-Stock"
                else:
                    return "Outbounded"
            except Exception:
                return "Outbounded"

    def get_outbound_class(row):
        try:
            if pd.isna(row.get('Outbound date')) or row.get('Outbound date') == pd.Timestamp(0):
                return "not-outbounded"
            elif pd.to_datetime(row.get('Outbound date')).date() > datetime.now().date():
                return "outbound-planned"
            else:
                return "outbounded"
        except Exception:
            return "not-outbounded"

    def get_release_status(row):
        if not pd.isna(row.get('Release date')) or row.get('Current_Status') == "Outbounded":
            return "Released"
        else:
            return "Not released"

    def is_outbounded_before_2025(row):
        if row.get('Outbound_status') == 'outbounded' and pd.notna(row.get('Outbound date')):
            try:
                outbound_year = pd.to_datetime(row.get('Outbound date')).year
                return outbound_year < 2025
            except Exception:
                return False
        return False

    def fix_agreed_delivery_date(row):
        try:
            agreed_val = row.get('Agreed Delivery date')
            if pd.isna(agreed_val) or str(agreed_val).strip() == "":
                return ""
            agreed_date = pd.to_datetime(agreed_val, errors='coerce')
            if pd.notna(agreed_date):
                if agreed_date.year in [2001, 2021, 2024]:
                    return ""
        except Exception:
            pass
        return row.get('Agreed Delivery date')

    # Fetch data from SQL Server
    try:
        query = "SELECT * FROM daily_report"
        df = pd.read_sql(query, conn)
        print(f"Fetched {len(df)} rows from daily_report.")
    except Exception as e:
        print(f"Error fetching data from daily_report: {e}")
        cursor.close()
        conn.close()
        return False

    if df.empty:
        print("No data found in daily_report.")
        cursor.close()
        conn.close()
        return False

    # 1. Remove rows where "Container No." is null or blank
    df = df[df["Container No."].notnull() & (df["Container No."].astype(str).str.strip() != "")]

    # 2. Clean specified columns
    for col in ["Container No.", "Release Number", "Import invoice"]:
        if col in df.columns:
            df[col] = df[col].apply(remove_special_chars)

    # 3. Calculate Power
    if "Piece" in df.columns and "Wattage" in df.columns:
        df["Power"] = pd.to_numeric(df["Piece"], errors='coerce') * pd.to_numeric(df["Wattage"], errors='coerce')
    else:
        df["Power"] = None

    # 4. Calculate MegaWattage
    df["MegaWattage"] = df["Power"] / 1_000_000

    # 5. Create Ref1
    if "Container No." in df.columns and "Release Number" in df.columns:
        df["Ref1"] = df["Release Number"].astype(str) + df["Container No."].astype(str)
    else:
        df["Ref1"] = None

    # 6. Create Ref2
    if "Container No." in df.columns and "Release Number" in df.columns and "Wattage" in df.columns:
        df["Ref2"] = df["Release Number"].astype(str) + df["Container No."].astype(str) + df["Wattage"].astype(str)
    else:
        df["Ref2"] = None

    # 7. Add Status columns
    df['Current_Status'] = df.apply(get_status, axis=1)
    df['Outbound_status'] = df.apply(get_outbound_class, axis=1)
    df['Release_Status'] = df.apply(get_release_status, axis=1)

    # 8. Add Delivery Status
    ref_counts = df['Ref1'].map(df['Ref1'].value_counts())
    df['Delivery_Status'] = ref_counts.apply(lambda x: "Partial_delivery" if x > 1 else "Full_delivery")

    # 9. Filter out old outbounded data
    df = df[~df.apply(is_outbounded_before_2025, axis=1)]

    # 10. Fix Agreed Delivery date
    if 'Agreed Delivery date' in df.columns:
        df['Agreed Delivery date'] = df.apply(fix_agreed_delivery_date, axis=1)

    # 11. Remove unnecessary columns
    for col in ['Total Wattage', 'MW']:
        if col in df.columns:
            df = df.drop(col, axis=1)

    # Prepare data for insertion
    df = df.where(pd.notnull(df), None)
    
    print(f"Transformed data. Ready to insert {len(df)} records into current_report.")

    # Insert data into SQL Server using pandas to_sql
    try:
        df.to_sql(
            'current_report',
            conn,
            if_exists='append',
            index=False,
            chunksize=500,
            method='multi'
        )
        print(f"Successfully inserted {len(df)} records into current_report.")
    except Exception as e:
        print(f"Error inserting data into current_report: {e}")
        cursor.close()
        conn.close()
        return False

    # Close the connection
    cursor.close()
    conn.close()
    
    print("Daily data transfer completed.")
    return True

if __name__ == "__main__":
    success = run_daily_data_transfer()
    sys.exit(0 if success else 1) 