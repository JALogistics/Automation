import os
import re
import sys
from datetime import datetime
from pathlib import Path
import pandas as pd
from supabase import create_client, Client
from dotenv import load_dotenv

def run_daily_data_transfer():
    """
    Main function to handle the daily data transfer process from daily_report to current_report.
    Includes data fetching, transformation, and insertion into Supabase.
    """
    # Load environment variables from default .env location if present
    env_path = Path(__file__).parents[1] / ".env"
    if env_path.exists():
        load_dotenv(env_path)
    else:
        load_dotenv()

    # Initialize Supabase client
    if not os.getenv("SUPABASE_URL") or not os.getenv("SUPABASE_KEY"):
        print("Missing Supabase credentials. Check your .env file.")
        return False

    supabase: Client = create_client(
        os.getenv("SUPABASE_URL"),
        os.getenv("SUPABASE_KEY")
    )

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

    # Fetch data in batches
    all_data = []
    batch_size = 1000
    offset = 0

    while True:
        try:
            response = supabase.table("daily_report").select("*").range(offset, offset + batch_size - 1).execute()
            batch = response.data
            if not batch:
                break
            all_data.extend(batch)
            if len(batch) < batch_size:
                break
            offset += batch_size
        except Exception as e:
            print(f"Error fetching data from daily_report: {e}")
            return False

    if not all_data:
        print("No data found in daily_report.")
        return False

    print(f"Fetched {len(all_data)} rows from daily_report.")

    # Convert to DataFrame and transform data
    df = pd.DataFrame(all_data)

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
    records = df.to_dict(orient="records")
    batch_size = 500

    print(f"Transformed data. Ready to insert {len(records)} records into current_report.")

    # Insert data in batches
    for i in range(0, len(records), batch_size):
        batch = records[i:i+batch_size]
        try:
            batch = [{k: (None if pd.isna(v) else v) for k, v in record.items()} for record in batch]
            response = supabase.table("current_report").insert(batch).execute()
            if not response.data:
                print(f"Error inserting batch {i//batch_size + 1}: {response.__dict__}")
            else:
                print(f"Batch {i//batch_size + 1} inserted successfully.")
        except Exception as e:
            print(f"Exception inserting batch {i//batch_size + 1}: {e}")
            return False

    print("Daily data transfer completed.")
    return True

if __name__ == "__main__":
    success = run_daily_data_transfer()
    sys.exit(0 if success else 1) 