import os
import sys
from pathlib import Path
import pandas as pd
from supabase import create_client, Client
from dotenv import load_dotenv
from datetime import datetime

def convert_date_column(series: pd.Series) -> pd.Series:
    """
    Convert date column handling both string dates and Excel serial numbers.
    
    Args:
        series: Pandas Series containing date values
        
    Returns:
        Series with properly formatted date strings (YYYY-MM-DD)
    """
    def parse_mixed_date(date_val):
        if pd.isna(date_val) or date_val is None or date_val == '':
            return None
        
        # If it's already a datetime object, format it
        if isinstance(date_val, (pd.Timestamp, datetime)):
            return date_val.strftime('%Y-%m-%d')
        
        # If it's a number, treat as Excel serial date
        if isinstance(date_val, (int, float)):
            try:
                # Excel's epoch starts at 1899-12-30
                converted_date = pd.to_datetime(date_val, origin='1899-12-30', unit='D')
                return converted_date.strftime('%Y-%m-%d')
            except:
                return None
        
        # Otherwise, try to parse as string
        try:
            parsed_date = pd.to_datetime(date_val, format='mixed', dayfirst=True, errors='coerce')
            if pd.notna(parsed_date):
                return parsed_date.strftime('%Y-%m-%d')
        except:
            pass
        
        return None
    
    return series.apply(parse_mixed_date)


def fetch_all_data(supabase: Client, table_name: str) -> list:
    """Fetch all data from a table in batches."""
    all_data = []
    batch_size = 1000
    offset = 0
    
    while True:
        try:
            response = supabase.table(table_name).select("*").range(offset, offset + batch_size - 1).execute()
            batch = response.data
            if not batch:
                break
            all_data.extend(batch)
            if len(batch) < batch_size:
                break
            offset += batch_size
            print(f"Fetched {len(all_data)} rows from {table_name} so far...")
        except Exception as e:
            print(f"Error fetching data from {table_name}: {e}")
            raise
    
    return all_data

def save_to_csv(df: pd.DataFrame, directory: str, filename: str) -> bool:
    """Save DataFrame to CSV at specified location."""
    try:
        os.makedirs(directory, exist_ok=True)
        filepath = os.path.join(directory, filename)
        print(f"Saving combined data to {filepath}...")
        df.to_csv(filepath, index=False)
        print(f"Successfully saved combined data to {filepath}")
        return True
    except Exception as e:
        print(f"Error saving to {filepath}: {e}")
        return False

def create_consolidated_report():
    """Main function to create consolidated report from current and archive data."""
    try:
        # Load environment variables
        env_path = Path(__file__).parents[1] / ".env"
        if env_path.exists():
            load_dotenv(env_path)
        else:
            load_dotenv()

        # Initialize Supabase client
        if not os.getenv("SUPABASE_URL") or not os.getenv("SUPABASE_KEY"):
            raise ValueError("Missing Supabase credentials. Check your .env file.")

        supabase: Client = create_client(
            os.getenv("SUPABASE_URL"),
            os.getenv("SUPABASE_KEY")
        )

        print("Starting consolidated report creation process.")

        # Fetch data from both tables
        print("Fetching data from current_report table...")
        current_data = fetch_all_data(supabase, "current_report")
        print(f"Fetched {len(current_data)} total rows from current_report.")

        print("Fetching data from archive_data table...")
        archive_data = fetch_all_data(supabase, "archive_data")
        print(f"Fetched {len(archive_data)} total rows from archive_data.")

        # Convert to DataFrames
        current_df = pd.DataFrame(current_data)
        archive_df = pd.DataFrame(archive_data)

        # Add source column
        if not current_df.empty:
            current_df["data_source"] = "current"
        if not archive_df.empty:
            archive_df["data_source"] = "archive"

        # Combine the DataFrames
        print("Combining data from both tables...")
        combined_df = pd.concat([current_df, archive_df], ignore_index=True)

        if combined_df.empty:
            print("No data found in either table. Exiting.")
            return

        print(f"Combined data has {len(combined_df)} total rows.")
        
        # Transform date columns
        print("Transforming date columns...")
        date_columns = [
            'Outbound date',
            'Agreed Delivery date',
            'Delivery date',
            'ETD date POL',
            'ATD date POL',
            'ETA date',
            'ATA date',
            'Import date',
            'Planned Inbound date',
            'Inbound date',
            'Release date from port (ATA date)',
            'Container Returned date',
            'Release date',
            'date CMR sent to JASolar'
        ]
        
        for col in date_columns:
            if col in combined_df.columns:
                combined_df[col] = convert_date_column(combined_df[col])
                print(f"Transformed {col}")
        
        print("Date transformation completed.")

        # Prepare output paths
        current_date = datetime.now().strftime("%Y-%m-%d")
        output_dir1 = r"C:\Users\DeepakSureshNidagund\Downloads\Reporting Application\Automation\automation\tests"
        output_dir2 = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Logistics Reporting\000_Master_Query_Reports\Automation_DB\CDR_Reports"
        output_dir3 = r"C:\Users\DeepakSureshNidagund\JA Solar GmbH\Power BI Setup - PowerBISetup\CDR"

        # Save to both locations
        save_to_csv(combined_df, output_dir1, "CDR.csv")
        save_to_csv(combined_df, output_dir2, f"CDR_{current_date}.csv")
        save_to_csv(combined_df, output_dir3, f"CDR_{current_date}.csv")

        # Prepare data for Supabase
        combined_df = combined_df.where(pd.notnull(combined_df), None)
        records = combined_df.to_dict(orient="records")
        batch_size = 500

        # Delete existing records
        try:
            print("Deleting existing records from consolidated_report table...")
            supabase.table("consolidated_report").delete().neq("id", 0).execute()
            print("Existing records deleted successfully.")
        except Exception as e:
            print(f"Error deleting existing records: {e}")

        # Insert new records in batches
        print(f"Ready to insert {len(records)} records into consolidated_report.")
        for i in range(0, len(records), batch_size):
            batch = records[i:i+batch_size]
            try:
                batch = [{k: (None if pd.isna(v) else v) for k, v in record.items()} for record in batch]
                response = supabase.table("consolidated_report").insert(batch).execute()
                if not response.data:
                    print(f"Error inserting batch {i//batch_size + 1}: {response.__dict__}")
                else:
                    print(f"Batch {i//batch_size + 1} inserted successfully.")
            except Exception as e:
                print(f"Exception inserting batch {i//batch_size + 1}: {e}")

        print("Consolidated report creation completed.")

    except Exception as e:
        print(f"Error in create_consolidated_report: {e}")
        raise

if __name__ == "__main__":
    try:
        create_consolidated_report()
    except Exception as e:
        print(f"Script failed: {e}")
        sys.exit(1) 