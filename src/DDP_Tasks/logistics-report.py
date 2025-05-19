import os
import sys
import glob
from pathlib import Path
from typing import List, Dict, Any
import datetime
import pandas as pd
from dotenv import load_dotenv

def generate_logistics_report():
    """
    Main function to generate logistics report by processing CDR data.
    Handles the entire workflow from finding the latest report to generating the final Excel output.
    """
    
    def load_environment():
        """Load environment variables from .env file"""
        env_path = Path(__file__).parents[2] / ".env"
        if env_path.exists():
            load_dotenv(env_path)
        else:
            load_dotenv()

    def find_latest_cdr_report() -> str:
        """
        Find the most recent CDR report file in the specified directory
        
        Returns:
            Path to the latest CDR report file
        """
        reports_dir = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Logistics Reporting\000_Master_Query_Reports\Automation_DB\CDR_Reports"
        
        # Check if directory exists
        if not os.path.exists(reports_dir):
            print(f"Directory does not exist: {reports_dir}")
            print("Creating directory...")
            os.makedirs(reports_dir, exist_ok=True)
            sys.exit(1)
        
        # List all files in the directory for debugging
        print(f"Listing all files in {reports_dir}:")
        all_files = os.listdir(reports_dir)
        for file in all_files:
            print(f"  - {file}")
        
        patterns = [
            "CDR_*.csv",  # Primary pattern
            "*CDR*.csv",  # Secondary pattern - any file with CDR in the name
            "*.csv"       # Fallback pattern - any CSV file
        ]
        
        report_files = []
        for pattern in patterns:
            full_pattern = os.path.join(reports_dir, pattern)
            found_files = glob.glob(full_pattern)
            if found_files:
                print(f"Found {len(found_files)} files with pattern '{pattern}'")
                report_files = found_files
                break
        
        if not report_files:
            print(f"No CSV files found in {reports_dir}")
            print("Please make sure there are CDR report files in the directory.")
            sys.exit(1)
        
        try:
            latest_file = max(report_files, key=lambda x: 
                            datetime.datetime.strptime(os.path.basename(x).split('_')[1].split('.')[0], 
                                                    "%Y-%m-%d"))
        except (IndexError, ValueError):
            print("Could not parse dates from filenames, using file modification time instead.")
            latest_file = max(report_files, key=os.path.getmtime)
        
        print(f"Using latest file: {latest_file}")
        return latest_file

    def read_cdr_data(file_path: str) -> pd.DataFrame:
        """
        Read data from the CDR report file
        
        Args:
            file_path: Path to the CSV or Excel file
            
        Returns:
            DataFrame containing the CDR report data
        """
        try:
            if file_path.lower().endswith('.csv'):
                encodings = ['utf-8', 'latin1', 'iso-8859-1']
                delimiters = [',', ';', '\t']
                
                for encoding in encodings:
                    for delimiter in delimiters:
                        try:
                            print(f"Trying to read CSV with encoding: {encoding}, delimiter: '{delimiter}'")
                            df = pd.read_csv(file_path, encoding=encoding, sep=delimiter)
                            print(f"Successfully read data from CDR report: {len(df)} records")
                            return df
                        except Exception as e:
                            print(f"Failed with encoding: {encoding}, delimiter: '{delimiter}', error: {e}")
                            continue
                
                print("Trying pandas auto-detection for CSV...")
                df = pd.read_csv(file_path)
                print(f"Successfully read data from CDR report: {len(df)} records")
                return df
            else:
                df = pd.read_excel(file_path)
                print(f"Successfully read data from CDR report: {len(df)} records")
                return df
        except Exception as e:
            print(f"Error reading CDR report: {e}")
            sys.exit(1)

    def format_date_columns(df: pd.DataFrame) -> pd.DataFrame:
        """
        Convert date columns to proper datetime format
        
        Args:
            df: DataFrame containing the columns to convert
            
        Returns:
            DataFrame with formatted date columns
        """
        date_columns = [
            'Release date', 'Outbound Date', 'Outbound date', 'Inbound date', 
            'Agreed Delivery date', 'Delivery date'
        ]
        
        for col in [c for c in date_columns if c in df.columns]:
            try:
                df[col] = pd.to_datetime(df[col], errors='coerce')
            except Exception as e:
                print(f"Error converting '{col}' to date format: {e}")
        
        return df

    print("Starting logistics report data extraction...")
    
    try:
        # Load environment variables
        load_environment()
        
        # Find and read the latest CDR report
        latest_file = find_latest_cdr_report()
        raw_df = read_cdr_data(latest_file)
        
        # Print the size of the data
        print(f"Total records in CDR report: {len(raw_df)}")
        
        # Format date columns first
        raw_df = format_date_columns(raw_df)
        
        # Apply transformations for status distribution reporting
        print("\nApplying data transformations...")
        if 'Current_Status' in raw_df.columns:
            status_counts = raw_df['Current_Status'].value_counts()
            print("\nStatus distribution:")
            for status, count in status_counts.items():
                megawattage = 0
                if 'MegaWattage' in raw_df.columns:
                    raw_df['MegaWattage_numeric'] = pd.to_numeric(raw_df['MegaWattage'], errors='coerce')
                    megawattage = raw_df[raw_df['Current_Status'] == status]['MegaWattage_numeric'].sum(skipna=True)
                print(f"  {status}: {count} records, Total MegaWattage: {float(megawattage):.2f} MW")

        # Remove rows where Outbound_status is "Outbounded" and Outbound Date is in year 2024,2023,2022
        outbound_date_col = 'Outbound Date' if 'Outbound Date' in raw_df.columns else 'Outbound date'
        if outbound_date_col in raw_df.columns and 'Outbound_status' in raw_df.columns:
            raw_df = raw_df[~((raw_df['Outbound_status'] == "Outbounded") & 
                             (raw_df[outbound_date_col].dt.year.isin([2024, 2023, 2022])))]

        # Create a copy of the raw DataFrame for transformations
        transformed_df = raw_df.copy()
        
        # Apply filters
        if 'Outbound_status' in transformed_df.columns:
            transformed_df = transformed_df[transformed_df['Outbound_status'].isin(["outbounded", "outbound-planned"])]   
        
        if 'Release Number' in transformed_df.columns:
            transformed_df = transformed_df[transformed_df['Release Number'].astype(str).str.startswith(('2', '5'))]
            
        if outbound_date_col in transformed_df.columns:
            current_month = datetime.datetime.now().month
            current_year = datetime.datetime.now().year
            
            transformed_df = transformed_df[
                # (transformed_df[outbound_date_col].dt.month == current_month) & 
                (transformed_df[outbound_date_col].dt.year == current_year)
            ]
            print(f"Filtered for outbound dates in current month only: {current_month}/{current_year}")
        
        # Remove specified columns
        columns_to_remove = [
            'House B/l', 'Bill of Lading', 'Shipping line', 'Vessel', 'ETD date POL',
            'ATD date POL', 'ETA date', 'ATA date', 'Import MRN', 'Import date',
            'Planned Inbound date', 'Inbound duration days (Inbound date-ATA date+1)', 'Inbound Status',
            'Dev. Planned to Real in days (Inbound date-Planned inbound date',
            'Release date from port (ATA date)', 'Contractual freetime for D&D combined', 
            'Free DM days', 'Free DT days', 'Free DM days remained', 'Free DT days remained',
            'Container Returned date', 'Factory JASolar', 'Storage time after release',
            'Delivery Duration', 'Dev. Between Agreed vs Real delivery date',
            'Sales Name', 'date CMR sent to JASolar', 'PTW / intermodel type',
            'Port fees (THC, ISPS, etc. )', 'DM cost', 'DT cost',
            'Port storage cost', 'Drayage costs (Port to WH)', 'Inbound costs',
            'Storage costs (fm IB to Today/OB)', 'Outbound costs', 'Transport costs','MegaWattage_numeric',
            'Ref1', 'Ref2','created_at','id','Stock Status', 'Stock age', 'Internal Outbound ref',
            'Destination Address', 'Destination Postal Code','Status',
        ]
        
        columns_to_remove = [col for col in columns_to_remove if col in transformed_df.columns]
        if columns_to_remove:
            transformed_df = transformed_df.drop(columns=columns_to_remove)
            
        print("\n")

        # Remove rows with specific years
        if outbound_date_col in raw_df.columns:
            raw_df = raw_df[~raw_df[outbound_date_col].dt.year.isin([2022, 2023, 2024])]
            print(f"Removed rows with {outbound_date_col} in years 2022, 2023, 2024 from raw_df. Remaining rows: {len(raw_df)}")

        if outbound_date_col in transformed_df.columns:
            transformed_df = transformed_df[~transformed_df[outbound_date_col].dt.year.isin([2022, 2023, 2024])]
            print(f"Removed rows with {outbound_date_col} in years 2022, 2023, 2024 from transformed_df. Remaining rows: {len(transformed_df)}")
        
        # Print status distributions
        print("\n")
        if 'Current_Status' in raw_df.columns:
            status_counts = raw_df['Current_Status'].value_counts()
            print("\nCurrent status after the filters:")
            for status, count in status_counts.items():
                megawattage = 0
                if 'MegaWattage' in raw_df.columns:
                    raw_df['MegaWattage_numeric'] = pd.to_numeric(raw_df['MegaWattage'], errors='coerce')
                    megawattage = raw_df[raw_df['Current_Status'] == status]['MegaWattage_numeric'].sum(skipna=True)
                print(f"  {status}: {count} records, Total MegaWattage: {float(megawattage):.2f} MW")
        
        if 'Outbound_status' in transformed_df.columns:
            status_counts = transformed_df['Outbound_status'].value_counts()
            print("\nCurrent status after the filters:")
            for status, count in status_counts.items():
                megawattage = 0
                if 'MegaWattage' in transformed_df.columns:
                    transformed_df['MegaWattage_numeric'] = pd.to_numeric(transformed_df['MegaWattage'], errors='coerce')
                    megawattage = transformed_df[transformed_df['Outbound_status'] == status]['MegaWattage_numeric'].sum(skipna=True)
                print(f"  {status}: {count} records, Total MegaWattage: {float(megawattage):.2f} MW")

            # Add current month summary
            current_month = datetime.datetime.now().month
            current_year = datetime.datetime.now().year
            current_month_df = transformed_df[
                (transformed_df[outbound_date_col].dt.month == current_month) & 
                (transformed_df[outbound_date_col].dt.year == current_year)
            ]
            
            current_month_status = current_month_df['Outbound_status'].value_counts()
            print(f"\nCurrent month ({datetime.datetime.now().strftime('%B %Y')}) status summary:")
            for status, count in current_month_status.items():
                megawattage = 0
                if 'MegaWattage' in current_month_df.columns:
                    current_month_df['MegaWattage_numeric'] = pd.to_numeric(current_month_df['MegaWattage'], errors='coerce')
                    megawattage = current_month_df[current_month_df['Outbound_status'] == status]['MegaWattage_numeric'].sum(skipna=True)
                print(f"  {status}: {count} records, Total MegaWattage: {float(megawattage):.2f} MW")

        # Save to Excel
        current_date = datetime.datetime.now().strftime("%Y-%m-%d")
        output_dir = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Logistics Reporting\000_Master_Query_Reports\Automation_DB\Logistics_Reports"
        output_file = os.path.join(output_dir, f"logistics-report_{current_date}.xlsx")
        
        os.makedirs(output_dir, exist_ok=True)
        
        if os.path.exists(output_file):
            os.remove(output_file)
            print(f"Deleted existing file: {output_file}")
            
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            raw_df.to_excel(writer, sheet_name='CDR Report', index=False)
            transformed_df.to_excel(writer, sheet_name='Outbound_logistics_report', index=False)
            
            workbook = writer.book
            date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})
            
            # Format date columns in both sheets
            for sheet_name, df in [('CDR Report', raw_df), ('Outbound_logistics_report', transformed_df)]:
                worksheet = writer.sheets[sheet_name]
                for col_num, col_name in enumerate(df.columns):
                    if col_name in ['Release date', 'Outbound Date', 'Inbound date', 'Agreed Delivery date', 'Delivery date']:
                        worksheet.set_column(col_num, col_num, 18, date_format)
        
        print(f"\nExcel workbook successfully saved at: {output_file}")
        print("Logistics report data extraction completed successfully.")
        
    except Exception as e:
        print(f"Error in logistics report generation: {e}")
        sys.exit(1)

if __name__ == "__main__":
    generate_logistics_report() 