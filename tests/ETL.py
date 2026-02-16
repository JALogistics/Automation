import os
import pandas as pd
from pathlib import Path
import pyodbc
from typing import List, Dict
import logging
from datetime import datetime
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Configure logging - only show errors
logging.basicConfig(
    level=logging.ERROR,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Configuration
SOURCE_FOLDER = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Logistics Reporting\0_Daily Reports_Warehouses (Excel files)"
TEMPLATE_FILE = r"C:\Users\DeepakSureshNidagund\Downloads\Reporting Application\Automation\automation\0_DailyReportTemplate (DO NOT DELETE,CHANGE).xlsx"
TABLE_NAME = 'daily_report'

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
        logger.error(error_msg)
        raise Exception(error_msg)

# Database column names (exact match with database schema)
DB_COLUMNS = [
    'Name JA partner', 'WH location', 'Type of WH (bonded/non)',
    'Container No.', 'Product type', 'Product reference', 'Port of Loading',
    'Port of destination', 'Inbound ref.', 'Import invoice', 'House B/l',
    'Bill of Lading', 'Shipping line', 'Vessel', 'ETD date POL',
    'ATD date POL', 'ETA date', 'ATA date', 'Import MRN', 'Import date',
    'Planned Inbound date', 'Inbound date',
    'Inbound duration days (Inbound date-ATA date+1)', 'Inbound Status',
    'Dev. Planned to Real in days Inbound date-Planned inbound date',  # Note: NO parentheses
    'Release date from port (ATA date)',
    'Contractual freetime for D&D combined', 'Free DM days', 'Free DT days',
    'Free DM days remained', 'Free DT days remained',
    'Container Returned date', 'Factory JASolar', 'Pallets', 'Piece',
    'Wattage', 'Total Wattage', 'MW', 'Stock Status', 'Stock age',
    'Release Number', 'Release type', 'Incoterm', 'Release date',
    'Internal Outbound ref', 'Outbound date', 'Outbound Status',
    'Agreed Delivery date', 'Delivery date', 'Storage time after release',
    'Delivery Duration', 'Dev. Between Agreed vs Real delivery date',
    'Sales Name', 'date CMR sent to JASolar', 'Customer Name',
    'Customer Country', 'Consignee name', 'Destination Address',
    'Destination Postal Code', 'Destination City', 'Destination Country',
    'Sales invoice', 'PTW / intermodel type',
    'Port fees (THC, ISPS, etc. )', 'DM cost', 'DT cost',
    'Port storage cost', 'Drayage costs (Port to WH)', 'Inbound costs',
    'Storage costs (fm IB to Today/OB)', 'Outbound costs',
    'Transport costs', 'Comments'
]


def get_template_headers(template_path: str) -> List[str]:
    """
    Read the template file to understand the structure, but return database column names.
    
    Args:
        template_path: Path to the template Excel file
        
    Returns:
        List of database column names
    """
    # We use the database column names directly to ensure exact match
    return DB_COLUMNS


def normalize_column_name(col: str) -> str:
    """
    Normalize column names for comparison by removing newlines, extra spaces, and converting to lowercase.
    Also handles special cases like missing parentheses.
    
    Args:
        col: Column name to normalize
        
    Returns:
        Normalized column name
    """
    if pd.isna(col):
        return ""
    # Remove newlines, carriage returns, and extra spaces
    cleaned = str(col).replace('\n', ' ').replace('\r', ' ').strip()
    # Replace multiple spaces with single space
    cleaned = ' '.join(cleaned.split())
    # Convert to lowercase for comparison
    normalized = cleaned.lower()
    
    # Handle special case: remove trailing parentheses for comparison
    # This helps match columns like "Dev. Planned to Real in days (Inbound date-Planned inbound date)" 
    # with "Dev. Planned to Real in days (Inbound date-Planned inbound date"
    normalized = normalized.rstrip(')')
    
    return normalized


def clean_column_name(col: str) -> str:
    """
    Clean column name by removing newlines and extra spaces but preserving case.
    
    Args:
        col: Column name to clean
        
    Returns:
        Cleaned column name
    """
    if pd.isna(col):
        return ""
    # Remove newlines and carriage returns, replace with space, then clean up spaces
    cleaned = str(col).replace('\n', ' ').replace('\r', ' ')
    # Replace multiple spaces with single space
    cleaned = ' '.join(cleaned.split())
    return cleaned.strip()


def map_columns(df: pd.DataFrame, template_headers: List[str]) -> pd.DataFrame:
    """
    Map DataFrame columns to template headers.
    
    Args:
        df: DataFrame to map
        template_headers: List of template column headers
        
    Returns:
        DataFrame with columns mapped to template headers
    """
    # Create normalized mapping
    normalized_template = {normalize_column_name(col): col for col in template_headers}
    
    # Create column mapping
    column_mapping = {}
    for col in df.columns:
        # Clean the source column name first
        cleaned_col = clean_column_name(col)
        normalized_col = normalize_column_name(cleaned_col)
        if normalized_col in normalized_template:
            column_mapping[col] = normalized_template[normalized_col]
    
    # Rename columns based on mapping
    df_mapped = df.rename(columns=column_mapping)
    
    # Add missing columns from template with None values
    for template_col in template_headers:
        if template_col not in df_mapped.columns:
            df_mapped[template_col] = None
    
    # Reorder columns to match template
    df_mapped = df_mapped[template_headers]
    
    return df_mapped


def read_excel_files(source_folder: str, template_headers: List[str]) -> tuple[pd.DataFrame, int]:
    """
    Read all Excel files from source folder and combine them.
    
    Args:
        source_folder: Path to folder containing Excel files
        template_headers: List of template column headers
        
    Returns:
        Tuple of (Combined DataFrame with all data, number of files processed)
    """
    all_data = []
    folder_path = Path(source_folder)
    
    if not folder_path.exists():
        logger.error(f"Source folder does not exist: {source_folder}")
        raise FileNotFoundError(f"Source folder not found: {source_folder}")
    
    # Get all Excel files
    excel_files = list(folder_path.glob('*.xlsx')) + list(folder_path.glob('*.xls'))
    excel_files = [f for f in excel_files if not f.name.startswith('~$')]  # Exclude temp files
    
    total_files = len(excel_files)
    
    files_processed = 0
    for file_path in excel_files:
        try:
            # Read Excel file
            df = pd.read_excel(file_path)
            
            # Skip empty files
            if df.empty:
                continue
            
            # Map columns to template
            df_mapped = map_columns(df, template_headers)
            
            # Add metadata
            df_mapped['source_file'] = file_path.name
            df_mapped['processed_at'] = datetime.now().isoformat()
            
            all_data.append(df_mapped)
            files_processed += 1
            
        except Exception as e:
            logger.error(f"Error processing {file_path.name}: {e}")
            continue
    
    if not all_data:
        return pd.DataFrame(), 0
    
    # Combine all DataFrames
    combined_df = pd.concat(all_data, ignore_index=True)
    
    return combined_df, files_processed


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


def prepare_data_for_database(df: pd.DataFrame) -> pd.DataFrame:
    """
    Prepare DataFrame for SQL Server insertion.
    
    Args:
        df: DataFrame to prepare
        
    Returns:
        DataFrame ready for SQL Server
    """
    # Replace NaN with None for proper NULL handling in database
    df_clean = df.where(pd.notnull(df), None)
    
    # Transform date columns
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
        if col in df_clean.columns:
            df_clean[col] = convert_date_column(df_clean[col])
    
    # Convert all values to strings as per table schema, except source metadata and dates
    for col in df_clean.columns:
        if col not in ['source_file', 'processed_at'] and col not in date_columns:
            df_clean[col] = df_clean[col].astype(str).replace('None', None)
    
    # Remove source_file and processed_at as they're not in the schema
    df_clean = df_clean.drop(columns=['source_file', 'processed_at'], errors='ignore')
    
    # Fix the problematic column name - template has parentheses, DB doesn't
    if 'Dev. Planned to Real in days (Inbound date-Planned inbound date)' in df_clean.columns:
        df_clean = df_clean.rename(columns={
            'Dev. Planned to Real in days (Inbound date-Planned inbound date)': 
            'Dev. Planned to Real in days Inbound date-Planned inbound date'
        })
    elif 'Dev. Planned to Real in days (Inbound date-Planned inbound date' in df_clean.columns:
        df_clean = df_clean.rename(columns={
            'Dev. Planned to Real in days (Inbound date-Planned inbound date': 
            'Dev. Planned to Real in days Inbound date-Planned inbound date'
        })
    
    return df_clean


def upload_to_database(df: pd.DataFrame, batch_size: int = 1000) -> int:
    """
    Upload DataFrame to SQL Server in batches.
    
    Args:
        df: DataFrame to upload
        batch_size: Number of records per batch
        
    Returns:
        Number of successfully uploaded records
    """
    try:
        # Initialize SQL Server connection
        conn = get_sql_connection()
        print("Connected to SQL Server successfully.")
        
        # Upload data using pandas to_sql
        df.to_sql(
            TABLE_NAME,
            conn,
            if_exists='append',
            index=False,
            chunksize=batch_size,
            method='multi'
        )
        
        uploaded_count = len(df)
        
        # Close connection
        conn.close()
        
        return uploaded_count
        
    except Exception as e:
        logger.error(f"Error uploading to SQL Server: {e}")
        raise


def main():
    """
    Main ETL process.
    """
    try:
        # Step 1: Get template headers
        template_headers = get_template_headers(TEMPLATE_FILE)
        
        # Step 2: Read and combine Excel files
        combined_df, files_processed = read_excel_files(SOURCE_FOLDER, template_headers)
        
        if combined_df.empty:
            print("No data to upload")
            return
        
        # Step 3: Prepare data for SQL Server
        prepared_df = prepare_data_for_database(combined_df)
        
        # Step 4: Upload to SQL Server
        uploaded_count = upload_to_database(prepared_df)
        
        # Final Summary
        print(f"\n{'='*80}")
        print(f"ETL PROCESS COMPLETED SUCCESSFULLY")
        print(f"{'='*80}")
        print(f"Total Files Processed: {files_processed}")
        print(f"Total Records: {len(prepared_df)}")
        print(f"Successfully Uploaded Records: {uploaded_count}")
        print(f"{'='*80}\n")
        
    except Exception as e:
        print(f"\nETL process failed: {e}\n")
        logger.error(f"ETL process failed: {e}")
        raise


if __name__ == "__main__":
    main()
