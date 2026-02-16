import os
import pandas as pd
from supabase import create_client, Client
from typing import List, Dict
import logging
from datetime import datetime

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('data_upload.log')
    ]
)
logger = logging.getLogger(__name__)

# Supabase configuration
SUPABASE_URL = "https://iibczgyjzczqsecarqym.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImlpYmN6Z3lqemN6cXNlY2FycXltIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjU3MjEwNzMsImV4cCI6MjA4MTI5NzA3M30._drYhu8ajELkd19l7GtagD6yiSe46admrnqUrBkMAOU"
TABLE_NAME = "wms_stock"  # Change this to your actual table name

# File path
FILE_PATH = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Documents - Sales Dashboards (BI Solution)\WMS_stock\Stock+Report (5).xlsx"

# Expected table columns based on schema
TABLE_COLUMNS = [
    "Organization Name", "Organization Code", "Base name", "warehouseCode",
    "Wharea", "Location", "port", "destinationWarehouse", "EXW", "ETD",
    "ETA", "ATA", "ATD", "inboundDate", "Status of cargo logistics",
    "inboundStatus", "B/L NO", "vessel", "Release or not", "soldDate",
    "Release number", "Internal invoice number", "Type", "Container number",
    "Quantity", "Quantity (trust)", "Wattage", "Single-chip power",
    "Total wattage", "megawatt", "sku", "componentTechnology",
    "Status of  Goods Contract", "status", "Quality Status", "customer",
    "salesSupport", "sales support", "Region info（internal）", "deliveryTerms",
    "Currency", "Internal contract number", "internalSalesContract",
    "External contract number", "externalSalesContract",
    "External invoice number", "invNo", "countryCode",
    "Estimated Delivery Time", "Internal order number", "Component Model",
    "Back panel color", "borderColor", "junctionBox", "length", "voltage",
    "External dimensions", "soldWeek", "Inventory Age Days",
    "Inventory Age Months", "relatedTransactionCompany",
    "relatedTransactionTerm", "Internal Watt Unit Price",
    "externalRelatedPrice", "shipDate", "PO Number", "carbonFootprint",
    "Salesman(External)", "Region info(External)", "portCode", "portName",
    "borderSize", "batteryType", "glassThickness"
]


def initialize_supabase() -> Client:
    """Initialize and return Supabase client."""
    try:
        supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
        logger.info("Successfully connected to Supabase")
        return supabase
    except Exception as e:
        logger.error(f"Failed to connect to Supabase: {e}")
        raise


def read_excel_file(file_path: str) -> pd.DataFrame:
    """Read Excel file and return DataFrame."""
    try:
        logger.info(f"Reading Excel file: {file_path}")
        df = pd.read_excel(file_path)
        logger.info(f"Successfully read {len(df)} rows from Excel file")
        logger.info(f"Excel columns: {list(df.columns)}")
        return df
    except FileNotFoundError:
        logger.error(f"File not found: {file_path}")
        raise
    except Exception as e:
        logger.error(f"Error reading Excel file: {e}")
        raise


def map_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Map Excel columns to database columns."""
    logger.info("Mapping columns...")
    
    # Create a mapping dictionary for columns that might have different names
    # This will match Excel headers to database column names
    column_mapping = {}
    
    for excel_col in df.columns:
        # Try to find exact match first
        if excel_col in TABLE_COLUMNS:
            column_mapping[excel_col] = excel_col
        else:
            # Try case-insensitive match
            excel_col_lower = excel_col.lower().strip()
            for table_col in TABLE_COLUMNS:
                if table_col.lower().strip() == excel_col_lower:
                    column_mapping[excel_col] = table_col
                    break
    
    logger.info(f"Mapped {len(column_mapping)} columns")
    logger.info(f"Column mapping: {column_mapping}")
    
    # Rename columns based on mapping
    df_mapped = df.rename(columns=column_mapping)
    
    # Keep only columns that exist in the table schema
    available_columns = [col for col in df_mapped.columns if col in TABLE_COLUMNS]
    df_mapped = df_mapped[available_columns]
    
    logger.info(f"Final DataFrame has {len(available_columns)} columns")
    
    return df_mapped


def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """Clean and prepare data for upload."""
    logger.info("Cleaning data...")
    
    # Replace NaN with None for proper JSON serialization
    df = df.where(pd.notna(df), None)
    
    # Convert all data to string type (since all columns are VARCHAR in schema)
    for col in df.columns:
        df[col] = df[col].astype(str)
        # Replace 'nan' string with None
        df[col] = df[col].replace('nan', None)
        df[col] = df[col].replace('None', None)
    
    logger.info("Data cleaning completed")
    return df


def upload_in_batches(supabase: Client, df: pd.DataFrame, batch_size: int = 100) -> None:
    """Upload data to Supabase in batches."""
    total_rows = len(df)
    logger.info(f"Starting upload of {total_rows} rows in batches of {batch_size}")
    
    success_count = 0
    error_count = 0
    
    # Convert DataFrame to list of dictionaries
    records = df.to_dict('records')
    
    # Upload in batches
    for i in range(0, total_rows, batch_size):
        batch = records[i:i + batch_size]
        batch_num = (i // batch_size) + 1
        total_batches = (total_rows + batch_size - 1) // batch_size
        
        try:
            logger.info(f"Uploading batch {batch_num}/{total_batches} ({len(batch)} rows)")
            response = supabase.table(TABLE_NAME).insert(batch).execute()
            success_count += len(batch)
            logger.info(f"Batch {batch_num} uploaded successfully")
        except Exception as e:
            error_count += len(batch)
            logger.error(f"Error uploading batch {batch_num}: {e}")
            # Continue with next batch instead of stopping
            continue
    
    logger.info(f"Upload completed: {success_count} rows successful, {error_count} rows failed")


def clear_table(supabase: Client) -> None:
    """Clear all data from the table (optional - use with caution)."""
    try:
        logger.warning(f"Clearing all data from table: {TABLE_NAME}")
        # This is a destructive operation - uncomment only if needed
        # response = supabase.table(TABLE_NAME).delete().neq('id', 0).execute()
        # logger.info("Table cleared successfully")
        pass
    except Exception as e:
        logger.error(f"Error clearing table: {e}")
        raise


def main():
    """Main function to orchestrate the data upload process."""
    try:
        start_time = datetime.now()
        logger.info("="*60)
        logger.info("Starting data upload process")
        logger.info("="*60)
        
        # Initialize Supabase client
        supabase = initialize_supabase()
        
        # Read Excel file
        df = read_excel_file(FILE_PATH)
        
        if df.empty:
            logger.warning("Excel file is empty. Nothing to upload.")
            print("❌ Excel file is empty. Nothing to upload.")
            return
        
        # Map columns to database schema
        df_mapped = map_columns(df)
        
        if df_mapped.empty or len(df_mapped.columns) == 0:
            logger.error("No matching columns found between Excel and database schema")
            print("❌ No matching columns found between Excel and database schema")
            return
        
        # Clean data
        df_clean = clean_data(df_mapped)
        
        # Optional: Clear existing data (uncomment if needed)
        # clear_table(supabase)
        
        # Upload data in batches
        upload_in_batches(supabase, df_clean, batch_size=100)
        
        end_time = datetime.now()
        duration = (end_time - start_time).total_seconds()
        
        logger.info("="*60)
        logger.info(f"Data upload process completed in {duration:.2f} seconds")
        logger.info("="*60)
        
        # Final success message
        print(f"\n✅ Data upload completed successfully!")
        print(f"📊 Total rows processed: {len(df_clean)}")
        print(f"⏱️  Time taken: {duration:.2f} seconds")
        print(f"📁 Detailed logs saved to: data_upload.log\n")
        
    except Exception as e:
        logger.error(f"Fatal error in main process: {e}")
        print(f"\n❌ Error: {e}")
        print(f"📁 Check data_upload.log for details\n")
        raise


if __name__ == "__main__":
    main()

