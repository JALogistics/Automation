import os
import glob
import logging
from datetime import datetime
import re
import pandas as pd

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('database_connection.log'),
        logging.StreamHandler()
    ]
)

def get_latest_cdr_file(source_dir):
    """
    Get the latest CDR file from the source directory based on the date in filename.
    
    Args:
        source_dir (str): Path to the source directory containing CDR files
        
    Returns:
        str: Path to the latest CDR file or None if no matching files found
    """
    try:
        # Create the pattern for matching CDR files
        pattern = os.path.join(source_dir, "CDR_????-??-??.csv")
        
        # Get all matching files
        cdr_files = glob.glob(pattern)
        
        if not cdr_files:
            logging.warning(f"No CDR files found in {source_dir}")
            return None
        
        # Extract dates from filenames and find the latest one
        latest_file = max(cdr_files, key=lambda x: re.search(r'CDR_(\d{4}-\d{2}-\d{2})', x).group(1))
        
        logging.info(f"Latest CDR file found: {os.path.basename(latest_file)}")
        return latest_file
    
    except Exception as e:
        logging.error(f"Error while getting latest CDR file: {str(e)}")
        return None

def process_and_save_file(source_file, destination_dir):
    """
    Process the CSV file, apply filters, and save as CSV file.
    
    Args:
        source_file (str): Path to the source CSV file
        destination_dir (str): Path to the destination directory
        
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        if not source_file or not os.path.exists(source_file):
            logging.error("Source file does not exist")
            return False
            
        # Create destination directory if it doesn't exist
        os.makedirs(destination_dir, exist_ok=True)
        
        # Read the CSV file
        logging.info("Reading CSV file...")
        df = pd.read_csv(source_file)
        
        # Apply filters
        logging.info("Applying filters to the data...")
        filtered_df = df[
            (df['data_source'].str.lower() == 'current') & 
            (df['Current_Status'].str.lower() == 'in-stock')
        ]
        
        # Get the original filename
        base_filename = os.path.splitext(os.path.basename(source_file))[0]
        
        # Create the destination path with csv extension
        destination_path = os.path.join(destination_dir, f"{base_filename}.csv")
        
        # Save as CSV file
        logging.info("Saving filtered data to CSV...")
        filtered_df.to_csv(destination_path, index=False)
        
        logging.info(f"File successfully saved to: {destination_path}")
        return True
        
    except Exception as e:
        logging.error(f"Error while processing and saving file: {str(e)}")
        return False

def main():
    """Main execution function"""
    try:
        # Define source and destination directories
        source_dir = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Logistics Reporting\000_Master_Query_Reports\Automation_DB\CDR_Reports"
        dest_dir = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Documents - Sales Dashboards (BI Solution)\A_WMS Reports\CDR_Stock_Report"
        
        # Get the latest CDR file
        latest_file = get_latest_cdr_file(source_dir)
        
        if latest_file:
            # Process and save the file
            if process_and_save_file(latest_file, dest_dir):
                logging.info("File processing completed successfully")
            else:
                logging.error("Failed to process and save file")
        else:
            logging.error("No valid CDR file found to process")
            
    except Exception as e:
        logging.error(f"Error in main execution: {str(e)}")

if __name__ == "__main__":
    main()