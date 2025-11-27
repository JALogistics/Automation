"""
Script to transfer latest RNO and EU Stock reports
- Transfers latest Sales_RNO_Report as CSV
- Transfers latest EU_Stock_Report to destination folder
"""

import os
import glob
import shutil
import pandas as pd
from pathlib import Path
from datetime import datetime


def get_latest_file(folder_path, pattern="*"):
    """
    Get the latest file from a folder based on modification time
    
    Args:
        folder_path: Path to the folder to search
        pattern: File pattern to match (e.g., "*.xlsx")
    
    Returns:
        Path to the latest file or None if no files found
    """
    search_pattern = os.path.join(folder_path, pattern)
    files = glob.glob(search_pattern)
    
    if not files:
        print(f"No files found in {folder_path} matching pattern {pattern}")
        return None
    
    # Get the latest file based on modification time
    latest_file = max(files, key=os.path.getmtime)
    return latest_file


def transfer_rno_report():
    """
    Transfer the latest Sales_RNO_Report as CSV
    """
    print("\n" + "="*60)
    print("Processing Sales RNO Report")
    print("="*60)
    
    source_folder = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Logistics Reporting\000_Master_Query_Reports\Automation_DB\Sales_RNO_Report"
    dest_folder = r"C:\Users\DeepakSureshNidagund\JA Solar GmbH\Power BI Setup - PowerBISetup"
    dest_filename = "Latest_RNO_report.csv"
    
    # Ensure destination folder exists
    os.makedirs(dest_folder, exist_ok=True)
    
    # Get the latest file
    latest_file = get_latest_file(source_folder, "Sales_RNO_Report_*.xlsx")
    
    if not latest_file:
        print("ERROR: No Sales_RNO_Report files found!")
        return False
    
    print(f"Source file: {latest_file}")
    print(f"File size: {os.path.getsize(latest_file) / 1024:.2f} KB")
    print(f"Last modified: {datetime.fromtimestamp(os.path.getmtime(latest_file))}")
    
    try:
        # Read the Excel file
        print("Reading Excel file...")
        df = pd.read_excel(latest_file)
        
        # Save as CSV
        dest_path = os.path.join(dest_folder, dest_filename)
        print(f"Saving as CSV to: {dest_path}")
        df.to_csv(dest_path, index=False, encoding='utf-8-sig')
        
        print(f"✓ Successfully transferred RNO report")
        print(f"  Rows: {len(df)}, Columns: {len(df.columns)}")
        return True
        
    except Exception as e:
        print(f"ERROR: Failed to transfer RNO report - {str(e)}")
        return False


def transfer_eu_stock_report():
    """
    Transfer the latest EU_Stock_Report to destination folder
    """
    print("\n" + "="*60)
    print("Processing EU Stock Report")
    print("="*60)
    
    source_folder = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Logistics Reporting\000_Master_Query_Reports\Automation_DB\EU_Stock_report"
    dest_folder = r"C:\Users\DeepakSureshNidagund\JA Solar GmbH\Power BI Setup - PowerBISetup\EU_Stock_report"
    
    # Ensure destination folder exists
    os.makedirs(dest_folder, exist_ok=True)
    
    # Get the latest file (supports both .xlsx and .csv)
    latest_file = get_latest_file(source_folder, "Stock_Report_*.*")
    
    if not latest_file:
        print("ERROR: No EU Stock Report files found!")
        return False
    
    print(f"Source file: {latest_file}")
    print(f"File size: {os.path.getsize(latest_file) / 1024:.2f} KB")
    print(f"Last modified: {datetime.fromtimestamp(os.path.getmtime(latest_file))}")
    
    try:
        # Get the filename and destination path
        filename = os.path.basename(latest_file)
        dest_path = os.path.join(dest_folder, filename)
        
        # Copy the file
        print(f"Copying to: {dest_path}")
        shutil.copy2(latest_file, dest_path)
        
        print(f"✓ Successfully transferred EU Stock report")
        return True
        
    except Exception as e:
        print(f"ERROR: Failed to transfer EU Stock report - {str(e)}")
        return False


def main():
    """
    Main function to execute both transfers
    """
    print("\n" + "="*60)
    print("File Transfer Script - RNO & EU Stock Reports")
    print("="*60)
    print(f"Started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    results = []
    
    # Transfer RNO Report
    results.append(("Sales RNO Report", transfer_rno_report()))
    
    # Transfer EU Stock Report
    results.append(("EU Stock Report", transfer_eu_stock_report()))
    
    # Summary
    print("\n" + "="*60)
    print("Transfer Summary")
    print("="*60)
    for report_name, success in results:
        status = "✓ SUCCESS" if success else "✗ FAILED"
        print(f"{report_name}: {status}")
    
    print(f"\nCompleted at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*60)
    
    # Return exit code
    all_success = all(result[1] for result in results)
    return 0 if all_success else 1


if __name__ == "__main__":
    exit(main())

