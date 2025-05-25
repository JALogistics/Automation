import pandas as pd
import os
import re
import glob
import datetime
from typing import Tuple, Optional
from pathlib import Path

def remove_special_chars(val):
    if pd.isnull(val):
        return val
    # Remove special characters and then strip leading/trailing spaces
    return re.sub(r'[^A-Za-z0-9 ]+', '', str(val)).strip()

def read_eu_stock_report(file_path: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Read and process the EU stock report Excel file.
    
    Args:
        file_path (str): Path to the Excel file
        
    Returns:
        Tuple[pd.DataFrame, pd.DataFrame]: Processed dataframe and Released data
        
    Raises:
        FileNotFoundError: If the Excel file is not found
        Exception: For other processing errors
    """
    try:
        # Check if file exists
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Excel file not found at: {file_path}")
            
        # Read the Excel file
        df = pd.read_excel(file_path, sheet_name="Summary-Europe")
        
        # 1. Clean specified columns
        columns_to_clean = ['Container', 'DN', 'Inv.']
        for col in columns_to_clean:
            if col in df.columns:
                df[col] = df[col].apply(remove_special_chars)

        # 2. Rename the columns
        df = df.rename(columns={
            'Inv.': 'Invoice Number',
            'Container': 'Container Number',
            'DN': 'Release Number'
        })

        # 3. Remove specified columns
        columns_to_remove = [
            'Factory', 'Related transaction company', 'Related transaction Term', 'Currency', 'Inv No.', 'C2 --> C1 Date',
            'Handover Date', 'Contractual Delivery Week', 'Country Code', '状态', 'Internal related price', 'Battery type',
            'Border Color', 'Junction box', 'length', 'Voltage', 'Storage duration', 'Sold（Week）', 'Storage duration（days）',
            'original WH', 'Warehouse after transfer', 'ETD month', 'Sold month', 'outbound quantity', 'Rest quantity',
            '是否为5.0组件技术(间隙贴膜)''Released on the sea', 'Booking No.', 'EWX Week', 'Type.2', 'Auxiliary column', 
            '成品料号', 'Inv&type', '是否签收', '型号', 'Unnamed: 65', 'Unnamed: 66', 'Unnamed: 67','Unnamed: 68', 'Unnamed: 69'
        ]
        
        # Drop the specified columns
        df = df.drop(columns=columns_to_remove, errors='ignore')

        # Create a new column 'Ref1' as concatenation of 'Release Number' and 'Container Number'
        if 'Release Number' in df.columns and 'Container Number' in df.columns:
            df['Ref1'] = df['Release Number'].astype(str) + df['Container Number'].astype(str)
        else:
            df['Ref1'] = None

        # Separate data into two datasets based on 'Release Number' and 'Not Released'
        Released_data = df[
            (df['Release Number'].notna() & (df['Release Number'].astype(str).str.strip() != "")) & 
            (df['Sold Date'].notna() & (df['Sold Date'].astype(str).str.strip() != ""))
        ]
        
        return df, Released_data
        
    except Exception as e:
        print(f"An error occurred while processing the EU stock report: {e}")
        raise

def read_cdr_and_merge(reports_dir: str, released_data: pd.DataFrame) -> pd.DataFrame:
    """
    Read the latest CDR report and merge it with Released data.
    
    Args:
        reports_dir (str): Directory containing CDR reports
        released_data (pd.DataFrame): DataFrame containing Released data
        
    Returns:
        pd.DataFrame: Merged and processed DataFrame
        
    Raises:
        FileNotFoundError: If no CDR reports are found
        Exception: For other processing errors
    """
    try:
        # Check if directory exists
        if not os.path.exists(reports_dir):
            print(f"Directory does not exist: {reports_dir}")
            print("Creating directory...")
            os.makedirs(reports_dir, exist_ok=True)

        # Try different patterns to find CDR reports
        patterns = [
            "CDR_*.csv",      # Primary pattern
            "*CDR*.csv",      # Secondary pattern - any file with CDR in the name
            "*.csv"           # Fallback pattern - any CSV file
        ]

        report_files = []
        for pattern in patterns:
            full_pattern = os.path.join(reports_dir, pattern)
            found_files = glob.glob(full_pattern)
            if found_files:
                report_files = found_files
                break

        if not report_files:
            raise FileNotFoundError(f"No CSV files found in {reports_dir}")

        try:
            # Try to get the latest file based on date in filename
            latest_file = max(report_files, key=lambda x: 
                            datetime.datetime.strptime(os.path.basename(x).split('_')[1].split('.')[0], 
                                                    "%Y-%m-%d"))
        except (IndexError, ValueError):
            print("Could not parse dates from filenames, using file modification time instead.")
            latest_file = max(report_files, key=os.path.getmtime)

        print(f"\nUsing latest CDR report: {os.path.basename(latest_file)}")
        print(f"Full path: {latest_file}")

        # Read the CDR file
        cdr = pd.read_csv(latest_file)

        # Select required columns from CDR
        cdr_columns = ["Ref1", "Current_Status", "Outbound date", "Agreed Delivery date", "Delivery date", "Delivery_Status"]
        cdr_selected = cdr[cdr_columns]

        # Merge the Released data with CDR data
        rno = pd.merge(
            released_data,
            cdr_selected,
            on="Ref1",
            how="left"  # Using left join to keep all records from Released data
        )

        # Process piece quantities from CDR
        cdr['Piece'] = pd.to_numeric(cdr['Piece'], errors='coerce')

        # Create separate summaries for each status
        status_summaries = {
            'On Sea': 'On_Sea_Pcs',
            'In-Stock': 'In_Stock_Pcs',
            'Outbounded': 'Outbounded_Pcs'
        }

        # Merge each status summary
        for status, col_name in status_summaries.items():
            summary = (cdr[cdr['Current_Status'] == status]
                      .groupby('Ref1')['Piece']
                      .sum()
                      .reset_index()
                      .rename(columns={'Piece': col_name}))
            
            rno = pd.merge(rno, summary, on="Ref1", how="left")
            rno[col_name] = rno[col_name].fillna(0)

        return rno

    except Exception as e:
        print(f"An error occurred while processing CDR report: {e}")
        raise

if __name__ == "__main__":
    # File paths
    eu_stock_path = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Documents - Sales Dashboards (BI Solution)\Y_EU Report\Europe Stock 最新版.xlsx"
    cdr_reports_dir = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Logistics Reporting\000_Master_Query_Reports\Automation_DB\CDR_Reports"

    # Read EU stock report
    df, released_data = read_eu_stock_report(eu_stock_path)
    
    # Read CDR and merge with released data
    final_report = read_cdr_and_merge(cdr_reports_dir, released_data)
    
    print("Processing completed successfully")