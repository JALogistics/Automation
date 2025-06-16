import pandas as pd
import numpy as np
import os
import glob
from datetime import datetime
import openpyxl
from pathlib import Path

# Source file path
europe_stock_path = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Documents - Sales Dashboards (BI Solution)\Y_EU Report\Europe Stock 最新版.xlsx"

# def load_europe_stock_data(file_path):
#     """
#     Load and process the Europe Stock data from Excel file.
#     """
#     try:
#         if not os.path.exists(file_path):
#             raise FileNotFoundError(f"Excel file not found at: {file_path}")
            
#         df = pd.read_excel(file_path, sheet_name="Summary-Europe")
#         return df
#     except Exception as e:
#         print(f"An error occurred while loading Europe stock data: {e}")
#         return None
def load_europe_stock_data(file_path):
    """
    Load and process the Europe Stock data from Excel file.
    """
    try:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Excel file not found at: {file_path}")
        
        # Add error handling for dates
        excel_options = {
            'sheet_name': "Summary-Europe",
            'na_values': ['#N/A', '#N/A N/A', '#NA', '-1.#IND', '-1.#QNAN', 
                         'N/A', 'n/a', 'NA', '<NA>', '#VALUE!'],
            'keep_default_na': True,
            'dtype': {'Sold Date': 'object'}  # Read date columns as object first
        }
            
        df = pd.read_excel(file_path, **excel_options)
        
        # Convert known date columns to datetime with error handling
        date_columns = ['Sold Date']  # Add other date columns if needed
        for col in date_columns:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
        
        return df
    except Exception as e:
        print(f"An error occurred while loading Europe stock data: {e}")
        return None

def clean_and_prepare_data(df):
    """
    Clean and prepare the data by removing special characters and renaming columns.
    """
    try:
        # Remove special characters from specified columns
        def remove_special_chars(val):
            if pd.isnull(val):
                return val
            return ''.join(c for c in str(val) if c.isalnum() or c.isspace()).strip()

        columns_to_clean = ['Container', 'DN', 'Inv.']
        for col in columns_to_clean:
            if col in df.columns:
                df[col] = df[col].apply(remove_special_chars)

        # Rename columns
        df = df.rename(columns={
            'Inv.': 'Invoice Number',
            'Container': 'Container Number',
            'DN': 'Release Number'
        })

        # Remove specified columns
        columns_to_remove = [
            'Factory', 'Related transaction company', 'Related transaction Term', 'Currency', 
            'Inv No.', 'C2 --> C1 Date', 'Handover Date', 'Contractual Delivery Week', 
            'Country Code', '状态', 'Internal related price', 'Battery type', 'Border Color',
            'Junction box', 'length', 'Voltage', 'Storage duration', 'Sold（Week）', 
            'Storage duration（days）', 'original WH', 'Warehouse after transfer', 'ETD month',
            'Sold month', 'outbound quantity', 'Rest quantity', 'Released on the sea',
            'Booking No.', 'EWX Week', 'Type.2', 'Auxiliary column', 'Inv&type', '是否签收',
            '型号', 'LRF', 'Unnamed: 66', 'Unnamed: 67', 'Unnamed: 68', 'Unnamed: 69','Unnamed: 70', 'Unnamed: 71'
        ]
        df = df.drop(columns=columns_to_remove, errors='ignore')

        # Create Ref1 column
        if 'Release Number' in df.columns and 'Container Number' in df.columns:
            df['Ref1'] = df['Release Number'].astype(str) + df['Container Number'].astype(str)
        else:
            df['Ref1'] = None

        return df
    except Exception as e:
        print(f"An error occurred while cleaning data: {e}")
        return None

def save_not_released_data(not_released_df):
    """
    Save the Not Released data to an Excel file with date in the filename.
    """
    try:
        # Reorder columns
        front_columns = ["Ref1", "Container Number", "Release Number", "DestinationWarehouse"]
        
        # Get remaining columns (excluding front columns)
        remaining_columns = [col for col in not_released_df.columns if col not in front_columns]
        
        # Create new column order
        new_column_order = front_columns + remaining_columns
        
        # Reorder the dataframe
        not_released_df = not_released_df[new_column_order]

        # Create the target directory if it doesn't exist
        target_dir = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Logistics Reporting\000_Master_Query_Reports\Automation_DB\EU_Stock_Report"
        os.makedirs(target_dir, exist_ok=True)

        # Generate filename with current date
        current_date = datetime.now().strftime("%Y%m%d")
        filename = f"EU_Stock_Report_{current_date}.xlsx"
        file_path = os.path.join(target_dir, filename)

        # Save the dataframe to Excel
        not_released_df.to_excel(file_path, index=False)
        print(f"Successfully saved Not Released data to: {file_path}")
        return True
    except Exception as e:
        print(f"An error occurred while saving Not Released data: {e}")
        return False

def split_released_data(df):
    """
    Split data into Released and Not Released datasets.
    """
    try:
        Released_data = df[
            (df['Release Number'].notna() & (df['Release Number'].astype(str).str.strip() != "")) & 
            (df['Sold Date'].notna() & (df['Sold Date'].astype(str).str.strip() != ""))
        ]

        Not_Released_data = df[
            (df['Release Number'].isna() | (df['Release Number'].astype(str).str.strip() == "")) & 
            (df['Sold Date'].isna() | (df['Sold Date'].astype(str).str.strip() == ""))
        ]

        # Save Not Released data to Excel
        save_not_released_data(Not_Released_data)

        return Released_data, Not_Released_data
    except Exception as e:
        print(f"An error occurred while splitting data: {e}")
        return None, None

def main():
    """
    Main execution function to process Europe Stock data
    """
    try:
        # Load the data
        df = load_europe_stock_data(europe_stock_path)
        if df is None:
            print("Failed to load Europe Stock data")
            return

        # Clean and prepare the data
        cleaned_df = clean_and_prepare_data(df)
        if cleaned_df is None:
            print("Failed to clean and prepare data")
            return

        # Split and save the data
        released_data, not_released_data = split_released_data(cleaned_df)
        if released_data is None or not_released_data is None:
            print("Failed to split data")
            return

        print("Process completed successfully!")

    except Exception as e:
        print(f"An error occurred in main execution: {e}")

if __name__ == "__main__":
    main()