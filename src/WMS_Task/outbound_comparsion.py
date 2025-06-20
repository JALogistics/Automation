import os
import glob
from datetime import datetime
import pandas as pd
import numpy as np

def get_latest_excel_file(folder_path):
    """
    Get the latest Excel file from the specified folder
    """
    # Create a pattern to match .xlsx files or .csv files
    pattern = os.path.join(folder_path, '*.xlsx')

    
    # Get all Excel files in the directory
    excel_files = glob.glob(pattern)
    
    if not excel_files:
        print(f"No Excel files found in {folder_path}")
        return None
    
    # Get the latest file based on modification time
    latest_file = max(excel_files, key=os.path.getmtime)
    return latest_file

def get_wms_data():
    """
    Get and process WMS data
    """
    wms_folder = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Documents - Sales Dashboards (BI Solution)\a_Combinded WM_Report\Outbound_WMS_Report"
    wms_file = get_latest_excel_file(wms_folder)
    
    if wms_file:
        print(f"\nReading WMS file: {os.path.basename(wms_file)}")
        try:
            # Read WMS file
            wms_df = pd.read_excel(wms_file)
            
            # Group by Ref1 and sum Quantity
            wms_summary = wms_df.groupby('Ref1')['Quantity'].sum().reset_index()
            wms_summary = wms_summary.rename(columns={'Quantity': 'Outbound_Pcs_wms'})
            
            return wms_summary
        except Exception as e:
            print(f"Error processing WMS file: {str(e)}")
            return None
    return None

def get_wms_stock_data():
    """
    Get and process WMS Stock data
    """
    wms_stock_folder = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Logistics Reporting\000_Master_Query_Reports\Automation_DB\WMS_Stock_Report"
    wms_stock_file = get_latest_excel_file(wms_stock_folder)
    
    if wms_stock_file:
        print(f"\nReading WMS Stock file: {os.path.basename(wms_stock_file)}")
        try:
            # Read WMS stock file
            wms_stock_df = pd.read_excel(wms_stock_file)
            
            # Group by Ref1 and sum Quantity
            wms_stock_summary = wms_stock_df.groupby('Ref1')['Quantity'].sum().reset_index()
            wms_stock_summary = wms_stock_summary.rename(columns={'Quantity': 'Stock_Pcs_wms'})
            
            return wms_stock_summary
        except Exception as e:
            print(f"Error processing WMS Stock file: {str(e)}")
            return None
    return None

def process_excel_file(file_path):
    """
    Read and process the Excel file
    """
    try:
        # Read the specific sheet
        df = pd.read_excel(file_path, sheet_name="BMO Report MTD")
        
        # Create new Ref1 column by combining Release Number and Container No.
        df['Ref1'] = df['Release Number'].astype(str) + df['Container No.'].astype(str)
        
        # Get WMS data summary
        wms_summary = get_wms_data()
        if wms_summary is not None:
            # Merge WMS quantity with main dataframe
            df = pd.merge(df, wms_summary, on='Ref1', how='left')
            # Fill NaN values with 0 for Pcs_wms column
            df['Outbound_Pcs_wms'] = df['Outbound_Pcs_wms'].fillna(0)
        
        # Get WMS Stock data summary
        wms_stock_summary = get_wms_stock_data()
        if wms_stock_summary is not None:
            # Merge WMS stock quantity with main dataframe
            df = pd.merge(df, wms_stock_summary, on='Ref1', how='left')
            df['Stock_Pcs_wms'] = df['Stock_Pcs_wms'].fillna(0)
        
        # Add Case_1 column comparing Piece with Pcs_wms
        df['Case_1'] = np.where(df['Piece'] == df['Outbound_Pcs_wms'], True, False)

        
        # Get all column indices
        all_columns = list(range(len(df.columns)))
        
        # Columns to remove (0-based indexing, so subtract 1 from each number)
        columns_to_remove = [1,2,4,5,6,7,8,9,11,16,17,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35]
        #columns_to_remove = [1,2,4,5,6,7,8,9,11,,16,17,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35]
        
        # Get columns to keep
        columns_to_keep = [i for i in all_columns if i not in columns_to_remove]
        
        # Add the new Ref1 column to the columns to keep
        df_transformed = df.iloc[:, columns_to_keep]
        
        # Reorder columns to put Ref1 at the beginning and Pcs_wms, Case_1 at the end
        cols = df_transformed.columns.tolist()
        cols = ['Ref1'] + [col for col in cols if col not in ['Ref1', 'Outbound_Pcs_wms', 'Case_1', 'Stock_Pcs_wms']] + ['Outbound_Pcs_wms', 'Stock_Pcs_wms', 'Case_1']
        df_transformed = df_transformed[cols]

        # Define output directory and create if it doesn't exist
        output_dir = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Logistics Reporting\000_Master_Query_Reports\Automation_DB\WMS_compared_report"
        os.makedirs(output_dir, exist_ok=True)
        
        # Create output filename with only date
        current_date = datetime.now().strftime("%Y%m%d")
        output_filename = f"compared_data_{current_date}.xlsx"
        output_path = os.path.join(output_dir, output_filename)
        
        # Print summary of Case_1 results with MegaWattage calculations
        print("\nCase_1 Summary with MegaWattage:")
        
        # Calculate for matching records (True)
        matching_records = df_transformed[df_transformed['Case_1'] == True]
        matching_count = len(matching_records)
        matching_megawatt = (matching_records['MegaWattage_numeric']).sum()
        
        # Calculate for non-matching records (False)
        non_matching_records = df_transformed[df_transformed['Case_1'] == False]
        non_matching_count = len(non_matching_records)
        non_matching_megawatt = (non_matching_records['MegaWattage_numeric']).sum()
        
        print(f"Matching records (True):")
        print(f"  Count: {matching_count}")
        print(f"  Total MegaWattage: {matching_megawatt:.2f} MW")
        print(f"\nNon-matching records (False):")
        print(f"  Count: {non_matching_count}")
        print(f"  Total MegaWattage: {non_matching_megawatt:.2f} MW")
        print(f"\nTotal MegaWattage: {(matching_megawatt + non_matching_megawatt):.2f} MW")

        # Filter out matching records (keep only mismatches)
        df_transformed = df_transformed[df_transformed['Case_1'] == False]
        
        # Remove the Case_1 column as it's no longer needed
        df_transformed = df_transformed.drop(columns=['Case_1'])
        
        # Save the transformed data (only mismatches)
        df_transformed.to_excel(output_path, index=False)
        
        # print("\nSaved file contains only mismatched records (Case_1 = False)")
        print(f"Number of mismatched records saved: {len(df_transformed)}")
        # print(f"File location: {output_path}")
        
        return output_path
        
    except Exception as e:
        print(f"Error processing file: {str(e)}")
        return None

def main():
    # Define the folder path
    folder_path = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Logistics Reporting\000_Master_Query_Reports\Automation_DB\BMO_Reports"
    
    # Get latest file
    latest_file = get_latest_excel_file(folder_path)
    
    if latest_file:
        mod_time = datetime.fromtimestamp(os.path.getmtime(latest_file))
        print(f"\nProcessing latest file from BMO_Reports folder:")
        print(f"File: {os.path.basename(latest_file)}")
        print(f"Last modified: {mod_time}")
        
        # Process the file
        processed_file = process_excel_file(latest_file)
        if processed_file:
            print(f"Processing completed successfully!")

if __name__ == "__main__":
    main()
