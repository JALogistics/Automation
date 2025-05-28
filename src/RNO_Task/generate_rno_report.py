import pandas as pd
import numpy as np
import os
import glob
from datetime import datetime
import openpyxl
from pathlib import Path

def load_europe_stock_data(file_path):
    """
    Load and process the Europe Stock data from Excel file.
    """
    try:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Excel file not found at: {file_path}")
            
        df = pd.read_excel(file_path, sheet_name="Summary-Europe")
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
            '型号', 'Unnamed: 65', 'Unnamed: 66', 'Unnamed: 67', 'Unnamed: 68', 'Unnamed: 69'
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

        return Released_data, Not_Released_data
    except Exception as e:
        print(f"An error occurred while splitting data: {e}")
        return None, None

def load_cdr_data(reports_dir):
    """
    Load the latest CDR report data.
    """
    try:
        if not os.path.exists(reports_dir):
            os.makedirs(reports_dir, exist_ok=True)

        patterns = ["CDR_*.csv", "*CDR*.csv", "*.csv"]
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
            latest_file = max(report_files, key=lambda x: 
                            datetime.strptime(os.path.basename(x).split('_')[1].split('.')[0], 
                                            "%Y-%m-%d"))
        except (IndexError, ValueError):
            latest_file = max(report_files, key=os.path.getmtime)

        cdr = pd.read_csv(latest_file)
        return cdr
    except Exception as e:
        print(f"An error occurred while loading CDR data: {e}")
        return None

def process_rno_data(Released_data, cdr):
    """
    Process RNO data by combining Released data with CDR data.
    """
    try:
        # Select relevant columns from CDR
        cdr_columns = ["Ref1", "Current_Status", "Outbound date", "Agreed Delivery date", 
                      "Delivery date", "Delivery_Status"]
        cdr_selected = cdr[cdr_columns]

        # Merge with Released data
        rno = pd.merge(Released_data, cdr_selected, on="Ref1", how="left")

        # Process CDR quantities
        cdr['Piece'] = pd.to_numeric(cdr['Piece'], errors='coerce')

        # Create status-based summaries
        status_summaries = {
            'On Sea': 'On_Sea_Pcs',
            'In-Stock': 'In_Stock_Pcs',
            'Outbounded': 'Outbounded_Pcs'
        }

        for status, col_name in status_summaries.items():
            summary = cdr[cdr['Current_Status'] == status].groupby('Ref1')['Piece'].sum().reset_index()
            summary = summary.rename(columns={'Piece': col_name})
            rno = pd.merge(rno, summary, on="Ref1", how="left")
            rno[col_name] = rno[col_name].fillna(0)

        return rno
    except Exception as e:
        print(f"An error occurred while processing RNO data: {e}")
        return None

def add_comparison_cases(rno):
    """
    Add comparison cases to the RNO data.
    """
    try:
        # Convert columns to numeric
        rno['Qty(PC)'] = pd.to_numeric(rno['Qty(PC)'], errors='coerce')
        rno['Outbounded_Pcs'] = pd.to_numeric(rno['Outbounded_Pcs'], errors='coerce')

        # Create comparison column
        rno['Outbound_Comparison'] = np.where(
            rno['Current_Status'] == 'Outbounded',
            rno['Outbounded_Pcs'] - rno['Qty(PC)'],
            rno['Qty(PC)']
        )

        # Replace negative values
        rno['Outbound_Comparison'] = np.where(
            rno['Outbound_Comparison'] < 0,
            rno['Qty(PC)'],
            rno['Outbound_Comparison']
        )

        # Add Case1 using np.where for consistency
        rno['Case1'] = np.where(rno['Outbound_Comparison'] == rno['Qty(PC)'], True, False)

        return rno
    except Exception as e:
        print(f"An error occurred while adding comparison cases: {e}")
        return None

def process_container_level_cdr(rno, cdr):
    """
    Process CDR data at container level and add container-level metrics.
    """
    try:
        # First ensure we have the container number in both dataframes
        if 'Container No.' in cdr.columns:
            container_col = 'Container No.'
        else:
            print("Warning: 'Container No.' column not found in CDR data")
            return None

        # Get total pieces by container
        cdr_total_summary = cdr.groupby(container_col)['Piece'].sum().reset_index()
        cdr_total_summary = cdr_total_summary.rename(columns={
            container_col: 'Container Number',
            'Piece': 'cnt_Total_Pcs_cdr'
        })

        # Get outbounded pieces by container
        cdr_outbound_summary = cdr[cdr['Current_Status'] == 'Outbounded'].groupby(container_col)['Piece'].sum().reset_index()
        cdr_outbound_summary = cdr_outbound_summary.rename(columns={
            container_col: 'Container Number',
            'Piece': 'cnt_Outbound_Pcs_cdr'
        })

        # Merge both summaries with rno DataFrame
        rno = pd.merge(rno, cdr_total_summary, on='Container Number', how='left')
        rno = pd.merge(rno, cdr_outbound_summary, on='Container Number', how='left')

        # Fill NaN values with 0 for both new columns
        rno['cnt_Total_Pcs_cdr'] = rno['cnt_Total_Pcs_cdr'].fillna(0)
        rno['cnt_Outbound_Pcs_cdr'] = rno['cnt_Outbound_Pcs_cdr'].fillna(0)

        # Convert to numeric
        rno['cnt_Total_Pcs_cdr'] = pd.to_numeric(rno['cnt_Total_Pcs_cdr'], errors='coerce')
        rno['cnt_Outbound_Pcs_cdr'] = pd.to_numeric(rno['cnt_Outbound_Pcs_cdr'], errors='coerce')

        # Add Case4 using np.where for consistency
        rno['Case4'] = np.where(rno['cnt_Outbound_Pcs_cdr'] == rno['Qty(PC)'], True, False)

        return rno
    except Exception as e:
        print(f"An error occurred while processing container level CDR data: {e}")
        return None

def process_wms_data(rno, wms_file_path):
    """
    Process WMS data and merge with RNO data.
    """
    try:
        # Read WMS data
        wms_Outbound = pd.read_excel(wms_file_path)
        
        # Create Ref1 in WMS data if not exists
        if 'Ref1' not in wms_Outbound.columns:
            if 'Release number' in wms_Outbound.columns and 'Container number' in wms_Outbound.columns:
                wms_Outbound['Ref1'] = wms_Outbound['Release number'].astype(str) + wms_Outbound['Container number'].astype(str)
            else:
                print("Required columns 'Release number' and 'Container number' not found in WMS data")
                return None
        
        # Create WMS summary
        wms_summary = wms_Outbound.groupby('Ref1')['Quantity'].sum().reset_index()
        wms_summary = wms_summary.rename(columns={'Quantity': 'Pcs_from_wms'})

        # Merge with RNO data
        rno = pd.merge(rno, wms_summary, on='Ref1', how='left')
        rno['Pcs_from_wms'] = pd.to_numeric(rno['Pcs_from_wms'], errors='coerce')
        rno['Pcs_from_wms'] = rno['Pcs_from_wms'].fillna(0)

        # Add Case2 and Case3
        rno['Case2'] = np.where(rno['Pcs_from_wms'] == rno['Outbounded_Pcs'],True,False)
        rno['Case3'] = np.where(rno['Pcs_from_wms'] == rno['Qty(PC)'] ,True,False)

        return rno
    except Exception as e:
        print(f"An error occurred while processing WMS data: {e}")
        return None

def apply_filters(rno, remove_data_path):
    """
    Apply all filters to the RNO data.
    """
    try:
        # Create a combined filter mask for all conditions
        keep_mask = ~(
            # Condition 1: Outbounded with In_Stock_Pcs = 0 and Case2 = True
            ((rno['Current_Status'] == 'Outbounded') & 
             (rno['In_Stock_Pcs'] == 0) & 
             (rno['Case2'] == True)) |
            
            # Condition 2: Outbounded with In_Stock_Pcs = 0 and Case3 = True
            ((rno['Current_Status'] == 'Outbounded') & 
             (rno['In_Stock_Pcs'] == 0) & 
             (rno['Case3'] == True)) |
            
            # Condition 3: Outbounded with In_Stock_Pcs = 0 and Case4 = True
            ((rno['Current_Status'] == 'Outbounded') & 
             (rno['In_Stock_Pcs'] == 0) & 
             (rno['Case4'] == True)) |
            
            # Condition 4: Outbounded with Outbound_Comparison = 0
            ((rno['Current_Status'] == 'Outbounded') & 
             (rno['Outbound_Comparison'] == 0)) |
            
            # Condition 5: Outbounded with Case2 = True
            ((rno['Current_Status'] == 'Outbounded') & 
             (rno['Case2'] == True)) |
            
            # Condition 6: All Outbounded records
            (rno['Current_Status'] == 'Outbounded')
        )

        # Apply the combined filter using loc
        rno = rno.loc[keep_mask].copy()

        # Remove rows with empty Container Number using loc
        container_mask = rno['Container Number'].notna() & (rno['Container Number'].str.strip() != '')
        rno = rno.loc[container_mask].copy()

        # Remove data based on Remove_data file
        if os.path.exists(remove_data_path):
            remove_df = pd.read_excel(remove_data_path)
            ref1_to_remove = remove_df['Ref1'].tolist()
            rno = rno.loc[~rno['Ref1'].isin(ref1_to_remove)].copy()

        # Reset index after all filtering
        rno = rno.reset_index(drop=True)

        return rno
    except Exception as e:
        print(f"An error occurred while applying filters: {e}")
        return None

def save_to_excel(rno, output_path):
    """
    Save the final RNO data to Excel.
    """
    try:
        # Reorder columns
        first_columns = ['Ref1', 'Container Number', 'Release Number']
        other_columns = [col for col in rno.columns if col not in first_columns]
        rno = rno[first_columns + other_columns]

        # Convert date columns
        date_columns = ['Sold Date', 'Outbound date', 'Agreed Delivery date', 'Delivery date']
        for col in date_columns:
            if col in rno.columns:
                rno[col] = pd.to_datetime(rno[col])

        # Save to Excel
        if os.path.exists(output_path):
            book = openpyxl.load_workbook(output_path)
            sheet = book['RNO Report']
            
            # Delete all rows except header
            while sheet.max_row > 1:
                sheet.delete_rows(2)
            
            book.save(output_path)
            book.close()

            # Write new data
            with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', 
                              if_sheet_exists='overlay') as writer:
                rno.to_excel(writer, sheet_name='RNO Report', index=False, 
                           header=False, startrow=1)
        else:
            rno.to_excel(output_path, sheet_name='RNO Report', index=False)

    except Exception as e:
        print(f"An error occurred while saving to Excel: {e}")

def main():
    """
    Main function to orchestrate the RNO report generation process.
    """
    try:
        # Define file paths
        europe_stock_path = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Documents - Sales Dashboards (BI Solution)\Y_EU Report\Europe Stock 最新版.xlsx"
        cdr_reports_dir = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Logistics Reporting\000_Master_Query_Reports\Automation_DB\CDR_Reports"
        wms_reports_dir = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Documents - Sales Dashboards (BI Solution)\a_Combinded WM_Report\Outbound_WMS_Report"
        remove_data_path = r"C:\Users\DeepakSureshNidagund\Downloads\Reporting Application\Automation\automation\tests\Remove_data.xlsx"
        output_path = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Logistics Reporting\000_Master_Query_Reports\Automation_DB\RNO_Report\RNO_Report.xlsx"

        # Load and process data
        print("Loading Europe stock data...")
        df = load_europe_stock_data(europe_stock_path)
        if df is None:
            return

        print("Cleaning and preparing data...")
        df = clean_and_prepare_data(df)
        if df is None:
            return

        print("Splitting released data...")
        Released_data, Not_Released_data = split_released_data(df)
        if Released_data is None or Not_Released_data is None:
            return

        print("Loading CDR data...")
        cdr = load_cdr_data(cdr_reports_dir)
        if cdr is None:
            return

        # Print CDR columns for debugging
        # print("CDR columns:", cdr.columns.tolist())

        print("Processing RNO data...")
        rno = process_rno_data(Released_data, cdr)
        if rno is None:
            return

        print("Adding comparison cases...")
        rno = add_comparison_cases(rno)
        if rno is None:
            return

        print("Processing container level CDR data...")
        rno = process_container_level_cdr(rno, cdr)
        if rno is None:
            return

        # Get latest WMS file
        print("Processing WMS data...")
        wms_files = glob.glob(os.path.join(wms_reports_dir, "*.xlsx"))
        if not wms_files:
            print("No WMS files found")
            return
        latest_wms_file = max(wms_files, key=os.path.getmtime)
        print(f"Using WMS file: {os.path.basename(latest_wms_file)}")

        rno = process_wms_data(rno, latest_wms_file)
        if rno is None:
            return

        print("Applying filters...")
        rno = apply_filters(rno, remove_data_path)
        if rno is None:
            return

        print("Saving to Excel...")
        save_to_excel(rno, output_path)
        print("RNO report generated successfully!")

    except Exception as e:
        print(f"An error occurred in the main function: {e}")

if __name__ == "__main__":
    main()
