import os
import sys
import glob
import pandas as pd
from pathlib import Path
from datetime import datetime, date

def generate_bmo_logistics_report():
    """
    Generate BMO logistics report by processing the latest logistics report file.
    This function handles the entire process from finding the latest report to generating
    and saving the BMO report with YTD and MTD data.
    
    The function will:
    1. Find the latest logistics report
    2. Read the outbound logistics data
    3. Process the data for BMO report
    4. Generate and save YTD and MTD reports
    
    Returns:
        str: Path to the saved BMO report file
    """
    print("Starting BMO report generation...")
    
    try:
        # Find the latest logistics report file
        reports_dir = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Logistics Reporting\000_Master_Query_Reports\Automation_DB\Logistics_Reports"
        pattern = os.path.join(reports_dir, "logistics-report_*.xlsx")
        report_files = glob.glob(pattern)
        
        if not report_files:
            print(f"No logistics report files found in {reports_dir}")
            sys.exit(1)
        
        latest_file = max(report_files, key=os.path.getmtime)
        print(f"Found latest logistics report: {latest_file}")
        
        # Read the Outbound_logistics_report data
        try:
            df = pd.read_excel(latest_file, sheet_name='Outbound_logistics_report')
            print(f"Successfully read data from Outbound_logistics_report sheet: {len(df)} records")
        except Exception as e:
            print(f"Error reading Outbound_logistics_report sheet: {e}")
            sys.exit(1)
            
        # Create a copy for BMO report processing
        bmo_df = df.copy()
        
        # Process YTD data
        def filter_ytd_data(df):
            if 'Outbound Dates' in df.columns:
                column_name = 'Outbound Dates'
            elif 'Outbound date' in df.columns:
                column_name = 'Outbound date'
            else:
                date_columns = [col for col in df.columns if 'date' in col.lower() and 'outbound' in col.lower()]
                if date_columns:
                    column_name = date_columns[0]
                else:
                    print("Warning: Could not find outbound date column. Using all data for YTD report.")
                    return df.copy()

            if not pd.api.types.is_datetime64_any_dtype(df[column_name]):
                df[column_name] = pd.to_datetime(df[column_name], errors='coerce')
            
            current_date = datetime.now()
            start_date = datetime(current_date.year, 1, 1)  # Use current year instead of hardcoded 2025
            
            ytd_df = df[
                (df[column_name] >= start_date) & 
                (df[column_name] <= current_date)
            ].copy()
            
            print(f"YTD Filter Details:")
            print(f"- Date column used: {column_name}")
            print(f"- Start date: {start_date.strftime('%Y-%m-%d')}")
            print(f"- End date: {current_date.strftime('%Y-%m-%d')}")
            print(f"- Total records before filter: {len(df)}")
            print(f"- Total records after YTD filter: {len(ytd_df)}")
            
            # Calculate and print total MegaWattage for YTD
            if 'MegaWattage_numeric' in ytd_df.columns:
                total_mw_ytd = ytd_df['MegaWattage_numeric'].sum()
                print(f"- Total YTD MegaWattage: {total_mw_ytd:.2f} MW")
            else:
                print("Warning: MegaWattage_numeric column not found in YTD data")
            
            return ytd_df
        
        # Process MTD data
        def filter_mtd_data(df):
            if 'Outbound Dates' in df.columns:
                column_name = 'Outbound Dates'
            elif 'Outbound date' in df.columns:
                column_name = 'Outbound date'
            else:
                date_columns = [col for col in df.columns if 'date' in col.lower() and 'outbound' in col.lower()]
                if date_columns:
                    column_name = date_columns[0]
                else:
                    print("Warning: Could not find outbound date column. Using all data for MTD report.")
                    return df.copy()

            if not pd.api.types.is_datetime64_any_dtype(df[column_name]):
                df[column_name] = pd.to_datetime(df[column_name], errors='coerce')
            
            current_date = datetime.now()
            start_date = datetime(current_date.year, current_date.month, 1)  # First day of current month
            
            mtd_df = df[
                (df[column_name] >= start_date) & 
                (df[column_name] <= current_date)
            ].copy()
            
            print(f"\nMTD Filter Details:")
            print(f"- Date column used: {column_name}")
            print(f"- Start date: {start_date.strftime('%Y-%m-%d')}")
            print(f"- End date: {current_date.strftime('%Y-%m-%d')}")
            print(f"- Total records before filter: {len(df)}")
            print(f"- Total records after MTD filter: {len(mtd_df)}")
            
            # Calculate and print total MegaWattage for MTD
            if 'MegaWattage_numeric' in mtd_df.columns:
                total_mw_mtd = mtd_df['MegaWattage_numeric'].sum()
                print(f"- Total MTD MegaWattage: {total_mw_mtd:.2f} MW")
            else:
                print("Warning: MegaWattage_numeric column not found in MTD data")
            
            return mtd_df
        
        # Filter data for reports
        ytd_df = filter_ytd_data(bmo_df)
        mtd_df = filter_mtd_data(bmo_df)
        
        # Save the report
        current_date = datetime.now().strftime("%Y-%m-%d")
        output_dir = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Logistics Reporting\000_Master_Query_Reports\Automation_DB\BMO_Reports"
        output_file = os.path.join(output_dir, f"bmo-report_{current_date}.xlsx")
        
        if os.path.exists(output_file):
            os.remove(output_file)
            print(f"Deleted existing file: {output_file}")
        
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            ytd_df.to_excel(writer, sheet_name='BMO Report YTD', index=False)
            mtd_df.to_excel(writer, sheet_name='BMO Report MTD', index=False)
        
        print(f"BMO report saved to: {output_file}")
        return output_file
        
    except Exception as e:
        print(f"Error in BMO report generation: {e}")
        sys.exit(1)

if __name__ == "__main__":
    generate_bmo_logistics_report() 