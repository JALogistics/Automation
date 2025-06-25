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
    and saving the BMO report with MTD data only.
    
    The function will:
    1. Find the latest logistics report
    2. Read the outbound logistics data
    3. Process the data for BMO report
    4. Generate and save MTD report
    
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
        
        # Ensure outbound date columns are always dt.date for consistency
        for col in ['Outbound Dates', 'Outbound date']:
            if col in bmo_df.columns:
                bmo_df[col] = pd.to_datetime(bmo_df[col], errors='coerce').dt.date
        
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
            start_date = pd.to_datetime(current_date.replace(day=1))  # First day of current month
            
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
        
        # Filter data for MTD report only
        mtd_df = filter_mtd_data(bmo_df)
        
        # Remove specified columns if they exist
        columns_to_remove = [
            "data_source",
            "MegaWattage",
            "Current_Status",
            "Release_Status",
            "Delivery_Status","Outbound_status"
        ]
        mtd_df = mtd_df.drop(columns=[col for col in columns_to_remove if col in mtd_df.columns], errors='ignore')
        
        # Save the report
        current_date = datetime.now().strftime("%Y-%m-%d")
        output_dir = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Logistics Reporting\000_Master_Query_Reports\Automation_DB\BMO_Reports"
        output_file = os.path.join(output_dir, f"bmo-report_{current_date}.xlsx")
        
        if os.path.exists(output_file):
            os.remove(output_file)
            print(f"Deleted existing file: {output_file}")
        
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            mtd_df.to_excel(writer, sheet_name='BMO Report MTD', index=False)
        
        print(f"BMO report saved to: {output_file}")
        
        # Generate and display summary
        summary = generate_bmo_summary(mtd_df)
        # Add Sales/Logistics to Push to summary
        push_data = get_sales_and_logistics_to_push()
        if push_data:
            summary.update(push_data)
            # Add Accumlate & Planned Outbound and Gap Target
            accumulated_outbound_mtd = summary.get('Accumulated_Outbound_MTD', 0)
            outbound_planned_this_month = summary.get('Outbound_Planned', 0)
            target_of_month = summary.get('Month_Target', 0)
            accum_and_planned = round(float(accumulated_outbound_mtd) + float(outbound_planned_this_month), 2)
            gap_target = round(float(target_of_month) - accum_and_planned, 2)
            if gap_target < 0:
                gap_target = abs(gap_target)
            summary['Accumlate_&_Planned_Outbound'] = accum_and_planned
            summary['Gap_Target'] = gap_target
            # Add Vs Target Outbound (%) and Vs Target Outbound Planned (%)
            vs_target_outbound = round((float(accumulated_outbound_mtd) / float(target_of_month)) * 100, 2) if float(target_of_month) else 0.0
            vs_target_outbound_planned = round((accum_and_planned / float(target_of_month)) * 100, 2) if float(target_of_month) else 0.0
            summary['Vs_Target_Outbound_(%)'] = f"{vs_target_outbound}%"
            summary['Vs_Target_Outbound_Planned_(%)'] = f"{vs_target_outbound_planned}%"
        # Convert all summary values to string
        summary = {k: str(v) for k, v in summary.items()}
        print("\nBMO Summary:")
        for k, v in summary.items():
            print(f"{k}: {v}")
        # Save summary to a new sheet in the same Excel file (as rows, not columns)
        summary_df = pd.DataFrame(list(summary.items()), columns=["JA Solar Logistics Report", "Mwps"])
        with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
        print(f"Summary saved to: {output_file} (sheet: 'Summary')")
        return output_file
        
    except Exception as e:
        print(f"Error in BMO report generation: {e}")
        sys.exit(1)

def generate_bmo_summary(df: pd.DataFrame, today: date = None) -> dict:
    """
    Generate BMO summary metrics for the report with fixed configuration values.
    """
    # Fixed values (edit as needed)
    target_of_month = 820.0
    total_working_days = 20
    working_days_till_today = 16

    if today is None:
        today = date.today()

    # Ensure 'Outbound date' is datetime
    df['Outbound date'] = pd.to_datetime(df['Outbound date'], errors='coerce').dt.date

    # Outbound today
    outbound_today = round(df.loc[df['Outbound date'] == today, 'MegaWattage_numeric'].sum(), 2)

    # Accumulated outbound till today (MTD)
    start_of_month = today.replace(day=1)
    mask_mtd = (df['Outbound date'] >= start_of_month) & (df['Outbound date'] <= today)
    accumulated_outbound_mtd = round(df.loc[mask_mtd, 'MegaWattage_numeric'].sum(), 2)

    # Outbound needed to achieve target
    outbound_needed_to_achieve_target = round(target_of_month - accumulated_outbound_mtd, 2)

    return {
        "Month_Target": target_of_month,
        "Total_Working_Days": total_working_days,
        "Working_Days_Till_Today": working_days_till_today,
        "Outbound_Today": outbound_today,
        "Accumulated_Outbound_MTD": accumulated_outbound_mtd,
        "Outbound_Needed_To_Achieve_Target": outbound_needed_to_achieve_target,
        
    }

def get_sales_and_logistics_to_push():
    sales_rno_dir = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Logistics Reporting\000_Master_Query_Reports\Automation_DB\Sales_RNO_Report"
    pattern = os.path.join(sales_rno_dir, "*.xlsx")
    files = glob.glob(pattern)
    if not files:
        print(f"No files found in {sales_rno_dir}")
        return None
    latest_file = max(files, key=os.path.getmtime)
    print(f"Latest Sales RNO Report: {latest_file}")
    df = pd.read_excel(latest_file, sheet_name='Main Data')
    # Sales to Push: Status Check == B or D
    sales_to_push = df[df['Status Check '].isin(['B', 'D'])]['Final_MWp'].sum()
    # Logistics to Push: Status Check == A or C
    logistics_to_push = df[df['Status Check '].isin(['A', 'C'])]['Final_MWp'].sum()
    # Outbound Planned for this month 
    outbound_planned_this_month = df[df['Final_Outbound_Plan'].dt.month == datetime.now().month]['Final_MWp'].sum()
    print(f"Sales to Push (B & D): {sales_to_push:.2f} MWp")
    print(f"Logistics to Push (A & C): {logistics_to_push:.2f} MWp")
    print(f"Outbound Planned: {outbound_planned_this_month:.2f} MWp")
    return {
        'Sales to Push': round(sales_to_push, 2),
        'Logistics to Push': round(logistics_to_push, 2),
        'Outbound_Planned': round(outbound_planned_this_month, 2)
    }

if __name__ == "__main__":
    generate_bmo_logistics_report() 