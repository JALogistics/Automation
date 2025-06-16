import pandas as pd
from datetime import datetime
import os
from calendar import month_abbr

def get_month_order():
    # Get current month number (1-12)
    current_month = datetime.now().month
    
    # Create ordered list of months starting from current month
    months = list(month_abbr)[1:]  # Get all month abbreviations (Jan-Dec)
    ordered_months = months[current_month-1:] + months[:current_month-1]
    return ordered_months

def process_rno_report():
    # Input file path
    input_file = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Logistics Reporting\000_Master_Query_Reports\Automation_DB\RNO_Report\Final_RNO_Report.xlsx"
    
    # Output directory
    output_dir = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Logistics Reporting\000_Master_Query_Reports\Automation_DB\Sales_RNO_Report"
    
    # Columns to drop
    columns_to_drop = [
        "Outbound date",
        "Agreed Delivery date",
        "Delivery date",
        "On_Sea_Pcs",
        "In_Stock_Pcs",
        "Outbounded_Pcs",
        "Outbound_Comparison",
        "Case1",
        "cnt_Total_Pcs_cdr",
        "cnt_Outbound_Pcs_cdr",
        "Case4",
        "Pcs_from_wms",
        "Case2",
        "Case3",
        "Column1",
        "Current_Status",
        "Ref1",
        "3pls_Planned",
        "Email_Plan_Outbound"
    ]
    
    try:
        # Read the Excel file
        df = pd.read_excel(input_file, sheet_name="RNO Report")
        
        # Drop specified columns
        df = df.drop(columns=columns_to_drop, errors='ignore')
        
        # Filter out rows where Status Check matches F, H, J, or G
        status_to_remove = ['F', 'H', 'J', 'G']
        df = df[~df['Status Check '].isin(status_to_remove)]
        
        # Create output directory if it doesn't exist
        os.makedirs(output_dir, exist_ok=True)
        
        # Generate output filename with current date
        current_date = datetime.now().strftime("%Y%m%d")
        output_file = os.path.join(output_dir, f"Sales_RNO_Report_{current_date}.xlsx")
        
        # Create summary statistics
        # 1. Summary by Current_Status
        status_summary = df.groupby('Status Check ').agg({
            'Final_Pcs': 'sum',
            'Final_MWp': 'sum'
        }).round(2)
        
        # Add grand total to status summary
        status_summary.loc['Grand Total'] = status_summary.sum()
        
        # 2. Total Final_MWp by Outbound_Month where Current_Status is 'E'
        monthly_summary = df[df['Status Check '] == 'E'].groupby('Outbound_Month')['Final_MWp'].sum().round(2)
        
        # 3. Total Final_MWp by Salesman and Region info where Current_Status is 'B' or 'D'
        sales_summary = df[df['Status Check '].isin(['B', 'D'])].pivot_table(
            values='Final_MWp',
            index='Salesman',
            columns='Region info',
            aggfunc='sum',
            fill_value=0,
            margins=True,
            margins_name='Grand Total'
        ).round(2)
        
        # 4. Total Final_MWp by Status (Location) and Status (Standard/Due)
        status_location_summary = df.pivot_table(
            values='Final_MWp',
            index='Status (Location)',
            columns='Status (Standard/Due)',
            aggfunc='sum',
            fill_value=0,
            margins=True,
            margins_name='Grand Total'
        ).round(2)
        
        # Create Excel writer object
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Save main data
            df.to_excel(writer, sheet_name='Main Data', index=False)
            
            # Save status summary
            status_summary.to_excel(writer, sheet_name='Status Summary')
            
            # Save sales summary
            sales_summary.to_excel(writer, sheet_name='Sales to Push(B_D)')
            
            # Save status location summary
            status_location_summary.to_excel(writer, sheet_name='Current Location Goods')
        
        print(f"File successfully saved as: {output_file}")
        print("\nSummary Statistics:")
        print("\n1. Summary by Status Check:")
        print(status_summary)
        print("\n2. Monthly Summary (Status Check = 'E'):")
        print(monthly_summary)
        print("\n4. Status Location Matrix:")
        print(status_location_summary)
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    process_rno_report()
