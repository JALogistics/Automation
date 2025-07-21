import pandas as pd
import os
from datetime import datetime
import glob

def get_latest_sales_rno_report():
    """
    Get the latest Sales RNO Report file from the specified directory
    """
    folder_path = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Logistics Reporting\000_Master_Query_Reports\Automation_DB\Sales_RNO_Report"
    pattern = os.path.join(folder_path, "Sales_RNO_Report_*.xlsx")
    files = glob.glob(pattern)
    
    if not files:
        raise FileNotFoundError(f"No Sales RNO Report files found in {folder_path}")
    
    # Get the latest file based on filename (which contains the date)
    latest_file = max(files)
    return latest_file

def read_template_headers(template_path):
    """
    Read the template file and extract headers from the first sheet
    """
    try:
        template_df = pd.read_excel(template_path, sheet_name=0)
        headers = template_df.columns.tolist()
        return headers
    except Exception as e:
        raise

def read_source_data(source_path, headers, sheet_name='Sheet1'):
    """
    Read source data and filter columns based on template headers
    """
    try:
        source_df = pd.read_excel(source_path, sheet_name=sheet_name)
        
        # Find matching columns between source and template headers
        matching_columns = []
        for header in headers:
            if header in source_df.columns:
                matching_columns.append(header)
        
        # Filter source data to only include matching columns
        if matching_columns:
            filtered_df = source_df[matching_columns].copy()
            return filtered_df
        else:
            return pd.DataFrame(columns=headers)
            
    except Exception as e:
        raise

def append_data_to_template(template_path, source1_path, source2_path, output_path):
    """
    Main function to append data from two sources to template
    """
    try:
        # Read template headers
        template_headers = read_template_headers(template_path)
        
        # Read source data
        source1_data = read_source_data(source1_path, template_headers, sheet_name='Main Data')
        source2_data = read_source_data(source2_path, template_headers)
        
        # Combine data from both sources
        combined_data = pd.concat([source1_data, source2_data], ignore_index=True)
        
        # Remove rows where "Container Number" is blank
        if "Container Number" in combined_data.columns:
            # Remove rows where Container Number is NaN, empty string, or whitespace
            combined_data = combined_data.dropna(subset=["Container Number"])
            combined_data = combined_data[combined_data["Container Number"].astype(str).str.strip() != ""]
            print(f"Removed rows with blank Container Number. Remaining records: {len(combined_data)}")
        else:
            print("Warning: 'Container Number' column not found in the data")
        
        # Drop specified columns
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
        
        # Drop columns that exist in the dataframe
        existing_columns_to_drop = [col for col in columns_to_drop if col in combined_data.columns]
        if existing_columns_to_drop:
            combined_data = combined_data.drop(columns=existing_columns_to_drop)
            print(f"Dropped {len(existing_columns_to_drop)} columns: {existing_columns_to_drop}")
        else:
            print("No specified columns found to drop")
        
        # Create output directory if it doesn't exist
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # Save combined data to output file
        combined_data.to_excel(output_path, index=False, sheet_name='Stock_Report')
        
        # Print summary
        print(f"\n=== Stock Report Generation Summary ===")
        print(f"Template headers: {len(template_headers)}")
        print(f"Source 1 records: {len(source1_data)}")
        print(f"Source 2 records: {len(source2_data)}")
        print(f"Total combined records: {len(combined_data)}")
        print(f"Output file: {output_path}")
        
        return True
        
    except Exception as e:
        raise

def main():
    """
    Main execution function with file paths
    """
    # File paths
    template_path = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Logistics Reporting\000_Master_Query_Reports\Automation_DB\EU_Stock_report\Temp_file\Templete.xlsx"
    try:
        source1_path = get_latest_sales_rno_report()
    except FileNotFoundError as e:
        print(f"\n❌ Error: {e}")
        return
    source2_path = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Logistics Reporting\000_Master_Query_Reports\Automation_DB\Not_Released_Goods\Not_Released_Goods.xlsx"
    
    # Generate output path with timestamp
    timestamp = datetime.now().strftime("%Y%m%d")
    output_dir = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Logistics Reporting\000_Master_Query_Reports\Automation_DB\EU_Stock_report"
    output_path = os.path.join(output_dir, f"Stock_Report_{timestamp}.xlsx")
    
    try:
        # Validate input files exist
        for file_path, file_name in [(template_path, "Template"), (source1_path, "Source 1"), (source2_path, "Source 2")]:
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"{file_name} file not found: {file_path}")
        
        # Generate stock report
        success = append_data_to_template(template_path, source1_path, source2_path, output_path)
        
        if success:
            print(f"\n✅ Stock report generated successfully!")
            print(f"📁 Sales rno report: {(source1_path)}")
            print(f"📁 Output file: {output_path}")
        else:
            print("\n❌ Stock report generation failed")
            
    except FileNotFoundError as e:
        print(f"\n❌ Error: {e}")
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")

if __name__ == "__main__":
    main()
