import pandas as pd
import os
import glob
from datetime import datetime

def consolidate_stock_reports():
    source_dir = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Logistics Reporting\000_Master_Query_Reports\Automation_DB\CDR_Reports"
    dest_dir = r"C:\Users\DeepakSureshNidagund\JA Solar GmbH\Power BI Setup - PowerBISetup"
    template_path = r"C:\Users\DeepakSureshNidagund\JA Solar GmbH\Power BI Setup - PowerBISetup\Stock_ Template.xlsx"
    
    # Read template to get structure
    template_df = pd.read_excel(template_path)
    template_columns = template_df.columns.tolist()
    
    csv_files = glob.glob(os.path.join(source_dir, "*.csv"))
    all_data = []
    
    for file in csv_files:
        try:
            df = pd.read_csv(file, low_memory=False)
            # Only keep columns that exist in template
            df = df[df.columns.intersection(template_columns)]
            all_data.append(df)
        except:
            continue
    
    if all_data:
        combined_df = pd.concat(all_data, ignore_index=True)
        filtered_df = combined_df[combined_df['Current_Status'].str.strip() == 'In-Stock']
        
        # Ensure all template columns exist (fill with NaN if missing)
        for col in template_columns:
            if col not in filtered_df.columns:
                filtered_df[col] = pd.NA
                
        # Reorder columns to match template
        filtered_df = filtered_df.reindex(columns=template_columns)
        
        output_filename = "consolidated_stock_report.xlsx"
        output_path = os.path.join(dest_dir, output_filename)
        
        # Create destination directory if it doesn't exist
        os.makedirs(dest_dir, exist_ok=True)
        
        # Delete existing file if it exists
        if os.path.exists(output_path):
            try:
                os.remove(output_path)
            except:
                pass
                
        # Save the new file
        filtered_df.to_excel(output_path, index=False)
        return output_path
    
    return None

if __name__ == "__main__":
    consolidate_stock_reports()
