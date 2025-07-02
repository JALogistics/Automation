import pandas as pd
import os
from pathlib import Path

def consolidate_project_files():
    # Source and target directories
    source_dir = r"C:\Users\DeepakSureshNidagund\JA Solar GmbH\Projects - Documents and Tracking\Project Report and Report Genetation\On-going Projects"
    target_dir = r"C:\Users\DeepakSureshNidagund\JA Solar GmbH\Projects - Documents and Tracking\Project Report and Report Genetation\CDR_Proejct_Report"
    
    # Required columns
    required_columns = [
        'Name_partner', 'Project', 'Customer', 'ERP_no_Sales_Invoice', 'Delivery_Terms',
        'Internal_Contract', 'External_Contract', 'Module-Type', 'Container_No', 'Pieces',
        'Wattage', 'Power', 'MW', 'Sea_Freight_Agent', 'B/L_no', 'ETD', 'ETA',
        'ETD (UPDATE)', 'ETA (UPDATE)', 'Destination_Country', 'Destination_Port',
        'Terminal Storage free days', 'Free_DM+DT days', 'Custom_Clearance_Date',
        'Custom_Clearance_No', 'Departure_Date ( From Port)', 'Transport_Mode',
        'Updated_Container No', 'Inbound_Date', 'Planned_Delivery Date',
        'Actual_Delivery Date', 'POD sent (Y/N)', 'Container_Returned',
        'Damgage Claim', 'Remark/Comments'
    ]

    # Create target directory if it doesn't exist
    os.makedirs(target_dir, exist_ok=True)

    # List to store all dataframes
    all_dfs = []

    # Process each Excel file in the source directory
    for file in os.listdir(source_dir):
        if file.endswith(('.xlsx', '.xls')):
            file_path = os.path.join(source_dir, file)
            try:
                # Read Excel file
                df = pd.read_excel(file_path)
                
                # Check if required columns exist
                missing_cols = [col for col in required_columns if col not in df.columns]
                if missing_cols:
                    print(f"Warning: File {file} is missing columns: {missing_cols}")
                    # Add missing columns with NaN values
                    for col in missing_cols:
                        df[col] = pd.NA
                
                # Reorder columns to match required columns
                df = df.reindex(columns=required_columns)
                
                all_dfs.append(df)
                print(f"Successfully processed: {file}")
                
            except Exception as e:
                print(f"Error processing file {file}: {str(e)}")

    if not all_dfs:
        print("No files were processed successfully.")
        return

    # Combine all dataframes
    consolidated_df = pd.concat(all_dfs, ignore_index=True)

    # Generate output filename with timestamp
    timestamp = pd.Timestamp.now().strftime("%Y%m%d")
    output_file = os.path.join(target_dir, f"Consolidated_Project_Report_{timestamp}.xlsx")

    # Save consolidated file
    try:
        consolidated_df.to_excel(output_file, index=False)
        print(f"\nConsolidated file saved successfully at:\n{output_file}")
        print(f"\nTotal records processed: {len(consolidated_df)}")
    except Exception as e:
        print(f"Error saving consolidated file: {str(e)}")

if __name__ == "__main__":
    consolidate_project_files()
