import os
import pandas as pd
import glob
from datetime import datetime
import warnings

# Suppress openpyxl warnings about default styles
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl.styles.stylesheet')

def combine_outbound_wms_files():
    # Path to the directory containing the files
    directory = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Documents - Sales Dashboards (BI Solution)\A_WMS Reports\Outbound_WMS"
    
    # Path for saving the combined output file
    output_directory = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Documents - Sales Dashboards (BI Solution)\a_Combinded WM_Report\Outbound_WMS_Report"
    
    # Create the output directory if it doesn't exist
    os.makedirs(output_directory, exist_ok=True)
    
    # Get a list of all files in the directory
    all_files = glob.glob(os.path.join(directory, "*.*"))
    
    # Filter for common data file extensions if needed
    # Uncomment and modify the next line if you need to filter for specific file types
    all_files = [f for f in all_files if f.endswith(('.xlsx', '.xls', '.csv'))]
    
    if not all_files:
        print("No files found in the directory.")
        return
    
    print(f"Found {len(all_files)} files to process.")
    
    # List to store all dataframes
    all_dfs = []
    
    # Process each file
    for file in all_files:
        try:
            # Determine file type and read accordingly
            if file.endswith(('.xlsx', '.xls')):
                # Skip first 2 rows, use 3rd row (index 2) as header
                df = pd.read_excel(file, skiprows=2, header=0, engine='openpyxl')
            elif file.endswith('.csv'):
                df = pd.read_csv(file, skiprows=2, header=0)
            else:
                print(f"Skipping unsupported file format: {file}")
                continue
                
            # Add source filename as a column for tracking
            df['Source_File'] = os.path.basename(file)

            # Create new Ref1 column by combining Release Number and Container No.
            df['Ref1'] = df['Release number'].astype(str) + df['Container number'].astype(str)
            
            # Append to our list of dataframes
            all_dfs.append(df)
            print(f"Processed file: {os.path.basename(file)}")
            
        except Exception as e:
            print(f"Error processing file {os.path.basename(file)}: {str(e)}")
    
    if not all_dfs:
        print("No data was successfully processed from any files.")
        return
    
    # Combine all dataframes
    combined_df = pd.concat(all_dfs, ignore_index=True)
    
    # Generate output file name with timestamp
    timestamp = datetime.now().strftime("%Y%m%d")
    output_file = os.path.join(output_directory, f"Combined_Outbound_WMS_{timestamp}.xlsx")
    
    # Save combined data to a new Excel file
    combined_df.to_excel(output_file, index=False, engine='openpyxl')
    
    print(f"Successfully combined {len(all_dfs)} files.")
    print(f"Combined data saved to: {output_file}")
    print(f"Total rows in combined file: {len(combined_df)}")
    
if __name__ == "__main__":
    combine_outbound_wms_files()
