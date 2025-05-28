import os
import pandas as pd
from collections import defaultdict

# Folder containing the .xlsx files
FOLDER = r"C:\Users\DeepakSureshNidagund\JA Solar GmbH\Projects - Documents and Tracking\Project Report and Report Genetation\Generated Reports\Pre-files"
OUTPUT_FOLDER = r"C:\Users\DeepakSureshNidagund\JA Solar GmbH\Projects - Documents and Tracking\Project Report and Report Genetation\Generated Reports\Generated_Files"
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Define columns to highlight
highlight_columns = [
    "Inbound_Date",
    "Terminal Storage free days",
    "Planned_Delivery Date",
    "Actual_Delivery Date",
    "POD sent (Y/N)",
    "Container_Returned",
    "Damgage Container Status"
]

# Group files by first 16 characters
groups = defaultdict(list)
for fname in os.listdir(FOLDER):
    if fname.lower().endswith('.xlsx') and not fname.startswith('~$'):  # Exclude temp files
        key = fname.split(';')[0]  # Use substring up to first semicolon
        groups[key].append(os.path.join(FOLDER, fname))

for key, files in groups.items():
    combined = []
    for file in files:
        try:
            df = pd.read_excel(file, engine='openpyxl')
            df['__source_file__'] = os.path.basename(file)  # Optional: track source
            combined.append(df)
        except Exception as e:
            print(f"Error reading {file}: {e}")
    
    if combined:
        result = pd.concat(combined, ignore_index=True)
        
        try:
            name_partner = str(result['Name_partner'].iloc[0]) if 'Name_partner' in result.columns else 'Unknown_Partner'
            project = str(result['Project'].iloc[0]) if 'Project' in result.columns else 'Unknown_Project'
            customer = str(result['Customer'].iloc[0]) if 'Customer' in result.columns else 'Unknown_Customer'
            
            # Clean the values to make them filename-safe
            name_partner = ''.join(c for c in name_partner if c.isalnum() or c in '_- ')
            project = ''.join(c for c in project if c.isalnum() or c in '_- ')
            customer = ''.join(c for c in customer if c.isalnum() or c in '_- ')
            
            # Create the filename
            filename = f"{name_partner}_{project}_{customer}.xlsx"
            output_path = os.path.join(OUTPUT_FOLDER, filename)
            
            # Create a Pandas Excel writer with xlsxwriter engine
            writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
            result.to_excel(writer, index=False, sheet_name='Sheet1')
            
            # Get the xlsxwriter workbook and worksheet objects
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            
            # Define the highlight format
            highlight_format = workbook.add_format({
                'bg_color': '#FFE699',  # Light yellow background
                'font_color': '#000000'  # Black text
            })
            
            # Apply highlighting to specified columns
            for col_num, col_name in enumerate(result.columns):
                if col_name in highlight_columns:
                    # Convert column number to Excel column letter
                    worksheet.set_column(col_num, col_num, None, highlight_format)
            
            writer.close()
            print(f"Combined file saved with highlighting: {output_path}")
            
        except Exception as e:
            print(f"Error creating filename or saving file: {e}")
            # Fallback to original key-based filename if there's an error
            output_path = os.path.join(OUTPUT_FOLDER, f"{key}.xlsx")
            result.to_excel(output_path, index=False, engine='openpyxl')
            print(f"Saved with fallback filename: {output_path}")
