import os
import pandas as pd
from collections import defaultdict

# Folder containing the .xlsx files
FOLDER = r"C:\Users\DeepakSureshNidagund\JA Solar GmbH\Projects - Documents and Tracking\Project Report and Report Genetation\Generated Reports\Pre-files"
OUTPUT_FOLDER = r"C:\Users\DeepakSureshNidagund\JA Solar GmbH\Projects - Documents and Tracking\Project Report and Report Genetation\Generated Reports\Generated_Files"
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

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
        output_path = os.path.join(OUTPUT_FOLDER, f"{key}.xlsx")
        try:
            result.to_excel(output_path, index=False, engine='openpyxl')
            print(f"Combined file saved: {output_path}")
        except Exception as e:
            print(f"Error saving {output_path}: {e}")
