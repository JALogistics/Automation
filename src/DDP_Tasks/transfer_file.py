import os
import shutil

def transfer_latest_file():
    try:
        # Source and destination paths
        source_dir = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Logistics Reporting\000_Master_Query_Reports\Automation_DB\CDR_Reports"
        dest_dir = r"C:\Users\DeepakSureshNidagund\JA Solar GmbH\Power BI Setup - PowerBISetup"
        
        # Ensure directories exist
        if not os.path.exists(source_dir):
            print(f"Source directory does not exist: {source_dir}")
            return
        if not os.path.exists(dest_dir):
            print(f"Destination directory does not exist: {dest_dir}")
            return

        # Get list of files in source directory
        files = [os.path.join(source_dir, f) for f in os.listdir(source_dir) if os.path.isfile(os.path.join(source_dir, f))]
        
        if not files:
            print("No files found in source directory")
            return

        # Get the latest file
        latest_file = max(files, key=os.path.getmtime)
        
        # Get file extension from the latest file
        _, file_extension = os.path.splitext(latest_file)
        
        # Create destination filename
        dest_filename = f"CDR{file_extension}"
        dest_path = os.path.join(dest_dir, dest_filename)

        # Copy the file
        shutil.copy2(latest_file, dest_path)
        print(f"Successfully transferred file from {latest_file} to {dest_path}")

    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    transfer_latest_file()
