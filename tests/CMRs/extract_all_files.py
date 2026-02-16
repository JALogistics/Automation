"""
Extract All Files from Multiple Folders
========================================
This script scans multiple folders (including subfolders) and extracts all filenames
to a single Excel file.

Author: Automated System
Date: 2026-01-26
Python Version: 3.8+
"""

import os
import pandas as pd
from pathlib import Path
from datetime import datetime
import logging


# ============================================================================
# CONFIGURATION
# ============================================================================

CONFIG = {
    # List of source folders to scan
    "source_folders": [
        r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Shared Documents - Supplier Management\Abreu\CMRs\CMR 2025",
        r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Shared Documents - Supplier Management\Access World\CMRs",
        r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Shared Documents - Supplier Management\Bollore\CMRs",
        r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Shared Documents - Supplier Management\CEVA PT\CMRs",
        r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Shared Documents - Supplier Management\Eurocom\CMRs",
        r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Shared Documents - Supplier Management\Fox\CMRs",
        r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Shared Documents - Supplier Management\Glovis\CMRs\2025",
        r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Shared Documents - Supplier Management\HIH\CMRs\2025 HIH",
        r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Shared Documents - Supplier Management\Intereuropa\CMRs",
        r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Shared Documents - Supplier Management\KNNL\CMRs",
        r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Shared Documents - Supplier Management\KRL\CMRs",
        r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Shared Documents - Supplier Management\LGI\CMRs",
        r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Shared Documents - Supplier Management\Odin Warehousing\CMRs\2025",
        r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Shared Documents - Supplier Management\Panex\CMRs\2025 Panex",
        r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Shared Documents - Supplier Management\Real\CMRs",
        r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Shared Documents - Supplier Management\Seacon\CMRs"
    ],
    
    # Output Excel file
    "output_file": "all_files_extracted.xlsx",
    
    # File extensions to include (leave empty [] to include ALL files)
    "file_extensions": [".pdf", ".jpg", ".jpeg", ".png"],  # Change to [] for all files
}


# ============================================================================
# LOGGING SETUP
# ============================================================================

def setup_logging():
    """Configure logging for progress tracking."""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )


# ============================================================================
# FILE EXTRACTOR
# ============================================================================

class FileExtractor:
    """Extracts all files from multiple folders."""
    
    def __init__(self, source_folders, file_extensions=None):
        """
        Initialize the File Extractor.
        
        Args:
            source_folders: List of folder paths to scan
            file_extensions: List of extensions to filter (None or [] for all files)
        """
        self.source_folders = source_folders
        self.file_extensions = [ext.lower() for ext in file_extensions] if file_extensions else None
        self.all_files = []
    
    def scan_all_folders(self):
        """Scan all configured folders and extract file information."""
        logging.info("="*70)
        logging.info("FILE EXTRACTION - STARTING")
        logging.info("="*70)
        logging.info(f"Scanning {len(self.source_folders)} folder(s)...")
        
        if self.file_extensions:
            logging.info(f"Filtering for extensions: {', '.join(self.file_extensions)}")
        else:
            logging.info("Including ALL file types")
        
        logging.info("")
        
        total_files = 0
        
        for idx, folder_path in enumerate(self.source_folders, 1):
            logging.info(f"[{idx}/{len(self.source_folders)}] Scanning: {folder_path}")
            
            if not os.path.exists(folder_path):
                logging.warning(f"  ⚠ Folder does not exist - SKIPPED")
                continue
            
            if not os.path.isdir(folder_path):
                logging.warning(f"  ⚠ Not a directory - SKIPPED")
                continue
            
            try:
                folder_files = self._scan_folder(folder_path)
                logging.info(f"  ✓ Found {len(folder_files)} file(s)")
                total_files += len(folder_files)
            except Exception as e:
                logging.error(f"  ✗ Error: {e}")
        
        logging.info("")
        logging.info(f"Scan complete! Total files found: {total_files}")
        return self.all_files
    
    def _scan_folder(self, folder_path):
        """
        Scan a single folder recursively.
        
        Args:
            folder_path: Path to folder to scan
            
        Returns:
            List of file dictionaries found in this folder
        """
        folder_files = []
        
        try:
            for root, dirs, files in os.walk(folder_path):
                for file in files:
                    file_path = Path(root) / file
                    file_extension = file_path.suffix.lower()
                    
                    # Filter by extension if specified
                    if self.file_extensions and file_extension not in self.file_extensions:
                        continue
                    
                    # Determine file type
                    if file_extension in [".jpg", ".jpeg"]:
                        file_type = "JPEG"
                    elif file_extension == ".pdf":
                        file_type = "PDF"
                    else:
                        file_type = file_extension.upper().replace(".", "") if file_extension else "NO EXTENSION"
                    
                    # Get relative path from the source folder
                    try:
                        relative_path = file_path.relative_to(folder_path)
                    except ValueError:
                        relative_path = file_path.name
                    
                    # Extract file information
                    file_info = {
                        "File Name (with extension)": file_path.name,
                        "File Name (no extension)": file_path.stem.upper(),
                        "File Type": file_type,
                        "Extension": file_extension if file_extension else "N/A",
                        "Folder Path": str(folder_path),
                        "Relative Path": str(relative_path),
                        "Full Path": str(file_path),
                        "File Size (bytes)": file_path.stat().st_size if file_path.exists() else 0,
                        "Modified Date": datetime.fromtimestamp(file_path.stat().st_mtime).strftime('%Y-%m-%d %H:%M:%S') if file_path.exists() else ""
                    }
                    
                    self.all_files.append(file_info)
                    folder_files.append(file_info)
                    
        except PermissionError as e:
            logging.error(f"Permission denied: {e}")
        except Exception as e:
            logging.error(f"Error scanning folder: {e}")
        
        return folder_files
    
    def export_to_excel(self, output_path):
        """
        Export all collected files to Excel.
        
        Args:
            output_path: Path to output Excel file
        """
        logging.info("")
        logging.info(f"Exporting to Excel: {output_path}")
        
        if not self.all_files:
            logging.warning("No files found to export!")
            return
        
        try:
            # Create DataFrame
            df = pd.DataFrame(self.all_files)
            
            # Sort by folder and filename
            df = df.sort_values(["Folder Path", "File Name (with extension)"]).reset_index(drop=True)
            
            # Create Excel file with multiple sheets
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # Main sheet with all details
                df.to_excel(writer, index=False, sheet_name="All Files")
                
                # Summary sheet with statistics
                summary_data = self._generate_summary(df)
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, index=False, sheet_name="Summary")
                
                # Simplified sheet with just filenames
                simple_df = df[["File Name (no extension)", "File Type"]].copy()
                simple_df.to_excel(writer, index=False, sheet_name="Simple List")
            
            logging.info(f"✓ Successfully exported {len(df)} files to {output_path}")
            logging.info("")
            logging.info("Excel file contains 3 sheets:")
            logging.info("  1. 'All Files' - Complete details with paths")
            logging.info("  2. 'Summary' - Statistics by folder and file type")
            logging.info("  3. 'Simple List' - Just filenames and types")
            
        except Exception as e:
            logging.error(f"Error exporting to Excel: {e}")
            raise
    
    def _generate_summary(self, df):
        """Generate summary statistics."""
        summary = []
        
        # Overall statistics
        summary.append({"Category": "OVERALL", "Metric": "Total Files", "Count": len(df)})
        summary.append({"Category": "OVERALL", "Metric": "Total Folders Scanned", "Count": df["Folder Path"].nunique()})
        summary.append({"Category": "OVERALL", "Metric": "Unique Filenames", "Count": df["File Name (no extension)"].nunique()})
        summary.append({"Category": "", "Metric": "", "Count": ""})
        
        # Files by type
        summary.append({"Category": "BY FILE TYPE", "Metric": "Type", "Count": "Count"})
        for file_type, count in df["File Type"].value_counts().items():
            summary.append({"Category": "BY FILE TYPE", "Metric": file_type, "Count": count})
        
        summary.append({"Category": "", "Metric": "", "Count": ""})
        
        # Files by folder
        summary.append({"Category": "BY FOLDER", "Metric": "Folder", "Count": "Count"})
        for folder, count in df["Folder Path"].value_counts().items():
            folder_name = Path(folder).name
            summary.append({"Category": "BY FOLDER", "Metric": folder_name, "Count": count})
        
        return summary


# ============================================================================
# MAIN EXECUTION
# ============================================================================

def main():
    """Main entry point."""
    setup_logging()
    
    # Create extractor
    extractor = FileExtractor(
        source_folders=CONFIG["source_folders"],
        file_extensions=CONFIG["file_extensions"]
    )
    
    # Scan folders
    extractor.scan_all_folders()
    
    # Export to Excel
    extractor.export_to_excel(CONFIG["output_file"])
    
    logging.info("="*70)
    logging.info("FILE EXTRACTION - COMPLETED SUCCESSFULLY")
    logging.info("="*70)


if __name__ == "__main__":
    main()
