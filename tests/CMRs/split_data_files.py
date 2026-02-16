"""
Split Missing CMRs Data by 3PLS
================================
This script reads the Missing CMRs report and splits it into separate files
for each 3PLS provider.

Author: Automated System
Date: 2026-01-26
Python Version: 3.8+
"""

import pandas as pd
import os
from pathlib import Path
import logging


# ============================================================================
# CONFIGURATION
# ============================================================================

CONFIG = {
    # Input file
    "input_file": r"C:\Users\DeepakSureshNidagund\Downloads\Reporting Application\Automation\automation\Missing CMRs compared to daily report.xlsx",
    
    # Column to split by
    "split_column": "3pls",
    
    # Output folder
    "output_folder": r"C:\Users\DeepakSureshNidagund\Downloads\Missing CMRs",
    
    # Column to treat as text (to preserve leading zeros)
    "text_columns": ["Release no"],
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
# DATA SPLITTER CLASS
# ============================================================================

class DataSplitter:
    """Splits Excel data by a specified column into multiple files."""
    
    def __init__(self, input_file, split_column, output_folder, text_columns=None):
        """
        Initialize the Data Splitter.
        
        Args:
            input_file (str): Path to input Excel file
            split_column (str): Column name to split data by
            output_folder (str): Folder to save output files
            text_columns (list): Column names to treat as text
        """
        self.input_file = input_file
        self.split_column = split_column
        self.output_folder = output_folder
        self.text_columns = text_columns or []
        self.data = None
    
    def load_data(self):
        """Load data from Excel file."""
        logging.info(f"Loading data from: {self.input_file}")
        
        try:
            # Read Excel with text columns as string
            if self.text_columns:
                dtype_dict = {col: str for col in self.text_columns}
                self.data = pd.read_excel(
                    self.input_file, 
                    dtype=dtype_dict
                )
            else:
                self.data = pd.read_excel(self.input_file)
            
            logging.info(f"Loaded {len(self.data)} rows from input file")
            logging.info(f"Columns: {', '.join(self.data.columns)}")
            
            # Check if split column exists
            if self.split_column not in self.data.columns:
                raise ValueError(
                    f"Column '{self.split_column}' not found in Excel. "
                    f"Available columns: {', '.join(self.data.columns)}"
                )
            
            return self.data
            
        except FileNotFoundError:
            logging.error(f"Input file not found: {self.input_file}")
            raise
        except Exception as e:
            logging.error(f"Error loading data: {e}")
            raise
    
    def split_and_save(self):
        """Split data by specified column and save to separate files."""
        logging.info("")
        logging.info(f"Splitting data by column: '{self.split_column}'")
        
        # Create output folder if it doesn't exist
        os.makedirs(self.output_folder, exist_ok=True)
        logging.info(f"Output folder: {self.output_folder}")
        logging.info("")
        
        # Get unique values in split column
        unique_values = self.data[self.split_column].dropna().unique()
        logging.info(f"Found {len(unique_values)} unique {self.split_column} values")
        
        # Also handle rows with NaN/empty values
        has_empty = self.data[self.split_column].isna().any()
        
        files_created = []
        
        # Process each unique value
        for idx, value in enumerate(unique_values, 1):
            # Filter data for this value
            filtered_data = self.data[self.data[self.split_column] == value].copy()
            
            # Create safe filename
            safe_filename = self._create_safe_filename(str(value))
            output_file = os.path.join(
                self.output_folder, 
                f"{safe_filename}_missing_cmrs.xlsx"
            )
            
            # Save to Excel
            self._save_to_excel(filtered_data, output_file, str(value))
            files_created.append(output_file)
            
            logging.info(f"[{idx}/{len(unique_values)}] Created: {safe_filename}_missing_cmrs.xlsx ({len(filtered_data)} rows)")
        
        # Handle empty/NaN values if any
        if has_empty:
            filtered_data = self.data[self.data[self.split_column].isna()].copy()
            output_file = os.path.join(
                self.output_folder, 
                "Unknown_3PLS_missing_cmrs.xlsx"
            )
            self._save_to_excel(filtered_data, output_file, "Unknown")
            files_created.append(output_file)
            logging.info(f"[EXTRA] Created: Unknown_3PLS_missing_cmrs.xlsx ({len(filtered_data)} rows)")
        
        logging.info("")
        logging.info(f"Total files created: {len(files_created)}")
        
        return files_created
    
    def _create_safe_filename(self, name):
        """
        Create a safe filename from a string.
        
        Args:
            name (str): Original name
            
        Returns:
            str: Safe filename
        """
        # Replace invalid characters with underscore
        invalid_chars = '<>:"/\\|?*'
        safe_name = name
        for char in invalid_chars:
            safe_name = safe_name.replace(char, '_')
        
        # Remove leading/trailing spaces and dots
        safe_name = safe_name.strip('. ')
        
        # Replace multiple spaces with single underscore
        safe_name = '_'.join(safe_name.split())
        
        return safe_name
    
    def _save_to_excel(self, data, output_file, sheet_name):
        """
        Save DataFrame to Excel file.
        
        Args:
            data (pd.DataFrame): Data to save
            output_file (str): Output file path
            sheet_name (str): Name for the sheet
        """
        try:
            # Ensure text columns are treated as text (prepend with apostrophe if needed)
            data_to_save = data.copy()
            
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                data_to_save.to_excel(
                    writer, 
                    index=False, 
                    sheet_name=sheet_name[:31]  # Excel sheet name limit
                )
            
        except Exception as e:
            logging.error(f"Error saving file {output_file}: {e}")
            raise
    
    def generate_summary(self):
        """Generate summary statistics."""
        logging.info("")
        logging.info("=" * 70)
        logging.info("SUMMARY")
        logging.info("=" * 70)
        
        # Overall stats
        logging.info(f"Total rows in input: {len(self.data)}")
        
        # Stats by split column
        value_counts = self.data[self.split_column].value_counts()
        logging.info(f"\nBreakdown by {self.split_column}:")
        for value, count in value_counts.items():
            logging.info(f"  - {value}: {count} rows")
        
        # Check for NaN
        nan_count = self.data[self.split_column].isna().sum()
        if nan_count > 0:
            logging.info(f"  - Unknown/Empty: {nan_count} rows")


# ============================================================================
# MAIN APPLICATION CLASS
# ============================================================================

class DataSplitterApp:
    """Main application orchestrator."""
    
    def __init__(self, config):
        """
        Initialize the application.
        
        Args:
            config (dict): Configuration dictionary
        """
        self.config = config
        self.splitter = DataSplitter(
            input_file=config["input_file"],
            split_column=config["split_column"],
            output_folder=config["output_folder"],
            text_columns=config.get("text_columns", [])
        )
    
    def run(self):
        """Execute the complete data splitting workflow."""
        try:
            logging.info("=" * 70)
            logging.info("DATA SPLITTER - STARTING")
            logging.info("=" * 70)
            logging.info("")
            
            # Step 1: Load data
            self.splitter.load_data()
            
            # Step 2: Split and save
            files_created = self.splitter.split_and_save()
            
            # Step 3: Generate summary
            self.splitter.generate_summary()
            
            logging.info("")
            logging.info("=" * 70)
            logging.info("DATA SPLITTER - COMPLETED SUCCESSFULLY")
            logging.info("=" * 70)
            
            return files_created
            
        except Exception as e:
            logging.error(f"Application error: {e}")
            raise


# ============================================================================
# MAIN EXECUTION
# ============================================================================

def main():
    """Main entry point."""
    # Setup logging
    setup_logging()
    
    # Display configuration
    logging.info("Configuration:")
    logging.info(f"  - Input file: {CONFIG['input_file']}")
    logging.info(f"  - Split by column: {CONFIG['split_column']}")
    logging.info(f"  - Output folder: {CONFIG['output_folder']}")
    logging.info(f"  - Text columns: {', '.join(CONFIG['text_columns'])}")
    logging.info("")
    
    # Create and run application
    app = DataSplitterApp(config=CONFIG)
    app.run()


if __name__ == "__main__":
    main()
