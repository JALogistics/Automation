import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import os
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def split_data_by_month(file_path):
    """
    Split Excel data based on Month'Year format in column A into 12 monthly sheets.
    
    Args:
        file_path (str): Path to the Excel file
    """
    try:
        # Check if file exists
        if not os.path.exists(file_path):
            logger.error(f"File not found: {file_path}")
            return
        
        # Read the Excel file
        logger.info(f"Reading file: {file_path}")
        df = pd.read_excel(file_path)
        
        # Check if dataframe is empty
        if df.empty:
            logger.error("Excel file is empty")
            return
        
        # Get the first column name (column A)
        first_column = df.columns[0]
        logger.info(f"First column name: {first_column}")
        
        # Parse the month and year from column A (format: month'year)
        def parse_month_year(value):
            """Parse month'year format into month and year"""
            try:
                if pd.isna(value):
                    return None, None
                value_str = str(value).strip()
                if "'" in value_str:
                    parts = value_str.split("'")
                    month = int(parts[0])
                    year = int(parts[1])
                    return month, year
                return None, None
            except Exception as e:
                logger.warning(f"Could not parse value: {value} - {e}")
                return None, None
        
        # Add month and year columns
        df[['Month', 'Year']] = df[first_column].apply(
            lambda x: pd.Series(parse_month_year(x))
        )
        
        # Remove rows where month or year could not be parsed
        df_valid = df.dropna(subset=['Month', 'Year'])
        
        if df_valid.empty:
            logger.error("No valid month'year data found in column A")
            return
        
        logger.info(f"Found {len(df_valid)} valid rows with month'year data")
        
        # Create a new Excel writer
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
            # Group by month and create separate sheets
            months_created = []
            
            for month in range(1, 13):
                month_data = df_valid[df_valid['Month'] == month].copy()
                
                # Drop the temporary Month and Year columns
                month_data = month_data.drop(columns=['Month', 'Year'])
                
                if not month_data.empty:
                    sheet_name = get_month_name(month)
                    logger.info(f"Creating sheet '{sheet_name}' with {len(month_data)} rows")
                    month_data.to_excel(writer, sheet_name=sheet_name, index=False)
                    months_created.append(sheet_name)
                else:
                    # Create empty sheet with headers
                    sheet_name = get_month_name(month)
                    logger.info(f"Creating empty sheet '{sheet_name}'")
                    empty_df = pd.DataFrame(columns=df.columns[:-2])  # Exclude Month and Year columns
                    empty_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    months_created.append(sheet_name)
        
        logger.info(f"Successfully created {len(months_created)} monthly sheets: {', '.join(months_created)}")
        logger.info(f"File saved: {file_path}")
        
    except Exception as e:
        logger.error(f"Error processing file: {e}", exc_info=True)

def get_month_name(month_num):
    """Convert month number to month name"""
    month_names = {
        1: "January",
        2: "February",
        3: "March",
        4: "April",
        5: "May",
        6: "June",
        7: "July",
        8: "August",
        9: "September",
        10: "October",
        11: "November",
        12: "December"
    }
    return month_names.get(month_num, f"Month_{month_num}")

if __name__ == "__main__":
    # File path
    file_path = r"C:\Users\DeepakSureshNidagund\Downloads\CostPerWatt.xlsx"
    
    # Split data by month
    split_data_by_month(file_path)
