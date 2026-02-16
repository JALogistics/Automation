# Daily Report ETL Script

## Overview
This script combines multiple Excel files from a source folder and uploads them to a Supabase database table called `daily_report`.

## Features
- **Column Mapping**: Automatically maps columns from source files to template headers (case-insensitive matching)
- **Flexible Structure**: Handles Excel files with columns in any order
- **Batch Upload**: Uploads data to Supabase in batches of 1000 records for optimal performance
- **Error Handling**: Comprehensive logging and error handling
- **Data Validation**: Cleans and validates data before upload

## Requirements
All dependencies are listed in `requirements.txt`:
- pandas>=2.0.0
- openpyxl>=3.1.0
- supabase>=2.0.0
- python-dotenv>=1.0.0

## Configuration

### Environment Variables
Set the following in your `.env` file:
```
SUPABASE_URL=your_supabase_url
SUPABASE_KEY=your_supabase_key
```

### File Paths
The script uses the following paths (configured in the script):
- **SOURCE_FOLDER**: `C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Logistics Reporting\0_Daily Reports_Warehouses (Excel files)`
- **TEMPLATE_FILE**: `0_DailyReportTemplate (DO NOT DELETE,CHANGE).xlsx`

## Usage

### Installation
```bash
pip install -r requirements.txt
```

### Running the Script
```bash
python tests/ETL.py
```

## How It Works

1. **Read Template Headers**: Loads the template Excel file to get the standard column headers
2. **Process Excel Files**: 
   - Scans the source folder for all Excel files (.xlsx, .xls)
   - Skips temporary files (starting with ~$)
   - For each file:
     - Reads the data
     - Maps columns to template headers using case-insensitive matching
     - Adds missing columns with NULL values
     - Reorders columns to match template
3. **Combine Data**: Concatenates all processed files into a single DataFrame
4. **Prepare for Upload**: 
   - Converts data to appropriate format
   - Handles NULL values
   - Converts all values to strings (as per table schema)
5. **Upload to Supabase**: Uploads data in batches of 1000 records

## Column Mapping
The script performs intelligent column mapping:
- Normalizes column names (removes extra spaces, converts to lowercase)
- Matches source columns to template columns
- Adds missing columns with NULL values
- Warns about unmatched columns

## Logging
The script provides detailed logging:
- INFO: Progress updates and successful operations
- WARNING: Non-critical issues (empty files, unmatched columns)
- ERROR: Critical errors that stop processing

## Error Handling
- File-level errors don't stop the entire process
- Invalid files are logged and skipped
- Database upload errors are caught and logged
- Missing credentials raise clear error messages

## Table Schema
The script uploads to the `daily_report` table with 71 columns including:
- Partner and warehouse information
- Container and shipping details
- Dates and timelines
- Stock and inventory data
- Costs and financial data
- Customer information

All columns are stored as TEXT in the database.
