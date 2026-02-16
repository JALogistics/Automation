# CMR File Validation System

A robust Python application to scan multiple folders for CMR (Clinical Medical Records) files and validate them against a target Excel file.

## Features

- 📁 **Recursive folder scanning** - Scans multiple folders including all nested subfolders
- 📄 **Multi-format support** - Handles PDF and JPEG/JPG files
- ✅ **Duplicate handling** - Automatically keeps first occurrence of duplicate filenames
- 📊 **Comprehensive reporting** - Generates two detailed Excel reports
- 🔍 **Statistics dashboard** - Summary sheet with validation metrics
- 🛡️ **Error handling** - Robust error handling for production use
- 📝 **Detailed logging** - Real-time progress messages during execution

## Requirements

- Python 3.8 or higher
- Required packages (see requirements.txt):
  - pandas >= 2.0.0
  - openpyxl >= 3.1.0

## Installation

1. **Clone or download this repository**

2. **Install required packages:**
   ```bash
   pip install -r requirements.txt
   ```

## Configuration

Before running the script, you need to configure the paths in the `cmr_check.py` file:

### 1. Locate the CONFIGURATION SECTION (lines 22-45)

```python
CONFIG = {
    # List of source folders to scan (add or remove paths as needed)
    "source_folders": [
        r"C:\CMR_Files\Folder1",
        r"C:\CMR_Files\Folder2",
        r"C:\CMR_Files\Folder3",
    ],
    
    # Target Excel file containing container names
    "target_excel_path": r"C:\CMR_Files\container_list.xlsx",
    
    # Column name in Excel that contains container names
    "container_column": "Container Name",
    
    # Output files
    "output_found_files": "found_files_list.xlsx",
    "output_validation_report": "cmr_validation_report.xlsx",
    
    # Supported file extensions
    "supported_extensions": [".pdf", ".jpg", ".jpeg"],
}
```

### 2. Update the following settings:

#### Source Folders:
```python
"source_folders": [
    r"C:\Your\First\Folder\Path",
    r"C:\Your\Second\Folder\Path",
    r"D:\Network\Share\CMR_Files",
],
```

**Tips:**
- Add as many folders as needed
- Use raw strings (prefix with `r`) to avoid backslash issues
- Can be local drives, network shares, or external drives
- Non-existent folders will be skipped with a warning

#### Target Excel File:
```python
"target_excel_path": r"C:\Path\To\Your\container_list.xlsx",
```

**Requirements:**
- Must be a valid Excel file (.xlsx or .xls)
- Must contain a column with container names to validate

#### Container Column Name:
```python
"container_column": "Container Name",
```

**Note:** Change this to match your Excel column name exactly (case-sensitive)

#### Output Files (Optional):
```python
"output_found_files": "found_files_list.xlsx",
"output_validation_report": "cmr_validation_report.xlsx",
```

**Default:** Files are saved in the same directory as the script

## Usage

### Basic Usage

1. **Configure the script** (see Configuration section above)

2. **Run the script:**
   ```bash
   python cmr_check.py
   ```

3. **Monitor the console output** - You'll see real-time progress messages

4. **Check the output files** - Two Excel files will be generated in the script directory

### Example Console Output

```
2026-01-26 10:30:45 - INFO - ======================================================================
2026-01-26 10:30:45 - INFO - CMR FILE VALIDATION SYSTEM - STARTING
2026-01-26 10:30:45 - INFO - ======================================================================
2026-01-26 10:30:45 - INFO - Configuration:
2026-01-26 10:30:45 - INFO -   - Source folders: 3 folder(s)
2026-01-26 10:30:45 - INFO -     1. C:\CMR_Files\Folder1
2026-01-26 10:30:45 - INFO -     2. C:\CMR_Files\Folder2
2026-01-26 10:30:45 - INFO -     3. C:\CMR_Files\Folder3
2026-01-26 10:30:45 - INFO - Starting folder scan...
2026-01-26 10:30:45 - INFO - Scanning folder: C:\CMR_Files\Folder1
2026-01-26 10:30:46 - INFO - Found 150 files in C:\CMR_Files\Folder1
2026-01-26 10:30:46 - INFO - Scan complete. Total files found: 450
2026-01-26 10:30:46 - INFO - Exporting found files list to: found_files_list.xlsx
2026-01-26 10:30:47 - INFO - Loading container names from: C:\CMR_Files\container_list.xlsx
2026-01-26 10:30:47 - INFO - Loaded 500 unique container names
2026-01-26 10:30:47 - INFO - Starting validation process...
2026-01-26 10:30:47 - INFO - Validation complete
2026-01-26 10:30:47 - INFO - Summary - Total: 500, Found: 445, Missing: 55, Match Rate: 89.00%
2026-01-26 10:30:47 - INFO - ======================================================================
2026-01-26 10:30:47 - INFO - CMR FILE VALIDATION SYSTEM - COMPLETED SUCCESSFULLY
2026-01-26 10:30:47 - INFO - ======================================================================
```

## Output Files

### 1. found_files_list.xlsx

Contains all unique files discovered during the scan.

| File Name | File Type |
|-----------|-----------|
| CMR_12345 | PDF |
| CMR_12346 | JPEG |
| CMR_12347 | PDF |

**Columns:**
- **File Name**: Filename without extension
- **File Type**: PDF or JPEG

### 2. cmr_validation_report.xlsx

Contains two sheets:

#### Sheet 1: Validation Results

| Container Name | File Exists | File Type |
|----------------|-------------|-----------|
| CMR_12345 | Yes | PDF |
| CMR_12346 | Yes | JPEG |
| CMR_99999 | No | N/A |

**Columns:**
- **Container Name**: Name from target Excel file
- **File Exists**: "Yes" if file found, "No" if missing
- **File Type**: PDF, JPEG, or N/A (if not found)

#### Sheet 2: Summary

| Metric | Value |
|--------|-------|
| Total Container Names | 500 |
| Files Found | 445 |
| Files Missing | 55 |
| Match Rate (%) | 89.00% |

## Edge Cases Handled

✅ Non-existent folder paths - Skipped with warning  
✅ Empty folders - Handled gracefully  
✅ Duplicate filenames - First occurrence kept  
✅ Missing target Excel file - Clear error message  
✅ Empty target Excel file - Handled gracefully  
✅ Both .jpg and .jpeg extensions - Treated as same type (JPEG)  
✅ Permission errors - Logged and skipped  
✅ Invalid column names - Clear error message with available columns  

## Troubleshooting

### Error: "Target Excel file not found"
**Solution:** Check that the `target_excel_path` is correct and the file exists

### Error: "Column 'X' not found in Excel"
**Solution:** Verify that the `container_column` name matches your Excel column exactly (case-sensitive)

### Warning: "Folder does not exist"
**Solution:** Check the folder path in `source_folders` is correct

### No files found
**Solution:** 
- Verify folders contain PDF or JPEG files
- Check file extensions are .pdf, .jpg, or .jpeg
- Ensure you have read permissions for the folders

### Script runs but output files are empty
**Solution:**
- Check that target Excel file contains data in the specified column
- Verify source folders contain supported file types

## Code Structure

The application uses Object-Oriented Programming with four main classes:

- **CMRFileScanner**: Handles folder scanning and file discovery
- **CMRValidator**: Performs validation logic against target Excel
- **ExcelExporter**: Manages Excel file exports
- **CMRFileChecker**: Main orchestrator that coordinates the workflow

## Best Practices

1. **Test with small dataset first** - Use a single folder with a few files to verify configuration
2. **Backup important data** - Although this script only reads files, always backup before bulk operations
3. **Use absolute paths** - Avoid relative paths to prevent confusion
4. **Check logs** - Review console output for warnings or errors
5. **Validate output** - Spot-check the generated Excel files to ensure accuracy

## Support

For issues or questions:
1. Review the troubleshooting section above
2. Check the console output for error messages
3. Verify your configuration matches the requirements

## License

This script is provided as-is for CMR file validation purposes.

---

**Version:** 1.0.0  
**Last Updated:** January 26, 2026  
**Python Version:** 3.8+
