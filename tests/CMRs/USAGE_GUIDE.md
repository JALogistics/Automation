# CMR Validation System - Summary

## 🎯 What It Does

Validates container names from an Excel file against CMR files stored across multiple folders, generating a comprehensive validation report.

## 📁 Key Files

1. **`extract_all_files.py`** - Scans folders and creates `all_files_extracted.xlsx`
2. **`cmr_check.py`** - Main validation script (validates container names)
3. **`test_extraction.py`** - Test script to verify pattern extraction
4. **`all_files_extracted.xlsx`** - Pre-extracted file database (19,101 files)
5. **`cmr_validation_report.xlsx`** - Final validation report with 3 sheets

## 🚀 How to Use

### Step 1: Extract Files (One-time or when files change)
```bash
python extract_all_files.py
```
This scans all configured folders and creates `all_files_extracted.xlsx`

### Step 2: Validate Containers
```bash
python cmr_check.py
```
This validates container names against extracted files and creates `cmr_validation_report.xlsx`

### Step 3: Review Results
Open `cmr_validation_report.xlsx` to see:
- **Sheet 1:** Full validation results (21,264 containers)
- **Sheet 2:** Summary statistics
- **Sheet 3:** Missing containers only (6,500 containers)

## 🔍 Smart Matching Features

### 1. Container Pattern Detection
Automatically finds standard container numbers (4 letters + 7 digits):
- Example: `2505130072-EISU8486100-FCA-14-504` → Finds `EISU8486100`

### 2. Multi-Part Splitting
Splits filenames by separators (-, _, space, comma):
- Example: `2311160002- TRHU7369936` → Finds both `2311160002` and `TRHU7369936`

### 3. Special Character Cleaning
Removes all special characters for consistent matching:
- Container: `TRHU-7369936` matches File: `TRHU 7369936`

## 📊 Current Results

- **Total Containers:** 21,264
- **Files Found:** 14,764 (69.43%)
- **Files Missing:** 6,500 (30.57%)
- **Searchable Index:** 42,745 entries

## ⚙️ Configuration

Edit the `CONFIG` section in each file:

### In `extract_all_files.py`:
```python
"source_folders": [
    r"C:\Path\To\CMR\Folder1",
    r"C:\Path\To\CMR\Folder2",
]
```

### In `cmr_check.py`:
```python
"extracted_files_path": r"path\to\all_files_extracted.xlsx",
"target_excel_path": r"path\to\check for CMRs.xlsx",
"container_column": "Container Name",
```

## 📝 Output Report Structure

### Validation Results Sheet
| Container Name | File Exists | File Type | Matched File Name |
|----------------|-------------|-----------|-------------------|
| EISU8486100 | Yes | PDF | 2505130072-EISU8486100-FCA-14-504 |
| TRHU7369936 | Yes | PDF | 2311160002- TRHU7369936 |
| XXXX9999999 | No | N/A | N/A |

### Summary Sheet
| Metric | Value |
|--------|-------|
| Total Container Names | 21,264 |
| Files Found | 14,764 |
| Files Missing | 6,500 |
| Match Rate (%) | 69.43% |
| Found - PDF Files | 14,500 |
| Found - JPEG Files | 264 |

### Missing Containers Sheet
Lists all containers that couldn't be matched (for follow-up)

## 🔧 Testing

Run the test script to verify pattern extraction:
```bash
python test_extraction.py
```

This shows how different filename formats are parsed.

## 📈 Performance

- **File Extraction:** ~5-7 seconds for 19,101 files
- **Validation:** ~4-5 seconds for 21,264 containers
- **Memory Usage:** Moderate (~43K lookup entries)

## 🎓 Technical Highlights

### Pattern Recognition
Uses regex to find container numbers: `[A-Z]{4}\d{7}`

### Lookup Index
Creates multiple search keys per file for fast matching:
- Original: 19,101 files
- Index entries: 42,745 (average 2.2 keys per file)

### Error Handling
- Missing folders → Warning, continues with other folders
- Missing files → Clear error messages
- Empty data → Graceful handling

## 📚 Dependencies

```txt
pandas>=2.0.0
openpyxl>=3.1.0
```

Install with:
```bash
pip install -r requirements.txt
```

## 💡 Tips

1. **Run extraction first** when files change in source folders
2. **Run validation frequently** - it's fast and uses pre-extracted data
3. **Check missing containers** sheet to identify gaps
4. **File naming matters** - standard container format (4 letters + 7 digits) works best

## 🐛 Troubleshooting

**Problem:** Low match rate  
**Solution:** Check if container names in Excel match filename patterns

**Problem:** Files not found  
**Solution:** Re-run `extract_all_files.py` to refresh file database

**Problem:** Encoding errors  
**Solution:** Ensure Excel files are saved in UTF-8 or standard format

---

**Version:** 2.1  
**Last Updated:** 2026-01-26  
**Status:** Production Ready ✓
