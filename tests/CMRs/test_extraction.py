"""
Test script to verify container number extraction from complex filenames.
This demonstrates how the extract_filename_parts() function works.
"""

import re
import pandas as pd


def extract_filename_parts(filename):
    """
    Extract multiple parts from filename for matching.
    Handles cases like:
    - "2311160002- TRHU7369936" -> extracts "TRHU7369936"
    - "2505130072-EISU8486100-FCA-14-504" -> extracts "EISU8486100"
    
    Special handling for container number patterns (4 letters + 7 digits).
    """
    if pd.isna(filename):
        return []
    
    filename_str = str(filename).upper()
    cleaned_parts = []
    
    # METHOD 1: Extract container number patterns (4 letters + 7 digits)
    container_pattern = r'[A-Z]{4}\d{7}'
    container_matches = re.findall(container_pattern, filename_str)
    for match in container_matches:
        if match not in cleaned_parts:
            cleaned_parts.append(match)
    
    # METHOD 2: Split by common separators and extract parts
    parts = re.split(r'[-_\s,]+', filename_str)
    for part in parts:
        cleaned = re.sub(r'[^A-Z0-9]', '', part)
        if len(cleaned) >= 4 and cleaned not in cleaned_parts:
            cleaned_parts.append(cleaned)
    
    # METHOD 3: Add the fully cleaned version (all parts combined)
    full_cleaned = re.sub(r'[^A-Z0-9]', '', filename_str)
    if full_cleaned and full_cleaned not in cleaned_parts:
        cleaned_parts.insert(0, full_cleaned)
    
    return cleaned_parts


# Test cases
test_cases = [
    "2505130072-EISU8486100-FCA-14-504",
    "2311160002- TRHU7369936",
    "CMAU1234567",
    "ABC_DEFG1234567_XYZ",
    "2407240018 - CSNU8273953",
    "TEST-ABCD1234567-END",
    "MULTIPLE-WXYZ9876543-ITEMS-PQRS5432109"
]

print("=" * 80)
print("CONTAINER NUMBER EXTRACTION TEST")
print("=" * 80)
print()

for filename in test_cases:
    parts = extract_filename_parts(filename)
    print(f"Filename: {filename}")
    print(f"Extracted Parts: {parts}")
    
    # Check for container pattern (4 letters + 7 digits)
    container_pattern = r'[A-Z]{4}\d{7}'
    containers_found = [p for p in parts if re.match(f'^{container_pattern}$', p)]
    
    if containers_found:
        print(f"[OK] Container Numbers Found: {containers_found}")
    else:
        print("[NOTE] No standard container number pattern found")
    
    print("-" * 80)
    print()

print("=" * 80)
print("TEST COMPLETED")
print("=" * 80)
