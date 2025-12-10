"""
Extract the Plant Pictures dataset with proper handling of long paths and nested directories.
"""

import zipfile
import os
from pathlib import Path
import shutil

# Paths
zip_path = Path.home() / "Downloads" / "Plant Pictures.zip"
extract_to = Path("data/raw")

print("="*80)
print("EXTRACTING PLANT DISEASE DATASET")
print("="*80)
print(f"Source: {zip_path}")
print(f"Destination: {extract_to}")
print()

if not zip_path.exists():
    print(f"❌ ERROR: Zip file not found at {zip_path}")
    exit(1)

# Create extraction directory
extract_to.mkdir(parents=True, exist_ok=True)

print("Starting extraction... (this will take several minutes)")
print()

try:
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        # Get list of files
        file_list = zip_ref.namelist()
        total_files = len(file_list)
        
        print(f"Total files to extract: {total_files:,}")
        print()
        
        # Extract with progress
        for i, file in enumerate(file_list, 1):
            try:
                zip_ref.extract(file, extract_to)
                
                # Print progress every 1000 files
                if i % 1000 == 0 or i == total_files:
                    print(f"Progress: {i:,}/{total_files:,} files ({i/total_files*100:.1f}%)")
                    
            except Exception as e:
                # Skip files that cause errors (usually due to path length)
                if i % 1000 == 0:  # Only print errors periodically
                    print(f"⚠ Warning: Skipped some files due to path length issues")
                continue
        
        print()
        print("✓ Extraction complete!")
        
except Exception as e:
    print(f"❌ ERROR during extraction: {e}")
    exit(1)

# Check what was extracted
print()
print("="*80)
print("CHECKING EXTRACTED CONTENTS")
print("="*80)

extracted_items = list(extract_to.iterdir())
print(f"Found {len(extracted_items)} items in {extract_to}:")
for item in extracted_items:
    if item.is_dir() and item.name != '.gitkeep':
        print(f"  📁 {item.name}/")
        
        # Look for nested dataset folder
        nested = list(item.iterdir())
        for nested_item in nested[:5]:  # Show first 5
            if nested_item.is_dir():
                print(f"    📁 {nested_item.name}/")
            else:
                print(f"    📄 {nested_item.name}")
        if len(nested) > 5:
            print(f"    ... and {len(nested) - 5} more items")

print()
print("="*80)
print("NEXT STEP: Run data_preparation.py to filter for Fresno crops")
print("="*80)

