"""
Download the Plant Disease dataset from Kaggle using kagglehub.
"""

import kagglehub
import shutil
from pathlib import Path
import config

print("="*80)
print("DOWNLOADING PLANT DISEASE DATASET FROM KAGGLE")
print("="*80)
print()

# Download latest version
print("Starting download... (this may take several minutes)")
path = kagglehub.dataset_download("vipoooool/new-plant-diseases-dataset")

print(f"\n✓ Dataset downloaded to: {path}")
print()

# The dataset will be downloaded to a kaggle cache directory
# Let's check what's in it and provide instructions
downloaded_path = Path(path)
print(f"Contents of downloaded dataset:")
for item in downloaded_path.iterdir():
    if item.is_dir():
        print(f"  📁 {item.name}/")
        # Show subdirectories
        for subitem in item.iterdir():
            if subitem.is_dir():
                print(f"    📁 {subitem.name}/")
    else:
        print(f"  📄 {item.name}")

print()
print("="*80)
print("NEXT STEPS")
print("="*80)
print(f"The dataset has been downloaded to Kaggle's cache: {path}")
print(f"")
print(f"You can either:")
print(f"1. Run data_preparation.py and point it to this location, OR")
print(f"2. Copy the dataset to data/raw/ in this project")
print()
print(f"Recommended: Update config.py to use the downloaded path directly")
print(f"or copy with:")
print(f'  RAW_DATA_DIR = Path(r"{path}")')
print("="*80)

