"""
Data preparation script for filtering and preprocessing the Plant Disease dataset.
This script filters out non-Fresno crops and prepares the data for training.
"""

import os
import shutil
import random
from pathlib import Path
from typing import List, Dict, Tuple
from collections import defaultdict
import argparse

import config
from utils import set_seed


def get_dataset_structure(data_dir: Path) -> Dict[str, List[str]]:
    """
    Analyze the structure of the raw dataset.
    
    Args:
        data_dir: Path to raw data directory
        
    Returns:
        Dictionary mapping crop names to their disease classes
    """
    dataset_structure = defaultdict(list)
    
    if not data_dir.exists():
        print(f"[!] Warning: Data directory not found: {data_dir}")
        print("Please download the Plant Disease dataset and place it in data/raw/")
        return dataset_structure
    
    # Assuming structure: data/raw/PlantVillage/{Crop}_{Disease}/
    for class_dir in data_dir.iterdir():
        if class_dir.is_dir():
            class_name = class_dir.name
            
            # Parse crop name (before first underscore or hyphen)
            if '_' in class_name:
                crop_name = class_name.split('_')[0]
            else:
                crop_name = class_name
            
            dataset_structure[crop_name].append(class_name)
    
    return dataset_structure


def filter_fresno_crops(
    raw_data_dir: Path,
    processed_data_dir: Path,
    fresno_crops: List[str],
) -> Dict[str, int]:
    """
    Filter dataset to include only Fresno-relevant crops.
    
    Args:
        raw_data_dir: Path to raw data
        processed_data_dir: Path to save processed data
        fresno_crops: List of crops to include
        
    Returns:
        Dictionary with statistics about filtered data
    """
    stats = {
        'total_classes': 0,
        'total_images': 0,
        'classes_per_crop': defaultdict(int),
        'images_per_crop': defaultdict(int),
    }
    
    # Create processed data directory
    processed_data_dir.mkdir(parents=True, exist_ok=True)
    
    print("\n" + "="*60)
    print("FILTERING DATASET FOR FRESNO-RELEVANT CROPS")
    print("="*60)
    
    # Iterate through all class directories
    for class_dir in raw_data_dir.iterdir():
        if not class_dir.is_dir():
            continue
        
        class_name = class_dir.name
        
        # Determine crop name
        if '___' in class_name:
            crop_name = class_name.split('___')[0]
        elif '_' in class_name:
            crop_name = class_name.split('_')[0]
        else:
            crop_name = class_name
        
        # Check if crop is in Fresno list (case-insensitive)
        crop_match = None
        for fresno_crop in fresno_crops:
            if crop_name.lower() == fresno_crop.lower():
                crop_match = fresno_crop
                break
        
        if crop_match:
            # Copy this class to processed directory
            dest_dir = processed_data_dir / class_name
            
            if dest_dir.exists():
                print(f"  [!] {class_name} already exists, skipping...")
                continue
            
            # Count images
            image_files = list(class_dir.glob('*.jpg')) + list(class_dir.glob('*.png')) + list(class_dir.glob('*.jpeg'))
            num_images = len(image_files)
            
            if num_images > 0:
                shutil.copytree(class_dir, dest_dir)
                print(f"  [+] Copied {class_name}: {num_images} images")
                
                stats['total_classes'] += 1
                stats['total_images'] += num_images
                stats['classes_per_crop'][crop_match] += 1
                stats['images_per_crop'][crop_match] += num_images
        else:
            print(f"  [-] Excluded {class_name} (not in Fresno crops)")
    
    print("\n" + "="*60)
    print("FILTERING COMPLETE")
    print("="*60)
    print(f"Total classes retained: {stats['total_classes']}")
    print(f"Total images retained: {stats['total_images']}")
    print("\nImages per crop:")
    for crop, count in sorted(stats['images_per_crop'].items()):
        print(f"  {crop}: {count} images ({stats['classes_per_crop'][crop]} classes)")
    print("="*60 + "\n")
    
    return stats


def split_dataset(
    processed_data_dir: Path,
    train_ratio: float = 0.7,
    val_ratio: float = 0.15,
    test_ratio: float = 0.15,
    seed: int = 42,
) -> Dict[str, Dict[str, int]]:
    """
    Split dataset into train, validation, and test sets.
    
    Args:
        processed_data_dir: Path to processed data
        train_ratio: Ratio for training set
        val_ratio: Ratio for validation set
        test_ratio: Ratio for test set
        seed: Random seed
        
    Returns:
        Dictionary with split statistics
    """
    set_seed(seed)
    
    split_stats = {
        'train': defaultdict(int),
        'val': defaultdict(int),
        'test': defaultdict(int),
    }
    
    print("\n" + "="*60)
    print("SPLITTING DATASET INTO TRAIN/VAL/TEST")
    print("="*60)
    print(f"Split ratios - Train: {train_ratio}, Val: {val_ratio}, Test: {test_ratio}")
    print("="*60 + "\n")
    
    # Create split directories
    for split in ['train', 'val', 'test']:
        split_dir = processed_data_dir.parent / split
        split_dir.mkdir(parents=True, exist_ok=True)
    
    # Iterate through each class
    for class_dir in processed_data_dir.iterdir():
        if not class_dir.is_dir():
            continue
        
        class_name = class_dir.name
        
        # Get all image files
        image_files = (
            list(class_dir.glob('*.jpg')) +
            list(class_dir.glob('*.png')) +
            list(class_dir.glob('*.jpeg')) +
            list(class_dir.glob('*.JPG'))
        )
        
        # Shuffle images
        random.shuffle(image_files)
        
        total_images = len(image_files)
        train_count = int(total_images * train_ratio)
        val_count = int(total_images * val_ratio)
        
        # Split indices
        train_images = image_files[:train_count]
        val_images = image_files[train_count:train_count + val_count]
        test_images = image_files[train_count + val_count:]
        
        # Copy images to respective splits
        for split, images in [('train', train_images), ('val', val_images), ('test', test_images)]:
            split_class_dir = processed_data_dir.parent / split / class_name
            split_class_dir.mkdir(parents=True, exist_ok=True)
            
            for img_path in images:
                dest_path = split_class_dir / img_path.name
                shutil.copy2(img_path, dest_path)
            
            split_stats[split][class_name] = len(images)
        
        print(f"  {class_name}: Train={len(train_images)}, Val={len(val_images)}, Test={len(test_images)}")
    
    print("\n" + "="*60)
    print("SPLITTING COMPLETE")
    print("="*60)
    for split in ['train', 'val', 'test']:
        total = sum(split_stats[split].values())
        print(f"{split.upper()}: {total} images across {len(split_stats[split])} classes")
    print("="*60 + "\n")
    
    return split_stats


def main(args):
    """Main function to run data preparation."""
    
    print("\n" + "="*60)
    print("CROP DISEASE PREDICTION - DATA PREPARATION")
    print("="*60 + "\n")
    
    # Set seed for reproducibility
    set_seed(config.RANDOM_SEED)
    
    # Get dataset structure
    print("Analyzing raw dataset structure...")
    # Use the actual location of the dataset
    raw_data_path = config.RAW_DATA_DIR / "New Plant Diseases Dataset(Augmented)" / "New Plant Diseases Dataset(Augmented)" / "train"
    dataset_structure = get_dataset_structure(raw_data_path)
    
    if not dataset_structure:
        print("\n[!] ERROR: No data found in raw directory.")
        print(f"Please download the dataset and extract it to: {config.RAW_DATA_DIR}")
        return
    
    print(f"\nFound {len(dataset_structure)} crop types in raw dataset:")
    for crop, classes in sorted(dataset_structure.items()):
        print(f"  {crop}: {len(classes)} classes")
    
    # Filter for Fresno crops
    if args.skip_filter:
        print("\n[!] Skipping filtering step (--skip-filter flag set)")
    else:
        stats = filter_fresno_crops(
            raw_data_path,
            config.PROCESSED_DATA_DIR,
            config.FRESNO_CROPS,
        )
    
    # Split dataset
    if args.skip_split:
        print("\n[!] Skipping split step (--skip-split flag set)")
    else:
        split_stats = split_dataset(
            config.PROCESSED_DATA_DIR,
            train_ratio=config.DATA_SPLIT['train'],
            val_ratio=config.DATA_SPLIT['val'],
            test_ratio=config.DATA_SPLIT['test'],
            seed=config.RANDOM_SEED,
        )
    
    print("\n[+] Data preparation complete!")
    print(f"  Processed data: {config.PROCESSED_DATA_DIR}")
    print(f"  Train data: {config.PROCESSED_DATA_DIR.parent / 'train'}")
    print(f"  Val data: {config.PROCESSED_DATA_DIR.parent / 'val'}")
    print(f"  Test data: {config.PROCESSED_DATA_DIR.parent / 'test'}")
    print()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Prepare Plant Disease dataset for training")
    parser.add_argument(
        "--skip-filter",
        action="store_true",
        help="Skip filtering step (if already done)",
    )
    parser.add_argument(
        "--skip-split",
        action="store_true",
        help="Skip splitting step (if already done)",
    )
    
    args = parser.parse_args()
    main(args)

