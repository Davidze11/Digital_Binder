"""
Master script to prepare data and start training the plant disease classifier.
This script handles the complete training pipeline with continuous learning support.
"""

import os
import sys
import subprocess
from pathlib import Path

# Add src to path
sys.path.append(str(Path(__file__).parent / 'src'))

import config
import argparse


def check_data_prepared():
    """Check if data has been prepared."""
    train_dir = config.PROCESSED_DATA_DIR.parent / 'train'
    val_dir = config.PROCESSED_DATA_DIR.parent / 'val'
    test_dir = config.PROCESSED_DATA_DIR.parent / 'test'
    
    return train_dir.exists() and val_dir.exists() and test_dir.exists()


def main():
    """Main function to run the complete training pipeline."""
    print("="*80)
    print("PLANT DISEASE CLASSIFICATION - CONTINUOUS LEARNING TRAINING SYSTEM")
    print("="*80)
    print("\nThis system implements continuous learning with:")
    print("  [+] Automatic checkpoint saving after each epoch")
    print("  [+] Automatic resume from last checkpoint")
    print("  [+] Training history preservation across sessions")
    print("  [+] Best model tracking and saving")
    print("  [+] Early stopping to prevent overfitting")
    print("="*80 + "\n")
    
    # Create necessary directories
    config.create_directories()
    
    # Check if data is prepared
    if not check_data_prepared():
        print("\n[!] Data not prepared. Preparing data first...")
        print("This may take a few minutes...\n")
        
        # Prepare data - call it directly with subprocess to avoid import issues
        result = subprocess.run(
            [sys.executable, str(Path(__file__).parent / 'src' / 'data_preparation.py')],
            capture_output=False
        )
        
        if result.returncode != 0:
            print("\n[!] Error during data preparation")
            return
        
        print("\n[+] Data preparation complete!")
    else:
        print("[+] Data already prepared")
    
    # Check for existing checkpoints
    checkpoint_dir = config.CHECKPOINTS_DIR
    existing_checkpoints = list(checkpoint_dir.glob("*_latest.pth"))
    
    if existing_checkpoints:
        print(f"\n📁 Found {len(existing_checkpoints)} existing checkpoint(s):")
        for ckpt in existing_checkpoints:
            print(f"  - {ckpt.name}")
        print("\n[+] Training will automatically resume from the latest checkpoint")
    else:
        print("\n[+] No existing checkpoints found. Starting fresh training.")
    
    # Start training
    print("\n" + "="*80)
    print("STARTING TRAINING")
    print("="*80 + "\n")
    
    # Set up training arguments
    parser = argparse.ArgumentParser(description="Train plant disease classifier")
    
    # Model arguments
    parser.add_argument(
        '--architecture',
        type=str,
        default='resnet34',
        choices=['resnet34', 'resnet50', 'vgg16'],
        help='Model architecture (default: resnet34)',
    )
    parser.add_argument(
        '--pretrained',
        action='store_true',
        default=True,
        help='Use pretrained weights (default: True)',
    )
    
    # Training arguments
    parser.add_argument(
        '--epochs',
        type=int,
        default=30,
        help='Number of epochs to train (default: 30)',
    )
    parser.add_argument(
        '--batch-size',
        type=int,
        default=32,
        help='Batch size (default: 32)',
    )
    parser.add_argument(
        '--learning-rate',
        type=float,
        default=0.001,
        help='Learning rate (default: 0.001)',
    )
    parser.add_argument(
        '--resume',
        action='store_true',
        default=True,
        help='Resume from latest checkpoint (default: True)',
    )
    parser.add_argument(
        '--use-tensorboard',
        action='store_true',
        default=True,
        help='Use TensorBoard logging (default: True)',
    )
    
    args = parser.parse_args()
    
    print(f"Training Configuration:")
    print(f"  Architecture: {args.architecture}")
    print(f"  Pretrained: {args.pretrained}")
    print(f"  Total Epochs: {args.epochs}")
    print(f"  Batch Size: {args.batch_size}")
    print(f"  Learning Rate: {args.learning_rate}")
    print(f"  Resume Training: {args.resume}")
    print(f"  TensorBoard: {args.use_tensorboard}")
    print()
    
    # Train model - call it via subprocess as well
    print("\n" + "="*80)
    print("STARTING TRAINING")
    print("="*80 + "\n")
    
    # Set up training arguments as list
    train_args = [
        sys.executable,
        str(Path(__file__).parent / 'src' / 'train_model.py'),
        '--architecture', args.architecture,
        '--epochs', str(args.epochs),
        '--batch-size', str(args.batch_size),
        '--learning-rate', str(args.learning_rate),
    ]
    
    if args.pretrained:
        train_args.append('--pretrained')
    
    if args.resume:
        train_args.append('--resume')
    
    if args.use_tensorboard:
        train_args.append('--use-tensorboard')
    
    result = subprocess.run(train_args, capture_output=False)
    
    if result.returncode != 0:
        print("\n[!] Error during training")
        return
    
    print("\n" + "="*80)
    print("TRAINING PIPELINE COMPLETE")
    print("="*80)
    print("\nYour model has been trained and saved!")
    print(f"  Checkpoints: {config.CHECKPOINTS_DIR}")
    print(f"  Training plots: {config.FIGURES_DIR}")
    if args.use_tensorboard:
        print(f"  TensorBoard logs: {config.LOGGING_CONFIG['tensorboard_dir']}")
        print("\nTo view TensorBoard, run:")
        print(f"  tensorboard --logdir={config.LOGGING_CONFIG['tensorboard_dir']}")
    
    # Generate educational summary
    print("\n" + "="*80)
    print("GENERATING EDUCATIONAL TRAINING SUMMARY")
    print("="*80)
    print("\nCreating beginner-friendly reports and visualizations...")
    
    summary_result = subprocess.run(
        [sys.executable, str(Path(__file__).parent / 'src' / 'training_summary.py')],
        capture_output=False
    )
    
    if summary_result.returncode == 0:
        print("\n[+] Educational summary generated successfully!")
        print("\nCheck these files:")
        print(f"  • Report: {config.RESULTS_DIR / 'TRAINING_REPORT_FOR_BEGINNERS.txt'}")
        print(f"  • Charts: {config.FIGURES_DIR}")
    
    print("\nTo continue training, simply run this script again!")
    print("The model will automatically resume from where it left off.")
    print("="*80 + "\n")


if __name__ == "__main__":
    main()

