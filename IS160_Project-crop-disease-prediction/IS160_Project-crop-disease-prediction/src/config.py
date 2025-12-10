"""
Configuration file for the Crop Disease Prediction project.
Contains all hyperparameters, paths, and settings.
"""

import os
from pathlib import Path

# Project root directory
PROJECT_ROOT = Path(__file__).parent.parent

# Data paths
DATA_DIR = PROJECT_ROOT / "data"
RAW_DATA_DIR = DATA_DIR / "raw"
PROCESSED_DATA_DIR = DATA_DIR / "processed"
AUGMENTED_DATA_DIR = DATA_DIR / "augmented"

# Model paths
MODELS_DIR = PROJECT_ROOT / "models"
CHECKPOINTS_DIR = MODELS_DIR / "checkpoints"

# Results paths
RESULTS_DIR = PROJECT_ROOT / "results"
FIGURES_DIR = RESULTS_DIR / "figures"
METRICS_DIR = RESULTS_DIR / "metrics"

# Fresno-relevant crops to include
FRESNO_CROPS = [
    "Tomato",
    "Apple",
    "Grape",
    "Peach",
    "Potato",
    "Squash",
    "Strawberry",
    "Orange",
    "Cherry",
    "Pepper",
    "Corn",
]

# Crops to exclude from the dataset
EXCLUDED_CROPS = [
    "Soybean",
    "Raspberry",
    "Blueberry",
]

# Model hyperparameters
MODEL_CONFIG = {
    "architecture": "resnet34",  # Options: resnet34, resnet50, vgg16, custom_cnn
    "pretrained": True,
    "num_classes": None,  # Will be set based on filtered dataset
    "input_size": 224,  # Image size for model input
}

# Training hyperparameters
TRAIN_CONFIG = {
    "batch_size": 32,
    "num_epochs": 50,
    "learning_rate": 0.001,
    "weight_decay": 1e-4,
    "optimizer": "adam",  # Options: adam, sgd, adamw
    "scheduler": "step",  # Options: step, cosine, plateau
    "step_size": 10,  # For StepLR scheduler
    "gamma": 0.1,  # Learning rate decay factor
    "patience": 5,  # For early stopping
    "num_workers": 4,  # DataLoader workers
    "pin_memory": True,
}

# Data split ratios
DATA_SPLIT = {
    "train": 0.7,
    "val": 0.15,
    "test": 0.15,
}

# Data augmentation parameters
AUGMENTATION_CONFIG = {
    "train": {
        "horizontal_flip": 0.5,
        "vertical_flip": 0.3,
        "rotation_range": 30,
        "brightness": 0.2,
        "contrast": 0.2,
        "saturation": 0.2,
        "hue": 0.1,
        "random_crop": True,
        "normalize": True,
    },
    "val_test": {
        "normalize": True,
        "resize": True,
    }
}

# Normalization parameters (ImageNet stats)
NORMALIZATION = {
    "mean": [0.485, 0.456, 0.406],
    "std": [0.229, 0.224, 0.225],
}

# Evaluation thresholds
EVALUATION_CONFIG = {
    "confidence_threshold": 0.5,
    "top_k": 3,  # For top-k accuracy
}

# Random seed for reproducibility
RANDOM_SEED = 42

# Device configuration
DEVICE = "cuda"  # Options: cuda, cpu, mps (for Apple Silicon)

# Logging
LOGGING_CONFIG = {
    "log_interval": 10,  # Log every N batches
    "save_interval": 5,  # Save checkpoint every N epochs
    "use_tensorboard": True,
    "tensorboard_dir": PROJECT_ROOT / "runs",
}

# Create directories if they don't exist
def create_directories():
    """Create all necessary directories for the project."""
    dirs = [
        RAW_DATA_DIR,
        PROCESSED_DATA_DIR,
        AUGMENTED_DATA_DIR,
        CHECKPOINTS_DIR,
        FIGURES_DIR,
        METRICS_DIR,
        LOGGING_CONFIG["tensorboard_dir"],
    ]
    for dir_path in dirs:
        dir_path.mkdir(parents=True, exist_ok=True)
    print("✓ All directories created successfully.")


if __name__ == "__main__":
    create_directories()
    print(f"Project root: {PROJECT_ROOT}")
    print(f"Data directory: {DATA_DIR}")
    print(f"Models directory: {MODELS_DIR}")
    print(f"Results directory: {RESULTS_DIR}")

