"""
Utility functions for the Crop Disease Prediction project.
"""

import os
import random
import numpy as np
import torch
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path
from typing import Dict, List, Tuple
import json


def set_seed(seed: int = 42):
    """
    Set random seed for reproducibility.
    
    Args:
        seed: Random seed value
    """
    random.seed(seed)
    np.random.seed(seed)
    torch.manual_seed(seed)
    if torch.cuda.is_available():
        torch.cuda.manual_seed(seed)
        torch.cuda.manual_seed_all(seed)
        torch.backends.cudnn.deterministic = True
        torch.backends.cudnn.benchmark = False
    print(f"[+] Random seed set to {seed}")


def get_device():
    """
    Get the best available device (CUDA, MPS, or CPU).
    
    Returns:
        torch.device: The device to use for training
    """
    if torch.cuda.is_available():
        device = torch.device("cuda")
        print(f"[+] Using CUDA: {torch.cuda.get_device_name(0)}")
    elif torch.backends.mps.is_available():
        device = torch.device("mps")
        print("[+] Using Apple Silicon MPS")
    else:
        device = torch.device("cpu")
        print("[+] Using CPU")
    return device


def count_parameters(model: torch.nn.Module) -> int:
    """
    Count the number of trainable parameters in a model.
    
    Args:
        model: PyTorch model
        
    Returns:
        int: Number of trainable parameters
    """
    return sum(p.numel() for p in model.parameters() if p.requires_grad)


def save_checkpoint(
    model: torch.nn.Module,
    optimizer: torch.optim.Optimizer,
    epoch: int,
    loss: float,
    accuracy: float,
    filepath: Path,
    scheduler: torch.optim.lr_scheduler._LRScheduler = None,
):
    """
    Save model checkpoint with full training state.
    
    Args:
        model: PyTorch model
        optimizer: Optimizer
        epoch: Current epoch
        loss: Current loss
        accuracy: Current accuracy
        filepath: Path to save checkpoint
        scheduler: Optional learning rate scheduler
    """
    checkpoint = {
        "epoch": epoch,
        "model_state_dict": model.state_dict(),
        "optimizer_state_dict": optimizer.state_dict(),
        "scheduler_state_dict": scheduler.state_dict() if scheduler else None,
        "loss": loss,
        "accuracy": accuracy,
    }
    torch.save(checkpoint, filepath)
    # Don't print save message for every checkpoint (too verbose)


def load_checkpoint(
    model: torch.nn.Module,
    optimizer: torch.optim.Optimizer,
    filepath: Path,
    device: torch.device,
    scheduler: torch.optim.lr_scheduler._LRScheduler = None,
) -> Tuple[torch.nn.Module, torch.optim.Optimizer, int, float]:
    """
    Load model checkpoint with full training state.
    
    Args:
        model: PyTorch model
        optimizer: Optimizer
        filepath: Path to checkpoint
        device: Device to load model to
        scheduler: Optional learning rate scheduler
        
    Returns:
        Tuple of (model, optimizer, epoch, loss)
    """
    checkpoint = torch.load(filepath, map_location=device)
    model.load_state_dict(checkpoint["model_state_dict"])
    optimizer.load_state_dict(checkpoint["optimizer_state_dict"])
    
    if scheduler and "scheduler_state_dict" in checkpoint and checkpoint["scheduler_state_dict"]:
        scheduler.load_state_dict(checkpoint["scheduler_state_dict"])
    
    epoch = checkpoint["epoch"]
    loss = checkpoint["loss"]
    accuracy = checkpoint.get("accuracy", 0.0)
    
    print(f"[+] Checkpoint loaded from {filepath}")
    print(f"  Epoch: {epoch}, Loss: {loss:.4f}, Accuracy: {accuracy:.4f}")
    
    return model, optimizer, epoch, loss


def plot_training_history(
    history: Dict[str, List[float]],
    save_path: Path = None,
):
    """
    Plot training and validation metrics over epochs.
    
    Args:
        history: Dictionary containing training metrics
        save_path: Optional path to save the plot
    """
    fig, axes = plt.subplots(1, 2, figsize=(15, 5))
    
    # Plot loss
    axes[0].plot(history['train_loss'], label='Train Loss', marker='o')
    axes[0].plot(history['val_loss'], label='Val Loss', marker='s')
    axes[0].set_xlabel('Epoch')
    axes[0].set_ylabel('Loss')
    axes[0].set_title('Training and Validation Loss')
    axes[0].legend()
    axes[0].grid(True, alpha=0.3)
    
    # Plot accuracy
    axes[1].plot(history['train_acc'], label='Train Accuracy', marker='o')
    axes[1].plot(history['val_acc'], label='Val Accuracy', marker='s')
    axes[1].set_xlabel('Epoch')
    axes[1].set_ylabel('Accuracy (%)')
    axes[1].set_title('Training and Validation Accuracy')
    axes[1].legend()
    axes[1].grid(True, alpha=0.3)
    
    plt.tight_layout()
    
    if save_path:
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        print(f"[+] Training history plot saved to {save_path}")
    
    plt.show()


def plot_confusion_matrix(
    cm: np.ndarray,
    class_names: List[str],
    save_path: Path = None,
    figsize: Tuple[int, int] = (12, 10),
):
    """
    Plot confusion matrix with annotations.
    
    Args:
        cm: Confusion matrix as numpy array
        class_names: List of class names
        save_path: Optional path to save the plot
        figsize: Figure size
    """
    plt.figure(figsize=figsize)
    
    # Normalize confusion matrix
    cm_normalized = cm.astype('float') / cm.sum(axis=1)[:, np.newaxis]
    
    sns.heatmap(
        cm_normalized,
        annot=True,
        fmt='.2f',
        cmap='Blues',
        xticklabels=class_names,
        yticklabels=class_names,
        cbar_kws={'label': 'Percentage'},
    )
    
    plt.xlabel('Predicted Label')
    plt.ylabel('True Label')
    plt.title('Confusion Matrix (Normalized)')
    plt.xticks(rotation=45, ha='right')
    plt.yticks(rotation=0)
    plt.tight_layout()
    
    if save_path:
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        print(f"[+] Confusion matrix saved to {save_path}")
    
    plt.show()


def save_metrics(metrics: Dict, filepath: Path):
    """
    Save evaluation metrics to JSON file.
    
    Args:
        metrics: Dictionary of metrics
        filepath: Path to save metrics
    """
    with open(filepath, 'w') as f:
        json.dump(metrics, f, indent=4)
    print(f"[+] Metrics saved to {filepath}")


def load_metrics(filepath: Path) -> Dict:
    """
    Load evaluation metrics from JSON file.
    
    Args:
        filepath: Path to metrics file
        
    Returns:
        Dictionary of metrics
    """
    with open(filepath, 'r') as f:
        metrics = json.load(f)
    return metrics


def format_time(seconds: float) -> str:
    """
    Format time in seconds to human-readable string.
    
    Args:
        seconds: Time in seconds
        
    Returns:
        Formatted time string
    """
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    secs = int(seconds % 60)
    
    if hours > 0:
        return f"{hours}h {minutes}m {secs}s"
    elif minutes > 0:
        return f"{minutes}m {secs}s"
    else:
        return f"{secs}s"


class AverageMeter:
    """Computes and stores the average and current value."""
    
    def __init__(self):
        self.reset()
    
    def reset(self):
        self.val = 0
        self.avg = 0
        self.sum = 0
        self.count = 0
    
    def update(self, val, n=1):
        self.val = val
        self.sum += val * n
        self.count += n
        self.avg = self.sum / self.count

