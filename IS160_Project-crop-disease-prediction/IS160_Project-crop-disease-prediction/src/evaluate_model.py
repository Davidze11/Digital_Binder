"""
Model evaluation script for crop disease classification.
Evaluates model performance and generates detailed metrics and visualizations.
"""

import argparse
from pathlib import Path
import json

import torch
import torch.nn as nn
from torch.utils.data import DataLoader
from torchvision import datasets, transforms
import numpy as np
from sklearn.metrics import (
    classification_report,
    confusion_matrix,
    accuracy_score,
    precision_recall_fscore_support,
)
from tqdm import tqdm
import matplotlib.pyplot as plt
import seaborn as sns

import config
from train_model import get_model, get_transforms
from utils import (
    set_seed,
    get_device,
    plot_confusion_matrix,
    save_metrics,
)


def load_model(model_path: Path, num_classes: int, architecture: str, device):
    """
    Load trained model from checkpoint.
    
    Args:
        model_path: Path to model checkpoint
        num_classes: Number of classes
        architecture: Model architecture
        device: Device to load model to
        
    Returns:
        Loaded model
    """
    model = get_model(architecture=architecture, num_classes=num_classes, pretrained=False)
    
    checkpoint = torch.load(model_path, map_location=device)
    model.load_state_dict(checkpoint['model_state_dict'])
    model = model.to(device)
    model.eval()
    
    print(f"✓ Model loaded from: {model_path}")
    print(f"  Epoch: {checkpoint['epoch']}")
    print(f"  Val Accuracy: {checkpoint['accuracy']:.4f}")
    
    return model


def get_test_loader(batch_size=32, num_workers=4):
    """
    Create test data loader.
    
    Args:
        batch_size: Batch size
        num_workers: Number of worker processes
        
    Returns:
        Tuple of (test_loader, class_names)
    """
    test_dir = config.PROCESSED_DATA_DIR.parent / 'test'
    
    if not test_dir.exists():
        raise FileNotFoundError(
            f"Test directory not found: {test_dir}\n"
            f"Please run data_preparation.py first."
        )
    
    test_dataset = datasets.ImageFolder(test_dir, transform=get_transforms('test'))
    
    test_loader = DataLoader(
        test_dataset,
        batch_size=batch_size,
        shuffle=False,
        num_workers=num_workers,
        pin_memory=True,
    )
    
    class_names = test_dataset.classes
    
    print(f"\n✓ Test loader created:")
    print(f"  Test: {len(test_dataset)} images, {len(test_loader)} batches")
    print(f"  Number of classes: {len(class_names)}")
    
    return test_loader, class_names


def evaluate_model(model, test_loader, device, class_names):
    """
    Evaluate model on test set and compute detailed metrics.
    
    Args:
        model: Trained model
        test_loader: Test data loader
        device: Device to evaluate on
        class_names: List of class names
        
    Returns:
        Dictionary containing evaluation metrics
    """
    model.eval()
    
    all_labels = []
    all_predictions = []
    all_probabilities = []
    
    print("\n" + "="*60)
    print("EVALUATING MODEL ON TEST SET")
    print("="*60 + "\n")
    
    with torch.no_grad():
        for images, labels in tqdm(test_loader, desc="Evaluating"):
            images = images.to(device)
            
            outputs = model(images)
            probabilities = torch.softmax(outputs, dim=1)
            _, predicted = outputs.max(1)
            
            all_labels.extend(labels.cpu().numpy())
            all_predictions.extend(predicted.cpu().numpy())
            all_probabilities.extend(probabilities.cpu().numpy())
    
    all_labels = np.array(all_labels)
    all_predictions = np.array(all_predictions)
    all_probabilities = np.array(all_probabilities)
    
    # Calculate metrics
    accuracy = accuracy_score(all_labels, all_predictions)
    precision, recall, f1, support = precision_recall_fscore_support(
        all_labels, all_predictions, average='weighted'
    )
    
    # Per-class metrics
    precision_per_class, recall_per_class, f1_per_class, support_per_class = (
        precision_recall_fscore_support(all_labels, all_predictions, average=None)
    )
    
    # Confusion matrix
    cm = confusion_matrix(all_labels, all_predictions)
    
    # Classification report
    report = classification_report(
        all_labels,
        all_predictions,
        target_names=class_names,
        digits=4,
    )
    
    print("\n" + "="*60)
    print("EVALUATION RESULTS")
    print("="*60)
    print(f"\nOverall Metrics:")
    print(f"  Accuracy:  {accuracy:.4f}")
    print(f"  Precision: {precision:.4f}")
    print(f"  Recall:    {recall:.4f}")
    print(f"  F1-Score:  {f1:.4f}")
    print("\n" + "="*60)
    print("\nDetailed Classification Report:")
    print(report)
    print("="*60 + "\n")
    
    # Prepare metrics dictionary
    metrics = {
        'overall': {
            'accuracy': float(accuracy),
            'precision': float(precision),
            'recall': float(recall),
            'f1_score': float(f1),
        },
        'per_class': {},
        'confusion_matrix': cm.tolist(),
        'class_names': class_names,
    }
    
    # Add per-class metrics
    for i, class_name in enumerate(class_names):
        metrics['per_class'][class_name] = {
            'precision': float(precision_per_class[i]),
            'recall': float(recall_per_class[i]),
            'f1_score': float(f1_per_class[i]),
            'support': int(support_per_class[i]),
        }
    
    return metrics, cm, all_labels, all_predictions, all_probabilities


def analyze_disease_detection(metrics, class_names):
    """
    Analyze model's ability to detect diseased vs healthy crops.
    
    Args:
        metrics: Evaluation metrics dictionary
        class_names: List of class names
    """
    print("\n" + "="*60)
    print("DISEASE vs HEALTHY ANALYSIS")
    print("="*60 + "\n")
    
    # Separate healthy and diseased classes
    healthy_classes = []
    diseased_classes = []
    
    for class_name in class_names:
        if 'healthy' in class_name.lower():
            healthy_classes.append(class_name)
        else:
            diseased_classes.append(class_name)
    
    print(f"Healthy classes: {len(healthy_classes)}")
    print(f"Diseased classes: {len(diseased_classes)}")
    print()
    
    # Calculate average performance for healthy vs diseased
    healthy_f1 = []
    diseased_f1 = []
    
    for class_name in healthy_classes:
        if class_name in metrics['per_class']:
            healthy_f1.append(metrics['per_class'][class_name]['f1_score'])
    
    for class_name in diseased_classes:
        if class_name in metrics['per_class']:
            diseased_f1.append(metrics['per_class'][class_name]['f1_score'])
    
    if healthy_f1:
        avg_healthy_f1 = np.mean(healthy_f1)
        print(f"Average F1-Score for HEALTHY classes: {avg_healthy_f1:.4f}")
    
    if diseased_f1:
        avg_diseased_f1 = np.mean(diseased_f1)
        print(f"Average F1-Score for DISEASED classes: {avg_diseased_f1:.4f}")
    
    print("\n" + "="*60 + "\n")


def plot_per_class_performance(metrics, save_path=None):
    """
    Plot per-class F1 scores.
    
    Args:
        metrics: Evaluation metrics dictionary
        save_path: Path to save plot
    """
    class_names = list(metrics['per_class'].keys())
    f1_scores = [metrics['per_class'][name]['f1_score'] for name in class_names]
    
    # Sort by F1 score
    sorted_indices = np.argsort(f1_scores)
    sorted_classes = [class_names[i] for i in sorted_indices]
    sorted_f1 = [f1_scores[i] for i in sorted_indices]
    
    # Create color map (healthy = green, diseased = red)
    colors = ['green' if 'healthy' in name.lower() else 'red' for name in sorted_classes]
    
    plt.figure(figsize=(12, len(class_names) * 0.4))
    bars = plt.barh(range(len(sorted_classes)), sorted_f1, color=colors, alpha=0.7)
    plt.yticks(range(len(sorted_classes)), sorted_classes)
    plt.xlabel('F1-Score')
    plt.title('Per-Class F1-Score Performance')
    plt.xlim([0, 1.0])
    plt.grid(axis='x', alpha=0.3)
    plt.tight_layout()
    
    if save_path:
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        print(f"✓ Per-class performance plot saved to {save_path}")
    
    plt.show()


def identify_worst_performing_classes(metrics, top_n=5):
    """
    Identify the worst performing classes.
    
    Args:
        metrics: Evaluation metrics dictionary
        top_n: Number of worst classes to show
    """
    print("\n" + "="*60)
    print(f"TOP {top_n} WORST PERFORMING CLASSES")
    print("="*60 + "\n")
    
    class_f1 = [(name, data['f1_score']) for name, data in metrics['per_class'].items()]
    class_f1.sort(key=lambda x: x[1])
    
    for i, (class_name, f1_score) in enumerate(class_f1[:top_n], 1):
        class_data = metrics['per_class'][class_name]
        print(f"{i}. {class_name}")
        print(f"   F1: {f1_score:.4f}, Precision: {class_data['precision']:.4f}, "
              f"Recall: {class_data['recall']:.4f}, Support: {class_data['support']}")
    
    print("\n" + "="*60 + "\n")


def simulate_crop_pulling(all_labels, all_predictions, all_probabilities, class_names):
    """
    Simulate the scenario of pulling out diseased crops based on model predictions.
    
    Args:
        all_labels: True labels
        all_predictions: Predicted labels
        all_probabilities: Prediction probabilities
        class_names: List of class names
    """
    print("\n" + "="*60)
    print("CROP PULLING SIMULATION")
    print("="*60 + "\n")
    
    # Determine which classes are diseased vs healthy
    healthy_indices = [i for i, name in enumerate(class_names) if 'healthy' in name.lower()]
    diseased_indices = [i for i, name in enumerate(class_names) if 'healthy' not in name.lower()]
    
    # Convert to binary (diseased=1, healthy=0)
    true_binary = np.array([1 if label in diseased_indices else 0 for label in all_labels])
    pred_binary = np.array([1 if pred in diseased_indices else 0 for pred in all_predictions])
    
    # Calculate metrics
    tp = np.sum((true_binary == 1) & (pred_binary == 1))  # Correctly identified diseased
    tn = np.sum((true_binary == 0) & (pred_binary == 0))  # Correctly identified healthy
    fp = np.sum((true_binary == 0) & (pred_binary == 1))  # False alarm (healthy marked as diseased)
    fn = np.sum((true_binary == 1) & (pred_binary == 0))  # Missed diseased crops
    
    total_diseased = np.sum(true_binary == 1)
    total_healthy = np.sum(true_binary == 0)
    total_samples = len(true_binary)
    
    print(f"Total samples: {total_samples}")
    print(f"  Diseased: {total_diseased} ({total_diseased/total_samples*100:.1f}%)")
    print(f"  Healthy: {total_healthy} ({total_healthy/total_samples*100:.1f}%)")
    print()
    print("Confusion Matrix (Binary: Diseased vs Healthy):")
    print(f"  True Positives (Diseased correctly flagged):  {tp}")
    print(f"  True Negatives (Healthy correctly kept):      {tn}")
    print(f"  False Positives (Healthy wrongly flagged):    {fp}")
    print(f"  False Negatives (Diseased missed):            {fn}")
    print()
    
    # Calculate rates
    if total_diseased > 0:
        detection_rate = tp / total_diseased
        miss_rate = fn / total_diseased
        print(f"Disease Detection Rate: {detection_rate:.2%} (caught {tp}/{total_diseased} diseased crops)")
        print(f"Disease Miss Rate: {miss_rate:.2%} (missed {fn}/{total_diseased} diseased crops)")
    
    if total_healthy > 0:
        false_positive_rate = fp / total_healthy
        print(f"False Positive Rate: {false_positive_rate:.2%} (wrongly flagged {fp}/{total_healthy} healthy crops)")
    
    print("\n" + "="*60 + "\n")


def main(args):
    """Main evaluation function."""
    
    # Set seed
    set_seed(config.RANDOM_SEED)
    
    # Get device
    device = get_device()
    
    # Get test loader
    test_loader, class_names = get_test_loader(
        batch_size=args.batch_size,
        num_workers=args.num_workers,
    )
    
    # Load model
    model = load_model(
        model_path=Path(args.model_path),
        num_classes=len(class_names),
        architecture=args.architecture,
        device=device,
    )
    
    # Evaluate model
    metrics, cm, all_labels, all_predictions, all_probabilities = evaluate_model(
        model, test_loader, device, class_names
    )
    
    # Save metrics
    metrics_path = config.METRICS_DIR / f"{args.architecture}_test_metrics.json"
    save_metrics(metrics, metrics_path)
    
    # Plot confusion matrix
    cm_path = config.FIGURES_DIR / f"{args.architecture}_confusion_matrix.png"
    plot_confusion_matrix(cm, class_names, save_path=cm_path, figsize=(16, 14))
    
    # Plot per-class performance
    perf_path = config.FIGURES_DIR / f"{args.architecture}_per_class_performance.png"
    plot_per_class_performance(metrics, save_path=perf_path)
    
    # Analyze disease detection
    analyze_disease_detection(metrics, class_names)
    
    # Identify worst performing classes
    identify_worst_performing_classes(metrics, top_n=10)
    
    # Simulate crop pulling
    simulate_crop_pulling(all_labels, all_predictions, all_probabilities, class_names)
    
    print("✓ Evaluation complete!")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Evaluate crop disease classification model")
    
    parser.add_argument(
        '--model-path',
        type=str,
        required=True,
        help='Path to trained model checkpoint',
    )
    parser.add_argument(
        '--architecture',
        type=str,
        default='resnet34',
        choices=['resnet34', 'resnet50', 'vgg16'],
        help='Model architecture',
    )
    parser.add_argument(
        '--batch-size',
        type=int,
        default=32,
        help='Batch size',
    )
    parser.add_argument(
        '--num-workers',
        type=int,
        default=4,
        help='Number of data loader workers',
    )
    
    args = parser.parse_args()
    main(args)

