"""
Model training script for crop disease classification.
"""

import argparse
import time
from pathlib import Path
from datetime import datetime

import torch
import torch.nn as nn
import torch.optim as optim
from torch.utils.data import DataLoader
from torch.utils.tensorboard import SummaryWriter
from torchvision import datasets, transforms, models
from tqdm import tqdm

import config
from utils import (
    set_seed,
    get_device,
    save_checkpoint,
    count_parameters,
    plot_training_history,
    AverageMeter,
    format_time,
)


def get_transforms(split='train'):
    """
    Get data transforms for training, validation, or test.
    
    Args:
        split: 'train', 'val', or 'test'
        
    Returns:
        torchvision.transforms.Compose object
    """
    if split == 'train':
        return transforms.Compose([
            transforms.Resize((config.MODEL_CONFIG['input_size'], config.MODEL_CONFIG['input_size'])),
            transforms.RandomHorizontalFlip(p=config.AUGMENTATION_CONFIG['train']['horizontal_flip']),
            transforms.RandomVerticalFlip(p=config.AUGMENTATION_CONFIG['train']['vertical_flip']),
            transforms.RandomRotation(config.AUGMENTATION_CONFIG['train']['rotation_range']),
            transforms.ColorJitter(
                brightness=config.AUGMENTATION_CONFIG['train']['brightness'],
                contrast=config.AUGMENTATION_CONFIG['train']['contrast'],
                saturation=config.AUGMENTATION_CONFIG['train']['saturation'],
                hue=config.AUGMENTATION_CONFIG['train']['hue'],
            ),
            transforms.ToTensor(),
            transforms.Normalize(
                mean=config.NORMALIZATION['mean'],
                std=config.NORMALIZATION['std'],
            ),
        ])
    else:  # val or test
        return transforms.Compose([
            transforms.Resize((config.MODEL_CONFIG['input_size'], config.MODEL_CONFIG['input_size'])),
            transforms.ToTensor(),
            transforms.Normalize(
                mean=config.NORMALIZATION['mean'],
                std=config.NORMALIZATION['std'],
            ),
        ])


def get_dataloaders(batch_size=32, num_workers=4):
    """
    Create data loaders for train, validation, and test sets.
    
    Args:
        batch_size: Batch size
        num_workers: Number of worker processes
        
    Returns:
        Tuple of (train_loader, val_loader, test_loader, num_classes, class_names)
    """
    data_dir = config.PROCESSED_DATA_DIR.parent
    
    train_dir = data_dir / 'train'
    val_dir = data_dir / 'val'
    test_dir = data_dir / 'test'
    
    # Check if directories exist
    for split_dir, name in [(train_dir, 'train'), (val_dir, 'val'), (test_dir, 'test')]:
        if not split_dir.exists():
            raise FileNotFoundError(
                f"{name.capitalize()} directory not found: {split_dir}\n"
                f"Please run data_preparation.py first."
            )
    
    # Create datasets
    train_dataset = datasets.ImageFolder(train_dir, transform=get_transforms('train'))
    val_dataset = datasets.ImageFolder(val_dir, transform=get_transforms('val'))
    test_dataset = datasets.ImageFolder(test_dir, transform=get_transforms('test'))
    
    # Create data loaders
    train_loader = DataLoader(
        train_dataset,
        batch_size=batch_size,
        shuffle=True,
        num_workers=num_workers,
        pin_memory=config.TRAIN_CONFIG['pin_memory'],
    )
    
    val_loader = DataLoader(
        val_dataset,
        batch_size=batch_size,
        shuffle=False,
        num_workers=num_workers,
        pin_memory=config.TRAIN_CONFIG['pin_memory'],
    )
    
    test_loader = DataLoader(
        test_dataset,
        batch_size=batch_size,
        shuffle=False,
        num_workers=num_workers,
        pin_memory=config.TRAIN_CONFIG['pin_memory'],
    )
    
    num_classes = len(train_dataset.classes)
    class_names = train_dataset.classes
    
    print(f"\n[+] Data loaders created:")
    print(f"  Train: {len(train_dataset)} images, {len(train_loader)} batches")
    print(f"  Val: {len(val_dataset)} images, {len(val_loader)} batches")
    print(f"  Test: {len(test_dataset)} images, {len(test_loader)} batches")
    print(f"  Number of classes: {num_classes}")
    
    return train_loader, val_loader, test_loader, num_classes, class_names


def get_model(architecture='resnet34', num_classes=38, pretrained=True):
    """
    Create model based on specified architecture.
    
    Args:
        architecture: Model architecture name
        num_classes: Number of output classes
        pretrained: Whether to use pretrained weights
        
    Returns:
        PyTorch model
    """
    print(f"\n[+] Creating model: {architecture}")
    print(f"  Pretrained: {pretrained}")
    print(f"  Number of classes: {num_classes}")
    
    if architecture == 'resnet34':
        model = models.resnet34(pretrained=pretrained)
        num_features = model.fc.in_features
        model.fc = nn.Linear(num_features, num_classes)
    
    elif architecture == 'resnet50':
        model = models.resnet50(pretrained=pretrained)
        num_features = model.fc.in_features
        model.fc = nn.Linear(num_features, num_classes)
    
    elif architecture == 'vgg16':
        model = models.vgg16(pretrained=pretrained)
        num_features = model.classifier[6].in_features
        model.classifier[6] = nn.Linear(num_features, num_classes)
    
    else:
        raise ValueError(f"Unknown architecture: {architecture}")
    
    print(f"  Total parameters: {count_parameters(model):,}")
    
    return model


def train_one_epoch(model, train_loader, criterion, optimizer, device, epoch):
    """
    Train model for one epoch.
    
    Args:
        model: PyTorch model
        train_loader: Training data loader
        criterion: Loss function
        optimizer: Optimizer
        device: Device to train on
        epoch: Current epoch number
        
    Returns:
        Tuple of (average loss, average accuracy)
    """
    model.train()
    
    losses = AverageMeter()
    accuracies = AverageMeter()
    
    pbar = tqdm(train_loader, desc=f"Epoch {epoch} [Train]")
    
    for images, labels in pbar:
        images, labels = images.to(device), labels.to(device)
        
        # Forward pass
        outputs = model(images)
        loss = criterion(outputs, labels)
        
        # Backward pass
        optimizer.zero_grad()
        loss.backward()
        optimizer.step()
        
        # Calculate accuracy
        _, predicted = outputs.max(1)
        accuracy = predicted.eq(labels).sum().item() / labels.size(0)
        
        # Update meters
        losses.update(loss.item(), images.size(0))
        accuracies.update(accuracy, images.size(0))
        
        # Update progress bar
        pbar.set_postfix({
            'loss': f'{losses.avg:.4f}',
            'acc': f'{accuracies.avg:.4f}',
        })
    
    return losses.avg, accuracies.avg


def validate(model, val_loader, criterion, device, epoch):
    """
    Validate model.
    
    Args:
        model: PyTorch model
        val_loader: Validation data loader
        criterion: Loss function
        device: Device to validate on
        epoch: Current epoch number
        
    Returns:
        Tuple of (average loss, average accuracy)
    """
    model.eval()
    
    losses = AverageMeter()
    accuracies = AverageMeter()
    
    pbar = tqdm(val_loader, desc=f"Epoch {epoch} [Val]  ")
    
    with torch.no_grad():
        for images, labels in pbar:
            images, labels = images.to(device), labels.to(device)
            
            # Forward pass
            outputs = model(images)
            loss = criterion(outputs, labels)
            
            # Calculate accuracy
            _, predicted = outputs.max(1)
            accuracy = predicted.eq(labels).sum().item() / labels.size(0)
            
            # Update meters
            losses.update(loss.item(), images.size(0))
            accuracies.update(accuracy, images.size(0))
            
            # Update progress bar
            pbar.set_postfix({
                'loss': f'{losses.avg:.4f}',
                'acc': f'{accuracies.avg:.4f}',
            })
    
    return losses.avg, accuracies.avg


def load_latest_checkpoint(args, model, optimizer, device):
    """
    Load the latest checkpoint if it exists for continuous learning.
    
    Args:
        args: Command line arguments
        model: PyTorch model
        optimizer: Optimizer
        device: Device to load checkpoint to
        
    Returns:
        Tuple of (start_epoch, best_val_acc, history)
    """
    from utils import load_checkpoint
    import json
    
    # Check for resume checkpoint
    checkpoint_path = config.CHECKPOINTS_DIR / f"{args.architecture}_latest.pth"
    history_path = config.CHECKPOINTS_DIR / f"{args.architecture}_history.json"
    
    if args.resume and checkpoint_path.exists():
        print("\n" + "="*60)
        print("RESUMING FROM CHECKPOINT")
        print("="*60)
        
        # Load checkpoint
        model, optimizer, last_epoch, last_loss = load_checkpoint(
            model, optimizer, checkpoint_path, device
        )
        
        # Load history if exists
        history = {
            'train_loss': [],
            'train_acc': [],
            'val_loss': [],
            'val_acc': [],
        }
        
        if history_path.exists():
            with open(history_path, 'r') as f:
                saved_history = json.load(f)
                history.update(saved_history)
            print(f"[+] Loaded training history with {len(history['train_loss'])} previous epochs")
        
        # Get best validation accuracy from history
        best_val_acc = max(history['val_acc']) if history['val_acc'] else 0.0
        
        print(f"[+] Resuming from epoch {last_epoch}")
        print(f"[+] Best validation accuracy so far: {best_val_acc:.4f}")
        print("="*60 + "\n")
        
        return last_epoch, best_val_acc, history
    else:
        print("\n[+] Starting training from scratch")
        return 0, 0.0, {
            'train_loss': [],
            'train_acc': [],
            'val_loss': [],
            'val_acc': [],
        }


def save_training_state(model, optimizer, scheduler, epoch, val_loss, val_acc, 
                        history, architecture, is_best=False):
    """
    Save complete training state including model, optimizer, scheduler, and history.
    
    Args:
        model: PyTorch model
        optimizer: Optimizer
        scheduler: Learning rate scheduler
        epoch: Current epoch
        val_loss: Validation loss
        val_acc: Validation accuracy
        history: Training history dictionary
        architecture: Model architecture name
        is_best: Whether this is the best model so far
    """
    import json
    
    # Save latest checkpoint
    checkpoint = {
        "epoch": epoch,
        "model_state_dict": model.state_dict(),
        "optimizer_state_dict": optimizer.state_dict(),
        "scheduler_state_dict": scheduler.state_dict() if scheduler else None,
        "loss": val_loss,
        "accuracy": val_acc,
    }
    
    latest_path = config.CHECKPOINTS_DIR / f"{architecture}_latest.pth"
    torch.save(checkpoint, latest_path)
    
    # Save training history
    history_path = config.CHECKPOINTS_DIR / f"{architecture}_history.json"
    with open(history_path, 'w') as f:
        json.dump(history, f, indent=4)
    
    # If best model, save separately
    if is_best:
        best_path = config.CHECKPOINTS_DIR / f"{architecture}_best.pth"
        torch.save(checkpoint, best_path)
        print(f"  [+] New best model saved! Val Acc: {val_acc:.4f}")


def train(args):
    """
    Main training function with continuous learning support.
    
    Args:
        args: Command line arguments
    """
    # Set seed
    set_seed(config.RANDOM_SEED)
    
    # Get device
    device = get_device()
    
    # Create data loaders
    train_loader, val_loader, test_loader, num_classes, class_names = get_dataloaders(
        batch_size=args.batch_size,
        num_workers=args.num_workers,
    )
    
    # Create model
    model = get_model(
        architecture=args.architecture,
        num_classes=num_classes,
        pretrained=args.pretrained,
    )
    model = model.to(device)
    
    # Loss function
    criterion = nn.CrossEntropyLoss()
    
    # Optimizer
    if args.optimizer == 'adam':
        optimizer = optim.Adam(
            model.parameters(),
            lr=args.learning_rate,
            weight_decay=args.weight_decay,
        )
    elif args.optimizer == 'sgd':
        optimizer = optim.SGD(
            model.parameters(),
            lr=args.learning_rate,
            momentum=0.9,
            weight_decay=args.weight_decay,
        )
    elif args.optimizer == 'adamw':
        optimizer = optim.AdamW(
            model.parameters(),
            lr=args.learning_rate,
            weight_decay=args.weight_decay,
        )
    
    # Learning rate scheduler
    scheduler = optim.lr_scheduler.StepLR(
        optimizer,
        step_size=args.step_size,
        gamma=args.gamma,
    )
    
    # Load checkpoint if resuming
    start_epoch, best_val_acc, history = load_latest_checkpoint(
        args, model, optimizer, device
    )
    
    # Restore scheduler state if resuming
    if start_epoch > 0:
        for _ in range(start_epoch):
            scheduler.step()
    
    # TensorBoard writer
    if args.use_tensorboard:
        if start_epoch > 0:
            # Use existing log directory
            timestamp = "resumed_" + datetime.now().strftime("%Y%m%d_%H%M%S")
        else:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_dir = config.LOGGING_CONFIG['tensorboard_dir'] / f"{args.architecture}_{timestamp}"
        writer = SummaryWriter(log_dir)
        print(f"\n[+] TensorBoard logging to: {log_dir}")
    
    patience_counter = 0
    
    print("\n" + "="*60)
    print("STARTING TRAINING")
    print("="*60)
    print(f"Training from epoch {start_epoch + 1} to {args.epochs}")
    print(f"Best validation accuracy: {best_val_acc:.4f}")
    print("="*60 + "\n")
    
    start_time = time.time()
    
    # Training loop
    for epoch in range(start_epoch + 1, args.epochs + 1):
        print(f"\nEpoch {epoch}/{args.epochs}")
        print("-" * 60)
        
        # Train
        train_loss, train_acc = train_one_epoch(
            model, train_loader, criterion, optimizer, device, epoch
        )
        
        # Validate
        val_loss, val_acc = validate(
            model, val_loader, criterion, device, epoch
        )
        
        # Update scheduler
        scheduler.step()
        
        # Save history
        history['train_loss'].append(train_loss)
        history['train_acc'].append(train_acc)
        history['val_loss'].append(val_loss)
        history['val_acc'].append(val_acc)
        
        # TensorBoard logging
        if args.use_tensorboard:
            writer.add_scalar('Loss/train', train_loss, epoch)
            writer.add_scalar('Loss/val', val_loss, epoch)
            writer.add_scalar('Accuracy/train', train_acc, epoch)
            writer.add_scalar('Accuracy/val', val_acc, epoch)
            writer.add_scalar('Learning_rate', optimizer.param_groups[0]['lr'], epoch)
        
        # Print epoch summary
        print(f"\nEpoch {epoch} Summary:")
        print(f"  Train Loss: {train_loss:.4f}, Train Acc: {train_acc:.4f}")
        print(f"  Val Loss: {val_loss:.4f}, Val Acc: {val_acc:.4f}")
        print(f"  LR: {optimizer.param_groups[0]['lr']:.6f}")
        
        # Check if best model
        is_best = val_acc > best_val_acc
        if is_best:
            best_val_acc = val_acc
            patience_counter = 0
        else:
            patience_counter += 1
        
        # Save training state (model learns and remembers)
        save_training_state(
            model, optimizer, scheduler, epoch, val_loss, val_acc,
            history, args.architecture, is_best=is_best
        )
        print(f"  [+] Training state saved (epoch {epoch})")
        
        # Early stopping
        if patience_counter >= args.patience:
            print(f"\n[!] Early stopping triggered after {epoch} epochs")
            print(f"  No improvement for {args.patience} epochs")
            break
        
        # Save milestone checkpoint at intervals
        if epoch % args.save_interval == 0:
            milestone_path = config.CHECKPOINTS_DIR / f"{args.architecture}_epoch_{epoch}.pth"
            checkpoint = {
                "epoch": epoch,
                "model_state_dict": model.state_dict(),
                "optimizer_state_dict": optimizer.state_dict(),
                "scheduler_state_dict": scheduler.state_dict(),
                "loss": val_loss,
                "accuracy": val_acc,
            }
            torch.save(checkpoint, milestone_path)
            print(f"  [+] Milestone checkpoint saved: epoch {epoch}")
    
    # Training complete
    training_time = time.time() - start_time
    
    print("\n" + "="*60)
    print("TRAINING COMPLETE")
    print("="*60)
    print(f"Total training time: {format_time(training_time)}")
    print(f"Best validation accuracy: {best_val_acc:.4f}")
    print(f"Total epochs trained: {len(history['train_loss'])}")
    print("="*60 + "\n")
    
    # Close TensorBoard writer
    if args.use_tensorboard:
        writer.close()
    
    # Plot training history
    plot_path = config.FIGURES_DIR / f"{args.architecture}_training_history.png"
    plot_training_history(history, save_path=plot_path)
    
    return model, history


def main():
    parser = argparse.ArgumentParser(description="Train crop disease classification model")
    
    # Model arguments
    parser.add_argument(
        '--architecture',
        type=str,
        default=config.MODEL_CONFIG['architecture'],
        choices=['resnet34', 'resnet50', 'vgg16'],
        help='Model architecture',
    )
    parser.add_argument(
        '--pretrained',
        action='store_true',
        default=config.MODEL_CONFIG['pretrained'],
        help='Use pretrained weights',
    )
    
    # Training arguments
    parser.add_argument(
        '--resume',
        action='store_true',
        default=True,
        help='Resume training from latest checkpoint if available',
    )
    parser.add_argument(
        '--epochs',
        type=int,
        default=config.TRAIN_CONFIG['num_epochs'],
        help='Number of training epochs',
    )
    parser.add_argument(
        '--batch-size',
        type=int,
        default=config.TRAIN_CONFIG['batch_size'],
        help='Batch size',
    )
    parser.add_argument(
        '--learning-rate',
        type=float,
        default=config.TRAIN_CONFIG['learning_rate'],
        help='Learning rate',
    )
    parser.add_argument(
        '--weight-decay',
        type=float,
        default=config.TRAIN_CONFIG['weight_decay'],
        help='Weight decay',
    )
    parser.add_argument(
        '--optimizer',
        type=str,
        default=config.TRAIN_CONFIG['optimizer'],
        choices=['adam', 'sgd', 'adamw'],
        help='Optimizer',
    )
    parser.add_argument(
        '--step-size',
        type=int,
        default=config.TRAIN_CONFIG['step_size'],
        help='Step size for learning rate scheduler',
    )
    parser.add_argument(
        '--gamma',
        type=float,
        default=config.TRAIN_CONFIG['gamma'],
        help='Gamma for learning rate scheduler',
    )
    parser.add_argument(
        '--patience',
        type=int,
        default=config.TRAIN_CONFIG['patience'],
        help='Patience for early stopping',
    )
    parser.add_argument(
        '--num-workers',
        type=int,
        default=config.TRAIN_CONFIG['num_workers'],
        help='Number of data loader workers',
    )
    parser.add_argument(
        '--save-interval',
        type=int,
        default=config.LOGGING_CONFIG['save_interval'],
        help='Save checkpoint every N epochs',
    )
    parser.add_argument(
        '--use-tensorboard',
        action='store_true',
        default=config.LOGGING_CONFIG['use_tensorboard'],
        help='Use TensorBoard logging',
    )
    
    args = parser.parse_args()
    
    # Train model
    train(args)


if __name__ == "__main__":
    main()

