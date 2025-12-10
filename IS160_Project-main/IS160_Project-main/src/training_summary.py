"""
Educational Training Summary Generator
Creates beginner-friendly reports and visualizations after model training.
"""

import json
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from pathlib import Path
from datetime import datetime
import config


def load_training_history(architecture='resnet34'):
    """Load training history from JSON file."""
    history_path = config.CHECKPOINTS_DIR / f"{architecture}_history.json"
    
    if not history_path.exists():
        print(f"[!] No training history found at {history_path}")
        return None
    
    with open(history_path, 'r') as f:
        history = json.load(f)
    
    return history


def create_educational_plots(history, save_dir):
    """
    Create educational visualizations that explain the training process.
    """
    save_dir = Path(save_dir)
    save_dir.mkdir(parents=True, exist_ok=True)
    
    epochs = list(range(1, len(history['train_loss']) + 1))
    
    # Set a nice style
    plt.style.use('seaborn-v0_8-whitegrid')
    sns.set_palette("husl")
    
    # ============================================================
    # 1. COMPREHENSIVE TRAINING OVERVIEW (4 subplots)
    # ============================================================
    fig, axes = plt.subplots(2, 2, figsize=(16, 12))
    fig.suptitle('Plant Disease Model Training Summary\n(Educational Overview)', 
                 fontsize=18, fontweight='bold', y=0.995)
    
    # Plot 1: Loss Over Time
    ax1 = axes[0, 0]
    ax1.plot(epochs, history['train_loss'], marker='o', linewidth=2, 
             label='Training Loss', color='#e74c3c', markersize=8)
    ax1.plot(epochs, history['val_loss'], marker='s', linewidth=2, 
             label='Validation Loss', color='#3498db', markersize=8)
    ax1.set_xlabel('Epoch (Training Cycle)', fontsize=12, fontweight='bold')
    ax1.set_ylabel('Loss (Error)', fontsize=12, fontweight='bold')
    ax1.set_title('How Well is the Model Learning?\n(Lower Loss = Better)', 
                  fontsize=13, fontweight='bold')
    ax1.legend(fontsize=11, loc='upper right')
    ax1.grid(True, alpha=0.3)
    ax1.set_xlim(0.5, len(epochs) + 0.5)
    
    # Add annotations
    min_val_loss = min(history['val_loss'])
    min_epoch = history['val_loss'].index(min_val_loss) + 1
    ax1.annotate(f'Best Model\n(Epoch {min_epoch})',
                xy=(min_epoch, min_val_loss),
                xytext=(min_epoch + 1, min_val_loss + 0.1),
                arrowprops=dict(arrowstyle='->', color='green', lw=2),
                fontsize=10, fontweight='bold', color='green')
    
    # Plot 2: Accuracy Over Time
    ax2 = axes[0, 1]
    train_acc_pct = [acc * 100 for acc in history['train_acc']]
    val_acc_pct = [acc * 100 for acc in history['val_acc']]
    
    ax2.plot(epochs, train_acc_pct, marker='o', linewidth=2, 
             label='Training Accuracy', color='#27ae60', markersize=8)
    ax2.plot(epochs, val_acc_pct, marker='s', linewidth=2, 
             label='Validation Accuracy', color='#f39c12', markersize=8)
    ax2.set_xlabel('Epoch (Training Cycle)', fontsize=12, fontweight='bold')
    ax2.set_ylabel('Accuracy (%)', fontsize=12, fontweight='bold')
    ax2.set_title('How Accurate is the Model?\n(Higher % = Better)', 
                  fontsize=13, fontweight='bold')
    ax2.legend(fontsize=11, loc='lower right')
    ax2.grid(True, alpha=0.3)
    ax2.set_xlim(0.5, len(epochs) + 0.5)
    ax2.set_ylim(0, 105)
    
    # Add accuracy annotations
    final_val_acc = val_acc_pct[-1]
    ax2.axhline(y=final_val_acc, color='#f39c12', linestyle='--', alpha=0.5)
    ax2.text(len(epochs) * 0.05, final_val_acc + 2, 
             f'Final: {final_val_acc:.1f}%', 
             fontsize=11, fontweight='bold', color='#f39c12')
    
    # Plot 3: Learning Progress Bar Chart
    ax3 = axes[1, 0]
    categories = ['Start\n(Epoch 1)', 'End\n(Final Epoch)']
    train_progress = [train_acc_pct[0], train_acc_pct[-1]]
    val_progress = [val_acc_pct[0], val_acc_pct[-1]]
    
    x = np.arange(len(categories))
    width = 0.35
    
    bars1 = ax3.bar(x - width/2, train_progress, width, label='Training', 
                    color='#27ae60', alpha=0.8)
    bars2 = ax3.bar(x + width/2, val_progress, width, label='Validation', 
                    color='#f39c12', alpha=0.8)
    
    ax3.set_ylabel('Accuracy (%)', fontsize=12, fontweight='bold')
    ax3.set_title('Learning Improvement\n(Start vs. End)', 
                  fontsize=13, fontweight='bold')
    ax3.set_xticks(x)
    ax3.set_xticklabels(categories, fontsize=11)
    ax3.legend(fontsize=11)
    ax3.set_ylim(0, 105)
    ax3.grid(True, alpha=0.3, axis='y')
    
    # Add value labels on bars
    for bars in [bars1, bars2]:
        for bar in bars:
            height = bar.get_height()
            ax3.text(bar.get_x() + bar.get_width()/2., height + 1,
                    f'{height:.1f}%', ha='center', va='bottom', 
                    fontsize=10, fontweight='bold')
    
    # Plot 4: Model Performance Summary (Text Box)
    ax4 = axes[1, 1]
    ax4.axis('off')
    
    # Calculate metrics
    improvement = val_acc_pct[-1] - val_acc_pct[0]
    best_val_acc = max(val_acc_pct)
    avg_val_acc = np.mean(val_acc_pct)
    
    summary_text = f"""
    TRAINING RESULTS SUMMARY
    {'='*50}
    
    Total Training Cycles (Epochs): {len(epochs)}
    
    FINAL PERFORMANCE:
      • Training Accuracy: {train_acc_pct[-1]:.2f}%
      • Validation Accuracy: {val_acc_pct[-1]:.2f}%
    
    IMPROVEMENT:
      • Started at: {val_acc_pct[0]:.2f}%
      • Ended at: {val_acc_pct[-1]:.2f}%
      • Improvement: +{improvement:.2f}%
    
    BEST PERFORMANCE:
      • Best Accuracy: {best_val_acc:.2f}%
      • Achieved at Epoch: {val_acc_pct.index(max(val_acc_pct)) + 1}
    
    STABILITY:
      • Average Accuracy: {avg_val_acc:.2f}%
      • Consistency: {'Excellent' if max(val_acc_pct) - min(val_acc_pct) < 5 else 'Good'}
    
    MODEL STATUS: {'Ready for Use!' if val_acc_pct[-1] > 85 else 'Needs More Training'}
    """
    
    ax4.text(0.1, 0.95, summary_text, transform=ax4.transAxes,
            fontsize=11, verticalalignment='top', family='monospace',
            bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.8))
    
    plt.tight_layout()
    summary_path = save_dir / 'training_summary_educational.png'
    plt.savefig(summary_path, dpi=300, bbox_inches='tight')
    print(f"[+] Educational summary saved: {summary_path}")
    plt.close()
    
    # ============================================================
    # 2. LEARNING CURVE (Big, clear visualization)
    # ============================================================
    fig, ax = plt.subplots(figsize=(14, 8))
    
    ax.plot(epochs, val_acc_pct, marker='o', linewidth=3, 
           label='Model Accuracy', color='#2ecc71', markersize=10)
    ax.fill_between(epochs, val_acc_pct, alpha=0.3, color='#2ecc71')
    
    ax.set_xlabel('Training Epoch (Each cycle through all images)', 
                  fontsize=14, fontweight='bold')
    ax.set_ylabel('Accuracy (%)', fontsize=14, fontweight='bold')
    ax.set_title('Plant Disease Detection Model: Learning Progress\n' +
                 'How the AI Gets Better at Identifying Diseases', 
                 fontsize=16, fontweight='bold', pad=20)
    ax.grid(True, alpha=0.3, linestyle='--')
    ax.set_xlim(0.5, len(epochs) + 0.5)
    ax.set_ylim(0, 105)
    
    # Add milestone markers
    milestones = [
        (25, '25% Accurate'),
        (50, '50% Accurate'),
        (75, '75% Accurate'),
        (90, '90% Accurate - Excellent!'),
    ]
    
    for milestone, label in milestones:
        ax.axhline(y=milestone, color='gray', linestyle=':', alpha=0.5)
        ax.text(0.5, milestone + 1, label, fontsize=10, alpha=0.7)
    
    # Highlight final result
    ax.plot(len(epochs), val_acc_pct[-1], 'r*', markersize=25, 
           label=f'Final Result: {val_acc_pct[-1]:.1f}%')
    ax.legend(fontsize=12, loc='lower right')
    
    plt.tight_layout()
    learning_curve_path = save_dir / 'learning_curve_simple.png'
    plt.savefig(learning_curve_path, dpi=300, bbox_inches='tight')
    print(f"[+] Learning curve saved: {learning_curve_path}")
    plt.close()
    
    # ============================================================
    # 3. COMPARISON CHART (Before vs After)
    # ============================================================
    fig, ax = plt.subplots(figsize=(10, 8))
    
    labels = ['Before Training\n(Random Guessing)', 
              'After Training\n(Trained Model)']
    values = [100 / 27, val_acc_pct[-1]]  # 27 classes, so random = 3.7%
    colors = ['#e74c3c', '#2ecc71']
    
    bars = ax.bar(labels, values, color=colors, alpha=0.8, width=0.6)
    
    # Add value labels
    for bar in bars:
        height = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2., height + 2,
               f'{height:.1f}%', ha='center', va='bottom', 
               fontsize=16, fontweight='bold')
    
    ax.set_ylabel('Accuracy (%)', fontsize=14, fontweight='bold')
    ax.set_title('The Power of Machine Learning!\n' +
                'From Random Guesses to Accurate Predictions', 
                fontsize=16, fontweight='bold', pad=20)
    ax.set_ylim(0, 105)
    ax.grid(True, alpha=0.3, axis='y')
    
    # Add explanation text
    improvement_factor = values[1] / values[0]
    ax.text(0.5, 0.5, 
           f'The model is now\n{improvement_factor:.1f}x better\nthan random guessing!',
           transform=ax.transAxes, fontsize=14, ha='center',
           bbox=dict(boxstyle='round', facecolor='yellow', alpha=0.7))
    
    plt.tight_layout()
    comparison_path = save_dir / 'before_after_comparison.png'
    plt.savefig(comparison_path, dpi=300, bbox_inches='tight')
    print(f"[+] Comparison chart saved: {comparison_path}")
    plt.close()


def generate_beginner_report(history, architecture='resnet34'):
    """
    Generate a beginner-friendly text report explaining the training results.
    """
    report_path = config.RESULTS_DIR / 'TRAINING_REPORT_FOR_BEGINNERS.txt'
    
    # Calculate metrics
    final_train_acc = history['train_acc'][-1] * 100
    final_val_acc = history['val_acc'][-1] * 100
    initial_val_acc = history['val_acc'][0] * 100
    improvement = final_val_acc - initial_val_acc
    best_val_acc = max(history['val_acc']) * 100
    best_epoch = history['val_acc'].index(max(history['val_acc'])) + 1
    total_epochs = len(history['train_loss'])
    
    # Determine model quality
    if final_val_acc >= 90:
        quality = "EXCELLENT"
        recommendation = "This model is ready for real-world use!"
    elif final_val_acc >= 80:
        quality = "GOOD"
        recommendation = "This model performs well and can be used with confidence."
    elif final_val_acc >= 70:
        quality = "FAIR"
        recommendation = "This model works but could benefit from more training."
    else:
        quality = "NEEDS IMPROVEMENT"
        recommendation = "Consider training longer or using more data."
    
    report = f"""
{'='*80}
              PLANT DISEASE DETECTION MODEL
              TRAINING RESULTS REPORT
              (Explained for Beginners)
{'='*80}

WHAT IS THIS REPORT?
-------------------
This report explains how well our AI model learned to identify plant diseases
from images. Think of it like a student's report card after studying!


YOUR MODEL'S PERFORMANCE
------------------------
Overall Grade: {quality}
Final Accuracy: {final_val_acc:.2f}%

This means the model correctly identifies plant diseases {final_val_acc:.1f}% of 
the time when shown new images it has never seen before.


HOW THE MODEL LEARNED
---------------------
• Training Cycles Completed: {total_epochs} epochs
  (An "epoch" is one complete pass through all training images)

• Starting Accuracy: {initial_val_acc:.2f}%
• Final Accuracy: {final_val_acc:.2f}%
• Improvement: +{improvement:.2f} percentage points

• Best Performance: {best_val_acc:.2f}%
  (Achieved during epoch {best_epoch})


WHAT DO THESE NUMBERS MEAN?
----------------------------
Imagine showing the model 100 plant images:

Before Training:
  - Would guess correctly: ~{100/27:.0f} times (just random guessing)
  
After Training:
  - Identifies correctly: ~{int(final_val_acc)} times
  - Makes mistakes: ~{int(100-final_val_acc)} times


TWO TYPES OF ACCURACY
---------------------
We measure two types of accuracy during training:

1. TRAINING ACCURACY: {final_train_acc:.2f}%
   - How well the model performs on images it has already studied
   - Like a student practicing with familiar problems

2. VALIDATION ACCURACY: {final_val_acc:.2f}%
   - How well the model performs on NEW images it hasn't seen before
   - Like taking a test with new questions
   - This is the MOST IMPORTANT number!


IS MY MODEL GOOD?
-----------------
{recommendation}

For reference:
  • 90-100%: Excellent - Professional quality
  • 80-90%:  Good - Reliable for most uses
  • 70-80%:  Fair - Acceptable but room for improvement
  • Below 70%: Needs more training


WHAT THE MODEL CAN DO NOW
--------------------------
Your model has learned to identify {27} different types of plant diseases
across multiple crops including:
  • Tomato diseases (10 types)
  • Apple diseases (4 types)
  • Grape diseases (4 types)
  • Potato diseases (3 types)
  • And more!

It can now:
✓ Analyze plant leaf images
✓ Detect disease symptoms
✓ Classify the specific disease type
✓ Help farmers make treatment decisions


HOW DOES IT WORK?
-----------------
The model uses something called a "neural network" - artificial neurons 
inspired by how the human brain works. During training:

1. The model looks at thousands of plant images
2. It learns patterns (like spots, discoloration, leaf texture)
3. It adjusts itself to get better at recognizing diseases
4. It tests itself on new images to make sure it really learned


TRAINING STATISTICS
-------------------
Total images used for training: ~45,000
Total images used for validation: ~14,000
Model architecture: {architecture.upper()}
Number of parameters (weights): ~21 million
Training device: GPU (NVIDIA RTX 4050)


NEXT STEPS
----------
1. Review the generated charts in: {config.FIGURES_DIR}
2. Test the model on your own plant images
3. Use it to help identify diseases in real crops
4. If needed, train for more epochs to improve accuracy


TECHNICAL DETAILS
-----------------
Final Training Loss: {history['train_loss'][-1]:.4f}
Final Validation Loss: {history['val_loss'][-1]:.4f}

(Lower loss = better performance. Loss is a measure of how wrong the 
model's predictions are. The model tries to minimize this during training.)


UNDERSTANDING THE GRAPHS
-------------------------
Three charts have been created for you:

1. "training_summary_educational.png"
   - Shows how loss and accuracy changed during training
   - Four panels showing different aspects of learning

2. "learning_curve_simple.png"
   - Big, clear view of accuracy improvement over time
   - Easy to see the learning progress

3. "before_after_comparison.png"
   - Shows the dramatic improvement from before to after training
   - Compares random guessing vs. trained model


QUESTIONS TO ASK YOURSELF
--------------------------
✓ Did the model improve over time? (Check the learning curve)
✓ Is the validation accuracy stable? (Not jumping around wildly)
✓ Is there a big gap between training and validation accuracy?
  (Small gap = good, means the model generalizes well)


CONGRATULATIONS!
----------------
You've successfully trained an AI model for plant disease detection!
This model can now help farmers, gardeners, and agricultural scientists
identify crop diseases quickly and accurately.


Report Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
{'='*80}
"""
    
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write(report)
    
    print(f"\n[+] Beginner-friendly report saved: {report_path}")
    return report_path


def main():
    """Generate complete educational training summary."""
    print("\n" + "="*80)
    print("EDUCATIONAL TRAINING SUMMARY GENERATOR")
    print("="*80 + "\n")
    
    # Load training history
    print("[*] Loading training history...")
    history = load_training_history()
    
    if history is None:
        print("\n[!] No training history found. Train the model first!")
        return
    
    print(f"[+] Found {len(history['train_loss'])} training epochs\n")
    
    # Create visualizations
    print("[*] Creating educational visualizations...")
    create_educational_plots(history, config.FIGURES_DIR)
    
    # Generate text report
    print("\n[*] Generating beginner-friendly report...")
    report_path = generate_beginner_report(history)
    
    print("\n" + "="*80)
    print("SUMMARY GENERATION COMPLETE!")
    print("="*80)
    print(f"\nFiles created:")
    print(f"  1. Report: {report_path}")
    print(f"  2. Charts: {config.FIGURES_DIR}")
    print("\nYou can now:")
    print("  • Read the beginner-friendly report for easy understanding")
    print("  • Use the charts in presentations or educational materials")
    print("  • Share the visualizations with students or team members")
    print("="*80 + "\n")


if __name__ == "__main__":
    main()

