"""
Overfitting Analysis Tool
Analyzes training history to detect overfitting and provides beginner-friendly explanations.
"""

import json
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from pathlib import Path
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


def analyze_overfitting(history):
    """
    Analyze training history for signs of overfitting.
    Returns a dictionary with analysis results.
    """
    train_acc = np.array(history['train_acc']) * 100
    val_acc = np.array(history['val_acc']) * 100
    train_loss = np.array(history['train_loss'])
    val_loss = np.array(history['val_loss'])
    
    # Calculate key metrics
    final_train_acc = train_acc[-1]
    final_val_acc = val_acc[-1]
    final_gap = final_train_acc - final_val_acc
    
    # Check for overfitting indicators
    overfitting_score = 0
    warnings = []
    good_signs = []
    
    # Indicator 1: Accuracy gap
    if abs(final_gap) < 1:
        good_signs.append("Excellent! Training and validation accuracy are nearly identical (<1% gap)")
        overfitting_score += 0
    elif abs(final_gap) < 3:
        good_signs.append("Very good! Small gap between training and validation accuracy (<3%)")
        overfitting_score += 1
    elif abs(final_gap) < 5:
        warnings.append(f"Moderate gap between training and validation accuracy ({final_gap:.2f}%)")
        overfitting_score += 2
    else:
        warnings.append(f"Large gap between training and validation accuracy ({final_gap:.2f}%)")
        overfitting_score += 3
    
    # Indicator 2: Validation performance higher than training (VERY GOOD SIGN!)
    if final_val_acc > final_train_acc:
        good_signs.append(f"Exceptional! Validation accuracy ({final_val_acc:.2f}%) > Training accuracy ({final_train_acc:.2f}%)")
        good_signs.append("This means the model generalizes BETTER than it memorizes!")
    
    # Indicator 3: Loss trend analysis
    loss_gap = train_loss[-1] - val_loss[-1]
    if val_loss[-1] < train_loss[-1]:
        good_signs.append(f"Great! Validation loss ({val_loss[-1]:.4f}) < Training loss ({train_loss[-1]:.4f})")
    elif abs(loss_gap) < 0.1:
        good_signs.append("Good! Training and validation losses are very close")
    else:
        warnings.append(f"Training loss is lower than validation loss by {abs(loss_gap):.4f}")
        overfitting_score += 1
    
    # Indicator 4: Check for divergence in later epochs
    if len(train_acc) >= 10:
        # Compare last 5 epochs with previous 5
        recent_train = train_acc[-5:]
        recent_val = val_acc[-5:]
        earlier_train = train_acc[-10:-5]
        earlier_val = val_acc[-10:-5]
        
        recent_gap = np.mean(recent_train - recent_val)
        earlier_gap = np.mean(earlier_train - earlier_val)
        
        if abs(recent_gap) < abs(earlier_gap):
            good_signs.append("Excellent! The gap is decreasing over time (getting better!)")
        elif abs(recent_gap - earlier_gap) < 1:
            good_signs.append("Good! The gap remains stable (not getting worse)")
        else:
            warnings.append(f"Gap is widening: was {earlier_gap:.2f}%, now {recent_gap:.2f}%")
            overfitting_score += 2
    
    # Indicator 5: Validation loss trend
    if len(val_loss) >= 5:
        recent_val_loss = val_loss[-5:]
        if np.all(np.diff(recent_val_loss) <= 0.01):  # Generally decreasing or stable
            good_signs.append("Good! Validation loss is still improving or stable")
        else:
            # Check if increasing
            if recent_val_loss[-1] > recent_val_loss[0]:
                warnings.append("Validation loss is increasing in recent epochs")
                overfitting_score += 2
    
    # Overall assessment
    if overfitting_score == 0:
        overall = "NO OVERFITTING DETECTED - Excellent!"
        status = "EXCELLENT"
        color = "green"
    elif overfitting_score <= 2:
        overall = "MINIMAL OVERFITTING - Very Good"
        status = "VERY GOOD"
        color = "lightgreen"
    elif overfitting_score <= 4:
        overall = "SLIGHT OVERFITTING - Acceptable"
        status = "ACCEPTABLE"
        color = "yellow"
    else:
        overall = "SIGNIFICANT OVERFITTING DETECTED"
        status = "NEEDS ATTENTION"
        color = "orange"
    
    return {
        'final_train_acc': final_train_acc,
        'final_val_acc': final_val_acc,
        'final_gap': final_gap,
        'train_loss': train_loss,
        'val_loss': val_loss,
        'train_acc': train_acc,
        'val_acc': val_acc,
        'overfitting_score': overfitting_score,
        'warnings': warnings,
        'good_signs': good_signs,
        'overall': overall,
        'status': status,
        'color': color
    }


def create_overfitting_visualization(history, analysis, save_dir):
    """Create comprehensive overfitting analysis visualizations."""
    save_dir = Path(save_dir)
    save_dir.mkdir(parents=True, exist_ok=True)
    
    epochs = list(range(1, len(history['train_loss']) + 1))
    
    # Set style
    plt.style.use('seaborn-v0_8-whitegrid')
    
    # ============================================================
    # Main Overfitting Analysis Dashboard
    # ============================================================
    fig = plt.figure(figsize=(18, 12))
    gs = fig.add_gridspec(3, 3, hspace=0.3, wspace=0.3)
    
    fig.suptitle('Overfitting Analysis Dashboard\n(Beginner-Friendly Evaluation)', 
                 fontsize=18, fontweight='bold')
    
    # Plot 1: Accuracy Comparison
    ax1 = fig.add_subplot(gs[0, :2])
    ax1.plot(epochs, analysis['train_acc'], 'o-', linewidth=2, 
             label='Training Accuracy', color='#e74c3c', markersize=6)
    ax1.plot(epochs, analysis['val_acc'], 's-', linewidth=2, 
             label='Validation Accuracy', color='#3498db', markersize=6)
    ax1.fill_between(epochs, analysis['train_acc'], analysis['val_acc'], 
                     alpha=0.2, color='gray')
    ax1.set_xlabel('Epoch', fontsize=12, fontweight='bold')
    ax1.set_ylabel('Accuracy (%)', fontsize=12, fontweight='bold')
    ax1.set_title('Training vs Validation Accuracy\n(Gap indicates overfitting)', 
                  fontsize=13, fontweight='bold')
    ax1.legend(fontsize=11, loc='lower right')
    ax1.grid(True, alpha=0.3)
    ax1.set_ylim(0, 105)
    
    # Add annotation for the gap
    final_epoch = len(epochs)
    gap = analysis['final_gap']
    mid_point = (analysis['train_acc'][-1] + analysis['val_acc'][-1]) / 2
    
    if abs(gap) > 0.5:
        ax1.annotate(f'Gap: {gap:.2f}%',
                    xy=(final_epoch, mid_point),
                    xytext=(final_epoch - len(epochs)*0.15, mid_point),
                    arrowprops=dict(arrowstyle='->', color='red', lw=2),
                    fontsize=11, fontweight='bold', color='red',
                    bbox=dict(boxstyle='round', facecolor='yellow', alpha=0.7))
    
    # Plot 2: Loss Comparison
    ax2 = fig.add_subplot(gs[1, :2])
    ax2.plot(epochs, analysis['train_loss'], 'o-', linewidth=2, 
             label='Training Loss', color='#e74c3c', markersize=6)
    ax2.plot(epochs, analysis['val_loss'], 's-', linewidth=2, 
             label='Validation Loss', color='#3498db', markersize=6)
    ax2.fill_between(epochs, analysis['train_loss'], analysis['val_loss'], 
                     alpha=0.2, color='gray')
    ax2.set_xlabel('Epoch', fontsize=12, fontweight='bold')
    ax2.set_ylabel('Loss (Error)', fontsize=12, fontweight='bold')
    ax2.set_title('Training vs Validation Loss\n(Divergence indicates overfitting)', 
                  fontsize=13, fontweight='bold')
    ax2.legend(fontsize=11, loc='upper right')
    ax2.grid(True, alpha=0.3)
    
    # Plot 3: Gap Progression
    ax3 = fig.add_subplot(gs[2, :2])
    accuracy_gap = analysis['train_acc'] - analysis['val_acc']
    colors = ['green' if abs(g) < 3 else 'orange' if abs(g) < 5 else 'red' for g in accuracy_gap]
    ax3.bar(epochs, accuracy_gap, color=colors, alpha=0.7, edgecolor='black')
    ax3.axhline(y=0, color='black', linestyle='--', linewidth=2)
    ax3.axhline(y=3, color='orange', linestyle=':', linewidth=1, alpha=0.5)
    ax3.axhline(y=-3, color='orange', linestyle=':', linewidth=1, alpha=0.5)
    ax3.axhline(y=5, color='red', linestyle=':', linewidth=1, alpha=0.5)
    ax3.axhline(y=-5, color='red', linestyle=':', linewidth=1, alpha=0.5)
    ax3.set_xlabel('Epoch', fontsize=12, fontweight='bold')
    ax3.set_ylabel('Accuracy Gap (%)', fontsize=12, fontweight='bold')
    ax3.set_title('Training-Validation Gap Over Time\n(Closer to 0 is better)', 
                  fontsize=13, fontweight='bold')
    ax3.grid(True, alpha=0.3, axis='y')
    
    # Add reference lines labels
    ax3.text(len(epochs)*0.02, 3.3, 'Moderate', fontsize=9, color='orange')
    ax3.text(len(epochs)*0.02, 5.3, 'High', fontsize=9, color='red')
    
    # Plot 4: Overall Status
    ax4 = fig.add_subplot(gs[0, 2])
    ax4.axis('off')
    
    status_text = f"""
OVERFITTING STATUS

{analysis['overall']}

Score: {analysis['overfitting_score']}/10
(Lower is better)

Final Metrics:
• Train Acc: {analysis['final_train_acc']:.2f}%
• Val Acc: {analysis['final_val_acc']:.2f}%
• Gap: {analysis['final_gap']:.2f}%

Interpretation:
• Gap < 1%: Excellent
• Gap < 3%: Very Good
• Gap < 5%: Acceptable
• Gap > 5%: Concern
    """
    
    ax4.text(0.1, 0.95, status_text, transform=ax4.transAxes,
            fontsize=11, verticalalignment='top', family='monospace',
            bbox=dict(boxstyle='round', facecolor=analysis['color'], 
                     alpha=0.8, edgecolor='black', linewidth=2))
    
    # Plot 5: Good Signs
    ax5 = fig.add_subplot(gs[1, 2])
    ax5.axis('off')
    
    good_text = "POSITIVE SIGNS:\n" + "="*30 + "\n\n"
    if analysis['good_signs']:
        for i, sign in enumerate(analysis['good_signs'][:5], 1):
            good_text += f"{i}. {sign}\n\n"
    else:
        good_text += "None detected"
    
    ax5.text(0.05, 0.95, good_text, transform=ax5.transAxes,
            fontsize=9, verticalalignment='top', family='monospace',
            bbox=dict(boxstyle='round', facecolor='lightgreen', alpha=0.6),
            wrap=True)
    
    # Plot 6: Warnings
    ax6 = fig.add_subplot(gs[2, 2])
    ax6.axis('off')
    
    warn_text = "WARNINGS:\n" + "="*30 + "\n\n"
    if analysis['warnings']:
        for i, warning in enumerate(analysis['warnings'][:5], 1):
            warn_text += f"{i}. {warning}\n\n"
    else:
        warn_text += "None! Model looks great!"
    
    ax6.text(0.05, 0.95, warn_text, transform=ax6.transAxes,
            fontsize=9, verticalalignment='top', family='monospace',
            bbox=dict(boxstyle='round', facecolor='lightyellow', alpha=0.8),
            wrap=True)
    
    plt.tight_layout()
    dashboard_path = save_dir / 'overfitting_analysis_dashboard.png'
    plt.savefig(dashboard_path, dpi=300, bbox_inches='tight')
    print(f"[+] Overfitting dashboard saved: {dashboard_path}")
    plt.close()


def generate_overfitting_report(analysis, architecture='resnet34'):
    """Generate a beginner-friendly overfitting analysis report."""
    report_path = config.RESULTS_DIR / 'OVERFITTING_ANALYSIS_REPORT.txt'
    
    report = f"""
{'='*80}
              OVERFITTING ANALYSIS REPORT
              (Beginner-Friendly Explanation)
{'='*80}

WHAT IS OVERFITTING?
--------------------
Overfitting happens when a model "memorizes" the training data instead of 
learning general patterns. It's like a student who memorizes answers without 
understanding concepts - they do great on practice tests but fail on real exams.

A good model should perform similarly on BOTH:
  • Training data (what it studied)
  • Validation data (new, unseen data - the "real test")


YOUR MODEL'S OVERFITTING STATUS
--------------------------------
{analysis['overall']}

Overfitting Score: {analysis['overfitting_score']}/10
(0 = No overfitting, 10 = Severe overfitting)


KEY METRICS
-----------
Training Accuracy:    {analysis['final_train_acc']:.2f}%
Validation Accuracy:  {analysis['final_val_acc']:.2f}%
Gap:                  {analysis['final_gap']:.2f}%

Training Loss:        {analysis['train_loss'][-1]:.4f}
Validation Loss:      {analysis['val_loss'][-1]:.4f}


WHAT DOES THIS MEAN?
--------------------
"""
    
    # Add interpretation
    if analysis['final_val_acc'] > analysis['final_train_acc']:
        report += f"""
🌟 EXCEPTIONAL RESULT! 🌟

Your validation accuracy ({analysis['final_val_acc']:.2f}%) is HIGHER than 
training accuracy ({analysis['final_train_acc']:.2f}%)!

This is the BEST possible scenario. It means:
  • The model is NOT memorizing training data
  • It generalizes BETTER to new data than to training data
  • It has learned true underlying patterns, not noise
  • The model is READY for real-world use!

This is actually BETTER than no overfitting - it's "underfitting" in the best 
way possible, meaning the model is conservative and reliable.
"""
    elif abs(analysis['final_gap']) < 1:
        report += f"""
✓ EXCELLENT - NO OVERFITTING DETECTED

Your model performs almost identically on training and validation data
(only {abs(analysis['final_gap']):.2f}% difference).

This means:
  • The model has learned general patterns, not specific examples
  • It will perform well on new, real-world data
  • The accuracy you see is reliable and trustworthy
"""
    elif abs(analysis['final_gap']) < 3:
        report += f"""
✓ VERY GOOD - MINIMAL OVERFITTING

The gap is small ({abs(analysis['final_gap']):.2f}%), which is normal and acceptable.

This means:
  • The model is slightly better at training data (expected)
  • It should still perform very well on new data
  • The difference is within acceptable limits
"""
    elif abs(analysis['final_gap']) < 5:
        report += f"""
⚠ ACCEPTABLE - SLIGHT OVERFITTING

The gap is moderate ({abs(analysis['final_gap']):.2f}%).

This means:
  • The model knows training data better than new data
  • Performance on new data should still be good
  • Consider more training or regularization if accuracy drops
"""
    else:
        report += f"""
⚠ ATTENTION NEEDED - SIGNIFICANT OVERFITTING

The gap is large ({abs(analysis['final_gap']):.2f}%).

This means:
  • The model may be memorizing training data
  • Performance on new data might be significantly lower
  • Consider: more data, data augmentation, or regularization
"""
    
    report += f"""


DETAILED ANALYSIS
-----------------

✓ POSITIVE SIGNS:
{'='*60}
"""
    
    if analysis['good_signs']:
        for i, sign in enumerate(analysis['good_signs'], 1):
            report += f"{i}. {sign}\n"
    else:
        report += "No particularly positive signs detected.\n"
    
    report += f"""

⚠ WARNING SIGNS:
{'='*60}
"""
    
    if analysis['warnings']:
        for i, warning in enumerate(analysis['warnings'], 1):
            report += f"{i}. {warning}\n"
    else:
        report += "No warning signs detected! Your model looks excellent!\n"
    
    report += f"""


HOW TO INTERPRET THESE RESULTS
-------------------------------

1. THE GAP:
   • < 1%: Perfect! Model generalizes excellently
   • 1-3%: Great! Normal and expected behavior
   • 3-5%: Good! Slight preference for training data
   • > 5%: Concerning - model may be overfitting

2. VALIDATION ACCURACY:
   • If validation > training: EXCEPTIONAL (like your model!)
   • If validation ≈ training: EXCELLENT
   • If validation < training: Check the gap size

3. LOSS BEHAVIOR:
   • Both decreasing: Learning continues
   • Both stable: Model has converged
   • Training↓ but Validation↑: Overfitting warning


WHAT TO DO NEXT
----------------

"""
    
    if analysis['overfitting_score'] <= 2:
        report += """✓ Your model looks GREAT! No action needed.

Recommendations:
  • Use this model confidently for predictions
  • Deploy it to production if needed
  • Document its excellent performance
  • Consider this as your final model
"""
    elif analysis['overfitting_score'] <= 4:
        report += """✓ Your model is GOOD but could be slightly better.

Optional improvements:
  • Try training with more data augmentation
  • Add dropout layers (if not already present)
  • Use early stopping (may already be active)
  • Collect more diverse training data
"""
    else:
        report += """⚠ Consider these improvements:

Recommended actions:
  1. Add more training data if possible
  2. Increase data augmentation strength
  3. Add dropout layers or increase dropout rate
  4. Use L2 regularization (weight decay)
  5. Try early stopping
  6. Reduce model complexity if needed
"""
    
    report += f"""


TECHNICAL DETAILS
-----------------
Model Architecture: {architecture}
Total Epochs: {len(analysis['train_acc'])}
Best Validation Accuracy: {max(analysis['val_acc']):.2f}%
Best Validation Loss: {min(analysis['val_loss']):.4f}


CONCLUSION
----------
"""
    
    if analysis['final_val_acc'] > analysis['final_train_acc']:
        report += f"""
Your model shows EXCEPTIONAL generalization! The validation accuracy being 
HIGHER than training accuracy is the gold standard. This model is ready for 
real-world deployment with confidence.

Final Verdict: ⭐⭐⭐⭐⭐ (5/5 stars) - Outstanding!
"""
    elif analysis['overfitting_score'] <= 2:
        report += f"""
Your model shows excellent generalization with minimal to no overfitting.
The performance you see is reliable and should transfer well to real-world data.

Final Verdict: ⭐⭐⭐⭐ (4/5 stars) - Excellent!
"""
    else:
        report += f"""
Your model shows some signs of overfitting. While the current performance is 
good, consider the recommendations above to improve generalization.

Final Verdict: ⭐⭐⭐ (3/5 stars) - Good, with room for improvement
"""
    
    report += f"""

Report Generated: {config.datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
{'='*80}
"""
    
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write(report)
    
    print(f"\n[+] Overfitting analysis report saved: {report_path}")
    return report_path


def main():
    """Run complete overfitting analysis."""
    print("\n" + "="*80)
    print("OVERFITTING ANALYSIS TOOL")
    print("="*80 + "\n")
    
    # Load training history
    print("[*] Loading training history...")
    history = load_training_history()
    
    if history is None:
        print("\n[!] No training history found. Train the model first!")
        return
    
    print(f"[+] Found {len(history['train_loss'])} training epochs\n")
    
    # Analyze overfitting
    print("[*] Analyzing for overfitting patterns...")
    analysis = analyze_overfitting(history)
    
    # Create visualizations
    print("\n[*] Creating overfitting analysis dashboard...")
    create_overfitting_visualization(history, analysis, config.FIGURES_DIR)
    
    # Generate text report
    print("\n[*] Generating detailed overfitting report...")
    report_path = generate_overfitting_report(analysis)
    
    # Print summary
    print("\n" + "="*80)
    print("ANALYSIS COMPLETE!")
    print("="*80)
    print(f"\nOverfitting Status: {analysis['status']}")
    print(f"Overall Assessment: {analysis['overall']}")
    print(f"\nFiles created:")
    print(f"  1. Dashboard: {config.FIGURES_DIR / 'overfitting_analysis_dashboard.png'}")
    print(f"  2. Report: {report_path}")
    print("\nKey Finding:")
    if analysis['final_val_acc'] > analysis['final_train_acc']:
        print("  ⭐ EXCEPTIONAL! Validation accuracy > Training accuracy")
        print("     Your model generalizes better than it memorizes!")
    elif analysis['overfitting_score'] <= 2:
        print("  ✓ EXCELLENT! No significant overfitting detected")
    else:
        print(f"  ⚠ Overfitting score: {analysis['overfitting_score']}/10")
        print("     Review the report for recommendations")
    print("="*80 + "\n")


if __name__ == "__main__":
    import datetime
    config.datetime = datetime
    main()

