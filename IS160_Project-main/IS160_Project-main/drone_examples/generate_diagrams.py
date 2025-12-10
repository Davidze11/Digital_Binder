"""
Generate a visual architecture diagram for the drone processing system
"""

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import numpy as np


def create_architecture_diagram():
    """Create a visual diagram showing the drone processing workflow"""
    
    fig, ax = plt.subplots(1, 1, figsize=(16, 12))
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 12)
    ax.axis('off')
    
    # Title
    ax.text(5, 11.5, 'Agricultural Drone Processing System Architecture', 
            ha='center', va='top', fontsize=20, fontweight='bold')
    
    # Colors
    input_color = '#E3F2FD'
    process_color = '#FFF9C4'
    output_color = '#C8E6C9'
    module_color = '#FFE0B2'
    
    # Box styling
    def draw_box(x, y, width, height, text, color, fontsize=10):
        box = FancyBboxPatch((x, y), width, height, 
                            boxstyle="round,pad=0.1", 
                            facecolor=color, edgecolor='black', linewidth=2)
        ax.add_patch(box)
        ax.text(x + width/2, y + height/2, text, 
               ha='center', va='center', fontsize=fontsize, fontweight='bold')
    
    # Arrow styling
    def draw_arrow(x1, y1, x2, y2, label=''):
        arrow = FancyArrowPatch((x1, y1), (x2, y2),
                              arrowstyle='->', mutation_scale=30, 
                              linewidth=2, color='black')
        ax.add_patch(arrow)
        if label:
            mid_x, mid_y = (x1 + x2) / 2, (y1 + y2) / 2
            ax.text(mid_x + 0.2, mid_y, label, fontsize=8, style='italic')
    
    # ========== INPUT LAYER ==========
    ax.text(5, 10.5, '1. INPUT', ha='center', fontsize=14, fontweight='bold', color='#1976D2')
    draw_box(1, 9.5, 2, 0.8, 'Drone\nImage', input_color, 10)
    draw_box(4, 9.5, 2, 0.8, 'Configuration\nParameters', input_color, 9)
    draw_box(7, 9.5, 2, 0.8, 'Trained\nModel', input_color, 10)
    
    # ========== PROCESSING LAYER ==========
    ax.text(5, 8.8, '2. PROCESSING', ha='center', fontsize=14, fontweight='bold', color='#F57C00')
    
    # Main processor box
    draw_box(2, 6.5, 6, 2, '', module_color)
    ax.text(5, 8.2, 'DroneImageProcessor', ha='center', fontsize=12, fontweight='bold')
    
    # Sub-components
    draw_box(2.2, 7.3, 1.8, 0.5, 'Image Loading', '#FFE0B2', 8)
    draw_box(4.1, 7.3, 1.8, 0.5, 'Tile Splitting', '#FFE0B2', 8)
    draw_box(6.0, 7.3, 1.8, 0.5, 'AI Inference', '#FFE0B2', 8)
    draw_box(3.0, 6.7, 2.0, 0.5, 'Confidence Filtering', '#FFE0B2', 8)
    draw_box(5.2, 6.7, 2.0, 0.5, 'Analysis Generation', '#FFE0B2', 8)
    
    # ========== INTERMEDIATE OUTPUT ==========
    ax.text(5, 5.9, '3. ANALYSIS DATA', ha='center', fontsize=14, fontweight='bold', color='#388E3C')
    draw_box(3.5, 5.0, 3, 0.7, 'Comprehensive JSON Analysis', process_color, 10)
    ax.text(5, 4.7, '(Detections, Statistics, Recommendations)', 
           ha='center', fontsize=8, style='italic')
    
    # ========== VISUALIZATION LAYER ==========
    ax.text(5, 4.2, '4. VISUALIZATION', ha='center', fontsize=14, fontweight='bold', color='#F57C00')
    
    draw_box(1, 2.5, 8, 1.5, '', module_color)
    ax.text(5, 3.8, 'DroneVisualizer', ha='center', fontsize=12, fontweight='bold')
    
    # Visualization components
    draw_box(1.2, 3.1, 1.7, 0.5, 'Heat Map\nGenerator', '#FFE0B2', 8)
    draw_box(3.0, 3.1, 1.7, 0.5, 'Priority Zone\nClustering', '#FFE0B2', 8)
    draw_box(4.8, 3.1, 1.7, 0.5, 'Dashboard\nCreator', '#FFE0B2', 8)
    draw_box(6.6, 3.1, 1.7, 0.5, 'Comparison\nReports', '#FFE0B2', 8)
    
    draw_box(2.5, 2.6, 2, 0.4, 'Color Mapping', '#FFE0B2', 7)
    draw_box(4.6, 2.6, 2, 0.4, 'Statistical Charts', '#FFE0B2', 7)
    
    # ========== OUTPUT LAYER ==========
    ax.text(5, 1.9, '5. OUTPUTS', ha='center', fontsize=14, fontweight='bold', color='#388E3C')
    
    draw_box(0.5, 0.5, 2, 0.8, 'Heat Map\nImage', output_color, 9)
    draw_box(2.7, 0.5, 2, 0.8, 'Priority Zones\nImage', output_color, 9)
    draw_box(4.9, 0.5, 2, 0.8, 'Statistical\nDashboard', output_color, 9)
    draw_box(7.1, 0.5, 2, 0.8, 'JSON\nReport', output_color, 9)
    
    # ========== ARROWS ==========
    # Input to Processing
    draw_arrow(2, 9.5, 3, 8.5)
    draw_arrow(5, 9.5, 5, 8.5)
    draw_arrow(8, 9.5, 7, 8.5)
    
    # Processing to Analysis
    draw_arrow(5, 6.5, 5, 5.7)
    
    # Analysis to Visualization
    draw_arrow(5, 5.0, 5, 4.0)
    
    # Visualization to Outputs
    draw_arrow(2, 2.5, 1.5, 1.3)
    draw_arrow(3.5, 2.5, 3.7, 1.3)
    draw_arrow(6.5, 2.5, 5.9, 1.3)
    draw_arrow(7.5, 2.5, 8.1, 1.3)
    
    # ========== LEGEND ==========
    legend_elements = [
        mpatches.Patch(facecolor=input_color, edgecolor='black', label='Input Data'),
        mpatches.Patch(facecolor=module_color, edgecolor='black', label='Processing Module'),
        mpatches.Patch(facecolor=process_color, edgecolor='black', label='Intermediate Data'),
        mpatches.Patch(facecolor=output_color, edgecolor='black', label='Output Files')
    ]
    ax.legend(handles=legend_elements, loc='upper right', fontsize=9)
    
    # ========== WORKFLOW NOTES ==========
    notes = """
    KEY FEATURES:
    • Tile-based processing for large images
    • GPU acceleration support
    • Confidence-based filtering
    • Multi-format output
    • Batch processing ready
    • Treatment recommendations
    """
    
    ax.text(0.2, 10.5, notes, fontsize=8, verticalalignment='top',
           bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.3))
    
    plt.tight_layout()
    plt.savefig('drone_examples/architecture_diagram.png', dpi=300, bbox_inches='tight')
    print("[+] Architecture diagram saved: drone_examples/architecture_diagram.png")
    plt.close()


def create_workflow_diagram():
    """Create a simple workflow diagram"""
    
    fig, ax = plt.subplots(1, 1, figsize=(14, 10))
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 12)
    ax.axis('off')
    
    ax.text(5, 11.5, 'Typical Drone Processing Workflow', 
            ha='center', va='top', fontsize=18, fontweight='bold')
    
    steps = [
        (10.5, "📸 Fly Drone & Capture Images", '#E3F2FD'),
        (9.5, "💾 Download Images to Computer", '#F3E5F5'),
        (8.5, "⚙️ Run process_drone_image.py", '#FFF9C4'),
        (7.5, "🔍 AI Analyzes Each Tile", '#FFE0B2'),
        (6.5, "📊 Generate Analysis Report", '#C8E6C9'),
        (5.5, "🗺️ Create Heat Maps", '#B2EBF2'),
        (4.5, "🎯 Identify Priority Zones", '#FFCCBC'),
        (3.5, "📈 Generate Dashboard", '#D1C4E9'),
        (2.5, "📋 Review Recommendations", '#C5E1A5'),
        (1.5, "🚜 Plan Treatment Actions", '#FFECB3'),
        (0.5, "📆 Schedule Next Flight", '#CFD8DC')
    ]
    
    for i, (y, text, color) in enumerate(steps):
        # Draw box
        box = FancyBboxPatch((1, y-0.3), 8, 0.6, 
                            boxstyle="round,pad=0.05", 
                            facecolor=color, edgecolor='black', linewidth=2)
        ax.add_patch(box)
        
        # Add text
        ax.text(5, y, f"{i+1}. {text}", ha='center', va='center', 
               fontsize=11, fontweight='bold')
        
        # Add arrow
        if i < len(steps) - 1:
            arrow = FancyArrowPatch((5, y-0.35), (5, steps[i+1][0]+0.35),
                                  arrowstyle='->', mutation_scale=25, 
                                  linewidth=2, color='#37474F')
            ax.add_patch(arrow)
    
    # Add time estimates
    time_estimates = [
        "30-60 min", "5 min", "2-5 min", "1-3 min",
        "< 1 min", "< 1 min", "< 1 min", "< 1 min",
        "10-30 min", "Variable", "Weekly"
    ]
    
    for i, (y, _, _) in enumerate(steps):
        ax.text(9.5, y, f"⏱ {time_estimates[i]}", ha='left', va='center',
               fontsize=8, style='italic', color='#666')
    
    # Add cycle arrow
    arc = mpatches.FancyArrowPatch((5, 0.2), (5, 11),
                                  connectionstyle="arc3,rad=.5",
                                  arrowstyle='->', mutation_scale=20,
                                  linewidth=2, color='#1976D2',
                                  linestyle='--', alpha=0.3)
    ax.add_patch(arc)
    ax.text(0.5, 6, 'Continuous\nMonitoring\nCycle', rotation=90,
           ha='center', va='center', fontsize=10, color='#1976D2',
           fontweight='bold', alpha=0.7)
    
    plt.tight_layout()
    plt.savefig('drone_examples/workflow_diagram.png', dpi=300, bbox_inches='tight')
    print("[+] Workflow diagram saved: drone_examples/workflow_diagram.png")
    plt.close()


if __name__ == "__main__":
    print("Generating architecture diagrams...")
    print()
    
    try:
        create_architecture_diagram()
        create_workflow_diagram()
        print()
        print("✓ All diagrams generated successfully!")
        print("  • architecture_diagram.png - System architecture")
        print("  • workflow_diagram.png - User workflow")
    except Exception as e:
        print(f"Error generating diagrams: {e}")
        import traceback
        traceback.print_exc()

