"""
Quick demo to test the drone processing system with a sample image
Generates a synthetic test image to demonstrate functionality
"""

import cv2
import numpy as np
from pathlib import Path
import sys

# Add src to path
sys.path.insert(0, str(Path(__file__).parent.parent / 'src'))

from drone_processor import DroneImageProcessor, get_class_names_from_checkpoint
from drone_visualizer import DroneVisualizer


def create_test_drone_image(output_path: str, width: int = 2000, height: int = 1500):
    """
    Create a synthetic drone-like image for testing.
    Simulates a field with some diseased and healthy areas.
    """
    print("[+] Generating synthetic test drone image...")
    
    # Create base image (field texture)
    image = np.random.randint(40, 80, (height, width, 3), dtype=np.uint8)
    
    # Add some greenish tint (simulate vegetation)
    image[:, :, 1] = np.clip(image[:, :, 1] + 80, 0, 255)  # More green
    
    # Add some texture variation
    noise = np.random.randint(-20, 20, (height, width, 3), dtype=np.int16)
    image = np.clip(image.astype(np.int16) + noise, 0, 255).astype(np.uint8)
    
    # Add some "diseased" patches (darker/brownish areas)
    num_disease_patches = 15
    for _ in range(num_disease_patches):
        x = np.random.randint(100, width - 300)
        y = np.random.randint(100, height - 300)
        radius = np.random.randint(50, 150)
        
        # Create circular patch
        y_grid, x_grid = np.ogrid[:height, :width]
        mask = (x_grid - x)**2 + (y_grid - y)**2 <= radius**2
        
        # Make it brownish (reduce green, add some red)
        image[mask, 1] = np.clip(image[mask, 1] - 40, 0, 255)  # Less green
        image[mask, 0] = np.clip(image[mask, 0] - 20, 0, 255)  # Less blue
        image[mask, 2] = np.clip(image[mask, 2] + 20, 0, 255)  # More red
    
    # Add some healthy bright patches
    num_healthy_patches = 10
    for _ in range(num_healthy_patches):
        x = np.random.randint(100, width - 300)
        y = np.random.randint(100, height - 300)
        radius = np.random.randint(30, 100)
        
        y_grid, x_grid = np.ogrid[:height, :width]
        mask = (x_grid - x)**2 + (y_grid - y)**2 <= radius**2
        
        # Make it brighter green (healthy vegetation)
        image[mask, 1] = np.clip(image[mask, 1] + 30, 0, 255)
    
    # Add some row structure (simulate planted rows)
    for row_y in range(0, height, 60):
        cv2.line(image, (0, row_y), (width, row_y), (50, 70, 50), 2)
    
    # Save image
    cv2.imwrite(output_path, image)
    print(f"[+] Test image saved: {output_path}")
    print(f"    Size: {width}x{height} pixels")
    
    return output_path


def run_demo():
    """Run a complete demo of the drone processing system"""
    
    print("""
    ========================================================================
                  DRONE PROCESSING SYSTEM - DEMO                           
    ========================================================================
    
    This demo will:
    1. Create a synthetic test drone image
    2. Process it with the disease detection model
    3. Generate heat maps and visualizations
    4. Create a comprehensive analysis dashboard
    
    """)
    
    # Setup paths
    demo_dir = Path('drone_examples/demo_output')
    demo_dir.mkdir(parents=True, exist_ok=True)
    
    test_image_path = demo_dir / 'test_drone_image.jpg'
    model_path = Path('models/checkpoints/resnet34_best.pth')
    
    # Check if model exists
    if not model_path.exists():
        print("[!] ERROR: Model not found!")
        print(f"[!] Expected location: {model_path}")
        print("[!] Please train a model first using: python run_training.py")
        print("\n[i] Available models:")
        checkpoint_dir = Path('models/checkpoints')
        if checkpoint_dir.exists():
            models = list(checkpoint_dir.glob('*.pth'))
            if models:
                for model in models:
                    print(f"    • {model.name}")
            else:
                print("    No models found. Train one first!")
        return
    
    # Step 1: Create test image
    print("\n" + "="*80)
    print("STEP 1: Creating Test Image")
    print("="*80)
    create_test_drone_image(str(test_image_path), width=2000, height=1500)
    
    # Step 2: Initialize processor
    print("\n" + "="*80)
    print("STEP 2: Initializing Disease Detection System")
    print("="*80)
    
    class_names = get_class_names_from_checkpoint(str(model_path))
    print(f"[+] Loaded {len(class_names)} disease classes")
    
    processor = DroneImageProcessor(
        model_path=str(model_path),
        class_names=class_names,
        tile_size=224,
        overlap=0.1,
        confidence_threshold=0.6  # Lower threshold for demo
    )
    
    # Step 3: Process image
    print("\n" + "="*80)
    print("STEP 3: Processing Drone Image")
    print("="*80)
    
    analysis = processor.process_drone_image(
        image_path=str(test_image_path),
        output_dir=str(demo_dir)
    )
    
    # Step 4: Generate visualizations
    print("\n" + "="*80)
    print("STEP 4: Generating Visualizations")
    print("="*80)
    
    visualizer = DroneVisualizer(output_dir=str(demo_dir))
    
    print("\n[1/3] Creating heat map...")
    heat_map_path = visualizer.create_heat_map(
        str(test_image_path),
        analysis,
        output_name='demo_heatmap.jpg'
    )
    
    print("\n[2/3] Creating priority zones...")
    try:
        zones_path = visualizer.create_priority_zones(
            str(test_image_path),
            analysis,
            output_name='demo_priority_zones.jpg'
        )
    except Exception as e:
        print(f"[!] Note: Could not create priority zones (this is normal for test images): {e}")
    
    print("\n[3/3] Creating dashboard...")
    dashboard_path = visualizer.create_dashboard(
        analysis,
        output_name='demo_dashboard.png'
    )
    
    # Step 5: Display results
    print("\n" + "="*80)
    print("DEMO COMPLETE!")
    print("="*80)
    
    print(f"\n[ANALYSIS RESULTS]")
    print(f"   - Disease Coverage: {analysis['summary']['coverage_percentage']:.2f}%")
    print(f"   - Total Tiles Analyzed: {analysis['summary']['total_tiles_analyzed']}")
    print(f"   - Detections Found: {analysis['summary']['tiles_with_detections']}")
    print(f"   - Healthy Tiles: {analysis['summary']['healthy_tiles']}")
    print(f"   - Diseased Tiles: {analysis['summary']['diseased_tiles']}")
    
    if analysis['disease_distribution']:
        print(f"\n[TOP DETECTIONS]")
        sorted_diseases = sorted(
            analysis['disease_distribution'].items(),
            key=lambda x: x[1],
            reverse=True
        )
        for i, (disease, count) in enumerate(sorted_diseases[:5], 1):
            disease_name = disease.split('___')[-1]
            print(f"   {i}. {disease_name}: {count} tiles")
    
    print(f"\n[OUTPUT FILES]")
    print(f"   - Test Image: {test_image_path}")
    print(f"   - Analysis JSON: {demo_dir / 'analysis_*.json'}")
    print(f"   - Heat Map: {heat_map_path}")
    print(f"   - Dashboard: {dashboard_path}")
    
    print(f"\n[RECOMMENDATIONS]")
    for i, rec in enumerate(analysis['recommendations'], 1):
        print(f"   {i}. {rec}")
    
    print("\n" + "="*80)
    print("SUCCESS! Demo completed successfully!")
    print(f"   Check {demo_dir} for all generated files.")
    print("\n[Next Steps]")
    print("   1. Try processing your own drone images:")
    print("      python process_drone_image.py --image your_image.jpg --visualize")
    print("   2. Check out more examples in drone_examples/example_usage.py")
    print("   3. Read the full documentation in drone_examples/README.md")
    print("="*80)


if __name__ == "__main__":
    try:
        run_demo()
    except Exception as e:
        print(f"\n[!] ERROR during demo: {e}")
        import traceback
        traceback.print_exc()
        print("\n[!] Demo failed. Please check that:")
        print("    1. Model exists at: models/checkpoints/resnet34_best.pth")
        print("    2. All dependencies are installed: pip install -r requirements.txt")
        print("    3. PyTorch is properly installed")

