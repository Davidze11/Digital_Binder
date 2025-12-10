"""
Example usage of the drone processing system
Demonstrates different use cases and workflows
"""

import sys
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent.parent / 'src'))

from drone_processor import DroneImageProcessor, get_class_names_from_checkpoint
from drone_visualizer import DroneVisualizer


def example_basic_processing():
    """Example 1: Basic drone image processing"""
    
    print("\n" + "="*80)
    print("EXAMPLE 1: Basic Drone Image Processing")
    print("="*80)
    
    # Configuration
    model_path = 'models/checkpoints/resnet34_best.pth'
    image_path = 'path/to/your/drone_image.jpg'  # Replace with your image
    output_dir = 'results/drone_analysis/example1'
    
    # Load class names
    class_names = get_class_names_from_checkpoint(model_path)
    
    # Initialize processor
    processor = DroneImageProcessor(
        model_path=model_path,
        class_names=class_names,
        tile_size=224,
        overlap=0.1,
        confidence_threshold=0.7
    )
    
    # Process image
    analysis = processor.process_drone_image(
        image_path=image_path,
        output_dir=output_dir
    )
    
    # Print results
    print(f"\n✓ Disease Coverage: {analysis['summary']['coverage_percentage']:.2f}%")
    print(f"✓ Diseased Areas: {analysis['summary']['diseased_tiles']} tiles")
    print(f"✓ Healthy Areas: {analysis['summary']['healthy_tiles']} tiles")
    
    return analysis


def example_with_visualizations():
    """Example 2: Processing with full visualizations"""
    
    print("\n" + "="*80)
    print("EXAMPLE 2: Processing with Visualizations")
    print("="*80)
    
    model_path = 'models/checkpoints/resnet34_best.pth'
    image_path = 'path/to/your/drone_image.jpg'
    output_dir = 'results/drone_analysis/example2'
    
    # Process image
    class_names = get_class_names_from_checkpoint(model_path)
    processor = DroneImageProcessor(model_path, class_names)
    analysis = processor.process_drone_image(image_path, output_dir)
    
    # Create visualizations
    visualizer = DroneVisualizer(output_dir=output_dir)
    
    # Heat map
    print("\n[+] Generating heat map...")
    heat_map_path = visualizer.create_heat_map(image_path, analysis)
    print(f"✓ Heat map saved: {heat_map_path}")
    
    # Priority zones
    print("\n[+] Generating priority zones...")
    zones_path = visualizer.create_priority_zones(image_path, analysis)
    if zones_path:
        print(f"✓ Priority zones saved: {zones_path}")
    
    # Dashboard
    print("\n[+] Generating dashboard...")
    dashboard_path = visualizer.create_dashboard(analysis)
    print(f"✓ Dashboard saved: {dashboard_path}")
    
    return analysis


def example_before_after_comparison():
    """Example 3: Before/After treatment comparison"""
    
    print("\n" + "="*80)
    print("EXAMPLE 3: Before/After Treatment Comparison")
    print("="*80)
    
    model_path = 'models/checkpoints/resnet34_best.pth'
    before_image = 'path/to/field_before_treatment.jpg'
    after_image = 'path/to/field_after_treatment.jpg'
    output_dir = 'results/drone_analysis/comparison'
    
    # Initialize
    class_names = get_class_names_from_checkpoint(model_path)
    processor = DroneImageProcessor(model_path, class_names)
    visualizer = DroneVisualizer(output_dir=output_dir)
    
    # Process before image
    print("\n[1/3] Processing BEFORE treatment image...")
    before_analysis = processor.process_drone_image(before_image, output_dir)
    print(f"     Disease coverage: {before_analysis['summary']['coverage_percentage']:.2f}%")
    
    # Process after image
    print("\n[2/3] Processing AFTER treatment image...")
    after_analysis = processor.process_drone_image(after_image, output_dir)
    print(f"     Disease coverage: {after_analysis['summary']['coverage_percentage']:.2f}%")
    
    # Generate comparison
    print("\n[3/3] Generating comparison report...")
    comparison_path = visualizer.create_comparison_report(
        before_image=before_image,
        after_image=after_image,
        before_analysis=before_analysis,
        after_analysis=after_analysis
    )
    
    # Calculate improvement
    before_diseased = before_analysis['summary']['diseased_tiles']
    after_diseased = after_analysis['summary']['diseased_tiles']
    improvement = (before_diseased - after_diseased) / before_diseased * 100
    
    print(f"\n✓ Comparison report saved: {comparison_path}")
    print(f"✓ Disease reduction: {improvement:.1f}%")
    
    return before_analysis, after_analysis


def example_batch_processing():
    """Example 4: Batch process multiple images"""
    
    print("\n" + "="*80)
    print("EXAMPLE 4: Batch Processing Multiple Images")
    print("="*80)
    
    model_path = 'models/checkpoints/resnet34_best.pth'
    input_dir = Path('path/to/drone_images_folder')
    output_dir = Path('results/drone_analysis/batch')
    
    # Get all images
    image_files = list(input_dir.glob('*.jpg')) + list(input_dir.glob('*.png'))
    print(f"\n[+] Found {len(image_files)} images to process")
    
    # Initialize
    class_names = get_class_names_from_checkpoint(model_path)
    processor = DroneImageProcessor(model_path, class_names)
    visualizer = DroneVisualizer(output_dir=str(output_dir))
    
    results = []
    
    for i, image_path in enumerate(image_files, 1):
        print(f"\n[{i}/{len(image_files)}] Processing {image_path.name}...")
        
        try:
            # Process image
            analysis = processor.process_drone_image(
                image_path=str(image_path),
                output_dir=str(output_dir / image_path.stem)
            )
            
            # Generate visualizations
            visualizer.create_heat_map(str(image_path), analysis,
                                      output_name=f"{image_path.stem}_heatmap.jpg")
            visualizer.create_dashboard(analysis,
                                       output_name=f"{image_path.stem}_dashboard.png")
            
            results.append({
                'image': image_path.name,
                'coverage': analysis['summary']['coverage_percentage'],
                'diseased_tiles': analysis['summary']['diseased_tiles'],
                'status': 'success'
            })
            
            print(f"    ✓ Coverage: {analysis['summary']['coverage_percentage']:.2f}%")
            
        except Exception as e:
            print(f"    ✗ Error: {e}")
            results.append({
                'image': image_path.name,
                'status': 'failed',
                'error': str(e)
            })
    
    # Summary
    print("\n" + "="*80)
    print("BATCH PROCESSING SUMMARY")
    print("="*80)
    successful = sum(1 for r in results if r['status'] == 'success')
    print(f"Processed: {successful}/{len(image_files)} images")
    
    if successful > 0:
        avg_coverage = sum(r['coverage'] for r in results if r['status'] == 'success') / successful
        print(f"Average disease coverage: {avg_coverage:.2f}%")
    
    return results


def example_custom_configuration():
    """Example 5: Custom configuration for specific needs"""
    
    print("\n" + "="*80)
    print("EXAMPLE 5: Custom Configuration")
    print("="*80)
    
    model_path = 'models/checkpoints/resnet34_best.pth'
    image_path = 'path/to/your/drone_image.jpg'
    output_dir = 'results/drone_analysis/custom'
    
    class_names = get_class_names_from_checkpoint(model_path)
    
    # Custom configuration for high-resolution images
    processor = DroneImageProcessor(
        model_path=model_path,
        class_names=class_names,
        tile_size=256,           # Larger tiles for high-res images
        overlap=0.2,             # More overlap for better coverage
        confidence_threshold=0.8  # Higher confidence for fewer false positives
    )
    
    analysis = processor.process_drone_image(image_path, output_dir)
    
    print(f"\n✓ Processed with custom settings")
    print(f"✓ Tile size: 256x256")
    print(f"✓ Overlap: 20%")
    print(f"✓ Confidence threshold: 0.8")
    print(f"✓ Results: {analysis['summary']['tiles_with_detections']} high-confidence detections")
    
    return analysis


def print_analysis_summary(analysis: dict):
    """Helper function to print analysis summary"""
    
    print("\n" + "="*80)
    print("ANALYSIS SUMMARY")
    print("="*80)
    
    summary = analysis['summary']
    print(f"\nField Coverage:")
    print(f"  • Total tiles analyzed: {summary['total_tiles_analyzed']}")
    print(f"  • Disease coverage: {summary['coverage_percentage']:.2f}%")
    print(f"  • Healthy areas: {summary['healthy_tiles']} tiles")
    print(f"  • Diseased areas: {summary['diseased_tiles']} tiles")
    
    print(f"\nTop Diseases:")
    disease_dist = analysis['disease_distribution']
    sorted_diseases = sorted(disease_dist.items(), key=lambda x: x[1], reverse=True)
    
    for i, (disease, count) in enumerate(sorted_diseases[:5], 1):
        disease_name = disease.split('___')[-1]
        print(f"  {i}. {disease_name}: {count} detections")
    
    print(f"\nRecommendations:")
    for i, rec in enumerate(analysis['recommendations'], 1):
        print(f"  {i}. {rec}")
    
    print("="*80)


if __name__ == "__main__":
    print("""
    ╔═══════════════════════════════════════════════════════════════════════╗
    ║         AGRICULTURAL DRONE PROCESSING - EXAMPLE SCRIPTS               ║
    ╚═══════════════════════════════════════════════════════════════════════╝
    
    This script demonstrates various use cases for the drone processing system.
    
    Available Examples:
    1. Basic Processing - Simple image analysis
    2. Full Visualizations - Heat maps, zones, and dashboard
    3. Before/After Comparison - Treatment effectiveness
    4. Batch Processing - Process multiple images
    5. Custom Configuration - Advanced settings
    
    To run an example, uncomment the function call below and update the image paths.
    """)
    
    # Uncomment the example you want to run:
    
    # analysis = example_basic_processing()
    # print_analysis_summary(analysis)
    
    # analysis = example_with_visualizations()
    # print_analysis_summary(analysis)
    
    # before, after = example_before_after_comparison()
    
    # results = example_batch_processing()
    
    # analysis = example_custom_configuration()
    # print_analysis_summary(analysis)
    
    print("\n[!] Update the image paths in this script and uncomment an example to run.")
    print("[!] Make sure you have a trained model at: models/checkpoints/resnet34_best.pth")

