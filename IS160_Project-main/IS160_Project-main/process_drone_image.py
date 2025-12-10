"""
Main script to process drone images for agricultural disease detection
Usage: python process_drone_image.py --image path/to/drone_image.jpg
"""

import argparse
from pathlib import Path
import sys

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / 'src'))

from drone_processor import DroneImageProcessor, get_class_names_from_checkpoint
from drone_visualizer import DroneVisualizer


def main():
    parser = argparse.ArgumentParser(
        description='Process drone imagery for plant disease detection'
    )
    
    parser.add_argument(
        '--image',
        type=str,
        required=True,
        help='Path to drone image to process'
    )
    
    parser.add_argument(
        '--model',
        type=str,
        default='models/checkpoints/resnet34_best.pth',
        help='Path to trained model checkpoint'
    )
    
    parser.add_argument(
        '--output',
        type=str,
        default='results/drone_analysis',
        help='Output directory for results'
    )
    
    parser.add_argument(
        '--tile-size',
        type=int,
        default=224,
        help='Size of tiles to extract from image'
    )
    
    parser.add_argument(
        '--overlap',
        type=float,
        default=0.1,
        help='Overlap percentage between tiles (0.0 to 1.0)'
    )
    
    parser.add_argument(
        '--confidence',
        type=float,
        default=0.7,
        help='Minimum confidence threshold for detections'
    )
    
    parser.add_argument(
        '--visualize',
        action='store_true',
        help='Generate visualizations (heat maps, priority zones, dashboard)'
    )
    
    parser.add_argument(
        '--skip-heatmap',
        action='store_true',
        help='Skip heat map generation'
    )
    
    parser.add_argument(
        '--skip-zones',
        action='store_true',
        help='Skip priority zones generation'
    )
    
    parser.add_argument(
        '--skip-dashboard',
        action='store_true',
        help='Skip dashboard generation'
    )
    
    args = parser.parse_args()
    
    # Validate inputs
    image_path = Path(args.image)
    if not image_path.exists():
        print(f"[!] Error: Image not found: {image_path}")
        return
    
    model_path = Path(args.model)
    if not model_path.exists():
        print(f"[!] Error: Model not found: {model_path}")
        print(f"[!] Please train a model first using: python run_training.py")
        return
    
    # Create output directory
    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    print("="*80)
    print("AGRICULTURAL DRONE IMAGE PROCESSING SYSTEM")
    print("="*80)
    print(f"Image: {image_path}")
    print(f"Model: {model_path}")
    print(f"Output: {output_dir}")
    print(f"Tile Size: {args.tile_size}x{args.tile_size}")
    print(f"Overlap: {args.overlap*100}%")
    print(f"Confidence Threshold: {args.confidence}")
    print("="*80)
    
    # Get class names
    print("\n[+] Loading class names...")
    class_names = get_class_names_from_checkpoint(str(model_path))
    print(f"[+] Detected {len(class_names)} disease classes")
    
    # Initialize processor
    processor = DroneImageProcessor(
        model_path=str(model_path),
        class_names=class_names,
        tile_size=args.tile_size,
        overlap=args.overlap,
        confidence_threshold=args.confidence
    )
    
    # Process image
    analysis = processor.process_drone_image(
        image_path=str(image_path),
        output_dir=str(output_dir)
    )
    
    # Print summary
    print("\n" + "="*80)
    print("ANALYSIS COMPLETE")
    print("="*80)
    print(f"Disease Coverage: {analysis['summary']['coverage_percentage']:.2f}%")
    print(f"Healthy Tiles: {analysis['summary']['healthy_tiles']}")
    print(f"Diseased Tiles: {analysis['summary']['diseased_tiles']}")
    print("\nTop Diseases Detected:")
    
    disease_dist = analysis['disease_distribution']
    sorted_diseases = sorted(disease_dist.items(), key=lambda x: x[1], reverse=True)
    for i, (disease, count) in enumerate(sorted_diseases[:5], 1):
        print(f"  {i}. {disease}: {count} detections")
    
    print("\nRecommendations:")
    for i, rec in enumerate(analysis['recommendations'], 1):
        print(f"  {i}. {rec}")
    
    # Generate visualizations
    if args.visualize or not (args.skip_heatmap and args.skip_zones and args.skip_dashboard):
        print("\n" + "="*80)
        print("GENERATING VISUALIZATIONS")
        print("="*80)
        
        visualizer = DroneVisualizer(output_dir=str(output_dir))
        
        if not args.skip_heatmap:
            heat_map_path = visualizer.create_heat_map(
                original_image_path=str(image_path),
                analysis=analysis
            )
        
        if not args.skip_zones:
            try:
                zones_path = visualizer.create_priority_zones(
                    original_image_path=str(image_path),
                    analysis=analysis
                )
            except Exception as e:
                print(f"[!] Could not generate priority zones: {e}")
        
        if not args.skip_dashboard:
            dashboard_path = visualizer.create_dashboard(analysis=analysis)
    
    print("\n" + "="*80)
    print("PROCESSING COMPLETE!")
    print("="*80)
    print(f"Results saved to: {output_dir}")
    print("\nGenerated files:")
    print(f"  • Analysis JSON: analysis_*.json")
    if not args.skip_heatmap:
        print(f"  • Heat Map: heatmap_*.jpg")
    if not args.skip_zones:
        print(f"  • Priority Zones: priority_zones_*.jpg")
    if not args.skip_dashboard:
        print(f"  • Dashboard: dashboard_*.png")
    print("="*80)


if __name__ == "__main__":
    main()

