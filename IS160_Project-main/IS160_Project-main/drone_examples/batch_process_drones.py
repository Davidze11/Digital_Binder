"""
Batch process multiple drone images
Useful for processing entire flight missions or multiple fields
"""

import argparse
import sys
from pathlib import Path
from datetime import datetime
import json

# Add src to path
sys.path.insert(0, str(Path(__file__).parent.parent / 'src'))

from drone_processor import DroneImageProcessor, get_class_names_from_checkpoint
from drone_visualizer import DroneVisualizer


def batch_process(
    input_dir: str,
    output_dir: str,
    model_path: str,
    file_pattern: str = "*.jpg",
    visualize: bool = True,
    **processor_kwargs
):
    """
    Batch process multiple drone images.
    
    Args:
        input_dir: Directory containing drone images
        output_dir: Directory to save results
        model_path: Path to trained model
        file_pattern: File pattern to match (default: *.jpg)
        visualize: Whether to generate visualizations
        **processor_kwargs: Additional arguments for DroneImageProcessor
    """
    input_path = Path(input_dir)
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)
    
    # Find all images
    image_files = list(input_path.glob(file_pattern))
    if not image_files:
        print(f"[!] No images found matching pattern: {file_pattern}")
        return
    
    print("="*80)
    print("BATCH DRONE IMAGE PROCESSING")
    print("="*80)
    print(f"Input directory: {input_path}")
    print(f"Output directory: {output_path}")
    print(f"Found {len(image_files)} images to process")
    print("="*80)
    
    # Initialize processor
    class_names = get_class_names_from_checkpoint(model_path)
    processor = DroneImageProcessor(
        model_path=model_path,
        class_names=class_names,
        **processor_kwargs
    )
    
    if visualize:
        visualizer = DroneVisualizer(output_dir=str(output_path))
    
    # Process each image
    results = []
    start_time = datetime.now()
    
    for i, image_path in enumerate(image_files, 1):
        print(f"\n{'='*80}")
        print(f"[{i}/{len(image_files)}] Processing: {image_path.name}")
        print(f"{'='*80}")
        
        image_start = datetime.now()
        
        try:
            # Create output subdirectory for this image
            image_output_dir = output_path / image_path.stem
            image_output_dir.mkdir(parents=True, exist_ok=True)
            
            # Process image
            analysis = processor.process_drone_image(
                image_path=str(image_path),
                output_dir=str(image_output_dir)
            )
            
            # Generate visualizations
            if visualize:
                print("\n[+] Generating visualizations...")
                
                try:
                    visualizer.create_heat_map(
                        str(image_path),
                        analysis,
                        output_name=f"{image_path.stem}_heatmap.jpg"
                    )
                except Exception as e:
                    print(f"[!] Heat map error: {e}")
                
                try:
                    visualizer.create_priority_zones(
                        str(image_path),
                        analysis,
                        output_name=f"{image_path.stem}_zones.jpg"
                    )
                except Exception as e:
                    print(f"[!] Priority zones error: {e}")
                
                try:
                    visualizer.create_dashboard(
                        analysis,
                        output_name=f"{image_path.stem}_dashboard.png"
                    )
                except Exception as e:
                    print(f"[!] Dashboard error: {e}")
            
            # Record results
            processing_time = (datetime.now() - image_start).total_seconds()
            
            result = {
                'image': image_path.name,
                'status': 'success',
                'processing_time_seconds': round(processing_time, 2),
                'summary': analysis['summary'],
                'top_diseases': dict(sorted(
                    analysis['disease_distribution'].items(),
                    key=lambda x: x[1],
                    reverse=True
                )[:5])
            }
            
            results.append(result)
            
            print(f"\n✓ Successfully processed in {processing_time:.1f}s")
            print(f"  • Disease coverage: {analysis['summary']['coverage_percentage']:.2f}%")
            print(f"  • Diseased tiles: {analysis['summary']['diseased_tiles']}")
            
        except Exception as e:
            print(f"\n✗ Error processing image: {e}")
            results.append({
                'image': image_path.name,
                'status': 'failed',
                'error': str(e)
            })
    
    # Generate summary report
    total_time = (datetime.now() - start_time).total_seconds()
    
    print("\n" + "="*80)
    print("BATCH PROCESSING SUMMARY")
    print("="*80)
    
    successful = [r for r in results if r['status'] == 'success']
    failed = [r for r in results if r['status'] == 'failed']
    
    print(f"\nTotal images processed: {len(image_files)}")
    print(f"  ✓ Successful: {len(successful)}")
    print(f"  ✗ Failed: {len(failed)}")
    print(f"  ⏱ Total time: {total_time:.1f}s ({total_time/60:.1f} minutes)")
    
    if successful:
        avg_coverage = sum(r['summary']['coverage_percentage'] for r in successful) / len(successful)
        total_diseased = sum(r['summary']['diseased_tiles'] for r in successful)
        avg_time = sum(r['processing_time_seconds'] for r in successful) / len(successful)
        
        print(f"\nStatistics:")
        print(f"  • Average disease coverage: {avg_coverage:.2f}%")
        print(f"  • Total diseased tiles: {total_diseased}")
        print(f"  • Average processing time: {avg_time:.1f}s per image")
        
        # Most common diseases across all images
        all_diseases = {}
        for r in successful:
            for disease, count in r['top_diseases'].items():
                all_diseases[disease] = all_diseases.get(disease, 0) + count
        
        if all_diseases:
            print(f"\nMost common diseases across all fields:")
            sorted_diseases = sorted(all_diseases.items(), key=lambda x: x[1], reverse=True)
            for i, (disease, count) in enumerate(sorted_diseases[:5], 1):
                disease_name = disease.split('___')[-1]
                print(f"  {i}. {disease_name}: {count} total detections")
    
    if failed:
        print(f"\nFailed images:")
        for r in failed:
            print(f"  • {r['image']}: {r['error']}")
    
    # Save summary to JSON
    summary_file = output_path / f"batch_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    with open(summary_file, 'w') as f:
        json.dump({
            'timestamp': datetime.now().isoformat(),
            'input_directory': str(input_path),
            'total_images': len(image_files),
            'successful': len(successful),
            'failed': len(failed),
            'total_processing_time': total_time,
            'results': results
        }, f, indent=2)
    
    print(f"\n✓ Summary saved to: {summary_file}")
    print("="*80)


def main():
    parser = argparse.ArgumentParser(
        description='Batch process multiple drone images'
    )
    
    parser.add_argument(
        '--input-dir',
        type=str,
        required=True,
        help='Directory containing drone images'
    )
    
    parser.add_argument(
        '--output-dir',
        type=str,
        default='results/drone_batch',
        help='Output directory for results'
    )
    
    parser.add_argument(
        '--model',
        type=str,
        default='models/checkpoints/resnet34_best.pth',
        help='Path to trained model checkpoint'
    )
    
    parser.add_argument(
        '--pattern',
        type=str,
        default='*.jpg',
        help='File pattern to match (e.g., *.jpg, *.png)'
    )
    
    parser.add_argument(
        '--tile-size',
        type=int,
        default=224,
        help='Size of tiles to extract'
    )
    
    parser.add_argument(
        '--overlap',
        type=float,
        default=0.1,
        help='Overlap percentage between tiles'
    )
    
    parser.add_argument(
        '--confidence',
        type=float,
        default=0.7,
        help='Minimum confidence threshold'
    )
    
    parser.add_argument(
        '--no-visualize',
        action='store_true',
        help='Skip visualization generation (faster)'
    )
    
    args = parser.parse_args()
    
    # Validate inputs
    if not Path(args.input_dir).exists():
        print(f"[!] Error: Input directory not found: {args.input_dir}")
        return
    
    if not Path(args.model).exists():
        print(f"[!] Error: Model not found: {args.model}")
        return
    
    # Run batch processing
    batch_process(
        input_dir=args.input_dir,
        output_dir=args.output_dir,
        model_path=args.model,
        file_pattern=args.pattern,
        visualize=not args.no_visualize,
        tile_size=args.tile_size,
        overlap=args.overlap,
        confidence_threshold=args.confidence
    )


if __name__ == "__main__":
    main()

