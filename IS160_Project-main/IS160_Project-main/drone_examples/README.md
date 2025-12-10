# Agricultural Drone Image Processing

This module provides tools to process aerial drone imagery for large-scale plant disease detection.

## Features

✅ **Tile-based Processing**: Split large drone images into smaller tiles for efficient processing
✅ **Disease Heat Maps**: Visualize disease distribution across your field
✅ **Priority Zones**: Identify and prioritize treatment areas using clustering
✅ **Before/After Analysis**: Track treatment effectiveness over time
✅ **Comprehensive Reports**: Generate detailed analysis with recommendations

## Quick Start

### 1. Basic Processing

Process a single drone image:

```bash
python process_drone_image.py --image path/to/drone_image.jpg --visualize
```

### 2. Custom Configuration

```bash
python process_drone_image.py \
  --image drone_field_photo.jpg \
  --model models/checkpoints/resnet34_best.pth \
  --tile-size 256 \
  --overlap 0.15 \
  --confidence 0.75 \
  --visualize
```

### 3. Batch Processing

Process multiple images:

```bash
python drone_examples/batch_process_drones.py --input-dir drone_images/ --output-dir results/batch_analysis/
```

## Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| `--image` | *required* | Path to drone image |
| `--model` | `models/checkpoints/resnet34_best.pth` | Path to trained model |
| `--output` | `results/drone_analysis` | Output directory |
| `--tile-size` | `224` | Size of tiles (pixels) |
| `--overlap` | `0.1` | Overlap between tiles (0.0-1.0) |
| `--confidence` | `0.7` | Minimum confidence threshold |
| `--visualize` | `False` | Generate all visualizations |

## Output Files

Each processing run generates:

1. **`analysis_[timestamp].json`** - Complete analysis data
2. **`heatmap_[timestamp].jpg`** - Disease heat map overlay
3. **`priority_zones_[timestamp].jpg`** - Treatment priority zones
4. **`dashboard_[timestamp].png`** - Statistical dashboard

## Python API Usage

### Basic Usage

```python
from src.drone_processor import DroneImageProcessor, get_class_names_from_checkpoint
from src.drone_visualizer import DroneVisualizer

# Load model and class names
model_path = 'models/checkpoints/resnet34_best.pth'
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
    image_path='drone_image.jpg',
    output_dir='results/analysis'
)

# Generate visualizations
visualizer = DroneVisualizer(output_dir='results/analysis')
visualizer.create_heat_map('drone_image.jpg', analysis)
visualizer.create_priority_zones('drone_image.jpg', analysis)
visualizer.create_dashboard(analysis)
```

### Before/After Comparison

```python
# Process before and after images
before_analysis = processor.process_drone_image('field_before_treatment.jpg')
after_analysis = processor.process_drone_image('field_after_treatment.jpg')

# Create comparison report
visualizer.create_comparison_report(
    before_image='field_before_treatment.jpg',
    after_image='field_after_treatment.jpg',
    before_analysis=before_analysis,
    after_analysis=after_analysis
)
```

## Understanding the Output

### Analysis JSON Structure

```json
{
  "metadata": {
    "image_name": "field_001.jpg",
    "image_size": {"width": 4000, "height": 3000},
    "timestamp": "2025-11-29T10:30:00",
    "tile_size": 224,
    "confidence_threshold": 0.7
  },
  "summary": {
    "total_tiles_analyzed": 2340,
    "tiles_with_detections": 450,
    "coverage_percentage": 19.23,
    "healthy_tiles": 280,
    "diseased_tiles": 170,
    "disease_ratio": 0.38
  },
  "disease_distribution": {
    "Tomato___Early_blight": 85,
    "Tomato___Late_blight": 45,
    "Tomato___healthy": 280
  },
  "recommendations": [
    "⚠ Moderate infection rate. Consider targeted treatment.",
    "Primary concern: Tomato___Early_blight (85 detections)",
    "Generate heat map to identify priority treatment zones."
  ]
}
```

### Priority Zones

Priority zones are calculated based on:
- **Area**: Number of diseased tiles in cluster
- **Confidence**: Average detection confidence
- **Priority Score**: `area × confidence`

Higher priority zones should be treated first.

## Use Cases

### 1. Field Scouting
- Weekly or bi-weekly flights over fields
- Early disease detection before visible to naked eye
- Quantify disease spread over time

### 2. Treatment Planning
- Identify hotspots that need immediate attention
- Calculate treatment costs based on affected area
- Optimize pesticide/fungicide application

### 3. Treatment Verification
- Compare before/after treatment images
- Measure treatment effectiveness
- Document results for records/insurance

### 4. Research & Analytics
- Collect longitudinal data on disease progression
- Study effectiveness of different treatments
- Generate reports for stakeholders

## Best Practices

### Image Capture
- **Altitude**: 20-50 meters for detail
- **Resolution**: Minimum 12MP camera
- **Lighting**: Consistent lighting, avoid harsh shadows
- **Overlap**: 60-70% overlap between flight passes
- **Weather**: Clear days, minimal wind

### Processing
- **Tile Size**: 224-256 pixels for best accuracy
- **Overlap**: 10-20% to avoid missing edge cases
- **Confidence**: 0.6-0.8 depending on use case
- **GPU**: Use CUDA-enabled GPU for faster processing

### Interpretation
- Cross-reference with ground truth
- Consider environmental factors (weather, irrigation)
- Track trends over multiple flights
- Consult with agronomists for treatment decisions

## Integration with Drone Software

### DJI Drones
Compatible with images from DJI Phantom, Mavic, and Matrice series.

### Processing Workflow
1. Capture images using DJI Pilot or similar
2. Export georeferenced images
3. Process with this system
4. Import results back to mission planning software

## Performance

### Processing Speed
- CPU: ~0.5-1 second per tile
- GPU (CUDA): ~0.05-0.1 second per tile

### Example: 4000x3000 image
- Tiles generated: ~2400
- GPU processing time: ~2-4 minutes
- CPU processing time: ~20-40 minutes

**Recommendation**: Use GPU for production use.

## Troubleshooting

### Out of Memory Error
- Reduce tile size
- Process smaller image sections
- Use a machine with more RAM/VRAM

### Low Detection Count
- Lower confidence threshold
- Check image quality and lighting
- Ensure model is trained on similar crops

### Inaccurate Detections
- Increase confidence threshold
- Retrain model with more diverse data
- Verify image resolution is sufficient

## Future Enhancements

🔄 **Planned Features**:
- Real-time processing for drone video feeds
- Multi-spectral image support (IR, NDVI)
- GPS coordinate integration
- Automated treatment recommendations
- Integration with precision agriculture systems
- Mobile app for field inspection

## Support

For issues or questions:
1. Check the documentation
2. Review example scripts in `drone_examples/`
3. Open an issue on GitHub

## License

Same as parent project license.

