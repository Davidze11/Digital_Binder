"""
Agricultural Drone Image Processing System
Processes aerial imagery from drones to detect plant diseases at scale
"""

import torch
import torch.nn as nn
from torchvision import models, transforms
from PIL import Image
import numpy as np
import cv2
from pathlib import Path
from typing import List, Tuple, Dict, Optional
import json
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')


class DroneImageProcessor:
    """
    Process large drone images by splitting them into smaller tiles,
    running disease detection on each tile, and generating comprehensive reports.
    """
    
    def __init__(
        self,
        model_path: str,
        class_names: List[str],
        device: str = "cuda" if torch.cuda.is_available() else "cpu",
        tile_size: int = 224,
        overlap: float = 0.1,
        confidence_threshold: float = 0.7
    ):
        """
        Initialize the drone image processor.
        
        Args:
            model_path: Path to the trained PyTorch model
            class_names: List of disease class names
            device: Device to run inference on (cuda/cpu)
            tile_size: Size of each tile to extract from the image
            overlap: Overlap percentage between tiles (0.0 to 1.0)
            confidence_threshold: Minimum confidence for disease detection
        """
        self.device = device
        self.tile_size = tile_size
        self.overlap = overlap
        self.confidence_threshold = confidence_threshold
        self.class_names = class_names
        
        # Load the model
        print(f"[+] Loading model from {model_path}...")
        self.model = self._load_model(model_path, len(class_names))
        self.model.eval()
        
        # Define image transformations
        self.transform = transforms.Compose([
            transforms.Resize((224, 224)),
            transforms.ToTensor(),
            transforms.Normalize(
                mean=[0.485, 0.456, 0.406],
                std=[0.229, 0.224, 0.225]
            )
        ])
        
        print(f"[+] Drone processor initialized on {device}")
        print(f"[+] Tile size: {tile_size}x{tile_size}, Overlap: {overlap*100}%")
    
    def _load_model(self, model_path: str, num_classes: int) -> nn.Module:
        """Load the trained model."""
        model = models.resnet34(pretrained=False)
        model.fc = nn.Linear(model.fc.in_features, num_classes)
        
        checkpoint = torch.load(model_path, map_location=self.device)
        
        # Handle different checkpoint formats
        if 'model_state_dict' in checkpoint:
            model.load_state_dict(checkpoint['model_state_dict'])
        else:
            model.load_state_dict(checkpoint)
        
        model = model.to(self.device)
        return model
    
    def split_into_tiles(
        self,
        image: np.ndarray
    ) -> List[Dict]:
        """
        Split a large image into smaller overlapping tiles.
        
        Args:
            image: Input image as numpy array (H, W, C)
            
        Returns:
            List of tile dictionaries with image data and coordinates
        """
        height, width = image.shape[:2]
        stride = int(self.tile_size * (1 - self.overlap))
        
        tiles = []
        tile_id = 0
        
        for y in range(0, height - self.tile_size + 1, stride):
            for x in range(0, width - self.tile_size + 1, stride):
                # Extract tile
                tile = image[y:y+self.tile_size, x:x+self.tile_size]
                
                tiles.append({
                    'id': tile_id,
                    'image': tile,
                    'x': x,
                    'y': y,
                    'width': self.tile_size,
                    'height': self.tile_size,
                    'center_x': x + self.tile_size // 2,
                    'center_y': y + self.tile_size // 2
                })
                tile_id += 1
        
        print(f"[+] Split image into {len(tiles)} tiles")
        return tiles
    
    def predict_tile(self, tile_image: np.ndarray) -> Tuple[str, float, np.ndarray]:
        """
        Run disease prediction on a single tile.
        
        Args:
            tile_image: Tile image as numpy array
            
        Returns:
            Tuple of (predicted_class, confidence, probabilities)
        """
        # Convert BGR to RGB if needed
        if len(tile_image.shape) == 3 and tile_image.shape[2] == 3:
            tile_image = cv2.cvtColor(tile_image, cv2.COLOR_BGR2RGB)
        
        # Convert to PIL Image
        pil_image = Image.fromarray(tile_image)
        
        # Transform and add batch dimension
        input_tensor = self.transform(pil_image).unsqueeze(0).to(self.device)
        
        # Run inference
        with torch.no_grad():
            outputs = self.model(input_tensor)
            probabilities = torch.softmax(outputs, dim=1)
            confidence, predicted_idx = torch.max(probabilities, 1)
        
        predicted_class = self.class_names[predicted_idx.item()]
        confidence_value = confidence.item()
        probs_array = probabilities.cpu().numpy()[0]
        
        return predicted_class, confidence_value, probs_array
    
    def process_drone_image(
        self,
        image_path: str,
        output_dir: Optional[str] = None
    ) -> Dict:
        """
        Process a complete drone image and generate analysis.
        
        Args:
            image_path: Path to the drone image
            output_dir: Directory to save outputs (optional)
            
        Returns:
            Dictionary containing analysis results
        """
        print(f"\n{'='*80}")
        print(f"PROCESSING DRONE IMAGE")
        print(f"{'='*80}")
        print(f"Image: {image_path}")
        
        # Load image
        image = cv2.imread(str(image_path))
        if image is None:
            raise ValueError(f"Could not load image: {image_path}")
        
        height, width = image.shape[:2]
        print(f"Image size: {width}x{height} pixels")
        
        # Split into tiles
        tiles = self.split_into_tiles(image)
        
        # Process each tile
        print(f"[+] Running disease detection on {len(tiles)} tiles...")
        results = []
        
        for i, tile_info in enumerate(tiles):
            if (i + 1) % 100 == 0:
                print(f"    Processed {i+1}/{len(tiles)} tiles...")
            
            predicted_class, confidence, probabilities = self.predict_tile(tile_info['image'])
            
            # Only keep results above confidence threshold
            if confidence >= self.confidence_threshold:
                results.append({
                    'tile_id': tile_info['id'],
                    'x': tile_info['x'],
                    'y': tile_info['y'],
                    'center_x': tile_info['center_x'],
                    'center_y': tile_info['center_y'],
                    'predicted_class': predicted_class,
                    'confidence': float(confidence),
                    'probabilities': probabilities.tolist()
                })
        
        print(f"[+] Found {len(results)} tiles with diseases (confidence >= {self.confidence_threshold})")
        
        # Generate comprehensive analysis
        analysis = self._generate_analysis(
            results,
            image.shape,
            Path(image_path).name
        )
        
        # Save outputs if directory provided
        if output_dir:
            output_path = Path(output_dir)
            output_path.mkdir(parents=True, exist_ok=True)
            
            # Save analysis JSON
            analysis_file = output_path / f"analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            with open(analysis_file, 'w') as f:
                json.dump(analysis, f, indent=2)
            print(f"[+] Analysis saved to {analysis_file}")
        
        return analysis
    
    def _generate_analysis(
        self,
        results: List[Dict],
        image_shape: Tuple,
        image_name: str
    ) -> Dict:
        """Generate comprehensive analysis from tile results."""
        
        # Count diseases
        disease_counts = {}
        for result in results:
            disease = result['predicted_class']
            disease_counts[disease] = disease_counts.get(disease, 0) + 1
        
        # Calculate coverage
        total_tiles = (image_shape[0] // self.tile_size) * (image_shape[1] // self.tile_size)
        diseased_tiles = len(results)
        coverage_percentage = (diseased_tiles / total_tiles * 100) if total_tiles > 0 else 0
        
        # Identify healthy vs diseased areas
        healthy_classes = [cls for cls in self.class_names if 'healthy' in cls.lower()]
        diseased_classes = [cls for cls in self.class_names if 'healthy' not in cls.lower()]
        
        healthy_count = sum(disease_counts.get(cls, 0) for cls in healthy_classes)
        diseased_count = sum(disease_counts.get(cls, 0) for cls in diseased_classes)
        
        analysis = {
            'metadata': {
                'image_name': image_name,
                'image_size': {
                    'width': image_shape[1],
                    'height': image_shape[0]
                },
                'timestamp': datetime.now().isoformat(),
                'tile_size': self.tile_size,
                'confidence_threshold': self.confidence_threshold
            },
            'summary': {
                'total_tiles_analyzed': total_tiles,
                'tiles_with_detections': diseased_tiles,
                'coverage_percentage': round(coverage_percentage, 2),
                'healthy_tiles': healthy_count,
                'diseased_tiles': diseased_count,
                'disease_ratio': round(diseased_count / max(diseased_tiles, 1), 2)
            },
            'disease_distribution': disease_counts,
            'detections': results,
            'recommendations': self._generate_recommendations(disease_counts, diseased_count, total_tiles)
        }
        
        return analysis
    
    def _generate_recommendations(
        self,
        disease_counts: Dict,
        diseased_count: int,
        total_tiles: int
    ) -> List[str]:
        """Generate treatment recommendations based on analysis."""
        recommendations = []
        
        # Calculate severity
        infection_rate = (diseased_count / total_tiles * 100) if total_tiles > 0 else 0
        
        if infection_rate < 5:
            recommendations.append("✓ Low infection rate. Monitor and spot treat affected areas.")
        elif infection_rate < 15:
            recommendations.append("⚠ Moderate infection rate. Consider targeted treatment of affected zones.")
        else:
            recommendations.append("⚠⚠ High infection rate. Immediate field-wide treatment recommended.")
        
        # Most common diseases
        if disease_counts:
            sorted_diseases = sorted(disease_counts.items(), key=lambda x: x[1], reverse=True)
            top_disease = sorted_diseases[0][0]
            
            if 'healthy' not in top_disease.lower():
                recommendations.append(f"Primary concern: {top_disease} ({disease_counts[top_disease]} detections)")
        
        # Priority zones
        if diseased_count > 0:
            recommendations.append("Generate heat map to identify priority treatment zones.")
            recommendations.append("Consider economic threshold before treatment.")
        
        return recommendations


def get_class_names_from_checkpoint(checkpoint_path: str) -> List[str]:
    """Extract class names from checkpoint file if available."""
    try:
        checkpoint = torch.load(checkpoint_path, map_location='cpu')
        if 'class_names' in checkpoint:
            return checkpoint['class_names']
    except:
        pass
    
    # Default class names (update based on your model)
    return [
        'Apple___Apple_scab', 'Apple___Black_rot', 'Apple___Cedar_apple_rust', 'Apple___healthy',
        'Grape___Black_rot', 'Grape___Esca_(Black_Measles)', 'Grape___healthy', 
        'Grape___Leaf_blight_(Isariopsis_Leaf_Spot)',
        'Orange___Haunglongbing_(Citrus_greening)',
        'Peach___Bacterial_spot', 'Peach___healthy',
        'Potato___Early_blight', 'Potato___healthy', 'Potato___Late_blight',
        'Squash___Powdery_mildew',
        'Strawberry___healthy', 'Strawberry___Leaf_scorch',
        'Tomato___Bacterial_spot', 'Tomato___Early_blight', 'Tomato___healthy',
        'Tomato___Late_blight', 'Tomato___Leaf_Mold', 'Tomato___Septoria_leaf_spot',
        'Tomato___Spider_mites Two-spotted_spider_mite', 'Tomato___Target_Spot',
        'Tomato___Tomato_mosaic_virus', 'Tomato___Tomato_Yellow_Leaf_Curl_Virus'
    ]


if __name__ == "__main__":
    # Example usage
    print("Drone Image Processor - Ready for initialization")
    print("Import this module and create a DroneImageProcessor instance to get started")

