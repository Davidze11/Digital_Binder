"""
Flask web application for plant disease detection.
Provides a user-friendly interface for uploading plant images and detecting diseases.
"""

import os
import json
from pathlib import Path
from datetime import datetime
import io

import torch
import torch.nn as nn
from torchvision import transforms, models
from PIL import Image
from flask import Flask, render_template, request, jsonify
from werkzeug.utils import secure_filename

# Configuration
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['UPLOAD_FOLDER'] = Path('uploads')
app.config['UPLOAD_FOLDER'].mkdir(exist_ok=True)

# Allowed image extensions
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'JPG', 'PNG', 'JPEG'}

# Model configuration
PROJECT_ROOT = Path(__file__).parent.parent
MODEL_PATH = PROJECT_ROOT / 'models' / 'checkpoints' / 'resnet34_best.pth'
DISEASE_INFO_PATH = Path(__file__).parent / 'disease_info.json'

# Image preprocessing configuration
IMG_SIZE = 224
NORMALIZE_MEAN = [0.485, 0.456, 0.406]
NORMALIZE_STD = [0.229, 0.224, 0.225]

# Global variables for model
model = None
device = None
class_names = None
disease_info = None


def allowed_file(filename):
    """Check if file extension is allowed."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def load_disease_info():
    """Load disease information and treatment suggestions."""
    global disease_info
    
    with open(DISEASE_INFO_PATH, 'r', encoding='utf-8') as f:
        disease_info = json.load(f)
    
    print(f"[+] Loaded information for {len(disease_info)} disease categories")


def get_class_names():
    """Get class names from the training dataset structure."""
    train_dir = PROJECT_ROOT / 'data' / 'train'
    
    if not train_dir.exists():
        raise FileNotFoundError(f"Training directory not found: {train_dir}")
    
    # Get sorted list of class names
    classes = sorted([d.name for d in train_dir.iterdir() if d.is_dir()])
    return classes


def load_model():
    """Load the trained ResNet34 model."""
    global model, device, class_names
    
    print("\n[+] Loading model...")
    
    # Get device
    device = torch.device('cuda' if torch.cuda.is_available() else 'cpu')
    print(f"  Using device: {device}")
    
    # Get class names
    class_names = get_class_names()
    num_classes = len(class_names)
    print(f"  Number of classes: {num_classes}")
    
    # Create model architecture
    model = models.resnet34(pretrained=False)
    num_features = model.fc.in_features
    model.fc = nn.Linear(num_features, num_classes)
    
    # Load trained weights
    if not MODEL_PATH.exists():
        raise FileNotFoundError(
            f"Model checkpoint not found: {MODEL_PATH}\n"
            "Please train the model first using train_model.py"
        )
    
    checkpoint = torch.load(MODEL_PATH, map_location=device)
    model.load_state_dict(checkpoint['model_state_dict'])
    model = model.to(device)
    model.eval()
    
    print(f"[+] Model loaded successfully!")
    print(f"  Checkpoint epoch: {checkpoint.get('epoch', 'unknown')}")
    print(f"  Validation accuracy: {checkpoint.get('accuracy', 'unknown'):.4f}")
    
    return model


def get_image_transforms():
    """Get image preprocessing transforms."""
    return transforms.Compose([
        transforms.Resize((IMG_SIZE, IMG_SIZE)),
        transforms.ToTensor(),
        transforms.Normalize(mean=NORMALIZE_MEAN, std=NORMALIZE_STD),
    ])


def preprocess_image(image_file):
    """
    Preprocess uploaded image for model inference.
    
    Args:
        image_file: File object from Flask request
        
    Returns:
        Preprocessed image tensor
    """
    # Load image
    image = Image.open(image_file).convert('RGB')
    
    # Apply transforms
    transform = get_image_transforms()
    image_tensor = transform(image).unsqueeze(0)  # Add batch dimension
    
    return image_tensor, image


def predict_disease(image_tensor):
    """
    Predict plant disease from preprocessed image.
    
    Args:
        image_tensor: Preprocessed image tensor
        
    Returns:
        Dictionary containing prediction results
    """
    global model, device, class_names
    
    # Move to device
    image_tensor = image_tensor.to(device)
    
    # Inference
    with torch.no_grad():
        outputs = model(image_tensor)
        probabilities = torch.nn.functional.softmax(outputs, dim=1)
        confidence, predicted = torch.max(probabilities, 1)
        
        # Get top 3 predictions
        top_probs, top_indices = torch.topk(probabilities, k=min(3, len(class_names)))
        top_probs = top_probs[0].cpu().numpy()
        top_indices = top_indices[0].cpu().numpy()
    
    # Get prediction details
    predicted_class = class_names[predicted.item()]
    confidence_score = confidence.item()
    
    # Parse class name
    plant, disease = predicted_class.split('___')
    plant = plant.replace('_', ' ')
    disease = disease.replace('_', ' ')
    
    # Get top 3 predictions
    top_predictions = []
    for i, (prob, idx) in enumerate(zip(top_probs, top_indices)):
        cls = class_names[idx]
        p, d = cls.split('___')
        top_predictions.append({
            'rank': i + 1,
            'plant': p.replace('_', ' '),
            'disease': d.replace('_', ' '),
            'confidence': float(prob),
            'class_name': cls
        })
    
    result = {
        'plant': plant,
        'disease': disease,
        'is_healthy': disease.lower() == 'healthy',
        'confidence': confidence_score,
        'class_name': predicted_class,
        'top_predictions': top_predictions,
        'timestamp': datetime.now().isoformat()
    }
    
    return result


def get_disease_details(class_name):
    """
    Get detailed information about the detected disease.
    
    Args:
        class_name: Predicted class name
        
    Returns:
        Dictionary containing disease information
    """
    global disease_info
    
    if class_name in disease_info:
        return disease_info[class_name]
    else:
        # Default info if not found
        return {
            'description': 'No detailed information available for this condition.',
            'symptoms': ['Information not available'],
            'causes': ['Information not available'],
            'treatment': ['Consult with a local agricultural extension service for guidance.'],
            'prevention': ['Maintain good plant health practices']
        }


@app.route('/')
def index():
    """Render the main page."""
    return render_template('index.html')


@app.route('/predict', methods=['POST'])
def predict():
    """
    Handle image upload and return prediction.
    
    Returns:
        JSON response with prediction results
    """
    try:
        # Check if file is present
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        
        # Check if file is selected
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        # Check file extension
        if not allowed_file(file.filename):
            return jsonify({'error': 'Invalid file type. Please upload a JPG, JPEG, or PNG image.'}), 400
        
        # Save uploaded file
        filename = secure_filename(file.filename)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"{timestamp}_{filename}"
        filepath = app.config['UPLOAD_FOLDER'] / filename
        file.save(filepath)
        
        # Preprocess image
        with open(filepath, 'rb') as f:
            image_tensor, original_image = preprocess_image(f)
        
        # Make prediction
        prediction = predict_disease(image_tensor)
        
        # Get disease details
        disease_details = get_disease_details(prediction['class_name'])
        
        # Combine results
        result = {
            'success': True,
            'prediction': prediction,
            'details': disease_details,
            'filename': filename
        }
        
        return jsonify(result)
    
    except Exception as e:
        print(f"[!] Error during prediction: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'Prediction failed: {str(e)}'}), 500


@app.route('/health', methods=['GET'])
def health_check():
    """Health check endpoint."""
    return jsonify({
        'status': 'healthy',
        'model_loaded': model is not None,
        'device': str(device),
        'num_classes': len(class_names) if class_names else 0
    })


if __name__ == '__main__':
    print("\n" + "="*70)
    print("PLANT DISEASE DETECTION WEB APPLICATION")
    print("="*70)
    
    # Load disease information
    load_disease_info()
    
    # Load model
    load_model()
    
    print("\n[+] Starting Flask server...")
    print("  Access the application at: http://localhost:5000")
    print("="*70 + "\n")
    
    # Run Flask app
    app.run(debug=True, host='0.0.0.0', port=5000)

