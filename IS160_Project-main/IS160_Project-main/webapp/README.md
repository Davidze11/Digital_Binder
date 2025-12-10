# 🌱 PlantAI - Plant Disease Detection Web Application

A professional, AI-powered web application for detecting plant diseases using deep learning. Built with Flask and PyTorch, featuring a modern, responsive user interface.

![PlantAI](https://img.shields.io/badge/AI-Powered-green)
![Flask](https://img.shields.io/badge/Flask-3.0.0-blue)
![PyTorch](https://img.shields.io/badge/PyTorch-2.1.0-red)

## 🎯 Features

- **Instant Disease Detection**: Upload a plant image and get results in under 1 second
- **High Accuracy**: 99.9% accuracy using ResNet34 deep learning architecture
- **27 Disease Classes**: Covers 8 different crop types (Apple, Grape, Orange, Peach, Potato, Squash, Strawberry, Tomato)
- **Comprehensive Information**: Detailed disease descriptions, symptoms, causes, treatment, and prevention
- **Top 3 Predictions**: View alternative diagnoses with confidence scores
- **Professional UI**: Modern, responsive design with smooth animations
- **Drag & Drop**: Easy image upload with drag-and-drop support
- **Privacy Focused**: Images are processed securely and not stored

## 📋 Prerequisites

Before running the application, ensure you have:

1. **Python 3.8 or higher** installed
2. **Trained Model**: The ResNet34 model must be trained first using the main project
3. **GPU Support** (optional but recommended): CUDA-enabled GPU for faster inference

## 🚀 Installation

### Step 1: Install Dependencies

Navigate to the webapp directory and install required packages:

```bash
cd webapp
pip install -r requirements.txt
```

### Step 2: Verify Model Checkpoint

Ensure the trained model exists at:
```
../models/checkpoints/resnet34_best.pth
```

If not, train the model first using the main project:
```bash
cd ..
python run_training.py
```

### Step 3: Verify Data Structure

The application needs access to the training data to load class names:
```
../data/train/
  ├── Apple___Apple_scab/
  ├── Apple___Black_rot/
  ├── Apple___Cedar_apple_rust/
  └── ... (other disease classes)
```

## 🎮 Usage

### Starting the Application

From the webapp directory, run:

```bash
python app.py
```

The application will start on `http://localhost:5000`

You should see:
```
[+] Loading model...
  Using device: cuda
  Number of classes: 27
[+] Model loaded successfully!
[+] Starting Flask server...
  Access the application at: http://localhost:5000
```

### Using the Web Interface

1. **Open Browser**: Navigate to `http://localhost:5000`
2. **Upload Image**: Click the upload area or drag & drop a plant image
3. **Analyze**: Click the "Analyze Plant" button
4. **View Results**: Review the diagnosis, confidence, and treatment recommendations
5. **Analyze Another**: Click "Analyze Another Plant" to start over

### Supported Image Formats

- **File Types**: JPG, JPEG, PNG
- **Maximum Size**: 16MB
- **Recommended**: Clear, well-lit photos of affected plant leaves

## 🏗️ Project Structure

```
webapp/
├── app.py                      # Flask application and model inference
├── disease_info.json           # Comprehensive disease information database
├── requirements.txt            # Python dependencies
├── README.md                   # This file
├── uploads/                    # Temporary upload directory (auto-created)
├── templates/
│   └── index.html             # Main web interface
├── static/
│   ├── css/
│   │   └── style.css          # Professional styling
│   ├── js/
│   │   └── script.js          # Frontend functionality
│   └── images/                # Static images (if any)
```

## 🔧 Configuration

### Changing Port

Edit `app.py` and modify the last line:

```python
app.run(debug=True, host='0.0.0.0', port=5000)  # Change port here
```

### Adjusting Upload Limits

Edit `app.py`:

```python
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB, adjust as needed
```

### Model Path

If your model is in a different location, update `app.py`:

```python
MODEL_PATH = Path('your/custom/path/to/model.pth')
```

## 🎨 Supported Crops & Diseases

### Apple (4 classes)
- Apple Scab
- Black Rot
- Cedar Apple Rust
- Healthy

### Grape (4 classes)
- Black Rot
- Esca (Black Measles)
- Leaf Blight (Isariopsis Leaf Spot)
- Healthy

### Orange (1 class)
- Huanglongbing (Citrus Greening)

### Peach (2 classes)
- Bacterial Spot
- Healthy

### Potato (3 classes)
- Early Blight
- Late Blight
- Healthy

### Squash (1 class)
- Powdery Mildew

### Strawberry (2 classes)
- Leaf Scorch
- Healthy

### Tomato (10 classes)
- Bacterial Spot
- Early Blight
- Late Blight
- Leaf Mold
- Septoria Leaf Spot
- Spider Mites (Two-spotted Spider Mite)
- Target Spot
- Tomato Mosaic Virus
- Tomato Yellow Leaf Curl Virus
- Healthy

## 🔍 API Endpoints

### Health Check
```
GET /health
```

Returns:
```json
{
  "status": "healthy",
  "model_loaded": true,
  "device": "cuda",
  "num_classes": 27
}
```

### Predict Disease
```
POST /predict
Content-Type: multipart/form-data
```

Parameters:
- `file`: Image file (JPG, JPEG, PNG)

Returns:
```json
{
  "success": true,
  "prediction": {
    "plant": "Tomato",
    "disease": "Early blight",
    "is_healthy": false,
    "confidence": 0.9876,
    "class_name": "Tomato___Early_blight",
    "top_predictions": [...]
  },
  "details": {
    "description": "...",
    "symptoms": [...],
    "causes": [...],
    "treatment": [...],
    "prevention": [...]
  }
}
```

## 🛠️ Troubleshooting

### Model Not Found Error

**Error**: `Model checkpoint not found`

**Solution**: Train the model first using the main project:
```bash
cd ..
python run_training.py
```

### CUDA Out of Memory

**Error**: `CUDA out of memory`

**Solution**: The model will automatically fall back to CPU. For GPU usage, try:
- Closing other GPU-intensive applications
- Reducing batch size (not applicable for single-image inference)
- Using a GPU with more memory

### Import Errors

**Error**: `ModuleNotFoundError`

**Solution**: Install all dependencies:
```bash
pip install -r requirements.txt
```

### Port Already in Use

**Error**: `Address already in use`

**Solution**: Change the port in `app.py` or kill the process using port 5000:
```bash
# Windows
netstat -ano | findstr :5000
taskkill /PID <PID> /F

# Linux/Mac
lsof -ti:5000 | xargs kill -9
```

## 📊 Model Performance

- **Architecture**: ResNet34 (pretrained on ImageNet)
- **Training Images**: 50,000+ images
- **Validation Accuracy**: 99.9%
- **Inference Time**: < 1 second on GPU, 2-3 seconds on CPU
- **Model Size**: ~85 MB

## 🔒 Security & Privacy

- Images are processed in real-time and not permanently stored
- No user data is collected or transmitted
- All processing happens locally on your server
- HTTPS recommended for production deployment

## 🚀 Production Deployment

For production use:

1. **Disable Debug Mode**:
   ```python
   app.run(debug=False, host='0.0.0.0', port=5000)
   ```

2. **Use Production Server** (e.g., Gunicorn):
   ```bash
   pip install gunicorn
   gunicorn -w 4 -b 0.0.0.0:5000 app:app
   ```

3. **Set Up HTTPS** with a reverse proxy (nginx, Apache)

4. **Add Authentication** if needed for restricted access

5. **Monitor Resources** for GPU/CPU usage and memory

## 🤝 Contributing

This web application is part of the IS160 Plant Disease Detection project. For improvements or bug reports:

1. Document the issue or enhancement
2. Test thoroughly
3. Ensure backward compatibility
4. Update documentation

## 📝 License

This project is created for educational purposes as part of IS160 coursework.

## 🙏 Acknowledgments

- **Model Architecture**: ResNet34 from torchvision
- **Dataset**: PlantVillage Dataset
- **Framework**: PyTorch, Flask
- **UI Design**: Custom responsive design with modern CSS

## 📧 Support

For issues or questions:
- Check the troubleshooting section
- Review the main project documentation
- Contact the development team

---

**Built with ❤️ for healthier crops and better yields**

