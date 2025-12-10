# 🚀 Quick Start Guide - PlantAI Web Application

## Installation Steps

### Option 1: Automatic Setup (Windows)

1. **Double-click `start.bat`** - This will automatically:
   - Check Python installation
   - Install all dependencies
   - Start the web server

### Option 2: Automatic Setup (Linux/Mac)

1. **Make the script executable:**
   ```bash
   chmod +x start.sh
   ```

2. **Run the script:**
   ```bash
   ./start.sh
   ```

### Option 3: Manual Setup

1. **Navigate to the webapp directory:**
   ```bash
   cd webapp
   ```

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```
   
   Or if you're using pip3:
   ```bash
   pip3 install -r requirements.txt
   ```

3. **Start the application:**
   ```bash
   python app.py
   ```
   
   Or if you're using python3:
   ```bash
   python3 app.py
   ```

4. **Open your browser:**
   Navigate to `http://localhost:5000`

## ⚠️ Important Prerequisites

### 1. Trained Model Required

Before running the web app, you MUST have a trained model. The app expects the model at:
```
../models/checkpoints/resnet34_best.pth
```

If you don't have this file, train the model first:
```bash
cd ..  # Go to project root
python run_training.py
```

### 2. Data Directory Required

The app needs access to the training data structure to load class names:
```
../data/train/
  ├── Apple___Apple_scab/
  ├── Apple___Black_rot/
  └── ... (other classes)
```

If you don't have this, run data preparation first:
```bash
cd ..  # Go to project root
python src/data_preparation.py
```

## 🎯 First Time Setup Checklist

- [ ] Python 3.8+ installed
- [ ] PyTorch installed (with CUDA if you have a GPU)
- [ ] Training data prepared in `data/train/` directory
- [ ] Model trained and saved in `models/checkpoints/resnet34_best.pth`
- [ ] Web app dependencies installed (`pip install -r requirements.txt`)

## 🧪 Testing the Installation

1. **Test dependencies:**
   ```bash
   python -c "import flask; import torch; import torchvision; from PIL import Image; print('✓ All dependencies installed!')"
   ```

2. **Test model existence:**
   ```bash
   # Windows
   if exist "..\models\checkpoints\resnet34_best.pth" echo ✓ Model found!
   
   # Linux/Mac
   [ -f "../models/checkpoints/resnet34_best.pth" ] && echo "✓ Model found!"
   ```

3. **Test data directory:**
   ```bash
   # Windows
   if exist "..\data\train" echo ✓ Training data found!
   
   # Linux/Mac
   [ -d "../data/train" ] && echo "✓ Training data found!"
   ```

## 🐛 Common Issues

### Issue: "No module named 'flask'"
**Solution:** Install dependencies:
```bash
pip install -r requirements.txt
```

### Issue: "Model checkpoint not found"
**Solution:** Train the model first:
```bash
cd ..
python run_training.py
```

### Issue: "Training directory not found"
**Solution:** Prepare the data first:
```bash
cd ..
python src/data_preparation.py
```

### Issue: "CUDA out of memory"
**Solution:** The app will automatically use CPU. This is slower but works fine.

### Issue: Port 5000 already in use
**Solution:** Either:
- Kill the process using port 5000
- Edit `app.py` and change the port number in the last line

## 📱 Using the Application

1. **Open browser** → `http://localhost:5000`
2. **Upload image** → Click upload area or drag & drop
3. **Analyze** → Click "Analyze Plant" button
4. **View results** → See diagnosis, confidence, and treatment recommendations

## 🎨 Supported Image Types

- **Formats:** JPG, JPEG, PNG
- **Max Size:** 16MB
- **Best Results:** Clear, well-lit photos of plant leaves

## 💡 Tips for Best Results

1. Take photos in good lighting
2. Focus on affected leaves
3. Ensure the leaf fills most of the frame
4. Avoid blurry or dark images
5. Use images similar to the training data

## 🔧 Advanced Configuration

### Change Port
Edit `app.py`, line ~250:
```python
app.run(debug=True, host='0.0.0.0', port=5000)  # Change 5000 to your port
```

### Enable Production Mode
Edit `app.py`, line ~250:
```python
app.run(debug=False, host='0.0.0.0', port=5000)
```

### Change Upload Size Limit
Edit `app.py`, line ~30:
```python
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # Change to your limit
```

## 📞 Need Help?

1. Check the main `README.md` for detailed information
2. Review the troubleshooting section
3. Ensure all prerequisites are met
4. Check that the model is trained and accessible

---

**Ready to protect your plants? Let's get started! 🌱**

