"""
Script to verify the drone processing system is properly installed
Run this to check all dependencies and configurations
"""

import sys
from pathlib import Path


def check_python_version():
    """Check Python version"""
    version = sys.version_info
    print(f"✓ Python {version.major}.{version.minor}.{version.micro}")
    if version.major < 3 or (version.major == 3 and version.minor < 8):
        print("  ⚠ Warning: Python 3.8+ recommended")
        return False
    return True


def check_dependencies():
    """Check if all required packages are installed"""
    required = {
        'torch': 'PyTorch',
        'torchvision': 'TorchVision',
        'cv2': 'OpenCV',
        'PIL': 'Pillow',
        'numpy': 'NumPy',
        'matplotlib': 'Matplotlib',
        'seaborn': 'Seaborn',
        'sklearn': 'scikit-learn',
        'scipy': 'SciPy',
        'tqdm': 'tqdm'
    }
    
    print("\nChecking dependencies...")
    all_good = True
    
    for module, name in required.items():
        try:
            __import__(module)
            print(f"✓ {name}")
        except ImportError:
            print(f"✗ {name} - NOT INSTALLED")
            all_good = False
    
    return all_good


def check_cuda():
    """Check CUDA availability"""
    try:
        import torch
        print(f"\nCUDA available: {torch.cuda.is_available()}")
        if torch.cuda.is_available():
            print(f"✓ CUDA device: {torch.cuda.get_device_name(0)}")
            print(f"  CUDA version: {torch.version.cuda}")
        else:
            print("  ℹ Running on CPU (slower, but works)")
        return True
    except Exception as e:
        print(f"✗ Error checking CUDA: {e}")
        return False


def check_model():
    """Check if trained model exists"""
    model_paths = [
        Path('models/checkpoints/resnet34_best.pth'),
        Path('models/checkpoints/resnet34_latest.pth'),
    ]
    
    print("\nChecking for trained model...")
    for model_path in model_paths:
        if model_path.exists():
            size_mb = model_path.stat().st_size / (1024 * 1024)
            print(f"✓ Model found: {model_path} ({size_mb:.1f} MB)")
            return True
    
    print("✗ No trained model found")
    print("  Run: python run_training.py --epochs 30")
    return False


def check_drone_modules():
    """Check if drone processing modules are present"""
    modules = [
        Path('src/drone_processor.py'),
        Path('src/drone_visualizer.py'),
        Path('process_drone_image.py'),
        Path('drone_examples/demo.py')
    ]
    
    print("\nChecking drone processing modules...")
    all_good = True
    
    for module in modules:
        if module.exists():
            print(f"✓ {module}")
        else:
            print(f"✗ {module} - MISSING")
            all_good = False
    
    return all_good


def check_directories():
    """Check/create required directories"""
    dirs = [
        Path('results/drone_analysis'),
        Path('drone_examples'),
        Path('models/checkpoints')
    ]
    
    print("\nChecking directories...")
    for dir_path in dirs:
        dir_path.mkdir(parents=True, exist_ok=True)
        print(f"✓ {dir_path}")
    
    return True


def run_verification():
    """Run all verification checks"""
    print("="*80)
    print("DRONE PROCESSING SYSTEM - VERIFICATION")
    print("="*80)
    
    checks = [
        ("Python Version", check_python_version),
        ("Dependencies", check_dependencies),
        ("CUDA/GPU", check_cuda),
        ("Trained Model", check_model),
        ("Drone Modules", check_drone_modules),
        ("Directories", check_directories)
    ]
    
    results = {}
    for name, check_func in checks:
        try:
            results[name] = check_func()
        except Exception as e:
            print(f"\n✗ Error in {name}: {e}")
            results[name] = False
    
    # Summary
    print("\n" + "="*80)
    print("VERIFICATION SUMMARY")
    print("="*80)
    
    all_passed = all(results.values())
    
    if all_passed:
        print("\n✅ All checks passed! System is ready to use.")
        print("\nNext steps:")
        print("  1. Run demo: python drone_examples/demo.py")
        print("  2. Process your image: python process_drone_image.py --image YOUR_IMAGE.jpg --visualize")
    else:
        print("\n⚠ Some checks failed. Please fix the issues above.")
        
        if not results.get("Dependencies"):
            print("\nTo install missing dependencies:")
            print("  pip install -r requirements.txt")
            print("  pip install -r drone_requirements.txt")
        
        if not results.get("Trained Model"):
            print("\nTo train a model:")
            print("  python run_training.py --epochs 30 --architecture resnet34")
    
    print("="*80)
    
    return all_passed


if __name__ == "__main__":
    try:
        run_verification()
    except Exception as e:
        print(f"\n✗ Verification failed with error: {e}")
        import traceback
        traceback.print_exc()

