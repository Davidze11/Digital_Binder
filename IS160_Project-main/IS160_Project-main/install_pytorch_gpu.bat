@echo off
echo ============================================================
echo Installing PyTorch with CUDA 12.1 Support
echo ============================================================
echo.
echo This will uninstall the CPU-only version and install GPU version
echo Your RTX 4050 Laptop GPU will be enabled for training!
echo.
pause

pip uninstall -y torch torchvision torchaudio

echo.
echo Installing PyTorch with CUDA 12.1 support...
echo.

pip install torch torchvision torchaudio --index-url https://download.pytorch.org/whl/cu121

echo.
echo ============================================================
echo Installation complete!
echo ============================================================
echo.
echo Testing GPU availability...
python -c "import torch; print('CUDA available:', torch.cuda.is_available()); print('GPU Name:', torch.cuda.get_device_name(0) if torch.cuda.is_available() else 'N/A')"

echo.
echo Press any key to exit...
pause > nul

