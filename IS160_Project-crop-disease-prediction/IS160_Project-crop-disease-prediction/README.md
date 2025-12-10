# AI-Based Crop Disease Prediction for Fresno County

## Project Overview

This project aims to build an AI model that can predict which crops show signs of disease from leaf images, focusing on agricultural species that are relevant to Fresno County. This is an approach to the development of plant disease recognition based on leaf image classification using deep convolutional networks.

Once diseased plants are flagged, the model can theoretically help "pull out" or quarantine those crops, preventing further spread. Although we will not deploy this in the field, the work will demonstrate how AI could be used as an early-warning system in precision agriculture.

## Research Questions

1. **Can we train a model to reliably distinguish between healthy and diseased plants** for Fresno-grown species (tomato, grape, peach, potato, squash, strawberry, apple, orange, bell pepper, corn, cherry)?

2. **Which visual characteristics** (such as discoloration, spotting, or irregular texture) are most indicative of specific crop diseases within Fresno County's major crops?

3. **How effective is the model** in identifying the "bad crops" under simulated conditions, and what are its limitations (false positives, false negatives)?

## Dataset

We use the **Plant Disease Dataset** from Kaggle, which includes 38 different types of plant diseases across 14 different plants.

**Dataset Link**: [Plant Disease PyTorch VGG16 and ResNet34](https://www.kaggle.com/code/aniketkolte04/plant-disease-pytorch-vgg16-and-resnet34)

### Fresno-Relevant Crops Included

The following crops overlap with those grown in Fresno County:

- **Tomato**: Bacterial Spot, Early Blight, Late Blight, Leaf Mold, Septoria Leaf Spot, Spider Mites, Target Spot, Mosaic Virus, Yellow Leaf Curl Virus, Healthy
- **Apple**: Apple Scab, Black Rot, Cedar Apple Rust, Healthy
- **Grape**: Black Rot, Esca (Black Measles), Leaf Blight (Isariopsis Leaf Spot), Healthy
- **Peach**: Bacterial Spot, Healthy
- **Potato**: Early Blight, Late Blight, Healthy
- **Squash**: Powdery Mildew, Healthy
- **Strawberry**: Leaf Scorch, Healthy
- **Orange**: Huanglongbing (Citrus Greening), Healthy
- **Cherry**: Powdery Mildew, Healthy
- **Pepper**: Bell Pepper (Bacterial Spot), Healthy
- **Corn (Maize)**: Northern Leaf Blight, Cercospora Leaf Spot, Gray Leaf Spot, Common Rust, Healthy

**Note**: Crops not relevant to Fresno (soybean, raspberry, blueberry) will be excluded during data preparation.

## Methodology

We will use **Convolutional Neural Networks (CNN)** or variants (e.g., ResNet34, VGG16) because of their strong performance in image classification and pattern recognition tasks.

### AI Pipeline Workflow

1. **Data Preparation** – Filter the dataset to include only classes for Fresno-relevant crops; clean, balance (or augment) classes.
2. **Exploratory Data Analysis (EDA)** – Visualize disease vs. healthy samples, class imbalances, feature distributions.
3. **Model Training** – Choose architectures (e.g., CNN, transfer learning with ResNet) and train to classify images into healthy vs. diseased or into specific disease classes.
4. **Model Evaluation** – Use metrics like accuracy, precision, recall, F1, and confusion matrices to assess performance.
5. **Simulation of Crop Pulling** – Use model predictions to simulate identifying "bad crops" and measure error rates and potential yield loss avoided.

## Project Structure

```
IS160_Project/
├── data/
│   ├── raw/              # Original dataset (download here)
│   ├── processed/        # Filtered and preprocessed data
│   └── augmented/        # Augmented training data
├── models/
│   └── checkpoints/      # Saved model weights
├── notebooks/
│   ├── 01_eda.ipynb                    # Exploratory Data Analysis
│   └── 02_model_experiments.ipynb      # Model experimentation
├── src/
│   ├── data_preparation.py   # Data filtering and preprocessing
│   ├── train_model.py        # Model training script
│   ├── evaluate_model.py     # Model evaluation script
│   ├── utils.py              # Utility functions
│   └── config.py             # Configuration settings
├── results/
│   ├── figures/          # Plots and visualizations
│   └── metrics/          # Evaluation metrics and reports
├── docs/                 # Additional documentation
├── .gitignore
├── requirements.txt
└── README.md
```

## Setup Instructions

### 1. Clone the Repository

```bash
git clone <repository-url>
cd IS160_Project
```

### 2. Create Virtual Environment

```bash
# Using venv
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Or using conda
conda create -n crop-disease python=3.10
conda activate crop-disease
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

### 4. Download Dataset

1. Download the Plant Disease dataset from [Kaggle](https://www.kaggle.com/datasets/vipoooool/new-plant-diseases-dataset)
2. Extract and place it in `data/raw/`

### 5. Prepare Data

```bash
python src/data_preparation.py
```

## Usage

### Exploratory Data Analysis

```bash
jupyter notebook notebooks/01_eda.ipynb
```

### Train Model

```bash
python src/train_model.py --model resnet34 --epochs 50 --batch-size 32
```

### Evaluate Model

```bash
python src/evaluate_model.py --model-path models/checkpoints/best_model.pth
```

## Expected Outcomes

1. A working classification model that can flag diseased plants among Fresno-relevant crops
2. Performance metrics (accuracy, precision, recall, F1-score)
3. Confusion matrices showing model strengths and weaknesses
4. Analysis of which disease types are harder to detect
5. Tradeoff analysis between false positives vs. false negatives
6. Simulation results demonstrating potential yield loss prevention

## Technologies Used

- **Python 3.10+**
- **PyTorch** – Deep learning framework
- **torchvision** – Pre-trained models and image transformations
- **NumPy & Pandas** – Data manipulation
- **Matplotlib & Seaborn** – Data visualization
- **scikit-learn** – Evaluation metrics
- **OpenCV & Pillow** – Image processing

## Team Members

*Add team member names and contributions here*

## License

*Add license information here*

## Acknowledgments

- Plant Disease Dataset from Kaggle
- Fresno County Agricultural Commissioner's Office for crop information
- IS 160 Course Instructors and TAs

