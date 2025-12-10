// DOM Elements
const fileInput = document.getElementById('fileInput');
const uploadArea = document.getElementById('uploadArea');
const uploadContent = document.getElementById('uploadContent');
const imagePreview = document.getElementById('imagePreview');
const previewImg = document.getElementById('previewImg');
const removeImage = document.getElementById('removeImage');
const analyzeBtn = document.getElementById('analyzeBtn');
const analyzeBtnText = document.getElementById('analyzeBtnText');
const loadingContainer = document.getElementById('loadingContainer');
const resultsContainer = document.getElementById('resultsContainer');
const analyzeAnother = document.getElementById('analyzeAnother');

// State
let selectedFile = null;

// Initialize
document.addEventListener('DOMContentLoaded', () => {
    setupEventListeners();
    setupDragAndDrop();
});

// Event Listeners
function setupEventListeners() {
    fileInput.addEventListener('change', handleFileSelect);
    removeImage.addEventListener('click', handleRemoveImage);
    analyzeBtn.addEventListener('click', handleAnalyze);
    analyzeAnother.addEventListener('click', handleAnalyzeAnother);
    
    // Click on upload area triggers file input
    uploadArea.addEventListener('click', (e) => {
        if (e.target !== removeImage && !removeImage.contains(e.target)) {
            if (!imagePreview.classList.contains('active')) {
                fileInput.click();
            }
        }
    });
}

// Drag and Drop
function setupDragAndDrop() {
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        uploadArea.addEventListener(eventName, preventDefaults);
        document.body.addEventListener(eventName, preventDefaults);
    });
    
    ['dragenter', 'dragover'].forEach(eventName => {
        uploadArea.addEventListener(eventName, () => {
            uploadArea.classList.add('drag-over');
        });
    });
    
    ['dragleave', 'drop'].forEach(eventName => {
        uploadArea.addEventListener(eventName, () => {
            uploadArea.classList.remove('drag-over');
        });
    });
    
    uploadArea.addEventListener('drop', handleDrop);
}

function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
}

function handleDrop(e) {
    const dt = e.dataTransfer;
    const files = dt.files;
    
    if (files.length > 0) {
        handleFile(files[0]);
    }
}

function handleFileSelect(e) {
    const files = e.target.files;
    if (files.length > 0) {
        handleFile(files[0]);
    }
}

function handleFile(file) {
    // Validate file type
    const validTypes = ['image/jpeg', 'image/jpg', 'image/png'];
    if (!validTypes.includes(file.type)) {
        showError('Please upload a valid image file (JPG, JPEG, or PNG)');
        return;
    }
    
    // Validate file size (16MB max)
    const maxSize = 16 * 1024 * 1024;
    if (file.size > maxSize) {
        showError('File size must be less than 16MB');
        return;
    }
    
    selectedFile = file;
    displayImagePreview(file);
    analyzeBtn.disabled = false;
}

function displayImagePreview(file) {
    const reader = new FileReader();
    
    reader.onload = (e) => {
        previewImg.src = e.target.result;
        uploadContent.style.display = 'none';
        imagePreview.classList.add('active');
        uploadArea.style.cursor = 'default';
    };
    
    reader.readAsDataURL(file);
}

function handleRemoveImage(e) {
    e.stopPropagation();
    selectedFile = null;
    fileInput.value = '';
    previewImg.src = '';
    uploadContent.style.display = 'block';
    imagePreview.classList.remove('active');
    analyzeBtn.disabled = true;
    uploadArea.style.cursor = 'pointer';
}

async function handleAnalyze() {
    if (!selectedFile) {
        showError('Please select an image first');
        return;
    }
    
    // Show loading state
    loadingContainer.classList.add('active');
    resultsContainer.classList.remove('active');
    analyzeBtn.disabled = true;
    analyzeBtnText.textContent = 'Analyzing...';
    
    // Scroll to loading container
    loadingContainer.scrollIntoView({ behavior: 'smooth', block: 'center' });
    
    // Prepare form data
    const formData = new FormData();
    formData.append('file', selectedFile);
    
    try {
        // Send request to backend
        const response = await fetch('/predict', {
            method: 'POST',
            body: formData
        });
        
        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.error || 'Prediction failed');
        }
        
        const result = await response.json();
        
        // Hide loading, show results
        setTimeout(() => {
            loadingContainer.classList.remove('active');
            displayResults(result);
        }, 500); // Small delay for better UX
        
    } catch (error) {
        console.error('Error:', error);
        loadingContainer.classList.remove('active');
        showError(error.message || 'An error occurred during analysis. Please try again.');
        analyzeBtn.disabled = false;
        analyzeBtnText.textContent = 'Analyze Plant';
    }
}

function displayResults(result) {
    const { prediction, details } = result;
    
    // Update result header
    const resultIcon = document.getElementById('resultIcon');
    const resultTitle = document.getElementById('resultTitle');
    const resultSubtitle = document.getElementById('resultSubtitle');
    
    if (prediction.is_healthy) {
        resultIcon.innerHTML = `
            <svg viewBox="0 0 20 20" fill="currentColor">
                <path fill-rule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z"/>
            </svg>
        `;
        resultIcon.className = 'result-icon healthy';
        resultTitle.textContent = 'Healthy Plant Detected!';
        resultSubtitle.textContent = 'Your plant appears to be in good health';
    } else {
        resultIcon.innerHTML = `
            <svg viewBox="0 0 20 20" fill="currentColor">
                <path fill-rule="evenodd" d="M8.257 3.099c.765-1.36 2.722-1.36 3.486 0l5.58 9.92c.75 1.334-.213 2.98-1.742 2.98H4.42c-1.53 0-2.493-1.646-1.743-2.98l5.58-9.92zM11 13a1 1 0 11-2 0 1 1 0 012 0zm-1-8a1 1 0 00-1 1v3a1 1 0 002 0V6a1 1 0 00-1-1z"/>
            </svg>
        `;
        resultIcon.className = 'result-icon diseased';
        resultTitle.textContent = 'Disease Detected';
        resultSubtitle.textContent = 'We recommend taking action to treat this condition';
    }
    
    // Update result info
    document.getElementById('plantType').textContent = prediction.plant;
    document.getElementById('diseaseType').textContent = prediction.disease;
    
    const confidencePercent = (prediction.confidence * 100).toFixed(1);
    document.getElementById('confidenceFill').style.width = `${confidencePercent}%`;
    document.getElementById('confidenceText').textContent = `${confidencePercent}%`;
    
    // Update alternative predictions
    displayAlternativePredictions(prediction.top_predictions);
    
    // Update disease details
    displayDiseaseDetails(details);
    
    // Show results with animation
    resultsContainer.classList.add('active', 'fade-in');
    resultsContainer.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function displayAlternativePredictions(predictions) {
    const alternativesList = document.getElementById('alternativesList');
    
    // Skip the first prediction (main prediction) and show next 2
    const alternatives = predictions.slice(1, 3);
    
    if (alternatives.length === 0) {
        document.getElementById('alternativePredictions').style.display = 'none';
        return;
    }
    
    document.getElementById('alternativePredictions').style.display = 'block';
    
    alternativesList.innerHTML = alternatives.map(pred => `
        <div class="alternative-item">
            <div style="display: flex; align-items: center; flex: 1;">
                <div class="alternative-rank">${pred.rank}</div>
                <div class="alternative-info">
                    <div class="alternative-plant">${pred.plant}</div>
                    <div class="alternative-disease">${pred.disease}</div>
                </div>
            </div>
            <div class="alternative-confidence">${(pred.confidence * 100).toFixed(1)}%</div>
        </div>
    `).join('');
}

function displayDiseaseDetails(details) {
    const diseaseDetails = document.getElementById('diseaseDetails');
    
    if (!details) {
        diseaseDetails.style.display = 'none';
        return;
    }
    
    diseaseDetails.style.display = 'block';
    
    // Description
    document.getElementById('descriptionText').textContent = details.description;
    
    // Symptoms
    const symptomsList = document.getElementById('symptomsList');
    symptomsList.innerHTML = details.symptoms.map(symptom => 
        `<li>${symptom}</li>`
    ).join('');
    
    // Causes
    const causesList = document.getElementById('causesList');
    causesList.innerHTML = details.causes.map(cause => 
        `<li>${cause}</li>`
    ).join('');
    
    // Treatment
    const treatmentList = document.getElementById('treatmentList');
    treatmentList.innerHTML = details.treatment.map(treatment => 
        `<li>${treatment}</li>`
    ).join('');
    
    // Prevention
    const preventionList = document.getElementById('preventionList');
    preventionList.innerHTML = details.prevention.map(prevention => 
        `<li>${prevention}</li>`
    ).join('');
}

function handleAnalyzeAnother() {
    // Reset everything
    handleRemoveImage(new Event('click'));
    resultsContainer.classList.remove('active');
    
    // Scroll back to upload section
    document.querySelector('.upload-card').scrollIntoView({ 
        behavior: 'smooth', 
        block: 'center' 
    });
    
    // Reset analyze button
    analyzeBtn.disabled = true;
    analyzeBtnText.textContent = 'Analyze Plant';
}

function showError(message) {
    // Create error toast
    const toast = document.createElement('div');
    toast.className = 'error-toast';
    toast.innerHTML = `
        <svg viewBox="0 0 20 20" fill="currentColor" style="width: 20px; height: 20px; flex-shrink: 0;">
            <path fill-rule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.707 7.293a1 1 0 00-1.414 1.414L8.586 10l-1.293 1.293a1 1 0 101.414 1.414L10 11.414l1.293 1.293a1 1 0 001.414-1.414L11.414 10l1.293-1.293a1 1 0 00-1.414-1.414L10 8.586 8.707 7.293z"/>
        </svg>
        <span>${message}</span>
    `;
    
    // Add styles if not already present
    if (!document.querySelector('#error-toast-styles')) {
        const style = document.createElement('style');
        style.id = 'error-toast-styles';
        style.textContent = `
            .error-toast {
                position: fixed;
                top: 20px;
                right: 20px;
                background: #fee2e2;
                color: #991b1b;
                padding: 1rem 1.5rem;
                border-radius: 0.5rem;
                box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);
                display: flex;
                align-items: center;
                gap: 0.75rem;
                max-width: 400px;
                z-index: 9999;
                animation: slideInRight 0.3s ease-out;
                border: 1px solid #fca5a5;
            }
            
            @keyframes slideInRight {
                from {
                    transform: translateX(100%);
                    opacity: 0;
                }
                to {
                    transform: translateX(0);
                    opacity: 1;
                }
            }
            
            @keyframes slideOutRight {
                from {
                    transform: translateX(0);
                    opacity: 1;
                }
                to {
                    transform: translateX(100%);
                    opacity: 0;
                }
            }
        `;
        document.head.appendChild(style);
    }
    
    document.body.appendChild(toast);
    
    // Remove after 5 seconds
    setTimeout(() => {
        toast.style.animation = 'slideOutRight 0.3s ease-out';
        setTimeout(() => {
            document.body.removeChild(toast);
        }, 300);
    }, 5000);
}

// Smooth scroll for nav links
document.querySelectorAll('a[href^="#"]').forEach(anchor => {
    anchor.addEventListener('click', function (e) {
        e.preventDefault();
        const target = document.querySelector(this.getAttribute('href'));
        if (target) {
            target.scrollIntoView({
                behavior: 'smooth',
                block: 'start'
            });
        }
    });
});

// Add scroll-to-top button behavior
let lastScrollTop = 0;
window.addEventListener('scroll', () => {
    const scrollTop = window.pageYOffset || document.documentElement.scrollTop;
    
    // Add shadow to header on scroll
    const header = document.querySelector('.header');
    if (scrollTop > 10) {
        header.style.boxShadow = '0 4px 6px -1px rgba(0, 0, 0, 0.1)';
    } else {
        header.style.boxShadow = '0 1px 2px 0 rgba(0, 0, 0, 0.05)';
    }
    
    lastScrollTop = scrollTop;
});

// Add loading state variations
const loadingMessages = [
    {
        text: 'Analyzing your plant image...',
        subtext: 'Our AI is examining leaf patterns, colors, and textures'
    },
    {
        text: 'Processing image data...',
        subtext: 'Comparing with thousands of plant disease patterns'
    },
    {
        text: 'Running deep learning model...',
        subtext: 'Using ResNet34 architecture for accurate detection'
    }
];

let loadingMessageIndex = 0;
let loadingInterval;

// Start rotating loading messages when loading starts
const originalAddClass = loadingContainer.classList.add.bind(loadingContainer.classList);
loadingContainer.classList.add = function(className) {
    originalAddClass(className);
    if (className === 'active') {
        loadingMessageIndex = 0;
        updateLoadingMessage();
        loadingInterval = setInterval(() => {
            loadingMessageIndex = (loadingMessageIndex + 1) % loadingMessages.length;
            updateLoadingMessage();
        }, 2000);
    }
};

// Stop rotating when loading ends
const originalRemoveClass = loadingContainer.classList.remove.bind(loadingContainer.classList);
loadingContainer.classList.remove = function(className) {
    originalRemoveClass(className);
    if (className === 'active') {
        clearInterval(loadingInterval);
    }
};

function updateLoadingMessage() {
    const message = loadingMessages[loadingMessageIndex];
    const loadingText = document.querySelector('.loading-text');
    const loadingSubtext = document.querySelector('.loading-subtext');
    
    if (loadingText && loadingSubtext) {
        loadingText.style.opacity = '0';
        loadingSubtext.style.opacity = '0';
        
        setTimeout(() => {
            loadingText.textContent = message.text;
            loadingSubtext.textContent = message.subtext;
            loadingText.style.transition = 'opacity 0.3s';
            loadingSubtext.style.transition = 'opacity 0.3s';
            loadingText.style.opacity = '1';
            loadingSubtext.style.opacity = '1';
        }, 300);
    }
}

