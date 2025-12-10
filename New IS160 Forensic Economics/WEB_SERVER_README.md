# Web Server - Forensic Economics AI

## 🚀 Running the Web Application

The Forensic Economics AI system is now available as a web application running on **localhost:8000**.

### Quick Start

#### Option 1: Using the Batch File (Windows)
```bash
run_server.bat
```

#### Option 2: Using PowerShell Script
```powershell
.\run_server.ps1
```

#### Option 3: Direct Python Command
```bash
python app.py
```

### Access the Application

Once the server is running, open your web browser and navigate to:

```
http://localhost:8000
```

## 📋 Features

### Web Interface
- **User-friendly form** for inputting case data
- **Real-time analysis** with progress indicators
- **Results display** with key metrics
- **Excel report download** functionality
- **Example data loader** for quick testing

### API Endpoints

#### 1. Main Page
- **URL:** `http://localhost:8000/`
- **Method:** GET
- **Description:** Returns the main web interface

#### 2. Run Analysis
- **URL:** `http://localhost:8000/api/analyze`
- **Method:** POST
- **Content-Type:** application/json
- **Description:** Runs forensic economic analysis

**Request Body:**
```json
{
  "name": "John Doe",
  "dob": "1980-01-15",
  "dod": "2024-03-20",
  "occupation": "Software Engineer",
  "annual_salary": 120000,
  "sex": "Male",
  "education_level": "Bachelor's",
  "home_county": "Los Angeles",
  "home_state": "California",
  "status": "Active"
}
```

**Response:**
```json
{
  "status": "success",
  "total_economic_loss": 2307881.36,
  "total_economic_loss_formatted": "$2,307,881.36",
  "output_filename": "forensic_economic_report_John_Doe_20251108_133508.xlsx",
  "summary": {
    "age_at_death": 44,
    "remaining_life_years": 33.20,
    "worklife_years": 20.52,
    "base_salary": 120000
  }
}
```

#### 3. Download Excel Report
- **URL:** `http://localhost:8000/api/download/<filename>`
- **Method:** GET
- **Description:** Downloads the generated Excel report

#### 4. Get Example Data
- **URL:** `http://localhost:8000/api/example`
- **Method:** GET
- **Description:** Returns example input data

#### 5. Health Check
- **URL:** `http://localhost:8000/api/health`
- **Method:** GET
- **Description:** Returns server health status

## 🗂️ Directory Structure

```
.
├── app.py                 # Flask web application
├── templates/
│   └── index.html        # Web interface
├── output/               # Generated Excel reports
├── uploads/              # Uploaded files (if needed)
└── run_server.bat        # Windows batch file to start server
└── run_server.ps1        # PowerShell script to start server
```

## 🔧 Configuration

The server runs on:
- **Host:** localhost
- **Port:** 8000
- **Debug Mode:** Enabled (for development)

To change the port, edit `app.py`:
```python
app.run(host='localhost', port=8000, debug=True)
```

## 📊 Usage

### Using the Web Interface

1. **Start the server** using one of the methods above
2. **Open your browser** to `http://localhost:8000`
3. **Fill in the form** with case information:
   - Personal Information (name, DOB, DOD)
   - Employment Information (occupation, salary)
   - Demographics (sex, education, location, status)
4. **Click "Load Example"** to populate with sample data
5. **Click "Run Analysis"** to start the analysis
6. **Wait for results** (typically 20-30 seconds)
7. **Download the Excel report** when analysis completes

### Using the API

#### Python Example
```python
import requests

url = "http://localhost:8000/api/analyze"
data = {
    "name": "John Doe",
    "dob": "1980-01-15",
    "dod": "2024-03-20",
    "occupation": "Software Engineer",
    "annual_salary": 120000,
    "sex": "Male",
    "education_level": "Bachelor's",
    "home_county": "Los Angeles",
    "home_state": "California",
    "status": "Active"
}

response = requests.post(url, json=data)
result = response.json()

if result['status'] == 'success':
    print(f"Total Economic Loss: {result['total_economic_loss_formatted']}")
    # Download Excel file
    filename = result['output_filename']
    download_url = f"http://localhost:8000/api/download/{filename}"
    file_response = requests.get(download_url)
    with open(filename, 'wb') as f:
        f.write(file_response.content)
```

#### cURL Example
```bash
curl -X POST http://localhost:8000/api/analyze \
  -H "Content-Type: application/json" \
  -d '{
    "name": "John Doe",
    "dob": "1980-01-15",
    "dod": "2024-03-20",
    "occupation": "Software Engineer",
    "annual_salary": 120000,
    "sex": "Male",
    "education_level": "Bachelor'\''s",
    "home_county": "Los Angeles",
    "home_state": "California",
    "status": "Active"
  }'
```

## 🛠️ Troubleshooting

### Port Already in Use
If port 8000 is already in use, you can:
1. Change the port in `app.py`
2. Or stop the process using port 8000

### Server Won't Start
- Check if Flask is installed: `pip install flask`
- Check if all dependencies are installed: `pip install -r requirements.txt`
- Check the console for error messages

### Analysis Fails
- Check the browser console for errors
- Check server logs in the terminal
- Verify all required fields are filled in the form
- Check `forensic_economics.log` for detailed error messages

## 🔒 Security Notes

- The server runs on **localhost only** (not accessible from network)
- Uses port 8000 (safe port, not blocked by browsers)
- Debug mode is enabled (disable in production)
- No authentication required (add for production use)
- Files are stored locally in the `output/` directory

## 📝 Notes

- The server processes one analysis at a time
- Analysis typically takes 20-30 seconds
- Excel reports are saved in the `output/` directory
- Server logs are displayed in the terminal
- Detailed logs are saved to `forensic_economics.log`

## 🚀 Production Deployment

For production deployment:
1. Disable debug mode: `debug=False`
2. Use a production WSGI server (e.g., Gunicorn, uWSGI)
3. Add authentication/authorization
4. Configure proper error handling
5. Set up logging
6. Use environment variables for configuration
7. Enable HTTPS

## 📞 Support

For issues or questions:
- Check server logs in the terminal
- Check `forensic_economics.log` for detailed logs
- Review the main README.md for system documentation

---

**Server Status:** ✅ Running on http://localhost:8000

