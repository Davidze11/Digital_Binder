# 🚀 AI PowerPoint Generator

Transform your documents and ideas into professional PowerPoint presentations using AI!

## ✨ Features

- **📁 Multiple File Formats**: Word (.docx), Excel (.xlsx), PowerPoint (.pptx), PDF (.pdf), Text (.txt), OpenDocument (.odt)
- **🤖 AI-Powered Analysis**: Uses Ollama AI for intelligent content summarization
- **🎨 Professional Themes**: Random PowerPoint design themes for polished presentations
- **📊 Smart Charts**: Automatic chart generation from data
- **💻 Multiple Interfaces**: Desktop GUI, Web interface, or Command line

## 🎯 Three Ways to Use

### 1. 🖥️ Desktop GUI (Recommended)
- Native desktop application with file browser
- Tabbed interface for files and text input
- Real-time progress tracking

### 2. 🌐 Web Interface (Modern)
- Runs in your web browser
- Drag & drop file uploads
- Mobile-friendly design

### 3. 💻 Command Line (Advanced)
- Text-based interactive interface
- Full feature access
- Scriptable for automation

## 🚀 Quick Start

### Option 1: Use the Launcher (Easiest)
```bash
python launcher.py
```
Then choose your preferred interface (1, 2, or 3).

### Option 2: Direct Launch
```bash
# Desktop GUI
python simple_gui.py

# Web Interface  
streamlit run web_gui.py

# Command Line
python wrapper.py
```

## 📦 Installation

All required packages are already installed in your virtual environment:
- `python-pptx` - PowerPoint generation
- `python-docx` - Word document processing  
- `pandas`, `openpyxl` - Excel file handling
- `matplotlib`, `seaborn`, `plotly` - Chart generation
- `requests`, `pillow` - Web and image processing
- `ollama` - AI integration
- `tkinterdnd2` - Drag & drop support
- `streamlit` - Web interface

## 🎨 How It Works

### File Processing Mode
1. **Upload/Select Files**: Choose your documents
2. **AI Analysis**: Content is analyzed and summarized
3. **Structure Creation**: Information is organized into slides
4. **Theme Application**: Professional design is applied
5. **Generation**: Your PowerPoint is created

### Text Input Mode
1. **Enter Content**: Type or paste your ideas
2. **AI Processing**: Content is structured and enhanced
3. **Slide Creation**: Automatic slide layout and formatting
4. **Download**: Get your finished presentation

## 📋 Supported Content Types

- **Business Documents**: Reports, proposals, strategic plans
- **Research Papers**: Studies, findings, methodologies
- **Project Plans**: Timelines, budgets, resource allocation
- **Data Analysis**: Excel spreadsheets with charts and insights
- **Meeting Notes**: Action items, decisions, next steps
- **Training Materials**: Procedures, guidelines, knowledge bases

## 🤖 AI Features

When Ollama is available, the system provides:
- **Executive Summaries**: Key points extraction
- **Business Insights**: Strategic recommendations
- **Content Organization**: Logical slide structure
- **Professional Language**: Presentation-ready text

## 📁 File Structure

```
AI POWERPOINT WRAPPER/
├── wrapper.py              # Main processing engine
├── launcher.py             # Interface selector
├── simple_gui.py           # Desktop GUI interface  
├── web_gui.py              # Web interface
├── gui_wrapper.py          # Advanced GUI with drag-drop
├── test_sample.txt         # Sample test file
└── wrapper test/           # Sample files and output folder
```

## 🎯 Example Use Cases

### Business Presentations
- Convert quarterly reports to executive summaries
- Transform project plans into stakeholder presentations
- Generate training materials from documentation

### Academic Research  
- Create presentation slides from research papers
- Visualize data analysis results
- Present study findings and methodologies

### Project Management
- Convert project documentation to status presentations
- Create milestone review materials
- Generate client update presentations

## 🔧 Troubleshooting

### Common Issues

**GUI won't start:**
```bash
# Install missing packages
pip install tkinterdnd2 streamlit
```

**Web interface issues:**
```bash
# Restart Streamlit
streamlit run web_gui.py --server.port 8502
```

**AI not working:**
- Ollama AI is optional - built-in analysis works without it
- For full AI features, install Ollama separately

### File Processing Issues
- **Unsupported format**: Check file extension (.docx, .xlsx, etc.)
- **Corrupted files**: Try opening the file in its native application first
- **Large files**: Processing may take longer for files over 10MB

## 💡 Tips for Best Results

1. **File Organization**: Use descriptive filenames
2. **Content Quality**: Well-structured documents produce better slides
3. **Title Selection**: Choose meaningful presentation titles
4. **Multiple Files**: Process related documents together for comprehensive presentations

## 🆘 Support

If you encounter issues:
1. Check that all files are in supported formats
2. Ensure virtual environment is activated
3. Verify all packages are installed correctly
4. Try the command line interface for detailed error messages

## 🎉 Success!

Your AI PowerPoint Generator is ready to use! Choose your preferred interface and start creating professional presentations from your documents and ideas.

---
*Generated with ❤️ by AI PowerPoint Generator*