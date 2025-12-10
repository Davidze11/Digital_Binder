import streamlit as st
import os
import tempfile
import threading
from pathlib import Path
import sys

# Add current directory to path for imports
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

try:
    from wrapper import OllamaDocumentToPowerPoint
except ImportError as e:
    st.error(f"Could not import wrapper module: {e}")
    st.stop()

def main():
    # Page configuration
    st.set_page_config(
        page_title="🚀 AI PowerPoint Generator",
        page_icon="🚀",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    # Custom CSS for better styling
    st.markdown("""
    <style>
    .main-header {
        text-align: center;
        padding: 2rem 0;
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        color: white;
        border-radius: 10px;
        margin-bottom: 2rem;
    }
    .upload-box {
        border: 2px dashed #cccccc;
        border-radius: 10px;
        padding: 2rem;
        text-align: center;
        background-color: #f8f9fa;
        margin: 1rem 0;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
    </style>
    """, unsafe_allow_html=True)
    
    # Header
    st.markdown("""
    <div class="main-header">
        <h1>🚀 AI-Powered PowerPoint Generator</h1>
        <p>Transform your documents and ideas into professional presentations</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar for settings
    with st.sidebar:
        st.header("⚙️ Settings")
        
        # Presentation title
        presentation_title = st.text_input(
            "Presentation Title",
            value="AI Generated Presentation",
            help="Enter a title for your presentation"
        )
        
        # AI Analysis toggle
        enable_ai = st.checkbox(
            "🤖 Enable AI Analysis",
            value=True,
            help="Use Ollama AI for intelligent content analysis and generation"
        )
        
        # Output format options
        st.subheader("📊 Output Options")
        include_charts = st.checkbox("Include Charts", value=True)
        include_images = st.checkbox("Include Images", value=True)
        
        st.markdown("---")
        
        # Help section
        with st.expander("ℹ️ Help & Instructions"):
            st.markdown("""
            **Supported File Types:**
            - Word Documents (.docx)
            - Excel Files (.xlsx)
            - PowerPoint Files (.pptx)
            - PDF Files (.pdf)
            - Text Files (.txt)
            - OpenDocument (.odt)
            
            **How to Use:**
            1. Choose between file upload or text input
            2. Upload files or enter your content
            3. Customize settings in the sidebar
            4. Click 'Generate Presentation'
            5. Download your PowerPoint file
            """)
    
    # Main content area with tabs
    tab1, tab2 = st.tabs(["📁 File Upload", "📝 Text Input"])
    
    with tab1:
        st.header("📁 File Processing")
        st.write("Upload documents to automatically generate a PowerPoint presentation")
        
        # File uploader
        uploaded_files = st.file_uploader(
            "Choose files to process",
            type=['docx', 'xlsx', 'pptx', 'pdf', 'txt', 'odt'],
            accept_multiple_files=True,
            help="Drag and drop files here or click to browse"
        )
        
        if uploaded_files:
            st.success(f"✅ {len(uploaded_files)} file(s) uploaded successfully!")
            
            # Display uploaded files
            for file in uploaded_files:
                col1, col2, col3 = st.columns([3, 1, 1])
                with col1:
                    st.write(f"📄 {file.name}")
                with col2:
                    st.write(f"{file.size / 1024:.1f} KB")
                with col3:
                    st.write(file.type or "Unknown")
        
        # Generate button for files
        if st.button("🚀 Generate from Files", key="generate_files", disabled=not uploaded_files):
            generate_from_files(uploaded_files, presentation_title, enable_ai)
    
    with tab2:
        st.header("📝 Text Input")
        st.write("Enter your content below to create a presentation")
        
        # Text input area
        user_text = st.text_area(
            "Your Content",
            height=300,
            placeholder="""Enter your ideas, notes, or content here...

Example:
Project Overview:
• Develop customer portal
• Implement automation
• Create mobile app

Budget: $50,000
Timeline: 6 months
Expected ROI: 25%

Key Benefits:
- Increased efficiency
- Better customer experience
- Cost savings""",
            help="Type or paste your content here. The AI will structure it into slides."
        )
        
        # Sample text buttons
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("📊 Business Sample"):
                st.session_state.sample_text = get_business_sample()
        with col2:
            if st.button("🔬 Research Sample"):
                st.session_state.sample_text = get_research_sample()
        with col3:
            if st.button("🎯 Project Sample"):
                st.session_state.sample_text = get_project_sample()
        
        # Display sample text if selected
        if 'sample_text' in st.session_state:
            user_text = st.text_area(
                "Sample Content (you can edit this)",
                value=st.session_state.sample_text,
                height=300,
                key="sample_text_area"
            )
        
        # Generate button for text
        if st.button("🚀 Generate from Text", key="generate_text", disabled=not user_text.strip()):
            generate_from_text(user_text, presentation_title, enable_ai)

def generate_from_files(uploaded_files, title, enable_ai):
    """Generate presentation from uploaded files"""
    with st.spinner("Processing files and generating presentation..."):
        try:
            # Create temporary directory for uploaded files
            with tempfile.TemporaryDirectory() as temp_dir:
                # Save uploaded files to temp directory
                temp_files = []
                for uploaded_file in uploaded_files:
                    temp_path = os.path.join(temp_dir, uploaded_file.name)
                    with open(temp_path, "wb") as f:
                        f.write(uploaded_file.getvalue())
                    temp_files.append(temp_path)
                
                # Initialize converter
                converter = OllamaDocumentToPowerPoint()
                
                # Categorize files by extension
                file_categories = {
                    'docx_files': [f for f in temp_files if f.endswith('.docx')],
                    'xlsx_files': [f for f in temp_files if f.endswith('.xlsx')],
                    'pptx_files': [f for f in temp_files if f.endswith('.pptx')],
                    'pdf_files': [f for f in temp_files if f.endswith('.pdf')],
                    'txt_files': [f for f in temp_files if f.endswith('.txt')],
                    'odt_files': [f for f in temp_files if f.endswith('.odt')]
                }
                
                # Process files
                st.write("📊 Processing uploaded files...")
                all_content = converter.process_files(**file_categories)
                
                if all_content:
                    st.write("🎨 Generating presentation...")
                    output_file = converter.create_presentation(
                        all_content=all_content,
                        presentation_title=title
                    )
                    
                    if output_file and os.path.exists(output_file):
                        # Provide download link
                        with open(output_file, "rb") as file:
                            st.success("✅ Presentation generated successfully!")
                            st.download_button(
                                label="📥 Download PowerPoint",
                                data=file.read(),
                                file_name=f"{title}.pptx",
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                            )
                    else:
                        st.error("❌ Failed to generate presentation")
                else:
                    st.error("❌ No content could be extracted from uploaded files")
                    
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")

def generate_from_text(user_text, title, enable_ai):
    """Generate presentation from text input"""
    with st.spinner("Analyzing text and generating presentation..."):
        try:
            # Initialize converter
            converter = OllamaDocumentToPowerPoint()
            
            st.write("🤖 Analyzing your text...")
            output_file = converter.create_presentation_from_text(
                user_text=user_text,
                presentation_title=title,
                enable_ai_analysis=enable_ai
            )
            
            if output_file and os.path.exists(output_file):
                # Provide download link
                with open(output_file, "rb") as file:
                    st.success("✅ Presentation generated successfully!")
                    st.download_button(
                        label="📥 Download PowerPoint",
                        data=file.read(),
                        file_name=f"{title}.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
            else:
                st.error("❌ Failed to generate presentation")
                
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")

def get_business_sample():
    return """Business Strategy Presentation

Executive Summary:
Our company is implementing digital transformation initiatives to improve efficiency and customer satisfaction.

Key Initiatives:
• Customer Experience Platform
  - 360-degree customer view
  - Personalized experiences
  - Omnichannel support

• Process Automation
  - Workflow optimization
  - Reduced manual tasks
  - Faster response times

• Data Analytics Platform
  - Real-time insights
  - Predictive analytics
  - Performance monitoring

Financial Impact:
Investment: $150,000
Expected ROI: 30% in 18 months
Annual savings: $200,000

Timeline:
Q1: Planning and setup
Q2: Implementation begins
Q3: Testing and training
Q4: Full deployment

Success Metrics:
- Customer satisfaction: +25%
- Process efficiency: +40%
- Cost reduction: 20%"""

def get_research_sample():
    return """Research Findings Report

Study Overview:
Analysis of customer behavior patterns in digital retail environments over 12-month period.

Methodology:
• Sample size: 10,000 customers
• Data collection: Online surveys and behavioral tracking
• Analysis period: January - December 2024
• Statistical methods: Regression analysis, clustering

Key Findings:
1. Mobile Usage Increase
   - 65% of purchases via mobile devices
   - 40% increase from previous year
   - Peak usage during evening hours

2. Personalization Impact
   - Personalized recommendations increase sales by 35%
   - Customer engagement improves by 50%
   - Return rate decreases by 25%

3. Customer Journey Insights
   - Average 3.2 touchpoints before purchase
   - Social media influences 70% of decisions
   - Reviews impact 85% of buying decisions

Recommendations:
• Invest in mobile optimization
• Enhance personalization algorithms
• Improve social media presence
• Streamline customer journey

Conclusion:
Digital retail behavior shows clear trends toward mobile-first, personalized experiences."""

def get_project_sample():
    return """Project Management Plan

Project: Customer Portal Development

Objective:
Create a comprehensive customer self-service portal to improve user experience and reduce support costs.

Scope:
• User registration and authentication
• Account management features
• Order tracking and history
• Support ticket system
• Knowledge base integration

Team Structure:
- Project Manager: Sarah Johnson
- Lead Developer: Mike Chen
- UI/UX Designer: Emma Davis
- QA Engineer: Alex Rodriguez
- Business Analyst: Lisa Parker

Timeline (16 weeks):
Phase 1: Requirements & Design (4 weeks)
Phase 2: Development (8 weeks)
Phase 3: Testing & Deployment (3 weeks)
Phase 4: Launch & Support (1 week)

Budget Breakdown:
Development: $80,000
Design: $25,000
Testing: $15,000
Infrastructure: $10,000
Total: $130,000

Risk Management:
• Technical risks: Prototype early, regular testing
• Resource risks: Cross-training team members
• Timeline risks: Buffer time built into schedule

Success Criteria:
- 90% user adoption within 3 months
- 50% reduction in support tickets
- User satisfaction score > 4.5/5
- System uptime > 99.5%"""

if __name__ == "__main__":
    main()