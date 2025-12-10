import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import threading
from pathlib import Path
import sys

# Import our wrapper class
try:
    sys.path.append(os.path.dirname(os.path.abspath(__file__)))
    from wrapper import OllamaDocumentToPowerPoint
except ImportError as e:
    print(f"Error: Could not import wrapper.py: {e}")
    print("Make sure wrapper.py is in the same directory.")

class SimplePowerPointGUI:
    def __init__(self):
        # Create main window
        self.root = tk.Tk()
        self.root.title("🚀 AI PowerPoint Generator - Simple Version")
        self.root.geometry("800x600")
        self.root.configure(bg='#f0f0f0')
        
        # Configure style
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Variables
        self.files_list = []
        self.converter = None
        
        self.setup_ui()
        
    def setup_ui(self):
        """Setup the user interface"""
        # Main container
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text="🚀 AI-Powered PowerPoint Generator", 
                               font=('Arial', 18, 'bold'))
        title_label.pack(pady=(0, 20))
        
        # Create notebook for tabs
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        # Tab 1: File Processing
        self.file_frame = ttk.Frame(notebook, padding="20")
        notebook.add(self.file_frame, text="📁 File Processing")
        
        # Tab 2: Text Input
        self.text_frame = ttk.Frame(notebook, padding="20")
        notebook.add(self.text_frame, text="📝 Text Input")
        
        # Setup tabs
        self.setup_file_tab()
        self.setup_text_tab()
        
        # Bottom controls
        self.setup_bottom_controls(main_frame)
        
    def setup_file_tab(self):
        """Setup the file processing tab"""
        # Instructions
        instructions = ttk.Label(self.file_frame, 
                               text="Click 'Browse Files' to select documents\n" +
                                    "Supported: .docx, .xlsx, .pptx, .pdf, .txt, .odt",
                               font=('Arial', 10))
        instructions.pack(pady=(0, 10))
        
        # File selection area
        file_area = tk.Frame(self.file_frame, bg='#e8f4f8', relief='solid', bd=2, height=150)
        file_area.pack(fill=tk.X, pady=(0, 10))
        file_area.pack_propagate(False)
        
        # Browse button
        browse_btn = ttk.Button(file_area, text="📂 Browse Files", 
                               command=self.browse_files)
        browse_btn.pack(pady=20)
        
        # Info label
        info_label = tk.Label(file_area,
                             text="Select Word, Excel, PowerPoint, PDF, or Text files\n" +
                                  "to automatically generate presentations",
                             bg='#e8f4f8',
                             fg='#666666',
                             font=('Arial', 10))
        info_label.pack()
        
        # Files listbox
        list_label = ttk.Label(self.file_frame, text="Selected Files:")
        list_label.pack(anchor=tk.W, pady=(10, 5))
        
        # Listbox frame
        list_frame = ttk.Frame(self.file_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        self.files_listbox = tk.Listbox(list_frame, height=8)
        self.files_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.files_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.files_listbox.configure(yscrollcommand=scrollbar.set)
        
        # File control buttons
        btn_frame = ttk.Frame(self.file_frame)
        btn_frame.pack(fill=tk.X, pady=(5, 0))
        
        remove_btn = ttk.Button(btn_frame, text="❌ Remove Selected", 
                               command=self.remove_selected_file)
        remove_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        clear_btn = ttk.Button(btn_frame, text="🗑️ Clear All", 
                              command=self.clear_files)
        clear_btn.pack(side=tk.LEFT)
        
    def setup_text_tab(self):
        """Setup the text input tab"""
        # Instructions
        text_instructions = ttk.Label(self.text_frame, 
                                    text="Type or paste your content below to create a PowerPoint presentation",
                                    font=('Arial', 10))
        text_instructions.pack(pady=(0, 10))
        
        # Text input area
        text_label = ttk.Label(self.text_frame, text="Your Content:")
        text_label.pack(anchor=tk.W, pady=(0, 5))
        
        self.text_input = scrolledtext.ScrolledText(self.text_frame, 
                                                   height=15, 
                                                   wrap=tk.WORD,
                                                   font=('Arial', 10))
        self.text_input.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Placeholder text
        placeholder_text = """Example: Type your ideas here...

Project Overview:
• Develop new customer portal
• Implement automated workflows  
• Create mobile applications

Timeline: 6 months
Budget: $75,000
Team: 5 developers

Key Benefits:
- 40% efficiency improvement
- Better customer experience
- Reduced operational costs

Next Steps:
1. Finalize requirements
2. Begin development phase
3. User testing and feedback"""
        
        self.text_input.insert(tk.INSERT, placeholder_text)
        self.text_input.bind('<FocusIn>', self.on_text_focus_in)
        
        # Text control buttons
        text_btn_frame = ttk.Frame(self.text_frame)
        text_btn_frame.pack(fill=tk.X, pady=(5, 0))
        
        clear_text_btn = ttk.Button(text_btn_frame, text="🗑️ Clear Text", 
                                   command=self.clear_text)
        clear_text_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        sample_btn = ttk.Button(text_btn_frame, text="📝 Load Sample", 
                               command=self.load_sample_text)
        sample_btn.pack(side=tk.LEFT)
        
    def setup_bottom_controls(self, parent):
        """Setup bottom control panel"""
        controls_frame = ttk.LabelFrame(parent, text="Presentation Settings", padding="15")
        controls_frame.pack(fill=tk.X, pady=(10, 0))
        
        # First row - Title and AI setting
        row1 = ttk.Frame(controls_frame)
        row1.pack(fill=tk.X, pady=(0, 10))
        
        title_label = ttk.Label(row1, text="Title:")
        title_label.pack(side=tk.LEFT)
        
        self.title_var = tk.StringVar(value="AI Generated Presentation")
        title_entry = ttk.Entry(row1, textvariable=self.title_var, width=40)
        title_entry.pack(side=tk.LEFT, padx=(10, 20), fill=tk.X, expand=True)
        
        self.ai_enabled = tk.BooleanVar(value=True)
        ai_check = ttk.Checkbutton(row1, text="🤖 Enable AI Analysis", 
                                  variable=self.ai_enabled)
        ai_check.pack(side=tk.RIGHT)
        
        # Progress section
        self.progress_var = tk.StringVar(value="Ready to generate presentation")
        progress_label = ttk.Label(controls_frame, textvariable=self.progress_var)
        progress_label.pack(anchor=tk.W, pady=(0, 5))
        
        self.progress_bar = ttk.Progressbar(controls_frame, mode='indeterminate')
        self.progress_bar.pack(fill=tk.X, pady=(0, 10))
        
        # Generate button
        self.generate_btn = ttk.Button(controls_frame, text="🚀 Generate PowerPoint", 
                                     command=self.generate_presentation)
        self.generate_btn.pack(pady=(5, 0))
        
        # Configure button style
        self.generate_btn.configure(style='Accent.TButton')
        self.style.configure('Accent.TButton', font=('Arial', 11, 'bold'))
        
    def browse_files(self):
        """Open file browser"""
        filetypes = [
            ("All Supported", "*.docx *.xlsx *.pptx *.pdf *.txt *.odt"),
            ("Word Documents", "*.docx"),
            ("Excel Files", "*.xlsx"),
            ("PowerPoint Files", "*.pptx"),
            ("PDF Files", "*.pdf"),
            ("Text Files", "*.txt"),
            ("OpenDocument", "*.odt"),
            ("All Files", "*.*")
        ]
        
        files = filedialog.askopenfilenames(
            title="Select files to process",
            filetypes=filetypes
        )
        
        for file_path in files:
            self.add_file(file_path)
        
    def add_file(self, file_path):
        """Add file to the list"""
        if file_path not in self.files_list:
            # Check if file extension is supported
            ext = Path(file_path).suffix.lower()
            supported = ['.docx', '.xlsx', '.pptx', '.pdf', '.txt', '.odt']
            
            if ext in supported:
                self.files_list.append(file_path)
                self.files_listbox.insert(tk.END, os.path.basename(file_path))
            else:
                messagebox.showwarning("Unsupported File", 
                                     f"File type '{ext}' is not supported.\n" +
                                     "Supported: .docx, .xlsx, .pptx, .pdf, .txt, .odt")
        
    def remove_selected_file(self):
        """Remove selected file from list"""
        selection = self.files_listbox.curselection()
        if selection:
            index = selection[0]
            self.files_listbox.delete(index)
            self.files_list.pop(index)
            
    def clear_files(self):
        """Clear all files"""
        self.files_list.clear()
        self.files_listbox.delete(0, tk.END)
    
    def on_text_focus_in(self, event):
        """Clear placeholder text when text area is focused"""
        current_text = self.text_input.get("1.0", tk.END).strip()
        if "Example: Type your ideas here..." in current_text:
            self.text_input.delete("1.0", tk.END)
    
    def clear_text(self):
        """Clear text input area"""
        self.text_input.delete("1.0", tk.END)
    
    def load_sample_text(self):
        """Load sample text"""
        sample = """Business Innovation Strategy

Executive Summary:
Our organization is launching three strategic initiatives to enhance competitiveness and drive sustainable growth in the digital marketplace.

Strategic Initiatives:

1. Digital Customer Experience Platform
   • 360-degree customer view
   • Personalized interaction engines  
   • Omnichannel communication
   • Real-time analytics dashboard

2. AI-Powered Operations Optimization
   • Automated workflow management
   • Predictive maintenance systems
   • Supply chain intelligence
   • Quality control automation

3. Employee Digital Empowerment
   • Remote collaboration tools
   • Skills development platform
   • Performance analytics
   • Flexible work arrangements

Financial Projections:
• Total Investment: $200,000
• Expected ROI: 35% over 24 months
• Break-even: 14 months
• Annual savings: $180,000

Implementation Roadmap:
Phase 1 (Months 1-3): Foundation & Planning
Phase 2 (Months 4-8): Development & Testing  
Phase 3 (Months 9-12): Deployment & Training
Phase 4 (Months 13-15): Optimization & Scale

Success Metrics:
✓ Customer satisfaction: +25%
✓ Operational efficiency: +30%  
✓ Employee engagement: +20%
✓ Cost reduction: 15%

Risk Mitigation:
• Phased rollout approach
• Comprehensive training programs
• Regular stakeholder communication
• Continuous monitoring and adjustment"""
        
        self.text_input.delete("1.0", tk.END)
        self.text_input.insert("1.0", sample)
    
    def generate_presentation(self):
        """Generate PowerPoint presentation"""
        # Get active tab
        notebook = None
        for child in self.root.winfo_children():
            for grandchild in child.winfo_children():
                if isinstance(grandchild, ttk.Notebook):
                    notebook = grandchild
                    break
        
        if not notebook:
            messagebox.showerror("Error", "Could not determine active tab")
            return
            
        current_tab = notebook.select()
        tab_text = notebook.tab(current_tab, "text")
        
        # Validate inputs
        title = self.title_var.get().strip()
        if not title:
            title = "AI Generated Presentation"
        
        if "File Processing" in tab_text:
            if not self.files_list:
                messagebox.showwarning("No Files", "Please select at least one file to process.")
                return
        elif "Text Input" in tab_text:
            text_content = self.text_input.get("1.0", tk.END).strip()
            if not text_content or "Example: Type your ideas here..." in text_content:
                messagebox.showwarning("No Content", "Please enter some text content.")
                return
        
        # Disable button and start progress
        self.generate_btn.config(state='disabled')
        self.progress_bar.start()
        
        # Run generation in separate thread
        thread = threading.Thread(target=self._generate_presentation_thread, args=(tab_text, title))
        thread.daemon = True
        thread.start()
    
    def _generate_presentation_thread(self, tab_text, title):
        """Generate presentation in separate thread"""
        try:
            # Initialize converter
            self.converter = OllamaDocumentToPowerPoint()
            output_file = None
            
            if "File Processing" in tab_text:
                # Process files
                self.update_progress("Processing files...")
                
                # Categorize files by extension
                file_categories = {
                    'docx_files': [f for f in self.files_list if f.endswith('.docx')],
                    'xlsx_files': [f for f in self.files_list if f.endswith('.xlsx')],
                    'pptx_files': [f for f in self.files_list if f.endswith('.pptx')],
                    'pdf_files': [f for f in self.files_list if f.endswith('.pdf')],
                    'txt_files': [f for f in self.files_list if f.endswith('.txt')],
                    'odt_files': [f for f in self.files_list if f.endswith('.odt')]
                }
                
                # Process all files
                all_content = self.converter.process_files(**file_categories)
                
                if all_content:
                    self.update_progress("Generating presentation...")
                    output_file = self.converter.create_presentation(
                        all_content=all_content,
                        presentation_title=title
                    )
                
            elif "Text Input" in tab_text:
                # Process text input
                text_content = self.text_input.get("1.0", tk.END).strip()
                self.update_progress("Analyzing text content...")
                output_file = self.converter.create_presentation_from_text(
                    user_text=text_content,
                    presentation_title=title,
                    enable_ai_analysis=self.ai_enabled.get()
                )
            
            # Update UI on main thread
            self.root.after(0, self._generation_complete, output_file)
            
        except Exception as e:
            error_msg = f"Error generating presentation: {str(e)}"
            self.root.after(0, self._generation_error, error_msg)
    
    def update_progress(self, message):
        """Update progress message (thread-safe)"""
        self.root.after(0, lambda: self.progress_var.set(message))
    
    def _generation_complete(self, output_file):
        """Handle successful generation"""
        self.progress_bar.stop()
        self.generate_btn.config(state='normal')
        
        if output_file and os.path.exists(output_file):
            self.progress_var.set(f"✅ Presentation created: {os.path.basename(output_file)}")
            
            result = messagebox.askyesnocancel(
                "Success!", 
                f"PowerPoint presentation created successfully!\n\n" +
                f"File: {os.path.basename(output_file)}\n\n" +
                "Would you like to:\n" +
                "• Yes: Open the presentation\n" +
                "• No: Open containing folder\n" +
                "• Cancel: Do nothing"
            )
            
            if result is True:  # Yes - open presentation
                try:
                    os.startfile(output_file)
                except:
                    os.startfile(os.path.dirname(output_file))
            elif result is False:  # No - open folder
                os.startfile(os.path.dirname(output_file))
        else:
            self.progress_var.set("❌ Failed to create presentation")
            messagebox.showerror("Error", "Failed to create presentation. Please check your inputs.")
    
    def _generation_error(self, error_msg):
        """Handle generation error"""
        self.progress_bar.stop()
        self.generate_btn.config(state='normal')
        self.progress_var.set("❌ Error occurred")
        messagebox.showerror("Error", error_msg)
    
    def run(self):
        """Start the GUI"""
        self.root.mainloop()

def main():
    """Main function to run the GUI"""
    try:
        print("🚀 Starting AI PowerPoint Generator GUI...")
        app = SimplePowerPointGUI()
        app.run()
    except Exception as e:
        print(f"Error starting GUI: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()