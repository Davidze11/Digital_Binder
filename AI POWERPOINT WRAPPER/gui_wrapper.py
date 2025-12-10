import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import tkinter.dnd as dnd
from tkinterdnd2 import DND_FILES, TkinterDnD
import os
import threading
from pathlib import Path
import sys

# Import our wrapper class
try:
    from wrapper import OllamaDocumentToPowerPoint
except ImportError:
    print("Error: Could not import wrapper.py. Make sure it's in the same directory.")
    sys.exit(1)

class PowerPointGeneratorGUI:
    def __init__(self):
        # Create main window with drag-and-drop support
        self.root = TkinterDnD.Tk()
        self.root.title("🚀 AI PowerPoint Generator")
        self.root.geometry("800x700")
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
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="🚀 AI-Powered PowerPoint Generator", 
                               font=('Arial', 18, 'bold'))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Create notebook for tabs
        notebook = ttk.Notebook(main_frame)
        notebook.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 20))
        main_frame.rowconfigure(1, weight=1)
        
        # Tab 1: File Processing
        self.file_frame = ttk.Frame(notebook, padding="20")
        notebook.add(self.file_frame, text="📁 File Processing")
        
        # Tab 2: Text Input
        self.text_frame = ttk.Frame(notebook, padding="20")
        notebook.add(self.text_frame, text="📝 Text Input")
        
        # Setup file processing tab
        self.setup_file_tab()
        
        # Setup text input tab
        self.setup_text_tab()
        
        # Bottom controls
        self.setup_bottom_controls(main_frame)
        
    def setup_file_tab(self):
        """Setup the file processing tab"""
        # Instructions
        instructions = ttk.Label(self.file_frame, 
                               text="Drag & Drop files here or click 'Browse Files'\n" +
                                    "Supported: Word, Excel, PowerPoint, PDF, Text, OpenDocument",
                               font=('Arial', 10))
        instructions.grid(row=0, column=0, columnspan=2, pady=(0, 10))
        
        # Drag and drop area
        self.drop_frame = tk.Frame(self.file_frame, 
                                  bg='#e8f4f8', 
                                  relief='ridge', 
                                  bd=2,
                                  height=200)
        self.drop_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        self.file_frame.columnconfigure(0, weight=1)
        
        # Drop label
        self.drop_label = tk.Label(self.drop_frame,
                                  text="📁 Drop files here\n\nSupported formats:\n" +
                                       "Word files, Excel files\n" +
                                       "PowerPoint files, PDF files\n" +
                                       "Text files, OpenOffice files",
                                  bg='#e8f4f8',
                                  fg='#666666',
                                  font=('Arial', 12))
        self.drop_label.place(relx=0.5, rely=0.5, anchor='center')
        
        # Configure drag and drop
        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind('<<Drop>>', self.on_file_drop)
        
        # Browse button
        browse_btn = ttk.Button(self.file_frame, text="📂 Browse Files", 
                               command=self.browse_files)
        browse_btn.grid(row=2, column=0, pady=(0, 10), sticky=tk.W)
        
        # Clear files button
        clear_btn = ttk.Button(self.file_frame, text="🗑️ Clear Files", 
                              command=self.clear_files)
        clear_btn.grid(row=2, column=1, pady=(0, 10), sticky=tk.E)
        
        # Files listbox
        list_label = ttk.Label(self.file_frame, text="Selected Files:")
        list_label.grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=(10, 5))
        
        # Listbox with scrollbar
        list_frame = ttk.Frame(self.file_frame)
        list_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        self.file_frame.rowconfigure(4, weight=1)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        self.files_listbox = tk.Listbox(list_frame, height=6)
        self.files_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.files_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.files_listbox.configure(yscrollcommand=scrollbar.set)
        
        # Remove selected file button
        remove_btn = ttk.Button(self.file_frame, text="❌ Remove Selected", 
                               command=self.remove_selected_file)
        remove_btn.grid(row=5, column=0, columnspan=2, pady=(5, 0))
        
    def setup_text_tab(self):
        """Setup the text input tab"""
        # Instructions
        text_instructions = ttk.Label(self.text_frame, 
                                    text="Type or paste your content below to create a PowerPoint presentation",
                                    font=('Arial', 10))
        text_instructions.grid(row=0, column=0, columnspan=2, pady=(0, 10))
        
        # Text input area
        text_label = ttk.Label(self.text_frame, text="Your Content:")
        text_label.grid(row=1, column=0, sticky=tk.W, pady=(0, 5))
        
        self.text_input = scrolledtext.ScrolledText(self.text_frame, 
                                                   height=15, 
                                                   wrap=tk.WORD,
                                                   font=('Arial', 10))
        self.text_input.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        self.text_frame.columnconfigure(0, weight=1)
        self.text_frame.rowconfigure(2, weight=1)
        
        # Placeholder text
        placeholder_text = """Example: Type your ideas here...

Project Ideas:
• Implement AI-powered customer service
• Develop mobile app for remote work  
• Create automated reporting dashboard

Budget: $50,000
Timeline: 6 months
Expected ROI: 25%

Key Benefits:
- Increased efficiency
- Cost savings
- Better customer satisfaction"""
        
        self.text_input.insert(tk.INSERT, placeholder_text)
        self.text_input.bind('<FocusIn>', self.on_text_focus_in)
        
        # Clear text button
        clear_text_btn = ttk.Button(self.text_frame, text="🗑️ Clear Text", 
                                   command=self.clear_text)
        clear_text_btn.grid(row=3, column=0, pady=(5, 0), sticky=tk.W)
        
        # Sample text button
        sample_btn = ttk.Button(self.text_frame, text="📝 Load Sample Text", 
                               command=self.load_sample_text)
        sample_btn.grid(row=3, column=1, pady=(5, 0), sticky=tk.E)
        
    def setup_bottom_controls(self, parent):
        """Setup bottom control panel"""
        controls_frame = ttk.LabelFrame(parent, text="Presentation Settings", padding="10")
        controls_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        controls_frame.columnconfigure(1, weight=1)
        
        # Presentation title
        title_label = ttk.Label(controls_frame, text="Title:")
        title_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        self.title_var = tk.StringVar(value="AI Generated Presentation")
        title_entry = ttk.Entry(controls_frame, textvariable=self.title_var, width=40)
        title_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        # AI Analysis checkbox
        self.ai_enabled = tk.BooleanVar(value=True)
        ai_check = ttk.Checkbutton(controls_frame, text="🤖 Enable AI Analysis", 
                                  variable=self.ai_enabled)
        ai_check.grid(row=0, column=2, sticky=tk.W)
        
        # Progress bar
        self.progress_var = tk.StringVar(value="Ready to generate presentation")
        progress_label = ttk.Label(controls_frame, textvariable=self.progress_var)
        progress_label.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(10, 5))
        
        self.progress_bar = ttk.Progressbar(controls_frame, mode='indeterminate')
        self.progress_bar.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Generate button
        self.generate_btn = ttk.Button(controls_frame, text="🚀 Generate PowerPoint", 
                                     command=self.generate_presentation,
                                     style='Accent.TButton')
        self.generate_btn.grid(row=3, column=0, columnspan=3, pady=(5, 0))
        
        # Style the generate button
        self.style.configure('Accent.TButton', font=('Arial', 11, 'bold'))
        
    def on_file_drop(self, event):
        """Handle file drop event"""
        files = self.root.tk.splitlist(event.data)
        for file_path in files:
            if os.path.isfile(file_path):
                self.add_file(file_path)
        self.update_drop_label()
        
    def browse_files(self):
        """Open file browser"""
        filetypes = [
            ("All Supported", "*.docx;*.xlsx;*.pptx;*.pdf;*.txt;*.odt"),
            ("Word Files", "*.docx"),
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
        self.update_drop_label()
        
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
                                     "Supported: Word, Excel, PowerPoint, PDF, Text, OpenDocument")
        
    def remove_selected_file(self):
        """Remove selected file from list"""
        selection = self.files_listbox.curselection()
        if selection:
            index = selection[0]
            self.files_listbox.delete(index)
            self.files_list.pop(index)
            self.update_drop_label()
            
    def clear_files(self):
        """Clear all files"""
        self.files_list.clear()
        self.files_listbox.delete(0, tk.END)
        self.update_drop_label()
        
    def update_drop_label(self):
        """Update the drop area label"""
        if self.files_list:
            self.drop_label.config(text=f"📁 {len(self.files_list)} file(s) selected\n\nDrop more files or use Browse")
        else:
            self.drop_label.config(text="📁 Drop files here\n\nSupported formats:\n" +
                                      "Word files, Excel files\n" +
                                      "PowerPoint files, PDF files\n" +
                                      "Text files, OpenOffice files")
    
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
        sample = """Business Strategy Presentation

Executive Summary:
Our company is implementing three key initiatives to drive growth and improve efficiency.

Key Projects:
• AI-Powered Customer Service
  - 24/7 automated support
  - 50% reduction in response time
  - Improved customer satisfaction

• Mobile Remote Work Platform
  - Seamless team collaboration
  - Video conferencing integration
  - Offline capability with auto-sync

• Automated Reporting Dashboard
  - Real-time data visualization
  - Custom KPI tracking
  - Alert system for critical metrics

Financial Overview:
- Total Investment: $150,000
- Expected ROI: 30% within 18 months
- Break-even point: 12 months

Implementation Timeline:
Q1: Planning and team assembly
Q2: Development phase begins
Q3: Testing and refinement
Q4: Launch and optimization

Success Metrics:
- Customer satisfaction score > 95%
- Employee productivity increase 25%
- Cost reduction of 20%"""
        
        self.text_input.delete("1.0", tk.END)
        self.text_input.insert("1.0", sample)
    
    def generate_presentation(self):
        """Generate PowerPoint presentation"""
        # Disable button during generation
        self.generate_btn.config(state='disabled')
        
        # Start progress bar
        self.progress_bar.start()
        
        # Run generation in separate thread to prevent UI freezing
        thread = threading.Thread(target=self._generate_presentation_thread)
        thread.daemon = True
        thread.start()
    
    def _generate_presentation_thread(self):
        """Generate presentation in separate thread"""
        try:
            # Initialize converter
            self.converter = OllamaDocumentToPowerPoint()
            
            # Get current tab
            notebook = self.root.nametowidget(self.root.winfo_children()[0].winfo_children()[1])
            current_tab = notebook.select()
            tab_text = notebook.tab(current_tab, "text")
            
            output_file = None
            title = self.title_var.get() or "AI Generated Presentation"
            
            if "File Processing" in tab_text and self.files_list:
                # Process files
                self.update_progress("Processing files...")
                
                # Categorize files
                word_files = [f for f in self.files_list if f.endswith('.docx')]
                excel_files = [f for f in self.files_list if f.endswith('.xlsx')]
                pptx_files = [f for f in self.files_list if f.endswith('.pptx')]
                pdf_files = [f for f in self.files_list if f.endswith('.pdf')]
                odt_files = [f for f in self.files_list if f.endswith('.odt')]
                txt_files = [f for f in self.files_list if f.endswith('.txt')]
                
                # Process structured files through process_files method
                all_content = self.converter.process_files(
                    word_files=word_files,
                    excel_files=excel_files,
                    pptx_files=pptx_files,
                    pdf_files=pdf_files,
                    odt_files=odt_files
                )
                
                # Process text files by reading their content
                if txt_files:
                    self.update_progress("Processing text files...")
                    text_content = ""
                    for txt_file in txt_files:
                        try:
                            with open(txt_file, 'r', encoding='utf-8') as f:
                                content = f.read()
                                text_content += f"\n\n--- Content from {os.path.basename(txt_file)} ---\n{content}"
                        except Exception as e:
                            print(f"Error reading {txt_file}: {e}")
                    
                    if text_content.strip():
                        # If we have text content and no other files, use text method
                        if not all_content and text_content.strip():
                            self.update_progress("Generating presentation from text...")
                            output_file = self.converter.create_presentation_from_text(
                                user_text=text_content,
                                presentation_title=title,
                                enable_ai_analysis=self.ai_enabled.get()
                            )
                            # Skip the normal processing since we're done
                            self.root.after(0, self._generation_complete, output_file)
                            return
                
                if all_content:
                    self.update_progress("Generating presentation...")
                    output_file = self.converter.create_presentation(
                        all_content=all_content,
                        presentation_title=title
                    )
                else:
                    # No content could be extracted from files
                    self.update_progress("No content found in files...")
                    self.root.after(0, self._generation_error, 
                                   "No content could be extracted from the selected files. Please check that your files contain readable content.")
                    return
                
            elif "Text Input" in tab_text:
                # Process text input
                text_content = self.text_input.get("1.0", tk.END).strip()
                if text_content and "Example: Type your ideas here..." not in text_content:
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
        
        if output_file:
            self.progress_var.set(f"✅ Presentation created: {os.path.basename(output_file)}")
            messagebox.showinfo("Success!", 
                              f"PowerPoint presentation created successfully!\n\n" +
                              f"File: {output_file}\n\n" +
                              "Would you like to open the containing folder?")
            
            # Ask if user wants to open folder
            if messagebox.askyesno("Open Folder?", "Open the folder containing your presentation?"):
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
        app = PowerPointGeneratorGUI()
        app.run()
    except Exception as e:
        print(f"Error starting GUI: {e}")
        print("Make sure all required packages are installed:")
        print("pip install tkinterdnd2")

if __name__ == "__main__":
    main()