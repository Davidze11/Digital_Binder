"""
🚀 AI PowerPoint Generator - Interface Launcher
Choose your preferred interface for creating PowerPoint presentations
"""

import os
import sys
import subprocess

def main():
    print("🚀 AI PowerPoint Generator")
    print("=" * 50)
    print()
    print("Choose your preferred interface:")
    print()
    print("1. 🖥️  Desktop GUI (Tkinter)")
    print("   - Native desktop application")
    print("   - File browser and text input")
    print("   - Works offline")
    print()
    print("2. 🌐 Web Interface (Streamlit)")
    print("   - Modern web-based interface")
    print("   - Drag & drop file upload")
    print("   - Runs in your browser")
    print()
    print("3. 💻 Command Line (Original)")
    print("   - Text-based interface")
    print("   - Interactive menu system")
    print("   - Full feature access")
    print()
    print("4. ❌ Exit")
    print()
    
    while True:
        try:
            choice = input("Enter your choice (1-4): ").strip()
            
            if choice == "1":
                print("\n🖥️ Starting Desktop GUI...")
                run_desktop_gui()
                break
                
            elif choice == "2":
                print("\n🌐 Starting Web Interface...")
                run_web_interface()
                break
                
            elif choice == "3":
                print("\n💻 Starting Command Line Interface...")
                run_command_line()
                break
                
            elif choice == "4":
                print("\n👋 Goodbye!")
                break
                
            else:
                print("❌ Invalid choice. Please enter 1, 2, 3, or 4.")
                
        except (KeyboardInterrupt, EOFError):
            print("\n👋 Goodbye!")
            break

def run_desktop_gui():
    """Launch the desktop GUI interface"""
    try:
        # Get the directory where this launcher script is located
        script_dir = os.path.dirname(os.path.abspath(__file__))
        
        # Try the advanced GUI first, fall back to simple version
        gui_script = os.path.join(script_dir, "simple_gui.py")
        advanced_gui = os.path.join(script_dir, "gui_wrapper.py")
        
        if os.path.exists(advanced_gui):
            try:
                import tkinterdnd2
                gui_script = advanced_gui
                print("Using advanced GUI with drag & drop support")
            except ImportError:
                print("Using simple GUI (drag & drop not available)")
        
        subprocess.run([sys.executable, gui_script], check=True)
        
    except subprocess.CalledProcessError as e:
        print(f"❌ Error running GUI: {e}")
        print("Make sure all required packages are installed.")
    except FileNotFoundError:
        print("❌ GUI script not found. Make sure simple_gui.py exists.")

def run_web_interface():
    """Launch the web interface using Streamlit"""
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        web_script = os.path.join(script_dir, "web_gui.py")
        
        print("📱 Opening web interface in your browser...")
        print("   (If browser doesn't open automatically, go to: http://localhost:8501)")
        print("   Press Ctrl+C to stop the web server")
        print()
        
        subprocess.run([
            sys.executable, "-m", "streamlit", "run", web_script,
            "--server.headless", "false",
            "--server.port", "8501"
        ], check=True)
        
    except subprocess.CalledProcessError as e:
        print(f"❌ Error running web interface: {e}")
        print("Make sure Streamlit is installed: pip install streamlit")
    except FileNotFoundError:
        print("❌ Web GUI script not found. Make sure web_gui.py exists.")

def run_command_line():
    """Launch the original command line interface"""
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        wrapper_script = os.path.join(script_dir, "wrapper.py")
        
        subprocess.run([sys.executable, wrapper_script], check=True)
        
    except subprocess.CalledProcessError as e:
        print(f"❌ Error running command line interface: {e}")
    except FileNotFoundError:
        print("❌ Command line script not found. Make sure wrapper.py exists.")

if __name__ == "__main__":
    main()