#!/usr/bin/env python3
"""
STANDALONE TEXT INPUT SCRIPT
Double-click this file or run: python text_to_ppt.py
"""

def main():
    print("=" * 50)
    print("TEXT TO POWERPOINT CONVERTER")
    print("=" * 50)
    print()
    print("Type your content below:")
    print("Press Enter twice when finished typing")
    print()
    
    # Get text input
    lines = []
    empty_lines = 0
    
    print("Start typing your content:")
    while True:
        try:
            line = input()
            if not line.strip():
                empty_lines += 1
                if empty_lines >= 2:
                    break
                lines.append(line)
            else:
                empty_lines = 0
                lines.append(line)
        except (KeyboardInterrupt, EOFError):
            break
    
    # Clean up the text
    text = "\n".join(lines).strip()
    
    if not text:
        print("No text entered. Goodbye!")
        return
    
    print(f"\nCaptured {len(text)} characters of text.")
    
    # Get presentation title
    title = input("Enter presentation title: ").strip()
    if not title:
        title = "My Presentation"
    
    # Create the presentation
    print(f"\nCreating presentation '{title}'...")
    
    try:
        # Try to import from the same directory
        import sys
        import os
        
        # Add current directory to path
        current_dir = os.path.dirname(os.path.abspath(__file__))
        sys.path.insert(0, current_dir)
        
        # Check if wrapper.py exists
        wrapper_path = os.path.join(current_dir, 'wrapper.py')
        if not os.path.exists(wrapper_path):
            print(f"ERROR: wrapper.py not found at {wrapper_path}")
            print("Make sure this script is in the same folder as wrapper.py")
            input("Press Enter to exit...")
            return
        
        from wrapper import OllamaDocumentToPowerPoint  # type: ignore
        
        converter = OllamaDocumentToPowerPoint()
        output_file = converter.create_presentation_from_text(
            user_text=text,
            presentation_title=title
        )
        
        print(f"SUCCESS! Created: {output_file}")
        input("Press Enter to exit...")
        
    except ImportError:
        print("ERROR: Could not find wrapper.py")
        print("Make sure this script is in the same folder as wrapper.py")
        input("Press Enter to exit...")
    except Exception as e:
        print(f"ERROR: {e}")
        input("Press Enter to exit...")

if __name__ == "__main__":
    main()