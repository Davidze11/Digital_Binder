#!/usr/bin/env python3
"""
SIMPLE TEXT TO POWERPOINT CONVERTER
Just run this script and type your ideas!
"""

def simple_text_input():
    """Simplified text input for PowerPoint generation"""
    
    print("\n" + "="*50)
    print("🎯 TEXT TO POWERPOINT CONVERTER")
    print("="*50)
    print("Type your content below:")
    print("(Press Enter twice when finished)")
    print("-"*50)
    
    # Get text input
    text_lines = []
    print("Start typing your content:")
    
    while True:
        try:
            line = input()
            if not line.strip():  # Empty line
                break
            text_lines.append(line)
        except (KeyboardInterrupt, EOFError):
            break
    
    if not text_lines:
        print("❌ No text entered!")
        return
    
    user_text = "\n".join(text_lines)
    
    # Get title
    title = input("\n📝 Presentation title: ").strip()
    if not title:
        title = "My Presentation"
    
    # Create presentation
    print(f"\n🚀 Creating presentation '{title}'...")
    
    try:
        # Import and use the converter
        import sys
        import os
        
        # Add current directory to path
        current_dir = os.path.dirname(os.path.abspath(__file__))
        sys.path.insert(0, current_dir)
        
        # Check if wrapper.py exists
        wrapper_path = os.path.join(current_dir, 'wrapper.py')
        if not os.path.exists(wrapper_path):
            print("❌ Error: wrapper.py not found in current directory")
            print(f"Looking for: {wrapper_path}")
            print("Make sure wrapper.py is in the same folder as this script")
            return
        
        # Import the converter
        from wrapper import OllamaDocumentToPowerPoint  # type: ignore
        
        converter = OllamaDocumentToPowerPoint()
        output_file = converter.create_presentation_from_text(
            user_text=user_text,
            presentation_title=title,
            output_file=f"{title.replace(' ', '_')}.pptx"
        )
        
        print(f"✅ SUCCESS! Created: {output_file}")
        
    except ImportError as e:
        print("❌ Error importing wrapper module")
        print(f"Details: {e}")
        print("Make sure wrapper.py is in the same folder as this script")
    except Exception as e:
        print(f"❌ Error creating presentation: {e}")

if __name__ == "__main__":
    simple_text_input()