#!/usr/bin/env python3
"""
Simple Text Input Interface for PowerPoint Generation
Creates presentations from your custom text input - no files needed!
"""

from wrapper import OllamaDocumentToPowerPoint
import os

def main():
    """Main interface for text input and presentation generation"""
    print("=" * 60)
    print("🎯 POWERPOINT GENERATOR - TEXT INPUT MODE")
    print("=" * 60)
    print("Enter your text content below (ideas, notes, plans, etc.)")
    print("When finished, press Enter on an empty line twice")
    print("-" * 60)
    
    # Collect user text input
    lines = []
    empty_line_count = 0
    
    while True:
        try:
            line = input()
            if line.strip() == "":
                empty_line_count += 1
                if empty_line_count >= 2:
                    break
                lines.append(line)
            else:
                empty_line_count = 0
                lines.append(line)
        except KeyboardInterrupt:
            print("\n\n❌ Input cancelled by user")
            return
        except EOFError:
            break
    
    # Join all lines into text
    user_text = "\n".join(lines).strip()
    
    if not user_text:
        print("❌ No text entered. Please try again.")
        return
    
    print(f"\n✅ Captured {len(user_text)} characters of text")
    print("-" * 60)
    
    # Get presentation title
    while True:
        title = input("📝 Enter presentation title: ").strip()
        if title:
            break
        print("Please enter a title for your presentation")
    
    # Get output filename (optional)
    default_filename = f"{title.replace(' ', '_').lower()}.pptx"
    filename = input(f"💾 Output filename (default: {default_filename}): ").strip()
    if not filename:
        filename = default_filename
    
    # Ensure .pptx extension
    if not filename.lower().endswith('.pptx'):
        filename += '.pptx'
    
    print("\n🚀 Creating your presentation...")
    print("-" * 60)
    
    try:
        # Create the presentation
        converter = OllamaDocumentToPowerPoint()
        output_file = converter.create_presentation_from_text(
            user_text=user_text,
            presentation_title=title,
            output_file=filename
        )
        
        print(f"✅ SUCCESS! Presentation created: {output_file}")
        
        # Show file info
        if os.path.exists(output_file):
            file_size = os.path.getsize(output_file) / 1024  # KB
            print(f"📊 File size: {file_size:.1f} KB")
            print(f"📁 Location: {os.path.abspath(output_file)}")
        
        print("\n🎉 Your PowerPoint presentation is ready!")
        
    except Exception as e:
        print(f"❌ ERROR: Failed to create presentation")
        print(f"Details: {str(e)}")
        return

def quick_demo():
    """Quick demo with sample text"""
    print("=" * 60)
    print("🎯 QUICK DEMO - SAMPLE TEXT INPUT")
    print("=" * 60)
    
    sample_text = """
    Project Management Best Practices

    1. Planning Phase
       - Define clear objectives and scope
       - Identify stakeholders and resources
       - Create detailed timeline with milestones
       - Establish communication protocols

    2. Execution Phase
       - Monitor progress regularly
       - Maintain open communication channels
       - Address issues promptly
       - Document decisions and changes

    3. Quality Control
       - Regular quality checkpoints
       - Stakeholder feedback sessions
       - Testing and validation procedures
       - Continuous improvement processes

    4. Risk Management
       - Identify potential risks early
       - Develop mitigation strategies
       - Monitor risk indicators
       - Have contingency plans ready

    Key Success Factors:
    - Strong leadership and clear vision
    - Effective team communication
    - Proper resource allocation
    - Adaptability to change
    - Regular performance reviews
    """
    
    print("Using sample project management text...")
    
    try:
        converter = OllamaDocumentToPowerPoint()
        output_file = converter.create_presentation_from_text(
            user_text=sample_text,
            presentation_title="Project Management Best Practices",
            output_file="demo_project_management.pptx"
        )
        
        print(f"✅ Demo presentation created: {output_file}")
        print("🎉 Check out the generated PowerPoint!")
        
    except Exception as e:
        print(f"❌ Demo failed: {str(e)}")

if __name__ == "__main__":
    print("Choose an option:")
    print("1. Enter your own text")
    print("2. Run quick demo")
    
    while True:
        choice = input("\nEnter choice (1 or 2): ").strip()
        if choice == "1":
            main()
            break
        elif choice == "2":
            quick_demo()
            break
        else:
            print("Please enter 1 or 2")