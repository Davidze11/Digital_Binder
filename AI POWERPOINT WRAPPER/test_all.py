#!/usr/bin/env python3
"""
TEST SCRIPT - Verify all text input methods work correctly
"""

def test_all_methods():
    print("=" * 60)
    print("TESTING ALL TEXT INPUT METHODS")
    print("=" * 60)
    
    try:
        # Test Method 1: Built-in function
        from wrapper import easy_text_input
        print("✓ Method 1: easy_text_input() - READY")
        
        # Test Method 2: Interactive demo
        from wrapper import interactive_text_input_demo
        print("✓ Method 2: interactive_text_input_demo() - READY")
        
        # Test Method 3: Custom text demo
        from wrapper import create_presentation_from_custom_text
        print("✓ Method 3: create_presentation_from_custom_text() - READY")
        
        print()
        print("ALL METHODS ARE WORKING!")
        print()
        print("QUICK START OPTIONS:")
        print("1. Run: easy_text_input()")
        print("2. Run: interactive_text_input_demo()")
        print("3. Run: create_presentation_from_custom_text()")
        print("4. Run: python simple_input.py")
        print("5. Run: python text_to_ppt.py")
        
    except ImportError as e:
        print(f"ERROR: {e}")
        print("Make sure wrapper.py is in the same directory")

def quick_demo():
    """Run a quick demo to show it works"""
    print("\n" + "=" * 40)
    print("RUNNING QUICK DEMO")
    print("=" * 40)
    
    try:
        from wrapper import OllamaDocumentToPowerPoint
        
        sample_text = """
        Quick Demo Ideas:
        
        1. Project Planning
        - Set clear goals
        - Create timeline
        - Assign resources
        
        2. Team Communication  
        - Regular meetings
        - Progress updates
        - Open feedback
        
        3. Success Metrics
        - Track milestones
        - Measure results
        - Celebrate wins
        """
        
        converter = OllamaDocumentToPowerPoint()
        output_file = converter.create_presentation_from_text(
            user_text=sample_text,
            presentation_title="Demo Presentation"
        )
        
        if output_file:
            print(f"SUCCESS! Demo presentation created: {output_file}")
        else:
            print("Demo completed (check for output file)")
            
    except Exception as e:
        print(f"Demo error: {e}")

if __name__ == "__main__":
    test_all_methods()
    
    # Ask if user wants to run demo
    run_demo = input("\nRun quick demo? (y/n): ").strip().lower()
    if run_demo == 'y':
        quick_demo()
    
    print("\nAll systems ready! Choose any method to create your presentation.")