#!/usr/bin/env python3
"""
Test script to check text processing capabilities
"""

import sys
import os

# Add current directory to path
sys.path.append('.')

try:
    from wrapper import OllamaDocumentToPowerPoint
    
    # Create instance
    converter = OllamaDocumentToPowerPoint()
    print("✅ Successfully created converter instance")
    
    # Check available methods
    methods = [method for method in dir(converter) if not method.startswith('_')]
    print("\nAvailable methods:")
    for method in methods:
        if 'text' in method.lower() or 'file' in method.lower():
            print(f"  🔍 {method}")
    
    # Test creating presentation from text directly
    print("\n🔍 Testing create_presentation_from_text...")
    try:
        test_text = "This is a test document. We have a project overview and key features."
        result = converter.create_presentation_from_text(
            user_text=test_text,
            presentation_title="Test Presentation",
            enable_ai_analysis=True
        )
        print(f"✅ create_presentation_from_text completed. Result: {result}")
        
    except Exception as e:
        print(f"❌ Error in create_presentation_from_text: {e}")
        import traceback
        traceback.print_exc()
        
    # Test if we can read text files and process them
    print("\n🔍 Testing text file reading...")
    try:
        if os.path.exists("test_document.txt"):
            with open("test_document.txt", "r", encoding="utf-8") as f:
                content = f.read()
            print(f"✅ Read text file: {len(content)} characters")
            
            # Try to create presentation from this content
            result2 = converter.create_presentation_from_text(
                user_text=content,
                presentation_title="Test from File",
                enable_ai_analysis=True
            )
            print(f"✅ Created presentation from file content. Result: {result2}")
        else:
            print("❌ test_document.txt not found")
            
    except Exception as e:
        print(f"❌ Error processing text file: {e}")
        import traceback
        traceback.print_exc()
        
except Exception as e:
    print(f"❌ General error: {e}")
    import traceback
    traceback.print_exc()