#!/usr/bin/env python3
"""
Test script to debug file processing with sample files
"""

import sys
import os
import traceback

# Add current directory to path
sys.path.append('.')

try:
    from wrapper import OllamaDocumentToPowerPoint
    print("✅ Successfully imported OllamaDocumentToPowerPoint")
    
    # Create instance
    converter = OllamaDocumentToPowerPoint()
    print("✅ Successfully created converter instance")
    
    # Test with empty file lists (like when no files are selected)
    print("\n🔍 Testing with empty file lists...")
    try:
        result = converter.process_files(
            docx_files=[],
            xlsx_files=[],
            pptx_files=[],
            pdf_files=[],
            txt_files=[],
            odt_files=[]
        )
        print(f"✅ process_files completed. Result: {result}")
        print(f"Result type: {type(result)}")
        if result:
            print(f"Content length: {len(str(result))}")
        else:
            print("❌ Result is None/empty - this might be the issue!")
            
    except Exception as e:
        print(f"❌ Error in process_files: {e}")
        traceback.print_exc()
    
    # Test create_presentation method
    print("\n🔍 Testing create_presentation with empty content...")
    try:
        result2 = converter.create_presentation(
            all_content="",
            presentation_title="Test Presentation"
        )
        print(f"✅ create_presentation completed. Result: {result2}")
        
    except Exception as e:
        print(f"❌ Error in create_presentation: {e}")
        traceback.print_exc()
        
except Exception as e:
    print(f"❌ General error: {e}")
    traceback.print_exc()