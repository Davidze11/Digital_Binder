#!/usr/bin/env python3
"""
Test script with correct parameter names
"""

import sys

# Add current directory to path
sys.path.append('.')

try:
    from wrapper import OllamaDocumentToPowerPoint
    
    # Create instance
    converter = OllamaDocumentToPowerPoint()
    print("✅ Successfully created converter instance")
    
    # Test with correct parameter names
    print("\n🔍 Testing with correct parameter names...")
    try:
        result = converter.process_files(
            word_files=[],
            excel_files=[],
            pptx_files=[],
            pdf_files=[],
            odt_files=[]
        )
        print(f"✅ process_files completed. Result: {result}")
        print(f"Result type: {type(result)}")
        if result:
            print(f"Content length: {len(str(result))}")
        else:
            print("❌ Result is None/empty")
            
    except Exception as e:
        print(f"❌ Error in process_files: {e}")
        import traceback
        traceback.print_exc()
        
except Exception as e:
    print(f"❌ General error: {e}")
    import traceback
    traceback.print_exc()