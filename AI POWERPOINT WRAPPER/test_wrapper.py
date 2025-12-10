#!/usr/bin/env python3
"""
Test script to debug wrapper file processing
"""

import sys
import os

# Add current directory to path
sys.path.append('.')

try:
    from wrapper import OllamaDocumentToPowerPoint
    print("✅ Successfully imported OllamaDocumentToPowerPoint")
    
    # Create instance
    converter = OllamaDocumentToPowerPoint()
    print("✅ Successfully created converter instance")
    
    # Check if process_files method exists
    if hasattr(converter, 'process_files'):
        print("✅ process_files method exists")
        print(f"Method signature: {converter.process_files.__doc__}")
    else:
        print("❌ process_files method does NOT exist")
        print("Available methods:")
        methods = [method for method in dir(converter) if not method.startswith('_')]
        for method in methods:
            print(f"  - {method}")
    
    # Check if create_presentation method exists
    if hasattr(converter, 'create_presentation'):
        print("✅ create_presentation method exists")
    else:
        print("❌ create_presentation method does NOT exist")
        
except ImportError as e:
    print(f"❌ Import error: {e}")
except Exception as e:
    print(f"❌ Error: {e}")