#!/usr/bin/env python3
"""
Test script to find correct parameter names
"""

import sys
import inspect

# Add current directory to path
sys.path.append('.')

try:
    from wrapper import OllamaDocumentToPowerPoint
    
    # Create instance
    converter = OllamaDocumentToPowerPoint()
    
    # Get the signature of process_files method
    sig = inspect.signature(converter.process_files)
    print("process_files method parameters:")
    for param_name, param in sig.parameters.items():
        print(f"  - {param_name}: {param}")
        
except Exception as e:
    print(f"Error: {e}")
    import traceback
    traceback.print_exc()