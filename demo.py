#!/usr/bin/env python3
"""
Demo script showing how to use the Markdown to DOCX converter
"""

import os
from md_to_docx import MarkdownToDocxConverter

def demo_conversion():
    """Demonstrate the converter usage"""
    print("Markdown to DOCX Converter Demo")
    print("=" * 40)
    
    # List available markdown files
    md_files = [f for f in os.listdir('.') if f.endswith('.md')]
    
    if not md_files:
        print("No Markdown files found in current directory")
        return
    
    print("Available Markdown files:")
    for i, file in enumerate(md_files, 1):
        print(f"{i}. {file}")
    
    try:
        choice = int(input("\nSelect a file to convert (number): ")) - 1
        if 0 <= choice < len(md_files):
            input_file = md_files[choice]
            output_file = input_file.replace('.md', '_converted.docx')
            
            print(f"\nConverting {input_file} to {output_file}...")
            
            converter = MarkdownToDocxConverter()
            success = converter.convert(input_file, output_file)
            
            if success:
                print(f"\n✅ Successfully converted to {output_file}")
                print(f"📁 File size: {os.path.getsize(output_file)} bytes")
            else:
                print("❌ Conversion failed")
        else:
            print("Invalid selection")
            
    except ValueError:
        print("Please enter a valid number")
    except KeyboardInterrupt:
        print("\nOperation cancelled")

if __name__ == "__main__":
    demo_conversion()
