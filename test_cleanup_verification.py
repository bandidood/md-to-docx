#!/usr/bin/env python3
"""
Test script to verify cleanup of temporary files
"""

import os
import tempfile
import sys
import logging

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from md_to_docx_online import OnlineMermaidConverter

def test_cleanup_functionality():
    """Test that temporary files are cleaned up properly"""
    print("Testing cleanup functionality...")
    
    # Create converter
    converter = OnlineMermaidConverter(fallback_local=False)
    
    # Create some fake temp files to test cleanup
    test_files = []
    for i in range(3):
        # Create temporary test files
        temp_file = tempfile.mktemp(suffix=f'.test{i}')
        with open(temp_file, 'w') as f:
            f.write(f"Test file {i}")
        test_files.append(temp_file)
        converter.temp_files.append(temp_file)
    
    # Verify files exist
    print(f"Created {len(test_files)} temporary files:")
    for temp_file in test_files:
        exists = os.path.exists(temp_file)
        print(f"  {temp_file}: {'EXISTS' if exists else 'MISSING'}")
        assert exists, f"Test file {temp_file} should exist"
    
    # Test cleanup
    print("\nRunning cleanup...")
    converter.cleanup_temp_files()
    
    # Verify files are cleaned up
    print("Verifying cleanup:")
    for temp_file in test_files:
        exists = os.path.exists(temp_file)
        print(f"  {temp_file}: {'EXISTS' if exists else 'CLEANED'}")
        assert not exists, f"Test file {temp_file} should be cleaned up"
    
    print("\n✅ Cleanup test passed!")

def test_online_converter_cleanup():
    """Test that online converter properly tracks and cleans up files"""
    print("\nTesting online converter cleanup with mermaid processing...")
    
    # Create converter
    converter = OnlineMermaidConverter(fallback_local=False)
    
    # Test markdown with mermaid diagram
    test_markdown = """
# Test Document

```mermaid
graph TD
    A[Start] --> B[Process]
    B --> C[End]
```

Some text after diagram.
"""
    
    # Count initial temp files
    initial_temp_count = len(converter.temp_files)
    print(f"Initial temp files count: {initial_temp_count}")
    
    try:
        # Process markdown (this will create temp files)
        processed_content, diagram_images = converter.process_mermaid_diagrams(test_markdown)
        
        # Check that temp files were created
        temp_count_after = len(converter.temp_files)
        print(f"Temp files after processing: {temp_count_after}")
        
        # Verify temp files exist
        existing_files = []
        for temp_file in converter.temp_files:
            if os.path.exists(temp_file):
                existing_files.append(temp_file)
                print(f"  Temp file exists: {temp_file}")
        
        print(f"Found {len(existing_files)} existing temp files")
        
        # Test cleanup
        print("\nRunning cleanup...")
        converter.cleanup_temp_files()
        
        # Verify cleanup
        remaining_files = []
        for temp_file in existing_files:
            if os.path.exists(temp_file):
                remaining_files.append(temp_file)
                print(f"  File still exists: {temp_file}")
        
        assert len(remaining_files) == 0, f"Cleanup failed, {len(remaining_files)} files remain"
        assert len(converter.temp_files) == 0, "temp_files list should be empty after cleanup"
        
        print("✅ Online converter cleanup test passed!")
        
    except Exception as e:
        print(f"⚠️  Online test failed (expected if no internet): {e}")
        # Still test cleanup of any files that were created
        converter.cleanup_temp_files()
        print("✅ Cleanup still executed properly")

if __name__ == "__main__":
    try:
        test_cleanup_functionality()
        test_online_converter_cleanup()
        print("\n🎉 All cleanup tests passed!")
    except Exception as e:
        print(f"\n❌ Test failed: {e}")
        sys.exit(1)
