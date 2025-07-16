#!/usr/bin/env python3
"""
Markdown to DOCX Converter with Online Mermaid Support
"""

import requests
import base64
import zlib
import urllib.parse
from io import BytesIO
from PIL import Image
import tempfile
import os
import re
import sys
import logging
import subprocess

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Optional SVG to PNG conversion support
# PIL cannot open SVG, so we need cairosvg for conversion
try:
    import cairosvg
    HAS_CAIROSVG = True
except ImportError:
    HAS_CAIROSVG = False
    logger.warning("cairosvg not available. SVG diagrams will be skipped if PNG format fails.")

# Hériter de la classe principale et remplacer la méthode de génération Mermaid
from md_to_docx import MarkdownToDocxConverter as BaseConverter, find_mmdc

# Custom exception for online generation failures
class OnlineMermaidError(Exception):
    """Exception raised when online Mermaid generation fails"""
    pass

class OnlineMermaidConverter(BaseConverter):
    def __init__(self, fallback_local=False):
        super().__init__()
        # Pas besoin de mmdc pour la version en ligne
        self.mmdc_cmd = "online"
        self.fallback_local = fallback_local
        self.local_mmdc_cmd = None
        
        # If fallback is enabled, try to find local mmdc
        if fallback_local:
            self.local_mmdc_cmd = find_mmdc()
            if self.local_mmdc_cmd:
                logger.info(f"Local mmdc found at: {self.local_mmdc_cmd}")
            else:
                logger.warning("Local mmdc not found. Fallback will not be available.")
        
    def generate_mermaid_image(self, mermaid_code, output_path):
        """Generate PNG image from Mermaid code using online service with fallback"""
        try:
            # Try online generation first
            if self._generate_online(mermaid_code, output_path):
                return True
            
            # If online fails and fallback is enabled, try local mmdc
            if self.fallback_local and self.local_mmdc_cmd:
                logger.info("Online generation failed, trying local mmdc fallback...")
                return self._generate_local_fallback(mermaid_code, output_path)
            
            # If no fallback or fallback also fails, raise exception
            error_msg = "Online Mermaid generation failed"
            if self.fallback_local:
                error_msg += " and local fallback is not available"
            
            logger.error(error_msg)
            raise OnlineMermaidError(error_msg)
            
        except OnlineMermaidError:
            # Re-raise our custom exception
            raise
        except Exception as e:
            logger.error(f"Unexpected error in generate_mermaid_image: {e}")
            raise OnlineMermaidError(f"Unexpected error: {e}")
    
    def _generate_online(self, mermaid_code, output_path):
        """Generate image using online mermaid.ink service"""
        try:
            logger.debug(f"Generating Mermaid image online for code: {mermaid_code[:50]}...")
            
            # Use deflate and URL-safe base64 encoding (pako compatible)
            deflated = zlib.compress(mermaid_code.encode('utf-8'))[2:-4]  # Skip zlib headers and checksum
            encoded_mermaid = base64.urlsafe_b64encode(deflated).decode('utf-8').rstrip('=')
            
            # Try PNG first, fallback to SVG if needed
            formats_to_try = [
                ('png', f"https://mermaid.ink/img/pako:{encoded_mermaid}?type=png"),
                ('svg', f"https://mermaid.ink/svg/pako:{encoded_mermaid}")
            ]
            
            # If cairosvg is not available, only try PNG
            if not HAS_CAIROSVG:
                formats_to_try = [('png', f"https://mermaid.ink/img/pako:{encoded_mermaid}?type=png")]
            
            for format_type, url in formats_to_try:
                logger.debug(f"Trying {format_type} format: {url}")
                
                try:
                    response = requests.get(url, timeout=30)
                    
                    logger.debug(f"HTTP status: {response.status_code}, Content-Type: {response.headers.get('Content-Type', 'Unknown')}")
                    
                    response.raise_for_status()
                    
                    # Handle SVG to PNG conversion if needed
                    if format_type == 'svg':
                        if not HAS_CAIROSVG:
                            logger.debug("cairosvg not available, skipping SVG format")
                            continue
                        
                        logger.debug("Converting SVG to PNG")
                        png_data = self._convert_svg_to_png(response.content)
                        
                        # Save PNG data
                        with open(output_path, 'wb') as f:
                            f.write(png_data)
                    else:
                        # Save the response content directly
                        with open(output_path, 'wb') as f:
                            f.write(response.content)
                    
                    # Verify that the image is valid
                    if self._verify_image(output_path):
                        logger.info(f"Successfully generated Mermaid image using {format_type} format")
                        return True
                        
                except requests.RequestException as e:
                    logger.debug(f"Request failed for {format_type}: {e}")
                    continue
            
            logger.warning("All online formats failed")
            return False
            
        except Exception as e:
            logger.error(f"Error in online generation: {e}")
            return False
    
    def _convert_svg_to_png(self, svg_content):
        """Convert SVG content to PNG using cairosvg"""
        try:
            # Get SVG dimensions for proper scaling
            svg_text = svg_content.decode('utf-8')
            width_match = re.search(r'width=["\']([^"\']*)["\']', svg_text)
            height_match = re.search(r'height=["\']([^"\']*)["\']', svg_text)
            
            # Extract numeric values from width/height
            width = None
            height = None
            if width_match:
                width_str = width_match.group(1)
                width_num = re.search(r'(\d+(?:\.\d+)?)', width_str)
                if width_num:
                    width = float(width_num.group(1))
            
            if height_match:
                height_str = height_match.group(1)
                height_num = re.search(r'(\d+(?:\.\d+)?)', height_str)
                if height_num:
                    height = float(height_num.group(1))
            
            # Convert SVG to PNG with proper dimensions
            if width and height:
                logger.debug(f"SVG dimensions: {width}x{height}")
                return cairosvg.svg2png(bytestring=svg_content, 
                                      output_width=width, 
                                      output_height=height)
            else:
                logger.debug("SVG dimensions not found, using default")
                return cairosvg.svg2png(bytestring=svg_content)
                
        except Exception as e:
            logger.debug(f"SVG parsing error: {e}, using fallback conversion")
            return cairosvg.svg2png(bytestring=svg_content)
    
    def _verify_image(self, image_path):
        """Verify that the generated image is valid"""
        try:
            with Image.open(image_path) as img:
                # Check that image has reasonable size
                if img.size[0] > 50 and img.size[1] > 50:
                    logger.debug(f"Image verification successful: {img.size}")
                    return True
                else:
                    logger.debug(f"Image too small: {img.size}")
                    return False
        except Exception as e:
            logger.debug(f"Image verification failed: {e}")
            return False
    
    def _generate_local_fallback(self, mermaid_code, output_path):
        """Generate image using local mmdc as fallback"""
        try:
            logger.info("Using local mmdc fallback")
            
            # Create temporary file for mermaid code with UTF-8 encoding
            with tempfile.NamedTemporaryFile(mode='w', suffix='.mmd', delete=False, encoding='utf-8') as temp_mmd:
                temp_mmd.write(mermaid_code)
                temp_mmd_path = temp_mmd.name
                self.temp_files.append(temp_mmd_path)
            
            # Run mmdc to generate PNG
            cmd = [
                self.local_mmdc_cmd,
                '-i', temp_mmd_path,
                '-o', output_path,
                '-t', 'default',  # default theme (colored)
                '-b', 'white',    # white background
                '--width', '1200',
                '--height', '800'
            ]
            
            logger.debug(f"Executing local mmdc command: {' '.join(cmd)}")
            
            result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
            
            if result.returncode != 0:
                logger.error(f"Local mmdc failed: {result.stderr}")
                if result.stdout:
                    logger.debug(f"mmdc stdout: {result.stdout}")
                return False
            
            if os.path.exists(output_path) and self._verify_image(output_path):
                logger.info("Successfully generated image using local mmdc fallback")
                return True
            else:
                logger.error("Local mmdc did not generate valid image")
                return False
                
        except Exception as e:
            logger.error(f"Error in local fallback: {e}")
            return False
    
    def process_mermaid_diagrams(self, markdown_content):
        """Process and replace Mermaid diagrams with placeholders, with proper error handling"""
        diagrams = self.extract_mermaid_diagrams(markdown_content)
        
        if not diagrams:
            return markdown_content, []
            
        logger.debug(f"Extracted {len(diagrams)} diagrams")

        # Generate images for each diagram
        diagram_images = []
        processed_content = markdown_content
        
        # Process diagrams in reverse order to maintain correct indices
        for diagram in reversed(diagrams):
            # Create temporary image file
            image_path = tempfile.mktemp(suffix='.png')
            self.temp_files.append(image_path)
            
            try:
                if self.generate_mermaid_image(diagram['code'], image_path):
                    diagram_images.insert(0, {
                        'path': image_path,
                        'index': diagram['index']
                    })

                    logger.debug(f"Successfully generated image for diagram {diagram['index']}")
                    
                    # Replace mermaid code block with placeholder
                    placeholder = f"[MERMAID_DIAGRAM_{diagram['index']}]"
                    processed_content = (processed_content[:diagram['start']] + 
                                       placeholder + 
                                       processed_content[diagram['end']:])
                else:
                    logger.warning(f"Failed to generate image for diagram {diagram['index']}")
                    
            except OnlineMermaidError as e:
                logger.error(f"Mermaid generation failed for diagram {diagram['index']}: {e}")
                # Re-raise to propagate the error up
                raise
            except Exception as e:
                logger.error(f"Unexpected error processing diagram {diagram['index']}: {e}")
                raise OnlineMermaidError(f"Unexpected error processing diagram {diagram['index']}: {e}")
                
        return processed_content, diagram_images

def main():
    import argparse
    import sys
    
    parser = argparse.ArgumentParser(description='Convert Markdown to DOCX with Online Mermaid support')
    parser.add_argument('input', help='Input Markdown file')
    parser.add_argument('output', help='Output DOCX file')
    parser.add_argument('--fallback-local', action='store_true', 
                       help='Enable local mmdc fallback if online generation fails')
    parser.add_argument('--verbose', '-v', action='store_true',
                       help='Enable verbose logging')
    
    args = parser.parse_args()
    
    # Configure logging level based on verbose flag
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    
    # Check if input file exists
    if not os.path.exists(args.input):
        logger.error(f"Input file '{args.input}' not found")
        sys.exit(1)
    
    logger.info("Using online Mermaid service (mermaid.ink)")
    if args.fallback_local:
        logger.info("Local mmdc fallback enabled")
    
    try:
        # Convert file
        converter = OnlineMermaidConverter(fallback_local=args.fallback_local)
        success = converter.convert(args.input, args.output)
        
        if success:
            logger.info(f"Successfully converted '{args.input}' to '{args.output}'")
        else:
            logger.error("Conversion failed")
            sys.exit(1)
            
    except OnlineMermaidError as e:
        logger.error(f"Mermaid generation error: {e}")
        sys.exit(1)
    except Exception as e:
        logger.error(f"Unexpected error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
