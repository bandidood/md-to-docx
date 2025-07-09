#!/usr/bin/env python3
"""
Web Application for Markdown to DOCX Converter
"""

import os
import tempfile
import uuid
from datetime import datetime, timedelta
from pathlib import Path

from flask import Flask, render_template, request, send_file, jsonify, flash, redirect, url_for
from werkzeug.utils import secure_filename
import markdown

# Import our existing converter
from md_to_docx import MarkdownToDocxConverter

app = Flask(__name__)
app.secret_key = 'your-secret-key-change-this'  # Change this in production

# Configuration
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
ALLOWED_EXTENSIONS = {'md', 'txt', 'markdown'}
MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB max file size

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH

# Create directories if they don't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def allowed_file(filename):
    """Check if file extension is allowed"""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def cleanup_old_files():
    """Clean up files older than 1 hour"""
    cutoff_time = datetime.now() - timedelta(hours=1)
    
    for folder in [UPLOAD_FOLDER, OUTPUT_FOLDER]:
        for file_path in Path(folder).glob('*'):
            if file_path.is_file():
                file_time = datetime.fromtimestamp(file_path.stat().st_mtime)
                if file_time < cutoff_time:
                    try:
                        file_path.unlink()
                    except Exception as e:
                        print(f"Error deleting old file {file_path}: {e}")

@app.route('/')
def index():
    """Main page with upload form and text area"""
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert_markdown():
    """Convert markdown to DOCX"""
    cleanup_old_files()
    
    markdown_content = ""
    input_filename = "document"
    
    # Check if file was uploaded
    if 'file' in request.files and request.files['file'].filename != '':
        file = request.files['file']
        
        if file and allowed_file(file.filename):
            # Save uploaded file
            filename = secure_filename(file.filename)
            input_filename = Path(filename).stem
            unique_filename = f"{uuid.uuid4().hex}_{filename}"
            input_path = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
            file.save(input_path)
            
            # Read file content
            try:
                with open(input_path, 'r', encoding='utf-8') as f:
                    markdown_content = f.read()
            except Exception as e:
                flash(f'Error reading uploaded file: {str(e)}', 'error')
                return redirect(url_for('index'))
        else:
            flash('Invalid file type. Please upload a .md, .txt, or .markdown file.', 'error')
            return redirect(url_for('index'))
    
    # Check if text was provided in textarea
    elif 'markdown_text' in request.form and request.form['markdown_text'].strip():
        markdown_content = request.form['markdown_text']
        input_filename = "pasted_content"
    
    else:
        flash('Please provide markdown content either by uploading a file or pasting text.', 'error')
        return redirect(url_for('index'))
    
    if not markdown_content.strip():
        flash('The provided markdown content is empty.', 'error')
        return redirect(url_for('index'))
    
    # Create temporary input file
    temp_input = tempfile.NamedTemporaryFile(mode='w', suffix='.md', delete=False, encoding='utf-8')
    temp_input.write(markdown_content)
    temp_input.close()
    
    # Generate output filename
    output_filename = f"{input_filename}_{uuid.uuid4().hex[:8]}.docx"
    output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
    
    try:
        # Convert using our existing converter
        converter = MarkdownToDocxConverter()
        success = converter.convert(temp_input.name, output_path)
        
        if success and os.path.exists(output_path):
            # Return the file for download
            return send_file(
                output_path,
                as_attachment=True,
                download_name=f"{input_filename}.docx",
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
        else:
            flash('Error converting markdown to DOCX. Please check your markdown syntax and try again.', 'error')
            return redirect(url_for('index'))
            
    except Exception as e:
        flash(f'Error during conversion: {str(e)}', 'error')
        return redirect(url_for('index'))
    
    finally:
        # Clean up temporary input file
        try:
            os.unlink(temp_input.name)
        except Exception:
            pass

@app.route('/preview', methods=['POST'])
def preview_markdown():
    """Generate HTML preview of markdown content"""
    try:
        markdown_content = request.json.get('content', '')
        
        if not markdown_content.strip():
            return jsonify({'html': '<p><em>No content to preview</em></p>'})
        
        # Convert markdown to HTML for preview
        md = markdown.Markdown(extensions=['tables', 'fenced_code', 'codehilite'])
        html_content = md.convert(markdown_content)
        
        # Add some basic styling classes
        html_content = html_content.replace('<table>', '<table class="table table-striped">')
        html_content = html_content.replace('<code>', '<code class="bg-light">')
        
        return jsonify({'html': html_content})
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/health')
def health_check():
    """Health check endpoint"""
    return jsonify({'status': 'ok', 'timestamp': datetime.now().isoformat()})

if __name__ == '__main__':
    # Check if mmdc is available
    from md_to_docx import find_mmdc
    mmdc_cmd = find_mmdc()
    if not mmdc_cmd:
        print("Warning: Mermaid CLI (mmdc) not found. Mermaid diagrams will not be processed.")
        print("To install: npm install -g @mermaid-js/mermaid-cli")
    else:
        print(f"Mermaid CLI found at: {mmdc_cmd}")
    
    print("Starting Markdown to DOCX Converter Web App...")
    print("Access the application at: http://localhost:5000")
    
    app.run(debug=True, host='0.0.0.0', port=5000)
