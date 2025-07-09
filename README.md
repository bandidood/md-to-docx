# Markdown to DOCX Converter with Mermaid Support

This script converts Markdown files to DOCX format with embedded Mermaid diagrams.

## Prerequisites

Ensure the following are installed:

- Python 3.x
- `python-docx`
- `markdown`
- `requests`
- `pillow`
- Node.js
- Mermaid CLI (`mmdc`)

## Usage

```bash
python md_to_docx.py <input_markdown_file> <output_docx_file>
```

The script processes any Mermaid diagrams found in the Markdown and embeds them as images.
