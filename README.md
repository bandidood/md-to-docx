# Markdown to DOCX Converter with Mermaid Support

This script converts Markdown files to DOCX format with embedded Mermaid diagrams.

## Prerequisites

Ensure the following are installed:

- Python 3.x
- `python-docx`
- `markdown`
- `requests`
- `pillow`
- `cairosvg` (for SVG to PNG conversion in online mode)
- Node.js (for local mode)
- Mermaid CLI (`mmdc`) (for local mode)

### Installation

```bash
pip install python-docx markdown requests pillow cairosvg
```

For local mode, also install Mermaid CLI:
```bash
npm install -g @mermaid-js/mermaid-cli
```

## Usage

### Local Mode (Default)

Uses locally installed Mermaid CLI:

```bash
python md_to_docx.py <input_markdown_file> <output_docx_file>
```

### Online Mode

Uses the online mermaid.ink service with optional local fallback:

```bash
python md_to_docx_online.py <input_markdown_file> <output_docx_file>
```

With local fallback enabled:

```bash
python md_to_docx_online.py --fallback-local <input_markdown_file> <output_docx_file>
```

## Modes Comparison

### Local Mode (`md_to_docx.py`)
- **Pros**: Full control, works offline, supports all Mermaid features
- **Cons**: Requires Node.js and Mermaid CLI installation
- **Requirements**: Node.js, @mermaid-js/mermaid-cli

### Online Mode (`md_to_docx_online.py`)
- **Pros**: No Node.js required, automatic fallback options
- **Cons**: Requires internet connection, depends on external service
- **Requirements**: Internet connection, cairosvg (for SVG conversion)
- **Features**: 
  - Tries PNG format first, falls back to SVG with cairosvg conversion
  - Optional local fallback if online service fails
  - Handles SVG to PNG conversion automatically

The script processes any Mermaid diagrams found in the Markdown and embeds them as images.
