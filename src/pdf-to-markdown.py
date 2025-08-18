#!/usr/bin/env python3
"""
PDF to Markdown converter using markitdown library
"""

import sys
from pathlib import Path
from markitdown import MarkItDown


def convert_pdf_to_markdown(pdf_file_path: str, output_file_path: str = None):
    """
    Convert PDF file to Markdown format
    
    Args:
        pdf_file_path: Path to the PDF file
        output_file_path: Optional output file path. If not provided, prints to stdout
    """
    try:
        # Initialize MarkItDown
        md = MarkItDown()
        
        # Convert the PDF file
        result = md.convert(pdf_file_path)
        
        if output_file_path:
            # Write to file
            with open(output_file_path, 'w', encoding='utf-8') as f:
                f.write(result.text_content)
            print(f"Converted {pdf_file_path} to {output_file_path}")
        else:
            # Print to stdout
            print(result.text_content)
            
    except Exception as e:
        print(f"Error converting {pdf_file_path}: {e}", file=sys.stderr)
        sys.exit(1)


def main():
    if len(sys.argv) < 2:
        print("Usage: python pdf-to-markdown.py <pdf_file> [output_file]")
        print("Example: python pdf-to-markdown.py document.pdf output.md")
        sys.exit(1)
    
    pdf_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    # Check if input file exists
    if not Path(pdf_file).exists():
        print(f"Error: File {pdf_file} not found", file=sys.stderr)
        sys.exit(1)
    
    convert_pdf_to_markdown(pdf_file, output_file)


if __name__ == "__main__":
    main()