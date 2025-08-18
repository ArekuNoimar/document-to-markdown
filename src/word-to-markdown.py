#!/usr/bin/env python3
"""
Word to Markdown converter using markitdown library
"""

import sys
from pathlib import Path
from markitdown import MarkItDown


def convert_word_to_markdown(docx_file_path: str, output_file_path: str = None):
    """
    Convert Word file to Markdown format
    
    Args:
        docx_file_path: Path to the Word file
        output_file_path: Optional output file path. If not provided, prints to stdout
    """
    try:
        # Initialize MarkItDown
        md = MarkItDown()
        
        # Convert the Word file
        result = md.convert(docx_file_path)
        
        if output_file_path:
            # Write to file
            with open(output_file_path, 'w', encoding='utf-8') as f:
                f.write(result.text_content)
            print(f"Converted {docx_file_path} to {output_file_path}")
        else:
            # Print to stdout
            print(result.text_content)
            
    except Exception as e:
        print(f"Error converting {docx_file_path}: {e}", file=sys.stderr)
        sys.exit(1)


def main():
    if len(sys.argv) < 2:
        print("Usage: python word-to-markdown.py <docx_file> [output_file]")
        print("Example: python word-to-markdown.py document.docx output.md")
        sys.exit(1)
    
    docx_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    # Check if input file exists
    if not Path(docx_file).exists():
        print(f"Error: File {docx_file} not found", file=sys.stderr)
        sys.exit(1)
    
    convert_word_to_markdown(docx_file, output_file)


if __name__ == "__main__":
    main()