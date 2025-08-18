#!/usr/bin/env python3
"""
Excel to Markdown converter using markitdown library
"""

import sys
from pathlib import Path
from markitdown import MarkItDown


def convert_excel_to_markdown(excel_file_path: str, output_file_path: str = None):
    """
    Convert Excel file to Markdown format
    
    Args:
        excel_file_path: Path to the Excel file
        output_file_path: Optional output file path. If not provided, prints to stdout
    """
    try:
        # Initialize MarkItDown
        md = MarkItDown()
        
        # Convert the Excel file
        result = md.convert(excel_file_path)
        
        if output_file_path:
            # Write to file
            with open(output_file_path, 'w', encoding='utf-8') as f:
                f.write(result.text_content)
            print(f"Converted {excel_file_path} to {output_file_path}")
        else:
            # Print to stdout
            print(result.text_content)
            
    except Exception as e:
        print(f"Error converting {excel_file_path}: {e}", file=sys.stderr)
        sys.exit(1)


def main():
    if len(sys.argv) < 2:
        print("Usage: python excel-to-markdown.py <excel_file> [output_file]")
        print("Example: python excel-to-markdown.py data.xlsx output.md")
        sys.exit(1)
    
    excel_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    # Check if input file exists
    if not Path(excel_file).exists():
        print(f"Error: File {excel_file} not found", file=sys.stderr)
        sys.exit(1)
    
    convert_excel_to_markdown(excel_file, output_file)


if __name__ == "__main__":
    main()