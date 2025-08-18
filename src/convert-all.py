#!/usr/bin/env python3
"""
Batch document to Markdown converter using markitdown library
Supports Excel, Word, PowerPoint, PDF, and text files
"""

import argparse
import sys
from pathlib import Path
from markitdown import MarkItDown


def get_supported_extensions():
    """Return supported file extensions"""
    return {
        '.xlsx': 'Excel',
        '.xls': 'Excel', 
        '.docx': 'Word',
        '.doc': 'Word',
        '.pptx': 'PowerPoint',
        '.ppt': 'PowerPoint',
        '.pdf': 'PDF',
        '.txt': 'Text'
    }


def convert_file_to_markdown(input_file_path: str, output_dir: str):
    """
    Convert a single file to Markdown format
    
    Args:
        input_file_path: Path to the input file
        output_dir: Directory to save the converted markdown file
    
    Returns:
        bool: True if conversion was successful, False otherwise
    """
    try:
        # Initialize MarkItDown
        md = MarkItDown()
        
        # Convert the file
        result = md.convert(input_file_path)
        
        # Generate output filename
        input_path = Path(input_file_path)
        output_filename = f"{input_path.stem}.md"
        output_file_path = Path(output_dir) / output_filename
        
        # Write to file
        with open(output_file_path, 'w', encoding='utf-8') as f:
            f.write(result.text_content)
        
        print(f" Converted: {input_file_path} � {output_file_path}")
        return True
        
    except Exception as e:
        print(f"L Error converting {input_file_path}: {e}")
        return False


def process_directory(directory_path: str, output_dir: str):
    """
    Process all supported files in a directory
    
    Args:
        directory_path: Path to the directory containing files
        output_dir: Directory to save converted markdown files
    
    Returns:
        tuple: (success_count, total_count)
    """
    supported_extensions = get_supported_extensions()
    success_count = 0
    total_count = 0
    
    directory = Path(directory_path)
    
    if not directory.exists():
        print(f"L Directory not found: {directory_path}")
        return 0, 0
    
    if not directory.is_dir():
        print(f"L Path is not a directory: {directory_path}")
        return 0, 0
    
    # Find all supported files
    supported_files = []
    for file_path in directory.rglob('*'):
        if file_path.is_file() and file_path.suffix.lower() in supported_extensions:
            supported_files.append(file_path)
    
    if not supported_files:
        print(f"�  No supported files found in: {directory_path}")
        print(f"Supported extensions: {', '.join(supported_extensions.keys())}")
        return 0, 0
    
    print(f"=� Found {len(supported_files)} supported files in: {directory_path}")
    
    # Process each file
    for file_path in supported_files:
        total_count += 1
        file_type = supported_extensions[file_path.suffix.lower()]
        print(f"= Processing {file_type} file: {file_path}")
        
        if convert_file_to_markdown(str(file_path), output_dir):
            success_count += 1
    
    return success_count, total_count


def process_single_file(file_path: str, output_dir: str):
    """
    Process a single file
    
    Args:
        file_path: Path to the file
        output_dir: Directory to save converted markdown file
    
    Returns:
        bool: True if conversion was successful, False otherwise
    """
    supported_extensions = get_supported_extensions()
    
    file = Path(file_path)
    
    if not file.exists():
        print(f"L File not found: {file_path}")
        return False
    
    if not file.is_file():
        print(f"L Path is not a file: {file_path}")
        return False
    
    if file.suffix.lower() not in supported_extensions:
        print(f"L Unsupported file type: {file.suffix}")
        print(f"Supported extensions: {', '.join(supported_extensions.keys())}")
        return False
    
    file_type = supported_extensions[file.suffix.lower()]
    print(f"= Processing {file_type} file: {file_path}")
    
    return convert_file_to_markdown(file_path, output_dir)


def main():
    parser = argparse.ArgumentParser(
        description="Convert documents (Excel, Word, PowerPoint, PDF, Text) to Markdown",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Convert all files in a directory
  python convert-all.py --directorypath src/media
  
  # Convert a single file
  python convert-all.py --filepath src/media/document.pdf
  
  # Specify custom output directory
  python convert-all.py --directorypath src/media --output custom-output
        """
    )
    
    # Create mutually exclusive group for directory or file path
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument(
        '--directorypath', 
        type=str,
        help='Directory path containing documents to convert'
    )
    group.add_argument(
        '--filepath',
        type=str, 
        help='Single file path to convert'
    )
    
    parser.add_argument(
        '--output',
        type=str,
        default='converted-markdown',
        help='Output directory for converted markdown files (default: converted-markdown)'
    )
    
    args = parser.parse_args()
    
    # Create output directory
    output_dir = Path(args.output)
    output_dir.mkdir(exist_ok=True)
    print(f"=� Output directory: {output_dir.absolute()}")
    
    # Display supported file types
    supported_extensions = get_supported_extensions()
    print(f"=' Supported file types: {', '.join(f'{ext} ({ftype})' for ext, ftype in supported_extensions.items())}")
    print()
    
    if args.directorypath:
        # Process directory
        success_count, total_count = process_directory(args.directorypath, str(output_dir))
        
        print()
        print("=" * 50)
        print(f"=� Conversion Summary:")
        print(f"   Total files processed: {total_count}")
        print(f"   Successful conversions: {success_count}")
        print(f"   Failed conversions: {total_count - success_count}")
        
        if total_count > 0:
            success_rate = (success_count / total_count) * 100
            print(f"   Success rate: {success_rate:.1f}%")
        
        sys.exit(0 if success_count == total_count else 1)
        
    elif args.filepath:
        # Process single file
        success = process_single_file(args.filepath, str(output_dir))
        
        print()
        print("=" * 50)
        if success:
            print(" File conversion completed successfully!")
            sys.exit(0)
        else:
            print("L File conversion failed!")
            sys.exit(1)


if __name__ == "__main__":
    main()