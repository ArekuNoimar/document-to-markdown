#!/usr/bin/env python3
"""
Batch document to Markdown converter using markitdown library
Supports Excel, Word, PowerPoint, PDF, text files, and comprehensive programming language source code files
Converts 50+ programming languages to Markdown with syntax highlighting
"""

import argparse
import sys
import threading
from pathlib import Path
from markitdown import MarkItDown
from tqdm import tqdm


def get_supported_extensions():
    """Return supported file extensions"""
    document_extensions = {
        '.xlsx': 'Excel',
        '.xls': 'Excel', 
        '.docx': 'Word',
        '.doc': 'Word',
        '.pptx': 'PowerPoint',
        '.ppt': 'PowerPoint',
        '.pdf': 'PDF',
        '.txt': 'Text'
    }
    
    # Comprehensive programming language extensions
    programming_extensions = {
        '.abp': 'ABAP',
        '.as': 'ActionScript',
        '.asm': 'Assembly',
        '.bat': 'Batch',
        '.c': 'C',
        '.cc': 'C++',
        '.clj': 'Clojure',
        '.coffee': 'CoffeeScript',
        '.cpp': 'C++',
        '.cs': 'C#',
        '.css': 'CSS',
        '.cxx': 'C++',
        '.d': 'D',
        '.dart': 'Dart',
        '.erl': 'Erlang',
        '.forth': 'Forth',
        '.go': 'Go',
        '.groovy': 'Groovy',
        '.h': 'C/C++ Header',
        '.hpp': 'C++ Header',
        '.hs': 'Haskell',
        '.htm': 'HTML',
        '.html': 'HTML',
        '.hx': 'Haxe',
        '.ipynb': 'Jupyter Notebook',
        '.java': 'Java',
        '.js': 'JavaScript',
        '.jsx': 'JSX',
        '.kt': 'Kotlin',
        '.kts': 'Kotlin Script',
        '.lhs': 'Literate Haskell',
        '.lisp': 'Lisp',
        '.lsl': 'LSL',
        '.lua': 'Lua',
        '.m': 'MATLAB/Objective-C',
        '.mat': 'MATLAB',
        '.mjs': 'JavaScript Module',
        '.ml': 'OCaml',
        '.pas': 'Pascal',
        '.php': 'PHP',
        '.pl': 'Perl',
        '.pm': 'Perl Module',
        '.pro': 'Prolog',
        '.ps1': 'PowerShell',
        '.py': 'Python',
        '.pyc': 'Python Compiled',
        '.pyo': 'Python Optimized',
        '.r': 'R',
        '.rb': 'Ruby',
        '.rs': 'Rust',
        '.scala': 'Scala',
        '.scm': 'Scheme',
        '.sh': 'Shell Script',
        '.sql': 'SQL',
        '.swift': 'Swift',
        '.swi': 'SWI-Prolog',
        '.ts': 'TypeScript',
        '.v': 'Verilog',
        '.vbs': 'VBScript',
        '.xhtml': 'XHTML',
        '.xml': 'XML',
        '.xquery': 'XQuery'
    }
    
    # Combine document and programming extensions
    all_extensions = {**document_extensions, **programming_extensions}
    return all_extensions


def get_source_code_extensions():
    """Return source code file extensions"""
    return {
        '.abp', '.as', '.asm', '.bat', '.c', '.cc', '.clj', '.coffee', '.cpp', 
        '.cs', '.css', '.cxx', '.d', '.dart', '.erl', '.forth', '.go', '.groovy',
        '.h', '.hpp', '.hs', '.htm', '.html', '.hx', '.ipynb', '.java', '.js', 
        '.jsx', '.kt', '.kts', '.lhs', '.lisp', '.lsl', '.lua', '.m', '.mat',
        '.mjs', '.ml', '.pas', '.php', '.pl', '.pm', '.pro', '.ps1', '.py', 
        '.pyc', '.pyo', '.r', '.rb', '.rs', '.scala', '.scm', '.sh', '.sql',
        '.swift', '.swi', '.ts', '.v', '.vbs', '.xhtml', '.xml', '.xquery'
    }


def convert_source_code_to_markdown(input_file_path: str, output_dir: str):
    """
    Convert a source code file to Markdown format with code block formatting
    
    Args:
        input_file_path: Path to the source code file
        output_dir: Directory to save the converted markdown file
    
    Returns:
        bool: True if conversion was successful, False otherwise
    """
    try:
        input_path = Path(input_file_path)
        
        # Read the source code content
        with open(input_file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Determine language for syntax highlighting
        extension = input_path.suffix.lower()
        language_map = {
            '.abp': 'abap',
            '.as': 'actionscript',
            '.asm': 'assembly',
            '.bat': 'batch',
            '.c': 'c',
            '.cc': 'cpp',
            '.clj': 'clojure',
            '.coffee': 'coffeescript',
            '.cpp': 'cpp',
            '.cs': 'csharp',
            '.css': 'css',
            '.cxx': 'cpp',
            '.d': 'd',
            '.dart': 'dart',
            '.erl': 'erlang',
            '.forth': 'forth',
            '.go': 'go',
            '.groovy': 'groovy',
            '.h': 'c',
            '.hpp': 'cpp',
            '.hs': 'haskell',
            '.htm': 'html',
            '.html': 'html',
            '.hx': 'haxe',
            '.ipynb': 'json',
            '.java': 'java',
            '.js': 'javascript',
            '.jsx': 'jsx',
            '.kt': 'kotlin',
            '.kts': 'kotlin',
            '.lhs': 'haskell',
            '.lisp': 'lisp',
            '.lsl': 'lsl',
            '.lua': 'lua',
            '.m': 'matlab',
            '.mat': 'matlab',
            '.mjs': 'javascript',
            '.ml': 'ocaml',
            '.pas': 'pascal',
            '.php': 'php',
            '.pl': 'perl',
            '.pm': 'perl',
            '.pro': 'prolog',
            '.ps1': 'powershell',
            '.py': 'python',
            '.pyc': 'python',
            '.pyo': 'python',
            '.r': 'r',
            '.rb': 'ruby',
            '.rs': 'rust',
            '.scala': 'scala',
            '.scm': 'scheme',
            '.sh': 'bash',
            '.sql': 'sql',
            '.swift': 'swift',
            '.swi': 'prolog',
            '.ts': 'typescript',
            '.v': 'verilog',
            '.vbs': 'vbscript',
            '.xhtml': 'xml',
            '.xml': 'xml',
            '.xquery': 'xquery'
        }
        
        language = language_map.get(extension, '')
        
        # Create markdown content with code block
        markdown_content = f"# {input_path.name}\n\n```{language}\n{content}\n```\n"
        
        # Generate output filename
        output_filename = f"{input_path.stem}.md"
        output_file_path = Path(output_dir) / output_filename
        
        # Write to file
        with open(output_file_path, 'w', encoding='utf-8') as f:
            f.write(markdown_content)
        
        return True
        
    except Exception as e:
        print(f"✗ Error converting {input_file_path}: {e}")
        return False


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
        input_path = Path(input_file_path)
        source_code_extensions = get_source_code_extensions()
        
        # Check if this is a source code file
        if input_path.suffix.lower() in source_code_extensions:
            return convert_source_code_to_markdown(input_file_path, output_dir)
        
        # Use MarkItDown for other file types
        md = MarkItDown()
        
        # Convert the file
        result = md.convert(input_file_path)
        
        # Generate output filename
        output_filename = f"{input_path.stem}.md"
        output_file_path = Path(output_dir) / output_filename
        
        # Write to file
        with open(output_file_path, 'w', encoding='utf-8') as f:
            f.write(result.text_content)
        
        return True
        
    except Exception as e:
        print(f"✗ Error converting {input_file_path}: {e}")
        return False


def process_files_chunk(files_chunk, output_dir, supported_extensions, results_dict, thread_id):
    """
    Process a chunk of files in a separate thread
    
    Args:
        files_chunk: List of file paths to process
        output_dir: Directory to save converted markdown files
        supported_extensions: Dictionary of supported extensions
        results_dict: Dictionary to store results (thread-safe)
        thread_id: ID of the current thread
    """
    local_success = 0
    local_total = len(files_chunk)
    
    for file_path in tqdm(files_chunk, desc=f"Thread-{thread_id}", unit="file", position=thread_id):
        file_type = supported_extensions[file_path.suffix.lower()]
        
        if convert_file_to_markdown(str(file_path), output_dir):
            local_success += 1
    
    results_dict[thread_id] = (local_success, local_total)


def process_directory(directory_path: str, output_dir: str, num_threads: int = 1):
    """
    Process all supported files in a directory
    
    Args:
        directory_path: Path to the directory containing files
        output_dir: Directory to save converted markdown files
        num_threads: Number of threads to use for parallel processing
    
    Returns:
        tuple: (success_count, total_count)
    """
    supported_extensions = get_supported_extensions()
    
    directory = Path(directory_path)
    
    if not directory.exists():
        print(f"✗ Directory not found: {directory_path}")
        return 0, 0
    
    if not directory.is_dir():
        print(f"✗ Path is not a directory: {directory_path}")
        return 0, 0
    
    # Find all supported files
    supported_files = []
    for file_path in directory.rglob('*'):
        if file_path.is_file() and file_path.suffix.lower() in supported_extensions:
            supported_files.append(file_path)
    
    if not supported_files:
        print(f"ⓘ No supported files found in: {directory_path}")
        print(f"Supported extensions: {', '.join(supported_extensions.keys())}")
        return 0, 0
    
    print(f"📁 Found {len(supported_files)} supported files in: {directory_path}")
    
    # If only one thread, process sequentially
    if num_threads <= 1:
        success_count = 0
        total_count = 0
        
        for file_path in tqdm(supported_files, desc="Converting files", unit="file"):
            total_count += 1
            file_type = supported_extensions[file_path.suffix.lower()]
            
            if convert_file_to_markdown(str(file_path), output_dir):
                success_count += 1
        
        return success_count, total_count
    
    # Calculate chunk size for each thread
    total_files = len(supported_files)
    chunk_size = total_files // num_threads
    remainder = total_files % num_threads
    
    # Divide files among threads
    file_chunks = []
    start_idx = 0
    
    for i in range(num_threads):
        # Add one extra file to some threads to handle remainder
        current_chunk_size = chunk_size + (1 if i < remainder else 0)
        end_idx = start_idx + current_chunk_size
        
        if start_idx < total_files:
            file_chunks.append(supported_files[start_idx:end_idx])
            start_idx = end_idx
    
    # Process files in parallel
    threads = []
    results_dict = {}
    
    print(f"🚀 Using {len(file_chunks)} threads for parallel processing")
    
    for i, chunk in enumerate(file_chunks):
        if chunk:  # Only create thread if chunk is not empty
            thread = threading.Thread(
                target=process_files_chunk,
                args=(chunk, output_dir, supported_extensions, results_dict, i)
            )
            threads.append(thread)
            thread.start()
    
    # Wait for all threads to complete
    for thread in threads:
        thread.join()
    
    # Calculate total results
    total_success = sum(success for success, _ in results_dict.values())
    total_count = sum(count for _, count in results_dict.values())
    
    return total_success, total_count


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
        print(f"✗ File not found: {file_path}")
        return False
    
    if not file.is_file():
        print(f"✗ Path is not a file: {file_path}")
        return False
    
    if file.suffix.lower() not in supported_extensions:
        print(f"✗ Unsupported file type: {file.suffix}")
        print(f"Supported extensions: {', '.join(supported_extensions.keys())}")
        return False
    
    file_type = supported_extensions[file.suffix.lower()]
    print(f"🔄 Processing {file_type} file: {file_path}")
    
    return convert_file_to_markdown(file_path, output_dir)


def main():
    parser = argparse.ArgumentParser(
        description="Convert documents and source code (50+ programming languages) to Markdown",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Convert all files in a directory
  python convert-all.py --directorypath src/media
  
  # Convert a single file
  python convert-all.py --filepath src/media/document.pdf
  
  # Specify custom output directory
  python convert-all.py --directorypath src/media --output custom-output
  
  # Use parallel processing with 4 threads
  python convert-all.py --directorypath src/media --subthread 4
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
    
    parser.add_argument(
        '--subthread',
        type=int,
        default=1,
        help='Number of parallel threads for processing files (default: 1)'
    )
    
    args = parser.parse_args()
    
    # Create output directory
    output_dir = Path(args.output)
    output_dir.mkdir(exist_ok=True)
    print(f"📂 Output directory: {output_dir.absolute()}")
    
    # Display supported file types
    supported_extensions = get_supported_extensions()
    print(f"📄 Supported file types: {', '.join(f'{ext} ({ftype})' for ext, ftype in supported_extensions.items())}")
    print()
    
    if args.directorypath:
        # Process directory
        success_count, total_count = process_directory(args.directorypath, str(output_dir), args.subthread)
        
        print()
        print("=" * 50)
        print(f"📊 Conversion Summary:")
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
            print("✓ File conversion completed successfully!")
            sys.exit(0)
        else:
            print("✗ File conversion failed!")
            sys.exit(1)


if __name__ == "__main__":
    main()