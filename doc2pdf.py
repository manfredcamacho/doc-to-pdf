#!/usr/bin/env python3

"""
A Python script to batch convert .doc and .docx files to PDF
using the LibreOffice command-line interface.

Repository: https://github.com/manfredcamacho/doc-to-pdf
"""

import sys
import os
import subprocess
import argparse
import platform
import shutil

def find_lo_command():
    """Tries to find the LibreOffice executable."""
    system = platform.system()
    
    if system == "Windows":
        candidates = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            "soffice.exe",
            "soffice"
        ]
    elif system == "Darwin":  # macOS
        candidates = [
            r"/Applications/LibreOffice.app/Contents/MacOS/soffice",
            "soffice"
        ]
    else:  # Linux and other Unix-like
        candidates = ["soffice", "libreoffice"]

    for cmd in candidates:
        # Use shutil.which to check if cmd is in PATH and executable
        if shutil.which(cmd):
            return cmd  # Return the first one that works
            
    return None  # Return None if no command was found

def convert_file(file_path, output_dir, lo_command):
    """Converts a single file to PDF."""
    if not file_path.endswith(('.doc', '.docx')):
        print(f"Skipping non-document file: {os.path.basename(file_path)}", file=sys.stderr)
        return
        
    # If no specific output_dir is given, save in the same dir as the source
    final_out_dir = output_dir if output_dir else os.path.dirname(file_path)
    
    print(f"Converting: {os.path.basename(file_path)}...")
    
    try:
        subprocess.run(
            [lo_command, '--headless', '--convert-to', 'pdf', '--outdir', final_out_dir, file_path],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.PIPE
        )
    except subprocess.CalledProcessError as e:
        print(f"Error converting {file_path}: {e.stderr.decode()}", file=sys.stderr)

def main():
    # 1. Find LibreOffice first
    lo_command = find_lo_command()
    if not lo_command:
        print("--- ERROR: Required Program Not Found ---", file=sys.stderr)
        print("This script requires LibreOffice to function.", file=sys.stderr)
        print("Please install LibreOffice from https://www.libreoffice.org/", file=sys.stderr)
        print("If it is already installed, make sure 'soffice' or 'libreoffice'", file=sys.stderr)
        print("is available in your system's PATH.", file=sys.stderr)
        sys.exit(1)
        
    print(f"Using LibreOffice command: {lo_command}")

    # 2. Set up argument parsing
    parser = argparse.ArgumentParser(
        description="Convert .doc/.docx files to PDF using LibreOffice.",
        epilog="Example: python convert_pdf.py -o ./converted_pdfs ~/Documents/report.docx"
    )
    parser.add_argument(
        "inputs", 
        nargs="+", 
        help="One or more source files or directories."
    )
    parser.add_argument(
        "-o", "--output-dir", 
        help="Optional. Directory to save all converted PDFs. (Default: same as source file)"
    )
    
    args = parser.parse_args()

    # 3. Create output directory if it was specified
    if args.output_dir:
        try:
            os.makedirs(args.output_dir, exist_ok=True)
            print(f"Output directory set to: {os.path.abspath(args.output_dir)}")
        except OSError as e:
            print(f"Error creating output directory '{args.output_dir}': {e}", file=sys.stderr)
            sys.exit(1)

    # 4. Process all provided inputs
    for path_arg in args.inputs:
        abs_path = os.path.abspath(path_arg)

        if not os.path.exists(abs_path):
            print(f"Error: Path not found '{abs_path}'", file=sys.stderr)
            continue

        if os.path.isdir(abs_path):
            print(f"\nProcessing directory: {abs_path}")
            for filename in os.listdir(abs_path):
                file_to_convert = os.path.join(abs_path, filename)
                if os.path.isfile(file_to_convert):
                    convert_file(file_to_convert, args.output_dir, lo_command)
        
        elif os.path.isfile(abs_path):
            print(f"\nProcessing file: {abs_path}")
            convert_file(abs_path, args.output_dir, lo_command)

    print("\nConversion process finished.")

if __name__ == "__main__":
    main()