import argparse
import glob
import os
import sys

def validate_files(file_paths):
    """
    Validates that each file in file_paths exists and is readable.
    
    Args:
        file_paths (list): List of file paths to validate.
    
    Raises:
        FileNotFoundError: If any file does not exist.
        PermissionError: If any file is not readable.
    """
    for file in file_paths:
        if not os.path.isfile(file):
            raise FileNotFoundError(f"File not found: {file}")
        if not os.access(file, os.R_OK):
            raise PermissionError(f"File not readable: {file}")

def combine_logs(input_files, output_file, add_header=False):
    """
    Combines multiple log files into a single output file.
    
    Args:
        input_files (list): List of input log file paths.
        output_file (str): Path to the combined output file.
        add_header (bool): Whether to add headers for each log file.
    """
    try:
        with open(output_file, 'w', encoding='utf-8') as outfile:
            for idx, file in enumerate(input_files, 1):
                if add_header:
                    header = f"\n--- Start of {os.path.basename(file)} ---\n"
                    outfile.write(header)
                
                with open(file, 'r', encoding='utf-8') as infile:
                    for line in infile:
                        outfile.write(line)
                
                if add_header:
                    footer = f"\n--- End of {os.path.basename(file)} ---\n"
                    outfile.write(footer)
                
                print(f"Successfully added: {file}")
        
        print(f"\nAll log files have been combined into '{output_file}'.")
    except Exception as e:
        print(f"An error occurred while combining log files: {e}")
        sys.exit(1)

def parse_arguments():
    """
    Parses command-line arguments.
    
    Returns:
        argparse.Namespace: Parsed command-line arguments.
    """
    parser = argparse.ArgumentParser(
        description='Combine multiple log files into a single file.'
    )
    parser.add_argument(
        'input_files',
        nargs='+',
        help='Paths to input log files. Supports wildcards, e.g., *.log'
    )
    parser.add_argument(
        '-o', '--output',
        required=True,
        help='Path to the output combined log file.'
    )
    parser.add_argument(
        '-H', '--header',
        action='store_true',
        help='Add headers and footers for each log file in the combined output.'
    )
    return parser.parse_args()

def expand_file_patterns(file_patterns):
    """
    Expands file patterns (like wildcards) into a list of file paths.
    
    Args:
        file_patterns (list): List of file patterns to expand.
    
    Returns:
        list: List of expanded file paths.
    """
    expanded_files = []
    for pattern in file_patterns:
        matched = glob.glob(pattern)
        if not matched:
            print(f"Warning: No files matched the pattern '{pattern}'.")
        expanded_files.extend(matched)
    return expanded_files

def main():
    args = parse_arguments()
    
    input_files = expand_file_patterns(args.input_files)
    
    if not input_files:
        print("Error: No input files to process.")
        sys.exit(1)
    
    input_files = list(dict.fromkeys(input_files))
    
    print(f"Total files to combine: {len(input_files)}")

    try:
        validate_files(input_files)
    except (FileNotFoundError, PermissionError) as e:
        print(f"Error: {e}")
        sys.exit(1)
    
    output_dir = os.path.dirname(os.path.abspath(args.output))
    if output_dir and not os.path.exists(output_dir):
        try:
            os.makedirs(output_dir)
            print(f"Created output directory: {output_dir}")
        except Exception as e:
            print(f"Error creating output directory '{output_dir}': {e}")
            sys.exit(1)
    
    combine_logs(input_files, args.output, add_header=args.header)

if __name__ == "__main__":
    main()