import re
import argparse
import os

def extract_unique_location_ids(file_path):
    """
    Extracts all unique Location IDs from the log file where the error message is BMXAA2661E.
    Processes the file line by line for efficiency with large files.
    
    Args:
        file_path (str): Path to the log file.
    
    Returns:
        set: A set of unique extracted Location IDs.
    """
    pattern = re.compile(
        r'"message"\s*:\s*"BMXAA2661E\s*-\s*Location\s*([A-Za-z0-9]+)\s*is not a valid location."'
    )
    location_ids = set()
    
    with open(file_path, 'r', encoding='utf-8') as file:
        for line in file:
            matches = pattern.findall(line)
            if matches:
                location_ids.update(matches)
    
    return location_ids

def save_to_file(location_ids, output_file):
    """
    Saves the set of unique Location IDs to the specified output file.
    Each Location ID is written on a new line, sorted alphabetically.
    
    Args:
        location_ids (set): Set of unique Location IDs to save.
        output_file (str): Path to the output file.
    """
    with open(output_file, 'w', encoding='utf-8') as f:
        for loc in sorted(location_ids):
            f.write(loc + '\n')
    print(f"\nExtracted unique Location IDs have been saved to '{output_file}'.")

def main():
    parser = argparse.ArgumentParser(
        description='Extract, deduplicate, and save unique Location IDs from log files.'
    )
    parser.add_argument(
        'input_file',
        help='Path to the log file to be processed.'
    )
    parser.add_argument(
        '-o', '--output',
        help='Path to the output file. Defaults to "extracted_unique_locations.txt".',
        default='extracted_unique_locations.txt'
    )
    args = parser.parse_args()
    
    input_file = args.input_file
    output_file = args.output
    
    if not os.path.isfile(input_file):
        print(f"Error: The file '{input_file}' does not exist.")
        return
    
    unique_locations = extract_unique_location_ids(input_file)
    
    if unique_locations:
        print("Extracted Unique Location IDs:")
        for loc in sorted(unique_locations):
            print(loc)
        
        save_to_file(unique_locations, output_file)
    else:
        print("No matching Location IDs found.")

if __name__ == "__main__":
    main()