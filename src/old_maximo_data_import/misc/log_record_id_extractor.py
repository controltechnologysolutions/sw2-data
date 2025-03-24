import re
import sys

def extract_record_ids(file_path):
    try:
        with open(file_path, 'r') as file:
            log_data = file.read()
        
        pattern = r"Record (\d+) \(action=-mu\)"
        
        record_ids = re.findall(pattern, log_data)
        
        record_ids = [int(record_id) for record_id in record_ids]
        
        return record_ids
    except FileNotFoundError:
        print(f"File not found: {file_path}")
        return []
    except Exception as e:
        print(f"An error occurred: {e}")
        return []

def write_ids_to_file(ids, output_file):
    try:
        with open(output_file, 'w') as file:
            for record_id in ids:
                file.write(f"{record_id}\n")
        print(f"Record IDs written to {output_file}")
    except Exception as e:
        print(f"An error occurred while writing to file: {e}")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python extract_ids.py <input_file_path> <output_file_path>")
        sys.exit(1)

    input_file_path = sys.argv[1]
    output_file_path = sys.argv[2]
    
    record_ids = extract_record_ids(input_file_path)

    if record_ids:
        write_ids_to_file(record_ids, output_file_path)
    else:
        print("No Record IDs found.")