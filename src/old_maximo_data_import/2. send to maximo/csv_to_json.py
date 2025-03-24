import csv
import json
import threading
from queue import Queue
import argparse
import os
import re
from datetime import datetime

CHUNK_SIZE_DEFAULT = 10000
THREADS_DEFAULT = 4
MAX_FILE_SIZE = 100 * 1024 * 1024  # 100 MB

INDEXED_BRACKET_PATTERN = re.compile(r'^([^,\[\{]+)\[(\d+)\]\[([^,\]]+)\]$')
BRACKET_PATTERN = re.compile(r'^([^,\[\{]+)\[([^,\]]+)\]$')
BRACE_PATTERN   = re.compile(r'^([^,\[\{]+)\{([^,\}]+)\}$')

POSSIBLE_DATE_FORMATS = [
    '%Y-%m-%d',
    '%Y-%m-%d %H:%M:%S',
    '%m/%d/%Y',
    '%d/%m/%Y',
    '%m/%d/%Y %H:%M:%S',
    '%d/%m/%Y %H:%M:%S',
]


def parse_date_if_match(value):
    """
    Check the given `value` against known date/time formats.
    If it matches any format exactly, convert the date/time to UTC
    and return an ISO8601 string. Otherwise, return None.
    """

    trimmed = value.strip()
    if not trimmed:
        return None

    for fmt in POSSIBLE_DATE_FORMATS:
        try:
            dt = datetime.strptime(trimmed, fmt)
            return dt.isoformat()
        except ValueError:
            pass
    return None


def transform_person(value):
    """
    Transform a person name by taking the first letter of the first name
    and concatenating it with the entire last name, all in uppercase.
    For example, 'Karl Humphrey' becomes 'KHUMPHREY'.
    """
    if not value.strip():
        return value
    parts = value.strip().split()
    if len(parts) < 2:
        return value 
    return (parts[0][0] + parts[-1]).upper()


def parse_csv_chunk(chunk, headers, parse_dates=False, person_transform_columns=None):
    """
    Convert a list of CSV rows (chunk) into a list of row-objects (dicts).
      - Headers in the form field[subfield] go into row_dict["field"][0]["subfield"].
      - Headers in the form field{subfield} go into row_dict["field"]["subfield"].
      - All other headers go into row_dict[header].

    If 'parse_dates' is True, only fields matching any known date/time
    format are converted to UTC ISO8601 strings.

    If 'person_transform_columns' is provided (as a list of header names),
    then for any header in that list, the corresponding value is transformed
    using the person name rule (first letter of the first name + last name, all uppercase).
    """
    row_objects = []
    for row in chunk:
        row_dict = {}
    
        for h, val in zip(headers, row):
            if parse_dates and val.strip():
                parsed = parse_date_if_match(val)
                if parsed is not None:
                    val = parsed
            if person_transform_columns and h in person_transform_columns:
                val = transform_person(val)

            indexed_bracket_match = INDEXED_BRACKET_PATTERN.match(h)
            if indexed_bracket_match:
                field = indexed_bracket_match.group(1)
                index = int(indexed_bracket_match.group(2))
                subfield = indexed_bracket_match.group(3)
                if field not in row_dict:
                    row_dict[field] = []
                while len(row_dict[field]) <= index:
                    row_dict[field].append({})
                row_dict[field][index][subfield] = val
            else:
                bracket_match = BRACKET_PATTERN.match(h)
                brace_match   = BRACE_PATTERN.match(h)
                if bracket_match:
                    field    = bracket_match.group(1)
                    subfield = bracket_match.group(2)
                    if field not in row_dict:
                        row_dict[field] = [{}]
                    row_dict[field][0][subfield] = val
                elif brace_match:
                    field    = brace_match.group(1)
                    subfield = brace_match.group(2)
                    if field not in row_dict:
                        row_dict[field] = {}
                    row_dict[field][subfield] = val
                else:
                    row_dict[h] = val
        row_objects.append(row_dict)
    return row_objects


def worker(input_queue, output_queue, headers, parse_dates=False, person_transform_columns=None):
    """
    Worker thread function:
      - Receives chunks of rows from 'input_queue'
      - Converts them into a list of dictionaries
      - Places that list onto 'output_queue'
    """
    while True:
        chunk = input_queue.get()
        if chunk is None:
            input_queue.task_done()
            break

        row_objects = parse_csv_chunk(
            chunk,
            headers,
            parse_dates=parse_dates,
            person_transform_columns=person_transform_columns
        )
        output_queue.put(row_objects)
        input_queue.task_done()


def open_new_file(file_index, base_filename):
    """
    Helper to open a new JSON file named <base>_fileIndex.json
    and write the initial '['.
    Returns the file handle, plus a state dict for tracking.
    """
    prefix, ext = os.path.splitext(base_filename)
    if not ext:
        ext = '.json'
    new_filename = f"{prefix}_{file_index}{ext}"
    f_out = open(new_filename, 'w', encoding='utf-8')
    f_out.write("[")
    state = {
        'first_object': True,
        'current_size': 1
    }
    return f_out, state, new_filename


def writer(output_queue, base_filename):
    """
    Writer thread function:
      - Consumes lists of row-objects from 'output_queue'
      - Writes them into multiple JSON files (each a valid array)
        up to a maximum size limit (100 MB by default).
      - Each file: [object1,object2,...]
    """
    file_index = 1
    f_out, state, current_filename = open_new_file(file_index, base_filename)

    while True:
        chunk_of_rows = output_queue.get()
        if chunk_of_rows is None:
            output_queue.task_done()
            break

        for row_dict in chunk_of_rows:
            row_json = json.dumps(row_dict, ensure_ascii=False)
            size_needed = len(row_json)
            if not state['first_object']:
                size_needed += 1 

            if (state['current_size'] + size_needed + 1) > MAX_FILE_SIZE and not state['first_object']:
                f_out.write("]")
                f_out.close()

                file_index += 1
                f_out, state, current_filename = open_new_file(file_index, base_filename)

            if not state['first_object']:
                f_out.write(",")
                state['current_size'] += 1

            f_out.write(row_json)
            state['current_size'] += len(row_json)
            state['first_object'] = False

        output_queue.task_done()

    f_out.write("]")
    f_out.close()


def open_csv_with_fallback(filename, encodings=None):
    """
    Attempt to open the CSV file using a list of encodings in `encodings`.
    If `encodings` is None, use a default list of encodings.
    Returns a file handle and the encoding used if successful,
    otherwise raises an error.
    """
    if encodings is None:
        encodings = ['utf-8', 'utf-8-sig', 'utf-16', 'latin-1', 'cp1252']

    for enc in encodings:
        try:
            f_in = open(filename, 'r', encoding=enc, newline='')
            f_in.read(4096)
            f_in.seek(0)
            return f_in, enc
        except UnicodeDecodeError:
            try:
                f_in.close()
            except:
                pass

    raise UnicodeDecodeError(f"Could not decode {filename} using these encodings: {encodings}")


def csv_to_json_threads(input_file,
                        output_file,
                        num_threads=THREADS_DEFAULT,
                        chunk_size=CHUNK_SIZE_DEFAULT,
                        enc=None,
                        parse_dates=False,
                        person_transform_columns=None):
    """
    Convert a CSV to JSON array-of-objects, but split output into ~100 MB files.

    Args:
        input_file   (str): Path to the input CSV file.
        output_file  (str): Base path for the output JSON files.
        num_threads  (int): Number of worker threads for parallel chunk parsing.
        chunk_size   (int): Number of CSV rows to read per chunk.
        enc          (str): (Optional) Encoding to use. If None, try fallback encodings.
        parse_dates  (bool): If True, only parse values matching known date/time formats.
        person_transform_columns (list[str]): Optional list of CSV header names for which
            the person transformation should be applied.
    """
    input_queue = Queue(maxsize=num_threads * 2)
    output_queue = Queue(maxsize=num_threads * 2)

    if enc:
        print(f"Opening {input_file} using encoding='{enc}'")
        f_in = open(input_file, 'r', encoding=enc, newline='')
    else:
        f_in, used_enc = open_csv_with_fallback(input_file)
        print(f"Opened {input_file} successfully with encoding='{used_enc}'")

    with f_in:
        reader = csv.reader(f_in, delimiter=',')
        try:
            headers = next(reader)
            headers = [h.lstrip('\ufeff').lower() for h in headers]
        except StopIteration:
            raise ValueError("CSV file is empty or missing headers.")

        workers = []
        for _ in range(num_threads):
            t = threading.Thread(
                target=worker,
                args=(input_queue, output_queue, headers, parse_dates, person_transform_columns)
            )
            t.start()
            workers.append(t)

        writer_thread = threading.Thread(
            target=writer,
            args=(output_queue, output_file)
        )
        writer_thread.start()

        chunk_data = []
        for i, row in enumerate(reader, start=1):
            chunk_data.append(row)
            if i % chunk_size == 0:
                input_queue.put(chunk_data)
                chunk_data = []

        if chunk_data:
            input_queue.put(chunk_data)

    for _ in range(num_threads):
        input_queue.put(None)
    input_queue.join()

    output_queue.put(None)
    output_queue.join()

    writer_thread.join()


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Convert a comma-delimited CSV to JSON, interpreting headers of the form "
            "field[subfield] as arrays, and field{subfield} as objects in the output JSON. "
            "Output may be split into multiple ~100MB JSON files."
        )
    )
    parser.add_argument('input_csv', help='Path to the input CSV file')
    parser.add_argument('output_base', help='Base path/filename for the output JSON (e.g., output.json).')
    parser.add_argument('--threads', type=int, default=THREADS_DEFAULT,
                        help=f'Number of worker threads (default: {THREADS_DEFAULT})')
    parser.add_argument('--chunk-size', type=int, default=CHUNK_SIZE_DEFAULT,
                        help=f'Rows per chunk (default: {CHUNK_SIZE_DEFAULT})')
    parser.add_argument('--encoding', type=str, default=None,
                        help='Optional file encoding. If omitted, the script tries multiple encodings.')
    parser.add_argument('--parse-dates', action='store_true',
                        help=(
                            "If set, only fields matching known date/time formats "
                            "are converted to UTC ISO 8601 strings."
                        ))
    parser.add_argument('--person-transform', nargs='+', default=None,
                        help=(
                            "One or more CSV column names for which person transformation should be applied. "
                            "For each value in these columns, the first letter of the first name will be concatenated "
                            "with the entire last name, and the result will be in uppercase. "
                            "For example, 'Karl Humphrey' becomes 'KHUMPHREY'."
                        ))

    args = parser.parse_args()

    person_transform_columns = None
    if args.person_transform:
        person_transform_columns = [col.lower() for col in args.person_transform]

    csv_to_json_threads(
        input_file=args.input_csv,
        output_file=args.output_base,
        num_threads=args.threads,
        chunk_size=args.chunk_size,
        enc=args.encoding,
        parse_dates=args.parse_dates,
        person_transform_columns=person_transform_columns
    )


if __name__ == '__main__':
    main()