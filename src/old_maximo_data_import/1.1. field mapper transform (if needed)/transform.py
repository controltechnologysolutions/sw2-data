#!/usr/bin/env python3
import os
import re
import json
import glob
import argparse
import sys
import traceback

def find_input_files(input_path):
    """
    If input_path ends with '_<number>.json', gather all numbered parts
    (mydata_1.json, mydata_2.json, ...).
    Otherwise, return just [input_path].
    """
    directory = os.path.dirname(input_path) or '.'
    basename = os.path.basename(input_path)
    filename_no_ext, ext = os.path.splitext(basename)

    match = re.match(r'^(.*)_(\d+)$', filename_no_ext)
    if match:
        prefix = match.group(1)
        pattern = os.path.join(directory, f"{prefix}_*{ext}")
        all_parts = glob.glob(pattern)

        def extract_part_number(filepath):
            name = os.path.splitext(os.path.basename(filepath))[0]
            return int(name.split('_')[-1])

        all_parts.sort(key=extract_part_number)
        return all_parts
    else:
        return [input_path]

def load_json_file(path):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)

def get_value_by_path(obj, path_str):
    """
    Safely traverse nested dicts with dot notation, e.g. 'location.coords.lat'.
    Return None if missing.
    """
    if not path_str:
        return None
    parts = path_str.split('.')
    current = obj
    for p in parts:
        if isinstance(current, dict) and p in current:
            current = current[p]
        else:
            return None
    return current

def transform_array(input_array, item_map, value_mapping):
    """
    Transform each element of input_array using 'item_map' (a from_to spec),
    returning a new list.
    """
    result = []
    for item in input_array:
        transformed_item = apply_mapping(item, item_map, value_mapping)
        result.append(transformed_item)
    return result

def apply_mapping(input_obj, mapping_spec, value_mapping):
    """
    Recursively apply 'mapping_spec' to 'input_obj'.
    - If mapping_spec is a string => dot path (like "priority").
    - If mapping_spec is a dict with 'arrayPath' => transform an array.
    - Else, nested object mapping.
    - value_mapping is { fieldName: { rawVal: mappedVal } }.

    Example:
      if mapping_spec == "priority":
        raw_val = get_value_by_path(input_obj, "priority")
        if "priority" in value_mapping and raw_val in that map, use the mapped value.
    """
    if isinstance(mapping_spec, str):
        raw_val = get_value_by_path(input_obj, mapping_spec)
        if mapping_spec in value_mapping:
            sub_map = value_mapping[mapping_spec]
            if raw_val in sub_map:
                raw_val = sub_map[raw_val]
        return raw_val

    if isinstance(mapping_spec, dict):
        if "arrayPath" in mapping_spec and "itemMap" in mapping_spec:
            arr_path = mapping_spec["arrayPath"]
            arr_val = get_value_by_path(input_obj, arr_path)
            if isinstance(arr_val, list):
                return transform_array(arr_val, mapping_spec["itemMap"], value_mapping)
            else:
                return []
        else:
            output_obj = {}
            for out_field, sub_spec in mapping_spec.items():
                output_obj[out_field] = apply_mapping(input_obj, sub_spec, value_mapping)
            return output_obj

    return mapping_spec

def apply_defaults_with_skip(obj, defaults):
    """
    Merge 'defaults' into 'obj', skipping creation of new nested objects or arrays if they don't exist.
    For arrays, we do NOT append. Instead, we merge the default object (the first element of defaults array)
    into every item of the existing array, if they're dicts.

    Behavior:
      - If default_val is a scalar => set it if missing in obj.
      - If default_val is a dict => only apply if obj[k] is a dict too (and already exists).
      - If default_val is a list => only apply if obj[k] is a list too (and already exists).
          * Then, for each element in obj[k], if default_val has at least one item,
            we recursively merge that item into each element.
    """
    if not isinstance(obj, dict) or not isinstance(defaults, dict):
        return

    for k, default_val in defaults.items():
        if not isinstance(default_val, (dict, list)):
            if k not in obj:
                obj[k] = default_val
        else:
            if k not in obj:
                continue

            if isinstance(default_val, dict) and isinstance(obj[k], dict):
                apply_defaults_with_skip(obj[k], default_val)

            elif isinstance(default_val, list) and isinstance(obj[k], list):
                if len(default_val) == 0:
                    continue
                default_item = default_val[0]
                if isinstance(default_item, dict):
                    for element in obj[k]:
                        if isinstance(element, dict):
                            apply_defaults_with_skip(element, default_item)
            else:
                # type mismatch => do nothing
                pass

def main():
    try:
        parser = argparse.ArgumentParser(
            description=(
                "Transform JSON with nested fields & array mappings, apply value mappings, "
                "and only set nested defaults if parent object/array already exists. "
                "For arrays, merge default values (the first item) into each element, no append."
            )
        )
        parser.add_argument('--input-json', required=True,
                            help="Path to input JSON array (may be split). E.g. 'data_1.json'.")
        parser.add_argument('--from-to-json', required=True,
                            help="Mapping spec (nested/array) to transform input -> output.")
        parser.add_argument('--mapping-json', required=False,
                            help="Value mapping: { fieldName: { rawVal: mappedVal } }.")
        parser.add_argument('--default-values-json', required=False,
                            help="Nested default values, only set if the parent object/array already exists.")
        parser.add_argument('--output-json', required=True,
                            help="Destination file for transformed JSON array.")

        args = parser.parse_args()

        from_to_map = load_json_file(args.from_to_json)

        value_mapping = {}
        if args.mapping_json:
            value_mapping = load_json_file(args.mapping_json)

        default_values = {}
        if args.default_values_json:
            default_values = load_json_file(args.default_values_json)

        input_files = find_input_files(args.input_json)
        if not input_files:
            print(f"[ERROR] No matching input files for '{args.input_json}'. Exiting.")
            sys.exit(1)

        all_input_data = []
        for fp in input_files:
            part = load_json_file(fp)
            if not isinstance(part, list):
                print(f"[ERROR] File '{fp}' is not a JSON array. Exiting.")
                sys.exit(1)
            all_input_data.extend(part)

        transformed_data = []
        for obj in all_input_data:
            out_obj = apply_mapping(obj, from_to_map, value_mapping)
            apply_defaults_with_skip(out_obj, default_values)
            transformed_data.append(out_obj)

        with open(args.output_json, 'w', encoding='utf-8') as f:
            json.dump(transformed_data, f, ensure_ascii=False, indent=2)

        print(f"Done! Wrote {len(transformed_data)} record(s) to '{args.output_json}'.")
    except Exception as e:
        print("[FATAL ERROR]", e)
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()