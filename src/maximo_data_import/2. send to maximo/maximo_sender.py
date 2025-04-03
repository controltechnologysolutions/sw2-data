import sys
import json
import requests
import time

from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from requests.exceptions import RequestException

timestamp = datetime.now().timestamp()
MAXAUTH_TOKEN = "<your_maximo_token>"
FAILED_LOG_FILE = f"{timestamp}_failed_requests.log"

def load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def parse_response(response):
    """Try to parse as JSON. If 'Error' is present, treat as an error."""
    try:
        data = response.json()
    except json.JSONDecodeError:
        return False, response.text

    if isinstance(data, dict) and "Error" in data:
        return True, data
    return False, data

def build_oslc_query_url(config, record):
    """
    Build the GET URL for searching an existing record.
    Replace {key} with record[key] if present in oslc.where.
    """
    base_url = config["base_url"]
    obj_structure = config["obj_structure"]
    oslc_where_template = config["oslc.where"]
    oslc_select = config["oslc.select"]

    oslc_where = oslc_where_template
    for k, v in record.items():
        placeholder = f"{{{k}}}"
        if placeholder in oslc_where:
            oslc_where = oslc_where.replace(placeholder, str(v))

    return (
        f"{base_url}/{obj_structure}"
        f"?lean=1"
        f"&oslc.where={oslc_where}"
        f"&oslc.select={oslc_select}"
    )

def fetch_object_id(session, record, config, timeout=30):
    """
    For update/merge, fetch the existing record's ID from Maximo.
    Return None if not found or if an error occurs.
    """
    url = build_oslc_query_url(config, record)
    headers = {"maxauth": MAXAUTH_TOKEN}

    try:
        resp = session.get(url, headers=headers, timeout=timeout)
    except RequestException as ex:
        print(f"  RequestException during GET: {ex}")
        return None

    is_error, parsed_resp = parse_response(resp)
    if is_error or not isinstance(parsed_resp, dict):
        return None

    members = parsed_resp.get("member", [])
    if not members:
        return None

    obj_id_attr = config["obj_id_attr_name"]
    first_member = members[0]
    maybe_id = first_member.get(obj_id_attr)
    if isinstance(maybe_id, dict):
        return maybe_id.get("content")
    return maybe_id

def process_one_record(index, record, session, config, action, create_url, timeout_seconds):
    """
    Process a single record (create/update/merge).
    Log errors if they occur. Returns True on success, False on error.
    """
    request_body_str = json.dumps(record, ensure_ascii=False)
    
    if action == "-c":
        method = "POST"
        url = create_url
        headers = {
            "maxauth": MAXAUTH_TOKEN,
            "Content-Type": "application/json"
        }
    else:
        obj_id = fetch_object_id(session, record, config, timeout=timeout_seconds)
        if not obj_id:
            msg = (
                f"No existing record found for "
                f"{config['obj_search_attr']}={record.get(config['obj_search_attr'])}."
            )
            print(f"  {msg}")
            with open(FAILED_LOG_FILE, "a", encoding="utf-8") as fail_log:
                fail_log.write(msg + "\n")
            return False

        resource_url = f"{config['base_url']}/{config['obj_structure']}/{obj_id}?lean=1"
        url = resource_url

        if action == "-u":
            method = "POST"
            headers = {
                "maxauth": MAXAUTH_TOKEN,
                "x-method-override": "PATCH",
                "Content-Type": "application/json"
            }
        elif action == "-mu":
            method = "POST"
            headers = {
                "maxauth": MAXAUTH_TOKEN,
                "x-method-override": "PATCH",
                "patchtype": "MERGE",
                "Content-Type": "application/json"
            }
        elif action == "-d":
            method = "DELETE"
            headers = {
                "maxauth": MAXAUTH_TOKEN,
                "x-method-override": "DELETE"
            }

    try:
        resp = session.request(
            method=method,
            url=url,
            headers=headers,
            data=request_body_str,
            timeout=timeout_seconds
        )
    except RequestException as ex:
        err_msg = (
            f"Record {index} (action={action}) - RequestException:\n"
            f"  {ex}\n"
            f"  Request Body: {record}"
        )
        print(err_msg)
        with open(FAILED_LOG_FILE, "a", encoding="utf-8") as fail_log:
            fail_log.write(err_msg + "\n")
        return False

    is_error, parsed_resp = parse_response(resp)
    if is_error:
        error_data = parsed_resp["Error"]
        error_msg = error_data.get("message", "")
        if "already exists" in error_msg.lower():
            return True  # Not considered as a failure

        err_msg = (
            f"Record {index} (action={action}) had error.\n"
            f"  Request Body: {record}\n"
            f"  Response: {json.dumps(parsed_resp, indent=2, ensure_ascii=False)}"
        )
        print(err_msg)
        with open(FAILED_LOG_FILE, "a", encoding="utf-8") as fail_log:
            fail_log.write(err_msg + "\n")
        return False

    print(f"  Success for record {index}. Status code: {resp.status_code}")
    return True

def process_in_bulk(records_to_process, data_array, start_index, create_url):
    session = requests.Session()
    timeout_seconds = 1800

    if records_to_process:
        selected = [(i, data_array[i]) for i in records_to_process if 0 <= i < len(data_array)]
    else:
        selected = [(i, data_array[i]) for i in range(start_index, len(data_array))]

    chunk_size = 200
    total_responses = 0

    for chunk_start in range(0, len(selected), chunk_size):
        chunk = selected[chunk_start:chunk_start + chunk_size]
        payload_list = []
        indices_chunk = []
        for orig_index, rec in chunk:
            payload_list.append({"_data": rec})
            indices_chunk.append(orig_index)
        payload_str = json.dumps(payload_list, ensure_ascii=False)
        headers = {
            "maxauth": MAXAUTH_TOKEN,
            "Content-Type": "application/json",
            "x-method-override": "BULK"
        }
        try:
            resp = session.request(
                method="POST",
                url=create_url,
                headers=headers,
                data=payload_str,
                timeout=timeout_seconds
            )
        except RequestException as ex:
            print(f"Bulk create failed: {ex}")
            sys.exit(1)

        is_error, parsed_resp = parse_response(resp)
        try:
            response_list = parsed_resp if isinstance(parsed_resp, list) else json.loads(resp.text)
        except Exception as e:
            print(f"Failed to parse bulk create response: {e}")
            sys.exit(1)

        if not isinstance(response_list, list):
            print(f"Unexpected response format: {parsed_resp}")
            sys.exit(1)

        total_responses += len(response_list)
        for pos, item in enumerate(response_list):
            orig_index = indices_chunk[pos]
            status = item.get("_responsemeta", {}).get("status")
            if status != "201":
                log_message = (
                    f"Record {orig_index} (bulk create) had error.\n"
                    f"  Response: {json.dumps(item, indent=2, ensure_ascii=False)}"
                )
                with open(FAILED_LOG_FILE, "a", encoding="utf-8") as fail_log:
                    fail_log.write(log_message + "\n")
        print(f"Processed {len(response_list)} responses in current chunk.")
    print(f"Bulk create completed with {total_responses} responses processed.")
    sys.exit(0)

def main():
    """
    Usage:
      python maximo_sender.py <-c|-u|-mu|-d|-bc> config.json data.json [start_index]

    Also supports data.json containing either:
      1) A plain JSON array, or
      2) A JSON object with "records_to_process" (list of indices) and "data" (the array).
    If "records_to_process" is provided, only those indices will be processed.
    Otherwise, we process all records (optionally starting from start_index).
    """
    if len(sys.argv) < 4:
        print("Usage: python maximo_sender.py <-c|-u|-mu|-d|-bc> config.json data.json [start_index]")
        sys.exit(1)

    action = sys.argv[1]
    config_json = sys.argv[2]
    data_json = sys.argv[3]

    start_index = 0
    if len(sys.argv) == 5:
        try:
            start_index = int(sys.argv[4])
        except ValueError:
            print("start_index must be an integer.")
            sys.exit(1)

    if action not in ("-c", "-u", "-mu", "-d", "-bc"):
        print("Invalid action. Must be one of: -c, -u, -mu, -d, -bc")
        sys.exit(1)

    config = load_json(config_json)
    raw_data = load_json(data_json)

    if isinstance(raw_data, list):
        data_array = raw_data
        records_to_process = None
    elif isinstance(raw_data, dict):
        records_to_process = raw_data.get("records_to_process")
        data_array = raw_data.get("data")
        if data_array is None:
            print("No 'data' key found in JSON. Aborting.")
            sys.exit(1)
    else:
        print(f"The file '{data_json}' must contain a JSON array or object.")
        sys.exit(1)

    if not isinstance(data_array, list):
        print("The 'data' portion of the JSON is not an array. Aborting.")
        sys.exit(1)

    print(f"Action: {action}, Config: {config_json}, Data length: {len(data_array)}")
    print(f"Starting from index {start_index}...")

    base_url = config["base_url"]
    obj_structure = config["obj_structure"]
    create_url = f"{base_url}/{obj_structure}?lean=1"

    if action == "-bc":
       process_in_bulk(records_to_process, data_array, start_index, create_url)

    all_pairs = []
    if records_to_process:
        for i in records_to_process:
            if 0 <= i < len(data_array):
                all_pairs.append((i, data_array[i]))
    else:
        for i in range(start_index, len(data_array)):
            all_pairs.append((i, data_array[i]))

    print(f"Number of records to process: {len(all_pairs)}")

    session = requests.Session()
    timeout_seconds = 30

    max_workers = 3 # Process X at a time; NOT RECOMMENDED TO CHANGE SINCE MAXIMO SEEMS TO NOT HANDLE WELL MULTIPLE DATABASE CHANGES AT THE SAME TIME

    results = []
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_index = {}
        for (idx, rec) in all_pairs:
            time.sleep(0.1)
            fut = executor.submit(
                process_one_record,
                idx,
                rec,
                session,
                config,
                action,
                create_url,
                timeout_seconds
            )
            future_to_index[fut] = idx

        for fut in as_completed(future_to_index):
            i = future_to_index[fut]
            try:
                success = fut.result()
            except Exception as e:
                print(f"Record {i} raised an unexpected exception: {e}")
                success = False
            results.append((i, success))

    success_count = sum(1 for (_, ok) in results if ok)
    failure_count = len(results) - success_count

    print(f"Done! Processed {len(results)} records => {success_count} success, {failure_count} failure.")

if __name__ == "__main__":
    main()