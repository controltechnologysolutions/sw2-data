# Maximo Data Import Toolset Documentation

This guide explains how to convert CSV data to JSON, (optionally) transform the JSON data to match your target database schema, and finally send the processed data to your IBM Maximo instance. Additional utilities are available for log handling and data troubleshooting.

![Step by step diagram](https://i.imgur.com/OHeOEXQ.png)

## Prerequisites

- **Python 3 Installed:** Ensure Python 3 is properly installed and configured on your system.
- **Input Files:** Have your source CSV file ready.
- **Mapping Files:** Prepare your JSON mapping files (if needed) using the sample files as a reference:
  - `default_values.sample.json`
  - `from_to.sample.json`
  - `mapping.sample.json`
- **Maximo Configuration:** Review and update the `config.sample.json` with your Maximo instance details (e.g., base URL, object structure, and field attributes).

## Step 0: Object Discovery

Before importing data, it’s essential to identify which IBM Maximo object structure you need to use for creating or updating records. Here’s how to discover the correct object structure in Maximo:

1. **Access the Object Structures Application**  
   - In the Maximo navigation menu, click on the **Menu** icon, then go to **Integration** → **Object Structures**. [Example image](https://i.imgur.com/0CSy4Aq.png).

2. **Search for the Object You Want to Create**  
   - In the Object Structures list, use the search fields to look for keywords that match the type of record you plan to create (e.g., *location*, *asset*, *service address*, etc.). [Example image](https://i.imgur.com/PPRurmK.png).

3. **Review the Object Structure Details**  
   - Click on a matching object structure to open its details. Here, you’ll see the **Object Structure** name (e.g., **MXL_LOCATION**) and a description of what it’s used for. [Example image](https://i.imgur.com/TCFOL8m.png).
   - Scroll down to view the **Source Objects** and **Child Objects** included in this structure. Confirm it contains the data fields you need to create or update.

4. **Confirm Compatibility**  
   - Make sure the object structure you choose supports the operations (create, update, delete, etc.) you need. Some object structures may be read-only or intended for other specific processes.

Once you have identified the appropriate object structure (e.g., **MXL_LOCATION** for creating or updating location records), note its name and reference it later in your `config.json` file. This ensures your data will be correctly interpreted and processed by Maximo.

## Step 1: Convert CSV to JSON

1. **Navigate to the Folder:**
   - Go to the `1. convert csv to json` directory.

2. **Run the Conversion Script:**
   - Execute the script with the required parameters. For example:
     ```bash
     python3 csv_to_json.py /path/to/csv.csv /path/to/output.json --threads 4 --chunk-size 50000 --encoding utf-8
     ```
   - **Parameters Explained:**
      - **Input CSV File:** `/path/to/csv.csv`  
        - **Expected:** A string representing the file path to the input CSV file containing your raw data.
      - **Output Base:** `/path/to/output.json`  
        - **Expected:** A string representing the base path/filename for the output JSON file. The script may split the output into multiple files (e.g., output_1.json, output_2.json) if the file size exceeds a limit.
      - **Threads (--threads):** Number of worker threads for processing (default: 4).  
        - **Expected:** An integer specifying how many worker threads to use.
      - **Chunk Size (--chunk-size):** Number of CSV rows to process per chunk (default: 10000).  
        - **Expected:** An integer defining the number of rows to process in each chunk.
      - **Encoding (--encoding):** Optional file encoding. If omitted, the script tries multiple encodings (e.g., utf-8, utf-16, latin-1, etc.).  
        - **Expected:** A string indicating the file encoding (optional).
      - **Parse Dates (--parse-dates):** If set, only fields matching known date/time formats are converted to UTC ISO8601 strings.  
        - **Expected:** A boolean flag (no value needed) that, when present, enables date parsing.
      - **Person Transform (--person-transform):** One or more CSV column names for which person transformation should be applied. For each value in these columns, the first letter of the first name will be concatenated with the entire last name, and the result   will  be in uppercase (e.g., 'Karl Humphrey' becomes 'KHUMPHREY').  
        - **Expected:** A list of one or more strings representing the CSV column names.
   
3. **Verify the Output:**
   - Confirm that the JSON file has been generated successfully.

## Advanced: Adapting CSV Files for Direct Maximo Import

If your CSV files are already formatted to match the Maximo standard attribute names, you can bypass the optional transformation step (Step 2) and import the data directly into Maximo. By adapting your CSV header row to use specific patterns, the `csv_to_json.py` script will automatically generate JSON in the structure required by Maximo.

## CSV Header Patterns

1. **Braces `{ }` for Creating Nested Objects**

   - **Usage:** Use curly braces to specify that part of the CSV header should be grouped into a nested JSON object. Fields sharing the same prefix and differing only by the key inside the braces will be merged into a single object.
   - **Example:** If your CSV headers are:
     ```csv
     asset{id}, asset{name}
     ```
     The script will output a JSON object with a key `asset` containing an object with keys `id` and `name`, mapped to the corresponding values from the CSV row:
     ```json
     {
       "asset": {
         "id": "sample_id",
         "name": "sample_name"
       }
     }
     ```
   - This pattern is useful for creating structured nested objects directly in the CSV, reducing the need for a separate transformation step.

2. **Brackets `[ ]` for Creating Arrays of Objects**

   - **Usage:** Use square brackets to indicate that the corresponding fields in the CSV should be treated as an array of objects in the resulting JSON.
   - **Example:** If your CSV headers are:
     ```csv
     user[0][name], user[0][age], user[1][name], user[1][age]
     ```
     This will produce a JSON structure like:
     ```json
     {
       "user": [
         {
           "name": "John",
           "age": 30
         },
         {
           "name": "Jane",
           "age": 25
         }
       ]
     }
     ```
   - **Advanced Example – Multiple Object Arrays:**

     If your CSV contains multiple instances of a related object, you can index them using numeric indices inside the square brackets. For example, consider the following CSV headers:
     ```csv
     assetspec[0][assetattrid], assetspec[0][alnvalue], assetspec[1][assetattrid], assetspec[1][alnvalue]
     ```
     This will produce a JSON structure where `assetspec` is an array containing two objects, like:
     ```json
     {
       "assetspec": [
         {
           "assetattrid": "value from CSV",
           "alnvalue": "value from CSV"
         },
         {
           "assetattrid": "value from CSV",
           "alnvalue": "value from CSV"
         }
       ]
     }
     ```
     This method allows you to define arrays of objects directly from the CSV file by indexing each object with a numeric key.

### Benefits of Direct CSV Formatting

- **Simplified Workflow:** By adapting your CSV file headers to the Maximo standard, you eliminate the need for the field mapper transform step.
- **Direct Mapping:** Ensure that the attribute names in your CSV match the Maximo database fields exactly, resulting in a JSON output that is ready for import.
- **Flexibility:** Easily update your CSV templates (e.g., `assets_template.csv`, `locations_template.csv`, `service_addresses_template.csv`) to include these patterns, streamlining the data import process.

### Example CSV Headers for Maximo Import

- **Assets Template:**
  ```csv
  assetnum, assetspec[0][assetattrid], assetspec[0][alnvalue], assetspec[1][assetattrid], assetspec[1][alnvalue]
  ```

Adjust your CSV files accordingly before running the `csv_to_json.py` script, so that the generated JSON is in line with Maximo's expected format.

*Note:* If your CSV files are not formatted using these patterns, you can still use the transform script (Step 2) to map fields and adjust the data as needed.

## Step 2: Transform Data Using Field Mapper (Optional)

Use this step if your JSON needs to be adjusted (e.g., mapping CSV fields to database fields, applying default values, or transforming values).

1. **Navigate to the Folder:**
   - Go to the `1.1. field mapper transform (if needed)` directory.

2. **Prepare Your Mapping Files:**
   - Modify the sample files as needed:
     - `default_values.sample.json` – Set any default values for all entries.
     - `from_to.sample.json` – Define how fields from the original JSON map to the target fields.
     - `mapping.sample.json` – (Optional) Provide any value transformations.

3. **Run the Transform Script:**
   - Execute the command:
     ```bash
     python3 transform.py \
         --input-json /path/to/input.json \
         --from-to-json /path/to/from_to.json \
         --default-values-json /path/to/default_values.json \
         --mapping-json /path/to/mapping.json \
         --output-json /path/to/output.json
     ```
   - **Parameters Explained:**
     - **--input-json:** The raw JSON file (often the output from Step 1).
     - **--from-to-json:** File that maps original JSON fields to target field names.
     - **--default-values-json:** File containing default values for each record.
     - **--mapping-json:** File for value-specific mappings.
     - **--output-json:** The file path for the transformed JSON.

4. **Verify the Transformation:**
   - Check the generated output JSON file to ensure the fields and values are correctly mapped.

## Step 3: Send Data to Maximo

This step uses the transformed JSON data to interact with your Maximo system. Different commands are provided depending on whether you are creating, updating, or deleting records.

1. **Navigate to the Folder:**
   - Go to the `2. send to maximo` directory.

2. **Update Your Configuration File:**
   - Edit `config.sample.json` (rename it if needed to `config.json`) with your Maximo instance details:
     - **base_url:** Your Maximo server URL.
     - **obj_structure:** The object structure (e.g., `mxapiwodetail`).
     - **obj_search_attr:** Attribute used for searching (e.g., `wonum`).
     - Other fields such as `oslc.where` and `oslc.select` should match your query needs.

3. **Choose the Desired Operation:**
   - **Bulk Upload:**
     ```bash
     python3 maximo_sender.py -bc /path/to/config.json /path/to/data_to_send.json
     ```
   - **Create Records:**
     ```bash
     python3 maximo_sender.py -c /path/to/config.json /path/to/data_to_send.json
     ```
   - **Update Records:**
     ```bash
     python3 maximo_sender.py -u /path/to/config.json /path/to/data_to_send.json
     ```
   - **Merge Update Records:**
     ```bash
     python3 maximo_sender.py -mu /path/to/config.json /path/to/data_to_send.json
     ```
   - **Delete Records:**
     ```bash
     python3 maximo_sender.py -d /path/to/config.json /path/to/data_to_send.json
     ```

4. **(Optional) Handling Specific Records:**
   - If you need to send only a subset of records, structure your JSON as follows:
     ```json
     {
         "records_to_process": [1, 18, 39, 59, ...],
         "data": [
             {
                 "your_object_data": "here"
             },
             {
                 "another_record": "here"
             }
         ]
     }
     ```
   - This allows the script to process only the specified records.

5. **Verify the Process:**
   - Monitor the output and logs to ensure that the data has been sent successfully to Maximo.

## Additional Utilities (Optional)

The `misc` folder contains supplementary scripts that can help manage logs and extract information:

- **combine_logs.py:** Combine multiple log files for easier analysis.
- **location_extractor.py:** Extract location information from logs.
- **log_record_id_extractor.py:** Extract record IDs from log files.

These scripts can be useful for troubleshooting and tracking your data import process.

## Process Flow Overview

For a visual representation, refer to the attached diagram (`steps.svg`). It outlines the flow from CSV conversion to JSON, optional transformation, and finally the data transmission to Maximo.

## Final Notes

- **Validation:** Always validate your JSON files after each step to ensure data integrity.
- **Testing:** Run the process in a test environment before executing in production.
- **Customization:** Adjust script parameters and mapping files according to your data and Maximo configuration.

By following these steps, you can efficiently import and integrate your data into the Maximo system using the provided toolset.  