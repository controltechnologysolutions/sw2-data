# Maximo Data Import Toolset Documentation

This guide explains how to import data into IBM Maximo using our GUI application. The process supports direct CSV or JSON imports, with an optional data transformation step if needed.

![Step by step diagram](https://i.imgur.com/OHeOEXQ.png)

## Prerequisites

- **Python 3.7 or higher:** Ensure Python is properly installed and configured on your system
- **Input Files:** Have your source CSV or JSON file ready
- **Maximo Credentials:** Valid username and password for your Maximo instance
- **Required Python packages:** Install dependencies using `pip install -r requirements.txt`

## Step 0: Object Discovery

Before importing data, it's essential to identify which IBM Maximo object structure you need to use for creating or updating records. Here's how to discover the correct object structure in Maximo:

1. **Access the Object Structures Application**  
   - In the Maximo navigation menu, click on the **Menu** icon, then go to **Integration** → **Object Structures**. [Example image](https://i.imgur.com/0CSy4Aq.png).

2. **Search for the Object You Want to Create**  
   - In the Object Structures list, use the search fields to look for keywords that match the type of record you plan to create (e.g., *location*, *asset*, *service address*, etc.). [Example image](https://i.imgur.com/PPRurmK.png).

3. **Review the Object Structure Details**  
   - Click on a matching object structure to open its details. Here, you'll see the **Object Structure** name (e.g., **MXL_LOCATION**) and a description of what it's used for. [Example image](https://i.imgur.com/TCFOL8m.png).
   - Scroll down to view the **Source Objects** and **Child Objects** included in this structure. Confirm it contains the data fields you need to create or update.

4. **Confirm Compatibility**  
   - Make sure the object structure you choose supports the operations (create, update, delete, etc.) you need. Some object structures may be read-only or intended for other specific processes.

## Step 1: Using the GUI Application

1. **Launch the Application:**
   ```bash
   python maximo_sender_ui.py
   ```

2. **Configure the Import:**
   - **Select Data File:** Click "Browse" to choose your CSV or JSON file
   - **Choose Request Type:**
     - Bulk Create: For creating multiple records efficiently
     - Create: For creating individual records
     - Update: For updating existing records
     - Merge Update: For partial updates to existing records
     - Delete: For removing records
   - **Enter Maximo Instance:** Your Maximo instance name
   - **Enter Object Structure:** The object structure identified in Step 0
   - **Provide Credentials:** Enter your Maximo username and password

3. **Additional Configuration (for Update/Delete Operations):**
   - Search Attribute: Field used to find existing records
   - ID Attribute: Unique identifier field
   - OSLC Where: Query condition for finding records
   - OSLC Select: Fields to retrieve during search

4. **Start Processing:**
   - Click "Start Processing" to begin the import
   - Monitor progress in real-time
   - View summary when complete

## Advanced: CSV File Formatting

If preparing CSV files, you can use special header patterns for proper JSON structure:

1. **Braces `{ }` for Nested Objects**
   ```csv
   asset{id}, asset{name}
   ```
   Generates:
   ```json
   {
     "asset": {
       "id": "value",
       "name": "value"
     }
   }
   ```

2. **Brackets `[ ]` for Arrays of Objects**
   ```csv
   assetspec[0][assetattrid], assetspec[0][alnvalue]
   ```
   Generates:
   ```json
   {
     "assetspec": [
       {
         "assetattrid": "value",
         "alnvalue": "value"
       }
     ]
   }
   ```

## Optional: Transform Data Using Field Mapper

If your data needs transformation before import:

1. **Navigate to the Transform Directory:**
   - Go to the `1.1. field mapper transform (if needed)` directory

2. **Prepare Mapping Files:**
   - `default_values.json` – Set default values
   - `from_to.json` – Map source to target fields
   - `mapping.json` – Define value transformations

3. **Run the Transform Script:**
   ```bash
   python transform.py \
       --input-json input.json \
       --from-to-json from_to.json \
       --default-values-json default_values.json \
       --mapping-json mapping.json \
       --output-json output.json
   ```

4. **Use Transformed File:**
   - Select the transformed JSON file in the GUI application

## Troubleshooting

- Check the terminal for detailed error messages
- Failed operations are logged in `*_failed_requests.log`
- Verify your Maximo credentials and instance name
- Ensure your CSV/JSON data matches the expected format
- Confirm the object structure permissions in Maximo

## Final Notes

- Always test imports with a small dataset first
- Back up your data before performing updates or deletes
- Monitor the progress and summary sections for operation status
- Use the "Clear" button to reset the form for new imports

The GUI application simplifies the import process by handling CSV to JSON conversion automatically and providing real-time feedback on the import progress.