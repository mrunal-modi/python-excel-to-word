# Word Document Variable Update Tool

This tool automates updating multiple Word documents with values from an Excel spreadsheet. It's particularly useful for maintaining multiple documents that share similar variable placeholders but need different values.

## Prerequisites

- Python 3.x
- Required Python packages:
  ```bash
  pip install python-docx openpyxl
  ```

## File Structure

1. Excel File (`InputFields.xlsx`):
   - Each worksheet corresponds to a Word document
   - Worksheet names must match Word document names (without .docx extension)
   - Two columns in each worksheet:
     - Column 1: Variable names
     - Column 2: Values

Example Excel structure:
```
Worksheet "App-1":
Variable    | Value
app_name    | MyApp1
server_name | MyApp1-Server

Worksheet "App-2":
Variable    | Value
app_name    | MyApp2
server_name | MyApp2-Server
```

2. Word Documents:
   - Named to match worksheet names (e.g., `App-1.docx`, `App-2.docx`)
   - Use placeholders in format: `[variable_name]`
   - Example: "Application Name: [app_name] runs on [server_name]"

## Usage

1. Place your Excel file and Word documents in the same directory as the script
2. Run the script:
   ```bash
   python simple_doc_update.py
   ```

## How It Works

1. The script reads the Excel file worksheets
2. For each worksheet:
   - Finds the corresponding .docx file
   - Reads variables and values
   - Updates placeholders in the document with their values
   - Saves the updated document

## Example

Word document content before:
```
Application Name: [app_name]
Server: [server_name]
```

Word document content after:
```
Application Name: MyApp1
Server: MyApp1-Server
```

## Limitations

- Word documents must use square brackets for placeholders: `[variable_name]`
- Excel worksheet names must exactly match document names
- All files must be in the same directory as the script
- Formatting in Word documents may be affected when placeholders are replaced

## Troubleshooting

1. Document not found:
   - Ensure worksheet names match document names exactly
   - Check that .docx files are in the same directory as the script

2. Variables not updating:
   - Verify placeholder format is correct: `[variable_name]`
   - Check Excel for matching variable names
   - Make sure there are no extra spaces in variable names

3. Permission errors:
   - Ensure Word documents aren't open while running the scri