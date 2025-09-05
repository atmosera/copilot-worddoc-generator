# Word Document Generator

A PowerShell automation solution for generating Word documents from Excel data. This tool was created with GitHub Copilot a## Notes

- **Path Requirements**: Always use absolute paths for Excel files, Word templates, and output folders. COM objects used for Excel and Word automation require absolute paths to function reliably.
- The script uses COM objects for Excel and Word, so both applications must be installed
- The script will update the Excel file with document creation status if a matching status column is found
- Document generation is skipped for rows where the document already exists
- When files already exist, the script can still update the Excel status column
- The script includes proper cleanup of COM objects to prevent memory leaksce to demonstrate automation of document generation based on templates and structured data.

## Overview

This project provides a PowerShell script that:
- Reads data from Excel spreadsheets
- Uses Word document templates with custom properties
- Automatically generates individual documents based on the data
- Updates the Excel file with document creation status
- Creates detailed logs of document processing

## Project Structure

```
copilot-worddoc-generator/
│
├── code/
│   ├── CreateAnalysisDocuments.ps1  # Main PowerShell script
│   ├── data_column_map.txt          # Maps Excel columns to Word document properties
│   └── [log files]                  # Generated during script execution
│
├── data_source/
│   └── application_inventory.xlsx   # Example Excel data source
│
└── doc_templates/
    └── Type_1_Analysis.docx         # Word document template with custom properties
```

## Script Documentation

### CreateAnalysisDocuments.ps1

The main PowerShell script that processes Excel data and generates Word documents.

#### Parameters

| Parameter | Description | Required |
|-----------|-------------|----------|
| ExcelPath | Path to the Excel file containing source data | Yes |
| SheetName | Name of the Excel worksheet containing the data | Yes |
| DocumentType | Type of document to create ("Type1", "Type2", or "Type3") | Yes |
| TemplateDocPath | Path to the Word document template | Yes |
| OutputFolder | Folder where generated documents will be saved | Yes |
| TokenMapPath | Path to mapping file for Excel columns to Word properties | No (defaults to data_column_map.txt in script directory) |

#### Usage Examples

> **IMPORTANT**: Always use absolute paths for `ExcelPath`, `TemplateDocPath`, and `OutputFolder` parameters to ensure the script can locate files correctly regardless of the current working directory. Relative paths may cause errors when working with COM objects.

Basic usage with absolute paths (recommended):
```powershell
.\code\CreateAnalysisDocuments.ps1 `
    -ExcelPath "C:\Path\To\data_source\application_inventory.xlsx" `
    -SheetName "Sheet1" `
    -DocumentType "Type1" `
    -TemplateDocPath "C:\Path\To\doc_templates\Type_1_Analysis.docx" `
    -OutputFolder "C:\Path\To\output"
```

Using PowerShell environment variables for paths:
```powershell
$ProjectRoot = "C:\Path\To\copilot-worddoc-generator"
.\code\CreateAnalysisDocuments.ps1 `
    -ExcelPath "$ProjectRoot\data_source\application_inventory.xlsx" `
    -SheetName "Sheet1" `
    -DocumentType "Type1" `
    -TemplateDocPath "$ProjectRoot\doc_templates\Type_1_Analysis.docx" `
    -OutputFolder "$ProjectRoot\output" `
    -TokenMapPath "$ProjectRoot\code\data_column_map.txt"
```

#### Features

- **Document Creation**: Creates Word documents for each row in the Excel file
- **Document Organization**: Organizes generated documents into folders based on Excel data (e.g., by Division)
- **Status Tracking**: Updates the Excel file with document creation status
- **Skip Existing**: Skips document creation for documents that already exist
- **Custom Properties**: Updates Word document custom properties based on Excel data
- **Detailed Logging**: Creates separate logs for success, existing documents, and errors

## Supporting Files

### data_column_map.txt

This text file maps Excel column names to Word document custom properties. Each line contains a mapping in the format:

```
Excel Column Name, Word Property Name
```

Special mappings:
- `DOCUMENT_FOLDER`: Used to specify which Excel column value should be used as a subfolder name

Example:
```
Division, DOCUMENT_FOLDER
Application / System, System Name
Department, Department
Function, Business Function
```

### Word Document Templates

Templates must be standard Word (.docx) documents with custom document properties defined. The script will:

1. Copy the template for each row in the Excel file
2. Update the custom properties with values from Excel
3. Update any fields in the document (if they reference the custom properties)

#### Creating Templates

1. Create a Word document
2. Add custom document properties (File → Info → Properties → Advanced Properties → Custom)
3. Add fields in the document that reference these properties
4. Save as .docx format

### Excel Data Source

The Excel file should:
- Have a header row with column names
- Include a column named according to the document type (e.g., "Type 1 Analysis Created")
- Have consistent data format

The script looks for a column that matches:
- "Type 1 Analysis Created" for DocumentType "Type1"
- "Type 2 Analysis Created" for DocumentType "Type2"
- "Type 3 Analysis Created" for DocumentType "Type3"

## Log Files

The script generates three types of log files in the script directory:

1. **Success Log**: Records successfully created documents
   - Format: `[DocumentType]_Creation_Success_[timestamp].log`

2. **Existing Document Log**: Records documents that already existed
   - Format: `[DocumentType]_Creation_Existing_[timestamp].log`

3. **Error Log**: Records errors encountered during document creation
   - Format: `[DocumentType]_Creation_Errors_[timestamp].log`

## Requirements

- Windows OS
- PowerShell 5.1 or later
- Microsoft Word installed
- Microsoft Excel installed

## Notes

- The script uses COM objects for Excel and Word, so both applications must be installed
- The script will update the Excel file with document creation status if a matching status column is found
- Document generation is skipped for rows where the document already exists
- When files already exist, the script can still update the Excel status column
- The script includes proper cleanup of COM objects to prevent memory leaks

## License

See the LICENSE file for details.
