<#
.SYNOPSIS
    Automates the creation of Word documents based on Excel data for various document types.

.DESCRIPTION
    This script reads data from an Excel spreadsheet and creates Word documents based on specified templates.
    It supports multiple document types (Type1, Type2, Type3) and updates document content using placeholders
    defined in a mapping file. The script handles document creation, updating existing documents, and detailed logging.

.PARAMETER ExcelPath
    Path to the Excel file containing source data.

.PARAMETER SheetName
    Name of the Excel worksheet containing the data.

.PARAMETER DocumentType
    Type of document to create. Valid values are "Type1", "Type2", "Type3".

.PARAMETER TemplateDocPath
    Path to the Word document template to use.

.PARAMETER OutputFolder
    Folder where generated documents will be saved.

.PARAMETER TokenMapPath
    Path to the mapping file that defines Excel column to document placeholder relationships.
    Defaults to "data_column_map.txt" in the script's directory.

.EXAMPLE
    .\CreateAnalysisDocuments.ps1 -ExcelPath "absolute\path\to\data_source\application_inventory.xlsx" -SheetName "Sheet1" -DocumentType "Type1" -TemplateDocPath "absolute\path\to\doc_templates\Type_1_Analysis.docx" -OutputFolder "absolute\path\to\output"

.NOTES
    File Name    : CreateAnalysisDocuments.ps1
    Author       : Atmosera
    Requires     : PowerShell 5.1 or later
    Version      : 1.0
    Date         : September 5, 2025
#>

param (
    [Parameter(Mandatory=$true)]
    [string]$ExcelPath,
    
    [Parameter(Mandatory=$true)]
    [string]$SheetName,
    
    [Parameter(Mandatory=$true)]
    [ValidateSet("Type1", "Type2", "Type3")]
    [string]$DocumentType,
    
    [Parameter(Mandatory=$true)]
    [string]$TemplateDocPath,
    
    [Parameter(Mandatory=$true)]
    [string]$OutputFolder,
    
    [Parameter(Mandatory=$false)]
    [string]$TokenMapPath = "$PSScriptRoot\data_column_map.txt"
)

# Initialize logging
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$successLogPath = "$($PSScriptRoot)\$($DocumentType)_Creation_Success_$($timestamp).log"
$existingLogPath = "$($PSScriptRoot)\$($DocumentType)_Creation_Existing_$($timestamp).log"
$errorLogPath = "$($PSScriptRoot)\$($DocumentType)_Creation_Errors_$($timestamp).log"

# Create log files with headers
"Document Creation Log - $(Get-Date)" | Out-File -FilePath $successLogPath
"Application,Path,Properties" | Out-File -FilePath $successLogPath -Append

"Existing Documents Log - $(Get-Date)" | Out-File -FilePath $existingLogPath
"Application,Path,StatusUpdated" | Out-File -FilePath $existingLogPath -Append

"Document Creation Error Log - $(Get-Date)" | Out-File -FilePath $errorLogPath
"Application,Error" | Out-File -FilePath $errorLogPath -Append

# Function to log success
function Write-SuccessLog {
    param (
        [string]$ApplicationName,
        [string]$DocumentPath,
        [hashtable]$Properties
    )
    
    $propsString = ($Properties.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join "; "
    "$ApplicationName,$DocumentPath,$propsString" | Out-File -FilePath $successLogPath -Append
}

# Function to log existing documents
function Write-ExistingLog {
    param (
        [string]$ApplicationName,
        [string]$DocumentPath,
        [bool]$StatusUpdated
    )
    
    "$ApplicationName,$DocumentPath,$StatusUpdated" | Out-File -FilePath $existingLogPath -Append
}

# Function to log errors
function Write-ErrorLog {
    param (
        [string]$ApplicationName,
        [string]$ErrorMessage
    )
    
    "$ApplicationName,$ErrorMessage" | Out-File -FilePath $errorLogPath -Append
}

# Function to create directory if it doesn't exist
function Ensure-Directory {
    param (
        [string]$Path
    )
    
    if (-not (Test-Path -Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

# Function to get the Excel column name based on document type
function Get-StatusColumnName {
    param (
        [string]$DocType
    )
    
    switch ($DocType) {
        "Type1" { return "Type 1 Analysis Created" }
        "Type2" { return "Type 2 Analysis Created" }
        "Type3" { return "Type 3 Analysis Created" }
        default { throw "Unsupported document type: $DocType" }
    }
}

# Ensure output folder exists
Ensure-Directory -Path $OutputFolder

# Check if Excel file exists
if (-not (Test-Path -Path $ExcelPath)) {
    Write-Host "Error: Excel file not found at path: $ExcelPath" -ForegroundColor Red
    exit 1
}

# Check if template document exists
if (-not (Test-Path -Path $TemplateDocPath)) {
    Write-Host "Error: Template document not found at path: $TemplateDocPath" -ForegroundColor Red
    exit 1
}

# Check if token map exists
if (-not (Test-Path -Path $TokenMapPath)) {
    Write-Host "Error: Token map file not found at path: $TokenMapPath" -ForegroundColor Red
    exit 1
}

# Read token map file
$tokenMap = @{}
Get-Content -Path $TokenMapPath | ForEach-Object {
    if ($_ -match "^(.*?),\s*(.*?)$") {
        $columnName = $matches[1].Trim()
        $wordPropertyName = $matches[2].Trim()
        $tokenMap[$columnName] = $wordPropertyName
    }
}

# Get the status column name based on document type
$statusColumnName = Get-StatusColumnName -DocType $DocumentType

# Import Excel data and keep workbook open for updates
try {
    # Create Excel COM object
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    # Open the workbook and select the sheet
    $workbook = $excel.Workbooks.Open($ExcelPath)
    
    # Try to select the specified sheet
    try {
        $worksheet = $workbook.Worksheets.Item($SheetName)
    }
    catch {
        Write-Host "Error: Sheet '$SheetName' not found in Excel workbook. Available sheets:" -ForegroundColor Red
        foreach ($sheet in $workbook.Worksheets) {
            Write-Host "  - $($sheet.Name)" -ForegroundColor Yellow
        }
        $workbook.Close($false)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        exit 1
    }
    
    # Get used range
    $usedRange = $worksheet.UsedRange
    $rowCount = $usedRange.Rows.Count
    $columnCount = $usedRange.Columns.Count
    
    if ($rowCount -le 1) {
        Write-Host "Error: Excel sheet contains no data or only headers" -ForegroundColor Red
        $workbook.Close($false)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        exit 1
    }
    
    # Get headers from first row and find the status column index
    $headers = @()
    $statusColumnIndex = -1
    for ($col = 1; $col -le $columnCount; $col++) {
        $headerValue = $worksheet.Cells.Item(1, $col).Value2
        if ($headerValue) {
            $headers += $headerValue
            if ($headerValue.Trim() -eq $statusColumnName.Trim()) {
                $statusColumnIndex = $col
            }
        }
    }
    
    # Check if the status column exists
    if ($statusColumnIndex -eq -1) {
        Write-Host "Warning: Status column '$statusColumnName' not found in Excel sheet. Status updates will be skipped." -ForegroundColor Yellow
    }
    else {
        Write-Host "Found status column '$statusColumnName' at index $statusColumnIndex" -ForegroundColor Green
    }
    
    # Convert Excel data to objects similar to CSV import
    $excelData = @()
    $rowIndexMap = @{}  # To store row indexes for updating later
    
    for ($row = 2; $row -le $rowCount; $row++) {
        $rowData = New-Object PSObject
        $appName = $null
        
        for ($col = 1; $col -le $columnCount; $col++) {
            if ($col -le $headers.Count) {
                $cellValue = $worksheet.Cells.Item($row, $col).Value2
                $rowData | Add-Member -MemberType NoteProperty -Name $headers[$col-1] -Value $cellValue
                
                # Store the application name to use as a key
                if ($headers[$col-1] -eq "Application / System") {
                    $appName = $cellValue
                }
            }
        }
        
        # Store the row index mapped to application name
        if (-not [string]::IsNullOrEmpty($appName)) {
            $rowIndexMap[$appName] = $row
        }
        
        $excelData += $rowData
    }
    
    Write-Host "Successfully imported Excel data with $($excelData.Count) rows from sheet '$SheetName'" -ForegroundColor Green
}
catch {
    Write-Host "Error importing Excel data: $_" -ForegroundColor Red
    
    # Clean up Excel objects if they exist
    if ($worksheet) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    }
    if ($workbook) {
        try { $workbook.Close($false) } catch { }
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel) {
        try { $excel.Quit() } catch { }
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    exit 1
}

# Initialize Word COM object
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    Write-Host "Successfully initialized Word application" -ForegroundColor Green
}
catch {
    Write-Host "Error initializing Word application: $_" -ForegroundColor Red
    exit 1
}

# Process each row in Excel data
$processedCount = 0
$existingCount = 0
$errorCount = 0

foreach ($row in $excelData) {
    $applicationName = $row.'Application / System'
    
    if ([string]::IsNullOrEmpty($applicationName)) {
        Write-Host "Skipping row with empty application name" -ForegroundColor Yellow
        continue
    }
    
    $division = $row.Division
    
    # Create output subfolder based on Division
    $divisionFolder = Join-Path -Path $OutputFolder -ChildPath $division
    if ($tokenMap.ContainsKey("Division") -and $tokenMap["Division"] -eq "DOCUMENT_FOLDER") {
        Ensure-Directory -Path $divisionFolder
    }

    
    # Define output file name based on document type
    $outputFileName = switch ($DocumentType) {
        "Type1" { "$applicationName - Type 1 Analysis.docx" }
        "Type2" { "$applicationName - Type 2 Analysis.docx" }
        "Type3" { "$applicationName - Type 3 Analysis.docx" }
        default { "$applicationName - Default Analysis.docx" }
    }
    $outputFilePath = Join-Path -Path $divisionFolder -ChildPath $outputFileName

    Write-Host "Processing: $applicationName" -ForegroundColor Cyan
    
    # Check if document already exists
    if (Test-Path -Path $outputFilePath) {
        Write-Host "  Document already exists at $outputFilePath" -ForegroundColor Yellow
        
        # Check and update Excel status column if needed
        $statusUpdated = $false
        if ($statusColumnIndex -ne -1 -and $rowIndexMap.ContainsKey($applicationName)) {
            $excelRow = $rowIndexMap[$applicationName]
            $currentStatus = $worksheet.Cells.Item($excelRow, $statusColumnIndex).Value2
            
            if ($currentStatus -ne "Yes") {
                $worksheet.Cells.Item($excelRow, $statusColumnIndex).Value2 = "Yes"
                Write-Host "  Updated Excel status column for $applicationName to 'Yes'" -ForegroundColor Green
                $statusUpdated = $true
            }
            else {
                Write-Host "  Excel status column for $applicationName already set to 'Yes'" -ForegroundColor Gray
            }
        }
        
        # Log the existing document
        Write-ExistingLog -ApplicationName $applicationName -DocumentPath $outputFilePath -StatusUpdated $statusUpdated
        
        # Skip to next row
        Write-Host "  Skipping document creation" -ForegroundColor Yellow
        $existingCount++
        continue
    }
    
    try {
        # Open template document
        $doc = $word.Documents.Open($TemplateDocPath)
        
        # Prepare properties to update
        $propertiesUpdated = @{}
        
        # Update custom document properties based on token map
        foreach ($mapEntry in $tokenMap.GetEnumerator()) {
            $csvColumnName = $mapEntry.Key
            $wordPropertyName = $mapEntry.Value
            
            # Skip folder mapping entry
            if ($wordPropertyName -eq "DOCUMENT_FOLDER") {
                continue
            }
            
            $propertyValue = $row.$csvColumnName
            if (-not [string]::IsNullOrEmpty($propertyValue)) {
                # Try to get and update the custom property
                $customProperties = $doc.CustomDocumentProperties
                $binding = "System.Reflection.BindingFlags" -as [type]
                [array]$customPropData = $wordPropertyName, $false, 4, $propertyValue
                try {
                    [System.__ComObject].InvokeMember("add", $binding::InvokeMethod, $null, $customProperties, $customPropData) | out-null   
                    Write-Host "Added custom property $wordPropertyName with a value of '$propertyValue'" -ForegroundColor Green               
                }
                catch [system.exception]{
                    $customPropComObject = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $customProperties, @($wordPropertyName))
                    [System.__ComObject].InvokeMember("Delete", $binding::InvokeMethod, $null, $customPropComObject, $null)
                    [System.__ComObject].InvokeMember("add", $binding::InvokeMethod, $null, $customProperties, $customPropData) | out-null
                    Write-Host "Replace existing custom property $wordPropertyName with a value of '$propertyValue'" -ForegroundColor Green
                }
                $propertiesUpdated[$wordPropertyName] = $propertyValue
            }
        }
        
        # Update all fields in the document
        $doc.Fields.Update() | Out-Null
        
        # Save document
        $doc.SaveAs($outputFilePath)
        $doc.Close()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
        $doc = $null
        
        # Update Excel status column
        if ($statusColumnIndex -ne -1 -and $rowIndexMap.ContainsKey($applicationName)) {
            $excelRow = $rowIndexMap[$applicationName]
            $worksheet.Cells.Item($excelRow, $statusColumnIndex).Value2 = "Yes"
            Write-Host "  Updated Excel status column for $applicationName to 'Yes'" -ForegroundColor Green
        }
        
        Write-Host "  Success: Document created at $outputFilePath" -ForegroundColor Green
        Write-SuccessLog -ApplicationName $applicationName -DocumentPath $outputFilePath -Properties $propertiesUpdated
        $processedCount++
    }
    catch {
        $errorMessage = $_.Exception.Message
        Write-Host "  Error: Failed to process $applicationName - $errorMessage" -ForegroundColor Red
        Write-ErrorLog -ApplicationName $applicationName -ErrorMessage $errorMessage
        $errorCount++
        
        # Close document if open (cleanup)
        if ($doc) {
            try { $doc.Close($false) } catch { }
        }
    }
}

# Save and close Excel workbook
if ($statusColumnIndex -ne -1) {
    try {
        $workbook.Save()
        Write-Host "Excel workbook saved with status updates" -ForegroundColor Green
    } 
    catch {
        Write-Host "Warning: Failed to save Excel workbook: $_" -ForegroundColor Yellow
    }
}

# Cleanup Word and Excel applications
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null

$workbook.Close($true)  # True to save changes
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

# Report summary
Write-Host "`nDocument Creation Summary:" -ForegroundColor Cyan
Write-Host "  Successfully created: $processedCount documents" -ForegroundColor Green
Write-Host "  Existing documents found: $existingCount documents" -ForegroundColor Yellow
Write-Host "  Failed to create: $errorCount documents" -ForegroundColor $(if ($errorCount -gt 0) { "Red" } else { "Green" })
Write-Host "  Total processed: $($processedCount + $existingCount + $errorCount) documents" -ForegroundColor White
Write-Host "`nLog files:" -ForegroundColor Cyan
Write-Host "  Success log: $successLogPath" -ForegroundColor Green
Write-Host "  Existing document log: $existingLogPath" -ForegroundColor Yellow
Write-Host "  Error log: $errorLogPath" -ForegroundColor $(if ($errorCount -gt 0) { "Red" } else { "Green" })
