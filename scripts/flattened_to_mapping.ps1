# Stop the script immediately if PowerShell hits an error.
# Without this, some errors may be ignored and the script may keep running.
$ErrorActionPreference = "Stop"

# ============================================================
# INPUT PATHS
# ============================================================

# This is your raw Assets schema JSON export.
# This script does not need to read it yet, because the flattened Excel file already has the useful schema data.
# We still keep the path here so the script can check that the file exists.
$jsonPath = "%json_path%"

# This is the flattened Excel file you already created.
# It should contain two sheets:
#   1. ObjectTypes
#   2. Attributes
$flattenedPath = "%flattened_path%"

# Get the folder where the flattened Excel file lives.
# The new mapping workbook will be created in this same folder.
$outputFolder = Split-Path -Parent $flattenedPath

# This is the new mapping workbook this script will create.
# It will contain two sheets:
#   1. ObjectTypeMapping
#   2. AttributeMapping
$mappingWorkbookPath = Join-Path $outputFolder "assets_schema_5_mapping.xlsx"

# Create a debug log file in the Windows temp folder.
# This helps you troubleshoot if the script fails.
$debugLog = Join-Path $outputFolder ("build_mapping_workbook_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

# ============================================================
# DEBUG LOG FUNCTION
# ============================================================

function Write-DebugLog {
    param(
        [string]$Message
    )

    # Create a timestamp for each log message.
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    # Add the timestamped message to the debug log file.
    Add-Content -Path $debugLog -Value "[$timestamp] $Message"
}

# ============================================================
# CONFIGURATION SECTION
# ============================================================
# This section controls what mappings you want the script to try to build.
# You can edit this section without touching the rest of the script.

# These are the source Category values you want to support first.
# The script will try to find matching ObjectTypeName values in the ObjectTypes sheet.
$targetObjectTypeNames = @(
    "Product",
    "Perishable Product",
    "Equipment",
    "Vehicle"
)

# This maps your source CSV columns to possible Assets attribute names.
#
# Example:
#   Source column: SerialNumber
#   Possible Assets names: Serial Number, SerialNumber
#
# The script checks each possible name until it finds a match.
$sourceColumnToAttributeCandidates = [ordered]@{
    "ItemName"       = @("Name", "ItemName", "Item Name")
    "SourceRecordID" = @("SourceRecordID", "Source Record ID")
    "Model"          = @("Model")
    "SerialNumber"   = @("Serial Number", "SerialNumber")
    "Location"       = @("Location")
    "Status"         = @("Status")
    "Owner"          = @("Owner")
    "Notes"          = @("Notes")
    "JiraIssueKey"   = @("JiraIssueKey", "Jira Issue Key")
}

# ============================================================
# HELPER FUNCTION: NORMALIZE NAMES
# ============================================================

function Normalize-Name {
    param(
        [string]$Value
    )

    # If the value is empty, return an empty string.
    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ""
    }

    # Normalize names so small differences do not matter.
    #
    # Example:
    #   "Serial Number"  becomes "serialnumber"
    #   "SerialNumber"   becomes "serialnumber"
    #   "serial_number"  becomes "serialnumber"
    #
    # This makes matching less fragile.
    return ($Value.ToLowerInvariant() -replace "[^a-z0-9]", "")
}

# ============================================================
# HELPER FUNCTION: CONVERT REQUIRED VALUES TO YES/NO
# ============================================================

function Convert-ToYesNo {
    param(
        $Value
    )

    # Convert the incoming value to text.
    $text = [string]$Value

    # If the flattened file says True, Yes, or 1, treat it as required.
    if ($text -match "^(true|yes|1)$") {
        return "Yes"
    }

    # Otherwise, treat it as not required.
    return "No"
}

# ============================================================
# HELPER FUNCTION: READ AN EXCEL SHEET INTO POWERSHELL OBJECTS
# ============================================================

function Read-ExcelWorksheet {
    param(
        # The open Excel workbook object.
        [object]$Workbook,

        # The name of the sheet we want to read.
        [string]$SheetName
    )

    Write-DebugLog "Reading worksheet: $SheetName"

    # Get the worksheet by name.
    $sheet = $Workbook.Worksheets.Item($SheetName)

    # Get the used area of the worksheet.
    # This includes all rows/columns that contain data.
    $range = $sheet.UsedRange

    # Count the number of rows in the used range.
    $rowCount = $range.Rows.Count

    # Count the number of columns in the used range.
    $colCount = $range.Columns.Count

    # If the sheet has fewer than 2 rows, it means there are headers but no data.
    if ($rowCount -lt 2) {
        throw "Worksheet '$SheetName' does not contain data rows."
    }

    # This array will hold the column headers from row 1.
    $headers = @()

    # Loop through each column in row 1.
    for ($col = 1; $col -le $colCount; $col++) {

        # Read the header value.
        $header = [string]$range.Cells.Item(1, $col).Value2

        # If the header is blank, give it a generic name.
        if ([string]::IsNullOrWhiteSpace($header)) {
            $header = "Column$col"
        }

        # Add the cleaned header to the header list.
        $headers += $header.Trim()
    }

    # This array will hold all data rows as PowerShell objects.
    $rows = @()

    # Start reading from row 2 because row 1 contains headers.
    for ($row = 2; $row -le $rowCount; $row++) {

        # Ordered keeps the columns in the same order as the Excel sheet.
        $obj = [ordered]@{}

        # Track whether the whole row is empty.
        $isEmptyRow = $true

        # Loop through each column in the current row.
        for ($col = 1; $col -le $colCount; $col++) {

            # Get the cell value.
            $value = [string]$range.Cells.Item($row, $col).Value2

            # If any value exists, this is not an empty row.
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                $isEmptyRow = $false
            }

            # Add the cell value to the object using the matching header name.
            $obj[$headers[$col - 1]] = $value.Trim()
        }

        # Only add non-empty rows.
        if (-not $isEmptyRow) {
            $rows += [pscustomobject]$obj
        }
    }

    # Release Excel COM objects for this range and sheet.
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($range) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) | Out-Null

    # Return the rows.
    return $rows
}

# ============================================================
# HELPER FUNCTION: WRITE ROWS TO AN EXCEL SHEET
# ============================================================

function Write-RowsToWorksheet {
    param(
        # The worksheet where data will be written.
        [object]$Worksheet,

        # The column headers to write in row 1.
        [string[]]$Headers,

        # The rows to write under the headers.
        [object[]]$Rows
    )

    # Format the whole sheet as text.
    # This helps avoid Excel converting IDs into numbers/scientific notation.
    $Worksheet.Cells.NumberFormat = "@"

    # Write the header row.
    for ($col = 1; $col -le $Headers.Count; $col++) {

        # Put each header into row 1.
        $Worksheet.Cells.Item(1, $col).Value2 = $Headers[$col - 1]

        # Make the header bold.
        $Worksheet.Cells.Item(1, $col).Font.Bold = $true
    }

    # Write the data rows.
    for ($row = 0; $row -lt $Rows.Count; $row++) {

        # Get the current PowerShell object.
        $currentRow = $Rows[$row]

        # Excel row number.
        # Add 2 because:
        #   row index starts at 0 in PowerShell
        #   row 1 is used by headers
        $excelRow = $row + 2

        # Write each column value.
        for ($col = 1; $col -le $Headers.Count; $col++) {

            # Get the current header.
            $headerName = $Headers[$col - 1]

            # Get the value from the current row using the header name.
            $cellValue = [string]$currentRow.$headerName

            # Write the value into the Excel cell.
            $Worksheet.Cells.Item($excelRow, $col).Value2 = $cellValue
        }
    }

    # Freeze the top row.
    # This makes the headers stay visible when scrolling.
    $Worksheet.Application.ActiveWindow.SplitRow = 1
    $Worksheet.Application.ActiveWindow.FreezePanes = $true

    # Auto-fit the columns so the text is visible.
    $Worksheet.Columns.AutoFit() | Out-Null
}

# ============================================================
# MAIN SCRIPT
# ============================================================

try {
    Write-DebugLog "Script started"

    # Confirm that the raw JSON file exists.
    # Again, this script does not read it yet, but it is useful to verify the project files are present.
    if (-not (Test-Path $jsonPath)) {
        throw "JSON file not found: $jsonPath"
    }

    # Confirm that the flattened Excel file exists.
    if (-not (Test-Path $flattenedPath)) {
        throw "Flattened Excel file not found: $flattenedPath"
    }

    Write-DebugLog "Path checks passed"

    # Start Excel through COM automation.
    $excel = New-Object -ComObject Excel.Application

    # Keep Excel hidden while the script runs.
    $excel.Visible = $false

    # Disable popups like overwrite confirmations.
    $excel.DisplayAlerts = $false

    Write-DebugLog "Opening flattened workbook"

    # Open your flattened schema workbook.
    $sourceWorkbook = $excel.Workbooks.Open($flattenedPath)

    # Read the ObjectTypes sheet into PowerShell objects.
    $objectTypes = Read-ExcelWorksheet -Workbook $sourceWorkbook -SheetName "ObjectTypes"

    # Read the Attributes sheet into PowerShell objects.
    $attributes = Read-ExcelWorksheet -Workbook $sourceWorkbook -SheetName "Attributes"

    Write-DebugLog "ObjectTypes rows read: $($objectTypes.Count)"
    Write-DebugLog "Attributes rows read: $($attributes.Count)"

    # ========================================================
    # BUILD OBJECT TYPE MAPPING ROWS
    # ========================================================

    # This array will become the ObjectTypeMapping sheet.
    $objectTypeMappingRows = @()

    # Loop through each target object type name from the config section.
    foreach ($targetName in $targetObjectTypeNames) {

        # Normalize the target name for easier matching.
        $normalizedTargetName = Normalize-Name $targetName

        # Find the matching object type from the ObjectTypes sheet.
        $match = $objectTypes |
            Where-Object { (Normalize-Name $_.ObjectTypeName) -eq $normalizedTargetName } |
            Select-Object -First 1

        # If no match is found, create a row with blank ID.
        if ($null -eq $match) {
            $objectTypeMappingRows += [pscustomobject]@{
                SourceTypeValue      = $targetName
                TargetObjectTypeID   = ""
                TargetObjectTypeName = $targetName
                MappingStatus        = "MissingObjectType"
            }
        }
        else {
            # If a match is found, create a completed mapping row.
            $objectTypeMappingRows += [pscustomobject]@{
                SourceTypeValue      = $targetName
                TargetObjectTypeID   = $match.ObjectTypeID
                TargetObjectTypeName = $match.ObjectTypeName
                MappingStatus        = "Matched"
            }
        }
    }

    Write-DebugLog "Object type mapping rows created: $($objectTypeMappingRows.Count)"

    # ========================================================
    # BUILD ATTRIBUTE MAPPING ROWS
    # ========================================================

    # This array will become the AttributeMapping sheet.
    $attributeMappingRows = @()

    # Loop through each matched object type.
    foreach ($objectTypeRow in $objectTypeMappingRows) {

        # If the object type ID is blank, skip it.
        # We cannot map attributes until we know the object type ID.
        if ([string]::IsNullOrWhiteSpace($objectTypeRow.TargetObjectTypeID)) {
            Write-DebugLog "Skipping attributes for missing object type: $($objectTypeRow.TargetObjectTypeName)"
            continue
        }

        # Store the object type ID and name for readability.
        $objectTypeID = [string]$objectTypeRow.TargetObjectTypeID
        $objectTypeName = [string]$objectTypeRow.TargetObjectTypeName

        # Get only attributes that belong to this object type.
        $attributesForObjectType = $attributes |
            Where-Object { ([string]$_.ObjectTypeID) -eq $objectTypeID }

        # Loop through each source column we want to map.
        foreach ($sourceColumn in $sourceColumnToAttributeCandidates.Keys) {

            # Get possible Assets attribute names for this source column.
            $candidateAttributeNames = $sourceColumnToAttributeCandidates[$sourceColumn]

            # Start with no match.
            $attributeMatch = $null

            # Try each possible attribute name.
            foreach ($candidateName in $candidateAttributeNames) {

                # Normalize the candidate name.
                $normalizedCandidate = Normalize-Name $candidateName

                # Look for an attribute with the same normalized name.
                $attributeMatch = $attributesForObjectType |
                    Where-Object { (Normalize-Name $_.AttributeName) -eq $normalizedCandidate } |
                    Select-Object -First 1

                # If we found a match, stop checking candidate names.
                if ($null -ne $attributeMatch) {
                    break
                }
            }

            # If no attribute match was found, create a row with blank attribute ID.
            if ($null -eq $attributeMatch) {
                $attributeMappingRows += [pscustomobject]@{
                    TargetObjectTypeID   = $objectTypeID
                    TargetObjectTypeName = $objectTypeName
                    SourceColumn         = $sourceColumn
                    AssetsAttributeID    = ""
                    AssetsAttributeName  = ""
                    Required             = ""
                    MappingStatus        = "MissingAttribute"
                }
            }
            else {
                # If an attribute match was found, create a completed mapping row.
                $attributeMappingRows += [pscustomobject]@{
                    TargetObjectTypeID   = $objectTypeID
                    TargetObjectTypeName = $objectTypeName
                    SourceColumn         = $sourceColumn
                    AssetsAttributeID    = $attributeMatch.AttributeID
                    AssetsAttributeName  = $attributeMatch.AttributeName
                    Required             = Convert-ToYesNo $attributeMatch.Required
                    MappingStatus        = "Matched"
                }
            }
        }
    }

    Write-DebugLog "Attribute mapping rows created: $($attributeMappingRows.Count)"

    # ========================================================
    # CREATE THE NEW MAPPING WORKBOOK
    # ========================================================

    Write-DebugLog "Creating mapping workbook"

    # If a previous mapping workbook already exists, delete it.
    # This keeps the script simple and avoids Excel overwrite prompts.
    if (Test-Path $mappingWorkbookPath) {
        Remove-Item -Path $mappingWorkbookPath -Force
        Write-DebugLog "Deleted existing mapping workbook"
    }

    # Create a new blank workbook.
    $mappingWorkbook = $excel.Workbooks.Add()

    # Use the first sheet for ObjectTypeMapping.
    $objectSheet = $mappingWorkbook.Worksheets.Item(1)
    $objectSheet.Name = "ObjectTypeMapping"

    # Add a second sheet for AttributeMapping.
    $attributeSheet = $mappingWorkbook.Worksheets.Add($null, $objectSheet)
    $attributeSheet.Name = "AttributeMapping"

    # Delete any extra default sheets Excel created.
    for ($i = $mappingWorkbook.Worksheets.Count; $i -ge 1; $i--) {

        # Get the current worksheet.
        $sheet = $mappingWorkbook.Worksheets.Item($i)

        # Keep only the two sheets we actually want.
        if ($sheet.Name -ne "ObjectTypeMapping" -and $sheet.Name -ne "AttributeMapping") {
            $sheet.Delete()
        }

        # Release the temporary sheet reference.
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) | Out-Null
    }

    # Define the headers for ObjectTypeMapping.
    $objectTypeHeaders = @(
        "SourceTypeValue",
        "TargetObjectTypeID",
        "TargetObjectTypeName",
        "MappingStatus"
    )

    # Define the headers for AttributeMapping.
    $attributeHeaders = @(
        "TargetObjectTypeID",
        "TargetObjectTypeName",
        "SourceColumn",
        "AssetsAttributeID",
        "AssetsAttributeName",
        "Required",
        "MappingStatus"
    )

    # Write the object type mapping rows to the first sheet.
    Write-RowsToWorksheet -Worksheet $objectSheet -Headers $objectTypeHeaders -Rows $objectTypeMappingRows

    # Activate the attribute sheet before freezing panes.
    # Excel freeze panes applies to the active window.
    $attributeSheet.Activate()

    # Write the attribute mapping rows to the second sheet.
    Write-RowsToWorksheet -Worksheet $attributeSheet -Headers $attributeHeaders -Rows $attributeMappingRows

    # Activate the first sheet so the workbook opens there.
    $objectSheet.Activate()

    # Save the new mapping workbook.
    $mappingWorkbook.SaveAs($mappingWorkbookPath)

    Write-DebugLog "Mapping workbook saved: $mappingWorkbookPath"

    # Print a simple success message for Power Automate Desktop.
    Write-Output "SUCCESS|MAPPING_WORKBOOK=$mappingWorkbookPath|LOG=$debugLog"
}
catch {
    # If anything fails, write the error to the debug log.
    Write-DebugLog "ERROR: $($_.Exception.Message)"

    # Print a clean error message for Power Automate Desktop.
    Write-Output "ERROR|MESSAGE=$($_.Exception.Message)|LOG=$debugLog"
}
finally {
    # Close the source flattened workbook if it was opened.
    if ($sourceWorkbook -ne $null) {
        $sourceWorkbook.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sourceWorkbook) | Out-Null
    }

    # Close the new mapping workbook if it was created.
    if ($mappingWorkbook -ne $null) {
        $mappingWorkbook.Close($true)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($mappingWorkbook) | Out-Null
    }

    # Quit Excel if it was started.
    if ($excel -ne $null) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }

    # Force cleanup of leftover Excel COM references.
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

