<#
    Purpose:
    This script reads the raw Jira Assets JSON export and creates a flattened Excel file.


    Input:
    assets_schema_5_export.json


    Output:
    assets_schema_5_flattened.xlsx


    Excel sheets created:
    1. ObjectTypes
       - One row per Assets object type


    2. Attributes
       - One row per fillable attribute belonging to an object type


    Debugging:
    This version writes a debug log to the TEMP folder so Power Automate Desktop
    can still show where the script failed even if normal output is missing.
#>


# Stop the script if a command fails.
$ErrorActionPreference = "Stop"


# ----------------------------
# PAD input variables
# ----------------------------


# Path to the raw JSON file created by the previous API export script.
# Power Automate Desktop should replace this placeholder.
$jsonPath = "%json_file_path%"


# Folder where the flattened Excel file should be saved.
# Power Automate Desktop should replace this placeholder.
$outputFolder = "%folder_path%"

# Full path for the output Excel file.
$outputPath = Join-Path $outputFolder "assets_schema_5_flattened.xlsx"


# ----------------------------
# Debug logging setup
# ----------------------------


$debugLog = Join-Path $outputFolder "assets_flatten_debug.log"


function Write-DebugLog {
    param (
        [string]$Message
    )


    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $debugLog -Value "[$timestamp] $Message"
}


Write-DebugLog "Script started"


# ----------------------------
# Excel COM variables
# ----------------------------


$excel = $null
$workbook = $null
$objectTypesSheet = $null
$attributesSheet = $null


try {
    Write-DebugLog "jsonPath = $jsonPath"
    Write-DebugLog "outputFolder = $outputFolder"
    Write-DebugLog "outputPath = $outputPath"


    $percentChar = [string][char]37

    # Check whether PAD actually replaced the placeholders.
    if ($jsonPath.Contains($percentChar)) {
        throw "PAD did not replace json_file_path. Current value: $jsonPath"
    }


    if ($outputFolder.Contains($percentChar)) {
        throw "PAD did not replace folder_path. Current value: $outputFolder"
    }


    # Check that the JSON file exists.
    if (-not (Test-Path -Path $jsonPath)) {
        throw "JSON file not found: $jsonPath"
    }


    # Check that the output folder exists.
    if (-not (Test-Path -Path $outputFolder)) {
        throw "Output folder not found: $outputFolder"
    }


    Write-DebugLog "Path checks passed"


    # If the output file already exists, remove it so SaveAs does not get blocked.
    if (Test-Path -Path $outputPath) {
        Write-DebugLog "Existing output file found. Removing: $outputPath"
        Remove-Item -Path $outputPath -Force
    }


    # Read the full JSON file as one string.
    Write-DebugLog "Reading JSON file"
    $jsonText = Get-Content -Path $jsonPath -Raw
    Write-DebugLog "JSON text length: $($jsonText.Length)"


    if ([string]::IsNullOrWhiteSpace($jsonText)) {
        throw "JSON file is empty: $jsonPath"
    }


    # Convert the JSON text into PowerShell objects.
    Write-DebugLog "Converting JSON text to PowerShell object"
    $assetSchema = $jsonText | ConvertFrom-Json


    if ($null -eq $assetSchema) {
        throw "ConvertFrom-Json returned null. Check whether the JSON file is valid."
    }


    $objectTypeCount = @($assetSchema).Count
    Write-DebugLog "Object type count: $objectTypeCount"


    # Start Excel through PowerShell.
    Write-DebugLog "Starting Excel COM application"
    $excel = New-Object -ComObject Excel.Application


    # Keep Excel hidden while the script runs.
    $excel.Visible = $false


    # Prevent Excel popups during script execution.
    $excel.DisplayAlerts = $false


    # Create a new Excel workbook.
    Write-DebugLog "Creating Excel workbook"
    $workbook = $excel.Workbooks.Add()


    # Use the first worksheet for ObjectTypes.
    $objectTypesSheet = $workbook.Worksheets.Item(1)
    $objectTypesSheet.Name = "ObjectTypes"


    # Add a second worksheet for Attributes.
    $attributesSheet = $workbook.Worksheets.Add()
    $attributesSheet.Name = "Attributes"


    # -------------------------------
    # Create ObjectTypes sheet headers
    # -------------------------------


    $objectTypeHeaders = @(
        "ObjectSchemaID",
        "SchemaLabel",
        "ObjectTypeID",
        "ObjectTypeName",
        "ParentObjectTypeID",
        "Inherited",
        "AbstractObjectType",
        "ObjectCount",
        "Created",
        "Updated"
    )


    Write-DebugLog "Writing ObjectTypes headers"


    # Write each ObjectTypes header into row 1.
    for ($i = 0; $i -lt $objectTypeHeaders.Count; $i++) {
        $objectTypesSheet.Cells.Item(1, $i + 1).Value2 = $objectTypeHeaders[$i]
    }


    # Start writing ObjectTypes data on row 2.
    $objectTypeRow = 2


    Write-DebugLog "Writing ObjectTypes rows"


    # Loop through each object type in the JSON.
    foreach ($objectType in $assetSchema) {


        # Write one row for the current object type.
        $objectTypesSheet.Cells.Item($objectTypeRow, 1).Value2 = [string]$objectType.objectSchemaId
        $objectTypesSheet.Cells.Item($objectTypeRow, 2).Value2 = $objectType.schemaLabel
        $objectTypesSheet.Cells.Item($objectTypeRow, 3).Value2 = [string]$objectType.objectTypeId
        $objectTypesSheet.Cells.Item($objectTypeRow, 4).Value2 = $objectType.objectTypeName
        $objectTypesSheet.Cells.Item($objectTypeRow, 5).Value2 = [string]$objectType.parentObjectTypeId
        $objectTypesSheet.Cells.Item($objectTypeRow, 6).Value2 = [string]$objectType.inherited
        $objectTypesSheet.Cells.Item($objectTypeRow, 7).Value2 = [string]$objectType.abstractObjectType
        $objectTypesSheet.Cells.Item($objectTypeRow, 8).Value2 = [string]$objectType.objectCount
        $objectTypesSheet.Cells.Item($objectTypeRow, 9).Value2 = $objectType.created
        $objectTypesSheet.Cells.Item($objectTypeRow, 10).Value2 = $objectType.updated


        # Move to the next Excel row.
        $objectTypeRow++
    }


    Write-DebugLog "ObjectTypes rows written: $($objectTypeRow - 2)"


    # ----------------------------
    # Create Attributes sheet headers
    # ----------------------------


    $attributeHeaders = @(
        "ObjectTypeID",
        "ObjectTypeName",
        "AttributeID",
        "AttributeName",
        "AttributeType",
        "Required",
        "Editable",
        "System",
        "Label",
        "ReferenceObjectTypeID",
        "ReferenceObjectTypeName",
        "MinimumCardinality",
        "MaximumCardinality",
        "Options",
        "Position"
    )


    Write-DebugLog "Writing Attributes headers"


    # Write each Attributes header into row 1.
    for ($i = 0; $i -lt $attributeHeaders.Count; $i++) {
        $attributesSheet.Cells.Item(1, $i + 1).Value2 = $attributeHeaders[$i]
    }


    # Start writing Attributes data on row 2.
    $attributeRow = 2
    $skippedAttributeCount = 0


    Write-DebugLog "Writing fillable Attributes rows"


    # Loop through each object type again.
    foreach ($objectType in $assetSchema) {


        # Get the list of attributes for the current object type.
        # In this JSON, attributes are stored inside attributes.value.
        $attributes = @($objectType.attributes.value)


        Write-DebugLog "ObjectType '$($objectType.objectTypeName)' has raw attribute count: $($attributes.Count)"


        # Loop through each attribute for the current object type.
        foreach ($attribute in $attributes) {


            # Skip blank/null attribute records just in case.
            if ($null -eq $attribute) {
                $skippedAttributeCount++
                continue
            }


            # Only keep attributes that a user/script can actually fill in.
            # Skip system-generated, hidden, or non-editable fields.
            if (
                $attribute.editable -ne $true -or
                $attribute.system -eq $true -or
                $attribute.hidden -eq $true
            ) {
                $skippedAttributeCount++
                continue
            }


            # Decide the readable attribute type.
            # Normal fields usually have defaultType.name.
            # Reference fields usually have referenceObjectTypeId.
            if ($attribute.defaultType.name) {
                $attributeType = $attribute.defaultType.name
            }
            elseif ($attribute.referenceObjectTypeId) {
                $attributeType = "Reference"
            }
            elseif ($attribute.type -eq 7) {
                $attributeType = "Status"
            }
            else {
                $attributeType = "TypeCode_$($attribute.type)"
            }


            # Required means minimumCardinality is 1 or greater.
            $required = ($attribute.minimumCardinality -ge 1)


            # Write one row for this object type + attribute relationship.
            $attributesSheet.Cells.Item($attributeRow, 1).Value2 = $objectType.objectTypeId
            $attributesSheet.Cells.Item($attributeRow, 2).Value2 = $objectType.objectTypeName
            $attributesSheet.Cells.Item($attributeRow, 3).Value2 = [string]$attribute.id
            $attributesSheet.Cells.Item($attributeRow, 4).Value2 = $attribute.name
            $attributesSheet.Cells.Item($attributeRow, 5).Value2 = $attributeType
            $attributesSheet.Cells.Item($attributeRow, 6).Value2 = [string]$required
            $attributesSheet.Cells.Item($attributeRow, 7).Value2 = [string]$attribute.editable
            $attributesSheet.Cells.Item($attributeRow, 8).Value2 = [string]$attribute.system
            $attributesSheet.Cells.Item($attributeRow, 9).Value2 = [string]$attribute.label
            $attributesSheet.Cells.Item($attributeRow, 10).Value2 = [string]$attribute.referenceObjectTypeId
            $attributesSheet.Cells.Item($attributeRow, 11).Value2 = $attribute.referenceObjectType.name
            $attributesSheet.Cells.Item($attributeRow, 12).Value2 = [string]$attribute.minimumCardinality
            $attributesSheet.Cells.Item($attributeRow, 13).Value2 = [string]$attribute.maximumCardinality
            $attributesSheet.Cells.Item($attributeRow, 14).Value2 = $attribute.options
            $attributesSheet.Cells.Item($attributeRow, 15).Value2 = [string]$attribute.position


            # Move to the next Excel row.
            $attributeRow++
        }
    }


    Write-DebugLog "Fillable attribute rows written: $($attributeRow - 2)"
    Write-DebugLog "Skipped attribute count: $skippedAttributeCount"


    # ----------------------------
    # Basic Excel formatting
    # ----------------------------


    Write-DebugLog "Applying Excel formatting"


    # Bold the header row on both sheets.
    $objectTypesSheet.Rows.Item(1).Font.Bold = $true
    $attributesSheet.Rows.Item(1).Font.Bold = $true


    # Auto-size columns so the data is easier to read.
    $objectTypesSheet.Columns.AutoFit() | Out-Null
    $attributesSheet.Columns.AutoFit() | Out-Null


    # Delete extra blank sheets that Excel may have created.
    Write-DebugLog "Deleting extra blank Excel sheets"


    foreach ($sheet in @($workbook.Worksheets)) {
        if ($sheet.Name -ne "ObjectTypes" -and $sheet.Name -ne "Attributes") {
            $sheet.Delete()
        }
    }


    # Save as .xlsx.
    # 51 = Excel Open XML Workbook format.
    Write-DebugLog "Saving workbook"
    $workbook.SaveAs($outputPath, 51)


    Write-DebugLog "Workbook saved successfully"


    # Close the workbook.
    $workbook.Close($true)
    $workbook = $null


    # Quit Excel.
    $excel.Quit()


    Write-DebugLog "Excel closed successfully"


    # Output final result in a PAD-friendly format.
    Write-Output "SUCCESS|OUTPUT=$outputPath|LOG=$debugLog"
}
catch {
    Write-DebugLog "ERROR MESSAGE: $($_.Exception.Message)"
    Write-DebugLog "ERROR LINE NUMBER: $($_.InvocationInfo.ScriptLineNumber)"
    Write-DebugLog "ERROR COMMAND: $($_.InvocationInfo.Line)"


    # Output error in a PAD-friendly format.
    Write-Output "ERROR|MESSAGE=$($_.Exception.Message)|LOG=$debugLog"


    exit 1
}
finally {
    Write-DebugLog "Cleanup started"


    # Close workbook if it is still open.
    if ($null -ne $workbook) {
        try {
            $workbook.Close($false)
            Write-DebugLog "Workbook closed during cleanup"
        }
        catch {
            Write-DebugLog "Workbook cleanup failed: $($_.Exception.Message)"
        }
    }


    # Quit Excel if it is still open.
    if ($null -ne $excel) {
        try {
            $excel.Quit()
            Write-DebugLog "Excel quit during cleanup"
        }
        catch {
            Write-DebugLog "Excel cleanup failed: $($_.Exception.Message)"
        }
    }


    # Release Excel COM objects from memory.
    if ($null -ne $attributesSheet) {
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($attributesSheet) | Out-Null
            Write-DebugLog "Released attributesSheet COM object"
        }
        catch {}
    }


    if ($null -ne $objectTypesSheet) {
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($objectTypesSheet) | Out-Null
            Write-DebugLog "Released objectTypesSheet COM object"
        }
        catch {}
    }


    if ($null -ne $workbook) {
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
            Write-DebugLog "Released workbook COM object"
        }
        catch {}
    }


    if ($null -ne $excel) {
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
            Write-DebugLog "Released excel COM object"
        }
        catch {}
    }


    # Force garbage collection to fully clean up Excel in the background.
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()


    Write-DebugLog "Cleanup finished"
}