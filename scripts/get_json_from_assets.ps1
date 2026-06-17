<#  
    This script is for getting the cloud ID. The output to this script inside powershell should be 
    <cloud_id>
    ... This is necessary in order for the final script to initialize the cloud id accordingly.
#>

$response = Invoke-RestMethod -Uri "http://ochca.atlassian.net/_edge/tenant_info" -Method Get

$response.cloudId.ToString().Trim()

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

<# 
    This script is for getting the workspace ID. The output to this script inside powershell should be
    <workspace_id>
    ... This is necessary in order for the final script to initialize the workspace id accordingly.
#>

$email = "%email%"
$token = "%api_token%"
$jiraUrl = "https://ochca.atlassian.net/rest/servicedeskapi/assets/workspace"

$pair = "${email}:$token"
$encoded = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($pair))

$headers = @{
    Authorization = "Basic $encoded"
    Accept        = "application/json"
}

$response = Invoke-RestMethod -Uri $jiraUrl -Method Get -Headers $headers
$response.values[0].workspaceId.ToString().Trim()

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

<# 
    Purpose:
    This script exports all object types from one Jira Assets schema,
    then gets the attributes for each object type,
    then saves everything into one JSON file.

    Result:
    You will get a JSON structure like:

    [
        {
            objectTypeId: 110,
            objectTypeName: "Product",
            parentObjectTypeId: 55,
            attributes: [...]
        }
    ]

    This gives us a clean raw export before we decide how to organize Excel.
#>

# Your Atlassian account email.
$email = "%email%"

# API token passed in from Power Automate Desktop.
$token = "%api_token%"

# Cloud ID and Workspace ID passed in from previous PowerShell/PAD steps.
$cloudId = "%cloud_id%".Trim()
$workspaceId = "%workspace_id%".Trim()

# The Assets schema we are exporting.
# Based on your output, schema 5 = Inventory Warehouse HDP.
$objectSchemaId = "5"

# Build the base Assets API URL once so we do not repeat it everywhere.
$baseUrl = "https://api.atlassian.com/ex/jira/$cloudId/jsm/assets/workspace/$workspaceId/v1"

# Build Basic Auth header using email:token.
$pair = "${email}:$token"
$encoded = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($pair))

# Headers needed for the API call.
$headers = @{
    Authorization  = "Basic $encoded"
    Accept         = "application/json"
    "Content-Type" = "application/json"
}

# Endpoint that gets all object types inside the selected schema.
$objectTypesUrl = "$baseUrl/objectschema/$objectSchemaId/objecttypes"

try {
    # Get all object types in the schema.
    $objectTypes = Invoke-RestMethod -Uri $objectTypesUrl -Method Get -Headers $headers

    # This array will hold the final cleaned export.
    $export = @()

    # Loop through each object type returned from Assets.
    foreach ($objectType in $objectTypes) {

        # Grab the current object type ID.
        $objectTypeId = $objectType.id

        # Endpoint that gets attributes for this specific object type.
        $attributesUrl = "$baseUrl/objecttype/$objectTypeId/attributes"

        # Get attributes for the current object type.
        $attributes = Invoke-RestMethod -Uri $attributesUrl -Method Get -Headers $headers

        # Build a cleaner custom object that combines object type info with attributes.
        $exportItem = [PSCustomObject]@{
            workspaceId               = $objectType.workspaceId
            objectSchemaId            = $objectType.objectSchemaId
            schemaLabel               = $objectType.schemaLabel

            objectTypeId              = $objectType.id
            objectTypeName            = $objectType.name
            parentObjectTypeId        = $objectType.parentObjectTypeId

            inherited                 = $objectType.inherited
            abstractObjectType        = $objectType.abstractObjectType
            parentObjectTypeInherited = $objectType.parentObjectTypeInherited

            objectCount               = $objectType.objectCount
            created                   = $objectType.created
            updated                   = $objectType.updated

            attributes                = $attributes
        }

        # Add this object type + its attributes to the final export array.
        $export += $exportItem
    }

    # Convert final export to JSON.
    $jsonOutput = $export | ConvertTo-Json -Depth 50

    # Choose where to save the JSON file.
    # You can change this path if Power Automate Desktop needs a specific folder.
    $outputPath = "%folder_path%\assets_schema_5_export.json"

    # Save JSON to file.
    $jsonOutput | Out-File -FilePath $outputPath -Encoding UTF8

    # Print useful status messages for PAD.
    Write-Output "SUCCESS"
    Write-Output "Exported object types and attributes."
    Write-Output "Output file: $outputPath"
}
catch {
    # Print a clear error marker for PAD.
    Write-Output "ERROR"

    # Print the main PowerShell exception message.
    Write-Output $_.Exception.Message

    # If the API returned a response body, print it too.
    if ($_.Exception.Response) {
        $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
        $errorBody = $reader.ReadToEnd()

        Write-Output "ERROR BODY:"
        Write-Output $errorBody
    }
}



