$path = "Z:\01 Intern Work\Current Interns\Jorge C\project\MockAssets.csv"

function New-MockAssetRow {
    param (
        [string]$SourceRecordID,
        [string]$ItemName,
        [string]$Category,
        [string]$Color = "",
        [string]$Model = "",
        [string]$SerialNumber,
        [string]$Location,
        [string]$Status = "Available",
        [string]$Owner = "",
        [string]$Notes = "",
        [string]$CreateJiraItem = "No",
        [string]$CreateAssetObject = "Yes"
    )

    [PSCustomObject][ordered]@{
        SourceRecordID    = $SourceRecordID
        ItemName          = $ItemName
        Category          = $Category
        Color             = $Color
        Model             = $Model
        SerialNumber      = $SerialNumber
        Location          = $Location
        Status            = $Status
        Owner             = $Owner
        Notes             = $Notes
        CreateJiraItem    = $CreateJiraItem
        CreateAssetObject = $CreateAssetObject
        JiraIssueKey      = ""
        JiraIssueID       = ""
        AssetObjectID     = ""
        LastSyncDate      = ""
        SyncStatus        = "Pending"
        ErrorMessage      = ""
    }
}

$rows = @(
    New-MockAssetRow `
        -SourceRecordID "SRC-0001" `
        -ItemName "Nitrile Gloves" `
        -Category "Product" `
        -Model "Box of 100" `
        -SerialNumber "PROD-001" `
        -Location "OC HCA Main Distribution Center (Santa Ana) - Aisle A, Shelf 1" `
        -Status "In Stock" `
        -Notes "Sample consumable item."

    New-MockAssetRow `
        -SourceRecordID "SRC-0002" `
        -ItemName "N95 Respirator Masks" `
        -Category "Product" `
        -Model "Box of 50" `
        -SerialNumber "PROD-002" `
        -Location "North Field Station (Anaheim) - Aisle B, Shelf 2" `
        -Status "In Stock" `
        -Notes "Sample PPE supply item."

    New-MockAssetRow `
        -SourceRecordID "SRC-0003" `
        -ItemName "Tetanus Vaccines" `
        -Category "Perishable Product" `
        -Model "Cold storage, exp. 2025" `
        -SerialNumber "PER-001" `
        -Location "Coastal Staging Area (Dana Point) - Cold Storage Unit 1" `
        -Status "Cold Storage" `
        -Notes "Sample perishable medical item."

    New-MockAssetRow `
        -SourceRecordID "SRC-0004" `
        -ItemName "Insulin Pens" `
        -Category "Perishable Product" `
        -Model "Cold storage, exp. 2025" `
        -SerialNumber "PER-002" `
        -Location "Central Supply Depot (Garden Grove) - Cold Storage Unit 1" `
        -Status "Cold Storage" `
        -Notes "Sample perishable medical item."

    New-MockAssetRow `
        -SourceRecordID "SRC-0005" `
        -ItemName "Philips HeartStart Defibrillator" `
        -Category "Equipment" `
        -SerialNumber "EQP-001" `
        -Location "South Field Station (Irvine) - Outdoor Staging Pad" `
        -Status "Available" `
        -Notes "Tracked equipment sample."

    New-MockAssetRow `
        -SourceRecordID "SRC-0006" `
        -ItemName "Zoll Ventilator" `
        -Category "Equipment" `
        -SerialNumber "EQP-002" `
        -Location "East Field Station (Tustin) - Refrigerated Trailer Bay" `
        -Status "Available" `
        -Notes "Tracked equipment sample that should create a Jira item." `
        -CreateJiraItem "Yes"

    New-MockAssetRow `
        -SourceRecordID "SRC-0007" `
        -ItemName "Motorola Emergency Radio" `
        -Category "Equipment" `
        -Model "APX Series" `
        -SerialNumber "EQP-003" `
        -Location "Coastal Staging Area (Dana Point) - Quarantine Storage Room" `
        -Status "Available" `
        -Notes "Radio equipment sample."

    New-MockAssetRow `
        -SourceRecordID "SRC-0008" `
        -ItemName "2022 Ford E-450 Ambulance" `
        -Category "Vehicle" `
        -Model "2022 Ford E-450 Ambulance" `
        -SerialNumber "VEH-001" `
        -Location "Coastal Staging Area (Dana Point) - Vehicle Bay" `
        -Status "Available" `
        -Notes "Vehicle sample." `
        -CreateJiraItem "Yes"

    New-MockAssetRow `
        -SourceRecordID "SRC-0009" `
        -ItemName "2021 Freightliner Mobile Command" `
        -Category "Vehicle" `
        -Model "2021 Freightliner Mobile Command" `
        -SerialNumber "VEH-002" `
        -Location "Central Supply Depot (Garden Grove) - Vehicle Bay" `
        -Status "Available" `
        -Notes "Mobile command vehicle sample." `
        -CreateJiraItem "Yes"

    New-MockAssetRow `
        -SourceRecordID "SRC-0010" `
        -ItemName "Hand Sanitizer" `
        -Category "Product" `
        -Model "Gallon jugs" `
        -SerialNumber "PROD-003" `
        -Location "OC HCA Main Distribution Center (Santa Ana) - Hazmat Cage" `
        -Status "In Stock" `
        -Notes "Sample consumable supply item."
)

$rows | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8

Write-Host "Created mock source file: $path"

