# Inventory Automation Workflow

This document outlines the current workflow for automating inventory management between the source data file, Atlassian Assets, Jira, and future reporting tools.

---

## High-Level Flow

```mermaid
flowchart TD
    A([START]) --> B[1. Confirm Source Data]
    B --> C[2. Confirm Assets Schema]
    C --> D[3. Build Mapping Files]
    D --> E[4. Read Source Rows]
    E --> F[5. Validate Each Row]

    F -->|Invalid row| F1[Set SyncStatus = Error<br/>Write ErrorMessage<br/>Skip Row]
    F -->|Valid row| G[6. Build Assets Payload]

    G --> H[7. Check If Asset Already Exists]

    H -->|Found| H1[Update Existing Assets Object]
    H -->|Not Found| H2[Create New Assets Object]

    H1 --> I[8. Write Asset Result Back]
    H2 --> I

    I --> J[9. Create or Update Jira Issue]

    J -->|JiraIssueKey Exists| J1[Update Jira Issue]
    J -->|JiraIssueKey Blank| J2[Create Jira Issue]

    J1 --> K[10. Link Asset and Jira]
    J2 --> K

    K --> L[11. Final Write-Back]
    L --> M[12. Reporting / Display]
    M --> N([END])
```

---

## 1. Confirm Source Data

### Source Options

- Mock CSV or Excel file for testing
- Excel or SharePoint List later

### Required Source Columns

| Column | Purpose |
|---|---|
| `SourceRecordID` | Unique identifier from the source data |
| `ItemName` | Main name/title of the asset |
| `Category` | Used for mapping to Assets object types |
| `TargetObjectTypeID` | Assets object type where the row should be created |
| `AssetObjectID` | Stores the created or matched Assets object ID |
| `JiraIssueKey` | Stores the linked Jira issue key |
| `SyncStatus` | Tracks Pending, Success, or Error |
| `ErrorMessage` | Stores error details when something fails |

---

## 2. Confirm Assets Schema

Use the Assets API to:

- Get available schemas and object types
- Get attributes for each object type
- Save important IDs:
  - Object type IDs
  - Attribute IDs
  - Required attribute fields

---

## 3. Build Mapping Files

### Object Type Mapping

Maps source values to Assets object types.

| Source Value | TargetObjectTypeID |
|---|---|
| Laptop | 101 |
| Monitor | 102 |
| Printer | 103 |

### Attribute Mapping

Maps source columns to Assets attributes.

| Source Column | Assets Attribute |
|---|---|
| `ItemName` | Name |
| `SerialNumber` | Serial Number |
| `Location` | Location |

---

## 4. Read Source Rows

Only process rows where:

- `SyncStatus = Pending`
- `AssetObjectID` is blank
- `JiraIssueKey` is blank
- `ModifiedDate > LastSyncDate`

---

## 5. Validate Each Row

Check that each row has:

- `SourceRecordID`
- `ItemName`
- `TargetObjectTypeID`
- Required mapped attributes

If invalid:

```text
Set SyncStatus = Error
Write ErrorMessage
Skip row
```

---

## 6. Build Assets Payload

For each valid row:

- Use `TargetObjectTypeID`
- Load the matching attribute mappings
- Pull values from the source row
- Build the JSON body for the Assets API

---

## 7. Check If Asset Already Exists

Search Assets by:

- `SourceRecordID`, or
- `SerialNumber`

Then decide:

```text
If found:
    Update existing Assets object

If not found:
    Create new Assets object
```

---

## 8. Write Asset Result Back

Save the result back to the source file:

| Field | Value |
|---|---|
| `AssetObjectID` | Created or updated asset ID |
| `LastSyncDate` | Current timestamp |
| `SyncStatus` | Success or Error |
| `ErrorMessage` | Error details if failed |

---

## 9. Create or Update Jira Issue

```text
If JiraIssueKey exists:
    Update Jira issue

If JiraIssueKey is blank:
    Create Jira issue
```

---

## 10. Link Asset and Jira

Create a two-way relationship:

- Write `AssetObjectID` into a Jira field
- Write `JiraIssueKey` into an Assets attribute

---

## 11. Final Write-Back

Save final sync results:

| Field | Purpose |
|---|---|
| `JiraIssueKey` | Jira issue key |
| `JiraIssueID` | Jira internal issue ID |
| `AssetObjectID` | Assets object ID |
| `SyncStatus` | Final success/error state |
| `LastSyncDate` | Final sync timestamp |
| `ErrorMessage` | Cleared if successful |

---

## 12. Reporting / Display

Possible reporting layers:

- Jira dashboard
- Confluence page
- Later: Power BI or other reports

---

## Final Goal

The final goal is to create a repeatable Power Automate Desktop workflow that can read inventory data, validate it, create or update Atlassian Assets objects, create or update Jira issues, link them together, and write results back to the source file.

### Task List

Task 1:
  Confirm source columns.
  Create 5-row mock dataset.
  Confirm PowerShell can read the file.

Task 2:
  Pull object types from Assets.
  Pull attributes for each object type.
  Save ObjectTypeIDs and AttributeIDs.

Task 3:
  Build ObjectTypeMapping.csv.
  Build AttributeMapping.csv.
  Validate required fields.

Task 4:
  Build Assets JSON payload from one row.
  Create one object in Assets.
  Write AssetObjectID back to source.

Task 5:
  Add search-before-create logic.
  Update existing objects instead of duplicating.

Task 6:
  Add Jira issue creation.
  Save JiraIssueKey back to source.

Task 7:
  Link Jira issue and Assets object.

Task 8:
  Add logging, error handling, and documentation.

### Pending Implementation

#### Task 3 (In progress JC)

After obtaining the pre-mapping file, the "flattened" file, we need to create a mapping file that will essentially tell the mock data all the information needed to know where to place it in assets. This will possibly be two files, or one with two sheets. The source file (source of truth, in this case the `MockAssets.csv`) will need to be in a structure that mostly agrees with the information inside `issue_details.md`. We also need to ensure that the required information is at the minimum being filled in

#### Task 4

Prove that one source row can successfully become one Atlassian Assets object.

First, take one valid row from the source file and use its TargetObjectTypeID to decide what type of Assets object should be created. Then use the attribute mapping file to match each source column to the correct Assets attribute ID.

After the row values are matched to the correct Assets attributes, build the JSON payload required by the Assets API. This payload should include the object type ID and the list of attributes with their values.

Once the payload is built, send it to the Assets create-object API endpoint. If the request succeeds, Atlassian Assets will return the newly created object information, including the AssetObjectID.

Finally, write the returned AssetObjectID back into the original source row. Also update the row’s sync fields, such as setting SyncStatus to Success, updating LastSyncDate, and clearing ErrorMessage.

In the end, the workflow should be able to create one Assets object from one source row and record the result back in the source file.

#### Task 5

The goal is to prevent duplicate Assets objects from being created.

Before creating a new object, the workflow should search Atlassian Assets to check whether the asset already exists. The search can be based on a unique field such as SourceRecordID, SerialNumber, or another reliable identifier.

If a matching object is found, the workflow should update the existing Assets object instead of creating a new one. If no match is found, the workflow should continue with the create-object process from Task 4.

This task adds the decision logic needed to choose between creating a new object and updating an existing one.

In the end, the workflow should be able to process one row without creating duplicates.

#### Task 6

The goal is to add Jira issue creation after the Assets object has been created or updated.

Once the Assets step succeeds, the workflow should use the source row and asset result to build a Jira issue payload. This should include the required Jira fields, such as project key, issue type, summary, and description.

If the source row does not already have a JiraIssueKey, the workflow should create a new Jira issue. After the issue is created, Jira will return the issue key and internal issue ID.

The workflow should then write the JiraIssueKey and JiraIssueID back into the source row.

In the end, one source row should be able to create or update an Assets object and then create a related Jira issue.

#### Task 7

The goal is to connect the Jira issue and the Assets object together.

After both the Assets object and Jira issue exist, the workflow should create a link between them. This can be done by writing the AssetObjectID into a Jira Assets custom field and/or writing the JiraIssueKey into an Assets attribute.

This step makes the relationship visible from both sides. The Jira issue should show which asset it belongs to, and the Assets object should show which Jira issue is connected to it.

The workflow should also confirm that both IDs were saved correctly before marking the row as fully successful.

In the end, the system should create a complete connection between the source row, the Assets object, and the Jira issue.

#### Task 8

The goal is to improve reliability, logging, and documentation.

The workflow should handle common errors, such as missing required fields, failed API calls, invalid mappings, or duplicate matches. When an error happens, the workflow should write a clear message into ErrorMessage and set SyncStatus to Error.

Logging should also be added so that each major step can be reviewed later. This includes when the script starts, which row is being processed, what API action was attempted, whether it succeeded or failed, and where any error occurred.

The documentation should be updated to explain how the workflow works, what files are required, what columns are expected, and how to troubleshoot common problems.

In the end, the workflow should be easier to test, debug, and hand off to someone else.