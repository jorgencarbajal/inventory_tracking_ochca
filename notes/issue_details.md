# Design Planning

This file is a replica of the information provided by the boss man inside the issue provided for the scope of this project.

## Description

The assigned design specifically says to read source rows, validate data, create/update Assets, create/update Jira work items, link them, and store returned ID's back in the source.

### High-Level Design Options

**Option C:** Excel/SharePoint list -> Power Automate -> Assets Objects + Jira Issues

Best when each real-world item needs both:

- a structured database record, and
- an operational work item for workflow, assignments, or lifecycle tracking.

### Recommended Approach

Recommended: Option C

Use Power Automate to:

1. Read rows from an approved Microsoft source.
2. Validate the row data.
3. Create or update an Assets object where needed.
4. Create or update a related Jira work item.
5. Store the returned Atlassian IDs back into the Microsoft source.
6. Allow Confluence and Jira dashboards to display the results.

This gives both database structure and workflow control.

### Proposed Architecture

#### Source Layer

One of the following:

- Excel Online table in OneDrive
- Excel Online table in SharePoint
- Microsoft List
- SharePoint List

#### Automation Layer

Power Automate cloud flow

#### Target Systems

- Jira Cloud project for work items
- JSM Assets schema for object storage
- Confluence pages for reporting and display

#### Identity / Security

- Atlassian API token
- Service account or dedicated automation account preferred
- Secure credentials stored in Power Automate connection or secret mechanism

## Data Model

### Source Record Example

Each row should represent one item.

### Suggested columns

- SourceRecordID
- ItemName
- Category
- Color
- Model
- SerialNumber
- Location
- Status
- Owner
- Notes
- CreateJiraItem (Yes/No)
- CreateAssetObject (Yes/No)
- JiraIssueKey
- JiraIssueID
- AssetObjectID
- LastSyncDate
- SyncStatus
- ErrorMessage

### Example

- SourceRecordID: CCH-001
- ItemName: Couch A
- Category: Couch
- Color: Blue
- Model: Waiting Room Large
- SerialNumber: SN-10001
- Location: Building 1 Lobby
- Status: In Service
- Owner: Facilities

## Jira Design

### Suggested Jira Project Type

Use one dedicated project for imported operational records.

### Possible issue types

- Asset Record
- Inventory Item
- Maintenance Item
- Task
- Request

### Suggested Jira Fields

#### Standard fields

- Summary
- Description
- Issue Type
- Labels
- Assignee

#### Custom fields as needed

- Color
- Category
- Model
- Serial Number
- Location
- Asset Object ID
- Source Record ID
- Lifecycle Status

Important design note: Only create custom fields that are truly needed in Jira workflows or reporting. Avoid creating too many project-specific custom fields if the data can live in Assets instead.

## Assets Design

### Suggested Object Schema

Schema name example:

Facilities Inventory

### Suggested Object Types

- Couch
- Chair
- Desk
- Device
- Vehicle
- Room
- Site

### Example Couch Object Attributes

- Name
- Category
- Color
- Model
- Serial Number
- Location
- Status
- Owner
- Related Jira Issue
- Source Record ID

## Flow Design in Power Automate

### Flow Trigger Options

##### Manual trigger

- Best for controlled test imports.

##### Scheduled trigger

- Best for nightly or hourly sync.

##### Row created/modified trigger

- Best for near real-time updates from SharePoint or Lists.

### Recommended Initial Trigger

Start with a manual or scheduled flow for easier control.

## Detailed Flow Logic

### Step 1: Read Source Records

Connect to Excel Online or SharePoint List.

Retrieve rows where:

- JiraIssueKey is blank, or
- AssetObjectID is blank, or
- SyncStatus = Pending, or
- ModifiedDate > LastSyncDate

### Step 2: Validate Required Fields

Check for required values such as:

- ItemName
- Category
- Status
- Color if required by the object type

If validation fails:

- set SyncStatus = Error
- write an ErrorMessage
- do not continue for that row

### Step 3: Check for Existing Asset Object

Use a unique value such as:

- SourceRecordID, or
- SerialNumber

If an object already exists:

- update it

Else:

- create it

### Step 4: Check for Existing Jira Work Item

If JiraIssueKey already exists in source:

- update issue

Else:

- create new issue

### Step 5: Link Asset and Jira Item

Where applicable:

- write the Asset Object ID into Jira custom field
- write the Jira issue key into the asset object attribute

### Step 6: Write Results Back to Source

Update source row with:

- JiraIssueKey
- JiraIssueID
- AssetObjectID
- LastSyncDate
- SyncStatus = Success
- ErrorMessage cleared

## Power Automate Components

### Core Actions Needed

- Trigger
- List rows present in a table / Get items
- Apply to each
- Condition
- Compose
- HTTP action
- Parse JSON
- Update row / Update item
- Scope for error handling

### Why HTTP Action Matters

Power Automate may not have all Jira/Assets functions natively. The HTTP action allows direct calls to Atlassian REST endpoints without installing software.

## Authentication Design

### Recommended Method

- Atlassian account email
- Atlassian API token
- Basic authentication over HTTPS

### Storage

Store credentials in the Power Automate connection securely. If available in your Microsoft environment, a stronger design would use an approved secrets store.

### Service Account Recommendation

Use a dedicated automation account instead of a personal account so that:

- ownership remains with the organization
- flows continue working if staff change
- permissions can be tightly managed

## Error Handling Design

Use three status values in the source system:

- Pending
- Success
- Error

Capture these common failures:

- Missing required fields
- Duplicate source record
- Authentication failure
- Jira permission issue
- Assets schema mismatch
- API throttling or timeout

Record the reason in ErrorMessage.

## Idempotency / Duplicate Prevention

This is very important.

Use a stable unique key such as:

- SourceRecordID, or
- SerialNumber

Rules:

- Never create a new item if a matching SourceRecordID already exists.
- Always search before create when possible.
- Store returned JiraIssueKey and AssetObjectID back into the source immediately.

## Example Processing Scenario

A spreadsheet contains 200 rows for couches.

Example rows:

- Blue Couch 01
- Blue Couch 02
- Red Couch 01
- Red Couch 02

Power Automate reads each row and:

- creates a Couch object in Assets,
- creates a Jira work item called something like Couch - Blue - Building 1 Lobby,
- writes IDs back to the spreadsheet,
- marks the row as Success.

No manual issue creation is required.

## Why Not Use CSV Import Alone

CSV import is useful for one-time bulk loads, but it has limitations:

- less dynamic for ongoing synchronization
- harder to manage updates automatically
- weaker error feedback for each source row
- less flexible for linking Assets and Jira in a repeatable way

CSV import is still useful for the very first historical load. Power Automate is better for recurring operations.

## Recommended Implementation Phases

### Phase 1: Prototype

- Create a small source table with 5 to 10 records.
- Create one Power Automate flow.
- Create Jira issues only.
- Test field mapping and error handling.

### Phase 2: Add Assets

- Build the object schema.
- Add create/update object logic.
- Store AssetObjectID in the source.

### Phase 3: Add Linking

- Link Jira issues to Assets objects.
- Add Jira custom field for AssetObjectID.
- Add Assets attribute for JiraIssueKey.

### Phase 4: Add Reporting

- Build Jira dashboards.
- Surface Jira data in Confluence.
- Build Confluence pages for inventory or status views.

### Phase 5: Production Hardening

- Use service account.
- Add logging.
- Add retry logic.
- Add permissions review.
- Add operational runbook.

## Governance Recommendations

- Keep the source schema simple.
- Avoid creating too many Jira custom fields.
- Put descriptive/static data in Assets when possible.
- Put workflow data in Jira.
- Use naming standards for records.
- Document every mapping between source columns and Atlassian fields.
- Test in a non-production project first.

## Risks

- Power Automate premium licensing may be required for HTTP actions.
- Atlassian API changes may require flow updates.
- Too many custom fields in Jira can create long-term admin burden.
- Poor deduplication rules can create duplicate issues or objects.
- Excel can be fragile if multiple users edit it at once.

## Mitigations

- Use SharePoint List or Microsoft List instead of Excel for higher reliability if possible.
- Use SourceRecordID as a hard unique key.
- Start with limited issue types and fields.
- Test with small batches.
- Add clear SyncStatus and ErrorMessage columns.
- Separate test and production flows.

## Recommended Source Platform

For long-term operational use:

- Best: SharePoint List or Microsoft List
- Good: Excel Online table
- One-time migration only: CSV file

Reason: Lists are more stable than spreadsheets for automation and multi-user updates.

## Future Enhancements

- Add approval steps before record creation.
- Add update-only mode for existing records.
- Add delete/retire logic.
- Add change tracking and audit history.
- Add Power BI reporting.
- Add Confluence executive dashboard pages.
- Add automation rules in Jira after item creation.

## Sample Field Mapping Table

| Source Column | Jira Field / Asset Attribute | Notes |
|---|---|---|
| SourceRecordID | Source Record ID | Unique key |
| ItemName | Summary / Name | Main display value |
| Category | Category | Can drive object type |
| Color | Color | Example attribute |
| Model | Model | Optional |
| SerialNumber | Serial Number | Good dedupe field |
| Location | Location | Can also link to location object |
| Status | Status | Map carefully |
| Owner | Owner | Person or team |
| Notes | Description / Notes | Free text |
| JiraIssueKey | Jira issue key | Write-back value |
| AssetObjectID | Asset object ID | Write-back value |

## Operational Recommendation

For your environment, the simplest compliant path is:

- maintain the source in SharePoint List or Excel Online,
- use Power Automate cloud flows,
- call Jira and Assets APIs with HTTP actions,
- store returned IDs back in the source,
- display the result in Jira and Confluence.

This avoids local installs and keeps the solution within the tools your Microsoft-centered environment is most likely to allow.
