# WIP-to-WAR Process Automation

Excel VBA and SAP GUI Scripting automation for extracting WIP production-order data, identifying eligible orders, and updating their SAP status through `COR2`.

> Important: The workbook is named **WIP-to-WAR**, but the VBA code searches for and selects the SAP status value `VAR`. Confirm the business meaning and the `VAR`/`WAR` mapping before publishing or using this project.

## Project overview

This application automates a repetitive production-order workflow that would otherwise require a user to:

1. Run SAP transaction `ZCO05` for plant `1900` and user status `WIP`.
2. Export the result to Excel.
3. Filter the export using predefined scheduler, order-type, delivery, user-decision, and confirmation criteria.
4. Copy qualifying order numbers into a work queue.
5. Open each order in SAP transaction `COR2`.
6. Find and select the `VAR` status.
7. Save the change and record the result.

The workbook uses Excel as the staging and result-tracking layer. SAP GUI Scripting performs the SAP navigation and updates.

## Implemented features

- **One-click execution:** The visible `ZCO05` worksheet shape is assigned to the public VBA procedure `RunAll`.
- **SAP export automation:** Connects to an open SAP GUI session and runs transaction `ZCO05` with plant `1900` and user status `WIP`.
- **Export-file detection:** Looks in `C:\TEMP` for the newest file matching `EXPORT_*.xlsx`, waiting up to 30 seconds for the file to appear.
- **Workbook refresh:** Clears or creates the `Format` and `Temp` worksheets before each run.
- **Read-only source import:** Opens the newest SAP export as a read-only workbook and copies the first worksheet into `Format`.
- **Two filter profiles:** Applies two predefined Excel AutoFilter sequences to identify qualifying orders.
- **Order queue creation:** Copies visible order numbers from column C into `Temp` and removes duplicate order numbers.
- **SAP status update:** Opens each queued order in `COR2`, opens the status list, searches the first ten visible status rows for `VAR`, and selects it when found.
- **Result tracking:** Writes `Success`, `Error - Info Popup`, or `Error - VAR not found` to the `Results` column.
- **Completion summary:** Displays staged, successful, and exception counts in a message box.
- **Temporary-file cleanup:** Deletes the newest `EXPORT_*.xlsx` file from `C:\TEMP` after the processing loop.

## Application workflow

### Stage 1: Export WIP data from SAP

`Stage1_RunZCO05`:

1. Attaches to the first available SAP GUI connection and session through `GetObject("SAPGUI")`.
2. Opens transaction `/NZCO05`.
3. Sets plant to `1900`.
4. Sets the SAP user-status filter to `WIP`.
5. Executes the report.
6. Uses the SAP menu export flow to save an Excel file in `C:\TEMP`.
7. Waits for the newest `EXPORT_*.xlsx` file to become available.

### Stage 2: Clean and stage orders in Excel

`Stage2_CleanExcel`:

1. Locates the newest SAP export.
2. Clears or creates `Format` and `Temp`.
3. Copies the first sheet of the SAP export into `Format`.
4. Creates `Temp!A1:B1` with `Ord. No.` and `Results` headers.
5. Runs filter profile 1:
   - Scheduler values: `032`, `033`, `040`, `061`, `065`
   - Order types: `ZP01`, `ZP03`, `ZP04`, `ZP06`, `ZP07`, `ZP14`
   - Delivery (`DLV`) is not blank
   - User decision (`Usr Decsn.`) is not blank
   - Confirmation (`CNF`) is not blank
6. Copies visible order numbers from column C to `Temp`.
7. Runs filter profile 2:
   - Scheduler values: `060`, `066`, `070`, `080`, `085`, `100`
   - Delivery (`DLV`) is not blank
   - Confirmation (`CNF`) is not blank
8. Appends the second set of order numbers to `Temp`.
9. Removes duplicate order numbers and autofits the two output columns.

### Stage 3: Update SAP status

`Stage3_COR2_VARUpdate`:

1. Reads each order number from `Temp`.
2. Opens `/nCOR2` in SAP.
3. Selects the completion-status option and enters the order number.
4. Handles an information popup by recording `Error - Info Popup`.
5. Opens the order-status list.
6. Checks status rows 0 through 9 for the exact text `VAR`.
7. Selects `VAR` and saves the order when found.
8. Records `Success` when the status is found, or `Error - VAR not found` otherwise.
9. Shows a completion summary.
10. Deletes the newest matching export file from `C:\TEMP`.

## Workbook structure

```text
WIP-TO-WAR Process Automation.xlsm
├── Format
│   ├── Imported SAP production-order data
│   ├── 20 columns, A:T
│   └── Worksheet shape labeled ZCO05, assigned to RunAll
├── Temp
│   ├── Ord. No.
│   └── Results
└── VBA project
    ├── ThisWorkbook       Workbook document module; no custom procedure logic observed
    ├── Sheet1             Worksheet module; no custom procedure logic observed
    └── Module1            Main automation code and helper procedures
```

### `Format` columns

The imported data contains the following fields:

`Plant`, `Prod. schdlr`, `Ord. No.`, `Material no.`, `Ord. typ`, `GIE`, `DLV`, `Usr Decsn.`, `TECO`, `Usr stat.`, `Err stat`, `Fiscal yr.`, `Fiscal prd.`, `DLVE`, `DLVR`, `Del. flg`, `Del. indic.`, `CNF`, `PCNF`, and `Actual rel. date`.

### `Temp` columns

| Column | Purpose |
| --- | --- |
| `Ord. No.` | Deduplicated list of SAP order numbers selected by the filter profiles |
| `Results` | Per-order outcome written by the `COR2` processing stage |

## Technical details

### Languages and platforms

- Visual Basic for Applications (VBA)
- Microsoft Excel desktop, macro-enabled workbook format (`.xlsm`)
- Windows-only execution because the code declares the Windows `kernel32` `Sleep` API
- SAP GUI for Windows with SAP GUI Scripting enabled

### VBA procedures

| Procedure | Role |
| --- | --- |
| `RunAll` | Master entry point; runs all three stages |
| `Stage1_RunZCO05` | Runs SAP export and waits for the generated file |
| `Stage2_CleanExcel` | Imports, filters, stages, and deduplicates order numbers |
| `Stage3_COR2_VARUpdate` | Processes the staged queue in SAP and records results |
| `CopyVisibleColC` | Copies visible nonblank order numbers from `Format` to `Temp` |
| `GetLatestSAPExport` | Returns the newest `EXPORT_*.xlsx` file in a folder |
| `KillMessageFilter` | Clears an object variable used for message-filter tracking |

### Excel features used

- `Workbooks.Open` with `ReadOnly:=True`
- Worksheet creation, clearing, copying, and renaming
- `AutoFilter` with multi-value and nonblank criteria
- `SpecialCells(xlCellTypeVisible)`
- `RemoveDuplicates`
- Column autofit
- Worksheet shape-to-macro assignment
- `MsgBox` status and completion messages
- `Application.ScreenUpdating` and `Application.DisplayAlerts`

### External dependencies

- SAP GUI must already be open and logged in.
- SAP GUI Scripting must be enabled and permitted by the SAP environment.
- The user must have access to transactions `ZCO05` and `COR2`.
- `C:\TEMP` must exist and be writable.
- The SAP export must use the `EXPORT_*.xlsx` naming pattern.

No database, Outlook integration, web service, Power Query connection, user form, chart, pivot table, or formula-driven calculation was observed in the supplied workbook.

## Installation and setup

1. Download the repository as a ZIP or clone it with Git.
2. Install Microsoft Excel desktop on Windows.
3. Install SAP GUI for Windows and confirm that SAP GUI Scripting is enabled according to your organization's policy.
4. Confirm that your SAP user can run `ZCO05`, export the result, open `COR2`, change the required status, and save the order.
5. Create `C:\TEMP` if it does not already exist.
6. Place the sanitized `.xlsm` workbook in a local folder.
7. Open SAP GUI and log in before starting the workbook.
8. Open the workbook and allow macros only if the file is trusted and approved for use.
9. Click the `ZCO05` shape on the `Format` worksheet.
10. Review the `Temp` worksheet and its `Results` column after completion.

The current code contains hard-coded SAP field identifiers, filter values, plant/user values, and the `C:\TEMP` folder. These values must match the target SAP environment.

## Usage example

```text
1. Start SAP GUI and log in.
2. Open the Excel macro workbook.
3. Click the ZCO05 button on Format.
4. Wait for the staged queue to be processed.
5. Review Temp!A:B for order numbers and results.
6. Review the completion message for staged, successful, and exception counts.
```

The macro deletes the newest matching export file from `C:\TEMP` at the end of the pipeline. Preserve a backup or change that policy in an approved development copy if the raw export is needed for audit purposes.

## Business benefits

The implemented workflow is designed to:

- Reduce repeated manual navigation between SAP and Excel.
- Apply consistent selection rules to the WIP production-order population.
- Reduce copying and transcription errors when creating the SAP work queue.
- Remove duplicate order numbers before SAP updates.
- Provide a visible per-order result and a completion summary.
- Reduce the time spent on repetitive order-status maintenance.
- Make exception cases visible for follow-up instead of silently ignoring them.

No measured time saving, error-rate reduction, or throughput baseline was included in the workbook. Any quantified business claim should be supported by operational measurements.

## Recommended repository structure

```text
.
├── README.md
├── workbook/
│   ├── WIP-TO-WAR-Process-Automation-sanitized.xlsm
│   └── sample-input/
│       └── EXPORT_sample.xlsx
├── src/
│   ├── Module1.bas
│   ├── ThisWorkbook.cls
│   └── Sheet1.cls
├── docs/
│   ├── workflow.md
│   ├── data-dictionary.md
│   └── screenshots/
├── .gitignore
└── LICENSE
```

### Files suitable for GitHub

- `README.md`
- Sanitized exported VBA modules such as `src/Module1.bas`
- A sanitized workbook only if the macro project and sample data are approved for public release
- Dummy input data that preserves the column structure without using real order, material, supplier, or company data
- Redacted screenshots and workflow documentation

### Files to keep private or exclude

- The original workbook supplied for this review
- Real SAP exports and production-order data
- Real material numbers, supplier details, plant data, or internal status data
- Screenshots containing SAP user names, system names, host names, order numbers, or company records
- Internal SharePoint paths, author information, Templafy identifiers, or other document metadata
- Any workbook containing an internal macro project unless its publication is approved

Suggested `.gitignore` entries:

```gitignore
*.xlsx
*.xlsm
!workbook/WIP-TO-WAR-Process-Automation-sanitized.xlsm
!workbook/sample-input/EXPORT_sample.xlsx
workbook/private/
workbook/live-data/
*.export.xlsx
```

Adjust the allow-list to match the sanitized file name used by the repository.

## Screenshots to include

Use synthetic or fully redacted data in every public screenshot.

1. **Workbook landing view:** `Format` with the `ZCO05` button visible and dummy production-order rows.
2. **Filter/staging view:** `Temp` showing deduplicated order numbers before SAP processing.
3. **Results view:** `Temp` showing `Success`, `Error - Info Popup`, and `Error - VAR not found` examples.
4. **SAP export step:** A redacted `ZCO05` selection screen, only if screenshots are approved for public use.
5. **SAP update step:** A redacted `COR2` status list with `VAR` selected, only if approved.
6. **Completion message:** A sample summary message with dummy counts.

## Security and publication notes

Static review found no plain-text SAP password, API key, or other login secret in the readable VBA source. The VBA project contains standard project-protection metadata; that is not a substitute for secret management or access control.

Before publishing:

- Remove real order numbers, material numbers, dates, statuses, and plant records.
- Remove or replace document properties and custom metadata.
- Remove internal SharePoint or OneDrive paths.
- Remove author names and tenant/template identifiers unless intentionally public.
- Replace the workbook with dummy data and verify that the macro does not point to a real SAP environment.
- Review every hard-coded SAP transaction, field identifier, status value, and local path with the process owner.

## Future improvements

These items are suggestions, not features present in the supplied workbook:

- Move plant, SAP user, filter lists, status values, and export path into a configuration area.
- Validate SAP connection, session, transaction access, and `C:\TEMP` before starting.
- Replace `End`-based error exits with structured cleanup that restores Excel application settings.
- Add a dry-run or confirmation step before writing status changes to SAP.
- Preserve or archive exports instead of deleting them automatically.
- Use a run-specific export filename or a stricter freshness check to avoid processing a stale file.
- Search the complete SAP status list rather than only rows 0 through 9.
- Add an audit log with timestamp, order number, selected status, result, and error text.
- Add a dummy-data test harness and a documented SAP test-system procedure.
- Export VBA modules into source-controlled text files and keep the `.xlsm` as a packaged release artifact.
- Confirm and document whether SAP status `VAR` is the intended implementation of the business term `WAR`.

## Portfolio description

Developed an Excel VBA and SAP GUI Scripting automation that extracts WIP production-order data from SAP transaction `ZCO05`, filters and deduplicates eligible orders, and updates their status through `COR2`. The solution automates the staging, validation, result tracking, and cleanup steps of a procurement and production-operations workflow.

## License and ownership

Add a license only after confirming that the VBA code, SAP transaction logic, workbook design, sample data, and company-specific process knowledge may be shared publicly. The original workbook contains business-specific logic and operational data that may require employer approval before publication.
