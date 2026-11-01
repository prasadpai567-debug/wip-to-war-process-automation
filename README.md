# SAP WIP Production-Order Status Automation

Excel VBA and SAP GUI Scripting automation for filtering WIP production orders and updating an SAP COR2 user status.

> Portfolio note: This repository contains sanitized VBA source and documentation. The original macro-enabled workbook and production data are not included. The current implementation selects SAP status `VAR`; confirm its relationship to the project name “WIP-to-WAR” before operational use.

## Business problem

The process requires repeated movement between SAP and Excel:

1. Run SAP transaction `ZCO05`.
2. Export WIP production-order data.
3. Filter the export using operational selection rules.
4. Create a queue of eligible order numbers.
5. Open each order in SAP transaction `COR2`.
6. Select the required user status.
7. Record the result for follow-up.

Manual processing is repetitive and increases the risk of inconsistent filtering, duplicate order handling, and missed exceptions.

## What the automation does

The VBA workflow:

1. Connects to an existing SAP GUI session.
2. Runs `ZCO05` for plant `1900` and user status `WIP`.
3. Exports the report to a local `EXPORT_*.xlsx` file in `C:\TEMP`.
4. Imports the first worksheet into Excel worksheet `Format`.
5. Applies two predefined filter profiles.
6. Creates a deduplicated order queue in worksheet `Temp`.
7. Opens each queued order in `COR2`.
8. Searches the first ten visible status rows for `VAR`.
9. Selects the status and attempts to save the order.
10. Records a per-order result in `Temp!B:B`.

The business meaning of `VAR` and its relationship to “WAR” must be confirmed before the automation is used to update SAP records.

## Verified capabilities

- Excel VBA automation with `Option Explicit`
- SAP GUI Scripting through late-bound objects
- SAP `ZCO05` export navigation
- Excel worksheet refresh and staging
- AutoFilter processing with two selection profiles
- Duplicate order-number removal
- SAP `COR2` order navigation
- Status lookup and selection for `VAR`
- Per-order result tracking
- Completion summary message
- Temporary export-file handling

## Application workflow

### Stage 1: Export WIP data from SAP

Procedure: `Stage1_RunZCO05`

- Attaches to the first available SAP connection and session.
- Opens `/NZCO05`.
- Sets plant `1900`.
- Sets the SAP user-status filter to `WIP`.
- Executes the report.
- Uses the SAP menu export flow.
- Saves to `C:\TEMP`.
- Waits up to 30 seconds for a matching `EXPORT_*.xlsx` file.

### Stage 2: Filter and stage eligible orders

Procedure: `Stage2_CleanExcel`

The source workbook is opened read-only and copied into `Format`. The queue is rebuilt in `Temp`.

Filter profile 1:

- Scheduler: `032`, `033`, `040`, `061`, `065`
- Order type: `ZP01`, `ZP03`, `ZP04`, `ZP06`, `ZP07`, `ZP14`
- Delivery is not blank
- User decision is not blank
- Confirmation is not blank

Filter profile 2:

- Scheduler: `060`, `066`, `070`, `080`, `085`, `100`
- Delivery is not blank
- Confirmation is not blank

The resulting order numbers are appended to `Temp` and deduplicated.

### Stage 3: Update SAP status

Procedure: `Stage3_COR2_VARUpdate`

For each staged order, the procedure:

- Opens `/nCOR2`.
- Selects the completion-status option.
- Enters the order number.
- Handles one recognized information-popup case.
- Opens the order-status list.
- Searches status rows 0 through 9 for the exact text `VAR`.
- Selects `VAR` when found.
- Writes `Success`, `Error - Info Popup`, or `Error - VAR not found`.
- Attempts to save when `VAR` is found.
- Displays staged, successful, and exception counts.

## Repository structure

```text
.
├── README.md
├── .gitignore
├── docs/
│   ├── PORTFOLIO_SUMMARY.md
│   ├── REPOSITORY_AUDIT.md
│   ├── TEST_PLAN.md
│   └── VBA_APPLICATION_REVIEW.md
└── src/
    ├── Module1.bas
    ├── Sheet1.cls
    └── ThisWorkbook.cls
```

The repository currently stores exported VBA source and documentation. It does not include the original `.xlsm` workbook, SAP exports, or production data.

## Requirements

- Windows
- Microsoft Excel desktop
- SAP GUI for Windows
- SAP GUI Scripting enabled and permitted
- Access to SAP transactions `ZCO05` and `COR2`
- An approved macro-enabled workbook containing the required worksheets and button
- A writable export directory

The source files alone are not a complete runnable application. The workbook must be packaged separately only after it has been sanitized and approved for release.

## Setup and usage

1. Clone or download this repository.
2. Obtain an approved sanitized workbook package separately.
3. Install Microsoft Excel desktop and SAP GUI for Windows.
4. Confirm that SAP GUI Scripting is enabled according to organizational policy.
5. Confirm access to `ZCO05` and `COR2`.
6. Confirm that `C:\TEMP` exists and is writable.
7. Open SAP GUI and log in to the approved test system.
8. Open the approved workbook and enable macros only when the file is trusted.
9. Start the workbook entry point assigned to the `ZCO05` control.
10. Review the `Temp` worksheet and completion message.

The current source contains hard-coded SAP control IDs, filter values, plant and status values, and the `C:\TEMP` path. These values must match the intended SAP environment.

## Safety and data handling

This automation changes SAP records. It should be tested in an approved non-production environment until the following are confirmed:

- Correct SAP system, client, and session
- Correct plant and filter values
- Correct target status
- Correct export file
- Successful save confirmation
- Approved handling of exported data

Do not commit real order numbers, material numbers, supplier information, company records, SAP exports, screenshots, credentials, or internal paths.

The current implementation deletes the newest matching export file from `C:\TEMP` after processing. Preserve the export or change that policy in an approved development copy if the source data is needed for audit or reconciliation.

## Error handling and logging

The implementation displays message boxes and writes limited per-order results to `Temp`. It does not currently provide a durable timestamped audit log.

Known technical limitations include:

- Excel application settings are not protected by a single cleanup block.
- Some failure paths use VBA `End`.
- SAP connection and session selection is positional.
- Export freshness is inferred from the newest filename match.
- Only the first ten visible SAP status rows are searched.
- A successful status selection is recorded before save confirmation is verified.
- Results are not explicitly saved to the workbook by the macro.
- The source does not include an automated test suite.

See [REPOSITORY_AUDIT.md](docs/REPOSITORY_AUDIT.md) for the detailed findings and [TEST_PLAN.md](docs/TEST_PLAN.md) for the recommended validation approach.

## Business benefits

The workflow is designed to:

- Reduce repeated navigation between SAP and Excel.
- Apply consistent selection rules.
- Reduce manual copying and duplicate order handling.
- Make per-order exceptions visible.
- Standardize a repetitive production-order status process.

No measured time saving, error-rate reduction, or throughput result is claimed because the repository contains no operational baseline.

## Future improvements

- Confirm and document the `VAR`/“WAR” business mapping.
- Add a configuration sheet for paths, plant, filters, and target status.
- Add preflight checks for SAP, folder, file, and worksheet readiness.
- Replace `End` with structured error handling and cleanup.
- Validate export freshness against a run-start timestamp.
- Preserve or archive source exports.
- Add a dry-run and confirmation mode.
- Confirm SAP save success before writing `Success`.
- Search the complete SAP status list.
- Add a durable audit log.
- Add synthetic test data and a non-production SAP test procedure.
- Package an approved sanitized workbook as a release artifact.

## Portfolio summary

Developed an Excel VBA and SAP GUI Scripting automation for a WIP production-order process. The solution extracts data from SAP `ZCO05`, applies business filters in Excel, removes duplicate orders, stages a processing queue, and navigates SAP `COR2` to update the required order status.

The project demonstrates Excel VBA, SAP GUI Scripting, process automation, data filtering, exception tracking, operational risk awareness, and technical documentation.

## License and ownership

No license is currently included. Add one only after confirming that the VBA code, SAP transaction logic, workbook design, and company-specific process knowledge may be shared publicly.
