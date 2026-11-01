# VBA Application Review

## Scope and method

The supplied macro-enabled workbook was reviewed from a separate copy. The review covered the XLSM package structure, workbook XML, worksheets, drawing assignment, embedded VBA project streams, VBA source code, and the visible Excel workbook state. The source workbook was not modified, and the automation was not executed against SAP because execution would change SAP records and delete the generated export file.

## Executive conclusion

The workbook is an Excel-based SAP GUI automation for a WIP production-order process. Its visible entry point is a `ZCO05` shape on `Format`, but the assigned macro `RunAll` runs an end-to-end pipeline:

```text
SAP ZCO05 export
        ↓
Excel Format import and filtering
        ↓
Temp order queue and deduplication
        ↓
SAP COR2 status selection
        ↓
Success / exception results and cleanup
```

The code is process-specific rather than a general-purpose framework. It depends on SAP GUI Scripting, a logged-in SAP session, fixed SAP control IDs, plant `1900`, user status `WIP`, a local `C:\TEMP` folder, and an `EXPORT_*.xlsx` file naming pattern.

## Workbook findings

| Area | Finding |
| --- | --- |
| File type | Excel macro-enabled workbook (`.xlsm`) |
| Worksheets | `Format` and `Temp` |
| `Format` dimensions | `A1:T515`, including a header row and 514 data rows in the supplied file |
| `Temp` dimensions | `A1:B1` in the supplied saved state |
| `Format` data | SAP-style production-order fields including plant, scheduler, order number, material number, order type, delivery, confirmation, user status, and release date |
| Visible control | A rectangle labeled `ZCO05` assigned to `[0]!RunAll` |
| Formulas | None observed in the worksheet XML |
| Charts | None observed |
| User forms | None observed in the VBA project streams |
| Database/web connections | None observed |
| Shared strings | 760 unique strings, 9,788 total references |
| Defined names | Hidden `_FilterDatabase` for `Format!$A$1:$T$515` |

The supplied `Format` data includes real-looking operational identifiers and should be treated as confidential until it has been replaced with dummy values.

## VBA project findings

The VBA project contains three code modules and several internal compiled-cache streams:

- `ThisWorkbook`: document module with attributes only; no custom procedure code observed.
- `Sheet1`: worksheet module with attributes only; no custom procedure code observed.
- `Module1`: the main standard module containing 374 lines of VBA source.

The main procedures are:

| Procedure | Implemented responsibility |
| --- | --- |
| `RunAll` | Disables screen updating and alerts, then calls the three stages |
| `Stage1_RunZCO05` | Attaches to SAP, executes `ZCO05`, exports to `C:\TEMP`, waits for the newest export |
| `Stage2_CleanExcel` | Imports the export, clears/builds worksheets, applies filters, copies order numbers, removes duplicates |
| `Stage3_COR2_VARUpdate` | Reads the queue, opens each order in `COR2`, selects `VAR`, writes result status, summarizes and deletes export |
| `CopyVisibleColC` | Copies visible nonblank order numbers from `Format` column C |
| `GetLatestSAPExport` | Finds the newest `EXPORT_*.xlsx` file using `Dir` and `FileDateTime` |
| `KillMessageFilter` | Clears an object variable; no assignment to that variable was observed in the reviewed source |

## Detailed logic

### Stage 1: SAP `ZCO05`

The code uses late-bound SAP GUI objects:

```vb
Set SapGuiAuto = GetObject("SAPGUI")
Set Application = SapGuiAuto.GetScriptingEngine
Set Connection = Application.Children(0)
Set Session = Connection.Children(0)
```

It then sends `/NZCO05`, sets:

- Plant: `1900`
- User status field: `WIP`

The recorded SAP menu sequence exports to `C:\TEMP`. The macro checks for a file matching `EXPORT_*.xlsx` for up to 30 seconds. If no matching file is found, it displays an error and terminates the VBA execution with `End`.

### Stage 2: Excel filtering

The source workbook is opened read-only. Its first worksheet is copied to `Format`. `Temp` is reset with `Ord. No.` and `Results` headers.

The first filter profile requires:

- `Prod. schdlr` in `032`, `033`, `040`, `061`, `065`
- `Ord. typ` in `ZP01`, `ZP03`, `ZP04`, `ZP06`, `ZP07`, `ZP14`
- Nonblank `DLV`
- Nonblank `Usr Decsn.`
- Nonblank `CNF`

The second profile requires:

- `Prod. schdlr` in `060`, `066`, `070`, `080`, `085`, `100`
- Nonblank `DLV`
- Nonblank `CNF`

Each profile copies visible values from column C (`Ord. No.`) into `Temp`, then the combined queue is deduplicated on column A.

### Stage 3: SAP `COR2`

For each queued order, the macro:

1. Sends `/nCOR2`.
2. Selects the completion-status option.
3. Enters the order number.
4. Handles an information popup as `Error - Info Popup`.
5. Opens the status list.
6. Checks status rows 0 through 9 for exact text `VAR`.
7. Selects and saves the status when found.
8. Writes `Success` or `Error - VAR not found` to `Temp`.

After the loop, the code counts exact `Success` values and treats every other result as an exception. It deletes the newest export file and shows a completion message.

## Dependencies and environment assumptions

- Microsoft Excel desktop on Windows.
- VBA 7-capable Excel. The code includes a conditional `Sleep` declaration for VBA7 and older VBA versions.
- SAP GUI for Windows with the SAP GUI Scripting Engine available.
- A logged-in SAP session at connection index 0 and session index 0.
- Authorization for SAP transactions `ZCO05` and `COR2`.
- A writable `C:\TEMP` folder.
- SAP screens and field IDs that match the recorded scripting paths.

## Publication and confidentiality review

### Data and metadata that should be removed

The supplied file contains or exposes:

- Production-order numbers.
- Material numbers.
- Plant and scheduler values.
- SAP status and date fields.
- Author and last-modified-by metadata.
- An internal SharePoint-style absolute path in workbook metadata.
- Templafy tenant/template/profile identifiers in custom properties.
- Company-specific SAP transaction names, field IDs, filter values, and status logic.

No plain-text SAP password or API key was found in the readable VBA source. The VBA project contains `CMG`, `DPB`, and `GC` protection metadata in the `PROJECT` stream. This is project-protection metadata, not an application login credential and not a substitute for access control.

### Recommended public-release process

1. Make a separate sanitized workbook copy.
2. Replace all production-order and material data with synthetic values.
3. Remove document properties and custom XML metadata.
4. Remove internal paths and company-specific identifiers from screenshots and documentation.
5. Export the VBA modules to text files for source control.
6. Review the sanitized VBA for remaining internal SAP details.
7. Test only in a non-production SAP environment.
8. Obtain approval before publishing the workbook or source code.

## Risks and limitations observed

- `Application.ScreenUpdating` and `Application.DisplayAlerts` are restored only on the normal path. `End`-based error exits can bypass cleanup.
- SAP connection and session selection is positional (`Children(0)`), which is fragile when multiple connections or sessions are open.
- The export path, plant, SAP user status, filter values, and SAP control IDs are hard-coded.
- The newest matching file is selected by timestamp. A stale file could be processed if a new export is delayed or fails.
- The code deletes the newest export after processing, which reduces the audit trail.
- Only status rows 0 through 9 are searched.
- Errors are mainly reported through message boxes and status text; there is no durable log.
- End-to-end behavior was not executed during this review because it would make SAP changes and delete the export file.
- The workbook name says `WIP-TO-WAR`, while the code explicitly selects `VAR`. The process owner should confirm whether that is intentional.

## Recommended future improvements

These are not implemented features:

- Add a configuration sheet for paths, plant, SAP user status, order types, scheduler values, and target status.
- Add a preflight check for SAP scripting, transaction access, session state, and folder permissions.
- Replace `End` with structured error handling and a cleanup block.
- Add a dry-run mode and confirmation before writing SAP changes.
- Keep an archive of source exports and write a timestamped audit log.
- Validate file freshness using a run start timestamp or unique export name.
- Search all available SAP status rows.
- Add a controlled test harness using dummy data and a non-production SAP system.
- Document the business meaning of `VAR` and its relationship to the project name `WAR`.
