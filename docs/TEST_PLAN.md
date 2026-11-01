# Test Plan

## Test description

This plan validates the sanitized Excel VBA and SAP GUI Scripting workflow that exports WIP production-order data from SAP ZCO05, filters and deduplicates the data in Excel, and updates a confirmed SAP order status through COR2.

The plan covers offline Excel checks, controlled SAP integration checks, negative paths, evidence capture, and release gating. It does not authorize production execution and it does not change the VBA business logic.

## Current status

- Automated application tests: not included.
- Excel/VBA execution during the repository audit: not run.
- SAP integration execution: not run.
- Test environment: must be an approved non-production SAP system.
- Target status: VAR is present in the source; confirm the VAR/WAR business mapping before testing.

## Validation workflow

1. Obtain approval for the test system, client, plant, target status, export retention policy, and sanitized data.
2. Prepare an approved macro-enabled workbook with the expected Format and Temp worksheets.
3. Create a synthetic or sanitized export fixture containing representative headers, qualifying rows, excluded rows, blank values, duplicates, and malformed values.
4. Run preflight checks for Excel settings, worksheet availability, C:\TEMP access, SAP GUI Scripting, and the intended SAP session.
5. Run Stage 1 and verify that a new export is created after the run start rather than reusing a stale file.
6. Run Stage 2 and compare Format and Temp against the expected filter and deduplication results.
7. In the non-production SAP system, run Stage 3 with a small known-safe order set.
8. Verify the selected status, the SAP save result, the worksheet result, popup handling, and the completion counts.
9. Repeat negative-path tests and confirm that Excel settings, source files, and operator-visible errors are handled as expected.
10. Capture evidence, record pass/fail results, and approve release only when all critical tests pass.

## Test matrix

| ID | Test | Expected result | Status |
| --- | --- | --- | --- |
| T01 | SAP GUI unavailable | The workflow reports an actionable error and does not continue with SAP updates. | Not run |
| T02 | Wrong SAP system or client | Preflight or operator review prevents execution against an unintended system. | Not run |
| T03 | C:\TEMP missing or unwritable | Export stage stops with a clear error and Excel settings are restored. | Not run |
| T04 | Stale EXPORT_*.xlsx exists | The stale file is not accepted as the current run result. | Not run |
| T05 | New export appears within the wait window | The new export is selected and imported. | Not run |
| T06 | Export never appears | The workflow times out with an exception result and does not process old data. | Not run |
| T07 | Required headers present | Stage 2 copies the first worksheet into Format and applies the intended profiles. | Not run |
| T08 | Required header missing or moved | The issue is detected or documented as a controlled failure; no silent mis-filtering occurs. | Not run |
| T09 | Profile 1 qualifying row | The row contributes its order number to Temp. | Not run |
| T10 | Profile 1 excluded row | The row is not added to Temp. | Not run |
| T11 | Profile 2 qualifying row | The row contributes its order number to Temp. | Not run |
| T12 | Duplicate order numbers | Temp contains one queue entry per unique order number. | Not run |
| T13 | Blank or malformed order number | The row is excluded or reported without crashing the batch. | Not run |
| T14 | VAR is in the first ten status rows | The status is selected and the order proceeds to save verification. | Not run |
| T15 | VAR is outside the first ten rows | The limitation is visible and the order is not incorrectly marked successful. | Not run |
| T16 | Expected information popup | The documented popup path produces the expected exception result. | Not run |
| T17 | Save fails or save confirmation is absent | The order is not reported as Success. | Not run |
| T18 | Batch interruption or unexpected popup | Cleanup runs, the operator sees the failing order, and remaining data is not silently misreported. | Not run |

## Evidence to capture

- Test system, client, plant, date, and operator.
- Sanitized input export or synthetic fixture identifier.
- Expected versus actual row and order counts.
- Format and Temp worksheets before and after processing.
- SAP status and save confirmation for each test order.
- Popup text and exception result where applicable.
- Final completion message and any timestamped log available in the test copy.
- GitHub Actions run for documentation validation.

Do not capture or commit credentials, real order numbers, supplier information, internal hostnames, or screenshots containing business data.

## Release gate

The project is not production-ready until the VAR/WAR mapping is confirmed, T01 through T18 are reviewed, all critical and high-risk tests pass in an approved non-production system, and the export-retention and audit-log policies are documented.

The repository's GitHub Actions workflow validates documentation structure and prevents common workbook/export artifacts from being committed. It does not test Excel, VBA, or SAP behavior.
