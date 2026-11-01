# Portfolio Summary

## Project

SAP WIP Production-Order Status Automation is an Excel VBA and SAP GUI Scripting workflow for a repetitive production-order process. It extracts WIP data from SAP ZCO05, filters the export in Excel, creates a deduplicated queue, and navigates SAP COR2 to update a confirmed order status.

## Business problem

The manual process requires repeated SAP navigation, spreadsheet filtering, order-number copying, status selection, and result tracking. Repetition creates opportunities for inconsistent filters, duplicate orders, missed exceptions, and weak audit evidence.

## Technical contribution

- Designed a staged Excel workflow with export, filtering, queue, and SAP-update phases.
- Automated SAP GUI navigation through late-bound scripting objects.
- Implemented two operational filter profiles and duplicate removal.
- Added per-order result tracking and completion counts.
- Documented hard-coded dependencies, safety risks, data-handling rules, and test gates.
- Added repository-level documentation validation without changing SAP business behavior.

## Skills demonstrated

Excel VBA, SAP GUI Scripting, Excel AutoFilter, data staging, duplicate handling, process automation, exception tracking, operational risk assessment, technical writing, GitHub documentation, and safe release planning.

## Business value

The workflow is intended to reduce repetitive navigation and manual copying, apply consistent selection rules, expose exceptions, and standardize a production-order status process. No time-saving or accuracy metric is claimed because the repository contains no operational baseline.

## Portfolio positioning

This project is suitable for demonstrating automation work to SAP Procurement Specialist, Procurement Automation Specialist, Procurement Operations Analyst, and enterprise process-improvement audiences. It shows how an operational process can be translated into a controlled technical workflow while documenting the conditions required for safe use.

## Important limitation

The repository contains exported VBA source and documentation, not the original macro-enabled workbook or production data. The source currently targets SAP status VAR; the VAR/WAR business mapping and non-production validation must be confirmed before operational use.
