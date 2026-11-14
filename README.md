# EBSCO PubOps VBA Macros

A small collection of Excel VBA tools created to reduce repetitive work in an academic journal package quality-control workflow. The macros illustrate practical spreadsheet automation: preparing structured workbooks, reconciling publisher and product-code data, and moving or removing rows according to review rules.

This repository is a personal portfolio of work from a former role. It is not affiliated with or endorsed by EBSCO Information Services, and it is not maintained as a supported product. The examples reflect a specific historical workflow and should not be assumed to match current systems or procedures. No customer data or sample workbooks are included.

## Macros

| File | Purpose |
| --- | --- |
| [`PAC_sheet_prep.bas`](PAC_sheet_prep.bas) | Renames the first worksheet to `EBSCO`, creates fresh `PUBLISHER` and `TBR` worksheets, turns the source range into `EBSCO_Table`, creates an empty `TBR_Table` with matching headers, and adds a publisher reference table and status key. |
| [`PAC_VLOOKUP.bas`](PAC_VLOOKUP.bas) | Inserts two columns after column O on `EBSCO`, splits column O at hyphens, prompts for a publisher column heading, and adds lookup formulas using the `PUBLISHER` worksheet. |
| [`STATUS_cleanup.bas`](STATUS_cleanup.bas) | Moves rows with selected status codes from `EBSCO_Table` to `TBR_Table`. It also moves all rows for an `SPID NAME` when every row for that SPID has status `XR`. |
| [`purge_dates.bas`](purge_dates.bas) | On the active worksheet, deletes rows from row 2 onward when the value in column V does not end in `9999`. |
| [`purge_packages.bas`](purge_packages.bas) | On the active worksheet, deletes rows where column G contains exactly `P`. The scan includes row 1. |

## Workflow Assumptions

The two `PAC_` macros are designed to be used in order:

1. Start with a copy of the workbook and run `PAC_sheet_prep`.
2. Populate the publisher reference table as needed, then run `PAC_VLOOKUP`.
3. Review the results and run `STATUS_Cleanup` if its status rules apply.

The macros depend on a particular workbook layout, including worksheet names, table names, column positions, and headers such as `PUBL PRODUCT CODE`, `Journal Code`, `SPID NAME`, and `STATUS`. `purge_dates` and `purge_packages` operate on whichever worksheet is active, so confirm the active sheet and the relevant columns before running them.

## Requirements

- Desktop Microsoft Excel with VBA enabled.
- A workbook structured for the assumptions listed above.
- Familiarity with Excel tables and the workbook's data before running a macro.

The modules can be imported into the Visual Basic Editor (`Alt+F11`) or copied into a workbook module. Save a macro-enabled workbook (`.xlsm`) before testing. These files are source modules, not a packaged add-in or a ready-to-run workbook.

## Data Safety

These macros make changes directly to workbook data. `PAC_sheet_prep` deletes and recreates existing worksheets named `PUBLISHER` and `TBR`; the cleanup macros delete or move rows. Run them only on a duplicate workbook, inspect the assumptions and code first, and verify the output before using it further. The macros do not include a built-in undo, confirmation, or recovery mechanism.

The code is provided for review as a portfolio example, not as production-ready software. It has not been validated against current Excel versions, current business rules, or arbitrary workbooks. Adapt and test it against synthetic or otherwise authorized data before reuse. No license is currently specified; publication does not by itself grant permission to reuse the code. Confirm that you have the right to publish or reuse any work-related code and that doing so complies with applicable confidentiality and intellectual-property obligations.
