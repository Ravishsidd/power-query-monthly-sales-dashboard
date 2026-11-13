# Monthly Sales Analysis with Excel Power Query

A practical sales-data project demonstrating how to consolidate monthly Excel files, clean and validate sales records, enrich transactions with lookup data, calculate approved revenue, and build a refreshable Excel dashboard.

> **Project scope:** The supplied deliverable is an Excel / Power Query dashboard project. It uses PivotTables and PivotCharts for reporting. A Power BI Desktop `.pbix` report is not included in this repository.

## Business problem

A retail sales team receives a separate Excel workbook every month. Manually opening each workbook, copying rows into a master file, repeating cleaning steps, and rebuilding summaries is repetitive and can introduce errors.

The goal is to create a repeatable workflow: connect Power Query to a folder of monthly sales workbooks, apply cleaning and validation steps once, and refresh the prepared data and dashboard when a new month is added.

## Skills demonstrated

- **Excel:** structured tables, workbook organisation, PivotTables, PivotCharts, and dashboard reporting.
- **Power Query:** folder-based import, combining workbooks, transforming data, merging lookup tables, and refreshable queries.
- **Data quality:** trimming identifiers, standardising order statuses, fixing data types, handling conversion errors, identifying exact duplicates, and preserving exception records.
- **Power Query M:** custom columns and reusable expressions for data validation, revenue calculations, and month-start dates.
- **Reporting:** approved order counts and net revenue by month, with optional region and category analysis.

## Data workflow

1. **Extract** — connect to a folder containing monthly workbooks and combine the Excel table named `tblSales`.
2. **Transform** — standardise `CustomerID`, `ProductID`, and order statuses; set appropriate data types; handle invalid quantities; and remove exact duplicate source rows.
3. **Enrich** — merge customer and product lookup tables using left outer joins.
4. **Validate** — assign a `QualityFlag`, separating approved sales from records requiring investigation.
5. **Calculate** — calculate net revenue after discount and create a real date field for monthly grouping.
6. **Load and report** — load approved sales to Excel, then build PivotTables, PivotCharts, and interactive slicers.
7. **Refresh** — add a compatible monthly workbook to the source folder and use **Data → Refresh All**.

## Key business rules and calculations

Approved sales in the supplied training case are completed orders with valid positive quantities and matching customer and product records. The supplied guide also describes preserving rejected records in an exceptions query rather than silently discarding them.

Net revenue is calculated as:

```text
Quantity × UnitPrice_GBP × (1 − DiscountPct)
```

The result is rounded to two decimal places. `DiscountPct` is stored as a decimal (for example, `0.05` represents 5%).

## Expected results from the supplied learning materials

| Reporting period | Approved orders | Approved revenue |
|---|---:|---:|
| January 2026 | 27 | £3,493.83 |
| January–February 2026 | 61 | £9,256.54 |
| January–March 2026 | 94 | £15,770.14 |
| January–April 2026 | 117 | £20,212.54 |

These are the training pack's expected checkpoints, not independently audited commercial results. The dataset is described in the supplied notes as synthetic.

## Repository contents

```text
.
├── README.md
├── .gitignore
├── data/
│   └── Sales_2026_03_March.xlsx
├── docs/
│   ├── Power_Query_Live_Class_Guide.pdf
│   ├── Power_Query_Folder_Class_Notes.md
│   ├── Power_Query_M_Expressions.txt
│   └── READ_ME_FIRST.txt
└── workbook/
    ├── Sales_Analysis_Power_Query_Dashboard.xlsx
    └── Student_Starter.xlsx
```

## Reproducing the full folder-refresh exercise

The repository contains the March source workbook and the supplied dashboard / starter workbooks, but it does **not** currently contain every source dependency described in the training notes. In particular, the January and February monthly workbooks, the customer and product master workbooks, and the April homework workbook are not included here.

To reproduce the complete exercise, obtain those source files and arrange them as described in `docs/Power_Query_Folder_Class_Notes.md`. Keep the report workbook outside the `Monthly_Sales` input folder. Monthly source files should have the same column structure and an Excel table named `tblSales`. Do not put lookup workbooks or unrelated files in the monthly input folder.

Once the Power Query folder connection is configured, add a new monthly file to that folder and run **Data → Refresh All**. The new file is discovered on refresh; copying it into the folder does not necessarily trigger an immediate update.

## Tools

- Microsoft Excel Desktop
- Excel Power Query
- Power Query M
- Excel PivotTables and PivotCharts

## Notes before making the repository public

The live-class PDF is branded training material. Confirm that you have permission to redistribute it publicly before making this repository public. If you are unsure, keep the repository private or remove the branded guide and any other course materials you are not authorised to share.

All sales data is identified by the supplied learning materials as synthetic. Do not add real customer or commercially confidential data to a public repository.
