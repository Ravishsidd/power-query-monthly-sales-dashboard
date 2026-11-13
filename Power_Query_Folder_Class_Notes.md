# Power Query: Folder Consolidation, Data Cleaning & Auto-Refreshing Excel Dashboard
**Instructor-led session (3 hours) | Excel desktop | UK retail sales case study**

## Learning outcomes
By the end, learners will import a folder of similarly structured Excel workbooks, combine monthly data into one query, clean identifiers and statuses, handle type-conversion errors, remove exact duplicate records, merge two lookup tables, define approved revenue, create a PivotTable and PivotChart, and refresh the whole reporting pipeline after adding a new monthly file.

## Business scenario
You are an analyst at a UK office-products retailer. The sales operations team exports a separate `Sales_YYYY_MM_Month.xlsx` workbook every month. Management needs one report of **completed and approved** orders, monthly net revenue, category revenue, regional performance and order counts. They want to add a new workbook to the source folder and refresh, rather than copy-pasting rows or rebuilding a report.

## Pack contents
- `Monthly_Sales/`: Sales January and February (start here)
- `ADD_IN_CLASS/`: Sales March (add midway through class)
- `HOMEWORK_ADD_LATER/`: Sales April (learner independent exercise)
- `Reference_Data/Customer_Master.xlsx`: `tblCustomers` with customer region and city
- `Reference_Data/Product_Master.xlsx`: `tblProducts` with product name/category
- `Student_Starter.xlsx`: the blank dashboard workspace and learner checklist
- `Instructor_Only/Power_Query_Instructor_Guide_and_Answer_Key.xlsx`: script, source register, approved rows, exceptions, checks, and an expected-revenue chart

All sales workbooks contain one Excel Table called **tblSales** with identical column names (the order dates and quantities/discounts vary). Do not save the report workbook inside `Monthly_Sales`.

## Three-hour timetable
| Phase | Time | Live activity |
|---|---:|---|
| ETL and the analyst's monthly reporting problem | 15 min | Explain files → Power Query → clean result → pivot |
| Connect From Folder | 20 min | Combine January and February files |
| Data cleaning, duplicate controls, quality flags | 30 min | Trim, status, type errors, null and duplicates |
| Merge reference tables | 20 min | Expand CustomerName / Region / ProductName / Category |
| Revenue and time dimensions | 15 min | Add Revenue_GBP, MonthStart |
| Build PivotTable and PivotChart | 25 min | Revenue by month, count of orders |
| March file + Refresh All demonstration | 15 min | Third month appears; check totals |
| Learner hands-on, April challenge, debugging | 40 min | Repeat and discuss real-world pitfalls |
Total: **180 minutes**

## Script to introduce the class (say this aloud)
“Suppose I give you separate January and February Excel sales files. You've prepared a PivotTable and chart. Next month, March's file arrives. What would happen if you were managing this manually? You might open each workbook, copy its data, paste it into a master file, repeat your cleaning rules, and then refresh the report. Power Query lets us connect to the **folder** instead of selecting each workbook individually. We define our transformation steps once. When March is added to the folder, Refresh All re-runs the query and updates the underlying output. Our final report and PivotChart can then refresh from that table.”

Clarify that **Refresh All is still a trigger**. Adding a file does not universally trigger a live instant update. Advanced refresh-on-open and unattended scheduling have different prerequisites.

## Module 1: Prepare and understand the data (15 min)
1. Extract the pack and inspect its directory structure.
2. Open Sales_2026_01_January.xlsx and Sales_2026_02_February.xlsx.
3. Notice each workbook has the same headers and a table named `tblSales`.
4. Discuss column *grain*: one line is one synthetic order; in a real transaction dataset, there may be multiple lines per order, so don't remove duplicates using OrderID by default.
5. Demonstrate inconsistencies: leading/trailing spaces in IDs, ` COMPLETE ` vs completed, quantity `bad`, quantity 0, and unknown reference IDs.
6. Explain the difference between invalid data and normal business records: Cancelled, Returned and Pending may be legitimate records but are excluded from an **approved completed-sales** report. Returns would require separate treatment in a real accounting revenue model.

## Module 2: Connect to the folder (20 min)
**Click path (Windows Excel):**
Data → Get Data → From File → From Folder → browse to `Monthly_Sales` → Combine → Combine & Transform Data.
- Select the object `tblSales` in the Combine Files dialog.
- The sample-file query and transform function are automatically generated.
- Explain `Source.Name` as lineage of each transaction.
- Do not point the connector at the parent of Monthly_Sales; that parent includes the reference files, homework, workbook, and answer key.
- Verify only January and February are included.

**Mentor concept:** Append and From Folder both vertically consolidate consistent schemas. But From Folder discovers the files in a source location each time the query refreshes. You don't manually add each new month as a separate Append input.

## Module 3: Transformations and quality checks (30 min)
On the combined query in Power Query Editor:
1. Rename the final query `All_Sales_Raw`.
2. Trim `CustomerID` and `ProductID` (Transform → Format → Trim).
3. Trim and then lowercase `OrderStatus`.
4. Use Add Column → Custom Column for `StatusClean`:
```
if [OrderStatus] = "complete" then "completed" else [OrderStatus]
```
5. Change OrderDate to Date, Quantity to Whole Number, UnitPrice_GBP to Decimal Number, DiscountPct to Decimal Number. If 'bad' becomes an Error in Quantity, explain why this is an invalid value; **Replace Errors** with null. Keep the exception data visible in an audit query.
6. Remove exact duplicates across **all nine original source columns** (select those nine columns → Remove Rows → Remove Duplicates). Do not include query-added flags or arbitrary row indexes as unique keys. There is one identical duplicate row in February, and one additional identical duplicate in March.
7. Explain how removing duplicates before vs after trimming/normalization can change what Power Query considers a duplicate.

**Question:** “If OrderID appears twice because the customer bought two different products in one order, should we remove one?” Answer: No — the real business grain and keys determine what a duplicate means.

## Module 4: Merge two lookup tables (20 min)
- Data → Get Data → From Workbook. Import the `tblCustomers` table from Customer_Master.xlsx and `tblProducts` from Product_Master.xlsx.
- Ensure CustomerID and ProductID types are Text.
- Return to All_Sales_Raw, Merge Queries with Customers using CustomerID, **Left Outer Join**. Expand CustomerName and Region.
- Merge with Products using ProductID, **Left Outer Join**. Expand ProductName and Category.
- Leave unmatched IDs visible; source has `C999` and `P999`. You should not silently assign an unrelated customer or product.
- Show Left Anti as an alternate way to discover missing keys.
- Explain duplicate IDs in the lookup table can multiply sales rows during a Merge, inflating reported revenue.

## Module 5: Apply approval rules (15 min)
After lookup expansions, add column `QualityFlag`:
```
if [StatusClean] <> "completed" then "Not completed"
else if [Quantity] = null then "Invalid quantity"
else if [Quantity] <= 0 then "Non-positive quantity"
else if [CustomerName] = null then "Unknown customer"
else if [ProductName] = null then "Unknown product"
else "Approved"
```
These are teaching rules — in production, businesses formally define cancellations, returns, missing data, price validation and revenue-recognition policy.

Right-click the query → **Reference** → rename `Approved_Sales`, filter QualityFlag = Approved.
Make another **Reference** query → `Exception_Orders`, filter QualityFlag <> Approved. For strict reconciliation, count exact duplicates separately (duplicates are removed before the Approved/Exception split).

Add a `Revenue_GBP` custom column to Approved_Sales:
```
Number.Round(
    [Quantity] * [UnitPrice_GBP] * (1 - [DiscountPct]),
    2,
    RoundingMode.AwayFromZero
)
```
Change Revenue_GBP to Decimal Number. This gives net order revenue after discount.

Add MonthStart custom column:
```
Date.StartOfMonth(Date.From([OrderDate]))
```
Change MonthStart to Date. Use it for chronologically ordered month labels in PivotTables; display as `mmm yyyy`.

## Module 6: Load and build the report (25 min)
1. In Approved_Sales, choose Home → Close & Load To → Table → New Worksheet.
2. Select the loaded table in Excel; rename the Excel Table `tblCombinedSales` on Table Design.
3. Click inside this output Table → Insert → PivotTable.
4. Pivot fields: Rows = MonthStart; Values = Sum Revenue_GBP; Values = Count OrderID.
5. Add Region or Category to filters/slicers if desired.
6. Format revenue as GBP and MonthStart as `mmm yyyy` (or use MonthStart as an actual date and group appropriately).
7. PivotTable Analyze → PivotChart → Clustered Column. Place on the Dashboard sheet.
8. Create a second Pivot by Category for the optional chart.

**Interpretation prompt:** “Which month earned more? Does the month with the most orders always produce the most revenue? Why might average order value be more informative?”

## Module 7: Demonstrate refresh live (15 min)
Before adding March, the approved-sales report should show:
- 61 approved orders
- £9,256.54 total approved revenue

Now:
1. Copy **Sales_2026_03_March.xlsx** from ADD_IN_CLASS into Monthly_Sales.
2. In the Excel reporting workbook, select **Data → Refresh All**.
3. Wait for data load to finish. If needed, right-click the PivotTable → Refresh.
4. Verify the PivotChart has a March bar.
5. New totals: **94 approved orders** and **£15,770.14** across January–March.
6. March's own contribution: **33 approved orders** and **£6,513.60**.

**Mentor line:** “Notice what we did *not* do: we did not open March and copy-paste any orders into the master file. We didn't modify our Power Query to add March as a separately named workbook. The query knows to look in the folder every time it refreshes.”

## Module 8: Homework / challenge (40 min)
Do not copy April before the learner task.
1. Add Sales_2026_04_April.xlsx to Monthly_Sales and refresh the query/PivotTable.
2. Verify April's approved orders and revenue.
3. Build a PivotTable showing Revenue_GBP by Region, and another by Category.
4. Create a bar chart showing category contributions.
5. Show source file names in the combined result to prove data lineage.
6. Find one Invalid quantity, one Unknown customer, one Unknown product, and one Exact duplicate.
7. Explain a scenario where From Folder can be dangerous: wrong workbook added to source folder, changed column names, different table name, or duplicate transactions loaded twice.

## Exact instructor checkpoints
| Month | Rows in file | Full-row duplicates | Approved orders | Approved GBP revenue | Other excluded rows |
|---|---:|---:|---:|---:|---:|
| Jan 2026 | 52 | 0 | 27 | £3,493.83 | 25 |
| Feb 2026 | 57 | 1 | 34 | £5,762.71 | 22 |
| Mar 2026 | 61 | 1 | 33 | £6,513.60 | 27 |
| Apr 2026 | 48 | 0 | 23 | £4,442.40 | 25 |

**Cumulative checkpoints**
- Jan + Feb: 109 imported rows, 1 duplicate, 61 approved orders, £9,256.54.
- Jan + Feb + Mar: 170 imported rows, 2 duplicates, 94 approved orders, £15,770.14.
- Jan + Feb + Mar + Apr: 218 imported rows, 2 duplicates, 117 approved orders, £20,212.54.

The Instructor answer key workbook contains every approved transaction and the reason each excluded transaction is not counted. The supplied workbook deliberately uses stricter rules than a simplistic `Status = completed` filter so students learn to verify data quality.

## Recap questions for learners
1. What is the difference between Append Queries and From Folder?
2. Why do monthly files need the same schema and Excel Table name?
3. Why would a PivotChart remain unchanged even after March enters the source folder?
4. What is the purpose of Source.Name?
5. Why do we trim identifier columns before merging?
6. What is the difference between an error, null and zero?
7. Why is removing duplicates based only on OrderID dangerous?
8. What happens if a lookup table contains duplicate CustomerID values?
9. Is inserting a March workbook into the source folder an automatic real-time trigger?
10. How would you investigate a report total that doesn't match your expected revenue?

## Troubleshooting
- **March missing:** copy it into Monthly_Sales, not just ADD_IN_CLASS; confirm the folder path; use Excel Data → Refresh All.
- **Pivot still old:** refresh PivotTable again after the query finishes.
- **Query errors after new workbook:** verify the name of Excel Table (`tblSales`) and column names/types; ensure no stray file is in Monthly_Sales.
- **Lots of blank Region/Category values:** verify ID text types and whitespace; use Left Anti to find unmatched foreign keys.
- **Month labels sorted incorrectly:** use real MonthStart dates, not preformatted textual abbreviations.
- **Rows multiply unexpectedly:** check duplicate lookup IDs and accidental source-file duplicates.
- **Close & Load To greyed out:** close and load initially, then use Data → Queries & Connections → Load To.
- **Refresh does not show source changes:** Power Query Editor's Refresh Preview is not the same as refreshing the worksheet output.

## Next-class extension ideas
- Refresh a SharePoint-hosted folder of similar workbooks.
- Extract Year, Quarter, Year-Month and MonthStart from OrderDate.
- Build a reusable parameter for source folder path.
- Separate corrections from monthly history to prevent retroactive double-counting.
- Design a data quality dashboard: excluded row count, invalid quantity rate, unmatched key rate, duplicate rate.
- Compare Power Query folder refresh with scheduled data pipelines in larger systems.

## Microsoft reference
- https://support.microsoft.com/en-us/excel/import-data-from-a-folder-with-multiple-files-power-query
- https://support.microsoft.com/en-us/excel/refresh-an-external-data-connection-in-excel
- https://support.microsoft.com/en-us/excel/refresh-pivottable-data

*All data in this teaching package is fictional.*
