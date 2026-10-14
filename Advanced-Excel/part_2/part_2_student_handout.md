# Advanced Excel — Student Handout

**Batch:** 2nd Year MBA · **Session:** Advanced Excel for Business, Day 2 · **Duration:** 3 Hrs

This handout covers everything from today's session — definitions, worked examples, function reference tables, and the hands-on exercises — so you can revise it later without depending on your notes.

**Datasets for today.** Two files work together throughout this session:

- **`employee_master.csv`** — a 40-row employee reference table. Columns: `Emp ID, First Name, Last Name, Department, Designation, Hire Date, Base Salary, Bonus %, Region, Manager, Performance Rating, Email, Phone, Status`
- **`monthly_transactions.csv`** — a 180-row expense transaction log. Columns: `Txn ID, Emp ID, Transaction Date, Category, Amount, Payment Mode, Approved By, Remarks, Project Code`

The transactions reference employees by Emp ID — you'll use lookup functions to pull employee details into the transaction sheet, then aggregate and pivot the combined data.

## Learning Objectives

By the end of today you should be able to:

- Use VLOOKUP, XLOOKUP, and INDEX-MATCH to pull data from a reference table
- Aggregate data conditionally with SUMIFS, COUNTIFS, and AVERAGEIFS
- Manipulate text and dates with built-in functions
- Handle errors defensively with IFERROR, IFNA, and ISERROR
- Build PivotTables and PivotCharts to summarise and visualise data interactively

## Today's Agenda

| Time         | Topic                                 |
| ------------ | ------------------------------------- |
| 0:00 – 0:10 | Recap & why lookups replace nested IF |
| 0:10 – 0:35 | VLOOKUP & XLOOKUP                     |
| 0:35 – 0:55 | INDEX-MATCH                           |
| 0:55 – 1:05 | Break                                 |
| 1:05 – 1:30 | SUMIFS, COUNTIFS, AVERAGEIFS          |
| 1:30 – 1:50 | Text & date functions                 |
| 1:50 – 2:05 | Error handling                        |
| 2:05 – 2:15 | Break                                 |
| 2:15 – 2:50 | PivotTables & PivotCharts             |
| 2:50 – 3:00 | Recap & exit ticket                   |

---

## 1. Why Lookups Replace Nested IF

In Day 1 we wrote nested IFs to assign regional targets:

```
=IF(C2="North",180000,IF(C2="South",150000,IF(C2="East",200000,165000)))
```

This works for 4 regions but becomes unreadable at 6 or 10. A **lookup function** stores the values in a separate reference table and lets the formula look them up by key — cleaner, scalable, and far easier to update.

## 2. VLOOKUP

**Definition:** Searches the **first column** of a table range for a value and returns a value from a specified column in the same row.

```
=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
```

| Argument | Meaning |
| --- | --- |
| `lookup_value` | The key to search for (e.g. `"E1001"`) |
| `table_array` | The reference table range — first column must contain the keys |
| `col_index_num` | Which column of the table to return (1 = first, 2 = second, etc.) |
| `range_lookup` | `FALSE` for exact match (almost always); `TRUE` for approximate match |

**Example:** Pull the **Department** for a transaction's employee:

```
=VLOOKUP(B2, employee_master!$A$2:$N$41, 4, FALSE)
```

- `B2` — the Emp ID in the transaction row
- `$A$2:$N$41` — absolute-referenced so copying down doesn't shift the range
- `4` — column 4 = Department
- `FALSE` — exact match

**Three VLOOKUP limitations:**

1. The lookup key must be in the **first column** of `table_array` — no looking to the left
2. `col_index_num` is a hardcoded number — insert a column and it silently returns the wrong data
3. Only returns the **first** match

## 3. XLOOKUP (Excel 365 / 2021+)

**Definition:** The modern replacement for VLOOKUP — no first-column restriction, no fragile column index, and a built-in error message.

```
=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
```

| Argument | Meaning |
| --- | --- |
| `lookup_value` | The key to search for |
| `lookup_array` | The column to search in (any column — not restricted to the first) |
| `return_array` | The column to return from (can be to the left of the lookup column) |
| `if_not_found` | Custom message if no match (replaces the need for IFERROR) |
| `match_mode` | `0` = exact (default), `-1` = exact or next smaller, `1` = exact or next larger, `2` = wildcard |
| `search_mode` | `1` = first-to-last (default), `-1` = last-to-first |

**Example:** Same Department lookup:

```
=XLOOKUP(B2, employee_master!$A$2:$A$41, employee_master!$D$2:$D$41, "Not Found")
```

No column index. Lookup and return columns are specified independently. `"Not Found"` handles missing IDs without a separate error wrapper.

**Example — left-direction lookup** (VLOOKUP can't do this): Given an Email, find the Emp ID:

```
=XLOOKUP(L2, employee_master!$L$2:$L$41, employee_master!$A$2:$A$41, "Not Found")
```

## 4. INDEX-MATCH

**Definition:** A two-function combination that works in every version of Excel. `MATCH` finds the **position** of a value; `INDEX` returns the **value** at that position.

### MATCH

```
=MATCH(lookup_value, lookup_array, [match_type])
```

- `match_type`: `0` = exact match, `1` = largest ≤ value (ascending data), `-1` = smallest ≥ value (descending data)

```
=MATCH("E1010", employee_master!$A$2:$A$41, 0)   → 10
```

### INDEX

```
=INDEX(return_array, row_num)
```

```
=INDEX(employee_master!$D$2:$D$41, 10)   → "IT"
```

### Combined

```
=INDEX(employee_master!$D$2:$D$41, MATCH(B2, employee_master!$A$2:$A$41, 0))
```

"Find where `B2`'s value sits in the Emp ID column, then return that same row's value from the Department column."

**Example — pull the Manager:**

```
=INDEX(employee_master!$J$2:$J$41, MATCH(B2, employee_master!$A$2:$A$41, 0))
```

### When to use which

| Need | Use |
| --- | --- |
| Simple exact lookup, key in first column | VLOOKUP (FALSE) |
| Lookup from any column, built-in error | XLOOKUP (Excel 365+) |
| Works everywhere, any direction | INDEX-MATCH |
| Approximate match (tax brackets, grades) | VLOOKUP (TRUE) or XLOOKUP with match_mode -1 |

---

### Break (10 min)

---

## 5. SUMIFS, COUNTIFS, AVERAGEIFS

**Definition:** Conditional aggregation functions — they answer questions like "total travel spend by the Finance department" without filtering the data first.

### SUMIFS

```
=SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)
```

Adds up all values in `sum_range` where **every** criterion is satisfied simultaneously. The sum range comes **first**, then criteria pairs.

**Example — total Travel spend:**
```
=SUMIFS(E2:E181, D2:D181, "Travel")
```

**Example — Travel spend in Q1 2026 (Jan–Mar):**
```
=SUMIFS(E2:E181, D2:D181, "Travel", C2:C181, ">="&DATE(2026,1,1), C2:C181, "<="&DATE(2026,3,31))
```

> **The `">="&value` pattern:** When a criterion uses a comparison operator, write the operator as a text string and concatenate it with the value using `&`. This is required syntax — not optional.

### COUNTIFS

Same pattern, but counts rows instead of summing:

```
=COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2], ...)
```

**Example — number of Credit Card transactions in Marketing:**
```
=COUNTIFS(F2:F181, "Credit Card", D2:D181, "Marketing")
```

### AVERAGEIFS

Same as SUMIFS but returns the average:

```
=AVERAGEIFS(average_range, criteria_range1, criteria1, ...)
```

**Example — average Hardware purchase amount:**
```
=AVERAGEIFS(E2:E181, D2:D181, "Hardware")
```

### MAXIFS / MINIFS (Excel 2019+)

```
=MAXIFS(max_range, criteria_range1, criteria1, ...)
```

**Example — largest single Training expense:**
```
=MAXIFS(E2:E181, D2:D181, "Training")
```

## 6. Text Functions

Functions that manipulate the text content of cells — cleaning, extracting, combining.

| Function | Syntax | What it does | Example |
| --- | --- | --- | --- |
| `LEFT` | `=LEFT(text, n)` | First *n* characters | `=LEFT("E1001",1)` → `"E"` |
| `RIGHT` | `=RIGHT(text, n)` | Last *n* characters | `=RIGHT("E1001",4)` → `"1001"` |
| `MID` | `=MID(text, start, n)` | *n* chars from position *start* | `=MID("P-201",3,3)` → `"201"` |
| `LEN` | `=LEN(text)` | Character count | `=LEN("Priya Nair")` → `10` |
| `FIND` | `=FIND(find, within)` | Position of first match (case-sensitive) | `=FIND("@","a@b.com")` → `2` |
| `SEARCH` | `=SEARCH(find, within)` | Same as FIND, case-insensitive | `=SEARCH("NAIR","Priya Nair")` → `7` |
| `SUBSTITUTE` | `=SUBSTITUTE(text, old, new)` | Replaces all occurrences | `=SUBSTITUTE("P-201","P-","PRJ-")` → `"PRJ-201"` |
| `CONCATENATE` | `=CONCATENATE(a, b, ...)` | Joins strings | `=CONCATENATE(B2," ",C2)` → `"Priya Nair"` |
| `TEXTJOIN` | `=TEXTJOIN(delim, ignore_empty, ...)` | Joins with delimiter (Excel 2019+) | `=TEXTJOIN(", ",TRUE,A2:A5)` |
| `TEXT` | `=TEXT(value, format)` | Formats a value as text | `=TEXT(42000,"₹#,##0")` → `"₹42,000"` |

**Example — extract email username:**
```
=LEFT(L2, FIND("@", L2)-1)     → "priya.nair"
```

**Example — extract email domain:**
```
=MID(L2, FIND("@", L2)+1, LEN(L2))   → "corpsolutions.com"
```

**Example — full name as "Last, First":**
```
=C2&", "&B2     → "Nair, Priya"
```

## 7. Date Functions

| Function | Syntax | What it returns | Example |
| --- | --- | --- | --- |
| `YEAR` | `=YEAR(date)` | 4-digit year | `=YEAR("12-01-2022")` → `2022` |
| `MONTH` | `=MONTH(date)` | Month (1–12) | `=MONTH("12-01-2022")` → `1` |
| `DAY` | `=DAY(date)` | Day of month | `=DAY("12-01-2022")` → `12` |
| `TODAY` | `=TODAY()` | Current date | Updates daily |
| `DATEDIF` | `=DATEDIF(start, end, unit)` | Date difference | `=DATEDIF(F2, TODAY(), "Y")` → years |
| `EOMONTH` | `=EOMONTH(start, months)` | End of month *n* months away | `=EOMONTH(TODAY(), 6)` |
| `NETWORKDAYS` | `=NETWORKDAYS(start, end)` | Working days (excl. weekends) | `=NETWORKDAYS("01-01-2026","31-03-2026")` |
| `TEXT` | `=TEXT(date, format)` | Custom date display | `=TEXT(F2, "DD-MMM-YYYY")` → `"12-Jan-2022"` |

**Example — years of service:**
```
=DATEDIF(F2, TODAY(), "Y")
```

**Example — transaction quarter:**
```
="Q"&ROUNDUP(MONTH(C2)/3,0)&" "&YEAR(C2)    → "Q1 2026"
```

**Example — probation end date (6 months after hire):**
```
=EOMONTH(F2, 6)
```

**Example — working days in Q1 2026:**
```
=NETWORKDAYS(DATE(2026,1,1), DATE(2026,3,31))    → 64
```

## 8. Error Handling

Errors are signals, not bugs — but in a report, raw `#N/A` or `#DIV/0!` looks broken. Error-handling functions intercept them.

### Common errors

| Error | Typical cause |
| --- | --- |
| `#N/A` | Lookup can't find the value |
| `#VALUE!` | Wrong data type in a formula |
| `#REF!` | Referenced cell was deleted |
| `#DIV/0!` | Division by zero |
| `#NAME?` | Unrecognised function name |

### IFERROR — catches any error

```
=IFERROR(formula, value_if_error)
```

```
=IFERROR(VLOOKUP(B2, employee_master!$A$2:$N$41, 4, FALSE), "Unknown Dept")
```

### IFNA — catches only #N/A

```
=IFNA(VLOOKUP(B2, employee_master!$A$2:$N$41, 4, FALSE), "ID Not Found")
```

**Best practice:** Use `IFNA` for lookups (the expected failure is `#N/A`). Use `IFERROR` only when you genuinely want to suppress *all* error types.

### ISERROR / ISNA — for conditional branching

Returns `TRUE`/`FALSE` — useful inside `IF` when you want to *decide* based on an error, not just suppress it:

```
=IF(ISNA(VLOOKUP(B2, ..., 4, FALSE)), "Invalid ID", "OK")
```

**Warning:** Don't wrap everything in `IFERROR(..., "")` — this hides genuine errors that should be investigated. Only suppress errors you've thought through.

---

### Break (10 min)

---

## 9. PivotTables & PivotCharts

**Definition:** A **PivotTable** is an interactive summary of a dataset. You drag fields into four areas to instantly group, aggregate, and filter data — no formulas needed.

### The four areas

| Area | What goes here | Example |
| --- | --- | --- |
| **Rows** | Categories down the left side | Department, Category |
| **Columns** | Categories across the top | Payment Mode, Quarter |
| **Values** | Numbers being summarised | Sum of Amount, Count of Txn ID |
| **Filters** | Fields that filter the entire PivotTable | Project Code, Region |

### Building a PivotTable

1. Click anywhere inside your data (or select the full range)
2. **Insert → PivotTable** → choose "New Worksheet"
3. In the **PivotTable Field List** pane on the right, drag fields:
   - Drag **Category** to Rows
   - Drag **Amount** to Values (defaults to Sum)
4. Result: total Amount by Category — one row per category, summed automatically

### Making it more useful

**Cross-tab:** Drag **Payment Mode** to Columns → each payment method becomes a column, showing spend by Category × Payment Mode.

**Change aggregation:** Click the dropdown arrow on "Sum of Amount" in Values → **Value Field Settings** → choose Count, Average, Max, etc.

**Date grouping (Excel 365):** Drag a Date field to Rows → right-click any date → **Group** → choose Months, Quarters, or Years. No helper column needed.

**Show Values As:** Right-click a value cell → **Show Values As** → choose **% of Grand Total**, **% of Row Total**, **% of Column Total**, or **Running Total**. Turns raw numbers into proportions.

**Sort:** Click any value cell → sort descending to see the highest-spend categories first.

**Filter:** Use dropdown arrows on Row/Column labels. Or drag a field to the Filters area for a report-level filter.

**Refresh:** When source data changes, right-click the PivotTable → **Refresh**. If you created the PivotTable from a Table (`Ctrl+T`), new rows are included automatically on refresh.

### PivotCharts

1. Click inside the PivotTable
2. **PivotTable Analyze → PivotChart** (or **Insert → PivotChart**)
3. Choose a chart type: Column, Bar, Pie, Line
4. The chart has the same filter buttons — clicking a filter updates both chart and table
5. Rearranging fields in the PivotTable automatically updates the chart

**Example workflow:**
1. PivotTable: Rows = Category, Values = Sum of Amount → insert a Clustered Column chart
2. Add Payment Mode to Columns → chart updates to grouped columns
3. Filter to "Credit Card" and "Bank Transfer" only → chart and table filter together
4. Sort by Sum descending → chart reorders

**Drill-down:** Double-click any value cell in a PivotTable — Excel creates a new sheet with just the underlying rows. Useful for investigating outliers.

---

## 10. Exercises

### Exercise 1: Lookup Drills

Add these columns to `monthly_transactions.csv` by looking up each Emp ID in `employee_master.csv`:

| New Column | Function to Use | Return Column |
| --- | --- | --- |
| Employee Name | XLOOKUP (or VLOOKUP + concatenation) | First Name & Last Name |
| Department | XLOOKUP or INDEX-MATCH | Department |
| Designation | INDEX-MATCH | Designation |
| Hire Date | XLOOKUP | Hire Date |
| Base Salary | VLOOKUP | Base Salary |
| Region | INDEX-MATCH | Region |
| Manager | INDEX-MATCH | Manager |
| Performance Rating | XLOOKUP | Performance Rating |

Then answer:
a) Which employee (by name) has the most transactions?
b) What is the average transaction amount for employees rated 5?
c) Are there transactions by "On Leave" or "Resigned" employees?

### Exercise 2: Reverse and Approximate Lookups

a) Given a Manager name, find the first Emp ID they manage. Which lookup method works?

b) Create this tax-bracket table and use VLOOKUP (TRUE) or XLOOKUP (match_mode -1) to find each employee's tax rate:

| Income Slab | Tax Rate |
| --- | --- |
| 0 | 0% |
| 30000 | 5% |
| 50000 | 10% |
| 75000 | 15% |
| 100000 | 20% |

### Exercise 3: SUMIFS / COUNTIFS / AVERAGEIFS Dashboard

Build these metrics — each with a single formula, no filtering:

| # | Metric |
| --- | --- |
| 1 | Total spend on Travel |
| 2 | Total Marketing spend for Project P-403 |
| 3 | Number of Credit Card transactions |
| 4 | Transactions approved by "Fatima Sheikh" in Q1 (Jan–Mar 2026) |
| 5 | Average Hardware purchase amount |
| 6 | Total Finance department spend in Q2 (Apr–Jun 2026) |
| 7 | Petty Cash transactions under ₹3,000 |
| 8 | Total Training spend in the North region |
| 9 | Transactions where Amount > ₹30,000 |
| 10 | Average Software License cost paid by Bank Transfer |

### Exercise 4: Text Functions

Using `employee_master.csv`:

a) **Full Name:** `=C2&", "&B2` → `"Nair, Priya"`

b) **Email username:** `=LEFT(L2, FIND("@",L2)-1)` → `"priya.nair"`

c) **Email domain:** `=MID(L2, FIND("@",L2)+1, LEN(L2))` → `"corpsolutions.com"`

d) **Masked phone:** `="XXXXXX"&RIGHT(M2,4)` → `"XXXXXX3210"`

e) **Standardise "Sr." to "Senior":** `=SUBSTITUTE(E2, "Sr.", "Senior")`

f) **Project number from code:** `=MID(I2,3,3)` on `"P-201"` → `"201"`

g) **Narrative sentence:** `=B2&" "&C2&" from "&D2&" spent ₹"&TEXT(E2,"#,##0")&" on "&D2` (adapt columns as needed)

### Exercise 5: Date Functions

a) **Years of service:** `=DATEDIF([@[Hire Date]], TODAY(), "Y")`

b) **Tenure bucket:** `=IF(DATEDIF(F2,TODAY(),"Y")>=5, "5+ years", IF(DATEDIF(F2,TODAY(),"Y")>=2, "2–5 years", "< 2 years"))`

c) **Hire quarter:** `="Q"&ROUNDUP(MONTH(F2)/3,0)&" "&YEAR(F2)`

d) **Transaction quarter:** Same formula on the Transaction Date column

e) **Probation end:** `=EOMONTH(F2, 6)`

f) **Working days in Q1 2026:** `=NETWORKDAYS(DATE(2026,1,1), DATE(2026,3,31))`

g) **Formatted label:** `=TEXT(F2, "DD-MMM-YYYY")` → `"12-Jan-2022"`

### Exercise 6: Error Handling

a) Change one Emp ID to `"E9999"`. Observe the `#N/A`. Wrap in `IFNA` with `"ID Not Found"`.

b) Calculate Average Spend per Category = `SUMIFS / COUNTIFS`. Handle `#DIV/0!` with `IFERROR(..., 0)`.

c) Create a validation column: `=IF(ISNA(VLOOKUP(B2,...)), "Invalid ID", "OK")`

### Exercise 7: PivotTable & PivotChart

a) PivotTable: Rows = Category, Values = Sum of Amount. Which category spends the most?

b) Add Department to Rows, Category to Columns → cross-tab. Which dept-category pair dominates?

c) Add a Month column → Rows = Month, Columns = Category. Any spending trends?

d) Rows = Approved By, Values = Sum of Amount + Count of Txn ID. Who approves the most? The highest value?

e) Filter by Project Code = P-102. Which categories dominate?

f) Insert a Clustered Column PivotChart from the Category PivotTable. Add Payment Mode to Columns. Filter to Credit Card + Bank Transfer.

g) Build a Line PivotChart: Rows = Month, Values = Sum of Amount. Add a trendline. Trend direction?

h) Double-click the highest value cell — explore the drill-down detail sheet.

## 11. Exit Ticket

1. Write a VLOOKUP that finds `"E1015"` in the employee master and returns the Department. What changes for Base Salary?
2. Rewrite it using INDEX-MATCH. Why is INDEX-MATCH more robust when columns are inserted?
3. Write a SUMIFS for total "Software License" spend paid by "Bank Transfer".
4. Write a formula to extract the email username (before `@`) from cell L2.
5. A VLOOKUP returns `#N/A` for a new employee. Wrap it in an error handler. Would you use IFERROR or IFNA — and why?

## 12. Exercise Solutions

Worked against the actual `employee_master.csv` (40 employees) and `monthly_transactions.csv` (180 rows) saved alongside this handout — column letters below assume the raw files with headers in row 1, data starting row 2. `employee_master.csv` columns: `A Emp ID · B First Name · C Last Name · D Department · E Designation · F Hire Date · G Base Salary · H Bonus % · I Region · J Manager · K Performance Rating · L Email · M Phone · N Status`. `monthly_transactions.csv` columns: `A Txn ID · B Emp ID · C Transaction Date · D Category · E Amount · F Payment Mode · G Approved By · H Remarks · I Project Code`.

### Exercise 1: Lookup Drills

| New Column | Formula |
| --- | --- |
| Employee Name | `=XLOOKUP(B2,employee_master!$A:$A,employee_master!$B:$B)&" "&XLOOKUP(B2,employee_master!$A:$A,employee_master!$C:$C)` |
| Department | `=XLOOKUP(B2,employee_master!$A:$A,employee_master!$D:$D)` |
| Designation | `=INDEX(employee_master!$E:$E,MATCH(B2,employee_master!$A:$A,0))` |
| Hire Date | `=XLOOKUP(B2,employee_master!$A:$A,employee_master!$F:$F)` |
| Base Salary | `=VLOOKUP(B2,employee_master!$A:$N,7,FALSE)` |
| Region | `=INDEX(employee_master!$I:$I,MATCH(B2,employee_master!$A:$A,0))` |
| Manager | `=INDEX(employee_master!$J:$J,MATCH(B2,employee_master!$A:$A,0))` |
| Performance Rating | `=XLOOKUP(B2,employee_master!$A:$A,employee_master!$K:$K)` |

a) **Most transactions:** Neha Sharma (`E1003`) — 10 transactions, the most of any employee in the log.

b) **Average amount for rating-5 employees:** ₹28,600, across 53 transactions by employees rated 5. Formula once the Performance Rating column is pulled in: `=AVERAGEIFS(monthly_transactions!E:E,monthly_transactions!K:K,5)` (K here being wherever you placed the looked-up rating).

c) **Transactions by "On Leave"/"Resigned" employees:** none. The employee master has 2 "On Leave" and 1 "Resigned" employee, but cross-checking their Emp IDs against the transaction log turns up zero matching rows — the honest answer is "no," and the point of the question is confirming you actually checked rather than assumed a messy dataset must contain one.

### Exercise 2: Reverse and Approximate Lookups

a) **Manager → first Emp ID:** VLOOKUP can't do this — Emp ID (column A) sits to the *left* of Manager (column J), and VLOOKUP can only return columns to the right of its lookup column. Use INDEX-MATCH or XLOOKUP instead:

```
=INDEX(employee_master!$A:$A,MATCH("Sameer Joshi",employee_master!$J:$J,0))
```

For Manager `"Sameer Joshi"`, this returns `E1001` (Priya Nair) — the first row in the table where Manager equals that name.

b) **Tax bracket lookup:**

```
=VLOOKUP(G2,TaxTable!$A:$B,2,TRUE)
```

(or `=XLOOKUP(G2,TaxTable!$A:$A,TaxTable!$B:$B,,-1)` for the XLOOKUP equivalent — `TRUE`/`match_mode -1` both mean "find the largest bracket value that doesn't exceed the salary.") Worked example: employee `E1006`, Base Salary ₹62,000 — falls in the ₹50,000 bracket (the next one up, ₹75,000, is too high), so the lookup returns a **10%** tax rate.

### Exercise 3: SUMIFS / COUNTIFS / AVERAGEIFS Dashboard

| # | Formula | Answer |
| --- | --- | --- |
| 1 | `=SUMIFS(E:E,D:D,"Travel")` | ₹6,27,200 |
| 2 | `=SUMIFS(E:E,D:D,"Marketing",I:I,"P-403")` | ₹8,44,800 |
| 3 | `=COUNTIFS(F:F,"Credit Card")` | 81 |
| 4 | `=COUNTIFS(G:G,"Fatima Sheikh",C:C,">="&DATE(2026,1,1),C:C,"<="&DATE(2026,3,31))` | 12 |
| 5 | `=AVERAGEIFS(E:E,D:D,"Hardware")` | ₹38,214.29 |
| 6 | `=SUMIFS(E:E,J:J,"Finance",C:C,">="&DATE(2026,4,1),C:C,"<="&DATE(2026,6,30))` — needs the Department lookup column from Exercise 1 first, here assumed placed in column J | ₹3,38,600 |
| 7 | `=COUNTIFS(F:F,"Petty Cash",E:E,"<3000")` | 15 |
| 8 | `=SUMIFS(E:E,D:D,"Training",K:K,"North")` — needs the Region lookup column from Exercise 1, here assumed placed in column K | ₹1,49,500 |
| 9 | `=COUNTIFS(E:E,">30000")` | 37 |
| 10 | `=AVERAGEIFS(E:E,D:D,"Software License",F:F,"Bank Transfer")` | ₹42,916.67 |

**Watch #7:** "Petty Cash" is a value in the **Payment Mode** column, not the Category column — there's no "Petty Cash" category in this dataset at all. Criteria on the wrong column silently return 0 rather than erroring, so this one is a built-in trap for checking your criteria ranges match the actual field you mean.

### Exercise 4: Text Functions

Each sub-question already shows its formula and expected output inline above — restated as a group for reference:

| | Formula | Result |
| --- | --- | --- |
| a | `=C2&", "&B2` | `"Nair, Priya"` |
| b | `=LEFT(L2,FIND("@",L2)-1)` | `"priya.nair"` |
| c | `=MID(L2,FIND("@",L2)+1,LEN(L2))` | `"corpsolutions.com"` |
| d | `="XXXXXX"&RIGHT(M2,4)` | `"XXXXXX3210"` |
| e | `=SUBSTITUTE(E2,"Sr.","Senior")` | turns `"Sr. Analyst"` into `"Senior Analyst"` |
| f | `=MID(I2,3,3)` | `"201"` from `"P-201"` |

### Exercise 5: Date Functions

Also answered inline above; grounded against a real row — employee `E1001` (Priya Nair), hired `12-01-2022`:

| | Formula | Result (as of today) |
| --- | --- | --- |
| a | `=DATEDIF(F2,TODAY(),"Y")` | `4` years of service |
| b | `=IF(DATEDIF(F2,TODAY(),"Y")>=5,"5+ years",IF(DATEDIF(F2,TODAY(),"Y")>=2,"2–5 years","< 2 years"))` | `"2–5 years"` |
| c | `="Q"&ROUNDUP(MONTH(F2)/3,0)&" "&YEAR(F2)` | `"Q1 2022"` |
| e | `=EOMONTH(F2,6)` | `31-07-2022` |
| g | `=TEXT(F2,"DD-MMM-YYYY")` | `"12-Jan-2022"` |

### Exercise 6: Error Handling

a) `=IFNA(VLOOKUP("E9999",employee_master!$A:$N,4,FALSE),"ID Not Found")` → `E9999` doesn't exist in the master, so this returns `"ID Not Found"` instead of a raw `#N/A`.

b) `=IFERROR(SUMIFS(E:E,D:D,A2)/COUNTIFS(D:D,A2),0)` — any category with zero matching transactions would otherwise divide by zero; wrapping it returns `0` instead of `#DIV/0!`.

c) `=IF(ISNA(VLOOKUP(B2,employee_master!$A:$N,1,FALSE)),"Invalid ID","OK")` — flags any Emp ID in the transaction log that doesn't exist in the employee master.

### Exercise 7: PivotTable & PivotChart

a) **Top category by spend:** Marketing (₹8,44,800) — narrowly ahead of Hardware (₹8,02,500). Full ranking: Marketing → Hardware → Travel (₹6,27,200) → Training (₹6,27,000) → Software License (₹6,24,500) → Office Supplies (₹1,33,400).

b) **Dominant Department × Category pair:** Marketing dept × Marketing category (₹8,44,800) overall. Within IT, two categories stand out: IT × Hardware (₹8,02,500) and IT × Software License (₹4,14,500) — IT alone accounts for well over half of total spend once both are combined.

c) **Monthly trend:** spend climbs from ₹4.86L in January to a peak of ₹6.79L in April, dips slightly in May (₹5.91L), then recovers in June (₹6.48L) — a gentle rise through Q1 into Q2 rather than a sharp trend in either direction.

d) **By Approved By:** Kiran Patel approves the highest *total value* (₹14.25L) despite approving fewer transactions (39) than Sameer Joshi, who approves the *most transactions* (50) but a smaller total value (₹6.05L). Worth drawing out explicitly: highest count and highest value don't have to be the same person.

e) **Project P-102:** Hardware dominates (21 of 39 rows), followed by Software License (13) and Training (5) — consistent with P-102 reading as an IT infrastructure project.

f)–g) No single numeric answer — these are chart-building steps. The trend in the resulting Line PivotChart matches part (c): a gradual climb from January into April, then a slight dip and recovery.

h) **Drill-down:** the single highest transaction in the dataset is ₹85,000 — `T5066`, Hardware, approved by Kiran Patel, Project `P-102`, dated 18-03-2026. This is the row a drill-down on the highest value cell would surface first.

## 13. Exit Ticket — Solutions

**1. VLOOKUP for `E1015`'s Department, and what changes for Base Salary**

```
=VLOOKUP("E1015",employee_master!$A:$N,4,FALSE)
```

Returns `"Finance"`. For Base Salary, only the `col_index_num` changes — from `4` to `7` (Base Salary is the 7th column in the master table):

```
=VLOOKUP("E1015",employee_master!$A:$N,7,FALSE)
```

Returns `33000`. (`E1015` is Ritika Chawla, Finance, Base Salary ₹33,000.)

**2. INDEX-MATCH rewrite**

```
=INDEX(employee_master!$D:$D,MATCH("E1015",employee_master!$A:$A,0))
```

More robust because the column to return is specified directly (`$D:$D` for Department) rather than counted as "the 4th column." If someone inserts a new column anywhere in the table, Excel automatically updates the `$D:$D` reference to still point at Department — but a hardcoded `4` in VLOOKUP keeps pointing at whatever now happens to sit in position 4, silently returning the wrong field with no error at all.

**3. SUMIFS for Software License spend via Bank Transfer**

```
=SUMIFS(E:E,D:D,"Software License",F:F,"Bank Transfer")
```

→ ₹5,15,000.

**4. Extract email username**

```
=LEFT(L2,FIND("@",L2)-1)
```

**5. IFERROR or IFNA?**

Use `IFNA`. The anticipated failure here is specifically `#N/A` — "this ID doesn't exist." `IFERROR` would also silently swallow unrelated errors, like a `#REF!` from a deleted column or a `#VALUE!` from a typo elsewhere in the formula — exactly the kind of error you want to *see* and fix, not hide behind a friendly fallback value.

**Looking ahead:** Today's lookups, aggregation, and PivotTables are the analytical core — next session covers data validation, conditional formatting, what-if analysis, and dashboard design.
