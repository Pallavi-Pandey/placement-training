# Excel Foundations — Student Handout

**Batch:** 2nd Year MBA · **Session:** Advanced Excel for Business, Day 1 · **Duration:** 3 Hrs

This handout covers everything from today's session — definitions, worked examples, and the hands-on dataset — so you can revise it later without depending on your notes.

**A running example.** Most definitions below use this same 5-row mini table, so the same cells keep meaning the same thing across topics:

| Row | A — Product  | B — Region | C — Qty | D — Unit Price |
| --- | ------------- | ----------- | -------- | --------------- |
| 2   | Office Chair  | North       | 3        | 4200            |
| 3   | Desk Lamp     | South       | 5        | 850             |
| 4   | Whiteboard    | East        | 2        | 1100            |
| 5   | Standing Desk | West        | 1        | 9500            |
| 6   | Printer       | North       | 4        | 7200            |

## Learning Objectives

By the end of today you should be able to:

- Navigate the Excel interface and work faster with core keyboard shortcuts
- Format raw data and convert ranges into structured Excel Tables
- Write basic formulas and copy them correctly using the right reference type
- Use `IF`, `AND`, `OR` — including nested `IF` — to encode business decision rules
- Sort and filter a dataset, including multi-level sorts and combined conditions
- Clean a genuinely messy business export and produce a summary report from it

## Today's Agenda

| Time         | Topic                                 |
| ------------ | ------------------------------------- |
| 0:00 – 0:10 | Why Excel runs business               |
| 0:10 – 0:30 | Interface tour & shortcuts            |
| 0:30 – 0:50 | Formatting & Tables                   |
| 0:50 – 1:05 | Basic formulas                        |
| 1:05 – 1:15 | Break                                 |
| 1:15 – 1:35 | Relative, absolute & mixed references |
| 1:35 – 2:00 | IF, AND, OR                           |
| 2:00 – 2:15 | Sorting & filtering                   |
| 2:15 – 2:30 | Data cleaning                         |
| 2:30 – 2:55 | Capstone exercise                     |
| 2:55 – 3:00 | Recap & exit ticket                   |

---

## 1. Why Excel Runs Business

Before any dashboard, forecasting model, or BI tool enters the picture, the underlying data almost always starts life in a spreadsheet — finance uses it for models, marketing for campaign trackers, operations for inventory, HR for headcount reports. Today builds the foundation — interface fluency, formulas, logic, cleaning — that every later Excel skill (lookups, pivot tables, dashboards) is built on top of.

## 2. Excel Interface & Shortcuts

**Definitions**

| Term                        | Meaning                                                                                  |
| --------------------------- | ---------------------------------------------------------------------------------------- |
| **Workbook**          | The Excel file itself — can contain many sheets                                         |
| **Worksheet (Sheet)** | One grid of cells, accessed via a tab at the bottom                                      |
| **Cell**              | One box in the grid, addressed by column letter + row number, e.g.`B3`                 |
| **Range**             | A rectangular block of cells, written as`A1:C10`                                       |
| **Ribbon**            | The tabbed toolbar at the top — Home, Insert, Page Layout, Formulas, Data, Review, View |
| **Name Box**          | The box left of the formula bar — type a cell address to jump straight to it            |
| **Formula Bar**       | Shows the actual formula behind the active cell, not just its displayed result           |

**Example:** In the mini table above, `B4` refers to the single cell holding `"East"`. `C2:C6` refers to the whole Qty column, rows 2 through 6 — this is what you'd select before writing `=SUM(C2:C6)`.

### Shortcut cheat sheet

| Shortcut             | Action                            |
| -------------------- | --------------------------------- |
| `Ctrl+C / V / X`   | Copy / Paste / Cut                |
| `Ctrl+Z / Y`       | Undo / Redo                       |
| `Ctrl+Arrow`       | Jump to edge of data              |
| `Ctrl+Shift+Arrow` | Select to edge of data            |
| `Ctrl+Home / End`  | Go to A1 / last used cell         |
| `Ctrl+1`           | Format Cells dialog               |
| `Alt+=`            | AutoSum                           |
| `F2`               | Edit active cell                  |
| `F4`               | Toggle reference type ($ locking) |
| `Ctrl+T`           | Convert range to Table            |
| `Ctrl+Shift+L`     | Toggle AutoFilter                 |
| `Ctrl+F / H`       | Find / Replace                    |

## 3. Formatting & Tables

**Definitions**

- **Number format** — how a value *displays*, without changing the value stored underneath. Currency, Percentage, Comma, and Date are the most common; applied via `Ctrl+1`, never by typing `₹` or `%` into the cell.
- **Excel Table** (`Ctrl+T`) — turns a plain range into a structured object: banded rows, headers that stay visible on scroll, automatic expansion when new rows are added, and column names you can use directly in formulas instead of cell letters.
- **Freeze Panes** — keeps header rows visible while scrolling through a long dataset.

**Example:** Select `D2:D6` (the Unit Price column) and apply Currency format via `Ctrl+1` — `4200` displays as `₹4,200.00` while the underlying value used in calculations is still exactly `4200`. Select `A1:D6` and press `Ctrl+T` to turn the mini table into `Table1`; the Unit Price column can now be written as `Table1[Unit Price]` in a formula instead of `D2:D6`.

## 4. Basic Formulas

**Definition:** A formula always starts with `=` and combines cell references with operators (`+ − * / ^`) or functions. A **function** is a named, ready-made formula — e.g. `SUM` — that takes one or more **arguments** in brackets.

| Function                        | What it returns                            | Example (on`C2:C6`) | Result |
| ------------------------------- | ------------------------------------------ | --------------------- | ------ |
| `SUM(range)`                  | Total                                      | `=SUM(C2:C6)`       | 15     |
| `AVERAGE(range)`              | Mean                                       | `=AVERAGE(C2:C6)`   | 3      |
| `COUNT(range)`                | Count of numbers                           | `=COUNT(C2:C6)`     | 5      |
| `COUNTA(range)`               | Count of non-blank cells (numbers or text) | `=COUNTA(A2:A6)`    | 5      |
| `MIN(range)` / `MAX(range)` | Smallest / largest value                   | `=MAX(D2:D6)`       | 9500   |

**Example:** `=D2*C2` on row 2 gives `4200*3 = 12600` — the value of that one order line. `Alt+=` (AutoSum) writes a `SUM` formula for you automatically when you select a column of numbers and press it.

---

### Break (10 min)

---

## 5. Relative, Absolute & Mixed References

**Definitions**

| Type               | Written as                                                                         | Behaviour when copied                               |
| ------------------ | ---------------------------------------------------------------------------------- | --------------------------------------------------- |
| **Relative** | `A1`                                                                             | Both the column and row shift to match the new cell |
| **Absolute** | `$A$1`                                                                           | Nothing shifts — always points at the same cell    |
| **Mixed**    | `$A1` or `A$1` | Only the column (`$A1`) or only the row (`A$1`) is locked |                                                     |

Press `F4` while editing a formula to cycle a reference through all four states: `A1` → `$A$1` → `A$1` → `$A1` → back to `A1`.

**Example — relative:** Add a Total column in `E2` with `=C2*D2` (`3*4200=12600`). Copy it down to `E6`: Excel rewrites it as `=C3*D3`, `=C4*D4`, etc. automatically — exactly right, because each row needs its own Qty and Price.

**Example — absolute:** Put a fixed GST rate in `G1 = 18%`. In `F2` write `=E2*$G$1`. Copy down to `F6` and every row still multiplies by `G1` — none of them drift to `G2`, `G3`, etc. Forgetting the `$` here is the single most common way a formula silently breaks once copied: `=E2*G1` copied to row 3 becomes `=E3*G2`, and `G2` is empty, so Tax quietly becomes 0.

**Example — mixed:** Building a small GST-comparison grid where amounts run down the rows and GST rates run across the columns, the top-left formula needs the amount's *column* locked and the rate's *row* locked, so it can be dragged both right and down at once: `=$C2*D$1`.

**One rule to remember:** if a value should change with the row you copy to, leave it relative. If it should never change, lock it with `$`. If you're filling a grid in two directions, lock only the one axis that shouldn't move for that cell.

## 6. Decision Logic: IF, AND, OR

**Definitions**

| Function     | Syntax                                               | Meaning                                                             |
| ------------ | ---------------------------------------------------- | ------------------------------------------------------------------- |
| `IF`       | `=IF(logical_test, value_if_true, value_if_false)` | Returns one of two results depending on a condition                 |
| `AND`      | `=AND(condition1, condition2, ...)`                | `TRUE` only if *every* condition is true                        |
| `OR`       | `=OR(condition1, condition2, ...)`                 | `TRUE` if *any* condition is true                               |
| Nested`IF` | `IF(...,IF(...,IF(...)))`                          | Chains multiple`IF`s together for more than two possible outcomes |

**Example — IF:** `=IF(D2>5000,"Premium","Standard")` on row 2 (`D2=4200`) returns `"Standard"`; on row 5 (`D5=9500`) it returns `"Premium"`.

**Example — AND:** `=IF(AND(C2>=3,D2>1000),"Bulk Order","Regular")` on row 2 (`Qty=3`, `Price=4200`) — both conditions are true, so it returns `"Bulk Order"`.

**Example — OR:** `=IF(OR(B2="North",B2="East"),"Zone A","Zone B")` on row 2 (`Region="North"`) returns `"Zone A"`; on row 3 (`Region="South"`) it returns `"Zone B"`.

**Example — nested IF:** A 3-tier rating on Unit Price —

```
=IF(D2>=9000,"Gold",IF(D2>=4000,"Silver","Bronze"))
```

Row 2 (`4200`) → checks `>=9000`? No → checks `>=4000`? Yes → `"Silver"`. Row 5 (`9500`) → `>=9000`? Yes → `"Gold"`. Read it as a decision tree: the first branch that matches wins, and the very last `else` is the catch-all that needs no condition of its own.

## 7. Sorting & Filtering

**Definitions**

- **Sort** — reorders rows by one or more columns; a **multi-level sort** (Data → Sort → Add Level) breaks ties in the first column using a second, e.g. Region then Price.
- **Filter / AutoFilter** (`Ctrl+Shift+L`) — temporarily hides rows that don't match a condition, without deleting or reordering anything.
- **Compound filter** — filtering on more than one column at the same time, e.g. Region = North *and* Qty > 2.

**Example:** Sorting the mini table by Region (A→Z) then Price (high to low) puts the two "North" rows together, Printer (7200) above Office Chair (4200). Filtering to `Region = North` afterwards hides every other row, leaving just those two visible — the sort order is preserved underneath.

## 8. Basic Data Cleaning

**Definitions**

| Tool                                   | What it does                                                           |
| -------------------------------------- | ---------------------------------------------------------------------- |
| `TRIM(text)`                         | Removes extra spaces (leading, trailing, and repeated internal spaces) |
| `PROPER(text)`                       | Capitalises the first letter of each word                              |
| `UPPER` / `LOWER(text)`            | Converts to all caps / all lowercase                                   |
| `VALUE(text)`                        | Converts a number stored as text into an actual number                 |
| **Remove Duplicates** (Data tab) | Deletes rows that are fully identical to an earlier row                |
| **Text to Columns** (Data tab)   | Splits one column into several, e.g. by a comma or space               |
| **Find & Replace** (`Ctrl+H`)  | Finds a pattern across the sheet and replaces every instance           |

**Example:** A raw cell containing `"  office CHAIR  "` becomes `"Office Chair"` after `=TRIM(PROPER(A2))`. A price stored as text, `"Rs.4200"`, needs the `"Rs."` stripped (Find & Replace) before `VALUE()` or a plain multiplication will treat it as a number again.

**The one sequencing rule that matters:** `Remove Duplicates` only catches rows that are already *identical*. Two rows for the same order survive it if one customer name is `"rahul mehta"` and the other is `"Rahul Mehta "` — always clean text columns *first*, then remove duplicates, never the other way round.

---

## 9. Capstone Exercise: Clean a Real Dataset, Build a Report

**Files:** `retail_sales_q1_raw.csv` (the messy export) and `region_targets.csv` (a small reference table of Region → Monthly Target → Commission Rate).

`retail_sales_q1_raw.csv` columns: `Order ID, Customer, Region, Product, Category, Qty, Unit Price, Discount %, Order Date, Sales Rep, Repeat Customer, Payment Mode, Order Status`

This export has the kind of problems a real business file actually has — inconsistent text, a few duplicate and near-duplicate rows, some blank or invalid quantities, prices typed with currency symbols, and dates in more than one format. Nothing tells you in advance exactly what's wrong or how many issues there are — part of the exercise is finding them yourself using the tools from today.

### Steps

1. Import the CSV and convert the range to a **Table** (`Ctrl+T`); format **Unit Price** as currency and **Order Date** as a short date once it's genuinely numeric/date-typed (see step 4).
2. Clean **Customer** with `=TRIM(PROPER([@Customer]))`; standardise **Region** spelling with Find & Replace so every casing/spacing variant becomes one value, and decide a rule for any blank Regions (e.g. mark `"Unknown"` rather than guessing).
3. Fix any text-stored prices with Find & Replace to strip currency symbols and thousands separators, then confirm the column is numeric (right-aligned, and `SUM` on it actually returns a number).
4. Standardise **Order Date** into one consistent format — mixed formats won't sort or filter correctly.
5. Decide a rule for zero/negative Qty rows (these are returns, not typos) and for blank Qty rows (`COUNTBLANK` to find them); write your rule in a note cell rather than silently deleting rows.
6. Run **Data → Remove Duplicates** *now*, after the text is clean — cleaning first is what lets this step catch near-duplicates as well as exact ones.
7. Add a **Total Sales** column: `=Qty*[Unit Price]` (relative reference, filled down).
8. Add a **Net Sales** column applying each row's own **Discount %**: `=[@[Total Sales]]*(1-[@[Discount %]])`.
9. Using `region_targets.csv`, add a **Regional Target** column with a nested `IF` on Region, and a **Commission** column that multiplies Net Sales by that region's rate the same way.
10. Add a **Performance Flag**: `=IF([@[Net Sales]]>=[@[Regional Target]],"Achieved","Below Target")`.
11. Add a **High Value Repeat** flag: `=IF(AND([@[Net Sales]]>10000,[@[Repeat Customer]]="Yes"),"High Value","Standard")`.
12. Filter out **Cancelled** orders before computing any revenue KPI — a cancelled order is not revenue.
13. Sort by **Region**, then by **Net Sales** descending; filter to isolate one region, one Sales Rep, or one Order Status at a time.
14. Below the table, build a KPI summary excluding Cancelled orders: Total Net Revenue (`SUM`), Average Order Value (`AVERAGE`), Number of Orders (`COUNT`), Highest and Lowest Sale (`MAX`/`MIN`), and how many are `Achieved` vs `Below Target`.

**Deliverable:** One cleaned, formatted Table plus a KPI summary — the kind of one-page report a manager could act on without asking a clarifying question.

## 10. Exit Ticket

1. Write the absolute reference for cell `B2`, then the mixed reference that locks only its column.
2. Write one formula using `IF` and `AND` that marks a customer `"High Value"` when Order Value exceeds 50,000 *and* they are a repeat customer.
3. Write a nested `IF` that returns `"Gold"` for Sales above 100,000, `"Silver"` above 50,000, and `"Bronze"` otherwise.

**Looking ahead:** Today's foundation — references, logic, clean data — is what lookup functions, pivot tables and dashboards build on top of next.
