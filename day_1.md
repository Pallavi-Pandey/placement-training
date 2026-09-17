# Day 1 — Excel Basics & Foundation

**Batch:** 2nd Year MBA
**Duration:** 3 Hrs
**Format:** Lecture + Hands-on Lab
**Prerequisite:** None
**Tools (this copy):** LibreOffice Calc 7.x
**Tools (students):** Microsoft Excel 2016+ / 365

> This is your working/trainer copy, written for **LibreOffice Calc** — the app you actually prep and test in. Students are on Microsoft Excel, so hand them `day_1_student_handout.md` / `.pdf`, which uses genuine Excel Tables and structured references (`[@ColumnName]`). Both documents solve the same exercise with the tool each side actually has open — don't mix the mechanics between them.

A hands-on session that takes students from the ribbon and formula bar to a cleaned, calculated, sorted business dataset — the working vocabulary every later Excel module builds on.

## Learning Objectives

By the end of this session, students will be able to:

- Navigate the Excel interface confidently and work faster with core keyboard shortcuts
- Format raw data and convert ranges into structured Excel Tables
- Write basic formulas and copy them correctly using the right reference type
- Use `IF`, `AND`, `OR` — including nested `IF` ladders — to encode multi-tier business rules
- Sort and filter a dataset, including multi-level sorts and compound filter criteria, to answer specific business questions
- Diagnose and clean a genuinely messy, multi-column business export — mixed date formats, text-stored numbers, near-duplicates, invalid entries — and produce a decision-ready summary report from it

## Session Run-of-Show

| Time         | Duration | Block                                              | Type     |
| ------------ | -------- | -------------------------------------------------- | -------- |
| 0:00 – 0:10 | 10 min   | Welcome & why Excel runs business                  | Concept  |
| 0:10 – 0:30 | 20 min   | Excel interface tour & core shortcuts              | Concept  |
| 0:30 – 0:50 | 20 min   | Formatting data & building Tables                  | Concept  |
| 0:50 – 1:05 | 15 min   | Basic formulas: SUM, AVERAGE, COUNT, MIN/MAX       | Concept  |
| 1:05 – 1:15 | 10 min   | Break                                              | —       |
| 1:15 – 1:35 | 20 min   | Relative, absolute & mixed references              | Concept  |
| 1:35 – 2:00 | 25 min   | Decision logic: IF, AND, OR                        | Concept  |
| 2:00 – 2:15 | 15 min   | Sorting & filtering                                | Concept  |
| 2:15 – 2:30 | 15 min   | Basic data cleaning                                | Concept  |
| 2:30 – 2:55 | 25 min   | Capstone: clean a dataset & build a summary report | Hands-on |
| 2:55 – 3:00 | 5 min    | Recap, exit ticket & Q&A                           | Wrap-up  |

---

## 1. Welcome & Why Excel Runs Business (0:00 – 0:10)

Open with a quick show of hands: who has used Excel for coursework, an internship MIS report, or a budget? Ground the session in the fact that before any dashboarding tool, forecasting model or BI platform is introduced, the underlying data almost always starts life in a spreadsheet — finance uses it for models, marketing for campaign trackers, operations for inventory, HR for headcount reports.

Set expectations: this is a laptop-open, hands-on session. Today builds the foundation — interface fluency, formulas, logic, cleaning — that every later Excel module (lookups, pivot tables, dashboards) depends on.

**Opener:** Show a "before / after" screenshot of a messy sales sheet next to the cleaned, summarised version students will produce in today's capstone — this is the destination for the next three hours.

## 2. Excel Interface Tour & Core Shortcuts (0:10 – 0:30)

Walk through the anatomy of the screen before touching data: the **Ribbon** (Home, Insert, Page Layout, Formulas, Data, Review, View), the **Name Box** and **Formula Bar**, the grid of cells addressed by column-letter + row-number (`A1` notation), sheet tabs, and the status bar's live Sum/Average/Count.

- **Workbook → Worksheet → Cell**: one file can hold many sheets; one sheet is a grid of cells
- A **range** is a rectangle of cells, written as `A1:C10`
- The Name Box lets you jump straight to any cell or named range by typing its address

**Quick practice (3 min):** Open a blank sheet, use the Name Box to jump to `D20`, type a value, then use `Ctrl+Home` to return to `A1`.

### Shortcut cheat sheet (LibreOffice Calc)

| Shortcut                | Action                                                                                        |
| ------------------------ | ---------------------------------------------------------------------------------------------- |
| `Ctrl+C / V / X`        | Copy / Paste / Cut                                                                              |
| `Ctrl+Z / Y`            | Undo / Redo                                                                                     |
| `Ctrl+Arrow`            | Jump to edge of data                                                                            |
| `Ctrl+Shift+Arrow`      | Select to edge of data                                                                          |
| `Ctrl+Home / End`       | Go to A1 / last used cell                                                                       |
| `Ctrl+1`                | Format Cells dialog                                                                             |
| Formula Bar **Σ** button | Insert AutoSum — Excel's `Alt+=` isn't reliably bound in Calc                                  |
| `F2`                    | Edit active cell                                                                                |
| `Shift+F4`              | Toggle reference type ($ locking) — **plain `F4` opens the Data Sources window in Calc instead** |
| Data → AutoFilter        | Toggle column filter dropdown arrows — no default shortcut, unlike Excel's `Ctrl+Shift+L`      |
| `Ctrl+F / H`            | Find / Replace                                                                                  |

> **Note for this copy:** Excel's `Ctrl+T` "Table" shortcut has no dependable Calc equivalent — see Section 3 for how this session builds the same result without it.

## 3. Formatting Data & Building Tables (0:30 – 0:50)

Formatting is about making numbers readable and structure explicit — not decoration.

- **Number formats** — Currency, Percentage, Comma, Date — applied via `Ctrl+1` (Format Cells), never by typing symbols into the cell
- **Borders, fill colour, bold headers** — used sparingly to signal structure, not to decorate
- **Merge & Center** — and why to avoid it once a range will later be sorted or filtered
- **Table-like structure in Calc** — LibreOffice does have its own Table object (**Insert → Table**), but its behaviour varies release to release, and its structured references (`[@ColumnName]`) are exactly what threw the `Err:507` seen while testing this dataset. This trainer copy gets the same practical result from three plain steps instead: **Format → AutoFormat Styles** for banded rows, **Data → AutoFilter** for header dropdown arrows, and **View → Freeze Rows and Columns** to pin the header — no Table object, no structured-reference surprises.
- **Freeze Rows and Columns** (View menu) — keeping header rows visible while scrolling a long dataset; select the cell just below/right of what should stay frozen, then apply

**Quick practice (6 min):** Take a raw list of 15 expense rows, apply currency formatting to the amount column, bold and colour-fill the header row, then apply AutoFormat banding, turn on AutoFilter, and freeze the header row.

## 4. Basic Formulas (0:50 – 1:05)

Every formula starts with `=`, combines cell references with operators (`+ − * / ^`), and follows standard order of operations (BODMAS). Introduce the five workhorse functions with a small revenue-by-region table:

`SUM(range)` · `AVERAGE(range)` · `COUNT(range)` · `COUNTA(range)` · `MIN(range)` · `MAX(range)`

```
=SUM(C2:C25)      → total revenue
=AVERAGE(C2:C25)  → average deal size
```

Show the **Σ (Sum) button** in the Formula Bar for one-click AutoSum, and the fill handle (drag the small square at the bottom-right of a cell) to copy a formula down a column — this sets up the reference discussion right after the break.

**Quick practice (5 min):** On the revenue table, calculate Total Revenue, Average Deal Size, Number of Deals, and the Largest and Smallest deal.

---

### Break (1:05 – 1:15, 10 min)

---

## 5. Relative, Absolute & Mixed References (1:15 – 1:35)

This is the concept that most breaks copied formulas in real reports — worth slowing down for.

- **Relative** (`A1`) — the default; shifts when the formula is copied to a new cell
- **Absolute** (`$A$1`) — locked on both row and column; stays fixed no matter where the formula is copied. In Calc, press **`Shift+F4`** to cycle through reference types while editing a formula (plain `F4` opens the Data Sources window instead — this trips up anyone used to Excel)
- **Mixed** (`$A1` or `A$1`) — locks only the column or only the row

```
=B2*$B$1   // B2 (sale amount) shifts down the column; $B$1 (fixed commission rate) never moves
```

**Quick practice (8 min):** Build a 10-row price list where every row multiplies its own price by one fixed GST rate cell (e.g. `$B$1 = 18%`). Copy the formula down and confirm the rate cell never shifts.

### Walking through it on the capstone sheet

In this copy the raw import stays a plain range, columns A–M (no Table object, per the note in Section 3):

| Col | A | B | C | D | E | F | G | H | I | J | K | L | M |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Field | Order ID | Customer | Region | Product | Category | Qty | Unit Price | Discount % | Order Date | Sales Rep | Repeat Customer | Payment Mode | Order Status |

Every calculated column the capstone adds lives from **N onward** — Total Sales in N, Net Sales in O, Regional Target in P, Commission in Q, Performance Flag in R, High Value Repeat in S. This is exactly where reference type stops being academic and starts deciding whether a formula survives being copied down 210 rows.

**Relative reference — the default, and usually what you want.** In **N2** write:

```
=F2*G2        (Qty × Unit Price)
```

Copy it down, and Calc rewrites it for you: row 3 becomes `=F3*G3`, row 4 becomes `=F4*G4`, and so on. Neither `F` nor `2` is locked — both track the row the formula lives in. This is correct here because every row needs its own Qty and its own Unit Price.

**Absolute reference — for one fixed value used by every row.** Add a GST cell clearly outside the dataset's own columns — the sheet already uses A–M for raw data and N–S for calculated columns, so anything from **column U onward** is genuinely free — say `U1 = 18%`. Every row's Tax should use *that same cell*, never its own neighbour:

```
=N2*$U$1
```

The `$` before both the column letter and the row number locks it completely. Copied down 210 rows, every single one still points at `U1`. Drop the `$` signs and write `=N2*U1` instead, and copying it down turns row 3's formula into `U2` (empty), row 4's into `U3` (empty) — Tax silently becomes 0 for almost every row. This exact mistake is the single most common way a "working" formula quietly breaks the moment someone copies it. The same logic applies to the Commission Rate and Regional Target values from `region_targets.csv`, if typed as fixed constants rather than the nested `IF` used in the capstone below. **Always check which columns your own data already occupies before picking a "free" cell** — a stray fixed-rate cell placed inside the live data range is a common self-inflicted version of this exact bug.

**Mixed reference — locks only one axis.** `$U1` always reads column U, whatever row you copy to. `U$1` always reads row 1, whatever column you copy to. These rarely come up in a straight down-the-column fill like Total Sales — relative or fully absolute covers that. They earn their keep when a formula gets copied in **two directions at once**, across *and* down — for instance, a small GST-impact grid built somewhere clearly empty (columns A–S are already spoken for, so start this one at, say, column W) as a what-if check:

|   | 5% | 12% | 18% | 28% |
|---|---|---|---|---|
| **4,200** |  |  |  |  |
| **9,500** |  |  |  |  |
| **18,500** |  |  |  |  |

Say the amounts sit in `W3:W5` and the rates sit across `X2:AA2`. The formula in the top-left cell (`X3`) needs to always read *this row's* amount from column W, and always read *this column's* rate from row 2 — no matter which direction it's dragged:

```
=$W3*X$2
```

- `$W3` — column locked to W (dragging right never leaves the amount column), row free (dragging down moves to the next amount)
- `X$2` — row locked to 2 (dragging down never leaves the rate row), column free (dragging right moves to the next rate)

Fill that one formula both right and down, and all 12 cells compute correctly from a single entry — that's the entire reason mixed references exist.

**The one-line rule to leave students with:** if a value should change with the row you're copying to, leave it relative; if it should never change no matter where the formula is copied, lock it with `$`; if you're filling a grid rather than a column, work out which axis is fixed for *that* cell and lock only that one.

## 6. Decision Logic: IF, AND, OR (1:35 – 2:00)

Introduce `IF` as Excel's way of encoding a business rule directly into a cell:

```
=IF(logical_test, value_if_true, value_if_false)
=IF(Sales>Target,"Bonus Eligible","Not Eligible")
```

Then combine conditions with `AND` (all must be true) and `OR` (any one true):

```
=IF(AND(SalesValue>50000,RepeatCustomer="Yes"),"High Value","Standard")
=IF(OR(Region="North",Region="East"),"Zone A","Zone B")
```

Then go one level further: **nested IF**, for when a rule has more than two outcomes. Each `IF` becomes the "else" of the one before it — this is the technique the capstone uses to assign a different target to every region without a lookup function:

```
=IF(Region="North",180000,IF(Region="South",150000,IF(Region="East",200000,165000)))
```

Read it aloud as a decision tree: "if North, 180000 — otherwise, if South, 150000 — otherwise, if East, 200000 — otherwise (must be West), 165000." The last branch is the catch-all, so it never needs its own `IF`.

**Quick practice (8 min):** On a stock sheet, flag `"Reorder"` where Quantity is below Reorder Level, then flag `"Priority Reorder"` where Quantity is below Reorder Level *and* the item is marked Fast-Moving.

## 7. Sorting & Filtering (2:00 – 2:15)

- **Single-column sort** — ascending/descending on any column, including inside a filtered range
- **Multi-level sort** — Data → Sort → Add sort key, e.g. sort by Region, then by Sales descending within each region
- **AutoFilter** (**Data → AutoFilter** — Calc has no dependable default shortcut for this, unlike Excel's `Ctrl+Shift+L`) — filter by value, text contains, date range, or a custom condition (e.g. Sales > 10,000); use **Data → More Filters → Standard Filter** for a "Top 10"-style numeric condition, since Calc has no built-in Top 10 filter button
- **Compound filters** — filtering on more than one column at once (e.g. Region = North *and* Status = Delivered) to answer a question that a single filter can't
- Clearing filters and re-checking row counts before drawing conclusions from a filtered view — a filtered `COUNT` that looks too low is usually a sign a filter was left on somewhere else

**Quick practice (5 min):** Sort the revenue table by Region, then by Sales descending within each region, then filter to show only rows where Sales exceeds a chosen threshold *and* the deal is still open.

## 8. Basic Data Cleaning (2:15 – 2:30)

Real business exports are rarely clean. Name the usual suspects — extra spaces, inconsistent capitalisation, duplicate rows, blank cells, numbers and dates stored as text — then the fixes:

`TRIM(text)` · `PROPER / UPPER / LOWER(text)` · `Text to Columns` · `Find & Replace` · `VALUE(text)` · `DATEVALUE(text)`

```
=TRIM(PROPER(A2))   // strips extra spaces, capitalises each word
```

Calc has no one-click **Remove Duplicates** button the way Excel does. The direct equivalent is **Data → More Filters → Standard Filter**, ticking **"No duplications"** under Options and setting "Copy results to" a new range — this produces a de-duplicated copy without disturbing the original. Show this, and **Data → Text to Columns** (splitting a "First Last" name field, or a comma-separated import — same name, same place as Excel), live on the sample sheet that feeds directly into the capstone.

Flag one sequencing trap explicitly, since it is what separates a rushed cleanup from a correct one: **the "No duplications" filter only catches rows that are already identical.** Two rows for the same order will survive it if one customer name is `"rahul mehta"` and the other is `"Rahul Mehta "` — clean the text columns *first*, then de-duplicate, never the other way round.

## 9. Capstone: Clean the Dataset, Build the Report (2:30 – 2:55)

### Dataset: `retail_sales_q1_raw.csv` (+ `region_targets.csv`)

A 210-row raw sales export, saved alongside this file, with real-world flaws seeded throughout — deliberately messier than the mid-session practice sheets, so the capstone earns its place as the hardest 25 minutes of the day. In this copy it stays a plain range (no Table object, per Section 3), with columns:

`A Order ID · B Customer · C Region · D Product · E Category · F Qty · G Unit Price · H Discount % · I Order Date · J Sales Rep · K Repeat Customer · L Payment Mode · M Order Status`

`region_targets.csv` — a small reference table (Region, Monthly Target, Commission Rate) used to drive the nested-IF steps below, since this cohort hasn't been taught lookup functions yet.

### Data profile (seeded issues, exact counts for this file)

| Issue                                       | Count                                                                | Where                                                                                                 |
| ------------------------------------------- | --------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------ |
| Exact duplicate rows                        | 13                                                                    | fully identical rows                                                                                    |
| Near-duplicate rows                         | 6                                                                      | same order, Customer text styled differently — survive de-duplication until text is cleaned first     |
| Blank Region                                | 4                                                                      | Region column                                                                                          |
| Inconsistent Region spelling/casing/spacing | throughout                                                            | `North`, `north`, `NORTH`, `North `, `  North` (and similarly for South/East/West)                    |
| Inconsistent Customer casing/spacing        | throughout                                                            | upper, lower, leading/trailing spaces                                                                  |
| Unit Price stored as text                   | 38                                                                     | e.g. `Rs.4200`, `₹ 4,200`, `9,600.00`                                                                 |
| Blank Qty                                   | 7                                                                      | Qty column                                                                                              |
| Zero or negative Qty                        | 10                                                                     | returns/cancellations, not typing errors                                                               |
| Order Date in a non-standard format         | 60                                                                     | `DD/MM/YYYY`, `YYYY/MM/DD`, `5 Jan 2026`, or blank                                                    |
| Order Status breakdown                      | 135 Delivered · 27 Cancelled · 36 Returned · 12 Pending               | decides what counts as revenue                                                                          |

*Counts verified directly against the saved `retail_sales_q1_raw.csv` — 210 rows, 191 distinct Order IDs, 19 extra rows from the duplicate/near-duplicate seeding.*

Hand this table out **after** the exercise, not before — it's the answer key for "did I actually catch everything," not a checklist to follow mechanically.

### Steps (LibreOffice Calc)

1. Import the CSV. Apply **Format → AutoFormat Styles** for banding, **Data → AutoFilter** for header dropdowns, and freeze row 1 (**View → Freeze Rows and Columns**). Format column G as currency and column I as a date once both are genuinely numeric/date-typed (see step 4).
2. Clean **Customer** (column B): in a helper column, say `U2`, write `=TRIM(PROPER(B2))`, fill down, then **Copy → Paste Special → Values Only** back over column B, and delete the helper column. Standardise **Region** (column C) with Find & Replace (`Ctrl+H`) so every casing/spacing variant becomes one value, and decide a rule for the blank Regions (e.g. mark `"Unknown"` rather than guessing).
3. Fix the text-stored prices in column G (`Rs.4200`, `₹ 4,200`, `9,600.00`) with Find & Replace to strip currency symbols and thousands separators, then confirm the column is numeric (right-aligned, and `SUM` on it actually returns a number, not 0).
4. Standardise **Order Date** (column I): entries Calc doesn't auto-recognise (`5 Jan 2026`, `YYYY/MM/DD`) can be converted with a helper column using `=DATEVALUE(I2)`, then **Paste Special → Values Only** back over column I and format it as a date. Mixed formats won't sort or filter correctly until this is done.
5. Decide a rule for the zero/negative Qty rows in column F (these are returns, not typos) and for the blank Qty rows (`=COUNTBLANK(F2:F211)` to count them); document the rule in a note cell rather than silently deleting rows.
6. De-duplicate *now*, after the text is clean: **Data → More Filters → Standard Filter**, tick **"No duplications"**, copy the result to a new range — this is the step that catches the 11 near-duplicates as well as the 8 exact ones; doing it before step 2 would miss the near-duplicates entirely.
7. Add **Total Sales** in column N: `=F2*G2` (relative reference, filled down).
8. Add **Net Sales** in column O, applying each row's own **Discount %**: `=N2*(1-H2)` — same formula copied down, discount rate changes with each row while the formula stays identical.
9. Using `region_targets.csv`, add **Regional Target** in column P with a nested `IF` on Region: `=IF(C2="North",180000,IF(C2="South",150000,IF(C2="East",200000,165000)))`, and **Commission** in column Q the same way: `=O2*IF(C2="North",0.06,IF(C2="South",0.05,IF(C2="East",0.065,0.055)))`.
10. Add **Performance Flag** in column R: `=IF(O2>=P2,"Achieved","Below Target")` — each row is now judged against *its own* region's target, not one flat number.
11. Add **High Value Repeat** in column S: `=IF(AND(O2>10000,K2="Yes"),"High Value","Standard")`.
12. Filter out **Cancelled** orders (column M) before computing any revenue KPI — a cancelled order is not revenue, even though it has a Unit Price and Qty like any other row.
13. Sort by **Region** (C), then by **Net Sales** (O) descending; filter to isolate one region, one Sales Rep, or one Order Status at a time.
14. Below the range, build a KPI summary excluding Cancelled orders: Total Net Revenue (`SUM`), Average Order Value (`AVERAGE`), Number of Orders (`COUNT`), Highest and Lowest Sale (`MAX`/`MIN`), and how many are `Achieved` vs `Below Target`.

**Deliverable:** One cleaned, formatted range plus a KPI summary — the same shape as a one-page MIS report handed to a manager.

### Practical exercises students can run on this dataset

Use these as the guided capstone above, as extra reps for faster finishers, or as homework:

**Formatting**

- Apply currency format to columns G, N and O; short-date format to column I
- Freeze row 1; apply AutoFormat banding and AutoFilter
- Add a totals row underneath the range with `SUM` for columns N and F

**Formulas & references**

- Total Sales = `F2*G2` (relative reference)
- Net Sales = `N2*(1-H2)` (relative reference to a per-row rate)
- Commission = Net Sales × region-specific rate assigned via nested IF (constants baked into the formula, no lookup function)
- Variance to Target = `O2-P2`, to see not just pass/fail but by how much

**Logic**

- Regional Target and Commission Rate via nested `IF` on Region (4-way branch)
- Performance Flag: `Achieved` / `Below Target` against each row's own regional target
- High Value Repeat: `=IF(AND(O2>10000,K2="Yes"),"High Value","Standard")`
- Priority Region: `=IF(OR(C2="North",C2="East"),"Zone A","Zone B")`
- Order health flag: `=IF(OR(F2<=0,M2="Cancelled"),"Exclude","Include")` — one flag that catches two different kinds of bad row

**Sorting & filtering**

- Sort by Region, then Net Sales descending
- Compound filter: one Region *and* Order Status = Delivered
- Standard Filter with a numeric condition on Net Sales to isolate the highest-value delivered orders (Calc has no one-click Top 10 button)
- Filter to only `Below Target` *and* `Repeat Customer = Yes` — the customers most worth a follow-up call

**Data cleaning**

- `TRIM` + `PROPER` on Customer, *before* de-duplicating
- Find & Replace to standardise Region spelling and fix currency-symbol prices
- `DATEVALUE` to standardise the three different Order Date formats
- `COUNTBLANK` for missing Qty, and a documented rule for zero/negative Qty
- Standard Filter "No duplications" — and checking the near-duplicates were caught too
- `Text to Columns` to split Customer into First Name / Last Name

**Summary report (the capstone deliverable)**

- Total Net Revenue, Average Order Value, Order Count — Cancelled orders excluded (`SUM`, `AVERAGE`, `COUNT`)
- Highest and Lowest single order (`MAX`, `MIN`)
- Count of `Achieved` vs `Below Target` orders, region by region
- Net Revenue by Region (one line per region, computed by filtering + `SUM`)
- Top Sales Rep by Net Revenue, and by Commission earned

## 10. Recap & Exit Ticket (2:55 – 3:00)

**Three-question exit ticket:**

1. Write the absolute reference for cell `B2`, then the mixed reference that locks only its column.
2. Write one formula using `IF` and `AND` that marks a customer `"High Value"` when Order Value exceeds 50,000 *and* they are a repeat customer.
3. Write a nested `IF` that returns `"Gold"` for Sales above 100,000, `"Silver"` above 50,000, and `"Bronze"` otherwise.

**Looking ahead:** Today's foundation — references, logic, clean data — is what lookup functions, pivot tables and dashboards are built on top of in the following sessions. Everything from here reuses the same range and the same habits.

---

**Trainer prep:** Share `retail_sales_q1_raw.csv` and `region_targets.csv` with each student (or pre-load both as tabs in a starter workbook); a projector with the Ribbon visible at readable zoom; keep a cleaned reference copy on hand for before/after comparison at the open, and hold back the Data Profile table until after the capstone so it lands as an answer key, not a shortcut. **This copy is LibreOffice Calc throughout** — if you're demoing live on classroom machines running Microsoft Excel instead, switch to the mechanics in `day_1_student_handout.md` (real `Ctrl+T` Tables, `[@ColumnName]` structured references, `F4`, Excel's native Remove Duplicates) so what's on the projector matches what students see on their own screens.
