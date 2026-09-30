# Pizza Sales Analysis Dashboard — Power BI

> **Repository contents**
> - `PizzaSales.pbip`, `PizzaSales.Report/`, `PizzaSales.SemanticModel/` – Power BI project: semantic model (`model.bim`) + 3-page report
> - Report theme (navy / white / gray, tomato accent): `PizzaSales.Report/StaticResources/RegisteredResources/PizzaSales_Theme.json`
> - This README – cleaning log, data model, DAX, insights, interview prep
>
> **Open the Power BI project:** enable *File → Options → Preview features → Power BI Project (.pbip) save option*, open `PizzaSales.pbip`, then update the `stg_pizza_sales` source path to your local copy of the Excel file and click **Refresh**. The source Excel file is not included in this repository.
> - Project files were generated programmatically and were built to open in Power BI Desktop; report pages may need minor visual adjustments.
> - All figures in this document were queried from the model (calendar year 2015).

---

## 1. Business Problem

A pizza restaurant records every pizza line sold (48,620 lines, 21,350 orders, calendar year 2015) but has no consolidated view of performance. Management needs to answer:

- How much revenue are we making, and is it growing or shrinking month to month?
- Which pizzas, categories and sizes drive revenue and volume, and which underperform?
- When do customers order (hour, weekday, meal period) so staffing and prep can be planned?
- How concentrated is revenue in a few products?

## 2. Objectives

1. Analyse revenue performance and monthly trend
2. Identify top and bottom pizzas by revenue and by quantity
3. Understand ordering patterns (hour, day, meal period, weekday vs weekend)
4. Compare categories and sizes
5. Quantify revenue concentration (Top 10 share)
6. Deliver an interactive, decision-ready dashboard on a validated data model

## 3. Tools Used

Microsoft Excel (source) · Power Query (M) · Power BI Desktop · DAX · Star-schema modelling

---

## 4. Data Cleaning (Power Query)

**Source:** `pizza_sales_excel_file.xlsx`, sheet `pizza_sales`.
**Approach:** one **staging query** `stg_pizza_sales` (load disabled) does all cleaning once. The fact table and every dimension reference it, so cleaning logic lives in one place.

### 4.1 Steps in `stg_pizza_sales`

| # | Step | Why |
|---|------|-----|
| 1 | Promote headers | First row holds column names |
| 2 | Set data types (per your spec) | Correct aggregation, sorting and relationships |
| 3 | `order_time` changed from **DateTime → Time** | The original load had it as datetime; it's a time-of-day value |
| 4 | `Text.Trim` + `Text.Clean` on `pizza_name_id`, `pizza_name`, `pizza_ingredients` | Removes stray spaces / non-printable characters that break joins |
| 5 | `pizza_size` → UPPER, `pizza_category` → Proper case | Prevents "l"/"L" or "classic"/"Classic" splitting into two members |
| 6 | `Table.Distinct` (exact duplicate rows) | Removes only fully identical rows |

### 4.2 Data-quality audit (what I actually checked)

| Check | Result |
|-------|--------|
| Row count | **48,620** in → **48,620** out |
| Exact duplicate rows | **0** |
| Duplicate `pizza_id` | **0** (unique key) |
| Blank values, all 12 columns | **0** |
| Empty-string size / category / name | **0** |
| Text needing trim (name, size, category) | **0** |
| Quantity ≤ 0, unit_price ≤ 0, total_price ≤ 0 | **0** |
| `total_price ≠ quantity × unit_price` (tolerance $0.01) | **0** |
| Invalid dates | **0**; range 01-Jan-2015 to 31-Dec-2015 |
| Sizes present | S, M, L, XL, XXL (no variants) |
| Categories present | Chicken, Classic, Supreme, Veggie (no variants) |
| `pizza_name_id` → name/category/ingredients | 91 IDs = 91 distinct combinations (clean 1:1) |
| Orphan keys (fact rows with no matching dimension row) | **0** |
| Outliers | Max quantity per line = 4 (927 lines have quantity > 1); unit price $9.75–$35.95. The $35.95 is the XXL Greek pizza, a legitimate menu price, **so it was kept.** |

### 4.3 Audit statement

**No records were removed, altered or imputed.** The standardisation steps (trim, case) are safeguards that changed zero values on this file. The only material change is the `order_time` data type. Validation is preserved permanently in the model as DQ measures (Section 7.7), so it can be re-run after any refresh.

---

## 5. Data Model

### 5.1 Star schema

```
                DimDate ──┐
                          │ 1
DimPizza ── 1 ──*  FactPizzaSales  *── 1 ── DimSize
                          │ *
                          │
                      1   │
                    DimCategory
```

| Table | Type | Grain | Rows | Built with |
|-------|------|-------|------|-----------|
| FactPizzaSales | Fact | one pizza line | 48,620 | Power Query (from staging) |
| DimDate | Dimension | one day | 365 | DAX calculated table, **marked as date table** |
| DimPizza | Dimension | one `pizza_name_id` (pizza + size) | 91 | Power Query |
| DimSize | Dimension | one size | 5 | Power Query (+ hidden `size_sort`) |
| DimCategory | Dimension | one category | 4 | Power Query |
| _Measures | Measure holder | – | – | Hidden placeholder column |
| TopN Selector | What-if parameter | – | 18 (3–20) | GENERATESERIES |

### 5.2 Relationships (all single-direction, dimension filters fact)

| From (many) | To (one) | Cardinality |
|-------------|----------|-------------|
| FactPizzaSales[order_date] | DimDate[Date] | Many-to-One |
| FactPizzaSales[pizza_name_id] | DimPizza[pizza_name_id] | Many-to-One |
| FactPizzaSales[pizza_size] | DimSize[pizza_size] | Many-to-One |
| FactPizzaSales[pizza_category] | DimCategory[pizza_category] | Many-to-One |

**Design decisions to explain in interviews**
- `pizza_ingredients` was removed from the fact table (it's descriptive text, 48k repeats) and lives only in DimPizza.
- Fact keeps `pizza_size`, `pizza_category`, `pizza_name` per your spec, but they are **hidden**, so slicers always use dimension columns. `DimPizza[pizza_category]` is hidden too, so there is only one category slicer field.
- `DimSize` has a hidden `size_sort` column so sizes sort S → M → L → XL → XXL, not alphabetically. This is one column beyond your spec, added on purpose.
- `DimDate` has extra helper columns: `Day of Week Number` (hidden, sorts Day Name Mon→Sun) and `Day Type`.
- No bidirectional relationships anywhere.

### 5.3 Column classification

| Layer | Objects |
|-------|---------|
| **Source columns** | pizza_id, order_id, pizza_name_id, quantity, order_date, order_time, unit_price, total_price, pizza_size, pizza_category, pizza_ingredients, pizza_name |
| **Power Query transformations** | Types, trim/clean, size & category standardisation, distinct, remove `pizza_ingredients` from fact, build DimPizza/DimSize/DimCategory, `size_sort` |
| **Calculated columns (DAX)** | FactPizzaSales: `Order Hour`, `Meal Period`, `Meal Period Sort` · DimDate: Year, Month Number, Month Name, Quarter, Year-Month, Day, Day Name, Day of Week Number, Week Number, Day Type |
| **DAX measures** | Section 7 |

### 5.4 Meal-period definitions

Based on the actual hourly distribution (orders exist from 09:00 to 23:59, peaking 12:00–13:00 and 17:00–19:00):

| Meal Period | Hours | Rationale |
|-------------|-------|-----------|
| Morning | before 11:00 | Store opens ~09:52; almost no trade |
| Lunch | 11:00–13:59 | Captures the midday peak |
| Afternoon | 14:00–16:59 | The mid-day lull |
| Evening | 17:00 onward | Dinner peak through close |

### 5.5 Time-intelligence caveat

The data covers **one year (2015)**. `Revenue Previous Year` and `Revenue YoY Growth %` are built correctly but return **blank** here. Month-over-month growth works. Don't put YoY on the dashboard; mention it as future-proofing.

---

## 6. Dashboard Build Guide (Power BI Desktop)

### 6.1 Setup
1. **View → Themes → Browse for themes** → load `PizzaSales_Theme.json`.
2. **File → Options → Current File → Data Load** → untick **Auto date/time**. The hidden `LocalDateTable` objects have already been deleted from the model; unticking this option stops Power BI from regenerating them.
3. Canvas: 16:9, 1280×720. Page background `#F3F4F6`. Save as `.pbix`.

### 6.2 Number formats (already set on the measures)
Revenue `$#,##0` · AOV / prices `$#,##0.00` · Orders & quantity `#,##0` · Percent `0.0%` · Pizzas per order `0.00`.

### 6.3 Page 1 — "PIZZA SALES PERFORMANCE DASHBOARD"

| Visual | Fields |
|--------|--------|
| 5 KPI cards | `Total Revenue`, `Total Orders`, `Total Pizzas Sold`, `Average Order Value`, `Average Pizzas Per Order` |
| Line chart – Revenue trend | X: `DimDate[Year-Month]` · Y: `Total Revenue` |
| Donut – Sales by category | Legend: `DimCategory[pizza_category]` · Values: `Total Revenue` |
| Column – Sales by size | X: `DimSize[pizza_size]` · Y: `Total Revenue` |
| Bar – Top 10 pizzas | Y: `DimPizza[pizza_name]` · X: `Total Revenue` · Filters pane → Top N → 10 by `Total Revenue` |
| Column – Orders by hour | X: `FactPizzaSales[Order Hour]` · Y: `Total Orders` |
| Slicers | Date (between) on `DimDate[Date]`; `DimCategory[pizza_category]`; `DimSize[pizza_size]` |
| Nav buttons | Insert → Buttons → Navigator → Page navigator; plus Home button |

### 6.4 Page 2 — "PIZZA PRODUCT & CATEGORY ANALYSIS"

| Visual | Fields |
|--------|--------|
| Top 10 by revenue (bar) | `pizza_name` · `Total Revenue` · Top N 10 |
| Top 10 by quantity (bar) | `pizza_name` · `Total Pizzas Sold` · Top N 10 |
| Bottom 10 by revenue (bar) | `pizza_name` · `Total Revenue` · **Bottom** N 10 |
| Revenue by category (column) | `pizza_category` · `Total Revenue` |
| Quantity by category (column) | `pizza_category` · `Total Pizzas Sold` |
| Revenue by size (column) | `pizza_size` · `Total Revenue` |
| Performance matrix | Rows: `pizza_name` · Values: `Total Revenue`, `Total Pizzas Sold`, `Total Orders`, `Average Pizza Price`, `Revenue Contribution %` |
| Conditional formatting | Revenue: data bars · Contribution %: background colour gradient (light → `#E4572E`) |
| **Top N selector** | Slicer on `TopN Selector[Value]` (single-select, dropdown). On each Top-N bar visual: Filters → drag `Show In Top N` → "is 1" (replaces the fixed Top N filter) |
| Slicers | Category, Size, **Pizza name** (`DimPizza[pizza_name]`, dropdown) |

### 6.5 Page 3 — "ORDER & TIME PATTERN ANALYSIS"

| Visual | Fields |
|--------|--------|
| Revenue by month (line) | `DimDate[Month Name]` · `Total Revenue` |
| Orders by hour (column) | `Order Hour` · `Total Orders` |
| Revenue by day of week (column) | `DimDate[Day Name]` · `Total Revenue` |
| Orders by day of week (column) | `DimDate[Day Name]` · `Total Orders` |
| Revenue by meal period (donut) | `Meal Period` · `Total Revenue` |
| Weekday vs weekend (column) | `DimDate[Day Type]` · `Total Revenue` |
| **Heatmap** (matrix) | Rows: `DimDate[Day Name]` · Columns: `FactPizzaSales[Order Hour]` · Values: `Total Orders` · Conditional formatting → Background colour → Gradient (min `#EEF2F7`, max `#E4572E`) · turn off row/column subtotals |
| Peak Hour KPI | Card: `Peak Order Hour` (shows as `12:00`) · second card: `Peak Order Day` |

### 6.6 Interactivity
- **Visual interactions:** Format → Edit interactions. Set the Top-10 visuals to *not* cross-filter each other.
- **Drill-through:** create hidden page **"Pizza Detail"**; drop `DimPizza[pizza_name]` in *Drill-through fields*. Put KPI cards + revenue by size + revenue by month on it. Right-click any pizza on Pages 1–2 → Drill through.
- **Tooltip pages:** create page **"Tooltip"** → Page format → *Allow use as tooltip* (size: Tooltip). Add cards: `Total Revenue`, `Total Orders`, `Total Pizzas Sold`, `Average Order Value`, `Revenue Growth %` (MoM). On chart visuals: Format → General → Tooltips → Type *Report page* → choose "Tooltip".
- **Navigation:** page navigator on every page + a Home button (Insert → Buttons → Blank; Action → Page navigation → Page 1).

### 6.7 UI/UX rules applied
Navy `#0B2545` title bar, white cards on `#F3F4F6`, **one** accent colour (tomato `#E4572E`) reserved for highlights, Segoe UI throughout, 8-px grid spacing, card corner radius 6, titles on every visual, no gridline clutter, category/size charts sorted logically (sizes S→XXL, weekdays Mon→Sun, meal periods Morning→Evening).

---

## 7. DAX Measures (all live in table `_Measures`)

### 7.1 Sales KPIs
| Measure | Definition |
|---------|-----------|
| Total Revenue | `SUM ( FactPizzaSales[total_price] )` |
| Total Orders | `DISTINCTCOUNT ( FactPizzaSales[order_id] )` |
| Total Pizzas Sold | `SUM ( FactPizzaSales[quantity] )` |
| Average Order Value | `DIVIDE ( [Total Revenue], [Total Orders] )` |
| Average Pizza Price | `AVERAGE ( FactPizzaSales[unit_price] )` (menu-price level) |
| Average Pizzas Per Order | `DIVIDE ( [Total Pizzas Sold], [Total Orders] )` |
| Revenue Per Pizza | `DIVIDE ( [Total Revenue], [Total Pizzas Sold] )` |

### 7.2 Time intelligence
| Measure | Definition |
|---------|-----------|
| Revenue Previous Month | `CALCULATE ( [Total Revenue], DATEADD ( DimDate[Date], -1, MONTH ) )` |
| Revenue Previous Year | `CALCULATE ( [Total Revenue], SAMEPERIODLASTYEAR ( DimDate[Date] ) )` |
| Revenue Growth % | MoM: `VAR Cur = [Total Revenue] VAR Prev = [Revenue Previous Month] RETURN IF ( NOT ISBLANK ( Prev ), DIVIDE ( Cur - Prev, Prev ) )` |
| Revenue YoY Growth % | Same pattern on Previous Year (blank: 1 year of data) |
| Monthly Revenue | `CALCULATE ( [Total Revenue], ALLEXCEPT ( DimDate, Year, Quarter, Month Number, Month Name, Year-Month ) )` |
| Yearly Revenue | `CALCULATE ( [Total Revenue], ALLEXCEPT ( DimDate, DimDate[Year] ) )` |
| Cumulative Revenue | Running total using `FILTER ( ALLSELECTED ( DimDate[Date] ), DimDate[Date] <= MAX ( DimDate[Date] ) )`; respects the date slicer |
| YTD Revenue / MTD Revenue | `TOTALYTD` / `TOTALMTD` |

### 7.3 Product analytics
| Measure | Logic |
|---------|-------|
| Best Selling Pizza | `TOPN(1)` over pizza names by quantity, non-blank only |
| Worst Selling Pizza | Same, ascending |
| Top Pizza by Revenue | Same, by revenue |
| Best Pizza Category | Category with the highest revenue |

**"Revenue by Pizza Size / Category", "Quantity by Pizza", "Orders by Pizza"** are not separate measures. They are `[Total Revenue]`, `[Total Pizzas Sold]` and `[Total Orders]` placed against the size/category/pizza field in a visual. Duplicating them as identical measures would be clutter. Say this in an interview: it shows you understand filter context.

### 7.4 Order analytics
`Average Items Per Order` (line items per order: `COUNTROWS / Total Orders`, differs from pizzas per order because quantity can exceed 1) · `Orders Per Day` (orders ÷ trading days) · `Revenue Per Order` (= AOV) · `Peak Order Hour` · `Peak Order Day`.

### 7.5 Contribution
`Revenue Contribution %`, `Pizza Quantity Contribution %` (both use `ALLSELECTED()` so they work for any dimension and respect slicers), `Category Contribution %`, `Top 10 Pizzas Revenue %`.

### 7.6 Ranking / Top N
`TopN Value` (slicer choice, default 10) · `Pizza Rank by Revenue` · `Pizza Rank by Quantity` (`RANKX … DENSE`) · `Show In Top N` (1/0 flag for visual-level filter).

### 7.7 Data quality
`DQ Row Count` · `DQ Price Mismatch Rows` · `DQ Invalid Qty or Price Rows` · `DQ Duplicate Pizza IDs`. All currently **0** (row count 48,620).

---

## 8. Validation (results from the model)

| Check | Result | Reconciles? |
|-------|--------|-------------|
| Total Revenue | **$817,860.05** | Equals source `SUM(total_price)` ✔ |
| Total Orders | **21,350** | Equals distinct `order_id` ✔ |
| Total Quantity | **49,574** | Equals source `SUM(quantity)` ✔ |
| Average Order Value | **$38.31** | 817,860.05 ÷ 21,350 ✔ |
| Avg Pizzas Per Order | **2.32** | 49,574 ÷ 21,350 ✔ |
| Revenue by category (sum) | $220,053.10 + $208,197.00 + $195,919.50 + $193,690.45 = $817,860.05 | ✔ |
| Revenue by size (sum) | $178,076.50 + $249,382.25 + $375,318.70 + $14,076.00 + $1,006.60 = $817,860.05 | ✔ |
| Monthly revenue (12 months) | Cumulative Dec = $817,860.05 | ✔ |
| Orders by weekday (sum) | 21,350 | ✔ |
| Orders by meal period (sum) | 21,350 | ✔ |
| Revenue by weekday / meal period | Each sums to $817,860.05 | ✔ |

Tip for your dashboard: put a hidden "Validation" page with `DQ` measures as cards.

---

## 9. Performance Optimisation

| Practice | Why it matters | Applied |
|----------|----------------|---------|
| Star schema | Small dimensions filter one narrow fact table; simpler DAX and faster scans | ✔ |
| Staging query with load disabled | Cleaning runs once; no duplicate logic | ✔ |
| Dedicated, marked date table | Enables time-intelligence functions; avoids hidden auto date tables | ✔ (auto tables deleted; also untick Auto date/time) |
| Removed `pizza_ingredients` from fact | Long text repeated 48k times is the most expensive column to compress | ✔ |
| Hidden fact keys/text columns | Forces users toward dimension fields; cleaner field list | ✔ |
| Measures instead of stored columns | Calculated columns consume memory; measures compute on demand | ✔ (only 3 low-cardinality columns on fact) |
| `summarizeBy = None` on IDs/prices | Stops accidental auto-summing of keys | ✔ |
| Single-direction relationships | Avoids ambiguity and slower filter propagation | ✔ |
| Variables (`VAR`) in DAX | Readability, and each expression is evaluated once | ✔ |
| Avoid high-cardinality visual fields | `order_id` (21k) and `pizza_id` (48k) are hidden | ✔ |
| Display folders & measure table | Findable, maintainable model | ✔ |

---

## 10. Key KPIs Explained

| KPI | Meaning | Value |
|-----|---------|-------|
| Total Revenue | Money earned from all pizzas | $817,860 |
| Total Orders | Distinct customer orders | 21,350 |
| Total Pizzas Sold | Units sold | 49,574 |
| Average Order Value | Spend per order | $38.31 |
| Avg Pizzas per Order | Basket size in pizzas | 2.32 |
| Avg Items per Order | Basket size in line items | 2.28 |
| Revenue per Pizza | Realised price per unit | $16.50 |
| Orders per Day | Average orders per trading day (days with at least one sale; 358 of 365) | 59.6 |

---

## 11. Business Insights (from your data)

1. **Revenue is flat, not growing.** Monthly revenue stays in a narrow band of about $64K–$73K. Peak: **July ($72,558)**; lowest: **October ($64,028)**. Month-over-month swings range from −8.1% (December) to +9.9% (November); there is no clear upward trend across 2015.
2. **Top revenue pizza: The Thai Chicken Pizza — $43,434 (5.3% of revenue).** Next: Barbecue Chicken ($42,768) and California Chicken ($41,410). Chicken pizzas take the top 3 revenue spots.
3. **Top volume pizza: The Classic Deluxe — 2,453 units**, ahead of Barbecue Chicken (2,432) and Hawaiian (2,422). Volume leaders and revenue leaders differ: the Hawaiian is #3 in units but only #8 in revenue, and the Pepperoni is #4 in units but #11 in revenue, because they sell at lower prices.
4. **Revenue is spread across the menu, but the top 10 pizzas still generate 44.5% of revenue** (of 32 pizzas).
5. **Classic is the top category: $220,053 (26.9% of revenue)** and by far the top in volume (14,888 pizzas). Supreme 25.5%, Chicken 24.0%, Veggie 23.7%; the categories are close, so revenue isn't dependent on one. Chicken earns the most per pizza (about $17.73 vs $14.78 for Classic).
6. **Large is the dominant size: 45.9% of revenue** ($375,319), followed by M (30.5%) and S (21.8%). **XL (1.7%) and XXL (0.1%, just 28 pizzas)** barely sell, so they are candidates for menu review.
7. **Underperformers:** The Brie Carre Pizza is the lowest on both revenue ($11,589, 1.4%) and near-lowest on quantity (490). Also weak: Green Garden, Spinach Supreme, Mediterranean and Spinach Pesto (roughly $14K–$16K each).
8. **Two daily peaks: lunch and dinner.** The busiest hour is **12:00 (2,520 orders)**, followed by 13:00 (2,455), 18:00 (2,399) and 17:00 (2,336). Evening (17:00+) is the largest meal period at 45.5% of revenue; Lunch 32.1%; Afternoon 22.3%; Morning is negligible (9 orders in total).
9. **Friday is the busiest day: 3,538 orders, $136,074.** Thursday and Saturday follow ($123.5K and $123.2K). **Sunday is the weakest (2,624 orders, $99,203).** The single busiest hour-slot is Thursday 13:00 (438 orders); Thursday 12:00 has 434.
10. **Weekends are 27.2% of revenue** ($222,386 of $817,860). With two of seven days, that is slightly below an even share, so weekend trade is not stronger than weekday trade.

---

## 12. Resume Project Entry

**PROJECT: Pizza Sales Analysis Dashboard | Power BI**
**Technology:** Power BI, Power Query, DAX, Excel

- Cleaned and profiled 48,620 sales records in Power Query (type correction, text standardisation, duplicate/null/price-consistency checks) using an audit-friendly staging query, with zero records removed and total revenue reconciled to $817,860.
- Designed a star-schema data model (1 fact table, 4 dimensions incl. a marked date table) and built 40+ DAX measures covering revenue KPIs, MoM growth, YTD/MTD, ranking and contribution analysis.
- Built an interactive 3-page dashboard (executive overview, product analysis, order/time patterns) with Top-N selector, drill-through, report-page tooltips and an hour × weekday heatmap.
- Identified that the top 10 of 32 pizzas generate 44.5% of $818K revenue, that Large pizzas drive 46% of revenue, and that orders peak at 12:00 and 17:00–18:00 with Friday the busiest day.
- Added data-quality measures (price mismatch, invalid values, duplicate IDs) to re-validate data after every refresh.

---

## 13. Your 60-Second Answer

> "I built a Power BI sales dashboard for a pizza restaurant using about 48,000 order lines from Excel. In Power Query I set data types, standardised text and checked for duplicates, blanks and whether total price equals quantity times unit price. The data was clean, so I kept every record and documented that. I modelled it as a star schema with a fact table and date, pizza, size and category dimensions, with a marked date table. In DAX I built KPIs like revenue, orders and average order value, plus month-over-month growth, ranking and contribution measures. The dashboard has three pages: an executive overview, product and category analysis, and order timing patterns, with slicers, drill-through and tooltips. Some findings: revenue was about $818K, the top 10 pizzas made 44.5% of it, Large pizzas drove 46%, and orders peak at noon and around 5–6pm, with Friday the busiest day. I validated every total back to the source."

---

## 14. Interview Preparation

### 14.1 Power BI (15)

1. **What is a star schema and why use it?** One central fact table joined to descriptive dimension tables. It gives faster queries, simpler DAX, smaller models and clean filtering.
2. **Fact vs dimension table?** Facts hold measurable events (sales lines); dimensions hold descriptive attributes used to slice them (date, pizza, size).
3. **What relationship cardinality did you use?** Many-to-one from fact to each dimension, single-direction, so dimensions filter the fact.
4. **Why avoid bidirectional filters?** They create ambiguity, can slow queries, and can cause unexpected results. Use them only when necessary.
5. **Import vs DirectQuery?** Import loads data into memory (fast, scheduled refresh); DirectQuery queries the source live (fresh data, slower, limited DAX). I used Import.
6. **Calculated column vs measure?** A column is computed at refresh, stored per row, and usable in slicers/rows. A measure is computed at query time on the current filter context and takes no storage.
7. **What is filter context?** The set of filters (slicers, visual axes, page/report filters) applied to a calculation at a given cell.
8. **Why a dedicated date table?** Time-intelligence functions need a continuous, marked date table; auto date/time creates hidden bloated tables.
9. **What is row-level security?** Rules that restrict which rows a user sees, defined as DAX filters on roles.
10. **What is drill-through?** A detail page filtered by the field you right-click, e.g. drill through from any pizza to its detail page.
11. **What are tooltip pages?** Report pages used as custom hover tooltips to show extra measures without cluttering the visual.
12. **How do you optimise a slow report?** Star schema, remove unused columns, reduce high-cardinality columns, prefer measures, limit visuals per page, use Performance Analyzer.
13. **What is Performance Analyzer?** A Desktop tool that records how long each visual's DAX query and rendering take.
14. **Bookmarks vs buttons?** Bookmarks save a report state (filters, visibility); buttons trigger actions like navigation or applying a bookmark.
15. **How would you refresh this report in the Service?** Publish, use a gateway for the local Excel file (or move the file to OneDrive/SharePoint), and schedule refresh.

### 14.2 DAX (10)

1. **CALCULATE, what does it do?** Evaluates an expression after modifying the filter context. Nearly all advanced DAX uses it.
2. **SUM vs SUMX?** SUM aggregates one column; SUMX iterates rows evaluating an expression per row, then sums, needed for e.g. `quantity * price`.
3. **ALL vs ALLSELECTED vs ALLEXCEPT?** ALL removes all filters; ALLSELECTED removes visual filters but keeps slicer/outer filters (used for % of total); ALLEXCEPT removes all except the listed columns.
4. **What is context transition?** When a row context becomes a filter context, which happens when CALCULATE or a measure is evaluated inside an iterator.
5. **How did you build MoM growth?** `DATEADD(…,-1,MONTH)` for the previous month, then `DIVIDE(Current - Previous, Previous)` with variables, blank when no previous period.
6. **Why DIVIDE instead of `/`?** DIVIDE handles divide-by-zero safely and returns blank (or an alternate value).
7. **What are variables in DAX?** `VAR … RETURN` stores results once; improves readability and performance and avoids repeat evaluation.
8. **How does your Top N slicer work?** A GENERATESERIES parameter table feeds `TopN Value`; `RANKX` computes rank; `Show In Top N` returns 1/0 and is used as a visual filter.
9. **How did you compute Best Selling Pizza?** `TOPN(1)` over `VALUES(pizza_name)` with a quantity column via `ADDCOLUMNS`, then `CONCATENATEX` to return the name.
10. **TOTALYTD vs a manual YTD?** TOTALYTD is shorthand for CALCULATE with DATESYTD; it needs a proper date table.

### 14.3 SQL on this dataset (10) — table `pizza_sales`

1. **Total revenue?** `SELECT SUM(total_price) FROM pizza_sales;` → 817,860.05
2. **Total orders?** `SELECT COUNT(DISTINCT order_id) FROM pizza_sales;` → 21,350
3. **Average order value?** `SELECT SUM(total_price)/COUNT(DISTINCT order_id) FROM pizza_sales;` → 38.31
4. **Revenue by category?** `SELECT pizza_category, SUM(total_price) rev FROM pizza_sales GROUP BY pizza_category ORDER BY rev DESC;`
5. **Top 5 pizzas by revenue?** `SELECT pizza_name, SUM(total_price) rev FROM pizza_sales GROUP BY pizza_name ORDER BY rev DESC LIMIT 5;` (`TOP 5` in SQL Server)
6. **Monthly revenue?** `SELECT DATE_FORMAT(order_date,'%Y-%m') ym, SUM(total_price) FROM pizza_sales GROUP BY ym ORDER BY ym;` (`FORMAT`/`DATEPART` in SQL Server)
7. **Orders by hour?** `SELECT HOUR(order_time) h, COUNT(DISTINCT order_id) FROM pizza_sales GROUP BY h ORDER BY h;`
8. **% revenue by size?** `SELECT pizza_size, SUM(total_price)*100.0/SUM(SUM(total_price)) OVER() FROM pizza_sales GROUP BY pizza_size;`
9. **Find duplicate rows?** `SELECT pizza_id, COUNT(*) FROM pizza_sales GROUP BY pizza_id HAVING COUNT(*)>1;`
10. **Rows where price doesn't reconcile?** `SELECT * FROM pizza_sales WHERE ABS(quantity*unit_price - total_price) > 0.01;`

### 14.4 Power Query / data cleaning (10)

1. **What is Power Query?** A no-code/M-language ETL engine in Power BI for connecting to, cleaning and shaping data before it loads.
2. **What is M?** The functional language behind Power Query steps; each applied step is a line of M.
3. **How did you handle data types?** Set explicitly per column; corrected `order_time` from DateTime to Time.
4. **Trim vs Clean?** Trim removes leading/trailing spaces; Clean removes non-printable characters.
5. **How did you remove duplicates without losing valid data?** `Table.Distinct` only removes *fully identical* rows; I also checked `pizza_id` uniqueness. Zero were found.
6. **What is a staging query?** A load-disabled query holding the cleaned base; other queries reference it so logic isn't duplicated.
7. **Reference vs duplicate query?** Reference reuses the original's output (changes flow through); duplicate copies the steps independently.
8. **What is query folding?** When Power Query pushes transformations back to the source as native queries. Excel doesn't fold, so it doesn't apply to this project, but it matters with SQL sources.
9. **How would you handle missing values?** Depends on meaning: remove, replace with default, fill down, or flag. Never silently drop business records; document the choice.
10. **How do you make data cleaning auditable?** Keep row counts before/after, don't delete suspicious records silently, document each step, and keep DQ measures in the model.

### 14.5 Project-specific (10)

1. **What business problem does it solve?** Gives management one place to see revenue, product mix and ordering patterns for pricing, menu and staffing decisions.
2. **How did you validate accuracy?** Reconciled revenue, orders and quantity to the source, checked that category/size/month/weekday breakdowns each sum to the total, and used DQ measures.
3. **Why distinct count for orders?** An order can have several pizza lines, so counting rows would overstate orders.
4. **Why is Average Pizzas Per Order different from Average Items Per Order?** Pizzas sums quantity (2.32); items counts line rows (2.28), because 927 lines have quantity > 1.
5. **Why is YoY growth blank?** The data covers only 2015, so no prior year exists. The measure is built and will populate with more data.
6. **What was the most important insight?** Top 10 pizzas give 44.5% of revenue, Large gives 46%, and demand peaks at lunch and dinner, so staffing and prep should match those windows.
7. **What would you recommend?** Review XL/XXL and the lowest-revenue pizzas (Brie Carre), staff for 12:00–13:00 and 17:00–19:00, and consider promotions for Sunday and quiet months (Sep–Oct, Dec).
8. **Why did you define meal periods that way?** From the hourly order distribution: two peaks (lunch, dinner) with an afternoon dip, and almost no orders before 11:00.
9. **What limitations does the dataset have?** One year, no customer IDs, no cost/margin data, so no profit, retention or seasonality-over-years analysis.
10. **What would you do next?** Add cost data for margin, more years for YoY, customer IDs for retention/RFM, and publish to the Service with scheduled refresh.

---

## 15. Implementation Checklist (reproduce step-by-step)

**A. Load & clean**
- [ ] Power BI Desktop → Get Data → Excel → `pizza_sales_excel_file.xlsx` → sheet `pizza_sales` → Transform Data
- [ ] Rename query `stg_pizza_sales`; **untick Enable load**
- [ ] Set types (order_time = Time; prices = Decimal; ids = Whole; text as Text)
- [ ] Trim + Clean text columns; size → UPPERCASE; category → Capitalize Each Word
- [ ] Remove Duplicates (all columns)
- [ ] Profile: Column quality / distribution; confirm 0 errors, 0 empty
- [ ] Reference `stg_pizza_sales` → `FactPizzaSales` (remove `pizza_ingredients`)
- [ ] Reference → `DimPizza` (keep pizza_name_id, pizza_name, pizza_category, pizza_ingredients → remove duplicates on pizza_name_id)
- [ ] Reference → `DimSize` (pizza_size → remove duplicates → add `size_sort` 1–5)
- [ ] Reference → `DimCategory` (pizza_category → remove duplicates)
- [ ] Close & Apply

**B. Model**
- [ ] New table `DimDate` (Section 5.1 DAX) → Table tools → **Mark as date table** on `Date`
- [ ] Sort `Month Name` by `Month Number`; `Day Name` by `Day of Week Number`; `pizza_size` by `size_sort`
- [ ] Model view: create the 4 many-to-one, single-direction relationships
- [ ] Add `Order Hour`, `Meal Period` (+ sort) columns to fact
- [ ] Hide fact keys/text columns and helper columns; set IDs/prices to *Don't summarize*
- [ ] Options → Current File → Data Load → **untick Auto date/time**

**C. Measures**
- [ ] Create `_Measures` table; add all measures from Section 7 with display folders and formats
- [ ] Create `TopN Selector` (Modeling → New parameter → Numeric range 3–20, step 1)

**D. Report**
- [ ] Import `PizzaSales_Theme.json`
- [ ] Build Pages 1–3 (Section 6)
- [ ] Add slicers, sync date/category/size slicers across pages (View → Sync slicers)
- [ ] Create Tooltip page + Drill-through page; navigation and Home buttons
- [ ] Set alt text on all visuals; check colour contrast

**E. Validate & publish**
- [ ] Compare totals with Section 8; check DQ measures = 0
- [ ] Run Performance Analyzer
- [ ] Save `.pbix`; export PDF snapshot; screenshots for portfolio
- [ ] Publish to GitHub (README = this document + `.pbix` + theme + screenshots) and link on your resume/LinkedIn
