# Conditional Formatting in Excel 2024 — step-by-step guide on the Credit dataset

Companion workbook: `Credit_ISLR_conditional_formatting.xlsx`

- Sheet **Customers** — the 400 customers of `Credit_ISLR.csv` as the table `tblCustomers`, with 12 conditional-formatting rules already applied.
- Sheet **Rules** — one row per rule: menu path, dialog settings, equivalent formula, format and a live count of the cells (or rows) that match.

Everything below can be rebuilt from scratch in Excel 2024 by following the steps in order.

## 1. Starting point

`Credit_ISLR.csv` has 400 rows. Two small clean-ups are applied when we import it:

- the first column of the file is an unnamed row index that duplicates `ID`, so we drop it;
- the values `" Male"` in `Gender` carry a leading space, so we trim them.

The 12 columns we work with, and their position in the sheet:

| Col | Field | Type | Meaning |
|---|---|---|---|
| A | ID | number | customer identifier |
| B | Income | number | income in thousands of dollars |
| C | Limit | number | credit limit |
| D | Rating | number | credit rating |
| E | Cards | number | number of credit cards |
| F | Age | number | age in years |
| G | Education | number | years of education |
| H | Gender | text | Male / Female |
| I | Student | text | Yes / No |
| J | Married | text | Yes / No |
| K | Ethnicity | text | category |
| L | Balance | number | average credit-card balance |

Data rows occupy rows 2 to 401; row 1 holds the headers.

## 2. Importing the CSV and converting it into a table

### Option A — open the file and press Ctrl+T (static copy)

1. **File ▸ Open** ▸ `Credit_ISLR.csv` (or double-click the file).
2. We delete column A (the row index) and, if we want, fix the leading space with **Home ▸ Find & Select ▸ Replace** (find `" Male"`, replace with `Male`).
3. We click any cell with data and press **Ctrl+T** (or **Home ▸ Format as Table**). We confirm that *My table has headers* is ticked and click **OK**.
4. **Table Design ▸ Table Name** ▸ `tblCustomers`.
5. **File ▸ Save As** ▸ *Excel Workbook (\*.xlsx)* — a CSV cannot store tables or formatting.

### Option B — Power Query (live connection to the file)

1. We open a blank workbook and go to **Data ▸ Get & Transform Data ▸ From Text/CSV**.
2. We pick `Credit_ISLR.csv`. The preview shows *File Origin*, *Delimiter* (Comma) and *Data Type Detection*; we check that the numeric columns are detected as numbers.
3. We click **Transform Data** to open the Power Query Editor and we:
   - select the first (unnamed) column ▸ **Home ▸ Remove Columns**;
   - select `Gender` ▸ **Transform ▸ Format ▸ Trim**.
4. **Home ▸ Close & Load**. Excel creates a new sheet with the data already formatted as a table and a query we can refresh from **Data ▸ Refresh All** whenever the CSV changes.
5. We click inside the table, open the **Table Design** tab and type `tblCustomers` in the **Table Name** box (top-left, *Properties* group). We can also rename the sheet to `Customers`.

Why a table? Filter buttons, structured references (`tblCustomers[Balance]`) in worksheet formulas, and — what matters here — every conditional-formatting rule applied to a table column **extends automatically** to the rows we add later.

## 3. How conditional formatting works (before we click anything)

**Selecting the right cells.** A rule applies to the cells that are selected when we create it. To select the data body of a table column we move the pointer to the top edge of the header cell until it becomes a small black down-arrow, then we click once (one click = data only, two clicks = data plus header). Alternatively we click the first data cell and press **Ctrl+Shift+↓**, or we type the range (for example `B2:B401`) in the Name Box and press Enter.

**The menu.** Everything lives in **Home ▸ Styles ▸ Conditional Formatting**, which has five galleries plus three commands:

| Gallery / command | What it does |
|---|---|
| Highlight Cells Rules | Greater Than, Less Than, Between, Equal To, Text that Contains, A Date Occurring, Duplicate Values |
| Top/Bottom Rules | Top 10 Items, Top 10 %, Bottom 10 Items, Bottom 10 %, Above Average, Below Average |
| Data Bars | a bar proportional to the value, drawn inside the cell |
| Color Scales | a two- or three-colour gradient across the range |
| Icon Sets | arrows, shapes, flags, ratings, etc. |
| New Rule… | all rule types, including *Use a formula to determine which cells to format* |
| Clear Rules | from selected cells, entire sheet, this table |
| Manage Rules… | the Rules Manager: order, edit, delete, *Stop If True*, *Applies to* |

**Formulas inside the rules.** Every rule is, under the hood, a test that returns TRUE or FALSE for each cell. The built-in dialogs simply hide the formula. When we write the formula ourselves (*New Rule ▸ Use a formula…*):

- we write it **for the top-left cell of the selected range** (for `B2:B401` we write `=B2>…`); Excel shifts the relative references down for the other cells;
- `$L2` (column locked, row free) is what makes a **whole-row** rule work: every cell of a row looks at the `Balance` of *its own* row;
- `$B$2:$B$401` (fully locked) is what we use for aggregates such as `AVERAGE` or `LARGE`, so that every cell compares itself with the whole column;
- **structured references are not accepted** inside conditional-formatting formulas: `=tblCustomers[Balance]>1000` is rejected. We use plain A1 references, or `INDIRECT("tblCustomers[Balance]")` if we really need the table name;
- if we click on a cell while typing, Excel inserts an absolute reference (`$B$2`); we press **F4** to cycle through `$B$2 → B$2 → $B2 → B2`.

**Priority.** Rules are evaluated in the order shown in the Rules Manager (top = highest priority). When two rules set the *same* property (two fills, two font colours) on the same cell, the higher one wins; when they set *different* properties (a fill and a border, a bar and a font colour) both are applied. This is why several rules can live on the same column without fighting.

## 4. The rules, one by one

Each block gives: the cells we select, the clicks, the settings, and the formula that produces the same result if we prefer *New Rule ▸ Use a formula…*.

### Rule 1 — Data Bars on Income (B2:B401)

1. We select `B2:B401`.
2. **Conditional Formatting ▸ Data Bars ▸ Gradient Fill ▸ Blue Data Bar**.

That is all: the bar length is proportional to the value, from the lowest income (shortest bar) to the highest (full cell). Through **Data Bars ▸ More Rules…** we can change the *Minimum/Maximum* type (Lowest/Highest Value, Number, Percent, Percentile, Formula), pick a solid fill, or tick *Show Bar Only* to hide the number.

Equivalent formula: none — graphical rules read the whole range at once and do not take a formula.

### Rule 2 — Greater Than on Limit (C2:C401)

1. We select `C2:C401`.
2. **Conditional Formatting ▸ Highlight Cells Rules ▸ Greater Than…**
3. We type `10000` and, in the *with* list, we keep **Light Red Fill with Dark Red Text**. **OK**.

13 credit limits above 10,000 are highlighted.

Equivalent formula: `=C2>10000` — the "Less Than" rule is `=C2<…`, "Equal To" is `=C2=…`.

### Rule 3 — Duplicate Values on Limit (C2:C401)

1. Same selection, `C2:C401`.
2. **Conditional Formatting ▸ Highlight Cells Rules ▸ Duplicate Values…**
3. We leave **Duplicate** selected (the other choice, *Unique*, does the opposite), then in the *values with* list we pick **Custom Format…** ▸ tab *Border* ▸ colour dark red ▸ **Outline** ▸ **OK** ▸ **OK**.

26 cells belong to a limit that appears more than once. A border was chosen on purpose: it does not conflict with the fill of rule 2, so a cell can show both.

Equivalent formula: `=COUNTIF($C$2:$C$401,C2)>1` (and `=COUNTIF($C$2:$C$401,C2)=1` for unique values).

### Rule 4 — Colour Scale on Rating (D2:D401)

1. We select `D2:D401`.
2. **Conditional Formatting ▸ Color Scales ▸ Green – Yellow – Red Color Scale** (hovering over a preset shows its name).

The lowest rating is red, the median is yellow, the highest is green. **Color Scales ▸ More Rules…** lets us switch to a 2-colour scale, move the midpoint (Percentile 50 by default) or fix the endpoints to numbers. When a *high* value is the bad one we pick the mirror preset **Red – Yellow – Green**.

Equivalent formula: none (graphical rule).

### Rule 5 — Icon Set on Cards (E2:E401)

1. We select `E2:E401`.
2. **Conditional Formatting ▸ Icon Sets ▸ Shapes ▸ 3 Traffic Lights (Unrimmed)**. Excel assigns the icons by percentiles (top third, middle third, bottom third), which is not what we want.
3. **Conditional Formatting ▸ Manage Rules…** ▸ we select the icon rule ▸ **Edit Rule…**
4. In *Edit the Rule Description* we set **Type = Number** on both rows, then the thresholds: first icon **>= 5**, second icon **>= 3** (the third icon takes everything below 3).
5. We click **Reverse Icon Order** so that the red light is on top: red when Cards >= 5, yellow when 3–4, green when 1–2. **OK**, **OK**.

Equivalent formula: none (graphical rule). Ticking *Show Icon Only* in the same dialog hides the numbers.

### Rule 6 — Between on Age (F2:F401)

1. We select `F2:F401`.
2. **Conditional Formatting ▸ Highlight Cells Rules ▸ Between…**
3. We type `18` and `30`, we choose **Yellow Fill with Dark Yellow Text**. **OK**.

32 customers aged 18–30 (both limits included) are highlighted.

Equivalent formula: `=AND(F2>=18,F2<=30)`.

### Rule 7 — Above Average on Education (G2:G401)

1. We select `G2:G401`.
2. **Conditional Formatting ▸ Top/Bottom Rules ▸ Above Average…**
3. We choose **Green Fill with Dark Green Text**. **OK**.

215 cells are above the column average (13.45 years). The average is recomputed every time the data changes; **Below Average…** is the mirror rule.

Equivalent formula: `=G2>AVERAGE($G$2:$G$401)` — note the locked range.

### Rule 8 — Equal To on Student (I2:I401)

1. We select `I2:I401`.
2. **Conditional Formatting ▸ Highlight Cells Rules ▸ Equal To…**
3. We type `Yes`, we open the *with* list and pick **Custom Format…** ▸ tab *Font* ▸ Bold, colour dark blue ▸ tab *Fill* ▸ light blue ▸ **OK** ▸ **OK**.

The 40 students are highlighted. Text comparisons are not case-sensitive. For partial matches we use **Text that Contains…** instead (for example `Fem` on `Gender`).

Equivalent formula: `=I2="Yes"` — for "Text that Contains": `=ISNUMBER(SEARCH("Yes",I2))`.

### Rule 9 — Top 10 Items on Balance (L2:L401)

1. We select `L2:L401`.
2. **Conditional Formatting ▸ Top/Bottom Rules ▸ Top 10 Items…**
3. We keep `10`, we pick **Custom Format…** ▸ *Fill* gold ▸ *Font* Bold ▸ **OK** ▸ **OK**.

The 10 highest balances get a gold fill (ties, if any, are all included).

Equivalent formulas: `=L2>=LARGE($L$2:$L$401,10)`; Top 10 %: `=L2>=PERCENTILE($L$2:$L$401,0.9)`; Bottom 10 Items: `=L2<=SMALL($L$2:$L$401,10)`.

### Rule 10 — Formula rule on Income: statistical outliers (B2:B401)

1. We select `B2:B401`.
2. **Conditional Formatting ▸ New Rule… ▸ Use a formula to determine which cells to format**.
3. In *Format values where this formula is true* we type

   `=B2>AVERAGE($B$2:$B$401)+2*STDEV.S($B$2:$B$401)`

4. **Format…** ▸ *Font* ▸ Bold, colour red ▸ **OK** ▸ **OK**.

24 incomes are more than two standard deviations above the mean (z-score > 2). The rule sits on the same column as the data bars of rule 1: a font colour and a bar do not conflict, so both show. (The workbook stores the rule with `STDEV`, the older name of `STDEV.S`; the result is identical.)

### Rule 11 — Formula rule on the whole row: no balance (A2:L401)

1. We select the whole data body `A2:L401` (we type it in the Name Box, or we click a cell of the table and press **Ctrl+A** once).
2. **New Rule… ▸ Use a formula to determine which cells to format**.
3. Formula: `=$L2=0`
4. **Format…** ▸ *Font* ▸ Italic, colour grey ▸ **OK** ▸ **OK**.

90 rows with a zero balance turn grey. The `$` in front of `L` is the whole trick: the column is locked, the row is not, so cell A2 tests L2, cell B2 tests L2, … and cell A3 tests L3.

### Rule 12 — Formula rule on the whole row: two conditions (A2:L401)

1. Same selection, `A2:L401`.
2. **New Rule… ▸ Use a formula to determine which cells to format**.
3. Formula: `=AND($I2="Yes",$L2>1000)`
4. **Format…** ▸ *Fill* ▸ light purple ▸ **OK** ▸ **OK**.

20 rows — students with a balance above 1,000 — get a purple background. `OR(...)` and `NOT(...)` work the same way; a whole-row rule based on a text column is `=$I2="Yes"`.

Because rules 1–10 were created first they sit higher in the Rules Manager, so on *their* cells (Income, Limit, Rating, …) the column colour wins over the purple fill; the remaining cells of the row are purple.

## 5. Managing the rules

**Conditional Formatting ▸ Manage Rules…** opens the *Conditional Formatting Rules Manager*. In *Show formatting rules for* we choose **This Worksheet** (or **This Table**) to see all 12 rules at once. From here we can:

- change the order with the up/down arrows (higher = higher priority);
- **Edit Rule…** to change thresholds, formulas or formats; **Duplicate Rule** to copy one; **Delete Rule**;
- tick **Stop If True** so that the rules below are not evaluated for the cells that match — useful when two fills overlap and we want an explicit winner;
- edit the **Applies to** range of a rule directly.

**Clear Rules ▸ Clear Rules from Selected Cells / Entire Sheet / This Table** removes the formatting without touching the data.

## 6. Quick reference

| # | Column / range | Menu path | Settings | Formula equivalent | Format |
|---|---|---|---|---|---|
| 1 | Income · B2:B401 | Data Bars ▸ Gradient Fill ▸ Blue | min = lowest, max = highest | — | blue gradient bar |
| 2 | Limit · C2:C401 | Highlight Cells ▸ Greater Than | 10000 | `=C2>10000` | light red fill, dark red text |
| 3 | Limit · C2:C401 | Highlight Cells ▸ Duplicate Values | Duplicate | `=COUNTIF($C$2:$C$401,C2)>1` | dark red border |
| 4 | Rating · D2:D401 | Color Scales ▸ Green–Yellow–Red | min / 50th pct / max | — | red → yellow → green |
| 5 | Cards · E2:E401 | Icon Sets ▸ 3 Traffic Lights, then Edit Rule | Number: ≥5 red, ≥3 yellow, <3 green, reversed | — | traffic-light icon |
| 6 | Age · F2:F401 | Highlight Cells ▸ Between | 18 and 30 | `=AND(F2>=18,F2<=30)` | yellow fill, dark yellow text |
| 7 | Education · G2:G401 | Top/Bottom ▸ Above Average | — | `=G2>AVERAGE($G$2:$G$401)` | green fill, dark green text |
| 8 | Student · I2:I401 | Highlight Cells ▸ Equal To | Yes | `=I2="Yes"` | light blue fill, bold dark blue text |
| 9 | Balance · L2:L401 | Top/Bottom ▸ Top 10 Items | 10 | `=L2>=LARGE($L$2:$L$401,10)` | gold fill, bold |
| 10 | Income · B2:B401 | New Rule ▸ Use a formula | — | `=B2>AVERAGE($B$2:$B$401)+2*STDEV.S($B$2:$B$401)` | bold red text |
| 11 | whole row · A2:L401 | New Rule ▸ Use a formula | — | `=$L2=0` | grey italic text |
| 12 | whole row · A2:L401 | New Rule ▸ Use a formula | — | `=AND($I2="Yes",$L2>1000)` | light purple fill |

## 7. Common pitfalls

- **Wrong active cell.** A formula rule written for `B2` but applied while `B1` (the header) was part of the selection shifts every test by one row. We select data cells only, and we check *Applies to* in the Rules Manager.
- **Missing or extra `$`.** `=L2=0` on `A2:L401` would make column A test column L, column B test column M, and so on. Whole-row rules always lock the column: `=$L2=0`.
- **Structured references.** `tblCustomers[Balance]` is fine in a worksheet formula (the *Cells matched* column of the Rules sheet could use it) but is rejected inside a conditional-formatting rule.
- **Numbers stored as text.** A CSV imported with the wrong locale can turn `14.891` into text; numeric rules then match nothing. We check the *Data Type Detection* step of Power Query, or the green triangle Excel shows in the corner of text-numbers.
- **Overlapping fills.** Two fill rules on the same cells show only the higher one; we use the order in the Rules Manager (or *Stop If True*) to decide the winner, or we combine different properties (fill + border, bar + font) as in rules 2/3 and 1/10.
- **Copying formats.** The Format Painter copies conditional formatting together with the rest of the format; pasting a rule onto a different range recreates it with shifted references, which is sometimes what we want and sometimes not.
