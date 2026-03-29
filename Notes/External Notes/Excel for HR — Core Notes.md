# Excel for [[Human Resources (HR)|HR]] — Core Notes

## BSB40420 Certificate IV in Human Resource Management

---

## Module 1 — Text Functions, Formulas, Charts & Consolidation

---

### Text to Columns

Used when data exported from another system combines values that should be in separate columns (e.g. first name and last name in one cell).

**How to use it:**

1. Highlight the column you want to split
2. Go to Data tab → Data Tools → Text to Columns
3. Step 1: Choose "Delimited" (for spaces, commas, tabs) or "Fixed Width"
4. Step 2: Select the delimiter — tick "Space" to split on spaces
5. Step 3: Set the destination cell (e.g. C2) and press Finish

You can split one row at a time or the entire column at once.

---

### Flash Fill

Used to extract part of a value based on a pattern (e.g. extracting just the state abbreviation from a "City, State" column).

**How to use it:**

- Option 1: Start typing the pattern in the next column — Excel will detect it and offer to complete it. Press Enter to accept.
- Option 2: Type a couple of examples, then go to Data → Flash Fill
- Option 3: Shortcut — **Ctrl + E**

Flash Fill looks to the left of your cursor for a pattern. If it doesn't pick up the pattern right away, add more examples and run it again. If the output is wrong, correct it in the cell and Flash Fill will retrain.

---

### TODAY Function

Returns the current date automatically — useful for calculating length of service.

```
=TODAY()
```

**Calculating length of service in days:**

```
=TODAY() - G2
```

Where G2 contains the hire date. Dates are stored as numbers in Excel, so subtraction works.

**Important:** If you autofill this formula down, lock the TODAY cell with an absolute reference using F4:

```
=$M$2 - G2
```

The dollar signs lock the row and column so the reference doesn't shift when you fill down.

---

### Basic Formulas and Operators

|Operation|Operator|Example|
|---|---|---|
|Addition|+|=A2+B2|
|Subtraction|-|=A2-B2|
|Multiplication|*|=E2*F2|
|Division|/|=A2/B2|

**Gross pay example:**

```
=E2*F2
```

(Pay rate × Hours)

**Days remaining example (with absolute reference):**

```
=$K$2-H2
```

(Fixed days allotted minus days used — K2 is locked because it's the same for all employees)

**Autofill tip:** Double-click the fill handle (small square, bottom-right of a cell) to automatically fill down, as long as there's data in the column to the left.

---

### Using a Form to Manage Records

A built-in data entry form lets you scroll through, add, edit, and delete records without navigating the spreadsheet directly.

**How to add it:**

1. Click the Quick Access Toolbar dropdown → More Commands
2. Change "Popular Commands" to "All Commands"
3. Type F, find "Form," and click Add
4. Press OK

**How to use it:** Click inside your dataset, then click the Form button. You can scroll through records, edit fields, add new records, or delete records. Calculated fields (like gross pay) are display-only — you can't edit them in the form.

---

### IF Function

Checks whether a condition is true or false and returns a different value for each.

**Syntax:**

```
=IF(logical_test, value_if_true, value_if_false)
```

**Example:**

```
=IF(K5>=K6, "Yes", "No")
```

If the sale (K5) is greater than or equal to the goal (K6), return "Yes" — otherwise "No."

This is the foundation for COUNTIF, SUMIF, and AVERAGEIF.

---

### COUNTIF

Counts how many times a value appears in a range — e.g. how many training programs a specific employee has completed.

**Syntax:**

```
=COUNTIF(range, criteria)
```

**Example:**

```
=COUNTIF(C2:C51, G2)
```

Counts how many times the name in G2 appears in the employee column (C2:C51).

Use a cell reference for criteria (not a typed name) so you can change the input and get instant results.

**Tip:** Use Ctrl + Shift + Down to quickly highlight a long column range.

---

### SUMIF

Adds up values in one column where a matching condition is met in another column — e.g. total training cost for one employee or department.

**Syntax:**

```
=SUMIF(range, criteria, sum_range)
```

**Example — by employee:**

```
=SUMIF(C2:C51, G5, D2:D51)
```

Finds all rows where the employee column matches G5, then sums the corresponding costs in column D.

**Example — by department (using named ranges):**

```
=SUMIF(department, G8, cost)
```

---

### AVERAGEIF

Returns the average of values in a column where a condition is met — e.g. average training cost for the Leadership program.

**Syntax:**

```
=AVERAGEIF(range, criteria, average_range)
```

**Example:**

```
=AVERAGEIF(program, G10, cost)
```

Finds all rows where the program matches G10, then averages the cost column.

---

### Naming Ranges

Instead of repeatedly highlighting the same columns, you can give a range a name and refer to it by that name in any formula.

**How to name a range:**

1. Highlight the cells (e.g. D2:D51)
2. Click in the Name Box (top left, where the cell address shows)
3. Type a name (e.g. `cost`) and press Enter

**Using a named range in a formula:**

```
=SUM(cost)
=AVERAGEIF(program, G10, cost)
```

Named ranges work across the entire workbook, not just the sheet they're on. To view all named ranges: Formulas tab → Define Names → drop down.

---

### XLOOKUP

Searches a column for a value and returns a corresponding value from another column — e.g. enter an employee number and return their name, hire date, or pay rate.

**Syntax:**

```
=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found])
```

|Argument|What it means|
|---|---|
|lookup_value|What you're searching for (e.g. employee number in A2)|
|lookup_array|Where to search (e.g. the employee number column)|
|return_array|What to return (e.g. the name column)|
|if_not_found|Optional — custom message if no match (e.g. "No such ID")|

**Example:**

```
=XLOOKUP(A2, empnum, name)
=XLOOKUP(A2, empnum, hiredate,  "No such ID")
```

By default, XLOOKUP searches for an exact match. Use named ranges to avoid navigating between sheets each time. If the result shows a serial number instead of a date, right-click → Format Cells → Short Date.

---

### Consolidation Tool

Combines data from multiple worksheets into a single summary — e.g. merging training expense data from California, Washington, and Oregon sheets.

**How to use it:**

1. Go to the summary sheet and highlight the destination range (must match the layout of source sheets)
2. Data tab → Data Tools → Consolidate
3. Set Function to Sum
4. Click the arrow, go to the first source sheet, highlight the same range, click Add
5. Repeat for each sheet — Excel auto-selects the same range on each
6. Tick "Top row" and "Left column" under Labels
7. Press OK

**Key requirement:** All source sheets must use the same headers and layout. The tool is intuitive enough to handle unique labels if they differ across sheets.

---

### Inserting Charts

**Recommended Chart (easiest approach):**

1. Select your data range
2. Insert tab → Charts → Recommended Charts
3. Excel analyses your data and suggests the best chart types
4. Choose one and click OK

**Common chart types:**

- Clustered column — compare multiple series side by side (good for quarterly data by program)
- Line — show trends over time
- Pie — show proportion of a whole
- Tree map — bigger box = higher value

**Customising:**

- Chart Design tab → Chart Styles (paintbrush icon) for quick visual presets
- Switch Row/Column to flip how data is grouped
- Click on Move Chart to move it to its own worksheet

Charts remain dynamically linked to source data — change a value and the chart updates automatically.

---

### Sparklines

Mini charts that sit inside a single cell — used for quick trend analysis without a full chart.

**How to insert:**

1. Click the cell where you want the sparkline
2. Insert tab → Sparklines → Line (or Column)
3. Set the data range (e.g. B4:E4 for Q1–Q4 of one row)
4. Press OK

**Tips:**

- Make the row taller and column wider to see the sparkline clearly
- Use the Sparkline contextual tab to highlight high/low points and change the style
- Grab the fill handle and drag down to create sparklines for multiple rows at once
- Switch between Line and Column style from the Sparkline tab

---

## Module 2 — Tables, Conditional Formatting, Subtotals, Data Validation & Pivot Tables

---

### Converting a List to a Table

Tables give you automatic filters, slicers, a total row, and better formatting options.

**How to convert:**

1. Click inside your data
2. Home tab → Styles → Format as Table
3. Choose a style, confirm the range, press OK

Once converted, every column header has a filter dropdown. You can sort A–Z, Z–A, by colour, or search for specific values.

---

### Filtering

**Single filter:** Click a column header dropdown → select a value or use the search bar.

**Date filter:** Click the hire date dropdown → Date Filters → Between → enter start and end dates.

**Multiple filters:** Apply filters on more than one column at the same time — they stack. Excel shows how many records match (e.g. "10 of 51 records found").

**Clear a filter:** Click the filter icon on that column → Clear Filter from [column name].

---

### Slicers

Visual filter buttons — click once to filter the table, click again to clear. Better for repeated filtering than dropdown menus.

**How to insert:**

1. Click inside the table
2. Table Design tab → Insert Slicer
3. Tick the column(s) you want slicers for, press OK

**Tips:**

- Multiple slicers work together — filtering one updates the others
- Change the number of columns in the Slicer tab to make long lists more readable
- Resize and reposition slicers like any other object
- Delete a slicer with the Delete key

---

### Total Row

Adds a summary row at the bottom of the table with dropdown formulas — no manual formula writing needed.

**How to turn on:** Table Design tab → tick Total Row.

Each cell in the total row has a dropdown offering Sum, Average, Count, Max, Min, and more. The totals update automatically when you filter or use slicers.

---

### Conditional Formatting

Colour-codes cells based on their values to make patterns visible without moving or hiding data.

**Data bars:** Show relative size of values as in-cell bar charts.

- Highlight the column → Conditional Formatting → Data Bars → choose a style

**Colour scales:** Apply a gradient (e.g. green = high, red = low) across a range.

- Highlight the column → Conditional Formatting → Colour Scales

**Highlight cell rules:** Highlight cells that meet a specific condition.

- Highlight the column → Conditional Formatting → Highlight Cell Rules → Greater Than / Less Than / Between / etc.
- Choose your threshold value and formatting colour

**Top/Bottom rules:**

- Conditional Formatting → Top/Bottom Rules → Top 10% (or 10 items, bottom, etc.)

After applying conditional formatting, you can filter by colour: click the column dropdown → Filter by Colour.

---

### Subtotal Tool

Generates grouped subtotals within a list based on a column — e.g. total hours and length of service per state.

**Requirements:** Works on lists only (not formatted tables). Must sort the column of interest first.

**How to use it:**

1. Sort the column you want to group by (e.g. State, A–Z)
2. Data tab → Outline group → Subtotal
3. "At each change in" = your grouping column (e.g. State)
4. "Use function" = Sum (or Average, Count, etc.)
5. "Add subtotal to" = tick the columns to summarise (e.g. Length of Service, Hours)
6. Press OK

**Outline buttons** (top-left, numbered 1–3):

- Button 1 = grand total only
- Button 2 = subtotal per group + grand total
- Button 3 = all detail rows + subtotals

---

### Data Validation

Restricts what can be entered into a cell — e.g. forcing users to pick from a set list of turnover reasons.

**How to set up a dropdown list:**

1. Highlight the cells to restrict
2. Data tab → Data Validation
3. Under Allow, select "List"
4. In Source, either type options separated by commas, or highlight a cell range containing the options
5. Press OK

Cells will show a dropdown arrow. If someone tries to type something not on the list, Excel shows an error. You can customise the error message under the Error Alert tab.

---

### Importing Data from External Sources

Data tab → Get & Transform → choose your source type:

- Text/CSV — most common for HR system exports
- From Web — paste a URL
- From Table/Range — use data already in the workbook
- Get Data → for PDFs, folders, SharePoint, and more

When importing a CSV: check the preview, confirm the delimiter looks right, then click Load. Excel converts it to a table automatically.

---

### Pivot Tables

Pivot tables summarise large datasets without writing formulas — drag fields to instantly group, count, sum, or average data.

**How to insert:**

1. Click inside your data
2. Insert tab → Pivot Table (or Table Design → Summarise with Pivot Table)
3. Choose "New Worksheet," press OK
4. Drag fields into the Rows, Columns, Values, and Filters boxes on the right

**Recommended Pivot Tables:** Insert → Recommended Pivot Tables shows pre-built options based on your data.

**Changing the calculation type:**

- Click the dropdown in the Values box → Value Field Settings
- Change from Sum to Average, Count, Max, Min, etc.

**Changing number format:**

- Right-click a value in the pivot table → Value Field Settings → Number Format

**Show as percentage of grand total:**

- Value Field Settings → Show Values As → % of Grand Total

---

### Drill Down Reports

Double-click any value in a pivot table to create a new worksheet showing all the individual records that make up that number. Useful for investigating a specific department, program, or result.

---

### Report Filter Pages

Creates separate pivot table worksheets automatically — one per item in a filter field.

**How to use it:**

1. Drag a field into the Filters box (e.g. Position)
2. Pivot Table Analyze tab → Options dropdown → Show Report Filter Pages
3. Select the filter field, press OK

Excel creates one worksheet per unique value (e.g. one per job position), each with its own independent pivot table.

---

### Pivot Charts

A chart linked directly to a pivot table — filters on either one update both.

**How to insert:**

1. Click inside your pivot table
2. Pivot Table Analyze tab → Pivot Chart
3. Choose a chart type, press OK

Use the filter buttons on the pivot chart to filter by any field. Change chart style from the Chart Design tab. Works the same way as regular charts for resizing, moving, and styling.

---

## Quick Reference — Key Formulas

|Task|Formula|
|---|---|
|Today's date|`=TODAY()`|
|Length of service (days)|`=TODAY()-G2` or `=$M$2-G2`|
|Gross pay|`=E2*F2`|
|Days remaining|`=$K$2-H2`|
|Simple IF|`=IF(K5>=K6,"Yes","No")`|
|Count by criteria|`=COUNTIF(C2:C51,G2)`|
|Sum by criteria|`=SUMIF(range,criteria,sum_range)`|
|Average by criteria|`=AVERAGEIF(range,criteria,avg_range)`|
|Lookup by ID|`=XLOOKUP(A2,empnum,name,"No such ID")`|
|Sum a named range|`=SUM(cost)`|

**Absolute reference shortcut:** With the cell reference selected in a formula, press **F4** to add $ signs and lock the reference.