# E-Commerce Sales & Customer Retention Analysis Guide

**Step-by-step Excel instructions for analyzing the `orders.csv` dataset.**

---

## Dataset Overview

The `orders.csv` file contains **6,000 orders** across **18 months** (Jul 2024 – Dec 2025) for **2,270 unique customers**.

| Column | Example | Format |
|--------|---------|--------|
| `customer_id` | C-1108 | Text |
| `order_id` | O-1086 | Text |
| `order_date` | 01/01/2025 | MM/DD/YYYY |
| `order_amount` | 148.56 | Number (2 decimals) |
| `product_category` | Electronics | Text |
| `product_name` | Smart Watch Band | Text |
| `customer_acquisition_date` | 06/28/2025 | MM/DD/YYYY |

---

## Step 1: Load the Data

1. Open a **new Excel workbook**
2. Go to **Data** → **From Text/CSV**
3. Select `orders.csv`
4. In the preview window, verify:
   - Delimiter: **Comma**
   - File origin: **65001: Unicode (UTF-8)**
   - Data types detected automatically
5. Click **Load**
6. Rename the sheet to **Raw Data**

### Verify the data loaded correctly:
- Should have **6,000 rows** of data (plus header)
- All columns should be present
- Dates should be in MM/DD/YYYY format
- Order amounts should be numeric (no dollar signs or commas)

---

## Step 2: Clean the Data

Create a copy of the Raw Data sheet:

1. Right-click the **Raw Data** tab → **Move or Copy**
2. Check **Create a copy** → Click OK
3. Rename the copy to **Cleaned Data**

### Check for issues on the Cleaned Data sheet:

**Check for duplicates:**
- Select the `order_id` column
- **Data** → **Remove Duplicates** → OK
- If any duplicates are found, note the count

**Check for blank Customer IDs:**
- Select the `customer_id` column
- **Data** → **Filter** → Uncheck "(Blanks)"
- If blanks exist, delete those rows

**Standardize Product Categories:**
- Select the `product_category` column
- **Data** → **Filter** → Review all unique values
- Expected categories: Electronics, Clothing, Home & Kitchen, Beauty, Sports
- If any misspellings exist (e.g., "electronics"), use **Find & Replace** (Ctrl+H) to fix them

**Format as Table:**
- Select any cell in the data
- **Insert** → **Table** → Check "My table has headers" → OK
- This enables automatic formula referencing

---

## Step 3: Calculate Total Revenue

On a new sheet named **Sales Analysis**:

### KPI Summary Section (Top of sheet, cells A1:C6)

| Cell | Label | Formula |
|------|-------|---------|
| A1 | **Total Revenue** | |
| B1 | *(value)* | `=SUM(Cleaned Data[order_amount])` |
| A2 | **Total Orders** | |
| B2 | *(value)* | `=COUNTA(Cleaned Data[order_id])` |
| A3 | **Average Order Value** | |
| B3 | *(value)* | `=B1/B2` |

**Format B1** as Currency: `$#,##0.00`
**Format B3** as Currency: `$#,##0.00`

---

## Step 4: Analyze Monthly Sales

### Add a Helper Column

On the **Cleaned Data** sheet, add a new column header in **G1**: `Order Month`

In cell **G2**, enter this formula:
```
=TEXT([@order_date],"YYYY-MM")
```

Copy this formula down to all rows (if using a Table, it auto-fills).

### Create a Pivot Table for Monthly Sales

1. Select any cell in the Cleaned Data table
2. **Insert** → **PivotTable** → **New Worksheet** → OK
3. Rename the new sheet to **Monthly Analysis**

**Pivot Table Setup:**
- Drag `Order Month` to **Rows**
- Drag `order_amount` to **Values** (ensure it says "Sum of order_amount")
- Drag `order_id` to **Values** (ensure it says "Count of order_id")

The pivot table should look like:

| Order Month | Sum of order_amount | Count of order_id |
|-------------|--------------------:|-------------------:|
| 2024-07 | $XXX,XXX | XXX |
| 2024-08 | $XXX,XXX | XXX |
| ... | ... | ... |
| 2025-12 | $XXX,XXX | XXX |

### Add AOV Column

In the cell next to your pivot table (column D), add this formula for each row:
```
=[@[Sum of order_amount]]/[@[Count of order_id]]
```

Or manually: divide the Revenue cell by the Orders cell for each month.

### Create Charts

**Monthly Revenue Line Chart:**
1. Select the Order Month and Sum of order_amount columns from the pivot table
2. **Insert** → **Line Chart** → **2-D Line**
3. Title: "Monthly Revenue Trend"
4. Add data labels if desired

**Monthly Orders Column Chart:**
1. Select the Order Month and Count of order_id columns
2. **Insert** → **Column Chart** → **2-D Column**
3. Title: "Orders per Month"

### Identify Key Months

Look at the pivot table and note:
- **Highest revenue month:** _______________
- **Lowest revenue month:** _______________
- **General trend:** (increasing / decreasing / stable)

---

## Step 5: Analyze Revenue by Product Category

### Create a Pivot Table

1. Select any cell in the Cleaned Data table
2. **Insert** → **PivotTable** → **New Worksheet** → OK
3. Rename the new sheet to **Category Analysis**

**Pivot Table Setup:**
- Drag `product_category` to **Rows**
- Drag `order_amount` to **Values** (Sum of order_amount)
- Drag `order_id` to **Values** (Count of order_id)

The pivot table should look like:

| Product Category | Sum of order_amount | Count of order_id |
|------------------|--------------------:|-------------------:|
| Electronics | $XXX,XXX | XXX |
| Clothing | $XXX,XXX | XXX |
| Home & Kitchen | $XXX,XXX | XXX |
| Beauty | $XXX,XXX | XXX |
| Sports | $XXX,XXX | XXX |

### Create a Bar Chart

1. Select the Product Category and Sum of order_amount columns
2. **Insert** → **Bar Chart** → **2-D Bar**
3. Title: "Revenue by Product Category"
4. Sort bars from highest to lowest (right-click chart → **Sort** → **Sort Largest to Smallest**)

### Identify Top and Bottom Categories

- **Top category:** _______________
- **Bottom category:** _______________

---

## Step 6: Calculate Repeat Purchase Rate

### Create a Customer Summary Table

On the **Cleaned Data** sheet or a new sheet:

**Method: Use a Pivot Table**

1. **Insert** → **PivotTable** → **New Worksheet**
2. Rename to **Customer Analysis**

**Pivot Table Setup:**
- Drag `customer_id` to **Rows**
- Drag `order_id` to **Values** (Count of order_id)

This gives you each customer's total order count.

**Add a "Customer Type" Column:**
In the column next to the pivot table, use this formula (assuming order counts are in column C):
```
=IF([@[Count of order_id]]>=2, "Repeat Customer", "One-Time Customer")
```

### Calculate Repeat Purchase Rate

On the **Customer Analysis** sheet, add a summary section:

| Cell | Label | Formula |
|------|-------|---------|
| A1 | **Total Unique Customers** | |
| B1 | *(value)* | `=COUNTA(PivotTable[customer_id])` (exclude header/blank) |
| A2 | **Repeat Customers (2+ orders)** | |
| B2 | *(value)* | `=COUNTIF(PivotTable[Customer Type], "Repeat Customer")` |
| A3 | **Repeat Purchase Rate** | |
| B3 | *(value)* | `=B2/B1*100` |

**Format B3** as Percentage or Number with "%" suffix.

---

## Step 7: Identify Churned Customers

### Find the Most Recent Order Date

On the **Cleaned Data** sheet or in a helper cell:

```
=MAX(Cleaned Data[order_date])
```

Label this cell: **Analysis Date**

### Find Each Customer's Last Order Date

On the **Customer Analysis** sheet, add `order_date` to the Pivot Table Values as **Max of order_date**.

**Pivot Table Setup (updated):**
- Rows: `customer_id`
- Values: `Count of order_id`
- Values: `Max of order_date`

### Flag Churned Customers

In the column next to the pivot table, add this formula (assuming:
- Max of order_date is in column D
- Analysis Date is in a named cell or absolute reference like `$B$10`):

```
=IF(([@[Max of order_date]] - $B$10) < -90, "Churned", "Active")
```

**Note:** Excel stores dates as numbers, so subtraction gives days. If the result is less than -90 (meaning 90+ days ago), the customer is churned.

Alternatively, if your analysis date is in cell **H1** on the same sheet:
```
=IF(([@[Max of order_date]] - $H$1) < -90, "Churned", "Active")
```

### Count Churned Customers

| Cell | Label | Formula |
|------|-------|---------|
| A5 | **Potentially Churned Customers** | |
| B5 | *(value)* | `=COUNTIF(PivotTable[Status], "Churned")` |
| A6 | **Churn Percentage** | |
| B6 | *(value)* | `=B5/B1*100` |

---

## Step 8: Build the Dashboard

Create a new sheet named **Dashboard**.

### Layout Guide

Arrange the following elements on the Dashboard sheet:

```
+------------------------------------------------------------------+
|                    E-COMMERCE DASHBOARD                           |
+------------------------------------------------------------------+
|  [Total Revenue]  |  [Total Orders]  |  [Avg Order Value]       |
|    $XXX,XXX       |      XXXXX       |       $XXX.XX            |
+------------------------------------------------------------------+
|  [Repeat Purchase Rate]  |  [Churned Customers]                 |
|         XX.X%            |        XXX (XX.X%)                   |
+------------------------------------------------------------------+
|                                                                  |
|  [Monthly Revenue Line Chart]    |  [Revenue by Category Bar]   |
|                                  |                              |
|                                  |                              |
+------------------------------------------------------------------+
|  [Orders by Month Column Chart]  |  [Customer Status Pie/Bar]   |
|                                  |                              |
|                                  |                              |
+------------------------------------------------------------------+
```

### KPI Cards (Top Section)

Use **merged cells** to create large KPI display areas:

1. **Total Revenue**
   - Merge cells A2:C3
   - Enter: `='Sales Analysis'!B1`
   - Format as Currency, increase font size to 24pt
   - Add label "Total Revenue" above in smaller font

2. **Total Orders**
   - Merge cells D2:F3
   - Enter: `='Sales Analysis'!B2`
   - Format as Number with commas, font size 24pt
   - Add label "Total Orders"

3. **Average Order Value**
   - Merge cells G2:I3
   - Enter: `='Sales Analysis'!B3`
   - Format as Currency, font size 24pt
   - Add label "Average Order Value"

4. **Repeat Purchase Rate**
   - Merge cells A5:C6
   - Enter: `='Customer Analysis'!B3`
   - Format as Percentage, font size 24pt
   - Add label "Repeat Purchase Rate"

5. **Churned Customers**
   - Merge cells D5:F6
   - Enter: `='Customer Analysis'!B5` and append text
   - Format as Number, font size 24pt
   - Add label "Churned Customers"

### Charts

**Chart 1: Monthly Revenue Trend (Line Chart)**
1. Copy the monthly revenue pivot table data (not the pivot table itself)
2. Paste as values on the Dashboard sheet (or reference pivot table cells)
3. Select the Month and Revenue columns
4. **Insert** → **Line Chart** → **2-D Line**
5. Position below KPI cards, left side
6. Title: "Monthly Revenue Trend"

**Chart 2: Revenue by Category (Bar Chart)**
1. Copy the category revenue data
2. **Insert** → **Bar Chart** → **2-D Bar**
3. Position below KPI cards, right side
4. Title: "Revenue by Product Category"

**Chart 3: Orders by Month (Column Chart)**
1. Use the monthly order count data
2. **Insert** → **Column Chart** → **2-D Column**
3. Title: "Orders per Month"

**Chart 4: Customer Status (Pie Chart or Bar Chart)**
1. Create a small summary table:
   | Status | Count |
   |--------|-------|
   | Active | XXX |
   | Churned | XXX |
2. Select the data
3. **Insert** → **Pie Chart** or **Bar Chart**
4. Title: "Customer Status Distribution"

### Dashboard Formatting Tips

- Use a **consistent color scheme** (e.g., blue for positive metrics, orange/red for churn)
- Add **borders** around KPI cards
- Use **bold fonts** for numbers
- Keep **white space** between elements
- Add a **title** at the top: "E-Commerce Sales & Customer Retention Dashboard"
- Add a **date stamp**: "Data Period: Jul 2024 – Dec 2025"

---

## Step 9: Write the Summary Report

Create a new sheet named **Report** or add a text box to the Dashboard.

### Report Template

---

**E-Commerce Sales & Customer Retention Analysis**
**Reporting Period:** July 2024 – December 2025

---

**What We Found**

The company generated approximately $[TOTAL REVENUE] in total revenue over the 18-month period, with [TOTAL ORDERS] orders placed across all product categories. The average order value was $[AOV]. Sales showed [describe trend: e.g., "a general upward trend" / "seasonal fluctuations" / "steady growth"], with the strongest month being [MONTH] ($[AMOUNT]) and the weakest being [MONTH] ($[AMOUNT]).

**Customer Loyalty**

Approximately [XX]% of customers have made more than one purchase, indicating [good/moderate/low] customer retention. Out of [TOTAL CUSTOMERS] unique customers, [REPEAT CUSTOMERS] are repeat buyers. This suggests [brief interpretation of what this means for the business].

**Churn Analysis**

[XXX] customers ([XX.X]% of the customer base) have not made a purchase in the last 90 days and are considered potentially churned. These customers represent an opportunity for re-engagement through targeted promotional campaigns.

**Category Performance**

[Top Category] generated the highest revenue at $[AMOUNT], accounting for [XX]% of total sales. [Bottom Category] had the lowest revenue at $[AMOUNT], suggesting potential areas for improvement in product selection, pricing, or marketing.

---

**Recommendations**

1. **Re-engage Inactive Customers:** Launch an email campaign targeting the [XXX] churned customers with a limited-time discount offer to encourage them to return.

2. **Promote Top Categories:** Increase marketing visibility and inventory for [Top Category], which drives the most revenue. Consider bundling deals or featured placements.

3. **Improve Underperforming Categories:** Investigate why [Bottom Category] underperforms. Consider customer surveys, pricing reviews, or promotional support.

4. **Encourage Repeat Purchases:** Implement a loyalty program or offer first-time buyers a discount on their second purchase to increase the repeat purchase rate from [XX]% toward a target of [XX+10]%.

5. **Prepare for Peak Periods:** Based on the monthly trends, increase inventory and marketing spend before [strongest months] to capitalize on seasonal demand.

---

## Quick Reference: Excel Formulas Used

| Purpose | Formula |
|---------|---------|
| Total Revenue | `=SUM(Cleaned Data[order_amount])` |
| Total Orders | `=COUNTA(Cleaned Data[order_id])` |
| Average Order Value | `=Total Revenue / Total Orders` |
| Extract Month | `=TEXT([@order_date],"YYYY-MM")` |
| Customer Type | `=IF(Count>=2, "Repeat", "One-Time")` |
| Most Recent Order | `=MAX(order_date)` per customer |
| Churn Flag | `=IF((LastOrder - AnalysisDate) < -90, "Churned", "Active")` |
| Repeat Rate | `=Repeat Customers / Total Customers * 100` |
| Churn Count | `=COUNTIF(Status Column, "Churned")` |

---

## File Structure

Your final Excel workbook should have these sheets:

| Sheet Name | Contents |
|------------|----------|
| **Raw Data** | Original `orders.csv` data (untouched) |
| **Cleaned Data** | Verified and formatted data with Order Month helper column |
| **Sales Analysis** | KPI summary: Total Revenue, Total Orders, AOV |
| **Monthly Analysis** | Pivot table: monthly revenue, orders, and AOV |
| **Category Analysis** | Pivot table: revenue by product category |
| **Customer Analysis** | Pivot table: customer order counts, repeat/churn flags |
| **Dashboard** | KPI cards, charts, and visual summary |
| **Report** | Plain-language summary and recommendations |

---

## Time Estimate

| Step | Time |
|------|------|
| Load & clean data | 30 min |
| Total revenue & monthly analysis | 45 min |
| Average order value | 20 min |
| Category analysis | 30 min |
| Repeat purchase rate | 30 min |
| Churn identification | 30 min |
| Dashboard assembly | 45 min |
| Summary report | 30 min |
| Review & polish | 20 min |
| **Total** | **~5 hours** |
