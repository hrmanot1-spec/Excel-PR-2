# Sales Analysis Excel Project

## Project Overview

This Excel project analyzes sample sales data and completes the 10 main
tasks listed in the project instructions. It includes formula-based
analysis, conditional formatting, monthly trends, a region/product
summary, charts, and a dashboard.

> **Note:** The workbook uses sample data because the original
> instructions did not include a separate sales dataset. Replace the
> sample records with your own data if required, while keeping the
> column structure consistent.

## Files

-   `Excel_Main_Tasks_Only_With_Visible_Charts.xlsx` --- the Excel
    workbook.
-   `README.md` --- this project guide.

## Workbook Sheets

### 1. Dataset

Contains sample order records with: - Order ID and Order Date -
Customer, Region, and Product - Quantity, Unit Price, and Discount -
Total Sales, Cost, Profit, and Timestamp

Key formulas: - **Total Sales:** `=F2*G2` - **Profit:**
`=I2*(1-H2)-J2` - **Timestamp:** `=NOW()`

The `NOW()` formula recalculates when Excel recalculates the workbook.

### 2. Analysis

Contains the calculations for the main tasks: - Total sales, total
profit, order count, and average discount - Top 10 customers by total
purchase, using `SUMIF`, `LARGE`, `INDEX`, and `MATCH` - What-If
Analysis comparing different discount rates and estimated total profit -
Regression measures for Profit vs Sales using `SLOPE`, `INTERCEPT`,
`RSQ`, and `CORREL` - Descriptive statistics including mean, median, and
standard deviation

### 3. Monthly Growth

Summarizes sales by month and calculates month-to-month growth. Up/down
arrows indicate whether sales increased or decreased compared with the
previous month.

### 4. Sales by Region-Product

Summarizes total sales for each region and product using `SUMIFS`. This
is a formula-based, PivotTable-style summary.

### 5. Dashboard

Shows key metrics and provides an overview of sales performance.

### 6. Charts

Contains the three charts required by the project: - **Bar Chart:**
Total Sales by Region - **Line Chart:** Monthly Sales Trend - **Pie
Chart:** Sales Share by Product

## Main Tasks Checklist

1.  **Conditional Formatting:** Highlights the highest purchase totals
    in the customer analysis.
2.  **What-If Analysis:** Compares estimated total profit under
    different discount rates.
3.  **Linear Regression:** Uses Profit as the dependent variable and
    Total Sales as the independent variable.
4.  **Descriptive Statistics:** Calculates mean, median, and standard
    deviation for selected fields.
5.  **Monthly Sales Growth:** Displays monthly growth percentages and
    up/down arrows.
6.  **Timestamp:** Adds a timestamp with `NOW()`.
7.  **High-Value Customers:** Uses `INDEX()`, `MATCH()`, `LARGE()`, and
    `SUMIF()` to identify top customers.
8.  **Sales by Region and Product:** Uses `SUMIFS()` to build a summary
    table.
9.  **Charts:** Includes bar, line, and pie charts on the **Charts**
    sheet.
10. **Dashboard:** Summarizes key metrics and visual insights.

## Using the What-If Analysis

1.  Open the **Analysis** sheet.
2.  Find the **WHAT-IF ANALYSIS --- DISCOUNT VS TOTAL PROFIT** section.
3.  Change a discount rate in the first column.
4.  Excel recalculates the estimated profit using the updated rate.

The estimate applies the selected discount rate to total sales and
subtracts total cost.

## Using Excel Analysis ToolPak

The workbook includes formula-based regression and descriptive
statistics. To also generate the ToolPak reports requested in the
instructions:

1.  In Excel, go to **File → Options → Add-ins**.
2.  At the bottom, choose **Excel Add-ins** and click **Go**.
3.  Tick **Analysis ToolPak** and click **OK**.
4.  Open **Data → Data Analysis**.
5.  For regression, select **Regression**:
    -   **Input Y Range:** Profit column
    -   **Input X Range:** Total Sales column
    -   Include the header only if the **Labels** option is selected.
6.  For descriptive statistics, choose **Descriptive Statistics**,
    select the required numeric columns, tick **Summary statistics**,
    and choose an output range.

Menu names may differ slightly depending on your Excel version.

## How to View the Charts

Open the workbook and select the **Charts** worksheet tab. The bar,
line, and pie charts are arranged on that sheet.

## Important Notes

-   Open the workbook in Microsoft Excel or a compatible spreadsheet app
    to calculate formulas and render charts.
-   If formulas do not update, set calculation mode to **Automatic** and
    recalculate the workbook.
-   `NOW()` is a volatile formula and changes when Excel recalculates.
-   The sample records are illustrative; insights and results will
    change if you replace the data.
-   This project is designed to demonstrate the required Excel tasks,
    not to represent actual business results.

## Formula Reference

  --------------------------------------------------------------------------------------------------
  Purpose                             Example formula
  ----------------------------------- --------------------------------------------------------------
  Calculate total sales               `=F2*G2`

  Calculate profit                    `=I2*(1-H2)-J2`

  Total purchase by customer          `=SUMIF(Dataset!$C$2:$C$121,F2,Dataset!$I$2:$I$121)`

  Highest ranked value                `=LARGE($G$2:$G$31,1)`

  Find matching customer              `=INDEX($F$2:$F$31,MATCH(LARGE($G$2:$G$31,1),$G$2:$G$31,0))`

  Sum sales by region and product     `=SUMIFS(...)`

  Timestamp                           `=NOW()`

  Regression slope                    `=SLOPE(profit_range,sales_range)`

  Regression R-squared                `=RSQ(profit_range,sales_range)`
  --------------------------------------------------------------------------------------------------

------------------------------------------------------------------------

**Project:** Excel Sales Analysis\
**Purpose:** Demonstrate data analysis, formulas, conditional
formatting, charts, and dashboard reporting in Excel.
