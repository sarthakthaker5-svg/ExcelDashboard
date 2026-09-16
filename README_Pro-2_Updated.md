# 📊 Sales & Data Analytics Dashboard – Excel Project

<p align="center">
  <img src="https://readme-typing-svg.demolab.com?font=Fira+Code&size=25&pause=1000&color=36BCF7&center=true&vCenter=true&width=700&lines=Sales+%26+Data+Analytics+Dashboard;Excel+Data+Analysis+Project;Formulas+%7C+Pivot+Tables+%7C+Regression+%7C+Visualizations" alt="Typing SVG" />
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Microsoft-Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white" />
  <img src="https://img.shields.io/badge/Data-Analysis-blue?style=for-the-badge" />
  <img src="https://img.shields.io/badge/Pivot-Tables-orange?style=for-the-badge" />
  <img src="https://img.shields.io/badge/Visualization-green?style=for-the-badge" />
</p>

---

## 📌 Project Overview

The **Sales & Data Analytics Dashboard** is an Excel-based data analysis project created to analyze customer sales, quantities, discounts, order dates, profits, regions, and product categories.

The project demonstrates the practical use of **Microsoft Excel** for formula-based analysis, conditional formatting, What-If Analysis, statistical analysis, regression, Pivot Tables, charts, and dashboard creation.

The workbook contains multiple worksheets that organize the workflow from raw sales data to analysis, visualization, and the final dashboard.

---

# 🎯 Project Objectives

The main objectives of this project are:

- Analyze sales data using Microsoft Excel.
- Analyze customer-wise total purchases.
- Highlight the top 10 customers based on total purchase.
- Perform What-If Analysis to study the impact of discount changes on profit.
- Perform Linear Regression between Profit and Sales.
- Generate Descriptive Statistics for important numerical fields.
- Calculate monthly sales growth.
- Add timestamps using the `NOW()` function.
- Identify high-value customers using `INDEX`, `MATCH`, and `FILTER`.
- Analyze sales by region and product category using Pivot Tables.
- Create Bar, Line, and Pie Charts.
- Build a centralized Excel Dashboard.
- Present important sales KPIs and analytical results.

---

# 🚀 Project Highlights

- 📊 **200 sales records analyzed**
- 💰 **Total Sales: ₹195,217.76**
- 💵 **Total Profit: ₹68,287.01**
- 📦 **Total Quantity Sold: 1,999**
- 📈 **Average Transaction: ₹976.09**
- 🔎 Customer-wise purchase analysis
- 🏆 Top 10 customer analysis
- 🧮 What-If discount scenario analysis
- 📉 Profit vs Sales Linear Regression
- 📊 Descriptive Statistics
- 📅 Monthly Sales and Growth Analysis
- ⏱️ Timestamp using `NOW()`
- 🔍 `INDEX + MATCH` and `FILTER`
- 📌 Pivot Table analysis
- 📊 Bar, Line and Pie Charts
- 🎯 Interactive-style Excel Dashboard

---

# 🛠️ Technologies and Excel Features Used

- **Microsoft Excel**
- Excel Tables
- Excel Formulas
- `SUM`
- `SUMIF`
- `SUMIFS`
- `SUMPRODUCT`
- `AVERAGE`
- `RANK`
- `IFERROR`
- `EDATE`
- `DATE`
- `NOW`
- `FILTER`
- `UNIQUE`
- `INDEX`
- `MATCH`
- Conditional Formatting
- Icon Sets / Directional Arrows
- What-If Analysis
- Data Analysis ToolPak
- Linear Regression
- Descriptive Statistics
- Pivot Tables
- Bar Charts
- Line Charts
- Pie Charts
- Dashboard Design

---

# 📂 Workbook Structure

```text
Sales & Data Analytics Dashboard
│
├── Raw Data
├── Analysis
├── Monthly Analysis
├── Pivot Table
├── Profit vs Sales
├── Descriptive Statistics
├── Visualizations
└── Dashboard
```

---

# 📘 Workbook Information

| Information | Details |
|---|---|
| Project Type | Sales & Data Analytics |
| Tool Used | Microsoft Excel |
| Total Records | 200 |
| Total Sales | ₹195,217.76 |
| Total Profit | ₹68,287.01 |
| Total Quantity | 1,999 |
| Average Transaction | ₹976.09 |
| Date Range | April 2024 – April 2025 |
| Main Analysis | Customers, Sales, Profit, Discount, Region & Product |
| Visualization | Bar Chart, Line Chart, Pie Chart |
| Dashboard | Excel Dashboard |
| Data Table | `SaleData` |

---

# 📑 Worksheets

## 1️⃣ Raw Data

The **Raw Data** sheet contains the sales dataset used for the complete project.

### Dataset Fields

- Customer_ID
- Customer_Name
- Region
- Product_Category
- Sales
- Quantity
- Discount
- Order_Date
- Profit
- Timestamp

The dataset contains **200 sales records**.

The data is organized in an Excel Table named:

```text
SaleData
```

The `Timestamp` column uses the `NOW()` function.

---

## 2️⃣ Analysis

The **Analysis** sheet contains formula-based customer analysis, What-If Analysis, high-value customer analysis, and discount sensitivity calculations.

### 👥 Customer Purchase Analysis

Customer names are summarized with their total purchase values using:

```excel
=SUMIF('Raw Data'!B:B,A2,'Raw Data'!E:E)
```

### 🏆 Top 10 Customers

Conditional Formatting is used to highlight the top 10 customers based on their total purchase.

The customer ranking uses:

```excel
=RANK(H2,$H$2:$H$31,0)
```

High-value customers are filtered using:

```excel
=FILTER(G2:H31,I2:I31<=10)
```

### 🔎 INDEX + MATCH

The customer with the highest total purchase can be identified using:

```excel
=INDEX(G2:G31,MATCH(MAX(H2:H31),H2:H31,0))
```

### 💡 What-If Analysis

The sheet contains:

- Current Total Profit
- Discount Scenario
- Scenario Profit

The scenario profit calculation is based on the change between the scenario discount and the original discount:

```excel
=E2-SUMPRODUCT('Raw Data'!E2:E201,E3-'Raw Data'!G2:G201)
```

For a **10% discount scenario**, the calculated scenario profit is approximately:

```text
₹68,037.99
```

A discount sensitivity section also contains discount values such as:

```text
0%
5%
10%
15%
20%
```

---

## 3️⃣ Monthly Analysis

The **Monthly Analysis** sheet analyzes sales month by month from **April 2024 to April 2025**.

### Monthly Sales

Monthly sales are calculated using:

```excel
=SUMIFS('Raw Data'!$E:$E,'Raw Data'!$H:$H,">="&A2,'Raw Data'!$H:$H,"<"&EDATE(A2,1))
```

### Monthly Growth

Growth is calculated using:

```excel
=IFERROR((B3-C3)/C3,"")
```

Conditional Formatting with directional arrows is applied to the monthly growth values to visually indicate increases and decreases.

---

## 4️⃣ Pivot Table

The **Pivot Table** sheet summarizes total sales by **Region** and **Product Category**.

### Region Analysis

The Pivot Table includes:

- Central
- East
- North
- South
- West

### Product Categories

The analysis includes:

- Books
- Clothing
- Electronics
- Furniture
- Office Supplies

### Total Sales by Region

| Region | Total Sales |
|---|---:|
| Central | ₹43,468.33 |
| East | ₹32,441.64 |
| North | ₹34,975.42 |
| South | ₹40,248.79 |
| West | ₹44,083.58 |

### Product Category Totals

| Product Category | Total Sales |
|---|---:|
| Books | ₹47,002.46 |
| Clothing | ₹41,927.59 |
| Electronics | ₹41,044.00 |
| Furniture | ₹36,612.89 |
| Office Supplies | ₹28,630.82 |

The overall Pivot Table sales total is:

**₹195,217.76**

---

## 5️⃣ Profit vs Sales

The **Profit vs Sales** sheet contains the Linear Regression output created using the Excel **Data Analysis ToolPak**.

### Regression Variables

- **Dependent Variable (Y):** Profit
- **Independent Variable (X):** Sales
- **Observations:** 200

### Regression Statistics

| Metric | Value |
|---|---:|
| Multiple R | 0.7734 |
| R Square | 0.5982 |
| Adjusted R Square | 0.5961 |
| Standard Error | 86.2812 |
| Observations | 200 |

The regression analysis is used to examine the relationship between Sales and Profit.

---

## 6️⃣ Descriptive Statistics

The **Descriptive Statistics** sheet contains statistical summaries for:

- Sales
- Quantity
- Discount
- Order_Date
- Profit

The analysis includes statistical measures such as:

- Mean
- Standard Error
- Median
- Mode
- Standard Deviation
- Sample Variance
- Kurtosis
- Skewness
- Range
- Minimum
- Maximum

### Selected Statistics

| Variable | Mean | Standard Deviation |
|---|---:|---:|
| Sales | ₹976.09 | ₹299.27 |
| Quantity | 9.995 | 5.65 |
| Discount | 9.725% | 7.23% |
| Profit | ₹341.44 | ₹135.77 |

---

## 7️⃣ Visualizations

The **Visualizations** sheet contains three main charts used to present the analysis visually.

### 📊 Bar / Column Chart

Used to compare sales across the analyzed categories/regions.

### 📈 Line Chart

Used to visualize the **Monthly Sales Trend** from April 2024 to April 2025.

### 🥧 Pie Chart

Used to show the distribution of sales across product categories.

These charts make the analytical results easier to understand and present.

---

## 8️⃣ Dashboard

The **Dashboard** sheet provides a centralized presentation of the main project results.

### 📌 KPI Cards

The dashboard includes:

- **Total Sales**
- **Total Profit**
- **Total Quantity**
- **Average Transaction**

### 📊 Dashboard Visualizations

The dashboard presents the main sales visualizations, including:

- Regional Sales Analysis
- Monthly Sales Trend
- Product Category Sales Distribution

The dashboard provides a single place to review the major sales performance indicators and visual analysis.

---

# 🧮 Excel Functions Covered

## Lookup and Reference Functions

```text
UNIQUE
FILTER
INDEX
MATCH
RANK
```

Used to identify, rank, filter, and retrieve customer information.

## Logical and Error Handling

```text
IFERROR
```

Used to handle calculation errors in growth analysis.

## Date and Time Functions

```text
DATE
EDATE
NOW
```

Used for monthly analysis, date calculations, and timestamps.

## Aggregation and Calculation

```text
SUM
SUMIF
SUMIFS
SUMPRODUCT
AVERAGE
```

Used for sales, profit, customer purchase, monthly sales, and scenario calculations.

---

# 📊 Sales Analysis

The project analyzes sales performance through:

- Customer-wise Total Purchase
- Top 10 Customers
- Regional Sales
- Product Category Sales
- Monthly Sales
- Monthly Growth
- Total Sales
- Total Profit
- Total Quantity
- Average Transaction
- Discount Impact on Profit

---

# 👥 Customer Analysis

The customer analysis identifies total purchases for individual customers.

The analysis uses `SUMIF` to calculate customer-level sales and `RANK` to order customers according to total purchase.

The top 10 high-value customers identified in the workbook are:

| Rank | Customer | Total Purchase |
|---:|---|---:|
| 1 | James Wilson | ₹20,774.74 |
| 2 | Mary Davis | ₹17,143.29 |
| 3 | David Brown | ₹11,915.83 |
| 4 | Emma Davis | ₹11,398.32 |
| 5 | Linda Wilson | ₹9,970.20 |
| 6 | Lisa Thomas | ₹8,635.95 |
| 7 | Emma Wilson | ₹8,543.97 |
| 8 | Robert Anderson | ₹8,331.85 |
| 9 | Lisa Davis | ₹8,090.12 |
| 10 | Michael Anderson | ₹7,881.90 |

---

# 📉 Discount and Profit Analysis

The project includes a discount scenario analysis to examine how changing the discount can affect total profit.

The current total profit is:

```text
₹68,287.01
```

The workbook uses a scenario discount of:

```text
10%
```

The resulting scenario profit is approximately:

```text
₹68,037.99
```

The workbook also provides a discount sensitivity section for different discount levels.

---

# 📈 Regression Analysis

The project uses **Linear Regression** to analyze the relationship between Sales and Profit.

The regression output contains:

- Regression Statistics
- ANOVA
- Coefficients
- Residual-related regression output

The model uses **200 observations**.

---

# 📊 Pivot Table Analysis

The Pivot Table summarizes sales by:

- Region
- Product Category

This allows the dataset to be analyzed without manually calculating every regional and product total.

---

# 📅 Monthly Growth Analysis

Monthly sales are calculated from April 2024 through April 2025.

The workbook calculates the percentage growth compared with the previous month.

Directional conditional-formatting arrows are used to make positive and negative monthly growth easier to identify.

---

# ⏱️ Timestamp Analysis

The **Raw Data** sheet contains a Timestamp column generated using:

```excel
=NOW()
```

This records the current Excel date and time when the workbook recalculates.

---

# 📊 Key Dataset Statistics

| Metric | Value |
|---|---:|
| Total Records | 200 |
| Total Sales | ₹195,217.76 |
| Total Profit | ₹68,287.01 |
| Total Quantity | 1,999 |
| Average Sales per Record | ₹976.09 |
| Number of Regions | 5 |
| Number of Product Categories | 5 |
| Date Range | Apr 2024 – Apr 2025 |

---

# 🔎 Business Analysis Areas

The project focuses on the following questions:

### Customer Performance

Which customers have the highest total purchase values?

### Regional Performance

How are total sales distributed across different regions?

### Product Performance

How are sales distributed across different product categories?

### Monthly Performance

How do sales change from month to month?

### Discount Impact

How does changing the discount scenario affect estimated total profit?

### Sales-Profit Relationship

What relationship exists between Sales and Profit?

### Statistical Analysis

What are the main descriptive statistics of Sales, Quantity, Discount, Order Date, and Profit?

---

# 🔄 Project Workflow

The complete project follows this workflow:

### Step 1 – Raw Data

Organize the original sales dataset containing customer, region, product, sales, quantity, discount, date, and profit information.

### Step 2 – Customer Analysis

Calculate customer-wise total purchases and identify high-value customers.

### Step 3 – What-If Analysis

Create discount scenarios and evaluate their impact on profit.

### Step 4 – Statistical Analysis

Perform Descriptive Statistics and Linear Regression using the Data Analysis ToolPak.

### Step 5 – Monthly Analysis

Calculate monthly sales and month-over-month growth.

### Step 6 – Pivot Table

Summarize sales by region and product category.

### Step 7 – Visualization

Create Bar, Line, and Pie Charts.

### Step 8 – Dashboard

Combine the main KPIs and visualizations into a centralized dashboard.

---

# 📌 Excel Concepts Practiced

```text
✔ Excel Tables
✔ Data Analysis
✔ Customer Analysis
✔ Conditional Formatting
✔ Top 10 Analysis
✔ What-If Analysis
✔ Linear Regression
✔ Descriptive Statistics
✔ Lookup Functions
✔ Date & Time Functions
✔ Monthly Growth Analysis
✔ Pivot Tables
✔ Bar Charts
✔ Line Charts
✔ Pie Charts
✔ Dashboard Design
✔ Data Visualization
```

---

# 📚 Learning Outcomes

Through this project, the following concepts were practiced:

- Excel Data Organization
- Excel Tables
- Customer-wise Sales Analysis
- Conditional Formatting
- What-If Analysis
- Data Analysis ToolPak
- Linear Regression
- Descriptive Statistics
- Lookup and Reference Functions
- Date and Time Functions
- Monthly Growth Analysis
- Pivot Table Creation
- Chart Creation
- Dashboard Design
- Business Data Interpretation

---

# 💡 Skills Demonstrated

## Technical Skills

- Microsoft Excel
- Excel Formulas
- Excel Tables
- `SUMIF` / `SUMIFS`
- `SUMPRODUCT`
- `FILTER`
- `UNIQUE`
- `INDEX-MATCH`
- `RANK`
- `IFERROR`
- `DATE`
- `EDATE`
- `NOW`
- Conditional Formatting
- What-If Analysis
- Data Analysis ToolPak
- Linear Regression
- Descriptive Statistics
- Pivot Tables
- Excel Charts
- Dashboard Creation

## Analytical Skills

- Customer Sales Analysis
- Regional Sales Analysis
- Product Category Analysis
- Profit Analysis
- Discount Impact Analysis
- Monthly Sales Analysis
- Growth Analysis
- Statistical Analysis
- Sales-Profit Relationship Analysis
- Data Visualization
- Business Data Interpretation

---

# 🏆 Project Achievements

- Successfully analyzed **200 sales records**
- Calculated total sales of **₹195,217.76**
- Calculated total profit of **₹68,287.01**
- Analyzed **1,999 total quantities**
- Created customer-wise purchase analysis
- Identified top 10 high-value customers
- Performed discount scenario analysis
- Performed Linear Regression using the Data Analysis ToolPak
- Generated Descriptive Statistics
- Created monthly sales growth analysis
- Added timestamps using `NOW()`
- Created Pivot Table analysis
- Created Bar, Line, and Pie Charts
- Built a centralized Excel Dashboard

---

# 📋 Dataset Fields

```text
Customer_ID
Customer_Name
Region
Product_Category
Sales
Quantity
Discount
Order_Date
Profit
Timestamp
```

---

# 📊 Main KPIs

The main dashboard KPIs are:

```text
Total Sales       → ₹195,217.76
Total Profit      → ₹68,287.01
Total Quantity    → 1,999
Average Transaction → ₹976.09
```

---

# 📤 Project Output

The final workbook contains:

```text
Raw Data
     ↓
Analysis
     ↓
Monthly Analysis
     ↓
Pivot Table
     ↓
Profit vs Sales
     ↓
Descriptive Statistics
     ↓
Visualizations
     ↓
Dashboard
```

The project provides a complete Excel-based workflow from raw sales data to analysis, statistical evaluation, visualization, and dashboard presentation.

---

# 📝 How to Use

1. Open the Excel workbook.
2. Start from the **Raw Data** sheet.
3. Review the 200 sales records.
4. Open **Analysis** to view customer, ranking, high-value customer, and What-If calculations.
5. Open **Monthly Analysis** to view monthly sales and growth.
6. Open **Pivot Table** to view regional and product-category sales.
7. Open **Profit vs Sales** to view the regression analysis.
8. Open **Descriptive Statistics** to view statistical summaries.
9. Open **Visualizations** to view the charts.
10. Open **Dashboard** to view the final presentation.

---

# 📌 Project Summary

The **Sales & Data Analytics Dashboard** is a practical Microsoft Excel project that demonstrates how raw sales data can be transformed into structured analysis and visual insights.

The project combines:

```text
Raw Sales Data
      +
Excel Formulas
      +
Customer Analysis
      +
What-If Analysis
      +
Statistical Analysis
      +
Regression
      +
Pivot Tables
      +
Data Visualization
      +
Dashboard
      =
Sales & Data Analytics Solution
```

This project demonstrates practical knowledge of **Excel Data Analysis, statistical analysis, customer analysis, sales performance analysis, Pivot Tables, visualization, and dashboard creation**.

---

# 👨‍💻 Author

**Sarth Thakar**
