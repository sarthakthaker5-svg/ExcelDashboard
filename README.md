::: {align="center"}
# 📊 Sales & Data Analytics Dashboard

`<img src="https://readme-typing-svg.demolab.com?font=Poppins&weight=600&size=28&duration=3000&pause=1000&color=36BCF7&center=true&vCenter=true&width=750&lines=Excel+Data+Cleaning;Sales+Data+Analysis;Pivot+Tables;Excel+Dashboard+%26+Visualization" alt="Typing SVG" />`{=html}

`<br>`{=html}

`<img src="https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white"/>`{=html}
`<img src="https://img.shields.io/badge/Data_Analysis-Excel-blue?style=for-the-badge"/>`{=html}
`<img src="https://img.shields.io/badge/Pivot_Tables-Analysis-orange?style=for-the-badge"/>`{=html}
`<img src="https://img.shields.io/badge/Status-Completed-brightgreen?style=for-the-badge"/>`{=html}

`<br>`{=html}`<br>`{=html}

*A beginner-friendly Excel analytics project demonstrating data
organization, cleaning, formula-based analysis, Pivot Tables,
visualizations, and an interactive sales dashboard.*
:::

------------------------------------------------------------------------

# 📖 Project Overview

**Sales & Data Analytics Dashboard** is an Excel-based data analysis
project created to practice practical spreadsheet and business analytics
skills.

The workbook uses a sales/order dataset containing **300 records** and
transforms the data into a structured analytical workbook. It includes
raw and cleaned data, Excel formulas, Pivot Table summaries, charts, a
dashboard, and project documentation.

The project demonstrates how Excel can be used to move from **raw order
data → cleaned data → calculations → Pivot Table analysis →
visualizations → dashboard reporting**.

------------------------------------------------------------------------

# 🎯 Objectives

-   Organize a structured sales/order dataset.
-   Work with Excel Tables and named data structures.
-   Prepare a cleaned master dataset for analysis.
-   Calculate total order value using Quantity and Unit Price.
-   Use lookup functions such as **VLOOKUP** and **INDEX-MATCH**.
-   Apply logical functions such as **IF**, **AND**, and **OR**.
-   Analyze products, quantities, order status, payment methods, and
    referral sources.
-   Build Pivot Tables for business-oriented summaries.
-   Create charts and visual representations of analytical results.
-   Present important sales information through a dashboard.
-   Document the workbook structure and usage instructions.

------------------------------------------------------------------------

# ✨ Project Highlights

-   ✔ 300 sales/order records.
-   ✔ Separate **Raw Data** and **Cleaned Data** sheets.
-   ✔ Structured Excel Table named **SalesData** in the Cleaned Data
    sheet.
-   ✔ Calculated **TotalPrice** for orders.
-   ✔ Lookup analysis using **VLOOKUP**.
-   ✔ Customer-based lookup using **INDEX-MATCH**.
-   ✔ Discount eligibility logic using **IF**.
-   ✔ Order-priority classification using nested **IF**, **AND**, and
    **OR** logic.
-   ✔ Pivot Table summaries for products, order status, payment methods,
    and referral sources.
-   ✔ Standard Excel visualizations.
-   ✔ Dashboard containing charts and analytical visuals.
-   ✔ Slicer/timeline-related workbook controls are included in the
    workbook structure.
-   ✔ Dedicated Documentation sheet explaining the workbook.

------------------------------------------------------------------------

# 🛠️ Technologies Used

  Technology        Purpose
  ----------------- ----------------------------------------------------
  Microsoft Excel   Data preparation, analysis, formulas and reporting
  Excel Tables      Structured data management
  Pivot Tables      Summarizing and analyzing sales data
  Excel Formulas    Calculations, lookups and business logic
  Charts            Visual representation of analytical results
  Dashboard         Consolidated visual reporting

------------------------------------------------------------------------

# 📂 Project Structure

``` text
Sales_Data_Analytics/
│
├── SarthThakar_ExcelDashboard.xlsx
├── README.md
│
└── Workbook Sheets/
    ├── Raw Data
    ├── Cleaned Data
    ├── Pivot Table
    ├── Formulas
    ├── Visualization
    ├── Dashboard
    └── Documentation
```

------------------------------------------------------------------------

# 📑 Workbook Structure

## 1. 📥 Raw Data

The **Raw Data** sheet contains the original order dataset in an Excel
Table named **Table1**.

The dataset contains **300 orders** with fields covering order
identification, customer information, product details, payment, delivery
status, coupons, referrals, and calculated order value.

### Main Columns

  Column            Description
  ----------------- ---------------------------------------------
  OrderID           Unique order identifier
  Date              Order date
  CustomerID        Customer identifier
  Product           Product purchased
  Quantity          Number of units ordered
  UnitPrice         Price per unit
  ShippingAddress   Shipping address
  PaymentMethod     Payment method used
  OrderStatus       Current order status
  TrackingNumber    Shipment tracking identifier
  ItemsInCart       Number of items in cart
  CouponCode        Coupon used for the order
  ReferralSource    Source through which the order was referred
  TotalPrice        Total order value

------------------------------------------------------------------------

## 2. 🧹 Cleaned Data

The **Cleaned Data** sheet contains the cleaned master dataset.

It is formatted as an official Excel Table named **SalesData**, making
it suitable for formulas, Pivot Tables, charts, and further analysis.

The structure contains the same 14 analytical fields as the Raw Data
sheet, including the calculated **TotalPrice** field.

------------------------------------------------------------------------

## 3. 📊 Pivot Table

The **Pivot Table** sheet contains summarized sales analysis created
from the dataset.

The workbook includes Pivot Table analysis for:

-   Product-wise Total Price
-   Product-wise Quantity
-   Product-wise Order Count
-   Order Status-wise Order Count
-   Order Status-wise Total Price
-   Payment Method and Referral Source revenue analysis

### Example Summary

The Pivot Table data shows an overall total of approximately
**₹306,976.92** in TotalPrice across the 300 records.

------------------------------------------------------------------------

## 4. 🧮 Formulas

The **Formulas** sheet demonstrates custom Excel calculations and
lookup/logic functions.

### Functions Demonstrated

  Formula / Function   Purpose
  -------------------- ---------------------------------------------
  VLOOKUP              Find the Unit Price for a selected Order ID
  INDEX-MATCH          Find Order Status using Customer ID
  IF                   Determine discount eligibility
  AND                  Combine conditions for order classification
  OR                   Check multiple risk-related order statuses
  Nested IF            Create Order Priority categories

### Examples of Logic Used

**Discount Eligibility**

Orders with a coupon code other than `"None"` are classified as:

-   `Discount Applied`
-   `No Discount`

**Order Priority**

The workbook classifies orders into:

-   `High Value-In Transit`
-   `Revenue Risk`
-   `Standard Order`

based on order value and order status conditions.

------------------------------------------------------------------------

# 📈 Analysis Performed

The workbook provides analysis across several important business
dimensions.

### 🛍️ Product Analysis

-   Total sales value by product.
-   Quantity sold by product.
-   Number of orders by product.

### 📦 Order Status Analysis

-   Cancelled orders.
-   Delivered orders.
-   Pending orders.
-   Returned orders.
-   Shipped orders.

### 💳 Payment Analysis

Sales value is analyzed across payment methods such as:

-   Cash
-   Credit Card
-   Debit Card
-   Gift Card
-   Online

### 📣 Referral Source Analysis

The Pivot Table also compares sales value across referral sources,
including:

-   Email
-   Facebook
-   Google
-   Instagram
-   Referral

------------------------------------------------------------------------

# 📊 Visualizations

The workbook contains a dedicated **Visualization** sheet for individual
charts and a **Dashboard** sheet for consolidated visual reporting.

The workbook contains multiple Excel charts, including:

-   Bar chart
-   Pie chart
-   Line chart

These visualizations help communicate sales performance and categorical
distributions in an easy-to-understand format.

------------------------------------------------------------------------

# 🖥️ Dashboard

The **Dashboard** sheet brings important visual analysis together in one
place.

It is designed to provide a quick overview of the sales dataset through
charts and dashboard-style reporting.

The workbook also contains defined controls related to:

-   Payment Method
-   Referral Source
-   Date timeline

These controls support interactive analysis within Excel where supported
by the workbook version.

------------------------------------------------------------------------

# 🗃️ Dataset Information

  Attribute               Details
  ----------------------- --------------------
  Dataset Type            Sales / Order Data
  Number of Records       300
  Number of Fields        14
  Raw Data Table          Table1
  Cleaned Data Table      SalesData
  Main Calculated Field   TotalPrice
  Main Analysis Tool      Pivot Tables
  Visualization Tool      Excel Charts
  Reporting Tool          Excel Dashboard

------------------------------------------------------------------------

# 🔑 Excel Concepts Covered

✔ Excel Tables\
✔ Data Cleaning / Preparation\
✔ Structured References\
✔ VLOOKUP\
✔ INDEX-MATCH\
✔ IF Function\
✔ Nested IF\
✔ AND Function\
✔ OR Function\
✔ Pivot Tables\
✔ Data Summarization\
✔ Sales Analysis\
✔ Order Status Analysis\
✔ Payment Method Analysis\
✔ Referral Source Analysis\
✔ Charts\
✔ Dashboard Reporting

------------------------------------------------------------------------

# ⚡ Key Calculations

## 💰 Total Price

The dataset contains a **TotalPrice** field representing the total value
of an order based on quantity and unit price.

Conceptually:

``` text
TotalPrice = Quantity × UnitPrice
```

------------------------------------------------------------------------

## 🎟️ Discount Eligibility

The workbook checks the **CouponCode** field to determine whether a
discount is applied.

``` text
If CouponCode ≠ "None"
→ Discount Applied

Otherwise
→ No Discount
```

------------------------------------------------------------------------

## 🚚 Order Priority

The workbook uses order value and order status to classify orders.

``` text
If TotalPrice > 2000 AND OrderStatus = "Shipped"
→ High Value-In Transit

If OrderStatus = "Cancelled" OR "Returned"
→ Revenue Risk

Otherwise
→ Standard Order
```

------------------------------------------------------------------------

# 📸 Project Output

The workbook itself contains the completed analysis and visual outputs.

Recommended screenshots for a GitHub project repository include:

  Screenshot         Description
  ------------------ ------------------------------------
  📥 Raw Data        Original sales/order dataset
  🧹 Cleaned Data    Cleaned SalesData table
  📊 Pivot Table     Pivot-based sales analysis
  🧮 Formulas        Lookup and logical formula results
  📈 Visualization   Individual charts
  🖥️ Dashboard       Consolidated Excel dashboard

> **Note:** Add screenshots to a `Screenshots` folder and update the
> image paths below if you want to display them directly on GitHub.

------------------------------------------------------------------------

# 🎓 Learning Outcomes

After completing this project, I gained practical experience in:

-   Organizing and structuring Excel datasets.
-   Preparing data for analysis.
-   Working with Excel Tables.
-   Writing lookup formulas.
-   Applying logical conditions.
-   Creating business-oriented calculations.
-   Building Pivot Tables.
-   Summarizing sales data.
-   Creating charts from analytical results.
-   Designing a dashboard-style report.
-   Presenting data in a more understandable visual format.
-   Documenting an Excel analytics project for portfolio use.

------------------------------------------------------------------------

# 💼 Skills Demonstrated

### Excel Skills

-   Excel Tables
-   Formula Writing
-   Lookup Functions
-   Logical Functions
-   Data Preparation
-   Pivot Tables
-   Charts
-   Dashboard Creation

### Analytical Skills

-   Sales Analysis
-   Product Analysis
-   Order Status Analysis
-   Payment Method Analysis
-   Referral Source Analysis
-   Data Summarization
-   Business Reporting

### Reporting Skills

-   KPI-oriented reporting
-   Visual data presentation
-   Dashboard organization
-   Analytical documentation

------------------------------------------------------------------------

# 🏆 Project Achievements

✅ Created a structured sales analytics workbook.

✅ Organized 300 order records into a reusable Excel dataset.

✅ Created a cleaned master table named **SalesData**.

✅ Implemented VLOOKUP and INDEX-MATCH.

✅ Applied IF, AND, OR and nested logical conditions.

✅ Created Pivot Table summaries.

✅ Built multiple Excel visualizations.

✅ Created a consolidated dashboard sheet.

✅ Documented the workbook structure and analytical workflow.

------------------------------------------------------------------------

# 🚀 How to Use

### Step 1

Download and open:

``` text
SarthThakar_ExcelDashboard.xlsx
```

### Step 2

Start with the **Documentation** sheet to understand the workbook
structure.

### Step 3

Open **Raw Data** to view the source order records.

### Step 4

Open **Cleaned Data** to work with the structured **SalesData** table.

### Step 5

Open **Formulas** to review the lookup and logical calculations.

### Step 6

Open **Pivot Table** to explore summarized sales analysis.

### Step 7

Open **Visualization** to view individual charts.

### Step 8

Open **Dashboard** for the consolidated visual report.

------------------------------------------------------------------------

# 📌 Project Summary

**Sales & Data Analytics Dashboard** demonstrates how Microsoft Excel
can be used to transform structured order data into meaningful business
analysis.

The project combines **data preparation, Excel formulas, lookup
functions, logical calculations, Pivot Tables, charts, and dashboard
reporting** into a single portfolio-ready workbook.

It is designed as a practical learning project for students and
beginners who want hands-on experience with Excel-based data analytics
and business reporting.

------------------------------------------------------------------------

# 👨‍💻 Author

## Sarth Thakar

**Excel Data Analytics Project**
