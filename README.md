# 📊 Sales & Data Analytics Dashboard – Excel Project

<p align="center">
  <img src="https://readme-typing-svg.demolab.com?font=Fira+Code&size=25&pause=1000&color=36BCF7&center=true&vCenter=true&width=700&lines=Sales+%26+Data+Analytics+Dashboard;Excel+Data+Analysis+Project;Pivot+Tables+%7C+Formulas+%7C+Visualizations" alt="Typing SVG" />
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Microsoft-Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white" />
  <img src="https://img.shields.io/badge/Data-Analysis-blue?style=for-the-badge" />
  <img src="https://img.shields.io/badge/Pivot-Tables-orange?style=for-the-badge" />
  <img src="https://img.shields.io/badge/Visualization-green?style=for-the-badge" />
</p>

---

## 📌 Project Overview

The **Sales & Data Analytics Dashboard** is an Excel-based data analysis project designed to analyze sales orders, customer information, products, order status, payment methods, referral sources, and revenue.

The project demonstrates the practical use of **Microsoft Excel** for data cleaning, formula-based analysis, Pivot Tables, charts, and dashboard creation.

The workbook contains multiple sheets that organize the complete workflow from raw data to final visualization and documentation.

---

## 🎯 Objectives

The main objectives of this project are:

- Analyze sales and order data using Microsoft Excel.
- Clean and organize the raw dataset.
- Calculate and analyze total sales/revenue.
- Use Excel formulas for lookup and logical operations.
- Create Pivot Tables for business analysis.
- Analyze product-wise sales performance.
- Analyze order status distribution.
- Analyze revenue by payment method.
- Analyze revenue by referral source.
- Create charts and visualizations.
- Build a centralized Excel Dashboard.
- Understand practical applications of Excel Data Analytics.

---

## 🚀 Project Highlights

- 📊 300 sales orders analyzed
- 💰 Total Revenue: **306,976.92**
- 📦 Total Quantity Sold: **870**
- 🧹 Cleaned and structured sales dataset
- 🔎 VLOOKUP and INDEX-MATCH formulas
- 🧠 IF, AND and OR logical formulas
- 📈 Pivot Table analysis
- 📊 Bar, Pie and Line Charts
- 📋 Organized workbook structure
- 📑 Documentation sheet included
- 🎯 Sales and order performance analysis

---

## 🛠️ Technologies Used

- **Microsoft Excel**
- Excel Tables
- Excel Formulas
- VLOOKUP
- INDEX + MATCH
- IF
- AND
- OR
- Pivot Tables
- Charts
- Data Cleaning
- Data Visualization
- Dashboard Design

---

## 📂 Project Structure

```text
Sales & Data Analytics Dashboard
│
├── Raw Data
├── Cleaned Data
├── Pivot Table
├── Formulas
├── Visualization
├── Dashboard
└── Documentation
````

---

## 📘 Workbook Information

| Information    | Details                           |
| -------------- | --------------------------------- |
| Project Type   | Sales & Data Analytics            |
| Tool Used      | Microsoft Excel                   |
| Total Orders   | 300                               |
| Total Quantity | 870                               |
| Total Revenue  | 306,976.92                        |
| Main Analysis  | Sales, Orders, Products & Revenue |
| Visualization  | Bar Chart, Pie Chart, Line Chart  |
| Dashboard      | Excel Dashboard                   |
| Dataset        | Sales Order Dataset               |

---

# 📑 Worksheets

## 1️⃣ Raw Data

The **Raw Data** sheet contains the original sales order dataset.

It includes information such as:

* Order ID
* Date
* Customer ID
* Product
* Quantity
* Unit Price
* Shipping Address
* Payment Method
* Order Status
* Tracking Number
* Items in Cart
* Coupon Code
* Referral Source
* Total Price

The sheet contains **300 sales orders**.

---

## 2️⃣ Cleaned Data

The **Cleaned Data** sheet contains the organized and structured version of the sales dataset.

The data is stored in an Excel Table named:

```text
SalesData
```

This sheet is used as the main source for formulas, analysis, and visualizations.

---

## 3️⃣ Pivot Table

The **Pivot Table** sheet provides summarized business information.

The analysis includes:

### Product Analysis

* Product-wise Total Sales
* Product-wise Quantity
* Product-wise Order Count

### Order Status Analysis

* Number of orders by status
* Total revenue by order status

### Payment Method Analysis

* Revenue generated through different payment methods

### Referral Source Analysis

* Revenue generated through different referral sources

The overall analyzed revenue is:

**306,976.92**

---

## 4️⃣ Formulas

The **Formulas** sheet demonstrates the use of Excel formulas for different business requirements.

### 🔎 VLOOKUP

Used to find the Unit Price for a specific Order ID.

Example:

```excel
=VLOOKUP("ORD200000",SalesData[],6,FALSE)
```

### 🔎 INDEX + MATCH

Used to retrieve the Order Status using a Customer ID.

Example:

```excel
=INDEX('Cleaned Data'!I:I,MATCH("C72649",'Cleaned Data'!C:C,0))
```

### 🏷️ Discount Eligibility

Used to identify whether a discount was applied.

```excel
=IF('Cleaned Data'!L2<>"None","Discount Applied","No Discount")
```

### ⚡ Order Priority

Used to classify orders based on Total Price and Order Status.

```excel
=IF(AND('Cleaned Data'!N2>2000,'Cleaned Data'!I2="Shipped"),"High Value-In Transit",IF(OR('Cleaned Data'!I2="Cancelled",'Cleaned Data'!I2="Returned"),"Revenue Risk","Standard Order"))
```

---

## 5️⃣ Visualization

The **Visualization** sheet contains graphical representations of the sales data.

The project includes:

### 📊 Bar Chart

Shows **Total Price by Product**.

This helps compare the sales/revenue contribution of different products.

### 🥧 Pie Chart

Shows order-related distribution based on the Pivot Table analysis.

### 📈 Line Chart

Shows **Total Price over Date**.

This helps identify changes and trends in revenue over time.

---

## 6️⃣ Dashboard

The **Dashboard** sheet provides a centralized view of the major visualizations.

It contains the main charts used for presenting the sales analysis.

The dashboard helps users quickly understand:

* Product performance
* Sales trends
* Order information
* Revenue patterns

---

## 7️⃣ Documentation

The **Documentation** sheet explains the purpose and structure of each worksheet.

It provides a simple overview of:

* Sheet names
* Sheet purpose
* Data organization
* Analysis workflow
* Workbook structure

---

# 🧮 Excel Functions Covered

The project demonstrates the following Excel functions:

### Lookup Functions

* VLOOKUP
* INDEX
* MATCH

### Logical Functions

* IF
* AND
* OR

### Calculation

* Multiplication
* Aggregation
* Total calculations

---

# 📊 Formula Categories

## 🔍 Lookup & Reference

```text
VLOOKUP
INDEX
MATCH
```

Used to retrieve information from the sales dataset.

---

## 🧠 Logical Functions

```text
IF
AND
OR
```

Used to classify orders and identify specific business conditions.

---

## 💰 Sales Calculation

Total Price is calculated based on:

```text
Quantity × Unit Price
```

For example:

```text
5 × 570.62 = 2853.10
```

---

# 📈 Sales Analysis

The project analyzes sales performance using the following areas:

* Product-wise sales
* Quantity sold
* Number of orders
* Total revenue
* Order status
* Payment methods
* Referral sources

---

# 📦 Product Analysis

The dataset contains the following products:

| Product | Number of Orders |
| ------- | ---------------: |
| Printer |               52 |
| Chair   |               46 |
| Monitor |               43 |
| Tablet  |               43 |
| Desk    |               41 |
| Phone   |               40 |
| Laptop  |               35 |

This analysis helps identify the most frequently ordered products.

---

# 🚚 Order Status Analysis

The dataset contains five major order statuses:

| Order Status | Number of Orders |
| ------------ | ---------------: |
| Cancelled    |               73 |
| Shipped      |               59 |
| Delivered    |               58 |
| Pending      |               56 |
| Returned     |               54 |

This helps understand the overall order fulfillment situation.

---

# 💳 Payment Method Analysis

The project analyzes sales according to payment method.

| Payment Method | Number of Orders |
| -------------- | ---------------: |
| Online         |               73 |
| Credit Card    |               58 |
| Gift Card      |               57 |
| Cash           |               57 |
| Debit Card     |               55 |

This provides an overview of customer payment preferences.

---

# 📢 Referral Source Analysis

The project also analyzes how customers reached the business.

| Referral Source | Number of Orders |
| --------------- | ---------------: |
| Google          |               68 |
| Email           |               64 |
| Instagram       |               57 |
| Referral        |               56 |
| Facebook        |               55 |

This helps identify the major customer acquisition sources.

---

# 🎟️ Coupon Analysis

The dataset includes multiple coupon codes such as:

* SAVE10
* FREESHIP
* WINTER15
* None / No Coupon

Coupon information is used in the formula analysis to identify whether a discount was applied.

---

# 📊 Pivot Table Analysis

Pivot Tables are used to summarize large amounts of sales data.

The project uses Pivot Tables for:

* Product performance
* Quantity analysis
* Order count
* Order status
* Payment method
* Referral source
* Revenue analysis

Pivot Tables make it easier to identify important business patterns without manually calculating every value.

---

# 📉 Data Visualization

The project uses different chart types to make the analysis easier to understand.

### Bar Chart

Used for comparing product-wise Total Price.

### Pie Chart

Used for showing proportional order/revenue information.

### Line Chart

Used for showing Total Price trends across dates.

These visualizations make the final analysis more understandable and presentation-friendly.

---

# 📤 Project Output

The final workbook contains:

```text
Raw Data
       ↓
Cleaned Data
       ↓
Formulas
       ↓
Pivot Table
       ↓
Visualization
       ↓
Dashboard
       ↓
Documentation
```

The project provides a complete Excel-based workflow from raw sales data to business analysis and visualization.

---

# 📚 Learning Outcomes

Through this project, the following concepts were practiced:

* Excel Data Cleaning
* Excel Tables
* Lookup Functions
* Logical Functions
* Data Analysis
* Pivot Tables
* Data Visualization
* Dashboard Creation
* Business Data Interpretation
* Sales Performance Analysis

---

# 💡 Skills Demonstrated

### Technical Skills

* Microsoft Excel
* Excel Formulas
* VLOOKUP
* INDEX-MATCH
* IF
* AND
* OR
* Pivot Tables
* Excel Charts
* Data Cleaning
* Dashboard Creation

### Analytical Skills

* Sales Analysis
* Revenue Analysis
* Product Analysis
* Order Analysis
* Customer Data Analysis
* Payment Analysis
* Referral Source Analysis
* Business Data Interpretation

---

# 🏆 Project Achievements

* Successfully analyzed **300 sales orders**
* Calculated total revenue of **306,976.92**
* Analyzed **870 total quantities**
* Created a structured cleaned dataset
* Implemented multiple Excel formulas
* Created Pivot Table summaries
* Created multiple data visualizations
* Built a centralized Dashboard
* Documented the complete workbook structure

---

# 📋 Dataset Fields

The project uses the following fields:

```text
OrderID
Date
CustomerID
Product
Quantity
UnitPrice
ShippingAddress
PaymentMethod
OrderStatus
TrackingNumber
ItemsInCart
CouponCode
ReferralSource
TotalPrice
```

---

# 📊 Key Dataset Statistics

| Metric                     |      Value |
| -------------------------- | ---------: |
| Total Orders               |        300 |
| Total Quantity             |        870 |
| Total Revenue              | 306,976.92 |
| Number of Products         |          7 |
| Number of Order Statuses   |          5 |
| Number of Payment Methods  |          5 |
| Number of Referral Sources |          5 |

---

# 🔎 Business Analysis Areas

The project focuses on the following business questions:

### Product Performance

Which products generate more sales and orders?

### Order Performance

How many orders are Cancelled, Shipped, Delivered, Pending, or Returned?

### Payment Analysis

Which payment methods are used most frequently?

### Marketing Analysis

Which referral sources bring the most customers?

### Revenue Analysis

What is the overall revenue generated by the dataset?

### Order Classification

Which orders are high-value, revenue-risk, or standard orders?

---

# 🔄 Project Workflow

The complete project follows this workflow:

### Step 1 – Raw Data

Collect and organize the original sales order data.

### Step 2 – Data Cleaning

Prepare the dataset in a structured Excel Table.

### Step 3 – Formula Analysis

Use Excel formulas to perform lookups and classify orders.

### Step 4 – Pivot Table

Summarize important sales and order information.

### Step 5 – Visualization

Create charts from the analyzed data.

### Step 6 – Dashboard

Present the important visualizations in one place.

### Step 7 – Documentation

Document the workbook structure and purpose of each sheet.

---

# 📌 Excel Concepts Practiced

```text
✔ Excel Tables
✔ Data Cleaning
✔ Lookup Functions
✔ Logical Functions
✔ Conditional Analysis
✔ Pivot Tables
✔ Charts
✔ Dashboard
✔ Data Visualization
✔ Business Analysis
```

---

# 📝 How to Use

1. Open the Excel workbook.
2. Start from the **Raw Data** sheet.
3. Review the original sales dataset.
4. Open **Cleaned Data** to view the structured dataset.
5. Check the **Formulas** sheet for formula-based analysis.
6. Open **Pivot Table** to view summarized results.
7. Open **Visualization** to view charts.
8. Open **Dashboard** to view the final presentation.
9. Open **Documentation** to understand the workbook structure.

---

# 📌 Project Summary

The **Sales & Data Analytics Dashboard** is a practical Microsoft Excel project that demonstrates how raw sales data can be transformed into meaningful business insights.

The project combines:

```text
Data Cleaning
     +
Excel Formulas
     +
Pivot Tables
     +
Data Visualization
     +
Dashboard
     =
Sales & Data Analytics Solution
```

This project demonstrates practical knowledge of Excel Data Analysis and provides a structured approach for understanding sales performance, revenue, products, orders, payment methods, and referral sources.

---

# 👨‍💻 Author

**Sarth Thakar**

```
```
