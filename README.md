SALES & DATA ANALYTICS DASHBOARD

1. PROJECT OVERVIEW

The Sales & Data Analytics Dashboard is an Excel-based data analysis project designed to analyze and understand sales order data. The project uses Microsoft Excel to clean raw data, perform calculations, analyze sales information using Pivot Tables, create visualizations, and develop an interactive dashboard.

The workbook contains 300 sales records with information such as Order ID, Date, Customer ID, Product, Quantity, Unit Price, Shipping Address, Payment Method, Order Status, Tracking Number, Items in Cart, Coupon Code, Referral Source, and Total Price.

The main purpose of this project is to convert raw sales data into meaningful information that can help understand sales performance, customer orders, product performance, payment methods, order statuses, and referral sources.

2. PROJECT OBJECTIVES

The main objectives of this project are:

• To organize and analyze sales order data using Microsoft Excel.
• To clean and prepare the raw dataset for analysis.
• To calculate important sales-related values.
• To analyze product-wise sales and quantities.
• To analyze order status and revenue.
• To analyze payment methods and referral sources.
• To use Excel formulas for data lookup and classification.
• To create Pivot Tables for summarized analysis.
• To create charts and visualizations for better understanding.
• To develop an interactive sales dashboard.
• To present business information in a simple and understandable format.

3. TECHNOLOGIES AND TOOLS USED

Microsoft Excel

The complete project is developed using Microsoft Excel.

Excel Features Used:

• Excel Tables
• Data Cleaning
• Excel Formulas
• VLOOKUP
• INDEX
• MATCH
• IF
• AND
• OR
• Pivot Tables
• Charts
• Data Visualization
• Dashboard
• Slicers
• Timeline Controls
• Conditional Analysis
• Data Summarization

4. DATASET INFORMATION

The dataset contains 300 sales order records.

The main fields available in the dataset are:

OrderID – Unique identification number for each order.

Date – Date on which the order was placed.

CustomerID – Unique identification number of the customer.

Product – Product purchased by the customer.

Quantity – Number of units purchased.

UnitPrice – Price of one unit of the product.

ShippingAddress – Shipping location of the order.

PaymentMethod – Payment method used by the customer.

OrderStatus – Current status of the order.

TrackingNumber – Tracking information associated with the order.

ItemsInCart – Number of items present in the customer's cart.

CouponCode – Coupon or promotional code used by the customer.

ReferralSource – Source through which the customer reached the platform.

TotalPrice – Total price of the order.

5. WORKBOOK STRUCTURE

The workbook contains seven main worksheets.

5.1 RAW DATA

The Raw Data sheet contains the original sales dataset consisting of 300 records.

This sheet is used as the starting point for the project. It contains the original order information before the data preparation and analysis process.

5.2 CLEANED DATA

The Cleaned Data sheet contains the prepared sales dataset.

The data is organized into an Excel Table named SalesData, making it easier to use formulas, Pivot Tables, and other Excel analysis features.

This sheet is used as the main source for analysis.

5.3 PIVOT TABLE

The Pivot Table sheet contains summarized information generated from the sales dataset.

The analysis includes:

• Product-wise Total Price
• Product-wise Quantity
• Product-wise Order Count
• Order Status-wise Order Count
• Order Status-wise Total Price
• Payment Method and Referral Source analysis
• Overall sales summary

The Pivot Tables help convert the detailed 300-row dataset into meaningful summaries.

5.4 FORMULAS

The Formulas sheet demonstrates different Excel functions used for analysis.

The main calculations include:

• Unit Price Lookup
• Order Status by Customer ID
• Discount Eligibility
• Order Priority

These formulas demonstrate how Excel functions can be used to retrieve information and classify sales orders.

5.5 VISUALIZATION

The Visualization sheet contains individual charts and formatted visual representations of the sales data.

Charts are used to make the analysis easier to understand and to identify differences between products, order statuses, and other sales categories.

5.6 DASHBOARD

The Dashboard sheet provides a visual summary of the sales analysis.

It contains charts, KPI-style summaries, slicers, and timeline controls that allow users to interact with the data and understand the overall sales performance.

The dashboard provides a simple way to view important information without manually analyzing the complete dataset.

5.7 DOCUMENTATION

The Documentation sheet provides information about the workbook structure and the purpose of the different sheets used in the project.

It helps users understand how the Excel workbook is organized.

6. DATA CLEANING

The raw sales data was prepared and organized before performing analysis.

The cleaned dataset was converted into an Excel Table named SalesData. Using an Excel Table makes it easier to reference the data in formulas and Pivot Tables.

The cleaned data contains consistent fields for orders, customers, products, quantities, prices, payment methods, order statuses, coupon codes, and referral sources.

7. DATA ANALYSIS

The project performs different types of sales analysis.

Product Analysis:

Product-wise quantity and total revenue are analyzed to understand which products contribute more to sales.

Order Status Analysis:

Orders are categorized into different statuses such as Cancelled, Delivered, Pending, Returned, and Shipped. The number of orders and total revenue associated with each status are analyzed.

Payment Method Analysis:

Different payment methods such as Cash, Credit Card, Debit Card, Gift Card, and Online are analyzed to understand customer payment preferences.

Referral Source Analysis:

Referral sources such as Email, Facebook, Google, Instagram, and Referral are analyzed to understand how customers reach the sales platform.

Overall Sales Analysis:

The Pivot Table provides an overall summary of 300 orders with a total order value of 306,976.92.

8. FORMULAS USED

VLOOKUP

VLOOKUP is used to find the Unit Price of a specific order.

Example:

=VLOOKUP("ORD200000",SalesData[],6,FALSE)

INDEX AND MATCH

INDEX and MATCH are used together to find the Order Status associated with a particular Customer ID.

Example:

=INDEX('Cleaned Data'!I:I,MATCH("C72649",'Cleaned Data'!C:C,0))

IF FUNCTION

The IF function is used to determine whether a discount or coupon has been applied.

Example:

=IF('Cleaned Data'!L2<>"None","Discount Applied","No Discount")

AND, OR AND IF

These functions are combined to classify orders according to their value and status.

Example:

=IF(AND('Cleaned Data'!N2>2000,'Cleaned Data'!I2="Shipped"),"High Value-In Transit",IF(OR('Cleaned Data'!I2="Cancelled",'Cleaned Data'!I2="Returned"),"Revenue Risk","Standard Order"))

9. PIVOT TABLE ANALYSIS

Pivot Tables are used to summarize large amounts of sales data quickly.

The project uses Pivot Tables to analyze:

• Product performance
• Quantity sold
• Total sales value
• Number of orders
• Order status
• Payment methods
• Referral sources

Pivot Tables make it easier to identify patterns and compare different categories.

10. DATA VISUALIZATION

Charts are created to present the analyzed data visually.

The project includes visualizations that help users understand:

• Product performance
• Sales distribution
• Order status
• Revenue performance
• Customer and order-related information

Charts make the information easier to understand than viewing raw numbers alone.

11. INTERACTIVE DASHBOARD

The Dashboard is the main visual output of the project.

It combines important sales information into a single interface.

The dashboard includes:

• KPI summaries
• Sales charts
• Product analysis
• Order analysis
• Slicers
• Timeline controls
• Visual summaries

Slicers and Timeline controls allow users to filter the dashboard and explore the data interactively.

12. KEY CALCULATIONS

Total Price is calculated based on the quantity and unit price of an order.

Total Price = Quantity × Unit Price

For example, if an order contains 5 units and the unit price is 570.62:

Total Price = 5 × 570.62

Total Price = 2853.10

This calculation is used to determine the value of individual orders and is later summarized using Pivot Tables.

13. EXCEL CONCEPTS COVERED

The project demonstrates practical knowledge of:

• Data Cleaning
• Excel Tables
• Data Organization
• VLOOKUP
• INDEX and MATCH
• IF Function
• AND Function
• OR Function
• Conditional Classification
• Pivot Tables
• Data Summarization
• Charts
• Data Visualization
• Dashboard Creation
• Slicers
• Timeline Controls
• Sales Analysis

14. KEY FINDINGS

The analysis provides a summarized view of the sales dataset.

The workbook contains 300 orders with an overall Total Price of 306,976.92.

The analysis also shows differences in product performance, order status, payment methods, and referral sources.

The Pivot Tables and dashboard make it easier to compare these categories and identify important patterns in the sales data.

15. LEARNING OUTCOMES

Through this project, the following skills were developed:

• Understanding of Excel-based data analysis.
• Understanding of data cleaning and preparation.
• Practical use of Excel formulas.
• Understanding of lookup functions.
• Experience with Pivot Tables.
• Experience creating charts and visualizations.
• Understanding of dashboard design.
• Ability to summarize large datasets.
• Ability to present data in a meaningful way.
• Improved analytical and problem-solving skills.

16. SKILLS DEMONSTRATED

The project demonstrates the following skills:

Microsoft Excel
Data Analysis
Data Cleaning
Data Visualization
Pivot Table Analysis
Excel Formula Development
Dashboard Development
Business Data Interpretation
Reporting
Analytical Thinking

17. PROJECT OUTPUT

The final workbook provides a complete sales analysis solution starting from raw data and ending with an interactive dashboard.

The workflow of the project is:

Raw Data → Cleaned Data → Formulas → Pivot Table Analysis → Visualization → Dashboard

This structure demonstrates how raw business data can be transformed into useful analytical information using Microsoft Excel.

18. HOW TO USE THE WORKBOOK

Step 1: Open the Excel workbook.

Step 2: Review the Raw Data sheet to understand the original dataset.

Step 3: Open the Cleaned Data sheet to view the prepared dataset.

Step 4: Review the Formulas sheet to understand the calculations and lookup functions.

Step 5: Open the Pivot Table sheet to view summarized analysis.

Step 6: Open the Visualization sheet to view individual charts.

Step 7: Open the Dashboard sheet to view the final interactive dashboard.

Step 8: Use the available slicers and timeline controls to filter and explore the dashboard.

19. PROJECT SUMMARY

The Sales & Data Analytics Dashboard project demonstrates how Microsoft Excel can be used as a complete data analysis and reporting tool.

Starting with 300 raw sales records, the project organizes and prepares the data, performs formula-based calculations, creates summarized Pivot Table reports, develops visualizations, and combines important information into an interactive dashboard.

The project provides practical experience in Excel data analysis, visualization, reporting, and dashboard development.

20. AUTHOR

Sarth Thakar

Project: Sales & Data Analytics Dashboard

Tool Used: Microsoft Excel
