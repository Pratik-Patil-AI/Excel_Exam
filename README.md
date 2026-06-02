# Apex Product Analytics Dashboard

## Overview

Apex Product Analytics Dashboard is an interactive Excel dashboard built using Online Sales Data. The project demonstrates advanced Excel skills including data analysis, Pivot Tables, Pivot Charts, KPI reporting, lookup functions, logical functions, conditional formatting, and dashboard design.

The dashboard provides insights into product performance, regional sales distribution, customer payment preferences, and revenue trends through dynamic visualizations and interactive filters.

---

## Dataset Information

The dataset contains 240 online sales transactions with the following fields:

* Transaction ID
* Date
* Product Category
* Product Name
* Units Sold
* Unit Price
* Total Revenue
* Region
* Payment Method

---

## Project Objectives

* Analyze sales performance across different product categories.
* Identify high-performing regions.
* Monitor revenue trends over time.
* Compare customer payment preferences.
* Create an interactive and dynamic Excel dashboard.
* Demonstrate advanced Excel analytical techniques.

---

## Excel Skills Demonstrated

### Lookup Functions

#### VLOOKUP

Used to retrieve Product Category based on Product Name.

```excel
=VLOOKUP(A2,Lookup_Table,2,FALSE)
```

#### INDEX-MATCH

Used as an alternative to VLOOKUP for flexible lookups.

```excel
=INDEX(Category_Column,MATCH(Product_Name,Product_Column,0))
```

---

### Logical Functions

#### IF Function

```excel
=IF(TotalRevenue>=1000,"High","Low")
```

Used to classify transactions based on revenue.

#### Nested IF Function

```excel
=IF(TotalRevenue>=2000,"A",
IF(TotalRevenue>=1000,"B",
IF(TotalRevenue>=500,"C","D")))
```

Used to assign performance grades.

#### AND + IF Function

```excel
=IF(AND(TotalRevenue>=1000,TotalRevenue<=3000),
"Target Achieved",
"Below Target")
```

Used to evaluate business targets.

---

### Conditional Formatting

Applied to visually identify:

* High Revenue Transactions
* Revenue Performance Grades
* Target Achievement Status
* Data Patterns and Outliers

---

## Key Performance Indicators (KPIs)

The dashboard highlights important business metrics including:

### Total Revenue

Measures overall sales generated.

### Total Units Sold

Shows total product demand.

### Total Transactions

Displays the number of completed transactions.

### Average Revenue Per Transaction

Measures average transaction value.

### Product Category Count

Shows product portfolio diversity.

### Top Performing Region

Identifies the region generating maximum revenue.

---

## Pivot Table Analysis

### Revenue by Product Category

Analyzes contribution of each product category to total sales.

### Revenue by Region

Compares regional sales performance.

### Revenue by Payment Method

Evaluates customer payment preferences.

### Monthly Revenue Trend

Tracks revenue movement over time.

### Units Sold by Category

Measures product demand across categories.

### Top Products by Revenue

Identifies best-performing products.

---

## Dashboard Features

* Interactive Slicers
* Dynamic Pivot Tables
* Dynamic Pivot Charts
* KPI Cards
* Revenue Trend Analysis
* Category Performance Analysis
* Regional Performance Analysis
* Payment Method Insights
* Professional Dashboard Layout
* Automated Data Refresh

---

## Dynamic Dashboard Design

The dashboard is built using Excel Tables and Pivot Tables to ensure dynamic updates.

Features include:

* Automatic data expansion
* Refreshable Pivot Tables
* Interactive filtering
* Linked charts and KPIs
* Scalable reporting structure

---

## Tools Used

* Microsoft Excel 2021
* Pivot Tables
* Pivot Charts
* Slicers
* Conditional Formatting
* VLOOKUP
* INDEX-MATCH
* IF Functions
* Nested Functions
* Dashboard Design Techniques

---

## Business Insights Generated

* Electronics generated the highest revenue contribution.
* Regional performance varies significantly across markets.
* Customer payment behavior can be analyzed through payment method distribution.
* Revenue trends help identify seasonal patterns and sales opportunities.
* Top-performing products contribute a significant portion of total revenue.

---

## Dashboard Preview

The dashboard includes:

* KPI Summary Cards
* Revenue Trend Visualization
* Product Category Analysis
* Regional Sales Analysis
* Payment Method Breakdown
* Interactive Slicers

---

## Author

Patil Pratik

Excel Dashboard & Data Analytics Project demonstrating advanced reporting, business intelligence, and dashboard development techniques using Microsoft Excel.
