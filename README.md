# 📊 Excel-Based Sales Data Analysis

> Analysing real-world sales data using Microsoft Excel — pivot tables, charts, and business insights.

---

## Project Overview
This project focuses on analysing a real-world sales dataset using **Microsoft Excel**. The goal was to clean, explore, and visualise the data in order to extract meaningful business insights — covering top-selling products, regional performance, monthly sales trends, and customer value.

This project is part of my learning journey in data analysis using Excel, SQL, Power BI, and Python.

---

## 🎯 Project Objectives
- Clean and prepare raw sales data for analysis
- Analyse sales performance across products and regions
- Identify top-performing products and customers
- Explore monthly and yearly sales trends
- Build pivot tables and charts to visualise key insights

---

## 📂 Dataset Overview

| Detail | Info |
|--------|------|
| Source | Sample Sales Transaction Data |
| Period Covered | 2003 – 2005 |
| Total Records | 2,823 rows (raw) |
| Countries | 19 |
| Product Lines | 7 |
| Key Columns Used | Order Number, Total Sales, Order Date, Product Line, Country, Territory, Customer Name, Deal Size, Status |

---

## 🧹 Data Cleaning Steps

The raw dataset had **25 columns**. After review, **7 columns were dropped** as they were not relevant to the project objectives:

> `QTR_ID`, `PRODUCTCODE`, `PHONE`, `CITY`, `STATE`, `CONTACTLASTNAME`, `CONTACTFIRSTNAME`

The following cleaning steps were carried out on the **Cleaned Data** sheet (Raw Data was preserved untouched):

1. **Removed duplicate rows** using Excel's Remove Duplicates tool (key: ORDERNUMBER)
2. **Filled blank TERRITORY cells** — USA rows had no territory assigned; filled with `NA/US`
3. **Formatted ORDERDATE** column as proper Date format
4. **Formatted TOTALSALES and PRICEEACH** as Currency
5. **Validated STATUS values** to ensure no typos or unexpected entries
6. **Converted data to an Excel Table** named `SalesData` for structured referencing

**Four helper columns were added:**

| Column | Formula Used | Purpose |
|--------|-------------|---------|
| MONTH-YEAR | `=TEXT([@ORDERDATE],"MMM-YYYY")` | Readable date label for trend charts |
| REVENUE TIER | `=IF([@TOTALSALES]>=5000,"High",IF([@TOTALSALES]>=2000,"Medium","Low"))` | Classifies orders by revenue size |
| FULFILLED | `=IF([@STATUS]="Shipped","Yes","No")` | Flags successfully shipped orders |
| EST. PROFIT | `=[@TOTALSALES]*0.3` | Estimates profit at 30% margin |

---

## 📈 Pivot Tables Built

| # | Pivot Table | Rows | Columns | Values |
|---|------------|------|---------|--------|
| 1 | Sales by Product Line & Year | PRODUCTLINE | YEAR_ID | Sum of TOTALSALES |
| 2 | Monthly Sales Trend | MONTH_ID | YEAR_ID | Sum of TOTALSALES |
| 3 | Sales by Country (Top 10) | COUNTRY | — | Sum of TOTALSALES |
| 4 | Top 10 Customers | CUSTOMERNAME | — | Sum of TOTALSALES |
| 5 | Sales by Deal Size | DEALSIZE | YEAR_ID | Sum of TOTALSALES |

---

## 📉 Charts Created

| # | Chart Type | Based On | Insight Shown |
|---|-----------|----------|--------------|
| 1 | Clustered Column Chart | Pivot 1 | Revenue comparison across product lines per year |
| 2 | Line with Markers | Pivot 2 | Monthly sales trend across 2003–2005 |
| 3 | Horizontal Bar Chart | Pivot 3 | Top 10 countries ranked by revenue |
| 4 | Pie Chart | Pivot 5 | Revenue split by deal size |

## 📊 Power BI Dashboard Preview

![Dashboard](dashboard_screenshot.png)
> 💡 Download the `Sales_Performance_Dashboard.pbix` 
file above and open in Power BI Desktop to interact 
with the full dashboard including slicers and filters.
---

## 💡 Key Findings & Business Insights

#### 1. 🏆 Classic Cars Dominate Revenue
Classic Cars generated **$2,968,546.40** — accounting for over **35.8%** of total revenue. Vintage Cars came second at **$1,644,212.05 (19.83%)**, followed by Motorcycles at **$971,086.29 (11.71%)**. Trains was the weakest performer, contributing less than **3%**. The business is heavily dependent on one product line, which is a risk worth monitoring.

#### 2. 🌍 USA Leads Regional Performance
The USA was the single largest market with **$2,986,425.21 in revenue (36.02% of total)**, followed by Spain (**$1,021,705.97**), France (**$919,257.85**), Australia (**$521,598.46**), and the UK (**$413,203.34**). At the territory level, EMEA led with **49.79%**, just ahead of North America (NA) at **38.35%**.

#### 3. 📅 November is the Strongest Sales Month
November recorded the highest monthly revenue at **$1,744,682.45** — almost double any other month. October (**$919,036.70**) and May (**$756,812.91**) were the next strongest. This suggests a strong seasonal pattern likely driven by holiday buying behaviour. Businesses should plan inventory and marketing campaigns ahead of Q4.

#### 4. 📆 2004 Was the Best Performing Year
Sales grew from **$2,898,149.94 in 2003** to **$3,913,700.87 in 2004** — a growth of **35.04%**. The 2005 figure of **$1,479,035.98** appears lower but covers only part of the year (January to May), so it is not a fair comparison to full years.

#### 5. 💼 Medium Deals Drive the Business
Medium-sized deals accounted for **59.8% of total revenue ($4,961,736.68)**, making them the backbone of the business. Small deals contributed **31.2% ($2,590,392.20)** and Large deals **8.9% ($738,757.91)**. This tells us the business thrives on mid-range transactions and the contribution of Small deals cannot be overlooked as well.

#### 6. 👥 Top Customers Are Highly Concentrated
The top customer — **Euro Shopping Channel** — generated **$766,195.05** alone, representing **31.41% of total revenue**. The second customer, **Mini Gifts Distributors Ltd.**, contributed **$530,587.19 (21.75%)**. The top 5 customers together account for nearly **72.44% of all revenue** — a concentration risk the business should be aware of.

#### 7. ✅ Order Fulfilment Rate is Strong
**92.26% of all orders were successfully shipped** — indicating a reliable fulfilment operation. This is a positive signal for customer satisfaction.

---

## 🗂️ Workbook Structure

| Sheet | Description |
|-------|-------------|
| Raw_Data | Original dataset — untouched |
| Cleaned_Data | Cleaned and structured data with helper columns |
| Pivot-ProductLine | Sales by product line & year |
| Pivot-Monthly | Monthly sales trend by year |
| Pivot-Country | TOP 10 countries by revenue |
| Pivot-Customer | Top 10 customers by revenue |
| Pivot-DealSize | Sales by deal size & year |

---

## 🔧 Tools Used
- **Microsoft Excel** — Data cleaning, pivot tables, charts
- **Power BI** — Interactive dashboard, DAX measures, slicers

---

## 🚀 What's Next
-**The journey continues! I have just completed the IDA/3MTT Data Analysis Bootcamp — a three-month intensive program where I built three real-world capstone projects across retail, healthcare, and personal finance, analysing over 200,000 records using Python, SQL, Microsoft Excel, and Power BI. Full project documentation and files coming to this repository soon.

---

*⭐ Feel free to explore the workbook and leave a star if this was helpful!*
