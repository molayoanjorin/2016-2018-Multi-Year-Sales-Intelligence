# Project Overview

This project focuses on transforming three years of raw sales data (2016–2018) into a consolidated Business Intelligence and Decision Support System using Microsoft Excel, Power Query and Power Pivot.

The objective was to develop a reporting solution that enables management to monitor business performance, evaluate product and customer contributions, analyse sales trends and identify areas requiring attention.

The project covers the complete analytical process, from data preparation and integration to business performance analysis and executive reporting.

# Business Problem

The organisation maintained its sales records in separate annual datasets for 2016, 2017 and 2018. These records needed to be consolidated before meaningful multi-year analysis could be performed.

The project addressed the following questions:

How is the business performing across the three years?

How have revenue and gross profit changed over time?

Which products contribute most to business performance?

Which customers account for the largest share of revenue?

Which sales channels generate the most revenue?

What monthly and quarterly revenue patterns can be identified?

How can management use these findings to make better business decisions?

# Tools and Technologies

### Microsoft Excel
Data analysis, PivotTables, PivotCharts and executive dashboard development.

### Power Query

Data cleaning, standardisation, transformation and dataset integration.

### Power Pivot

Data modelling and analytical calculations.

### DAX

Business KPIs, year-on-year calculations, product performance scoring and contribution analysis.

# Dataset Overview

The project used three annual sales datasets: Sales2016, Sales2017 and Sales2018. Each dataset contained 3,134 records, giving a combined total of 9,402 records.

The datasets contained 13 columns: OrderNumber, OrderDate, Ship Date, Customer ID, Channel, Currency Code, Warehouse Code, Delivery Region ID, Product ID, Order Quantity, Unit Price, Total Unit Cost and Total Revenue.

The data contained five currencies: USD, NZD, AUD, EUR and GBP.

Important financial reporting consideration: The currencies were not converted into a common currency. Financial comparisons and totals were therefore analysed using individual currency filters rather than combining different currencies into one financial figure.

# Data Engineering and Preparation

## Data Importation

The three annual datasets were imported into Power Query as separate queries. The original datasets were retained, and cleaning and transformation processes were applied before integration.

## Data Cleaning and Standardisation

### Trim
Removed unnecessary leading/trailing spaces from OrderNumber.

### Proper Case

Standardised the Channel column.

### Uppercase

Standardised Currency Code, and Warehouse Code.

### Data Type Conversion

Assigned appropriate data types to date, text and numeric columns.

## Data Quality Checks

Power Query's data profiling and duplicate-checking features were used to inspect the datasets. The checks showed 0% empty values, 0% errors and no duplicate records across all columns. No records were removed during the cleaning process.

## Dataset Integration

The three cleaned annual datasets were combined using Append Queries as New in Power Query. The resulting query was named Consolidated Sales Data. The consolidated dataset contained 9,402 records, with 3,134 records from each year.

## Date Transformation

Additional date-related columns were created from OrderDate to support time-based analysis: Year, Month Name, Month Number, Quarter, Day of Week and Year-Month. The consolidated data was also prepared for use in the Power Pivot Data Model.

# Business Performance KPIs

Sales Revenue

Total Orders

Total Units Sold

Average Order Value

Total Cost

Gross Profit

Gross Profit Margin

# Year-on-Year Performance Analysis

The year-on-year analysis examined annual revenue and gross profit performance across 2016, 2017 and 2018. The analysis included Previous Year Revenue, Annual Revenue Growth, Previous Year Gross Profit and Annual Gross Profit Growth.

### Key Findings

Revenue for 2017 and 2018 was 60,137,900.20.

Gross profit for 2017 and 2018 was 22,477,296.88.

Revenue growth between 2017 and 2018 was 0%.

Gross profit growth between 2017 and 2018 was 0%.

Growth for 2016 could not be calculated because 2015 data was not available.

Business interpretation: The identical revenue and gross profit figures recorded in 2017 and 2018 indicate no year-on-year growth between those two years. Management should investigate the factors behind this lack of growth and identify opportunities to improve sales performance and profitability.

# Product Performance Index

A custom Product Performance Index (PPI) was developed to rank products based on their contribution to business performance.

Revenue Contribution = 50%

Units Sold Contribution = 30%

Gross Profit Contribution = 20%

Formula: Product Performance Index = (Revenue Contribution * 50%) + (Units Sold Contribution * 30%) + (Gross Profit Contribution * 20%).

### Business Interpretation
Revenue received the highest weight because it represents the product's contribution to sales. Units sold measures sales volume and provides an indication of product demand, while gross profit contribution considers the product's contribution to profit after costs.

The index provides management with a structured way to compare products across these three measures. A higher score indicates a stronger overall contribution across the selected metrics. It is a contribution-based ranking measure, not a standalone measure of product margin or product quality.

# Pareto Analysis

Pareto analysis was conducted to examine how revenue was distributed across products and customers, using revenue contribution and cumulative revenue contribution.

### Product Pareto Analysis

50% of products contributed 84.76% of revenue for the selected currency.

### Customer Pareto Analysis
76% of customers contributed 80.42% of revenue for the selected currency.

### Business Interpretation

The results help management understand the distribution of revenue across products and customers. Product analysis can support product prioritisation and inventory decisions, while customer analysis can help management identify important customer segments and develop retention strategies.

Note: These percentages relate to the selected currency and filter context. They should not be interpreted as combined figures across all five currencies.


# Executive Dashboard

An interactive, one-page executive dashboard was developed in Microsoft Excel.

### KPI Cards

Total Revenue

Total Orders

Average Order Value

Total Units Sold

Gross Profit

Gross Profit Margin

## Dashboard Visualisations

Monthly Revenue Trend

Yearly Performance and Growth

Top 10 Products by Revenue

Top 10 Delivery Region IDs by Revenue

Top 10 Customers by Revenue

Revenue Contribution by Sales Channel

## Dashboard Interactivity

The dashboard includes slicers for Year, Month Name, Product ID, Delivery Region ID, Channel and Currency Code. The slicers were connected to the relevant PivotTables and PivotCharts to support interactive analysis.

# Key Business Recommendations

1. Prioritise Product Performance Management

Management should use the Product Performance Index and Pareto analysis to identify products with stronger contributions and guide inventory, marketing and sales decisions.

2. Review Export Channel Performance

Since Export recorded the lowest revenue across all five currencies, management should investigate the factors affecting its performance, including customer demand, pricing and distribution.

3. Address the Lack of Annual Growth

Revenue and gross profit remained unchanged between 2017 and 2018. Management should investigate the reasons behind this stagnation and explore opportunities to improve sales performance and customer demand.

# Data Limitations and Considerations

The dataset contains five currencies: USD, NZD, AUD, EUR and GBP.

Currency conversion was not performed. Financial comparisons were made within individual currency filters.

Geographic analysis was conducted using Delivery Region ID. The dataset did not provide a verified mapping to named territories or cities.

Previous-year growth for 2016 could not be calculated because 2015 data was unavailable.

The Product Performance Index is a weighted contribution measure and should be interpreted alongside other business considerations.

# Technical Challenge

During the stress-test process, an issue occurred when connecting the existing Power Query process to the updated source file. The original query expected an Excel Table named Sales2018, while the updated source contained the data as a worksheet. This caused a navigation error.

The issue was resolved by creating raw Excel Tables inside the workbook and changing each query's source to Excel.CurrentWorkbook(){{Name="Raw_2018"]}[Content].

Tools used: Advanced Editor, Excel Tables, Queries & Connections, Refresh All.

# How to Use This Project

1. Download this workbook from the Excel/ folder
2. Open it in desktop Excel (Windows, Microsoft 365 or a version with Power Query and Power Pivot).
3. If prompted, click Enable Content.
4. To test automation:
   Add new rows to the bottom of a raw table (Raw_2016, Raw_2017, Raw_2018).
   Click Data, then Refresh All, wait for it to finish then click it once more.
   Check that the consolidated table, KPI cards, charts and slicers have updated.

 # Author
 
 ### Anjorin Molayo.

 Note: This is a project based analysis for my 1st semester exam at Alt School, Karatu Cohort.













