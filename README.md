#  Retail Sales Data Analysis (Excel)

A beginner-to-intermediate level Excel project that cleans, analyzes, and derives insights from a retail sales dataset using formulas, logical functions, and statistical analysis — built as part of my Data Analyst portfolio.

## Project Overview

This project works with 500 retail transaction records for an electronics retailer, covering order details, products, pricing, discounts, and regional sales. The goal was to practice real-world Excel skills: handling missing/zero values, building dynamic lookup-style formulas, and summarizing data using statistical functions — all without hardcoding any results.

## Dataset

**File:** `Retail_Sales_project.xlsx`
**Rows:** 500 transactions | **Columns:** 14

| Column | Description |
|---|---|
| Order ID | Unique identifier for each order |
| Customer Name | Name of the customer |
| Date | Order date |
| Product Name | Item purchased (Headphones, Charger, Tablet, Phone, Laptop) |
| Category | Product category |
| Brand | Product brand (Apple, HP, Dell, Samsung, Lenovo) |
| Quantity Sold | Units sold in the order |
| Unit Price | Listed price per unit |
| Total Sales | Quantity Sold × Unit Price |
| Discount | Discount percentage applied |
| Region | Sales region (North, South, East, West) |
| New Unit Price | Corrected price (see formulas below) |
| Mode | Most frequently occurring quantity sold |
| Median | Median quantity sold |

## Key Formulas & Logic

- **Missing Price Correction** — Some records had a `Unit Price` of 0. An `IF` + `AVERAGEIFS` formula fills these gaps by calculating the average price for the same **Brand + Product + Category** combination, so no sale is ever priced at zero:
  ```excel
  =IF(H1=0, AVERAGEIFS(H:H, F:F, F1, D:D, D1, E:E, E1, H:H, "<>0"), H1)
  ```
- **Mode of Quantity Sold** — Identifies the most commonly sold quantity per order using `MODE.SNGL`.
- **Median of Quantity Sold** — Calculates the median using `MEDIAN`, giving a robust central tendency measure unaffected by outliers.
- **Structured Table** — The full dataset is formatted as an Excel Table with headers, autofilters, and banded rows for easy sorting and filtering.

## Key Insights

- **Total revenue:** ₹7,24,000 across 500 orders (Jan 2023 – May 2024)
- **Top-selling region:** North (₹2,90,500), followed by East
- **Top product:** Headphones and Laptops tie for the highest total sales (₹1,62,000 each)
- **Top brand by revenue:** Apple (₹1,66,000), closely followed by HP and Lenovo
- **Regions covered:** North, South, East, West — with visible variance in performance, useful for regional strategy discussions

## Tools Used

- Microsoft Excel
- Logical Functions: `IF`
- Lookup/Aggregation: `AVERAGEIFS`
- Statistical Functions: `MODE.SNGL`, `MEDIAN`
- Excel Tables & AutoFilter

- Add Pivot Tables and charts for regional/brand-wise visual comparisons
- Build a summary dashboard sheet with KPIs (Total Sales, Avg Order Value, Top Region)
- Extend analysis with Power Query for data cleaning at scale

