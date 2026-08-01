# 📊 Superstore Sales Analytics Dashboard (Excel)

An interactive sales dashboard built in Excel using the [Sample Superstore dataset](https://www.kaggle.com/datasets/vivek468/superstore-dataset-final) — covering data cleaning, pivot tables, KPI reporting, and slicer-driven interactivity.

![Excel](https://img.shields.io/badge/Excel-217346?style=flat&logo=microsoft-excel&logoColor=white)

## 🔍 Overview

This project takes a raw 9,994-row retail transactions dataset and turns it into a single-page, interactive dashboard — with live KPIs, four connected charts, and filters that update everything at once.

## 📁 Dataset

- **Source:** Sample Superstore dataset (US retail orders, 2014–2017)
- **Rows:** 9,994 order line items
- **Columns:** 21 (Order/Ship dates, Customer, Region, Category, Sub-Category, Sales, Profit, Discount, Quantity, etc.)

## 🧹 Data Cleaning (Power Query)

- Verified and corrected data types across all columns (dates, currency, whole numbers)
- Audited for duplicates, blank/null values, and inconsistent categorical values — dataset came back clean, confirmed via Excel's Column Quality view
- **Fixed a Postal Code truncation bug:** zip codes were stored as numbers, silently dropping leading zeros (e.g. `02116` → `2116`) on 449 records. Corrected using `Text.PadStart()` to restore proper 5-digit formatting
- Ran business-logic sanity checks: no negative Sales/Quantity, no Order Date later than Ship Date, Discount range validated (0–80%)
- Loaded as a structured **Excel Table** (`SalesData`) for dynamic, self-expanding ranges

## 📈 Pivot Tables & Charts

| Pivot Table | Chart Type | Insight |
|---|---|---|
| `PT_SalesByRegion` | Clustered Column | Sales & Profit by Region |
| `PT_SalesByCategory` | Bar | Sales by Category & Sub-Category |
| `PT_Top10Customers` | Bar | Top 10 Customers by Sales |
| `PT_MonthlyTrend` | Line | Monthly Sales & Profit trend, 2014–2017 |

## 🎛️ Interactivity

- **Slicers:** Region, Category, Segment
- **Timeline:** filter by Order Date (Month/Quarter/Year)
- All slicers and the timeline are connected to all four pivot tables via Report Connections, so every chart updates together on a single click

## 🧮 KPI Cards

| KPI | Formula logic | Value |
|---|---|---|
| Total Sales | `SUM(SalesData[Sales])` | $2,297,200.86 |
| Total Profit | `SUM(SalesData[Profit])` | $286,397.02 |
| Total Orders | `SUMPRODUCT(1/COUNTIF(...))` — counts unique Order IDs, not line items | 5,009 |
| Avg Order Value | Total Sales ÷ Total Orders | $458.61 |
| Profit Margin | Total Profit ÷ Total Sales | 12.47% |

## 🛠️ Skills Demonstrated

- Power Query (data import, type correction, custom columns)
- Data cleaning & validation (nulls, duplicates, categorical consistency, business-logic checks)
- Excel Tables (structured, dynamic ranges)
- Pivot Tables & Pivot Charts
- Slicers & Timelines (cross-filtering multiple views)
- KPI design with SUMIFS/SUMPRODUCT/COUNTIF formulas
- Dashboard layout & design (single-page reporting view)

## 📂 Repository Contents

```
├── Sample_-_Superstore.csv     # Raw source data
├── Sales_Dashboard.xlsx        # Final interactive dashboard
└── README.md
```

## 🚀 How to Use

1. Download `Sales_Dashboard.xlsx`
2. Open in Excel (2016 or later recommended for full Slicer/Timeline support)
3. Go to the `Dashboard` sheet
4. Use the Region, Category, and Segment slicers, or the Order Date timeline, to filter all KPIs and charts interactively

## 📌 Notes

- Underlying pivot table sheets are hidden for a cleaner presentation but remain fully functional — unhide them from Excel's sheet tab right-click menu if you want to inspect the raw pivot structure.
- All KPI and summary values recalculate automatically if new rows are added to the `SalesData` table.
