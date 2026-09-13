<div align="center">

# -- ! Smart Home Tech Sales Analytics — Excel Dashboard Project ! --
### *A Multi-Sheet Excel Workbook for Descriptive Stats, Regression, Pivoting & Live Dashboards*

[![Excel](https://img.shields.io/badge/Excel-Workbook-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-us/microsoft-365/excel)
[![Data Analysis](https://img.shields.io/badge/Data%20Analysis-Toolpak-4479A1?style=for-the-badge&logo=googlesheets&logoColor=white)](https://support.microsoft.com/excel)
[![PivotTables](https://img.shields.io/badge/PivotTables-Region%20%2F%20Product-FF6F00?style=for-the-badge&logo=databricks&logoColor=white)](https://support.microsoft.com/excel)
[![Regression](https://img.shields.io/badge/Regression-Linear%20Model-9C27B0?style=for-the-badge&logo=airtable&logoColor=white)](https://support.microsoft.com/excel)
[![Dashboard](https://img.shields.io/badge/Dashboard-Live%20Reporting-4CAF50?style=for-the-badge&logo=googleanalytics&logoColor=white)](https://support.microsoft.com/excel)

<br/>

> *"200 rows of transactions — one dashboard that tells you where the profit actually comes from."*

</div>

---

## 📋 Table of Contents

- [📌 Overview](#-overview)
- [🎯 Problem Statement](#-problem-statement)
- [✨ Key Features](#-key-features)
- [🏗️ Project Structure](#️-project-structure)
- [🔄 Project Workflow](#-project-workflow)
- [🧱 Workbook Design — `Smart_Home_Tech_Sales_Project.xlsx`](#-workbook-design--smart_home_tech_sales_projectxlsx)
- [🖥️ Sheet-by-Sheet Screenshots](#️-sheet-by-sheet-screenshots)
- [🛠️ Tech Stack](#️-tech-stack)
- [📈 Results & Insights](#-results--insights)
- [🏆 Advantages](#-advantages)
- [📄 License](#-license)
- [👤 Author](#-author)
- [🙏 Acknowledgements](#-acknowledgements)

---

## 📌 Overview

**Smart Home Tech Sales Analytics** is a self-contained Excel workbook that models 200 transactions of a smart-home electronics business — security cameras, smart lighting, smart speakers, and smart thermostats — sold across five regions. Eight linked sheets turn the raw transaction log into descriptive statistics, a regression model, a pivot report, a what-if discount simulator, and a live executive Dashboard.

This project is designed to:
- Practice **descriptive statistics** (`Mean`, `Median`, `Mode`, `Skewness`, `Kurtosis`) via the Data Analysis ToolPak
- Build a **simple linear regression** to test how Quantity predicts Sales
- Summarize revenue with **PivotTables** sliced by Region and Product Category
- Model **what-if scenarios** for discount changes using formula-driven sensitivity analysis
- Rank and flag **top customers** with `RANK`, `SUMIFS`, `COUNTIF`, and `INDEX`/`MATCH`
- Track **month-over-month growth** in sales and profit with `SUMIFS` and conditional formatting arrows
- Tie every sheet together into a single **interactive Dashboard** with slicers, charts, and KPI cards

---

## 🎯 Problem Statement

> **Objective:** Build a multi-sheet Excel workbook (`Smart_Home_Tech_Sales_Project.xlsx`) that ingests a 200-row smart-home sales transaction log, then derives descriptive statistics, a regression model, a pivot summary, a discount what-if model, and a ranked customer summary — culminating in a single-page interactive Dashboard.

| 📂 Feature | 📄 Type | 🔍 Description |
|------------|---------|----------------|
| Transaction Log | Core | 200 rows of Customer, Region, Product, Sales, Quantity, Discount, Order Date, Profit |
| Descriptive Statistics | Core | Mean, median, mode, variance, skewness, kurtosis for Sales, Quantity, Discount |
| Regression Model | Core | Simple linear regression of Sales on Quantity via the Data Analysis ToolPak |
| Regional/Product Pivot | Core | `TOTAL SALES BY REGION AND PRODUCT` PivotTable with slicer |
| Monthly Trend Tracking | Core | `SUMIFS`-driven month-over-month Sales & Profit growth with trend arrows |
| Discount What-If Model | Structure | Sensitivity model projecting profit at a new proposed discount rate |
| Customer Ranking | Structure | `RANK`, `SUMIFS`, `COUNTIF`, `INDEX`/`MATCH` to surface the Top 10 customers |
| Executive Dashboard | Structure | Slicers, KPI cards, line/bar/pie charts, and written insights in one view |

The goal is to demonstrate **end-to-end Excel analytics** — from raw transactional data to statistical modeling to an executive-ready dashboard — entirely without macros.

---

## ✨ Key Features

| Feature | Description |
|--------|-------------|
| 🗄️ **Eight Linked Sheets** | `DATA`, `DescriptiveStats`, `Regression`, `PivotTable`, `MonthlySales`, `WhatIf`, `CustomerSummary`, and `Dashboard` |
| 📊 **Descriptive Stats Panel** | Full ToolPak output — mean, standard error, median, mode, variance, kurtosis, skewness, range — for Sales, Quantity, and Discount |
| 📈 **Regression Output** | R² of 0.83 and a significant coefficient showing Quantity strongly predicts Sales |
| 🧮 **Region × Product PivotTable** | Cross-tab of total sales by Region and Product Category with a Region slicer |
| 📅 **Month-over-Month Growth** | `SUMIFS` totals per month with a conditional-formatting arrow showing % growth vs. prior month |
| 🎚️ **Discount Sensitivity Model** | A single adjustable input cell recalculates projected profit at a new discount rate |
| 🏆 **Top-10 Customer Ranking** | `RANK` orders customers by total purchase; `INDEX`/`MATCH` pulls names into a clean Top-10 list |
| 🖥️ **Interactive Dashboard** | Region slicer, KPI cards (Total Sales, Profit, Avg Discount, Orders), trend line, bar chart, pie chart, and narrative insights |
| 🕒 **Live Timestamp Column** | `NOW()` stamps the DATA sheet so the workbook shows when it was last refreshed |

---

## 🏗️ Project Structure

```
📦 smart-home-tech-sales/
│
├── 📊 Smart_Home_Tech_Sales_Project.xlsx   ← Full workbook: 8 sheets, formulas + dashboard
│
├── 🖼️ screenshots/                          ← Sheet-by-sheet reference images
│   ├── Data.png
│   ├── DescriptiveStats.png
│   ├── Regression_.png
│   ├── pivotTable.png
│   ├── Monthlysales.png
│   ├── whatif.png
│   ├── Customersummry.png
│   └── Dashboard.png
│
└── 📄 README.md                             ← Project documentation
```

---

## 🔄 Project Workflow

```
Start
  │
  ▼
┌───────────────────────────┐
│  DATA sheet                 │  200 transactions → Sales, Quantity, Discount, Profit, Timestamp
└─────────────┬─────────────┘
              │
              ▼
┌───────────────────────────┐
│  DescriptiveStats sheet     │  ToolPak summary → Mean / Median / Skew / Kurtosis
└─────────────┬─────────────┘
              │
              ▼
┌───────────────────────────┐
│  Regression sheet           │  Sales ~ Quantity → R² = 0.834, coefficient = 0.257
└─────────────┬─────────────┘
              │
              ▼
┌───────────────────────────┐
│  PivotTable sheet           │  Region × Product Category → Grand totals
└─────────────┬─────────────┘
              │
              ▼
┌───────────────────────────┐
│  MonthlySales sheet         │  SUMIFS by month → Growth % vs prior month
└─────────────┬─────────────┘
              │
              ▼
┌───────────────────────────┐
│  WhatIf sheet               │  Discount input → Projected profit sensitivity
└─────────────┬─────────────┘
              │
              ▼
┌───────────────────────────┐
│  CustomerSummary sheet      │  RANK + INDEX/MATCH → Top 10 customers
└─────────────┬─────────────┘
              │
              ▼
┌─────────────────────────────────────────────────────────────┐
│   DASHBOARD → Slicers + KPI cards + Trend/Bar/Pie charts       │
│   → Reads live from all seven feeder sheets                   │
└─────────────────────────────┬───────────────────────────────┘
                               │
                               ▼
                     Executive Insights Panel
              (top product, regional leader, margin health)
                               │
                               ▼
                             Done ✅
```

---

## 🧱 Workbook Design — `Smart_Home_Tech_Sales_Project.xlsx`

The workbook contains eight sheets, with the Dashboard reading live from all of them:

| Sheet | Key Columns | Purpose |
|-------|-------------|---------|
| 🧾 **DATA** | `Customer_ID`, `Customer_Name`, `Region`, `Product_Category`, `Sales`, `Quantity`, `Discount`, `Order_Date`, `Month`, `Profit`, `Order Month`, `Timestamp` | The master 200-row transaction log every other sheet pulls from |
| 📊 **DescriptiveStats** | Mean, Standard Error, Median, Mode, Std Dev, Variance, Kurtosis, Skewness, Range, Min, Max, Sum, Count | ToolPak descriptive statistics for Sales, Discount, and Profit columns |
| 📈 **Regression** | Regression Statistics, ANOVA, Coefficients | Simple linear regression of Sales against Quantity |
| 🧮 **PivotTable** | `Product Category` × `Region` | Total Sales by Region and Product Category, with Grand Totals and a Region slicer |
| 📅 **MonthlySales** | `Month`, `Total Sales`, `Total Profit`, `Growth vs Prior Month` | Month-over-month totals with a trend chart and up/down growth arrows |
| 🎚️ **WhatIf** | `Average Current Discount`, `Total Current Sales`, `New Proposed Discount`, `Total Profit at New Discount` | One-input sensitivity model for discount-rate changes |
| 🏆 **CustomerSummary** | `Customer ID`, `Customer Name`, `Total Purchase`, `Total Profit`, `Order Count`, `Rank`, `Top 10?`, `Top Rank`, `High-Value Customer` | Every customer ranked by spend, with a clean Top-10 lookup table |
| 🖥️ **Dashboard** | — | KPI cards, Region slicer, trend line, bar chart, pie chart, and written executive insights |

| Concept | Where it's used |
|---------|-----------------|
| 📊 **Descriptive Statistics (ToolPak)** | Mean, median, mode, variance, skewness, kurtosis for Sales / Discount / Profit |
| 📈 **Linear Regression (ToolPak)** | Sales predicted from Quantity — R², ANOVA table, coefficient significance |
| 🧮 **PivotTable + Slicer** | Region × Product Category cross-tab with an interactive Region filter |
| 📅 **`SUMIFS`** | Month-over-month Sales and Profit totals on the MonthlySales sheet |
| 🎯 **Growth Formula** | `(This Month - Prior Month) / Prior Month` drives the conditional-format trend arrows |
| 🎚️ **Sensitivity Modeling** | A single discount-rate input cell recalculates projected total profit |
| 🏆 **`RANK` / `IF`** | Orders customers by Total Purchase and flags the Top 10 |
| 🔍 **`INDEX`/`MATCH`** | Pulls customer names into the Top-10 list from their rank |
| 📆 **`EOMONTH`** | Normalizes each Order Date down to the first of its month for grouping |
| 🕒 **`NOW()`** | Live timestamp on the DATA sheet showing last recalculation |

---

## 🖥️ Sheet-by-Sheet Screenshots

**DATA** — the master transaction log all other sheets read from:

![DATA sheet](screenshots/Data.png)

**DescriptiveStats** — ToolPak summary statistics for Sales, Discount, and Profit:

![Descriptive Statistics](screenshots/DescriptiveStats.png)

**Regression** — Sales modeled against Quantity, R² = 0.834:

![Regression Output](screenshots/Regression_.png)

**PivotTable** — Total Sales by Region and Product Category:

![PivotTable](screenshots/pivotTable.png)

**MonthlySales** — month-over-month Sales & Profit with growth trend arrows:

![Monthly Sales Trend](screenshots/Monthlysales.png)

**WhatIf** — discount-rate sensitivity model:

![What-If Discount Model](screenshots/whatif.png)

**CustomerSummary** — ranked customers and the Top-10 list:

![Customer Summary](screenshots/Customersummry.png)

**Dashboard** — the full executive view with slicers, KPIs, and charts:

![Executive Dashboard](screenshots/Dashboard.png)

---

## 🛠️ Tech Stack

| Tool | Version | Purpose |
|------|---------|---------|
| 📗 **Microsoft Excel** | 365 / 2021+ | Spreadsheet engine used to build and run the workbook |
| 📊 **Data Analysis ToolPak** | Add-in | Descriptive Statistics and Regression outputs |
| 🧮 **PivotTables & Slicers** | Built-in | Region × Product Category interactive summary |
| 📅 **Date Functions** | Built-in | `EOMONTH` for month bucketing, `NOW()` for live timestamps |
| 📊 **Aggregate Functions** | Built-in | `SUMIFS`, `COUNTIF`, `AVERAGE`, `RANK` for summaries and ranking |
| 🔍 **Lookup Functions** | Built-in | `INDEX`/`MATCH` for pulling Top-10 customer names |
| 📈 **Charts** | Built-in | Line chart (trend), bar chart (region/product), pie chart (category mix) |
| 🎚️ **Conditional Formatting** | Built-in | Up/down arrows on month-over-month growth |

---

## 📈 Results & Insights

Reading the Dashboard and feeder sheets produces:

- 💵 **₹1,75,963.27** total sales and **₹44,388.19** total profit across all 200 orders
- 🎯 **9.73%** average discount rate, holding margin around 25% company-wide
- 🏆 **Security Cameras** is the top-performing product category at **₹63,928.98** in sales, with Smart Lighting trailing at **₹19,879.79**
- 🌍 **West region leads all territories** at over ₹41,000 in sales, while **North is the weakest** at ₹24,062.51
- 📈 **Regression confirms Quantity strongly predicts Sales** — R² of **0.834**, coefficient of **0.257**, and a p-value effectively at zero
- 📅 **Monthly sales fluctuate** between roughly ₹22,000–₹37,600, with growth swinging from **-29% to +38%** month over month
- 🎚️ **What-if model shows** raising the average discount from 9.73% to a flat 10% would drop projected profitability to **₹44,252**
- 👑 **Riley Martinez ranks #1** among all customers with a single order of **₹2,300.33** in sales and **₹647.59** in profit

---

## 🏆 Advantages

| Advantage | Detail |
|-----------|--------|
| 🔗 **Fully Formula-Driven** | Every summary, rank, and KPI recalculates live from the DATA sheet — nothing is hardcoded |
| 📊 **Statistics + Business Reporting in One File** | Combines ToolPak-grade descriptive stats and regression with a business-facing dashboard |
| 🧮 **Interactive Filtering** | Region slicer on both the PivotTable and Dashboard updates every linked chart instantly |
| 🎚️ **Built-In Scenario Testing** | The WhatIf sheet lets anyone test a new discount rate without touching the source data |
| 🏆 **Actionable Customer Ranking** | Top-10 list surfaces high-value customers for retention or loyalty targeting |
| 🧪 **Extensible** | Easy to extend with additional feeder sheets (e.g. `Returns`, `Inventory`) or further dashboard KPIs |

---

## 📄 License

This project is licensed under the **MIT License** — see the [LICENSE](LICENSE) file for full details.

```
MIT License — Free to use, modify, and distribute with attribution.
```

---

## 👤 Author

<div align="center">

### Tejas Varma

[![GitHub](https://img.shields.io/badge/GitHub-Tejas14302-181717?style=for-the-badge&logo=github&logoColor=white)](https://github.com/Tejas14302)

> *"Every dashboard starts with a well-designed formula — just like every program starts with a single line."*

**🎓 Role:** Excel Analyst | Programming Enthusiast \
**📍 Location:** Surat, India \
**🛠️ Skills:** Excel · Data Analysis ToolPak · PivotTables · Regression · Dashboard Design

</div>

---

## 🙏 Acknowledgements

Special thanks to the following resources and communities that made this project possible:

- 📚 [Microsoft Excel Function Reference](https://support.microsoft.com/en-us/office/excel-functions-alphabetical-b3944572-255d-4efb-bb96-c6d90033e188) — Official function documentation
- 📊 [Excel Data Analysis ToolPak](https://support.microsoft.com/en-us/office/use-the-analysis-toolpak-to-perform-complex-data-analysis-6c67ccf0-f4a9-487c-8dec-bdb5a2cefab6) — Reference for Descriptive Statistics and Regression tools
- 🧮 [Excel PivotTables Overview](https://support.microsoft.com/en-us/office/create-a-pivottable-to-analyze-worksheet-data-a9a84538-bfe9-40a9-a8e9-f99134456576) — Reference for PivotTables and Slicers
- 🔍 [Excel Lookup & Reference Functions](https://support.microsoft.com/en-us/office/lookup-and-reference-functions-reference-841e8c17-19b3-4536-9cba-9a222247e1b1) — Reference for `INDEX`/`MATCH`
- 💬 [Stack Overflow Community](https://stackoverflow.com/) — Problem-solving support

---

<div align="center">

---

*Made with ❤️ and PivotTables — Last updated: 13 September, 2026*

</div>
