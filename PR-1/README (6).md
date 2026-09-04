<div align="center">

# -- ! Formula Forge — Advanced Excel Analytics Workbook ! --
### *A Multi-Sheet Excel Project for Logical, Lookup, Date, Text & Dynamic Array Functions*

[![Excel](https://img.shields.io/badge/Excel-Workbook-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-us/microsoft-365/excel)
[![Formulas](https://img.shields.io/badge/Formulas-Logical%20%2F%20Text%20%2F%20Date-4479A1?style=for-the-badge&logo=googlesheets&logoColor=white)](https://support.microsoft.com/excel)
[![Lookup Functions](https://img.shields.io/badge/Lookup-VLOOKUP%20%2F%20XLOOKUP%20%2F%20INDEX--MATCH-FF6F00?style=for-the-badge&logo=databricks&logoColor=white)](https://support.microsoft.com/excel)
[![Dynamic Arrays](https://img.shields.io/badge/Dynamic%20Arrays-FILTER%20%2F%20XMATCH-9C27B0?style=for-the-badge&logo=airtable&logoColor=white)](https://support.microsoft.com/excel)
[![Dashboard](https://img.shields.io/badge/Dashboard-Live%20Reporting-4CAF50?style=for-the-badge&logo=googleanalytics&logoColor=white)](https://support.microsoft.com/excel)

<br/>

> *"A formula connects the cells — a dashboard reveals what they're trying to tell you."*

</div>

---

## 📋 Table of Contents

- [📌 Overview](#-overview)
- [🎯 Problem Statement](#-problem-statement)
- [✨ Key Features](#-key-features)
- [🏗️ Project Structure](#️-project-structure)
- [🔄 Project Workflow](#-project-workflow)
- [🧱 Workbook Design — `PR-1.xlsx`](#-workbook-design--pr-1xlsx)
- [🖥️ Example Formula Output](#️-example-formula-output)
- [🛠️ Tech Stack](#️-tech-stack)
- [📈 Results & Insights](#-results--insights)
- [🏆 Advantages](#-advantages)
- [📄 License](#-license)
- [👤 Author](#-author)
- [🙏 Acknowledgements](#-acknowledgements)

---

## 📌 Overview

**Formula Forge** is a self-contained Excel workbook that models three real-world datasets — **student grades**, **regional sales**, and **employee records** — and ties them together with a live **Dashboard** sheet built entirely on formulas. It's a hands-on tour through the Excel functions an analyst reaches for daily, from simple conditionals to dynamic array lookups.

This project is designed to:
- Demonstrate **logical & conditional formulas** (`IF`, `AND`, `OR`) for grading and eligibility checks
- Practice **text functions** (`LEFT`, `FIND`, `UPPER`, `LOWER`) to extract and reshape names
- Apply **date functions** (`DATEDIF`, `DAYS`, `TODAY`) to compute ages and tenure
- Apply **math functions** (`ROUND`, `CEILING`, `FLOOR`) for salary formatting
- Use **aggregation functions** (`COUNTIFS`, `AVERAGEIFS`, `SUMIFS`) for conditional summaries
- Use **lookup functions** (`VLOOKUP`, `INDEX`/`MATCH`, `XLOOKUP`, `XMATCH`) to pull records across sheets
- Use **dynamic references & arrays** (`INDIRECT`, `OFFSET`, `FILTER`) for flexible, self-updating reports

---

## 🎯 Problem Statement

> **Objective:** Build a multi-sheet Excel workbook (`PR-1.xlsx`) that tracks student grades, sales transactions, and employee data — then wire up a Dashboard sheet that filters, looks up, and summarizes that data live using formulas alone (no manual calculation, no VBA).

| 📂 Feature | 📄 Type | 🔍 Description |
|------------|---------|----------------|
| Grade Calculation | Core | Compute averages and assign letter grades via nested `IF` |
| Eligibility Flags | Core | Flag students/sales above a threshold via `AND` / `OR` |
| Name Extraction | Core | Split full names and reformat casing with text functions |
| Age & Tenure Tracking | Core | Compute age and days-since-joining via date functions |
| Conditional Aggregation | Core | `COUNTIFS`, `AVERAGEIFS`, `SUMIFS` for filtered summaries |
| Cross-Sheet Lookups | Core | `VLOOKUP`, `INDEX`/`MATCH`, `XLOOKUP`, `XMATCH` across sheets |
| Dynamic Reporting | Structure | `INDIRECT`, `OFFSET`, `FILTER` for a self-updating Dashboard |

The goal is to demonstrate **comfort with intermediate-to-advanced Excel** — formulas, cross-sheet references, and dynamic arrays — entirely without macros.

---

## ✨ Key Features

| Feature | Description |
|--------|-------------|
| 🗄️ **Four Linked Sheets** | `Students grade`, `Sales Data`, `Employee Data`, and a `Dashboard` that reads from all three |
| 🎓 **Nested Grading Logic** | `IF`/`OR`/`AND` chains assign letter grades (A–F) from Math and Science scores |
| ✂️ **Text Extraction** | `LEFT` + `FIND` pull first names; `UPPER`/`LOWER` reformat casing |
| 📅 **Date Intelligence** | `DATEDIF` for student age, `DAYS`/`TODAY` for employee tenure |
| 🔢 **Rounding Suite** | `ROUND`, `CEILING`, `FLOOR` for three salary display formats |
| 🎯 **Discount Tiers** | Nested `IF` assigns discount % bands from sale price |
| 📊 **Conditional Aggregates** | `COUNTIFS`, `AVERAGEIFS`, `SUMIFS` power the Dashboard's summary cells |
| 🔍 **Classic & Modern Lookups** | `VLOOKUP` and `INDEX`/`MATCH` side by side with `XLOOKUP`/`XMATCH` |
| 🧮 **Dynamic References** | `INDIRECT` resolves a range from text; `OFFSET` builds a rolling last-N-sales total |
| 🌊 **Spilled Arrays** | `FILTER` returns every top-scoring student as a live, self-resizing list |

---

## 🏗️ Project Structure

```
📦 formula-forge/
│
├── 📊 PR-1.xlsx              ← Full workbook: 4 sheets, formulas + dashboard
│
└── 📄 README.md              ← Project documentation
```

---

## 🔄 Project Workflow

```
Start
  │
  ▼
┌───────────────────────────┐
│  Students grade sheet      │  Scores → Average → Grade → Age
└─────────────┬─────────────┘
              │
              ▼
┌───────────────────────────┐
│  Sales Data sheet          │  Price × Qty → Amount → Discount %
└─────────────┬─────────────┘
              │
              ▼
┌───────────────────────────┐
│  Employee Data sheet       │  Salary formatting → Tenure in days
└─────────────┬─────────────┘
              │
              ▼
┌─────────────────────────────────────────────────────────────┐
│   DASHBOARD → COUNTIFS / AVERAGEIFS / SUMIFS                  │
│   → VLOOKUP / INDEX-MATCH / XLOOKUP / XMATCH                  │
│   → INDIRECT / OFFSET / FILTER                                │
└─────────────────────────────┬───────────────────────────────┘
                               │
                               ▼
                       Live Cross-Sheet Report
              (counts, lookups, rolling totals, top-N list)
                               │
                               ▼
                             Done ✅
```

---

## 🧱 Workbook Design — `PR-1.xlsx`

The workbook contains four sheets, with the Dashboard reading live from the other three:

| Sheet | Key Columns | Purpose |
|-------|-------------|---------|
| 🎓 **Students grade** | `Student ID`, `Full Name`, `DOB`, `Math`, `Science`, `Average`, `Grade` | Scores, computed average, letter grade, and derived name/age fields |
| 💰 **Sales Data** | `Sales ID`, `Region`, `Product`, `Sales Person`, `Price`, `Quantity`, `Sales Amount`, `Discount %` | Per-transaction sales with computed revenue and discount tier |
| 🧑‍💼 **Employee Data** | `Employee ID`, `Name`, `Department`, `Salary`, `Joining Date`, tenure fields | Employee records with rounded salary variants and days since joining |
| 📊 **Dashboard** | — | Cross-sheet formulas: counts, averages, sums, lookups, and a filtered top-students list |

| Concept | Where it's used |
|---------|-----------------|
| 🧮 **Nested `IF`** | Grade bands (`A`–`F`) and discount tiers (`0%`–`15%`) |
| 🔗 **`AND` / `OR`** | `Above 80 in Both` flag; `Discount Eligible` flag |
| ✂️ **Text Functions** | `LEFT(name, FIND(" ", name)-1)` isolates first names; `UPPER`/`LOWER` reformat them |
| 📅 **Date Functions** | `DATEDIF(DOB, TODAY(), "y")` for age; `DAYS(TODAY(), JoinDate)` for tenure |
| 🔢 **Rounding Functions** | `ROUND`, `CEILING`, `FLOOR` — three salary display variants |
| 📊 **Conditional Aggregates** | `COUNTIFS`, `AVERAGEIFS`, `SUMIFS` drive the Dashboard's summary panel |
| 🔍 **Lookup Functions** | `VLOOKUP`, `INDEX`/`MATCH`, `XLOOKUP`, `XMATCH` — classic and modern side by side |
| 🧭 **Dynamic References** | `INDIRECT` resolves a sheet range typed in as text; `OFFSET` builds a rolling last-N total |
| 🌊 **Dynamic Array** | `FILTER` spills every student scoring above 80 in Math into a live list |

---

## 🖥️ Example Formula Output

```excel
' Students grade — Grade assignment via nested IF
G2: =IF(OR(D2<35,E2<35),"F",IF(AVERAGE(D2:E2)>=90,"A",IF(AVERAGE(D2:E2)>=75,"B",
     IF(AVERAGE(D2:E2)>=60,"C",IF(AVERAGE(D2:E2)>=50,"D","F")))))

+------------+-------------------+---------+---------+-------+
| Student ID | Full Name         | Average | Grade   | Age   |
+------------+-------------------+---------+---------+-------+
| 1          | Ishita Rathod     | 88.5    | B       | 20    |
| 2          | Ishaan Solanki    | 91      | A       | 18    |
| 6          | Kavya Chauhan     | 48.5    | F       | 19    |
| 10         | Sneha Deshmukh    | 95.5    | A       | 20    |
+------------+-------------------+---------+---------+-------+

' Sales Data — Discount tier via nested IF
K3: =IF(G3>=50000,"15%",IF(G3>=25000,"10%",IF(G3>=10000,"5%",IF(G3>=5000,"2%","0%"))))

+---------+--------+---------+--------------+-------------+------------+
| Sales ID| Region | Product | Sales Person | Sales Amount| Discount % |
+---------+--------+---------+--------------+-------------+------------+
| 2       | North  | Laptop  | Neha Kulkarni| 2,183,332   | 15%        |
| 5       | East   | Mouse   | Rahul Desai  | 12,900      | 0%         |
| 8       | East   | Monitor | Rahul Desai  | 560,450     | 5%         |
+---------+--------+---------+--------------+-------------+------------+

' Dashboard — FILTER spills every student scoring above 80 in Math
A6: =FILTER('Students grade'!A2:B13, 'Students grade'!D2:D13 > 80)

+------------+-------------------+
| Student ID | Full Name         |
+------------+-------------------+
| 1          | Ishita Rathod     |
| 2          | Ishaan Solanki    |
| 3          | Meera Kapoor      |
| 4          | Ritika Bhatt      |
| 7          | Tanvi Nair        |
| 10         | Sneha Deshmukh    |
+------------+-------------------+

' Dashboard — XLOOKUP a salary by Employee ID
E6: =XLOOKUP(E4, 'Employee Data'!A2:A11, 'Employee Data'!D2:D11, "Not Found")
→ 51116.54  (Employee ID 2 → Pooja Menon, Marketing)

' Dashboard — OFFSET a rolling total of the last N sales
H5: =SUM(OFFSET('Sales Data'!I13, -(H4-1), 0, H4, 1))
→ 828,708  (sum of the last 3 rows in the Sales Amount column)
```

---

## 🛠️ Tech Stack

| Tool | Version | Purpose |
|------|---------|---------|
| 📗 **Microsoft Excel** | 365 / 2021+ | Spreadsheet engine used to build and run the workbook |
| 🧮 **Logical Functions** | Built-in | `IF`, `AND`, `OR` for grading, eligibility, and discount tiers |
| ✂️ **Text Functions** | Built-in | `LEFT`, `FIND`, `UPPER`, `LOWER` for name extraction and casing |
| 📅 **Date Functions** | Built-in | `DATEDIF`, `DAYS`, `TODAY` for age and tenure calculations |
| 🔢 **Math Functions** | Built-in | `ROUND`, `CEILING`, `FLOOR` for salary display variants |
| 📊 **Aggregate Functions** | Built-in | `COUNTIFS`, `AVERAGEIFS`, `SUMIFS` for filtered summaries |
| 🔍 **Lookup Functions** | Built-in | `VLOOKUP`, `INDEX`/`MATCH`, `XLOOKUP`, `XMATCH` |
| 🧭 **Dynamic Functions** | Built-in (365) | `INDIRECT`, `OFFSET`, `FILTER` for live, self-updating references |

---

## 📈 Results & Insights

Reading the Dashboard sheet produces:

- ✅ **12 of 12 students** scored above 50 in Math (`COUNTIFS`)
- 🎯 **80.45** average score among students averaging above 60 (`AVERAGEIFS`)
- 💵 **8,599,081** total sales revenue from North-region Laptop transactions (`SUMIFS`)
- 🌊 **6 students** — Ishita, Ishaan, Meera, Ritika, Tanvi, and Sneha — spill live into the "top scorers" list via `FILTER`
- 🔍 **Cross-sheet lookups resolve correctly** — `VLOOKUP` and `XLOOKUP` return matching results (e.g. Product Code 101 → ₹53,252) for the same query
- 📈 **Rolling 3-sale total of 828,708** computed dynamically via `OFFSET`, updating automatically as new rows are appended
- 🧭 **`INDIRECT` resolves a typed-in range reference** (`'Sales Data'!I2:I13`) to a live sum of 10,461,802

---

## 🏆 Advantages

| Advantage | Detail |
|-----------|--------|
| 🔗 **No Manual Calculation** | Every derived column and dashboard cell is formula-driven, not hardcoded |
| 🧮 **Classic + Modern Lookups Side by Side** | `VLOOKUP`/`INDEX`-`MATCH` sit next to `XLOOKUP`/`XMATCH` for direct comparison |
| 📚 **Educational** | Solid single-file reference for practicing conditional, lookup, and dynamic array formulas together |
| 🧾 **Real-World Modeling** | Mirrors common analyst tasks — grading, discount tiering, tenure tracking, sales rollups |
| 🌊 **Self-Updating Reports** | `FILTER` and `OFFSET`-based cells resize and recalculate automatically as source data changes |
| 🧪 **Extensible** | Easy to extend with additional sheets (e.g. `Inventory`, `Attendance`) or further dashboard cells |

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
**🛠️ Skills:** Excel · Formulas · Lookup Functions · Dynamic Arrays · Data Transformation

</div>

---

## 🙏 Acknowledgements

Special thanks to the following resources and communities that made this project possible:

- 📚 [Microsoft Excel Function Reference](https://support.microsoft.com/en-us/office/excel-functions-alphabetical-b3944572-255d-4efb-bb96-c6d90033e188) — Official function documentation
- 📅 [Excel Date & Time Functions](https://support.microsoft.com/en-us/office/date-and-time-functions-reference-fd1b5961-c1ae-4677-be58-074152f97bfd) — Reference for `DATEDIF`, `DAYS`, and `TODAY`
- 🔍 [Excel Lookup & Reference Functions](https://support.microsoft.com/en-us/office/lookup-and-reference-functions-reference-841e8c17-19b3-4536-9cba-9a222247e1b1) — Reference for `VLOOKUP`, `XLOOKUP`, `INDEX`/`MATCH`, `XMATCH`
- 🌊 [Excel Dynamic Array Functions](https://support.microsoft.com/en-us/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531) — Reference for `FILTER` and spilled arrays
- 💬 [Stack Overflow Community](https://stackoverflow.com/) — Problem-solving support

---

<div align="center">

---

*Made with ❤️ and Dynamic Arrays — Last updated: 04 September, 2026*

</div>
