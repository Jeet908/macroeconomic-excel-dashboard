# 🌐 Interactive Country-Year Economic Dashboard (Advanced Excel)

[![Excel](https://img.shields.io/badge/Microsoft%20Excel-Office%20365%20%2F%202021-green?logo=microsoftexcel)](https://products.office.com/excel)
[![ETL](https://img.shields.io/badge/ETL-Power%20Query-blue)](#)

An executive-level macroeconomic reporting dashboard built on World Bank historical datasets (1960–2020), utilizing dynamic array lookups, multi-condition aggregations, Power Query ETL, and pivot analytics.

---

## 📌 Features & Architecture
- **Dynamic KPI Interface:** Single-cell drop-down parameter controls (`Country Code` and `Year`) that instantly update GDP, inflation rates, and exchange-rate trends without VBA macros.
- **Power Query (ETL) Automation:** Cleans, unpivots, and typesets time-series records across 190+ sovereign states.
- **Multi-Decade Pivot Analysis:** Configured custom pivot hierarchies (`HighestGDP`, `RegionalExchangeRate`) grouped into 10-year decade buckets with blank label suppression.
- **Defensive Error Handling:** Built-in `=IFERROR()` and `XLOOKUP` default strings to prevent formula breaks on unpopulated year-country intersections.

---

## 🛠️ Formulas & Engineering Highlights
| Formula | Function in Dashboard |
| :--- | :--- |
| `XLOOKUP` | Dynamic retrieval of sovereign region, currency name, and income tier. |
| `INDEX / MATCH` | Two-way cross-tabulation lookups across regional economic matrices. |
| `SUMIFS` | Multi-conditional numeric extraction across dual filters (Country + Year). |
| `Data Validation` | Dropdown lists restricting user inputs to verified ISO 3-letter alpha codes. |

---

## 📂 Repository Contents
- `/dashboard/Global_Macroeconomic_Dashboard.xlsx` — The complete spreadsheet containing raw data, ETL queries, pivot models, and dashboard UI.
- `/docs/formula_reference.md` — Technical breakdown of lookup formulas and conditional logic.
- `/scripts/build_sample_workbook.py` — Python script using `openpyxl` to reproduce the workbook programmatically.

---

## 👤 Author
**Sourojit Kar**  
- [GitHub Profile](https://github.com/Jeet908)
- [LinkedIn Profile](https://linkedin.com/in/sourojitkar)
