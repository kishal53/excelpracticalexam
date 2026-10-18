# -- ! Red Wine Quality Dashboard ! --
### *Excel-Based Data Cleaning, PivotTables, Formulas & Interactive Dashboard Project*

> *"Data is only as useful as the tools that analyse it — build those tools, and you command the information."*

---

## 📋 Table of Contents

- [📌 Overview]
- [🎯 Problem Statement]
- [✨ Key Features]
- [🏗️ Project Structure]
- [🔄 Project Workflow]
- [📥 Part A — Dataset & Preparation]
- [📊 Part B — Analysis, Visualisation & Dashboard]
- [🛠️ Tech Stack]
- [📈 Results & Insights]
- [🚀 How to Use]
- [🏆 Advantages]
- [⚠️ Limitations]
- [👤 Author]
- [🙏 Acknowledgements]

---

## 📌 Overview

The **Red Wine Quality Dashboard** is a full Microsoft Excel analysis and reporting workbook built on the **Red Wine Quality** dataset (UCI Machine Learning Repository / Kaggle): **1,599 Portuguese "Vinho Verde" red wines**, each described by 11 physico-chemical measurements and a sensory quality score. The workbook takes the raw CSV and turns it into a cleaned dataset, native PivotTables, advanced formulas, charts, a correlation heat map, an interactive dashboard and full documentation.

This project is designed to:
- Import and clean a real dataset (duplicates, missing values, data types)
- Build native PivotTables and formula-driven summary tables
- Apply VLOOKUP, INDEX-MATCH, IF and nested IF / AND formulas
- Create charts and use Conditional Formatting to show patterns
- Build an interactive dashboard with dropdown filters and dynamic ranges
- Document the workbook so anyone can understand and reuse it

---

## 🎯 Problem Statement

> **Objective:** Choose a dataset from Kaggle / UCI, clean it, analyse it and present the results in an interactive Excel dashboard, completing the tasks in the order listed.

| 📂 Task | 📄 Skills Tested | 🔍 What Was Done |
|---------|------------------|------------------|
| 1. Dataset Selection & Import | Data Import, Initial Cleaning | Imported `winequality-red.csv` (semicolon-delimited) into `1_Raw_Data` |
| 2. Data Cleaning & Preparation | Remove duplicates, missing values, data types | 240 duplicates removed, 0 missing values, formats applied, live cleaning log |
| 3. Create Pivot Tables | Data Analysis, Summarisation | 4 native PivotTables plus formula-driven summary tables |
| 4. Advanced Formulas | VLOOKUP, INDEX-MATCH, IF, nested functions | 7 formula columns on all 1,359 wines plus a Wine Finder |
| 5. Data Visualization | Charts, Conditional Formatting | 4 charts, a correlation heat map, data bars and colour scales |
| 6. Interactive Dashboard | Filters, dynamic ranges | 2 dropdown filters, KPI cards, 2 charts, dynamic named ranges |
| 7. Documentation | Sheet organisation, naming conventions, instructions | A full `7_Documentation` sheet |

---

## ✨ Key Features

| Feature | Description |
|--------|-------------|
| 🧹 **Data Cleaning** | Duplicates removed, missing values checked and every column verified as numeric |
| 🧾 **Cleaning Log** | Live formulas that audit raw rows, clean rows, duplicates, missing cells and column ranges |
| 📊 **Native PivotTables** | 4 PivotTables built from the cleaned data |
| 🧮 **Advanced Formulas** | INDEX-MATCH, VLOOKUP, IF and nested IF with AND across all rows |
| 🔎 **Wine Finder** | Type a Wine_ID and get its quality, alcohol, pH, grade and profile |
| 🌡️ **Correlation Heat Map** | Live CORREL matrix of all 12 variables with a red-white-blue colour scale |
| 🎚️ **Dropdown Filters** | Quality score and Alcohol band filters drive the whole dashboard |
| 🔄 **Dynamic Ranges** | Named ranges grow automatically when rows are added |
| 🚦 **Conditional Formatting** | Data bars, colour scales and rule-based highlights |
| 📖 **Documentation** | Column dictionary, cleaning steps, formula explanations and usage guide |

---

## 🏗️ Project Structure

```
📦 red-wine-quality-dashboard/
│
├── 📄 kishal_ExcelDashboard.xlsx    ← Main workbook (9 sheets)
│     ├── 1_Raw_Data                 ← Untouched import, 1,599 rows
│     ├── 2_Clean_Data               ← 1,359 unique rows with Wine_ID
│     ├── 2_Cleaning_Log             ← Live data-quality checks
│     ├── 3_Pivot_Tables             ← 4 native PivotTables
│     ├── 3_Summary_Tables           ← Formula-driven summaries and band cut-offs
│     ├── 4_Formulas                 ← Lookups, nested IFs, parameters, Wine Finder
│     ├── 5_Visualization            ← Correlation heat map and 4 charts
│     ├── 6_Dashboard                ← Filters, KPI cards, charts, profile table
│     └── 7_Documentation            ← Full workbook documentation
│
├── 📄 winequality-red.csv           ← Source dataset (semicolon-separated)
│
└── 📄 README.md                     ← Project documentation
```

---

## 🔄 Project Workflow

```
Raw Dataset (1,599 wines)
        │
        ▼
┌─────────────────────────────────┐
│  Cleaning (2_Clean_Data)        │  ← Remove duplicates, check missing, set formats
└──────────────┬──────────────────┘
               │
   ┌───────────┼──────────────────────────┐
   ▼           ▼                          ▼
┌──────────┐ ┌─────────────────┐  ┌──────────────────┐
│ Pivot    │ │ Formulas        │  │ Cleaning Log     │
│ & Summary│ │ (lookups, nested│  │ (live audit of   │
│ Tables   │ │  IF, Finder)    │  │  data quality)   │
└────┬─────┘ └────────┬────────┘  └──────────────────┘
     │                │
     └───────┬────────┘
             ▼
   ┌───────────────────┐
   │  Visualization    │  ← Heat map, 4 charts
   └─────────┬─────────┘
             ▼
   ┌───────────────────┐
   │     Dashboard     │  ← Dropdown filters, KPI cards,
   │  + Documentation  │     charts, dynamic ranges
   └───────────────────┘
```

---

## 📥 Part A — Dataset & Preparation

### 📝 1. Dataset

The **1_Raw_Data** sheet holds the original 1,599 samples. The cleaned table keeps these 12 columns plus a `Wine_ID`:

| Column | Unit | Meaning |
|--------|------|---------|
| Fixed_Acidity | g/dm³ | Non-volatile acids (mainly tartaric) |
| Volatile_Acidity | g/dm³ | Acetic acid; too much gives a vinegar taste |
| Citric_Acid | g/dm³ | Adds freshness and flavour |
| Residual_Sugar | g/dm³ | Sugar left after fermentation |
| Chlorides | g/dm³ | Salt content |
| Free_Sulfur_Dioxide | mg/dm³ | SO₂ free to protect against microbes |
| Total_Sulfur_Dioxide | mg/dm³ | Free plus bound SO₂ |
| Density | g/cm³ | Depends mainly on alcohol and sugar |
| pH | pH units | Acidity level |
| Sulphates | g/dm³ | Potassium sulphate preservative |
| Alcohol | % by volume | Alcohol content |
| Quality | score 3–8 | Median sensory rating by wine tasters |

### 🧹 2. Cleaning Steps

| Step | What Was Done |
|------|---------------|
| Import | Loaded the CSV with a semicolon delimiter into `1_Raw_Data` |
| Keep raw intact | Cleaned a copy in `2_Clean_Data` so the raw data stays as evidence |
| Remove duplicates | Removed rows identical in all 12 columns: **240 removed (15.0%)**, **1,359 remain** |
| Missing values | Checked every column with COUNTBLANK: **0 missing**, so nothing needed filling |
| Data types & formats | Every column numeric; fixed decimals per column; Quality as a whole number |
| Naming & ID | Headers renamed to `Title_Case_With_Underscores`; added `Wine_ID` (1–1,359) |
| Verify | `2_Cleaning_Log` shows "Numeric OK" for all 12 columns |

### ⚙️ 3. Settings

- **Dynamic named ranges** (`rngQuality`, `rngAlcohol`, `rngData` …) grow automatically with the data
- **PivotTables** need **Data → Refresh All** after the source data changes

---

## 📊 Part B — Analysis, Visualisation & Dashboard

### 📈 4. Pivot Tables

| PivotTable | Source | Rows / Columns | Values |
|------------|--------|----------------|--------|
| 1. Quality summary | 2_Clean_Data | Quality | Count, avg Alcohol, Volatile_Acidity, pH, Sulphates |
| 2. Alcohol band × Quality | 4_Formulas | Alcohol_Band / Quality | Wine count |
| 3. Wine profile summary | 4_Formulas | Wine_Profile | Count, avg Alcohol, avg Quality |
| 4. Grade × Acidity flag | 4_Formulas | Quality_Grade / Acidity_Flag | Wine count |

`3_Summary_Tables` holds formula-driven versions (COUNTIFS / AVERAGEIFS) that feed the charts and dashboard and store the editable alcohol-band cut-offs.

### 🧮 5. Advanced Formulas

| Column | Function | Purpose |
|--------|----------|---------|
| Quality, Alcohol, Volatile_Acidity, Total_SO2 | `INDEX-MATCH` | Pull values by Wine_ID |
| Quality_Grade | `VLOOKUP` | Translate score into Poor / Average / Above Average / Good / Excellent |
| Alcohol_Band | Nested `IF` | Low / Medium / High from editable cut-offs |
| Sulfur_Risk | Nested `IF` | Low / Moderate / High from total SO₂ |
| Acidity_Flag | `IF` | "Above limit" when volatile acidity exceeds 0.7 |
| Wine_Profile | Nested `IF` + `AND` | Premium / Solid / Standard / Poor |
| Wine Finder | `VLOOKUP`, `INDEX-MATCH`, `IFERROR` | Look up any wine by ID |

Blue cells on `4_Formulas` hold the editable thresholds (volatile acidity limit 0.7, total SO₂ moderate above 50, high above 100).

### 📉 6. Visualisation

| Visual | Shows |
|--------|-------|
| Correlation heat map | Live `CORREL` matrix of all 12 variables |
| Number of wines by quality score | Column chart of the quality distribution |
| Average alcohol % by quality score | Column chart showing alcohol rising with quality |
| Wine count by alcohol band, stacked by quality | Quality mix inside each alcohol band |
| Volatile acidity & sulphates by quality score | Line chart of two opposite trends |

### 🖥️ 7. Dashboard

- **Filters:** Quality score (All, 3–8) and Alcohol band (All, Low, Medium, High) dropdowns
- **KPI cards:** Wines in selection, % of all wines, Avg quality, Avg alcohol, Avg pH, Avg volatile acidity
- **Charts:** Wines by quality (follows the Alcohol band filter) and Wines by alcohol band (follows the Quality filter)
- **Tables:** Quality mix, Alcohol mix and a *Selection profile vs all wines* table with colour scale
- **Tested:** all 28 filter combinations (7 quality × 4 band) were checked against an independent calculation: 784 values, 0 mismatches

### 📝 8. Documentation

The `7_Documentation` sheet covers the purpose, dataset and column dictionary, sheet guide, cleaning steps, PivotTables, every formula, every chart, dashboard filter behaviour, naming conventions and assumptions.

---

## 🛠️ Tech Stack

| Tool | Purpose |
|------|---------|
| 📗 **Microsoft Excel** | Core platform |
| 🧮 **Formulas** | COUNTIFS, AVERAGEIFS, INDEX-MATCH, VLOOKUP, IF, AND, IFERROR, CORREL, COUNTBLANK |
| 📊 **PivotTables** | Native summaries by quality, band, profile and grade |
| 📈 **Charts** | Column, stacked column and line charts |
| 🚦 **Conditional Formatting** | Data bars, colour scales, rule-based highlights |
| ✅ **Data Validation** | Dropdown filters on the dashboard |
| 🏷️ **Named Ranges** | Dynamic ranges that grow with the data |

---

## 📈 Results & Insights

Based on the cleaned data (1,359 unique wines):

- 🍷 **Quality distribution:** most wines score 5 (577) or 6 (535), together about 82%; only 10 score 3 and 17 score 8
- 🥂 **Average quality:** 5.62, with average alcohol 10.43% and average pH 3.31
- 📈 **Alcohol and quality:** average alcohol rises from about 9.96% (quality 3) to 12.16% (quality 8); correlation with quality is **+0.48**, the strongest positive link
- 📉 **Volatile acidity and quality:** falls from 0.885 (quality 3) to 0.404 (quality 7); correlation is **-0.40**, the strongest negative link
- 🧪 **Sulphates:** rise steadily with quality, from 0.57 to 0.77
- 🏷️ **Alcohol bands:** Low 576 wines, Medium 536, High 247
- ⭐ **Wine profiles:** Premium 101, Solid 560, Standard 635, Poor 63
- ⚠️ **Acidity:** 202 wines exceed the 0.7 volatile acidity limit; 29 of the 63 "Poor"-grade wines are among them
- 🔎 **Combined view:** 89 wines are both quality 7 and in the High alcohol band

---

## 🚀 How to Use

1. Open `kishal_ExcelDashboard.xlsx` in **Microsoft Excel** and enable editing.
2. Press **Data → Refresh All** so the PivotTables load current values.
3. Go to the **6_Dashboard** sheet and use the yellow dropdown cells **C3** (Quality score) and **F3** (Alcohol band).
4. Watch the KPI cards, tables and both charts update. Set both back to **All** for the full dataset.
5. On **4_Formulas**, type a Wine_ID (1–1,359) in the yellow cell **M21** to use the Wine Finder.
6. To change the alcohol band cut-offs, edit the blue cells **C16:C18** on **3_Summary_Tables**.
7. Read **7_Documentation** for the full explanation of every sheet.

---

## 🏆 Advantages

| Advantage | Detail |
|-----------|--------|
| 🎓 **Complete Coverage** | Cleaning, pivots, formulas, charts and a dashboard in one workbook |
| 🔄 **Formula-Driven** | Dashboard, summaries and checks recalculate instantly when data changes |
| 🎚️ **Interactive** | Two dropdown filters control KPI cards, tables and charts |
| ✅ **Verified** | Cleaning log and an independent test of all 28 dashboard filter combinations |
| 📚 **Educational** | Each sheet maps to one exam task, in the order listed |
| 📖 **Documented** | Column dictionary, formula explanations and usage steps included |

---


---

## 👤 Author

### Gosai Kishal

**🎓 Role:** Student | Excel & Data Analysis Enthusiast
**📍 Location:** India
**🛠️ Skills:** Excel · PivotTables · Dashboards · Formulas · Data Cleaning · Data Storytelling

> *"Every dataset has a story — build the dashboard that tells it."*

---

## 🙏 Acknowledgements

- Dataset: P. Cortez, A. Cerdeira, F. Almeida, T. Matos and J. Reis, *"Modeling wine preferences by data mining from physicochemical properties"*, Decision Support Systems, 2009 (UCI Machine Learning Repository).

---

*Made with ❤️ and ☕ — Last updated: 18 October, 2026*
