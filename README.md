# E-commerce Sales & Ad Spend Dashboard

[![Python 3.9+](https://img.shields.io/badge/python-3.9+-blue.svg)](https://www.python.org/downloads/)
[![Streamlit](https://img.shields.io/badge/Streamlit-1.28-red.svg)](https://streamlit.io/)
[![Excel](https://img.shields.io/badge/Excel-PivotTables-green.svg)](https://www.microsoft.com/en-us/microsoft-365/excel)

A complete data analytics project that turns messy e-commerce transaction data into a professional, interactive business dashboard. Built for freelancers and analysts who need to deliver **client‑ready insights fast**.

---

## 🚀 What It Does

This project takes 12,000+ rows of raw e‑commerce transactions and advertising spend data, cleans them, calculates key business metrics, and delivers **two interactive dashboards** – one in Excel (manually polished with slicers and dark theme) and one as a web app (Streamlit).

### 📊 Two Delivery Formats

**1. Excel Dashboard (Showcase)**  
`dashboard.xlsx` – a fully formatted workbook with:
- **5 PivotCharts**: Revenue by Category & Status, Revenue by Country, Revenue by Payment Method, Monthly Revenue Trend, ROAS by Category.
- **4 Slicers**: Filter by Product Category, Country, Payment Method, and Revenue Type (Fulfilled vs. Cancelled).
- **Dark Theme**: Professional black background with white text, optimised for live client demos.
- *Just open the file and start filtering – no setup required.*

**2. Streamlit Web App (Interactive)**  
Run `streamlit run app.py` locally to explore the same insights in a web interface with sidebar filters – ideal for quick internal exploration.

### 📊 Key Metrics Calculated

| Metric | Value |
|--------|-------|
| **Total Revenue** | $1.40M |
| **Fulfilled Revenue** | $1.15M |
| **Cancelled/Returned Revenue** | $253K |
| **Total Ad Spend** | $375K |
| **Overall ROAS** | 2.94x |
| **Profit After Ads** | $760K |

---

## 📸 Dashboard Preview

### Excel Dashboard

| KPI Cards | Slicers |
|---|---|
| ![Excel KPI cards](screenshots/excel/01-kpi-cards.png) | ![Excel slicers](screenshots/excel/02-slicers.png) |

| Monthly Revenue Trend | ROAS by Category |
|---|---|
| ![Excel monthly revenue trend](screenshots/excel/03-monthly-revenue-trend.png) | ![Excel ROAS by category](screenshots/excel/04-roas-by-category.png) |

| Revenue by Payment Method | Revenue by Country |
|---|---|
| ![Excel revenue by payment method](screenshots/excel/05-revenue-by-payment-method.png) | ![Excel revenue by country](screenshots/excel/06-revenue-by-country.png) |

| Revenue by Category & Status |
|---|
| ![Excel revenue by category and status](screenshots/excel/07-revenue-by-category-status.png) |

### Streamlit Web App

| Header & KPIs | Filters Sidebar |
|---|---|
| ![Streamlit header KPIs](screenshots/webapp/01-header-kpis.png) | ![Streamlit filters sidebar](screenshots/webapp/02-filters-sidebar.png) |

| Monthly Revenue Trend | Revenue by Category & Country |
|---|---|
| ![Streamlit monthly revenue trend](screenshots/webapp/03-monthly-revenue-trend.png) | ![Streamlit revenue by category and country](screenshots/webapp/04-revenue-by-category-country.png) |

| Revenue by Payment Method | ROAS by Category |
|---|---|
| ![Streamlit revenue by payment method](screenshots/webapp/05-revenue-by-payment-method.png) | ![Streamlit ROAS by category](screenshots/webapp/06-roas-by-category.png) |

| Data Table Preview |
|---|
| ![Streamlit data table preview](screenshots/webapp/07-data-table-preview.png) |

### Data Cleaning — Before & After

| Before | After |
|---|---|
| ![Messy raw data](screenshots/excel/08-cleaning-before.png) | ![Cleaned data](screenshots/excel/09-cleaning-after.png) |

---

## 🧼 The Cleaning Problem

The raw export arrives inconsistent — the kind of file that breaks naive group-bys. The cleaning layer normalizes:

- **Payment methods** — `Net Banking`, `NetBanking`, `net banking` → one canonical value. Same for `5 Debit Card`, `5 debit card`, `DEBIT CARD` → `Debit Card`.
- **Countries** — `USA`, `U.S.A`, `United States`, `usa` → `USA`. Same for `UK` / `U.K.` / `United Kingdom`.
- **Order status** — mapped into two reporting buckets: `Fulfilled` and `Cancelled/Returned`.
- **Derived columns** — `Revenue` per line, and a `Month` period key for time-series aggregation.

---

## 🛠️ Setup & Reproduce

```bash
git clone https://github.com/MashhudFarah/Data-Analysis-Projects.git
cd Data-Analysis-Projects

# Install dependencies
pip install -r requirements.txt

# Step 1 — clean the raw transactions → produces clean_data.xlsx
python clean_data_script.py

# Step 2 — build the Excel deliverable → produces dashboard.xlsx
python build_dashboard.py

# Step 3 — launch the interactive web app
streamlit run app.py
```

**Requirements:** Python 3.9+, `pandas`, `openpyxl`, `streamlit`, `plotly` (see `requirements.txt`).

Steps 2 and 3 both read from `clean_data.xlsx`, so **step 1 must run first**.

---

## 📁 Project Structure

```bash
Data-Analysis-Projects/
├── README.md                      # Project documentation
├── LICENSE                        # MIT
├── .gitignore                     # Excludes clean_data.xlsx and caches
├── requirements.txt               # Python dependencies
│
├── 📊 ad_spend.csv                # Source: daily ad spend by category
├── 📊 messy_data.csv              # Source: raw e-commerce transactions
│
├── 📄 clean_data_script.py        # Step 1: Clean messy_data.csv → clean_data.xlsx
├── 📄 build_dashboard.py          # Step 2: Build Excel deliverable (dashboard.xlsx)
├── 📄 app.py                      # Step 3: Streamlit web app (interactive version)
│
├── 🔧 clean_data.xlsx             # Generated intermediate (ignored by Git)
├── 🏆 dashboard.xlsx              # Final polished Excel dashboard (tracked)
│
└── screenshots/
    ├── excel/                     # Excel dashboard + before/after cleaning shots
    └── webapp/                    # Streamlit web app screenshots
```

`clean_data.xlsx` is fully regenerable from `clean_data_script.py` and is **not** tracked in Git. `dashboard.xlsx` is tracked because it's a deliverable, not an intermediate.

---

## 🧾 License

Released under the MIT License — see [LICENSE](LICENSE).

---

## 👤 Author

**Mashhud Farah** — [github.com/MashhudFarah](https://github.com/MashhudFarah)
