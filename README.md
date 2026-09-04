# E-commerce Sales & Ad Spend Dashboard

[![Python 3.9+](https://img.shields.io/badge/python-3.9+-blue.svg)](https://www.python.org/downloads/)
[![Streamlit](https://img.shields.io/badge/Streamlit-1.28-red.svg)](https://streamlit.io/)
[![Excel](https://img.shields.io/badge/Excel-PivotTables-green.svg)](https://www.microsoft.com/en-us/microsoft-365/excel)

A complete data analytics project that turns messy e-commerce transaction data into a professional, interactive business dashboard. Built for freelancers and analysts who need to deliver **client‑ready insights fast**.

---

## 🚀 What It Does

This project takes 12,000+ rows of raw e‑commerce transactions and advertising spend data, cleans them, calculates key business metrics, and exports **two interactive dashboards** – one in Excel (client‑friendly) and one as a web app (technical).

### 🔧 Two Delivery Formats

**1. Excel Dashboard (Primary Deliverable)**  
`Sales_Dashboard_Polished.xlsx` – a fully formatted workbook with:
- **5 PivotCharts**: Revenue by Category & Status, Revenue by Country, Revenue by Payment Method, Monthly Revenue Trend, ROAS by Category.
- **4 Slicers**: Filter by Product Category, Country, Payment Method, and Revenue Type (Fulfilled vs. Cancelled).
- **Dark Theme**: Professional black background with white text, optimised for live client demos.
- *Just open the file and start filtering – no setup required.*

**2. Streamlit Web App (Interactive Hosted Version)**  
Run `streamlit run app.py` locally to get the same insights in a web interface with sidebar filters – ideal for quick internal exploration.

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

## 📁 Project Structure

```bash
Data-Analysis-Projects/
├── README.md                      # Project documentation (this file)
├── requirements.txt               # Python dependencies
│
├── 📊 ad_spend.csv                # Source: daily ad spend by category
├── 📊 messy_data.csv              # Source: raw e-commerce transactions
│
├── 📄 clean_data_script.py        # Step 1: Clean messy_data.csv → clean_data.xlsx
├── 📄 build_dashboard.py          # Step 2: Build Excel data source (dashboard.xlsx)
├── 📄 app.py                      # Streamlit web app (interactive version)
│
├── 🔧 clean_data.xlsx             # 👈 Generated intermediate (ignored by Git)
├── 🔧 dashboard.xlsx              # 👈 Generated intermediate (ignored by Git)
│
└── 🏆 Sales_Dashboard_Polished.xlsx # 👈 Client-ready deliverable (tracked)

