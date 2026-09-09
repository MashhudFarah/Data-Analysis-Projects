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
`dashboard.xlsx` – a fully formatted workbook with:
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

## 📸 Dashboard Preview

<img width="1269" height="345" alt="Screenshot_2026-09-09_11-11-52" src="https://github.com/user-attachments/assets/157052c6-5a6c-44dc-9d83-1e5bcc269054" />
<img width="1295" height="356" alt="Screenshot_2026-09-09_11-12-21" src="https://github.com/user-attachments/assets/9725c97a-3295-4523-903f-47f187491d28" />
<img width="1302" height="411" alt="Screenshot_2026-09-09_11-12-40" src="https://github.com/user-attachments/assets/bf2b3ef8-e715-4c02-865a-a43230251526" />
<img width="1249" height="383" alt="Screenshot_2026-09-09_11-12-59" src="https://github.com/user-attachments/assets/719a5564-26c2-4675-9352-11bdab628b0f" />
<img width="1237" height="409" alt="Screenshot_2026-09-09_11-13-22" src="https://github.com/user-attachments/assets/1b788724-560b-4bd5-a17d-633ca479ff20" />
<img width="1278" height="343" alt="Screenshot_2026-09-09_11-13-47" src="https://github.com/user-attachments/assets/6c75ae21-3a44-4464-b4ee-b9d48a798d38" />
<img width="751" height="499" alt="Screenshot_2026-09-09_11-20-06" src="https://github.com/user-attachments/assets/a59d3440-2916-4745-9aa6-0a66575c3a5f" />

### Streamlit Web App

<img width="1004" height="410" alt="Screenshot_2026-09-09_12-24-49" src="https://github.com/user-attachments/assets/b8dae8c2-66a3-4c71-9965-f71e61aee2ce" />
<img width="1011" height="461" alt="Screenshot_2026-09-09_12-25-04" src="https://github.com/user-attachments/assets/f00e89be-f026-41cc-ba4d-3557cbaf4a63" />
<img width="995" height="430" alt="Screenshot_2026-09-09_12-25-19" src="https://github.com/user-attachments/assets/83b78e14-ccca-46df-9a4f-b34b5b2ca2cd" />
<img width="288" height="623" alt="Screenshot_2026-09-09_12-25-43" src="https://github.com/user-attachments/assets/59500a79-646b-4c42-a992-10b228dd37ae" />
<img width="967" height="430" alt="Screenshot_2026-09-09_12-25-53" src="https://github.com/user-attachments/assets/cedc0dd1-9b31-4278-8985-5bc383d30627" />
<img width="970" height="466" alt="Screenshot_2026-09-09_12-25-59" src="https://github.com/user-attachments/assets/b40719f9-d168-44a2-ac93-723280e0c61e" />


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
├── 📄 build_dashboard.py          # Step 2: Build Excel data source (dashboard_data_source.xlsx)
├── 📄 app.py                      # Streamlit web app (interactive version)
│
├── 🔧 clean_data.xlsx             # 👈 Generated intermediate
│
└── 🏆 Sales_Dashboard_Polished.xlsx # 👈 Client-ready deliverable (tracked)

