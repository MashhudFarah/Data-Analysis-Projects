from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st


ROOT = Path(__file__).resolve().parent


@st.cache_data
def load_data():
    for path in [ROOT / "dashboard.xlsx", ROOT / "clean_data.xlsx", ROOT / "data.xlsx"]:
        if path.exists():
            df = pd.read_excel(path)
            if "Order_Date" in df.columns:
                df["Order_Date"] = pd.to_datetime(df["Order_Date"], errors="coerce").dt.strftime("%Y-%m-%d")
            if "Revenue" in df.columns:
                df["Revenue"] = pd.to_numeric(df["Revenue"], errors="coerce").fillna(0)
            return df

    st.error("No dashboard data file was found. Please ensure dashboard.xlsx or clean_data.xlsx exists in the same folder as app.py.")
    return pd.DataFrame()


@st.cache_data
def load_ad_spend():
    path = ROOT / "ad_spend.csv"
    if not path.exists():
        return pd.DataFrame(columns=["Date", "Product_Category", "Ad_Spend"])

    ad_df = pd.read_csv(path)
    if "Date" in ad_df.columns:
        ad_df["Date"] = pd.to_datetime(ad_df["Date"], errors="coerce").dt.strftime("%Y-%m-%d")
    if "Ad_Spend" in ad_df.columns:
        ad_df["Ad_Spend"] = pd.to_numeric(ad_df["Ad_Spend"], errors="coerce").fillna(0)
    if "Product_Category" in ad_df.columns:
        ad_df["Product_Category"] = ad_df["Product_Category"].astype(str).str.strip().str.title()
    return ad_df.dropna(subset=["Date", "Product_Category"]).groupby(["Date", "Product_Category"], as_index=False)["Ad_Spend"].sum()


st.set_page_config(page_title="Ecommerce Dashboard", page_icon="📊", layout="wide")

df = load_data()

if df.empty:
    st.stop()

st.sidebar.header("Filter Data")
selected_categories = st.sidebar.multiselect(
    "Product Category",
    options=df["Product_Category"].dropna().unique().tolist() if "Product_Category" in df.columns else [],
    default=df["Product_Category"].dropna().unique().tolist() if "Product_Category" in df.columns else [],
)
selected_countries = st.sidebar.multiselect(
    "Country",
    options=df["Country"].dropna().unique().tolist() if "Country" in df.columns else [],
    default=df["Country"].dropna().unique().tolist() if "Country" in df.columns else [],
)
selected_payment_methods = st.sidebar.multiselect(
    "Payment Method",
    options=df["Payment_Method"].dropna().unique().tolist() if "Payment_Method" in df.columns else [],
    default=df["Payment_Method"].dropna().unique().tolist() if "Payment_Method" in df.columns else [],
)

filtered_df = df.copy()
if "Product_Category" in filtered_df.columns and selected_categories:
    filtered_df = filtered_df[filtered_df["Product_Category"].isin(selected_categories)]
if "Country" in filtered_df.columns and selected_countries:
    filtered_df = filtered_df[filtered_df["Country"].isin(selected_countries)]
if "Payment_Method" in filtered_df.columns and selected_payment_methods:
    filtered_df = filtered_df[filtered_df["Payment_Method"].isin(selected_payment_methods)]

# Basic KPI calculations
revenue_total = float(filtered_df["Revenue"].sum()) if "Revenue" in filtered_df.columns else 0.0
orders_total = int(filtered_df.shape[0])
avg_order_value = float(revenue_total / orders_total) if orders_total else 0.0
category_revenue = (
    filtered_df.groupby("Product_Category", dropna=False)["Revenue"].sum().sort_values(ascending=False)
    if "Product_Category" in filtered_df.columns and "Revenue" in filtered_df.columns
    else pd.Series(dtype=float)
)
country_revenue = (
    filtered_df.groupby("Country", dropna=False)["Revenue"].sum().sort_values(ascending=False)
    if "Country" in filtered_df.columns and "Revenue" in filtered_df.columns
    else pd.Series(dtype=float)
)
monthly_revenue = (
    filtered_df.assign(Month=pd.to_datetime(filtered_df["Order_Date"], errors="coerce").dt.to_period("M").astype(str))
      .groupby("Month", as_index=False)["Revenue"].sum()
    if "Order_Date" in filtered_df.columns and "Revenue" in filtered_df.columns
    else pd.DataFrame(columns=["Month", "Revenue"])
)

ad_spend_df = load_ad_spend()
ad_performance = pd.DataFrame(columns=["Date", "Product_Category", "Fulfilled_Revenue", "Ad_Spend", "ROAS", "Profit_After_Ads"])
if not ad_spend_df.empty and "Order_Date" in filtered_df.columns and "Revenue" in filtered_df.columns and "Product_Category" in filtered_df.columns:
    revenue_type_col = "Revenue_Type" if "Revenue_Type" in filtered_df.columns else None
    fulfilled_df = filtered_df.copy()
    if revenue_type_col:
        fulfilled_df = fulfilled_df[fulfilled_df[revenue_type_col].astype(str).str.lower().eq("fulfilled")]

    ad_revenue = fulfilled_df[["Order_Date", "Product_Category", "Revenue"]].copy()
    ad_revenue = ad_revenue.rename(columns={"Order_Date": "Date"})
    ad_revenue["Date"] = pd.to_datetime(ad_revenue["Date"], errors="coerce").dt.strftime("%Y-%m-%d")
    ad_revenue["Product_Category"] = ad_revenue["Product_Category"].astype(str).str.strip().str.title()
    ad_revenue = (
        ad_revenue.dropna(subset=["Date", "Product_Category"])
        .groupby(["Date", "Product_Category"], as_index=False)["Revenue"]
        .sum()
        .rename(columns={"Revenue": "Fulfilled_Revenue"})
    )
    ad_performance = ad_revenue.merge(ad_spend_df, on=["Date", "Product_Category"], how="left")
    ad_performance["Ad_Spend"] = ad_performance["Ad_Spend"].fillna(0)
    ad_performance["ROAS"] = np.where(ad_performance["Ad_Spend"] > 0, ad_performance["Fulfilled_Revenue"] / ad_performance["Ad_Spend"], 0)
    ad_performance["Profit_After_Ads"] = ad_performance["Fulfilled_Revenue"] - ad_performance["Ad_Spend"]
    ad_performance = ad_performance[ad_performance["Ad_Spend"] > 0].copy()

st.title("Ecommerce Dashboard")
st.caption("A quick view of revenue performance and product trends.")

col1, col2, col3 = st.columns(3)
col1.metric("Total Revenue", f"${revenue_total:,.2f}")
col2.metric("Total Orders", f"{orders_total:,}")
col3.metric("Average Order Value", f"${avg_order_value:,.2f}")

if not ad_performance.empty:
    ad_spend_total = float(ad_performance["Ad_Spend"].sum())
    avg_roas = float(ad_performance["ROAS"].mean())
    profit_after_ads = float(ad_performance["Profit_After_Ads"].sum())
    ad_col1, ad_col2, ad_col3 = st.columns(3)
    ad_col1.metric("Ad Spend", f"${ad_spend_total:,.2f}")
    ad_col2.metric("Avg ROAS", f"{avg_roas:,.2f}x")
    ad_col3.metric("Profit After Ads", f"${profit_after_ads:,.2f}")

    ad_chart = px.bar(
        ad_performance.groupby("Product_Category", as_index=False)["ROAS"].mean().sort_values("ROAS", ascending=False),
        x="Product_Category",
        y="ROAS",
        title="ROAS by Product Category",
        color="Product_Category",
        template="plotly_white",
    )
    ad_chart.update_yaxes(tickformat=".2f")
    st.plotly_chart(ad_chart, use_container_width=True)

if not monthly_revenue.empty:
    monthly_chart = px.line(
        monthly_revenue,
        x="Month",
        y="Revenue",
        title="Monthly Revenue",
        markers=True,
        template="plotly_white",
    )
    monthly_chart.update_layout(xaxis_title="Month", yaxis_title="Revenue")
    monthly_chart.update_yaxes(tickprefix="$", tickformat=",.2f")
    st.plotly_chart(monthly_chart, use_container_width=True)

left_col, right_col = st.columns(2)

if not category_revenue.empty:
    category_chart = px.bar(
        category_revenue.reset_index().rename(columns={"index": "Product_Category", "Revenue": "Revenue"}),
        x="Product_Category",
        y="Revenue",
        title="Revenue by Category",
        color="Product_Category",
        template="plotly_white",
    )
    category_chart.update_yaxes(tickprefix="$", tickformat=",.2f")
    left_col.plotly_chart(category_chart, use_container_width=True)

if not country_revenue.empty:
    country_chart = px.bar(
        country_revenue.reset_index().rename(columns={"index": "Country", "Revenue": "Revenue"}),
        x="Revenue",
        y="Country",
        title="Revenue by Country",
        color="Country",
        orientation="h",
        template="plotly_white",
    )
    country_chart.update_xaxes(tickprefix="$", tickformat=",.2f")
    right_col.plotly_chart(country_chart, use_container_width=True)

if "Payment_Method" in filtered_df.columns and "Revenue" in filtered_df.columns:
    payment_revenue = filtered_df.groupby("Payment_Method", dropna=False)["Revenue"].sum().sort_values(ascending=False)
    payment_chart = px.pie(
        payment_revenue.reset_index().rename(columns={"index": "Payment_Method", "Revenue": "Revenue"}),
        names="Payment_Method",
        values="Revenue",
        title="Revenue by Payment Method",
        hole=0.35,
        template="plotly_white",
    )
    st.plotly_chart(payment_chart, use_container_width=True)

st.dataframe(filtered_df.head(20), use_container_width=True)
