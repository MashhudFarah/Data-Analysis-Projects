import os
import sys
from pathlib import Path
from collections import Counter
import pandas as pd
import numpy as np
from dateutil import parser

BASE_DIR = Path(__file__).resolve().parent
CSV_PATH = BASE_DIR / 'messy_data.csv'
OUTPUT_XLSX = BASE_DIR / 'clean_data.xlsx'

def find_column(df, keywords):
    """Return first column name that contains any of the keywords (case-insensitive)."""
    low = {c: c.lower() for c in df.columns}
    for kw in keywords:
        for c, lc in low.items():
            if kw.lower() in lc:
                return c
    return None


def standardize_payment(val):
    if pd.isna(val):
        return val
    s = str(val).strip().lower()
    s = s.replace('_', ' ').replace('-', ' ').replace('.', '')
    s = s.replace('netbanking', 'net banking')
    # common variants
    if any(x in s for x in ['credit card', 'creditcard', 'cc', 'creditcard']):
        return 'Credit Card'
    if any(x in s for x in ['debit card', 'debitcard']):
        return 'Debit Card'
    if 'net' in s and 'bank' in s:
        return 'Net Banking'
    if any(x in s for x in ['cash on delivery', 'cod', 'cash']):
        return 'Cash on Delivery' if 'delivery' in s or 'cod' in s else 'Cash'
    if 'upi' in s:
        return 'UPI'
    if 'paypal' in s:
        return 'PayPal'
    if 'bank' in s and 'transfer' in s:
        return 'Bank Transfer'
    # fallback title case
    return str(val).strip().title()


def standardize_country(val):
    if pd.isna(val):
        return val
    s = str(val).strip().lower().replace('.', '').replace(' ', '')
    if s in ('us','usa','us a','us a','usaa','u s a','u.s.a','u.s.a.') or 'unitedstates' in s:
        return 'United States'
    if s in ('uk','u.k','u.k.','unitedkingdom','greatbritain','england'):
        return 'United Kingdom'
    if s in ('ca','can','canada'):
        return 'Canada'
    if 'australia' in s or s in ('au','aus'):
        return 'Australia'
    if 'india' in s:
        return 'India'
    # default: title case words
    return ' '.join(w.title() for w in str(val).strip().split())


def standardize_category(val):
    if pd.isna(val):
        return val
    s = str(val).strip().lower()
    s = s.replace('&', 'and')
    if 'home' in s and ('kitchen' in s or 'kit' in s):
        return 'Home & Kitchen'
    if 'book' in s:
        return 'Books'
    if 'fashion' in s or 'clothing' in s or 'apparel' in s:
        return 'Fashion'
    if 'beauty' in s or 'cosmetic' in s:
        return 'Beauty'
    if 'elect' in s or 'phone' in s or 'gadget' in s:
        return 'Electronics'
    if 'toy' in s:
        return 'Toys'
    return str(val).strip().title()


def try_parse_date_series(series):
    # try pandas to_datetime (fast), then fallback to dateutil per value
    converted = pd.to_datetime(series, errors='coerce')
    n_converted = converted.notna().sum()
    # attempt to parse remaining with dayfirst True and individual parse
    mask = converted.isna() & series.notna()
    if mask.any():
        parsed = []
        for v in series[mask]:
            try:
                dt = parser.parse(str(v), dayfirst=False)
            except Exception:
                try:
                    dt = parser.parse(str(v), dayfirst=True)
                except Exception:
                    dt = pd.NaT
            parsed.append(dt)
        # assign back
        converted.loc[mask] = parsed
        n_converted = converted.notna().sum()
    return converted, n_converted


def autofit_and_format(ws, df, date_cols):
    from openpyxl.styles import Font
    # header style
    header_font = Font(bold=True)
    for cell in next(ws.iter_rows(min_row=1, max_row=1)):
        cell.font = header_font
    # set column widths
    for i, col in enumerate(df.columns, start=1):
        max_len = max(
            [len(str(col))] + [len(str(x)) for x in df[col].dropna().astype(str).sample(min(200, max(1, df.shape[0])))])
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = min(50, max_len + 2)
    # date formatting
    for dc in date_cols:
        if dc in df.columns:
            col_idx = df.columns.get_loc(dc) + 1
            for cell in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                for c in cell:
                    # openpyxl stores datetimes as datetime objects; excel format
                    try:
                        c.number_format = 'YYYY-MM-DD'
                    except Exception:
                        pass


def main():
    if not CSV_PATH.exists():
        print('CSV file not found:', CSV_PATH)
        sys.exit(1)

    df = pd.read_csv(CSV_PATH, dtype=str, keep_default_na=False, na_values=['', 'NA', 'N/A', 'null'])
    original_rows = len(df)
    original_cols = len(df.columns)

    issues = {}

    # 1) Standardize date formats: find columns containing 'date'
    date_cols = [c for c in df.columns if 'date' in c.lower()]
    date_conversions = {}
    for c in date_cols:
        converted, n_conv = try_parse_date_series(df[c])
        # converted is datetime64[ns]
        df[c] = converted.dt.date  # store as date objects or NaT
        date_conversions[c] = int(n_conv)
    issues['date_conversions'] = date_conversions

    # 2) Remove duplicate rows
    before_dup = len(df)
    df = df.drop_duplicates()
    after_dup = len(df)
    duplicates_removed = before_dup - after_dup
    issues['duplicates_removed'] = int(duplicates_removed)

    # 3) Handle missing values
    # discount -> fill 0
    discount_col = find_column(df, ['discount'])
    discount_filled = 0
    if discount_col:
        # convert numeric
        df[discount_col] = pd.to_numeric(df[discount_col], errors='coerce')
        mask = df[discount_col].isna()
        discount_filled = int(mask.sum())
        df.loc[mask, discount_col] = 0.0
    issues['discount_col'] = discount_col
    issues['discount_filled'] = discount_filled

    # city -> fill 'Unknown'
    city_col = find_column(df, ['city'])
    city_filled = 0
    if city_col:
        mask = df[city_col].isna() | (df[city_col].astype(str).str.strip() == '')
        city_filled = int(mask.sum())
        df.loc[mask, city_col] = 'Unknown'
    issues['city_col'] = city_col
    issues['city_filled'] = city_filled

    # rating -> fill median
    rating_col = find_column(df, ['rating', 'review'])
    rating_filled = 0
    rating_median = None
    if rating_col:
        df[rating_col] = pd.to_numeric(df[rating_col], errors='coerce')
        rating_median = float(df[rating_col].median(skipna=True)) if df[rating_col].dropna().size>0 else 0.0
        mask = df[rating_col].isna()
        rating_filled = int(mask.sum())
        df.loc[mask, rating_col] = rating_median
    issues['rating_col'] = rating_col
    issues['rating_filled'] = rating_filled
    issues['rating_median'] = rating_median

    # 4) Fix inconsistent text: Payment_Method, Country, Product_Category
    payment_col = find_column(df, ['payment', 'payment_method', 'payment method'])
    payment_before_uniques = None
    if payment_col:
        payment_before_uniques = df[payment_col].dropna().astype(str).str.strip().unique().tolist()
        df[payment_col] = df[payment_col].apply(standardize_payment)
    issues['payment_col'] = payment_col
    issues['payment_before_uniques'] = payment_before_uniques
    issues['payment_after_uniques'] = df[payment_col].dropna().unique().tolist() if payment_col else None

    country_col = find_column(df, ['country'])
    country_before_uniques = None
    if country_col:
        country_before_uniques = df[country_col].dropna().astype(str).str.strip().unique().tolist()
        df[country_col] = df[country_col].apply(standardize_country)
    issues['country_col'] = country_col
    issues['country_before_uniques'] = country_before_uniques
    issues['country_after_uniques'] = df[country_col].dropna().unique().tolist() if country_col else None

    category_col = find_column(df, ['category', 'product_category', 'product category'])
    category_before_uniques = None
    if category_col:
        category_before_uniques = df[category_col].dropna().astype(str).str.strip().unique().tolist()
        df[category_col] = df[category_col].apply(standardize_category)
    issues['category_col'] = category_col
    issues['category_before_uniques'] = category_before_uniques
    issues['category_after_uniques'] = df[category_col].dropna().unique().tolist() if category_col else None

    # 5) Remove invalid quantity entries (negative or zero)
    qty_col = find_column(df, ['quantity', 'qty'])
    qty_removed = 0
    if qty_col:
        df[qty_col] = pd.to_numeric(df[qty_col], errors='coerce')
        mask_invalid = (df[qty_col].isna()) | (df[qty_col] <= 0)
        qty_removed = int(mask_invalid.sum())
        df = df[~mask_invalid].copy()
    issues['qty_col'] = qty_col
    issues['qty_removed'] = qty_removed

    final_rows = len(df)
    final_cols = len(df.columns)

    # Export to Excel with formatting
    try:
        from openpyxl import load_workbook
        with pd.ExcelWriter(OUTPUT_XLSX, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Cleaned Data', index=False)
            # create summary dataframe
            summary = []
            summary.append(['Original Rows', original_rows])
            summary.append(['Final Rows', final_rows])
            summary.append(['Original Columns', original_cols])
            summary.append(['Final Columns', final_cols])
            summary.append(['Duplicates Removed', duplicates_removed])
            summary.append(['Discount Filled (null->0)', discount_filled])
            summary.append(['City Filled (null->Unknown)', city_filled])
            summary.append(['Rating Filled (null->median)', rating_filled])
            summary.append(['Rating Median Used', rating_median])
            summary.append(['Quantity Invalid Rows Removed', qty_removed])
            # date conversions per column
            for k, v in issues['date_conversions'].items():
                summary.append([f'Date conversions in {k}', v])
            # unique counts before/after for standardized fields
            if payment_col:
                summary.append(['Payment method uniques before', len(issues['payment_before_uniques'])])
                summary.append(['Payment method uniques after', len(issues['payment_after_uniques'])])
            if country_col:
                summary.append(['Country uniques before', len(issues['country_before_uniques'])])
                summary.append(['Country uniques after', len(issues['country_after_uniques'])])
            if category_col:
                summary.append(['Product category uniques before', len(issues['category_before_uniques'])])
                summary.append(['Product category uniques after', len(issues['category_after_uniques'])])

            summary_df = pd.DataFrame(summary, columns=['Metric', 'Value'])
            summary_df.to_excel(writer, sheet_name='Summary', index=False)

        # Now post-process workbook for formatting
        from openpyxl import load_workbook
        wb = load_workbook(OUTPUT_XLSX)
        ws = wb['Cleaned Data']
        # apply some simple formatting: bold headers and date format for date columns
        date_columns_in_df = date_cols
        autofit_and_format(ws, df, date_columns_in_df)
        # format summary sheet headers
        from openpyxl.styles import Font
        ws2 = wb['Summary']
        for cell in ws2[1]:
            cell.font = Font(bold=True)
        wb.save(OUTPUT_XLSX)
        print('Wrote cleaned data to', OUTPUT_XLSX)
    except Exception as e:
        print('Failed to write Excel file:', e)
        sys.exit(2)

    # print a short JSON-like summary
    print('\n=== Summary ===')
    print('Original rows:', original_rows)
    print('Final rows:', final_rows)
    print('Duplicates removed:', duplicates_removed)
    print('Discount filled:', discount_filled)
    print('City filled:', city_filled)
    print('Rating filled:', rating_filled, '(median used:', rating_median, ')')
    print('Quantity invalid rows removed:', qty_removed)
    print('Output file:', OUTPUT_XLSX)

if __name__ == '__main__':
    main()
