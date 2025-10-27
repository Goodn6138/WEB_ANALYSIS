import pandas as pd
from openpyxl import Workbook
from app.excel_writer import write_table
import tempfile

month_order = ['January','February','March','April','May','June',
               'July','August','September','October','November','December']

def s_p_month(df):
    monthly = df.groupby('Month')['Amount(Kes.)'].sum().reset_index()
    monthly['Month'] = pd.Categorical(monthly['Month'], categories=month_order, ordered=True)
    monthly = monthly.sort_values('Month')
    total = monthly['Amount(Kes.)'].sum()
    return pd.concat([monthly, pd.DataFrame({'Month':['Total'], 'Amount(Kes.)':[total]})], ignore_index=True)

def s_p_chain(df):
    c = df.groupby('Chain')['Amount(Kes.)'].sum().reset_index().sort_values('Amount(Kes.)', ascending=False)
    total = c['Amount(Kes.)'].sum()
    return pd.concat([c, pd.DataFrame({'Chain':['Total'], 'Amount(Kes.)':[total]})], ignore_index=True)

def s_p_variant(df):
    v = df.groupby('Variant')['Amount(Kes.)'].sum().reset_index().sort_values('Amount(Kes.)', ascending=False)
    total = v['Amount(Kes.)'].sum()
    return pd.concat([v, pd.DataFrame({'Variant':['Total'], 'Amount(Kes.)':[total]})], ignore_index=True)

def s_p_tier3(df):
    t3 = df[df['Chain'].str.contains('Tier 3', case=False, na=False)]
    t3 = t3.groupby('Outlet Name')['Amount(Kes.)'].sum().reset_index().sort_values('Amount(Kes.)', ascending=False)
    total = t3['Amount(Kes.)'].sum()
    return pd.concat([t3, pd.DataFrame({'Outlet Name':['Total'], 'Amount(Kes.)':[total]})], ignore_index=True)

def run_analysis_pipeline(file_path: str) -> str:
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip()
    for col in ['Month','Chain','Variant']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    wb = Workbook()
    ws = wb.active
    ws.title = "Dashboard"
    ws.freeze_panes = "A2"

    row = 2
    row = write_table(ws, s_p_month(df), row, 1, "Sales per Month")
    row = write_table(ws, s_p_chain(df), row, 1, "Sales per Chain")
    row = write_table(ws, s_p_variant(df), row, 1, "Sales per Variant")
    row = write_table(ws, s_p_tier3(df), row, 1, "Tier 3 Outlets")

    output = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
    wb.save(output)
    return output
