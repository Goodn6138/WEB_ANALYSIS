import pandas as pd
from openpyxl import Workbook
from app.excel_writer import write_table
import tempfile

month_order = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
]

# =========================================================
# 1️⃣ Sales per Month
# =========================================================
def s_p_month(df):
    monthly = df.groupby('Month')['Amount(Kes.)'].sum().reset_index()
    monthly['Month'] = pd.Categorical(monthly['Month'], categories=month_order, ordered=True)
    monthly = monthly.sort_values('Month')
    total = monthly['Amount(Kes.)'].sum()
    return pd.concat([monthly, pd.DataFrame({'Month':['Total'], 'Amount(Kes.)':[total]})], ignore_index=True)


# =========================================================
# 2️⃣ Sales per Chain (Year Total)
# =========================================================
def s_p_chain(df):
    c = df.groupby('Chain')['Amount(Kes.)'].sum().reset_index().sort_values('Amount(Kes.)', ascending=False)
    total = c['Amount(Kes.)'].sum()
    return pd.concat([c, pd.DataFrame({'Chain':['Total'], 'Amount(Kes.)':[total]})], ignore_index=True)


# =========================================================
# 3️⃣ Sales per Variant
# =========================================================
def s_p_variant(df):
    v = df.groupby('Variant')['Amount(Kes.)'].sum().reset_index().sort_values('Amount(Kes.)', ascending=False)
    total = v['Amount(Kes.)'].sum()
    return pd.concat([v, pd.DataFrame({'Variant':['Total'], 'Amount(Kes.)':[total]})], ignore_index=True)


# =========================================================
# 4️⃣ Sales per Tier-3 Outlet
# =========================================================
def s_p_tier3(df):
    t3 = df[df['Chain'].str.contains('Tier 3', case=False, na=False)]
    t3 = t3.groupby('Outlet Name')['Amount(Kes.)'].sum().reset_index().sort_values('Amount(Kes.)', ascending=False)
    total = t3['Amount(Kes.)'].sum()
    return pd.concat([t3, pd.DataFrame({'Outlet Name':['Total'], 'Amount(Kes.)':[total]})], ignore_index=True)


# =========================================================
# 5️⃣ Sales per Chain per Month (with totals per month)
# =========================================================
def s_p_c_m(df):
    cm = (df.groupby(['Month', 'Chain'], as_index=False)['Amount(Kes.)']
            .sum()
            .assign(Month=lambda x: pd.Categorical(x['Month'], categories=month_order, ordered=True))
            .sort_values(['Month', 'Chain']))

    # add totals per month
    out_rows = []
    for m in month_order:
        sub = cm[cm['Month'] == m]
        if sub.empty:
            continue
        for _, r in sub.iterrows():
            out_rows.append({'Month': m, 'Chain': r['Chain'], 'Amount(Kes.)': r['Amount(Kes.)']})
        out_rows.append({'Month': m, 'Chain': f'Total for {m}', 'Amount(Kes.)': sub['Amount(Kes.)'].sum()})

    return pd.DataFrame(out_rows)


# =========================================================
# 6️⃣ Run the complete analysis pipeline
# =========================================================
def run_analysis_pipeline(file_path: str) -> str:
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip()
    for col in ['Month', 'Chain', 'Variant']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    wb = Workbook()
    ws = wb.active
    ws.title = "Dashboard"
    ws.freeze_panes = "A2"

    row = 2
    row = write_table(ws, s_p_month(df),   row, 1, "Sales per Month")
    row = write_table(ws, s_p_chain(df),   row, 1, "Sales per Chain (Total)")
    row = write_table(ws, s_p_variant(df), row, 1, "Sales per Variant")
    row = write_table(ws, s_p_tier3(df),   row, 1, "Tier 3 Outlets")
    row = write_table(ws, s_p_c_m(df),     row, 1, "Sales per Chain per Month (Detailed)")

    output = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
    wb.save(output)
    return output
