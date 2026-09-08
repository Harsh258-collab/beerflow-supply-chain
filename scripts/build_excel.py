"""
BeerFlow — Excel Workbook Builder
Author: Harsh Raj Pandey
Builds a professional multi-sheet Excel workbook with:
  - Raw data sheets
  - Pivot summaries
  - EOQ / Inventory models
  - Distributor scorecard
  - Charts
"""

import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, numbers
)
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
import warnings
warnings.filterwarnings('ignore')

# ─── PATHS ────────────────────────────────────────────────────────────────────
BASE = r'C:\Users\BIT\.gemini\antigravity\scratch\beerflow-supply-chain'
DATA = f'{BASE}\\data'
OUT  = f'{BASE}\\excel\\BeerFlow_Supply_Chain_Analysis.xlsx'

# ─── STYLES ───────────────────────────────────────────────────────────────────
HEADER_FILL   = PatternFill("solid", fgColor="1F3864")   # dark navy
SUBHEAD_FILL  = PatternFill("solid", fgColor="2E75B6")   # blue
ALT_FILL      = PatternFill("solid", fgColor="D6E4F0")   # light blue
GREEN_FILL    = PatternFill("solid", fgColor="E2EFDA")
RED_FILL      = PatternFill("solid", fgColor="FCE4D6")
YELLOW_FILL   = PatternFill("solid", fgColor="FFF2CC")
WHITE_FILL    = PatternFill("solid", fgColor="FFFFFF")

HEADER_FONT   = Font(name='Calibri', bold=True, color="FFFFFF", size=11)
TITLE_FONT    = Font(name='Calibri', bold=True, color="1F3864", size=14)
SUBHEAD_FONT  = Font(name='Calibri', bold=True, color="FFFFFF", size=10)
BODY_FONT     = Font(name='Calibri', size=10)
BOLD_FONT     = Font(name='Calibri', bold=True, size=10)

THIN_BORDER = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)
MEDIUM_BORDER = Border(
    left=Side(style='medium'), right=Side(style='medium'),
    top=Side(style='medium'), bottom=Side(style='medium')
)

CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT   = Alignment(horizontal='left',   vertical='center')
RIGHT  = Alignment(horizontal='right',  vertical='center')


def style_header_row(ws, row, ncols, start_col=1):
    for c in range(start_col, start_col + ncols):
        cell = ws.cell(row=row, column=c)
        cell.fill   = HEADER_FILL
        cell.font   = HEADER_FONT
        cell.alignment = CENTER
        cell.border = THIN_BORDER


def style_data_rows(ws, start_row, end_row, ncols, start_col=1):
    for r in range(start_row, end_row + 1):
        fill = ALT_FILL if r % 2 == 0 else WHITE_FILL
        for c in range(start_col, start_col + ncols):
            cell = ws.cell(row=r, column=c)
            cell.fill      = fill
            cell.font      = BODY_FONT
            cell.border    = THIN_BORDER
            cell.alignment = LEFT


def auto_width(ws, min_w=8, max_w=35):
    for col in ws.columns:
        max_len = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(max_len + 2, min_w), max_w)


def add_title_block(ws, title, subtitle, row=1):
    ws.merge_cells(f'A{row}:J{row}')
    t = ws[f'A{row}']
    t.value     = title
    t.font      = Font(name='Calibri', bold=True, color="1F3864", size=16)
    t.alignment = CENTER
    t.fill      = PatternFill("solid", fgColor="D6E4F0")

    ws.merge_cells(f'A{row+1}:J{row+1}')
    s = ws[f'A{row+1}']
    s.value     = subtitle
    s.font      = Font(name='Calibri', color="44546A", size=10, italic=True)
    s.alignment = CENTER
    ws.row_dimensions[row].height   = 28
    ws.row_dimensions[row+1].height = 18
    return row + 3   # next available row


# ═══════════════════════════════════════════════════════════════════════════════
#  SHEET 1: COVER / DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════════

def build_cover(wb, df_orders, df_inv, df_dist):
    ws = wb.active
    ws.title = "Dashboard"
    ws.sheet_view.showGridLines = False

    # Title banner
    ws.merge_cells('A1:L3')
    banner = ws['A1']
    banner.value = "BEERFLOW — Beverage Supply Chain Optimization"
    banner.font  = Font(name='Calibri', bold=True, color="FFFFFF", size=20)
    banner.fill  = PatternFill("solid", fgColor="1F3864")
    banner.alignment = CENTER
    ws.row_dimensions[1].height = 50

    ws.merge_cells('A4:L4')
    sub = ws['A4']
    sub.value = "Author: Harsh Raj Pandey  |  Domain: Supply Chain & Logistics Analytics  |  Industry: Beverage / FMCG"
    sub.font  = Font(name='Calibri', italic=True, color="44546A", size=10)
    sub.fill  = PatternFill("solid", fgColor="D6E4F0")
    sub.alignment = CENTER
    ws.row_dimensions[4].height = 18

    # KPI tiles row
    total_rev    = df_orders['Revenue_INR'].sum()
    total_orders = len(df_orders)
    avg_margin   = df_orders['Gross_Margin_Pct'].mean()
    otif_rate    = (df_orders['OTIF'] == 'Yes').mean() * 100
    avg_lt       = df_orders['Lead_Time_Days'].mean()
    total_freight= df_orders['Freight_Cost_INR'].sum()

    kpis = [
        ("Total Revenue",       f"INR {total_rev/1e7:.2f} Cr",   "2E75B6"),
        ("Total Orders",        f"{total_orders:,}",              "1F3864"),
        ("Avg Gross Margin",    f"{avg_margin:.1f}%",             "375623"),
        ("OTIF Rate",           f"{otif_rate:.1f}%",              "843C0C" if otif_rate < 80 else "375623"),
        ("Avg Lead Time",       f"{avg_lt:.1f} Days",             "7030A0"),
        ("Total Freight Cost",  f"INR {total_freight/1e5:.1f} L", "C00000"),
    ]

    ws.row_dimensions[6].height = 14
    tile_cols = [1, 3, 5, 7, 9, 11]
    for idx, (label, val, color) in enumerate(kpis):
        c = tile_cols[idx]
        ws.merge_cells(start_row=7, start_column=c, end_row=7, end_column=c+1)
        ws.merge_cells(start_row=8, start_column=c, end_row=8, end_column=c+1)
        ws.merge_cells(start_row=9, start_column=c, end_row=9, end_column=c+1)

        lbl_cell = ws.cell(row=7, column=c, value=label)
        lbl_cell.font  = Font(name='Calibri', bold=True, color="FFFFFF", size=9)
        lbl_cell.fill  = PatternFill("solid", fgColor=color)
        lbl_cell.alignment = CENTER

        val_cell = ws.cell(row=8, column=c, value=val)
        val_cell.font  = Font(name='Calibri', bold=True, color="FFFFFF", size=14)
        val_cell.fill  = PatternFill("solid", fgColor=color)
        val_cell.alignment = CENTER

        # spacer
        ws.cell(row=9, column=c).fill = PatternFill("solid", fgColor=color)

        for r in [7, 8, 9]:
            ws.row_dimensions[r].height = 22

    # Region breakdown table
    ws.row_dimensions[11].height = 14
    region_sum = df_orders.groupby('Region').agg(
        Orders=('Order_ID','count'),
        Revenue=('Revenue_INR','sum'),
        Avg_Margin=('Gross_Margin_Pct','mean'),
        OTIF_Pct=('OTIF', lambda x: (x=='Yes').mean()*100)
    ).reset_index().sort_values('Revenue', ascending=False)

    headers = ['Region', 'Orders', 'Revenue (INR)', 'Avg Margin %', 'OTIF %']
    for col_i, h in enumerate(headers, start=1):
        c = ws.cell(row=12, column=col_i, value=h)
        c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = CENTER; c.border = THIN_BORDER
    ws.row_dimensions[12].height = 18

    for r_i, row in enumerate(region_sum.itertuples(), start=13):
        fill = ALT_FILL if r_i % 2 == 0 else WHITE_FILL
        vals = [row.Region, row.Orders, f"INR {row.Revenue:,.0f}",
                f"{row.Avg_Margin:.1f}%", f"{row.OTIF_Pct:.1f}%"]
        for c_i, v in enumerate(vals, start=1):
            cell = ws.cell(row=r_i, column=c_i, value=v)
            cell.fill = fill; cell.font = BODY_FONT
            cell.alignment = CENTER; cell.border = THIN_BORDER
        ws.row_dimensions[r_i].height = 16

    # SKU performance table
    sku_sum = df_orders.groupby('SKU').agg(
        Revenue=('Revenue_INR','sum'),
        Volume=('Quantity_Cases','sum'),
        Margin=('Gross_Margin_Pct','mean')
    ).reset_index().sort_values('Revenue', ascending=False)

    sku_headers = ['SKU', 'Revenue (INR)', 'Volume (Cases)', 'Avg Margin %']
    col_start = 7
    for col_i, h in enumerate(sku_headers, start=col_start):
        c = ws.cell(row=12, column=col_i, value=h)
        c.fill = SUBHEAD_FILL; c.font = SUBHEAD_FONT; c.alignment = CENTER; c.border = THIN_BORDER

    for r_i, row in enumerate(sku_sum.itertuples(), start=13):
        fill = ALT_FILL if r_i % 2 == 0 else WHITE_FILL
        vals = [row.SKU, f"INR {row.Revenue:,.0f}", f"{row.Volume:,}", f"{row.Margin:.1f}%"]
        for c_i, v in enumerate(vals, start=col_start):
            cell = ws.cell(row=r_i, column=c_i, value=v)
            cell.fill = fill; cell.font = BODY_FONT
            cell.alignment = CENTER; cell.border = THIN_BORDER

    for col in ws.columns:
        ws.column_dimensions[get_column_letter(col[0].column)].width = 18


# ═══════════════════════════════════════════════════════════════════════════════
#  SHEET 2: ORDERS DATA
# ═══════════════════════════════════════════════════════════════════════════════

def build_orders(wb, df):
    ws = wb.create_sheet("Orders Data")
    ws.sheet_view.showGridLines = False

    next_row = add_title_block(ws,
        "Order & Logistics Data — BeerFlow",
        "2,000 synthetic orders | Jan 2023 – Dec 2024 | 5 Regions | 7 SKUs"
    )

    headers = list(df.columns)
    for c_i, h in enumerate(headers, start=1):
        cell = ws.cell(row=next_row, column=c_i, value=h.replace('_',' '))
        cell.fill = HEADER_FILL; cell.font = HEADER_FONT
        cell.alignment = CENTER; cell.border = THIN_BORDER
    ws.row_dimensions[next_row].height = 18

    for r_i, row in enumerate(df.head(500).itertuples(index=False), start=next_row+1):
        fill = ALT_FILL if r_i % 2 == 0 else WHITE_FILL
        for c_i, val in enumerate(row, start=1):
            cell = ws.cell(row=r_i, column=c_i, value=val)
            cell.fill = fill; cell.font = BODY_FONT
            cell.alignment = CENTER; cell.border = THIN_BORDER
        ws.row_dimensions[r_i].height = 14

    auto_width(ws)
    ws.freeze_panes = f'A{next_row+1}'


# ═══════════════════════════════════════════════════════════════════════════════
#  SHEET 3: INVENTORY & EOQ MODEL
# ═══════════════════════════════════════════════════════════════════════════════

def build_inventory(wb, df):
    ws = wb.create_sheet("Inventory & EOQ Model")
    ws.sheet_view.showGridLines = False

    next_row = add_title_block(ws,
        "Inventory Optimization — EOQ Model",
        "Economic Order Quantity | Safety Stock | Reorder Points | Stockout Risk Assessment"
    )

    # EOQ formula box
    ws.merge_cells(f'A{next_row}:F{next_row}')
    cell = ws[f'A{next_row}']
    cell.value = "EOQ Formula:  EOQ = SQRT( 2 x Annual Demand x Ordering Cost / Holding Cost per Unit )"
    cell.font  = Font(name='Calibri', bold=True, color="7030A0", size=11, italic=True)
    cell.fill  = PatternFill("solid", fgColor="EAD1F5")
    cell.alignment = CENTER
    ws.row_dimensions[next_row].height = 20
    next_row += 2

    headers = [h.replace('_', ' ') for h in df.columns]
    for c_i, h in enumerate(headers, start=1):
        cell = ws.cell(row=next_row, column=c_i, value=h)
        cell.fill = HEADER_FILL; cell.font = HEADER_FONT
        cell.alignment = CENTER; cell.border = THIN_BORDER
    ws.row_dimensions[next_row].height = 18

    risk_colors = {'High': 'FCE4D6', 'Medium': 'FFF2CC', 'Low': 'E2EFDA'}
    for r_i, row in enumerate(df.itertuples(index=False), start=next_row+1):
        risk = row.Stockout_Risk
        rfill = PatternFill("solid", fgColor=risk_colors.get(risk, "FFFFFF"))
        base_fill = ALT_FILL if r_i % 2 == 0 else WHITE_FILL
        for c_i, val in enumerate(row, start=1):
            cell = ws.cell(row=r_i, column=c_i, value=val)
            cell.fill = rfill if c_i == len(headers) else base_fill
            cell.font = BODY_FONT; cell.alignment = CENTER; cell.border = THIN_BORDER
        ws.row_dimensions[r_i].height = 14

    auto_width(ws)
    ws.freeze_panes = f'A{next_row+1}'

    # Legend
    leg_row = next_row + len(df) + 3
    ws.cell(row=leg_row, column=1, value="LEGEND:").font = BOLD_FONT
    for i, (risk, color, label) in enumerate([
        ('High', 'FCE4D6', 'High Risk — Stock below Reorder Point'),
        ('Medium', 'FFF2CC', 'Medium Risk — Stock within 1.5x Reorder Point'),
        ('Low', 'E2EFDA', 'Low Risk — Adequate Stock Levels')
    ], start=1):
        c = ws.cell(row=leg_row + i, column=1, value=f"  {risk}")
        c.fill = PatternFill("solid", fgColor=color)
        c.font = BODY_FONT; c.border = THIN_BORDER
        ws.cell(row=leg_row + i, column=2, value=label).font = BODY_FONT


# ═══════════════════════════════════════════════════════════════════════════════
#  SHEET 4: DISTRIBUTOR SCORECARD
# ═══════════════════════════════════════════════════════════════════════════════

def build_distributor(wb, df):
    ws = wb.create_sheet("Distributor Scorecard")
    ws.sheet_view.showGridLines = False

    next_row = add_title_block(ws,
        "Distributor Performance Scorecard",
        "KPI: OTIF % | Fill Rate | Damage Rate | Turnaround Time | Composite Score | Grade"
    )

    # Scoring methodology box
    ws.merge_cells(f'A{next_row}:J{next_row}')
    cell = ws[f'A{next_row}']
    cell.value = ("Scoring: OTIF% x 35%  +  Fill Rate% x 30%  +  "
                  "(100 - Damage*10) x 20%  +  (100 - TAT*10) x 15%   |   "
                  "Grade: A >= 85  |  B >= 70  |  C >= 55  |  D < 55")
    cell.font  = Font(name='Calibri', color="7030A0", size=9, italic=True)
    cell.fill  = PatternFill("solid", fgColor="EAD1F5")
    cell.alignment = CENTER
    ws.row_dimensions[next_row].height = 18
    next_row += 2

    headers = [h.replace('_', ' ') for h in df.columns]
    for c_i, h in enumerate(headers, start=1):
        cell = ws.cell(row=next_row, column=c_i, value=h)
        cell.fill = HEADER_FILL; cell.font = HEADER_FONT
        cell.alignment = CENTER; cell.border = THIN_BORDER
    ws.row_dimensions[next_row].height = 18

    grade_colors = {'A': 'E2EFDA', 'B': 'FFF2CC', 'C': 'FCE4D6', 'D': 'FF0000'}
    for r_i, row in enumerate(df.sort_values('Performance_Score', ascending=False).itertuples(index=False), start=next_row+1):
        base_fill = ALT_FILL if r_i % 2 == 0 else WHITE_FILL
        for c_i, val in enumerate(row, start=1):
            cell = ws.cell(row=r_i, column=c_i, value=val)
            col_name = df.columns[c_i - 1]
            if col_name == 'Grade':
                cell.fill = PatternFill("solid", fgColor=grade_colors.get(str(val), "FFFFFF"))
                cell.font = Font(name='Calibri', bold=True, size=10)
            elif col_name == 'OTIF_Pct':
                cell.fill = GREEN_FILL if float(val) >= 90 else (YELLOW_FILL if float(val) >= 75 else RED_FILL)
                cell.font = BODY_FONT
            else:
                cell.fill = base_fill; cell.font = BODY_FONT
            cell.alignment = CENTER; cell.border = THIN_BORDER
        ws.row_dimensions[r_i].height = 16

    auto_width(ws)
    ws.freeze_panes = f'A{next_row+1}'


# ═══════════════════════════════════════════════════════════════════════════════
#  SHEET 5: DEMAND FORECAST
# ═══════════════════════════════════════════════════════════════════════════════

def build_forecast(wb, df):
    ws = wb.create_sheet("Demand Forecast")
    ws.sheet_view.showGridLines = False

    next_row = add_title_block(ws,
        "Demand Forecasting Model",
        "24-Month Trend Analysis | Seasonal Adjustment | MAPE Accuracy | SKU-Level Forecast"
    )

    headers = [h.replace('_', ' ') for h in df.columns]
    for c_i, h in enumerate(headers, start=1):
        cell = ws.cell(row=next_row, column=c_i, value=h)
        cell.fill = HEADER_FILL; cell.font = HEADER_FONT
        cell.alignment = CENTER; cell.border = THIN_BORDER
    ws.row_dimensions[next_row].height = 18

    for r_i, row in enumerate(df.itertuples(index=False), start=next_row+1):
        mape = float(row.MAPE_Pct)
        mape_fill = GREEN_FILL if mape < 10 else (YELLOW_FILL if mape < 20 else RED_FILL)
        base_fill = ALT_FILL if r_i % 2 == 0 else WHITE_FILL
        for c_i, val in enumerate(row, start=1):
            cell = ws.cell(row=r_i, column=c_i, value=val)
            cell.fill = mape_fill if c_i == 5 else base_fill
            cell.font = BODY_FONT; cell.alignment = CENTER; cell.border = THIN_BORDER
        ws.row_dimensions[r_i].height = 14

    auto_width(ws)
    ws.freeze_panes = f'A{next_row+1}'

    # MAPE legend
    leg_r = next_row + len(df) + 3
    ws.cell(row=leg_r, column=1, value="MAPE Legend:").font = BOLD_FONT
    for i, (color, label) in enumerate([
        ('E2EFDA', 'MAPE < 10% — Excellent Forecast Accuracy'),
        ('FFF2CC', 'MAPE 10–20% — Acceptable Forecast Accuracy'),
        ('FCE4D6', 'MAPE > 20% — Poor Forecast — Needs Recalibration'),
    ], start=1):
        c = ws.cell(row=leg_r+i, column=1, value="   ")
        c.fill = PatternFill("solid", fgColor=color); c.border = THIN_BORDER
        ws.cell(row=leg_r+i, column=2, value=label).font = BODY_FONT


# ═══════════════════════════════════════════════════════════════════════════════
#  SHEET 6: OTIF & KPI SUMMARY
# ═══════════════════════════════════════════════════════════════════════════════

def build_kpi_summary(wb, df_orders):
    ws = wb.create_sheet("OTIF & KPI Summary")
    ws.sheet_view.showGridLines = False

    next_row = add_title_block(ws,
        "OTIF & Key Performance Indicators",
        "On-Time In-Full Rate | Regional Breakdown | Channel Analysis | Shipment Mode Performance"
    )

    # ── OTIF by Region ──
    ws.cell(row=next_row, column=1, value="OTIF Performance by Region").font = Font(name='Calibri', bold=True, color="1F3864", size=12)
    ws.row_dimensions[next_row].height = 20
    next_row += 1

    region_otif = df_orders.groupby('Region').agg(
        Total_Orders=('Order_ID','count'),
        OTIF_Orders=('OTIF', lambda x: (x=='Yes').sum()),
    ).reset_index()
    region_otif['OTIF_Rate_%'] = (region_otif['OTIF_Orders'] / region_otif['Total_Orders'] * 100).round(1)
    region_otif['Status'] = region_otif['OTIF_Rate_%'].apply(
        lambda x: 'On Track' if x >= 90 else ('Needs Attention' if x >= 75 else 'Critical')
    )

    hdrs = ['Region', 'Total Orders', 'OTIF Orders', 'OTIF Rate %', 'Status']
    for c_i, h in enumerate(hdrs, start=1):
        cell = ws.cell(row=next_row, column=c_i, value=h)
        cell.fill = SUBHEAD_FILL; cell.font = SUBHEAD_FONT; cell.alignment = CENTER; cell.border = THIN_BORDER
    next_row += 1

    for r_i, row in enumerate(region_otif.itertuples(index=False), start=next_row):
        base = ALT_FILL if r_i % 2 == 0 else WHITE_FILL
        status_fill = GREEN_FILL if row.Status == 'On Track' else (YELLOW_FILL if row.Status == 'Needs Attention' else RED_FILL)
        for c_i, val in enumerate([row.Region, row.Total_Orders, row.OTIF_Orders, f"{row._3}%", row.Status], start=1):
            cell = ws.cell(row=r_i, column=c_i, value=val)
            cell.fill = status_fill if c_i == 5 else base
            cell.font = BODY_FONT; cell.alignment = CENTER; cell.border = THIN_BORDER
        next_row += 1

    next_row += 2

    # ── OTIF by Channel ──
    ws.cell(row=next_row, column=1, value="OTIF Performance by Channel").font = Font(name='Calibri', bold=True, color="1F3864", size=12)
    next_row += 1

    ch_otif = df_orders.groupby('Channel').agg(
        Orders=('Order_ID','count'),
        OTIF_Rate=('OTIF', lambda x: round((x=='Yes').mean()*100, 1)),
        Avg_Lead_Time=('Lead_Time_Days','mean'),
        Revenue=('Revenue_INR','sum')
    ).reset_index().sort_values('Revenue', ascending=False)

    hdrs2 = ['Channel', 'Orders', 'OTIF Rate %', 'Avg Lead Time (Days)', 'Revenue (INR)']
    for c_i, h in enumerate(hdrs2, start=1):
        cell = ws.cell(row=next_row, column=c_i, value=h)
        cell.fill = SUBHEAD_FILL; cell.font = SUBHEAD_FONT; cell.alignment = CENTER; cell.border = THIN_BORDER
    next_row += 1

    for r_i, row in enumerate(ch_otif.itertuples(index=False), start=next_row):
        base = ALT_FILL if r_i % 2 == 0 else WHITE_FILL
        otif_fill = GREEN_FILL if row.OTIF_Rate >= 90 else (YELLOW_FILL if row.OTIF_Rate >= 75 else RED_FILL)
        vals = [row.Channel, row.Orders, f"{row.OTIF_Rate}%", f"{row.Avg_Lead_Time:.1f}", f"INR {row.Revenue:,.0f}"]
        for c_i, v in enumerate(vals, start=1):
            cell = ws.cell(row=r_i, column=c_i, value=v)
            cell.fill = otif_fill if c_i == 3 else base
            cell.font = BODY_FONT; cell.alignment = CENTER; cell.border = THIN_BORDER
        next_row += 1

    auto_width(ws)


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    print("[...] Loading datasets...")
    df_orders = pd.read_csv(f'{DATA}/orders_data.csv')
    df_inv    = pd.read_csv(f'{DATA}/inventory_data.csv')
    df_dist   = pd.read_csv(f'{DATA}/distributor_scorecard.csv')
    df_fc     = pd.read_csv(f'{DATA}/demand_forecast.csv')

    wb = Workbook()

    print("[...] Building Dashboard sheet...")
    build_cover(wb, df_orders, df_inv, df_dist)

    print("[...] Building Orders Data sheet...")
    build_orders(wb, df_orders)

    print("[...] Building Inventory & EOQ sheet...")
    build_inventory(wb, df_inv)

    print("[...] Building Distributor Scorecard sheet...")
    build_distributor(wb, df_dist)

    print("[...] Building Demand Forecast sheet...")
    build_forecast(wb, df_fc)

    print("[...] Building OTIF & KPI Summary sheet...")
    build_kpi_summary(wb, df_orders)

    print(f"[...] Saving workbook to:\n  {OUT}")
    wb.save(OUT)
    print("[DONE] Excel workbook created successfully!")

if __name__ == '__main__':
    main()
