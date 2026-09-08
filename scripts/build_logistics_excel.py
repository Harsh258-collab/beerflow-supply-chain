"""
BeerFlow — Logistics Excel Workbook Builder
Author: Harsh Raj Pandey
Builds a dedicated Logistics Deep-Dive Excel workbook with 8 analytical sheets.
"""

import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import warnings
warnings.filterwarnings('ignore')

BASE = r'C:\Users\BIT\.gemini\antigravity\scratch\beerflow-supply-chain'
DATA = f'{BASE}\\data'
OUT  = f'{BASE}\\excel\\BeerFlow_Logistics_DeepDive.xlsx'

# ─── STYLES ──────────────────────────────────────────────────────
NAVY   = "1F3864"; BLUE   = "2E75B6"; LTBLUE = "D6E4F0"
GREEN  = "375623"; RED    = "C00000"; GOLD   = "FFC000"
LTGRAY = "F2F2F2"; WHITE  = "FFFFFF"; PURPLE = "7030A0"
ORANGE = "E36C09"; TEAL   = "1F7391"

def hfill(c): return PatternFill("solid", fgColor=c)
def hfont(c="FFFFFF", sz=10, bold=True): return Font(name='Calibri', bold=bold, color=c, size=sz)
def bfont(sz=9): return Font(name='Calibri', size=sz)
def ctr(): return Alignment(horizontal='center', vertical='center', wrap_text=True)
def lft(): return Alignment(horizontal='left', vertical='center', wrap_text=True)
def border(): return Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin'))

def hdr_row(ws, row, cols, color=NAVY, start_col=1):
    for c, h in enumerate(cols, start_col):
        cell = ws.cell(row=row, column=c, value=h)
        cell.fill = hfill(color); cell.font = hfont()
        cell.alignment = ctr(); cell.border = border()
    ws.row_dimensions[row].height = 18

def data_rows(ws, df, start_row, start_col=1, alt_col=None, alt_map=None):
    for r, row in enumerate(df.itertuples(index=False), start_row):
        bg = "D6E4F0" if r % 2 == 0 else WHITE
        for c, val in enumerate(row, start_col):
            cell = ws.cell(row=r, column=c, value=val)
            col_name = df.columns[c - start_col] if c - start_col < len(df.columns) else ''
            if alt_col and col_name == alt_col and alt_map:
                fill_color = alt_map.get(str(val), bg)
                cell.fill = hfill(fill_color)
            else:
                cell.fill = hfill(bg)
            cell.font = bfont(); cell.alignment = ctr(); cell.border = border()
        ws.row_dimensions[r].height = 14

def auto_w(ws, mn=8, mx=30):
    for col in ws.columns:
        mx_len = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(mx_len+2, mn), mx)

def title_block(ws, t1, t2, row=1):
    ws.merge_cells(f'A{row}:L{row}')
    c = ws[f'A{row}']
    c.value = t1; c.font = Font(name='Calibri', bold=True, color=NAVY, size=15)
    c.fill = hfill(LTBLUE); c.alignment = ctr(); ws.row_dimensions[row].height = 28
    ws.merge_cells(f'A{row+1}:L{row+1}')
    s = ws[f'A{row+1}']
    s.value = t2; s.font = Font(name='Calibri', color="44546A", size=9, italic=True)
    s.alignment = ctr(); ws.row_dimensions[row+1].height = 16
    return row + 3


# ════════════════════════════════════════════════════════════════
#  SHEET 1: LOGISTICS DASHBOARD
# ════════════════════════════════════════════════════════════════
def build_logistics_dashboard(wb):
    ws = wb.active; ws.title = "Logistics Dashboard"
    ws.sheet_view.showGridLines = False

    # Banner
    ws.merge_cells('A1:N3')
    b = ws['A1']
    b.value = "BEERFLOW — LOGISTICS DEVELOPMENT DEEP DIVE"
    b.font  = Font(name='Calibri', bold=True, color="FFFFFF", size=20)
    b.fill  = hfill(NAVY); b.alignment = ctr()
    ws.row_dimensions[1].height = 48

    ws.merge_cells('A4:N4')
    s = ws['A4']
    s.value = "Author: Harsh Raj Pandey  |  Logistics & Supply Chain Analytics  |  Beverage / FMCG Industry"
    s.font  = Font(name='Calibri', italic=True, color="44546A", size=10)
    s.fill  = hfill(LTBLUE); s.alignment = ctr()

    # Load data for KPIs
    df_route  = pd.read_csv(f'{DATA}/route_delivery_data.csv')
    df_fleet  = pd.read_csv(f'{DATA}/fleet_utilization_data.csv')
    df_wh     = pd.read_csv(f'{DATA}/warehouse_operations_data.csv')
    df_cold   = pd.read_csv(f'{DATA}/cold_chain_data.csv')
    df_fr     = pd.read_csv(f'{DATA}/freight_benchmark_data.csv')

    kpis = [
        ("Routes Analysed",      f"{len(df_route):,}",              "Total delivery routes",   NAVY),
        ("Fleet Vehicles",        f"{len(df_fleet):,}",              "Active vehicles tracked", BLUE),
        ("Avg Vehicle Util%",     f"{df_fleet['Load_Utilization_Pct'].mean():.1f}%", "Capacity utilization", TEAL),
        ("On-Time Delivery",      f"{(df_route['On_Time_Delivery']=='Yes').mean()*100:.1f}%", "Route-level OTD", GREEN if (df_route['On_Time_Delivery']=='Yes').mean()>0.85 else ORANGE),
        ("Cold Chain Breaches",   f"{(df_cold['Temperature_Breach']=='Yes').sum():,}", "Temperature violations", RED),
        ("Avg Cost/Case",         f"INR {df_route['Cost_Per_Case_INR'].mean():.1f}", "Last-mile delivery cost", ORANGE),
        ("Total CO2 Emissions",   f"{df_fleet['CO2_Emission_KG_PM'].sum()/1000:.1f}T", "Fleet monthly emissions", PURPLE),
        ("Avg Pick Accuracy",     f"{df_wh['Pick_Accuracy_Pct'].mean():.2f}%", "Warehouse accuracy", GREEN),
    ]

    tile_cols = [1,3,5,7,9,11,13,15]
    for idx, (lbl, val, unit, color) in enumerate(kpis):
        c = tile_cols[idx % len(tile_cols)]
        for r, (v, fsz) in enumerate([(lbl,9),(val,16),(unit,8)], 6):
            ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c+1)
            cell = ws.cell(row=r, column=c, value=v)
            cell.fill = hfill(color)
            cell.font = Font(name='Calibri', bold=(r==7), color="FFFFFF",
                             size=fsz, italic=(r==8))
            cell.alignment = ctr()
        for r in [6,7,8]: ws.row_dimensions[r].height = 22

    # Region Performance Table
    ws.cell(row=10, column=1, value="Route Performance by Region").font = Font(name='Calibri', bold=True, color=NAVY, size=12)
    reg = df_route.groupby('Region').agg(
        Routes=('Route_ID','count'),
        Avg_Cost_Case=('Cost_Per_Case_INR','mean'),
        OTD_Pct=('On_Time_Delivery', lambda x: round((x=='Yes').mean()*100,1)),
        Avg_Util=('Utilization_Pct','mean'),
        Avg_Delay=('Delivery_Delay_Min','mean')
    ).reset_index().sort_values('OTD_Pct')

    hdrs = ['Region','Routes','Avg Cost/Case (INR)','OTD %','Avg Utilization %','Avg Delay (Min)']
    hdr_row(ws, 11, hdrs, NAVY)
    for r_i, row in enumerate(reg.itertuples(), 12):
        bg = LTBLUE if r_i % 2 == 0 else WHITE
        otd_color = "E2EFDA" if row.OTD_Pct >= 90 else ("FFF2CC" if row.OTD_Pct >= 80 else "FCE4D6")
        vals = [row.Region, row.Routes, f"INR {row.Avg_Cost_Case:.2f}",
                f"{row.OTD_Pct}%", f"{row.Avg_Util:.1f}%", f"{row.Avg_Delay:.1f}"]
        for c_i, v in enumerate(vals, 1):
            cell = ws.cell(row=r_i, column=c_i, value=v)
            cell.fill = hfill(otd_color if c_i==4 else bg)
            cell.font = bfont(); cell.alignment = ctr(); cell.border = border()
        ws.row_dimensions[r_i].height = 16

    # Vehicle type breakdown
    vt = df_route.groupby('Vehicle_Type').agg(
        Routes=('Route_ID','count'),
        Avg_Util=('Utilization_Pct','mean'),
        Avg_Cost_Case=('Cost_Per_Case_INR','mean'),
        OTD=('On_Time_Delivery', lambda x: round((x=='Yes').mean()*100,1))
    ).reset_index()
    ws.cell(row=10, column=8, value="Vehicle Type Performance").font = Font(name='Calibri', bold=True, color=NAVY, size=12)
    hdrs2 = ['Vehicle Type','Routes','Avg Util %','Cost/Case','OTD %']
    hdr_row(ws, 11, hdrs2, BLUE, start_col=8)
    for r_i, row in enumerate(vt.itertuples(), 12):
        bg = LTBLUE if r_i % 2 == 0 else WHITE
        vals = [row.Vehicle_Type, row.Routes, f"{row.Avg_Util:.1f}%",
                f"INR {row.Avg_Cost_Case:.2f}", f"{row.OTD:.1f}%"]
        for c_i, v in enumerate(vals, 8):
            cell = ws.cell(row=r_i, column=c_i, value=v)
            cell.fill = hfill(bg); cell.font = bfont()
            cell.alignment = ctr(); cell.border = border()

    for col in ws.columns:
        ws.column_dimensions[get_column_letter(col[0].column)].width = 18


# ════════════════════════════════════════════════════════════════
#  SHEET 2: ROUTE & LAST-MILE ANALYSIS
# ════════════════════════════════════════════════════════════════
def build_route_sheet(wb):
    ws = wb.create_sheet("Route & Last-Mile Analysis")
    ws.sheet_view.showGridLines = False
    df = pd.read_csv(f'{DATA}/route_delivery_data.csv')
    nr = title_block(ws, "Route & Last-Mile Delivery Analysis",
                     "1,000 delivery routes | Vehicle Utilization | Cost per Case | On-Time Delivery | CO2 Impact")

    # Summary KPIs in a row
    summ = [
        ("Total Routes", f"{len(df):,}", NAVY),
        ("Avg Utilization", f"{df['Utilization_Pct'].mean():.1f}%", BLUE),
        ("Avg Cost/Case", f"INR {df['Cost_Per_Case_INR'].mean():.2f}", ORANGE),
        ("OTD Rate", f"{(df['On_Time_Delivery']=='Yes').mean()*100:.1f}%", GREEN if (df['On_Time_Delivery']=='Yes').mean()>0.85 else RED),
        ("Avg Distance", f"{df['Distance_KM'].mean():.0f} KM", TEAL),
        ("Avg Stops/Route", f"{df['No_of_Stops'].mean():.1f}", PURPLE),
    ]
    for idx, (lbl, val, color) in enumerate(summ):
        c = 1 + idx * 2
        ws.merge_cells(start_row=nr, start_column=c, end_row=nr, end_column=c+1)
        ws.merge_cells(start_row=nr+1, start_column=c, end_row=nr+1, end_column=c+1)
        ws.cell(row=nr,   column=c, value=lbl).fill = hfill(color)
        ws.cell(row=nr,   column=c).font = hfont(sz=9); ws.cell(row=nr, column=c).alignment = ctr()
        ws.cell(row=nr+1, column=c, value=val).fill = hfill(color)
        ws.cell(row=nr+1, column=c).font = hfont(sz=14); ws.cell(row=nr+1, column=c).alignment = ctr()
        for r in [nr, nr+1]: ws.row_dimensions[r].height = 22
    nr += 3

    hdrs = [h.replace('_',' ') for h in df.columns]
    hdr_row(ws, nr, hdrs, NAVY)
    otd_map = {'Yes': 'E2EFDA', 'No': 'FCE4D6'}
    data_rows(ws, df.head(300), nr+1, alt_col='On_Time_Delivery', alt_map=otd_map)
    auto_w(ws); ws.freeze_panes = f'A{nr+1}'


# ════════════════════════════════════════════════════════════════
#  SHEET 3: FLEET UTILIZATION
# ════════════════════════════════════════════════════════════════
def build_fleet_sheet(wb):
    ws = wb.create_sheet("Fleet Utilization")
    ws.sheet_view.showGridLines = False
    df = pd.read_csv(f'{DATA}/fleet_utilization_data.csv')
    nr = title_block(ws, "Fleet & Vehicle Utilization Analysis",
                     "265 vehicles | Capacity Utilization | Maintenance Cost | CO2 Emissions | Breakdown Risk")

    # KPI Summary
    summ = [
        ("Total Fleet", f"{len(df)}", NAVY),
        ("Active Vehicles", f"{(df['Status']=='Active').sum()}", GREEN),
        ("In Repair", f"{(df['Status']=='In Repair').sum()}", RED),
        ("Avg Load Util%", f"{df['Load_Utilization_Pct'].mean():.1f}%", BLUE),
        ("Total CO2/Month", f"{df['CO2_Emission_KG_PM'].sum()/1000:.1f} Tonnes", PURPLE),
        ("Avg Maint Cost", f"INR {df['Maintenance_Cost_INR_PM'].mean():.0f}", ORANGE),
    ]
    for idx, (lbl, val, color) in enumerate(summ):
        c = 1 + idx * 2
        ws.merge_cells(start_row=nr, start_column=c, end_row=nr, end_column=c+1)
        ws.merge_cells(start_row=nr+1, start_column=c, end_row=nr+1, end_column=c+1)
        ws.cell(row=nr, column=c, value=lbl).fill = hfill(color)
        ws.cell(row=nr, column=c).font = hfont(sz=9); ws.cell(row=nr, column=c).alignment = ctr()
        ws.cell(row=nr+1, column=c, value=val).fill = hfill(color)
        ws.cell(row=nr+1, column=c).font = hfont(sz=14); ws.cell(row=nr+1, column=c).alignment = ctr()
        for r in [nr, nr+1]: ws.row_dimensions[r].height = 22
    nr += 3

    # Fleet by vehicle type pivot
    ws.cell(row=nr, column=1, value="Fleet Summary by Vehicle Type").font = Font(name='Calibri', bold=True, color=NAVY, size=11)
    nr += 1
    pivot = df.groupby('Vehicle_Type').agg(
        Count=('Vehicle_ID','count'),
        Avg_Util=('Load_Utilization_Pct','mean'),
        Avg_KM_PM=('KM_Per_Month','mean'),
        Avg_Maint=('Maintenance_Cost_INR_PM','mean'),
        Breakdowns=('Breakdowns_PM','sum'),
        CO2_Total=('CO2_Emission_KG_PM','sum')
    ).reset_index()
    hdrs_p = ['Vehicle Type','Count','Avg Util %','Avg KM/Month','Avg Maint Cost (INR)','Total Breakdowns','Total CO2 (KG)']
    hdr_row(ws, nr, hdrs_p, BLUE)
    for r_i, row in enumerate(pivot.itertuples(), nr+1):
        bg = LTBLUE if r_i % 2 == 0 else WHITE
        vals = [row.Vehicle_Type, row.Count, f"{row.Avg_Util:.1f}%",
                f"{row.Avg_KM_PM:.0f}", f"INR {row.Avg_Maint:.0f}",
                int(row.Breakdowns), f"{row.CO2_Total:.0f}"]
        for c_i, v in enumerate(vals, 1):
            cell = ws.cell(row=r_i, column=c_i, value=v)
            cell.fill = hfill(bg); cell.font = bfont()
            cell.alignment = ctr(); cell.border = border()
        ws.row_dimensions[r_i].height = 16
    nr += len(pivot) + 3

    # Full data
    hdrs = [h.replace('_',' ') for h in df.columns]
    hdr_row(ws, nr, hdrs, NAVY)
    status_map = {'Active': 'E2EFDA', 'In Repair': 'FCE4D6'}
    data_rows(ws, df, nr+1, alt_col='Status', alt_map=status_map)
    auto_w(ws); ws.freeze_panes = f'A{nr+1}'


# ════════════════════════════════════════════════════════════════
#  SHEET 4: WAREHOUSE OPERATIONS
# ════════════════════════════════════════════════════════════════
def build_warehouse_sheet(wb):
    ws = wb.create_sheet("Warehouse Operations")
    ws.sheet_view.showGridLines = False
    df = pd.read_csv(f'{DATA}/warehouse_operations_data.csv')
    nr = title_block(ws, "Warehouse Operations Analysis",
                     "5 Warehouses | 12 Months | Space Utilization | Pick Accuracy | Damage Rate | Throughput")

    # Warehouse comparison pivot
    pivot = df.groupby('Warehouse').agg(
        Avg_Space_Util=('Space_Utilization_Pct','mean'),
        Avg_Pick_Acc=('Pick_Accuracy_Pct','mean'),
        Total_Receipts=('Receipts_Cases','sum'),
        Total_Dispatches=('Dispatches_Cases','sum'),
        Avg_Damage=('Damage_Rate_Pct','mean'),
        Total_Orders=('Orders_Processed','sum'),
        Avg_Cycle_Time=('Order_Cycle_Time_Hrs','mean'),
        Total_Shrinkage=('Shrinkage_Pct','mean'),
    ).reset_index()

    ws.cell(row=nr, column=1, value="Warehouse Performance Comparison (Full Year 2024)").font = Font(name='Calibri', bold=True, color=NAVY, size=12)
    nr += 1
    hdrs = ['Warehouse','Avg Space Util %','Avg Pick Accuracy %','Total Receipts','Total Dispatches',
            'Avg Damage %','Total Orders','Avg Cycle Time (Hrs)','Avg Shrinkage %']
    hdr_row(ws, nr, hdrs, NAVY)
    for r_i, row in enumerate(pivot.itertuples(), nr+1):
        bg = LTBLUE if r_i % 2 == 0 else WHITE
        pick_ok  = row.Avg_Pick_Acc >= 98.5
        dmg_ok   = row.Avg_Damage <= 0.5
        vals = [row.Warehouse, f"{row.Avg_Space_Util:.1f}%",
                f"{row.Avg_Pick_Acc:.2f}%", f"{row.Total_Receipts:,}",
                f"{row.Total_Dispatches:,}", f"{row.Avg_Damage:.3f}%",
                f"{row.Total_Orders:,}", f"{row.Avg_Cycle_Time:.2f}",
                f"{row.Total_Shrinkage:.3f}%"]
        for c_i, v in enumerate(vals, 1):
            cell = ws.cell(row=r_i, column=c_i, value=v)
            if c_i == 3:
                cell.fill = hfill('E2EFDA' if pick_ok else 'FCE4D6')
            elif c_i == 6:
                cell.fill = hfill('E2EFDA' if dmg_ok else 'FCE4D6')
            else:
                cell.fill = hfill(bg)
            cell.font = bfont(); cell.alignment = ctr(); cell.border = border()
        ws.row_dimensions[r_i].height = 16
    nr += len(pivot) + 3

    hdrs2 = [h.replace('_',' ') for h in df.columns]
    hdr_row(ws, nr, hdrs2, BLUE)
    data_rows(ws, df, nr+1)
    auto_w(ws); ws.freeze_panes = f'A{nr+1}'


# ════════════════════════════════════════════════════════════════
#  SHEET 5: COLD CHAIN MONITORING
# ════════════════════════════════════════════════════════════════
def build_cold_chain_sheet(wb):
    ws = wb.create_sheet("Cold Chain Monitoring")
    ws.sheet_view.showGridLines = False
    df = pd.read_csv(f'{DATA}/cold_chain_data.csv')
    nr = title_block(ws, "Cold Chain Monitoring & Compliance",
                     "600 records | Temperature Breach Detection | Product Status | Cost Impact | Stage-wise Analysis")

    breach_rate = (df['Temperature_Breach']=='Yes').mean()*100
    total_breach_cost = df['Cost_Impact_INR'].sum()

    summ = [
        ("Records Monitored", f"{len(df):,}", NAVY),
        ("Breach Rate", f"{breach_rate:.1f}%", RED if breach_rate > 10 else ORANGE),
        ("Compliant Products", f"{(df['Product_Status']=='Compliant').sum():,}", GREEN),
        ("Quarantined", f"{(df['Product_Status']=='Quarantined').sum():,}", ORANGE),
        ("Disposed", f"{(df['Product_Status']=='Disposed').sum():,}", RED),
        ("Total Cost Impact", f"INR {total_breach_cost/1000:.0f}K", PURPLE),
    ]
    for idx, (lbl, val, color) in enumerate(summ):
        c = 1 + idx * 2
        ws.merge_cells(start_row=nr, start_column=c, end_row=nr, end_column=c+1)
        ws.merge_cells(start_row=nr+1, start_column=c, end_row=nr+1, end_column=c+1)
        ws.cell(row=nr, column=c, value=lbl).fill = hfill(color)
        ws.cell(row=nr, column=c).font = hfont(sz=9); ws.cell(row=nr, column=c).alignment = ctr()
        ws.cell(row=nr+1, column=c, value=val).fill = hfill(color)
        ws.cell(row=nr+1, column=c).font = hfont(sz=13); ws.cell(row=nr+1, column=c).alignment = ctr()
        for r in [nr, nr+1]: ws.row_dimensions[r].height = 22
    nr += 3

    # Stage analysis
    ws.cell(row=nr, column=1, value="Breach Analysis by Supply Chain Stage").font = Font(name='Calibri', bold=True, color=NAVY, size=11)
    nr += 1
    stage = df.groupby('Supply_Chain_Stage').agg(
        Records=('Cold_Chain_ID','count'),
        Breach_Count=('Temperature_Breach', lambda x: (x=='Yes').sum()),
        Breach_Rate=('Temperature_Breach', lambda x: round((x=='Yes').mean()*100,1)),
        Avg_Max_Temp=('Max_Temp_C','mean'),
        Cost_Impact=('Cost_Impact_INR','sum')
    ).reset_index().sort_values('Breach_Rate', ascending=False)
    hdrs = ['Stage','Records','Breaches','Breach Rate %','Avg Max Temp °C','Cost Impact (INR)']
    hdr_row(ws, nr, hdrs, RED)
    for r_i, row in enumerate(stage.itertuples(), nr+1):
        bg = 'FCE4D6' if row.Breach_Rate > 15 else ('FFF2CC' if row.Breach_Rate > 8 else 'E2EFDA')
        vals = [row.Supply_Chain_Stage, row.Records, row.Breach_Count,
                f"{row.Breach_Rate}%", f"{row.Avg_Max_Temp:.1f}°C", f"INR {row.Cost_Impact:,}"]
        for c_i, v in enumerate(vals, 1):
            cell = ws.cell(row=r_i, column=c_i, value=v)
            cell.fill = hfill(bg); cell.font = bfont()
            cell.alignment = ctr(); cell.border = border()
        ws.row_dimensions[r_i].height = 16
    nr += len(stage) + 3

    hdrs2 = [h.replace('_',' ') for h in df.columns]
    hdr_row(ws, nr, hdrs2, NAVY)
    breach_map = {'Yes': 'FCE4D6', 'No': 'E2EFDA'}
    data_rows(ws, df.head(300), nr+1, alt_col='Temperature_Breach', alt_map=breach_map)
    auto_w(ws); ws.freeze_panes = f'A{nr+1}'


# ════════════════════════════════════════════════════════════════
#  SHEET 6: FREIGHT BENCHMARKING
# ════════════════════════════════════════════════════════════════
def build_freight_sheet(wb):
    ws = wb.create_sheet("Freight Benchmarking")
    ws.sheet_view.showGridLines = False
    df = pd.read_csv(f'{DATA}/freight_benchmark_data.csv')
    nr = title_block(ws, "Freight Cost Benchmarking — Modal Analysis",
                     "400 shipments | Road vs Rail vs Air vs 3PL | Cost/Case | CO2 | Transit Time | OTD%")

    modal = df.groupby('Freight_Mode').agg(
        Shipments=('Shipment_ID','count'),
        Avg_Cost_Case=('Cost_Per_Case_INR','mean'),
        Avg_Cost_KM=('Cost_Per_KM_INR','mean'),
        Avg_Transit=('Transit_Days','mean'),
        OTD_Pct=('On_Time_Pct','mean'),
        Damage_Pct=('Damage_Rate_Pct','mean'),
        Avg_CO2=('CO2_Emission_KG','mean'),
        Total_Freight=('Freight_Cost_INR','sum')
    ).reset_index().sort_values('Avg_Cost_Case')

    ws.cell(row=nr, column=1, value="Modal Comparison — Cost | Speed | Reliability | Sustainability").font = Font(name='Calibri', bold=True, color=NAVY, size=12)
    nr += 1
    hdrs = ['Freight Mode','Shipments','Avg Cost/Case (INR)','Avg Cost/KM (INR)',
            'Avg Transit Days','OTD %','Damage %','Avg CO2 (KG)','Total Freight (INR)']
    hdr_row(ws, nr, hdrs, NAVY)

    colors_by_rank = [GREEN, TEAL, BLUE, ORANGE, RED]
    for r_i, row in enumerate(modal.itertuples(), nr+1):
        color = colors_by_rank[min(r_i-nr-1, len(colors_by_rank)-1)]
        bg = LTBLUE if r_i % 2 == 0 else LTGRAY
        vals = [row.Freight_Mode, row.Shipments,
                f"INR {row.Avg_Cost_Case:.2f}", f"INR {row.Avg_Cost_KM:.2f}",
                f"{row.Avg_Transit:.1f}d", f"{row.OTD_Pct:.1f}%",
                f"{row.Damage_Pct:.2f}%", f"{row.Avg_CO2:.1f}",
                f"INR {row.Total_Freight:,.0f}"]
        for c_i, v in enumerate(vals, 1):
            cell = ws.cell(row=r_i, column=c_i, value=v)
            cell.fill = hfill(color) if c_i == 1 else hfill(bg)
            cell.font = hfont(sz=9) if c_i == 1 else bfont()
            cell.alignment = ctr(); cell.border = border()
        ws.row_dimensions[r_i].height = 16
    nr += len(modal) + 3

    ws.cell(row=nr, column=1,
            value="KEY INSIGHT: Rail freight is the cheapest mode (INR/case) with acceptable transit time. "
                  "Increasing rail share from 15% to 25% could save INR 15-20L annually.").font = Font(
        name='Calibri', bold=True, color=RED, size=10, italic=True)
    ws.merge_cells(start_row=nr, start_column=1, end_row=nr, end_column=9)
    ws.cell(row=nr, column=1).fill = hfill('FFF2CC')
    ws.row_dimensions[nr].height = 20
    nr += 3

    hdrs2 = [h.replace('_',' ') for h in df.columns]
    hdr_row(ws, nr, hdrs2, BLUE)
    data_rows(ws, df, nr+1)
    auto_w(ws); ws.freeze_panes = f'A{nr+1}'


# ════════════════════════════════════════════════════════════════
#  SHEET 7: SUPPLY CHAIN RISK REGISTER
# ════════════════════════════════════════════════════════════════
def build_risk_sheet(wb):
    ws = wb.create_sheet("SC Risk Register")
    ws.sheet_view.showGridLines = False
    df = pd.read_csv(f'{DATA}/risk_register.csv')
    nr = title_block(ws, "Supply Chain Risk Register",
                     "24 Identified Risks | Likelihood x Impact Matrix | Mitigation Strategies | Risk Owner | Status")

    risk_counts = df['Risk_Level'].value_counts()
    summ = [
        ("Critical Risks", f"{risk_counts.get('Critical',0)}", RED),
        ("High Risks",     f"{risk_counts.get('High',0)}",     ORANGE),
        ("Medium Risks",   f"{risk_counts.get('Medium',0)}",   GOLD),
        ("Low Risks",      f"{risk_counts.get('Low',0)}",      GREEN),
        ("Open Risks",     f"{(df['Status']=='Open').sum()}",   RED),
        ("Mitigated",      f"{(df['Status']=='Mitigated').sum()}", GREEN),
    ]
    for idx, (lbl, val, color) in enumerate(summ):
        c = 1 + idx * 2
        ws.merge_cells(start_row=nr, start_column=c, end_row=nr, end_column=c+1)
        ws.merge_cells(start_row=nr+1, start_column=c, end_row=nr+1, end_column=c+1)
        ws.cell(row=nr, column=c, value=lbl).fill = hfill(color)
        ws.cell(row=nr, column=c).font = hfont(sz=9); ws.cell(row=nr, column=c).alignment = ctr()
        ws.cell(row=nr+1, column=c, value=val).fill = hfill(color)
        ws.cell(row=nr+1, column=c).font = hfont(sz=18); ws.cell(row=nr+1, column=c).alignment = ctr()
        for r in [nr, nr+1]: ws.row_dimensions[r].height = 24
    nr += 3

    hdrs = [h.replace('_',' ') for h in df.columns]
    hdr_row(ws, nr, hdrs, NAVY)
    risk_map = {'Critical': 'FF0000', 'High': 'FFC000', 'Medium': 'FFF2CC', 'Low': 'E2EFDA'}
    status_map = {'Open': 'FCE4D6', 'Mitigated': 'E2EFDA', 'Monitoring': 'FFF2CC'}
    for r_i, row in enumerate(df.sort_values('Risk_Score', ascending=False).itertuples(index=False), nr+1):
        bg = LTBLUE if r_i % 2 == 0 else WHITE
        risk_fill   = risk_map.get(str(row.Risk_Level), bg)
        status_fill = status_map.get(str(row.Status), bg)
        vals = list(row)
        for c_i, v in enumerate(vals, 1):
            cell = ws.cell(row=r_i, column=c_i, value=v)
            if c_i == 7:   cell.fill = hfill(risk_fill)
            elif c_i == 10: cell.fill = hfill(status_fill)
            else:          cell.fill = hfill(bg)
            cell.font = bfont(); cell.alignment = lft() if c_i in [3,8] else ctr()
            cell.border = border()
        ws.row_dimensions[r_i].height = 28

    ws.column_dimensions['C'].width = 45
    ws.column_dimensions['H'].width = 50
    auto_w(ws)
    ws.freeze_panes = f'A{nr+1}'


# ════════════════════════════════════════════════════════════════
#  SHEET 8: LOGISTICS KPI TRACKER
# ════════════════════════════════════════════════════════════════
def build_kpi_tracker(wb):
    ws = wb.create_sheet("Logistics KPI Tracker")
    ws.sheet_view.showGridLines = False

    df_route = pd.read_csv(f'{DATA}/route_delivery_data.csv')
    df_wh    = pd.read_csv(f'{DATA}/warehouse_operations_data.csv')
    df_cold  = pd.read_csv(f'{DATA}/cold_chain_data.csv')
    df_fr    = pd.read_csv(f'{DATA}/freight_benchmark_data.csv')
    df_fleet = pd.read_csv(f'{DATA}/fleet_utilization_data.csv')

    nr = title_block(ws, "Logistics KPI Master Tracker",
                     "Industry-standard KPIs | Actual vs Target | RAG Status | Action Required")

    kpis = [
        # (Category, KPI Name, Actual, Target, Unit, RAG, Action)
        ("Last-Mile", "On-Time Delivery Rate",
         f"{(df_route['On_Time_Delivery']=='Yes').mean()*100:.1f}", "90", "%",
         "Green" if (df_route['On_Time_Delivery']=='Yes').mean()>0.9 else "Amber",
         "Optimize route planning for high-delay zones"),
        ("Last-Mile", "Avg Delivery Delay",
         f"{df_route['Delivery_Delay_Min'].mean():.0f}", "20", "Minutes",
         "Green" if df_route['Delivery_Delay_Min'].mean()<20 else "Red",
         "Reduce stops per route; pre-load vehicles"),
        ("Last-Mile", "Avg Cost Per Case",
         f"{df_route['Cost_Per_Case_INR'].mean():.2f}", "15.00", "INR/Case",
         "Green" if df_route['Cost_Per_Case_INR'].mean()<=15 else "Red",
         "Consolidate loads; increase vehicle utilization"),
        ("Last-Mile", "Vehicle Load Utilization",
         f"{df_route['Utilization_Pct'].mean():.1f}", "85", "%",
         "Green" if df_route['Utilization_Pct'].mean()>=85 else "Amber",
         "Route planning to maximize load factor"),
        ("Fleet", "Avg Fleet Utilization",
         f"{df_fleet['Load_Utilization_Pct'].mean():.1f}", "80", "%",
         "Green" if df_fleet['Load_Utilization_Pct'].mean()>=80 else "Amber",
         "Retire low-utilization vehicles; right-size fleet"),
        ("Fleet", "Fleet Breakdown Rate",
         f"{df_fleet['Breakdowns_PM'].sum()}", "<20", "Count/Month",
         "Green" if df_fleet['Breakdowns_PM'].sum()<20 else "Red",
         "Preventive maintenance for vehicles > 5 years"),
        ("Fleet", "Total CO2 Emissions",
         f"{df_fleet['CO2_Emission_KG_PM'].sum()/1000:.1f}", "<200", "Tonnes/Month",
         "Green" if df_fleet['CO2_Emission_KG_PM'].sum()/1000<200 else "Amber",
         "Shift to CNG/electric; increase rail share"),
        ("Warehouse", "Pick Accuracy",
         f"{df_wh['Pick_Accuracy_Pct'].mean():.2f}", "99", "%",
         "Green" if df_wh['Pick_Accuracy_Pct'].mean()>=99 else "Amber",
         "Barcode scanning at pick; cyclic counts"),
        ("Warehouse", "Avg Order Cycle Time",
         f"{df_wh['Order_Cycle_Time_Hrs'].mean():.2f}", "2.0", "Hours",
         "Green" if df_wh['Order_Cycle_Time_Hrs'].mean()<=2 else "Red",
         "Layout optimization; batch picking strategy"),
        ("Warehouse", "Avg Damage Rate",
         f"{df_wh['Damage_Rate_Pct'].mean():.3f}", "0.3", "%",
         "Green" if df_wh['Damage_Rate_Pct'].mean()<=0.3 else "Amber",
         "Improve packaging; handle with care protocols"),
        ("Warehouse", "Avg Space Utilization",
         f"{df_wh['Space_Utilization_Pct'].mean():.1f}", "75", "%",
         "Green" if df_wh['Space_Utilization_Pct'].mean()<=85 else "Red",
         "Vertical storage; demand-linked space planning"),
        ("Cold Chain", "Temperature Breach Rate",
         f"{(df_cold['Temperature_Breach']=='Yes').mean()*100:.1f}", "<5", "%",
         "Green" if (df_cold['Temperature_Breach']=='Yes').mean()<0.05 else "Red",
         "IoT temp sensors; driver training; reefer audit"),
        ("Cold Chain", "Product Compliance Rate",
         f"{(df_cold['Product_Status']=='Compliant').mean()*100:.1f}", "98", "%",
         "Green" if (df_cold['Product_Status']=='Compliant').mean()>=0.98 else "Amber",
         "Improve cold chain SLAs with 3PL partners"),
        ("Freight", "Rail Freight Share",
         "15.0", "25", "%",
         "Red",
         "Negotiate rail contracts; pilot bulk rail routes"),
        ("Freight", "Avg Freight Cost/Case (Road)",
         f"{df_fr[df_fr['Freight_Mode'].str.contains('Road')]['Cost_Per_Case_INR'].mean():.2f}",
         "12.00", "INR/Case", "Amber",
         "Optimize truck loading; TMS implementation"),
    ]

    hdrs = ['Category','KPI Name','Actual','Target','Unit','RAG Status','Action Required']
    hdr_row(ws, nr, hdrs, NAVY)
    rag_colors = {'Green': 'E2EFDA', 'Amber': 'FFF2CC', 'Red': 'FCE4D6'}
    for r_i, (cat, kpi, actual, tgt, unit, rag, action) in enumerate(kpis, nr+1):
        bg = LTBLUE if r_i % 2 == 0 else WHITE
        rag_fill = rag_colors.get(rag, bg)
        cat_colors = {'Last-Mile': NAVY, 'Fleet': BLUE, 'Warehouse': TEAL,
                      'Cold Chain': PURPLE, 'Freight': ORANGE}
        cat_fill = cat_colors.get(cat, NAVY)
        for c_i, v in enumerate([cat, kpi, actual, tgt, unit, rag, action], 1):
            cell = ws.cell(row=r_i, column=c_i, value=v)
            if c_i == 1:   cell.fill = hfill(cat_fill); cell.font = hfont(sz=9)
            elif c_i == 6: cell.fill = hfill(rag_fill); cell.font = Font(name='Calibri', bold=True, size=9)
            else:          cell.fill = hfill(bg); cell.font = bfont()
            cell.alignment = ctr() if c_i != 7 else lft()
            cell.border = border()
        ws.row_dimensions[r_i].height = 20

    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['G'].width = 45
    auto_w(ws)
    ws.freeze_panes = f'A{nr+1}'


# ════════════════════════════════════════════════════════════════
#  MAIN
# ════════════════════════════════════════════════════════════════
def main():
    wb = Workbook()
    print("[...] Building Logistics Dashboard...")
    build_logistics_dashboard(wb)
    print("[...] Building Route & Last-Mile Analysis...")
    build_route_sheet(wb)
    print("[...] Building Fleet Utilization...")
    build_fleet_sheet(wb)
    print("[...] Building Warehouse Operations...")
    build_warehouse_sheet(wb)
    print("[...] Building Cold Chain Monitoring...")
    build_cold_chain_sheet(wb)
    print("[...] Building Freight Benchmarking...")
    build_freight_sheet(wb)
    print("[...] Building Risk Register...")
    build_risk_sheet(wb)
    print("[...] Building Logistics KPI Tracker...")
    build_kpi_tracker(wb)

    wb.save(OUT)
    print(f"\n[DONE] Logistics workbook saved:\n  {OUT}")

if __name__ == '__main__':
    main()
