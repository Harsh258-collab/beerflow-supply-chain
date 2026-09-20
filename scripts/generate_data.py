"""
BeerFlow Supply Chain Dataset Generator
Author: Harsh Raj Pandey
Description: Generates realistic synthetic supply chain data for beverage logistics analysis.
             Inspired by real-world FMCG/beverage distribution patterns.
"""

import pandas as pd
import numpy as np
import random
from datetime import datetime, timedelta
import os

np.random.seed(42)
random.seed(42)

# ─── CONFIG ───────────────────────────────────────────────────────────────────
REGIONS        = ['North India', 'South India', 'East India', 'West India', 'Central India']
CITIES         = {
    'North India':   ['Delhi', 'Lucknow', 'Chandigarh', 'Jaipur', 'Amritsar'],
    'South India':   ['Bengaluru', 'Chennai', 'Hyderabad', 'Kochi', 'Mysuru'],
    'East India':    ['Kolkata', 'Bhubaneswar', 'Patna', 'Guwahati', 'Ranchi'],
    'West India':    ['Mumbai', 'Pune', 'Ahmedabad', 'Surat', 'Nagpur'],
    'Central India': ['Bhopal', 'Indore', 'Raipur', 'Jabalpur', 'Gwalior'],
}
SKUS           = ['Lager_500ml', 'Wheat_330ml', 'Stout_650ml', 'IPA_330ml',
                  'Premium_Lager_330ml', 'Strong_Beer_500ml', 'Light_Beer_330ml']
CHANNELS       = ['Modern Trade', 'General Trade', 'On-Premise (Bar/Restaurant)', 'E-Commerce']
SUPPLIERS      = ['BrewMalt India', 'HopCo Supplies', 'GrainSource Ltd', 'AquaPure Beverages', 'PackRight Solutions']
WAREHOUSES     = ['WH-DEL-01', 'WH-MUM-01', 'WH-BLR-01', 'WH-KOL-01', 'WH-HYD-01']
SHIPMENT_MODES = ['Road (Truck)', 'Road (Mini-Truck)', 'Rail Freight', 'Air Freight']
STATUS_OPTS    = ['Delivered On Time', 'Delivered Late', 'In Transit', 'Cancelled', 'Returned']

BASE_PRICE = {
    'Lager_500ml': 180, 'Wheat_330ml': 150, 'Stout_650ml': 220,
    'IPA_330ml': 160, 'Premium_Lager_330ml': 200,
    'Strong_Beer_500ml': 175, 'Light_Beer_330ml': 140,
}
COGS_RATIO = {'Lager_500ml': 0.42, 'Wheat_330ml': 0.38, 'Stout_650ml': 0.50,
              'IPA_330ml': 0.40, 'Premium_Lager_330ml': 0.45,
              'Strong_Beer_500ml': 0.43, 'Light_Beer_330ml': 0.36}

START_DATE = datetime(2023, 1, 1)
END_DATE   = datetime(2024, 12, 31)
N_ORDERS   = 2000

# ─── HELPER FUNCTIONS ─────────────────────────────────────────────────────────

def random_date(start, end):
    delta = end - start
    return start + timedelta(days=random.randint(0, delta.days))

def seasonal_multiplier(date):
    month = date.month
    # Peak: Oct-Dec (festive), Mar-Apr (summer starts); Low: Jul-Aug (monsoon)
    if month in [3, 4, 10, 11, 12]: return round(random.uniform(1.3, 1.7), 2)
    if month in [7, 8]:             return round(random.uniform(0.6, 0.85), 2)
    return round(random.uniform(0.9, 1.2), 2)

def lead_time(mode, region):
    base = {'Road (Truck)': 3, 'Road (Mini-Truck)': 2,
            'Rail Freight': 5, 'Air Freight': 1}[mode]
    region_penalty = {'North India': 0, 'South India': 1,
                      'East India': 2, 'West India': 0, 'Central India': 1}[region]
    return base + region_penalty + random.randint(-1, 2)

# ─── GENERATE ORDERS DATASET ──────────────────────────────────────────────────

def generate_orders():
    records = []
    for i in range(1, N_ORDERS + 1):
        region   = random.choice(REGIONS)
        city     = random.choice(CITIES[region])
        sku      = random.choice(SKUS)
        channel  = random.choice(CHANNELS)
        wh       = random.choice(WAREHOUSES)
        mode     = random.choices(SHIPMENT_MODES, weights=[55, 25, 15, 5])[0]
        order_dt = random_date(START_DATE, END_DATE)
        sm       = seasonal_multiplier(order_dt)
        qty      = int(random.randint(50, 500) * sm)
        unit_p   = BASE_PRICE[sku] + random.randint(-10, 20)
        cogs_u   = round(unit_p * COGS_RATIO[sku], 2)
        revenue  = round(qty * unit_p, 2)
        cogs     = round(qty * cogs_u, 2)
        margin   = round((revenue - cogs) / revenue * 100, 2)
        # Standard SLA Promised Delivery window: base transit + regional distance + 2-day order-to-dispatch buffer
        base_sla = {'Air Freight': 1, 'Road (Mini-Truck)': 2, 'Road (Truck)': 3, 'Rail Freight': 5}[mode]
        region_sla = {'North India': 0, 'West India': 0, 'South India': 1, 'Central India': 1, 'East India': 2}[region]
        promised_days = base_sla + region_sla + 2
        promised_dt   = order_dt + timedelta(days=promised_days)

        # Actual dispatch & transit
        dispatch_delay = random.choices([1, 2, 3], weights=[80, 15, 5])[0]
        ship_dt        = order_dt + timedelta(days=dispatch_delay)
        lt             = base_sla + region_sla + random.choices([-1, 0, 1, 2], weights=[20, 55, 20, 5])[0]
        delivery_dt    = ship_dt + timedelta(days=lt)

        # On-Time In-Full (OTIF) Calculation:
        # On-Time: delivery <= promised SLA date (~86% on-time)
        # In-Full: full order fulfilled without stockout cut (~95% fill rate)
        # OTIF = On-Time AND In-Full (~82% overall baseline vs 90%+ benchmark)
        on_time = delivery_dt <= promised_dt
        in_full = random.random() < 0.95
        otif    = 'Yes' if (on_time and in_full) else 'No'

        status = 'Delivered On Time' if on_time else ('Delivered Late' if random.random() < 0.85 else 'In Transit')
        freight_cost = round(qty * random.uniform(0.8, 2.5), 2)
        supplier = random.choice(SUPPLIERS)

        records.append({
            'Order_ID':         f'ORD-{i:05d}',
            'Order_Date':       order_dt.strftime('%Y-%m-%d'),
            'Ship_Date':        ship_dt.strftime('%Y-%m-%d'),
            'Delivery_Date':    delivery_dt.strftime('%Y-%m-%d'),
            'Promised_Date':    promised_dt.strftime('%Y-%m-%d'),
            'Region':           region,
            'City':             city,
            'SKU':              sku,
            'Channel':          channel,
            'Warehouse':        wh,
            'Supplier':         supplier,
            'Shipment_Mode':    mode,
            'Quantity_Cases':   qty,
            'Unit_Price_INR':   unit_p,
            'COGS_Per_Unit':    cogs_u,
            'Revenue_INR':      revenue,
            'COGS_INR':         cogs,
            'Gross_Margin_Pct': margin,
            'Freight_Cost_INR': freight_cost,
            'Lead_Time_Days':   lt,
            'OTIF':             otif,
            'Shipment_Status':  status,
            'Seasonal_Factor':  sm,
        })
    return pd.DataFrame(records)

# ─── GENERATE INVENTORY DATASET ───────────────────────────────────────────────

def generate_inventory():
    records = []
    for wh in WAREHOUSES:
        for sku in SKUS:
            daily_demand  = random.randint(30, 200)
            holding_cost  = round(BASE_PRICE[sku] * 0.02, 2)
            ordering_cost = random.randint(500, 2000)
            lt_days       = random.randint(2, 7)
            current_stock = random.randint(100, 2000)
            # EOQ formula: sqrt(2*D*S/H)
            eoq           = int(np.sqrt((2 * daily_demand * 365 * ordering_cost) / holding_cost))
            safety_stock  = int(daily_demand * lt_days * 1.5)
            reorder_pt    = int(daily_demand * lt_days + safety_stock)
            stockout_risk = 'High' if current_stock < reorder_pt else ('Medium' if current_stock < reorder_pt * 1.5 else 'Low')

            records.append({
                'Warehouse':         wh,
                'SKU':               sku,
                'Daily_Demand_Cases': daily_demand,
                'Annual_Demand':     daily_demand * 365,
                'Holding_Cost_INR':  holding_cost,
                'Ordering_Cost_INR': ordering_cost,
                'Lead_Time_Days':    lt_days,
                'Current_Stock':     current_stock,
                'EOQ_Cases':         eoq,
                'Safety_Stock':      safety_stock,
                'Reorder_Point':     reorder_pt,
                'Stockout_Risk':     stockout_risk,
            })
    return pd.DataFrame(records)

# ─── GENERATE DISTRIBUTOR SCORECARD ───────────────────────────────────────────

def generate_distributor_scorecard():
    distributors = [f'DIST-{r[:3].upper()}-{i:02d}' for r in REGIONS for i in range(1, 5)]
    records = []
    for dist in distributors:
        region = random.choice(REGIONS)
        otif   = round(random.uniform(60, 98), 1)
        fill   = round(random.uniform(75, 99), 1)
        damage = round(random.uniform(0.2, 3.5), 2)
        tat    = round(random.uniform(1.5, 5.0), 1)
        vol    = random.randint(5000, 50000)
        score  = round((otif * 0.35 + fill * 0.30 + (100 - damage * 10) * 0.20 + (100 - tat * 10) * 0.15), 1)
        grade  = 'A' if score >= 85 else ('B' if score >= 70 else ('C' if score >= 55 else 'D'))

        records.append({
            'Distributor_ID':     dist,
            'Region':             region,
            'OTIF_Pct':           otif,
            'Fill_Rate_Pct':      fill,
            'Damage_Rate_Pct':    damage,
            'Avg_TAT_Days':       tat,
            'Volume_Cases':       vol,
            'Performance_Score':  score,
            'Grade':              grade,
        })
    return pd.DataFrame(records)

# ─── GENERATE DEMAND FORECAST DATA ────────────────────────────────────────────

def generate_demand_forecast():
    records = []
    for sku in SKUS:
        base_demand = random.randint(5000, 20000)
        for month in range(1, 25):  # 24 months: Jan 2023 – Dec 2024
            year   = 2023 if month <= 12 else 2024
            m      = month if month <= 12 else month - 12
            dt     = datetime(year, m, 1)
            sm     = seasonal_multiplier(dt)
            actual = int(base_demand * sm + random.randint(-500, 500))
            trend  = base_demand * (1 + 0.03 * (month / 12))  # 3% annual growth
            forecast = int(trend * sm)
            error  = round(abs(actual - forecast) / actual * 100, 2)

            records.append({
                'Month':           dt.strftime('%b-%Y'),
                'SKU':             sku,
                'Actual_Demand':   actual,
                'Forecast_Demand': forecast,
                'MAPE_Pct':        error,
                'Seasonal_Factor': sm,
                'Trend_Base':      int(trend),
            })
    return pd.DataFrame(records)

# ─── MAIN ─────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    out = r'C:\Users\BIT\.gemini\antigravity\scratch\beerflow-supply-chain\data'
    os.makedirs(out, exist_ok=True)

    print("[...] Generating Orders Dataset...")
    df_orders = generate_orders()
    df_orders.to_csv(f'{out}/orders_data.csv', index=False)
    print(f"   [OK] orders_data.csv -- {len(df_orders)} records")

    print("[...] Generating Inventory Dataset...")
    df_inv = generate_inventory()
    df_inv.to_csv(f'{out}/inventory_data.csv', index=False)
    print(f"   [OK] inventory_data.csv -- {len(df_inv)} records")

    print("[...] Generating Distributor Scorecard...")
    df_dist = generate_distributor_scorecard()
    df_dist.to_csv(f'{out}/distributor_scorecard.csv', index=False)
    print(f"   [OK] distributor_scorecard.csv -- {len(df_dist)} records")

    print("[...] Generating Demand Forecast Data...")
    df_fc = generate_demand_forecast()
    df_fc.to_csv(f'{out}/demand_forecast.csv', index=False)
    print(f"   [OK] demand_forecast.csv -- {len(df_fc)} records")

    print("\n[DONE] All datasets generated successfully!")
    print(f"[PATH] Location: {out}")
