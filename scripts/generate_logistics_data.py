"""
BeerFlow — Logistics Deep-Dive Dataset Generator
Author: Harsh Raj Pandey
Generates 6 new logistics-specific datasets:
  1. Route & Last-Mile Delivery
  2. Fleet & Vehicle Utilization
  3. Warehouse Operations
  4. Cold Chain Monitoring
  5. Freight Cost Benchmarking
  6. Supply Chain Risk Register
"""

import pandas as pd
import numpy as np
import random
from datetime import datetime, timedelta
import os

np.random.seed(99)
random.seed(99)

BASE = r'C:\Users\BIT\.gemini\antigravity\scratch\beerflow-supply-chain'
OUT  = f'{BASE}\\data'

REGIONS    = ['North India', 'South India', 'East India', 'West India', 'Central India']
HUBS       = {
    'North India':   'Delhi Hub',
    'South India':   'Bengaluru Hub',
    'East India':    'Kolkata Hub',
    'West India':    'Mumbai Hub',
    'Central India': 'Bhopal Hub',
}
CITIES = {
    'North India':   ['Delhi', 'Lucknow', 'Chandigarh', 'Jaipur', 'Amritsar', 'Meerut', 'Agra'],
    'South India':   ['Bengaluru', 'Chennai', 'Hyderabad', 'Kochi', 'Mysuru', 'Coimbatore', 'Vizag'],
    'East India':    ['Kolkata', 'Bhubaneswar', 'Patna', 'Guwahati', 'Ranchi', 'Siliguri', 'Jamshedpur'],
    'West India':    ['Mumbai', 'Pune', 'Ahmedabad', 'Surat', 'Nagpur', 'Nashik', 'Vadodara'],
    'Central India': ['Bhopal', 'Indore', 'Raipur', 'Jabalpur', 'Gwalior', 'Ujjain', 'Bilaspur'],
}
VEHICLE_TYPES = ['32-ft Truck', '20-ft Truck', 'Mini Truck (14-ft)', 'Tata Ace', 'Tempo']
VEHICLE_CAP   = {'32-ft Truck': 800, '20-ft Truck': 500,
                 'Mini Truck (14-ft)': 250, 'Tata Ace': 120, 'Tempo': 80}
WAREHOUSES    = ['WH-DEL-01', 'WH-MUM-01', 'WH-BLR-01', 'WH-KOL-01', 'WH-HYD-01']
SKUS          = ['Lager_500ml', 'Wheat_330ml', 'Stout_650ml', 'IPA_330ml',
                 'Premium_Lager_330ml', 'Strong_Beer_500ml', 'Light_Beer_330ml']
START         = datetime(2024, 1, 1)
END           = datetime(2024, 12, 31)

def rand_date():
    return START + timedelta(days=random.randint(0, 364))

# ══════════════════════════════════════════════════════════════════
# DATASET 1: ROUTE & LAST-MILE DELIVERY
# ══════════════════════════════════════════════════════════════════

def gen_route_data():
    records = []
    for i in range(1, 1001):
        region   = random.choice(REGIONS)
        hub      = HUBS[region]
        dest     = random.choice(CITIES[region])
        vehicle  = random.choices(VEHICLE_TYPES, weights=[20, 25, 30, 15, 10])[0]
        cap      = VEHICLE_CAP[vehicle]
        load     = random.randint(int(cap * 0.45), cap)
        util_pct = round(load / cap * 100, 1)
        dist_km  = random.randint(15, 420)
        speed    = random.randint(28, 65)          # avg kmph
        travel_h = round(dist_km / speed, 2)
        stops    = random.randint(2, 12)
        time_per_stop = random.uniform(0.25, 0.75)
        total_h  = round(travel_h + stops * time_per_stop, 2)
        fuel_l   = round(dist_km / random.uniform(5, 9), 2)
        fuel_cost= round(fuel_l * 94.5, 2)         # INR/litre avg
        toll_cost= round(dist_km * random.uniform(0.8, 2.2), 2)
        driver_cost = round(total_h * random.uniform(120, 200), 2)
        total_cost  = round(fuel_cost + toll_cost + driver_cost, 2)
        cost_per_case = round(total_cost / load, 2)
        delay_min= max(0, int(random.gauss(12, 25)))
        on_time  = 'Yes' if delay_min < 30 else 'No'
        date     = rand_date()

        records.append({
            'Route_ID':         f'RT-{i:04d}',
            'Date':             date.strftime('%Y-%m-%d'),
            'Region':           region,
            'Origin_Hub':       hub,
            'Destination_City': dest,
            'Vehicle_Type':     vehicle,
            'Vehicle_Capacity_Cases': cap,
            'Load_Cases':       load,
            'Utilization_Pct':  util_pct,
            'Distance_KM':      dist_km,
            'Avg_Speed_KMPH':   speed,
            'Travel_Time_Hrs':  travel_h,
            'No_of_Stops':      stops,
            'Total_Time_Hrs':   total_h,
            'Fuel_Litres':      fuel_l,
            'Fuel_Cost_INR':    fuel_cost,
            'Toll_Cost_INR':    toll_cost,
            'Driver_Cost_INR':  driver_cost,
            'Total_Route_Cost_INR': total_cost,
            'Cost_Per_Case_INR': cost_per_case,
            'Delivery_Delay_Min': delay_min,
            'On_Time_Delivery': on_time,
        })
    return pd.DataFrame(records)


# ══════════════════════════════════════════════════════════════════
# DATASET 2: FLEET & VEHICLE UTILIZATION
# ══════════════════════════════════════════════════════════════════

def gen_fleet_data():
    records = []
    fleet_id = 1
    for wh in WAREHOUSES:
        region = [r for r, h in HUBS.items() if wh[3:6] in h.upper() or
                  (wh == 'WH-DEL-01' and r == 'North India') or
                  (wh == 'WH-MUM-01' and r == 'West India') or
                  (wh == 'WH-BLR-01' and r == 'South India') or
                  (wh == 'WH-KOL-01' and r == 'East India') or
                  (wh == 'WH-HYD-01' and r == 'South India')][0]
        for vtype in VEHICLE_TYPES:
            count = {'32-ft Truck': 8, '20-ft Truck': 10,
                     'Mini Truck (14-ft)': 15, 'Tata Ace': 12, 'Tempo': 8}[vtype]
            for v in range(1, count + 1):
                age_yrs  = random.randint(1, 9)
                trips_pm = random.randint(8, 28)
                km_pm    = trips_pm * random.randint(80, 350)
                maint_cost = round(age_yrs * random.uniform(1500, 4000), 0)
                breakdown  = random.randint(0, 3) if age_yrs > 5 else random.randint(0, 1)
                util_days  = random.randint(18, 28)
                idle_days  = 30 - util_days
                avg_load   = round(random.uniform(0.55, 0.95) * VEHICLE_CAP[vtype], 0)
                fuel_eff   = round(random.uniform(4.5, 9.5), 1)
                emission   = round(km_pm * (1/fuel_eff) * 2.68, 1)  # kg CO2

                records.append({
                    'Vehicle_ID':      f'VH-{fleet_id:04d}',
                    'Warehouse':       wh,
                    'Region':          region,
                    'Vehicle_Type':    vtype,
                    'Capacity_Cases':  VEHICLE_CAP[vtype],
                    'Age_Years':       age_yrs,
                    'Trips_Per_Month': trips_pm,
                    'KM_Per_Month':    km_pm,
                    'Avg_Load_Cases':  int(avg_load),
                    'Load_Utilization_Pct': round(avg_load / VEHICLE_CAP[vtype] * 100, 1),
                    'Utilization_Days_PM': util_days,
                    'Idle_Days_PM':    idle_days,
                    'Maintenance_Cost_INR_PM': int(maint_cost),
                    'Breakdowns_PM':   breakdown,
                    'Fuel_Efficiency_KMPL': fuel_eff,
                    'CO2_Emission_KG_PM': emission,
                    'Status':          'Active' if breakdown == 0 else ('In Repair' if breakdown >= 2 else 'Active'),
                })
                fleet_id += 1
    return pd.DataFrame(records)


# ══════════════════════════════════════════════════════════════════
# DATASET 3: WAREHOUSE OPERATIONS
# ══════════════════════════════════════════════════════════════════

def gen_warehouse_ops():
    records = []
    months = pd.date_range('2024-01', periods=12, freq='MS')
    for wh in WAREHOUSES:
        capacity = random.randint(8000, 20000)  # total case capacity
        for month in months:
            receipts    = random.randint(3000, 8000)
            dispatches  = random.randint(2800, 7800)
            closing_stk = random.randint(2000, capacity)
            util_pct    = round(closing_stk / capacity * 100, 1)
            pick_acc    = round(random.uniform(96.5, 99.8), 2)
            order_cycle = round(random.uniform(1.2, 3.8), 2)   # hours
            putaway_t   = round(random.uniform(0.5, 2.0), 2)
            labour_hrs  = random.randint(800, 2500)
            overtime_hrs= random.randint(0, 200)
            damage_cases= random.randint(5, 80)
            damage_pct  = round(damage_cases / dispatches * 100, 3)
            dock_turns  = random.randint(12, 35)           # per day
            orders_proc = random.randint(200, 600)
            returns     = random.randint(5, 60)
            shrinkage   = round(random.uniform(0.05, 0.8), 3)

            records.append({
                'Warehouse':          wh,
                'Month':              month.strftime('%b-%Y'),
                'Total_Capacity_Cases': capacity,
                'Receipts_Cases':     receipts,
                'Dispatches_Cases':   dispatches,
                'Closing_Stock_Cases': closing_stk,
                'Space_Utilization_Pct': util_pct,
                'Pick_Accuracy_Pct':  pick_acc,
                'Order_Cycle_Time_Hrs': order_cycle,
                'Putaway_Time_Hrs':   putaway_t,
                'Labour_Hours':       labour_hrs,
                'Overtime_Hours':     overtime_hrs,
                'Damaged_Cases':      damage_cases,
                'Damage_Rate_Pct':    damage_pct,
                'Dock_Turns_Per_Day': dock_turns,
                'Orders_Processed':   orders_proc,
                'Returns_Cases':      returns,
                'Shrinkage_Pct':      shrinkage,
            })
    return pd.DataFrame(records)


# ══════════════════════════════════════════════════════════════════
# DATASET 4: COLD CHAIN MONITORING
# ══════════════════════════════════════════════════════════════════

def gen_cold_chain():
    records = []
    for i in range(1, 601):
        region      = random.choice(REGIONS)
        wh          = random.choice(WAREHOUSES)
        sku         = random.choice(SKUS)
        date        = rand_date()
        stage       = random.choice(['Brewery', 'Primary Warehouse', 'Regional Hub', 'Distributor', 'Retail Outlet'])
        set_temp    = random.choice([2, 4, 6, 8])          # target °C
        avg_temp    = round(set_temp + random.gauss(0.5, 1.8), 1)
        min_temp    = round(avg_temp - random.uniform(0.5, 2.5), 1)
        max_temp    = round(avg_temp + random.uniform(0.5, 3.5), 1)
        breach      = 'Yes' if max_temp > set_temp + 4 or min_temp < set_temp - 2 else 'No'
        breach_dur  = random.randint(15, 180) if breach == 'Yes' else 0
        humidity    = round(random.uniform(55, 85), 1)
        product_ok  = 'Compliant' if breach == 'No' else random.choice(['Compliant', 'Quarantined', 'Disposed'])
        transport_h = round(random.uniform(1, 18), 1)
        cost_impact = round(breach_dur * random.uniform(50, 200), 0) if breach == 'Yes' else 0

        records.append({
            'Cold_Chain_ID':      f'CC-{i:04d}',
            'Date':               date.strftime('%Y-%m-%d'),
            'Region':             region,
            'Warehouse':          wh,
            'SKU':                sku,
            'Supply_Chain_Stage': stage,
            'Target_Temp_C':      set_temp,
            'Avg_Temp_C':         avg_temp,
            'Min_Temp_C':         min_temp,
            'Max_Temp_C':         max_temp,
            'Temperature_Breach': breach,
            'Breach_Duration_Min': breach_dur,
            'Humidity_Pct':       humidity,
            'Transport_Hours':    transport_h,
            'Product_Status':     product_ok,
            'Cost_Impact_INR':    int(cost_impact),
        })
    return pd.DataFrame(records)


# ══════════════════════════════════════════════════════════════════
# DATASET 5: FREIGHT COST BENCHMARKING
# ══════════════════════════════════════════════════════════════════

def gen_freight_benchmark():
    records = []
    modes   = ['Road (Truck)', 'Road (Mini-Truck)', 'Rail Freight', 'Air Freight', '3PL Partner']
    for i in range(1, 401):
        region   = random.choice(REGIONS)
        mode     = random.choice(modes)
        dist     = random.randint(50, 1200)
        weight   = random.randint(500, 12000)   # kg
        cases    = random.randint(50, 800)
        base_rate = {
            'Road (Truck)': 2.8, 'Road (Mini-Truck)': 3.5,
            'Rail Freight': 1.4, 'Air Freight': 12.0, '3PL Partner': 3.2
        }[mode]
        freight  = round(dist * weight / 1000 * base_rate * random.uniform(0.85, 1.25), 2)
        per_case = round(freight / cases, 2)
        transit  = {
            'Road (Truck)': random.randint(2, 6), 'Road (Mini-Truck)': random.randint(1, 4),
            'Rail Freight': random.randint(4, 9), 'Air Freight': 1, '3PL Partner': random.randint(2, 5)
        }[mode]
        damage   = round(random.uniform(0, 2.5) if mode != 'Air Freight' else random.uniform(0, 0.5), 2)
        on_time  = round(random.uniform(70, 99), 1)
        co2      = round(dist * weight / 1000 * {
            'Road (Truck)': 0.12, 'Road (Mini-Truck)': 0.15,
            'Rail Freight': 0.04, 'Air Freight': 0.60, '3PL Partner': 0.13
        }[mode], 2)
        date     = rand_date()

        records.append({
            'Shipment_ID':         f'FR-{i:04d}',
            'Date':                date.strftime('%Y-%m-%d'),
            'Region':              region,
            'Freight_Mode':        mode,
            'Distance_KM':         dist,
            'Weight_KG':           weight,
            'Cases_Shipped':       cases,
            'Freight_Cost_INR':    freight,
            'Cost_Per_Case_INR':   per_case,
            'Cost_Per_KM_INR':     round(freight / dist, 2),
            'Transit_Days':        transit,
            'Damage_Rate_Pct':     damage,
            'On_Time_Pct':         on_time,
            'CO2_Emission_KG':     co2,
        })
    return pd.DataFrame(records)


# ══════════════════════════════════════════════════════════════════
# DATASET 6: SUPPLY CHAIN RISK REGISTER
# ══════════════════════════════════════════════════════════════════

def gen_risk_register():
    risks = [
        # (Category, Risk Description, Likelihood 1-5, Impact 1-5, Mitigation)
        ('Demand Risk',     'Sudden demand spike due to unplanned event (cricket final, festival)',       4, 5, 'Maintain 15% buffer stock for top SKUs during high-risk periods'),
        ('Demand Risk',     'Demand forecast error exceeding 20% for slow-moving SKUs',                  3, 3, 'Switch to weekly rolling forecast; review MAPE monthly'),
        ('Demand Risk',     'Monsoon season demand trough deeper than expected',                         3, 2, 'Pre-negotiate flexible order quantities with distributors'),
        ('Supplier Risk',   'Key malt supplier disruption due to crop failure or logistics issue',       2, 5, 'Dual-source malt from at least 2 approved suppliers'),
        ('Supplier Risk',   'Packaging material shortage (cans, bottles) from single vendor',            2, 4, 'Maintain 30-day packaging buffer; qualify alternate vendor'),
        ('Supplier Risk',   'Supplier quality failure causing batch rejection',                          2, 4, '100% incoming quality check; supplier scorecard monitoring'),
        ('Logistics Risk',  'Fuel price spike increasing freight cost by 15%+',                         4, 3, 'Hedge freight contracts quarterly; shift volume to rail'),
        ('Logistics Risk',  'Major road/highway blockage disrupting North-South corridor',               2, 4, 'Pre-identify alternate routes; maintain regional safety stock'),
        ('Logistics Risk',  'Fleet breakdown during peak season (Oct-Dec)',                              3, 5, 'Preventive maintenance schedule; tie-up with backup fleet vendors'),
        ('Logistics Risk',  'Driver shortage during festive season',                                     3, 3, 'Retain drivers with festive bonus; 3PL backup agreements'),
        ('Logistics Risk',  'Cold chain temperature breach during summer months',                        4, 4, 'Real-time IoT temperature monitoring; SLA with cold chain vendors'),
        ('Logistics Risk',  'Port/rail strike disrupting inter-city movement',                           2, 5, 'Emergency road freight plan; 15-day buffer stock at key hubs'),
        ('Inventory Risk',  'Stockout of premium SKU during Diwali/New Year',                           3, 5, 'EOQ-based pre-season build with safety stock buffer'),
        ('Inventory Risk',  'Expiry/quality degradation of slow-moving SKU in warehouse',               2, 3, 'FEFO (First Expiry First Out) picking policy; monthly SKU audit'),
        ('Inventory Risk',  'Overstocking due to over-forecast causing high holding cost',               3, 3, 'Cap safety stock at 30 days; trigger liquidation at 45 days'),
        ('Distributor Risk','Top distributor (Grade A) losing license or exiting market',                1, 5, 'Identify and develop 2nd-tier distributors in each zone'),
        ('Distributor Risk','Distributor financial distress causing order default',                       2, 4, 'Monthly credit monitoring; cap credit limit at 30-day value'),
        ('Distributor Risk','Distributor selling competitor products over company products',              3, 3, 'Exclusive zone agreements; volume-linked incentive scheme'),
        ('Regulatory Risk', 'State government changing alcohol distribution regulations',                2, 5, 'Legal monitoring team; compliance buffer in business plan'),
        ('Regulatory Risk', 'GST rate change on beverages affecting margins',                            2, 3, 'Build 5% margin buffer; pass-through clause in distributor contracts'),
        ('External Risk',   'Natural disaster (flood/cyclone) disrupting East/South India hub',          2, 5, 'Geo-distributed hub strategy; disaster recovery plan'),
        ('External Risk',   'Counterfeit product flooding market in a region',                           2, 4, 'Track-and-trace barcoding; rapid market audit protocol'),
        ('Technology Risk', 'WMS (Warehouse Mgmt System) downtime during peak dispatch',                 2, 4, 'Manual backup SOP; daily data backup; vendor SLA < 2 hr response'),
        ('Technology Risk', 'GPS tracker failure on cold chain vehicles',                                3, 3, 'Dual-device installation; driver check-in protocol every 2 hours'),
    ]
    records = []
    for i, (cat, desc, like, imp, mit) in enumerate(risks, 1):
        risk_score = like * imp
        risk_level = 'Critical' if risk_score >= 15 else ('High' if risk_score >= 9 else ('Medium' if risk_score >= 4 else 'Low'))
        owner      = random.choice(['Logistics Head', 'Supply Chain Manager', 'Procurement Lead',
                                    'Sales Director', 'Warehouse Manager', 'Finance Controller'])
        status     = random.choice(['Open', 'Mitigated', 'Monitoring', 'Open'])
        records.append({
            'Risk_ID':            f'RSK-{i:03d}',
            'Category':           cat,
            'Risk_Description':   desc,
            'Likelihood_1to5':    like,
            'Impact_1to5':        imp,
            'Risk_Score':         risk_score,
            'Risk_Level':         risk_level,
            'Mitigation_Strategy': mit,
            'Risk_Owner':         owner,
            'Status':             status,
        })
    return pd.DataFrame(records)


# ══════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════

if __name__ == '__main__':
    os.makedirs(OUT, exist_ok=True)

    datasets = [
        ('gen_route_data',        'route_delivery_data.csv',     "Route & Last-Mile Delivery"),
        ('gen_fleet_data',        'fleet_utilization_data.csv',  "Fleet & Vehicle Utilization"),
        ('gen_warehouse_ops',     'warehouse_operations_data.csv',"Warehouse Operations"),
        ('gen_cold_chain',        'cold_chain_data.csv',         "Cold Chain Monitoring"),
        ('gen_freight_benchmark', 'freight_benchmark_data.csv',  "Freight Cost Benchmarking"),
        ('gen_risk_register',     'risk_register.csv',           "Supply Chain Risk Register"),
    ]

    for func_name, filename, label in datasets:
        print(f"[...] Generating {label}...")
        df = globals()[func_name]()
        df.to_csv(f'{OUT}/{filename}', index=False)
        print(f"   [OK] {filename} -- {len(df)} records")

    print("\n[DONE] All logistics datasets generated!")
