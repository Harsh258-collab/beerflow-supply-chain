# BeerFlow — Logistics Development Deep Dive

**Author:** Harsh Raj Pandey  
**Focus:** End-to-End Logistics Operations Analysis  
**Scope:** Last-Mile Delivery | Fleet | Warehouse | Cold Chain | Freight | Risk

---

## 1. Last-Mile Delivery Analysis

Last-mile is where most logistics cost is lost — and where the customer experience is made or broken.

### What I found:
- **1,000 delivery routes** analysed across 5 regions and 5 vehicle types
- Average vehicle load utilization: **~72%** — meaning 28% of capacity goes empty every trip
- Average cost per case delivered: **INR 16–20** depending on region
- On-time delivery rate: **~79%** — against an industry target of 90%+
- East India routes average **38 minutes of delay** per route — 2x the national average

### Root cause of low utilization:
Routes are planned by geography (same city) rather than by load optimization. A Tata Ace doing 3 drops at 40% capacity could be consolidated into 1 Mini Truck doing 6 drops at 85% capacity — cutting cost per case by ~35%.

### Recommendation:
Implement a **TMS (Transport Management System)** with dynamic route optimization. Even a basic Excel-based load planner can improve utilization by 15–20% in the first 3 months.

---

## 2. Fleet & Vehicle Utilization

### What I found:
- **265 vehicles** across 5 warehouses — mix of 32-ft trucks, 20-ft trucks, mini trucks, Tata Ace, and Tempos
- Vehicles older than 5 years show **2–3x higher breakdown rates** — causing unplanned downtime during peak periods
- 32-ft trucks have the **best cost-per-case** ratio for bulk inter-city routes but are over-used for short urban drops
- Fleet CO2 emissions: **~180 tonnes/month** — reduction target of 20% achievable by shifting 10% volume to CNG vehicles

### Vehicle mis-match problem:
Large 32-ft trucks are being deployed for short urban routes (under 50 km) where a Tata Ace or Mini Truck would cost 40% less per trip. Right-sizing vehicles to route type is the single biggest fleet cost lever.

### Recommendation:
- Retire vehicles older than 8 years in high-breakdown categories
- Introduce **route-to-vehicle matching rules**: 32-ft trucks → bulk inter-hub; Mini/Ace → urban last-mile
- Pilot 5 CNG Mini Trucks in Delhi and Mumbai to measure fuel savings

---

## 3. Warehouse Operations

### What I found:
- **60 monthly records** across 5 warehouses (full year 2024)
- Avg pick accuracy: **98.1%** — good but below the 99%+ target for beverage FMCG
- Avg order cycle time: **2.4 hours** — target is under 2 hours
- WH-KOL-01 (Kolkata) has the highest damage rate: **0.62%** — 2x other warehouses
- Space utilization peaks at **88%** in October-November — creating bottlenecks during festive dispatch

### The Kolkata problem:
Kolkata warehouse shows consistently high damage and low pick accuracy. Investigation points to manual picking without barcode scanning, and stacking of heavy SKUs (Stout 650ml) on top of lighter SKUs (Wheat 330ml).

### Recommendation:
1. **FEFO policy** (First Expiry First Out) across all warehouses — standardize picking protocol
2. **Barcode scan at pick** for 100% order verification — reduces mis-picks by ~60%
3. Pre-season space clearing in September: move slow-moving SKUs to auxiliary storage to free peak capacity
4. Dedicated zone for Kolkata — retrain warehouse staff on stacking rules

---

## 4. Cold Chain Monitoring

This is non-negotiable in beverages — temperature breach = product quality failure.

### What I found:
- **600 cold chain records** across 5 supply chain stages
- Temperature breach rate: **~18%** — far above the acceptable <5% threshold
- Highest breach stage: **Retail Outlet** (30%+ breach rate) — products sitting in warm display areas
- Summer months (April-June) see breach rates 2.5x higher than winter
- Total cost impact of breaches: **INR 2.8–3.5 Lakhs** for the monitored sample

### The retail gap:
We control the brewery, the warehouse, and even the truck — but once the product reaches a retail outlet or bar, temperature control collapses. Most small retailers don't have proper refrigeration or set it too warm to save electricity.

### Recommendation:
1. **IoT temperature sensors** in all company-owned refrigeration units at key retailers
2. **Cold chain SLA** in distributor contracts — with penalty clause for breach above 8%
3. Real-time temperature monitoring dashboard — alert system for breach > 30 minutes
4. Seasonal audits of cold chain infrastructure at Top 50 retail accounts in summer months

---

## 5. Freight Cost Benchmarking — Modal Analysis

### What I found (400 shipments benchmarked):

| Mode | Avg Cost/Case | Transit Days | CO2/KG | OTD% |
|---|---|---|---|---|
| Rail Freight | INR 8.20 | 5.5 days | 0.04 | 82% |
| Road (Truck) | INR 14.80 | 3.2 days | 0.12 | 88% |
| Road (Mini-Truck) | INR 18.50 | 2.1 days | 0.15 | 91% |
| 3PL Partner | INR 16.20 | 3.8 days | 0.13 | 85% |
| Air Freight | INR 68.40 | 1.0 days | 0.60 | 96% |

### The rail opportunity:
Rail freight costs **44% less per case** than road trucking. For bulk Lager and Strong Beer shipments from Mumbai/Delhi to Kolkata and Bhopal, rail is entirely viable — transit time is higher (5–6 days) but with proper demand planning, this lead time is manageable.

### Recommendation:
- Increase rail freight share from 15% → 25% for bulk inter-city SKU movement
- Negotiate annual volume contracts with Indian Railways for temperature-controlled containers
- Reserve air freight strictly for: (a) stockout emergency resupply, (b) new product launch quantities
- **Projected annual saving: INR 15–20 Lakhs** at 25% rail share

---

## 6. Supply Chain Risk Register

### Risk Heat Map Summary:

| Risk Level | Count | Key Risks |
|---|---|---|
| **Critical** | 3 | Peak season fleet breakdown, Key malt supplier disruption, Stockout during Diwali |
| **High** | 8 | Cold chain breach, Port/rail strike, Top distributor exit, Financial distress |
| **Medium** | 10 | Fuel price spike, Over-forecast, Driver shortage, WMS downtime |
| **Low** | 3 | GST change, Demand trough, Documentation compliance |

### The 3 risks I'd prioritize:
1. **Fleet breakdown in Oct-Dec** (Likelihood 3, Impact 5 = Score 15): This is the highest commercial risk. A truck breakdown during Diwali means lost sales that cannot be recovered. Mitigation: mandatory pre-season fleet servicing by September 30, and a backup fleet agreement with a regional 3PL.

2. **Cold chain breach during summer** (Likelihood 4, Impact 4 = Score 16): High frequency, high cost. IoT monitoring is the key investment here.

3. **Top distributor exit** (Likelihood 1, Impact 5 = Score 5): Low probability but catastrophic. Every region should have a Grade B distributor actively being developed as a backup.

---

## Logistics KPI Summary — Current vs Target

| KPI | Current | Target | Gap | Priority |
|---|---|---|---|---|
| On-Time Delivery % | 79% | 90% | -11% | 🔴 HIGH |
| Vehicle Utilization % | 72% | 85% | -13% | 🔴 HIGH |
| Cost per Case (INR) | 17.40 | 15.00 | -2.40 | 🟡 MED |
| Pick Accuracy % | 98.1% | 99% | -0.9% | 🟡 MED |
| Cold Chain Breach % | 18% | <5% | -13% | 🔴 HIGH |
| Rail Freight Share % | 15% | 25% | -10% | 🟡 MED |
| Fleet Breakdown/Month | 24 | <20 | -4 | 🟡 MED |
| Order Cycle Time (Hrs) | 2.4 | 2.0 | -0.4 | 🟢 LOW |

---

*"Logistics is the backbone of every FMCG company. The difference between a good supply chain and a great one is measured in minutes, degrees, and centimetres — not just cost."*  
— Harsh Raj Pandey
