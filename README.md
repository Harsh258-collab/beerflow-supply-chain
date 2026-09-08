# 🍺 BeerFlow — Beverage Supply Chain Optimization

> **Author:** Harsh Raj Pandey
> **Domain:** Supply Chain & Logistics | FMCG / Beverage Industry
> **Tools:** Python · Microsoft Excel · Microsoft PowerPoint
> **Data:** Synthetic dataset modelled on real FMCG industry patterns | Jan 2023 – Dec 2024

---

## 📌 Why I Built This

Honestly, it started with a question I couldn't stop thinking about after reading a few beverage industry case studies: *why do FMCG companies consistently run out of stock during the exact festivals and seasons when people want to buy the most?*

The more I dug in, the more I realized the answer wasn't a lack of data — it was a lack of the right analysis connecting demand, inventory, logistics, and distributor performance together. Most companies look at these in silos. I wanted to see what happens when you look at them as one system.

That's BeerFlow — an end-to-end supply chain analysis for a beverage company operating across India, built from the ground up using Python, Excel, and PowerPoint.

It answers 4 questions I genuinely wanted solved:

1. 📈 **Can we forecast demand accurately enough to actually avoid stockouts?**
2. 📦 **Are we ordering inventory optimally — or just guessing?**
3. 🚚 **Where exactly is our logistics network losing time and money?**
4. 🏆 **Which distributors are pulling their weight — and which need a serious conversation?**

---

## 🗂️ Repository Structure

```
beerflow-supply-chain/
│
├── README.md
│
├── data/
│   ├── orders_data.csv                ← 2,000 order records (Jan 2023 – Dec 2024)
│   ├── inventory_data.csv             ← EOQ model inputs per SKU & warehouse
│   ├── distributor_scorecard.csv      ← 20 distributor KPI records
│   ├── demand_forecast.csv            ← 24-month demand & forecast (7 SKUs)
│   ├── route_delivery_data.csv        ← 1,000 last-mile delivery routes
│   ├── fleet_utilization_data.csv     ← 265 vehicles — utilization, CO2, breakdowns
│   ├── warehouse_operations_data.csv  ← 5 warehouses × 12 months operations
│   ├── cold_chain_data.csv            ← 600 cold chain temperature records
│   ├── freight_benchmark_data.csv     ← 400 shipments — Rail vs Road vs Air vs 3PL
│   └── risk_register.csv             ← 24-risk supply chain risk register
│
├── excel/
│   ├── BeerFlow_Supply_Chain_Analysis.xlsx
│   │   ├── Sheet 1: Dashboard (KPI Overview)
│   │   ├── Sheet 2: Orders Data
│   │   ├── Sheet 3: Inventory & EOQ Model
│   │   ├── Sheet 4: Distributor Scorecard
│   │   ├── Sheet 5: Demand Forecast
│   │   └── Sheet 6: OTIF & KPI Summary
│   │
│   └── BeerFlow_Logistics_DeepDive.xlsx
│       ├── Sheet 1: Logistics Dashboard
│       ├── Sheet 2: Route & Last-Mile Analysis
│       ├── Sheet 3: Fleet Utilization
│       ├── Sheet 4: Warehouse Operations
│       ├── Sheet 5: Cold Chain Monitoring
│       ├── Sheet 6: Freight Benchmarking (Modal Analysis)
│       ├── Sheet 7: Supply Chain Risk Register
│       └── Sheet 8: Logistics KPI Tracker (Actual vs Target, RAG)
│
├── powerpoint/
│   └── BeerFlow_Strategy_Deck.pptx   ← 14-slide executive strategy deck
│
├── insights/
│   ├── key_findings.md               ← Overall supply chain findings
│   └── logistics_deep_dive.md        ← Detailed logistics analysis
│
└── scripts/
    ├── generate_data.py              ← Core supply chain data generator
    ├── generate_logistics_data.py    ← Logistics datasets generator
    ├── build_excel.py                ← Supply chain Excel builder
    ├── build_logistics_excel.py      ← Logistics Excel builder
    └── build_ppt.py                  ← PowerPoint deck builder
```

---

## 📊 The Analysis — What I Did & What I Found

### 1. Demand Forecasting
I built a 24-month SKU-level demand model with seasonal adjustment factors based on real FMCG seasonality patterns — festive demand spikes, monsoon troughs, and year-on-year growth trends.

**What I found:** Demand swings 30–70% above baseline in Oct–Dec and Mar–Apr. But ordering patterns barely adjust. That mismatch is the root cause of most stockouts I observed in the data.

I measured forecast accuracy using MAPE (Mean Absolute Percentage Error). Most SKUs came in around 12–18% MAPE — acceptable but not great. Lager 500ml was the most predictable; IPA 330ml was the hardest to forecast, likely due to its niche demand pattern.

### 2. Inventory Optimization (EOQ)
For every SKU × warehouse combination, I calculated:
- **EOQ** = sqrt(2 × Annual Demand × Ordering Cost / Holding Cost)
- **Safety Stock** = Daily Demand × Lead Time × 1.5
- **Reorder Point** = Daily Demand × Lead Time + Safety Stock

**What I found:** East India warehouses hold 30% below recommended safety stock. Stout 650ml is being over-ordered by ~45% above EOQ across the board — this ties up working capital and inflates holding costs unnecessarily.

### 3. Logistics — Last-Mile & Route Analysis
I analysed 1,000 delivery routes covering vehicle utilization, cost per case, on-time delivery, and CO2 emissions.

**What I found:** Vehicles run at ~72% capacity on average — 28% of capacity wasted per trip. Large 32-ft trucks are being deployed for short urban drops under 50km, where a Mini Truck would cost 35–40% less per case. This is a routing and vehicle allocation problem, not a fleet size problem.

### 4. Cold Chain
Temperature compliance is non-negotiable for beverages. I tracked 600 cold chain records across 5 supply chain stages.

**What I found:** 18% breach rate — far above the <5% target. The retail outlet stage is the worst offender (30%+ breach). We control the brewery and the truck, but once product hits a small retailer without proper refrigeration, quality is at risk.

### 5. Freight Modal Benchmarking
Compared road, rail, air, and 3PL across cost per case, transit time, OTD%, damage rate, and CO2.

**What I found:** Rail freight costs 44% less per case than road trucking for distances above 800km — but it's only 15% utilized. Shifting 10% more volume to rail could save INR 15–20 Lakhs annually.

### 6. Distributor Scorecard
I built a weighted KPI scoring model (OTIF 35%, Fill Rate 30%, Damage 20%, TAT 15%) and graded 20 distributors A–D.

**What I found:** 40% of distributors fall in Grade C or D. These aren't necessarily bad businesses — most are concentrated in East and Central India where infrastructure constraints suppress performance. The fix is a structured Distributor Performance Improvement Program (DPIP), not replacement.

### 7. Supply Chain Risk Register
24 risks identified across 6 categories (Demand, Supplier, Logistics, Inventory, Distributor, Regulatory/External), each scored on Likelihood × Impact.

**What I found:** The 3 highest-priority risks are festive season fleet breakdowns (Score 15), cold chain breaches in summer (Score 16), and key malt supplier disruption (Score 10). Each has a specific mitigation strategy documented.

---

## 🔑 Key Metrics

| Metric | Value |
|---|---|
| Total Orders Analysed | 2,000 |
| Delivery Routes Analysed | 1,000 |
| Fleet Vehicles Tracked | 265 |
| Cold Chain Records | 600 |
| Freight Shipments Benchmarked | 400 |
| Supply Chain Risks Documented | 24 |
| SKUs | 7 |
| Regions | 5 |
| Warehouses | 5 |
| Distributors | 20 |
| Data Period | Jan 2023 – Dec 2024 |
| Estimated Annual Savings Identified | INR 55–80 Lakhs |

---

## 💡 Top 3 Recommendations

**1. Dynamic Safety Stock + Seasonal Inventory Buffer**
Pre-build 25–30% extra stock for top SKUs by September each year. Implement EOQ-guided ordering to cut excess Stout 650ml stock by 45%.
*Expected: 40% fewer stockouts | +8–10% OTIF improvement*

**2. East India Logistics Hub + Rail Freight Expansion**
Set up a Kolkata consolidation hub. Shift bulk inter-city volume to rail freight (15% → 25%).
*Expected: Lead time 6.2 → 3.5 days | INR 15–20L freight savings*

**3. Distributor Performance Improvement Program (DPIP)**
QBRs for Grade C/D distributors. Tiered incentives for Grade A. Digital order tools for bottom quartile.
*Expected: OTIF 76% → 88% | Grade D count drops 50%*

---

## 🛠️ Tools Used

| Tool | How I Used It |
|---|---|
| **Python (pandas, numpy)** | Data generation, statistical modelling, seasonal simulation |
| **Microsoft Excel** | Dashboard, EOQ model, KPI tracker, distributor scorecard, freight analysis |
| **Microsoft PowerPoint** | 14-slide executive strategy deck |
| **EOQ Model** | `sqrt(2 × D × S / H)` — classic inventory optimization |
| **MAPE** | `|Actual − Forecast| / Actual × 100` — forecasting accuracy |
| **Weighted Scoring** | OTIF 35% + Fill Rate 30% + Damage 20% + TAT 15% |

---

## 🚀 Run It Yourself

```bash
pip install pandas openpyxl python-pptx numpy

# Generate all datasets
python scripts/generate_data.py
python scripts/generate_logistics_data.py

# Build Excel workbooks
python scripts/build_excel.py
python scripts/build_logistics_excel.py

# Build PowerPoint
python scripts/build_ppt.py
```

---

## 📬 Connect

- **GitHub:** [github.com/Harsh258-collab](https://github.com/Harsh258-collab)
- **LinkedIn:** [Harsh Raj Pandey](https://www.linkedin.com/in/harsh-raj-pandey-1a0319325)

---

> *"Good supply chains don't just move products — they build trust."*
