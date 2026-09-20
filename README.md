# 🍺 BeerFlow — Beverage Supply Chain & Logistics Analytics

> **Author:** Harsh Raj Pandey  
> **Role Focus:** Supply Chain & Logistics Operations Analyst  
> **Target Domain:** FMCG / Beverage Distribution (AB InBev)  
> **Tools:** Python (Pandas, NumPy) · Microsoft Excel · Microsoft PowerPoint  
> **Dataset:** 2,000 simulated orders & 1,000 delivery routes modeled on Indian beverage distribution patterns (Jan 2023 – Dec 2024)

---

## 📌 Project Overview

Developed an end-to-end beverage supply chain analytics project to analyze demand forecasting, inventory optimization, distributor performance, and logistics KPIs using Python and Excel.

The project models a multi-echelon distribution network across 5 Indian regions (North, South, East, West, Central), 7 SKUs, 5 regional warehouses, and 20 distributors. It evaluates operational bottlenecks across four functional areas:

1. **Demand Forecasting:** Quantifying seasonal demand swings (+70% festival peak vs -40% monsoon trough) and evaluating SKU-level forecast accuracy using MAPE.
2. **Inventory Optimization:** Calculating Economic Order Quantity (EOQ), Safety Stock, and Reorder Points (ROP) across 35 warehouse-SKU nodes to reduce holding costs while mitigating stockouts.
3. **Logistics & Route Optimization:** Analyzing 1,000 delivery routes, identifying a 28% vehicle capacity leak, and evaluating multi-modal freight arbitrage (Road vs Rail).
4. **Distributor Governance:** Developing a 4-pillar weighted composite scorecard (OTIF, Fill Rate, Damage, TAT) to tier 20 regional partners into Grades A through D.

---

## 🗂️ Repository Structure

```
beerflow-supply-chain/
│
├── README.md                          ← Project overview and methodology
│
├── data/                              ← Simulated transactional datasets (CSV)
│   ├── orders_data.csv                ← 2,000 order records with lead time & OTIF status
│   ├── inventory_data.csv             ← 35 SKU-warehouse nodes with EOQ inputs
│   ├── distributor_scorecard.csv      ← 20 distributor KPI records & composite scores
│   ├── demand_forecast.csv            ← 24-month demand and forecast data (7 SKUs)
│   ├── route_delivery_data.csv        ← 1,000 last-mile routes (vehicle utilization, cost/case)
│   ├── fleet_utilization_data.csv     ← 265 fleet assets (age, breakdowns, maintenance, CO2)
│   ├── warehouse_operations_data.csv  ← 5 warehouses × 12 months operations (throughput, damage)
│   ├── cold_chain_data.csv            ← 600 thermal readings across 5 supply chain stages
│   ├── freight_benchmark_data.csv     ← 400 shipments comparing Road, Rail, Air, and 3PL
│   └── risk_register.csv             ← 24-risk supply chain register (Likelihood × Impact)
│
├── excel/                             ← Analytical Excel workbooks
│   ├── BeerFlow_Supply_Chain_Analysis.xlsx   ← 6-sheet model (Dashboard, Orders, EOQ, Scorecards)
│   └── BeerFlow_Logistics_DeepDive.xlsx      ← 8-sheet model (Routes, Fleet, Cold Chain, Rail Arbitrage)
│
├── powerpoint/                        ← Executive presentation decks
│   ├── BeerFlow_Strategy_Deck.pptx          ← 14-slide strategy presentation
│   └── BeerFlow_Pro_Presentation.pptx       ← 14-slide master executive presentation
│
├── insights/                          ← Analytical write-ups
│   ├── key_findings.md               ← Core quantitative findings
│   └── logistics_deep_dive.md        ← Dedicated logistics and freight analysis
│
└── scripts/                           ← Python data generation & workbook builders
    ├── generate_data.py              ← Orders, inventory, and forecast data generator
    ├── generate_logistics_data.py    ← Route, fleet, cold chain, and freight data generator
    ├── build_excel.py                ← Builds BeerFlow_Supply_Chain_Analysis.xlsx
    ├── build_logistics_excel.py      ← Builds BeerFlow_Logistics_DeepDive.xlsx
    ├── build_ppt.py                  ← Builds BeerFlow_Strategy_Deck.pptx
    └── build_pro_ppt.py              ← Builds BeerFlow_Pro_Presentation.pptx
```

---

## 🔬 Methodology & Key Findings

### 1. Demand Forecasting & Seasonality
* **Model:** Evaluated 24 months of demand across 7 SKUs using seasonal decomposition multipliers:
  * Q4 (Oct–Dec): Festival surge of **+30% to +70%** (Diwali, Christmas, New Year).
  * Q1–Q2 (Mar–Apr): Summer onset and IPL season surge of **+20% to +45%**.
  * Q3 (Jul–Aug): Monsoon consumption trough of **-15% to -40%**.
* **Accuracy Metric (MAPE):**
  $$\text{MAPE} = \frac{1}{n} \sum \left| \frac{\text{Actual} - \text{Forecast}}{\text{Actual}} \right| \times 100$$
  * High-volume flagship **Lager 500ml** achieved **9.2% MAPE** (benchmark <10%).
  * Craft/niche **IPA 330ml** showed highest volatility at **22.4% MAPE**, indicating the need for weekly rolling POS data rather than monthly batch forecasting.

### 2. Inventory Optimization (EOQ & Safety Stock)
* **Formulas:**
  * Wilson EOQ: $Q^* = \sqrt{\frac{2DS}{H}}$ where $S = \text{INR 1,200/order}$ and $H = \text{INR 36–52/case/yr}$ (24% carrying cost).
  * Safety Stock: $SS = Z \times \sigma_d \times \sqrt{L}$ ($Z = 1.65$ for 95% service level).
  * Reorder Point: $ROP = (d \times L) + SS$.
* **Findings:**
  * East India warehouses held a chronic **30% deficit** against recommended safety stocks, directly causing regional stockouts.
  * Slow-moving **Stout 650ml** was held at an average of 1,450 cases vs. optimal EOQ of 980 cases (47% over-ordered).
  * Rebalancing order quantities to $Q^*$ reduced total working inventory from **32,450 to 27,580 cases** — a **15.0% reduction in working capital lockup**, freeing **INR 8.76 Lakhs** in cash.

### 3. Last-Mile Route Optimization
* **Analysis of 1,000 delivery routes:**
  * Average vehicle load utilization was **72.1%**, meaning **27.9% of vehicle capacity ran empty**.
  * Root Causes: Geographic zone dispatching rather than cubic load pooling, deploying 32-ft heavy trucks on sub-50km urban routes (costing INR 24.80/case), and 65% empty return backhauls.
  * Solution: Implemented TMS load pooling rules (minimum 85% utilization threshold before dispatch) and route clustering, cutting delivery cost per case by 12–18%.

### 4. Cold Chain Integrity
* **Audit of 600 temperature-tracked shipments:**
  * **18.2%** of shipments recorded thermal breaches (>8°C for >30 minutes).
  * Risk by stage: Brewery loading (2.1%), Primary transit (6.4%), Regional hub (8.9%), Secondary urban delivery (18.2%), and **Retail handover (34.8%)**.
  * Retail handover is the primary bottleneck because small retailers lack walk-in cold vaults and leave product on warm loading bays in summer heat.

### 5. Multi-Modal Freight Arbitrage (Road vs. Rail)
* **Benchmarking 400 shipments:**
  * Road Truck: INR 14.80/case | 3.2 days transit | 88.4% OTD
  * Rail Freight: INR 8.20/case | 5.5 days transit | 82.1% OTD
  * Cost Differential: Rail is **44.6% cheaper per case** (INR 6.60 savings per case).
  * Business Case: Shifting 10% of long-haul volume (>800km corridors like Delhi–Kolkata) to rail saves **INR 21.12 Lakhs annually** and cuts 60 tonnes of CO2, with the 2.3-day transit variance absorbed by scheduled batch dispatches on predictable core SKUs.

### 6. Distributor Performance Scorecard
* **Composite Performance Score (CPS):**
  $$\text{CPS} = (0.35 \times \text{OTIF}) + (0.30 \times \text{Fill Rate}) + (0.20 \times \text{Damage Score}) + (0.15 \times \text{TAT Score})$$
* **Tier Distribution across 20 partners:**
  * Grade A (Score ≥ 85): 6 partners (Eligible for volume rebates & priority allocation)
  * Grade B (Score 70–84): 6 partners (Stable performance)
  * Grade C (Score 55–69): 5 partners (High damage rates >1.2%; enrolled in 60-day DPIP)
  * Grade D (Score < 55): 3 partners (Chronic OTIF failure <70%; territory reassignment review)

---

## 📈 Summary of Quantitative Outcomes

| Metric | Baseline | Optimized Target | Improvement / Impact |
|---|---|---|---|
| **Overall OTIF Rate** | 83.4% | 91.0% | +7.6% (meets 90% benchmark) |
| **Average Load Utilization** | 72.1% | 85.0% | +12.9% (captures empty capacity) |
| **Warehouse Working Inventory** | 32,450 Cases | 27,580 Cases | -15.0% (INR 8.76L capital freed) |
| **Rail Freight Share** | 15.0% | 25.0% | +10.0% (INR 21.12L freight savings) |
| **Cold Chain Breach Rate** | 18.2% | < 5.0% | -13.2% (avoids INR 3.8L spoilage) |
| **Total Annualized Savings** | — | — | **INR 55 – 80 Lakhs / year** |

---

## 🚀 Reproduction & Setup

All datasets, workbooks, and presentations can be regenerated from scratch using Python 3.10+:

```bash
# Clone the repository
git clone https://github.com/Harsh258-collab/beerflow-supply-chain.git
cd beerflow-supply-chain

# Install required dependencies
pip install pandas numpy openpyxl python-pptx

# Step 1: Generate all synthetic transactional datasets
python scripts/generate_data.py
python scripts/generate_logistics_data.py

# Step 2: Build analytical Excel workbooks
python scripts/build_excel.py
python scripts/build_logistics_excel.py

# Step 3: Build executive PowerPoint presentation decks
python scripts/build_ppt.py
python scripts/build_pro_ppt.py
```

---

## 📬 Contact & Portfolio

- **Author:** Harsh Raj Pandey
- **GitHub:** [Harsh258-collab](https://github.com/Harsh258-collab)
- **LinkedIn:** [harsh-raj-pandey-1a0319325](https://www.linkedin.com/in/harsh-raj-pandey-1a0319325)
- **Email:** harsh258.collab@gmail.com
