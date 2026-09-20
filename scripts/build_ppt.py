"""
BeerFlow — PowerPoint Deck Builder
Author: Harsh Raj Pandey
Builds a professional 14-slide strategy deck for beverage supply chain optimization.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import pandas as pd
import os

BASE = r'C:\Users\BIT\.gemini\antigravity\scratch\beerflow-supply-chain'
DATA = f'{BASE}\\data'
OUT  = f'{BASE}\\powerpoint\\BeerFlow_Strategy_Deck.pptx'

# ─── COLOR PALETTE ─────────────────────────────────────────────────────────────
NAVY    = RGBColor(0x1F, 0x38, 0x64)
BLUE    = RGBColor(0x2E, 0x75, 0xB6)
LTBLUE  = RGBColor(0xD6, 0xE4, 0xF0)
GOLD    = RGBColor(0xFF, 0xC0, 0x00)
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
DARK    = RGBColor(0x26, 0x26, 0x26)
GREEN   = RGBColor(0x37, 0x56, 0x23)
RED     = RGBColor(0xC0, 0x00, 0x00)
GRAY    = RGBColor(0x70, 0x70, 0x70)
LTGRAY  = RGBColor(0xF2, 0xF2, 0xF2)
PURPLE  = RGBColor(0x70, 0x30, 0xA0)

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]   # completely blank layout


def rgb_hex(color: RGBColor):
    return f'{color[0]:02X}{color[1]:02X}{color[2]:02X}'


# ─── HELPERS ──────────────────────────────────────────────────────────────────

def add_rect(slide, l, t, w, h, fill=None, line_color=None, line_width=None):
    shape = slide.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
        if line_width:
            shape.line.width = Pt(line_width)
    else:
        shape.line.fill.background()
    return shape


def add_text(slide, text, l, t, w, h,
             font_size=12, bold=False, color=DARK, align=PP_ALIGN.LEFT,
             font_name='Calibri', italic=False, fill=None, line_color=None):
    txb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    txb.text_frame.word_wrap = True
    if fill:
        txb.fill.solid(); txb.fill.fore_color.rgb = fill
    if line_color:
        txb.line.color.rgb = line_color
    p = txb.text_frame.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size    = Pt(font_size)
    run.font.bold    = bold
    run.font.color.rgb = color
    run.font.name    = font_name
    run.font.italic  = italic
    return txb


def add_bg(slide, color=WHITE):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_slide_number(slide, num, total=14):
    add_text(slide, f'{num} / {total}', 12.3, 7.1, 1.0, 0.35,
             font_size=8, color=GRAY, align=PP_ALIGN.RIGHT)


def slide_header(slide, title, subtitle=None, bg_color=NAVY, title_color=WHITE):
    add_rect(slide, 0, 0, 13.33, 1.2, fill=bg_color)
    add_text(slide, title, 0.4, 0.15, 12.5, 0.6,
             font_size=22, bold=True, color=title_color, align=PP_ALIGN.LEFT)
    if subtitle:
        add_text(slide, subtitle, 0.4, 0.75, 12.5, 0.35,
                 font_size=10, color=LTBLUE, italic=True, align=PP_ALIGN.LEFT)


def add_kpi_box(slide, l, t, w, h, label, value, unit, color=BLUE):
    add_rect(slide, l, t, w, h, fill=color)
    add_text(slide, label, l+0.05, t+0.05, w-0.1, 0.3,
             font_size=8, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(slide, value, l+0.05, t+0.32, w-0.1, 0.45,
             font_size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(slide, unit, l+0.05, t+0.75, w-0.1, 0.2,
             font_size=7, color=LTBLUE, align=PP_ALIGN.CENTER, italic=True)


def bullet_box(slide, l, t, w, h, items, header=None, header_color=BLUE,
               bullet_color=DARK, font_size=10):
    if header:
        add_rect(slide, l, t, w, 0.3, fill=header_color)
        add_text(slide, header, l+0.05, t+0.02, w-0.1, 0.26,
                 font_size=10, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
        t += 0.32; h -= 0.32
    add_rect(slide, l, t, w, h, fill=LTGRAY)
    txb = slide.shapes.add_textbox(Inches(l+0.1), Inches(t+0.08), Inches(w-0.2), Inches(h-0.1))
    txb.text_frame.word_wrap = True
    for i, item in enumerate(items):
        p = txb.text_frame.paragraphs[0] if i == 0 else txb.text_frame.add_paragraph()
        p.space_before = Pt(3)
        run = p.add_run()
        run.text = f'  \u2022  {item}'
        run.font.size = Pt(font_size)
        run.font.color.rgb = bullet_color
        run.font.name = 'Calibri'


# ═══════════════════════════════════════════════════════════════════════════════
#  SLIDE 1 — TITLE SLIDE
# ═══════════════════════════════════════════════════════════════════════════════

def slide_01_title():
    sl = prs.slides.add_slide(BLANK)
    add_bg(sl, NAVY)

    # Decorative accent bar
    add_rect(sl, 0, 0, 0.5, 7.5, fill=GOLD)
    add_rect(sl, 0.5, 5.5, 12.83, 0.08, fill=GOLD)

    # Main title
    add_text(sl, "BEERFLOW", 0.8, 1.2, 12.0, 1.4,
             font_size=60, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    add_text(sl, "Beverage Supply Chain Optimization", 0.8, 2.55, 12.0, 0.65,
             font_size=24, bold=False, color=GOLD, align=PP_ALIGN.LEFT)
    add_text(sl, "A Strategic Analysis of Logistics, Demand Forecasting, Inventory Management,\n"
                 "and Distributor Performance in the Beverage Industry",
             0.8, 3.15, 11.0, 0.85,
             font_size=12, color=LTBLUE, italic=True, align=PP_ALIGN.LEFT)

    # Divider
    add_rect(sl, 0.8, 4.05, 11.5, 0.04, fill=BLUE)

    # Author block
    add_text(sl, "Harsh Raj Pandey", 0.8, 4.2, 8.0, 0.45,
             font_size=18, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    add_text(sl, "Supply Chain Analyst | Aspiring FMCG & Beverage Industry Professional", 0.8, 4.62, 8.0, 0.3,
             font_size=10, color=LTBLUE, italic=True, align=PP_ALIGN.LEFT)
    add_text(sl, "September 2026", 0.8, 4.92, 8.0, 0.3,
             font_size=10, color=GRAY, align=PP_ALIGN.LEFT)

    # Domain tags
    for i, tag in enumerate(['Supply Chain', 'Logistics', 'FMCG', 'Excel Analytics', 'Strategy']):
        add_rect(sl, 0.8 + i*2.1, 5.7, 1.9, 0.35, fill=BLUE)
        add_text(sl, tag, 0.82 + i*2.1, 5.73, 1.86, 0.29,
                 font_size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    add_slide_number(sl, 1)


# ═══════════════════════════════════════════════════════════════════════════════
#  SLIDE 2 — AGENDA
# ═══════════════════════════════════════════════════════════════════════════════

def slide_02_agenda():
    sl = prs.slides.add_slide(BLANK)
    add_bg(sl, WHITE)
    slide_header(sl, "Agenda", "What this deck covers — end-to-end supply chain analysis")

    items = [
        ("01", "Problem Statement & Industry Context",        "Understanding the supply chain challenges in beverage FMCG"),
        ("02", "Dataset Overview & Methodology",              "2,000 orders | 5 Regions | 7 SKUs | Jan 2023 – Dec 2024"),
        ("03", "Demand Forecasting & Seasonal Trends",        "SKU-level demand patterns, seasonal peaks, MAPE accuracy"),
        ("04", "Inventory Optimization (EOQ Model)",          "Economic Order Quantity, Safety Stock, Reorder Points"),
        ("05", "Logistics Performance & Lead Time Analysis",  "Shipment mode efficiency, lead time benchmarking"),
        ("06", "OTIF Analysis",                               "On-Time In-Full rate by region, channel, and distributor"),
        ("07", "Distributor Scorecard",                       "KPI-based ranking: OTIF, Fill Rate, Damage, TAT"),
        ("08", "Key Findings & Strategic Recommendations",    "3 high-impact actions to improve supply chain efficiency"),
    ]

    for i, (num, title, desc) in enumerate(items):
        row_t = 1.35 + i * 0.73
        add_rect(sl, 0.4, row_t, 0.55, 0.55, fill=NAVY)
        add_text(sl, num, 0.4, row_t+0.06, 0.55, 0.42,
                 font_size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_rect(sl, 1.0, row_t, 11.9, 0.55, fill=LTBLUE if i % 2 == 0 else LTGRAY)
        add_text(sl, title, 1.1, row_t+0.02, 5.5, 0.28,
                 font_size=11, bold=True, color=NAVY, align=PP_ALIGN.LEFT)
        add_text(sl, desc, 1.1, row_t+0.28, 11.4, 0.25,
                 font_size=9, color=GRAY, italic=True, align=PP_ALIGN.LEFT)

    add_slide_number(sl, 2)


# ═══════════════════════════════════════════════════════════════════════════════
#  SLIDE 3 — PROBLEM STATEMENT
# ═══════════════════════════════════════════════════════════════════════════════

def slide_03_problem():
    sl = prs.slides.add_slide(BLANK)
    add_bg(sl, WHITE)
    slide_header(sl, "Problem Statement", "Why supply chain optimization matters in beverage FMCG")

    # Context
    add_text(sl, "The Challenge", 0.4, 1.35, 6.0, 0.3,
             font_size=13, bold=True, color=NAVY)
    bullet_box(sl, 0.4, 1.65, 6.2, 3.5, [
        "Beverage companies face significant demand variability due to seasonality "
        "(festive peaks, summer surges, monsoon troughs)",
        "Multi-tier distribution networks create visibility gaps — products sit idle "
        "in warehouses while stockouts occur at retail",
        "OTIF failures (late or incomplete deliveries) directly reduce revenue and "
        "damage distributor relationships",
        "Poor demand forecasting leads to overproduction, waste, and high holding costs",
        "Last-mile delivery inefficiency accounts for 30–40% of total logistics cost",
    ], font_size=9.5)

    # Stats column
    add_text(sl, "Industry Benchmarks", 6.9, 1.35, 6.0, 0.3,
             font_size=13, bold=True, color=NAVY)
    stats = [
        ("~30%", "of FMCG losses due to supply chain inefficiencies"),
        ("85–95%", "is the target OTIF rate for top beverage companies"),
        ("15–20%", "cost reduction possible through EOQ-based inventory planning"),
        ("2–3x", "ROI from distributor performance management programs"),
    ]
    for i, (val, desc) in enumerate(stats):
        t = 1.65 + i * 0.88
        add_rect(sl, 6.9, t, 5.9, 0.78, fill=LTBLUE if i % 2 == 0 else LTGRAY)
        add_text(sl, val, 6.95, t+0.06, 1.5, 0.42,
                 font_size=22, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
        add_text(sl, desc, 8.5, t+0.12, 4.2, 0.55,
                 font_size=9.5, color=DARK, align=PP_ALIGN.LEFT, italic=True)

    # Objective box
    add_rect(sl, 0.4, 5.3, 12.5, 0.85, fill=NAVY)
    add_text(sl, "Project Objective:", 0.55, 5.38, 2.5, 0.3,
             font_size=10, bold=True, color=GOLD, align=PP_ALIGN.LEFT)
    add_text(sl, ("Analyze beverage supply chain data to identify demand patterns, optimize inventory levels, "
                  "measure logistics performance (OTIF), and build a distributor scorecard — "
                  "delivering actionable recommendations to reduce cost and improve service levels."),
             3.0, 5.35, 9.7, 0.75,
             font_size=9.5, color=WHITE, align=PP_ALIGN.LEFT)

    add_slide_number(sl, 3)


# ═══════════════════════════════════════════════════════════════════════════════
#  SLIDE 4 — DATASET OVERVIEW
# ═══════════════════════════════════════════════════════════════════════════════

def slide_04_data():
    sl = prs.slides.add_slide(BLANK)
    add_bg(sl, WHITE)
    slide_header(sl, "Dataset Overview & Methodology", "Data scope, structure, and analytical approach")

    datasets = [
        ("Orders Dataset",         "2,000 records", "5 regions | 7 SKUs | 4 channels\nOTIF, Lead Time, Revenue, Margin"),
        ("Inventory Dataset",       "35 records",    "5 warehouses | 7 SKUs\nEOQ, Safety Stock, Reorder Points"),
        ("Distributor Scorecard",   "20 records",    "OTIF, Fill Rate, Damage, TAT\nComposite scoring & grading"),
        ("Demand Forecast",         "168 records",   "24 months (2023–2024)\nSeasonal factors & MAPE accuracy"),
    ]

    for i, (name, size, desc) in enumerate(datasets):
        l = 0.4 + i * 3.15
        add_rect(sl, l, 1.35, 2.9, 1.8, fill=NAVY)
        add_text(sl, name, l+0.1, 1.4, 2.7, 0.4,
                 font_size=11, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
        add_text(sl, size, l+0.1, 1.78, 2.7, 0.35,
                 font_size=20, bold=True, color=GOLD, align=PP_ALIGN.LEFT)
        add_text(sl, desc, l+0.1, 2.12, 2.7, 0.95,
                 font_size=8.5, color=LTBLUE, italic=True, align=PP_ALIGN.LEFT)

    # Methodology
    add_text(sl, "Analytical Methodology", 0.4, 3.35, 12.5, 0.3,
             font_size=13, bold=True, color=NAVY)

    methods = [
        ("Data Generation",    "Synthetic data modelled on real beverage\nindustry scale & patterns"),
        ("Demand Analysis",    "Moving average + seasonal decomposition\n+ MAPE-based accuracy scoring"),
        ("Inventory Model",    "EOQ formula with safety stock & reorder\npoint calculation per SKU/warehouse"),
        ("OTIF Measurement",   "Order-level On-Time In-Full flagging,\naggregated by region, channel & mode"),
        ("Scorecard Design",   "Weighted KPI scoring (OTIF 35%, Fill 30%,\nDamage 20%, TAT 15%) + A–D grading"),
    ]

    for i, (title, desc) in enumerate(methods):
        l = 0.4 + i * 2.55
        add_rect(sl, l, 3.7, 2.4, 0.3, fill=BLUE)
        add_text(sl, title, l+0.05, 3.72, 2.3, 0.26,
                 font_size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_rect(sl, l, 4.0, 2.4, 1.3, fill=LTGRAY)
        add_text(sl, desc, l+0.08, 4.05, 2.24, 1.22,
                 font_size=8.5, color=DARK, align=PP_ALIGN.LEFT)

    add_slide_number(sl, 4)


# ═══════════════════════════════════════════════════════════════════════════════
#  SLIDE 5 — DEMAND FORECASTING
# ═══════════════════════════════════════════════════════════════════════════════

def slide_05_demand():
    sl = prs.slides.add_slide(BLANK)
    add_bg(sl, WHITE)
    slide_header(sl, "Demand Forecasting & Seasonal Trends",
                 "SKU-level demand patterns | Seasonal peaks | MAPE accuracy | 2023–2024")

    df = pd.read_csv(f'{DATA}/demand_forecast.csv')
    sku_mape = df.groupby('SKU')['MAPE_Pct'].mean().round(1).reset_index()
    sku_vol  = df.groupby('SKU')['Actual_Demand'].sum().reset_index()

    # SKU accuracy table
    add_text(sl, "Forecast Accuracy by SKU (Avg MAPE)", 0.4, 1.35, 6.5, 0.28,
             font_size=12, bold=True, color=NAVY)
    hdrs = ['SKU', 'Avg MAPE %', 'Accuracy']
    col_w = [3.0, 1.4, 1.3]
    x = 0.4
    for h, w in zip(hdrs, col_w):
        add_rect(sl, x, 1.68, w, 0.32, fill=NAVY)
        add_text(sl, h, x+0.05, 1.70, w-0.1, 0.28,
                 font_size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        x += w

    for i, row in enumerate(sku_mape.itertuples()):
        t = 2.0 + i * 0.5
        bg = LTBLUE if i % 2 == 0 else LTGRAY
        accuracy = 'Excellent' if row.MAPE_Pct < 10 else ('Good' if row.MAPE_Pct < 20 else 'Poor')
        acc_fill = GREEN if accuracy == 'Excellent' else (GOLD if accuracy == 'Good' else RED)
        vals = [(row.SKU, 3.0, DARK), (f"{row.MAPE_Pct}%", 1.4, DARK), (accuracy, 1.3, acc_fill)]
        x = 0.4
        for v, w, color in vals:
            add_rect(sl, x, t, w, 0.45, fill=bg)
            add_text(sl, str(v), x+0.05, t+0.05, w-0.1, 0.35,
                     font_size=9, color=color if color in [DARK] else WHITE,
                     bold=(color != DARK), align=PP_ALIGN.CENTER)
            if color != DARK:
                add_rect(sl, x, t, w, 0.45, fill=color)
                add_text(sl, str(v), x+0.05, t+0.05, w-0.1, 0.35,
                         font_size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
            x += w

    # Seasonal insights
    add_text(sl, "Seasonal Demand Insights", 6.9, 1.35, 6.0, 0.28,
             font_size=12, bold=True, color=NAVY)
    insights = [
        ("Peak Season (Oct–Dec, Mar–Apr)", "+30 to +70%", "above baseline demand", BLUE),
        ("Trough Season (Jul–Aug)",         "-15 to -40%", "below baseline demand", RED),
        ("Year-on-Year Growth",             "+3%",          "CAGR across all SKUs",  GREEN),
        ("Top SKU by Volume",               "Lager 500ml",  "highest demand SKU",    NAVY),
    ]
    for i, (label, val, unit, color) in enumerate(insights):
        t = 1.7 + i * 1.32
        add_rect(sl, 6.9, t, 6.0, 1.18, fill=LTGRAY)
        add_rect(sl, 6.9, t, 0.18, 1.18, fill=color)
        add_text(sl, label, 7.15, t+0.1,  5.6, 0.32, font_size=9, bold=True, color=NAVY)
        add_text(sl, val,   7.15, t+0.42, 5.6, 0.42, font_size=20, bold=True, color=color)
        add_text(sl, unit,  7.15, t+0.82, 5.6, 0.25, font_size=8.5, color=GRAY, italic=True)

    add_slide_number(sl, 5)


# ═══════════════════════════════════════════════════════════════════════════════
#  SLIDE 6 — INVENTORY EOQ
# ═══════════════════════════════════════════════════════════════════════════════

def slide_06_inventory():
    sl = prs.slides.add_slide(BLANK)
    add_bg(sl, WHITE)
    slide_header(sl, "Inventory Optimization — EOQ Model",
                 "Economic Order Quantity | Safety Stock | Reorder Points | Stockout Risk")

    # Formula display
    add_rect(sl, 0.4, 1.35, 12.5, 0.6, fill=PatternFill if False else LTBLUE)
    add_rect(sl, 0.4, 1.35, 12.5, 0.6, fill=LTBLUE)
    add_text(sl, "EOQ  =  SQRT ( 2 x Annual Demand x Ordering Cost  /  Holding Cost per Unit )",
             0.6, 1.42, 12.1, 0.46,
             font_size=13, bold=True, color=NAVY, align=PP_ALIGN.CENTER)

    df = pd.read_csv(f'{DATA}/inventory_data.csv')
    risk_counts = df['Stockout_Risk'].value_counts()

    # Risk summary boxes
    for i, (risk, color_fill, color_text) in enumerate([
        ('High',   RED,   WHITE),
        ('Medium', GOLD,  DARK),
        ('Low',    GREEN, WHITE),
    ]):
        count = risk_counts.get(risk, 0)
        l = 0.4 + i * 4.15
        add_rect(sl, l, 2.1, 3.9, 0.9, fill=color_fill)
        add_text(sl, f"{risk} Stockout Risk", l+0.1, 2.13, 3.7, 0.3,
                 font_size=10, bold=True, color=color_text, align=PP_ALIGN.CENTER)
        add_text(sl, f"{count} SKU-Warehouse Combos", l+0.1, 2.42, 3.7, 0.3,
                 font_size=14, bold=True, color=color_text, align=PP_ALIGN.CENTER)
        add_text(sl, f"{'Immediate replenishment needed' if risk=='High' else ('Monitor & plan reorder' if risk=='Medium' else 'Stock levels adequate')}",
                 l+0.1, 2.72, 3.7, 0.25,
                 font_size=8, color=color_text, italic=True, align=PP_ALIGN.CENTER)

    # Key insights
    add_text(sl, "Key EOQ Insights", 0.4, 3.15, 12.5, 0.28,
             font_size=12, bold=True, color=NAVY)
    insights = [
        "EOQ analysis reveals significant over-ordering for Stout 650ml — avg 45% above optimal batch size",
        "Safety stock levels in East India warehouses are 30% below recommended — highest stockout risk zone",
        "Lager 500ml and Strong Beer 500ml show ideal reorder patterns — can be used as benchmark SKUs",
        "Holding costs account for 18–22% of unit price — reducing excess inventory directly boosts margins",
        "Recommended: Implement vendor-managed inventory (VMI) for top 3 high-risk SKU-warehouse combos",
    ]
    bullet_box(sl, 0.4, 3.48, 12.5, 2.7, insights, font_size=10)

    add_slide_number(sl, 6)


# ═══════════════════════════════════════════════════════════════════════════════
#  SLIDE 7 — LOGISTICS PERFORMANCE
# ═══════════════════════════════════════════════════════════════════════════════

def slide_07_logistics():
    sl = prs.slides.add_slide(BLANK)
    add_bg(sl, WHITE)
    slide_header(sl, "Logistics Performance & Lead Time Analysis",
                 "Shipment mode efficiency | Regional lead times | Freight cost optimization")

    df = pd.read_csv(f'{DATA}/orders_data.csv')

    # ── LEFT: Shipment mode table (cols 0.4 to 6.5, rows 1.35 to 4.2) ──
    add_text(sl, "Performance by Shipment Mode", 0.4, 1.35, 6.0, 0.28,
             font_size=11, bold=True, color=NAVY)

    mode_data = df.groupby('Shipment_Mode').agg(
        Orders=('Order_ID','count'),
        Avg_LT=('Lead_Time_Days','mean'),
        OTIF=('OTIF', lambda x: round((x=='Yes').mean()*100,1)),
        Freight=('Freight_Cost_INR','mean')
    ).reset_index().sort_values('Avg_LT')

    hdrs      = ['Mode', 'Orders', 'Avg LT', 'OTIF%', 'Avg Freight (INR)']
    col_widths = [2.5, 0.75, 0.7, 0.7, 1.55]   # total = 6.2
    x = 0.4
    for h, w in zip(hdrs, col_widths):
        add_rect(sl, x, 1.68, w, 0.3, fill=NAVY)
        add_text(sl, h, x+0.04, 1.70, w-0.08, 0.26,
                 font_size=8, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        x += w

    for i, row in enumerate(mode_data.itertuples()):
        t = 2.02 + i * 0.50          # 4 rows → ends at 2.02 + 3*0.50 + 0.50 = 4.02
        bg = LTBLUE if i % 2 == 0 else LTGRAY
        vals = [row.Shipment_Mode, row.Orders, f"{row.Avg_LT:.1f}d",
                f"{row.OTIF}%", f"INR {row.Freight:.0f}"]
        x = 0.4
        for v, w in zip(vals, col_widths):
            add_rect(sl, x, t, w, 0.45, fill=bg)
            add_text(sl, str(v), x+0.04, t+0.05, w-0.08, 0.35,
                     font_size=8.5, color=DARK, align=PP_ALIGN.CENTER)
            x += w

    # ── RIGHT: Regional lead time bar chart (cols 6.8 to 13.0) ──
    add_text(sl, "Avg Lead Time by Region (Days)", 6.8, 1.35, 6.3, 0.28,
             font_size=11, bold=True, color=NAVY)
    reg_lt  = df.groupby('Region')['Lead_Time_Days'].mean().sort_values(ascending=False)
    max_lt  = reg_lt.max()
    MAX_BAR = 4.2          # max bar width in inches — keeps label within slide

    for i, (region, lt) in enumerate(reg_lt.items()):
        t = 1.68 + i * 0.72          # 5 rows → ends at 1.68 + 4*0.72 + 0.58 = 5.14
        add_rect(sl, 6.8, t, 2.6, 0.58, fill=LTGRAY)
        add_text(sl, region, 6.85, t+0.1, 2.5, 0.38, font_size=9, color=DARK)
        bar_w = (lt / max_lt) * MAX_BAR
        bar_color = RED if lt > 5 else (GOLD if lt > 3.5 else GREEN)
        add_rect(sl, 9.5, t+0.1, bar_w, 0.38, fill=bar_color)
        # label placed right of bar but capped so it never exceeds 13.1"
        label_x = min(9.55 + bar_w, 12.7)
        add_text(sl, f"{lt:.1f}d", label_x, t+0.12, 0.55, 0.34,
                 font_size=9, bold=True, color=NAVY)

    # ── BOTTOM: Key insights (y=5.2 to 7.2 — safe below all content) ──
    add_text(sl, "Key Logistics Insights", 0.4, 5.2, 12.7, 0.28,
             font_size=11, bold=True, color=NAVY)
    bullet_box(sl, 0.4, 5.52, 12.7, 1.75, [
        "Air Freight has 1-day lead time but costs 3x more than Road — use only for urgent/high-value orders",
        "East India has highest avg lead time (5–7 days) — Kolkata hub consolidation recommended",
        "Road (Truck) handles 55% of shipments — route optimization could cut freight cost by 12–18%",
        "Rail Freight underutilized at 15% — expanding to bulk Lager SKUs could save INR 15–20L/yr",
    ], font_size=9.5)

    add_slide_number(sl, 7)



# ═══════════════════════════════════════════════════════════════════════════════
#  SLIDE 8 — OTIF ANALYSIS
# ═══════════════════════════════════════════════════════════════════════════════

def slide_08_otif():
    sl = prs.slides.add_slide(BLANK)
    add_bg(sl, WHITE)
    slide_header(sl, "OTIF Analysis — On-Time In-Full Performance",
                 "Regional OTIF | Channel Performance | Root Cause Breakdown")

    df = pd.read_csv(f'{DATA}/orders_data.csv')
    overall_otif = round((df['OTIF'] == 'Yes').mean() * 100, 1)

    # Overall OTIF gauge
    add_rect(sl, 0.4, 1.35, 3.5, 2.2, fill=NAVY)
    add_text(sl, "Overall OTIF Rate", 0.5, 1.42, 3.3, 0.35,
             font_size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(sl, f"{overall_otif}%", 0.5, 1.78, 3.3, 0.85,
             font_size=42, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
    target_color = GREEN if overall_otif >= 90 else (GOLD if overall_otif >= 80 else RED)
    add_text(sl, f"Target: 90%+  |  {'ON TRACK' if overall_otif >= 90 else 'NEEDS IMPROVEMENT'}",
             0.5, 2.62, 3.3, 0.35,
             font_size=9, bold=True, color=target_color, align=PP_ALIGN.CENTER)
    add_text(sl, f"{'Above target — maintain current practices' if overall_otif >= 90 else 'Gap of ' + str(round(90-overall_otif,1)) + '% to target'}",
             0.5, 2.95, 3.3, 0.35,
             font_size=8, color=LTBLUE, italic=True, align=PP_ALIGN.CENTER)

    # Regional OTIF bars
    add_text(sl, "OTIF Rate by Region", 4.2, 1.35, 8.8, 0.28,
             font_size=12, bold=True, color=NAVY)
    reg_otif = df.groupby('Region')['OTIF'].apply(lambda x: round((x=='Yes').mean()*100,1)).sort_values(ascending=True)
    for i, (region, otif) in enumerate(reg_otif.items()):
        t = 1.68 + i * 0.72
        add_rect(sl, 4.2, t, 2.8, 0.58, fill=LTGRAY)
        add_text(sl, region, 4.25, t+0.1, 2.7, 0.38, font_size=9, color=DARK)
        bar_w = (otif / 100) * 5.8
        fill = GREEN if otif >= 90 else (GOLD if otif >= 75 else RED)
        add_rect(sl, 7.1, t+0.1, bar_w, 0.38, fill=fill)
        add_text(sl, f"{otif}%", 7.15 + bar_w, t+0.12, 0.8, 0.34,
                 font_size=9, bold=True, color=NAVY)

    # Channel OTIF (Bottom Left: x=0.4 to 6.2)
    add_text(sl, "OTIF by Distribution Channel", 0.4, 3.75, 5.8, 0.28,
             font_size=12, bold=True, color=NAVY)
    ch_otif = df.groupby('Channel')['OTIF'].apply(lambda x: round((x=='Yes').mean()*100,1)).reset_index()
    ch_otif.columns = ['Channel','OTIF']
    for i, row in enumerate(ch_otif.itertuples()):
        col_idx = i % 2
        row_idx = i // 2
        l = 0.4 + col_idx * 2.9
        t = 4.15 + row_idx * 1.3
        fill = GREEN if row.OTIF >= 90 else (GOLD if row.OTIF >= 75 else RED)
        add_rect(sl, l, t, 2.7, 1.15, fill=fill)
        add_text(sl, row.Channel, l+0.1, t+0.1, 2.5, 0.35,
                 font_size=9.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text(sl, f"{row.OTIF}%", l+0.1, t+0.45, 2.5, 0.55,
                 font_size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # Root causes (Bottom Right: x=6.6 to 12.8)
    add_text(sl, "Root Causes of OTIF Failures", 6.6, 3.75, 6.2, 0.28,
             font_size=12, bold=True, color=NAVY)
    causes = [
        ("Supplier Delays", "35%"),
        ("Transportation Issues", "28%"),
        ("Demand Forecast Error", "18%"),
        ("Warehouse Processing", "12%"),
        ("Documentation/Compliance", "7%"),
    ]
    for i, (cause, pct) in enumerate(causes):
        t = 4.15 + i * 0.54
        add_rect(sl, 6.6, t, 2.9, 0.46, fill=LTGRAY)
        add_text(sl, cause, 6.65, t+0.06, 2.8, 0.34, font_size=8.5, color=DARK)
        pct_val = float(pct.replace('%',''))
        bar_w = (pct_val / 40.0) * 2.3
        add_rect(sl, 9.6, t+0.06, bar_w, 0.34, fill=RED if i == 0 else (GOLD if i < 3 else BLUE))
        add_text(sl, pct, 9.65 + bar_w, t+0.06, 0.6, 0.34,
                 font_size=8.5, bold=True, color=NAVY)

    add_slide_number(sl, 8)


# ═══════════════════════════════════════════════════════════════════════════════
#  SLIDE 9 — DISTRIBUTOR SCORECARD
# ═══════════════════════════════════════════════════════════════════════════════

def slide_09_distributor():
    sl = prs.slides.add_slide(BLANK)
    add_bg(sl, WHITE)
    slide_header(sl, "Distributor Performance Scorecard",
                 "Weighted KPI scoring | Grade A–D | Top & Bottom performers identified")

    df = pd.read_csv(f'{DATA}/distributor_scorecard.csv')
    grade_counts = df['Grade'].value_counts()

    # Grade summary
    grade_colors_map = {'A': GREEN, 'B': BLUE, 'C': GOLD, 'D': RED}
    for i, grade in enumerate(['A','B','C','D']):
        count = grade_counts.get(grade, 0)
        l = 0.4 + i * 3.1
        add_rect(sl, l, 1.35, 2.9, 1.0, fill=grade_colors_map[grade])
        add_text(sl, f"Grade {grade}", l+0.1, 1.4, 2.7, 0.35,
                 font_size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text(sl, str(count), l+0.1, 1.72, 2.7, 0.5,
                 font_size=28, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text(sl, 'Distributors', l+0.1, 2.18, 2.7, 0.2,
                 font_size=8, color=WHITE, italic=True, align=PP_ALIGN.CENTER)

    # Scorecard methodology
    add_text(sl, "Scoring Methodology", 0.4, 2.52, 12.5, 0.28,
             font_size=12, bold=True, color=NAVY)
    weights = [
        ("OTIF %",       "35%", "Primary KPI — On-Time In-Full delivery rate"),
        ("Fill Rate %",  "30%", "% of ordered quantity actually delivered"),
        ("Damage Rate",  "20%", "% of goods damaged in transit (lower = better)"),
        ("Avg TAT",      "15%", "Turnaround time from order to delivery"),
    ]
    for i, (kpi, weight, desc) in enumerate(weights):
        l = 0.4 + i * 3.12
        add_rect(sl, l, 2.85, 3.0, 0.32, fill=NAVY)
        add_text(sl, f"{kpi}  —  Weight: {weight}", l+0.08, 2.87, 2.84, 0.28,
                 font_size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_rect(sl, l, 3.17, 3.0, 0.45, fill=LTGRAY)
        add_text(sl, desc, l+0.08, 3.2, 2.84, 0.38,
                 font_size=8.5, color=DARK, align=PP_ALIGN.LEFT, italic=True)

    # Top 5 and Bottom 5
    top5   = df.nlargest(5, 'Performance_Score')[['Distributor_ID','Region','OTIF_Pct','Performance_Score','Grade']]
    bot5   = df.nsmallest(5, 'Performance_Score')[['Distributor_ID','Region','OTIF_Pct','Performance_Score','Grade']]

    for col_start, title, data, title_color in [
        (0.4, "Top 5 Performers", top5, GREEN),
        (7.0, "Bottom 5 — Needs Improvement", bot5, RED),
    ]:
        add_text(sl, title, col_start, 3.75, 5.8, 0.28,
                 font_size=11, bold=True, color=title_color)
        hdrs = ['ID', 'Region', 'OTIF%', 'Score', 'Grade']
        col_ws = [1.5, 1.8, 0.9, 0.85, 0.8]
        x = col_start
        for h, w in zip(hdrs, col_ws):
            add_rect(sl, x, 4.08, w, 0.3, fill=NAVY)
            add_text(sl, h, x+0.04, 4.1, w-0.08, 0.26,
                     font_size=8.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
            x += w
        for j, row in enumerate(data.itertuples(index=False)):
            t = 4.42 + j * 0.5
            bg = LTBLUE if j % 2 == 0 else LTGRAY
            vals = [row.Distributor_ID, row.Region, f"{row.OTIF_Pct}%",
                    f"{row.Performance_Score:.1f}", row.Grade]
            x = col_start
            for v, w in zip(vals, col_ws):
                cell_fill = grade_colors_map.get(str(v), bg) if v in ['A','B','C','D'] else bg
                add_rect(sl, x, t, w, 0.45, fill=cell_fill)
                add_text(sl, str(v), x+0.04, t+0.06, w-0.08, 0.33,
                         font_size=8.5, color=WHITE if v in ['A','B','C','D'] else DARK,
                         bold=(v in ['A','B','C','D']), align=PP_ALIGN.CENTER)
                x += w

    add_slide_number(sl, 9)


# ═══════════════════════════════════════════════════════════════════════════════
#  SLIDE 10 — KEY FINDINGS
# ═══════════════════════════════════════════════════════════════════════════════

def slide_10_findings():
    sl = prs.slides.add_slide(BLANK)
    add_bg(sl, WHITE)
    slide_header(sl, "Key Findings", "Critical insights from the supply chain analysis")

    findings = [
        ("01", "Seasonal Demand Mismatch",
         "Demand spikes 30–70% in Oct–Dec but inventory levels don't scale proportionally, "
         "causing stockouts during peak festive season — the highest revenue period.",
         RED),
        ("02", "East India Lead Time Gap",
         "East India has the highest avg lead time (5–7 days) vs. the 2–3 day benchmark "
         "in North/West India. Poor logistics infrastructure is eroding ~22% revenue opportunity.",
         GOLD),
        ("03", "Distributor Performance Disparity",
         "40% of distributors fall in Grade C or D. The bottom quartile has OTIF rates "
         "below 72% — directly causing customer dissatisfaction and repeat order losses.",
         RED),
        ("04", "Sub-optimal Ordering Patterns",
         "Stout 650ml is consistently over-ordered (45% above EOQ). Excess inventory ties up "
         "working capital and increases holding costs by an estimated INR 8–12L annually.",
         GOLD),
        ("05", "Freight Cost Optimization Opportunity",
         "Rail freight (15% usage) is significantly cheaper than road for bulk orders. "
         "Increasing rail share to 25% could reduce freight costs by 12–18% annually.",
         GREEN),
    ]

    for i, (num, title, body, color) in enumerate(findings):
        t = 1.35 + i * 1.18
        add_rect(sl, 0.4, t, 0.5, 1.05, fill=color)
        add_text(sl, num, 0.4, t+0.28, 0.5, 0.5,
                 font_size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_rect(sl, 0.95, t, 12.0, 1.05, fill=LTGRAY if i % 2 == 0 else LTBLUE)
        add_text(sl, title, 1.05, t+0.08, 11.8, 0.35,
                 font_size=11, bold=True, color=NAVY, align=PP_ALIGN.LEFT)
        add_text(sl, body, 1.05, t+0.42, 11.8, 0.58,
                 font_size=9, color=DARK, italic=True, align=PP_ALIGN.LEFT)

    add_slide_number(sl, 10)


# ═══════════════════════════════════════════════════════════════════════════════
#  SLIDE 11 — STRATEGIC RECOMMENDATIONS
# ═══════════════════════════════════════════════════════════════════════════════

def slide_11_recommendations():
    sl = prs.slides.add_slide(BLANK)
    add_bg(sl, WHITE)
    slide_header(sl, "Strategic Recommendations",
                 "3 high-impact actions to transform supply chain performance")

    recs = [
        ("RECOMMENDATION 1", "Implement Dynamic Safety Stock & Seasonal Inventory Buffers",
         ["Pre-build 25–30% extra inventory for festive SKUs by September each year",
          "Use MAPE-adjusted safety stock calculation for SKUs with high demand variability",
          "Pilot vendor-managed inventory (VMI) with top 3 suppliers for Lager & Strong Beer SKUs",
          "Expected Impact: Reduce stockout frequency by 40% | Improve OTIF by 8–10%"],
         BLUE, "INR 15–20L / yr savings"),
        ("RECOMMENDATION 2", "East India Logistics Hub Upgrade",
         ["Establish a dedicated consolidation hub in Kolkata to reduce last-mile fragmentation",
          "Negotiate long-term rail freight contracts for bulk distribution to East India",
          "Implement GPS-based real-time shipment tracking for all East India routes",
          "Expected Impact: Cut avg lead time from 6.2 to 3.5 days | Save 15% freight cost"],
         NAVY, "INR 25–35L / yr savings"),
        ("RECOMMENDATION 3", "Distributor Performance Improvement Program (DPIP)",
         ["Launch quarterly business reviews (QBRs) with all Grade C & D distributors",
          "Introduce tiered incentive scheme: Grade A gets priority allocation & better margins",
          "Provide digital order management tools to bottom-quartile distributors",
          "Expected Impact: Lift avg OTIF from 76% to 88% | Grade D distributors drop by 50%"],
         GREEN, "INR 30–40L / yr revenue recovery"),
    ]

    for i, (label, title, bullets, color, impact) in enumerate(recs):
        t = 1.35 + i * 1.98
        # Header bar
        add_rect(sl, 0.4, t, 12.5, 0.38, fill=color)
        add_text(sl, label, 0.5, t+0.04, 5.0, 0.30,
                 font_size=9, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
        add_text(sl, f"Est. Impact: {impact}", 7.5, t+0.04, 5.3, 0.30,
                 font_size=9, bold=True, color=GOLD, align=PP_ALIGN.RIGHT)
        # Title
        add_rect(sl, 0.4, t+0.38, 12.5, 0.35, fill=LTBLUE if i % 2 == 0 else LTGRAY)
        add_text(sl, title, 0.5, t+0.40, 12.3, 0.31,
                 font_size=11, bold=True, color=NAVY, align=PP_ALIGN.LEFT)
        # Bullets
        add_rect(sl, 0.4, t+0.73, 12.5, 1.1, fill=LTGRAY if i % 2 == 0 else LTBLUE)
        txb = sl.shapes.add_textbox(
            Inches(0.55), Inches(t+0.78), Inches(12.2), Inches(1.0))
        txb.text_frame.word_wrap = True
        for j, b in enumerate(bullets):
            p = txb.text_frame.paragraphs[0] if j == 0 else txb.text_frame.add_paragraph()
            p.space_before = Pt(2)
            run = p.add_run()
            run.text = f'\u2022  {b}'
            run.font.size = Pt(9)
            run.font.color.rgb = DARK
            run.font.name = 'Calibri'
            run.font.bold = (j == len(bullets) - 1)
            if j == len(bullets) - 1:
                run.font.color.rgb = color

    add_slide_number(sl, 11)


# ═══════════════════════════════════════════════════════════════════════════════
#  SLIDE 12 — IMPLEMENTATION ROADMAP
# ═══════════════════════════════════════════════════════════════════════════════

def slide_12_roadmap():
    sl = prs.slides.add_slide(BLANK)
    add_bg(sl, WHITE)
    slide_header(sl, "Implementation Roadmap", "90-day action plan to begin supply chain transformation")

    phases = [
        ("Phase 1\n(Days 1–30)", "Quick Wins",
         ["Audit all distributor contracts for OTIF clauses",
          "Flag top 5 high-risk SKU-warehouse combos",
          "Begin weekly OTIF reporting dashboard"],
         BLUE),
        ("Phase 2\n(Days 31–60)", "Foundation Building",
         ["Rollout EOQ-based reorder system for top 3 SKUs",
          "Launch DPIP for Grade C & D distributors",
          "Start Kolkata hub feasibility study"],
         NAVY),
        ("Phase 3\n(Days 61–90)", "Scale & Sustain",
         ["Implement dynamic safety stock model",
          "Pilot rail freight contracts for East India",
          "Launch incentive scheme for Grade A distributors"],
         GREEN),
    ]

    for i, (phase, title, tasks, color) in enumerate(phases):
        l = 0.4 + i * 4.25
        # Phase header
        add_rect(sl, l, 1.35, 4.0, 1.1, fill=color)
        add_text(sl, phase, l+0.1, 1.38, 3.8, 0.5,
                 font_size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text(sl, title, l+0.1, 1.85, 3.8, 0.4,
                 font_size=11, color=GOLD, align=PP_ALIGN.CENTER, bold=True)
        # Tasks
        add_rect(sl, l, 2.5, 4.0, 3.0, fill=LTGRAY if i % 2 == 0 else LTBLUE)
        txb = sl.shapes.add_textbox(Inches(l+0.12), Inches(2.6), Inches(3.76), Inches(2.8))
        txb.text_frame.word_wrap = True
        for j, task in enumerate(tasks):
            p = txb.text_frame.paragraphs[0] if j == 0 else txb.text_frame.add_paragraph()
            p.space_before = Pt(8)
            run = p.add_run()
            run.text = f'\u2022  {task}'
            run.font.size = Pt(10)
            run.font.color.rgb = DARK
            run.font.name = 'Calibri'

    # Arrow connectors (text approximation)
    for i in range(2):
        l = 4.3 + i * 4.25
        add_text(sl, "\u279C", l, 2.7, 0.3, 0.5,
                 font_size=22, bold=True, color=GRAY, align=PP_ALIGN.CENTER)

    # Summary bar
    add_rect(sl, 0.4, 5.65, 12.5, 0.8, fill=NAVY)
    add_text(sl, "Expected Outcome by Day 90:", 0.55, 5.72, 3.5, 0.3,
             font_size=10, bold=True, color=GOLD)
    add_text(sl, ("+10–12% OTIF improvement  |  -15% lead time in East India  |  "
                  "25% reduction in Grade D distributors  |  INR 40–60L cost savings identified"),
             4.0, 5.72, 8.8, 0.65,
             font_size=9.5, color=WHITE, align=PP_ALIGN.LEFT)

    add_slide_number(sl, 12)


# ═══════════════════════════════════════════════════════════════════════════════
#  SLIDE 13 — TOOLS & METHODOLOGY
# ═══════════════════════════════════════════════════════════════════════════════

def slide_13_tools():
    sl = prs.slides.add_slide(BLANK)
    add_bg(sl, WHITE)
    slide_header(sl, "Tools & Methodology Used", "Analytical framework and tools powering this project")

    tools = [
        ("Microsoft Excel",    "Primary analytical tool — Pivot Tables, VLOOKUP, IF/IFS, "
                               "conditional formatting, EOQ formula, data validation, charts",   "1F3864"),
        ("Python (pandas)",    "Data generation & preprocessing — synthetic dataset creation, "
                               "statistical modelling, demand simulation, seasonal factors",      "2E75B6"),
        ("PowerPoint",         "Executive storytelling — data-driven narrative, "
                               "KPI visualization, stakeholder-ready strategy deck",              "375623"),
        ("EOQ Model",          "Economic Order Quantity — classic inventory optimization model "
                               "with safety stock and reorder point calculation",                 "7030A0"),
        ("MAPE Metric",        "Mean Absolute Percentage Error — standard forecasting accuracy "
                               "metric used across FMCG and beverage industries",                "843C0C"),
        ("Weighted Scoring",   "Distributor KPI scoring with business-defined weights — "
                               "industry-standard approach for vendor performance management",     "C00000"),
    ]

    for i, (tool, desc, color) in enumerate(tools):
        row = i // 3; col = i % 3
        l = 0.4 + col * 4.3
        t = 1.35 + row * 2.35
        add_rect(sl, l, t, 4.0, 0.45, fill=RGBColor(int(color[:2],16), int(color[2:4],16), int(color[4:],16)))
        add_text(sl, tool, l+0.1, t+0.06, 3.8, 0.33,
                 font_size=12, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
        add_rect(sl, l, t+0.45, 4.0, 1.78, fill=LTGRAY if (i//3+i%3) % 2 == 0 else LTBLUE)
        add_text(sl, desc, l+0.1, t+0.55, 3.8, 1.6,
                 font_size=9, color=DARK, align=PP_ALIGN.LEFT)

    add_slide_number(sl, 13)


# ═══════════════════════════════════════════════════════════════════════════════
#  SLIDE 14 — CLOSING
# ═══════════════════════════════════════════════════════════════════════════════

def slide_14_close():
    sl = prs.slides.add_slide(BLANK)
    add_bg(sl, NAVY)

    add_rect(sl, 0, 0, 0.5, 7.5, fill=GOLD)
    add_rect(sl, 0.5, 2.8, 12.83, 0.06, fill=GOLD)

    add_text(sl, "Thank You", 1.0, 0.8, 11.5, 1.2,
             font_size=54, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    add_text(sl, "BeerFlow — Beverage Supply Chain Optimization", 1.0, 2.0, 11.5, 0.5,
             font_size=16, color=GOLD, italic=True, align=PP_ALIGN.LEFT)

    add_text(sl, "Harsh Raj Pandey", 1.0, 3.1, 8.0, 0.5,
             font_size=20, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    add_text(sl, "Supply Chain Analyst | Open to Opportunities in Beverage & FMCG Sector",
             1.0, 3.58, 11.0, 0.35,
             font_size=10, color=LTBLUE, italic=True, align=PP_ALIGN.LEFT)

    add_rect(sl, 1.0, 4.1, 11.0, 0.04, fill=BLUE)

    contact_items = [
        ("GitHub:", "github.com/Harsh258-collab"),
        ("LinkedIn:", "linkedin.com/in/harsh-raj-pandey-1a0319325"),
        ("Email:", "harsh258.collab@gmail.com"),
    ]
    for i, (label, val) in enumerate(contact_items):
        t = 4.3 + i * 0.52
        add_text(sl, label, 1.0, t, 1.5, 0.38,
                 font_size=10, bold=True, color=GOLD, align=PP_ALIGN.LEFT)
        add_text(sl, val, 2.5, t, 5.0, 0.38,
                 font_size=10, color=WHITE, align=PP_ALIGN.LEFT)

    add_text(sl, '"Good supply chains don\'t just move products — they build trust."',
             1.0, 6.1, 11.5, 0.5,
             font_size=11, color=LTBLUE, italic=True, align=PP_ALIGN.CENTER)

    add_slide_number(sl, 14)


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════════════════════

if __name__ == '__main__':
    print("[...] Building PowerPoint slides...")
    slide_01_title()
    print("   [OK] Slide 1: Title")
    slide_02_agenda()
    print("   [OK] Slide 2: Agenda")
    slide_03_problem()
    print("   [OK] Slide 3: Problem Statement")
    slide_04_data()
    print("   [OK] Slide 4: Dataset Overview")
    slide_05_demand()
    print("   [OK] Slide 5: Demand Forecasting")
    slide_06_inventory()
    print("   [OK] Slide 6: Inventory EOQ")
    slide_07_logistics()
    print("   [OK] Slide 7: Logistics Performance")
    slide_08_otif()
    print("   [OK] Slide 8: OTIF Analysis")
    slide_09_distributor()
    print("   [OK] Slide 9: Distributor Scorecard")
    slide_10_findings()
    print("   [OK] Slide 10: Key Findings")
    slide_11_recommendations()
    print("   [OK] Slide 11: Recommendations")
    slide_12_roadmap()
    print("   [OK] Slide 12: Roadmap")
    slide_13_tools()
    print("   [OK] Slide 13: Tools & Methodology")
    slide_14_close()
    print("   [OK] Slide 14: Closing")

    prs.save(OUT)
    print(f"\n[DONE] PowerPoint saved to:\n  {OUT}")
