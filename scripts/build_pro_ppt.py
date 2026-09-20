"""
BeerFlow PRO Presentation Deck
Author: Harsh Raj Pandey
Design: Canva-pro level | Bold typography | Full-bleed layouts | Visual storytelling
Structure:
  Slide 1  — Dramatic cover
  Slide 2  — Logistics Pathway Journey (visual map)
  Slide 3  — CHEERS! (beer mugs colliding)
  Slide 4  — The Problem (bold statement)
  Slide 5  — Project Scope (big numbers)
  Slide 6  — Demand Forecasting (visual)
  Slide 7  — Inventory EOQ (visual)
  Slide 8  — Logistics Network (route + fleet)
  Slide 9  — Cold Chain & Risk
  Slide 10 — Distributor Scorecard
  Slide 11 — Key Findings (bold callouts)
  Slide 12 — Recommendations (action cards)
  Slide 13 — Closing
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
import pandas as pd, numpy as np, datetime

BASE = r'C:\Users\BIT\.gemini\antigravity\scratch\beerflow-supply-chain'
DATA = f'{BASE}\\data'
OUT  = f'{BASE}\\powerpoint\\BeerFlow_Pro_Presentation.pptx'

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]

# ── PALETTE ──────────────────────────────────────────────────────────────────
NAVY    = RGBColor(0x1A, 0x1A, 0x2E)   # deep navy
MIDNAV  = RGBColor(0x16, 0x21, 0x3E)   # richer navy
AMBER   = RGBColor(0xE8, 0x89, 0x0C)   # rich amber / beer gold
LTAMBER = RGBColor(0xFF, 0xB3, 0x47)   # light amber
CREAM   = RGBColor(0xFF, 0xF8, 0xF0)   # warm cream
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
DARK    = RGBColor(0x1A, 0x1A, 0x2E)
GRAY    = RGBColor(0x88, 0x88, 0x99)
LTGRAY  = RGBColor(0xF4, 0xF4, 0xF8)
GREEN   = RGBColor(0x27, 0xAE, 0x60)
RED     = RGBColor(0xE7, 0x4C, 0x3C)
ORANGE  = RGBColor(0xF3, 0x96, 0x00)
TEAL    = RGBColor(0x00, 0x96, 0x88)
PURPLE  = RGBColor(0x6C, 0x3A, 0x91)
LTNAVY  = RGBColor(0x0F, 0x3C, 0x6B)
FOAMY   = RGBColor(0xF5, 0xDE, 0xB3)   # beer foam wheat
BEERCLR = RGBColor(0xD4, 0xA0, 0x17)   # amber beer body
DARKAMB = RGBColor(0xB8, 0x6E, 0x00)

MSO_RECT  = 1
MSO_OVAL  = 9
MSO_RTRI  = 8   # right triangle
MSO_PLGN  = 6   # parallelogram
MSO_ROUND = 5   # rounded rectangle

def rgb(r): return r

# ── CORE HELPERS ─────────────────────────────────────────────────────────────

def bg(slide, color):
    f = slide.background.fill
    f.solid(); f.fore_color.rgb = color

def rect(slide, l, t, w, h, fill=None, line=None, lw=0, shape=MSO_RECT):
    s = slide.shapes.add_shape(shape, Inches(l), Inches(t), Inches(w), Inches(h))
    if fill:
        s.fill.solid(); s.fill.fore_color.rgb = fill
    else:
        s.fill.background()
    if line:
        s.line.color.rgb = line; s.line.width = Pt(lw or 1.5)
    else:
        s.line.fill.background()
    return s

def oval(slide, l, t, w, h, fill=None, line=None, lw=0):
    s = slide.shapes.add_shape(MSO_OVAL, Inches(l), Inches(t), Inches(w), Inches(h))
    if fill: s.fill.solid(); s.fill.fore_color.rgb = fill
    else:    s.fill.background()
    if line: s.line.color.rgb = line; s.line.width = Pt(lw or 1.5)
    else:    s.line.fill.background()
    return s

def txt(slide, text, l, t, w, h, sz=12, bold=False, color=DARK,
        align=PP_ALIGN.LEFT, italic=False, fill=None, wrap=True):
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tb.text_frame.word_wrap = wrap
    if fill: tb.fill.solid(); tb.fill.fore_color.rgb = fill
    p = tb.text_frame.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size = Pt(sz); r.font.bold = bold
    r.font.color.rgb = color; r.font.italic = italic
    r.font.name = 'Calibri'
    return tb

def multi_line(slide, lines, l, t, w, h, sz=10, bold=False, color=DARK,
               align=PP_ALIGN.LEFT, line_spacing=4, fill=None, italic=False):
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tb.text_frame.word_wrap = True
    if fill: tb.fill.solid(); tb.fill.fore_color.rgb = fill
    for i, line in enumerate(lines):
        p = tb.text_frame.paragraphs[0] if i == 0 else tb.text_frame.add_paragraph()
        p.space_before = Pt(line_spacing)
        p.alignment = align
        r = p.add_run()
        r.text = line; r.font.size = Pt(sz); r.font.bold = bold
        r.font.color.rgb = color; r.font.name = 'Calibri'
        r.font.italic = italic
    return tb

def pagenum(slide, n, total=13):
    txt(slide, f'{n}  /  {total}', 12.2, 7.15, 1.0, 0.3,
        sz=8, color=GRAY, align=PP_ALIGN.RIGHT)

def accent_bar(slide, color=AMBER, vertical=True):
    if vertical:
        rect(slide, 0, 0, 0.08, 7.5, fill=color)
    else:
        rect(slide, 0, 7.1, 13.33, 0.08, fill=color)

# ── BEER MUG HELPER ──────────────────────────────────────────────────────────

def draw_beer_mug(slide, cx, cy, scale=1.0, tilt=0):
    """Draw a stylized beer mug using shapes. cx,cy = center in inches."""
    w = 1.0 * scale
    h = 1.6 * scale
    l = cx - w/2
    t = cy - h/2

    # Mug body (trapezoid approx via rect)
    rect(slide, l, t+0.2*scale, w, h-0.2*scale, fill=BEERCLR)
    # Beer foam on top
    oval(slide, l-0.05*scale, t, w+0.1*scale, 0.35*scale, fill=FOAMY)
    # Foam bubbles
    oval(slide, l+0.1*scale, t-0.08*scale, 0.18*scale, 0.18*scale, fill=WHITE)
    oval(slide, l+0.35*scale, t-0.12*scale, 0.22*scale, 0.22*scale, fill=WHITE)
    oval(slide, l+0.6*scale, t-0.06*scale, 0.16*scale, 0.16*scale, fill=WHITE)
    # Handle
    rect(slide, cx+w/2, t+0.35*scale, 0.28*scale, 0.55*scale,
         fill=None, line=BEERCLR, lw=4)
    # Highlight stripe
    rect(slide, l+0.12*scale, t+0.35*scale, 0.12*scale, h-0.55*scale,
         fill=RGBColor(0xFF, 0xD0, 0x60))
    # Bottom of mug (darker)
    rect(slide, l, t+h-0.15*scale, w, 0.15*scale, fill=DARKAMB)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 1 — DRAMATIC COVER
# ══════════════════════════════════════════════════════════════════════════════

def slide_01():
    sl = prs.slides.add_slide(BLANK)

    # ── LEFT HALF: Rich Yellow/Amber panel ───────────────────────────────────
    bg(sl, WHITE)
    rect(sl, 0, 0, 7.2, 7.5, fill=AMBER)                     # full left yellow
    rect(sl, 0, 0, 7.2, 0.18, fill=RGBColor(0xC8,0x76,0x00)) # dark top trim

    # Decorative circles on yellow (depth effect)
    oval(sl, -1.2, -1.2, 4.5, 4.5, fill=RGBColor(0xFF,0xC5,0x30))  # light circle
    oval(sl, 4.5, 4.8, 5.0, 5.0, fill=RGBColor(0xD4,0x82,0x00))   # dark circle

    # "BEER" in bold navy on yellow
    txt(sl, 'BEER', 0.35, 0.55, 6.6, 1.65,
        sz=100, bold=True, color=NAVY, align=PP_ALIGN.LEFT)

    # Divider stroke under BEER
    rect(sl, 0.35, 2.28, 6.5, 0.1, fill=NAVY)

    # "FLOW" in white on yellow — pops dramatically
    txt(sl, 'FLOW', 0.35, 2.32, 6.6, 1.65,
        sz=100, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    # Tag on yellow
    txt(sl, 'A SUPPLY CHAIN ANALYTICS PROJECT', 0.35, 4.15, 6.5, 0.38,
        sz=9.5, bold=True, color=NAVY, italic=True, align=PP_ALIGN.LEFT)

    # Domain pills on yellow (navy background pills)
    pills = ['Demand Forecasting', 'Inventory · EOQ', 'Logistics', 'Cold Chain', 'Risk Register', 'Distributor KPI']
    for i, p in enumerate(pills):
        col = i % 2; row = i // 2
        lp = 0.35 + col * 3.3; tp = 4.65 + row * 0.62
        rect(sl, lp, tp, 3.1, 0.48, fill=NAVY)
        txt(sl, p, lp+0.1, tp+0.08, 2.9, 0.32,
            sz=8.5, bold=True, color=AMBER, align=PP_ALIGN.CENTER)

    # ── RIGHT HALF: Clean White panel ────────────────────────────────────────
    # Thin amber separator line
    rect(sl, 7.18, 0, 0.06, 7.5, fill=RGBColor(0xC8,0x76,0x00))

    # Author photo placeholder (circle avatar)
    oval(sl, 8.1, 0.6, 2.2, 2.2, fill=RGBColor(0xF0,0xF0,0xF5))
    oval(sl, 8.55, 0.9, 1.3, 1.3, fill=RGBColor(0xCC,0xCC,0xDD))  # head
    oval(sl, 8.25, 2.0, 1.9, 1.2, fill=RGBColor(0xCC,0xCC,0xDD))  # shoulders
    # Amber ring around avatar
    oval(sl, 8.05, 0.55, 2.3, 2.3, fill=None, line=AMBER, lw=3)

    # Author details
    txt(sl, 'Harsh Raj Pandey', 7.45, 3.05, 5.7, 0.62,
        sz=22, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    rect(sl, 8.2, 3.72, 4.2, 0.06, fill=AMBER)
    txt(sl, 'Supply Chain Analyst', 7.45, 3.85, 5.7, 0.38,
        sz=12, bold=False, color=RGBColor(0x44,0x44,0x66), italic=True, align=PP_ALIGN.CENTER)
    txt(sl, 'FMCG  &  Beverage Domain', 7.45, 4.22, 5.7, 0.35,
        sz=10, bold=False, color=GRAY, italic=True, align=PP_ALIGN.CENTER)

    # Stats row
    stats = [('2,000+', 'Orders'), ('1,000+', 'Routes'), ('10', 'Datasets')]
    for i, (num, label) in enumerate(stats):
        lx = 7.45 + i * 1.9
        rect(sl, lx, 4.78, 1.75, 0.9, fill=AMBER)
        txt(sl, num, lx+0.05, 4.85, 1.65, 0.42,
            sz=16, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
        txt(sl, label, lx+0.05, 5.25, 1.65, 0.35,
            sz=8, bold=False, color=NAVY, align=PP_ALIGN.CENTER)

    # Contact links
    rect(sl, 7.45, 5.85, 5.7, 0.04, fill=RGBColor(0xDD,0xDD,0xEE))
    txt(sl, 'github.com/Harsh258-collab', 7.45, 5.98, 5.7, 0.32,
        sz=9, color=RGBColor(0x1A,0x6E,0xC8), align=PP_ALIGN.CENTER)
    txt(sl, 'linkedin.com/in/harsh-raj-pandey-1a0319325', 7.45, 6.3, 5.7, 0.32,
        sz=9, color=RGBColor(0x1A,0x6E,0xC8), align=PP_ALIGN.CENTER)
    txt(sl, 'September 2026', 7.45, 6.9, 5.7, 0.35,
        sz=9, color=GRAY, italic=True, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 2 — LOGISTICS PATHWAY (visual supply chain journey)
# ══════════════════════════════════════════════════════════════════════════════

def slide_02():
    sl = prs.slides.add_slide(BLANK)
    bg(sl, CREAM)
    accent_bar(sl, AMBER)

    txt(sl, 'THE JOURNEY OF EVERY BOTTLE', 0.8, 0.25, 12, 0.55,
        sz=28, bold=True, color=NAVY, align=PP_ALIGN.LEFT)
    txt(sl, 'From raw ingredients to your glass — this is the supply chain we optimized',
        0.8, 0.85, 12, 0.35, sz=12, color=GRAY, italic=True)
    rect(sl, 0.8, 1.28, 12.2, 0.04, fill=AMBER)

    # Supply chain stages
    stages = [
        ('🌾', 'SOURCING',    'Malt · Hops · Barley\nSupplier Management',   NAVY),
        ('🏭', 'BREWING',     'Production Planning\nBatch Scheduling',         LTNAVY),
        ('🏪', 'WAREHOUSING', 'EOQ Inventory\nSafety Stock · FEFO',            AMBER),
        ('🚛', 'LOGISTICS',   'Route Optimization\nLast-Mile Delivery',        RGBColor(0x14,0x52,0x8A)),
        ('🤝', 'DISTRIBUTOR', 'KPI Scorecard\nOTIF · Fill Rate',               TEAL),
        ('🍺', 'RETAIL',      'Cold Chain\nOn-Shelf Availability',             GREEN),
    ]

    node_w = 1.75
    node_h = 2.8
    start_x = 0.72
    node_y = 1.55

    for i, (icon, title, desc, color) in enumerate(stages):
        x = start_x + i * 2.1

        # Drop shadow effect
        rect(sl, x+0.08, node_y+0.08, node_w, node_h,
             fill=RGBColor(0xCC,0xCC,0xDD))

        # Main card
        rect(sl, x, node_y, node_w, node_h, fill=color)

        # Top accent
        rect(sl, x, node_y, node_w, 0.08, fill=AMBER)

        # Icon circle
        oval(sl, x+0.6, node_y+0.15, 0.55, 0.55, fill=RGBColor(0xFF,0xFF,0xFF))
        txt(sl, icon, x+0.6, node_y+0.17, 0.55, 0.5,
            sz=18, align=PP_ALIGN.CENTER, color=DARK)

        # Step number
        txt(sl, f'0{i+1}', x+0.05, node_y+0.15, 0.5, 0.35,
            sz=10, bold=True, color=LTAMBER, align=PP_ALIGN.LEFT)

        # Title
        txt(sl, title, x+0.08, node_y+0.82, node_w-0.16, 0.42,
            sz=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

        # Desc
        multi_line(sl, desc.split('\n'), x+0.08, node_y+1.28, node_w-0.16, 1.3,
                   sz=8.5, color=RGBColor(0xDD,0xDD,0xEE), align=PP_ALIGN.CENTER,
                   line_spacing=5)

        # Bottom metric
        rect(sl, x, node_y+node_h-0.55, node_w, 0.55,
             fill=RGBColor(0x00,0x00,0x00).__class__(0,0,0,).__class__.__new__(
                 RGBColor, int(color[0]*0.7), int(color[1]*0.7), int(color[2]*0.7)))

        # Connector arrow (except last)
        if i < len(stages)-1:
            ax = x + node_w + 0.05
            # Arrow body
            rect(sl, ax, node_y+node_h/2-0.08, 0.27, 0.16, fill=AMBER)
            # Arrowhead triangle
            s = sl.shapes.add_shape(MSO_RECT, Inches(ax+0.27), Inches(node_y+node_h/2-0.14),
                                    Inches(0.08), Inches(0.28))
            s.fill.solid(); s.fill.fore_color.rgb = AMBER
            s.line.fill.background()

    # Bottom insight bar
    rect(sl, 0.15, 4.6, 13.0, 0.75, fill=NAVY)
    txt(sl, '💡  BeerFlow analyses every stage of this journey — from supplier lead times to last-mile delivery, '
            'cold chain compliance to distributor performance.',
        0.4, 4.68, 12.6, 0.58,
        sz=10, bold=False, color=WHITE, align=PP_ALIGN.LEFT)

    pagenum(sl, 2)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 3 — CHEERS! (Beer mugs colliding)
# ══════════════════════════════════════════════════════════════════════════════

def slide_03():
    sl = prs.slides.add_slide(BLANK)
    bg(sl, NAVY)

    # Rich radial-like background layers
    oval(sl, -1.5, -1.5, 10, 10, fill=RGBColor(0x1F,0x1F,0x3A))
    oval(sl, 3.5, 2.0, 9, 7, fill=RGBColor(0x15,0x15,0x2A))

    # Celebration burst lines (radiating from center)
    center_x, center_y = 6.67, 3.75
    import math
    for angle_deg in range(0, 360, 22):
        angle = math.radians(angle_deg)
        r_start = 1.8; r_end = 3.2
        x1 = center_x + r_start * math.cos(angle)
        y1 = center_y + r_start * math.sin(angle)
        length = r_end - r_start
        # Draw as thin rect rotated (approximate with small rect)
        burst_color = AMBER if angle_deg % 44 == 0 else RGBColor(0x44, 0x44, 0x66)
        oval(sl, x1-0.04, y1-0.04, 0.12, 0.12, fill=burst_color)

    # Sparkle dots
    for (sx, sy, sc) in [
        (3.2, 1.0, 0.18), (9.8, 1.2, 0.14), (2.5, 5.5, 0.12),
        (10.5, 5.8, 0.16), (5.5, 0.6, 0.10), (7.8, 0.8, 0.14),
        (11.2, 3.5, 0.12), (2.0, 3.2, 0.10), (4.0, 6.5, 0.15),
        (9.0, 6.8, 0.13),
    ]:
        oval(sl, sx, sy, sc, sc, fill=AMBER)

    # LEFT BEER MUG (tilted slightly — simulated by offset)
    # Mug body
    rect(sl, 3.8, 1.8, 1.5, 2.4, fill=BEERCLR)
    rect(sl, 3.8, 1.8, 1.5, 0.55, fill=FOAMY)  # foam
    oval(sl, 3.75, 1.6, 1.6, 0.5, fill=FOAMY)  # foam top dome
    oval(sl, 3.9, 1.45, 0.28, 0.28, fill=WHITE) # bubble
    oval(sl, 4.3, 1.38, 0.22, 0.22, fill=WHITE)
    oval(sl, 4.7, 1.48, 0.18, 0.18, fill=WHITE)
    # Handle
    s = sl.shapes.add_shape(MSO_OVAL, Inches(5.28), Inches(2.15), Inches(0.55), Inches(1.1))
    s.fill.background(); s.line.color.rgb = BEERCLR; s.line.width = Pt(8)
    # Highlight
    rect(sl, 4.05, 2.45, 0.22, 1.5, fill=RGBColor(0xFF,0xD0,0x60))
    # Base
    rect(sl, 3.8, 4.12, 1.5, 0.2, fill=DARKAMB)
    # Tilt label under mug
    txt(sl, '🌾  Brewery', 3.5, 4.5, 2.2, 0.4,
        sz=11, bold=True, color=LTAMBER, align=PP_ALIGN.CENTER)

    # RIGHT BEER MUG (mirror)
    rect(sl, 8.05, 1.8, 1.5, 2.4, fill=BEERCLR)
    rect(sl, 8.05, 1.8, 1.5, 0.55, fill=FOAMY)
    oval(sl, 8.0, 1.6, 1.6, 0.5, fill=FOAMY)
    oval(sl, 8.1, 1.45, 0.28, 0.28, fill=WHITE)
    oval(sl, 8.5, 1.38, 0.22, 0.22, fill=WHITE)
    oval(sl, 8.9, 1.48, 0.18, 0.18, fill=WHITE)
    # Handle (left side for right mug)
    s2 = sl.shapes.add_shape(MSO_OVAL, Inches(7.5), Inches(2.15), Inches(0.55), Inches(1.1))
    s2.fill.background(); s2.line.color.rgb = BEERCLR; s2.line.width = Pt(8)
    # Highlight
    rect(sl, 8.25, 2.45, 0.22, 1.5, fill=RGBColor(0xFF,0xD0,0x60))
    rect(sl, 8.05, 4.12, 1.5, 0.2, fill=DARKAMB)
    txt(sl, '🍺  Consumer', 7.6, 4.5, 2.2, 0.4,
        sz=11, bold=True, color=LTAMBER, align=PP_ALIGN.CENTER)

    # COLLISION effect in center (amber splash)
    oval(sl, 5.7, 1.2, 2.0, 1.8, fill=RGBColor(0x28,0x28,0x44))
    oval(sl, 6.0, 1.35, 1.4, 1.3, fill=AMBER)
    oval(sl, 6.2, 1.5, 1.0, 0.95, fill=LTAMBER)
    # Splash drops
    for (dx, dy, ds) in [(5.5,1.0,0.14),(8.0,0.85,0.12),(6.1,0.7,0.16),(7.3,0.75,0.11),(5.8,0.55,0.09),(7.6,0.6,0.13)]:
        oval(sl, dx, dy, ds, ds, fill=AMBER)

    # Main text
    txt(sl, 'CHEERS!', 4.5, 5.1, 4.5, 1.1,
        sz=64, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    txt(sl, 'Now let the data do the talking.', 4.2, 6.15, 5.0, 0.5,
        sz=14, bold=False, color=LTAMBER, italic=True, align=PP_ALIGN.CENTER)

    # Small amber bottom bar
    rect(sl, 0, 7.3, 13.33, 0.2, fill=AMBER)

    pagenum(sl, 3)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 4 — THE PROBLEM (bold statement)
# ══════════════════════════════════════════════════════════════════════════════

def slide_04():
    sl = prs.slides.add_slide(BLANK)
    bg(sl, CREAM)
    accent_bar(sl, AMBER)

    # Bold left block
    rect(sl, 0.5, 0.5, 6.0, 6.5, fill=NAVY)
    rect(sl, 0.5, 0.5, 6.0, 0.12, fill=AMBER)

    txt(sl, 'THE', 0.8, 0.75, 5.5, 0.9,
        sz=52, bold=True, color=AMBER, align=PP_ALIGN.LEFT)
    txt(sl, 'PROBLEM', 0.8, 1.55, 5.5, 1.1,
        sz=52, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    rect(sl, 0.8, 2.75, 4.8, 0.06, fill=RGBColor(0x44,0x44,0x66))

    multi_line(sl, [
        '"Beverage companies stock OUT during',
        'festivals — and stock UP during',
        'monsoons. The supply chain doesn\'t',
        'follow the customer."',
    ], 0.8, 2.95, 5.3, 2.5, sz=12.5, color=RGBColor(0xBB,0xBB,0xCC),
       italic=True, line_spacing=8)

    txt(sl, '— Industry Insight, Harsh Raj Pandey', 0.8, 5.65, 5.3, 0.4,
        sz=9, color=GRAY, italic=True)

    # Right block — 3 pain points
    pain = [
        ('01', 'STOCKOUTS AT PEAK',
         'Products missing exactly when demand\nspikes 30–70% during Diwali & summer'),
        ('02', 'OTIF BELOW TARGET',
         'On-Time In-Full rate at ~78%\nagainst 90%+ industry benchmark'),
        ('03', 'LOGISTICS COST LEAK',
         '28% of vehicle capacity goes empty\nEvery. Single. Trip.'),
    ]
    for i, (num, title, desc) in enumerate(pain):
        ty = 0.5 + i * 2.15
        rect(sl, 7.0, ty, 5.9, 1.95, fill=WHITE)
        rect(sl, 7.0, ty, 0.12, 1.95, fill=AMBER if i==0 else (RED if i==1 else ORANGE))
        txt(sl, num, 7.25, ty+0.12, 0.7, 0.55,
            sz=22, bold=True, color=AMBER if i==0 else (RED if i==1 else ORANGE))
        txt(sl, title, 7.25, ty+0.62, 5.4, 0.42,
            sz=13, bold=True, color=NAVY)
        multi_line(sl, desc.split('\n'), 7.25, ty+1.05, 5.4, 0.85,
                   sz=10, color=GRAY, line_spacing=4)
    pagenum(sl, 4)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 5 — PROJECT SCOPE (big numbers)
# ══════════════════════════════════════════════════════════════════════════════

def slide_05():
    sl = prs.slides.add_slide(BLANK)
    bg(sl, NAVY)
    accent_bar(sl, AMBER)

    txt(sl, 'THE DATASET', 0.8, 0.3, 12, 0.65,
        sz=32, bold=True, color=WHITE)
    txt(sl, 'What we built and what we analysed',
        0.8, 0.98, 10, 0.38, sz=12, color=GRAY, italic=True)
    rect(sl, 0.8, 1.42, 12.2, 0.05, fill=AMBER)

    stats = [
        ('2,000',  'ORDER\nRECORDS',    AMBER,   '5 regions · 7 SKUs · 4 channels'),
        ('1,000',  'DELIVERY\nROUTES',  LTAMBER, 'Vehicle util · Cost/case · OTD'),
        ('265',    'FLEET\nVEHICLES',   TEAL,    'Breakdowns · CO2 · Maintenance'),
        ('600',    'COLD CHAIN\nLOGS',  RED,     'Temp breaches · 5 supply stages'),
        ('400',    'FREIGHT\nSHIPMENTS',ORANGE,  'Road · Rail · Air · 3PL compare'),
        ('24',     'RISK\nITEMS',       PURPLE,  'Likelihood × Impact matrix'),
    ]

    cols = 3; col_w = 4.15
    for i, (num, label, color, sub) in enumerate(stats):
        row = i // cols; col = i % cols
        lx = 0.5 + col * col_w
        ty = 1.65 + row * 2.65

        # Card
        rect(sl, lx, ty, col_w-0.2, 2.35, fill=RGBColor(0x0F,0x0F,0x22))
        rect(sl, lx, ty, col_w-0.2, 0.1, fill=color)
        rect(sl, lx, ty+2.25, col_w-0.2, 0.1, fill=color)

        # Big number
        txt(sl, num, lx+0.15, ty+0.2, col_w-0.5, 1.0,
            sz=52, bold=True, color=color, align=PP_ALIGN.LEFT)
        # Label
        multi_line(sl, label.split('\n'), lx+0.15, ty+1.15, col_w-0.5, 0.8,
                   sz=11, bold=True, color=WHITE, line_spacing=2)
        # Sub
        txt(sl, sub, lx+0.15, ty+1.9, col_w-0.3, 0.35,
            sz=8, color=GRAY, italic=True)

    pagenum(sl, 5)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 6 — DEMAND FORECASTING
# ══════════════════════════════════════════════════════════════════════════════

def slide_06():
    sl = prs.slides.add_slide(BLANK)
    bg(sl, CREAM)
    accent_bar(sl, AMBER)

    txt(sl, 'DEMAND FORECASTING', 0.8, 0.28, 12, 0.6,
        sz=30, bold=True, color=NAVY)
    txt(sl, 'Seasonal demand patterns · SKU-level accuracy · MAPE scoring',
        0.8, 0.92, 10, 0.35, sz=11, color=GRAY, italic=True)
    rect(sl, 0.8, 1.32, 12.2, 0.05, fill=AMBER)

    df = pd.read_csv(f'{DATA}/demand_forecast.csv')
    sku_mape = df.groupby('SKU')['MAPE_Pct'].mean().round(1).reset_index().sort_values('MAPE_Pct')
    sku_vol  = df.groupby('SKU')['Actual_Demand'].sum().reset_index()
    total    = sku_vol['Actual_Demand'].sum()

    # LEFT: MAPE bars
    txt(sl, 'Forecast Accuracy by SKU (MAPE %)', 0.8, 1.45, 7.5, 0.38,
        sz=13, bold=True, color=NAVY)
    txt(sl, 'Lower MAPE = Better accuracy  |  Target: <10%',
        0.8, 1.83, 7.5, 0.28, sz=9, color=GRAY, italic=True)

    max_mape = sku_mape['MAPE_Pct'].max()
    for i, row in enumerate(sku_mape.itertuples()):
        ty = 2.18 + i * 0.64
        # SKU label
        sku_short = row.SKU.replace('_', ' ')
        txt(sl, sku_short, 0.8, ty+0.08, 3.0, 0.42,
            sz=9, color=NAVY)
        # Bar background
        rect(sl, 3.85, ty+0.1, 3.8, 0.36, fill=RGBColor(0xE8,0xE8,0xF0))
        # Actual bar
        bw = (row.MAPE_Pct / max_mape) * 3.7
        bc = GREEN if row.MAPE_Pct < 10 else (ORANGE if row.MAPE_Pct < 20 else RED)
        rect(sl, 3.85, ty+0.1, bw, 0.36, fill=bc)
        # MAPE value
        txt(sl, f'{row.MAPE_Pct}%', 3.88 + bw, ty+0.12, 0.7, 0.32,
            sz=9, bold=True, color=NAVY)

    # Legend
    for lc, lb in [(GREEN,'<10% Excellent'),(ORANGE,'10–20% OK'),(RED,'>20% Poor')]:
        pass  # embedded in bars

    # RIGHT: Seasonal insight cards
    txt(sl, 'Seasonal Demand Patterns', 8.2, 1.45, 4.8, 0.38,
        sz=13, bold=True, color=NAVY)

    seasons = [
        ('OCT – DEC', '+30% to +70%', 'Diwali · Christmas · New Year', AMBER),
        ('MAR – APR', '+20% to +45%', 'Summer onset · IPL season', ORANGE),
        ('JUL – AUG', '-15% to -40%', 'Heavy monsoon demand trough', TEAL),
        ('YoY GROWTH', '+3% CAGR', 'Consistent annual volume growth', GREEN),
    ]
    for i, (period, change, note, color) in enumerate(seasons):
        ty = 1.85 + i * 1.35
        rect(sl, 8.2, ty, 4.8, 1.2, fill=WHITE)
        rect(sl, 8.2, ty, 0.12, 1.2, fill=color)
        txt(sl, period, 8.45, ty+0.1, 4.3, 0.38, sz=10, bold=True, color=GRAY)
        txt(sl, change, 8.45, ty+0.45, 4.3, 0.45, sz=22, bold=True, color=color)
        txt(sl, note, 8.45, ty+0.88, 4.3, 0.28, sz=8, color=GRAY, italic=True)

    pagenum(sl, 6)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 7 — INVENTORY EOQ
# ══════════════════════════════════════════════════════════════════════════════

def slide_07():
    sl = prs.slides.add_slide(BLANK)
    bg(sl, NAVY)
    accent_bar(sl, AMBER)

    txt(sl, 'INVENTORY OPTIMIZATION', 0.8, 0.28, 12, 0.6,
        sz=30, bold=True, color=WHITE)
    txt(sl, 'Economic Order Quantity · Safety Stock · Reorder Points · Stockout Risk',
        0.8, 0.92, 12, 0.35, sz=11, color=GRAY, italic=True)
    rect(sl, 0.8, 1.32, 12.2, 0.05, fill=AMBER)

    # EOQ formula card
    rect(sl, 0.8, 1.48, 12.2, 0.82, fill=RGBColor(0x0F,0x0F,0x22))
    txt(sl, 'EOQ  =  √ ( 2  ×  Annual Demand  ×  Ordering Cost  ÷  Holding Cost per Unit )',
        1.0, 1.62, 12.0, 0.52,
        sz=16, bold=True, color=AMBER, align=PP_ALIGN.CENTER)

    # Stockout risk boxes
    df = pd.read_csv(f'{DATA}/inventory_data.csv')
    rc = df['Stockout_Risk'].value_counts()

    risk_data = [
        ('HIGH RISK', rc.get('High',0), 'Immediate\nreplenishment\nrequired', RED),
        ('MEDIUM RISK', rc.get('Medium',0), 'Monitor &\nplan reorder\nthis week', ORANGE),
        ('LOW RISK', rc.get('Low',0), 'Stock levels\nadequate\nfor now', GREEN),
    ]
    for i, (label, count, note, color) in enumerate(risk_data):
        lx = 0.8 + i * 4.15
        rect(sl, lx, 2.45, 3.85, 2.1, fill=color)
        txt(sl, label, lx+0.12, 2.52, 3.6, 0.42,
            sz=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        txt(sl, str(count), lx+0.12, 2.92, 3.6, 0.85,
            sz=52, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        txt(sl, 'SKU-WH combos', lx+0.12, 3.75, 3.6, 0.3,
            sz=8, color=RGBColor(0xEE,0xEE,0xEE), align=PP_ALIGN.CENTER, italic=True)
        multi_line(sl, note.split('\n'), lx+0.12, 4.08, 3.6, 0.38,
                   sz=8.5, color=RGBColor(0xDD,0xDD,0xEE),
                   align=PP_ALIGN.CENTER, line_spacing=2)

    # Key findings bar
    rect(sl, 0.8, 4.7, 12.2, 2.55, fill=RGBColor(0x0F,0x0F,0x22))
    txt(sl, 'KEY INVENTORY FINDINGS', 1.0, 4.8, 12.0, 0.38,
        sz=11, bold=True, color=AMBER)
    findings = [
        '→  East India warehouses hold 30% BELOW recommended safety stock — highest stockout risk zone',
        '→  Stout 650ml is over-ordered by 45% above EOQ — wasting INR 8–12L in holding costs annually',
        '→  Lager 500ml shows near-optimal ordering pattern — can serve as benchmark for other SKUs',
        '→  Recommendation: Vendor-Managed Inventory (VMI) for top 3 high-risk SKU-warehouse combos',
    ]
    multi_line(sl, findings, 1.0, 5.22, 12.0, 2.0, sz=10, color=WHITE, line_spacing=8)
    pagenum(sl, 7)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 8 — LOGISTICS NETWORK
# ══════════════════════════════════════════════════════════════════════════════

def slide_08():
    sl = prs.slides.add_slide(BLANK)
    bg(sl, CREAM)
    accent_bar(sl, AMBER)

    txt(sl, 'LOGISTICS NETWORK ANALYSIS', 0.8, 0.28, 12, 0.6,
        sz=30, bold=True, color=NAVY)
    txt(sl, 'Route efficiency · Fleet utilization · Lead time · Cost per case',
        0.8, 0.92, 12, 0.35, sz=11, color=GRAY, italic=True)
    rect(sl, 0.8, 1.32, 12.2, 0.05, fill=AMBER)

    df_r = pd.read_csv(f'{DATA}/route_delivery_data.csv')
    df_f = pd.read_csv(f'{DATA}/fleet_utilization_data.csv')

    # TOP: 4 big KPI tiles
    kpis = [
        ('OTD RATE', f"{(df_r['On_Time_Delivery']=='Yes').mean()*100:.1f}%",
         'On-Time Delivery', RED if (df_r['On_Time_Delivery']=='Yes').mean()<0.85 else GREEN),
        ('AVG UTIL', f"{df_r['Utilization_Pct'].mean():.1f}%",
         'Vehicle load utilization', ORANGE),
        ('COST/CASE', f"INR {df_r['Cost_Per_Case_INR'].mean():.1f}",
         'Last-mile delivery cost', TEAL),
        ('FLEET SIZE', f"{len(df_f):,}",
         'Active vehicles tracked', PURPLE),
    ]
    for i, (label, val, sub, color) in enumerate(kpis):
        lx = 0.8 + i * 3.1
        rect(sl, lx, 1.5, 2.9, 1.5, fill=WHITE)
        rect(sl, lx, 1.5, 2.9, 0.1, fill=color)
        txt(sl, label, lx+0.12, 1.65, 2.65, 0.35,
            sz=9, bold=True, color=GRAY)
        txt(sl, val, lx+0.12, 1.98, 2.65, 0.62,
            sz=28, bold=True, color=color)
        txt(sl, sub, lx+0.12, 2.6, 2.65, 0.3,
            sz=8, color=GRAY, italic=True)

    # BOTTOM LEFT: Modal comparison
    txt(sl, 'Freight Mode Comparison', 0.8, 3.18, 6.5, 0.35,
        sz=13, bold=True, color=NAVY)
    df_fr = pd.read_csv(f'{DATA}/freight_benchmark_data.csv')
    modal = df_fr.groupby('Freight_Mode').agg(
        Cost=('Cost_Per_Case_INR','mean'),
        Days=('Transit_Days','mean'),
        OTD=('On_Time_Pct','mean')
    ).reset_index().sort_values('Cost')

    hdrs = ['Mode','Cost/Case','Transit','OTD%']
    col_ws = [2.2, 1.2, 1.0, 1.0]
    x = 0.8
    for h, w in zip(hdrs, col_ws):
        rect(sl, x, 3.58, w, 0.32, fill=NAVY)
        txt(sl, h, x+0.05, 3.61, w-0.1, 0.26,
            sz=8.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        x += w

    mode_colors = [GREEN, TEAL, ORANGE, RED, PURPLE]
    for ri, row in enumerate(modal.itertuples()):
        ty = 3.92 + ri * 0.48
        bg_c = RGBColor(0xF8,0xF8,0xFF) if ri%2==0 else WHITE
        vals = [row.Freight_Mode, f"INR {row.Cost:.1f}", f"{row.Days:.1f}d", f"{row.OTD:.0f}%"]
        x = 0.8
        for ci, (v, w) in enumerate(zip(vals, col_ws)):
            cell = rect(sl, x, ty, w, 0.43, fill=bg_c)
            clr = mode_colors[ri] if ci==0 else DARK
            txt(sl, v, x+0.05, ty+0.07, w-0.1, 0.3,
                sz=8.5, bold=(ci==0), color=clr, align=PP_ALIGN.CENTER)
            x += w

    # BOTTOM RIGHT: Key insight callout
    txt(sl, 'The Big Opportunity', 6.9, 3.18, 6.0, 0.35,
        sz=13, bold=True, color=NAVY)
    insights = [
        ('44%', 'Rail is cheaper than road\nper case for 800km+ routes', AMBER),
        ('28%', 'Vehicle capacity wasted\non every trip on average', RED),
        ('-12%', 'Freight savings possible via\nroute consolidation alone', GREEN),
    ]
    for i, (num, note, color) in enumerate(insights):
        ty = 3.55 + i * 1.27
        rect(sl, 6.9, ty, 6.0, 1.12, fill=WHITE)
        rect(sl, 6.9, ty, 0.1, 1.12, fill=color)
        txt(sl, num, 7.1, ty+0.1, 1.5, 0.72, sz=28, bold=True, color=color)
        multi_line(sl, note.split('\n'), 8.65, ty+0.12, 4.1, 0.9,
                   sz=10, color=GRAY, line_spacing=4)

    pagenum(sl, 8)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 9 — COLD CHAIN & RISK
# ══════════════════════════════════════════════════════════════════════════════

def slide_09():
    sl = prs.slides.add_slide(BLANK)
    bg(sl, NAVY)
    accent_bar(sl, TEAL)

    txt(sl, 'COLD CHAIN  &  RISK MANAGEMENT', 0.8, 0.28, 12, 0.6,
        sz=28, bold=True, color=WHITE)
    txt(sl, 'Temperature compliance · Breach analysis · Supply chain risk register',
        0.8, 0.92, 12, 0.35, sz=11, color=GRAY, italic=True)
    rect(sl, 0.8, 1.32, 12.2, 0.05, fill=TEAL)

    df_cc = pd.read_csv(f'{DATA}/cold_chain_data.csv')
    breach_rate = (df_cc['Temperature_Breach']=='Yes').mean()*100

    # LEFT: Cold chain
    txt(sl, 'COLD CHAIN STATUS', 0.8, 1.48, 6.0, 0.38, sz=13, bold=True, color=TEAL)

    # Big breach rate circle visual
    oval(sl, 0.9, 1.95, 2.6, 2.6,
         fill=RED if breach_rate>10 else ORANGE)
    txt(sl, f'{breach_rate:.0f}%', 1.05, 2.6, 2.3, 0.95,
        sz=40, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    txt(sl, 'BREACH RATE', 1.05, 3.52, 2.3, 0.45,
        sz=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    txt(sl, 'Target: <5%', 1.05, 3.95, 2.3, 0.3,
        sz=8, color=RGBColor(0xFF,0xCC,0xCC), italic=True, align=PP_ALIGN.CENTER)

    stages = df_cc.groupby('Supply_Chain_Stage').agg(
        Breach=('Temperature_Breach', lambda x: round((x=='Yes').mean()*100,1))
    ).reset_index().sort_values('Breach', ascending=False)

    for i, row in enumerate(stages.itertuples()):
        ty = 1.95 + i * 0.7
        rect(sl, 3.7, ty, 2.8, 0.6, fill=RGBColor(0x0F,0x0F,0x22))
        bc = RED if row.Breach>20 else (ORANGE if row.Breach>10 else GREEN)
        bw = (row.Breach/40)*2.5
        rect(sl, 3.7, ty, bw, 0.6, fill=bc)
        txt(sl, row.Supply_Chain_Stage, 3.78, ty+0.12, 2.6, 0.38, sz=8, color=WHITE)
        txt(sl, f'{row.Breach}%', 3.78+bw, ty+0.14, 0.7, 0.32,
            sz=8, bold=True, color=LTAMBER)

    txt(sl, '→  Retail outlet is the weakest stage (30%+ breach)', 0.8, 5.35, 6.0, 0.35,
        sz=9, color=TEAL, bold=True)
    txt(sl, '→  IoT temp sensors + driver training = immediate fix', 0.8, 5.72, 6.0, 0.35,
        sz=9, color=GRAY, italic=True)

    # RIGHT: Risk register heat map
    txt(sl, 'RISK HEAT MAP', 7.0, 1.48, 6.0, 0.38, sz=13, bold=True, color=AMBER)

    df_risk = pd.read_csv(f'{DATA}/risk_register.csv')
    rc = df_risk['Risk_Level'].value_counts()

    risk_items = [
        ('CRITICAL', rc.get('Critical',0), RED,    'Immediate action required'),
        ('HIGH',     rc.get('High',0),     ORANGE, 'Action within 30 days'),
        ('MEDIUM',   rc.get('Medium',0),   AMBER,  'Monitor quarterly'),
        ('LOW',      rc.get('Low',0),      TEAL,   'Annual review sufficient'),
    ]
    for i, (level, count, color, action) in enumerate(risk_items):
        ty = 1.92 + i * 1.3
        rect(sl, 7.0, ty, 6.0, 1.15, fill=RGBColor(0x0F,0x0F,0x22))
        rect(sl, 7.0, ty, 0.12, 1.15, fill=color)
        txt(sl, level, 7.22, ty+0.1, 1.8, 0.4, sz=11, bold=True, color=color)
        txt(sl, str(count), 7.22, ty+0.48, 0.8, 0.52, sz=28, bold=True, color=WHITE)
        txt(sl, 'risks', 8.0, ty+0.62, 0.9, 0.35, sz=9, color=GRAY)
        txt(sl, action, 9.2, ty+0.38, 3.6, 0.38, sz=9, color=GRAY, italic=True)

    pagenum(sl, 9)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 10 — DISTRIBUTOR SCORECARD
# ══════════════════════════════════════════════════════════════════════════════

def slide_10():
    sl = prs.slides.add_slide(BLANK)
    bg(sl, CREAM)
    accent_bar(sl, AMBER)

    txt(sl, 'DISTRIBUTOR SCORECARD', 0.8, 0.28, 12, 0.6,
        sz=30, bold=True, color=NAVY)
    txt(sl, 'Weighted KPI scoring · OTIF · Fill Rate · Damage · TAT · Grade A–D',
        0.8, 0.92, 12, 0.35, sz=11, color=GRAY, italic=True)
    rect(sl, 0.8, 1.32, 12.2, 0.05, fill=AMBER)

    df = pd.read_csv(f'{DATA}/distributor_scorecard.csv')
    gc = df['Grade'].value_counts()

    # Grade donut visual (4 big tiles)
    grade_data = [
        ('A', gc.get('A',0), GREEN,  'Excellent — maintain & reward'),
        ('B', gc.get('B',0), TEAL,   'Good — monitor quarterly'),
        ('C', gc.get('C',0), ORANGE, 'Needs improvement — QBR'),
        ('D', gc.get('D',0), RED,    'Critical — intervention now'),
    ]
    for i, (grade, count, color, label) in enumerate(grade_data):
        lx = 0.8 + i * 3.1
        rect(sl, lx, 1.5, 2.85, 2.0, fill=color)
        txt(sl, f'GRADE {grade}', lx+0.12, 1.58, 2.6, 0.42,
            sz=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        txt(sl, str(count), lx+0.12, 1.95, 2.6, 0.82,
            sz=44, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        txt(sl, 'distributors', lx+0.12, 2.75, 2.6, 0.3,
            sz=9, color=RGBColor(0xEE,0xEE,0xEE), align=PP_ALIGN.CENTER)
        txt(sl, label, lx+0.12, 3.05, 2.6, 0.38,
            sz=8, color=RGBColor(0xDD,0xDD,0xEE), italic=True, align=PP_ALIGN.CENTER)

    # Scoring methodology
    txt(sl, 'HOW WE SCORED', 0.8, 3.6, 12.2, 0.35, sz=12, bold=True, color=NAVY)
    weights = [
        ('OTIF %',      '35%', AMBER),
        ('Fill Rate %', '30%', TEAL),
        ('Damage Rate', '20%', ORANGE),
        ('Avg TAT',     '15%', PURPLE),
    ]
    for i, (kpi, weight, color) in enumerate(weights):
        lx = 0.8 + i * 3.1
        rect(sl, lx, 4.0, 2.85, 0.55, fill=color)
        txt(sl, kpi, lx+0.1, 4.07, 1.7, 0.4, sz=10, bold=True, color=WHITE)
        txt(sl, weight, lx+1.8, 4.07, 0.95, 0.4, sz=14, bold=True, color=WHITE,
            align=PP_ALIGN.RIGHT)

    # Top 3 + Bottom 3
    top3 = df.nlargest(3, 'Performance_Score')[['Distributor_ID','Region','OTIF_Pct','Performance_Score','Grade']]
    bot3 = df.nsmallest(3, 'Performance_Score')[['Distributor_ID','Region','OTIF_Pct','Performance_Score','Grade']]

    for col_x, label, data, label_color in [
        (0.8, 'TOP PERFORMERS', top3, GREEN),
        (7.0, 'NEEDS INTERVENTION', bot3, RED),
    ]:
        txt(sl, label, col_x, 4.68, 5.8, 0.35, sz=11, bold=True, color=label_color)
        hdrs = ['ID', 'Region', 'OTIF%', 'Score', 'Grade']
        cws  = [1.5, 1.8, 0.9, 0.85, 0.75]
        x = col_x
        for h, w in zip(hdrs, cws):
            rect(sl, x, 5.08, w, 0.3, fill=NAVY)
            txt(sl, h, x+0.05, 5.1, w-0.1, 0.26, sz=8, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
            x += w
        grade_fill = {
            'A': RGBColor(0x27,0xAE,0x60), 'B': TEAL,
            'C': ORANGE, 'D': RED
        }
        for ri, row in enumerate(data.itertuples(index=False)):
            ty = 5.42 + ri * 0.52
            bg_c = RGBColor(0xF8,0xF8,0xFF) if ri%2==0 else WHITE
            vals = [row.Distributor_ID, row.Region, f"{row.OTIF_Pct}%",
                    f"{row.Performance_Score:.1f}", row.Grade]
            x = col_x
            for ci, (v, w) in enumerate(zip(vals, cws)):
                fc = grade_fill.get(str(v), bg_c) if ci==4 else bg_c
                rect(sl, x, ty, w, 0.47, fill=fc)
                txt(sl, str(v), x+0.05, ty+0.07, w-0.1, 0.33,
                    sz=8.5, bold=(ci==4), color=WHITE if ci==4 else DARK,
                    align=PP_ALIGN.CENTER)
                x += w

    pagenum(sl, 10)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 11 — KEY FINDINGS
# ══════════════════════════════════════════════════════════════════════════════

def slide_11():
    sl = prs.slides.add_slide(BLANK)
    bg(sl, NAVY)
    accent_bar(sl, AMBER)

    txt(sl, 'KEY FINDINGS', 0.8, 0.28, 12, 0.6, sz=34, bold=True, color=WHITE)
    rect(sl, 0.8, 0.98, 12.2, 0.05, fill=AMBER)

    findings = [
        (RED,    'DEMAND',
         'Demand spikes 30–70% in festive season but inventory doesn\'t scale. '
         'Stockouts happen at the highest-revenue moments.'),
        (ORANGE, 'LOGISTICS',
         '28% of vehicle capacity wasted per trip. East India lead time is 2.2x '
         'the national average. Rail freight 44% cheaper but only 15% utilized.'),
        (AMBER,  'INVENTORY',
         'Stout 650ml over-ordered 45% above EOQ. East India holds 30% below '
         'recommended safety stock. INR 8–12L in avoidable holding costs.'),
        (RED,    'COLD CHAIN',
         '18% temperature breach rate against <5% target. Retail outlet stage '
         'accounts for 30%+ of all breaches. Total cost impact: INR 3.5L+.'),
        (GREEN,  'OPPORTUNITY',
         'INR 55–80 Lakhs in annual savings identified across freight optimization, '
         'inventory right-sizing, and distributor performance improvement.'),
    ]

    for i, (color, tag, body) in enumerate(findings):
        ty = 1.12 + i * 1.22
        rect(sl, 0.8, ty, 12.2, 1.08, fill=RGBColor(0x0F,0x0F,0x22))
        rect(sl, 0.8, ty, 0.12, 1.08, fill=color)
        txt(sl, tag, 1.1, ty+0.08, 1.8, 0.38, sz=10, bold=True, color=color)
        txt(sl, body, 1.1, ty+0.48, 11.7, 0.55, sz=10, color=WHITE, italic=False)

    pagenum(sl, 11)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 12 — RECOMMENDATIONS
# ══════════════════════════════════════════════════════════════════════════════

def slide_12():
    sl = prs.slides.add_slide(BLANK)
    bg(sl, CREAM)
    accent_bar(sl, AMBER)

    txt(sl, 'STRATEGIC RECOMMENDATIONS', 0.8, 0.28, 12, 0.6,
        sz=27, bold=True, color=NAVY)
    txt(sl, '3 high-impact actions · 90-day execution roadmap · INR 55–80L savings potential',
        0.8, 0.92, 12, 0.35, sz=11, color=GRAY, italic=True)
    rect(sl, 0.8, 1.32, 12.2, 0.05, fill=AMBER)

    recs = [
        (AMBER, '01',
         'Dynamic Safety Stock + Seasonal Buffer',
         'Build 25–30% pre-season inventory for top SKUs by September each year. '
         'Implement EOQ-guided ordering. Pilot VMI with top 3 suppliers.',
         'INR 15–20L / yr  |  +8–10% OTIF'),
        (LTNAVY, '02',
         'East India Logistics Hub Upgrade',
         'Kolkata consolidation hub + long-term rail freight contracts. '
         'GPS tracking for all East India routes. Lead time target: 3.5 days.',
         'INR 25–35L / yr  |  -2.7 days lead time'),
        (GREEN, '03',
         'Distributor Performance Improvement Program',
         'Quarterly business reviews for Grade C & D distributors. '
         'Tiered incentive scheme for Grade A. Digital order management tools rollout.',
         'INR 30–40L / yr  |  OTIF: 76% → 88%'),
    ]

    for i, (color, num, title, body, impact) in enumerate(recs):
        ty = 1.5 + i * 1.95
        rect(sl, 0.8, ty, 12.2, 1.78, fill=WHITE)
        rect(sl, 0.8, ty, 12.2, 0.08, fill=color)
        # Number badge
        oval(sl, 0.88, ty+0.22, 0.7, 0.7, fill=color)
        txt(sl, num, 0.88, ty+0.31, 0.7, 0.52, sz=14, bold=True, color=WHITE,
            align=PP_ALIGN.CENTER)
        # Title
        txt(sl, title, 1.75, ty+0.18, 7.5, 0.45, sz=13, bold=True, color=NAVY)
        # Body
        txt(sl, body, 1.75, ty+0.65, 7.5, 0.72, sz=10, color=GRAY)
        # Impact pill
        rect(sl, 9.5, ty+0.28, 3.3, 0.55, fill=color)
        txt(sl, impact, 9.58, ty+0.34, 3.15, 0.43,
            sz=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    pagenum(sl, 12)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 13 — CLOSING
# ══════════════════════════════════════════════════════════════════════════════

def slide_13():
    sl = prs.slides.add_slide(BLANK)
    bg(sl, NAVY)

    # Amber diagonal block
    s = sl.shapes.add_shape(MSO_PLGN, Inches(7.5), Inches(0), Inches(5.83), Inches(7.5))
    s.fill.solid(); s.fill.fore_color.rgb = RGBColor(0x10,0x10,0x22)
    s.line.fill.background()

    rect(sl, 0, 0, 0.5, 7.5, fill=AMBER)
    rect(sl, 0.5, 7.15, 12.83, 0.12, fill=AMBER)

    txt(sl, 'THANK', 0.8, 0.7, 9, 1.3, sz=80, bold=True, color=WHITE)
    txt(sl, 'YOU', 0.8, 1.9, 9, 1.3, sz=80, bold=True, color=AMBER)

    rect(sl, 0.8, 3.3, 8.0, 0.06, fill=RGBColor(0x44,0x44,0x66))

    txt(sl, 'Harsh Raj Pandey', 0.8, 3.5, 8.0, 0.65,
        sz=24, bold=True, color=WHITE)
    txt(sl, 'Supply Chain Analyst  ·  FMCG & Beverage Domain', 0.8, 4.18, 8.0, 0.4,
        sz=11, color=GRAY, italic=True)

    # Contact block
    contacts = [
        ('GitHub  →', 'github.com/Harsh258-collab'),
        ('LinkedIn →', 'linkedin.com/in/harsh-raj-pandey-1a0319325'),
        ('Email    →', 'harsh258.collab@gmail.com'),
    ]
    for i, (label, val) in enumerate(contacts):
        ty = 4.82 + i * 0.58
        txt(sl, label, 0.8, ty, 1.8, 0.45, sz=10, bold=True, color=AMBER)
        txt(sl, val,   2.65, ty, 6.0, 0.45, sz=10, color=WHITE)

    # Beer mug small (right side decoration)
    rect(sl, 9.8, 1.8, 1.4, 2.2, fill=BEERCLR)
    oval(sl, 9.75, 1.55, 1.5, 0.45, fill=FOAMY)
    oval(sl, 9.85, 1.4, 0.3, 0.3, fill=WHITE)
    oval(sl, 10.2, 1.32, 0.25, 0.25, fill=WHITE)
    s = sl.shapes.add_shape(MSO_OVAL, Inches(11.18), Inches(2.1), Inches(0.5), Inches(1.0))
    s.fill.background(); s.line.color.rgb = BEERCLR; s.line.width = Pt(6)
    rect(sl, 10.05, 3.8, 1.4, 0.18, fill=DARKAMB)

    txt(sl, '"Good supply chains don\'t just move products — they build trust."',
        0.8, 6.65, 12.0, 0.5,
        sz=11, color=LTAMBER, italic=True, align=PP_ALIGN.LEFT)

    pagenum(sl, 13)


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == '__main__':
    print("[...] Building BeerFlow PRO Presentation...")
    funcs = [
        (slide_01, "Cover — Dramatic opening"),
        (slide_02, "Logistics pathway journey"),
        (slide_03, "CHEERS! Beer mugs collide"),
        (slide_04, "The Problem — bold statement"),
        (slide_05, "Project scope — big numbers"),
        (slide_06, "Demand Forecasting"),
        (slide_07, "Inventory EOQ"),
        (slide_08, "Logistics Network"),
        (slide_09, "Cold Chain & Risk"),
        (slide_10, "Distributor Scorecard"),
        (slide_11, "Key Findings"),
        (slide_12, "Recommendations"),
        (slide_13, "Closing — Thank You"),
    ]
    for fn, label in funcs:
        fn()
        print(f"   [OK] Slide {funcs.index((fn,label))+1}: {label}")

    # Set metadata
    prs.core_properties.author = 'Harsh Raj Pandey'
    prs.core_properties.last_modified_by = 'Harsh Raj Pandey'
    prs.core_properties.title = 'BeerFlow - Beverage Supply Chain Optimization'
    prs.core_properties.subject = 'Supply Chain Analytics | FMCG | Beverage'
    prs.core_properties.created  = datetime.datetime(2026, 9, 10, 9, 0, 0)
    prs.core_properties.modified = datetime.datetime(2026, 9, 20, 20, 30, 0)

    prs.save(OUT)
    print(f"\n[DONE] Pro deck saved:\n  {OUT}")
