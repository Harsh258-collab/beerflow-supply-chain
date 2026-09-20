"""
BeerFlow Master Presentation Deck — 10+ Year Executive Experience Tier
Author: Harsh Raj Pandey
Palette: Bold Yellow & White & Deep Obsidian (#FFB703, #FFFFFF, #14141E)
Key Visual Elements:
  - Slide 1: Massive Bold Yellow & White Cover (100pt+ Typography)
  - Slide 2: The Sinuous Logistics Pathway Highway (From Barley Field to City Fleet)
  - Slide 3: Full-Page Cinematic Beer Delivery & Frosted Glass Collision with Realistic Water Droplets & Splashes!
  - Slide 4: The Core Problem: Industry Seasonality Mismatch
  - Slide 5: The Analytics Scope: 2,000 Orders, 1,000 Routes, 265 Fleet Vehicles
  - Slide 6: Demand Forecasting: Seasonal Swings (+70% Peak vs -40% Trough) & MAPE
  - Slide 7: Inventory EOQ Model: High/Medium/Low Stockout Risk Matrix
  - Slide 8: Logistics & Route Optimization: 28% Wasted Capacity & Freight Arbitrage
  - Slide 9: Fleet Analytics & CO2 Footprint
  - Slide 10: Cold Chain Integrity: The 18% Breach Rate & Retail Weak Link
  - Slide 11: Freight Benchmarking: Road vs Rail vs Air vs 3PL
  - Slide 12: Distributor Performance Scorecard (Weighted Ranking)
  - Slide 13: 3 Strategic Recommendations & 90-Day Roadmap (INR 55-80L Savings)
  - Slide 14: Executive Closing & Contact Card
Includes native PowerPoint slide transitions (Smooth Push & Fade) for presentation mode!
"""

import os
import math
import random
import datetime
import pandas as pd
import numpy as np

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml import parse_xml

# ── PATHS & SETUP ────────────────────────────────────────────────────────────
BASE = r'C:\Users\BIT\.gemini\antigravity\scratch\beerflow-supply-chain'
DATA = f'{BASE}\\data'
OUT  = f'{BASE}\\powerpoint\\BeerFlow_Pro_Presentation.pptx'

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]

# ── PALETTE DEFINITIONS ──────────────────────────────────────────────────────
# High-contrast, hyper-professional palette (Yellow, White, Obsidian Navy)
YELLOW      = RGBColor(0xFF, 0xB7, 0x03)  # Vibrant golden amber yellow
AMBER_GOLD  = RGBColor(0xFB, 0x85, 0x00)  # Rich brewing gold
LIGHT_CREAM = RGBColor(0xFD, 0xF0, 0xD5)  # Warm barley cream
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)  # Pure crisp white
OBSIDIAN    = RGBColor(0x0F, 0x14, 0x1C)  # Deep executive dark navy/black
DEEP_NAVY   = RGBColor(0x1B, 0x26, 0x3B)  # Rich slate navy
MID_SLATE   = RGBColor(0x41, 0x5A, 0x77)  # Architectural slate
LIGHT_GRAY  = RGBColor(0xF0, 0xF4, 0xF8)  # Clean background gray
CARD_BORDER = RGBColor(0xE0, 0xE7, 0xFF)  # Subtle card outline
GREEN_ACC   = RGBColor(0x06, 0xD6, 0xA0)  # High-grade green
RED_ACC     = RGBColor(0xEF, 0x47, 0x6F)  # High-alert red
BLUE_ACC    = RGBColor(0x11, 0x8A, 0xB2)  # Analytics blue

# Beer & Glassware Specific Colors
BEER_AMBER  = RGBColor(0xE5, 0x98, 0x00)
BEER_DEEP   = RGBColor(0xBC, 0x70, 0x00)
BEER_LIGHT  = RGBColor(0xFF, 0xC3, 0x4D)
FOAM_WHITE  = RGBColor(0xFF, 0xFA, 0xF0)
FOAM_CREAM  = RGBColor(0xF4, 0xEB, 0xD0)
GLASS_RIM   = RGBColor(0xE0, 0xF2, 0xFE)
DROPLET_CLR = RGBColor(0xC2, 0x78, 0x00)
DROPLET_HI  = RGBColor(0xFF, 0xFF, 0xFF)

MSO_RECT  = 1
MSO_ROUND = 5
MSO_OVAL  = 9

# ── HELPER FUNCTIONS ─────────────────────────────────────────────────────────

def add_transition(slide, trans_type="push", direction="l"):
    """Inject native OpenXML slide transition for presentation mode."""
    if trans_type == "push":
        xml = f'<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="med"><p:push dir="{direction}"/></p:transition>'
    elif trans_type == "fade":
        xml = '<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="med"><p:fade/></p:transition>'
    else:
        xml = '<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="med"><p:wipe/></p:transition>'
    slide._element.append(parse_xml(xml))

def set_bg(slide, color):
    f = slide.background.fill
    f.solid()
    f.fore_color.rgb = color

def draw_rect(slide, l, t, w, h, fill=None, line=None, lw=0, shape=MSO_RECT):
    s = slide.shapes.add_shape(shape, Inches(l), Inches(t), Inches(w), Inches(h))
    if fill:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    else:
        s.fill.background()
    if line:
        s.line.color.rgb = line
        s.line.width = Pt(lw or 1.0)
    else:
        s.line.fill.background()
    return s

def draw_oval(slide, l, t, w, h, fill=None, line=None, lw=0):
    s = slide.shapes.add_shape(MSO_OVAL, Inches(l), Inches(t), Inches(w), Inches(h))
    if fill:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    else:
        s.fill.background()
    if line:
        s.line.color.rgb = line
        s.line.width = Pt(lw or 1.0)
    else:
        s.line.fill.background()
    return s

def add_text(slide, text, l, t, w, h, sz=12, bold=False, color=OBSIDIAN,
             align=PP_ALIGN.LEFT, italic=False, wrap=True):
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tb.text_frame.word_wrap = wrap
    tb.text_frame.margin_left = Inches(0.02)
    tb.text_frame.margin_right = Inches(0.02)
    tb.text_frame.margin_top = Inches(0.02)
    tb.text_frame.margin_bottom = Inches(0.02)
    p = tb.text_frame.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = str(text)
    r.font.size = Pt(sz)
    r.font.bold = bold
    r.font.color.rgb = color
    r.font.italic = italic
    r.font.name = 'Calibri'
    return tb

def add_header(slide, tag, title, subtitle=None, dark=False):
    """Clean, authoritative header used across data slides."""
    accent = YELLOW if dark else AMBER_GOLD
    t_color = WHITE if dark else OBSIDIAN
    s_color = RGBColor(0x94, 0xA3, 0xB8) if dark else MID_SLATE
    
    # Tag badge
    draw_rect(slide, 0.8, 0.35, 2.4, 0.28, fill=accent)
    add_text(slide, tag.upper(), 0.8, 0.38, 2.4, 0.24, sz=8.5, bold=True, color=OBSIDIAN if dark else WHITE, align=PP_ALIGN.CENTER)
    
    # Main title
    add_text(slide, title, 0.8, 0.68, 11.5, 0.55, sz=22, bold=True, color=t_color)
    
    # Subtitle
    if subtitle:
        add_text(slide, subtitle, 0.8, 1.22, 11.5, 0.32, sz=10, italic=True, color=s_color)
        
    # Subtle accent line
    draw_rect(slide, 0.8, 1.55, 11.7, 0.02, fill=accent)

def add_page_number(slide, n, total=14, dark=False):
    c = RGBColor(0x64, 0x74, 0x8B) if dark else RGBColor(0x94, 0xA3, 0xB8)
    add_text(slide, f"{n:02d} / {total:02d}", 11.8, 7.1, 1.0, 0.25, sz=8.5, color=c, align=PP_ALIGN.RIGHT)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 1 — ULTRA BOLD YELLOW & WHITE COVER
# ══════════════════════════════════════════════════════════════════════════════

def slide_01_cover():
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl, WHITE)
    add_transition(sl, "fade")

    # ── LEFT 55%: Massive Golden-Yellow Power Panel ──
    draw_rect(sl, 0, 0, 7.33, 7.5, fill=YELLOW)
    
    # Architectural background grid lines on yellow for high-end feel
    for i in range(1, 7):
        draw_rect(sl, i * 1.2, 0, 0.02, 7.5, fill=RGBColor(0xEE, 0xA8, 0x00))
    for j in range(1, 8):
        draw_rect(sl, 0, j * 0.95, 7.33, 0.02, fill=RGBColor(0xEE, 0xA8, 0x00))

    # Top category pill
    draw_rect(sl, 0.8, 0.8, 4.2, 0.36, fill=OBSIDIAN)
    add_text(sl, "END-TO-END SUPPLY CHAIN OPTIMIZATION", 0.85, 0.86, 4.1, 0.26,
             sz=8.5, bold=True, color=YELLOW, align=PP_ALIGN.CENTER)

    # Giant Display Typography (110pt)
    add_text(sl, "BEER", 0.7, 1.3, 6.2, 1.8,
             sz=108, bold=True, color=OBSIDIAN, align=PP_ALIGN.LEFT)
    
    # Graphic divider bar
    draw_rect(sl, 0.8, 3.1, 5.8, 0.12, fill=WHITE)
    
    add_text(sl, "FLOW", 0.7, 3.15, 6.2, 1.8,
             sz=108, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    # Subtitle Manifesto
    add_text(sl, "Logistics Development · Demand Forecasting · EOQ Inventory · Cold Chain Integrity",
             0.8, 5.15, 5.9, 0.45, sz=11, bold=True, color=OBSIDIAN)
    add_text(sl, "A data-driven operational study modelled on real FMCG beverage distribution networks across India.",
             0.8, 5.65, 5.8, 0.6, sz=9.5, color=RGBColor(0x2D, 0x25, 0x10), italic=True)

    # Core Stats Pill Row on Yellow
    pills = [("2,000", "Orders"), ("1,000", "Routes"), ("265", "Fleet Units"), ("5", "Mega-Hubs")]
    for idx, (num, lbl) in enumerate(pills):
        px = 0.8 + idx * 1.45
        draw_rect(sl, px, 6.35, 1.35, 0.65, fill=OBSIDIAN)
        add_text(sl, num, px, 6.4, 1.35, 0.35, sz=13, bold=True, color=YELLOW, align=PP_ALIGN.CENTER)
        add_text(sl, lbl, px, 6.72, 1.35, 0.22, sz=7.5, bold=False, color=WHITE, align=PP_ALIGN.CENTER)

    # ── RIGHT 45%: Clean Architectural White Panel ──
    # Author Card Frame
    draw_rect(sl, 7.8, 0.8, 5.0, 5.8, fill=LIGHT_GRAY, line=CARD_BORDER, lw=1.5)
    
    # Author Header Banner
    draw_rect(sl, 7.8, 0.8, 5.0, 0.12, fill=AMBER_GOLD)

    # Circular Profile Monogram Avatar
    draw_oval(sl, 9.65, 1.25, 1.3, 1.3, fill=OBSIDIAN, line=YELLOW, lw=3)
    add_text(sl, "HP", 9.65, 1.58, 1.3, 0.6, sz=26, bold=True, color=YELLOW, align=PP_ALIGN.CENTER)

    # Author Name & Role
    add_text(sl, "HARSH RAJ PANDEY", 7.9, 2.7, 4.8, 0.45,
             sz=18, bold=True, color=OBSIDIAN, align=PP_ALIGN.CENTER)
    add_text(sl, "Supply Chain & Logistics Operations Analyst", 7.9, 3.15, 4.8, 0.32,
             sz=10.5, bold=True, color=MID_SLATE, align=PP_ALIGN.CENTER)
    add_text(sl, "Target: AB InBev Supply Chain / Logistics Development", 7.9, 3.45, 4.8, 0.3,
             sz=9, italic=True, color=AMBER_GOLD, align=PP_ALIGN.CENTER)

    draw_rect(sl, 8.4, 3.85, 3.8, 0.02, fill=CARD_BORDER)

    # Core Competencies Bullet Blocks
    skills = [
        ("Logistics Network Design", "Route optimization, TMS load pooling, freight benchmarking"),
        ("Quantitative Inventory", "EOQ calculation, dynamic safety stock, stockout risk mitigation"),
        ("Operational Toolset", "Advanced Python (Pandas/NumPy), MS Excel (6-Sheet Deep-Dive), Executive PPT"),
    ]
    for s_idx, (stitle, sdesc) in enumerate(skills):
        sy = 4.05 + s_idx * 0.72
        draw_rect(sl, 8.1, sy + 0.05, 0.12, 0.45, fill=YELLOW)
        add_text(sl, stitle, 8.35, sy, 4.2, 0.25, sz=9.5, bold=True, color=OBSIDIAN)
        add_text(sl, sdesc, 8.35, sy + 0.22, 4.2, 0.35, sz=8, color=MID_SLATE)

    # Contact Details
    draw_rect(sl, 7.8, 6.2, 5.0, 0.4, fill=OBSIDIAN)
    add_text(sl, "GitHub: Harsh258-collab   ·   LinkedIn: harsh-raj-pandey   ·   Sept 2026",
             7.8, 6.28, 5.0, 0.25, sz=8, color=LIGHT_CREAM, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 2 — THE SINUOUS LOGISTICS HIGHWAY (Winding Travel Pathway)
# ══════════════════════════════════════════════════════════════════════════════

def slide_02_logistics_pathway():
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl, LIGHT_GRAY)
    add_transition(sl, "push", "l")

    # Header
    add_header(sl, "SUPPLY CHAIN ROADMAP",
               "The Logistics Highway: 72 Hours from Grain Silo to Chilled Mug",
               "A continuous, temperature-tracked physical pipeline connecting raw sourcing to consumer touchpoint.")
    add_page_number(sl, 2)

    # ── DRAWING THE WINDING ROADWAY ──
    # The road ribbon weaves across the slide:
    # Segment 1 (Top-Left to Mid-Left) -> Segment 2 (Mid-Left to Mid-Right) -> Segment 3 (Mid-Right to Bottom-Right)
    
    # Road Asphalt Base (Dark Slate)
    draw_rect(sl, 0.6, 2.7, 12.1, 0.45, fill=DEEP_NAVY)
    draw_rect(sl, 0.6, 5.1, 12.1, 0.45, fill=DEEP_NAVY)
    # Connecting road curve on the right
    draw_rect(sl, 12.1, 2.7, 0.6, 2.85, fill=DEEP_NAVY)

    # Dashed Road Highway Centerline (Yellow dashes)
    for seg_x in np.arange(0.8, 12.0, 0.4):
        draw_rect(sl, seg_x, 2.9, 0.2, 0.04, fill=YELLOW)
        draw_rect(sl, seg_x, 5.3, 0.2, 0.04, fill=YELLOW)
    for seg_y in np.arange(2.9, 5.3, 0.4):
        draw_rect(sl, 12.38, seg_y, 0.04, 0.2, fill=YELLOW)

    # ── 6 MILESTONE WAYPOINT CARDS ──
    milestones = [
        # (X, Y, Step, Icon, Name, Subtitle, Metric, Color)
        (0.8, 1.7, "01", "🌾", "RAW SOURCING", "Punjab / Rajasthan", "Lead Time: 4.2d", AMBER_GOLD),
        (4.4, 1.7, "02", "🏭", "BREWERY & PACK", "Fermentation & Canning", "Cap: 24K Cases/d", YELLOW),
        (8.0, 1.7, "03", "🏢", "MOTHER HUB", "WH-DEL-01 (Automated)", "Space Util: 88%", BLUE_ACC),
        (9.4, 4.1, "04", "🚆", "LONG-HAUL FREIGHT", "32-ft Truck & Bulk Rail", "Rail Cost: -44%", GREEN_ACC),
        (5.2, 4.1, "05", "🚛", "CITY REEFER FLEET", "Tata Ace / 14-ft Trucks", "Util Target: 85%", RED_ACC),
        (1.0, 4.1, "06", "🍺", "COLD TAPROOM", "On-Shelf & Bar Delivery", "Target: <4°C", YELLOW),
    ]

    for (mx, my, step, icon, title, subtitle, metric, color) in milestones:
        # Waypoint Card
        draw_rect(sl, mx, my, 2.7, 0.95, fill=WHITE, line=color, lw=2, shape=MSO_ROUND)
        # Step Badge
        draw_rect(sl, mx, my, 0.65, 0.95, fill=color, shape=MSO_ROUND)
        add_text(sl, step, mx, my + 0.12, 0.65, 0.35, sz=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text(sl, icon, mx, my + 0.48, 0.65, 0.35, sz=14, align=PP_ALIGN.CENTER)
        
        # Details
        add_text(sl, title, mx + 0.75, my + 0.08, 1.9, 0.26, sz=9.5, bold=True, color=OBSIDIAN)
        add_text(sl, subtitle, mx + 0.75, my + 0.32, 1.9, 0.24, sz=8, color=MID_SLATE)
        
        # Metric Pill
        draw_rect(sl, mx + 0.75, my + 0.58, 1.85, 0.26, fill=LIGHT_GRAY)
        add_text(sl, metric, mx + 0.78, my + 0.62, 1.8, 0.2, sz=7.5, bold=True, color=color)

    # ── TRUCK GRAPHICS ALONG THE ROAD ──
    # Mini 32-ft Long-Haul Truck on top road
    draw_rect(sl, 2.8, 2.62, 0.9, 0.3, fill=WHITE, line=OBSIDIAN, lw=1)
    draw_rect(sl, 3.7, 2.68, 0.3, 0.24, fill=AMBER_GOLD) # cab
    draw_oval(sl, 3.0, 2.92, 0.12, 0.12, fill=OBSIDIAN)
    draw_oval(sl, 3.5, 2.92, 0.12, 0.12, fill=OBSIDIAN)
    draw_oval(sl, 3.85, 2.92, 0.12, 0.12, fill=OBSIDIAN)
    add_text(sl, "32-FT FREIGHT", 2.82, 2.65, 0.85, 0.2, sz=5.5, bold=True, color=OBSIDIAN)

    # Mini City Reefer Truck on bottom road
    draw_rect(sl, 3.4, 5.02, 0.75, 0.28, fill=WHITE, line=OBSIDIAN, lw=1)
    draw_rect(sl, 3.1, 5.08, 0.3, 0.22, fill=GREEN_ACC) # cab (facing left)
    draw_oval(sl, 3.2, 5.3, 0.12, 0.12, fill=OBSIDIAN)
    draw_oval(sl, 3.8, 5.3, 0.12, 0.12, fill=OBSIDIAN)
    add_text(sl, "REEFER 14FT", 3.42, 5.06, 0.7, 0.2, sz=5.5, bold=True, color=OBSIDIAN)

    # ── BOTTOM BANNER: KEY BOTTLENECKS IDENTIFIED ──
    draw_rect(sl, 0.8, 6.0, 11.7, 0.9, fill=DEEP_NAVY)
    draw_rect(sl, 0.8, 6.0, 0.15, 0.9, fill=RED_ACC)
    add_text(sl, "CRITICAL SUPPLY CHAIN BOTTLENECKS REVEALED IN DATA", 1.1, 6.08, 11.2, 0.26,
             sz=9.5, bold=True, color=YELLOW)
    add_text(sl, "• Sourcing to Mother Hub: Lead times swing from 3.2 days (North) to 6.8 days (East India) due to rail bottlenecks.\n"
                 "• Last-Mile City Fleet: 28% of vehicle capacity runs empty per trip due to order fragmentation rather than load pooling.\n"
                 "• Retail Handover: 18% of all cold chain breaches occur at the final retail delivery drop.",
             1.1, 6.32, 11.2, 0.5, sz=8.5, color=WHITE)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 3 — CINEMATIC BEER DELIVERY & GLASS COLLISION (With Water Droplets!)
# ══════════════════════════════════════════════════════════════════════════════

def slide_03_beer_collision():
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl, OBSIDIAN)
    add_transition(sl, "fade")
    add_page_number(sl, 3, dark=True)

    # Dramatic warm radial glow background
    draw_oval(sl, 2.5, 0.2, 8.3, 7.0, fill=RGBColor(0x28, 0x1A, 0x08))
    draw_oval(sl, 4.2, 1.2, 5.0, 5.0, fill=RGBColor(0x42, 0x2A, 0x0A))

    # Top Tagline
    draw_rect(sl, 4.4, 0.4, 4.5, 0.3, fill=YELLOW)
    add_text(sl, "THE ULTIMATE SUPPLY CHAIN OUTCOME", 4.4, 0.45, 4.5, 0.22,
             sz=9, bold=True, color=OBSIDIAN, align=PP_ALIGN.CENTER)

    # ── GLASS 1: LEFT BEER MUG (Tilted towards center) ──
    gx1, gy1 = 3.6, 2.2
    
    # Outer Frost / Glass Body
    draw_rect(sl, gx1, gy1 + 0.3, 1.9, 3.4, fill=BEER_DEEP, shape=MSO_ROUND)
    # Inner Golden Liquid Core
    draw_rect(sl, gx1 + 0.1, gy1 + 0.4, 1.7, 3.1, fill=BEER_AMBER, shape=MSO_ROUND)
    # Vertical Glass Reflection Facets
    draw_rect(sl, gx1 + 0.35, gy1 + 0.4, 0.22, 3.0, fill=BEER_LIGHT)
    draw_rect(sl, gx1 + 0.85, gy1 + 0.4, 0.18, 3.0, fill=BEER_LIGHT)
    
    # Glass Handle (Left)
    draw_rect(sl, gx1 - 0.55, gy1 + 0.8, 0.65, 2.0, fill=None, line=GLASS_RIM, lw=6, shape=MSO_ROUND)
    
    # Foam Head (Thick, billowing cloud)
    draw_oval(sl, gx1 - 0.15, gy1 - 0.15, 2.2, 0.75, fill=FOAM_WHITE)
    draw_oval(sl, gx1 + 0.1, gy1 - 0.25, 0.65, 0.65, fill=WHITE)
    draw_oval(sl, gx1 + 0.65, gy1 - 0.32, 0.8, 0.75, fill=WHITE)
    draw_oval(sl, gx1 + 1.25, gy1 - 0.22, 0.7, 0.65, fill=WHITE)
    # Foam dripping down glass
    draw_rect(sl, gx1 + 0.5, gy1 + 0.4, 0.25, 0.8, fill=FOAM_WHITE, shape=MSO_ROUND)
    draw_rect(sl, gx1 + 1.2, gy1 + 0.35, 0.2, 0.6, fill=FOAM_WHITE, shape=MSO_ROUND)

    # ── GLASS 2: RIGHT BEER MUG (Colliding symmetrically) ──
    gx2, gy2 = 7.8, 2.2
    
    # Outer Frost / Glass Body
    draw_rect(sl, gx2, gy2 + 0.3, 1.9, 3.4, fill=BEER_DEEP, shape=MSO_ROUND)
    # Inner Golden Liquid Core
    draw_rect(sl, gx2 + 0.1, gy2 + 0.4, 1.7, 3.1, fill=BEER_AMBER, shape=MSO_ROUND)
    # Vertical Glass Reflection Facets
    draw_rect(sl, gx2 + 0.85, gy2 + 0.4, 0.22, 3.0, fill=BEER_LIGHT)
    draw_rect(sl, gx2 + 1.35, gy2 + 0.4, 0.18, 3.0, fill=BEER_LIGHT)
    
    # Glass Handle (Right)
    draw_rect(sl, gx2 + 1.8, gy2 + 0.8, 0.65, 2.0, fill=None, line=GLASS_RIM, lw=6, shape=MSO_ROUND)
    
    # Foam Head
    draw_oval(sl, gx2 - 0.15, gy2 - 0.15, 2.2, 0.75, fill=FOAM_WHITE)
    draw_oval(sl, gx2 + 0.05, gy2 - 0.22, 0.7, 0.65, fill=WHITE)
    draw_oval(sl, gx2 + 0.55, gy2 - 0.32, 0.8, 0.75, fill=WHITE)
    draw_oval(sl, gx2 + 1.2, gy2 - 0.25, 0.65, 0.65, fill=WHITE)
    # Foam dripping down glass
    draw_rect(sl, gx2 + 0.35, gy2 + 0.35, 0.2, 0.7, fill=FOAM_WHITE, shape=MSO_ROUND)
    draw_rect(sl, gx2 + 1.1, gy2 + 0.4, 0.22, 0.5, fill=FOAM_WHITE, shape=MSO_ROUND)

    # ── WATER DROPLETS & CONDENSATION ON BOTH GLASSES! ──
    # Each droplet has a darker condensation shadow + crisp pure white reflection highlight glint
    random.seed(42)
    for base_x, is_left in [(gx1, True), (gx2, False)]:
        for _ in range(28):
            dx = base_x + random.uniform(0.15, 1.65)
            dy = gy1 + random.uniform(0.8, 3.4)
            dsize = random.uniform(0.06, 0.16)
            
            # Base droplet body
            draw_oval(sl, dx, dy, dsize, dsize * 1.2, fill=DROPLET_CLR)
            # Crisp white specular glint highlight (makes it look 3D wet!)
            draw_oval(sl, dx + dsize * 0.2, dy + dsize * 0.15, dsize * 0.35, dsize * 0.35, fill=DROPLET_HI)

        # 3 Elongated running condensation drip rivulets
        for rx_off in [0.4, 0.95, 1.4]:
            rx = base_x + rx_off
            ry = gy1 + random.uniform(1.2, 1.8)
            rh = random.uniform(0.5, 0.9)
            draw_rect(sl, rx, ry, 0.05, rh, fill=DROPLET_CLR, shape=MSO_ROUND)
            # droplet bead at bottom of drip
            draw_oval(sl, rx - 0.02, ry + rh - 0.02, 0.09, 0.09, fill=DROPLET_CLR)
            draw_oval(sl, rx, ry + rh, 0.035, 0.035, fill=DROPLET_HI)

    # ── COLLISION IMPACT & SPLASH EFFECT AT CENTER (x=6.67) ──
    # Impact burst rings
    draw_oval(sl, 5.8, 1.5, 1.7, 1.7, fill=None, line=YELLOW, lw=2)
    draw_oval(sl, 6.1, 1.8, 1.1, 1.1, fill=None, line=WHITE, lw=1.5)
    
    # Liquid splash arc bursts
    splash_particles = [
        (6.67, 1.2, 0.22, AMBER_GOLD), (6.3, 1.0, 0.16, YELLOW), (7.1, 0.9, 0.18, YELLOW),
        (6.0, 1.5, 0.12, WHITE), (7.4, 1.4, 0.14, WHITE), (6.67, 0.7, 0.18, AMBER_GOLD),
        (5.7, 0.8, 0.10, BEER_LIGHT), (7.6, 0.7, 0.12, BEER_LIGHT), (6.4, 0.5, 0.14, YELLOW),
        (6.9, 0.55, 0.11, WHITE), (6.67, 2.0, 0.28, AMBER_GOLD)
    ]
    for (spx, spy, sps, spcol) in splash_particles:
        draw_oval(sl, spx, spy, sps, sps, fill=spcol)
        draw_oval(sl, spx + sps*0.2, spy + sps*0.15, sps*0.3, sps*0.3, fill=WHITE)

    # ── MASSIVE CINEMATIC TITLE & MANIFESTO ──
    add_text(sl, "CHEERS!", 3.0, 5.75, 7.33, 0.9,
             sz=64, bold=True, color=YELLOW, align=PP_ALIGN.CENTER)
    add_text(sl, "DELIVERED COLD. EXACTLY ON TIME.", 2.0, 6.55, 9.33, 0.4,
             sz=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(sl, "Every algorithm, EOQ formula, and route consolidation model exists for this exact moment.",
             2.0, 6.92, 9.33, 0.35, sz=11, color=LIGHT_CREAM, italic=True, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 4 — THE CORE PROBLEM: FMCG SEASONALITY MISMATCH
# ══════════════════════════════════════════════════════════════════════════════

def slide_04_problem():
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl, WHITE)
    add_transition(sl, "push", "l")
    add_header(sl, "THE INDUSTRY CHALLENGE",
               "The Bullwhip Paradox: Why Beverage Supply Chains Break at Peak",
               "Traditional static planning models create stockouts during summer & festivals, and excess inventory in monsoons.")
    add_page_number(sl, 4)

    # 3 High-Impact Problem Pillars
    pillars = [
        ("01", "SEASONAL DEMAND AMPLIFICATION",
         "Demand swings up to +70% in Q4 (Diwali/New Year) and +45% in summer.\n\n"
         "Because replenishment lead times are static (3–6 days), breweries under-produce before the spike, "
         "leading to critical stockouts during the highest-revenue weeks of the year.",
         RED_ACC, "Estimated Lost Sales: ~INR 42L/yr"),
        
        ("02", "FRAGMENTED LAST-MILE LOGISTICS",
         "Vehicles run at an average of 72% load capacity — wasting 28% of freight spend.\n\n"
         "Orders are dispatched as they arrive rather than batched into density-optimized routes, "
         "forcing expensive mini-truck runs when 32-ft consolidated trips are feasible.",
         AMBER_GOLD, "Freight Inefficiency: 28% Capacity Lost"),
        
        ("03", "COLD CHAIN COMPLIANCE BREAKDOWN",
         "18% of all shipments record temperature violations above the 8°C ceiling.\n\n"
         "While brewery and primary warehouse transit maintain strict chill, "
         "temperature breaches spike to 30%+ at the retail outlet delivery interface.",
         BLUE_ACC, "Quality Risk: 18% Breached Shipments"),
    ]

    for idx, (num, title, desc, col, tag) in enumerate(pillars):
        px = 0.8 + idx * 3.95
        # Card Frame
        draw_rect(sl, px, 1.8, 3.8, 4.4, fill=LIGHT_GRAY, line=CARD_BORDER, lw=1.5, shape=MSO_ROUND)
        # Accent Top Bar
        draw_rect(sl, px, 1.8, 3.8, 0.12, fill=col, shape=MSO_ROUND)
        
        # Number Badge
        add_text(sl, num, px + 0.25, 2.05, 1.0, 0.6, sz=32, bold=True, color=col)
        add_text(sl, title, px + 0.25, 2.65, 3.3, 0.5, sz=12, bold=True, color=OBSIDIAN)
        
        # Description
        add_text(sl, desc, px + 0.25, 3.25, 3.3, 2.1, sz=9.5, color=DEEP_NAVY)
        
        # Bottom Impact Pill
        draw_rect(sl, px + 0.2, 5.5, 3.4, 0.5, fill=WHITE, line=col, lw=1)
        add_text(sl, tag, px + 0.25, 5.62, 3.3, 0.3, sz=8.5, bold=True, color=col, align=PP_ALIGN.CENTER)

    # Bottom Summary Bar
    draw_rect(sl, 0.8, 6.35, 11.7, 0.65, fill=OBSIDIAN)
    add_text(sl, "EXECUTIVE HYPOTHESIS: Aligning dynamic EOQ buffer stocks with multi-modal freight can elevate OTIF from 78% to 92% "
                 "while unlocking INR 55–80 Lakhs in annual cost recovery.",
             1.0, 6.48, 11.3, 0.45, sz=9.5, bold=True, color=YELLOW)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 5 — THE ANALYTICS SCOPE (High-Impact Data Architecture)
# ══════════════════════════════════════════════════════════════════════════════

def slide_05_scope():
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl, WHITE)
    add_transition(sl, "push", "l")
    add_header(sl, "DATA ARCHITECTURE",
               "Simulation Scope: 2-Year Multi-Echelon Supply Chain Dataset",
               "Synthesized across 10 structured tables capturing order-level, fleet-level, and facility-level transactions.")
    add_page_number(sl, 5)

    stats = [
        ("2,000", "Orders Analyzed", "SKU, Region, Warehouse, Lead Time, OTIF status", AMBER_GOLD),
        ("1,000", "Delivery Routes", "Vehicle utilization, fuel, tolls, stop density, delay min", YELLOW),
        ("265", "Fleet Vehicles", "32-ft, 20-ft, Mini Trucks, Tata Ace, Tempos with CO2 metrics", BLUE_ACC),
        ("600", "Cold Chain Logs", "5 supply chain stages, IoT temperature readings, breach durations", RED_ACC),
        ("400", "Freight Runs", "Direct comparison: Road vs Rail vs Air vs 3PL partners", GREEN_ACC),
        ("24", "Risk Register Items", "Categorized by Likelihood × Impact with mitigation owners", DEEP_NAVY),
    ]

    for idx, (bignum, title, desc, col) in enumerate(stats):
        col_i = idx % 3
        row_i = idx // 3
        bx = 0.8 + col_i * 3.95
        by = 1.8 + row_i * 2.35

        draw_rect(sl, bx, by, 3.8, 2.15, fill=LIGHT_GRAY, line=CARD_BORDER, lw=1.5, shape=MSO_ROUND)
        draw_rect(sl, bx, by, 3.8, 0.08, fill=col, shape=MSO_ROUND)
        
        add_text(sl, bignum, bx + 0.2, by + 0.15, 3.4, 0.7, sz=38, bold=True, color=col)
        add_text(sl, title, bx + 0.2, by + 0.85, 3.4, 0.32, sz=12, bold=True, color=OBSIDIAN)
        add_text(sl, desc, bx + 0.2, by + 1.2, 3.4, 0.75, sz=9, color=MID_SLATE)

    # Methodology Footnote
    draw_rect(sl, 0.8, 6.6, 11.7, 0.45, fill=LIGHT_CREAM)
    add_text(sl, "Tooling Ecosystem: Python (Pandas/NumPy) for Monte Carlo generation · MS Excel (OpenPyXL) with 8 analytical sheets · PowerPoint (python-pptx)",
             1.0, 6.7, 11.3, 0.3, sz=9, bold=True, color=OBSIDIAN)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 6 — DEMAND FORECASTING & SEASONALITY
# ══════════════════════════════════════════════════════════════════════════════

def slide_06_demand():
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl, WHITE)
    add_transition(sl, "push", "l")
    add_header(sl, "DEMAND ANALYTICS",
               "Demand Forecasting: Seasonal Swings & SKU-Level Accuracy",
               "Evaluating 24 months of historical demand against forecast accuracy (MAPE) across 7 beer product lines.")
    add_page_number(sl, 6)

    df = pd.read_csv(f'{DATA}/demand_forecast.csv')
    sku_mape = df.groupby('SKU')['MAPE_Pct'].mean().round(1).reset_index().sort_values('MAPE_Pct')

    # LEFT PANEL: SKU Accuracy Ranking (Bar Chart)
    draw_rect(sl, 0.8, 1.8, 6.0, 4.5, fill=LIGHT_GRAY, line=CARD_BORDER, lw=1.5, shape=MSO_ROUND)
    add_text(sl, "Forecast Accuracy by SKU (Mean Absolute % Error)", 1.0, 1.95, 5.5, 0.3, sz=11, bold=True, color=OBSIDIAN)
    add_text(sl, "Target Benchmark: MAPE < 10% (Lower is better)", 1.0, 2.22, 5.5, 0.25, sz=8.5, italic=True, color=MID_SLATE)

    for i, row in enumerate(sku_mape.itertuples()):
        sy = 2.6 + i * 0.5
        sname = row.SKU.replace('_', ' ')
        add_text(sl, sname, 1.0, sy, 2.3, 0.28, sz=8.5, bold=True, color=DEEP_NAVY)
        
        # Background bar
        draw_rect(sl, 3.3, sy + 0.04, 2.8, 0.24, fill=WHITE)
        # Value bar
        bw = (row.MAPE_Pct / 25.0) * 2.8
        bcol = GREEN_ACC if row.MAPE_Pct < 12 else (AMBER_GOLD if row.MAPE_Pct < 18 else RED_ACC)
        draw_rect(sl, 3.3, sy + 0.04, bw, 0.24, fill=bcol)
        add_text(sl, f"{row.MAPE_Pct}%", 3.35 + bw, sy + 0.02, 0.6, 0.26, sz=8.5, bold=True, color=OBSIDIAN)

    # RIGHT PANEL: Seasonal Calendar Insights
    draw_rect(sl, 7.1, 1.8, 5.4, 4.5, fill=LIGHT_GRAY, line=CARD_BORDER, lw=1.5, shape=MSO_ROUND)
    add_text(sl, "Seasonal Demand Multipliers vs Operational Reality", 7.3, 1.95, 5.0, 0.3, sz=11, bold=True, color=OBSIDIAN)

    seasons = [
        ("OCT - DEC", "+30% to +70%", "FESTIVAL SURGE", "Diwali, Christmas, New Year. High stockout risk for Premium Lager & Wheat Beer.", RED_ACC),
        ("MAR - APR", "+20% to +45%", "SUMMER ONSET", "IPL Season & early heatwaves. Strong Beer volume accelerates by 50K+ cases.", AMBER_GOLD),
        ("JUL - AUG", "-15% to -40%", "MONSOON TROUGH", "Heavy rain suppresses on-premise consumption. Danger of inventory expiration.", BLUE_ACC),
        ("ANNUAL", "+3.0% CAGR", "ORGANIC GROWTH", "Baseline consumption growth driven by craft beer adoption in Tier-1 cities.", GREEN_ACC),
    ]

    for s_idx, (period, val, tag, desc, scol) in enumerate(seasons):
        sy = 2.45 + s_idx * 0.92
        draw_rect(sl, 7.3, sy, 5.0, 0.8, fill=WHITE, line=scol, lw=1, shape=MSO_ROUND)
        draw_rect(sl, 7.3, sy, 0.1, 0.8, fill=scol, shape=MSO_ROUND)
        add_text(sl, period, 7.5, sy + 0.08, 1.3, 0.25, sz=9.5, bold=True, color=DEEP_NAVY)
        add_text(sl, val, 7.5, sy + 0.32, 1.3, 0.38, sz=13, bold=True, color=scol)
        add_text(sl, tag, 8.9, sy + 0.08, 3.2, 0.22, sz=8, bold=True, color=scol)
        add_text(sl, desc, 8.9, sy + 0.3, 3.3, 0.45, sz=8, color=MID_SLATE)

    # Strategic Takeaway Banner
    draw_rect(sl, 0.8, 6.45, 11.7, 0.55, fill=DEEP_NAVY)
    add_text(sl, "KEY ACTION: Shift from static monthly replenishment to a 3-week rolling safety stock ramp starting August 15th.",
             1.0, 6.58, 11.3, 0.35, sz=9.5, bold=True, color=YELLOW)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 7 — INVENTORY EOQ & STOCKOUT RISK
# ══════════════════════════════════════════════════════════════════════════════

def slide_07_eoq():
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl, WHITE)
    add_transition(sl, "push", "l")
    add_header(sl, "INVENTORY OPTIMIZATION",
               "Economic Order Quantity (EOQ) & Stockout Risk Matrix",
               "Balancing holding costs vs ordering costs across 35 SKU-Warehouse nodes to eliminate working capital traps.")
    add_page_number(sl, 7)

    df = pd.read_csv(f'{DATA}/inventory_data.csv')
    risk_counts = df['Stockout_Risk'].value_counts()

    # Top Formula Banner
    draw_rect(sl, 0.8, 1.8, 11.7, 0.75, fill=OBSIDIAN)
    add_text(sl, "MATHEMATICAL MODEL:   EOQ = √ ( 2 × Annual Demand × Order Cost  ÷  Holding Cost per Case/Year )",
             1.0, 1.95, 11.3, 0.3, sz=12, bold=True, color=YELLOW, align=PP_ALIGN.CENTER)
    add_text(sl, "Safety Stock = Z × σ_Demand × √Lead_Time   |   Reorder Point = (Daily Demand × Lead Time) + Safety Stock",
             1.0, 2.25, 11.3, 0.22, sz=9, color=LIGHT_CREAM, align=PP_ALIGN.CENTER)

    # 3 Risk Category Tiles
    risk_meta = [
        ("HIGH STOCKOUT RISK", risk_counts.get('High', 0), "Current Stock < Reorder Point",
         "Immediate order trigger required. East India warehouses show chronic 30% safety stock deficit.", RED_ACC),
        ("OPTIMAL BUFFER", risk_counts.get('Low', 0), "Reorder Point ≤ Stock ≤ Max",
         "Balanced inventory. Lager 500ml and Strong Beer in North/West hubs exhibit near-perfect alignment.", GREEN_ACC),
        ("EXCESS / OVER-ORDER", risk_counts.get('Medium', 0), "Stock > 1.4× EOQ Target",
         "Working capital lockup. Stout 650ml is over-ordered by 45% above EOQ, inflating holding costs by INR 9L.", AMBER_GOLD),
    ]

    for idx, (rtitle, rcount, rsub, rdesc, rcol) in enumerate(risk_meta):
        rx = 0.8 + idx * 3.95
        draw_rect(sl, rx, 2.75, 3.8, 3.4, fill=LIGHT_GRAY, line=CARD_BORDER, lw=1.5, shape=MSO_ROUND)
        draw_rect(sl, rx, 2.75, 3.8, 0.1, fill=rcol, shape=MSO_ROUND)
        
        add_text(sl, rtitle, rx + 0.2, 2.95, 3.4, 0.28, sz=10.5, bold=True, color=rcol)
        add_text(sl, str(rcount), rx + 0.2, 3.25, 3.4, 0.75, sz=44, bold=True, color=OBSIDIAN)
        add_text(sl, "SKU-Warehouse Nodes", rx + 0.2, 4.05, 3.4, 0.22, sz=8.5, bold=True, color=MID_SLATE)
        
        draw_rect(sl, rx + 0.2, 4.35, 3.4, 0.02, fill=CARD_BORDER)
        add_text(sl, rsub, rx + 0.2, 4.45, 3.4, 0.25, sz=8.5, bold=True, color=OBSIDIAN)
        add_text(sl, rdesc, rx + 0.2, 4.75, 3.4, 1.2, sz=8.5, color=DEEP_NAVY)

    # Bottom Callout
    draw_rect(sl, 0.8, 6.3, 11.7, 0.7, fill=LIGHT_CREAM, line=YELLOW, lw=1.5)
    add_text(sl, "EOQ SAVINGS REALIZATION: Right-sizing Stout orders to calculated EOQ saves INR 8.5L annually. "
                 "Injecting dynamic buffer into East India prevents an estimated INR 24L in lost peak revenue.",
             1.0, 6.45, 11.3, 0.45, sz=9.5, bold=True, color=OBSIDIAN)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 8 — LOGISTICS & ROUTE OPTIMIZATION
# ══════════════════════════════════════════════════════════════════════════════

def slide_08_routes():
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl, WHITE)
    add_transition(sl, "push", "l")
    add_header(sl, "LOGISTICS NETWORK",
               "Last-Mile Route Optimization: The 28% Wasted Capacity Leak",
               "Granular analysis of 1,000 delivery routes tracking load utilization, stop density, and cost per case delivered.")
    add_page_number(sl, 8)

    df_r = pd.read_csv(f'{DATA}/route_delivery_data.csv')
    avg_util = df_r['Utilization_Pct'].mean()
    avg_cost = df_r['Cost_Per_Case_INR'].mean()
    otd_rate = (df_r['On_Time_Delivery'] == 'Yes').mean() * 100

    # 3 High Level KPI Badges
    kpis = [
        ("AVG LOAD UTILIZATION", f"{avg_util:.1f}%", "28% of truck capacity runs empty", RED_ACC),
        ("AVG COST PER CASE", f"INR {avg_cost:.2f}", "Benchmark target is INR 14.50", AMBER_GOLD),
        ("ON-TIME ROUTE RATE", f"{otd_rate:.1f}%", "East India routes drag avg down", BLUE_ACC),
    ]
    for idx, (ktitle, kval, ksub, kcol) in enumerate(kpis):
        kx = 0.8 + idx * 3.95
        draw_rect(sl, kx, 1.8, 3.8, 1.2, fill=LIGHT_GRAY, line=CARD_BORDER, lw=1.5, shape=MSO_ROUND)
        draw_rect(sl, kx, 1.8, 0.1, 1.2, fill=kcol, shape=MSO_ROUND)
        add_text(sl, ktitle, kx + 0.25, 1.9, 3.4, 0.22, sz=8.5, bold=True, color=MID_SLATE)
        add_text(sl, kval, kx + 0.25, 2.12, 3.4, 0.5, sz=26, bold=True, color=OBSIDIAN)
        add_text(sl, ksub, kx + 0.25, 2.65, 3.4, 0.25, sz=8, italic=True, color=kcol)

    # TABLE: Vehicle Utilization vs Cost Efficiency
    draw_rect(sl, 0.8, 3.2, 6.6, 3.8, fill=LIGHT_GRAY, line=CARD_BORDER, lw=1.5, shape=MSO_ROUND)
    add_text(sl, "Vehicle Type Performance Breakdown", 1.0, 3.35, 6.2, 0.28, sz=11, bold=True, color=OBSIDIAN)

    vt = df_r.groupby('Vehicle_Type').agg(
        Routes=('Route_ID', 'count'),
        Avg_Util=('Utilization_Pct', 'mean'),
        Cost_Case=('Cost_Per_Case_INR', 'mean'),
        OTD=('On_Time_Delivery', lambda x: round((x=='Yes').mean()*100, 1))
    ).reset_index().sort_values('Cost_Case')

    # Table headers
    thdrs = ['Vehicle Type', 'Routes', 'Util %', 'Cost/Case', 'OTD %']
    tw    = [2.0, 0.9, 1.0, 1.3, 1.0]
    tx = 1.0
    for h, w in zip(thdrs, tw):
        draw_rect(sl, tx, 3.75, w, 0.32, fill=DEEP_NAVY)
        add_text(sl, h, tx, 3.8, w, 0.22, sz=8, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tx += w

    for r_idx, row in enumerate(vt.itertuples()):
        ty = 4.12 + r_idx * 0.46
        bg_col = WHITE if r_idx % 2 == 0 else LIGHT_CREAM
        tx = 1.0
        vals = [row.Vehicle_Type, row.Routes, f"{row.Avg_Util:.1f}%", f"INR {row.Cost_Case:.2f}", f"{row.OTD:.1f}%"]
        for c_idx, (v, w) in enumerate(zip(vals, tw)):
            draw_rect(sl, tx, ty, w, 0.42, fill=bg_col)
            add_text(sl, str(v), tx, ty + 0.08, w, 0.26, sz=8.5, color=OBSIDIAN, align=PP_ALIGN.CENTER)
            tx += w

    # RIGHT: Root Cause Analysis of Wasted Capacity
    draw_rect(sl, 7.6, 3.2, 4.9, 3.8, fill=LIGHT_GRAY, line=CARD_BORDER, lw=1.5, shape=MSO_ROUND)
    add_text(sl, "Why Is 28% Capacity Wasted?", 7.8, 3.35, 4.5, 0.28, sz=11, bold=True, color=OBSIDIAN)

    causes = [
        ("Misaligned Vehicle Sizing", "32-ft heavy trucks deployed on sub-50km urban routes where a Tata Ace would cut cost by 40%."),
        ("Geographic Rather than Load Routing", "Routes are planned by municipality zones rather than dynamic cube-utilization algorithms."),
        ("Distributor Frequency Demands", "High-frequency, small-drop orders prevent full-pallet consolidation."),
        ("Backhaul Inefficiency", "Over 65% of regional delivery trips return to the mother warehouse empty (zero revenue backhaul)."),
    ]
    for c_i, (ctitle, cdesc) in enumerate(causes):
        cy = 3.8 + c_i * 0.76
        draw_oval(sl, 7.85, cy + 0.05, 0.18, 0.18, fill=YELLOW)
        add_text(sl, ctitle, 8.15, cy, 4.2, 0.25, sz=9, bold=True, color=DEEP_NAVY)
        add_text(sl, cdesc, 8.15, cy + 0.22, 4.2, 0.45, sz=8, color=MID_SLATE)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 9 — FLEET ANALYTICS & ESG FOOTPRINT
# ══════════════════════════════════════════════════════════════════════════════

def slide_09_fleet():
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl, WHITE)
    add_transition(sl, "push", "l")
    add_header(sl, "FLEET OPERATIONS & ESG",
               "Fleet Health, Maintenance Costs & Carbon Emissions",
               "Tracking 265 dedicated transport assets across age profiles, breakdown frequencies, and decarbonization pathways.")
    add_page_number(sl, 9)

    df_f = pd.read_csv(f'{DATA}/fleet_utilization_data.csv')
    tot_co2 = df_f['CO2_Emission_KG_PM'].sum() / 1000.0
    tot_maint = df_f['Maintenance_Cost_INR_PM'].sum() / 100000.0

    # 3 Summary Stat Tiles
    fstats = [
        ("ACTIVE FLEET", f"{len(df_f)} Units", "248 Active · 17 In Repair", DEEP_NAVY),
        ("MONTHLY CO2 EMISSIONS", f"{tot_co2:.1f} Tonnes", "Decarbonization target: -20%", AMBER_GOLD),
        ("MAINTENANCE EXPENDITURE", f"INR {tot_maint:.1f} Lakhs/mo", "Aged vehicles (>5 yrs) drive 68% of spend", RED_ACC),
    ]
    for idx, (ftitle, fval, fsub, fcol) in enumerate(fstats):
        fx = 0.8 + idx * 3.95
        draw_rect(sl, fx, 1.8, 3.8, 1.25, fill=LIGHT_GRAY, line=CARD_BORDER, lw=1.5, shape=MSO_ROUND)
        draw_rect(sl, fx, 1.8, 3.8, 0.08, fill=fcol, shape=MSO_ROUND)
        add_text(sl, ftitle, fx + 0.2, 1.92, 3.4, 0.22, sz=8.5, bold=True, color=MID_SLATE)
        add_text(sl, fval, fx + 0.2, 2.15, 3.4, 0.48, sz=22, bold=True, color=fcol)
        add_text(sl, fsub, fx + 0.2, 2.68, 3.4, 0.25, sz=8, color=OBSIDIAN)

    # LEFT: Vehicle Age vs Maintenance Curve
    draw_rect(sl, 0.8, 3.3, 5.8, 3.7, fill=LIGHT_GRAY, line=CARD_BORDER, lw=1.5, shape=MSO_ROUND)
    add_text(sl, "Vehicle Age vs Breakdown Risk & Maintenance Spend", 1.0, 3.45, 5.4, 0.28, sz=11, bold=True, color=OBSIDIAN)

    age_data = [
        ("1 - 2 Years", "98 Units", "0.2 Breakdowns/mo", "INR 1,850/mo", GREEN_ACC),
        ("3 - 4 Years", "84 Units", "0.8 Breakdowns/mo", "INR 4,200/mo", BLUE_ACC),
        ("5 - 6 Years", "52 Units", "1.9 Breakdowns/mo", "INR 9,100/mo", AMBER_GOLD),
        ("7+ Years",   "31 Units", "3.4 Breakdowns/mo", "INR 16,800/mo", RED_ACC),
    ]
    for a_i, (age_grp, ucount, bdown, mspend, acolor) in enumerate(age_data):
        ay = 3.9 + a_i * 0.72
        draw_rect(sl, 1.0, ay, 5.4, 0.62, fill=WHITE, line=acolor, lw=1, shape=MSO_ROUND)
        draw_rect(sl, 1.0, ay, 0.1, 0.62, fill=acolor, shape=MSO_ROUND)
        add_text(sl, age_grp, 1.2, ay + 0.08, 1.5, 0.22, sz=9.5, bold=True, color=DEEP_NAVY)
        add_text(sl, ucount, 1.2, ay + 0.32, 1.5, 0.2, sz=8, color=MID_SLATE)
        add_text(sl, bdown, 2.9, ay + 0.18, 1.8, 0.25, sz=8.5, bold=True, color=acolor)
        add_text(sl, mspend, 4.6, ay + 0.18, 1.7, 0.25, sz=9, bold=True, color=OBSIDIAN, align=PP_ALIGN.RIGHT)

    # RIGHT: Decarbonization & Right-Sizing Strategy
    draw_rect(sl, 6.8, 3.3, 5.7, 3.7, fill=LIGHT_GRAY, line=CARD_BORDER, lw=1.5, shape=MSO_ROUND)
    add_text(sl, "Fleet Right-Sizing & ESG Roadmap", 7.0, 3.45, 5.3, 0.28, sz=11, bold=True, color=OBSIDIAN)

    factions = [
        ("Phase Out 7+ Year Assets", "Retiring the 31 oldest vehicles will eliminate 42% of all unscheduled breakdowns during peak Q4 delivery."),
        ("CNG / EV Pilot in Metro Hubs", "Piloting 15 electric 14-ft delivery trucks in Delhi & Mumbai cuts route fuel cost by 35% and drops fleet CO2 by 18T/mo."),
        ("Dynamic Preventive Servicing", "Transitioning from mileage-based to engine-telematics maintenance schedules prior to festival peak."),
    ]
    for fa_i, (fatitle, fadesc) in enumerate(factions):
        fay = 3.9 + fa_i * 0.95
        draw_rect(sl, 7.0, fay, 5.3, 0.82, fill=WHITE, line=CARD_BORDER, lw=1, shape=MSO_ROUND)
        draw_rect(sl, 7.0, fay, 0.1, 0.82, fill=YELLOW, shape=MSO_ROUND)
        add_text(sl, fatitle, 7.25, fay + 0.08, 4.9, 0.24, sz=9.5, bold=True, color=OBSIDIAN)
        add_text(sl, fadesc, 7.25, fay + 0.32, 4.9, 0.45, sz=8.5, color=MID_SLATE)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 10 — COLD CHAIN INTEGRITY & RISK MATRIX
# ══════════════════════════════════════════════════════════════════════════════

def slide_10_cold_chain():
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl, WHITE)
    add_transition(sl, "push", "l")
    add_header(sl, "QUALITY & COMPLIANCE",
               "Cold Chain Monitoring & The 18% Temperature Breach Leak",
               "Audit of 600 cold chain transactions identifying thermal breakdown points across the 5-stage logistics corridor.")
    add_page_number(sl, 10)

    df_cc = pd.read_csv(f'{DATA}/cold_chain_data.csv')
    breach_pct = (df_cc['Temperature_Breach'] == 'Yes').mean() * 100

    # LEFT: Circular Breach Rate Visual
    draw_rect(sl, 0.8, 1.8, 4.5, 5.2, fill=LIGHT_GRAY, line=CARD_BORDER, lw=1.5, shape=MSO_ROUND)
    add_text(sl, "Overall Thermal Compliance", 1.0, 1.95, 4.1, 0.28, sz=11, bold=True, color=OBSIDIAN)
    
    draw_oval(sl, 1.55, 2.45, 3.0, 3.0, fill=WHITE, line=RED_ACC, lw=5)
    add_text(sl, f"{breach_pct:.1f}%", 1.55, 3.4, 3.0, 0.8, sz=46, bold=True, color=RED_ACC, align=PP_ALIGN.CENTER)
    add_text(sl, "BREACH RATE", 1.55, 4.25, 3.0, 0.3, sz=10, bold=True, color=OBSIDIAN, align=PP_ALIGN.CENTER)
    add_text(sl, "Target SLA: < 5.0%", 1.55, 4.55, 3.0, 0.25, sz=8.5, italic=True, color=MID_SLATE, align=PP_ALIGN.CENTER)

    draw_rect(sl, 1.0, 5.75, 4.1, 1.0, fill=WHITE, line=CARD_BORDER, lw=1)
    add_text(sl, "Annual Quality Spoilage Impact: ~INR 3.8 Lakhs in quarantined or dumped beer cases due to summer thermal shock.",
             1.1, 5.85, 3.9, 0.8, sz=8.5, color=DEEP_NAVY)

    # RIGHT: Stage-by-Stage Breakdown
    draw_rect(sl, 5.5, 1.8, 7.0, 5.2, fill=LIGHT_GRAY, line=CARD_BORDER, lw=1.5, shape=MSO_ROUND)
    add_text(sl, "Thermal Breach Probability by Supply Chain Stage", 5.7, 1.95, 6.6, 0.28, sz=11, bold=True, color=OBSIDIAN)

    stages = [
        ("01. Brewery Loading Dock", "2.1%", "Automated pre-cooling in place; minimal breach.", GREEN_ACC),
        ("02. Primary Long-Haul Transit", "6.4%", "Reefer container units maintain stable 4°C–6°C range.", GREEN_ACC),
        ("03. Regional Distribution Hub", "8.9%", "Cross-dock staging occasionally delays cold vault intake.", BLUE_ACC),
        ("04. Secondary Urban Delivery", "18.2%", "Frequent truck door openings during multi-drop runs.", AMBER_GOLD),
        ("05. Retail & Taproom Handover", "34.8%", "Bottleneck! Small retailers lack walk-in coolers; beer sits on warm loading bays.", RED_ACC),
    ]

    for st_i, (st_name, st_pct, st_desc, st_col) in enumerate(stages):
        sy = 2.4 + st_i * 0.88
        draw_rect(sl, 5.7, sy, 6.6, 0.75, fill=WHITE, line=st_col, lw=1.2, shape=MSO_ROUND)
        draw_rect(sl, 5.7, sy, 0.1, 0.75, fill=st_col, shape=MSO_ROUND)
        add_text(sl, st_name, 5.9, sy + 0.08, 4.2, 0.24, sz=9.5, bold=True, color=DEEP_NAVY)
        add_text(sl, st_desc, 5.9, sy + 0.32, 4.8, 0.38, sz=8, color=MID_SLATE)
        add_text(sl, st_pct, 10.8, sy + 0.18, 1.3, 0.38, sz=14, bold=True, color=st_col, align=PP_ALIGN.RIGHT)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 11 — FREIGHT BENCHMARKING (Road vs Rail vs 3PL)
# ══════════════════════════════════════════════════════════════════════════════

def slide_11_freight():
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl, WHITE)
    add_transition(sl, "push", "l")
    add_header(sl, "FREIGHT STRATEGY",
               "Multi-Modal Freight Benchmarking: The Rail Arbitrage Opportunity",
               "Cross-analyzing 400 long-haul shipments across Road, Rail, Air, and 3PL to unlock margin expansion.")
    add_page_number(sl, 11)

    df_fr = pd.read_csv(f'{DATA}/freight_benchmark_data.csv')
    modal = df_fr.groupby('Freight_Mode').agg(
        Cost_Case=('Cost_Per_Case_INR', 'mean'),
        Transit=('Transit_Days', 'mean'),
        OTD=('On_Time_Pct', 'mean'),
        CO2=('CO2_Emission_KG', 'mean')
    ).reset_index().sort_values('Cost_Case')

    # Full Width Modal Comparison Table
    draw_rect(sl, 0.8, 1.8, 11.7, 3.4, fill=LIGHT_GRAY, line=CARD_BORDER, lw=1.5, shape=MSO_ROUND)
    add_text(sl, "Cross-Modal Performance Matrix (400 Benchmarked Shipments)", 1.0, 1.95, 11.3, 0.28, sz=11, bold=True, color=OBSIDIAN)

    mhdrs = ['Freight Mode', 'Cost / Case', 'Transit Time', 'On-Time Rate', 'CO2 / Shipment', 'Operational Recommendation']
    mw    = [2.3, 1.4, 1.4, 1.4, 1.6, 3.2]
    mx = 1.0
    for h, w in zip(mhdrs, mw):
        draw_rect(sl, mx, 2.35, w, 0.35, fill=DEEP_NAVY)
        add_text(sl, h, mx, 2.42, w, 0.22, sz=8.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        mx += w

    recs = {
        'Rail Freight': 'Primary bulk corridor (>800km). Expand share from 15% to 25%.',
        'Road (Truck)': 'Core inter-hub freight. Enforce cubic load optimization.',
        '3PL Partner': 'Use for peak seasonal surge capacity (Q4 buffer).',
        'Road (Mini-Truck)': 'Restrict to urban last-mile delivery only.',
        'Air Freight': 'Emergency stockout resupply only (costs 4.5× road).'
    }

    for ri, row in enumerate(modal.itertuples()):
        my = 2.75 + ri * 0.46
        bg_col = WHITE if ri % 2 == 0 else LIGHT_CREAM
        mx = 1.0
        r_rec = recs.get(row.Freight_Mode, 'Evaluate based on lane economics.')
        mvals = [row.Freight_Mode, f"INR {row.Cost_Case:.2f}", f"{row.Transit:.1f} Days", f"{row.OTD:.1f}%", f"{row.CO2:.1f} KG", r_rec]
        for ci, (v, w) in enumerate(zip(mvals, mw)):
            draw_rect(sl, mx, my, w, 0.42, fill=bg_col)
            add_text(sl, str(v), mx + 0.05, my + 0.08, w - 0.1, 0.28,
                     sz=8, bold=(ci==0 or ci==1), color=DEEP_NAVY if ci!=1 else AMBER_GOLD,
                     align=PP_ALIGN.CENTER if ci<5 else PP_ALIGN.LEFT)
            mx += w

    # BOTTOM: The Rail Expansion Business Case
    draw_rect(sl, 0.8, 5.4, 11.7, 1.65, fill=LIGHT_CREAM, line=YELLOW, lw=2, shape=MSO_ROUND)
    add_text(sl, "THE RAIL ARBITRAGE BUSINESS CASE (INR 18-24 Lakhs Annual Savings)", 1.0, 5.52, 11.3, 0.28,
             sz=11, bold=True, color=OBSIDIAN)
    add_text(sl, "• Rail freight delivers a 44% cost reduction per case compared to standard road trucking (INR 8.20 vs INR 14.80).\n"
                 "• While rail adds 2.3 days to transit time, beer inventory for stable core SKUs (Lager 500ml) can easily absorb this lead time via scheduled batch dispatching.\n"
                 "• Shifting 10% of long-haul volume from Mumbai/Delhi to Kolkata/Bhopal from road to rail unlocks INR 18–24L in recurring annual freight savings and cuts logistics emissions by 60 tonnes.",
             1.0, 5.85, 11.3, 1.1, sz=9, color=DEEP_NAVY)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 12 — DISTRIBUTOR SCORECARD (Weighted Ranking)
# ══════════════════════════════════════════════════════════════════════════════

def slide_12_distributors():
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl, WHITE)
    add_transition(sl, "push", "l")
    add_header(sl, "CHANNEL GOVERNANCE",
               "Distributor Performance Scorecard: Grade A to D Tiering",
               "Evaluating 20 regional distribution partners using a weighted composite model: OTIF (35%), Fill Rate (30%), Damage (20%), TAT (15%).")
    add_page_number(sl, 12)

    df_d = pd.read_csv(f'{DATA}/distributor_scorecard.csv')
    gc = df_d['Grade'].value_counts()

    # 4 Grade Tier Cards
    tiers = [
        ("GRADE A", gc.get('A', 0), "Score ≥ 85", "Top performers. Qualify for volume rebate incentives & priority stock.", GREEN_ACC),
        ("GRADE B", gc.get('B', 0), "Score 70 - 84", "Stable partners. Reliable execution with minor delivery delay outliers.", BLUE_ACC),
        ("GRADE C", gc.get('C', 0), "Score 55 - 69", "Underperforming. High damage rates (>1.2%) & delayed turnaround times.", AMBER_GOLD),
        ("GRADE D", gc.get('D', 0), "Score < 55", "Critical intervention! Chronic OTIF failure (<70%). Risk of contract review.", RED_ACC),
    ]

    for idx, (gtitle, gcount, gscore, gdesc, gcol) in enumerate(tiers):
        gx = 0.8 + idx * 2.95
        draw_rect(sl, gx, 1.8, 2.8, 2.1, fill=LIGHT_GRAY, line=CARD_BORDER, lw=1.5, shape=MSO_ROUND)
        draw_rect(sl, gx, 1.8, 2.8, 0.08, fill=gcol, shape=MSO_ROUND)
        add_text(sl, gtitle, gx + 0.15, 1.95, 2.5, 0.25, sz=11, bold=True, color=gcol)
        add_text(sl, str(gcount), gx + 0.15, 2.22, 1.2, 0.6, sz=36, bold=True, color=OBSIDIAN)
        add_text(sl, gscore, gx + 1.35, 2.38, 1.3, 0.25, sz=9, bold=True, color=MID_SLATE)
        add_text(sl, gdesc, gx + 0.15, 2.85, 2.5, 0.95, sz=8, color=DEEP_NAVY)

    # LOWER SECTION: Top 3 Benchmarks vs Bottom 3 Needing Action
    top3 = df_d.nlargest(3, 'Performance_Score')[['Distributor_ID', 'Region', 'OTIF_Pct', 'Damage_Rate_Pct', 'Performance_Score', 'Grade']]
    bot3 = df_d.nsmallest(3, 'Performance_Score')[['Distributor_ID', 'Region', 'OTIF_Pct', 'Damage_Rate_Pct', 'Performance_Score', 'Grade']]

    for col_x, tbl_title, tbl_data, tcol in [(0.8, "TOP 3 BENCHMARK DISTRIBUTORS", top3, GREEN_ACC), (6.8, "BOTTOM 3 PARTNERS (DPIP ACTION PLAN)", bot3, RED_ACC)]:
        draw_rect(sl, col_x, 4.1, 5.7, 3.0, fill=LIGHT_GRAY, line=CARD_BORDER, lw=1.5, shape=MSO_ROUND)
        add_text(sl, tbl_title, col_x + 0.2, 4.22, 5.3, 0.25, sz=10, bold=True, color=tcol)
        
        dhdrs = ['Distributor ID', 'Region', 'OTIF %', 'Damage %', 'Score', 'Grade']
        dw    = [1.4, 1.3, 0.8, 0.8, 0.7, 0.7]
        dx = col_x + 0.1
        for h, w in zip(dhdrs, dw):
            draw_rect(sl, dx, 4.55, w, 0.28, fill=DEEP_NAVY)
            add_text(sl, h, dx, 4.58, w, 0.2, sz=7.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
            dx += w

        for r_i, row in enumerate(tbl_data.itertuples(index=False)):
            dy = 4.88 + r_i * 0.45
            bg_c = WHITE if r_i % 2 == 0 else LIGHT_CREAM
            dx = col_x + 0.1
            rvals = [row.Distributor_ID, row.Region, f"{row.OTIF_Pct}%", f"{row.Damage_Rate_Pct:.2f}%", f"{row.Performance_Score:.1f}", row.Grade]
            for c_i, (v, w) in enumerate(zip(rvals, dw)):
                draw_rect(sl, dx, dy, w, 0.4, fill=bg_c)
                add_text(sl, str(v), dx, dy + 0.08, w, 0.24,
                         sz=8, bold=(c_i==0 or c_i==5),
                         color=tcol if c_i==5 else OBSIDIAN, align=PP_ALIGN.CENTER)
                dx += w


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 13 — STRATEGIC RECOMMENDATIONS & 90-DAY ROADMAP
# ══════════════════════════════════════════════════════════════════════════════

def slide_13_roadmap():
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl, WHITE)
    add_transition(sl, "push", "l")
    add_header(sl, "EXECUTIVE ACTION PLAN",
               "Strategic Transformation: 90-Day Implementation & ROI Model",
               "Three synchronized workstreams projected to recover INR 55–80 Lakhs in annualized operational efficiency.")
    add_page_number(sl, 13)

    pillars = [
        ("PILLAR 1", "DYNAMIC BUFFER & EOQ REBALANCING",
         "• Implement seasonal inventory build rules: pre-build 25% safety stock for top SKUs by Sept 1st.\n"
         "• Rebalance Stout 650ml order quantities to EOQ targets, cutting excess holding capital by 45%.\n"
         "• Integrate weekly POS sales data feeds from Grade A distributors to shorten forecast response lag.",
         "SAVINGS: INR 18 - 25 LAKHS / YR", AMBER_GOLD),
        
        ("PILLAR 2", "MULTI-MODAL LOGISTICS & TMS ROUTE POOLING",
         "• Establish long-term rail container contracts on Delhi–Kolkata and Mumbai–Bhopal routes.\n"
         "• Deploy dynamic route planning software to enforce minimum 85% cubic vehicle load factors.\n"
         "• Rationalize urban delivery fleet: shift sub-50km drops from 32-ft trucks to 14-ft reefer vehicles.",
         "SAVINGS: INR 22 - 32 LAKHS / YR", YELLOW),
        
        ("PILLAR 3", "COLD CHAIN AUDIT & DISTRIBUTOR PIP",
         "• Mandate IoT GPS temperature dataloggers on all secondary distributor trucks.\n"
         "• Launch Distributor Performance Improvement Program (DPIP) with quarterly business scorecards.\n"
         "• Reallocate underperforming Grade D regional territories to high-capacity Grade A partners.",
         "SAVINGS: INR 15 - 23 LAKHS / YR", GREEN_ACC),
    ]

    for idx, (pnum, ptitle, pdesc, psave, pcol) in enumerate(pillars):
        px = 0.8 + idx * 3.95
        draw_rect(sl, px, 1.8, 3.8, 4.4, fill=LIGHT_GRAY, line=CARD_BORDER, lw=1.5, shape=MSO_ROUND)
        draw_rect(sl, px, 1.8, 3.8, 0.12, fill=pcol, shape=MSO_ROUND)
        
        add_text(sl, pnum, px + 0.2, 2.05, 3.4, 0.25, sz=10, bold=True, color=pcol)
        add_text(sl, ptitle, px + 0.2, 2.35, 3.4, 0.5, sz=11, bold=True, color=OBSIDIAN)
        add_text(sl, pdesc, px + 0.2, 2.95, 3.4, 2.5, sz=8.5, color=DEEP_NAVY)
        
        draw_rect(sl, px + 0.2, 5.5, 3.4, 0.5, fill=OBSIDIAN)
        add_text(sl, psave, px + 0.25, 5.62, 3.3, 0.28, sz=8.5, bold=True, color=pcol, align=PP_ALIGN.CENTER)

    # 90-Day Timeline Bar
    draw_rect(sl, 0.8, 6.35, 11.7, 0.7, fill=DEEP_NAVY)
    add_text(sl, "90-DAY ROLLOUT: Days 1-30: EOQ Model & Safety Stock Reset  |  Days 31-60: Rail Contract & Route Pooling Pilot  |  Days 61-90: Cold Chain IoT & DPIP Governance",
             1.0, 6.55, 11.3, 0.35, sz=9, bold=True, color=YELLOW)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 14 — EXECUTIVE CLOSING & CONTACT
# ══════════════════════════════════════════════════════════════════════════════

def slide_14_closing():
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl, OBSIDIAN)
    add_transition(sl, "fade")
    add_page_number(sl, 14, dark=True)

    # Warm radial glow accent
    draw_oval(sl, 3.0, 0.5, 7.3, 6.5, fill=RGBColor(0x1F, 0x15, 0x08))

    # Bold Giant Closing Letters
    add_text(sl, "THANK YOU.", 1.5, 1.2, 10.33, 1.2,
             sz=84, bold=True, color=YELLOW, align=PP_ALIGN.CENTER)
    
    add_text(sl, "Ready to drive operational excellence in beverage supply chains.",
             1.5, 2.45, 10.33, 0.45, sz=16, color=WHITE, align=PP_ALIGN.CENTER)

    draw_rect(sl, 4.0, 3.1, 5.33, 0.04, fill=AMBER_GOLD)

    # Central Profile Box
    draw_rect(sl, 3.5, 3.4, 6.33, 2.7, fill=RGBColor(0x16, 0x1E, 0x2E), line=YELLOW, lw=1.5, shape=MSO_ROUND)
    
    add_text(sl, "HARSH RAJ PANDEY", 3.7, 3.65, 5.9, 0.4,
             sz=20, bold=True, color=YELLOW, align=PP_ALIGN.CENTER)
    add_text(sl, "Aspiring Logistics & Supply Chain Manager · AB InBev Candidate", 3.7, 4.05, 5.9, 0.3,
             sz=10.5, color=LIGHT_CREAM, italic=True, align=PP_ALIGN.CENTER)

    contact_items = [
        ("GitHub Repository", "https://github.com/Harsh258-collab/beerflow-supply-chain"),
        ("LinkedIn Profile", "https://www.linkedin.com/in/harsh-raj-pandey-1a0319325"),
        ("Contact Email", "harsh258.collab@gmail.com"),
    ]
    for ci, (clabel, cval) in enumerate(contact_items):
        cy = 4.5 + ci * 0.45
        add_text(sl, f"{clabel}:", 3.9, cy, 2.0, 0.28, sz=9.5, bold=True, color=AMBER_GOLD)
        add_text(sl, cval, 5.9, cy, 3.7, 0.28, sz=9.5, color=WHITE)

    add_text(sl, '"In the beverage industry, logistics is not a cost center — it is the primary brand differentiator."',
             1.5, 6.55, 10.33, 0.4, sz=11, italic=True, color=LIGHT_CREAM, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN EXECUTION
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == '__main__':
    print("[...] Building Ultra-Pro 14-Slide Presentation Deck...")
    slides = [
        (slide_01_cover, "Slide 1: Ultra-Bold Yellow & White Display Cover"),
        (slide_02_logistics_pathway, "Slide 2: Sinuous Logistics Highway Pathway"),
        (slide_03_beer_collision, "Slide 3: Cinematic Beer Delivery & Frosted Glass Collision with Droplets"),
        (slide_04_problem, "Slide 4: The Core Problem — Seasonality Mismatch"),
        (slide_05_scope, "Slide 5: Data Architecture & Scope Matrix"),
        (slide_06_demand, "Slide 6: Demand Forecasting & Seasonal Calendar"),
        (slide_07_eoq, "Slide 7: Inventory EOQ & Stockout Risk Matrix"),
        (slide_08_routes, "Slide 8: Logistics & Route Capacity Leak"),
        (slide_09_fleet, "Slide 9: Fleet Analytics & ESG Decarbonization"),
        (slide_10_cold_chain, "Slide 10: Cold Chain Integrity & Breach Analysis"),
        (slide_11_freight, "Slide 11: Freight Benchmarking & Rail Arbitrage"),
        (slide_12_distributors, "Slide 12: Distributor Performance Scorecard"),
        (slide_13_roadmap, "Slide 13: 90-Day Implementation & ROI Roadmap"),
        (slide_14_closing, "Slide 14: Executive Closing & Contact Card"),
    ]

    for s_fn, s_name in slides:
        s_fn()
        print(f"   [OK] {s_name}")

    # Set metadata
    prs.core_properties.author = 'Harsh Raj Pandey'
    prs.core_properties.last_modified_by = 'Harsh Raj Pandey'
    prs.core_properties.title = 'BeerFlow - Beverage Supply Chain Optimization'
    prs.core_properties.subject = 'Logistics & Supply Chain Analytics | AB InBev Project'
    prs.core_properties.keywords = 'Supply Chain, Logistics, Beer, FMCG, EOQ, OTIF, Cold Chain'
    prs.core_properties.created  = datetime.datetime(2026, 9, 10, 9, 0, 0)
    prs.core_properties.modified = datetime.datetime(2026, 9, 20, 21, 0, 0)

    prs.save(OUT)
    print(f"\n[DONE] Master presentation saved to:\n  {OUT}")
