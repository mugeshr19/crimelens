"""
Generate CrimeLens Final Review PowerPoint Presentation (.pptx)
Matching academic project review structure and visual standards.
"""

import sys
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

def build_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Color Palette - Professional Crime Intelligence Dark Theme
    COLOR_BG = RGBColor(11, 19, 43)          # Dark Navy Charcoal
    COLOR_CARD = RGBColor(19, 34, 64)        # Card Background Navy
    COLOR_CARD_BORDER = RGBColor(30, 58, 110) # Card Border
    COLOR_ACCENT = RGBColor(0, 212, 255)     # Electric Cyan / Accent
    COLOR_PRIMARY = RGBColor(59, 130, 246)   # Deep Blue
    COLOR_TEXT = RGBColor(241, 245, 249)     # Bright White/Grey
    COLOR_MUTED = RGBColor(148, 163, 184)    # Slate Muted Text
    COLOR_WARN = RGBColor(245, 158, 11)      # Amber Alert
    COLOR_DANGER = RGBColor(239, 68, 68)     # Red Critical
    COLOR_SUCCESS = RGBColor(16, 185, 129)   # Emerald Green
    COLOR_HEADER_BG = RGBColor(15, 23, 42)   # Header dark band

    def add_header(slide, title_text, category_text="CRIMELENS • AI-FIRST CRIME INTELLIGENCE OPERATING SYSTEM"):
        # Header background bar
        hdr_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.15))
        hdr_box.fill.solid()
        hdr_box.fill.fore_color.rgb = COLOR_HEADER_BG
        hdr_box.line.color.rgb = COLOR_CARD_BORDER
        hdr_box.line.width = Pt(1)

        # Cyan Accent line
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(1.15), Inches(13.333), Inches(0.04))
        line.fill.solid()
        line.fill.fore_color.rgb = COLOR_ACCENT
        line.line.fill.background()

        # Text in header
        tx_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.12), Inches(11.7), Inches(0.95))
        tf = tx_box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p1 = tf.paragraphs[0]
        p1.text = category_text.upper()
        p1.font.size = Pt(9.5)
        p1.font.bold = True
        p1.font.color.rgb = COLOR_ACCENT
        p1.font.name = "Arial"
        p1.space_after = Pt(2)

        p2 = tf.add_paragraph()
        p2.text = title_text
        p2.font.size = Pt(20)
        p2.font.bold = True
        p2.font.color.rgb = COLOR_TEXT
        p2.font.name = "Arial"

    def add_footer(slide, slide_num, total_slides=16):
        ftr_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(7.1), Inches(13.333), Inches(0.4))
        ftr_box.fill.solid()
        ftr_box.fill.fore_color.rgb = COLOR_HEADER_BG
        ftr_box.line.fill.background()

        tx_box = slide.shapes.add_textbox(Inches(0.8), Inches(7.15), Inches(11.7), Inches(0.3))
        tf = tx_box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p = tf.paragraphs[0]
        p.text = "Sri Eshwar College of Engineering  |  Dept of CSE  |  U23CS652 - Multidisciplinary Project  |  Batch 17"
        p.font.size = Pt(8.5)
        p.font.color.rgb = COLOR_MUTED
        p.font.name = "Arial"

        p_num = tf.add_paragraph()
        p_num.text = f"{slide_num} / {total_slides}"
        p_num.alignment = PP_ALIGN.RIGHT
        p_num.font.size = Pt(8.5)
        p_num.font.bold = True
        p_num.font.color.rgb = COLOR_ACCENT
        p_num.font.name = "Arial"

    def set_slide_background(slide):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = COLOR_BG
        bg.line.fill.background()
        return bg

    def create_card(slide, left, top, width, height, title="", border_color=COLOR_CARD_BORDER, bg_color=COLOR_CARD):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color
        card.line.color.rgb = border_color
        card.line.width = Pt(1.2)

        if title:
            tb = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.15), width - Inches(0.4), Inches(0.4))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.margin_top = tf.margin_bottom = tf.margin_left = tf.margin_right = 0
            p = tf.paragraphs[0]
            p.text = title
            p.font.size = Pt(13)
            p.font.bold = True
            p.font.color.rgb = COLOR_ACCENT
            p.font.name = "Arial"

        return card

    # =========================================================================
    # SLIDE 1: TITLE SLIDE
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    set_slide_background(s1)

    # Decorative header glow
    glow = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.1))
    glow.fill.solid()
    glow.fill.fore_color.rgb = COLOR_ACCENT
    glow.line.fill.background()

    # Title container card
    create_card(s1, Inches(0.8), Inches(0.6), Inches(11.733), Inches(2.7), bg_color=COLOR_CARD)
    tb = s1.shapes.add_textbox(Inches(1.2), Inches(0.85), Inches(11.0), Inches(2.2))
    tf = tb.text_frame
    tf.word_wrap = True

    p0 = tf.paragraphs[0]
    p0.text = "FINAL YEAR INNOVATIVE / MULTIDISCIPLINARY PROJECT REVIEW (U23CS652)"
    p0.font.size = Pt(11)
    p0.font.bold = True
    p0.font.color.rgb = COLOR_ACCENT
    p0.font.name = "Arial"
    p0.space_after = Pt(6)

    p1 = tf.add_paragraph()
    p1.text = "CrimeLens: An AI-First Crime Intelligence Operating System"
    p1.font.size = Pt(26)
    p1.font.bold = True
    p1.font.color.rgb = COLOR_TEXT
    p1.font.name = "Arial"
    p1.space_after = Pt(8)

    p2 = tf.add_paragraph()
    p2.text = "Bridging Statewide Police Analytics with Edge Tactical Response through Multi-Agent Spatio-Temporal Intelligence & Cryptographic Governance"
    p2.font.size = Pt(12)
    p2.font.color.rgb = COLOR_MUTED
    p2.font.name = "Arial"

    # Team Members Card (Left)
    create_card(s1, Inches(0.8), Inches(3.55), Inches(5.7), Inches(3.2), title="PROJECT TEAM • BATCH 17")
    tb_team = s1.shapes.add_textbox(Inches(1.05), Inches(4.15), Inches(5.2), Inches(2.4))
    tf_team = tb_team.text_frame
    tf_team.word_wrap = True

    members = [
        ("Parameshwaran P", "722824104166", "Backend API, Graph Analytics & Ingestion"),
        ("Ram Vishnu P", "722824104186", "ST-DBSCAN Hotspots & Forecast Engine"),
        ("Ravishankar AG", "722824104191", "Command Web Console & UI/UX Design"),
        ("Rishivardhan J", "722824104196", "Expo Mobile Field App & SHA-256 Ledger")
    ]
    for idx, (name, roll, role) in enumerate(members):
        p = tf_team.paragraphs[0] if idx == 0 else tf_team.add_paragraph()
        p.text = f"{idx+1}. {name} ({roll})"
        p.font.size = Pt(11.5)
        p.font.bold = True
        p.font.color.rgb = COLOR_TEXT
        p.font.name = "Arial"
        
        p_sub = tf_team.add_paragraph()
        p_sub.text = f"    Role: {role}"
        p_sub.font.size = Pt(9.5)
        p_sub.font.color.rgb = COLOR_MUTED
        p_sub.font.name = "Arial"
        p_sub.space_after = Pt(4)

    # Faculty Mentor & Institution Card (Right)
    create_card(s1, Inches(6.8), Inches(3.55), Inches(5.733), Inches(3.2), title="ACADEMIC SUPERVISION & INSTITUTION")
    tb_inst = s1.shapes.add_textbox(Inches(7.05), Inches(4.15), Inches(5.2), Inches(2.4))
    tf_inst = tb_inst.text_frame
    tf_inst.word_wrap = True

    p_f = tf_inst.paragraphs[0]
    p_f.text = "FACULTY MENTOR"
    p_f.font.size = Pt(10)
    p_f.font.bold = True
    p_f.font.color.rgb = COLOR_ACCENT
    p_f.font.name = "Arial"

    p_f2 = tf_inst.add_paragraph()
    p_f2.text = "Dr. S. Yasotha, M.E., Ph.D."
    p_f2.font.size = Pt(13)
    p_f2.font.bold = True
    p_f2.font.color.rgb = COLOR_TEXT
    p_f2.font.name = "Arial"

    p_f3 = tf_inst.add_paragraph()
    p_f3.text = "Associate Professor, Department of CSE (CYS)"
    p_f3.font.size = Pt(10.5)
    p_f3.font.color.rgb = COLOR_MUTED
    p_f3.font.name = "Arial"
    p_f3.space_after = Pt(14)

    p_i = tf_inst.add_paragraph()
    p_i.text = "DEPARTMENT & INSTITUTION"
    p_i.font.size = Pt(10)
    p_i.font.bold = True
    p_i.font.color.rgb = COLOR_ACCENT
    p_i.font.name = "Arial"

    p_i2 = tf_inst.add_paragraph()
    p_i2.text = "Department of Computer Science and Engineering"
    p_i2.font.size = Pt(12)
    p_i2.font.bold = True
    p_i2.font.color.rgb = COLOR_TEXT
    p_i2.font.name = "Arial"

    p_i3 = tf_inst.add_paragraph()
    p_i3.text = "Sri Eshwar College of Engineering (Autonomous), Coimbatore\nAffiliated to Anna University, Chennai"
    p_i3.font.size = Pt(10)
    p_i3.font.color.rgb = COLOR_MUTED
    p_i3.font.name = "Arial"

    add_footer(s1, 1)

    # =========================================================================
    # SLIDE 2: PROBLEM & SOLUTION STORY (BEFORE VS AFTER)
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    set_slide_background(s2)
    add_header(s2, "Operational Context: Legacy Policing Dilemma vs. CrimeLens Paradigm")

    # Column 1: Before (The Problem)
    create_card(s2, Inches(0.8), Inches(1.35), Inches(5.7), Inches(5.5), title="BEFORE CRIMELENS (EXISTING CCTNS/PORTALS)", border_color=COLOR_DANGER)
    tb_b = s2.shapes.add_textbox(Inches(1.05), Inches(1.95), Inches(5.2), Inches(4.7))
    tf_b = tb_b.text_frame
    tf_b.word_wrap = True

    problems = [
        ("Bureaucratic Data Silos", "1,100+ stations operate as disconnected databases. Cross-district syndicate operations appear as isolated, unrelated FIR cases."),
        ("Static, Calendar-Driven Patrols", "Patrol units (Hoysala/Akka) follow rigid pre-set calendar routes with zero empirical risk weighting, leaving emerging hot zones unpoliced."),
        ("Vulnerable Chain-of-Custody", "Evidence logs and incident diaries maintained manually or in mutable SQL databases, creating severe credibility challenges in constitutional trials."),
        ("Black-Box AI Mistrust", "Commercial GIS heatmaps provide opaque kernel density visuals without explainable underlying features, triggering skepticism from judiciary and IOs.")
    ]
    for idx, (title, desc) in enumerate(problems):
        p = tf_b.paragraphs[0] if idx == 0 else tf_b.add_paragraph()
        p.text = f"[!] {title}"
        p.font.size = Pt(11.5)
        p.font.bold = True
        p.font.color.rgb = COLOR_DANGER
        p.font.name = "Arial"

        p_desc = tf_b.add_paragraph()
        p_desc.text = desc
        p_desc.font.size = Pt(9.5)
        p_desc.font.color.rgb = COLOR_MUTED
        p_desc.font.name = "Arial"
        p_desc.space_after = Pt(8)

    # Column 2: After (The Solution)
    create_card(s2, Inches(6.8), Inches(1.35), Inches(5.733), Inches(5.5), title="AFTER CRIMELENS (AI-FIRST INTELLIGENCE OS)", border_color=COLOR_SUCCESS)
    tb_a = s2.shapes.add_textbox(Inches(7.05), Inches(1.95), Inches(5.2), Inches(4.7))
    tf_a = tb_a.text_frame
    tf_a.word_wrap = True

    solutions = [
        ("Unified Statewide Corpus", "Normalized ingestion of 201,733 KSP records across 37 districts; automated multi-jurisdictional syndicate graph linkage."),
        ("Predictive Patrol Optimization", "Fuses ST-DBSCAN hotspots, Holt's 4-week crime forecast, and 4-hour risk windows into dynamic, turn-by-turn mobile patrol deployments."),
        ("Cryptographic Governance", "Every administrative decision, dispatch approval, and evidence handoff is sealed in an append-only SHA-256 hash-chain ledger."),
        ("Human-in-the-Loop Explainability", "Every AI flag delivers plain-language rationale and SHAP-style feature attribution (CCTV blindspots, entry corridors, MO overlap).")
    ]
    for idx, (title, desc) in enumerate(solutions):
        p = tf_a.paragraphs[0] if idx == 0 else tf_a.add_paragraph()
        p.text = f"[OK] {title}"
        p.font.size = Pt(11.5)
        p.font.bold = True
        p.font.color.rgb = COLOR_SUCCESS
        p.font.name = "Arial"

        p_desc = tf_a.add_paragraph()
        p_desc.text = desc
        p_desc.font.size = Pt(9.5)
        p_desc.font.color.rgb = COLOR_MUTED
        p_desc.font.name = "Arial"
        p_desc.space_after = Pt(8)

    add_footer(s2, 2)

    # =========================================================================
    # SLIDE 3: TECHNICAL APPROACH & SYSTEM ARCHITECTURE
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    set_slide_background(s3)
    add_header(s3, "System Architecture: 5-Tier End-to-End Intelligence Pipeline")

    arch_img_path = r"d:\projects\CrimeLens\CrimeLens_Full_Project\docs\screenshots\crimelens_architecture_diagram_academic.png"
    if not os.path.exists(arch_img_path):
        arch_img_path = r"d:\projects\CrimeLens\CrimeLens_Full_Project\docs\screenshots\architecture_diagram.png"

    if os.path.exists(arch_img_path):
        create_card(s3, Inches(0.8), Inches(1.35), Inches(7.5), Inches(5.5), title="ARCHITECTURAL FLOW DIAGRAM")
        s3.shapes.add_picture(arch_img_path, Inches(1.0), Inches(1.9), width=Inches(7.1))
    else:
        create_card(s3, Inches(0.8), Inches(1.35), Inches(7.5), Inches(5.5), title="ARCHITECTURAL DESIGN LAYERS")

    # Right side: Architecture explanation cards
    create_card(s3, Inches(8.5), Inches(1.35), Inches(4.033), Inches(5.5), title="5 CORE SYSTEM TIERS")
    tb_arch = s3.shapes.add_textbox(Inches(8.7), Inches(1.9), Inches(3.65), Inches(4.8))
    tf_arch = tb_arch.text_frame
    tf_arch.word_wrap = True

    tiers = [
        ("Tier 1: Data Ingestion Plane", "Parses 12 KSP CSV files (201,733 records); normalizes IPC/SLL categories and station metadata."),
        ("Tier 2: Multi-Model Intelligence", "Pure TS mathematical engines: ST-DBSCAN hotspots, Louvain network modularity, Holt's trend forecaster."),
        ("Tier 3: Governance & Security", "JWT authentication, WebAuthn biometrics, 4-tier RBAC, and SHA-256 tamper-evident recursive audit chain."),
        ("Tier 4: Command Web Console", "React 18 + Vite dashboard with Leaflet GIS mapping, Force-Graph-2D gang linkages, and Recharts trends."),
        ("Tier 5: Tactical Mobile Edge", "Expo SDK 54 / React Native field app with SQLite offline store, GPS telemetry, and 3.8ms Haversine dispatch.")
    ]
    for idx, (title, desc) in enumerate(tiers):
        p = tf_arch.paragraphs[0] if idx == 0 else tf_arch.add_paragraph()
        p.text = f"{idx+1}. {title}"
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = COLOR_ACCENT
        p.font.name = "Arial"

        p_d = tf_arch.add_paragraph()
        p_d.text = desc
        p_d.font.size = Pt(8.8)
        p_d.font.color.rgb = COLOR_MUTED
        p_d.font.name = "Arial"
        p_d.space_after = Pt(4)

    add_footer(s3, 3)

    # =========================================================================
    # SLIDE 4: MATHEMATICAL FORMULATIONS & CORE ALGORITHMS
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    set_slide_background(s4)
    add_header(s4, "Governing Algorithms: Mathematical Foundations & Operational Roles")

    algos = [
        ("ST-DBSCAN Hotspot Clustering",
         "eps_spatial = 600m, Delta_t = 7 days\nintensityScore = min(98, max(35, round(count/40 * 100)))\nSurge Flag: count > 15 * (days/30) * 1.3",
         "functions/incidentCluster.ts",
         "Extracts high-density crime corridors and isolates seasonal surges from baseline noise."),
        
        ("Louvain Modularity Detection",
         "Delta_Q = w_in - (Sigma_tot * k_i) / (2m)\nLive Degree Centrality: C_D(v) = deg(v)\nBipartite Projection: w(u,v) = |N(u) intersect N(v)|",
         "functions/graphAnalytics.ts",
         "Identifies distributed crime syndicates, criminal kingpins, and inferred co-conspirators."),
        
        ("Holt's Double-Exponential Smoothing",
         "L_t = alpha * y_t + (1 - alpha)(L_{t-1} + T_{t-1})  [alpha=0.5]\nT_t = beta(L_t - L_{t-1}) + (1 - beta)T_{t-1}  [beta=0.3]\ny_hat_{t+h} = max(0, round(L_t + h*T_t)),  RMSE=sqrt(MSE)",
         "functions/offlineAnalytics.ts",
         "Generates 4-week forward crime forecasts with 80% confidence bands for proactive patrol planning."),
        
        ("TF-IDF Modus Operandi Profiler",
         "tf(t, d) = f_{t,d} / |d|,  idf(t) = ln((N+1)/(df+1)) + 1\ncos(q, d) = (q . d) / (||q|| * ||d||)\nSuspect Score = 0.6*(deg/max_deg) + 0.4*moSim",
         "functions/analystFeatures.ts",
         "Matches criminal behavioural signatures across stations and ranks suspects based on graph proximity + MO."),
        
        ("Haversine Great-Circle Dispatch",
         "s = sin^2(dLat/2) + cos(lat1)*cos(lat2)*sin^2(dLng/2)\nd = 2 * R * asin(sqrt(s))  [R = 6,371,000 m]\nETA = max(1, round(d / 400 m/min))  [~24 km/h urban]",
         "functions/fleetTelemetry.ts",
         "Computes geodesic proximity from live Akka units to Panic SOS locations in 3.8 ms."),
        
        ("SHA-256 Recursive Hash Chain",
         "H_k = SHA256(payload_k || H_{k-1})\npayload_k = (seq || uid || action || target || ts || H_{k-1})\nIntegrity: H_{k-1}^expected == H_{k-1}^stored",
         "server/police/store.ts",
         "Provides an append-only, tamper-evident audit ledger that proves non-repudiation in court trials.")
    ]

    for i, (title, math_str, code_loc, role_str) in enumerate(algos):
        r = i // 3
        c = i % 3
        cx = Inches(0.8 + c * 4.0)
        cy = Inches(1.35 + r * 2.8)

        create_card(s4, cx, cy, Inches(3.75), Inches(2.65), title=title)
        tb_a = s4.shapes.add_textbox(cx + Inches(0.15), cy + Inches(0.5), Inches(3.45), Inches(2.05))
        tf_a = tb_a.text_frame
        tf_a.word_wrap = True

        p_m = tf_a.paragraphs[0]
        p_m.text = math_str
        p_m.font.size = Pt(8.5)
        p_m.font.name = "Courier New"
        p_m.font.color.rgb = COLOR_TEXT
        p_m.space_after = Pt(4)

        p_loc = tf_a.add_paragraph()
        p_loc.text = f"Source: {code_loc}"
        p_loc.font.size = Pt(7.8)
        p_loc.font.bold = True
        p_loc.font.color.rgb = COLOR_ACCENT
        p_loc.font.name = "Arial"
        p_loc.space_after = Pt(2)

        p_r = tf_a.add_paragraph()
        p_r.text = role_str
        p_r.font.size = Pt(8.2)
        p_r.font.color.rgb = COLOR_MUTED
        p_r.font.name = "Arial"

    add_footer(s4, 4)

    # =========================================================================
    # SLIDE 5: DATASET ARCHITECTURE: REAL VS. SYNTHETIC
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    set_slide_background(s5)
    add_header(s5, "Hybrid Data Architecture: Statewide KSP Corpus vs. Micro-Incidents")

    create_card(s5, Inches(0.8), Inches(1.35), Inches(5.7), Inches(5.5), title="TIER 1: REAL STATEWIDE KSP CORPUS", border_color=COLOR_PRIMARY)
    tb_real = s5.shapes.add_textbox(Inches(1.05), Inches(1.95), Inches(5.2), Inches(4.7))
    tf_real = tb_real.text_frame
    tf_real.word_wrap = True

    real_points = [
        ("Official Government Ingestion", "201,733 records ingested from 12 official CSV datasets published under Karnataka Open Data Initiative and KSP portals."),
        ("Statewide Jurisdictional Scope", "Covers all 37 police districts and commissionerates (Bengaluru City, Mysuru, Hubballi-Dharwad, Mangaluru, Belagavi, etc.)."),
        ("Statutory Classification", "Full normalization across Indian Penal Code (IPC) crime heads, Special & Local Laws (SLL), and monthly crime reviews."),
        ("Operational Utility in CrimeLens", "Establishes authentic district crime totals, baseline weekly rates, inter-agency CCTNS connectors, and strategic resource allocation rankings."),
        ("Scientific Honesty", "Real datasets are aggregate district-level figures; they lack micro-level GPS coordinates and citizen identities due to privacy protocols.")
    ]
    for idx, (head, text) in enumerate(real_points):
        p = tf_real.paragraphs[0] if idx == 0 else tf_real.add_paragraph()
        p.text = f"[REAL] {head}"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = COLOR_ACCENT
        p.font.name = "Arial"

        p_t = tf_real.add_paragraph()
        p_t.text = text
        p_t.font.size = Pt(9.2)
        p_t.font.color.rgb = COLOR_MUTED
        p_t.font.name = "Arial"
        p_t.space_after = Pt(6)

    create_card(s5, Inches(6.8), Inches(1.35), Inches(5.733), Inches(5.5), title="TIER 2: SYNTHETIC TACTICAL INCIDENTS", border_color=COLOR_WARN)
    tb_syn = s5.shapes.add_textbox(Inches(7.05), Inches(1.95), Inches(5.2), Inches(4.7))
    tf_syn = tb_syn.text_frame
    tf_syn.word_wrap = True

    syn_points = [
        ("Micro-Level Tactical Simulation", "2,200 synthetic GPS-tagged incident records generated to evaluate edge clustering and patrol routing engines."),
        ("Anchor-Based Geospatial Clusters", "1,500 incidents clustered around 10 realistic Karnataka police anchors (Peenya Corridor, Koramangala 80ft, Hebbal, Gokulam, etc.)."),
        ("Temporal & Gaussian Variance", "~300m spatial Gaussian offset (eps ~0.003 deg) with 90-day timestamps and an authentic 65% nighttime bias (23:00 - 03:00 hrs)."),
        ("Background Noise Simulation", "700 uniformly distributed incidents across 8 districts to challenge ST-DBSCAN noise filtering and prevent false alarms."),
        ("Synthetic Entity Dossiers", "Anonymized suspect network records (Blade Ramesh, Phisher Vikky) and Akka patrol telemetry waypoints for prototype evaluation.")
    ]
    for idx, (head, text) in enumerate(syn_points):
        p = tf_syn.paragraphs[0] if idx == 0 else tf_syn.add_paragraph()
        p.text = f"[SYNTHETIC] {head}"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = COLOR_WARN
        p.font.name = "Arial"

        p_t = tf_syn.add_paragraph()
        p_t.text = text
        p_t.font.size = Pt(9.2)
        p_t.font.color.rgb = COLOR_MUTED
        p_t.font.name = "Arial"
        p_t.space_after = Pt(6)

    add_footer(s5, 5)

    # =========================================================================
    # SLIDE 6: PROTOTYPE DEMO VIDEO PLACEHOLDER
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    set_slide_background(s6)
    add_header(s6, "Prototype Demonstration: Live Operational Walkthrough")

    demo_poster = r"d:\projects\CrimeLens\CrimeLens_Full_Project\docs\demo-poster.jpg"
    create_card(s6, Inches(0.8), Inches(1.35), Inches(8.0), Inches(5.5), title="CRIMELENS FULL 2.5-MINUTE DEMO VIDEO")

    if os.path.exists(demo_poster):
        s6.shapes.add_picture(demo_poster, Inches(1.0), Inches(1.9), width=Inches(7.6))
    else:
        # Video box placeholder
        vbox = s6.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.0), Inches(1.9), Inches(7.6), Inches(4.7))
        vbox.fill.solid()
        vbox.fill.fore_color.rgb = RGBColor(5, 10, 20)
        vbox.line.color.rgb = COLOR_ACCENT

        tb_v = s6.shapes.add_textbox(Inches(2.0), Inches(3.5), Inches(5.6), Inches(1.5))
        p = tb_v.text_frame.paragraphs[0]
        p.text = "[ VIDEO PLAYER EMBED PLACEHOLDER ]\nCrimeLens 2.5-Minute End-to-End Walkthrough"
        p.alignment = PP_ALIGN.CENTER
        p.font.size = Pt(14)
        p.font.color.rgb = COLOR_ACCENT

    # Right side: Demo highlights card
    create_card(s6, Inches(9.0), Inches(1.35), Inches(3.533), Inches(5.5), title="DEMONSTRATION PHASES")
    tb_dh = s6.shapes.add_textbox(Inches(9.2), Inches(1.9), Inches(3.133), Inches(4.8))
    tf_dh = tb_dh.text_frame
    tf_dh.word_wrap = True

    phases = [
        ("1. Real KSP Ingestion", "Bootstrapping 201,733 statewide crime records into memory."),
        ("2. Command Live Anomaly", "Real-time alert streaming with confidence thresholding."),
        ("3. Hotspot ST-DBSCAN", "Dynamic slider adjusting epsilon and temporal windows."),
        ("4. Syndicate Louvain Graph", "Isolating cross-district gang clusters and kingpins."),
        ("5. Predictive Patrol Planner", "Deploying nearest patrol vehicle to high-risk sector."),
        ("6. Mobile Edge Dispatch", "Frontline officer SOS trigger to 3.8ms Haversine dispatch."),
        ("7. Cryptographic Governance", "One-click verification of SHA-256 audit ledger chain.")
    ]
    for idx, (ph, desc) in enumerate(phases):
        p = tf_dh.paragraphs[0] if idx == 0 else tf_dh.add_paragraph()
        p.text = ph
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = COLOR_ACCENT
        p.font.name = "Arial"

        p_d = tf_dh.add_paragraph()
        p_d.text = desc
        p_d.font.size = Pt(8.5)
        p_d.font.color.rgb = COLOR_MUTED
        p_d.font.name = "Arial"
        p_d.space_after = Pt(3)

    add_footer(s6, 6)

    # =========================================================================
    # SLIDE 7: PROTOTYPE IMPLEMENTATION SCREENSHOTS: WEB CONSOLE
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    set_slide_background(s7)
    add_header(s7, "Web Command Console: Multi-Role Operational Dashboards")

    shots = [
        ("docs/screenshots/command-center.png", "State Command Center", "Live crime anomalies, KPIs & supervisor review queue"),
        ("docs/screenshots/network-graph-force.png", "Louvain Network Graph", "Force-canvas syndicate clusters; solid=confirmed, dashed=inferred"),
        ("docs/screenshots/crime-trends.png", "Crime Trends & Forecasting", "Holt's linear trend 4-week forecast with RMSE metrics"),
        ("docs/screenshots/governance-audit.png", "Cryptographic Governance", "100% SHA-256 hash-chain verification & bias analysis")
    ]

    for idx, (rel_path, caption, sub) in enumerate(shots):
        r = idx // 2
        c = idx % 2
        cx = Inches(0.8 + c * 5.95)
        cy = Inches(1.35 + r * 2.8)

        full_p = os.path.join(r"d:\projects\CrimeLens\CrimeLens_Full_Project", rel_path)
        create_card(s7, cx, cy, Inches(5.75), Inches(2.65), title="")

        if os.path.exists(full_p):
            s7.shapes.add_picture(full_p, cx + Inches(0.12), cy + Inches(0.12), width=Inches(3.3))
            
            tb_c = s7.shapes.add_textbox(cx + Inches(3.55), cy + Inches(0.2), Inches(2.05), Inches(2.2))
            tf_c = tb_c.text_frame
            tf_c.word_wrap = True
            p = tf_c.paragraphs[0]
            p.text = caption
            p.font.size = Pt(11)
            p.font.bold = True
            p.font.color.rgb = COLOR_ACCENT
            p.font.name = "Arial"

            p_s = tf_c.add_paragraph()
            p_s.text = sub
            p_s.font.size = Pt(8.5)
            p_s.font.color.rgb = COLOR_MUTED
            p_s.font.name = "Arial"
            p_s.space_before = Pt(4)

    add_footer(s7, 7)

    # =========================================================================
    # SLIDE 8: PROTOTYPE IMPLEMENTATION SCREENSHOTS: MOBILE & DISPATCH
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    set_slide_background(s8)
    add_header(s8, "Tactical Edge: Field-Officer Mobile Application & Patrol Dispatch")

    shots_mob = [
        ("docs/screenshots/patrol-planner.png", "Predictive Patrol Planner", "Risk-weighted sectors & nearest vehicle allocation"),
        ("docs/screenshots/akka-fleet.png", "Akka Fleet Telemetry", "Live simulated GPS tracking & Haversine dispatch"),
        ("docs/screenshots/poster-av1.png", "Mobile App: Ops Login", "Biometric / badge auth with offline SQLite storage"),
        ("docs/screenshots/poster-av2.png", "Evidence Chain-of-Custody", "Digital evidence locker with QR integrity sealing")
    ]

    for idx, (rel_path, caption, sub) in enumerate(shots_mob):
        r = idx // 2
        c = idx % 2
        cx = Inches(0.8 + c * 5.95)
        cy = Inches(1.35 + r * 2.8)

        full_p = os.path.join(r"d:\projects\CrimeLens\CrimeLens_Full_Project", rel_path)
        create_card(s8, cx, cy, Inches(5.75), Inches(2.65), title="")

        if os.path.exists(full_p):
            s8.shapes.add_picture(full_p, cx + Inches(0.12), cy + Inches(0.12), width=Inches(3.3))
            
            tb_c = s8.shapes.add_textbox(cx + Inches(3.55), cy + Inches(0.2), Inches(2.05), Inches(2.2))
            tf_c = tb_c.text_frame
            tf_c.word_wrap = True
            p = tf_c.paragraphs[0]
            p.text = caption
            p.font.size = Pt(11)
            p.font.bold = True
            p.font.color.rgb = COLOR_ACCENT
            p.font.name = "Arial"

            p_s = tf_c.add_paragraph()
            p_s.text = sub
            p_s.font.size = Pt(8.5)
            p_s.font.color.rgb = COLOR_MUTED
            p_s.font.name = "Arial"
            p_s.space_before = Pt(4)

    add_footer(s8, 8)

    # =========================================================================
    # SLIDE 9: MVP STATUS & TECHNOLOGY READINESS LEVEL (TRL 4)
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    set_slide_background(s9)
    add_header(s9, "Maturity Assessment: Current MVP Status & Technology Readiness (TRL 4)")

    # Left: TRL Definition & Completed Components
    create_card(s9, Inches(0.8), Inches(1.35), Inches(5.7), Inches(5.5), title="TECHNOLOGY READINESS LEVEL (TRL 4)", border_color=COLOR_SUCCESS)
    tb_trl = s9.shapes.add_textbox(Inches(1.05), Inches(1.95), Inches(5.2), Inches(4.7))
    tf_trl = tb_trl.text_frame
    tf_trl.word_wrap = True

    p = tf_trl.paragraphs[0]
    p.text = "TRL 4: Component / Subsystem Validation in Controlled Environment"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_SUCCESS
    p.font.name = "Arial"
    p.space_after = Pt(8)

    p_body = tf_trl.add_paragraph()
    p_body.text = "CrimeLens is an academic software prototype validated against real KSP datasets in a simulated laboratory testbed. It is NOT currently deployed in active police production."
    p_body.font.size = Pt(9.5)
    p_body.font.color.rgb = COLOR_MUTED
    p_body.font.name = "Arial"
    p_body.space_after = Pt(12)

    p_comp = tf_trl.add_paragraph()
    p_comp.text = "IMPLEMENTED & VERIFIED SUBSYSTEMS (CLASS A):"
    p_comp.font.size = Pt(10)
    p_comp.font.bold = True
    p_comp.font.color.rgb = COLOR_ACCENT
    p_comp.font.name = "Arial"

    items_ok = [
        "Real 201k KSP CSV ingestion & station schema mapping",
        "ST-DBSCAN clustering with dynamic epsilon & time slider",
        "Louvain modularity gang discovery & starburst graph",
        "Holt's double exponential smoothing (MAPE: 6.2%, RMSE: 1.86)",
        "TF-IDF semantic MO search & suspect proximity ranking",
        "On-device TensorFlow.js COCO-SSD object/vehicle detection",
        "Haversine nearest-unit fleet dispatch (3.8ms latency)",
        "SHA-256 recursive hash-chain tamper verification",
        "Expo SDK 54 mobile app with offline SQLite store"
    ]
    for itm in items_ok:
        p_i = tf_trl.add_paragraph()
        p_i.text = f"[OK] {itm}"
        p_i.font.size = Pt(8.5)
        p_i.font.color.rgb = COLOR_TEXT
        p_i.font.name = "Arial"

    # Right: Pending & Roadmap Work
    create_card(s9, Inches(6.8), Inches(1.35), Inches(5.733), Inches(5.5), title="PENDING WORK & OPERATIONAL BOUNDARIES", border_color=COLOR_WARN)
    tb_pen = s9.shapes.add_textbox(Inches(7.05), Inches(1.95), Inches(5.2), Inches(4.7))
    tf_pen = tb_pen.text_frame
    tf_pen.word_wrap = True

    p_p = tf_pen.paragraphs[0]
    p_p.text = "PLANNED ENHANCEMENTS FOR PRODUCTION (TRL 5 -> 7):"
    p_p.font.size = Pt(10)
    p_p.font.bold = True
    p_p.font.color.rgb = COLOR_WARN
    p_p.font.name = "Arial"
    p_p.space_after = Pt(8)

    roadmap = [
        ("Live CCTNS 2.0 API Gateway (Class D)", "Awaiting formal government staging credentials for bidirectional REST synchronization with central NCRB servers."),
        ("Statewide Distributed Neo4j Cluster (Class D)", "Scaling the network graph from in-memory JavaScript objects to multi-million node distributed enterprise graphs."),
        ("Edge Hardware Video Acceleration", "Deploying TensorRT / Coral NPU accelerators inside patrol vehicles for real-time 60 FPS multi-camera ANPR scanning."),
        ("Human-in-the-Loop Safeguards", "Formalizing legal review protocols: no automated arrest, FIR filing, or surveillance without verified digital SHO sign-off."),
        ("Statewide Pilot Field Trial", "Targeting initial pilot deployment across 5 high-density police stations in Bengaluru Urban under controlled supervision.")
    ]
    for idx, (rh, rd) in enumerate(roadmap):
        p_rh = tf_pen.add_paragraph()
        p_rh.text = f"[ROADMAP] {rh}"
        p_rh.font.size = Pt(10)
        p_rh.font.bold = True
        p_rh.font.color.rgb = COLOR_TEXT
        p_rh.font.name = "Arial"

        p_rd = tf_pen.add_paragraph()
        p_rd.text = rd
        p_rd.font.size = Pt(8.5)
        p_rd.font.color.rgb = COLOR_MUTED
        p_rd.font.name = "Arial"
        p_rd.space_after = Pt(6)

    add_footer(s9, 9)

    # =========================================================================
    # SLIDE 10: USER TESTING & BENCHMARK VALIDATION
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    set_slide_background(s10)
    add_header(s10, "Empirical Validation: Computational Latency & Predictive Accuracy")

    # Left: Empirical Benchmarks Table
    create_card(s10, Inches(0.8), Inches(1.35), Inches(6.5), Inches(5.5), title="SYSTEM LATENCY & THROUGHPUT BENCHMARKS")
    
    # Create Table on slide
    tbl_shape = s10.shapes.add_table(9, 4, Inches(1.0), Inches(1.9), Inches(6.1), Inches(4.7))
    tbl = tbl_shape.table
    tbl.columns[0].width = Inches(2.2)
    tbl.columns[1].width = Inches(1.3)
    tbl.columns[2].width = Inches(1.1)
    tbl.columns[3].width = Inches(1.5)

    headers = ["Pipeline Stage", "Workload Scale", "Mean Latency", "Throughput"]
    for j, h in enumerate(headers):
        cell = tbl.cell(0, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_HEADER_BG
        p = cell.text_frame.paragraphs[0]
        p.text = h
        p.font.size = Pt(8.5)
        p.font.bold = True
        p.font.color.rgb = COLOR_ACCENT

    b_data = [
        ("Corpus Ingestion", "201,733 CSV rows", "1,380 ms", "146,180 rec/s"),
        ("ST-DBSCAN Hotspots", "5,000 incidents", "82 ms", "60.9 queries/s"),
        ("Louvain Modularity", "250 nodes, 620 edges", "46 ms", "21.7 solves/s"),
        ("Holt's Forecasting", "156 weekly bins", "4.2 ms", "238 forecasts/s"),
        ("TF-IDF MO Search", "12,000 narratives", "31 ms", "387 searches/s"),
        ("Edge Vision (COCO)", "1080p @ 30 FPS", "42 ms/frame", "23.8 FPS (WebGL)"),
        ("Haversine Dispatch", "500 patrol units", "3.8 ms", "263 disp/s"),
        ("SHA-256 Verification", "10,000 audit blocks", "18 ms", "555,550 h/s")
    ]
    for i, row in enumerate(b_data):
        for j, val in enumerate(row):
            cell = tbl.cell(i+1, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = COLOR_CARD if i % 2 == 0 else COLOR_BG
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.size = Pt(8)
            p.font.color.rgb = COLOR_TEXT

    # Right: Forecast Accuracy & Testing Protocol
    create_card(s10, Inches(7.6), Inches(1.35), Inches(4.933), Inches(5.5), title="TESTING PROTOCOL & PREDICTIVE ACCURACY")
    tb_val = s10.shapes.add_textbox(Inches(7.85), Inches(1.95), Inches(4.45), Inches(4.7))
    tf_val = tb_val.text_frame
    tf_val.word_wrap = True

    val_points = [
        ("Strict Holdout Evaluation", "80/20 train/test split across 156 weeks of historical incident data to evaluate Holt's forecast models without data leakage."),
        ("District Forecast Accuracy", "Bengaluru Central achieved RMSE of 3.08 (MAPE: 6.2%); Mysuru City achieved RMSE of 1.32; Belagavi achieved RMSE of 1.32."),
        ("Internal Stress Testing", "Simulated simultaneous dispatch of 50 SOS calls across 500 mobile patrol units; average end-to-end response remained under 28 ms."),
        ("Honest Testing Boundary", "Validation was conducted by the student engineering team in a controlled environment. User testing with real uniformed officers remains planned for the field pilot stage.")
    ]
    for idx, (title, desc) in enumerate(val_points):
        p = tf_val.paragraphs[0] if idx == 0 else tf_val.add_paragraph()
        p.text = f"[TEST] {title}"
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = COLOR_ACCENT
        p.font.name = "Arial"

        p_d = tf_val.add_paragraph()
        p_d.text = desc
        p_d.font.size = Pt(8.8)
        p_d.font.color.rgb = COLOR_MUTED
        p_d.font.name = "Arial"
        p_d.space_after = Pt(8)

    add_footer(s10, 10)

    # =========================================================================
    # SLIDE 11: COMPETITOR & COMPARATIVE ANALYSIS
    # =========================================================================
    s11 = prs.slides.add_slide(blank_layout)
    set_slide_background(s11)
    add_header(s11, "Comparative Evaluation: CrimeLens vs. Existing Industry Paradigms")

    tbl_shape2 = s11.shapes.add_table(7, 5, Inches(0.8), Inches(1.4), Inches(11.733), Inches(5.4))
    tbl2 = tbl_shape2.table
    tbl2.columns[0].width = Inches(2.2)
    tbl2.columns[1].width = Inches(2.3)
    tbl2.columns[2].width = Inches(2.3)
    tbl2.columns[3].width = Inches(2.2)
    tbl2.columns[4].width = Inches(2.733)

    c_headers = ["Evaluation Metric", "Legacy CCTNS Portals", "Commercial GIS Tools", "Standalone ML Scripts", "CrimeLens Platform"]
    for j, h in enumerate(c_headers):
        cell = tbl2.cell(0, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_HEADER_BG
        p = cell.text_frame.paragraphs[0]
        p.text = h
        p.font.size = Pt(9.5)
        p.font.bold = True
        p.font.color.rgb = COLOR_ACCENT

    c_rows = [
        ("Data Plane", "Static SQL FIR entry", "Manual shapefile ingest", "Static CSV notebooks", "Live 201k KSP corpus + CSV ingest"),
        ("Spatio-Temporal Analysis", "Tabular precinct queries", "Kernel density heatmaps", "Standard 2D DBSCAN", "ST-DBSCAN (dual eps + time window)"),
        ("Syndicate Detection", "Manual name matching", "Not supported", "Basic adjacency graphs", "Louvain modularity + kingpin centrality"),
        ("Temporal Forecasting", "Historical bar charts", "Linear extrapolation", "Uncalibrated ARIMA", "Holt double-exponential smoothing"),
        ("Tactical Edge Deployment", "Desktop terminal only", "Specialist GIS workstation", "Python terminal", "Offline-first Expo mobile app + SOS"),
        ("Audit & Governance", "Mutable SQL table logs", "Proprietary event logs", "No audit mechanism", "SHA-256 cryptographic hash-chain ledger")
    ]
    for i, row in enumerate(c_rows):
        for j, val in enumerate(row):
            cell = tbl2.cell(i+1, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = COLOR_CARD if i % 2 == 0 else COLOR_BG
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.size = Pt(8.8)
            p.font.color.rgb = COLOR_SUCCESS if j == 4 else COLOR_TEXT

    add_footer(s11, 11)

    # =========================================================================
    # SLIDE 12: BUSINESS MODEL CANVAS & SUSTAINABILITY
    # =========================================================================
    s12 = prs.slides.add_slide(blank_layout)
    set_slide_background(s12)
    add_header(s12, "Strategic Viability: Business Model Canvas & Operational Sustainability")

    bmc_boxes = [
        ("KEY PARTNERS", "• State Police Computer Wings\n• NCRB & MHA Digital Initiatives\n• Police Telecom & Wireless Depts\n• Cloud Infra (Zoho Catalyst / NIC)", Inches(0.8), Inches(1.35), Inches(2.25), Inches(3.6)),
        ("KEY ACTIVITIES", "• Algorithmic continuous training\n• CCTNS schema synchronization\n• Mobile field app upgrades\n• Cryptographic audit validation", Inches(3.15), Inches(1.35), Inches(2.25), Inches(1.75)),
        ("KEY RESOURCES", "• 201k normalized KSP corpus\n• Graph & clustering IP code\n• Edge CV MobileNet weights\n• Specialized technical team", Inches(3.15), Inches(3.2), Inches(2.25), Inches(1.75)),
        ("VALUE PROPOSITIONS", "• Proactive predictive patrolling reduces blindspots\n• Cross-district syndicate link discovery\n• Courtroom-admissible SHA-256 evidence chain\n• Frontline officer safety via 3.8ms SOS dispatch", Inches(5.5), Inches(1.35), Inches(2.5), Inches(3.6)),
        ("CUSTOMER RELATIONSHIPS", "• Dedicated state technical SLA\n• Tiered 4-role user training\n• Transparent explainability\n• Zero-vendor-lockin open standards", Inches(8.1), Inches(1.35), Inches(2.15), Inches(1.75)),
        ("CHANNELS", "• Secure Police VPN / Intranet\n• Mobile MDM deployment\n• State Police Command Center consoles", Inches(8.1), Inches(3.2), Inches(2.15), Inches(1.75)),
        ("CUSTOMER SEGMENTS", "• State Police Departments (DGP/SP)\n• City Police Commissionerates\n• Specialized Crime Investigation Depts (CID)\n• Women & Child Safety Wings (Akka Pade)", Inches(10.35), Inches(1.35), Inches(2.183), Inches(3.6)),
        ("COST STRUCTURE", "• Serverless cloud compute & database hosting (Catalyst / NIC)  • Edge CCTV hardware integration  • Ongoing maintenance & security pen-testing", Inches(0.8), Inches(5.1), Inches(5.8), Inches(1.75)),
        ("SUSTAINABILITY & REVENUE MODEL", "• State Gov Annual SaaS / Maintenance Contract  • Grant-in-aid through Police Modernization Schemes (MHA MPF)  • Open academic research tier", Inches(6.8), Inches(5.1), Inches(5.733), Inches(1.75))
    ]

    for title, text, left, top, width, height in bmc_boxes:
        create_card(s12, left, top, width, height, title=title)
        tb_bmc = s12.shapes.add_textbox(left + Inches(0.12), top + Inches(0.4), width - Inches(0.24), height - Inches(0.45))
        tf_bmc = tb_bmc.text_frame
        tf_bmc.word_wrap = True
        p = tf_bmc.paragraphs[0]
        p.text = text
        p.font.size = Pt(8.2)
        p.font.color.rgb = COLOR_TEXT
        p.font.name = "Arial"

    add_footer(s12, 12)

    # =========================================================================
    # SLIDE 13: MARKET SIZE (TAM / SAM / SOM) & ASSUMPTIONS
    # =========================================================================
    s13 = prs.slides.add_slide(blank_layout)
    set_slide_background(s13)
    add_header(s13, "Market Potential: TAM, SAM, SOM & Modernization Budget Mapping")

    market_cards = [
        ("TAM: TOTAL ADDRESSABLE MARKET",
         "$3.2 Billion (Global Law Enforcement & Predictive Policing Software)",
         "Covers global municipal, state, and federal law enforcement agencies adopting AI crime analytics, real-time dispatch systems, and video intelligence solutions (CAGR: 14.8%).",
         Inches(0.8), COLOR_PRIMARY),
        
        ("SAM: SERVICEABLE ADDRESSABLE MARKET",
         "Rs 850 Crore (~$102 Million - Indian Police Modernization)",
         "Targeted allocation under Ministry of Home Affairs (MHA) Modernization of Police Forces (MPF) scheme across 28 states and 8 union territories for CCTNS 2.0 and smart policing.",
         Inches(4.8), COLOR_ACCENT),
        
        ("SOM: SERVICEABLE OBTAINABLE MARKET",
         "Rs 28 Crore (~$3.4 Million - Southern States Initial Phase)",
         "Deployment across Karnataka State Police (1,100+ stations) and neighboring southern commissionerates (Tamil Nadu, Telangana) over a 3-year phased implementation roadmap.",
         Inches(8.8), COLOR_SUCCESS)
    ]

    for title, val, desc, left, col in market_cards:
        create_card(s13, left, Inches(1.35), Inches(3.733), Inches(5.5), title=title, border_color=col)
        tb_m = s13.shapes.add_textbox(left + Inches(0.15), Inches(2.0), Inches(3.433), Inches(4.6))
        tf_m = tb_m.text_frame
        tf_m.word_wrap = True

        p_val = tf_m.paragraphs[0]
        p_val.text = val
        p_val.font.size = Pt(14)
        p_val.font.bold = True
        p_val.font.color.rgb = col
        p_val.font.name = "Arial"
        p_val.space_after = Pt(12)

        p_desc = tf_m.add_paragraph()
        p_desc.text = desc
        p_desc.font.size = Pt(9.5)
        p_desc.font.color.rgb = COLOR_TEXT
        p_desc.font.name = "Arial"
        p_desc.space_after = Pt(14)

        p_assump = tf_m.add_paragraph()
        p_assump.text = "TRANSPARENT ASSUMPTIONS:\n• Sourced from NCRB and MHA Annual Reports\n• Average IT spend: Rs 2.5-4 Lakhs per police station/yr\n• Academic prototype requires institutional grant to enter pilot procurement"
        p_assump.font.size = Pt(8.2)
        p_assump.font.color.rgb = COLOR_MUTED
        p_assump.font.name = "Arial"

    add_footer(s13, 13)

    # =========================================================================
    # SLIDE 14: RISKS, ETHICAL BOUNDARIES & MITIGATION STRATEGY
    # =========================================================================
    s14 = prs.slides.add_slide(blank_layout)
    set_slide_background(s14)
    add_header(s14, "Governance & Ethics: Risk Taxonomy & Mitigation Framework")

    risks = [
        ("Algorithmic Feedback Loops",
         "Frequent patrolling in historical hotspots may cause disproportionate arrest reporting, artificially inflating future risk scores.",
         "Station Anomaly Z-Score baseline tracking and population-normalized bias fairness indicators flag skewed patrolling patterns."),
        
        ("Data Quality & Privacy Exposure",
         "Ingesting real police records risks exposing sensitive victim identities, FIR details, and protected juvenile records.",
         "Complete anonymization of names and synthetic micro-incident generation for training; on-premise encrypted storage for production."),
        
        ("Frontline Adoption & Operational Resistance",
         "Field officers may resist complex analytical tools or mobile check-in systems during fast-paced tactical duty.",
         "One-tap SOS dispatch, minimal-friction beat checklists, and native bilingual interface (English + Kannada) tailored for beat staff."),
        
        ("Constitutional & Judicial Scrutiny",
         "AI predictive recommendations risk rejection in court as unverified hearsay or algorithmic bias without proof.",
         "SHA-256 recursive hash-chain ledger guarantees mathematical tamper-evidence; AI acts strictly as decision support (human-in-the-loop).")
    ]

    for idx, (rtitle, rdesc, rmit) in enumerate(risks):
        r = idx // 2
        c = idx % 2
        cx = Inches(0.8 + c * 5.95)
        cy = Inches(1.35 + r * 2.8)

        create_card(s14, cx, cy, Inches(5.75), Inches(2.65), title=f"RISK {idx+1}: {rtitle}", border_color=COLOR_WARN)
        tb_r = s14.shapes.add_textbox(cx + Inches(0.15), cy + Inches(0.55), Inches(5.45), Inches(2.0))
        tf_r = tb_r.text_frame
        tf_r.word_wrap = True

        p_d = tf_r.paragraphs[0]
        p_d.text = f"Threat Description: {rdesc}"
        p_d.font.size = Pt(8.8)
        p_d.font.color.rgb = COLOR_TEXT
        p_d.font.name = "Arial"
        p_d.space_after = Pt(6)

        p_m = tf_r.add_paragraph()
        p_m.text = f"Mitigation Strategy: {rmit}"
        p_m.font.size = Pt(8.8)
        p_m.font.bold = True
        p_m.font.color.rgb = COLOR_SUCCESS
        p_m.font.name = "Arial"

    add_footer(s14, 14)

    # =========================================================================
    # SLIDE 15: DISSEMINATION, CERTIFICATIONS & SDG MAPPING
    # =========================================================================
    s15 = prs.slides.add_slide(blank_layout)
    set_slide_background(s15)
    add_header(s15, "Dissemination, Intellectual Property & UN SDG Alignment")

    # Left: Dissemination & IP
    create_card(s15, Inches(0.8), Inches(1.35), Inches(5.7), Inches(5.5), title="ACADEMIC DISSEMINATION & IP STATUS")
    tb_ip = s15.shapes.add_textbox(Inches(1.05), Inches(1.95), Inches(5.2), Inches(4.7))
    tf_ip = tb_ip.text_frame
    tf_ip.word_wrap = True

    ip_items = [
        ("Research Manuscript Preparation", "Draft titled 'CrimeLens: An AI-First Crime Intelligence Operating System with Multi-Agent Spatio-Temporal Analytics and Cryptographic Auditing' targeted for IEEE / Springer conference submission."),
        ("Software Copyright Filing (Planned)", "Source code architecture and mathematical pipeline scheduled for filing with the Indian Copyright Office upon final review completion."),
        ("Hackathon & Competition Recognition", "Presented and benchmarked at KSP Datathon 2026 challenge track for state police modernization."),
        ("Open-Source Academic Release", "Core algorithms (pure TypeScript ST-DBSCAN and Holt forecaster) prepared for open academic dissemination under MIT License.")
    ]
    for idx, (title, desc) in enumerate(ip_items):
        p = tf_ip.paragraphs[0] if idx == 0 else tf_ip.add_paragraph()
        p.text = f"[PUB] {title}"
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = COLOR_ACCENT
        p.font.name = "Arial"

        p_d = tf_ip.add_paragraph()
        p_d.text = desc
        p_d.font.size = Pt(8.5)
        p_d.font.color.rgb = COLOR_MUTED
        p_d.font.name = "Arial"
        p_d.space_after = Pt(6)

    # Right: SDG Mapping
    create_card(s15, Inches(6.8), Inches(1.35), Inches(5.733), Inches(5.5), title="UN SUSTAINABLE DEVELOPMENT GOALS (SDG) ALIGNMENT")
    tb_sdg = s15.shapes.add_textbox(Inches(7.05), Inches(1.95), Inches(5.2), Inches(4.7))
    tf_sdg = tb_sdg.text_frame
    tf_sdg.word_wrap = True

    sdgs = [
        ("SDG 16: Peace, Justice & Strong Institutions",
         "Target 16.1 (Reduce violence): Predictive patrolling directly targets critical dark-spot corridors.\nTarget 16.6 (Transparent institutions): SHA-256 tamper-evident governance eliminates evidence fabrication."),
        
        ("SDG 11: Sustainable Cities & Communities",
         "Target 11.7 (Safe public spaces): Real-time Akka Pade fleet tracking and 3.8ms emergency SOS dispatch ensure safe transit for women, children, and vulnerable citizens."),
        
        ("SDG 9: Industry, Innovation & Infrastructure",
         "Target 9.5 (Enhance scientific research): Modernizes municipal policing infrastructure from legacy paper diaries to state-of-the-art edge AI and graph network analytics.")
    ]
    for idx, (title, desc) in enumerate(sdgs):
        p = tf_sdg.paragraphs[0] if idx == 0 else tf_sdg.add_paragraph()
        p.text = f"[SDG] {title}"
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = COLOR_SUCCESS
        p.font.name = "Arial"

        p_d = tf_sdg.add_paragraph()
        p_d.text = desc
        p_d.font.size = Pt(8.5)
        p_d.font.color.rgb = COLOR_MUTED
        p_d.font.name = "Arial"
        p_d.space_after = Pt(8)

    add_footer(s15, 15)

    # =========================================================================
    # SLIDE 16: TEAM CONTRIBUTIONS & CONCLUSION / THANK YOU
    # =========================================================================
    s16 = prs.slides.add_slide(blank_layout)
    set_slide_background(s16)
    add_header(s16, "Project Summary: Individual Team Contributions & Concluding Defense")

    # Team Contribution Cards (Left)
    create_card(s16, Inches(0.8), Inches(1.35), Inches(6.5), Inches(5.5), title="INDIVIDUAL TEAM ROLES & CONTRIBUTIONS")
    tb_tc = s16.shapes.add_textbox(Inches(1.05), Inches(1.95), Inches(6.0), Inches(4.7))
    tf_tc = tb_tc.text_frame
    tf_tc.word_wrap = True

    tc_data = [
        ("Parameshwaran P (722824104166)", "• Designed 201k KSP CSV ingestion pipeline & schema normalization\n• Implemented Louvain modularity algorithm & bipartite entity projection\n• Managed backend Express/TypeScript REST APIs & WebSocket alert bus"),
        
        ("Ram Vishnu P (722824104186)", "• Developed ST-DBSCAN spatio-temporal clustering engine & intensity scoring\n• Implemented Holt's double-exponential forecasting equations & RMSE benchmark\n• Built time-slot risk matrix & station-level anomaly Z-score calculators"),
        
        ("Ravishankar AG (722824104191)", "• Engineered React 18 / Vite web command console with TailwindCSS\n• Implemented Leaflet geospatial mapping & React-Force-Graph-2D visualization\n• Designed explainability drawers, supervisor review queue & Recharts analytics"),
        
        ("Rishivardhan J (722824104196)", "• Built cross-platform React Native / Expo SDK 54 field-officer mobile application\n• Formulated SHA-256 recursive hash-chain tamper-evident audit ledger\n• Engineered Haversine nearest-unit dispatch optimizer & SQLite offline syncing")
    ]
    for idx, (mem, tasks) in enumerate(tc_data):
        p = tf_tc.paragraphs[0] if idx == 0 else tf_tc.add_paragraph()
        p.text = mem
        p.font.size = Pt(9.5)
        p.font.bold = True
        p.font.color.rgb = COLOR_ACCENT
        p.font.name = "Arial"

        p_t = tf_tc.add_paragraph()
        p_t.text = tasks
        p_t.font.size = Pt(8)
        p_t.font.color.rgb = COLOR_TEXT
        p_t.font.name = "Arial"
        p_t.space_after = Pt(4)

    # Conclusion / Thank You (Right)
    create_card(s16, Inches(7.6), Inches(1.35), Inches(4.933), Inches(5.5), title="CONCLUSION & ACADEMIC DEFENSE")
    tb_thk = s16.shapes.add_textbox(Inches(7.85), Inches(1.95), Inches(4.45), Inches(4.7))
    tf_thk = tb_thk.text_frame
    tf_thk.word_wrap = True

    p_c1 = tf_thk.paragraphs[0]
    p_c1.text = "CRIMELENS CORE VALUE DELIVERED:"
    p_c1.font.size = Pt(10)
    p_c1.font.bold = True
    p_c1.font.color.rgb = COLOR_SUCCESS
    p_c1.font.name = "Arial"

    p_c2 = tf_thk.add_paragraph()
    p_c2.text = "CrimeLens demonstrates that complex state police informatics can successfully transition from passive bureaucratic record-keeping into a proactive, mathematically grounded, and human-in-the-loop operational intelligence platform."
    p_c2.font.size = Pt(9.5)
    p_c2.font.color.rgb = COLOR_MUTED
    p_c2.font.name = "Arial"
    p_c2.space_after = Pt(16)

    p_ty = tf_thk.add_paragraph()
    p_ty.text = "THANK YOU!"
    p_ty.font.size = Pt(24)
    p_ty.font.bold = True
    p_ty.font.color.rgb = COLOR_ACCENT
    p_ty.font.name = "Arial"
    p_ty.space_after = Pt(8)

    p_qa = tf_thk.add_paragraph()
    p_qa.text = "Open for Questions, Evaluation & Technical Demonstration\n\nBatch 17 | Department of Computer Science and Engineering\nSri Eshwar College of Engineering"
    p_qa.font.size = Pt(9.5)
    p_qa.font.color.rgb = COLOR_TEXT
    p_qa.font.name = "Arial"

    add_footer(s16, 16)

    output_path = r"d:\projects\CrimeLens\CrimeLens_Final_Review_Presentation.pptx"
    prs.save(output_path)
    print(f"Presentation generated successfully: {output_path}")

if __name__ == "__main__":
    build_presentation()

