"""
CrimeLens - Official Final Review Presentation Generator
Recreating exactly the Sri Eshwar Innovation Studio template style and 16-slide sequence.
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

    # Exact Institutional Palette from Sample Template
    COLOR_WHITE = RGBColor(255, 255, 255)
    COLOR_NAVY_DARK = RGBColor(10, 25, 68)       # #0A1944 Main Dark Navy
    COLOR_NAVY_BANNER = RGBColor(11, 27, 74)     # Dark Blue Banner
    COLOR_ROYAL_BLUE = RGBColor(24, 76, 120)     # #184C78 Institutional Blue
    COLOR_TEXT_DARK = RGBColor(24, 30, 42)       # Heading text dark
    COLOR_TEXT_RED = RGBColor(190, 18, 30)       # Template Subheading Red
    COLOR_TEXT_MUTED = RGBColor(100, 116, 139)   # Subtitle gray
    COLOR_BORDER_LIGHT = RGBColor(203, 213, 225) # Light card outline
    COLOR_CARD_BG = RGBColor(238, 244, 255)      # Light blue soft card background
    COLOR_RED_CARD_BG = RGBColor(254, 242, 242)  # Light pinkish red before card
    COLOR_RED_BORDER = RGBColor(252, 165, 165)
    COLOR_YELLOW_CARD_BG = RGBColor(254, 252, 232)
    COLOR_YELLOW_BORDER = RGBColor(254, 240, 138)
    COLOR_GREEN_CARD_BG = RGBColor(240, 253, 244)
    COLOR_GREEN_BORDER = RGBColor(187, 247, 208)
    COLOR_YELLOW_ACCENT = RGBColor(250, 204, 21) # #FACC15 Accent Yellow
    COLOR_ORANGE_ACCENT = RGBColor(249, 115, 22) # Innovation Studio Orange

    def add_top_brand(slide):
        # Top-left institutional text block
        tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.28), Inches(4.5), Inches(0.55))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = "Sri Eshwar College of Engineering"
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = COLOR_ROYAL_BLUE
        p.font.name = "Calibri"

        p_sub = tf.add_paragraph()
        p_sub.text = "Coimbatore | Autonomous Institution | NAAC 'A' Grade • NBA • NIRF"
        p_sub.font.size = Pt(7.5)
        p_sub.font.color.rgb = COLOR_TEXT_MUTED
        p_sub.font.name = "Calibri"

        # Top-right Innovation Studio branding block
        tb_r = slide.shapes.add_textbox(Inches(10.2), Inches(0.28), Inches(2.6), Inches(0.55))
        tf_r = tb_r.text_frame
        tf_r.word_wrap = True
        tf_r.margin_left = tf_r.margin_top = tf_r.margin_right = tf_r.margin_bottom = 0
        p_r = tf_r.paragraphs[0]
        p_r.text = "SRI ESHWAR"
        p_r.font.size = Pt(8.5)
        p_r.font.bold = True
        p_r.font.color.rgb = COLOR_TEXT_DARK
        p_r.font.name = "Calibri"
        p_r.alignment = PP_ALIGN.RIGHT

        p_r2 = tf_r.add_paragraph()
        p_r2.text = "Innovation Studio"
        p_r2.font.size = Pt(15)
        p_r2.font.bold = True
        p_r2.font.color.rgb = COLOR_ORANGE_ACCENT
        p_r2.font.name = "Calibri"
        p_r2.alignment = PP_ALIGN.RIGHT

    def add_page_header(slide, title_text, subtitle_text=""):
        add_top_brand(slide)

        # Title
        tb_t = slide.shapes.add_textbox(Inches(0.6), Inches(0.95), Inches(12.0), Inches(0.55))
        tf_t = tb_t.text_frame
        tf_t.word_wrap = True
        tf_t.margin_left = tf_t.margin_top = tf_t.margin_right = tf_t.margin_bottom = 0
        p = tf_t.paragraphs[0]
        p.text = title_text
        p.font.size = Pt(24)
        p.font.bold = True
        p.font.color.rgb = COLOR_TEXT_DARK
        p.font.name = "Calibri"

        # Subtitle (Red italic, matching template)
        if subtitle_text:
            tb_s = slide.shapes.add_textbox(Inches(0.6), Inches(1.5), Inches(12.0), Inches(0.35))
            tf_s = tb_s.text_frame
            tf_s.word_wrap = True
            tf_s.margin_left = tf_s.margin_top = tf_s.margin_right = tf_s.margin_bottom = 0
            p_s = tf_s.paragraphs[0]
            p_s.text = subtitle_text
            p_s.font.size = Pt(11)
            p_s.font.italic = True
            p_s.font.color.rgb = COLOR_TEXT_RED
            p_s.font.name = "Calibri"

    def add_footer(slide, page_num):
        tb = slide.shapes.add_textbox(Inches(0.6), Inches(7.1), Inches(12.133), Inches(0.3))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = "Sri Eshwar College of Engineering  •  Design Thinking & Innovation  •  Department of CSE"
        p.font.size = Pt(8.5)
        p.font.color.rgb = COLOR_TEXT_MUTED
        p.font.name = "Calibri"

        p_num = tf.add_paragraph()
        p_num.text = str(page_num)
        p_num.alignment = PP_ALIGN.RIGHT
        p_num.font.size = Pt(8.5)
        p_num.font.color.rgb = COLOR_TEXT_MUTED
        p_num.font.name = "Calibri"

    def create_card(slide, left, top, width, height, bg_color=COLOR_WHITE, border_color=COLOR_BORDER_LIGHT, radius=None):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color
        card.line.color.rgb = border_color
        card.line.width = Pt(1)
        return card

    # =========================================================================
    # SLIDE 1: TITLE SLIDE (Exact matching of template layout and colors)
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    add_top_brand(s1)

    # Centered Department & Course Title
    tb_c = s1.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(11.333), Inches(1.2))
    tf_c = tb_c.text_frame
    tf_c.word_wrap = True
    tf_c.margin_left = tf_c.margin_top = tf_c.margin_right = tf_c.margin_bottom = 0

    p_dept = tf_c.paragraphs[0]
    p_dept.text = "Department of Computer Science and Engineering"
    p_dept.font.size = Pt(21)
    p_dept.font.bold = True
    p_dept.font.color.rgb = COLOR_TEXT_RED
    p_dept.font.name = "Calibri"
    p_dept.alignment = PP_ALIGN.CENTER

    p_course = tf_c.add_paragraph()
    p_course.text = "U23CS652 – Innovative / Multidisciplinary Project"
    p_course.font.size = Pt(20)
    p_course.font.bold = True
    p_course.font.color.rgb = COLOR_ROYAL_BLUE
    p_course.font.name = "Calibri"
    p_course.alignment = PP_ALIGN.CENTER
    p_course.space_before = Pt(4)

    # Final Review Banner
    p_rev = tf_c.add_paragraph()
    p_rev.text = "Final Review"
    p_rev.font.size = Pt(30)
    p_rev.font.bold = True
    p_rev.font.color.rgb = COLOR_NAVY_DARK
    p_rev.font.name = "Calibri"
    p_rev.alignment = PP_ALIGN.CENTER
    p_rev.space_before = Pt(10)

    p_proj = tf_c.add_paragraph()
    p_proj.text = "CrimeLens: Crime Analytics System using AI"
    p_proj.font.size = Pt(22)
    p_proj.font.bold = True
    p_proj.font.color.rgb = RGBColor(133, 77, 14) # Olive/Gold Project Title
    p_proj.font.name = "Calibri"
    p_proj.alignment = PP_ALIGN.CENTER
    p_proj.space_before = Pt(8)

    # Bottom Two Dark Navy Boxes (Faculty Mentor & Team)
    # Left Card: Faculty Mentor
    create_card(s1, Inches(0.8), Inches(4.7), Inches(5.5), Inches(2.2), bg_color=COLOR_NAVY_DARK, border_color=COLOR_NAVY_DARK)
    tb_fm = s1.shapes.add_textbox(Inches(1.1), Inches(4.9), Inches(4.9), Inches(1.8))
    tf_fm = tb_fm.text_frame
    tf_fm.word_wrap = True
    tf_fm.margin_left = tf_fm.margin_top = tf_fm.margin_right = tf_fm.margin_bottom = 0

    p_fm0 = tf_fm.paragraphs[0]
    p_fm0.text = "FACULTY MENTOR"
    p_fm0.font.size = Pt(11)
    p_fm0.font.bold = True
    p_fm0.font.color.rgb = COLOR_WHITE
    p_fm0.font.name = "Calibri"
    p_fm0.space_after = Pt(12)

    p_fm1 = tf_fm.add_paragraph()
    p_fm1.text = "Name:  Dr. S. Yasotha, M.E., Ph.D."
    p_fm1.font.size = Pt(11.5)
    p_fm1.font.color.rgb = COLOR_WHITE
    p_fm1.font.name = "Calibri"
    p_fm1.space_after = Pt(6)

    p_fm2 = tf_fm.add_paragraph()
    p_fm2.text = "Designation:  Associate Professor"
    p_fm2.font.size = Pt(11.5)
    p_fm2.font.color.rgb = COLOR_WHITE
    p_fm2.font.name = "Calibri"
    p_fm2.space_after = Pt(6)

    p_fm3 = tf_fm.add_paragraph()
    p_fm3.text = "Department:  Computer Science and Engineering (CYS)"
    p_fm3.font.size = Pt(11.5)
    p_fm3.font.color.rgb = COLOR_WHITE
    p_fm3.font.name = "Calibri"

    # Right Card: Team
    create_card(s1, Inches(6.8), Inches(4.7), Inches(5.7), Inches(2.2), bg_color=COLOR_NAVY_DARK, border_color=COLOR_NAVY_DARK)
    tb_tm = s1.shapes.add_textbox(Inches(7.1), Inches(4.85), Inches(5.1), Inches(1.9))
    tf_tm = tb_tm.text_frame
    tf_tm.word_wrap = True
    tf_tm.margin_left = tf_tm.margin_top = tf_tm.margin_right = tf_tm.margin_bottom = 0

    p_tm0 = tf_tm.paragraphs[0]
    p_tm0.text = "TEAM"
    p_tm0.font.size = Pt(11)
    p_tm0.font.bold = True
    p_tm0.font.color.rgb = COLOR_WHITE
    p_tm0.font.name = "Calibri"
    p_tm0.space_after = Pt(6)

    p_tm1 = tf_tm.add_paragraph()
    p_tm1.text = "Team Name / Batch No:  Batch 17"
    p_tm1.font.size = Pt(10.5)
    p_tm1.font.color.rgb = COLOR_WHITE
    p_tm1.font.name = "Calibri"
    p_tm1.space_after = Pt(3)

    p_tm2 = tf_tm.add_paragraph()
    p_tm2.text = "Year & Dept / Section:  III Year, CSE"
    p_tm2.font.size = Pt(10.5)
    p_tm2.font.color.rgb = COLOR_WHITE
    p_tm2.font.name = "Calibri"
    p_tm2.space_after = Pt(4)

    team_list = [
        "1. Parameshwaran P    -  722824104166",
        "2. Ram Vishnu P          -  722824104186",
        "3. Ravishankar AG       -  722824104191",
        "4. Rishivardhan J        -  722824104196"
    ]
    p_m_label = tf_tm.add_paragraph()
    p_m_label.text = "Members: " + team_list[0]
    p_m_label.font.size = Pt(10)
    p_m_label.font.color.rgb = COLOR_WHITE
    p_m_label.font.name = "Calibri"

    for member_str in team_list[1:]:
        p_mem = tf_tm.add_paragraph()
        p_mem.text = "                 " + member_str
        p_mem.font.size = Pt(10)
        p_mem.font.color.rgb = COLOR_WHITE
        p_mem.font.name = "Calibri"

    add_footer(s1, 1)

    # =========================================================================
    # SLIDE 2: PROBLEM & SOLUTION STORY
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    add_page_header(s2, "Problem & Solution Story", "Tell it as a story — your user before, and after your product. Link a short video of them using it.")

    # Left Container: BEFORE (Light Pinkish Red Background)
    create_card(s2, Inches(0.6), Inches(1.9), Inches(5.8), Inches(3.2), bg_color=COLOR_RED_CARD_BG, border_color=COLOR_RED_BORDER)
    tb_b_title = s2.shapes.add_textbox(Inches(0.8), Inches(2.05), Inches(5.4), Inches(0.3))
    p = tb_b_title.text_frame.paragraphs[0]
    p.text = "BEFORE  —  A DAY WITHOUT OUR PRODUCT"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_TEXT_RED
    p.alignment = PP_ALIGN.CENTER

    # 3 Small white boxes inside BEFORE
    b_steps = [
        ("1", "Siloed Station Records", "1,100+ police stations act as data islands. Cross-district crime syndicates go unnoticed."),
        ("2", "Rigid Calendar Patrols", "Units follow uniform calendar routes; critical dark spots remain unpoliced."),
        ("3", "Tamper-Prone Audits", "Manual diary logs leave digital evidence vulnerable to defense tampering claims in court.")
    ]
    for i, (num, h, t) in enumerate(b_steps):
        bx = Inches(0.78 + i * 1.8)
        create_card(s2, bx, Inches(2.45), Inches(1.68), Inches(1.8), bg_color=COLOR_WHITE, border_color=COLOR_BORDER_LIGHT)
        # Number badge
        num_circ = s2.shapes.add_shape(MSO_SHAPE.OVAL, bx + Inches(0.1), Inches(2.55), Inches(0.28), Inches(0.28))
        num_circ.fill.solid()
        num_circ.fill.fore_color.rgb = COLOR_NAVY_DARK
        num_circ.line.fill.background()
        p = num_circ.text_frame.paragraphs[0]
        p.text = num
        p.font.size = Pt(9)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE
        p.alignment = PP_ALIGN.CENTER

        tb_box = s2.shapes.add_textbox(bx + Inches(0.08), Inches(2.95), Inches(1.52), Inches(1.2))
        tf = tb_box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p_h = tf.paragraphs[0]
        p_h.text = h
        p_h.font.size = Pt(9)
        p_h.font.bold = True
        p_h.font.color.rgb = COLOR_TEXT_DARK
        p_t = tf.add_paragraph()
        p_t.text = t
        p_t.font.size = Pt(7.5)
        p_t.font.color.rgb = COLOR_TEXT_MUTED
        p_t.space_before = Pt(2)

    tb_b_footer = s2.shapes.add_textbox(Inches(0.8), Inches(4.35), Inches(5.4), Inches(0.65))
    p = tb_b_footer.text_frame.paragraphs[0]
    p.text = "Police analysts receive district reports with disconnected spreadsheets and late submissions. Syndicates are missed, patrol deployments are reactive, and public trust in crime prevention drops."
    p.font.size = Pt(8.5)
    p.font.italic = True
    p.font.color.rgb = COLOR_TEXT_RED
    p.alignment = PP_ALIGN.CENTER

    # Right Container: AFTER (Light Blue Background)
    create_card(s2, Inches(6.6), Inches(1.9), Inches(6.1), Inches(3.2), bg_color=COLOR_CARD_BG, border_color=RGBColor(191, 219, 254))
    tb_a_title = s2.shapes.add_textbox(Inches(6.8), Inches(2.05), Inches(5.7), Inches(0.3))
    p = tb_a_title.text_frame.paragraphs[0]
    p.text = "AFTER  —  A DAY WITH OUR PRODUCT"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_ROYAL_BLUE
    p.alignment = PP_ALIGN.CENTER

    a_steps = [
        ("1", "Statewide Corpus Ingest", "201,733 KSP records mapped into a unified real-time spatial intelligence plane."),
        ("2", "AI Predict & Network", "ST-DBSCAN hotspots + Louvain graphs automatically link gang kingpins and forecast trends."),
        ("3", "Tactical Edge Dispatch", "Haversine dispatch to nearest units; every action sealed in a SHA-256 hash ledger.")
    ]
    for i, (num, h, t) in enumerate(a_steps):
        bx = Inches(6.8 + i * 1.9)
        create_card(s2, bx, Inches(2.45), Inches(1.78), Inches(1.8), bg_color=COLOR_WHITE, border_color=COLOR_BORDER_LIGHT)
        num_circ = s2.shapes.add_shape(MSO_SHAPE.OVAL, bx + Inches(0.1), Inches(2.55), Inches(0.28), Inches(0.28))
        num_circ.fill.solid()
        num_circ.fill.fore_color.rgb = COLOR_ROYAL_BLUE
        num_circ.line.fill.background()
        p = num_circ.text_frame.paragraphs[0]
        p.text = num
        p.font.size = Pt(9)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE
        p.alignment = PP_ALIGN.CENTER

        tb_box = s2.shapes.add_textbox(bx + Inches(0.08), Inches(2.95), Inches(1.62), Inches(1.2))
        tf = tb_box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p_h = tf.paragraphs[0]
        p_h.text = h
        p_h.font.size = Pt(9)
        p_h.font.bold = True
        p_h.font.color.rgb = COLOR_TEXT_DARK
        p_t = tf.add_paragraph()
        p_t.text = t
        p_t.font.size = Pt(7.5)
        p_t.font.color.rgb = COLOR_TEXT_MUTED
        p_t.space_before = Pt(2)

    tb_a_footer = s2.shapes.add_textbox(Inches(6.8), Inches(4.35), Inches(5.7), Inches(0.65))
    p = tb_a_footer.text_frame.paragraphs[0]
    p.text = "Every incident is analyzed against ST-DBSCAN clustering and network graphs. Supervisory staff verify transparent recommendations, while frontline officers receive predictive routes and one-tap SOS safety."
    p.font.size = Pt(8.5)
    p.font.italic = True
    p.font.color.rgb = COLOR_ROYAL_BLUE
    p.alignment = PP_ALIGN.CENTER

    # Middle Arrow
    arr = s2.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(6.38), Inches(3.2), Inches(0.22), Inches(0.3))
    arr.fill.solid()
    arr.fill.fore_color.rgb = COLOR_YELLOW_ACCENT
    arr.line.fill.background()

    # Bottom Row: Customer & User Bar
    create_card(s2, Inches(0.6), Inches(5.25), Inches(5.8), Inches(0.75), bg_color=COLOR_CARD_BG, border_color=RGBColor(219, 234, 254))
    tb_c = s2.shapes.add_textbox(Inches(0.8), Inches(5.32), Inches(5.4), Inches(0.6))
    p = tb_c.text_frame.paragraphs[0]
    p.text = "CUSTOMER / BUYER  —  who pays for it"
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY_DARK
    p_sub = tb_c.text_frame.add_paragraph()
    p_sub.text = "State Police Headquarter & Police Modernization (MHA MPF Scheme)"
    p_sub.font.size = Pt(8.5)
    p_sub.font.italic = True
    p_sub.font.color.rgb = COLOR_TEXT_MUTED

    create_card(s2, Inches(6.6), Inches(5.25), Inches(6.1), Inches(0.75), bg_color=COLOR_GREEN_CARD_BG, border_color=COLOR_GREEN_BORDER)
    tb_u = s2.shapes.add_textbox(Inches(6.8), Inches(5.32), Inches(5.7), Inches(0.6))
    p = tb_u.text_frame.paragraphs[0]
    p.text = "USER / CONSUMER  —  who actually uses it"
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = RGBColor(22, 101, 52)
    p_sub = tb_u.text_frame.add_paragraph()
    p_sub.text = "Command center analysts, District SPs, Station IOs, and frontline patrol officers"
    p_sub.font.size = Pt(8.5)
    p_sub.font.italic = True
    p_sub.font.color.rgb = COLOR_TEXT_MUTED

    # Bottom Dark Banner: Story Video Link
    create_card(s2, Inches(0.6), Inches(6.12), Inches(12.1), Inches(0.85), bg_color=COLOR_NAVY_BANNER, border_color=COLOR_NAVY_BANNER)
    tb_sv = s2.shapes.add_textbox(Inches(1.8), Inches(6.25), Inches(6.5), Inches(0.6))
    p = tb_sv.text_frame.paragraphs[0]
    p.text = "STORY VIDEO"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_YELLOW_ACCENT
    p2 = tb_sv.text_frame.add_paragraph()
    p2.text = "A short video walkthrough demonstrating CrimeLens solving cross-border syndicates and field dispatch."
    p2.font.size = Pt(8.5)
    p2.font.color.rgb = COLOR_WHITE

    # Play Icon Circle
    play_c = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.95), Inches(6.25), Inches(0.55), Inches(0.55))
    play_c.fill.solid()
    play_c.fill.fore_color.rgb = COLOR_WHITE
    play_c.line.fill.background()

    # Link button on right
    create_card(s2, Inches(8.8), Inches(6.25), Inches(3.6), Inches(0.55), bg_color=COLOR_WHITE, border_color=COLOR_YELLOW_ACCENT)
    tb_lnk = s2.shapes.add_textbox(Inches(8.9), Inches(6.32), Inches(3.4), Inches(0.4))
    p = tb_lnk.text_frame.paragraphs[0]
    p.text = "Story video link (CrimeLens 2.5 min Walkthrough)"
    p.font.size = Pt(8.5)
    p.font.color.rgb = COLOR_ROYAL_BLUE
    p.alignment = PP_ALIGN.CENTER

    add_footer(s2, 2)

    # =========================================================================
    # SLIDE 3: TECHNICAL APPROACH
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    add_page_header(s3, "Technical Approach", "Block diagram, the tech stack you used, and the method you followed to build it.")

    # Left Container: System Flow (Matching Layered Diagram Box)
    create_card(s3, Inches(0.6), Inches(1.9), Inches(7.2), Inches(4.9), bg_color=COLOR_CARD_BG, border_color=RGBColor(219, 234, 254))
    tb_bf = s3.shapes.add_textbox(Inches(0.9), Inches(2.05), Inches(6.6), Inches(0.3))
    p = tb_bf.text_frame.paragraphs[0]
    p.text = "BLOCK DIAGRAM / SYSTEM FLOW"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY_DARK

    layers = [
        ("Layer 1 - Data Sources", "12 KSP CSV files (201,733 records) | Synthetic GPS incidents (2,200) | CCTV video feeds"),
        ("Layer 2 - Ingestion & Preparation", "district-loader.ts | IPC/SLL statutory normalizer | Temporal binning | CSV Parser"),
        ("Layer 3 - Storage & Data Models", "PostgreSQL (Drizzle ORM) | In-memory spatial index | SQLite (Field mobile offline)"),
        ("Layer 4 - AI & Analytics Engine", "ST-DBSCAN (eps=600m) | Louvain modularity | Holt's linear trend | TF-IDF Cosine MO"),
        ("Layer 5 - Explainability & Governance", "Recursive SHA-256 hash chain ledger | SHAP-style feature attribution | Anomaly Z-Score"),
        ("Layer 6 - Security & Edge Engine", "JWT 4-tier RBAC | WebAuthn biometric login | TensorFlow.js COCO-SSD on-device CV"),
        ("Layer 7 - Backend API & Telemetry", "Node.js 20 Express 5 REST API | WebSocket ws real-time dispatch | Haversine distance"),
        ("Layer 8 - Presentation (React & Expo)", "React 18 + Vite command console (Leaflet, Force Graph) | Expo SDK 54 mobile field app"),
        ("Layer 9 - Human-in-the-Loop Review", "Supervisor sign-off queue | IO warrant generator | Akka Pade turn-by-turn dispatch")
    ]
    for i, (lh, ld) in enumerate(layers):
        ly = Inches(2.4 + i * 0.47)
        c_layer = create_card(s3, Inches(0.9), ly, Inches(6.6), Inches(0.42), bg_color=COLOR_WHITE, border_color=COLOR_BORDER_LIGHT)
        tb_l = s3.shapes.add_textbox(Inches(0.95), ly + Inches(0.02), Inches(6.5), Inches(0.38))
        tf = tb_l.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = f"{lh}: "
        p.font.size = Pt(8.5)
        p.font.bold = True
        p.font.color.rgb = COLOR_NAVY_DARK
        run = p.add_run()
        run.text = ld
        run.font.size = Pt(7.5)
        run.font.bold = False
        run.font.color.rgb = COLOR_TEXT_MUTED

    # Right Container: Tech Stack Used
    create_card(s3, Inches(8.0), Inches(1.9), Inches(4.7), Inches(4.9), bg_color=COLOR_CARD_BG, border_color=RGBColor(219, 234, 254))
    tb_ts = s3.shapes.add_textbox(Inches(8.3), Inches(2.05), Inches(4.1), Inches(0.3))
    p = tb_ts.text_frame.paragraphs[0]
    p.text = "TECH STACK USED"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY_DARK

    tech_items = [
        ("Hardware / Components", "Desktop workstations, patrol tablets, smartphone edge"),
        ("Software / Languages", "TypeScript, Node.js 20, Python 3.11, SQL"),
        ("Frameworks & Libraries", "React 18, Vite 5, TailwindCSS, React Native 0.81, Expo SDK 54"),
        ("Mapping & Visuals", "Leaflet, React-Leaflet, React Force Graph 2D, Recharts"),
        ("AI / ML Analytics", "Pure TS ST-DBSCAN, Louvain, Holt's Trend, TF-IDF, TensorFlow.js COCO-SSD"),
        ("Tools & Platforms", "PostgreSQL, Drizzle ORM, Zoho Catalyst, esbuild, WebSockets"),
        ("Standards Followed", "NIST SHA-256, 4-tier RBAC, CCTNS 2.0 schema, Human-in-the-Loop")
    ]
    for i, (th, td) in enumerate(tech_items):
        ty = Inches(2.45 + i * 0.6)
        tb_item = s3.shapes.add_textbox(Inches(8.3), ty, Inches(4.1), Inches(0.55))
        tf = tb_item.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = th
        p.font.size = Pt(9.5)
        p.font.bold = True
        p.font.color.rgb = COLOR_ROYAL_BLUE
        p_d = tf.add_paragraph()
        p_d.text = td
        p_d.font.size = Pt(8.5)
        p_d.font.italic = True
        p_d.font.color.rgb = COLOR_TEXT_MUTED

    add_footer(s3, 3)

    # =========================================================================
    # SLIDE 4: PROTOTYPE DEMO VIDEO
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    add_page_header(s4, "Prototype Demo Video", "Play a 2–3 minute video of the prototype actually working.")

    # Left: Large Video Player Area (Dark Gray Rectangle, matching template)
    create_card(s4, Inches(0.6), Inches(1.9), Inches(7.8), Inches(4.9), bg_color=RGBColor(60, 64, 67), border_color=RGBColor(60, 64, 67))
    demo_poster = r"d:\projects\CrimeLens\CrimeLens_Full_Project\docs\demo-poster.jpg"
    if os.path.exists(demo_poster):
        s4.shapes.add_picture(demo_poster, Inches(0.65), Inches(1.95), width=Inches(7.7))
    else:
        tb_v = s4.shapes.add_textbox(Inches(1.5), Inches(3.8), Inches(6.0), Inches(1.0))
        p = tb_v.text_frame.paragraphs[0]
        p.text = "[ VIDEO SCREEN EMBED / PLAY AREA ]\nCrimeLens 2.5-Minute End-to-End Walkthrough"
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE
        p.alignment = PP_ALIGN.CENTER

    # Right: Video Checklist (Matching Yellow Card in Template)
    create_card(s4, Inches(8.7), Inches(1.9), Inches(4.0), Inches(4.9), bg_color=COLOR_YELLOW_CARD_BG, border_color=COLOR_YELLOW_BORDER)
    tb_vc = s4.shapes.add_textbox(Inches(9.0), Inches(2.1), Inches(3.4), Inches(0.4))
    p = tb_vc.text_frame.paragraphs[0]
    p.text = "VIDEO CHECKLIST"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY_DARK

    checklist = [
        "Length 2–3 minutes, clearly audible narration",
        "Show prototype switching ON and ingesting 201k KSP records",
        "Show a real input -> real output (ST-DBSCAN hotspot detection)",
        "Show Louvain criminal syndicate graph & kingpin identification",
        "Demonstrate one full success case: Panic SOS -> 3.8ms Haversine dispatch",
        "No edits that hide failures — keep it honest (TRL 4 academic prototype)",
        "Have a backup: offline video file + live local server"
    ]
    for i, item in enumerate(checklist):
        tb_chk = s4.shapes.add_textbox(Inches(9.0), Inches(2.6 + i * 0.58), Inches(3.4), Inches(0.5))
        tf = tb_chk.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = "✓  " + item
        p.font.size = Pt(8.5)
        p.font.color.rgb = COLOR_TEXT_DARK

    add_footer(s4, 4)

    # =========================================================================
    # SLIDE 5: PROTOTYPE SCREENSHOTS — SOFTWARE TEAM
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    add_page_header(s5, "Prototype Screenshots — Software Team", "Real screens from the working web prototype (demonstration data).")

    shots_s5 = [
        ("docs/screenshots/command-center.png", "STATE COMMAND CENTER", "Real-time anomaly stream, statewide crime KPIs, and human review queue"),
        ("docs/screenshots/governance-audit.png", "INTEGRITY & GOVERNANCE", "Recursive SHA-256 hash-chain verification with 100% tamper detection"),
        ("docs/screenshots/patrol-planner.png", "PREDICTIVE PATROL PLANNER", "Risk-weighted route optimization fusing ST-DBSCAN with Holt forecasts"),
        ("docs/screenshots/case-explorer.png", "CASE WORKSPACE & MO SEARCH", "Natural language intent query with TF-IDF similarity across districts")
    ]
    for idx, (rel_path, cap, sub) in enumerate(shots_s5):
        r = idx // 2
        c = idx % 2
        cx = Inches(0.6 + c * 6.2)
        cy = Inches(1.9 + r * 2.5)

        create_card(s5, cx, cy, Inches(5.9), Inches(2.35), bg_color=COLOR_CARD_BG, border_color=RGBColor(219, 234, 254))
        tb_cap = s5.shapes.add_textbox(cx + Inches(0.2), cy + Inches(0.08), Inches(5.5), Inches(0.3))
        p = tb_cap.text_frame.paragraphs[0]
        p.text = cap
        p.font.size = Pt(9.5)
        p.font.bold = True
        p.font.color.rgb = COLOR_NAVY_DARK

        full_p = os.path.join(r"d:\projects\CrimeLens\CrimeLens_Full_Project", rel_path)
        if os.path.exists(full_p):
            s5.shapes.add_picture(full_p, cx + Inches(0.2), cy + Inches(0.42), width=Inches(3.3), height=Inches(1.8))
        
        tb_desc = s5.shapes.add_textbox(cx + Inches(3.6), cy + Inches(0.45), Inches(2.1), Inches(1.7))
        tf = tb_desc.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = sub
        p.font.size = Pt(8)
        p.font.color.rgb = COLOR_TEXT_MUTED

    add_footer(s5, 5)

    # =========================================================================
    # SLIDE 6: PROTOTYPE PHOTOGRAPHS — SOFTWARE TEAMS
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    add_page_header(s6, "Prototype Photographs — Software Teams", "For web, app and AI/ML projects. Show real screens and real output — not mock-ups or stock images.")

    # 3 Column Cards: Web App, Mobile App, AI/ML Model
    # Col 1: Web Application
    create_card(s6, Inches(0.6), Inches(1.9), Inches(3.8), Inches(4.9), bg_color=COLOR_CARD_BG, border_color=RGBColor(219, 234, 254))
    tb_w = s6.shapes.add_textbox(Inches(0.9), Inches(2.1), Inches(3.2), Inches(0.35))
    p = tb_w.text_frame.paragraphs[0]
    p.text = "WEB APPLICATION"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY_DARK

    w_items = [
        "Local Demo: localhost:3000 (React 18 + Vite)",
        "Backend API: localhost:5000 (Node 20 Express)",
        "7 dedicated role workspaces (SP, Analyst, IO, etc.)",
        "ST-DBSCAN live cluster extraction (82 ms)",
        "Louvain network modularity graph (46 ms)",
        "Full cryptographic audit log verification"
    ]
    for i, itm in enumerate(w_items):
        tb_i = s6.shapes.add_textbox(Inches(0.9), Inches(2.6 + i * 0.48), Inches(3.2), Inches(0.4))
        p = tb_i.text_frame.paragraphs[0]
        p.text = "▪  " + itm
        p.font.size = Pt(8.5)
        p.font.color.rgb = COLOR_TEXT_DARK

    img_web = r"d:\projects\CrimeLens\CrimeLens_Full_Project\docs\screenshots\network-graph-force.png"
    if os.path.exists(img_web):
        s6.shapes.add_picture(img_web, Inches(0.85), Inches(5.1), width=Inches(3.3), height=Inches(1.5))

    # Col 2: Mobile Application
    create_card(s6, Inches(4.7), Inches(1.9), Inches(3.8), Inches(4.9), bg_color=COLOR_CARD_BG, border_color=RGBColor(219, 234, 254))
    tb_m = s6.shapes.add_textbox(Inches(5.0), Inches(2.1), Inches(3.2), Inches(0.35))
    p = tb_m.text_frame.paragraphs[0]
    p.text = "MOBILE APPLICATION"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY_DARK

    m_items = [
        "Built on Expo SDK 54 & React Native 0.81",
        "Frontline officer pocket intelligence",
        "Biometric & badge authentication",
        "Offline-first SQLite local store",
        "One-tap Panic SOS -> 3.8 ms Haversine dispatch",
        "Digital evidence locker with QR custody seal"
    ]
    for i, itm in enumerate(m_items):
        tb_i = s6.shapes.add_textbox(Inches(5.0), Inches(2.6 + i * 0.48), Inches(3.2), Inches(0.4))
        p = tb_i.text_frame.paragraphs[0]
        p.text = "▪  " + itm
        p.font.size = Pt(8.5)
        p.font.color.rgb = COLOR_TEXT_DARK

    img_mob = r"d:\projects\CrimeLens\CrimeLens_Full_Project\docs\screenshots\poster-av2.png"
    if os.path.exists(img_mob):
        s6.shapes.add_picture(img_mob, Inches(4.95), Inches(5.1), width=Inches(3.3), height=Inches(1.5))

    # Col 3: AI / ML Model
    create_card(s6, Inches(8.8), Inches(1.9), Inches(3.9), Inches(4.9), bg_color=COLOR_YELLOW_CARD_BG, border_color=COLOR_YELLOW_BORDER)
    tb_ai = s6.shapes.add_textbox(Inches(9.1), Inches(2.1), Inches(3.3), Inches(0.35))
    p = tb_ai.text_frame.paragraphs[0]
    p.text = "AI / ML MODEL & DATA"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY_DARK

    ai_items = [
        "Real Corpus: 201,733 records across 37 districts",
        "Synthetic Incident Set: 2,200 GPS-tagged records",
        "10 Karnataka Anchor Clusters (65% night bias)",
        "ST-DBSCAN: 600m spatial, 7-day temporal window",
        "Holt's Smoothing: 4-wk forecast (RMSE: 1.32–3.08)",
        "TensorFlow.js COCO-SSD client-side object detection",
        "TF-IDF cosine similarity for modus operandi matching",
        "SHA-256 recursive 256-bit hash-chain tamper ledger"
    ]
    for i, itm in enumerate(ai_items):
        tb_i = s6.shapes.add_textbox(Inches(9.1), Inches(2.55 + i * 0.52), Inches(3.3), Inches(0.48))
        tf = tb_i.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = "▪  " + itm
        p.font.size = Pt(8.5)
        p.font.color.rgb = COLOR_TEXT_DARK

    add_footer(s6, 6)

    # =========================================================================
    # SLIDE 7: MVP STATUS & TECHNOLOGY READINESS LEVEL
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    add_page_header(s7, "MVP Status & Technology Readiness Level", "What works today, what is still pending, and how far the technology has traveled.")

    # Col 1: Working Now (Green Card)
    create_card(s7, Inches(0.6), Inches(1.9), Inches(3.8), Inches(4.9), bg_color=COLOR_GREEN_CARD_BG, border_color=COLOR_GREEN_BORDER)
    tb_wn = s7.shapes.add_textbox(Inches(0.9), Inches(2.1), Inches(3.2), Inches(0.35))
    p = tb_wn.text_frame.paragraphs[0]
    p.text = "WORKING NOW"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = RGBColor(22, 101, 52)

    wn_items = [
        "201,733 KSP CSV record ingestion & mapping",
        "ST-DBSCAN hotspot detection with risk score",
        "Louvain community detection & gang kingpins",
        "Holt's double exponential forecasting (RMSE 1.32)",
        "TF-IDF semantic modus operandi matching",
        "Haversine nearest-unit dispatch (3.8 ms latency)",
        "SHA-256 cryptographic hash-chain verification",
        "Client-side TensorFlow.js COCO-SSD object vision",
        "Expo SDK 54 field officer mobile app prototype"
    ]
    for i, itm in enumerate(wn_items):
        tb_i = s7.shapes.add_textbox(Inches(0.9), Inches(2.55 + i * 0.46), Inches(3.2), Inches(0.42))
        tf = tb_i.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = "✓  " + itm
        p.font.size = Pt(8.2)
        p.font.color.rgb = COLOR_TEXT_DARK

    # Col 2: Still In Progress (Yellow Card)
    create_card(s7, Inches(4.7), Inches(1.9), Inches(3.8), Inches(4.9), bg_color=COLOR_YELLOW_CARD_BG, border_color=COLOR_YELLOW_BORDER)
    tb_ip = s7.shapes.add_textbox(Inches(5.0), Inches(2.1), Inches(3.2), Inches(0.35))
    p = tb_ip.text_frame.paragraphs[0]
    p.text = "STILL IN PROGRESS"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = RGBColor(133, 77, 14)

    ip_items = [
        ("Live CCTNS 2.0 API gateway sync", "Class D - awaiting official police staging credentials"),
        ("Statewide distributed Neo4j cluster", "Class D - architecturally planned for million-node graphs"),
        ("Real-world officer testing", "Internal team testing completed; police pilot pending"),
        ("Edge hardware video accelerator", "Integrating TensorRT / Coral NPU for in-vehicle cameras"),
        ("Formal institutional legal review", "Establishing judicial admissibility protocols with legal experts")
    ]
    for i, (head, note) in enumerate(ip_items):
        tb_i = s7.shapes.add_textbox(Inches(5.0), Inches(2.55 + i * 0.8), Inches(3.2), Inches(0.75))
        tf = tb_i.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = "⏳  " + head
        p.font.size = Pt(8.5)
        p.font.bold = True
        p.font.color.rgb = COLOR_TEXT_DARK
        p_n = tf.add_paragraph()
        p_n.text = note
        p_n.font.size = Pt(7.5)
        p_n.font.color.rgb = COLOR_TEXT_MUTED
        p_n.space_before = Pt(2)

    # Col 3: TRL Breakdown & Our Level
    create_card(s7, Inches(8.8), Inches(1.9), Inches(3.9), Inches(4.9), bg_color=COLOR_CARD_BG, border_color=RGBColor(219, 234, 254))
    tb_trl_h = s7.shapes.add_textbox(Inches(9.1), Inches(2.1), Inches(3.3), Inches(0.35))
    p = tb_trl_h.text_frame.paragraphs[0]
    p.text = "TECHNOLOGY READINESS LEVEL"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY_DARK

    trl_bands = [
        ("TRL 1 – 3", "Idea & lab proof of concept"),
        ("TRL 4 – 6", "Working prototype tested in a close-to-real setting"),
        ("TRL 7 – 9", "Demonstrated / in use in the real operational environment")
    ]
    for i, (th, td) in enumerate(trl_bands):
        create_card(s7, Inches(9.1), Inches(2.5 + i * 0.75), Inches(3.3), Inches(0.65), bg_color=COLOR_WHITE, border_color=COLOR_BORDER_LIGHT)
        tb_tb = s7.shapes.add_textbox(Inches(9.2), Inches(2.55 + i * 0.75), Inches(3.1), Inches(0.55))
        tf = tb_tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = th
        p.font.size = Pt(9)
        p.font.bold = True
        p.font.color.rgb = COLOR_ROYAL_BLUE
        p_d = tf.add_paragraph()
        p_d.text = td
        p_d.font.size = Pt(7.5)
        p_d.font.color.rgb = COLOR_TEXT_MUTED

    # Bottom Yellow Highlight Box: Our TRL 4
    create_card(s7, Inches(9.1), Inches(5.0), Inches(3.3), Inches(1.6), bg_color=COLOR_YELLOW_CARD_BG, border_color=COLOR_YELLOW_BORDER)
    tb_our = s7.shapes.add_textbox(Inches(9.2), Inches(5.1), Inches(3.1), Inches(1.4))
    tf = tb_our.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = "Our TRL: TRL 4"
    p.font.size = Pt(11.5)
    p.font.bold = True
    p.font.color.rgb = RGBColor(133, 77, 14)
    p_d = tf.add_paragraph()
    p_d.text = "Working prototype validated in a simulated environment using real Karnataka State Police historical datasets and synthetic incident sets. It has not yet been deployed in active police field operations."
    p_d.font.size = Pt(8)
    p_d.font.color.rgb = COLOR_TEXT_DARK
    p_d.space_before = Pt(4)

    add_footer(s7, 7)

    # =========================================================================
    # SLIDE 8: USER TESTING & VALIDATION
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    add_page_header(s8, "User Testing & Validation", "Criterion 2 — take the prototype back to the users from Review I. Write who, what task, and what they actually said.")

    # 3 Column Layout exactly matching sample
    # Card 1: WHO WE TESTED WITH
    create_card(s8, Inches(0.6), Inches(1.9), Inches(3.8), Inches(4.9), bg_color=COLOR_CARD_BG, border_color=RGBColor(219, 234, 254))
    tb_w1 = s8.shapes.add_textbox(Inches(0.9), Inches(2.1), Inches(3.2), Inches(0.35))
    p = tb_w1.text_frame.paragraphs[0]
    p.text = "WHO WE TESTED WITH"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY_DARK

    tb_w1_body = s8.shapes.add_textbox(Inches(0.9), Inches(2.6), Inches(3.2), Inches(4.0))
    tf = tb_w1_body.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = "Not yet tested with end-user police officers.\n\nSo far, the internal engineering team evaluated the prototype in a controlled lab testbed across:"
    p.font.size = Pt(9.5)
    p.font.color.rgb = COLOR_TEXT_DARK

    tests_list = [
        "201,733 statewide KSP corpus records",
        "2,200 synthetic GPS-tagged incident rows",
        "Simulated 500 patrol units responding to concurrent Panic SOS alerts",
        "Holdout validation across 156 weekly incident bins"
    ]
    for itm in tests_list:
        p_i = tf.add_paragraph()
        p_i.text = "• " + itm
        p_i.font.size = Pt(8.5)
        p_i.font.color.rgb = COLOR_TEXT_MUTED
        p_i.space_before = Pt(4)

    # Card 2: WHAT WE ASKED THEM TO DO
    create_card(s8, Inches(4.7), Inches(1.9), Inches(3.8), Inches(4.9), bg_color=COLOR_CARD_BG, border_color=RGBColor(219, 234, 254))
    tb_w2 = s8.shapes.add_textbox(Inches(5.0), Inches(2.1), Inches(3.2), Inches(0.35))
    p = tb_w2.text_frame.paragraphs[0]
    p.text = "WHAT WE ASKED THEM TO DO"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY_DARK

    tb_w2_body = s8.shapes.add_textbox(Inches(5.0), Inches(2.6), Inches(3.2), Inches(4.0))
    tf = tb_w2_body.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = "Planned Pilot Evaluation Protocol:"
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_ROYAL_BLUE

    protocol = [
        ("Task 1: Hotspot Review", "Station House Officers verify ST-DBSCAN cluster boundaries against known beat crime sprees."),
        ("Task 2: Syndicate Tracing", "Crime Analysts inspect Louvain modularity clusters and verify inferred co-accused links."),
        ("Task 3: Patrol Deployment", "Supervisors approve automated predictive patrol shifts and evaluate vehicle dispatch timing."),
        ("Task 4: SOS Field Response", "Field officers test one-tap Panic SOS on Expo mobile app and measure response latency.")
    ]
    for h, d in protocol:
        p_h = tf.add_paragraph()
        p_h.text = h
        p_h.font.size = Pt(8.5)
        p_h.font.bold = True
        p_h.font.color.rgb = COLOR_TEXT_DARK
        p_h.space_before = Pt(6)
        p_d = tf.add_paragraph()
        p_d.text = d
        p_d.font.size = Pt(7.5)
        p_d.font.color.rgb = COLOR_TEXT_MUTED

    # Card 3: WHAT THEY SAID
    create_card(s8, Inches(8.8), Inches(1.9), Inches(3.9), Inches(4.9), bg_color=COLOR_YELLOW_CARD_BG, border_color=COLOR_YELLOW_BORDER)
    tb_w3 = s8.shapes.add_textbox(Inches(9.1), Inches(2.1), Inches(3.3), Inches(0.35))
    p = tb_w3.text_frame.paragraphs[0]
    p.text = "WHAT THEY SAID"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = RGBColor(133, 77, 14)

    tb_w3_body = s8.shapes.add_textbox(Inches(9.1), Inches(2.6), Inches(3.3), Inches(4.0))
    tf = tb_w3_body.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = "“No official police user feedback collected yet. Honest reporting: quotes, field usability struggles, and feedback-driven iterations will be recorded during the planned station pilot stage.”"
    p.font.size = Pt(9.5)
    p.font.italic = True
    p.font.color.rgb = RGBColor(133, 77, 14)

    p_f = tf.add_paragraph()
    p_f.text = "\nInternal Developer Verification Findings:\n• ST-DBSCAN dynamic slider allows analysts to isolate transient from chronic crime sprees in under 100 ms.\n• SHA-256 tamper verification provides 100% detection of manual payload mutations.\n• Mobile app stays responsive offline with SQLite."
    p_f.font.size = Pt(8.2)
    p_f.font.color.rgb = COLOR_TEXT_DARK

    add_footer(s8, 8)

    # =========================================================================
    # SLIDE 9: COMPETITOR ANALYSIS
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    add_page_header(s9, "Competitor Analysis", "Compare yourself with what users use today. Every tick and cross needs 2–4 words of proof beside it.")

    # Table matching template layout exactly
    tbl_shape = s9.shapes.add_table(7, 4, Inches(0.6), Inches(1.9), Inches(12.133), Inches(3.9))
    tbl = tbl_shape.table
    tbl.columns[0].width = Inches(2.7)
    tbl.columns[1].width = Inches(3.1)
    tbl.columns[2].width = Inches(3.1)
    tbl.columns[3].width = Inches(3.233)

    c_headers = ["FEATURE / CRITERIA", "OUR SOLUTION (CrimeLens)", "ALTERNATIVE 1 — Manual CCTNS / Spreadsheets", "ALTERNATIVE 2 — Commercial GIS Tools"]
    for j, h in enumerate(c_headers):
        cell = tbl.cell(0, j)
        cell.fill.solid()
        if j == 0:
            cell.fill.fore_color.rgb = COLOR_NAVY_DARK
        elif j == 1:
            cell.fill.fore_color.rgb = RGBColor(22, 101, 52) # Dark green
        else:
            cell.fill.fore_color.rgb = COLOR_ROYAL_BLUE
        p = cell.text_frame.paragraphs[0]
        p.text = h
        p.font.size = Pt(9)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE

    c_data = [
        ("Cost to the user", "✔ ₹0 software license (open-source core)", "✖ High staff hours spent in manual entry", "✖ Expensive proprietary licenses per seat"),
        ("Time taken / speed", "✔ Automated real-time streaming (<100ms)", "✖ Slow, batch-based delayed queries", "✖ Manual shapefile import required"),
        ("Accuracy / quality of result", "✔ 6.2% MAPE (Holt) + dual ST-DBSCAN", "✖ Inconsistent, uncalibrated estimates", "✖ Static density maps without time window"),
        ("Syndicate link detection", "✔ Louvain modularity + kingpin analysis", "✖ Not supported; manual name matching", "✖ Limited basic geographic overlays"),
        ("Works with low power / offline", "✔ Offline-capable Expo mobile app", "✖ Desktop terminal only; requires LAN", "✖ Specialist workstation only"),
        ("Audit & judicial governance", "✔ SHA-256 tamper-evident hash ledger", "✖ Mutable SQL table logs", "✖ Proprietary closed-source event logs")
    ]
    for i, row in enumerate(c_data):
        for j, val in enumerate(row):
            cell = tbl.cell(i+1, j)
            cell.fill.solid()
            if j == 1:
                cell.fill.fore_color.rgb = COLOR_GREEN_CARD_BG
            elif i % 2 == 0:
                cell.fill.fore_color.rgb = COLOR_WHITE
            else:
                cell.fill.fore_color.rgb = RGBColor(248, 250, 252)
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.size = Pt(8.5)
            p.font.color.rgb = COLOR_TEXT_DARK

    # Bottom Unfair Advantage Callout Box
    create_card(s9, Inches(0.6), Inches(6.0), Inches(12.133), Inches(0.9), bg_color=COLOR_YELLOW_CARD_BG, border_color=COLOR_YELLOW_BORDER)
    tb_ua = s9.shapes.add_textbox(Inches(0.8), Inches(6.05), Inches(11.7), Inches(0.8))
    tf = tb_ua.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = "🔑  UNFAIR ADVANTAGE • WHY IT IS HARD TO COPY"
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = RGBColor(133, 77, 14)
    p_d = tf.add_paragraph()
    p_d.text = "Unified statewide corpus ingestion paired with pure TypeScript mathematical engines (ST-DBSCAN, Louvain, Holt) and immutable SHA-256 governance. It connects central state command with frontline officers in one tamper-evident loop."
    p_d.font.size = Pt(8.5)
    p_d.font.color.rgb = COLOR_TEXT_DARK
    p_d.space_before = Pt(2)

    add_footer(s9, 9)

    # =========================================================================
    # SLIDE 10: BUSINESS MODEL CANVAS
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    add_page_header(s10, "Business Model Canvas", "The full picture on one page. One or two honest lines per block — no invented figures.")

    # BMC Grid matching template
    bmc_blocks = [
        ("KEY PARTNERS", "• State Police Computer Wings\n• NCRB & MHA Modernization Cells\n• College Mentor & Institutional Labs\n• Open-source TypeScript/React communities", Inches(0.6), Inches(1.9), Inches(2.2), Inches(2.5)),
        ("KEY ACTIVITIES", "• Algorithmic model calibration\n• CCTNS schema synchronization\n• Mobile field app testing & offline sync\n• Cryptographic audit integrity runs", Inches(2.9), Inches(1.9), Inches(2.1), Inches(1.2)),
        ("KEY RESOURCES", "• 201,733 normalized KSP corpus\n• ST-DBSCAN & Louvain algorithms\n• TensorFlow.js MobileNet weights\n• Dedicated student engineering team", Inches(2.9), Inches(3.2), Inches(2.1), Inches(1.2)),
        ("VALUE PROPOSITION", "• Proactive patrol optimization cuts crime blindspots\n• Louvain discovers cross-district criminal syndicates\n• Tamper-evident SHA-256 evidence chain of custody\n• Frontline officer safety via 3.8ms Panic SOS dispatch", Inches(5.1), Inches(1.9), Inches(2.5), Inches(2.5), True), # Yellow
        ("CUSTOMER RELATIONS", "• Dedicated state technical SLA\n• Tiered 4-role user training\n• Transparent plain-language explainability", Inches(7.7), Inches(1.9), Inches(2.1), Inches(1.2)),
        ("CHANNELS", "• Secure State Police Intranet / VPN\n• Dedicated Mobile MDM store\n• Command center web console", Inches(7.7), Inches(3.2), Inches(2.1), Inches(1.2)),
        ("CUSTOMER SEGMENTS", "• State Police Departments (DG/IGP)\n• City Police Commissionerates (SP/DCP)\n• Specialized Crime Investigation Depts (CID)\n• Women & Child Safety Patrols (Akka Pade)", Inches(9.9), Inches(1.9), Inches(2.833), Inches(2.5)),
        ("COST STRUCTURE", "Prototype: ₹0 software licenses (open-source stack). Team time; cloud hosting cost depends on deployment scale (Catalyst / NIC).", Inches(0.6), Inches(4.5), Inches(5.8), Inches(1.2)),
        ("REVENUE / SUSTAINABILITY", "Funded by academic institution / government grant; B2G police modernization deployment support is planned.", Inches(6.6), Inches(4.5), Inches(6.133), Inches(1.2))
    ]

    for item in bmc_blocks:
        title, text, left, top, width, height = item[:6]
        is_yellow = len(item) > 6 and item[6]
        bg = COLOR_YELLOW_CARD_BG if is_yellow else COLOR_CARD_BG
        brd = COLOR_YELLOW_BORDER if is_yellow else RGBColor(219, 234, 254)
        create_card(s10, left, top, width, height, bg_color=bg, border_color=brd)

        tb = s10.shapes.add_textbox(left + Inches(0.1), top + Inches(0.08), width - Inches(0.2), height - Inches(0.16))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(8.5)
        p.font.bold = True
        p.font.color.rgb = RGBColor(133, 77, 14) if is_yellow else COLOR_NAVY_DARK
        p_t = tf.add_paragraph()
        p_t.text = text
        p_t.font.size = Pt(7.5)
        p_t.font.color.rgb = COLOR_TEXT_DARK
        p_t.space_before = Pt(3)

    # Bottom Checkbox Row: WHICH BUSINESS MODEL ARE WE?
    create_card(s10, Inches(0.6), Inches(5.85), Inches(12.133), Inches(1.05), bg_color=COLOR_CARD_BG, border_color=RGBColor(219, 234, 254))
    tb_bm = s10.shapes.add_textbox(Inches(0.8), Inches(5.92), Inches(11.7), Inches(0.95))
    tf = tb_bm.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = "WHICH BUSINESS MODEL ARE WE? (tick one)"
    p.font.size = Pt(9)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY_DARK

    boxes_text = "☐ B2C — a person buys it        ☐ B2B — a company buys it        ☑ B2G — a government body buys it        ☐ CSR / NGO — sponsor pays"
    p2 = tf.add_paragraph()
    p2.text = boxes_text
    p2.font.size = Pt(9.5)
    p2.font.bold = True
    p2.font.color.rgb = COLOR_ROYAL_BLUE
    p2.space_before = Pt(4)

    p3 = tf.add_paragraph()
    p3.text = "Value Proposition and Customer Segments match the law-enforcement problem: the state police department is the buyer, while frontline police officers are the end-users."
    p3.font.size = Pt(7.5)
    p3.font.italic = True
    p3.font.color.rgb = COLOR_TEXT_MUTED
    p3.space_before = Pt(2)

    add_footer(s10, 10)

    # =========================================================================
    # SLIDE 11: MARKET SIZE — TAM, SAM, SOM
    # =========================================================================
    s11 = prs.slides.add_slide(blank_layout)
    add_page_header(s11, "Market Size — TAM, SAM, SOM", "How many people have this problem? Use real numbers with a source. A small sourced number beats a big guess.")

    # Left Column Labels
    labels = [
        ("TAM", Inches(2.2)),
        ("SAM", Inches(3.8)),
        ("SOM", Inches(5.4))
    ]
    for lbl, y in labels:
        tb_l = s11.shapes.add_textbox(Inches(1.2), y, Inches(1.8), Inches(0.8))
        p = tb_l.text_frame.paragraphs[0]
        p.text = lbl
        p.font.size = Pt(28)
        p.font.bold = True
        p.font.color.rgb = COLOR_NAVY_DARK

    # Right 3 Cards exactly matching sample
    # TAM Card
    create_card(s11, Inches(3.6), Inches(1.9), Inches(9.133), Inches(1.4), bg_color=COLOR_CARD_BG, border_color=RGBColor(219, 234, 254))
    tb = s11.shapes.add_textbox(Inches(3.8), Inches(2.0), Inches(8.7), Inches(1.2))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = "TAM — Total Addressable Market"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY_DARK
    p2 = tf.add_paragraph()
    p2.text = "Everyone in the country who has this problem. The whole pie."
    p2.font.size = Pt(8.5)
    p2.font.color.rgb = COLOR_TEXT_MUTED
    p3 = tf.add_paragraph()
    p3.text = "All police stations in India = 17,535 (BPR&D, as of 01.01.2023)"
    p3.font.size = Pt(9.5)
    p3.font.bold = True
    p3.font.color.rgb = COLOR_ROYAL_BLUE
    p3.space_before = Pt(4)

    # SAM Card
    create_card(s11, Inches(3.6), Inches(3.5), Inches(9.133), Inches(1.4), bg_color=COLOR_CARD_BG, border_color=RGBColor(219, 234, 254))
    tb = s11.shapes.add_textbox(Inches(3.8), Inches(3.6), Inches(8.7), Inches(1.2))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = "SAM — Serviceable Available Market"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY_DARK
    p2 = tf.add_paragraph()
    p2.text = "The slice you could realistically reach with your solution and channel."
    p2.font.size = Pt(8.5)
    p2.font.color.rgb = COLOR_TEXT_MUTED
    p3 = tf.add_paragraph()
    p3.text = "Police stations in Karnataka = 1,100+ (Karnataka State Police Policy Note, 2024)"
    p3.font.size = Pt(9.5)
    p3.font.bold = True
    p3.font.color.rgb = COLOR_ROYAL_BLUE
    p3.space_before = Pt(4)

    # SOM Card (Yellow highlight card)
    create_card(s11, Inches(3.6), Inches(5.1), Inches(9.133), Inches(1.6), bg_color=COLOR_YELLOW_CARD_BG, border_color=COLOR_YELLOW_BORDER)
    tb = s11.shapes.add_textbox(Inches(3.8), Inches(5.2), Inches(8.7), Inches(1.4))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = "SOM — Serviceable Obtainable Market"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = RGBColor(133, 77, 14)
    p2 = tf.add_paragraph()
    p2.text = "What you could actually capture in the first 1–2 years. Be honest and small."
    p2.font.size = Pt(8.5)
    p2.font.color.rgb = COLOR_TEXT_MUTED
    p3 = tf.add_paragraph()
    p3.text = "Pilot with 5 police stations in Bengaluru Urban commissionerate in Year 1 (our academic target, not sourced)"
    p3.font.size = Pt(9.5)
    p3.font.bold = True
    p3.font.color.rgb = COLOR_ROYAL_BLUE
    p3.space_before = Pt(4)

    add_footer(s11, 11)

    # =========================================================================
    # SLIDE 12: RISKS & MITIGATION
    # =========================================================================
    s12 = prs.slides.add_slide(blank_layout)
    add_page_header(s12, "Risks & Mitigation", "What could stop this from working, and what you are doing about it. Reviewers trust teams that name their risks.")

    # 4 Column Risk Table matching template
    tbl_shape = s12.shapes.add_table(5, 4, Inches(0.6), Inches(1.9), Inches(12.133), Inches(4.9))
    tbl = tbl_shape.table
    tbl.columns[0].width = Inches(2.2)
    tbl.columns[1].width = Inches(3.2)
    tbl.columns[2].width = Inches(3.3)
    tbl.columns[3].width = Inches(3.433)

    r_headers = ["RISK", "WHY IT MATTERS", "OUR MITIGATION", "HOW WE WILL VALIDATE"]
    for j, h in enumerate(r_headers):
        cell = tbl.cell(0, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_NAVY_DARK
        p = cell.text_frame.paragraphs[0]
        p.text = h
        p.font.size = Pt(9)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE

    r_data = [
        ("Technical\nalgorithmic bias / false clusters",
         "Over-policing historically high-arrest wards creates self-fulfilling prediction loops.",
         "Station Anomaly Z-Score baseline tracking and population-normalized fairness indicators flag skewed patrolling.",
         "Holdout evaluation across 156 weekly bins; compare cluster rates against census population baselines."),
        
        ("Cost / supply\nserverless compute scale-up",
         "Processing statewide CCTV video feeds centrally could cause runaway hosting costs.",
         "Client-side TensorFlow.js COCO-SSD runs on edge browser / device; only alerts sent to central server.",
         "Latency and resource profiling on low-cost devices; verify CPU stays below 45%."),
        
        ("User adoption\nfrontline officer resistance",
         "If the mobile UI is clunky, beat officers will bypass digital checklists during fast-paced duty.",
         "One-tap Panic SOS, clean low-friction beat checklists, and native bilingual interface (English + Kannada).",
         "Usability task timing during student review; timed simulated SOS dispatch under 4 seconds."),
        
        ("Data / safety\nsensitive citizen FIR disclosure",
         "Leaking juvenile records or victim identities violates constitutional and data protection laws.",
         "Complete synthetic incident generation for demonstrations; on-premise encrypted storage for production; SHA-256 audit.",
         "Automated vulnerability scan; verify zero PII exists in client bundles; cryptographic chain integrity check.")
    ]

    for i, row in enumerate(r_data):
        for j, val in enumerate(row):
            cell = tbl.cell(i+1, j)
            cell.fill.solid()
            if j == 0:
                cell.fill.fore_color.rgb = COLOR_RED_CARD_BG
            elif i % 2 == 0:
                cell.fill.fore_color.rgb = COLOR_WHITE
            else:
                cell.fill.fore_color.rgb = RGBColor(248, 250, 252)
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.size = Pt(8.2)
            p.font.color.rgb = COLOR_TEXT_RED if j == 0 else COLOR_TEXT_DARK

    add_footer(s12, 12)

    # =========================================================================
    # SLIDE 13: DISSEMINATION, CERTIFICATIONS & IP
    # =========================================================================
    s13 = prs.slides.add_slide(blank_layout)
    add_page_header(s13, "Dissemination, Certifications & IP", "Criteria 3 & 4 — certifications completed, contests entered, funding, publication or patent. Mark the rest as planned.")

    # 6 Card Grid matching template
    ip_grid = [
        ("HACKATHONS / CONTESTS", "KSP Datathon 2026 track submission — entered / benchmarked\nState Police Modernization Challenge — planned", Inches(0.6), Inches(1.9), Inches(5.8), Inches(1.5)),
        ("FUNDING APPLIED", "Tamil Nadu Student Project Scheme (TNSCST) — planned\nInstitutional Innovation Seed Grant — applied / under review", Inches(6.6), Inches(1.9), Inches(6.133), Inches(1.5)),
        ("PAPER / PUBLICATION", "IEEE / Springer conference paper drafted: 'CrimeLens: An AI-First Crime Intelligence Operating System with Spatio-Temporal Analytics' — drafted, not yet submitted", Inches(0.6), Inches(3.6), Inches(5.8), Inches(1.5)),
        ("PATENT / IP", "Software Copyright filing with Indian Copyright Office for CrimeLens Core Processing Engine & Hash Ledger — planned upon review completion", Inches(6.6), Inches(3.6), Inches(6.133), Inches(1.5)),
        ("CERTIFICATIONS COMPLETED", "Team certifications in Full-Stack Development, React Native, Machine Learning with TensorFlow & Python — completed by team members", Inches(0.6), Inches(5.3), Inches(5.8), Inches(1.5), True), # Green
        ("YUKTHI SUBMISSION", "Yukthi Innovation Challenge Submission ID: To be added by the team upon portal upload", Inches(6.6), Inches(5.3), Inches(6.133), Inches(1.5), False, True) # Yellow
    ]

    for item in ip_grid:
        title, text, left, top, width, height = item[:6]
        is_green = len(item) > 6 and item[6]
        is_yellow = len(item) > 7 and item[7]
        bg = COLOR_GREEN_CARD_BG if is_green else COLOR_YELLOW_CARD_BG if is_yellow else COLOR_CARD_BG
        brd = COLOR_GREEN_BORDER if is_green else COLOR_YELLOW_BORDER if is_yellow else RGBColor(219, 234, 254)
        create_card(s13, left, top, width, height, bg_color=bg, border_color=brd)

        tb = s13.shapes.add_textbox(left + Inches(0.2), top + Inches(0.12), width - Inches(0.4), height - Inches(0.24))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(9.5)
        p.font.bold = True
        p.font.color.rgb = RGBColor(22, 101, 52) if is_green else RGBColor(133, 77, 14) if is_yellow else COLOR_NAVY_DARK
        p_t = tf.add_paragraph()
        p_t.text = text
        p_t.font.size = Pt(8.5)
        p_t.font.color.rgb = COLOR_TEXT_DARK
        p_t.space_before = Pt(4)

    add_footer(s13, 13)

    # =========================================================================
    # SLIDE 14: IMPACT & SDG MAPPING
    # =========================================================================
    s14 = prs.slides.add_slide(blank_layout)
    add_page_header(s14, "Impact & SDG Mapping", "Who benefits, by how much, and which global goal it supports.")

    # Top Row: 2 Cards (Who Benefits & Measured Impact)
    create_card(s14, Inches(0.6), Inches(1.9), Inches(4.5), Inches(3.3), bg_color=COLOR_CARD_BG, border_color=RGBColor(219, 234, 254))
    tb_wb = s14.shapes.add_textbox(Inches(0.9), Inches(2.1), Inches(3.9), Inches(2.9))
    tf = tb_wb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = "WHO BENEFITS"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY_DARK

    p_d1 = tf.add_paragraph()
    p_d1.text = "Direct Beneficiaries:\nState police analysts, district SPs, station IOs, and Akka Pade frontline patrol officers."
    p_d1.font.size = Pt(8.5)
    p_d1.font.color.rgb = COLOR_TEXT_DARK
    p_d1.space_before = Pt(8)

    p_d2 = tf.add_paragraph()
    p_d2.text = "Indirect Beneficiaries:\nCitizens, women transit commuters, and local businesses benefiting from reduced crime blindspots and accountable, auditable policing."
    p_d2.font.size = Pt(8.5)
    p_d2.font.color.rgb = COLOR_TEXT_MUTED
    p_d2.space_before = Pt(8)

    create_card(s14, Inches(5.4), Inches(1.9), Inches(7.333), Inches(3.3), bg_color=COLOR_CARD_BG, border_color=RGBColor(219, 234, 254))
    tb_mi = s14.shapes.add_textbox(Inches(5.7), Inches(2.1), Inches(6.7), Inches(2.9))
    tf = tb_mi.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = "MEASURED IMPACT (LABORATORY BENCHMARKS)"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY_DARK

    m_items = [
        "Ingested and indexed 201,733 statewide KSP records in 1,380 ms (146k rec/s throughput).",
        "ST-DBSCAN detected all 10 seeded hotspot zones in 82 ms without false geographical fragmentation.",
        "Louvain modularity successfully isolated all 3 synthetic cross-district gang syndicates (modularity Q > 0.60).",
        "Holt's time-series model achieved 6.2% MAPE across 156 historical weekly incident bins.",
        "Nearest patrol unit dispatch computed in 3.8 ms using Haversine distance matrix optimizer.",
        "SHA-256 recursive ledger verified 10,000 chained blocks in 18 ms with 100% detection of manual payload mutations.",
        "Synthetic data only — real-world crime reduction percentage not yet measured."
    ]
    for itm in m_items:
        p_i = tf.add_paragraph()
        p_i.text = "• " + itm
        p_i.font.size = Pt(8.2)
        p_i.font.color.rgb = COLOR_TEXT_DARK
        p_i.space_before = Pt(3)

    # Bottom Row: UN SDG Box matching template
    create_card(s14, Inches(0.6), Inches(5.4), Inches(12.133), Inches(1.5), bg_color=COLOR_CARD_BG, border_color=RGBColor(219, 234, 254))
    tb_sdg_h = s14.shapes.add_textbox(Inches(0.9), Inches(5.5), Inches(11.5), Inches(0.3))
    p = tb_sdg_h.text_frame.paragraphs[0]
    p.text = "UN SUSTAINABLE DEVELOPMENT GOALS"
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY_DARK

    sdg_boxes = [
        ("1 No Poverty", False), ("2 Zero Hunger", False), ("3 Good Health", False),
        ("4 Quality Education", False), ("5 Gender Equality", False), ("6 Clean Water", False),
        ("7 Clean Energy", False), ("8 Decent Work", False), ("9 Industry & Innovation", True),
        ("10 Reduced Inequalities", False), ("11 Sustainable Cities", True), ("12 Responsible Consumption", False),
        ("13 Climate Action", False), ("14 Life Below Water", False), ("15 Life on Land", False),
        ("16 Peace & Justice", True), ("17 Partnerships", False)
    ]
    for idx, (sdg_name, is_target) in enumerate(sdg_boxes):
        r = idx // 9
        c = idx % 9
        bx = Inches(0.9 + c * 1.25)
        by = Inches(5.82 + r * 0.38)
        bg = COLOR_GREEN_CARD_BG if is_target else COLOR_WHITE
        brd = COLOR_GREEN_BORDER if is_target else COLOR_BORDER_LIGHT
        c_sdg = create_card(s14, bx, by, Inches(1.18), Inches(0.32), bg_color=bg, border_color=brd)
        tb_s = s14.shapes.add_textbox(bx, by + Inches(0.02), Inches(1.18), Inches(0.28))
        tf = tb_s.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = sdg_name
        p.font.size = Pt(6.8)
        p.font.bold = is_target
        p.font.color.rgb = RGBColor(22, 101, 52) if is_target else COLOR_TEXT_MUTED
        p.alignment = PP_ALIGN.CENTER

    tb_sdg_sum = s14.shapes.add_textbox(Inches(0.9), Inches(6.58), Inches(11.5), Inches(0.25))
    p = tb_sdg_sum.text_frame.paragraphs[0]
    p.text = "Our solution maps to SDG 16 (Peace, Justice & Strong Institutions) as primary, because tamper-evident analytics support transparent justice, and SDG 11 (Sustainable Cities) for safer public spaces."
    p.font.size = Pt(8)
    p.font.bold = True
    p.font.color.rgb = COLOR_ROYAL_BLUE

    add_footer(s14, 14)

    # =========================================================================
    # SLIDE 15: TEAM CONTRIBUTION
    # =========================================================================
    s15 = prs.slides.add_slide(blank_layout)
    add_page_header(s15, "Team Contribution", "Criterion 5 — every member should be able to speak about their own part of the build.")

    # 4 Row Table matching sample template
    tbl_shape = s15.shapes.add_table(5, 4, Inches(0.6), Inches(1.9), Inches(12.133), Inches(4.9))
    tbl = tbl_shape.table
    tbl.columns[0].width = Inches(2.6)
    tbl.columns[1].width = Inches(1.8)
    tbl.columns[2].width = Inches(2.8)
    tbl.columns[3].width = Inches(4.933)

    t_headers = ["Member Name", "Roll No.", "Role in the Build", "What they delivered for the prototype"]
    for j, h in enumerate(t_headers):
        cell = tbl.cell(0, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_NAVY_DARK
        p = cell.text_frame.paragraphs[0]
        p.text = h
        p.font.size = Pt(9.5)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE

    t_rows = [
        ("Parameshwaran P", "722824104166", "Backend API & Graph Lead",
         "• Ingested and normalized 201,733 statewide KSP CSV records\n• Implemented Louvain modularity algorithm & bipartite entity projection\n• Built Node.js 20 Express REST APIs & WebSocket live alert bus"),
        
        ("Ram Vishnu P", "722824104186", "Spatial Analytics & Forecast Lead",
         "• Developed ST-DBSCAN spatio-temporal clustering & intensity scoring\n• Implemented Holt's double-exponential smoothing equations (6.2% MAPE)\n• Engineered station anomaly Z-score engine & time-slot risk matrix"),
        
        ("Ravishankar AG", "722824104191", "Command Console & UI/UX Lead",
         "• Engineered React 18 + Vite web command console with TailwindCSS\n• Implemented Leaflet geospatial mapping & React Force Graph 2D\n• Designed explainability drawers, supervisor review queue & Recharts analytics"),
        
        ("Rishivardhan J", "722824104196", "Tactical Mobile & Security Lead",
         "• Built cross-platform React Native / Expo SDK 54 mobile application\n• Formulated SHA-256 recursive hash-chain tamper-evident audit ledger\n• Engineered Haversine nearest-unit dispatch optimizer (3.8ms) & SQLite offline sync")
    ]
    for i, row in enumerate(t_rows):
        for j, val in enumerate(row):
            cell = tbl.cell(i+1, j)
            cell.fill.solid()
            if i % 2 == 0:
                cell.fill.fore_color.rgb = COLOR_WHITE
            else:
                cell.fill.fore_color.rgb = RGBColor(248, 250, 252)
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.size = Pt(8.5)
            p.font.color.rgb = COLOR_TEXT_DARK

    add_footer(s15, 15)

    # =========================================================================
    # SLIDE 16: THANK YOU (Dark Blue Background matching sample)
    # =========================================================================
    s16 = prs.slides.add_slide(blank_layout)

    # Background: Solid Dark Navy #0A1944
    bg = s16.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = COLOR_NAVY_DARK
    bg.line.fill.background()

    # Subtle decorative circular arcs matching template aesthetic
    c1 = s16.shapes.add_shape(MSO_SHAPE.OVAL, Inches(10.5), Inches(-1.5), Inches(5.0), Inches(5.0))
    c1.fill.solid()
    c1.fill.fore_color.rgb = RGBColor(16, 36, 94)
    c1.line.fill.background()

    c2 = s16.shapes.add_shape(MSO_SHAPE.OVAL, Inches(8.5), Inches(4.5), Inches(6.0), Inches(6.0))
    c2.fill.solid()
    c2.fill.fore_color.rgb = RGBColor(14, 32, 84)
    c2.line.fill.background()

    # Center Text Frame
    tb_ty = s16.shapes.add_textbox(Inches(2.0), Inches(2.6), Inches(9.333), Inches(2.5))
    tf = tb_ty.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

    p_ty = tf.paragraphs[0]
    p_ty.text = "Thank You!"
    p_ty.font.size = Pt(44)
    p_ty.font.bold = True
    p_ty.font.color.rgb = COLOR_YELLOW_ACCENT # Exact matching yellow
    p_ty.font.name = "Calibri"
    p_ty.alignment = PP_ALIGN.CENTER

    p_sub = tf.add_paragraph()
    p_sub.text = "Questions, feedback & ideas are always welcome."
    p_sub.font.size = Pt(17)
    p_sub.font.color.rgb = COLOR_WHITE
    p_sub.font.name = "Calibri"
    p_sub.alignment = PP_ALIGN.CENTER
    p_sub.space_before = Pt(14)

    p_team = tf.add_paragraph()
    p_team.text = "Batch 17  •  Department of Computer Science and Engineering\nSri Eshwar College of Engineering (Autonomous), Coimbatore"
    p_team.font.size = Pt(12)
    p_team.font.color.rgb = RGBColor(148, 163, 184)
    p_team.font.name = "Calibri"
    p_team.alignment = PP_ALIGN.CENTER
    p_team.space_before = Pt(20)

    # Page number bottom right
    tb_pn = s16.shapes.add_textbox(Inches(12.2), Inches(7.0), Inches(0.8), Inches(0.3))
    p = tb_pn.text_frame.paragraphs[0]
    p.text = "16"
    p.font.size = Pt(8.5)
    p.font.color.rgb = RGBColor(148, 163, 184)
    p.alignment = PP_ALIGN.RIGHT

    output_path = r"d:\projects\CrimeLens\CrimeLens_Official_Template_Final_Review.pptx"
    prs.save(output_path)
    print(f"Official template-matching presentation generated: {output_path}")

if __name__ == "__main__":
    build_presentation()

