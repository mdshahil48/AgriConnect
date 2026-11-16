import os
import sys
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

    # Theme Colors
    DARK_GREEN = RGBColor(23, 74, 44)       # #174A2C
    PRIMARY_GREEN = RGBColor(46, 125, 50)    # #2E7D32
    LIGHT_GREEN = RGBColor(232, 245, 233)    # #E8F5E9
    NAVY = RGBColor(16, 42, 67)              # #102A43
    OFF_WHITE = RGBColor(248, 250, 247)      # #F8FAF7
    ACCENT_AMBER = RGBColor(242, 166, 90)    # #F2A65A
    WHITE = RGBColor(255, 255, 255)
    BORDER_COLOR = RGBColor(200, 225, 205)   # #C8E1CD
    BORDER_LIGHT = RGBColor(220, 230, 222)
    TEXT_MUTED = RGBColor(90, 105, 115)
    CARD_BG = RGBColor(255, 255, 255)
    RED_ACCENT = RGBColor(198, 40, 40)

    def set_slide_bg(slide):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = OFF_WHITE
        bg.line.color.rgb = OFF_WHITE
        return bg

    def add_header(slide, title_text, category_text="AGRICONNECT | PROJECT PROGRESS SEMINAR – II", slide_num=1):
        # Header banner text
        cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(10), Inches(0.35))
        tf_cat = cat_box.text_frame
        tf_cat.word_wrap = True
        tf_cat.margin_left = tf_cat.margin_top = tf_cat.margin_right = tf_cat.margin_bottom = 0
        p_cat = tf_cat.paragraphs[0]
        p_cat.text = category_text.upper()
        p_cat.font.size = Pt(10)
        p_cat.font.bold = True
        p_cat.font.color.rgb = PRIMARY_GREEN
        p_cat.font.name = "Segoe UI"

        # Main Title
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.72), Inches(10), Inches(0.7))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        tf_title.margin_left = tf_title.margin_top = tf_title.margin_right = tf_title.margin_bottom = 0
        p_title = tf_title.paragraphs[0]
        p_title.text = title_text
        p_title.font.size = Pt(24)
        p_title.font.bold = True
        p_title.font.color.rgb = NAVY
        p_title.font.name = "Segoe UI"

        # Slide Number Badge (Top Right)
        badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(11.8), Inches(0.45), Inches(0.75), Inches(0.4))
        badge.fill.solid()
        badge.fill.fore_color.rgb = LIGHT_GREEN
        badge.line.color.rgb = BORDER_COLOR
        tf_b = badge.text_frame
        tf_b.margin_left = tf_b.margin_top = tf_b.margin_right = tf_b.margin_bottom = 0
        p_b = tf_b.paragraphs[0]
        p_b.text = f"{slide_num:02d}"
        p_b.font.size = Pt(12)
        p_b.font.bold = True
        p_b.font.color.rgb = DARK_GREEN
        p_b.alignment = PP_ALIGN.CENTER
        p_b.font.name = "Segoe UI"

        # Subtle bottom footer
        foot_box = slide.shapes.add_textbox(Inches(0.8), Inches(7.05), Inches(11.733), Inches(0.3))
        tf_foot = foot_box.text_frame
        tf_foot.word_wrap = True
        tf_foot.margin_left = tf_foot.margin_top = tf_foot.margin_right = tf_foot.margin_bottom = 0
        p_f = tf_foot.paragraphs[0]
        p_f.text = "AgriConnect Platform  •  Transparent Agricultural Supply Chain  •  https://agriconnect-q20v.onrender.com"
        p_f.font.size = Pt(9)
        p_f.font.color.rgb = TEXT_MUTED
        p_f.font.name = "Segoe UI"

    def add_notes(slide, notes_text):
        notes_slide = slide.notes_slide
        text_frame = notes_slide.notes_text_frame
        text_frame.text = notes_text.strip()

    # =========================================================================
    # SLIDE 1: TITLE SLIDE
    # =========================================================================
    slide1 = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide1)

    # Top accent bar
    top_bar = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.15))
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = PRIMARY_GREEN
    top_bar.line.fill.background()

    # Hero Card (Left & Center)
    hero_card = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.8), Inches(11.733), Inches(3.6))
    hero_card.fill.solid()
    hero_card.fill.fore_color.rgb = DARK_GREEN
    hero_card.line.color.rgb = DARK_GREEN

    # Seminar Tag inside Hero
    sem_box = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.3), Inches(1.2), Inches(3.4), Inches(0.38))
    sem_box.fill.solid()
    sem_box.fill.fore_color.rgb = ACCENT_AMBER
    sem_box.line.fill.background()
    tf_sb = sem_box.text_frame
    tf_sb.margin_left = tf_sb.margin_top = tf_sb.margin_right = tf_sb.margin_bottom = 0
    p_sb = tf_sb.paragraphs[0]
    p_sb.text = "PROJECT PROGRESS SEMINAR – II"
    p_sb.font.size = Pt(11)
    p_sb.font.bold = True
    p_sb.font.color.rgb = NAVY
    p_sb.alignment = PP_ALIGN.CENTER
    p_sb.font.name = "Segoe UI"

    # Project Title
    title_box = slide1.shapes.add_textbox(Inches(1.3), Inches(1.75), Inches(10.6), Inches(0.9))
    tf_t = title_box.text_frame
    tf_t.word_wrap = True
    tf_t.margin_left = tf_t.margin_top = tf_t.margin_right = tf_t.margin_bottom = 0
    p_t = tf_t.paragraphs[0]
    p_t.text = "AgriConnect"
    p_t.font.size = Pt(36)
    p_t.font.bold = True
    p_t.font.color.rgb = WHITE
    p_t.font.name = "Segoe UI"

    # Subtitle
    sub_box = slide1.shapes.add_textbox(Inches(1.3), Inches(2.55), Inches(10.6), Inches(0.75))
    tf_s = sub_box.text_frame
    tf_s.word_wrap = True
    tf_s.margin_left = tf_s.margin_top = tf_s.margin_right = tf_s.margin_bottom = 0
    p_s = tf_s.paragraphs[0]
    p_s.text = "Transparent Digital Agricultural Supply Chain Platform with Assisted-Digital Access"
    p_s.font.size = Pt(17)
    p_s.font.color.rgb = RGBColor(220, 245, 230)
    p_s.font.name = "Segoe UI"

    # Hero visual flow strip inside hero card
    flow_strip = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.3), Inches(3.35), Inches(10.7), Inches(0.65))
    flow_strip.fill.solid()
    flow_strip.fill.fore_color.rgb = RGBColor(16, 52, 31)
    flow_strip.line.color.rgb = RGBColor(46, 125, 50)
    tf_fs = flow_strip.text_frame
    tf_fs.margin_top = Inches(0.12)
    p_fs = tf_fs.paragraphs[0]
    p_fs.text = "🌱 Farmer Production   ➔   📱 Assisted VDF & IVR Access   ➔   🌐 Digital Marketplace   ➔   🚚 Cold Logistics   ➔   🛒 Urban Consumers"
    p_fs.font.size = Pt(11)
    p_fs.font.bold = True
    p_fs.font.color.rgb = WHITE
    p_fs.alignment = PP_ALIGN.CENTER
    p_fs.font.name = "Segoe UI"

    # Team Members Header Label
    tm_lbl = slide1.shapes.add_textbox(Inches(0.8), Inches(4.6), Inches(6), Inches(0.3))
    tf_tml = tm_lbl.text_frame
    tf_tml.margin_left = tf_tml.margin_top = tf_tml.margin_right = tf_tml.margin_bottom = 0
    p_tml = tf_tml.paragraphs[0]
    p_tml.text = "PROJECT TEAM MEMBERS"
    p_tml.font.size = Pt(11)
    p_tml.font.bold = True
    p_tml.font.color.rgb = PRIMARY_GREEN
    p_tml.font.name = "Segoe UI"

    # Team Members 4 Columns
    team_members = [
        {"name": "Md Aamir Alam", "enroll": "AJU/231047", "roll": "Btcs/074", "role": "Project Leader & UI Architecture"},
        {"name": "Md Shail", "enroll": "AJU/230656", "roll": "Btcs/048", "role": "Backend & Database Systems"},
        {"name": "Rokkam Madhava Rao", "enroll": "AJU/230194", "roll": "Btcs/017", "role": "Logistics & Telephony Engine"},
        {"name": "N. Rohit Kumar", "enroll": "AJU/230686", "roll": "Btcs/051", "role": "ML Analytics & QA Testing"}
    ]

    card_w = Inches(2.78)
    gap = Inches(0.2)
    start_x = Inches(0.8)

    for i, tm in enumerate(team_members):
        cx = start_x + i * (card_w + gap)
        card = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, Inches(4.95), card_w, Inches(1.85))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = BORDER_COLOR

        tb = slide1.shapes.add_textbox(cx + Inches(0.18), Inches(5.1), card_w - Inches(0.36), Inches(1.55))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p1 = tf.paragraphs[0]
        p1.text = tm["name"]
        p1.font.size = Pt(13)
        p1.font.bold = True
        p1.font.color.rgb = NAVY
        p1.font.name = "Segoe UI"

        p2 = tf.add_paragraph()
        p2.text = tm["role"]
        p2.font.size = Pt(10)
        p2.font.bold = True
        p2.font.color.rgb = PRIMARY_GREEN
        p2.font.name = "Segoe UI"
        p2.space_after = Pt(6)

        p3 = tf.add_paragraph()
        p3.text = f"Enrollment: {tm['enroll']}\nClass Roll: {tm['roll']}"
        p3.font.size = Pt(9.5)
        p3.font.color.rgb = TEXT_MUTED
        p3.font.name = "Segoe UI"

    # Slide 1 Footer
    foot = slide1.shapes.add_textbox(Inches(0.8), Inches(7.05), Inches(11.733), Inches(0.3))
    tf_f = foot.text_frame
    p_f = tf_f.paragraphs[0]
    p_f.text = "Department of Computer Science & Engineering  •  Live Deployment: https://agriconnect-q20v.onrender.com"
    p_f.font.size = Pt(9)
    p_f.font.color.rgb = TEXT_MUTED
    p_f.alignment = PP_ALIGN.CENTER
    p_f.font.name = "Segoe UI"

    add_notes(slide1, """Good morning respected faculty members and evaluators. 
Welcome to our Project Progress Seminar – II presentation. 
Our project is titled 'AgriConnect: Transparent Digital Agricultural Supply Chain Platform with Assisted-Digital Access'.
The core purpose of AgriConnect is to remove exploitative middlemen from agricultural trade by building an end-to-end transparent supply chain. It connects rural farmers directly with urban consumers, while solving the digital divide through assisted village facilitation and telephony IVR access.
Our team consists of Md Aamir Alam serving as Project Leader and leading Frontend Architecture, Md Shail heading Backend and Database Systems, Rokkam Madhava Rao heading Logistics and the Telephony Engine, and N. Rohit Kumar leading ML Analytics and QA Testing.
Let us now take you through the detailed progress and architectural implementation of AgriConnect.""")

    # =========================================================================
    # SLIDE 2: PROJECT INTRODUCTION
    # =========================================================================
    slide2 = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide2)
    add_header(slide2, "What is AgriConnect?", slide_num=2)

    # Left Summary Card
    intro_card = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(4.5), Inches(5.3))
    intro_card.fill.solid()
    intro_card.fill.fore_color.rgb = WHITE
    intro_card.line.color.rgb = BORDER_COLOR

    tb_ic = slide2.shapes.add_textbox(Inches(1.1), Inches(1.8), Inches(3.9), Inches(4.7))
    tf_ic = tb_ic.text_frame
    tf_ic.word_wrap = True
    tf_ic.margin_left = tf_ic.margin_top = tf_ic.margin_right = tf_ic.margin_bottom = 0

    p_ih = tf_ic.paragraphs[0]
    p_ih.text = "Centralized Web Platform"
    p_ih.font.size = Pt(16)
    p_ih.font.bold = True
    p_ih.font.color.rgb = DARK_GREEN
    p_ih.font.name = "Segoe UI"
    p_ih.space_after = Pt(10)

    p_ib = tf_ic.add_paragraph()
    p_ib.text = ("AgriConnect is a digital agricultural supply-chain platform designed to connect farmers, "
                 "field-support facilitators, consumers, delivery personnel, and administrators through a unified, "
                 "transparent web-based system.\n\n"
                 "Key Focus Pillars:\n"
                 "• Direct Agricultural Marketplace\n"
                 "• Farmgate Price & Cost Breakdown Transparency\n"
                 "• Assisted-Digital Access for Non-Tech Farmers\n"
                 "• Real-time Order & Logistics Coordination\n"
                 "• Role-Based Security & Permissions\n"
                 "• Data-Driven Analytics and Demand Monitoring")
    p_ib.font.size = Pt(11)
    p_ib.font.color.rgb = NAVY
    p_ib.font.name = "Segoe UI"

    # Right Diagram Area (Central Hub Architecture)
    hub_card = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.6), Inches(1.5), Inches(6.933), Inches(5.3))
    hub_card.fill.solid()
    hub_card.fill.fore_color.rgb = WHITE
    hub_card.line.color.rgb = BORDER_COLOR

    # Title of diagram
    tb_hd = slide2.shapes.add_textbox(Inches(5.9), Inches(1.75), Inches(6.3), Inches(0.35))
    tf_hd = tb_hd.text_frame
    p_hd = tf_hd.paragraphs[0]
    p_hd.text = "MULTI-STAKEHOLDER INTERACTION TOPOLOGY"
    p_hd.font.size = Pt(11)
    p_hd.font.bold = True
    p_hd.font.color.rgb = PRIMARY_GREEN
    p_hd.font.name = "Segoe UI"

    # Center Hub Box
    center_hub = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.75), Inches(3.4), Inches(2.6), Inches(1.5))
    center_hub.fill.solid()
    center_hub.fill.fore_color.rgb = DARK_GREEN
    center_hub.line.color.rgb = PRIMARY_GREEN
    tf_ch = center_hub.text_frame
    tf_ch.margin_left = tf_ch.margin_top = tf_ch.margin_right = tf_ch.margin_bottom = 0
    p_ch1 = tf_ch.paragraphs[0]
    p_ch1.text = "AgriConnect Core"
    p_ch1.font.size = Pt(14)
    p_ch1.font.bold = True
    p_ch1.font.color.rgb = WHITE
    p_ch1.alignment = PP_ALIGN.CENTER
    p_ch1.font.name = "Segoe UI"
    p_ch2 = tf_ch.add_paragraph()
    p_ch2.text = "Centralized Web Engine\nOrder & Inventory Engine\nPostgreSQL Relational DB"
    p_ch2.font.size = Pt(9)
    p_ch2.font.color.rgb = RGBColor(200, 240, 210)
    p_ch2.alignment = PP_ALIGN.CENTER
    p_ch2.font.name = "Segoe UI"

    # 5 Satellite Stakeholders
    satellites = [
        {"name": "🌾 Farmer", "sub": "Produce Listings\nOrder Payouts", "x": 8.0, "y": 2.1, "w": 2.1, "h": 0.95, "bg": LIGHT_GREEN, "border": PRIMARY_GREEN},
        {"name": "🤝 VDF / Field Support", "sub": "Assisted Digital Listing\nPhysical Verification", "x": 5.85, "y": 3.4, "w": 1.7, "h": 1.5, "bg": LIGHT_GREEN, "border": PRIMARY_GREEN},
        {"name": "🛒 Urban Consumer", "sub": "Transparent Marketplace\nDirect Purchases", "x": 10.55, "y": 3.4, "w": 1.75, "h": 1.5, "bg": LIGHT_GREEN, "border": PRIMARY_GREEN},
        {"name": "🚚 Delivery Agent", "sub": "Hub Aggregation\nMilestone Updates", "x": 6.7, "y": 5.2, "w": 2.2, "h": 1.0, "bg": LIGHT_GREEN, "border": PRIMARY_GREEN},
        {"name": "⚙️ System Admin", "sub": "Governance, Analytics\nAudit Oversight", "x": 9.3, "y": 5.2, "w": 2.2, "h": 1.0, "bg": LIGHT_GREEN, "border": PRIMARY_GREEN}
    ]

    for sat in satellites:
        s_box = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(sat["x"]), Inches(sat["y"]), Inches(sat["w"]), Inches(sat["h"]))
        s_box.fill.solid()
        s_box.fill.fore_color.rgb = sat["bg"]
        s_box.line.color.rgb = sat["border"]
        tf_s = s_box.text_frame
        tf_s.word_wrap = True
        tf_s.margin_left = tf_s.margin_top = tf_s.margin_right = tf_s.margin_bottom = 0
        p_s1 = tf_s.paragraphs[0]
        p_s1.text = sat["name"]
        p_s1.font.size = Pt(11)
        p_s1.font.bold = True
        p_s1.font.color.rgb = DARK_GREEN
        p_s1.alignment = PP_ALIGN.CENTER
        p_s1.font.name = "Segoe UI"
        p_s2 = tf_s.add_paragraph()
        p_s2.text = sat["sub"]
        p_s2.font.size = Pt(8.5)
        p_s2.font.color.rgb = NAVY
        p_s2.alignment = PP_ALIGN.CENTER
        p_s2.font.name = "Segoe UI"

    add_notes(slide2, """In this slide, we introduce what AgriConnect truly represents.
Traditional agricultural platforms only provide a catalog or a mobile app that an illiterate or rural farmer cannot easily operate. AgriConnect is different: it is an integrated web-based platform serving five key user groups.
At the center is the AgriConnect digital platform running on Django and PostgreSQL. Around it, we connect:
First, the Farmer who supplies fresh produce.
Second, the Village Digital Facilitator—a trained local youth who acts as an assisted-digital interface for farmers lacking smartphones.
Third, the Consumer who enjoys complete price transparency.
Fourth, the Delivery Personnel managing hub pickups and final delivery.
And fifth, the Administrator who monitors the overall health and analytics of the supply chain.
This creates a seamless ecosystem that ensures no user is left behind due to the digital divide.""")

    # =========================================================================
    # SLIDE 3: PROBLEM STATEMENT
    # =========================================================================
    slide3 = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide3)
    add_header(slide3, "Problem Statement: Broken Supply Chains vs. AgriConnect", slide_num=3)

    # Left: Traditional Fragmented System
    trad_card = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(5.7), Inches(5.3))
    trad_card.fill.solid()
    trad_card.fill.fore_color.rgb = WHITE
    trad_card.line.color.rgb = RGBColor(230, 180, 180)

    # Top red tag
    tag_trad = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.1), Inches(1.75), Inches(3.2), Inches(0.35))
    tag_trad.fill.solid()
    tag_trad.fill.fore_color.rgb = RGBColor(255, 235, 235)
    tag_trad.line.color.rgb = RED_ACCENT
    tf_tt = tag_trad.text_frame
    p_tt = tf_tt.paragraphs[0]
    p_tt.text = "TRADITIONAL FRAGMENTED SYSTEM"
    p_tt.font.size = Pt(9.5)
    p_tt.font.bold = True
    p_tt.font.color.rgb = RED_ACCENT
    p_tt.alignment = PP_ALIGN.CENTER
    p_tt.font.name = "Segoe UI"

    problems = [
        ("Exploitative Middlemen Multi-Tiers", "Farmers receive only 20–30% of the final retail price due to cascading middleman commissions."),
        ("Zero Price & Cost Transparency", "Neither consumers nor farmers know actual transportation, handling, or mandi price markups."),
        ("Severe Digital Exclusion Barrier", "Keypad-phone users and illiterate farmers cannot use complex smartphone mobile apps."),
        ("Uncoordinated Logistics & Waste", "25–35% of fresh vegetables rot post-harvest due to lack of synchronized cluster pickups."),
        ("Manual & Paper-Based Operations", "Orders, payments, and inventory are tracked on paper registers, leading to disputes.")
    ]

    start_py = 2.25
    for title, desc in problems:
        p_box = slide3.shapes.add_textbox(Inches(1.1), Inches(start_py), Inches(5.1), Inches(0.8))
        tf_p = p_box.text_frame
        tf_p.word_wrap = True
        tf_p.margin_left = tf_p.margin_top = tf_p.margin_right = tf_p.margin_bottom = 0
        p1 = tf_p.paragraphs[0]
        p1.text = f"❌  {title}"
        p1.font.size = Pt(11)
        p1.font.bold = True
        p1.font.color.rgb = RED_ACCENT
        p1.font.name = "Segoe UI"
        p2 = tf_p.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = NAVY
        p2.font.name = "Segoe UI"
        start_py += 0.88

    # Right: Connected AgriConnect System
    conn_card = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.5), Inches(5.7), Inches(5.3))
    conn_card.fill.solid()
    conn_card.fill.fore_color.rgb = WHITE
    conn_card.line.color.rgb = BORDER_COLOR

    tag_conn = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.1), Inches(1.75), Inches(3.2), Inches(0.35))
    tag_conn.fill.solid()
    tag_conn.fill.fore_color.rgb = LIGHT_GREEN
    tag_conn.line.color.rgb = PRIMARY_GREEN
    tf_tc = tag_conn.text_frame
    p_tc = tf_tc.paragraphs[0]
    p_tc.text = "THE AGRICONNECT SOLUTION"
    p_tc.font.size = Pt(9.5)
    p_tc.font.bold = True
    p_tc.font.color.rgb = DARK_GREEN
    p_tc.alignment = PP_ALIGN.CENTER
    p_tc.font.name = "Segoe UI"

    solutions = [
        ("Direct Farm-to-Consumer Market", "Farmers receive 60–70% of consumer spend; transactions are direct and fair."),
        ("100% Granular Rupee Breakdown", "Every product page shows farmgate price, logistics cost, quality fee, and platform charge."),
        ("Assisted-Digital & IVR Telephony", "Village Facilitators (VDF) and an interactive 1800 IVR toll-free engine bridge the digital divide."),
        ("Optimized Cluster Aggregation", "Orders are batched at local rural hubs with verified transport assignment."),
        ("Centralized Web Architecture", "Real-time order state machine, automated digital invoicing, and PostgreSQL transaction integrity.")
    ]

    start_sy = 2.25
    for title, desc in solutions:
        s_box = slide3.shapes.add_textbox(Inches(7.1), Inches(start_sy), Inches(5.1), Inches(0.8))
        tf_s = s_box.text_frame
        tf_s.word_wrap = True
        tf_s.margin_left = tf_s.margin_top = tf_s.margin_right = tf_s.margin_bottom = 0
        p1 = tf_s.paragraphs[0]
        p1.text = f"✅  {title}"
        p1.font.size = Pt(11)
        p1.font.bold = True
        p1.font.color.rgb = PRIMARY_GREEN
        p1.font.name = "Segoe UI"
        p2 = tf_s.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = NAVY
        p2.font.name = "Segoe UI"
        start_sy += 0.88

    add_notes(slide3, """On Slide 3, we address the exact problem statement that motivated this project.
In traditional agricultural supply chains, smallholder farmers face severe exploitation. Multiple layers of middlemen capture over 70% of the consumer rupee. Furthermore, farmers have zero visibility into real mandi rates, and post-harvest spoilage is rampant due to lack of coordinated logistics. Most importantly, millions of small farmers in rural areas use basic keypad phones and cannot use high-tech smartphone apps.
AgriConnect systematically tackles these issues. We provide direct trade where farmers earn up to 65% of the consumer price. We implement open price transparency where every consumer can inspect the cost breakdown of every rupee spent. And through Village Facilitators and an IVR telephony simulation, even a farmer with a 2G keypad phone can sell their harvest without touching a smartphone.""")

    # =========================================================================
    # SLIDE 4: PROPOSED SOLUTION
    # =========================================================================
    slide4 = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide4)
    add_header(slide4, "Our Proposed Solution & Pipeline Architecture", slide_num=4)

    # Top Description Card
    desc_card = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.4), Inches(11.733), Inches(1.0))
    desc_card.fill.solid()
    desc_card.fill.fore_color.rgb = WHITE
    desc_card.line.color.rgb = BORDER_COLOR
    tf_dc = desc_card.text_frame
    tf_dc.margin_top = Inches(0.15)
    tf_dc.margin_left = Inches(0.25)
    p_dc = tf_dc.paragraphs[0]
    p_dc.text = "AgriConnect delivers an integrated digital platform uniting product management, assisted cataloging, transparent ordering, and cold-chain dispatch into one synchronized workflow."
    p_dc.font.size = Pt(12)
    p_dc.font.bold = True
    p_dc.font.color.rgb = DARK_GREEN
    p_dc.font.name = "Segoe UI"

    # Pipeline Process Flow (5 Horizontal Blocks)
    pipeline_steps = [
        {"num": "STEP 1", "title": "Farmer / VDF Hub", "body": "• Produce listing creation\n• Voice & IVR assisted intake\n• Quality grade assessment\n• APMC Mandi benchmarking"},
        {"num": "STEP 2", "title": "AgriConnect Core", "body": "• Catalog publishing\n• Price corridor validation\n• Transparent cost breakdown\n• Real-time stock reservation"},
        {"num": "STEP 3", "title": "Consumer Discovery", "body": "• Clean marketplace browsing\n• Group buying pool aggregation\n• Direct basket checkout\n• Instant order confirmation"},
        {"num": "STEP 4", "title": "Logistics Dispatch", "body": "• Hub vehicle assignment\n• Driver pickup verification\n• Cold mini-van transit\n• Delivery milestone tracking"},
        {"num": "STEP 5", "title": "Fulfillment & Payout", "body": "• Doorstep consumer delivery\n• Quality confirmation\n• Direct farmer ledger payout\n• Closed-loop analytics feed"}
    ]

    step_w = Inches(2.18)
    step_gap = Inches(0.2)
    start_x = Inches(0.8)

    for i, step in enumerate(pipeline_steps):
        sx = start_x + i * (step_w + step_gap)
        box = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, sx, Inches(2.6), step_w, Inches(3.0))
        box.fill.solid()
        box.fill.fore_color.rgb = WHITE
        box.line.color.rgb = PRIMARY_GREEN if i in [0, 1] else BORDER_COLOR

        # Header tag on box
        tag = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, sx + Inches(0.12), Inches(2.72), step_w - Inches(0.24), Inches(0.35))
        tag.fill.solid()
        tag.fill.fore_color.rgb = DARK_GREEN if i in [0, 1] else LIGHT_GREEN
        tag.line.fill.background()
        tf_tg = tag.text_frame
        p_tg = tf_tg.paragraphs[0]
        p_tg.text = step["num"]
        p_tg.font.size = Pt(9.5)
        p_tg.font.bold = True
        p_tg.font.color.rgb = WHITE if i in [0, 1] else DARK_GREEN
        p_tg.alignment = PP_ALIGN.CENTER
        p_tg.font.name = "Segoe UI"

        tb = slide4.shapes.add_textbox(sx + Inches(0.12), Inches(3.18), step_w - Inches(0.24), Inches(2.3))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p1 = tf.paragraphs[0]
        p1.text = step["title"]
        p1.font.size = Pt(11)
        p1.font.bold = True
        p1.font.color.rgb = NAVY
        p1.font.name = "Segoe UI"
        p1.space_after = Pt(6)

        p2 = tf.add_paragraph()
        p2.text = step["body"]
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = TEXT_MUTED
        p2.font.name = "Segoe UI"

    # Bottom Foundation Bar: Admin Governance & ML Analytics
    bot_card = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(5.8), Inches(11.733), Inches(0.95))
    bot_card.fill.solid()
    bot_card.fill.fore_color.rgb = LIGHT_GREEN
    bot_card.line.color.rgb = PRIMARY_GREEN
    tf_bc = bot_card.text_frame
    tf_bc.margin_top = Inches(0.15)
    tf_bc.margin_left = Inches(0.25)
    p_bc1 = tf_bc.paragraphs[0]
    p_bc1.text = "Cross-Cutting System Infrastructure: System Admin Governance & Real-Time Analytics"
    p_bc1.font.size = Pt(11.5)
    p_bc1.font.bold = True
    p_bc1.font.color.rgb = DARK_GREEN
    p_bc1.font.name = "Segoe UI"
    p_bc2 = tf_bc.add_paragraph()
    p_bc2.text = "Role-based permission barriers  •  Audit trails  •  Regression-based demand projection  •  Production PostgreSQL ACID transactions"
    p_bc2.font.size = Pt(9.5)
    p_bc2.font.color.rgb = NAVY
    p_bc2.font.name = "Segoe UI"

    add_notes(slide4, """Slide 4 presents our proposed solution as a structured end-to-end pipeline.
Instead of treating supply chain stages in isolation, AgriConnect coordinates all phases through a unified state machine:
In Step 1, produce is registered either directly or through the Village Facilitator.
In Step 2, the AgriConnect backend evaluates quality grades, checks mandi benchmark rates, and publishes verified listings.
In Step 3, urban consumers browse fresh produce, view exact price breakdowns, and order individually or in group buying pools.
In Step 4, our logistics module assigns a designated vehicle and driver for aggregate pickup from rural aggregation centers.
In Step 5, doorstep delivery is completed, order status updates in real-time, and payouts are credited to the farmer.
Underpinning this entire pipeline are our Admin Governance tools and ML analytics engine.""")

    # =========================================================================
    # SLIDE 5: PROJECT OBJECTIVES
    # =========================================================================
    slide5 = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide5)
    add_header(slide5, "Strategic Project Objectives", slide_num=5)

    objectives = [
        ("1. Digitalize Agricultural Operations", "Convert manual registers, paper records, and fragmented phone calls into structured, searchable digital databases with automated inventory and tracking.", "💻"),
        ("2. Supply Chain Price Transparency", "Expose complete breakdown of every consumer rupee spent (farmer pay, transit, packaging, platform fee) to eliminate exploitative intermediary margins.", "🔍"),
        ("3. Multi-Stakeholder Integration", "Connect farmers, facilitators, buyers, drivers, and administrators under one unified, role-governed web application with tailored dashboards.", "🌐"),
        ("4. Enable Assisted-Digital Access", "Bridge the rural digital divide via Village Digital Facilitators (VDF) and an interactive 1800-toll-free IVR simulator for keypad-phone farmers.", "📞"),
        ("5. Streamline Logistics Coordination", "Automate aggregation hub dispatch, vehicle assignment, status updates, and delivery confirmations to minimize post-harvest perishability loss.", "🚚"),
        ("6. Support Data-Driven Decisions", "Provide operational dashboards, 7-day demand forecasting algorithms, and mandi price recommendations to guide harvesting and fair pricing.", "📊")
    ]

    card_w = Inches(3.7)
    card_h = Inches(2.4)
    cols = 3

    for i, obj in enumerate(objectives):
        col = i % cols
        row = i // cols
        cx = Inches(0.8) + col * Inches(3.7 + 0.31)
        cy = Inches(1.5) + row * Inches(2.4 + 0.35)

        card = slide5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, cy, card_w, card_h)
        card.fill.solid()
        card.fill.fore_color.rgb = WHITE
        card.line.color.rgb = BORDER_COLOR

        tb = slide5.shapes.add_textbox(cx + Inches(0.2), cy + Inches(0.2), card_w - Inches(0.4), card_h - Inches(0.4))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p_icon = tf.paragraphs[0]
        p_icon.text = obj[2]
        p_icon.font.size = Pt(20)
        p_icon.space_after = Pt(4)

        p1 = tf.add_paragraph()
        p1.text = obj[0]
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.font.color.rgb = DARK_GREEN
        p1.font.name = "Segoe UI"
        p1.space_after = Pt(6)

        p2 = tf.add_paragraph()
        p2.text = obj[1]
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = NAVY
        p2.font.name = "Segoe UI"

    add_notes(slide5, """Here on Slide 5, we have outlined the six core objectives of our project.
Objective 1 is to digitalize agricultural operations: converting paper-based mandi records into a reliable digital ledger.
Objective 2 is supply chain price transparency: making sure every stakeholder knows where every rupee goes.
Objective 3 is connecting multiple stakeholders under a unified, role-based web architecture.
Objective 4 is assisted digital access: empowering rural farmers who lack smartphones through local facilitators and voice IVR.
Objective 5 is logistics optimization: coordinating aggregation hubs and cold vans to slash post-harvest food waste.
And Objective 6 is supporting data-driven decisions: utilizing historical mandi prices and predictive algorithms to help farmers price their harvests profitably.""")

    # =========================================================================
    # SLIDE 6: STAKEHOLDER ECOSYSTEM
    # =========================================================================
    slide6 = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide6)
    add_header(slide6, "Stakeholder Ecosystem & Interaction Architecture", slide_num=6)

    # Center circle for AgriConnect
    cx_center = Inches(5.4)
    cy_center = Inches(3.2)
    c_w = Inches(2.6)
    c_h = Inches(2.6)
    center_circle = slide6.shapes.add_shape(MSO_SHAPE.OVAL, cx_center, cy_center, c_w, c_h)
    center_circle.fill.solid()
    center_circle.fill.fore_color.rgb = DARK_GREEN
    center_circle.line.color.rgb = PRIMARY_GREEN
    center_circle.line.width = Pt(3)
    tf_cc = center_circle.text_frame
    tf_cc.margin_top = Inches(0.65)
    p_cc1 = tf_cc.paragraphs[0]
    p_cc1.text = "AGRICONNECT"
    p_cc1.font.size = Pt(14)
    p_cc1.font.bold = True
    p_cc1.font.color.rgb = WHITE
    p_cc1.alignment = PP_ALIGN.CENTER
    p_cc1.font.name = "Segoe UI"
    p_cc2 = tf_cc.add_paragraph()
    p_cc2.text = "Centralized Web\nEcosystem Hub"
    p_cc2.font.size = Pt(10)
    p_cc2.font.color.rgb = ACCENT_AMBER
    p_cc2.alignment = PP_ALIGN.CENTER
    p_cc2.font.name = "Segoe UI"

    # 5 Stakeholder Detailed Cards surrounding center
    stakeholders_pos = [
        {"role": "FARMER", "sub": "Produce Supplier", "bullets": "• Direct harvest listing\n• Price discovery & mandi rates\n• Real-time order monitoring\n• Payout tracking & SMS receipts", "x": 0.8, "y": 1.5, "w": 3.4, "h": 2.2},
        {"role": "VDF / FIELD SUPPORT", "sub": "Assisted-Digital Agent", "bullets": "• Onboards non-tech farmers\n• Voice & assisted produce entry\n• Quality grading verification\n• Hub aggregation liaison", "x": 0.8, "y": 4.1, "w": 3.4, "h": 2.2},
        {"role": "CONSUMER", "sub": "Urban Buyer", "bullets": "• Fresh farm produce catalog\n• 100% price transparency view\n• Group buying discount pools\n• Real-time dispatch tracking", "x": 9.133, "y": 1.5, "w": 3.4, "h": 2.2},
        {"role": "DELIVERY AGENT", "sub": "Logistics Partner", "bullets": "• Hub pickup route assignment\n• Cold mini-van dispatch\n• Delivery milestone updates\n• Proof-of-delivery confirmation", "x": 9.133, "y": 4.1, "w": 3.4, "h": 2.2},
        {"role": "SYSTEM ADMIN", "sub": "Platform Governance", "bullets": "• User role & credential control\n• Catalog & dispute resolution\n• System audit trail monitoring\n• Operational analytics oversight", "x": 4.6, "y": 6.0, "w": 4.2, "h": 1.25}
    ]

    for s in stakeholders_pos:
        card = slide6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(s["x"]), Inches(s["y"]), Inches(s["w"]), Inches(s["h"]))
        card.fill.solid()
        card.fill.fore_color.rgb = WHITE
        card.line.color.rgb = BORDER_COLOR

        tb = slide6.shapes.add_textbox(Inches(s["x"]) + Inches(0.18), Inches(s["y"]) + Inches(0.15), Inches(s["w"]) - Inches(0.36), Inches(s["h"]) - Inches(0.3))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p1 = tf.paragraphs[0]
        p1.text = s["role"]
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.font.color.rgb = DARK_GREEN
        p1.font.name = "Segoe UI"

        p2 = tf.add_paragraph()
        p2.text = s["sub"]
        p2.font.size = Pt(9)
        p2.font.bold = True
        p2.font.color.rgb = PRIMARY_GREEN
        p2.font.name = "Segoe UI"
        p2.space_after = Pt(4)

        p3 = tf.add_paragraph()
        p3.text = s["bullets"]
        p3.font.size = Pt(9)
        p3.font.color.rgb = NAVY
        p3.font.name = "Segoe UI"

    add_notes(slide6, """Slide 6 illustrates our Stakeholder Ecosystem.
AgriConnect's innovation lies in how it balances the needs of five distinct actors:
The Farmer gains market access, fair pricing, and clear payment records without losing commission to middlemen.
The Village Digital Facilitator—or VDF—acts as the human catalyst, registering harvests and providing physical quality checks for farmers in remote hamlets.
The Consumer receives freshly harvested vegetables with absolute certainty regarding quality and ethical pricing.
The Delivery Personnel receive structured pickup routes from village aggregation hubs directly to city distribution points.
And the Administrator maintains system integrity, resolves quality disputes, and oversees platform security.
Every interaction is captured within the AgriConnect central platform.""")

    # =========================================================================
    # SLIDE 7: SYSTEM ARCHITECTURE
    # =========================================================================
    slide7 = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide7)
    add_header(slide7, "Multi-Tier Layered System Architecture", slide_num=7)

    layers = [
        ("LAYER 1: USER ACTORS", "Farmer (Direct / IVR)  |  Village Facilitator (VDF)  |  Consumer  |  Delivery Personnel  |  Platform Admin", PRIMARY_GREEN, LIGHT_GREEN, DARK_GREEN),
        ("LAYER 2: PRESENTATION LAYER", "Responsive HTML5, Vanilla CSS3 (Custom Design System, Glassmorphic Cards), JavaScript (ES6+), AudioContext (DTMF Web Audio API), Web Speech API", NAVY, WHITE, NAVY),
        ("LAYER 3: APPLICATION & BUSINESS LOGIC (DJANGO 5.x)", "Django Framework (MTV Architecture)  •  Custom Role Authentication & Decorators  •  Order State Machine  •  Inventory Management  •  IVR Simulation Controller  •  REST Endpoints", DARK_GREEN, LIGHT_GREEN, DARK_GREEN),
        ("LAYER 4: DATA PERSISTENCE LAYER (POSTGRESQL)", "PostgreSQL Relational DB  •  Django ORM  •  ACID Transactions  •  Normalized Relational Schemas (CustomUser, Product, Order, DeliveryAssignment, FarmerSMSLog)", NAVY, WHITE, NAVY),
        ("LAYER 5: ANALYTICS, MONITORING & CLOUD DEPLOYMENT", "Regression Demand Forecaster  •  APMC Mandi Price Analysis  •  Gunicorn WSGI Server  •  WhiteNoise Static Engine  •  Render Cloud PaaS (Auto CI/CD)", DARK_GREEN, LIGHT_GREEN, DARK_GREEN)
    ]

    ly_start = 1.45
    for i, (ltitle, ldesc, bcolor, fcolor, tcolor) in enumerate(layers):
        lcard = slide7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(ly_start), Inches(11.733), Inches(0.95))
        lcard.fill.solid()
        lcard.fill.fore_color.rgb = fcolor
        lcard.line.color.rgb = bcolor
        lcard.line.width = Pt(1.5)

        tb = slide7.shapes.add_textbox(Inches(1.1), Inches(ly_start + 0.12), Inches(11.133), Inches(0.75))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p1 = tf.paragraphs[0]
        p1.text = ltitle
        p1.font.size = Pt(11)
        p1.font.bold = True
        p1.font.color.rgb = tcolor
        p1.font.name = "Segoe UI"

        p2 = tf.add_paragraph()
        p2.text = ldesc
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = NAVY if tcolor != NAVY else DARK_GREEN
        p2.font.name = "Segoe UI"

        ly_start += 1.08

    add_notes(slide7, """Slide 7 presents the formal System Architecture of AgriConnect.
We have designed a clean, production-grade 5-layer architecture:
At Layer 1, diverse users access the platform through smartphones, computers, or feature phone telephone simulations.
At Layer 2, the presentation layer uses modern responsive HTML5, pure CSS3 design tokens, and JavaScript utilizing the Web Audio API for DTMF telephony and the Web Speech API for voice synthesis.
At Layer 3, our Django 5 application handles business logic, security decorators, role-based routing, and order processing.
At Layer 4, PostgreSQL ensures strict relational integrity with ACID transactions across users, orders, produce listings, and delivery assignments.
At Layer 5, our analytics engine runs predictive demand algorithms, while Gunicorn, WhiteNoise, and Render Cloud handle scalable deployment.
This decoupled design guarantees high maintainability and security.""")

    # =========================================================================
    # SLIDE 8: CORE MODULES
    # =========================================================================
    slide8 = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide8)
    add_header(slide8, "Core Functional Modules", slide_num=8)

    modules = [
        ("1. Authentication & Role Control", "• CustomUser model with 5 role types\n• Secure role-based view decorators\n• Profile isolation & session management\n• Fast role-switching simulator for demos"),
        ("2. Farmer & Produce Hub", "• Direct produce listing with image uploads\n• Produce grading & stock management\n• APMC Mandi daily rate benchmark comparator\n• Order status & real-time payout ledger"),
        ("3. Consumer & Order Engine", "• Filterable marketplace catalog\n• Complete price transparency cost breakdown\n• Dynamic shopping basket & checkout\n• Group buying pool volume discounts"),
        ("4. Logistics & Delivery Engine", "• Village aggregation hub routing\n• Vehicle & driver dispatch assignment\n• Pickup verification & driver notes\n• Delivery milestone tracking (Ordered ➔ Delivered)"),
        ("5. Assisted Access & Telephony (IVR)", "• Village Facilitator (VDF) assisted entry\n• Interactive 1800-890-KISAN phone dialer\n• DTMF key tones & multi-lingual speech output\n• Automated transactional SMS logging"),
        ("6. Analytics, QA & Admin Governance", "• 7-day regression demand forecast engine\n• Platform transaction & revenue monitoring\n• Consumer quality dispute resolution system\n• Automated end-to-end regression test suite")
    ]

    m_w = Inches(3.7)
    m_h = Inches(2.4)
    for i, mod in enumerate(modules):
        col = i % 3
        row = i // 3
        mx = Inches(0.8) + col * Inches(3.7 + 0.31)
        my = Inches(1.5) + row * Inches(2.4 + 0.35)

        mcard = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, mx, my, m_w, m_h)
        mcard.fill.solid()
        mcard.fill.fore_color.rgb = WHITE
        mcard.line.color.rgb = BORDER_COLOR

        tb = slide8.shapes.add_textbox(mx + Inches(0.2), my + Inches(0.2), m_w - Inches(0.4), m_h - Inches(0.4))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p1 = tf.paragraphs[0]
        p1.text = mod[0]
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.font.color.rgb = DARK_GREEN
        p1.font.name = "Segoe UI"
        p1.space_after = Pt(6)

        p2 = tf.add_paragraph()
        p2.text = mod[1]
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = NAVY
        p2.font.name = "Segoe UI"

    add_notes(slide8, """Slide 8 covers the six core modules that make up AgriConnect.
Module 1 handles Authentication and Role Management, ensuring farmers, consumers, facilitators, delivery drivers, and admins only see authorized views.
Module 2 is the Farmer and Produce Hub, supporting harvest registration, APMC mandi benchmark pricing, and farmer sales history.
Module 3 is the Consumer Marketplace and Order Engine, featuring our signature transparent cost breakdown and group buying pools.
Module 4 is the Logistics and Delivery Engine, managing vehicle assignments and delivery lifecycle statuses.
Module 5 is our unique Assisted Access and Telephony module, which includes the interactive IVR feature-phone simulator and facilitator dashboard.
And Module 6 is the Analytics and Admin Governance module, providing automated demand forecasting and end-to-end quality assurance.""")

    # =========================================================================
    # SLIDE 9: END-TO-END WORKFLOW
    # =========================================================================
    slide9 = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide9)
    add_header(slide9, "End-to-End AgriConnect Business Workflow", slide_num=9)

    flow_steps = [
        ("01", "Produce Intake", "Farmer lists crop or VDF registers via voice / IVR intake.", DARK_GREEN),
        ("02", "Verification", "Quality grade checked; Mandi benchmark comparison.", PRIMARY_GREEN),
        ("03", "Catalog Active", "Produce published with 100% transparent cost breakdown.", DARK_GREEN),
        ("04", "Order Placed", "Consumer adds to basket or joins Group Buying pool.", PRIMARY_GREEN),
        ("05", "Hub Aggregation", "Produce consolidated at Village Digital Aggregation Hub.", DARK_GREEN),
        ("06", "Dispatch Assigned", "Delivery driver & cold-chain vehicle allocated.", PRIMARY_GREEN),
        ("07", "Fulfillment", "Consumer delivery completed with digital signature.", DARK_GREEN),
        ("08", "Payout & Feed", "Farmer ledger credited; Data feeds ML demand model.", PRIMARY_GREEN),
    ]

    # Two rows of 4 steps
    card_fw = Inches(2.7)
    card_fh = Inches(2.3)
    gap_fx = Inches(0.31)
    gap_fy = Inches(0.4)

    for i, (num, title, desc, colr) in enumerate(flow_steps):
        row = i // 4
        col = i % 4
        fx = Inches(0.8) + col * (card_fw + gap_fx)
        fy = Inches(1.6) + row * (card_fh + gap_fy)

        fcard = slide9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, fx, fy, card_fw, card_fh)
        fcard.fill.solid()
        fcard.fill.fore_color.rgb = WHITE
        fcard.line.color.rgb = BORDER_COLOR

        # Number badge
        nbadge = slide9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, fx + Inches(0.18), fy + Inches(0.18), Inches(0.65), Inches(0.4))
        nbadge.fill.solid()
        nbadge.fill.fore_color.rgb = colr
        nbadge.line.fill.background()
        tf_nb = nbadge.text_frame
        p_nb = tf_nb.paragraphs[0]
        p_nb.text = num
        p_nb.font.size = Pt(11)
        p_nb.font.bold = True
        p_nb.font.color.rgb = WHITE
        p_nb.alignment = PP_ALIGN.CENTER
        p_nb.font.name = "Segoe UI"

        tb = slide9.shapes.add_textbox(fx + Inches(0.18), fy + Inches(0.7), card_fw - Inches(0.36), card_fh - Inches(0.85))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p1 = tf.paragraphs[0]
        p1.text = title
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.font.color.rgb = NAVY
        p1.font.name = "Segoe UI"
        p1.space_after = Pt(4)

        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = TEXT_MUTED
        p2.font.name = "Segoe UI"

    add_notes(slide9, """Slide 9 details our complete 8-step business workflow.
Let us walk through an order lifecycle:
First, a farmer in a village like Angara registers 500 kg of fresh tomatoes, either via web or assisted voice dialer.
Second, the listing undergoes quality grading and is compared against live APMC mandi benchmarks.
Third, the item is published on our marketplace with full rupee cost transparency.
Fourth, an urban buyer places an order or joins a neighborhood group pool.
Fifth, produce is gathered at the village aggregation center.
Sixth, a logistics vehicle is dispatched.
Seventh, the order is delivered to the customer's doorstep in Ranchi.
And eighth, the payment is credited directly into the farmer's account, while transaction data feeds our ML demand models for future seasonal planning.
This closed-loop system ensures efficiency at every link of the chain.""")

    # =========================================================================
    # SLIDE 10: TEAM STRUCTURE
    # =========================================================================
    slide10 = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide10)
    add_header(slide10, "Team Structure & Engineering Hierarchy", slide_num=10)

    # Top Leader Card
    leader_card = slide10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.4), Inches(1.5), Inches(4.5), Inches(1.6))
    leader_card.fill.solid()
    leader_card.fill.fore_color.rgb = DARK_GREEN
    leader_card.line.color.rgb = PRIMARY_GREEN
    tf_lc = leader_card.text_frame
    tf_lc.margin_top = Inches(0.18)
    p_lc1 = tf_lc.paragraphs[0]
    p_lc1.text = "PROJECT LEADER"
    p_lc1.font.size = Pt(11)
    p_lc1.font.bold = True
    p_lc1.font.color.rgb = ACCENT_AMBER
    p_lc1.alignment = PP_ALIGN.CENTER
    p_lc1.font.name = "Segoe UI"

    p_lc2 = tf_lc.add_paragraph()
    p_lc2.text = "Md Aamir Alam"
    p_lc2.font.size = Pt(15)
    p_lc2.font.bold = True
    p_lc2.font.color.rgb = WHITE
    p_lc2.alignment = PP_ALIGN.CENTER
    p_lc2.font.name = "Segoe UI"

    p_lc3 = tf_lc.add_paragraph()
    p_lc3.text = "Enrollment: AJU/231047  |  Class Roll: Btcs/074"
    p_lc3.font.size = Pt(9.5)
    p_lc3.font.color.rgb = RGBColor(220, 240, 225)
    p_lc3.alignment = PP_ALIGN.CENTER
    p_lc3.font.name = "Segoe UI"

    # 4 Vertical Engineering Cards
    verticals = [
        {"role": "Frontend & UI Architecture", "lead": "Md Aamir Alam", "enroll": "AJU/231047", "roll": "Btcs/074", "focus": "• Design system & tokens\n• Responsive layout & components\n• Telephony simulation UI\n• Multi-persona role dashboards"},
        {"role": "Backend & Database", "lead": "Md Shail", "enroll": "AJU/230656", "roll": "Btcs/048", "focus": "• Django MTV architecture\n• PostgreSQL ORM & migrations\n• Auth & role-based access\n• REST endpoints & order logic"},
        {"role": "Logistics & Telephony", "lead": "Rokkam Madhava Rao", "enroll": "AJU/230194", "roll": "Btcs/017", "focus": "• Vehicle dispatch workflows\n• Order fulfillment tracking\n• IVR DTMF keypad engine\n• Voice & SMS assisted intake"},
        {"role": "ML Analytics & QA", "lead": "N. Rohit Kumar", "enroll": "AJU/230686", "roll": "Btcs/051", "focus": "• 7-day regression demand model\n• APMC Mandi price benchmarks\n• Automated test suite (186+ checks)\n• Integration & regression QA"}
    ]

    card_vw = Inches(2.78)
    gap_v = Inches(0.2)
    start_vx = Inches(0.8)

    for i, vert in enumerate(verticals):
        vx = start_vx + i * (card_vw + gap_v)
        vcard = slide10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, vx, Inches(3.6), card_vw, Inches(3.2))
        vcard.fill.solid()
        vcard.fill.fore_color.rgb = WHITE
        vcard.line.color.rgb = BORDER_COLOR

        tb = slide10.shapes.add_textbox(vx + Inches(0.15), Inches(3.75), card_vw - Inches(0.3), Inches(2.9))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p1 = tf.paragraphs[0]
        p1.text = vert["role"]
        p1.font.size = Pt(11)
        p1.font.bold = True
        p1.font.color.rgb = DARK_GREEN
        p1.font.name = "Segoe UI"
        p1.space_after = Pt(2)

        p2 = tf.add_paragraph()
        p2.text = vert["lead"]
        p2.font.size = Pt(12)
        p2.font.bold = True
        p2.font.color.rgb = NAVY
        p2.font.name = "Segoe UI"

        p3 = tf.add_paragraph()
        p3.text = f"{vert['enroll']} | {vert['roll']}"
        p3.font.size = Pt(8.5)
        p3.font.color.rgb = PRIMARY_GREEN
        p3.font.name = "Segoe UI"
        p3.space_after = Pt(6)

        p4 = tf.add_paragraph()
        p4.text = vert["focus"]
        p4.font.size = Pt(9)
        p4.font.color.rgb = TEXT_MUTED
        p4.font.name = "Segoe UI"

    add_notes(slide10, """Slide 10 presents our Team Structure and Engineering Hierarchy.
To execute a project of this scale effectively, we divided our engineering responsibilities across four distinct pillars under unified leadership.
Md Aamir Alam serves as the Project Leader, driving the overall vision, milestone tracking, and cross-module integration, while also personally architecting the Frontend & UI layer.
Md Shail is in charge of the Backend and Database Systems, building our Django services and PostgreSQL models.
Rokkam Madhava Rao leads the Logistics & Telephony Engine, managing delivery dispatch models and assisted IVR voice integration.
And N. Rohit Kumar heads ML Analytics & QA Testing, responsible for demand forecasting models and our automated end-to-end testing suite.
In the subsequent slides, each team member's role and contributions are presented in depth.""")

    # =========================================================================
    # SLIDE 11: ROLE: PROJECT LEADER
    # =========================================================================
    slide11 = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide11)
    add_header(slide11, "Role Profile: Md Aamir Alam — Project Leader", slide_num=11)

    # Left: Identity & Scope Card
    id_card = slide11.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(3.8), Inches(5.3))
    id_card.fill.solid()
    id_card.fill.fore_color.rgb = DARK_GREEN
    id_card.line.color.rgb = PRIMARY_GREEN

    tb_id = slide11.shapes.add_textbox(Inches(1.05), Inches(1.75), Inches(3.3), Inches(4.8))
    tf_id = tb_id.text_frame
    tf_id.word_wrap = True
    tf_id.margin_left = tf_id.margin_top = tf_id.margin_right = tf_id.margin_bottom = 0

    p_id1 = tf_id.paragraphs[0]
    p_id1.text = "Md Aamir Alam"
    p_id1.font.size = Pt(18)
    p_id1.font.bold = True
    p_id1.font.color.rgb = WHITE
    p_id1.font.name = "Segoe UI"

    p_id2 = tf_id.add_paragraph()
    p_id2.text = "PROJECT LEADER"
    p_id2.font.size = Pt(11)
    p_id2.font.bold = True
    p_id2.font.color.rgb = ACCENT_AMBER
    p_id2.font.name = "Segoe UI"
    p_id2.space_after = Pt(10)

    p_id3 = tf_id.add_paragraph()
    p_id3.text = "Official Credentials:\n• Enrollment No: AJU/231047\n• Class Roll No: Btcs/074\n• Degree: B.Tech CSE\n\nLeadership Focus:\n• Architectural Vision & PRD\n• Cross-Functional Coordination\n• Sprint & Milestone Review\n• Quality & Release Readiness\n• College Evaluation Prep"
    p_id3.font.size = Pt(10.5)
    p_id3.font.color.rgb = RGBColor(220, 245, 230)
    p_id3.font.name = "Segoe UI"

    # Right: Detailed Responsibilities & Execution Pipeline
    resp_card = slide11.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.9), Inches(1.5), Inches(7.633), Inches(5.3))
    resp_card.fill.solid()
    resp_card.fill.fore_color.rgb = WHITE
    resp_card.line.color.rgb = BORDER_COLOR

    tb_r = slide11.shapes.add_textbox(Inches(5.2), Inches(1.75), Inches(7.0), Inches(4.8))
    tf_r = tb_r.text_frame
    tf_r.word_wrap = True
    tf_r.margin_left = tf_r.margin_top = tf_r.margin_right = tf_r.margin_bottom = 0

    p_rh = tf_r.paragraphs[0]
    p_rh.text = "CORE RESPONSIBILITIES & DELIVERABLES"
    p_rh.font.size = Pt(12)
    p_rh.font.bold = True
    p_rh.font.color.rgb = DARK_GREEN
    p_rh.font.name = "Segoe UI"
    p_rh.space_after = Pt(6)

    duties_leader = [
        "1. Project Vision & Objectives", "Authored comprehensive Project Requirements Document (PRD) and architecture roadmap aligning technical features with rural agricultural needs.",
        "2. Milestone & Sprint Planning", "Scheduled engineering phases across Semester milestones; established task backlogs and technical deliverables.",
        "3. Cross-Module Coordination", "Coordinated team members across frontend, backend, telephony, logistics, and analytics to prevent architectural bottlenecks.",
        "4. Integration Oversight", "Supervised API contract alignment between Django backend views and frontend responsive interfaces.",
        "5. Quality & Release Verification", "Conducted peer code reviews, deployment verification on Render Cloud, and managed documentation integrity."
    ]

    for i in range(0, len(duties_leader), 2):
        pd = tf_r.add_paragraph()
        pd.text = duties_leader[i]
        pd.font.size = Pt(11)
        pd.font.bold = True
        pd.font.color.rgb = NAVY
        pd.font.name = "Segoe UI"

        pb = tf_r.add_paragraph()
        pb.text = duties_leader[i+1]
        pb.font.size = Pt(9.5)
        pb.font.color.rgb = TEXT_MUTED
        pb.font.name = "Segoe UI"
        pb.space_after = Pt(4)

    # Visual Flow at bottom
    flow_box = slide11.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.2), Inches(5.85), Inches(7.0), Inches(0.75))
    flow_box.fill.solid()
    flow_box.fill.fore_color.rgb = LIGHT_GREEN
    flow_box.line.color.rgb = PRIMARY_GREEN
    tf_fb = flow_box.text_frame
    tf_fb.margin_top = Inches(0.12)
    p_fb = tf_fb.paragraphs[0]
    p_fb.text = "Project Lifecycle: Planning  ➔  Architecture  ➔  Integration  ➔  QA Testing  ➔  Cloud Deployment"
    p_fb.font.size = Pt(10)
    p_fb.font.bold = True
    p_fb.font.color.rgb = DARK_GREEN
    p_fb.alignment = PP_ALIGN.CENTER
    p_fb.font.name = "Segoe UI"

    add_notes(slide11, """As the Project Leader, my primary responsibility has been to steer the overall vision, engineering coordination, and execution of AgriConnect.
I established our Project Requirements Document and mapped out our development milestones so that every module—from backend models to frontend dashboards, telephony, and analytics—developed in synchronization.
I actively coordinated between our team members, resolved architectural friction points between Django endpoints and UI templates, verified our production deployment on Render, and ensured our system strictly delivers what rural farmers and urban consumers actually need.
Through structured sprint reviews and testing oversight, I ensured that our solution reached a high standard of quality, reliability, and readiness for today's evaluation.""")

    # =========================================================================
    # SLIDE 12: ROLE: FRONTEND & UI ARCHITECTURE
    # =========================================================================
    slide12 = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide12)
    add_header(slide12, "Role Profile: Md Aamir Alam — Frontend & UI Architecture", slide_num=12)

    # Left: Identity Card
    id_card12 = slide12.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(3.8), Inches(5.3))
    id_card12.fill.solid()
    id_card12.fill.fore_color.rgb = NAVY
    id_card12.line.color.rgb = PRIMARY_GREEN

    tb_id12 = slide12.shapes.add_textbox(Inches(1.05), Inches(1.75), Inches(3.3), Inches(4.8))
    tf_id12 = tb_id12.text_frame
    tf_id12.word_wrap = True
    tf_id12.margin_left = tf_id12.margin_top = tf_id12.margin_right = tf_id12.margin_bottom = 0

    p_id12_1 = tf_id12.paragraphs[0]
    p_id12_1.text = "Md Aamir Alam"
    p_id12_1.font.size = Pt(18)
    p_id12_1.font.bold = True
    p_id12_1.font.color.rgb = WHITE
    p_id12_1.font.name = "Segoe UI"

    p_id12_2 = tf_id12.add_paragraph()
    p_id12_2.text = "FRONTEND & UI ARCHITECTURE"
    p_id12_2.font.size = Pt(11)
    p_id12_2.font.bold = True
    p_id12_2.font.color.rgb = ACCENT_AMBER
    p_id12_2.font.name = "Segoe UI"
    p_id12_2.space_after = Pt(10)

    p_id12_3 = tf_id12.add_paragraph()
    p_id12_3.text = "Official Credentials:\n• Enrollment No: AJU/231047\n• Class Roll No: Btcs/074\n• Technology: HTML5, CSS3, ES6 JS\n\nUI Architecture Highlights:\n• Custom AgriTech Design System\n• High-Contrast Rural Accessibility\n• Interactive Phone Frame Simulation\n• Role-Specific Dashboard Views\n• Mobile-Responsive Viewports"
    p_id12_3.font.size = Pt(10.5)
    p_id12_3.font.color.rgb = RGBColor(220, 240, 255)
    p_id12_3.font.name = "Segoe UI"

    # Right: Responsibilities
    resp_card12 = slide12.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.9), Inches(1.5), Inches(7.633), Inches(5.3))
    resp_card12.fill.solid()
    resp_card12.fill.fore_color.rgb = WHITE
    resp_card12.line.color.rgb = BORDER_COLOR

    tb_r12 = slide12.shapes.add_textbox(Inches(5.2), Inches(1.75), Inches(7.0), Inches(4.8))
    tf_r12 = tb_r12.text_frame
    tf_r12.word_wrap = True
    tf_r12.margin_left = tf_r12.margin_top = tf_r12.margin_right = tf_r12.margin_bottom = 0

    p_rh12 = tf_r12.paragraphs[0]
    p_rh12.text = "TECHNICAL UI SPECIFICATIONS & ARCHITECTURE"
    p_rh12.font.size = Pt(12)
    p_rh12.font.bold = True
    p_rh12.font.color.rgb = DARK_GREEN
    p_rh12.font.name = "Segoe UI"
    p_rh12.space_after = Pt(6)

    duties_ui = [
        "1. Design System & Style Tokens", "Built a cohesive design system using curated palette tokens (#174A2C forest green, #2E7D32 leaf green, #F8FAF7 background), avoiding generic UI cliches.",
        "2. Role-Based Navigation & Dashboards", "Constructed custom UI dashboards tailored to each persona: Farmer sales portal, VDF facilitation hub, Consumer catalog, Delivery route panel, and Admin overview.",
        "3. Interactive IVR Phone Simulator UI", "Engineered an on-screen feature phone dialer with CSS-rendered casing, green LCD monitor, dynamic audio waveform bars, and responsive keypad buttons.",
        "4. Price Transparency Component", "Designed the interactive cost breakdown widget exposing farmer base pay, transit fees, and platform margins for total transparency.",
        "5. Usability & Cross-Device Testing", "Validated UI responsiveness across mobile, tablet, and widescreen viewports using modern flexbox, CSS grid, and touch-friendly controls."
    ]

    for i in range(0, len(duties_ui), 2):
        pd = tf_r12.add_paragraph()
        pd.text = duties_ui[i]
        pd.font.size = Pt(11)
        pd.font.bold = True
        pd.font.color.rgb = NAVY
        pd.font.name = "Segoe UI"

        pb = tf_r12.add_paragraph()
        pb.text = duties_ui[i+1]
        pb.font.size = Pt(9.5)
        pb.font.color.rgb = TEXT_MUTED
        pb.font.name = "Segoe UI"
        pb.space_after = Pt(4)

    # UI Pipeline Flow Box
    flow_box12 = slide12.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.2), Inches(5.85), Inches(7.0), Inches(0.75))
    flow_box12.fill.solid()
    flow_box12.fill.fore_color.rgb = LIGHT_GREEN
    flow_box12.line.color.rgb = PRIMARY_GREEN
    tf_fb12 = flow_box12.text_frame
    tf_fb12.margin_top = Inches(0.12)
    p_fb12 = tf_fb12.paragraphs[0]
    p_fb12.text = "UI Hierarchy: User Intent  ➔  Nav Bar  ➔  Role View  ➔  Component Cards  ➔  Django MTV Bridge"
    p_fb12.font.size = Pt(10)
    p_fb12.font.bold = True
    p_fb12.font.color.rgb = DARK_GREEN
    p_fb12.alignment = PP_ALIGN.CENTER
    p_fb12.font.name = "Segoe UI"

    add_notes(slide12, """In my second role as Frontend and UI Architect, my focus was to build an interface that was both aesthetically modern and highly intuitive for rural users.
Rather than using heavy frameworks or cookie-cutter templates, I built a lightweight, high-performance UI using semantic HTML5, pure CSS3, and JavaScript.
I established a custom AgriTech color system with high-contrast accessibility so sunlight on rural fields wouldn't hinder legibility.
I engineered the specialized UI for our IVR feature-phone simulator, complete with realistic dialer buttons and sound-wave animations.
I also created the distinct dashboards for all five user personas and designed our flagship 'Where Does Your Rupee Go?' price transparency widget.
Every template is fully responsive, lightweight, and cleanly wired into the Django backend.""")

    # =========================================================================
    # SLIDE 13: ROLE: BACKEND & DATABASE
    # =========================================================================
    slide13 = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide13)
    add_header(slide13, "Role Profile: Md Shail — Backend & Database Systems", slide_num=13)

    # Left: Identity Card
    id_card13 = slide13.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(3.8), Inches(5.3))
    id_card13.fill.solid()
    id_card13.fill.fore_color.rgb = DARK_GREEN
    id_card13.line.color.rgb = PRIMARY_GREEN

    tb_id13 = slide13.shapes.add_textbox(Inches(1.05), Inches(1.75), Inches(3.3), Inches(4.8))
    tf_id13 = tb_id13.text_frame
    tf_id13.word_wrap = True
    tf_id13.margin_left = tf_id13.margin_top = tf_id13.margin_right = tf_id13.margin_bottom = 0

    p_id13_1 = tf_id13.paragraphs[0]
    p_id13_1.text = "Md Shail"
    p_id13_1.font.size = Pt(18)
    p_id13_1.font.bold = True
    p_id13_1.font.color.rgb = WHITE
    p_id13_1.font.name = "Segoe UI"

    p_id13_2 = tf_id13.add_paragraph()
    p_id13_2.text = "BACKEND & DATABASE SYSTEMS"
    p_id13_2.font.size = Pt(11)
    p_id13_2.font.bold = True
    p_id13_2.font.color.rgb = ACCENT_AMBER
    p_id13_2.font.name = "Segoe UI"
    p_id13_2.space_after = Pt(10)

    p_id13_3 = tf_id13.add_paragraph()
    p_id13_3.text = "Official Credentials:\n• Enrollment No: AJU/230656\n• Class Roll No: Btcs/048\n• Stack: Python 3.12, Django 5.x, PostgreSQL\n\nDatabase & Backend Highlights:\n• Normalized Relational Schema\n• ACID Transaction Management\n• Role-Based Access Control (RBAC)\n• Django ORM Optimization\n• Secure Production Migrations"
    p_id13_3.font.size = Pt(10.5)
    p_id13_3.font.color.rgb = RGBColor(220, 245, 230)
    p_id13_3.font.name = "Segoe UI"

    # Right: Responsibilities
    resp_card13 = slide13.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.9), Inches(1.5), Inches(7.633), Inches(5.3))
    resp_card13.fill.solid()
    resp_card13.fill.fore_color.rgb = WHITE
    resp_card13.line.color.rgb = BORDER_COLOR

    tb_r13 = slide13.shapes.add_textbox(Inches(5.2), Inches(1.75), Inches(7.0), Inches(4.8))
    tf_r13 = tb_r13.text_frame
    tf_r13.word_wrap = True
    tf_r13.margin_left = tf_r13.margin_top = tf_r13.margin_right = tf_r13.margin_bottom = 0

    p_rh13 = tf_r13.paragraphs[0]
    p_rh13.text = "BACKEND ARCHITECTURE & DATABASE SYSTEMS"
    p_rh13.font.size = Pt(12)
    p_rh13.font.bold = True
    p_rh13.font.color.rgb = DARK_GREEN
    p_rh13.font.name = "Segoe UI"
    p_rh13.space_after = Pt(6)

    duties_be = [
        "1. Relational Database Modeling & Schema", "Designed normalized PostgreSQL schemas: CustomUser (Role Enum), FarmerProfile, Product, Category, Order, OrderItem, QualityDispute, and GroupBuyingPool.",
        "2. Django Application & Business Logic", "Built robust Django views and business logic controllers governing order placement, produce pricing calculations, and transaction settlement.",
        "3. Authentication & RBAC Security", "Implemented role-based permission verification using custom decorators (@login_required, role checks) ensuring strict boundary separation.",
        "4. Data Integrity & Migrations", "Constructed comprehensive Django migrations, foreign key cascades, and atomic database transactions (select_for_update) preventing race conditions.",
        "5. PostgreSQL Integration & Cloud Config", "Configured dj-database-url and psycopg2-binary for production cloud database connectivity on Render, ensuring fast indexed queries."
    ]

    for i in range(0, len(duties_be), 2):
        pd = tf_r13.add_paragraph()
        pd.text = duties_be[i]
        pd.font.size = Pt(11)
        pd.font.bold = True
        pd.font.color.rgb = NAVY
        pd.font.name = "Segoe UI"

        pb = tf_r13.add_paragraph()
        pb.text = duties_be[i+1]
        pb.font.size = Pt(9.5)
        pb.font.color.rgb = TEXT_MUTED
        pb.font.name = "Segoe UI"
        pb.space_after = Pt(4)

    # Database Pipeline Box
    flow_box13 = slide13.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.2), Inches(5.85), Inches(7.0), Inches(0.75))
    flow_box13.fill.solid()
    flow_box13.fill.fore_color.rgb = LIGHT_GREEN
    flow_box13.line.color.rgb = PRIMARY_GREEN
    tf_fb13 = flow_box13.text_frame
    tf_fb13.margin_top = Inches(0.12)
    p_fb13 = tf_fb13.paragraphs[0]
    p_fb13.text = "Data Pipeline: HTTP Request  ➔  Security Decorators  ➔  Business Views  ➔  Django ORM  ➔  PostgreSQL"
    p_fb13.font.size = Pt(10)
    p_fb13.font.bold = True
    p_fb13.font.color.rgb = DARK_GREEN
    p_fb13.alignment = PP_ALIGN.CENTER
    p_fb13.font.name = "Segoe UI"

    add_notes(slide13, """Hello evaluators, I am Md Shail, responsible for Backend and Database Systems.
My role focused on architecting a reliable data persistence layer and high-performance server logic using Django and PostgreSQL.
I designed our relational database schema across ten core tables, ensuring normalized relationships between users, farmer profiles, products, orders, and delivery records.
To prevent concurrency issues during simultaneous checkout of perishable harvest batches, I implemented atomic transactions and database constraints.
I also developed our role-based authentication system, writing custom access decorators so that only verified actors can execute privileged operations like order dispatch or price modifications.
Finally, I configured our production database on Render using dj-database-url and psycopg2, ensuring reliable and secure data operations.""")

    # =========================================================================
    # SLIDE 14: ROLE: LOGISTICS & TELEPHONY ENGINE
    # =========================================================================
    slide14 = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide14)
    add_header(slide14, "Role Profile: Rokkam Madhava Rao — Logistics & Telephony", slide_num=14)

    # Left: Identity Card
    id_card14 = slide14.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(3.8), Inches(5.3))
    id_card14.fill.solid()
    id_card14.fill.fore_color.rgb = DARK_GREEN
    id_card14.line.color.rgb = PRIMARY_GREEN

    tb_id14 = slide14.shapes.add_textbox(Inches(1.05), Inches(1.75), Inches(3.3), Inches(4.8))
    tf_id14 = tb_id14.text_frame
    tf_id14.word_wrap = True
    tf_id14.margin_left = tf_id14.margin_top = tf_id14.margin_right = tf_id14.margin_bottom = 0

    p_id14_1 = tf_id14.paragraphs[0]
    p_id14_1.text = "Rokkam Madhava Rao"
    p_id14_1.font.size = Pt(17)
    p_id14_1.font.bold = True
    p_id14_1.font.color.rgb = WHITE
    p_id14_1.font.name = "Segoe UI"

    p_id14_2 = tf_id14.add_paragraph()
    p_id14_2.text = "LOGISTICS & TELEPHONY ENGINE"
    p_id14_2.font.size = Pt(11)
    p_id14_2.font.bold = True
    p_id14_2.font.color.rgb = ACCENT_AMBER
    p_id14_2.font.name = "Segoe UI"
    p_id14_2.space_after = Pt(10)

    p_id14_3 = tf_id14.add_paragraph()
    p_id14_3.text = "Official Credentials:\n• Enrollment No: AJU/230194\n• Class Roll No: Btcs/017\n• Technology: Telephony Simulation, Web Audio\n\nDual Engineering Verticals:\n1. Cold-Chain Logistics Engine\n• Vehicle assignment models\n• Hub aggregation workflows\n• Delivery state tracking\n2. Telephony & Assisted Access\n• 1800-890-KISAN simulator\n• DTMF keypress processing\n• Speech & SMS confirmation"
    p_id14_3.font.size = Pt(10)
    p_id14_3.font.color.rgb = RGBColor(220, 245, 230)
    p_id14_3.font.name = "Segoe UI"

    # Right: Responsibilities
    resp_card14 = slide14.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.9), Inches(1.5), Inches(7.633), Inches(5.3))
    resp_card14.fill.solid()
    resp_card14.fill.fore_color.rgb = WHITE
    resp_card14.line.color.rgb = BORDER_COLOR

    tb_r14 = slide14.shapes.add_textbox(Inches(5.2), Inches(1.75), Inches(7.0), Inches(4.8))
    tf_r14 = tb_r14.text_frame
    tf_r14.word_wrap = True
    tf_r14.margin_left = tf_r14.margin_top = tf_r14.margin_right = tf_r14.margin_bottom = 0

    p_rh14 = tf_r14.paragraphs[0]
    p_rh14.text = "LOGISTICS COORDINATION & TELEPHONY ENGINE"
    p_rh14.font.size = Pt(12)
    p_rh14.font.bold = True
    p_rh14.font.color.rgb = DARK_GREEN
    p_rh14.font.name = "Segoe UI"
    p_rh14.space_after = Pt(6)

    duties_log = [
        "1. Logistics Delivery Assignment Workflow", "Designed the DeliveryAssignment model linking confirmed orders with delivery drivers, vehicle identifiers (Tata Ace Cold Van JH-01-EF-4921), and rural aggregation hubs.",
        "2. Order Fulfillment State Machine", "Structured step-by-step delivery milestone tracking: Pending ➔ Packed ➔ Picked Up ➔ In Transit ➔ Delivered with driver notes.",
        "3. Interactive IVR Telephony Engine", "Built the full interactive logic for the 1800-890-KISAN helpline simulator: authentic DTMF audio tones, automated call menus, and multi-language voice prompts.",
        "4. Voice Harvest Intake & SMS Alerting", "Implemented simulated voice harvest intake: when a farmer registers 500 kg tomatoes via IVR, an automated confirmation SMS is generated in the database.",
        "5. Assisted-Digital Facilitator Integration", "Connected telephony logs with Village Digital Facilitator (VDF) workflows, ensuring local agents receive immediate alerts for harvest pickups."
    ]

    for i in range(0, len(duties_log), 2):
        pd = tf_r14.add_paragraph()
        pd.text = duties_log[i]
        pd.font.size = Pt(11)
        pd.font.bold = True
        pd.font.color.rgb = NAVY
        pd.font.name = "Segoe UI"

        pb = tf_r14.add_paragraph()
        pb.text = duties_log[i+1]
        pb.font.size = Pt(9.5)
        pb.font.color.rgb = TEXT_MUTED
        pb.font.name = "Segoe UI"
        pb.space_after = Pt(4)

    # Dual Pipeline Box
    flow_box14 = slide14.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.2), Inches(5.85), Inches(7.0), Inches(0.75))
    flow_box14.fill.solid()
    flow_box14.fill.fore_color.rgb = LIGHT_GREEN
    flow_box14.line.color.rgb = PRIMARY_GREEN
    tf_fb14 = flow_box14.text_frame
    tf_fb14.margin_top = Inches(0.12)
    p_fb14 = tf_fb14.paragraphs[0]
    p_fb14.text = "Fulfillment Pipeline: Order Created  ➔  Hub Logistics  ➔  Vehicle Allocated  ➔  Transit  ➔  Doorstep Delivery"
    p_fb14.font.size = Pt(10)
    p_fb14.font.bold = True
    p_fb14.font.color.rgb = DARK_GREEN
    p_fb14.alignment = PP_ALIGN.CENTER
    p_fb14.font.name = "Segoe UI"

    add_notes(slide14, """Respected evaluators, I am Rokkam Madhava Rao, leading the Logistics and Telephony Engine.
My responsibilities span two critical physical and communications layers of AgriConnect.
On the logistics front, agricultural produce is highly perishable. I designed the DeliveryAssignment workflow that matches confirmed orders with specific vehicles and drivers, managing pickup from rural aggregation centers like the Angara Hub through to doorstep delivery.
On the telephony front, I built the Assisted-Digital Telephony simulator for 1800-890-KISAN. This system generates dual-tone DTMF sounds using the Web Audio API, handles simulated keypad options like checking mandi prices or requesting a harvest pickup, and triggers automated transactional SMS logs.
This bridge ensures that even farmers without smartphones remain actively connected to our digital supply chain.""")

    # =========================================================================
    # SLIDE 15: ROLE: ML ANALYTICS & QA
    # =========================================================================
    slide15 = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide15)
    add_header(slide15, "Role Profile: N. Rohit Kumar — ML Analytics & QA Testing", slide_num=15)

    # Left: Identity Card
    id_card15 = slide15.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(3.8), Inches(5.3))
    id_card15.fill.solid()
    id_card15.fill.fore_color.rgb = DARK_GREEN
    id_card15.line.color.rgb = PRIMARY_GREEN

    tb_id15 = slide15.shapes.add_textbox(Inches(1.05), Inches(1.75), Inches(3.3), Inches(4.8))
    tf_id15 = tb_id15.text_frame
    tf_id15.word_wrap = True
    tf_id15.margin_left = tf_id15.margin_top = tf_id15.margin_right = tf_id15.margin_bottom = 0

    p_id15_1 = tf_id15.paragraphs[0]
    p_id15_1.text = "N. Rohit Kumar"
    p_id15_1.font.size = Pt(18)
    p_id15_1.font.bold = True
    p_id15_1.font.color.rgb = WHITE
    p_id15_1.font.name = "Segoe UI"

    p_id15_2 = tf_id15.add_paragraph()
    p_id15_2.text = "ML ANALYTICS & QA TESTING"
    p_id15_2.font.size = Pt(11)
    p_id15_2.font.bold = True
    p_id15_2.font.color.rgb = ACCENT_AMBER
    p_id15_2.font.name = "Segoe UI"
    p_id15_2.space_after = Pt(10)

    p_id15_3 = tf_id15.add_paragraph()
    p_id15_3.text = "Official Credentials:\n• Enrollment No: AJU/230686\n• Class Roll No: Btcs/051\n• Domain: Machine Learning & Software QA\n\nDual Engineering Verticals:\n1. Supply Chain ML & Analytics\n• 7-day demand forecasting model\n• Price elasticity calculation\n• Mandi pricing corridors\n2. Quality Assurance & Testing\n• End-to-end automated suite\n• Persona switching validation\n• Regression & boundary tests"
    p_id15_3.font.size = Pt(10)
    p_id15_3.font.color.rgb = RGBColor(220, 245, 230)
    p_id15_3.font.name = "Segoe UI"

    # Right: Responsibilities
    resp_card15 = slide15.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.9), Inches(1.5), Inches(7.633), Inches(5.3))
    resp_card15.fill.solid()
    resp_card15.fill.fore_color.rgb = WHITE
    resp_card15.line.color.rgb = BORDER_COLOR

    tb_r15 = slide15.shapes.add_textbox(Inches(5.2), Inches(1.75), Inches(7.0), Inches(4.8))
    tf_r15 = tb_r15.text_frame
    tf_r15.word_wrap = True
    tf_r15.margin_left = tf_r15.margin_top = tf_r15.margin_right = tf_r15.margin_bottom = 0

    p_rh15 = tf_r15.paragraphs[0]
    p_rh15.text = "ML ANALYTICS & QUALITY ASSURANCE SPECIFICATION"
    p_rh15.font.size = Pt(12)
    p_rh15.font.bold = True
    p_rh15.font.color.rgb = DARK_GREEN
    p_rh15.font.name = "Segoe UI"
    p_rh15.space_after = Pt(6)

    duties_qa = [
        "1. 7-Day Regression Demand Forecasting Model", "Developed Python predictive model (demand_predictor.py) incorporating baseline urban consumption, price elasticity coefficients (-1.25 for perishables), and weekend surge weights.",
        "2. Optimal Price Corridor Algorithm", "Created recommendation logic calculating recommended farmer min (Mandi benchmark * 1.15) to max (Retail benchmark * 0.78) for fair price discovery.",
        "3. Automated Full-System Test Suite", "Authored comprehensive test_full_suite.py utilizing Django Test Client, covering 186+ assertions across endpoints, APIs, and business workflows.",
        "4. Role-Based Persona Testing", "Simulated and validated all 5 user personas (vdf_sunil, farmer_ramesh, consumer_priya, delivery_rajesh, admin) ensuring session security and view isolation.",
        "5. Regression & Defect Verification", "Conducted continuous regression cycles, testing cart checkout, dispute filing, and IVR DTMF responses to guarantee zero regressions."
    ]

    for i in range(0, len(duties_qa), 2):
        pd = tf_r15.add_paragraph()
        pd.text = duties_qa[i]
        pd.font.size = Pt(11)
        pd.font.bold = True
        pd.font.color.rgb = NAVY
        pd.font.name = "Segoe UI"

        pb = tf_r15.add_paragraph()
        pb.text = duties_qa[i+1]
        pb.font.size = Pt(9.5)
        pb.font.color.rgb = TEXT_MUTED
        pb.font.name = "Segoe UI"
        pb.space_after = Pt(4)

    # QA & Analytics Box
    flow_box15 = slide15.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.2), Inches(5.85), Inches(7.0), Inches(0.75))
    flow_box15.fill.solid()
    flow_box15.fill.fore_color.rgb = LIGHT_GREEN
    flow_box15.line.color.rgb = PRIMARY_GREEN
    tf_fb15 = flow_box15.text_frame
    tf_fb15.margin_top = Inches(0.12)
    p_fb15 = tf_fb15.paragraphs[0]
    p_fb15.text = "Dual Lifecycle: [Data ➔ Predict ➔ Advisory]  |  [Test Case ➔ Execution ➔ Defect Fix ➔ Verified Release]"
    p_fb15.font.size = Pt(10)
    p_fb15.font.bold = True
    p_fb15.font.color.rgb = DARK_GREEN
    p_fb15.alignment = PP_ALIGN.CENTER
    p_fb15.font.name = "Segoe UI"

    add_notes(slide15, """Respected evaluators, I am N. Rohit Kumar, heading ML Analytics and QA Testing.
My work directly impacts both the intelligence and the stability of AgriConnect.
In the analytics domain, I developed our 7-day demand forecasting engine in Python. This model calculates price elasticity for crops like tomatoes and onions, projects consumption volumes across the Ranchi urban cluster, and determines the optimal price corridor between wholesale mandi and supermarket retail prices.
In the QA domain, I engineered our automated test suite in test_full_suite.py. Using the Django Test Client, I automated end-to-end verification across every public view, role dashboard, checkout flow, and IVR API endpoint.
I validated that role switches between our 5 personas remain isolated and secure, ensuring our release is robust, bug-free, and thoroughly validated.""")

    # =========================================================================
    # SLIDE 16: TECHNOLOGY STACK
    # =========================================================================
    slide16 = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide16)
    add_header(slide16, "Comprehensive Technology Stack", slide_num=16)

    tech_cards = [
        ("Frontend Technologies", "• HTML5 & Semantic Web\n• Modern Vanilla CSS3\n• ES6+ JavaScript\n• Web Audio API (DTMF)\n• Web Speech Synthesis API", DARK_GREEN, LIGHT_GREEN),
        ("Backend Framework", "• Python 3.12 Runtime\n• Django 5.x Web Framework\n• Django REST Framework\n• Gunicorn 22.0 WSGI Server\n• WhiteNoise 6.5 Static Engine", NAVY, WHITE),
        ("Database Systems", "• PostgreSQL 16 (Production)\n• SQLite (Local Testing)\n• Django ORM\n• dj-database-url Engine\n• psycopg2-binary Driver", DARK_GREEN, LIGHT_GREEN),
        ("Logistics & Telephony", "• Dual-Tone Multi-Frequency (DTMF)\n• Web Audio Oscillator Matrix\n• Simulated SMS Database Log\n• Cold Mini-Van Dispatch Models\n• Cloud Telephony Ready (Twilio/Exotel)", NAVY, WHITE),
        ("Analytics & ML", "• Python Mathematical Modeling\n• Demand Regression Predictor\n• APMC Mandi Benchmark Engine\n• Price Elasticity Modeling\n• Chart.js Visual Reporting", DARK_GREEN, LIGHT_GREEN),
        ("DevOps & QA Testing", "• Render Cloud PaaS\n• GitHub CI/CD Repository\n• Python-Dotenv Security\n• Django Test Client Suite\n• End-to-End Test Suite (186+ checks)", NAVY, WHITE)
    ]

    t_w = Inches(3.7)
    t_h = Inches(2.4)
    for i, (t_title, t_desc, fcol, bgcol) in enumerate(tech_cards):
        col = i % 3
        row = i // 3
        tx = Inches(0.8) + col * Inches(3.7 + 0.31)
        ty = Inches(1.5) + row * Inches(2.4 + 0.35)

        card = slide16.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, tx, ty, t_w, t_h)
        card.fill.solid()
        card.fill.fore_color.rgb = bgcol
        card.line.color.rgb = BORDER_COLOR

        tb = slide16.shapes.add_textbox(tx + Inches(0.2), ty + Inches(0.2), t_w - Inches(0.4), t_h - Inches(0.4))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p1 = tf.paragraphs[0]
        p1.text = t_title
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.font.color.rgb = fcol
        p1.font.name = "Segoe UI"
        p1.space_after = Pt(6)

        p2 = tf.add_paragraph()
        p2.text = t_desc
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = NAVY if fcol != NAVY else DARK_GREEN
        p2.font.name = "Segoe UI"

    add_notes(slide16, """Slide 16 details the complete technology stack powering AgriConnect.
On the frontend, we chose modern vanilla CSS and JavaScript over bulky frontend frameworks to keep the payload ultralight for rural 2G/3G connectivity, taking advantage of standard browser APIs like Web Audio for DTMF simulation and Web Speech for voice generation.
On the backend, Python 3.12 and Django 5 provide an industrial-strength foundation with Django REST Framework for JSON APIs and Gunicorn for WSGI serving.
For data persistence, we run PostgreSQL 16 on Render Cloud, utilizing psycopg2 and dj-database-url.
Our logistics and telephony layer simulates realistic telecom DTMF frequencies and handles harvest dispatch.
Our analytics engine uses pure Python math modeling for demand curves and mandi benchmarking.
And our deployment pipeline uses GitHub connected to Render Cloud for seamless continuous integration.""")

    # =========================================================================
    # SLIDE 17: SECURITY & DEPLOYMENT
    # =========================================================================
    slide17 = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide17)
    add_header(slide17, "Security, Cloud Deployment & Production Operations", slide_num=17)

    # 3 Strategic Cards
    sec_cards = [
        ("Application Security & Governance", [
            "Role-Based Access Control (RBAC): Strict custom view decorators prevent unauthorized access between personas.",
            "CSRF & XSS Protection: Full Django CSRF tokens applied to all POST endpoints, forms, and IVR actions.",
            "Environment Variable Isolation: Secret keys, database credentials, and debug toggles stored in isolated .env.",
            "Password Hashing & Session Security: PBKDF2 with SHA-256 password storage and secure session cookies."
        ]),
        ("Relational Database & Migrations", [
            "PostgreSQL Managed Database: Deployed on Render Cloud with automated connection pooling via dj-database-url.",
            "Atomic Transaction Blocks: Critical checkout operations use select_for_update to avoid stock overselling.",
            "Normalized Data Schemas: Foreign key constraints and cascade rules safeguard data integrity.",
            "Comprehensive Migrations: Automated schema versioning ensuring reproducible database states."
        ]),
        ("Cloud Deployment Architecture", [
            "Render Cloud PaaS: Auto-deploys on every git push to the main branch via render.yaml configuration.",
            "Gunicorn WSGI & WhiteNoise: Multithreaded request handling paired with efficient static asset caching.",
            "Live Production Access: Fully accessible at https://agriconnect-q20v.onrender.com.",
            "Continuous Health Monitoring: Automated error logging and status tracking on cloud infrastructure."
        ])
    ]

    sc_w = Inches(3.7)
    sc_h = Inches(5.1)
    for i, (title, points) in enumerate(sec_cards):
        sc_x = Inches(0.8) + i * Inches(3.7 + 0.31)
        card = slide17.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, sc_x, Inches(1.5), sc_w, sc_h)
        card.fill.solid()
        card.fill.fore_color.rgb = WHITE
        card.line.color.rgb = BORDER_COLOR

        tb = slide17.shapes.add_textbox(sc_x + Inches(0.2), Inches(1.75), sc_w - Inches(0.4), sc_h - Inches(0.5))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p1 = tf.paragraphs[0]
        p1.text = title
        p1.font.size = Pt(13)
        p1.font.bold = True
        p1.font.color.rgb = DARK_GREEN
        p1.font.name = "Segoe UI"
        p1.space_after = Pt(10)

        for pt in points:
            pp = tf.add_paragraph()
            pp.text = f"• {pt}"
            pp.font.size = Pt(9.5)
            pp.font.color.rgb = NAVY
            pp.font.name = "Segoe UI"
            pp.space_after = Pt(8)

    add_notes(slide17, """Slide 17 covers Security, Database Management, and Production Deployment.
Security is built into AgriConnect from the ground up:
Every endpoint is protected against Cross-Site Request Forgery and unauthorized role access. Credentials and secret keys are kept completely out of source code using python-dotenv.
Our database is a production PostgreSQL instance running on Render Cloud, utilizing atomic database transactions to ensure that stock reservations and order confirmations are always consistent.
Finally, our deployment runs on Render Cloud using Gunicorn and WhiteNoise, triggered automatically through Git CI/CD.
Evaluators can access and test the live application right now at https://agriconnect-q20v.onrender.com.""")

    # =========================================================================
    # SLIDE 18: CHALLENGES & SOLUTIONS
    # =========================================================================
    slide18 = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide18)
    add_header(slide18, "Engineering Challenges & Implemented Solutions", slide_num=18)

    # Table layout: 6 rows
    ch_rows = [
        ("Multi-Role State Complexity", "Five distinct user personas sharing the same web application caused navigation and permission confusion.", "Engineered role-isolated dashboards, custom Django view decorators, and a quick-switch persona bar for demonstrations."),
        ("Rural Digital Exclusion Barrier", "Illiterate smallholders and farmers with 2G keypad feature phones cannot interact with smartphone apps.", "Implemented the Village Digital Facilitator (VDF) assisted workflow paired with an interactive 1800 IVR telephony dialer."),
        ("Post-Harvest Perishability Waste", "Uncoordinated farmer pickups result in vegetable spoilage and heavy financial losses.", "Built the DeliveryAssignment logistics engine linking rural aggregation hubs to cold-chain vans with milestone updates."),
        ("Price Discovery & Middleman Markups", "Farmers have no benchmark data to negotiate fair prices against exploitative mandi buyers.", "Integrated daily APMC mandi rates and built the price corridor recommender calculating optimal min-max price ranges."),
        ("Cross-Module Integration Integrity", "Simultaneous development across frontend, backend, telephony, and analytics risked integration failures.", "Authored an automated full-system regression suite (test_full_suite.py) validating 186+ endpoint assertions continuously."),
        ("Production Cloud Configuration", "Deploying PostgreSQL and Django static assets seamlessly on cloud hosting without downtime.", "Configured render.yaml, WhiteNoise compression, dj-database-url, and Gunicorn for automated deployment.")
    ]

    # Two column header
    h_col1 = slide18.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(5.7), Inches(0.45))
    h_col1.fill.solid()
    h_col1.fill.fore_color.rgb = DARK_GREEN
    h_col1.line.fill.background()
    tf_h1 = h_col1.text_frame
    p_h1 = tf_h1.paragraphs[0]
    p_h1.text = "TECHNICAL CHALLENGE"
    p_h1.font.size = Pt(11)
    p_h1.font.bold = True
    p_h1.font.color.rgb = WHITE
    p_h1.alignment = PP_ALIGN.CENTER
    p_h1.font.name = "Segoe UI"

    h_col2 = slide18.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.5), Inches(5.7), Inches(0.45))
    h_col2.fill.solid()
    h_col2.fill.fore_color.rgb = PRIMARY_GREEN
    h_col2.line.fill.background()
    tf_h2 = h_col2.text_frame
    p_h2 = tf_h2.paragraphs[0]
    p_h2.text = "IMPLEMENTED ENGINEERING SOLUTION"
    p_h2.font.size = Pt(11)
    p_h2.font.bold = True
    p_h2.font.color.rgb = WHITE
    p_h2.alignment = PP_ALIGN.CENTER
    p_h2.font.name = "Segoe UI"

    row_y = 2.05
    for ch_title, ch_desc, sol_desc in ch_rows:
        # Challenge Card
        c_box = slide18.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(row_y), Inches(5.7), Inches(0.72))
        c_box.fill.solid()
        c_box.fill.fore_color.rgb = WHITE
        c_box.line.color.rgb = RGBColor(230, 200, 200)
        tf_c = c_box.text_frame
        tf_c.word_wrap = True
        tf_c.margin_left = Inches(0.15)
        tf_c.margin_top = Inches(0.08)
        p_c1 = tf_c.paragraphs[0]
        p_c1.text = ch_title
        p_c1.font.size = Pt(10)
        p_c1.font.bold = True
        p_c1.font.color.rgb = RED_ACCENT
        p_c1.font.name = "Segoe UI"
        p_c2 = tf_c.add_paragraph()
        p_c2.text = ch_desc
        p_c2.font.size = Pt(8.5)
        p_c2.font.color.rgb = NAVY
        p_c2.font.name = "Segoe UI"

        # Solution Card
        s_box = slide18.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(row_y), Inches(5.7), Inches(0.72))
        s_box.fill.solid()
        s_box.fill.fore_color.rgb = WHITE
        s_box.line.color.rgb = BORDER_COLOR
        tf_s = s_box.text_frame
        tf_s.word_wrap = True
        tf_s.margin_left = Inches(0.15)
        tf_s.margin_top = Inches(0.08)
        p_s1 = tf_s.paragraphs[0]
        p_s1.text = "Engineered Resolution:"
        p_s1.font.size = Pt(10)
        p_s1.font.bold = True
        p_s1.font.color.rgb = PRIMARY_GREEN
        p_s1.font.name = "Segoe UI"
        p_s2 = tf_s.add_paragraph()
        p_s2.text = sol_desc
        p_s2.font.size = Pt(8.5)
        p_s2.font.color.rgb = NAVY
        p_s2.font.name = "Segoe UI"

        row_y += 0.8

    add_notes(slide18, """Slide 18 outlines the key engineering challenges we tackled and solved.
The first challenge was managing five distinct user roles without cross-polluting permissions. We solved this with custom view decorators and dedicated dashboards.
The second challenge was rural digital literacy. We overcame this with our Village Facilitator model and our 1800 IVR simulator.
The third challenge was reducing post-harvest waste, solved by our aggregation hub dispatch engine.
The fourth challenge was price discovery, resolved by our APMC mandi benchmark analyzer.
The fifth challenge was maintaining integration quality across four team members' work, solved by our automated 186-assertion test suite.
And the sixth was cloud deployment, solved with Render, WhiteNoise, and automated migrations.
These solutions demonstrate our practical problem-solving capability.""")

    # =========================================================================
    # SLIDE 19: FUTURE SCOPE
    # =========================================================================
    slide19 = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide19)
    add_header(slide19, "Future Scope & Strategic Technology Roadmap", slide_num=19)

    # 5 Roadmap Horizons
    roadmap_horizons = [
        ("PHASE 1 (CURRENT)", "Production Foundation", "• Fully operational web platform\n• Multi-role dashboard engine\n• Web Audio IVR phone simulator\n• Mandi benchmark comparator\n• Automated test suite & Render PaaS", DARK_GREEN, LIGHT_GREEN),
        ("PHASE 2 (SHORT TERM)", "Mobile & Voice Gateway", "• Progressive Web App (PWA)\n• Real-world Cloud Telephony APIs (Twilio/Exotel)\n• Multi-lingual voice intake (Hindi, Bengali, Santhali)\n• UPI & escrows payment integration\n• Automated push & SMS gateways", PRIMARY_GREEN, WHITE),
        ("PHASE 3 (MEDIUM TERM)", "AI/ML Demand & Pricing", "• Multi-variate LSTM demand forecasting\n• Satellite weather risk assessment\n• Dynamic dynamic price matching\n• Farmer crop rotation advisories\n• Yield prediction using historical records", NAVY, LIGHT_GREEN),
        ("PHASE 4 (LONG TERM)", "IoT & Smart Cold-Chain", "• Real-time GPS vehicle tracking\n• IoT temperature & humidity cold-van sensors\n• QR-code batch traceability\n• Automated quality dispute arbitration\n• Smart warehouse inventory robotics", DARK_GREEN, WHITE),
        ("PHASE 5 (VISION)", "Pan-India Agri Grid", "• Direct integration with e-NAM national mandis\n• Cooperative Farmer Producer Org (FPO) tooling\n• Micro-lending & crop insurance credit rating\n• Carbon credit tracking for organic farming\n• Unified national food security analytics", PRIMARY_GREEN, LIGHT_GREEN)
    ]

    h_w = Inches(2.18)
    h_gap = Inches(0.2)
    start_hx = Inches(0.8)

    for i, (h_tag, h_title, h_body, bcol, fcol) in enumerate(roadmap_horizons):
        hx = start_hx + i * (h_w + h_gap)
        card = slide19.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, hx, Inches(1.5), h_w, Inches(5.3))
        card.fill.solid()
        card.fill.fore_color.rgb = fcol
        card.line.color.rgb = bcol
        card.line.width = Pt(1.5)

        # Top Tag
        tag = slide19.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, hx + Inches(0.1), Inches(1.65), h_w - Inches(0.2), Inches(0.4))
        tag.fill.solid()
        tag.fill.fore_color.rgb = bcol
        tag.line.fill.background()
        tf_t = tag.text_frame
        p_t = tf_t.paragraphs[0]
        p_t.text = h_tag
        p_t.font.size = Pt(8.5)
        p_t.font.bold = True
        p_t.font.color.rgb = WHITE
        p_t.alignment = PP_ALIGN.CENTER
        p_t.font.name = "Segoe UI"

        tb = slide19.shapes.add_textbox(hx + Inches(0.12), Inches(2.2), h_w - Inches(0.24), Inches(4.4))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p1 = tf.paragraphs[0]
        p1.text = h_title
        p1.font.size = Pt(11)
        p1.font.bold = True
        p1.font.color.rgb = NAVY
        p1.font.name = "Segoe UI"
        p1.space_after = Pt(8)

        p2 = tf.add_paragraph()
        p2.text = h_body
        p2.font.size = Pt(9)
        p2.font.color.rgb = TEXT_MUTED
        p2.font.name = "Segoe UI"

    add_notes(slide19, """Slide 19 outlines our Future Scope and Strategic Roadmap.
We clearly distinguish between what is already implemented today in Phase 1 and our planned future enhancements:
Phase 1 is complete: the live web platform, role dashboards, IVR simulation, price transparency engine, and cloud deployment.
In Phase 2, we will develop a lightweight Progressive Web App and link actual PSTN cloud telephony APIs such as Exotel or Twilio for live phone calls.
In Phase 3, we plan to enhance our analytics into deep-learning LSTM models for regional demand forecasting and yield prediction.
In Phase 4, we will introduce IoT temperature sensors into cold-storage vans with QR-code farm-to-table batch traceability.
And in Phase 5, our vision is to integrate with national initiatives like e-NAM, building a unified digital backbone for Indian agriculture.""")

    # =========================================================================
    # SLIDE 20: CONCLUSION & Q&A
    # =========================================================================
    slide20 = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide20)

    # Top accent bar
    top_bar20 = slide20.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.15))
    top_bar20.fill.solid()
    top_bar20.fill.fore_color.rgb = PRIMARY_GREEN
    top_bar20.line.fill.background()

    # Left Summary Card
    sum_card = slide20.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.8), Inches(6.8), Inches(5.8))
    sum_card.fill.solid()
    sum_card.fill.fore_color.rgb = WHITE
    sum_card.line.color.rgb = BORDER_COLOR

    tb_sc = slide20.shapes.add_textbox(Inches(1.1), Inches(1.1), Inches(6.2), Inches(5.2))
    tf_sc = tb_sc.text_frame
    tf_sc.word_wrap = True
    tf_sc.margin_left = tf_sc.margin_top = tf_sc.margin_right = tf_sc.margin_bottom = 0

    p_sch = tf_sc.paragraphs[0]
    p_sch.text = "Key Project Achievements"
    p_sch.font.size = Pt(20)
    p_sch.font.bold = True
    p_sch.font.color.rgb = DARK_GREEN
    p_sch.font.name = "Segoe UI"
    p_sch.space_after = Pt(12)

    concl_points = [
        "1. Centralized Transparent Marketplace: Successfully designed and deployed an end-to-end platform connecting farmers directly to consumers, cutting out exploitative middlemen.",
        "2. Innovative Assisted-Digital Access: Bridged the rural digital divide using Village Facilitators (VDF) and an interactive 1800-toll-free IVR simulator for keypad-phone farmers.",
        "3. Coordinated Logistics & Cold-Chain: Implemented structured aggregation hub dispatch and delivery milestone tracking to minimize post-harvest food waste.",
        "4. Data-Driven Pricing & Demand Analytics: Created mandi benchmark comparators and 7-day regression demand models for fair farmgate remuneration.",
        "5. Production-Ready Deployment: Engineered with Django 5, PostgreSQL, automated test suites (186+ assertions), live at https://agriconnect-q20v.onrender.com."
    ]

    for pt in concl_points:
        p = tf_sc.add_paragraph()
        p.text = pt
        p.font.size = Pt(10)
        p.font.color.rgb = NAVY
        p.font.name = "Segoe UI"
        p.space_after = Pt(8)

    # Right Hero Card: Final Message & Thank You
    thank_card = slide20.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.9), Inches(0.8), Inches(4.633), Inches(5.8))
    thank_card.fill.solid()
    thank_card.fill.fore_color.rgb = DARK_GREEN
    thank_card.line.color.rgb = PRIMARY_GREEN

    tb_tc = slide20.shapes.add_textbox(Inches(8.2), Inches(1.3), Inches(4.033), Inches(4.8))
    tf_tc = tb_tc.text_frame
    tf_tc.word_wrap = True
    tf_tc.margin_left = tf_tc.margin_top = tf_tc.margin_right = tf_tc.margin_bottom = 0

    p_tc1 = tf_tc.paragraphs[0]
    p_tc1.text = "“Connecting Agriculture,\nTechnology and People.”"
    p_tc1.font.size = Pt(19)
    p_tc1.font.bold = True
    p_tc1.font.color.rgb = ACCENT_AMBER
    p_tc1.alignment = PP_ALIGN.CENTER
    p_tc1.font.name = "Segoe UI"
    p_tc1.space_after = Pt(20)

    p_tc2 = tf_tc.add_paragraph()
    p_tc2.text = "THANK YOU!"
    p_tc2.font.size = Pt(32)
    p_tc2.font.bold = True
    p_tc2.font.color.rgb = WHITE
    p_tc2.alignment = PP_ALIGN.CENTER
    p_tc2.font.name = "Segoe UI"
    p_tc2.space_after = Pt(14)

    p_tc3 = tf_tc.add_paragraph()
    p_tc3.text = "Questions & Answers"
    p_tc3.font.size = Pt(16)
    p_tc3.font.bold = True
    p_tc3.font.color.rgb = RGBColor(220, 245, 230)
    p_tc3.alignment = PP_ALIGN.CENTER
    p_tc3.font.name = "Segoe UI"
    p_tc3.space_after = Pt(20)

    p_tc4 = tf_tc.add_paragraph()
    p_tc4.text = "Project Team:\nMd Aamir Alam (Project Leader & UI)\nMd Shail (Backend & Database)\nRokkam Madhava Rao (Logistics & IVR)\nN. Rohit Kumar (ML Analytics & QA)\n\nLive Demo:\nhttps://agriconnect-q20v.onrender.com"
    p_tc4.font.size = Pt(9.5)
    p_tc4.font.color.rgb = RGBColor(200, 230, 210)
    p_tc4.alignment = PP_ALIGN.CENTER
    p_tc4.font.name = "Segoe UI"

    # Slide 20 Footer
    foot20 = slide20.shapes.add_textbox(Inches(0.8), Inches(7.05), Inches(11.733), Inches(0.3))
    tf_f20 = foot20.text_frame
    p_f20 = tf_f20.paragraphs[0]
    p_f20.text = "AgriConnect  •  Project Progress Seminar – II  •  Department of Computer Science & Engineering"
    p_f20.font.size = Pt(9)
    p_f20.font.color.rgb = TEXT_MUTED
    p_f20.alignment = PP_ALIGN.CENTER
    p_f20.font.name = "Segoe UI"

    add_notes(slide20, """To conclude our Project Progress Seminar – II presentation:
AgriConnect demonstrates that digital technology can be designed to include everyone—from illiterate smallholders to urban consumers.
By connecting transparent pricing, village-level assisted access, telephony simulation, coordinated cold-chain logistics, and predictive analytics into a unified Django-PostgreSQL architecture, we have built a functional platform that directly addresses India's agricultural supply chain crisis.
Our live platform is active and ready on Render, backed by extensive automated test coverage.
We sincerely thank our guide, faculty members, and evaluators for their time and guidance.
We are now open to your questions and feedback.""")

    output_path = os.path.abspath("AgriConnect_Project_Progress_Seminar_II.pptx")
    prs.save(output_path)
    print(f"Presentation successfully built and saved to: {output_path}")

if __name__ == "__main__":
    build_presentation()
