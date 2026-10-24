"""
Run this script to generate:
  centre_of_mass_FYBSc.pptx

Requirements:
  pip install python-pptx

Then run:
  python3 generate_ppt.py
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

# ── Colours ──────────────────────────────────────────────
BG_DARK   = RGBColor(0x1A, 0x1A, 0x2E)   # dark navy
ACCENT    = RGBColor(0xE9, 0x45, 0x60)   # red-pink
BLUE      = RGBColor(0x4C, 0xC9, 0xF0)   # light blue
GREEN     = RGBColor(0x48, 0xC7, 0x8E)   # mint green
ORANGE    = RGBColor(0xFF, 0xA5, 0x00)   # orange
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GREY= RGBColor(0xC0, 0xC8, 0xD8)

def set_bg(slide, color):
    """Fill slide background with a solid colour."""
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_textbox(slide, text, left, top, width, height,
                font_size=18, bold=False, color=WHITE,
                align=PP_ALIGN.LEFT, italic=False):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return txBox

def add_colored_box(slide, left, top, width, height, fill_color, line_color=None):
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        left, top, width, height
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(1.2)
    else:
        shape.line.fill.background()
    return shape

def add_bullet_box(slide, title, bullets, left, top, width, height,
                   title_color=GREEN, accent_color=GREEN):
    # Card background
    card = add_colored_box(slide, left, top, width, height,
                           RGBColor(0x0D, 0x1F, 0x0D), accent_color)

    # Title
    add_textbox(slide, title,
                left + Inches(0.15), top + Inches(0.1),
                width - Inches(0.3), Inches(0.45),
                font_size=14, bold=True, color=title_color)

    # Bullets
    txBox = slide.shapes.add_textbox(
        left + Inches(0.15), top + Inches(0.55),
        width - Inches(0.3), height - Inches(0.65)
    )
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, b in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text = "▸  " + b
        run.font.size = Pt(11)
        run.font.color.rgb = LIGHT_GREY

# ── Create presentation ───────────────────────────────────
prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)

blank_layout = prs.slide_layouts[6]  # completely blank

# ═══════════════════════════════════════════════════════════
# SLIDE 1 – Title
# ═══════════════════════════════════════════════════════════
slide1 = prs.slides.add_slide(blank_layout)
set_bg(slide1, BG_DARK)

# Top accent bar
add_colored_box(slide1, 0, 0, prs.slide_width, Inches(0.08), ACCENT)

# MH Study brand
add_textbox(slide1, "📚  MH Study",
            Inches(0.4), Inches(0.15), Inches(3), Inches(0.5),
            font_size=13, bold=True, color=ACCENT)

# Subject tag box
tag = add_colored_box(slide1,
                      prs.slide_width - Inches(3.2), Inches(0.1),
                      Inches(3), Inches(0.45),
                      RGBColor(0x3A, 0x0A, 0x18), ACCENT)

add_textbox(slide1, "FY BSc  ·  Physics",
            prs.slide_width - Inches(3.1), Inches(0.12),
            Inches(2.8), Inches(0.4),
            font_size=11, color=ACCENT, align=PP_ALIGN.CENTER)

# Main title
add_textbox(slide1, "CENTRE OF MASS",
            Inches(1), Inches(1.8), Inches(11.33), Inches(1.6),
            font_size=54, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# Red underline
add_colored_box(slide1, Inches(3.5), Inches(3.55), Inches(6.33), Inches(0.06), ACCENT)

# Definition
add_textbox(slide1,
            "The centre of mass is the unique point where the entire mass of a body\n"
            "or system of particles appears to be concentrated.\n"
            "It is the weighted average position of all mass in the system.",
            Inches(1.5), Inches(3.75), Inches(10.33), Inches(1.5),
            font_size=16, color=LIGHT_GREY, align=PP_ALIGN.CENTER)

# Chapter tag
add_textbox(slide1, "⚛️   FY BSc  ·  Physics  ·  Systems of Particles",
            Inches(3.5), Inches(5.5), Inches(6.33), Inches(0.5),
            font_size=12, color=ACCENT, align=PP_ALIGN.CENTER)

# Bottom bar
add_colored_box(slide1, 0, prs.slide_height - Inches(0.08),
                prs.slide_width, Inches(0.08), ACCENT)

# Slide number
add_textbox(slide1, "1 / 5",
            prs.slide_width - Inches(1), prs.slide_height - Inches(0.45),
            Inches(0.9), Inches(0.35),
            font_size=10, color=RGBColor(0x55, 0x55, 0x77), align=PP_ALIGN.RIGHT)


# ═══════════════════════════════════════════════════════════
# SLIDE 2 – Formula
# ═══════════════════════════════════════════════════════════
slide2 = prs.slides.add_slide(blank_layout)
set_bg(slide2, RGBColor(0x1A, 0x1A, 0x40))
add_colored_box(slide2, 0, 0, prs.slide_width, Inches(0.08), BLUE)

add_textbox(slide2, "📚  MH Study  ·  FY BSc Physics",
            Inches(0.4), Inches(0.15), Inches(5), Inches(0.4),
            font_size=11, color=RGBColor(0x66, 0x77, 0x88))

add_textbox(slide2, "🔢   Formula",
            Inches(0.5), Inches(0.7), Inches(10), Inches(0.7),
            font_size=28, bold=True, color=BLUE)

add_colored_box(slide2, Inches(0.5), Inches(1.45), Inches(12.33), Inches(0.04),
                RGBColor(0x2A, 0x3A, 0x55))

formulas = [
    ("For two particles",
     "x_cm  =  (m₁x₁ + m₂x₂) / (m₁ + m₂)",
     "where m₁, m₂ are masses and x₁, x₂ are their positions"),
    ("For many particles (general form)",
     "x_cm  =  Σ(mᵢ xᵢ) / Σmᵢ",
     "Summation over all particles in the system"),
    ("3D position vector",
     "r_cm  =  Σ(mᵢ rᵢ) / M",
     "M = total mass of the system"),
]

y = Inches(1.6)
for label, formula, note in formulas:
    # Card
    add_colored_box(slide2, Inches(0.5), y, Inches(12.33), Inches(1.35),
                    RGBColor(0x10, 0x18, 0x35), BLUE)
    add_textbox(slide2, label,
                Inches(0.75), y + Inches(0.1), Inches(11.8), Inches(0.35),
                font_size=10, color=RGBColor(0x88, 0x99, 0xAA))
    add_textbox(slide2, formula,
                Inches(0.75), y + Inches(0.4), Inches(11.8), Inches(0.5),
                font_size=18, bold=True, color=BLUE)
    add_textbox(slide2, note,
                Inches(0.75), y + Inches(0.88), Inches(11.8), Inches(0.35),
                font_size=10, color=RGBColor(0x88, 0x99, 0xAA), italic=True)
    y += Inches(1.55)

add_colored_box(slide2, 0, prs.slide_height - Inches(0.08),
                prs.slide_width, Inches(0.08), BLUE)
add_textbox(slide2, "2 / 5",
            prs.slide_width - Inches(1), prs.slide_height - Inches(0.45),
            Inches(0.9), Inches(0.35),
            font_size=10, color=RGBColor(0x55, 0x55, 0x77), align=PP_ALIGN.RIGHT)


# ═══════════════════════════════════════════════════════════
# SLIDE 3 – Important Points
# ═══════════════════════════════════════════════════════════
slide3 = prs.slides.add_slide(blank_layout)
set_bg(slide3, RGBColor(0x1A, 0x2E, 0x1A))
add_colored_box(slide3, 0, 0, prs.slide_width, Inches(0.08), GREEN)

add_textbox(slide3, "📚  MH Study  ·  FY BSc Physics",
            Inches(0.4), Inches(0.15), Inches(5), Inches(0.4),
            font_size=11, color=RGBColor(0x66, 0x77, 0x88))

add_textbox(slide3, "⚡   Important Points",
            Inches(0.5), Inches(0.7), Inches(10), Inches(0.7),
            font_size=28, bold=True, color=GREEN)

add_colored_box(slide3, Inches(0.5), Inches(1.45), Inches(12.33), Inches(0.04),
                RGBColor(0x2A, 0x45, 0x2A))

points = [
    "The centre of mass always lies closer to the heavier mass in a two-body system.",
    "For a uniform symmetric body (sphere, ring, rod), the CM coincides with the geometric centre.",
    "The centre of mass can lie outside the material of an object (e.g., a ring or hollow sphere).",
    "It does not depend on the choice of coordinate system — only on mass distribution.",
    "In the absence of external forces, the centre of mass moves with constant velocity (or stays at rest).",
]

y = Inches(1.6)
for i, pt in enumerate(points):
    add_colored_box(slide3, Inches(0.5), y, Inches(12.33), Inches(0.85),
                    RGBColor(0x0D, 0x20, 0x0D), GREEN)
    add_textbox(slide3, f"  {i+1}",
                Inches(0.6), y + Inches(0.15), Inches(0.45), Inches(0.55),
                font_size=14, bold=True, color=GREEN, align=PP_ALIGN.CENTER)
    add_textbox(slide3, pt,
                Inches(1.15), y + Inches(0.15), Inches(11.4), Inches(0.6),
                font_size=13, color=LIGHT_GREY)
    y += Inches(1.0)

add_colored_box(slide3, 0, prs.slide_height - Inches(0.08),
                prs.slide_width, Inches(0.08), GREEN)
add_textbox(slide3, "3 / 5",
            prs.slide_width - Inches(1), prs.slide_height - Inches(0.45),
            Inches(0.9), Inches(0.35),
            font_size=10, color=RGBColor(0x55, 0x55, 0x77), align=PP_ALIGN.RIGHT)


# ═══════════════════════════════════════════════════════════
# SLIDE 4 – Centre of Mass & Momentum
# ═══════════════════════════════════════════════════════════
slide4 = prs.slides.add_slide(blank_layout)
set_bg(slide4, RGBColor(0x2E, 0x1A, 0x0D))
add_colored_box(slide4, 0, 0, prs.slide_width, Inches(0.08), ORANGE)

add_textbox(slide4, "📚  MH Study  ·  FY BSc Physics",
            Inches(0.4), Inches(0.15), Inches(5), Inches(0.4),
            font_size=11, color=RGBColor(0x77, 0x66, 0x55))

add_textbox(slide4, "🚀   Centre of Mass & Momentum",
            Inches(0.5), Inches(0.7), Inches(12), Inches(0.7),
            font_size=26, bold=True, color=ORANGE)

add_colored_box(slide4, Inches(0.5), Inches(1.45), Inches(12.33), Inches(0.04),
                RGBColor(0x45, 0x30, 0x15))

# Left column
add_colored_box(slide4, Inches(0.5), Inches(1.6), Inches(5.9), Inches(4.8),
                RGBColor(0x22, 0x10, 0x02), ORANGE)
add_textbox(slide4, "Linear Momentum",
            Inches(0.7), Inches(1.7), Inches(5.5), Inches(0.5),
            font_size=16, bold=True, color=ORANGE)
add_colored_box(slide4, Inches(0.7), Inches(2.22), Inches(5.5), Inches(0.03),
                RGBColor(0x55, 0x35, 0x10))

left_pts = [
    "Total momentum = total mass × velocity of CM.",
    "If no external force acts, total momentum is conserved.",
    "CM continues in uniform motion (Newton's 1st Law).",
]
y = Inches(2.35)
for pt in left_pts:
    add_textbox(slide4, "▸  " + pt,
                Inches(0.75), y, Inches(5.4), Inches(0.65),
                font_size=12, color=LIGHT_GREY)
    y += Inches(0.72)

# Formula box left
add_colored_box(slide4, Inches(0.7), Inches(5.0), Inches(5.5), Inches(0.7),
                RGBColor(0x40, 0x25, 0x05), ORANGE)
add_textbox(slide4, "P_total  =  M · v_cm",
            Inches(0.7), Inches(5.05), Inches(5.5), Inches(0.6),
            font_size=18, bold=True, color=ORANGE, align=PP_ALIGN.CENTER)

# Right column
add_colored_box(slide4, Inches(6.93), Inches(1.6), Inches(5.9), Inches(4.8),
                RGBColor(0x22, 0x10, 0x02), ORANGE)
add_textbox(slide4, "Newton's 2nd Law for a System",
            Inches(7.1), Inches(1.7), Inches(5.5), Inches(0.5),
            font_size=16, bold=True, color=ORANGE)
add_colored_box(slide4, Inches(7.1), Inches(2.22), Inches(5.5), Inches(0.03),
                RGBColor(0x55, 0x35, 0x10))

right_pts = [
    "External force = total mass × acceleration of CM.",
    "Internal forces cancel in pairs (Newton's 3rd Law).",
    "Only external forces change the state of the CM.",
]
y = Inches(2.35)
for pt in right_pts:
    add_textbox(slide4, "▸  " + pt,
                Inches(7.15), y, Inches(5.4), Inches(0.65),
                font_size=12, color=LIGHT_GREY)
    y += Inches(0.72)

add_colored_box(slide4, Inches(7.1), Inches(5.0), Inches(5.5), Inches(0.7),
                RGBColor(0x40, 0x25, 0x05), ORANGE)
add_textbox(slide4, "F_ext  =  M · a_cm",
            Inches(7.1), Inches(5.05), Inches(5.5), Inches(0.6),
            font_size=18, bold=True, color=ORANGE, align=PP_ALIGN.CENTER)

add_colored_box(slide4, 0, prs.slide_height - Inches(0.08),
                prs.slide_width, Inches(0.08), ORANGE)
add_textbox(slide4, "4 / 5",
            prs.slide_width - Inches(1), prs.slide_height - Inches(0.45),
            Inches(0.9), Inches(0.35),
            font_size=10, color=RGBColor(0x55, 0x55, 0x77), align=PP_ALIGN.RIGHT)


# ═══════════════════════════════════════════════════════════
# SLIDE 5 – Summary
# ═══════════════════════════════════════════════════════════
slide5 = prs.slides.add_slide(blank_layout)
set_bg(slide5, RGBColor(0x2E, 0x0D, 0x1A))
add_colored_box(slide5, 0, 0, prs.slide_width, Inches(0.08), ACCENT)

add_textbox(slide5, "📚  MH Study  ·  FY BSc Physics",
            Inches(0.4), Inches(0.15), Inches(5), Inches(0.4),
            font_size=11, color=RGBColor(0x88, 0x66, 0x77))

add_textbox(slide5, "📋   Summary",
            Inches(0.5), Inches(0.7), Inches(12), Inches(0.7),
            font_size=28, bold=True, color=ACCENT)

add_colored_box(slide5, Inches(0.5), Inches(1.45), Inches(12.33), Inches(0.04),
                RGBColor(0x55, 0x20, 0x35))

summary_items = [
    ("⚖️", "Weighted Average Position",
     "Centre of mass is the mass-weighted average position of a system."),
    ("📐", "Mass Distribution",
     "Its position depends on mass distribution, not the shape alone."),
    ("🔵", "Can Lie Outside",
     "It can lie outside the material of an object (e.g., a ring or horseshoe)."),
    ("🚀", "Momentum Relation",
     "Total momentum is related to the velocity of the CM:  P = M·v_cm"),
]

positions = [
    (Inches(0.5),  Inches(1.6)),
    (Inches(6.93), Inches(1.6)),
    (Inches(0.5),  Inches(4.1)),
    (Inches(6.93), Inches(4.1)),
]

for (icon_text, card_title, card_body), (lft, tp) in zip(summary_items, positions):
    add_colored_box(slide5, lft, tp, Inches(5.9), Inches(2.3),
                    RGBColor(0x25, 0x08, 0x14), ACCENT)
    add_textbox(slide5, icon_text + "  " + card_title,
                lft + Inches(0.2), tp + Inches(0.15),
                Inches(5.5), Inches(0.5),
                font_size=14, bold=True, color=ACCENT)
    add_colored_box(slide5, lft + Inches(0.2), tp + Inches(0.68),
                    Inches(5.5), Inches(0.03),
                    RGBColor(0x55, 0x20, 0x35))
    add_textbox(slide5, card_body,
                lft + Inches(0.2), tp + Inches(0.8),
                Inches(5.5), Inches(1.3),
                font_size=12, color=LIGHT_GREY)

add_textbox(slide5,
            "✅  End of Presentation  ·  MH Study  ·  FY BSc Physics",
            Inches(1), Inches(6.8), Inches(11.33), Inches(0.45),
            font_size=11, color=RGBColor(0x88, 0x66, 0x77), align=PP_ALIGN.CENTER)

add_colored_box(slide5, 0, prs.slide_height - Inches(0.08),
                prs.slide_width, Inches(0.08), ACCENT)
add_textbox(slide5, "5 / 5",
            prs.slide_width - Inches(1), prs.slide_height - Inches(0.45),
            Inches(0.9), Inches(0.35),
            font_size=10, color=RGBColor(0x55, 0x55, 0x77), align=PP_ALIGN.RIGHT)


# ── Save ─────────────────────────────────────────────────
output = "centre_of_mass_FYBSc.pptx"
prs.save(output)
print(f"✅  Saved: {output}")
print("   You can now share this .pptx file directly via WhatsApp, email, etc.")
