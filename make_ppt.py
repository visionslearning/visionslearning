"""
Generate Centre_of_Mass.pptx
Run: python3 make_ppt.py
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pptx.oxml.ns import qn
from lxml import etree
import copy

# ── Colours ──────────────────────────────────────────
BG          = RGBColor(0x0a, 0x0e, 0x1a)   # dark navy
PANEL       = RGBColor(0x11, 0x18, 0x27)   # slide card bg
ACCENT1     = RGBColor(0x6c, 0x63, 0xff)   # purple
ACCENT2     = RGBColor(0x00, 0xd4, 0xff)   # cyan
ACCENT3     = RGBColor(0xff, 0x6b, 0x6b)   # red
ACCENT4     = RGBColor(0xff, 0xd9, 0x3d)   # yellow
GREEN       = RGBColor(0x06, 0xd6, 0xa0)   # green
WHITE       = RGBColor(0xff, 0xff, 0xff)
MUTED       = RGBColor(0x94, 0xa3, 0xb8)
CARD_BG     = RGBColor(0x1e, 0x27, 0x3a)
DARK_CARD   = RGBColor(0x16, 0x1d, 0x2e)

# ── Slide size: widescreen 16:9 ───────────────────────
W = Inches(13.33)
H = Inches(7.5)

prs = Presentation()
prs.slide_width  = W
prs.slide_height = H

BLANK_LAYOUT = prs.slide_layouts[6]   # truly blank


# ════════════════════════════════════════════════════════
# Helper utilities
# ════════════════════════════════════════════════════════

def add_bg(slide, color: RGBColor):
    """Fill slide background with solid colour."""
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def box(slide, l, t, w, h,
        fill=None, line_color=None, line_width=Pt(1),
        radius=None):
    """Add a rectangle (optionally rounded) with optional fill/border."""
    shape = slide.shapes.add_shape(
        1,          # MSO_SHAPE_TYPE.RECTANGLE
        Inches(l), Inches(t), Inches(w), Inches(h)
    )
    shape.line.width = line_width
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    else:
        shape.fill.background()   # transparent
    if line_color:
        shape.line.color.rgb = line_color
    else:
        shape.line.fill.background()
    # rounded corners via XML
    if radius:
        sp = shape.element
        spPr = sp.find(qn('p:spPr'))
        prstGeom = spPr.find(qn('a:prstGeom'))
        if prstGeom is not None:
            prstGeom.attrib['prst'] = 'roundRect'
            avLst = prstGeom.find(qn('a:avLst'))
            if avLst is None:
                avLst = etree.SubElement(prstGeom, qn('a:avLst'))
            else:
                for gd in avLst.findall(qn('a:gd')):
                    avLst.remove(gd)
            gd = etree.SubElement(avLst, qn('a:gd'))
            gd.set('name', 'adj')
            gd.set('fmla', f'val {radius}')
    return shape


def txt(slide, text, l, t, w, h,
        size=18, bold=False, color=WHITE,
        align=PP_ALIGN.LEFT, italic=False, wrap=True):
    """Add a text box."""
    txb = slide.shapes.add_textbox(
        Inches(l), Inches(t), Inches(w), Inches(h)
    )
    txb.word_wrap = wrap
    tf = txb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size  = Pt(size)
    run.font.bold  = bold
    run.font.color.rgb = color
    run.font.italic = italic
    return txb


def multiline_txt(slide, lines, l, t, w, h,
                  size=14, color=MUTED, align=PP_ALIGN.LEFT,
                  line_spacing=None):
    """Add a textbox with multiple paragraphs."""
    txb = slide.shapes.add_textbox(
        Inches(l), Inches(t), Inches(w), Inches(h)
    )
    txb.word_wrap = True
    tf = txb.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if isinstance(line, dict):
            run = p.add_run()
            run.text = line['text']
            run.font.size  = Pt(line.get('size', size))
            run.font.bold  = line.get('bold', False)
            run.font.color.rgb = line.get('color', color)
            run.font.italic = line.get('italic', False)
        else:
            run = p.add_run()
            run.text = line
            run.font.size  = Pt(size)
            run.font.color.rgb = color
    return txb


def gradient_title(slide, text, l, t, w, h, size=36):
    """Title with purple-ish colour (pptx can't do per-char gradient easily)."""
    txt(slide, text, l, t, w, h, size=size, bold=True, color=ACCENT1, align=PP_ALIGN.LEFT)


def slide_label(slide, num, topic):
    txt(slide, f"SLIDE {num:02d} / 05  ·  {topic}",
        0.5, 0.18, 12, 0.35, size=9, bold=True, color=ACCENT1)


def tag_pill(slide, text, l, t):
    """Small tag pill."""
    box(slide, l, t, len(text)*0.085 + 0.3, 0.28,
        fill=RGBColor(0x1a, 0x15, 0x44),
        line_color=ACCENT1, line_width=Pt(0.75), radius=20000)
    txt(slide, text, l+0.08, t+0.02, len(text)*0.085+0.15, 0.26,
        size=8, bold=True, color=ACCENT1)


def divider(slide, t):
    """Thin horizontal rule."""
    box(slide, 0.5, t, 12.33, 0.02, fill=RGBColor(0x2a,0x35,0x50),
        line_color=None)


def formula_card(slide, icon, label, eq, desc, l, t, w=3.9):
    """Draw a formula card block."""
    h_card = 1.85
    box(slide, l, t, w, h_card, fill=DARK_CARD,
        line_color=ACCENT2, line_width=Pt(0.8), radius=15000)
    # icon
    txt(slide, icon, l+0.18, t+0.12, 0.5, 0.4, size=18)
    # label
    txt(slide, label, l+0.18, t+0.55, w-0.36, 0.28,
        size=8, bold=True, color=ACCENT2)
    # equation box
    box(slide, l+0.18, t+0.85, w-0.36, 0.42,
        fill=RGBColor(0x08,0x10,0x22),
        line_color=ACCENT2, line_width=Pt(1.2), radius=8000)
    txt(slide, eq, l+0.28, t+0.88, w-0.56, 0.38,
        size=11, bold=True, color=WHITE)
    # desc
    txt(slide, desc, l+0.18, t+1.33, w-0.36, 0.45,
        size=8.5, color=MUTED, wrap=True)


def point_card(slide, num, text, t):
    """Draw a numbered point row."""
    row_h = 0.5
    box(slide, 0.5, t, 12.33, row_h, fill=DARK_CARD,
        line_color=RGBColor(0x2a,0x35,0x50), line_width=Pt(0.6), radius=10000)
    # left accent stripe (simulated with thin box)
    box(slide, 0.5, t, 0.06, row_h,
        fill=ACCENT1, line_color=None, radius=0)
    # number circle
    box(slide, 0.72, t+0.09, 0.32, 0.32,
        fill=ACCENT1, line_color=None, radius=99999)
    txt(slide, str(num), 0.72, t+0.09, 0.32, 0.32,
        size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    # text
    txt(slide, text, 1.18, t+0.07, 11.4, row_h-0.1,
        size=11, color=WHITE, wrap=True)


def momentum_card(slide, icon_title, bullets, eq, l, t, w=5.9, accent=GREEN):
    h_card = 3.0
    box(slide, l, t, w, h_card, fill=DARK_CARD,
        line_color=accent, line_width=Pt(0.8), radius=15000)
    # header
    txt(slide, icon_title, l+0.2, t+0.18, w-0.4, 0.4,
        size=13, bold=True, color=accent)
    # bullets
    by = t + 0.65
    for b in bullets:
        txt(slide, f"▸  {b}", l+0.25, by, w-0.45, 0.42,
            size=9.5, color=MUTED, wrap=True)
        by += 0.42
    # equation box
    box(slide, l+0.2, t+2.45, w-0.4, 0.42,
        fill=RGBColor(0x08,0x10,0x22),
        line_color=accent, line_width=Pt(1), radius=8000)
    txt(slide, eq, l+0.2, t+2.48, w-0.4, 0.38,
        size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


def summary_card(slide, icon, text, l, t, w=2.9, h=1.9):
    box(slide, l, t, w, h, fill=DARK_CARD,
        line_color=RGBColor(0x2a,0x35,0x50), line_width=Pt(0.6), radius=15000)
    txt(slide, icon, l+w/2-0.25, t+0.18, 0.5, 0.5,
        size=22, align=PP_ALIGN.CENTER)
    txt(slide, text, l+0.18, t+0.76, w-0.36, h-0.85,
        size=9.5, color=MUTED, wrap=True, align=PP_ALIGN.CENTER)


# ════════════════════════════════════════════════════════
# SLIDE 1 — Title / Introduction
# ════════════════════════════════════════════════════════
s1 = prs.slides.add_slide(BLANK_LAYOUT)
add_bg(s1, BG)

# Card background
box(s1, 0.4, 0.25, 12.53, 7.0, fill=PANEL,
    line_color=RGBColor(0x3a,0x33,0x88), line_width=Pt(1), radius=20000)

slide_label(s1, 1, "FY BSc Physics")
gradient_title(s1, "Centre of Mass", 0.65, 0.6, 9, 1.1, size=44)

divider(s1, 1.85)

tag_pill(s1, "⚛️  Systems of Particles", 0.65, 2.0)
tag_pill(s1, "📚  MH Study",              3.45, 2.0)

# Body text
txt(s1,
    "The centre of mass is the unique point where the entire\n"
    "mass of a body — or system of particles — appears to\n"
    "be concentrated.",
    0.65, 2.55, 7.2, 1.4, size=15, color=MUTED)

txt(s1,
    "It represents the mass-weighted average position of all\n"
    "matter in the system, and is central to analysing motion.",
    0.65, 3.95, 7.2, 1.2, size=15, color=MUTED)

# Right visual — two mass system diagram
box(s1, 9.1, 1.8, 3.5, 4.8, fill=DARK_CARD,
    line_color=RGBColor(0x2a,0x35,0x50), line_width=Pt(0.6), radius=20000)

# m1 circle
box(s1, 9.4, 2.2, 1.1, 1.1, fill=ACCENT1,
    line_color=None, radius=99999)
txt(s1, "m₁", 9.4, 2.2, 1.1, 1.1,
    size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# m2 circle (smaller)
box(s1, 11.0, 2.45, 0.8, 0.8, fill=ACCENT3,
    line_color=None, radius=99999)
txt(s1, "m₂", 11.0, 2.45, 0.8, 0.8,
    size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# CM star marker
txt(s1, "★", 10.5, 3.55, 0.5, 0.5,
    size=22, bold=True, color=ACCENT4, align=PP_ALIGN.CENTER)
txt(s1, "Centre of Mass", 9.9, 4.1, 2.8, 0.4,
    size=9, color=ACCENT4, align=PP_ALIGN.CENTER)

# dashed line (simulated with thin box)
box(s1, 9.9, 2.78, 1.65, 0.06,
    fill=RGBColor(0x55,0x55,0x88), line_color=None)

txt(s1, "← → arrow keys or click arrows to navigate",
    0.65, 6.6, 12, 0.4, size=9, color=RGBColor(0x44,0x55,0x77),
    align=PP_ALIGN.CENTER)


# ════════════════════════════════════════════════════════
# SLIDE 2 — Formulae
# ════════════════════════════════════════════════════════
s2 = prs.slides.add_slide(BLANK_LAYOUT)
add_bg(s2, BG)

box(s2, 0.4, 0.25, 12.53, 7.0, fill=PANEL,
    line_color=RGBColor(0x00,0x88,0xaa), line_width=Pt(1), radius=20000)

slide_label(s2, 2, "Key Formulae")
gradient_title(s2, "🔢  Formulae", 0.65, 0.6, 9, 0.9, size=38)
divider(s2, 1.6)
txt(s2, "Mathematical expressions for Centre of Mass position",
    0.65, 1.68, 12, 0.4, size=12, color=MUTED)

formula_card(s2, "👥", "TWO PARTICLES",
             "xcm = (m₁x₁ + m₂x₂) / (m₁ + m₂)",
             "Weighted average of positions of two masses along a line.",
             0.55, 2.2, w=3.95)

formula_card(s2, "🔣", "GENERAL FORM  (N PARTICLES)",
             "xcm = Σ(mᵢ xᵢ) / Σmᵢ",
             "Summation over all particles in the system.",
             4.69, 2.2, w=3.95)

formula_card(s2, "📐", "3D POSITION VECTOR",
             "r⃗cm = Σ(mᵢ r⃗ᵢ) / M",
             "M = total mass. Extends the formula to all three dimensions.",
             8.83, 2.2, w=3.95)

# Bottom note
box(s2, 0.55, 4.25, 12.23, 0.7, fill=RGBColor(0x0a,0x14,0x28),
    line_color=RGBColor(0x00,0x88,0xaa), line_width=Pt(0.5), radius=10000)
txt(s2,
    "💡  Note: All three forms are equivalent — the vector form covers the most general case in 3D space.",
    0.75, 4.3, 12, 0.58, size=10.5, color=MUTED)


# ════════════════════════════════════════════════════════
# SLIDE 3 — Key Points
# ════════════════════════════════════════════════════════
s3 = prs.slides.add_slide(BLANK_LAYOUT)
add_bg(s3, BG)

box(s3, 0.4, 0.25, 12.53, 7.0, fill=PANEL,
    line_color=RGBColor(0xaa,0x33,0x33), line_width=Pt(1), radius=20000)

slide_label(s3, 3, "Important Points")
gradient_title(s3, "⚡  Key Points", 0.65, 0.6, 9, 0.9, size=38)
divider(s3, 1.6)
txt(s3, "Must-know concepts for exams and viva",
    0.65, 1.68, 12, 0.4, size=12, color=MUTED)

points = [
    "CM always lies closer to the heavier mass in a two-body system.",
    "For a uniform symmetric body (sphere, ring, rod), CM coincides with the geometric centre.",
    "CM can lie outside the material of an object  (e.g., a ring or hollow sphere).",
    "Position of CM is independent of the coordinate system — only mass distribution matters.",
    "Without external forces, CM moves with constant velocity (or stays at rest).",
]
for i, p in enumerate(points):
    point_card(s3, i+1, p, 2.15 + i * 0.65)


# ════════════════════════════════════════════════════════
# SLIDE 4 — CM & Momentum
# ════════════════════════════════════════════════════════
s4 = prs.slides.add_slide(BLANK_LAYOUT)
add_bg(s4, BG)

box(s4, 0.4, 0.25, 12.53, 7.0, fill=PANEL,
    line_color=RGBColor(0x06,0x88,0x66), line_width=Pt(1), radius=20000)

slide_label(s4, 4, "CM & Motion")
gradient_title(s4, "🚀  CM & Momentum", 0.65, 0.6, 9, 0.9, size=38)
divider(s4, 1.6)
txt(s4, "Connecting Newton's Laws to the Centre of Mass",
    0.65, 1.68, 12, 0.4, size=12, color=MUTED)

momentum_card(s4,
    "⚡  Linear Momentum",
    [
        "Total momentum = total mass × velocity of CM.",
        "If no external force acts, total momentum is conserved.",
        "CM continues in uniform motion.",
    ],
    "P⃗total  =  M · v⃗cm",
    l=0.55, t=2.15, w=6.0, accent=GREEN)

momentum_card(s4,
    "🔬  Newton's 2nd Law (System)",
    [
        "Net external force = total mass × acceleration of CM.",
        "Internal forces cancel in pairs (Newton's 3rd Law).",
        "Only external forces change the state of the CM.",
    ],
    "F⃗ext  =  M · a⃗cm",
    l=6.78, t=2.15, w=6.0, accent=ACCENT2)


# ════════════════════════════════════════════════════════
# SLIDE 5 — Summary
# ════════════════════════════════════════════════════════
s5 = prs.slides.add_slide(BLANK_LAYOUT)
add_bg(s5, BG)

box(s5, 0.4, 0.25, 12.53, 7.0, fill=PANEL,
    line_color=RGBColor(0x88,0x77,0x00), line_width=Pt(1), radius=20000)

slide_label(s5, 5, "Summary")
gradient_title(s5, "📋  Summary", 0.65, 0.6, 9, 0.9, size=38)
divider(s5, 1.6)
txt(s5, "Everything at a glance",
    0.65, 1.68, 12, 0.4, size=12, color=MUTED)

summaries = [
    ("⚖️",  "Centre of mass is the mass-weighted average position of a system."),
    ("📐",  "Its position depends on mass distribution, not shape alone."),
    ("🔵",  "CM can lie outside the material of an object (ring, horseshoe, hollow sphere)."),
    ("🚀",  "Total momentum relates to CM velocity:\nP = M·v⃗cm"),
]
sx = [0.55, 3.65, 6.75, 9.85]
for (icon, text), x in zip(summaries, sx):
    summary_card(s5, icon, text, l=x, t=2.2, w=2.95, h=2.5)

# End badge
box(s5, 2.0, 5.2, 9.33, 0.7, fill=RGBColor(0x0a,0x0e,0x22),
    line_color=ACCENT1, line_width=Pt(0.8), radius=20000)
txt(s5,
    "✅  End of Presentation  ·  MH Study  ·  FY BSc Physics",
    2.0, 5.28, 9.33, 0.55,
    size=12, bold=True, color=ACCENT1, align=PP_ALIGN.CENTER)


# ════════════════════════════════════════════════════════
# Save
# ════════════════════════════════════════════════════════
OUT = "Centre_of_Mass.pptx"
prs.save(OUT)
print(f"Saved → {OUT}")
