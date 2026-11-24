#!/usr/bin/env python3
"""Build the SRS Ambiguity Detector mini-project presentation.

Design language: warm paper + ink + teal + amber, matching the product UI.
Output: SRS_Ambiguity_Detector_Presentation.pptx (16:9, 10 slides).
"""
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parent.parent
SHOTS = ROOT / "screenshots"
OUT = ROOT / "SRS_Ambiguity_Detector_Presentation.pptx"

# ---------------------------------------------------------------- palette ---
PAPER = RGBColor(0xFB, 0xF8, 0xF1)
CARD = RGBColor(0xFF, 0xFF, 0xFC)
INK = RGBColor(0x17, 0x1C, 0x24)
MUTED = RGBColor(0x5C, 0x66, 0x70)
FAINT = RGBColor(0x8A, 0x91, 0x98)
LINE = RGBColor(0xDD, 0xD6, 0xC8)
TEAL = RGBColor(0x14, 0x7D, 0x73)
TEAL_D = RGBColor(0x0C, 0x5A, 0x53)
TEAL_SOFT = RGBColor(0xE1, 0xEF, 0xEA)
AMBER = RGBColor(0xC5, 0x8A, 0x1B)
AMBER_SOFT = RGBColor(0xF7, 0xEB, 0xD2)
ORANGE = RGBColor(0xC2, 0x41, 0x0C)
RED_SOFT = RGBColor(0xF9, 0xE7, 0xDE)
NAVY = RGBColor(0x0D, 0x1B, 0x2E)
NAVY2 = RGBColor(0x13, 0x2C, 0x49)
CREAM_DIM = RGBColor(0xC9, 0xD4, 0xDE)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

HEAD = "Calibri"
BODY = "Calibri"
MONO = "Consolas"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]
SW, SH = 13.333, 7.5


# ---------------------------------------------------------------- helpers ---
def slide_new(bg=PAPER):
    s = prs.slides.add_slide(BLANK)
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), prs.slide_width, prs.slide_height)
    r.fill.solid(); r.fill.fore_color.rgb = bg
    r.line.fill.background(); r.shadow.inherit = False
    return s


def shp(s, kind, x, y, w, h, fill=None, line=None, lw=0.75, adj=None):
    sp = s.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line; sp.line.width = Pt(lw)
    sp.shadow.inherit = False
    if adj is not None:
        try:
            sp.adjustments[0] = adj
        except Exception:
            pass
    return sp


def txt(s, x, y, w, h, paras, valign=MSO_ANCHOR.TOP, wrap=True):
    """paras: list of dicts {runs:[(text,size,bold,color)], align, space_after, line}"""
    box = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = valign
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, para in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = para.get("align", PP_ALIGN.LEFT)
        if para.get("space_after") is not None:
            p.space_after = Pt(para["space_after"])
        if para.get("space_before") is not None:
            p.space_before = Pt(para["space_before"])
        if para.get("line") is not None:
            p.line_spacing = para["line"]
        for (t, size, bold, color, *rest) in para["runs"]:
            r = p.add_run(); r.text = t
            r.font.size = Pt(size); r.font.bold = bold
            r.font.color.rgb = color; r.font.name = rest[0] if rest else BODY
            if len(rest) > 1 and rest[1]:
                r.font.italic = True
    return box


def P(text, size=14, bold=False, color=INK, align=PP_ALIGN.LEFT, after=4, line=1.06, font=BODY, italic=False):
    return {"runs": [(text, size, bold, color, font, italic)], "align": align, "space_after": after, "line": line}


def header(s, kicker, title, page, title_size=27):
    txt(s, 0.62, 0.34, 9.5, 0.3, [P(kicker.upper(), 10, True, TEAL, after=0, font=MONO)])
    txt(s, 0.62, 0.60, 11.6, 0.75, [P(title, title_size, True, INK, after=0, line=0.98, font=HEAD)])
    shp(s, MSO_SHAPE.RECTANGLE, 0.64, 1.28, 0.62, 0.045, fill=AMBER)
    shp(s, MSO_SHAPE.RECTANGLE, 1.30, 1.28, 11.42, 0.016, fill=LINE)
    footer(s, page)


def footer(s, page, dark=False):
    c1, c2 = (CREAM_DIM, CREAM_DIM) if dark else (FAINT, FAINT)
    txt(s, 0.62, 7.12, 8.5, 0.25, [P("SRS AMBIGUITY DETECTOR  •  MINI PROJECT  •  KANPUR INSTITUTE OF TECHNOLOGY", 7.5, False, c1, after=0, font=MONO)])
    chip = shp(s, MSO_SHAPE.OVAL, 12.55, 7.02, 0.34, 0.34, fill=(NAVY2 if dark else TEAL))
    txt(s, 12.55, 7.075, 0.34, 0.25, [P(str(page), 9, True, WHITE, align=PP_ALIGN.CENTER, after=0)])


def card(s, x, y, w, h, fill=CARD, line=LINE, adj=0.055):
    return shp(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill=fill, line=line, lw=1.0, adj=adj)


def chip(s, x, y, w, h, text, fill, color, size=10.5):
    shp(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill=fill, adj=0.5)
    txt(s, x + 0.08, y, w - 0.16, h, [P(text, size, True, color, align=PP_ALIGN.CENTER, after=0)],
        valign=MSO_ANCHOR.MIDDLE)


def arrow_r(s, x, y, color=TEAL, w=0.34):
    shp(s, MSO_SHAPE.CHEVRON, x, y, w, 0.26, fill=color, adj=0.55)


def browser(s, x, y, w, img, label, h_img=None, crop_bottom=0.0):
    """Framed app screenshot with a slim browser chrome bar."""
    h_img = h_img or w / 1.6
    card(s, x, y, w, h_img + 0.34, fill=WHITE, line=LINE)
    bar = shp(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, 0.34, fill=NAVY, adj=0.16)
    for i, c in enumerate([RGBColor(0xF2, 0x76, 0x5E), RGBColor(0xF5, 0xBF, 0x4F), RGBColor(0x61, 0xC4, 0x54)]):
        shp(s, MSO_SHAPE.OVAL, x + 0.14 + i * 0.17, y + 0.115, 0.105, 0.105, fill=c)
    txt(s, x + 0.72, y + 0.015, w - 0.85, 0.3, [P(label, 8.5, False, CREAM_DIM, after=0, font=MONO)],
        valign=MSO_ANCHOR.MIDDLE)
    pic = s.shapes.add_picture(str(img), Inches(x + 0.055), Inches(y + 0.345), Inches(w - 0.11), Inches(h_img - 0.06))
    if crop_bottom:
        pic.crop_bottom = crop_bottom
    pic.shadow.inherit = False
    return h_img + 0.34


def notes(s, text):
    s.notes_slide.notes_text_frame.text = text


def transition(s, kind="fade", dur=650):
    sld = s._element
    for el in sld.findall(qn("p:transition")):
        sld.remove(el)
    inner = {
        "fade": "<p:fade/>",
        "push": '<p:push dir="l"/>',
        "wipe": '<p:wipe dir="l"/>',
        "cover": '<p:cover dir="l"/>',
    }[kind]
    el = parse_xml(f'<p:transition {nsdecls("p")} spd="med" dur="{dur}">{inner}</p:transition>')
    idx = 0
    for i, child in enumerate(sld):
        if child.tag in (qn("p:cSld"), qn("p:clrMapOvr")):
            idx = i + 1
    sld.insert(idx, el)


def connector(s, x1, y1, x2, y2, color=TEAL, w=1.4):
    ln = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    ln.line.color.rgb = color; ln.line.width = Pt(w); ln.shadow.inherit = False
    return ln


# ============================================================== SLIDE 1 =====
s = slide_new(NAVY)
shp(s, MSO_SHAPE.OVAL, 9.3, -2.5, 6.4, 6.4, fill=NAVY2)
ring = shp(s, MSO_SHAPE.OVAL, 10.15, -1.65, 4.6, 4.6, fill=None, line=TEAL, lw=1.4)
shp(s, MSO_SHAPE.OVAL, -1.3, 5.6, 3.4, 3.4, fill=NAVY2)
txt(s, 0.72, 0.55, 7.5, 0.3, [P("MINI PROJECT  •  B.TECH CSE — SECTION A  •  7TH SEMESTER", 10.5, True, RGBColor(0x6E, 0xC8, 0xBE), after=0, font=MONO)])
logo_bg = shp(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.72, 0.95, 1.62, 0.62, fill=WHITE, adj=0.12)
s.shapes.add_picture(str(ROOT / "synopsis-src/assets/kit-logo.png"), Inches(0.82), Inches(1.03), Inches(1.42), Inches(0.46))
txt(s, 0.72, 1.95, 7.6, 1.9, [
    {"runs": [("SRS AMBIGUITY", 43, True, WHITE, HEAD)], "space_after": 0, "line": 0.95},
    {"runs": [("DETECTOR", 43, True, RGBColor(0x6E, 0xC8, 0xBE), HEAD)], "space_after": 0, "line": 0.95},
])
shp(s, MSO_SHAPE.RECTANGLE, 0.74, 3.62, 0.85, 0.06, fill=AMBER)
txt(s, 0.72, 3.86, 7.3, 1.35, [
    P("Clarity before code.", 21, True, RGBColor(0xF2, 0xC14, 0xE) if False else RGBColor(0xF2, 0xC1, 0x4E), after=6),
    P("A deterministic web application that detects vague, incomplete and unverifiable "
      "requirements — and explains exactly how to fix them, before they become defects.",
      14.5, False, CREAM_DIM, after=0, line=1.22),
])
# screenshot, slightly tilted
back = shp(s, MSO_SHAPE.ROUNDED_RECTANGLE, 8.05, 1.92, 4.62, 3.42, fill=WHITE, adj=0.035)
back.rotation = -1.6
pic = s.shapes.add_picture(str(SHOTS / "08-analysis-result--1440x900.png"), Inches(8.16), Inches(2.03), Inches(4.40), Inches(3.20))
pic.rotation = -1.6; pic.shadow.inherit = False
chip(s, 8.30, 5.48, 1.34, 0.40, "11 DETECTORS", TEAL, WHITE, 8.8)
chip(s, 9.76, 5.48, 1.34, 0.40, "73 / 100 DEMO", AMBER, NAVY, 8.8)
chip(s, 11.22, 5.48, 1.46, 0.40, "1,023 TESTS", NAVY2, WHITE, 8.8)
pres = card(s, 0.72, 5.62, 7.05, 1.18, fill=NAVY2, line=NAVY2)
txt(s, 0.98, 5.82, 6.7, 0.85, [
    {"runs": [("Presented by   ", 11, False, CREAM_DIM), ("Ashish Kumar", 14, True, WHITE),
              ("   •   Roll No. 2301650100054", 11.5, False, CREAM_DIM)], "space_after": 5},
    {"runs": [("Guided by   ", 11, False, CREAM_DIM), ("Mr. Suraj Srivastava", 12.5, True, WHITE),
              ("   •   Kanpur Institute of Technology, Kanpur", 11, False, CREAM_DIM)], "space_after": 0},
])
txt(s, 0.72, 6.98, 9.0, 0.3, [P("AFFILIATED TO DR. A.P.J. ABDUL KALAM TECHNICAL UNIVERSITY, LUCKNOW", 8, False, RGBColor(0x7E, 0x90, 0xA4), after=0, font=MONO)])
transition(s, "fade", 800)
notes(s, "Good morning, respected faculty. My mini project is the SRS Ambiguity Detector — a working web "
         "application that reviews Software Requirements Specifications and flags ambiguous requirements "
         "before coding begins. My premise in one line: clarity before code. In the next few minutes I will "
         "show the problem, the solution, the architecture, the detection engine, live output, results, "
         "security, and future scope. [~40 seconds]")

# ============================================================== SLIDE 2 =====
s = slide_new()
header(s, "01 • The problem", "One vague sentence can fork the whole project", 2)
txt(s, 0.62, 1.48, 12.1, 0.4, [P("Requirements are written in natural language — and natural language quietly hides multiple meanings.", 14.5, False, MUTED, after=0)])
examples = [
    ("“The system shall respond quickly.”", "quickly", "Subjective + unmeasurable", "no limit, no load, no test."),
    ("“The system shall support many users.”", "many", "Unknown capacity", "10 users, or 10,000?"),
    ("“It shall generate reports as needed.”", "it  •  as needed", "Unclear actor + trigger", "who acts, and when?"),
]
y = 2.02
for quote, flag, tag, desc in examples:
    card(s, 0.62, y, 6.55, 0.92)
    shp(s, MSO_SHAPE.RECTANGLE, 0.62, y, 0.07, 0.92, fill=ORANGE)
    txt(s, 0.92, y + 0.10, 6.05, 0.4, [P(quote, 14.5, True, INK, after=0, font=MONO)])
    txt(s, 0.92, y + 0.52, 6.05, 0.32, [
        {"runs": [(flag + "  →  ", 10.5, True, ORANGE, MONO), (tag, 10.5, True, TEAL_D),
                  ("  —  " + desc, 10.5, False, MUTED)], "space_after": 0}])
    y += 1.06
card(s, 0.62, 5.28, 12.09, 0.78, fill=AMBER_SOFT, line=AMBER_SOFT)
txt(s, 0.95, 5.44, 11.5, 0.5, [
    {"runs": [("The gap:  ", 14, True, ORANGE), ("manual review is expert work — but slow, subjective and easy to miss at scale. "
              "The same sentence is read differently by the analyst, developer and tester.", 13.5, False, INK)], "space_after": 0, "line": 1.15}])
# right impact chain
steps = [("AMBIGUOUS REQUIREMENT", "written once, read many ways"), ("DIVERGENT INTERPRETATION", "each stakeholder guesses"),
         ("WRONG BUILD • UNTESTABLE CASE", "code passes, intent fails"), ("LATE REWORK & DISPUTES", "costliest place to discover it")]
yy = 2.02
for i, (t1, t2) in enumerate(steps):
    shp(s, MSO_SHAPE.ROUNDED_RECTANGLE, 7.55, yy, 5.16, 0.62, fill=(NAVY if i == 3 else CARD), line=(NAVY if i == 3 else LINE), lw=1.0, adj=0.14)
    txt(s, 7.82, yy + 0.075, 4.7, 0.5, [
        P(t1, 11.5, True, (WHITE if i == 3 else INK), after=1, font=MONO),
        P(t2, 10, False, (CREAM_DIM if i == 3 else FAINT), after=0)])
    if i < 3:
        shp(s, MSO_SHAPE.ISOSCELES_TRIANGLE, 9.98, yy + 0.63, 0.26, 0.18, fill=AMBER).rotation = 180
    yy += 0.83
transition(s, "wipe")
notes(s, "The problem: an SRS is the contract between client, developer and tester, yet it is written in "
         "imprecise natural language. Three real patterns from my sample data: 'quickly' cannot be tested; "
         "'many users' hides capacity; 'it' and 'as needed' hide the actor and trigger. Each stakeholder "
         "fills the gap with a different assumption, so the defect is discovered after coding — the most "
         "expensive point. Manual review catches some of this, but it is slow and inconsistent. [~60 seconds]")

# ============================================================== SLIDE 3 =====
s = slide_new()
header(s, "02 • The solution", "Paste an SRS. Get a quality report in seconds.", 3)
flow = [("01", "PASTE / UPLOAD", "Text up to 200,000 chars, or PDF • DOCX • TXT up to 10 MB"),
        ("02", "SEGMENT", "Split into individual, ordered requirement statements"),
        ("03", "DETECT", "11 deterministic detectors scan every requirement"),
        ("04", "EXPLAIN", "Phrase, severity, reason and a concrete recommendation"),
        ("05", "SCORE", "0–100 score, ambiguity band and 4 health dimensions"),
        ("06", "SAVE & TRACK", "Reports, searchable history and dashboard trends")]
x = 0.62
for i, (num, t1, t2) in enumerate(flow):
    card(s, x, 1.60, 1.88, 1.62, fill=(TEAL if i == 2 else CARD), line=(TEAL if i == 2 else LINE))
    txt(s, x + 0.16, 1.76, 1.56, 1.32, [
        P(num, 15, True, (AMBER if i == 2 else AMBER), after=3, font=MONO),
        P(t1, 11, True, (WHITE if i == 2 else INK), after=4, font=MONO),
        P(t2, 9.5, False, (RGBColor(0xD7, 0xEC, 0xE7) if i == 2 else MUTED), after=0, line=1.12)])
    if i < 5:
        arrow_r(s, x + 1.90, 2.28, TEAL if i != 2 else AMBER, 0.16)
    x += 2.035
browser(s, 0.62, 3.52, 6.55, SHOTS / "07-analyzer-input--1440x900.png", "srs-detector.app/analyzer", h_img=3.20, crop_bottom=0.20)
benefits = [("EXPLAINABLE, NOT A BLACK BOX", "Every finding names the exact phrase, why it is risky, and how to rewrite it."),
            ("DETERMINISTIC & REPEATABLE", "Same input → same findings and same score. No randomness, no hidden weighting."),
            ("A COMPLETE PRODUCT, NOT A SCRIPT", "Verified accounts, saved reports, history, dashboard, privacy controls and optional provider-assisted rewrites.")]
yy = 3.58
for t1, t2 in benefits:
    card(s, 7.48, yy, 5.23, 0.98)
    shp(s, MSO_SHAPE.OVAL, 7.70, yy + 0.17, 0.30, 0.30, fill=TEAL_SOFT)
    txt(s, 7.70, yy + 0.205, 0.30, 0.25, [P("✓", 11, True, TEAL_D, align=PP_ALIGN.CENTER, after=0)])
    txt(s, 8.18, yy + 0.13, 4.35, 0.75, [P(t1, 11.5, True, INK, after=2, font=MONO), P(t2, 10.5, False, MUTED, after=0, line=1.1)])
    yy += 1.12
transition(s, "push")
notes(s, "The solution is a full web application. The user pastes SRS text or uploads a PDF, DOCX or TXT "
         "document. The system segments it into requirements, runs eleven deterministic detectors, and "
         "returns explanations plus a transparent score. Three design promises: explainable findings, "
         "deterministic repeatability, and a complete product with accounts, saved reports and trends — "
         "not a one-off script. The screen below is the actual analyzer page. [~55 seconds]")

# ============================================================== SLIDE 4 =====
s = slide_new()
header(s, "03 • System architecture", "A layered system with a pure, deterministic core", 4)
# diagram
card(s, 0.62, 1.62, 12.09, 4.62, fill=WHITE)
# tier labels
for x0, t1 in [(0.90, "CLIENT"), (3.30, "PRESENTATION"), (5.95, "APPLICATION • /api/v1"), (9.15, "DATA & SERVICES")]:
    txt(s, x0, 1.80, 2.6, 0.25, [P(t1, 9, True, FAINT, after=0, font=MONO)])
# browser
shp(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.90, 3.02, 1.78, 1.30, fill=NAVY, adj=0.07)
txt(s, 1.06, 3.28, 1.48, 0.85, [P("USER'S", 11, True, WHITE, align=PP_ALIGN.CENTER, after=1, font=MONO),
                               P("BROWSER", 11, True, WHITE, align=PP_ALIGN.CENTER, after=4, font=MONO),
                               P("desktop • mobile", 8.5, False, CREAM_DIM, align=PP_ALIGN.CENTER, after=0)])
# frontend
shp(s, MSO_SHAPE.ROUNDED_RECTANGLE, 3.18, 2.72, 2.20, 1.90, fill=TEAL, adj=0.055)
txt(s, 3.38, 2.92, 1.82, 1.55, [P("NEXT.JS FRONT END", 11.5, True, WHITE, after=4, font=MONO),
                               P("React 19 • TypeScript", 10.5, False, WHITE, after=2),
                               P("Tailwind • Recharts", 10.5, False, WHITE, after=6),
                               P("public pages + signed-in app", 9, False, RGBColor(0xD7, 0xEC, 0xE7), after=0, line=1.1)])
# backend layers
shp(s, MSO_SHAPE.ROUNDED_RECTANGLE, 6.02, 2.20, 2.68, 2.95, fill=CARD, line=TEAL, lw=1.4, adj=0.045)
txt(s, 6.24, 2.34, 2.25, 0.3, [P("FASTAPI BACK END", 11.5, True, TEAL_D, after=0, font=MONO)])
layers = [("ROUTERS", "validate + authorize"), ("SERVICES", "business logic"), ("REPOSITORIES", "owner-scoped SQL")]
yy = 2.72
for t1, t2 in layers:
    shp(s, MSO_SHAPE.ROUNDED_RECTANGLE, 6.24, yy, 2.24, 0.52, fill=TEAL_SOFT, adj=0.12)
    txt(s, 6.40, yy + 0.065, 1.95, 0.4, [P(t1, 9.5, True, TEAL_D, after=0, font=MONO), P(t2, 8.5, False, MUTED, after=0)])
    yy += 0.62
shp(s, MSO_SHAPE.ROUNDED_RECTANGLE, 6.24, yy + 0.02, 2.24, 0.52, fill=AMBER_SOFT, adj=0.12)
txt(s, 6.40, yy + 0.085, 1.95, 0.4, [P("DETECTION ENGINE", 9.5, True, ORANGE, after=0, font=MONO), P("pure functions • zero I/O", 8.5, False, MUTED, after=0)])
# data/services
shp(s, MSO_SHAPE.ROUNDED_RECTANGLE, 9.35, 2.20, 2.95, 1.28, fill=NAVY, adj=0.055)
txt(s, 9.60, 2.38, 2.5, 0.95, [P("POSTGRESQL 16", 12, True, WHITE, after=3, font=MONO),
                              P("10 tables • Alembic migrations 0001–0006", 9.5, False, CREAM_DIM, after=2),
                              P("owner_id on every user-owned row", 9.5, False, CREAM_DIM, after=0)])
services = [("DOCUMENT EXTRACTION", "pypdf • python-docx • TXT"), ("EMAIL & SECURITY", "verification • reset • Turnstile-ready"),
            ("OPTIONAL ENHANCEMENT", "user-key provider overview + rewrites")]
yy = 3.66
for t1, t2 in services:
    shp(s, MSO_SHAPE.ROUNDED_RECTANGLE, 9.35, yy, 2.95, 0.48, fill=CARD, line=LINE, adj=0.12)
    txt(s, 9.56, yy + 0.055, 2.6, 0.38, [P(t1, 8.8, True, INK, after=0, font=MONO), P(t2, 8.2, False, MUTED, after=0)])
    yy += 0.58
# arrows
for (x1, x2) in [(2.68, 3.18), (5.38, 6.02), (8.70, 9.35)]:
    arrow_r(s, x1 + 0.03, 3.52, TEAL, 0.34)
txt(s, 2.62, 3.24, 0.62, 0.2, [P("HTTPS", 7.5, True, FAINT, align=PP_ALIGN.CENTER, after=0, font=MONO)])
txt(s, 5.34, 3.24, 0.70, 0.2, [P("REST JSON", 7.5, True, FAINT, align=PP_ALIGN.CENTER, after=0, font=MONO)])
pills = [("PURE ENGINE", "identical input → identical output"), ("OWNER-SCOPED DATA", "one user's data is structurally separate"),
         ("VERSIONED SCHEMA", "migrations only — never hand edits"), ("TYPED CONTRACTS", "Pydantic + TypeScript end to end")]
x = 0.62
for t1, t2 in pills:
    shp(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, 6.42, 2.94, 0.52, fill=TEAL_SOFT, adj=0.16)
    txt(s, x + 0.14, 6.485, 2.68, 0.4, [{"runs": [(t1 + "  ", 9, True, TEAL_D, MONO), (t2, 8.3, False, MUTED)], "space_after": 0, "line": 1.0}],
        valign=MSO_ANCHOR.MIDDLE)
    x += 3.05
transition(s, "fade")
notes(s, "Architecture: browser to Next.js front end, then a versioned REST API into FastAPI. The back end "
         "is layered — routers validate, services hold business logic, repositories own all SQL. The "
         "detection engine is a set of pure functions with no database or network calls, which makes it "
         "fast, testable and deterministic. PostgreSQL holds ten tables; every user-owned row carries "
         "owner_id, and the schema evolves only through Alembic migrations. [~60 seconds]")

# ============================================================== SLIDE 5 =====
s = slide_new()
header(s, "04 • Detection engine", "Eleven detectors. Every finding explained.", 5)
detectors = [
    ("VAGUE QUANTIFIER", "“many” • “several”"), ("SUBJECTIVE TERM", "“fast” • “user-friendly”"),
    ("MISSING MEASURABLE CRITERIA", "a claim with no number"), ("PRONOUN REFERENCE", "“it” • “they” — who?"),
    ("OPTIONAL LANGUAGE", "“may” • “as needed”"), ("AMBIGUOUS OPERATOR", "“and/or” • “etc.”"),
    ("UNDEFINED TERMINOLOGY", "“Premium User” — defined?"), ("ABSOLUTE LANGUAGE", "“always” • “never”"),
    ("PASSIVE ACTOR", "“shall be approved” — by whom?"), ("MISSING CONSTRAINT", "no limit • timing • retention"),
    ("INCOMPLETE REQUIREMENT", "fragment • “TBD” • dangling modal"),
]
x0, y0, cw, chh, gap = 0.62, 1.60, 2.94, 1.06, 0.115
for i, (t1, t2) in enumerate(detectors):
    x = x0 + (i % 4) * (cw + gap); y = y0 + (i // 4) * (chh + gap)
    card(s, x, y, cw, chh)
    shp(s, MSO_SHAPE.OVAL, x + 0.14, y + 0.13, 0.34, 0.34, fill=TEAL)
    txt(s, x + 0.14, y + 0.185, 0.34, 0.25, [P(f"{i+1:02d}", 8.5, True, WHITE, align=PP_ALIGN.CENTER, after=0, font=MONO)])
    txt(s, x + 0.58, y + 0.12, cw - 0.70, 0.82, [
        P(t1, 9.8, True, INK, after=3, font=MONO, line=1.02),
        P(t2, 9.8, False, MUTED, after=0, font=MONO)])
# scoring card (12th cell)
x = x0 + 3 * (cw + gap); y = y0 + 2 * (chh + gap)
shp(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, cw, chh, fill=NAVY, adj=0.055)
txt(s, x + 0.18, y + 0.11, cw - 0.36, 0.86, [
    P("TRANSPARENT SCORING", 9.8, True, RGBColor(0x6E, 0xC8, 0xBE), after=3, font=MONO),
    {"runs": [("100 − Σ deductions", 12.5, True, WHITE, MONO)], "space_after": 2},
    P("Low 5 • Med 10 • High 15 • Critical 20", 8.8, False, CREAM_DIM, after=0, font=MONO)])
# bottom band
card(s, 0.62, 5.22, 12.09, 1.18, fill=TEAL_SOFT, line=TEAL_SOFT)
txt(s, 0.92, 5.40, 11.5, 0.85, [
    {"runs": [("FROM FINDINGS TO FOUR HEALTH SIGNALS    ", 10.5, True, TEAL_D, MONO),
              ("Measurability", 13, True, INK), ("  (can we test it?)   •   ", 11.5, False, MUTED),
              ("Specificity", 13, True, INK), ("  (is it precise?)   •   ", 11.5, False, MUTED),
              ("Clarity", 13, True, INK), ("  (is it unambiguous?)   •   ", 11.5, False, MUTED),
              ("Completeness", 13, True, INK), ("  (is anything missing?)", 11.5, False, MUTED)], "space_after": 5, "line": 1.15},
    P("Bands: 80–100 Low  •  60–79 Moderate  •  40–59 High  •  0–39 Very high ambiguity. Same input always produces the same score.",
      11, False, MUTED, after=0)])
transition(s, "wipe")
notes(s, "The engine is a registry of eleven independent detectors. Each returns the matched phrase with "
         "character offsets, severity, reason and recommendation. Findings are deduplicated "
         "deterministically. Scoring is deliberately transparent: start at 100 and subtract fixed points "
         "per severity — there is no hidden model. The same deductions are grouped into four health "
         "dimensions: measurability, specificity, clarity and completeness. [~65 seconds]")

# ============================================================== SLIDE 6 =====
s = slide_new()
header(s, "05 • Working product", "Actual screens from the running application", 6)
browser(s, 0.62, 1.58, 6.70, SHOTS / "09-analysis-report--1440x900.png", "srs-detector.app/analysis/report")
browser(s, 7.62, 1.58, 5.09, SHOTS / "06-dashboard--1440x900.png", "srs-detector.app/dashboard", h_img=3.18)
callouts = [("1", "SCORE & BAND", "73/100 with severity counts at a glance"),
            ("2", "CATEGORIES + HEALTH", "where the document is weak, dimension by dimension"),
            ("3", "TRENDS OVER TIME", "saved analyses turn quality into a trackable metric")]
x = 7.62
yy = 5.10
for n, t1, t2 in callouts[:1]:
    pass
items = [("1  SCORE & BAND", "73 / 100, Moderate — with Low 3 • Medium 10 • High 5"),
         ("2  EXPLAINED FINDINGS", "category counts and four health dimensions per report"),
         ("3  HISTORY & DASHBOARD", "every run saved; quality trends visible across drafts")]
yy = 5.06
for t1, t2 in items:
    shp(s, MSO_SHAPE.ROUNDED_RECTANGLE, 7.62, yy, 5.09, 0.52, fill=CARD, line=LINE, adj=0.14)
    txt(s, 7.84, yy + 0.065, 4.7, 0.4, [P(t1, 9.5, True, TEAL_D, after=0, font=MONO), P(t2, 9, False, MUTED, after=0)])
    yy += 0.615
txt(s, 0.62, 6.22, 6.7, 0.55, [
    {"runs": [("Same content, two input paths:  ", 12, True, INK), ("pasted text and an uploaded TXT file produced identical reports (73/100, 18 issues) — one shared pipeline.", 12, False, MUTED)], "space_after": 0, "line": 1.18}])
transition(s, "fade")
notes(s, "These are real screens from the application running locally with PostgreSQL. Left: a saved "
         "report — score ring, severity counts, category distribution and health dimensions. Right: the "
         "dashboard aggregating saved analyses over time. The same sample submitted as pasted text and "
         "as an uploaded file produced identical results, proving both input paths share one pipeline. "
         "If time permits after the slides, I can demonstrate this live. [~55 seconds]")

# ============================================================== SLIDE 7 =====
s = slide_new()
header(s, "06 • Results & validation", "Measured on a deliberately weak sample SRS", 7)
# left metric panel
card(s, 0.62, 1.58, 3.55, 4.92)
txt(s, 0.92, 1.80, 3.0, 0.3, [P("DEMO RUN — 7 REQUIREMENTS", 9.5, True, FAINT, after=0, font=MONO)])
ringc = shp(s, MSO_SHAPE.OVAL, 1.28, 2.22, 1.62, 1.62, fill=CARD, line=AMBER, lw=9)
txt(s, 1.28, 2.72, 1.62, 0.65, [P("73", 34, True, INK, align=PP_ALIGN.CENTER, after=0), P("/ 100", 10, False, FAINT, align=PP_ALIGN.CENTER, after=0)])
txt(s, 0.92, 3.98, 2.95, 0.35, [P("MODERATE AMBIGUITY", 12.5, True, ORANGE, align=PP_ALIGN.CENTER, after=0, font=MONO)])
# severity segmented bar
segs = [("Low", 3, TEAL, "3"), ("Medium", 10, AMBER, "10"), ("High", 5, ORANGE, "5"), ("Critical", 0, RGBColor(0xD8, 0xD2, 0xC6), "0")]
x = 0.92; total_w = 2.95
for name, val, col, lab in segs:
    w = max(total_w * val / 18, 0.16) if val else 0.16
    shp(s, MSO_SHAPE.RECTANGLE, x, 4.48, w, 0.26, fill=col)
    x += w
labels = [("LOW 3", 0.92), ("MEDIUM 10", 1.62), ("HIGH 5", 2.78), ("CRIT 0", 3.42)]
for t1, lx in labels:
    txt(s, lx, 4.80, 0.85, 0.2, [P(t1, 7.5, True, MUTED, after=0, font=MONO)])
txt(s, 0.92, 5.14, 2.95, 1.1, [P("18 findings — each with phrase, severity, reason and recommendation.", 11, False, MUTED, after=4, line=1.15),
                               P("Text and file-upload runs matched exactly.", 11, True, TEAL_D, after=0, line=1.15)])
# middle health panel
card(s, 4.42, 1.58, 4.35, 4.92)
txt(s, 4.72, 1.80, 3.8, 0.3, [P("REQUIREMENT HEALTH ( / 100 )", 9.5, True, FAINT, after=0, font=MONO)])
health = [("Measurability", 0, ORANGE), ("Clarity", 45, AMBER), ("Specificity", 65, TEAL), ("Completeness", 100, TEAL)]
yy = 2.28
for name, val, col in health:
    txt(s, 4.72, yy, 2.6, 0.25, [P(name, 11.5, True, INK, after=0)])
    txt(s, 7.85, yy, 0.62, 0.25, [P(str(val), 11.5, True, col, align=PP_ALIGN.RIGHT, after=0, font=MONO)])
    shp(s, MSO_SHAPE.ROUNDED_RECTANGLE, 4.72, yy + 0.30, 3.75, 0.14, fill=RGBColor(0xED, 0xE7, 0xDB), adj=0.5)
    if val:
        shp(s, MSO_SHAPE.ROUNDED_RECTANGLE, 4.72, yy + 0.30, 3.75 * val / 100, 0.14, fill=col, adj=0.5)
    yy += 0.78
txt(s, 4.72, 5.42, 3.78, 0.9, [
    {"runs": [("Reading the result:  ", 11, True, INK), ("every performance claim lacked a metric (Measurability 0), while no requirement was structurally incomplete (Completeness 100).", 11, False, MUTED)], "space_after": 0, "line": 1.18}])
# right verification panel
shp(s, MSO_SHAPE.ROUNDED_RECTANGLE, 9.02, 1.58, 3.69, 4.92, fill=NAVY, adj=0.045)
txt(s, 9.32, 1.80, 3.1, 0.3, [P("VERIFICATION EVIDENCE", 9.5, True, RGBColor(0x6E, 0xC8, 0xBE), after=0, font=MONO)])
stats = [("1,023", "automated tests across back-end (591) and front-end (432) suites"),
         ("87 / 87", "core detector, segmentation & scoring tests re-run for this review"),
         ("19", "front-end routes in the production build, fully type-checked"),
         ("~45,000", "lines of code including tests and engineering documentation")]
yy = 2.24
for big, small in stats:
    txt(s, 9.32, yy, 3.15, 0.75, [P(big, 20, True, WHITE, after=0, font=MONO), P(small, 9.5, False, CREAM_DIM, after=0, line=1.08)])
    yy += 0.98
    if big != "~45,000":
        shp(s, MSO_SHAPE.RECTANGLE, 9.32, yy - 0.10, 3.05, 0.012, fill=NAVY2)
transition(s, "push")
notes(s, "Results from running the actual system on seven deliberately weak requirements: score 73 out "
         "of 100, Moderate band, 18 findings. Health dimensions tell the story — measurability is zero "
         "because every quality claim lacked a number, while completeness is 100. On verification: the "
         "repository carries 1,023 automated tests, the production build generates 19 routes, and I "
         "re-ran the 87 core engine tests while preparing this presentation — all passed. [~65 seconds]")

# ============================================================== SLIDE 8 =====
s = slide_new()
header(s, "07 • Security & quality", "Built like production software — not a demo script", 8)
sec = [("IDENTITY YOU CAN TRUST", "Argon2id password hashing, email verification, HttpOnly cookie sessions with refresh-token rotation and reuse detection."),
       ("OWNERSHIP ON EVERY QUERY", "Every user-owned row carries owner_id; repositories filter by owner, so another user's report simply returns “not found”."),
       ("AN ENCRYPTED KEY VAULT", "Optional provider keys are stored only as Fernet ciphertext; the UI ever shows just the last four characters."),
       ("HOSTILE-FILE HANDLING", "Uploads must pass type, magic-byte, 10 MB and extraction-budget checks; downloads use short-lived signed links."),
       ("ABUSE & CSRF CONTROLS", "Per-route rate limits, server-side Origin checks on cookie-authenticated requests, security headers and scrubbed logging."),
       ("PRIVACY AS A FEATURE", "Data export, history purge, retention controls and full account deletion are built in — not bolted on.")]
pos = [(0.62, 1.58), (4.80, 1.58), (8.98, 1.58), (0.62, 3.62), (4.80, 3.62), (8.98, 3.62)]
for (t1, t2), (x, y) in zip(sec, pos):
    card(s, x, y, 3.73, 1.82)
    shp(s, MSO_SHAPE.OVAL, x + 0.18, y + 0.18, 0.34, 0.34, fill=TEAL_SOFT)
    txt(s, x + 0.18, y + 0.235, 0.34, 0.25, [P("◆", 10, True, TEAL_D, align=PP_ALIGN.CENTER, after=0)])
    txt(s, x + 0.66, y + 0.16, 2.9, 1.5, [P(t1, 10.8, True, INK, after=4, font=MONO, line=1.05), P(t2, 11.0, False, MUTED, after=0, line=1.2)])
stack = [("FRONT END", "Next.js 16 • React 19 • TypeScript • Tailwind • Recharts"),
         ("BACK END", "FastAPI • Pydantic v2 • SQLAlchemy 2.0 (async)"),
         ("DATA", "PostgreSQL 16 • Alembic migrations"),
         ("TESTING", "Pytest • Vitest • Testing Library")]
x = 0.62
for t1, t2 in stack:
    shp(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, 5.72, 2.94, 0.78, fill=NAVY, adj=0.10)
    txt(s, x + 0.16, 5.84, 2.65, 0.55, [P(t1, 9, True, RGBColor(0x6E, 0xC8, 0xBE), after=2, font=MONO), P(t2, 8.8, False, WHITE, after=0, line=1.05)])
    x += 3.05
transition(s, "fade")
notes(s, "Security was treated as a core requirement. Passwords use Argon2id; sessions are HttpOnly "
         "cookies with rotating refresh tokens; ownership is enforced inside the SQL queries themselves; "
         "provider keys exist only as Fernet ciphertext; uploads are treated as hostile until type, "
         "magic bytes and size are proven. Privacy lifecycle — export, purge, retention, deletion — is "
         "implemented end to end. The stack is shown at the bottom. [~60 seconds]")

# ============================================================== SLIDE 9 =====
s = slide_new()
header(s, "08 • Honest limits → future scope", "A trustworthy baseline, with a clear road ahead", 9)
card(s, 0.62, 1.58, 5.72, 4.55, fill=AMBER_SOFT, line=AMBER_SOFT)
txt(s, 0.94, 1.80, 5.1, 0.35, [P("WHAT VERSION 1 DELIBERATELY DOES NOT CLAIM", 10.5, True, ORANGE, after=0, font=MONO)])
limits = [("Heuristic, not semantic", "Pattern-based findings can over- or under-flag; the analyst keeps final judgement."),
          ("English-first", "Detectors and segmentation signals are tuned for English SRS documents."),
          ("No OCR", "Scanned or image-only PDFs cannot be read yet."),
          ("Paragraph-level segmentation", "Multi-sentence paragraphs stay as one requirement in this version.")]
yy = 2.30
for t1, t2 in limits:
    txt(s, 0.94, yy, 5.15, 0.75, [
        {"runs": [("▲  ", 11, True, ORANGE), (t1, 13, True, INK)], "space_after": 2},
        P(t2, 11, False, MUTED, after=0, line=1.12)])
    yy += 0.92
shp(s, MSO_SHAPE.ROUNDED_RECTANGLE, 6.52, 3.42, 0.42, 0.72, fill=TEAL, adj=0.18)
txt(s, 6.52, 3.60, 0.42, 0.4, [P("›", 20, True, WHITE, align=PP_ALIGN.CENTER, after=0)])
card(s, 7.12, 1.58, 5.59, 4.55, fill=TEAL_SOFT, line=TEAL_SOFT)
txt(s, 7.44, 1.80, 5.0, 0.35, [P("THE ROAD AHEAD", 10.5, True, TEAL_D, after=0, font=MONO)])
road = [("Semantic intelligence", "An NLP/ML layer beside the deterministic core for context-aware ambiguity."),
        ("Measured accuracy", "A labelled benchmark corpus to report precision/recall and calibrate severity."),
        ("Richer documents", "OCR for scanned PDFs, plus table and figure-aware extraction."),
        ("Collaboration", "Team workspaces, shared reviews, comments and approval states."),
        ("In the workflow", "CI checks and editor/API integrations; multilingual requirement sets.")]
yy = 2.30
for idx, (t1, t2) in enumerate(road, 1):
    shp(s, MSO_SHAPE.OVAL, 7.44, yy + 0.02, 0.34, 0.34, fill=TEAL)
    txt(s, 7.44, yy + 0.075, 0.34, 0.25, [P(str(idx), 9.5, True, WHITE, align=PP_ALIGN.CENTER, after=0, font=MONO)])
    txt(s, 7.92, yy, 4.55, 0.62, [P(t1, 12.5, True, INK, after=1), P(t2, 10.3, False, MUTED, after=0, line=1.08)])
    yy += 0.74
txt(s, 0.62, 6.34, 12.09, 0.5, [
    {"runs": [("Design principle:  ", 12.5, True, TEAL_D), ("keep the deterministic core as the trustworthy baseline — add intelligence around it, never instead of it.", 12.5, False, MUTED)], "space_after": 0}])
transition(s, "wipe")
notes(s, "I want to be explicit about limits: the engine is heuristic, so a flagged phrase can occasionally "
         "be acceptable in context; it is English-first; there is no OCR. The roadmap follows one "
         "principle — add a semantic layer and measured accuracy around the deterministic core, plus OCR, "
         "collaboration and CI integration. Professors usually ask about false positives here; the honest "
         "answer is that the tool assists the analyst and explains itself, rather than pretending to "
         "replace judgement. [~55 seconds]")

# ============================================================= SLIDE 10 =====
s = slide_new(NAVY)
shp(s, MSO_SHAPE.OVAL, 10.0, -2.3, 5.8, 5.8, fill=NAVY2)
shp(s, MSO_SHAPE.OVAL, -1.5, 5.4, 3.8, 3.8, fill=NAVY2)
txt(s, 0.72, 0.62, 8.5, 0.3, [P("09 • CONCLUSION", 10.5, True, RGBColor(0x6E, 0xC8, 0xBE), after=0, font=MONO)])
txt(s, 0.72, 1.02, 8.6, 1.5, [
    P("From ambiguous text to an", 31, True, WHITE, after=0, line=1.02, font=HEAD),
    P("actionable quality signal.", 31, True, RGBColor(0x6E, 0xC8, 0xBE), after=0, line=1.02, font=HEAD)])
takeaways = [("DETECT EARLY", "Ambiguity is flagged at the document stage — where fixing it is cheapest."),
             ("EXPLAIN CLEARLY", "Every finding carries its phrase, severity, reason and recommendation."),
             ("IMPROVE MEASURABLY", "Scores and health dimensions make each revision visibly better.")]
x = 0.72
for t1, t2 in takeaways:
    shp(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, 2.78, 2.62, 1.62, fill=NAVY2, adj=0.07)
    shp(s, MSO_SHAPE.RECTANGLE, x, 2.78, 2.62, 0.055, fill=AMBER)
    txt(s, x + 0.20, 3.00, 2.25, 1.2, [P(t1, 11.5, True, WHITE, after=5, font=MONO), P(t2, 10.5, False, CREAM_DIM, after=0, line=1.16)])
    x += 2.78
txt(s, 0.72, 4.78, 8.2, 0.75, [P("Thank you.", 30, True, WHITE, after=2, font=HEAD),
                              P("I welcome your questions — and a live demonstration, if time permits.", 13.5, False, CREAM_DIM, after=0)])
card(s, 0.72, 5.78, 8.05, 1.02, fill=WHITE, line=WHITE)
s.shapes.add_picture(str(ROOT / "synopsis-src/assets/kit-logo.png"), Inches(0.92), Inches(6.00), Inches(1.42), Inches(0.46))
txt(s, 2.55, 5.94, 6.0, 0.75, [
    {"runs": [("Ashish Kumar", 13.5, True, INK), ("  •  Roll No. 2301650100054  •  B.Tech CSE (Section A)", 10.5, False, MUTED)], "space_after": 3},
    {"runs": [("Guide: Mr. Suraj Srivastava", 10.5, True, TEAL_D), ("  •  Kanpur Institute of Technology, Kanpur (AKTU)", 10.5, False, MUTED)], "space_after": 0}])
back = shp(s, MSO_SHAPE.ROUNDED_RECTANGLE, 9.25, 1.55, 3.55, 2.85, fill=WHITE, adj=0.035)
back.rotation = 1.8
pic = s.shapes.add_picture(str(SHOTS / "08-analysis-result--1440x900.png"), Inches(9.36), Inches(1.66), Inches(3.33), Inches(2.63))
pic.rotation = 1.8; pic.shadow.inherit = False
chip(s, 9.55, 4.62, 3.05, 0.44, "DETERMINISTIC  •  EXPLAINABLE  •  TESTED", TEAL, WHITE, 8.5)
txt(s, 9.25, 5.35, 3.55, 0.9, [P("PROJECT", 9, True, RGBColor(0x6E, 0xC8, 0xBE), after=2, font=MONO),
                              P("SRS Ambiguity Detector\nMini Project • 2026–27", 11, False, CREAM_DIM, after=0, line=1.15)])
txt(s, 9.25, 6.55, 3.55, 0.3, [P("QUESTIONS & DISCUSSION", 9.5, True, AMBER, after=0, font=MONO)])
transition(s, "fade", 800)
notes(s, "To conclude: the SRS Ambiguity Detector turns ambiguous requirement text into an actionable, "
         "measurable quality signal. It detects early, explains clearly, and lets authors improve draft "
         "by draft. It is a complete, tested, security-conscious web application. Thank you — I welcome "
         "your questions, and I can run the live demonstration immediately if the panel wishes. [~35 seconds]")

# ---------------------------------------------------------------- finish ----
prs.core_properties.title = "SRS Ambiguity Detector — Mini Project Presentation"
prs.core_properties.author = "Ashish Kumar"
prs.core_properties.subject = "B.Tech CSE Mini Project, Kanpur Institute of Technology, Kanpur"
prs.core_properties.keywords = "SRS, ambiguity detection, requirements engineering, FastAPI, Next.js, PostgreSQL"
prs.save(OUT)
print(f"saved {OUT} slides={len(prs.slides.__iter__.__self__._sldIdLst)}")
