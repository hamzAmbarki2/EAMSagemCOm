"""Builds presentation/technical_defense_final.pptx.

Visual language follows the reference deck (white pages, thin black rule,
red numbered kicker, navy heavy titles, rounded outlined boxes). Every
diagram is drawn from native shapes so it stays editable and imports into
Canva. Architecture facts come from docker-compose*.yml, Caddyfile,
app/frontend/nginx.conf, infra/helm/eam and infra/terraform.

    python presentation/build_final_deck.py
"""
import os
from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
IMG = os.environ.get("EAM_IMG_DIR", os.path.join(HERE, "deck_assets"))
OUT = os.path.join(HERE, "technical_defense_final.pptx")

NAVY, RED, TEAL, BROWN = "0A1045", "D0101A", "0B8CA8", "5B5243"
PINK, LIME, GREEN, GREY = "FDE8E8", "D6FF7F", "8CC63F", "7A7A7A"
BLUE, GREENL, ORANGE, PURPLE = "2FA4E7", "4FBF5F", "F39C2C", "7C4DDB"
YELLOW, BLACK, WHITE, LIGHT = "F5B82E", "000000", "FFFFFF", "F4F4F4"
TITLE_FONT, BODY_FONT = "Arial Black", "Arial"

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
BLANK = prs.slide_layouts[6]
_page = [0]


def rgb(h):
    return RGBColor.from_string(h)


def _fmt_run(run, size, bold, color, font, italic):
    run.font.size, run.font.bold, run.font.italic = Pt(size), bold, italic
    run.font.color.rgb = rgb(color)
    run.font.name = font


def txt(sl, x, y, w, h, paras, size=14, bold=False, color=BLACK, align="l",
        font=BODY_FONT, anchor="m", italic=False, spc=None, fill=None, margin=0.05):
    """Text box. paras: str | list of (str | [(text, {opts}), ...])."""
    tb = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(margin)
    tf.margin_top = tf.margin_bottom = Inches(0.02)
    tf.vertical_anchor = {"m": MSO_ANCHOR.MIDDLE, "t": MSO_ANCHOR.TOP, "b": MSO_ANCHOR.BOTTOM}[anchor]
    if fill:
        tb.fill.solid()
        tb.fill.fore_color.rgb = rgb(fill)
    if isinstance(paras, str):
        paras = paras.split("\n")
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
        runs = [(p, {})] if isinstance(p, str) else p
        for t, o in runs:
            r = para.add_run()
            r.text = t
            _fmt_run(r, o.get("size", size), o.get("bold", bold), o.get("color", color),
                     o.get("font", font), o.get("italic", italic))
            if spc:
                r._r.get_or_add_rPr().set("spc", str(spc))
    return tb


def box(sl, x, y, w, h, line=NAVY, fill=None, lw=1.75, radius=0.12, dash=False, shape=None):
    s = sl.shapes.add_shape(shape or MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    if shape is None:
        s.adjustments[0] = radius
    if fill:
        s.fill.solid()
        s.fill.fore_color.rgb = rgb(fill)
    else:
        s.fill.background()
    if line:
        s.line.color.rgb = rgb(line)
        s.line.width = Pt(lw)
        if dash:
            ln = s.line._get_or_add_ln()
            d = etree.SubElement(ln, qn("a:prstDash"))
            d.set("val", "dash")
    else:
        s.line.fill.background()
    s.shadow.inherit = False
    return s


def card(sl, x, y, w, h, title, body=None, line=NAVY, fill=None, tsize=12, bsize=10,
         tcolor=None, dash=False, anchor="m", bcolor="333333"):
    """Rounded box with a bold title and optional body lines."""
    box(sl, x, y, w, h, line=line, fill=fill, dash=dash)
    paras = [[(title, {"bold": True, "size": tsize, "color": tcolor or line})]]
    for b in (body or "").split("\n") if body else []:
        paras.append([(b, {"size": bsize, "color": bcolor})])
    txt(sl, x + 0.04, y + 0.02, w - 0.08, h - 0.04, paras, align="c", anchor=anchor)


def line(sl, x1, y1, x2, y2, color=BLACK, lw=1.0):
    c = sl.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = rgb(color)
    c.line.width = Pt(lw)
    etree.SubElement(c._element.spPr, qn("a:effectLst"))
    return c


def arrow(sl, x1, y1, x2, y2, color=GREY, lw=1.75, dash=False, both=False):
    c = line(sl, x1, y1, x2, y2, color, lw)
    ln = c.line._get_or_add_ln()
    if dash:
        d = etree.SubElement(ln, qn("a:prstDash"))
        d.set("val", "dash")
    if both:
        etree.SubElement(ln, qn("a:headEnd")).set("type", "triangle")
    etree.SubElement(ln, qn("a:tailEnd")).set("type", "triangle")
    return c


def tag(sl, x, y, w, text, size=9, color=GREY, h=0.22):
    return txt(sl, x, y, w, h, text, size=size, color=color, align="c", fill=WHITE, italic=True)


def page_no(sl):
    _page[0] += 1
    txt(sl, 12.3, 7.05, 0.7, 0.3, str(_page[0]), size=12, align="r")


def base(kicker, title, tsize=30):
    sl = prs.slides.add_slide(BLANK)
    txt(sl, 0.5, 0.22, 9, 0.4, kicker, size=14, bold=True, color=RED, spc=200)
    line(sl, 0.5, 0.75, 12.83, 0.75, BLACK, 1.0)
    txt(sl, 0.5, 0.82, 12.33, 0.75, title, size=tsize, font=TITLE_FONT, color=NAVY, align="c")
    page_no(sl)
    return sl


def divider(num, title):
    sl = prs.slides.add_slide(BLANK)
    line(sl, 4.0, 0, 4.0, 7.5, BLACK, 1.25)
    txt(sl, 4.45, 1.6, 8.4, 4.4, [num] + title.split("\n"), size=48, font=TITLE_FONT,
        color=BLACK, align="l", anchor="m")
    page_no(sl)
    return sl


def notes(sl, text):
    sl.notes_slide.notes_text_frame.text = text


def pic(sl, name, x, y, w=None, h=None):
    path = os.path.join(IMG, name)
    kw = {}
    if w:
        kw["width"] = Inches(w)
    if h:
        kw["height"] = Inches(h)
    return sl.shapes.add_picture(path, Inches(x), Inches(y), **kw)


def mark(sl, x, y, kind, size=22):
    sym, col = {"ok": ("✓", GREEN), "no": ("✗", "C8203A"), "warn": ("⚠", YELLOW)}[kind]
    txt(sl, x, y, 0.5, 0.45, sym, size=size, bold=True, color=col, align="c", font="DejaVu Sans")


def table(sl, x, y, w, col_w, rows, header_fill=NAVY, size=11, row_h=0.4, first_col_bold=True):
    shp = sl.shapes.add_table(len(rows), len(rows[0]), Inches(x), Inches(y), Inches(w), Inches(row_h * len(rows)))
    tbl = shp.table
    tblPr = tbl._tbl.tblPr
    sid = tblPr.find(qn("a:tableStyleId"))
    if sid is None:
        sid = etree.SubElement(tblPr, qn("a:tableStyleId"))
    sid.text = "{5940675A-B579-460E-94D1-54222C63F5DA}"  # No Style, Table Grid
    for i, cw in enumerate(col_w):
        tbl.columns[i].width = Inches(cw)
    for r, row in enumerate(rows):
        tbl.rows[r].height = Inches(row_h)
        for c, val in enumerate(row):
            cell = tbl.cell(r, c)
            cell.margin_left = cell.margin_right = Inches(0.06)
            cell.margin_top = cell.margin_bottom = Inches(0.02)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            colr, sym = BLACK, val
            if val in ("✓", "✗", "⚠"):
                colr = {"✓": GREEN, "✗": "C8203A", "⚠": "D9A000"}[val]
                p.alignment = PP_ALIGN.CENTER
            if r == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = rgb(header_fill)
                colr = WHITE
                p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = sym
            sym_cell = val in ("✓", "✗", "⚠")
            _fmt_run(run, 16 if sym_cell else size, r == 0 or sym_cell or (c == 0 and first_col_bold),
                     colr, "DejaVu Sans" if sym_cell else BODY_FONT, False)
    return tbl


# ───────────────────────── 1 · Title ─────────────────────────
sl = prs.slides.add_slide(BLANK)
pic(sl, "esprit.jpg", 0.55, 0.3, w=2.6)
pic(sl, "sagemcom.png", 11.0, 0.15, w=1.7)
line(sl, 6.67, 0, 6.67, 1.2, BLACK, 1.25)
line(sl, 6.67, 6.6, 6.67, 7.5, BLACK, 1.25)
txt(sl, 4, 1.45, 5.33, 0.4, "2025 - 2026", size=14, bold=True, align="c", spc=300)
txt(sl, 1.2, 2.0, 10.9, 2.2, "EAM Platform for Intelligent\nPredictive Maintenance", size=40,
    font="Georgia", bold=True, italic=True, color=NAVY, align="c")
txt(sl, 1.2, 4.2, 10.9, 0.8, "Operational management, seven ML models and a RAG assistant", size=24,
    font=TITLE_FONT, color=RED, align="c")
txt(sl, 1.2, 5.3, 10.9, 1.3, [
    [("Host company: ", {}), ("SagemCom Ezzahra", {"bold": True})],
    [("Realised by: ", {}), ("Hamza Mbarki", {"bold": True})],
], size=16, align="c")
page_no(sl)

# ───────────────────────── 2 · Summary ─────────────────────────
sl = prs.slides.add_slide(BLANK)
line(sl, 1.3, 0, 1.3, 7.5, BLACK, 1.25)
txt(sl, 2.6, 0.9, 8, 1.2, "Summary", size=54, font=TITLE_FONT, align="l")
items = ["General Context", "Project Study", "Requirements & Analysis",
         "Design & Implementation", "Demo", "Conclusion & Perspectives"]
for i, it in enumerate(items):
    col, row = divmod(i, 3)
    x = 2.4 + col * 5.5
    y = 2.7 + row * 1.35
    txt(sl, x, y, 1.0, 0.8, f"0{i + 1}", size=34, font="Georgia", align="l")
    txt(sl, x + 1.05, y, 4.2, 0.8, it, size=19, align="l")
page_no(sl)

# ═════════════════ 01 · GENERAL CONTEXT ═════════════════
divider("01", "General\nContext")

sl = base("01. General Context", "Host Company", 32)
pic(sl, "sagemcom.png", 0.7, 1.95, w=2.4)
for i, t in enumerate([
    [("Industrial group", {"bold": True}), (" designing and manufacturing high-value communication equipment", {})],
    [("Products: ", {"bold": True}), ("set-top boxes, internet boxes, smart meters", {})],
    [("Host department: ", {"bold": True}), ("Software Engineering & Innovation, in-house digital tools for equipment maintenance", {})],
]):
    box(sl, 4.0, 1.75 + i * 1.05, 8.6, 0.85, line=RED, lw=2, shape=MSO_SHAPE.HEXAGON)
    txt(sl, 4.6, 1.75 + i * 1.05, 7.4, 0.85, [t], size=15, align="c")
txt(sl, 0.7, 5.0, 8, 0.4, "Production lines served by the platform", size=14, bold=True, color=RED, align="l")
for i, (n, d) in enumerate([("SMT assembly", "surface-mount lines"), ("Functional testing", "board & unit tests"),
                            ("Automated quality control", "inspection stations")]):
    x = 0.9 + i * 3.9
    card(sl, x, 5.45, 3.1, 1.0, n, d, line=RED, fill=RED, tcolor=WHITE, tsize=15, bsize=11, bcolor=WHITE)
    if i < 2:
        arrow(sl, x + 3.15, 5.95, x + 3.85, 5.95, BLACK, 1.5)
for i, t in enumerate(["7 months", "7 ML models (P1-P7)", "4 user roles"]):
    mark(sl, 2.2 + i * 3.4, 6.55, "ok", 18)
    txt(sl, 2.65 + i * 3.4, 6.55, 2.8, 0.45, t, size=13, italic=True, align="l")

sl = base("01. General Context", "Background & Motivation", 32)
cols = [("Reactive", "Repair after the failure", "Line stoppage, emergency parts ordering, maximum total cost.", "C8203A"),
        ("Preventive", "Repair on a calendar", "Fixed intervals regardless of real condition: parts replaced too early, or failures between two visits.", ORANGE),
        ("Predictive", "Repair before the failure", "Continuous reading of the machine's own signals (temperature, torque, wear) to anticipate it.", GREEN)]
for i, (h, s, d, c) in enumerate(cols):
    x = 0.75 + i * 4.15
    txt(sl, x, 1.75, 3.7, 0.9, str(i + 1), size=40, font=TITLE_FONT, color=RED, align="c")
    line(sl, x, 2.75, x + 0.5, 2.75, RED, 3)
    txt(sl, x + 0.55, 2.55, 2.6, 0.4, h + ":", size=18, bold=True, align="c")
    line(sl, x + 3.15, 2.75, x + 3.7, 2.75, RED, 3)
    txt(sl, x, 3.1, 3.7, 0.5, s, size=17, bold=True, italic=True, color="1A1A9C", align="c")
    txt(sl, x + 0.1, 3.7, 3.5, 1.3, d, size=14, align="c", anchor="t")
    box(sl, x + 0.3, 5.15, 3.1, 0.12, line=None, fill=c, radius=0.5)
box(sl, 1.3, 5.7, 10.7, 1.05, line=RED, lw=2)
txt(sl, 1.4, 5.7, 10.5, 1.05, [[("In most cases a machine's degradation is visible in its sensor data well before the stoppage. ", {}),
    ("That is the premise the whole ML pipeline is built on.", {"bold": True, "color": RED})]], size=17, align="c")

# ═════════════════ 02 · PROJECT STUDY ═════════════════
divider("02", "Project\nStudy")

sl = base("02. Project Study", "Problem Statement", 32)
probs = ["Operations and prediction live in separate tools", "Failures are found after the stoppage",
         "Predictions technicians cannot understand", "Documentation is disconnected from live machine status"]
for i, p in enumerate(probs):
    col, row = divmod(i, 2)
    x, y = 1.2 + col * 5.9, 2.0 + row * 1.5
    txt(sl, x, y, 0.9, 0.9, "X", size=44, font=TITLE_FONT, color="C8203A", align="c")
    txt(sl, x + 0.95, y, 4.6, 0.9, p, size=19, align="l")
txt(sl, 1.0, 5.2, 11.3, 1.6, [[("How can one platform cover classic operational management ", {}),
    ("and", {"color": RED}), (" predictive maintenance, with predictions ", {}),
    ("people trust", {"color": RED}), (" and ", {}), ("automation they control", {"color": RED}), ("?", {})]],
    size=26, font=TITLE_FONT, color=NAVY, align="c")

sl = base("02. Project Study", "Study of the Existant", 32)
exist = [("ERP / EAM suites", "SAP PM · IBM Maximo", "Proven EAM coverage of assets and work orders", "Predictive ML is a costly add-on, heavy licensing"),
         ("Industrial IoT / APM", "GE Digital Predix · Maximo APM", "Sensor analytics and reliability at scale", "No assistant tied to live machine status, integration cost"),
         ("CMMS & spreadsheets", "Generic CMMS · Excel", "Cheap and lightweight", "No prediction, no explanation, no traceability")]
for i, (h, ex, good, bad) in enumerate(exist):
    x = 0.45 + i * 4.2
    box(sl, x, 2.35, 3.9, 4.45, line=RED, lw=2.5, radius=0.1)
    box(sl, x + 0.35, 1.95, 3.2, 0.8, line=None, fill=PINK, radius=0.2)
    txt(sl, x + 0.35, 1.95, 3.2, 0.8, h, size=16, bold=True, align="c")
    mark(sl, x + 0.2, 3.05, "ok", 30)
    txt(sl, x + 0.8, 2.95, 2.95, 0.95, good, size=14, align="l")
    mark(sl, x + 0.2, 4.2, "no", 30)
    txt(sl, x + 0.8, 4.1, 2.95, 1.1, bad, size=14, align="l")
    txt(sl, x + 0.2, 5.6, 3.5, 0.9, ex, size=14, bold=True, italic=True, color=NAVY, align="c")

sl = base("02. Project Study", "Study of the Existant", 32)
rows = [["Criteria", "EAM suites", "IoT / APM", "CMMS / Excel", "Our solution"],
        ["Full EAM coverage", "✓", "⚠", "⚠", "✓"],
        ["Predictive ML built in", "⚠", "✓", "✗", "✓"],
        ["Predictions explained in plain language", "✗", "⚠", "✗", "✓"],
        ["Assistant tied to live machine status", "✗", "✗", "✗", "✓"],
        ["No heavy licensing", "✗", "✗", "✓", "✓"],
        ["Automation guarded by human approval", "⚠", "⚠", "✗", "✓"]]
table(sl, 0.9, 1.85, 11.5, [4.7, 1.7, 1.7, 1.7, 1.7], rows, row_h=0.62, size=15)
txt(sl, 0.9, 6.35, 11.5, 0.4, "Qualitative assessment from public product positioning; ✓ yes · ⚠ partial or add-on · ✗ no",
    size=11, italic=True, color=GREY, align="l")

sl = base("02. Project Study", "Our Proposed Solution", 32)
sol = [("Full EAM foundation", "Users, machines, work orders, intervention orders, scheduling, archiving"),
       ("7 predictive models (P1-P7)", "Failure, failure type, remaining life, anomaly, priority, schedule, parts"),
       ("Evidence fusion (DST)", "4 advanced signals combined into one honest, unified health score"),
       ("Systematic explainability", "Every prediction ships with a plain-language explanation"),
       ("Conversational RAG assistant", "Documentation + real-time machine status merged into every answer"),
       ("Guarded automation", "Parts proposals generated, never ordered without human approval")]
for i, (h, d) in enumerate(sol):
    col, row = divmod(i, 3)
    x, y = 0.7 + col * 6.1, 1.85 + row * 1.7
    mark(sl, x, y + 0.45, "ok", 34)
    box(sl, x + 0.65, y, 5.2, 1.45, line=None, fill=LIME, radius=0.12)
    txt(sl, x + 0.8, y + 0.05, 4.95, 1.35, [[(h, {"bold": True, "size": 17})], [(d, {"size": 13})]], align="l")

sl = base("02. Project Study", "Working Methodology", 32)
for i, (r, d) in enumerate([("Product Owner", "SagemCom supervisor: prioritises business needs and validates each deliverable"),
                            ("Scrum Master", "Intern: drives the specify → plan → build → verify cycle"),
                            ("Development team", "Intern, paired with an AI assistant for implementation and testing")]):
    card(sl, 0.7, 1.85 + i * 1.35, 4.9, 1.15, r, d, line=TEAL, fill=None, tsize=15, bsize=12)
for i, s in enumerate(["Specify", "Plan", "Build", "Verify"]):
    x = 6.1 + (i % 2) * 3.4
    y = 2.1 + (i // 2) * 1.45
    box(sl, x, y, 2.9, 1.05, line=None, fill="C5CBFA", radius=0.3)
    txt(sl, x, y, 2.9, 1.05, [[(f"{i + 1}  ", {"color": RED, "bold": True}), (s, {"bold": True})]], size=20, align="c")
arrow(sl, 9.05, 2.62, 9.45, 2.62, NAVY)
arrow(sl, 7.55, 3.2, 7.55, 3.5, NAVY)
arrow(sl, 9.45, 4.07, 9.05, 4.07, NAVY)
txt(sl, 6.1, 5.0, 6.3, 1.7, [[("7 monthly sprints, ", {"bold": True}), ("aligned with the roadmap.", {})],
    [("Every feature has a written specification and an implementation plan ", {}),
     ("before any code is written", {"bold": True, "color": RED}), (": full traceability of technical decisions.", {})]],
    size=15, align="l", anchor="t")

# ═════════════════ 03 · REQUIREMENTS & ANALYSIS ═════════════════
divider("03", "Requirements\n& Analysis")

sl = base("03. Requirements & Analysis", "The Four Platform Roles", 32)
roles = [("AD", "Administrator", "Infrastructure & data", "Manages users, machines and the document base; overall view of the system."),
         ("OP", "Operations Manager", "Oversight", "Consults machine labels and statuses at a glance."),
         ("HE", "Head Technician", "Planning & approval", "Plans interventions, arbitrates priorities, approves procurement proposals."),
         ("TE", "Technician", "Field execution", "Executes intervention orders, checks health status, queries the AI assistant.")]
for i, (ab, n, sub, d) in enumerate(roles):
    col, row = divmod(i, 2)
    x, y = 0.6 + col * 6.2, 1.8 + row * 2.5
    circ = box(sl, x, y + 0.25, 1.3, 1.3, line=NAVY, fill=PINK, shape=MSO_SHAPE.OVAL)
    txt(sl, x, y + 0.25, 1.3, 1.3, ab, size=26, font=TITLE_FONT, color=NAVY, align="c")
    txt(sl, x + 1.5, y, 4.4, 0.45, n, size=19, bold=True, align="l")
    txt(sl, x + 1.5, y + 0.45, 4.4, 0.4, sub, size=14, bold=True, italic=True, color=RED, align="l")
    box(sl, x + 1.5, y + 0.95, 4.4, 1.2, line=NAVY, lw=2)
    txt(sl, x + 1.6, y + 0.95, 4.2, 1.2, d, size=13, align="l")

sl = base("03. Requirements & Analysis", "Use Case Diagram", 32)
pic(sl, "usecase.png", 4.4, 1.6, h=5.4)

sl = base("03. Requirements & Analysis", "Functional Requirements: 9 Areas", 30)
fr = ["Authentication & authorization", "Machine management", "Order management", "Scheduling",
      "ML predictions (P1-P7)", "Alerts & notifications", "Conversational assistant", "Dashboards",
      "Archiving & traceability"]
for i, f in enumerate(fr):
    col, row = divmod(i, 3)
    x, y = 0.8 + col * 4.1, 1.8 + row * 1.2
    box(sl, x, y, 3.8, 0.95, line=NAVY, lw=2)
    txt(sl, x + 0.1, y, 0.7, 0.95, f"{i + 1}", size=24, font=TITLE_FONT, color=RED, align="c")
    txt(sl, x + 0.8, y, 2.95, 0.95, f, size=14, bold=True, align="l")
box(sl, 0.8, 5.55, 11.9, 1.2, line=RED, lw=2)
txt(sl, 0.95, 5.55, 11.6, 1.2, [[("Two rules cut across all nine: ", {"bold": True, "color": RED}),
    ("every prediction explains itself in plain language, and the system degrades honestly, reporting missing telemetry instead of inventing a value.", {})]],
    size=15, align="c")

sl = base("03. Requirements & Analysis", "Non-Functional Requirements", 32)
nfr = [("Performance", "< 2 s common operations; < 5 s full ML inference"),
       ("Availability", "EAM features keep working if the ML service is down"),
       ("Maintainability", "Layered separation; versioned Alembic migrations"),
       ("Portability", "Docker Compose / Helm; configuration in one .env"),
       ("Usability", "Interface adapted to role and device (desktop or tablet)"),
       ("Integrity & confidentiality", "Role-scoped data; intervention history immutable"),
       ("Security", "OWASP Top 10 hardened; scanned on every push")]
for i, (h, d) in enumerate(nfr):
    row, col = divmod(i, 2)
    x, y = 0.55 + col * 6.3, 1.75 + row * 1.25
    box(sl, x, y, 5.9, 1.0, line=BROWN, lw=2.25)
    txt(sl, x + 0.1, y + 0.05, 5.7, 0.4, h, size=16, bold=True, align="c")
    txt(sl, x + 0.1, y + 0.45, 5.7, 0.5, d, size=12, align="c")

sl = base("03. Requirements & Analysis", "Used Tools & Technologies", 32)
groups = [("Frontend", "React 19 · TypeScript\nVite · Tailwind · shadcn/ui\nnginx (static + /api proxy)"),
          ("Backend", "FastAPI · SQLAlchemy\nAlembic · Celery\nJWT role-based access"),
          ("ML microservice", "FastAPI · XGBoost\nscikit-learn · Cox PH\nDempster-Shafer fusion"),
          ("RAG & AI", "pgvector · BAAI/bge-m3\nbge-reranker-v2-m3\nGroq LLM"),
          ("Data & messaging", "PostgreSQL 15 + pgvector\nMinIO (S3) · RabbitMQ"),
          ("Infrastructure", "Docker Compose · Caddy\nTerraform · AKS · Helm\nKey Vault · ACR"),
          ("DevSecOps", "GitLab CI · Gitleaks · Semgrep\nSonarQube · Trivy · OWASP ZAP\nPrometheus · Grafana")]
pos = [(0.5, 1.85), (4.6, 1.85), (8.7, 1.85), (0.5, 4.2), (4.6, 4.2), (8.7, 4.2)]
for (h, d), (x, y) in zip(groups[:6], pos):
    box(sl, x, y + 0.25, 4.0, 1.95, line=BROWN, lw=2)
    box(sl, x + 0.5, y, 3.0, 0.5, line=None, fill=PINK, radius=0.4)
    txt(sl, x + 0.5, y, 3.0, 0.5, h, size=14, bold=True, align="c", spc=50)
    txt(sl, x + 0.15, y + 0.6, 3.7, 1.55, d.split("\n"), size=13, align="c")
h, d = groups[6]
box(sl, 4.6, 6.45, 4.0, 0.001, line=None)
sl.shapes[-1]._element.getparent().remove(sl.shapes[-1]._element)
txt(sl, 0.5, 6.65, 12.3, 0.4, [[("DevSecOps: ", {"bold": True, "color": RED}), (d.replace("\n", "  ·  "), {})]],
    size=12, align="c")

# ═════════════════ 04 · DESIGN & IMPLEMENTATION ═════════════════
divider("04", "Design &\nImplementation")

# ---------- Logical architecture (redrawn: layers, dependencies, no hosts/ports) ----------
sl = base("04. Design: Logical Architecture", "Logical Architecture: Layers & Dependencies", 26)


def layer(y, h, label, color, x=1.0, w=8.95):
    box(sl, x, y, w, h, line=color, lw=2.0, dash=True, radius=0.07)
    txt(sl, x + 0.1, y + 0.01, 4, 0.24, label, size=9, bold=True, color=color, align="l", spc=100)


def mod(x, y, w, h, t, color, sub=None):
    card(sl, x, y, w, h, t, sub, line=color, tsize=10.5, bsize=8.5)


box(sl, 1.85, 1.5, 7.2, 0.42, line=BLUE, lw=2, radius=0.3)
txt(sl, 1.85, 1.5, 7.2, 0.42, "Users: Administrator · Operations Manager · Head Technician · Technician", size=11.5, bold=True, color=BLUE, align="c")
arrow(sl, 5.45, 1.92, 5.45, 2.2, BLUE)
tag(sl, 5.55, 1.97, 1.9, "HTTPS · browser")

layer(2.2, 0.7, "PRESENTATION", BLUE)
txt(sl, 1.15, 2.42, 8.7, 0.45, [[("React 19 + TypeScript single-page app", {"bold": True}),
    (" · role-based spaces (Admin, Head Technician, Technician) · ML Intelligence panel · AI assistant", {})]], size=11, align="c")
arrow(sl, 5.45, 2.9, 5.45, 3.15, GREENL)
tag(sl, 5.55, 2.93, 2.2, "REST / JSON + JWT")

layer(3.15, 1.45, "APPLICATION: BACKEND API (FastAPI)", GREENL)
m1 = ["Auth & RBAC", "Machines & assets", "Work & intervention orders", "Planning & scheduling"]
m2 = ["Inventory & procurement", "Alerts & notifications", "ML gateway (unified health)", "AI chat orchestrator"]
for i, t in enumerate(m1):
    mod(1.15 + i * 2.2, 3.45, 2.1, 0.46, t, GREENL)
for i, t in enumerate(m2):
    mod(1.15 + i * 2.2, 3.98, 2.1, 0.46, t, GREENL)

layer(4.95, 0.9, "SPECIALIST SERVICES", ORANGE)
mod(1.15, 5.22, 2.85, 0.56, "ML microservice (FastAPI)", ORANGE, "P1-P7 · Wave 2 signals · DST fusion")
mod(4.15, 5.22, 2.85, 0.56, "RAG service (FastAPI)", ORANGE, "ingest · hybrid retrieval · rerank · bge-m3")
mod(7.15, 5.22, 2.7, 0.56, "Celery worker + beat", ORANGE, "consumes tasks · uses PostgreSQL")
arrow(sl, 2.5, 4.6, 2.5, 5.22, ORANGE)
tag(sl, 2.58, 4.7, 1.3, "HTTP /predict")
arrow(sl, 5.55, 4.6, 5.55, 5.22, ORANGE)
tag(sl, 5.63, 4.7, 1.45, "HTTP /retrieve")

layer(6.15, 0.92, "DATA", PURPLE)
mod(1.15, 6.42, 1.85, 0.58, "Model artifacts", PURPLE, ".pkl, read-only")
mod(3.15, 6.42, 2.7, 0.58, "PostgreSQL + pgvector", PURPLE, "business data · ML logs · chunks")
mod(6.0, 6.42, 1.75, 0.58, "Object storage", PURPLE, "MinIO / S3")
mod(7.9, 6.42, 1.95, 0.58, "Message broker", PURPLE, "RabbitMQ")
arrow(sl, 2.0, 5.78, 2.0, 6.42, PURPLE)
arrow(sl, 5.0, 5.78, 5.0, 6.42, PURPLE)
tag(sl, 5.08, 5.86, 1.1, "SQL (vectors)")
arrow(sl, 8.9, 5.78, 8.9, 6.42, PURPLE)
# backend -> data bus (left)
line(sl, 1.0, 3.9, 0.6, 3.9, PURPLE, 1.75)
line(sl, 0.6, 3.9, 0.6, 6.6, PURPLE, 1.75)
arrow(sl, 0.6, 6.6, 1.0, 6.6, PURPLE)
txt(sl, 0.02, 4.9, 0.95, 0.6, ["SQL ·", "S3 API ·", "AMQP"], size=8, italic=True, color=PURPLE, align="c", fill=WHITE)

box(sl, 10.45, 3.0, 2.45, 3.0, line=PURPLE, lw=1.75, dash=True, radius=0.07)
txt(sl, 10.5, 3.02, 2.35, 0.26, "EXTERNAL SERVICES", size=9, bold=True, color=PURPLE, align="c", spc=100)
mod(10.6, 3.4, 2.15, 0.6, "Groq LLM API", PURPLE, "chat completions")
mod(10.6, 5.15, 2.15, 0.6, "SMTP server", PURPLE, "e-mail notifications")
arrow(sl, 9.95, 3.7, 10.6, 3.7, PURPLE)
tag(sl, 9.99, 3.46, 0.6, "HTTPS", size=8)
arrow(sl, 9.95, 5.45, 10.6, 5.45, PURPLE)
tag(sl, 9.99, 5.2, 0.6, "SMTP", size=8)
txt(sl, 10.4, 6.15, 2.6, 0.9, "Backend is the only caller of the ML and RAG services. RAG reads PostgreSQL only; ML reads model files only.",
    size=9, italic=True, color=GREY, align="l", anchor="t")
notes(sl, "LOGICAL view: responsibilities and dependencies only. No hosts, ports or containers here (that is the physical view). "
          "The browser runs a React SPA. All traffic goes to the FastAPI backend, which holds the business logic (auth/RBAC, machines, orders, planning, "
          "inventory, alerts, ML gateway, AI chat orchestrator). The backend is the only caller of the two specialist services: the ML microservice "
          "(FastAPI, seven models + Wave 2 signals + Dempster-Shafer fusion, reading model files) and the RAG service (FastAPI, hybrid retrieval, "
          "reranker, bge-m3 embeddings, reading PostgreSQL/pgvector only). Celery workers consume tasks from RabbitMQ for e-mails and scheduled jobs. "
          "Data: PostgreSQL+pgvector, MinIO/S3 for attachments and RAG documents. External: Groq LLM (called by the backend chat) and SMTP (called by Celery).")

# ---------- Request flow ----------
sl = base("04. Design: Logical Architecture", "Request Flow: Unified Machine Health", 28)
steps = [("1", "Frontend asks", "GET /machines/{id}/unified-health"),
         ("2", "Backend reads telemetry", "Latest real sensor data; none found → telemetry_available: false, stop"),
         ("3", "ML microservice", "P1-P7 in parallel on the sensor snapshot"),
         ("4", "Wave 2 signals", "Run concurrently on the telemetry history"),
         ("5", "DST fusion", "One unified score and verdict"),
         ("6", "Log & render", "Stored in ml_prediction_logs; panel updates")]
for i, (n, h, d) in enumerate(steps):
    row, col = divmod(i, 3)
    x, y = 0.55 + col * 4.2, 1.95 + row * 2.0
    box(sl, x, y, 3.8, 1.55, line=TEAL, lw=2)
    box(sl, x + 0.15, y - 0.2, 0.55, 0.5, line=None, fill=RED, radius=0.5)
    txt(sl, x + 0.15, y - 0.2, 0.55, 0.5, n, size=16, bold=True, color=WHITE, align="c")
    txt(sl, x + 0.2, y + 0.25, 3.4, 0.45, h, size=15, bold=True, color=NAVY, align="c")
    txt(sl, x + 0.2, y + 0.7, 3.4, 0.8, d, size=12, align="c", anchor="t")
    if col < 2:
        arrow(sl, x + 3.85, y + 0.78, x + 4.15, y + 0.78, NAVY)
box(sl, 0.55, 6.0, 12.2, 0.8, line=RED, lw=2)
txt(sl, 0.7, 6.0, 11.9, 0.8, [[("Graceful degradation: ", {"bold": True, "color": RED}),
    ("no telemetry on record never means a fabricated prediction; the platform falls back to a maintenance-history-only score.", {})]],
    size=14, align="c")

# ---------- Physical architecture 1: Docker Compose host ----------
sl = base("04. Design: Physical Architecture", "Physical Architecture: Docker Compose Host", 26)
card(sl, 0.4, 3.2, 1.45, 0.8, "Users", "browser", line=BLUE, tsize=13)
arrow(sl, 1.85, 3.6, 2.5, 3.6, BLUE)
tag(sl, 1.88, 3.3, 0.45, "HTTPS", size=8)

box(sl, 2.35, 1.55, 8.85, 5.5, line=BROWN, lw=2, dash=True, radius=0.04)
txt(sl, 2.45, 1.57, 8.6, 0.28, "DOCKER HOST: developer workstation (dev) or Azure VM (demo) · docker compose", size=9, bold=True, color=BROWN, align="l", spc=60)
card(sl, 2.5, 3.0, 1.5, 1.0, "Caddy 2.8", ":80 / :443\nTLS (Let's Encrypt)\nprod overlay only", line=RED, tsize=12, bsize=9)
txt(sl, 2.45, 4.1, 1.7, 0.95, ["/  → frontend:80", "/api/* → backend:8000", "/grafana/* → grafana:3000"], size=8.5, color="333333", align="l")
arrow(sl, 4.0, 3.5, 4.35, 3.5, RED)

box(sl, 4.35, 1.95, 6.75, 3.05, line=BLUE, lw=2, radius=0.05)
txt(sl, 4.45, 1.97, 6.5, 0.26, "asset_management_network (bridge): containers resolve each other by name", size=9, bold=True, color=BLUE, align="l")
cw = 1.58
cols_x = [4.45 + i * (cw + 0.1) for i in range(4)]
row1 = [("frontend", "nginx, React SPA\n:80 (dev → :3000)"), ("backend", "FastAPI / uvicorn\n:8000"),
        ("ml-service", "FastAPI\n:8000 (dev → :8001)"), ("rag-service", "FastAPI + bge-m3\n:8003")]
row2 = [("celery_worker", "backend image\nCelery worker"), ("celery_beat", "backend image\nCelery scheduler"),
        ("rabbitmq", "AMQP :5672\n(mgmt :15672)"), ("minio", "S3 API :9000\nconsole :9001")]
for x, (n, d) in zip(cols_x, row1):
    card(sl, x, 2.28, cw, 0.85, n, d, line=NAVY, tsize=11, bsize=8.5)
for x, (n, d) in zip(cols_x, row2):
    card(sl, x, 3.23, cw, 0.85, n, d, line=NAVY, tsize=11, bsize=8.5)
card(sl, cols_x[0], 4.18, 2 * cw + 0.1, 0.72, "postgres", "pgvector/pg15 :5432 · volume postgres_data", line=PURPLE, tsize=11, bsize=8.5)
card(sl, cols_x[2], 4.18, cw, 0.72, "pgadmin", "dev only, host :5050", line=GREY, tsize=11, bsize=8.5, dash=True)
card(sl, cols_x[3], 4.18, cw, 0.72, "models/*.pkl", "bind mount, read-only", line=PURPLE, tsize=11, bsize=8.5, dash=True)

box(sl, 4.35, 5.15, 3.3, 1.0, line=ORANGE, lw=2, radius=0.06)
txt(sl, 4.45, 5.17, 3.1, 0.26, "monitoring_network (compose -f monitoring)", size=8.5, bold=True, color=ORANGE, align="l")
txt(sl, 4.45, 5.45, 3.1, 0.7, ["Prometheus :9095 · Pushgateway :9091", "Grafana :3001 · Trivy scanner"], size=9.5, align="l")
box(sl, 7.8, 5.15, 3.3, 1.0, line=PURPLE, lw=2, radius=0.06, dash=True)
txt(sl, 7.9, 5.17, 3.1, 0.26, "on-demand stacks (own compose files)", size=8.5, bold=True, color=PURPLE, align="l")
txt(sl, 7.9, 5.45, 3.1, 0.7, ["SonarQube CE :9090 + Postgres 15", "Jupyter notebooks :8888 (offline ML research)"], size=9.5, align="l")
box(sl, 2.5, 6.3, 8.6, 0.65, line=GREEN, lw=1.5, radius=0.1)
txt(sl, 2.6, 6.3, 8.4, 0.65, [[("Prod overlay: ", {"bold": True, "color": GREEN}),
    ("every service un-published except Caddy; images pulled from the container registry (pull_policy: always).", {})]], size=10, align="c")

card(sl, 11.55, 2.2, 1.6, 0.95, "Groq LLM API", "HTTPS 443\nfrom backend", line=PURPLE, tsize=11, bsize=9)
card(sl, 11.55, 3.45, 1.6, 0.95, "SMTP (Gmail)", ":587\nfrom Celery", line=PURPLE, tsize=11, bsize=9)
card(sl, 11.55, 4.7, 1.6, 0.95, "Container registry", "images for\nprod overlay", line=PURPLE, tsize=11, bsize=9)
arrow(sl, 11.2, 2.68, 11.55, 2.68, PURPLE)
arrow(sl, 11.2, 3.93, 11.55, 3.93, PURPLE)
arrow(sl, 11.55, 5.18, 11.2, 5.18, PURPLE, dash=True)
notes(sl, "PHYSICAL view 1/2: one Docker host running docker compose, either a developer workstation or the Azure demo VM. "
          "Application containers share asset_management_network and reach each other by service name (backend calls ml-service:8000 and rag-service:8003). "
          "Host ports in dev: frontend 3000, backend 8000, ml-service 8001 (container port 8000), rag 8003, postgres 5432, rabbitmq 5672/16200, minio 9000/9001, pgadmin 5050. "
          "In the demo VM the prod overlay publishes nothing except Caddy on 80/443, which routes / to the frontend, /api to the backend and /grafana to Grafana. "
          "Monitoring (Prometheus, Pushgateway, Grafana, Trivy) sits on its own network; SonarQube and Jupyter are separate on-demand stacks. "
          "The frontend is a static React/Vite build served by nginx, which also proxies /api to the backend.")

# ---------- Physical architecture 2: AKS ----------
sl = base("04. Design: Physical Architecture", "Physical Architecture: Production on Azure AKS", 26)
card(sl, 0.35, 3.05, 1.35, 0.8, "Users", "browser", line=BLUE, tsize=13)
arrow(sl, 1.7, 3.45, 2.05, 3.45, BLUE)
card(sl, 2.05, 2.85, 1.55, 1.2, "ngrok edge", "public HTTPS\nstatic domain", line=PURPLE, tsize=12, bsize=9.5)

box(sl, 3.85, 1.5, 9.2, 5.55, line=BLUE, lw=2, dash=True, radius=0.03)
txt(sl, 3.95, 1.52, 9, 0.26, "MICROSOFT AZURE: resource group eam-prod-rg · Poland Central", size=9, bold=True, color=BLUE, align="l", spc=60)
box(sl, 4.0, 1.85, 5.25, 4.2, line=TEAL, lw=1.75, radius=0.04)
txt(sl, 4.08, 1.87, 5.1, 0.25, "VNet 10.10.0.0/16 · AKS subnet 10.10.1.0/24", size=8.5, bold=True, color=TEAL, align="l")
box(sl, 4.12, 2.15, 5.0, 3.8, line=GREENL, lw=2, radius=0.04)
txt(sl, 4.2, 2.17, 4.85, 0.25, "AKS cluster · Azure CNI + network policy · OIDC / workload identity", size=8.5, bold=True, color=GREENL, align="l")
box(sl, 4.22, 2.5, 4.8, 2.4, line=NAVY, lw=1.5, radius=0.05)
txt(sl, 4.3, 2.52, 4.65, 0.25, "namespaces eam-prod and eam-staging (same layout)", size=8.5, bold=True, color=NAVY, align="l")
pods = [("ngrok", "tunnel agent"), ("frontend", "nginx :80"), ("backend", "FastAPI :8000"),
        ("ml-service", ":8000"), ("rag-service", ":8003"), ("celery-worker", "backend image"),
        ("celery-beat", "backend image"), ("rabbitmq", ":5672"), ("minio", ":9000 · PVC 5Gi")]
for i, (n, d) in enumerate(pods):
    row, col = divmod(i, 3)
    card(sl, 4.32 + col * 1.58, 2.85 + row * 0.67, 1.48, 0.58, n, d, line=NAVY, tsize=10, bsize=8)
card(sl, 4.22, 5.0, 2.35, 0.8, "System node pool", "critical add-ons only", line=GREENL, tsize=10, bsize=8.5)
card(sl, 6.67, 5.0, 2.35, 0.8, "User node pool", "autoscale 2-4 · Standard_B2s_v2", line=GREENL, tsize=10, bsize=8.5)

card(sl, 9.65, 1.9, 3.3, 1.15, "PostgreSQL Flexible Server 15", "pgvector + pg_trgm · burstable B2s\nDBs: asset_management / _prod", line=PURPLE, tsize=11, bsize=9)
card(sl, 9.65, 3.25, 3.3, 1.05, "Azure Key Vault", "JWT · Groq · SMTP · DB password\nSecrets Store CSI driver", line=PURPLE, tsize=11, bsize=9)
card(sl, 9.65, 4.5, 3.3, 0.95, "Azure Container Registry", "images pulled by kubelet identity (AcrPull)", line=PURPLE, tsize=11, bsize=9)
arrow(sl, 9.12, 2.5, 9.65, 2.5, PURPLE)
tag(sl, 9.18, 2.26, 0.5, "5432", size=8)
arrow(sl, 9.65, 3.78, 9.12, 3.78, PURPLE, dash=True)
tag(sl, 9.1, 3.55, 0.6, "secrets", size=8)
arrow(sl, 9.65, 4.98, 9.12, 4.98, PURPLE, dash=True)
tag(sl, 9.1, 4.75, 0.6, "pull", size=8)

arrow(sl, 3.6, 3.45, 4.32, 3.15, PURPLE, dash=True, both=True)
tag(sl, 3.35, 2.82, 1.45, "tunnel (pod dials out)", size=8)
card(sl, 0.35, 4.7, 3.2, 0.95, "Groq LLM API · SMTP", "HTTPS / SMTP egress via the AKS outbound IP", line=PURPLE, tsize=10.5, bsize=8.5)
arrow(sl, 4.12, 5.15, 3.55, 5.15, PURPLE)
txt(sl, 0.35, 5.85, 3.3, 1.15, "Azure's inbound public Load Balancer is blocked at subscription level, so the ngrok pod dials out to its edge and relays public traffic to frontend:80 (nginx proxies /api to the backend).",
    size=8.5, italic=True, color=GREY, align="l", anchor="t")

for i, (h, d) in enumerate([("GitLab CI", "secret scan · SCA · SAST · SonarQube · build · Trivy · gate"),
                            ("deploy-aks", "docker build → push to ACR → rollout restart"),
                            ("Terraform + Helm", "Azure resources (state on HCP Terraform) · chart eam, values per env")]):
    card(sl, 4.0 + i * 3.0, 6.2, 2.9, 0.75, h, d, line=ORANGE, tsize=10, bsize=8)
notes(sl, "PHYSICAL view 2/2: production target on Azure, provisioned with Terraform. Public users reach the app through an ngrok tunnel because the Azure inbound Load Balancer is blocked at subscription level; "
          "the ngrok pod dials out and relays to the frontend (nginx), which proxies /api to the backend. "
          "AKS (Poland Central, Azure CNI, network policy, workload identity) hosts the eam-staging and eam-prod namespaces: frontend, backend, ml-service, rag-service, celery worker/beat, plus in-cluster RabbitMQ and MinIO. "
          "PostgreSQL is a managed Azure Flexible Server with pgvector; secrets come from Key Vault through the Secrets Store CSI driver; images come from ACR. "
          "Delivery: GitLab CI gates the code, deploy-aks builds and pushes images then restarts deployments, Terraform and Helm define the infrastructure and the chart.")

# ---------- DevSecOps ----------
sl = base("04. Implementation: DevSecOps", "Security & Quality Pipeline", 30)
stages = [("1", "Gitleaks", "Secret scanning\nof git history"), ("2", "SCA", "pip-audit and\npnpm audit"),
          ("3", "SAST", "Semgrep rule sets\n(auto, explicit, custom)"), ("4", "SonarQube", "Static analysis\nand hotspots"),
          ("5", "Build + Trivy", "Docker images\nand IaC scan"), ("6", "Quality gate", "Final blocking\ncheck")]
for i, (n, h, d) in enumerate(stages):
    x = 0.45 + i * 2.1
    card(sl, x, 1.85, 1.95, 1.7, f"{n} · {h}", d, line=TEAL, tsize=12, bsize=10)
    if i < 5:
        arrow(sl, x + 1.95, 2.7, x + 2.1, 2.7, NAVY, 1.5)
box(sl, 0.45, 3.85, 12.4, 0.62, line=RED, lw=2)
txt(sl, 0.6, 3.85, 12.1, 0.62, [[("Blocking: ", {"bold": True, "color": RED}),
    ("stages 1-6.  ", {}), ("Non-blocking: ", {"bold": True}), ("lint.  ", {}),
    ("OWASP ZAP DAST ", {"bold": True}), ("logs in as all 4 roles and probes the running stack; results feed Grafana.", {})]], size=12, align="c")
for i, (v, l) in enumerate([("0", "new violations"), ("100%", "hotspots reviewed"), ("2.6%", "duplication (was 10.3%)"), ("74.8%", "coverage (target 70%)")]):
    x = 0.9 + i * 3.1
    box(sl, x, 4.85, 2.8, 1.65, line=GREEN, lw=2.25)
    txt(sl, x, 4.95, 2.8, 0.9, v, size=34, font=TITLE_FONT, color=GREEN, align="c")
    txt(sl, x, 5.85, 2.8, 0.55, l, size=13, align="c")
txt(sl, 0.45, 6.6, 12.4, 0.4, "Quality gate: fully green on the EAMSagemCom project (branch Phase_2)", size=12, italic=True, color=GREY, align="c")

sl = base("04. Implementation: DevSecOps", "Security Pipeline: Live Evidence", 30)
for i, (f, cap) in enumerate([("sonar.png", "SonarQube quality gate"), ("trivy.png", "Trivy image scan"), ("zap.png", "OWASP ZAP DAST")]):
    x = 0.45 + i * 4.2
    pic(sl, f, x, 2.0, w=4.0)
    txt(sl, x, 4.2, 4.0, 0.4, cap, size=14, bold=True, color=NAVY, align="c")
txt(sl, 0.45, 5.0, 12.4, 1.4, "Live dashboards from the running SonarQube and Grafana instances of this project, not mockups.",
    size=16, italic=True, align="c")

# ---------- ML pipeline ----------
sl = base("04. Implementation: Machine Learning", "Machine Learning Pipeline", 30)
card(sl, 0.4, 1.9, 2.15, 4.4, "Sensor telemetry", "air temperature\nprocess temperature\nrotational speed\ntorque\ntool wear", line=BLUE, tsize=13, bsize=11)
arrow(sl, 2.55, 4.1, 2.9, 4.1, NAVY)
card(sl, 2.9, 3.4, 1.85, 1.4, "Feature pipeline", "heat build-up and\nworkload", line=NAVY, tsize=12, bsize=10)
arrow(sl, 4.75, 4.1, 5.1, 4.1, NAVY)
box(sl, 5.1, 1.75, 3.55, 2.3, line=ORANGE, lw=2.25, radius=0.06)
txt(sl, 5.15, 1.78, 3.45, 0.3, "7 models in parallel (P1-P7)", size=12, bold=True, color=ORANGE, align="c")
p7 = ["P1 Failure chance", "P2 Failure type", "P3 Life left", "P4 Odd behavior", "P5 Urgency", "P6 Best timing", "P7 Parts needed"]
for i, t in enumerate(p7):
    col, row = divmod(i, 2)
    txt(sl, 5.2 + col * 0.0, 2.1 + i * 0.27, 3.3, 0.27, t, size=10.5, align="l")
box(sl, 5.1, 4.2, 3.55, 2.1, line=TEAL, lw=2.25, radius=0.06)
txt(sl, 5.15, 4.23, 3.45, 0.3, "4 Wave 2 signals (history)", size=12, bold=True, color=TEAL, align="c")
for i, t in enumerate(["PINN: keeps estimates realistic", "Cox PH: how long it can run", "Mahalanobis: off its own normal", "CUSUM + Kalman: slow decline"]):
    txt(sl, 5.2, 4.58 + i * 0.4, 3.4, 0.4, t, size=10.5, align="l")
arrow(sl, 8.65, 2.9, 9.05, 3.9, NAVY)
arrow(sl, 8.65, 5.25, 9.05, 4.3, NAVY)
card(sl, 9.05, 3.5, 1.85, 1.2, "DST fusion", "Dempster-Shafer;\nsays 'not sure'\non conflict", line=RED, tsize=12, bsize=9.5)
arrow(sl, 10.9, 4.1, 11.2, 4.1, NAVY)
card(sl, 11.2, 3.3, 1.75, 1.6, "Health score", "0-100 +\nplain-language\nverdict", line=GREEN, tsize=12, bsize=10)
txt(sl, 0.4, 6.5, 12.5, 0.45, "A score is shown only when the evidence is confident; every result is logged for audit and retraining.",
    size=12, italic=True, color=GREY, align="c")

sl = base("04. Implementation: Machine Learning", "Models & Why They Were Chosen", 28)
rows = [["Problem", "Model used", "Alternatives considered", "Why this one"],
        ["P1 Failure probability", "XGBoost + SMOTE", "Random Forest, logistic regression", "Non-linear, imbalanced sensor data"],
        ["P2 Failure type", "Multi-output XGBoost", "Multi-output Random Forest", "Higher F1 on rare failure modes"],
        ["P3 Remaining life", "XGBoost + quantile heads", "Hand-coded linear regression", "Calibrated uncertainty range (MAE 14.95)"],
        ["P4 Anomaly", "4-method ensemble", "Isolation Forest or autoencoder alone", "Catches shapes a single method misses"],
        ["P5 Priority", "XGBoost multi-class", "Rule-based heuristic", "Learns patterns, not fixed thresholds"],
        ["P6 Schedule", "XGBoost classifier", "Prophet, ARIMA, LSTM", "Robust on sparse per-machine data"],
        ["P7 Parts demand", "Survival + Croston", "Reorder point, moving average", "Built for bursty, intermittent demand"],
        ["Wave 2 fusion", "Dempster-Shafer", "Weighted average, majority vote", "Reports uncertainty, no forced consensus"]]
table(sl, 0.5, 1.7, 12.3, [2.5, 2.9, 3.4, 3.5], rows, row_h=0.56, size=12)

# ---------- RAG ----------
sl = base("04. Implementation: RAG Assistant", "Conversational RAG Assistant", 30)
txt(sl, 0.5, 1.7, 6, 0.4, "Ingestion (admin, once per document)", size=14, bold=True, color=RED, align="l")
ing = [("Upload", "PDF / docs via\nbackend (ADMIN)"), ("MinIO", "bucket rag-docs\nstored first"), ("Chunk + embed", "BAAI/bge-m3\nvia RAG service"), ("pgvector", "chunks stored in\nPostgreSQL")]
for i, (h, d) in enumerate(ing):
    x = 0.5 + i * 3.1
    card(sl, x, 2.15, 2.7, 1.2, h, d, line=TEAL, tsize=14, bsize=10.5)
    if i < 3:
        arrow(sl, x + 2.7, 2.75, x + 3.1, 2.75, NAVY)
txt(sl, 0.5, 3.75, 6, 0.4, "Query (every chat message)", size=14, bold=True, color=RED, align="l")
card(sl, 0.5, 4.2, 2.1, 1.3, "Question", "+ machine_id\nfrom the machine page", line=BLUE, tsize=13, bsize=10.5)
arrow(sl, 2.6, 4.85, 3.0, 4.15, NAVY)
arrow(sl, 2.6, 4.85, 3.0, 5.55, NAVY)
card(sl, 3.0, 3.6, 3.0, 1.1, "Document retrieval", "hybrid (vector + keyword)\n+ cross-encoder rerank", line=ORANGE, tsize=12, bsize=10)
card(sl, 3.0, 5.0, 3.0, 1.1, "Live ML snapshot", "unified health of the\nmachine, in parallel", line=ORANGE, tsize=12, bsize=10)
txt(sl, 3.15, 4.72, 2.8, 0.26, "fetched in parallel", size=9, italic=True, color=GREY, align="c")
arrow(sl, 6.0, 4.15, 6.5, 4.85, NAVY)
arrow(sl, 6.0, 5.55, 6.5, 4.85, NAVY)
card(sl, 6.5, 4.2, 2.5, 1.3, "Prompt + Groq LLM", "ML block injected just\nbefore the question", line=PURPLE, tsize=12, bsize=10)
arrow(sl, 9.0, 4.85, 9.4, 4.85, NAVY)
card(sl, 9.4, 4.2, 3.4, 1.3, "Grounded answer", "docs + what the machine is\ndoing right now · ml_context_used", line=GREEN, tsize=12, bsize=10)
txt(sl, 0.5, 6.3, 12.3, 0.6, "If the ML service is down the assistant still answers: a rule-based fallback score keeps the context block valid.",
    size=12, italic=True, color=GREY, align="c")

sl = base("04. Design: Data Model", "Class Diagram", 32)
pic(sl, "class.png", 4.6, 1.6, h=5.4)

# ═════════════════ 05 · DEMO ═════════════════
divider("05", "Demo")

sl = base("05. Demo", "Features Delivered", 32)
for i, (f, cap) in enumerate([("feat1.png", "Head Technician dashboard"), ("feat2.png", "Per-machine ML intelligence"),
                              ("feat3.png", "Conversational RAG assistant"), ("feat4.png", "Secure authentication")]):
    col, row = divmod(i, 2)
    x, y = 0.9 + col * 6.1, 1.75 + row * 2.65
    pic(sl, f, x, y, w=4.0)
    txt(sl, x, y + 2.08, 4.0, 0.4, cap, size=13, bold=True, color=NAVY, align="c")

sl = base("05. Demo", "Live Walkthrough of the Nine Areas", 30)
for i, f in enumerate(fr):
    col, row = divmod(i, 3)
    x, y = 0.8 + col * 4.1, 1.85 + row * 1.0
    box(sl, x, y, 3.8, 0.8, line=NAVY, lw=2)
    txt(sl, x + 0.1, y, 0.6, 0.8, f"{i + 1}", size=20, font=TITLE_FONT, color=RED, align="c")
    txt(sl, x + 0.7, y, 3.05, 0.8, f, size=13, bold=True, align="l")
box(sl, 3.2, 5.2, 6.9, 1.3, line=RED, lw=3, radius=0.2)
txt(sl, 3.2, 5.2, 6.9, 1.3, [[("▶  Narrated demo video", {"bold": True, "color": RED, "size": 24})],
                             [("end-to-end · all four roles · 6:47", {"size": 14})]], align="c")

# ═════════════════ 06 · CONCLUSION ═════════════════
divider("06", "Conclusion\n& Perspectives")

sl = base("06. Conclusion", "", 30)
txt(sl, 0.9, 1.7, 11.5, 1.7, "A solid EAM foundation and an honest predictive AI, secured and deployable on Azure",
    size=26, font=TITLE_FONT, align="c")
for i, t in enumerate(["Operational management + 7 ML models + RAG assistant, in one platform",
                       "Every feature specified, planned, built, then verified, with full traceability",
                       "Honest rigor: label leakage was detected and reported, not hidden",
                       "Secured by a 6-stage pipeline; staging and prod namespaces running on AKS"]):
    box(sl, 1.6, 3.55 + i * 0.85, 10.1, 0.68, line=TEAL, lw=2.25, radius=0.5)
    txt(sl, 1.9, 3.55 + i * 0.85, 9.6, 0.68, [[("→  ", {}), (t, {"bold": True})]], size=14, align="l")

sl = base("06. Perspectives", "", 30)
box(sl, 4.4, 2.3, 4.5, 4.5, line=None, fill=PINK, shape=MSO_SHAPE.OVAL)
box(sl, 5.2, 3.1, 2.9, 2.9, line=None, fill=WHITE, shape=MSO_SHAPE.OVAL)
txt(sl, 5.2, 3.1, 2.9, 2.9, "What's\nNext?", size=26, font=TITLE_FONT, color=NAVY, align="c")
for (x, y, t) in [(0.5, 1.75, "Real ground-truth data to resolve the\nleaked labels (P1 / P2 / P5)"),
                  (8.9, 1.75, "Public inbound path on AKS\n(resolve the Load Balancer quota)"),
                  (0.5, 5.6, "Production hardening: HA database,\nautoscaling, centralized logs"),
                  (8.9, 5.6, "Managed object storage and\nsecret rotation")]:
    box(sl, x, y, 3.95, 1.25, line=RED, lw=2.25, shape=MSO_SHAPE.HEXAGON)
    txt(sl, x + 0.5, y, 3.0, 1.25, t.replace("\n", " "), size=12, bold=True, align="c")

sl = prs.slides.add_slide(BLANK)
pic(sl, "esprit.jpg", 0.55, 0.3, w=2.6)
pic(sl, "sagemcom.png", 11.0, 0.15, w=1.7)
line(sl, 6.67, 0, 6.67, 1.2, BLACK, 1.25)
line(sl, 6.67, 6.6, 6.67, 7.5, BLACK, 1.25)
txt(sl, 1.2, 2.5, 10.9, 1.8, "Thank You for your\nAttention", size=44, font=TITLE_FONT, color=NAVY, align="c")
txt(sl, 1.2, 4.6, 10.9, 0.8, "Any Questions?", size=30, font=TITLE_FONT, color=RED, align="c")
page_no(sl)

prs.save(OUT)
print("saved", OUT, "slides:", len(prs.slides))
