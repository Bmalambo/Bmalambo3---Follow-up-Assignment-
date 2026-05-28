"""
build_presentation.py  --  Inflation Is Not Always Bad
Professional redesign: 16:9 widescreen, min 20pt body text,
one idea per slide, Storytelling with Data 3-Minute Story structure.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ── Palette ──────────────────────────────────────────────────────────────────
NAVY       = RGBColor(0x0A, 0x16, 0x28)
CARD       = RGBColor(0x11, 0x22, 0x3A)
CARD2      = RGBColor(0x15, 0x2A, 0x48)
GREEN      = RGBColor(0x00, 0xA8, 0x6B)
LIME       = RGBColor(0x7E, 0xD3, 0x21)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
LGREY      = RGBColor(0xCC, 0xD6, 0xE0)
MGREY      = RGBColor(0x88, 0x99, 0xAA)
YELLOW     = RGBColor(0xFF, 0xC4, 0x00)
AMBER      = RGBColor(0xFF, 0x8C, 0x00)
RED        = RGBColor(0xFF, 0x4A, 0x4A)
BLUE_DARK  = RGBColor(0x00, 0x4E, 0x9A)
TEAL       = RGBColor(0x00, 0x7A, 0x5E)

# ── Slide size: 16:9 widescreen ──────────────────────────────────────────────
W = 13.33   # inches
H = 7.5

# ── Font ─────────────────────────────────────────────────────────────────────
FONT = "Calibri"


# ── Low-level helpers ─────────────────────────────────────────────────────────
def new_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])   # blank layout


def bg(slide, colour=NAVY):
    f = slide.background.fill
    f.solid()
    f.fore_color.rgb = colour


def rect(slide, x, y, w, h, colour, line=False):
    from pptx.util import Inches
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    s = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = colour
    if line:
        s.line.color.rgb = colour
    else:
        s.line.fill.background()
    return s


def tb(slide, text, x, y, w, h,
       size=22, bold=False, italic=False,
       colour=WHITE, align=PP_ALIGN.LEFT, wrap=True, spacing_before=0):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    if spacing_before:
        p.space_before = Pt(spacing_before)
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = colour
    run.font.name = FONT
    return box


def multi_tb(slide, lines, x, y, w, h,
             size=22, bold=False, colour=WHITE,
             align=PP_ALIGN.LEFT, gap=4):
    """
    lines : list of str  OR  list of (str, bool)  [text, bold_override]
    """
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    first = True
    for item in lines:
        txt, b = (item, bold) if isinstance(item, str) else item
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.space_before = Pt(0 if first else gap)
        run = p.add_run()
        run.text = txt
        run.font.size = Pt(size)
        run.font.bold = b
        run.font.color.rgb = colour
        run.font.name = FONT
    return box


def bar_top(slide, colour=GREEN):
    rect(slide, 0, 0, W, 0.07, colour)


def bar_bot(slide, colour=GREEN):
    rect(slide, 0, H - 0.07, W, 0.07, colour)


def page_label(slide, text):
    tb(slide, text, 0.4, H - 0.42, W - 0.8, 0.32,
       size=11, colour=MGREY, align=PP_ALIGN.RIGHT)


def section_label(slide, text, y=0.18):
    tb(slide, text.upper(), 0.55, y, W - 1.1, 0.38,
       size=13, bold=True, colour=GREEN)


def divider(slide, y, colour=GREEN, x=0.55, w=None):
    if w is None:
        w = W - 1.1
    rect(slide, x, y, w, 0.03, colour)


# ════════════════════════════════════════════════════════════════════════════
# SLIDE 1  -  Cover
# ════════════════════════════════════════════════════════════════════════════
def s01_cover(prs):
    slide = new_slide(prs)
    bg(slide)
    bar_top(slide)
    bar_bot(slide)

    # Left green accent panel
    rect(slide, 0, 0.07, 0.35, H - 0.14, CARD)

    # OECD badge
    rect(slide, 0.55, 0.5, 2.8, 1.4, GREEN)
    tb(slide, "OECD", 0.65, 0.55, 2.6, 1.3,
       size=58, bold=True, colour=NAVY, align=PP_ALIGN.CENTER)

    # Main title  -- large, bold, readable
    tb(slide, "Inflation Is Not Always Bad",
       0.55, 2.2, 12.3, 1.5,
       size=48, bold=True, colour=WHITE)

    # Subtitle
    tb(slide, "When Rising Prices Are a Sign of a Healthy Economy",
       0.55, 3.85, 12.3, 0.75,
       size=26, colour=LGREY)

    divider(slide, 4.75, x=0.55, w=10)

    # Context
    tb(slide, "Follow-up to: Driving Economic Growth  |  Investment, Inflation & FDI (2000-2024)  |  OECD Cohort  2026",
       0.55, 4.88, 12.3, 0.5,
       size=14, italic=True, colour=MGREY)

    page_label(slide, "Slide 1 of 9")
    return slide


# ════════════════════════════════════════════════════════════════════════════
# SLIDE 2  -  The Big Idea
# ════════════════════════════════════════════════════════════════════════════
def s02_big_idea(prs):
    slide = new_slide(prs)
    bg(slide)
    bar_top(slide)
    bar_bot(slide)

    section_label(slide, "The Big Idea")

    # Large quote card
    rect(slide, 0.4, 0.75, W - 0.8, 5.7, CARD)
    rect(slide, 0.4, 0.75, 0.18, 5.7, GREEN)   # left border

    # Main statement -- 3 lines, 32pt each
    tb(slide, "Moderate inflation -- kept between 1.5% and 3% --",
       0.8, 1.0, W - 1.4, 0.85,
       size=30, bold=True, colour=WHITE)
    tb(slide, "is not the enemy.",
       0.8, 1.85, W - 1.4, 0.75,
       size=30, bold=True, colour=GREEN)

    divider(slide, 2.75, x=0.8, w=W - 2.0)

    # Supporting sentences -- 22pt
    tb(slide, "It rewards borrowers, motivates spending, and signals that businesses are growing.",
       0.8, 2.9, W - 1.4, 0.75,
       size=22, colour=LGREY)
    tb(slide, "It erodes the real burden of debt and gives central banks room to act in a crisis.",
       0.8, 3.7, W - 1.4, 0.75,
       size=22, colour=LGREY)

    divider(slide, 4.55, x=0.8, w=W - 2.0, colour=YELLOW)

    # Punchline
    tb(slide, "The real danger is not inflation itself -- it is losing control of it.",
       0.8, 4.72, W - 1.4, 0.75,
       size=22, bold=True, italic=True, colour=YELLOW)

    page_label(slide, "Storytelling with Data  |  Big Idea  |  Slide 2 of 9")
    return slide


# ════════════════════════════════════════════════════════════════════════════
# SLIDE 3  -  The Setup: Fear vs Reality
# ════════════════════════════════════════════════════════════════════════════
def s03_setup(prs):
    slide = new_slide(prs)
    bg(slide)
    bar_top(slide)
    bar_bot(slide)

    section_label(slide, "The Setup  |  Why We Fear Inflation -- And Why Economics Disagrees")

    # ---- LEFT PANEL: The Fear ----
    rect(slide, 0.4, 0.75, 5.9, 5.75, CARD)
    rect(slide, 0.4, 0.75, 5.9, 0.55, RED)
    tb(slide, "The Fear", 0.6, 0.78, 5.6, 0.48,
       size=22, bold=True, colour=WHITE, align=PP_ALIGN.CENTER)

    fear_points = [
        "Prices rise -- purchasing power falls",
        "Savings erode if interest rates lag behind",
        "Fixed-income earners lose real wealth",
        "Hyperinflation collapses economies\n(Zimbabwe 2008, Venezuela 2018)",
        "Businesses struggle to forecast and plan",
    ]
    for i, pt in enumerate(fear_points):
        rect(slide, 0.55, 1.45 + i * 0.97, 0.25, 0.25, RED)
        tb(slide, pt, 0.95, 1.38 + i * 0.97, 5.1, 0.85,
           size=20, colour=LGREY)

    # ---- RIGHT PANEL: The Reality ----
    rect(slide, 7.05, 0.75, 5.9, 5.75, CARD)
    rect(slide, 7.05, 0.75, 5.9, 0.55, GREEN)
    tb(slide, "The Economic Reality", 7.25, 0.78, 5.6, 0.48,
       size=22, bold=True, colour=WHITE, align=PP_ALIGN.CENTER)

    reality_points = [
        "Moderate inflation signals strong demand",
        "Encourages spending today, not tomorrow",
        "Borrowers repay debt in cheaper future money",
        "Central banks target 2% -- by design",
        "Our 8 OECD economies averaged 2.39%",
    ]
    for i, pt in enumerate(reality_points):
        rect(slide, 7.2, 1.45 + i * 0.97, 0.25, 0.25, GREEN)
        tb(slide, pt, 7.6, 1.38 + i * 0.97, 5.1, 0.85,
           size=20, colour=LGREY)

    # VS divider
    tb(slide, "VS", 6.15, 3.45, 1.0, 0.75,
       size=32, bold=True, colour=YELLOW, align=PP_ALIGN.CENTER)

    page_label(slide, "3-Minute Story  |  Setup  |  Slide 3 of 9")
    return slide


# ════════════════════════════════════════════════════════════════════════════
# SLIDE 4  -  Four Mechanisms (Why Moderate Inflation Works)
# ════════════════════════════════════════════════════════════════════════════
def s04_mechanisms(prs):
    slide = new_slide(prs)
    bg(slide)
    bar_top(slide)
    bar_bot(slide)

    section_label(slide, "The Economics  |  4 Reasons Moderate Inflation Works For Us")

    mechs = [
        ("01", GREEN,  "The Spending Incentive",
         "If prices will be higher tomorrow, consumers and businesses spend today.\nThis keeps demand flowing and prevents a deflationary freeze."),
        ("02", LIME,   "The Debt Erosion Effect",
         "A fixed-rate loan becomes cheaper to repay as inflation rises.\nMortgages, student loans, and business debt all shrink in real terms."),
        ("03", AMBER,  "The Wage Adjustment Buffer",
         "It is easier to give a 3% raise when inflation is 2% (a 1% real raise)\nthan to cut nominal wages. Inflation gives employers and workers flexibility."),
        ("04", YELLOW, "The Monetary Policy Room",
         "Central banks cut interest rates to stimulate the economy.\nIf inflation is 0%, rates hit zero and the central bank runs out of tools."),
    ]

    row_h = 1.48
    start_y = 0.78

    for i, (num, col, title, body) in enumerate(mechs):
        y = start_y + i * row_h
        rect(slide, 0.4, y, W - 0.8, row_h - 0.1, CARD)
        rect(slide, 0.4, y, 1.0, row_h - 0.1, col)
        tb(slide, num, 0.42, y + 0.22, 0.95, 0.9,
           size=36, bold=True, colour=NAVY, align=PP_ALIGN.CENTER)
        tb(slide, title, 1.6, y + 0.1, 4.5, 0.55,
           size=24, bold=True, colour=WHITE)
        tb(slide, body, 1.6, y + 0.62, W - 2.4, 0.82,
           size=19, colour=LGREY)

    page_label(slide, "3-Minute Story  |  Rising Action  |  Slide 4 of 9")
    return slide


# ════════════════════════════════════════════════════════════════════════════
# SLIDE 5  -  Who Benefits
# ════════════════════════════════════════════════════════════════════════════
def s05_benefits(prs):
    slide = new_slide(prs)
    bg(slide)
    bar_top(slide)
    bar_bot(slide)

    section_label(slide, "Who Benefits  |  Real Winners When Inflation Is Moderate & Stable")

    groups = [
        ("Homeowners",          GREEN,
         "Mortgage debt erodes in real terms while property values rise with inflation."),
        ("Governments",         LIME,
         "National debt becomes cheaper to service. Inflation quietly reduces wartime or pandemic borrowing."),
        ("Businesses",          AMBER,
         "Rising output prices widen margins, enabling investment, hiring, and expansion."),
        ("Workers",             YELLOW,
         "Wage negotiations in an inflationary environment allow real incomes to grow."),
        ("Equity Investors",    GREEN,
         "Companies with pricing power protect margins. Stocks outperform cash in moderate inflation."),
        ("Export Nations",      LIME,
         "Mild inflation can weaken the currency, making exports more competitive globally."),
    ]

    xs = [0.4, 4.65, 8.9]
    ys = [0.75, 3.85]
    cell_w = 4.0
    cell_h = 2.85

    for idx, (name, col, body) in enumerate(groups):
        col_i = idx % 3
        row_i = idx // 3
        x = xs[col_i]
        y = ys[row_i]

        rect(slide, x, y, cell_w, cell_h, CARD)
        rect(slide, x, y, cell_w, 0.55, col)
        tb(slide, name, x + 0.15, y + 0.07, cell_w - 0.25, 0.48,
           size=22, bold=True, colour=NAVY if col == YELLOW else WHITE)
        tb(slide, body, x + 0.15, y + 0.68, cell_w - 0.25, 2.0,
           size=18, colour=LGREY)

    page_label(slide, "3-Minute Story  |  Rising Action  |  Slide 5 of 9")
    return slide


# ════════════════════════════════════════════════════════════════════════════
# SLIDE 6  -  The Goldilocks Zone
# ════════════════════════════════════════════════════════════════════════════
def s06_goldilocks(prs):
    slide = new_slide(prs)
    bg(slide)
    bar_top(slide)
    bar_bot(slide)

    section_label(slide, "The Goldilocks Zone  |  Not Too Cold, Not Too Hot")

    # ---- Spectrum bar ----
    bar_y = 0.9
    bar_h = 1.1
    zones = [
        (0.4,  2.0, BLUE_DARK, "Below 0%",        "Deflation",       WHITE),
        (2.4,  1.8, TEAL,      "0% to 1.5%",      "Too Low",         WHITE),
        (4.2,  3.0, GREEN,     "1.5% to 3%",      "SWEET SPOT",      NAVY),
        (7.2,  2.5, AMBER,     "3% to 6%",        "Too High",        WHITE),
        (9.7,  3.2, RED,       "Above 6%",        "Hyperinflation",  WHITE),
    ]

    for zx, zw, zcol, zrange, zlabel, ztxt in zones:
        rect(slide, zx, bar_y, zw, bar_h, zcol)
        tb(slide, zrange, zx + 0.05, bar_y + 0.08, zw - 0.1, 0.42,
           size=14, bold=False, colour=ztxt, align=PP_ALIGN.CENTER)
        tb(slide, zlabel, zx + 0.05, bar_y + 0.52, zw - 0.1, 0.52,
           size=17, bold=True, colour=ztxt, align=PP_ALIGN.CENTER)

    # Arrow pointing to sweet spot
    tb(slide, "^", 5.35, 2.05, 0.6, 0.55,
       size=26, bold=True, colour=GREEN, align=PP_ALIGN.CENTER)
    tb(slide, "Central bank target zone\n(Fed, ECB, RBA all target ~2%)",
       3.9, 2.55, 3.5, 0.75,
       size=15, italic=True, colour=GREEN, align=PP_ALIGN.CENTER)

    # ---- Bottom: 3 consequence cards ----
    card_y = 3.55
    card_h = 3.25
    cards = [
        ("Deflation",          BLUE_DARK,
         "Consumers delay spending -- waiting for lower prices.\n\nDemand collapses, businesses cut jobs, prices fall further.\n\nJapan lost a decade to deflation in the 1990s."),
        ("Sweet Spot  1.5-3%", GREEN,
         "Spending is encouraged. Debt is manageable.\n\nCentral banks have room to stimulate if needed.\n\nThe OECD average of 2.39% sits right here."),
        ("Hyperinflation",     RED,
         "Money loses meaning. Wages cannot keep up.\n\nSupply chains collapse. Social contracts break.\n\nZimbabwe 2008: prices doubled every 24 hours."),
    ]

    cx = [0.4, 4.65, 8.9]
    for i, (ctitle, ccol, cbody) in enumerate(cards):
        rect(slide, cx[i], card_y, 4.0, card_h, CARD)
        rect(slide, cx[i], card_y, 4.0, 0.55, ccol)
        tb(slide, ctitle, cx[i] + 0.15, card_y + 0.07, 3.7, 0.48,
           size=20, bold=True,
           colour=NAVY if ccol == GREEN else WHITE,
           align=PP_ALIGN.CENTER)
        tb(slide, cbody, cx[i] + 0.2, card_y + 0.68, 3.6, 2.4,
           size=18, colour=LGREY)

    page_label(slide, "3-Minute Story  |  Climax  |  Slide 6 of 9")
    return slide


# ════════════════════════════════════════════════════════════════════════════
# SLIDE 7  -  The OECD Evidence
# ════════════════════════════════════════════════════════════════════════════
def s07_evidence(prs):
    slide = new_slide(prs)
    bg(slide)
    bar_top(slide)
    bar_bot(slide)

    section_label(slide, "The OECD Evidence  |  Our 8 Economies (2000-2024)")

    # ---- 4 large KPI cards ----
    kpis = [
        ("2.39%",  "Average Inflation",               GREEN),
        ("1.14%",  "Avg GDP per Capita Growth",        LIME),
        ("3.82%",  "Avg FDI Inflows (% of GDP)",       AMBER),
        ("23.43%", "Avg Gross Fixed Capital Formation", YELLOW),
    ]
    kw = (W - 1.0) / 4
    for i, (val, lbl, col) in enumerate(kpis):
        kx = 0.5 + i * kw
        rect(slide, kx, 0.72, kw - 0.12, 1.85, CARD)
        rect(slide, kx, 0.72, kw - 0.12, 0.06, col)
        tb(slide, val, kx + 0.1, 0.85, kw - 0.3, 0.9,
           size=42, bold=True, colour=col, align=PP_ALIGN.CENTER)
        tb(slide, lbl, kx + 0.1, 1.75, kw - 0.3, 0.7,
           size=14, colour=LGREY, align=PP_ALIGN.CENTER)

    # ---- Main insight ----
    rect(slide, 0.4, 2.78, W - 0.8, 1.2, CARD)
    rect(slide, 0.4, 2.78, 0.18, 1.2, GREEN)
    tb(slide, "An average inflation of 2.39% across 8 economies over 24 years sits squarely in the Goldilocks zone.",
       0.75, 2.88, W - 1.4, 0.55,
       size=22, bold=True, colour=WHITE)
    tb(slide, "Three shocks tested that stability: the 2008 Credit Crisis, the 2020 COVID collapse, and the 2022 inflation surge.",
       0.75, 3.45, W - 1.4, 0.45,
       size=19, colour=LGREY)

    # ---- Country callouts ----
    rows = [
        ("Czechia",   "Highest avg inflation (3.2%) -- also the strongest GDP growth signal in the group"),
        ("France",    "Lowest avg inflation (1.7%) -- economic activity remained moderate and steady"),
        ("Belgium",   "Highest FDI inflows (9.53% avg) -- foreign capital flows towards price-stable economies"),
        ("Denmark",   "Lowest inflation volatility -- a consistent and credible policy anchor over 24 years"),
    ]
    for i, (country, note) in enumerate(rows):
        ry = 4.15 + i * 0.68
        rect(slide, 0.4, ry, W - 0.8, 0.62, CARD2)
        tb(slide, country, 0.6, ry + 0.09, 1.6, 0.46,
           size=18, bold=True, colour=GREEN)
        tb(slide, note, 2.35, ry + 0.09, W - 3.1, 0.46,
           size=18, colour=LGREY)

    page_label(slide, "3-Minute Story  |  Evidence  |  Slide 7 of 9")
    return slide


# ════════════════════════════════════════════════════════════════════════════
# SLIDE 8  -  3 Takeaways (Resolution + Call to Action)
# ════════════════════════════════════════════════════════════════════════════
def s08_takeaways(prs):
    slide = new_slide(prs)
    bg(slide)
    bar_top(slide)
    bar_bot(slide)

    section_label(slide, "The Resolution  |  Three Things to Remember")

    takeaways = [
        (GREEN,  "1",
         "Ask the Rate -- Not Just 'Is There Inflation?'",
         "2% and 20% are both 'inflation' -- but they are opposites in effect.\n"
         "The number matters. Context matters. Always ask: how much?"),
        (AMBER,  "2",
         "Know Your Position -- Borrower or Saver?",
         "Inflation transfers wealth from lenders to borrowers.\n"
         "A homeowner with a mortgage benefits. A retiree on fixed savings loses."),
        (YELLOW, "3",
         "Watch the Trend -- Stable vs. Rising",
         "Inflation stable at 2% for a decade is an economic success story.\n"
         "Inflation rising from 2% to 5% is an early warning signal."),
    ]

    for i, (col, num, heading, body) in enumerate(takeaways):
        y = 0.82 + i * 1.82
        rect(slide, 0.4, y, W - 0.8, 1.7, CARD)
        rect(slide, 0.4, y, 1.1, 1.7, col)
        tb(slide, num, 0.42, y + 0.35, 1.05, 0.9,
           size=44, bold=True, colour=NAVY, align=PP_ALIGN.CENTER)
        tb(slide, heading, 1.7, y + 0.1, W - 2.5, 0.55,
           size=24, bold=True, colour=WHITE)
        tb(slide, body, 1.7, y + 0.68, W - 2.5, 0.9,
           size=20, colour=LGREY)

    # Big Idea restated
    divider(slide, 6.35, x=0.4, w=W - 0.8, colour=GREEN)
    tb(slide,
       '"Moderate inflation is the economy\'s heartbeat -- steady, controlled, and necessary for growth."',
       0.55, 6.45, W - 1.1, 0.65,
       size=19, bold=True, italic=True, colour=GREEN, align=PP_ALIGN.CENTER)

    page_label(slide, "3-Minute Story  |  Resolution & Call to Action  |  Slide 8 of 9")
    return slide


# ════════════════════════════════════════════════════════════════════════════
# SLIDE 9  -  References
# ════════════════════════════════════════════════════════════════════════════
def s09_references(prs):
    slide = new_slide(prs)
    bg(slide)
    bar_top(slide)
    bar_bot(slide)

    section_label(slide, "References & Sources")

    sources = [
        ("Data",    "OECD Data Explorer -- GDP per capita growth, GFCF, FDI net inflows, Inflation CPI (2000-2024)"),
        ("Data",    "World Bank -- World Development Indicators (WDI)"),
        ("Policy",  "Federal Reserve -- Why Does the Fed Aim for 2% Inflation? (2021)"),
        ("Policy",  "European Central Bank -- Inflation and the ECB 2% Target Explained"),
        ("Policy",  "Reserve Bank of Australia -- Inflation Target and Monetary Policy (2023 Review)"),
        ("Textbook","Mankiw, N.G. -- Principles of Economics, 9th Edition (Chapters 11, 16, 30)"),
        ("Textbook","Krugman & Wells -- Macroeconomics, 5th Edition -- Inflation and the Price Level"),
        ("Research","Knuuti et al. -- The Benefits of Inflation Targeting, IMF Working Paper (2018)"),
        ("Research","Bank for International Settlements -- BIS Quarterly Review: Inflation Dynamics"),
        ("Design",  "Cole Nussbaumer Knaflic -- Storytelling with Data (2015) -- narrative structure & visualisation"),
    ]

    tag_colours = {
        "Data":     GREEN,
        "Policy":   AMBER,
        "Textbook": LIME,
        "Research": YELLOW,
        "Design":   RGBColor(0x88, 0xBB, 0xFF),
    }

    for i, (tag, text) in enumerate(sources):
        ry = 0.72 + i * 0.64
        rect(slide, 0.4, ry, W - 0.8, 0.57, CARD if i % 2 == 0 else CARD2)
        rect(slide, 0.4, ry, 1.2, 0.57, tag_colours[tag])
        tb(slide, tag, 0.42, ry + 0.09, 1.16, 0.4,
           size=13, bold=True, colour=NAVY, align=PP_ALIGN.CENTER)
        tb(slide, text, 1.75, ry + 0.09, W - 2.5, 0.4,
           size=14, colour=LGREY)

    page_label(slide, "OECD Cohort  |  bmalambo  |  2026  |  Slide 9 of 9")
    return slide


# ════════════════════════════════════════════════════════════════════════════
# MAIN
# ════════════════════════════════════════════════════════════════════════════
def build():
    prs = Presentation()
    prs.slide_width  = Inches(W)
    prs.slide_height = Inches(H)

    s01_cover(prs)
    s02_big_idea(prs)
    s03_setup(prs)
    s04_mechanisms(prs)
    s05_benefits(prs)
    s06_goldilocks(prs)
    s07_evidence(prs)
    s08_takeaways(prs)
    s09_references(prs)

    output = "inflation_not_always_bad.pptx"
    prs.save(output)
    print(f"[OK] Saved: {output}  ({len(prs.slides)} slides)")
    print(f"     Location: C:\\Users\\bmalambo\\Bmalambo3---Follow-up-Assignment-\\{output}")


if __name__ == "__main__":
    build()
