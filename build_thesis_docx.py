"""
Build the Nordic property maintenance roll-up investment thesis as a
properly formatted Word document.

Reframe vs. the prior markdown: this is a buy-and-build play where KLAR
creates the platform via aggressive M&A, the same way Norvestor built PHM.
The investable universe is anchor candidates + bolt-on supply, not
sponsor-owned adjacent platforms.
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# --------------------------------------------------------------------------------------
# Styling
# --------------------------------------------------------------------------------------
NAVY = RGBColor(0x1F, 0x3A, 0x5F)
DARK_GREY = RGBColor(0x33, 0x33, 0x33)
MID_GREY = RGBColor(0x66, 0x66, 0x66)
LIGHT_BG = "F2F4F7"
HEADER_BG = "1F3A5F"
HEADER_FG = RGBColor(0xFF, 0xFF, 0xFF)


def set_cell_bg(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)


def set_cell_borders(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        b = OxmlElement(f'w:{edge}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '4')
        b.set(qn('w:color'), 'BFBFBF')
        tcBorders.append(b)
    tcPr.append(tcBorders)


def add_para(doc, text, size=10.5, bold=False, color=None, italic=False,
             space_before=0, space_after=4, align=None):
    p = doc.add_paragraph()
    if align:
        p.alignment = align
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.font.size = Pt(size)
    r.bold = bold
    r.italic = italic
    if color:
        r.font.color.rgb = color
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    return p


def add_runs(doc, segments, size=10.5, space_before=0, space_after=4):
    """segments: list of (text, bold_bool) tuples."""
    p = doc.add_paragraph()
    for text, bold in segments:
        r = p.add_run(text)
        r.font.name = "Calibri"
        r.font.size = Pt(size)
        r.bold = bold
        r.font.color.rgb = DARK_GREY
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    return p


def h1(doc, text):
    add_para(doc, text, size=16, bold=True, color=NAVY, space_before=14, space_after=6)


def h2(doc, text):
    add_para(doc, text, size=13, bold=True, color=NAVY, space_before=10, space_after=4)


def h3(doc, text):
    add_para(doc, text, size=11.5, bold=True, color=DARK_GREY, space_before=8, space_after=3)


def body(doc, text):
    add_para(doc, text, size=10.5, color=DARK_GREY, space_after=4)


def bullet(doc, text, level=0):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.left_indent = Cm(0.6 + 0.6 * level)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.font.size = Pt(10.5)
    r.font.color.rgb = DARK_GREY


def build_table(doc, headers, rows, widths_cm=None):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.autofit = False
    table.allow_autofit = False

    # Header row
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = ""
        p = hdr_cells[i].paragraphs[0]
        r = p.add_run(h)
        r.font.name = "Calibri"
        r.font.size = Pt(9.5)
        r.font.bold = True
        r.font.color.rgb = HEADER_FG
        set_cell_bg(hdr_cells[i], HEADER_BG)
        set_cell_borders(hdr_cells[i])
        hdr_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # Data rows
    for r_idx, row in enumerate(rows, start=1):
        cells = table.rows[r_idx].cells
        for c_idx, val in enumerate(row):
            cells[c_idx].text = ""
            p = cells[c_idx].paragraphs[0]
            run = p.add_run(str(val) if val is not None else "")
            run.font.name = "Calibri"
            run.font.size = Pt(9.5)
            run.font.color.rgb = DARK_GREY
            set_cell_borders(cells[c_idx])
            cells[c_idx].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if r_idx % 2 == 0:
                set_cell_bg(cells[c_idx], LIGHT_BG)

    if widths_cm:
        for i, w in enumerate(widths_cm):
            for row in table.rows:
                row.cells[i].width = Cm(w)

    # Spacing after table
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return table


# --------------------------------------------------------------------------------------
# Document
# --------------------------------------------------------------------------------------
doc = Document()

# Page margins
for section in doc.sections:
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.8)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)

# Set default font for the document Normal style
style = doc.styles['Normal']
style.font.name = 'Calibri'
style.font.size = Pt(10.5)


# --------------------------------------------------------------------------------------
# COVER BLOCK
# --------------------------------------------------------------------------------------
add_para(doc, "KLAR PARTNERS - SCREENING MEMO", size=9, bold=True,
         color=MID_GREY, space_after=2)
add_para(doc, "Nordic Property Maintenance Roll-Up",
         size=22, bold=True, color=NAVY, space_after=2)
add_para(doc, "Building the #2 platform behind PHM Group",
         size=13, italic=True, color=MID_GREY, space_after=10)

meta = doc.add_table(rows=4, cols=2)
meta.autofit = False
meta_data = [
    ("Sector", "Property maintenance / janitorial (vaktmester, eiendomsdrift, fastighetsskötsel)"),
    ("Geography", "Norway anchor; Nordic platform optionality"),
    ("Reference", "PHM Group (Norvestor VIII, platform vintage 2017)"),
    ("Decision", "Conditional proceed to Phase 2 origination (anchor + pipeline)"),
]
for i, (k, v) in enumerate(meta_data):
    c0, c1 = meta.rows[i].cells
    c0.text = ""
    c1.text = ""
    r0 = c0.paragraphs[0].add_run(k)
    r0.bold = True
    r0.font.size = Pt(10)
    r0.font.color.rgb = NAVY
    r1 = c1.paragraphs[0].add_run(v)
    r1.font.size = Pt(10)
    r1.font.color.rgb = DARK_GREY
    c0.width = Cm(3.5)
    c1.width = Cm(13.5)

doc.add_paragraph()


# --------------------------------------------------------------------------------------
# 1. RECOMMENDATION
# --------------------------------------------------------------------------------------
h1(doc, "1.  Recommendation")

body(doc,
    "Proceed to Phase 2 origination on the basis that this is a buy-and-build "
    "platform construction, not the acquisition of a finished #2 asset. "
    "The thesis is to anchor on a mid-sized founder-led janitor business "
    "(NOK 250-500m revenue) and execute 25-40 bolt-ons over a five-year "
    "hold to reach NOK 2.0-2.5bn revenue at exit, mirroring Norvestor's "
    "PHM playbook with a differentiated commercial-property / pan-Nordic "
    "positioning.")

h3(doc, "Headline in three lines")
bullet(doc,
    "Norwegian property maintenance is structurally attractive (NOK 15-25bn TAM, "
    "fragmented, recurring, regulation-shielded), with PHM at <5% share leaving >95% "
    "of the market open. The Nordic TAM under the same definition is NOK 55-80bn.")
bullet(doc,
    "The roll-up arithmetic works: 200+ named bolt-on targets between NOK 30m "
    "and NOK 400m revenue, typical bolt-on entry at 4-6x EBITDA, platform exit "
    "at 11-13x. Buy-vs-build spread of 4-7 turns is the core return engine.")
bullet(doc,
    "Right to win rests on KLAR's services playbook (Nimlas, Ocab adjacency), "
    "an explicit non-PHM positioning (commercial property + light technical "
    "bundling + Nordic geography), and three credible founder-led anchor "
    "candidates we can approach now.")

h3(doc, "What would kill this deal")
bullet(doc, "Inability to land an anchor at <8.5x within 12-18 months.")
bullet(doc, "Bolt-on price discipline breaks (PHM pushing transaction multiples >7x for sub-NOK 100m targets).")
bullet(doc, "Labour-cost step change (Renholdsoverenskomsten or minimum-wage shock) compressing sector margin >150 bps.")


# --------------------------------------------------------------------------------------
# 2. MARKET THESIS
# --------------------------------------------------------------------------------------
h1(doc, "2.  Market Thesis")

h2(doc, "2.1  Sizing the addressable market")

build_table(doc,
    headers=["Sub-segment", "Norway TAM (NOKbn)", "Basis"],
    rows=[
        ["Daily janitor / building care (vaktmester, eiendomsdrift)", "8 - 12",
         "NACE 81.10 / 81.21; residential + light commercial; ~12k registered entities"],
        ["Cleaning (industrial + residential)", "12 - 16",
         "Statistics Norway services sub-sector; BoldData top-50 aggregate ~NOK 20bn FY23"],
        ["Outdoor / grounds / snow", "2 - 4",
         "Seasonal but contracted; lumpy but visible"],
        ["Light technical FM (HVAC service, electrical, plumbing visits)", "5 - 8",
         "Adjacent; Nimlas-relevant; partial overlap with platform thesis"],
        ["Addressable for this roll-up (excl. catering, security, heavy IFM)", "15 - 25",
         "Brief baseline; consistent with bottom-up"],
    ],
    widths_cm=[5.5, 3.0, 8.5])

body(doc,
    "Sweden (fastighetsskötsel) adds SEK 25-35bn equivalent under the same definition; "
    "Denmark and Finland together add an estimated NOK 15-20bn. Pan-Nordic TAM for the "
    "addressable segment is NOK 55-80bn. Nominal growth 4-6% p.a. (population, housing "
    "stock, wage-driven price-throughput); real growth 1-2%.")

h2(doc, "2.2  Structural attractiveness (KLAR sector test)")

build_table(doc,
    headers=["Test", "Read", "Evidence"],
    rows=[
        ["Mission-critical", "Pass",
         "Snow clearing and building safety inspections are legally required; cleaning frequency mandated"],
        ["Non-discretionary demand", "Pass",
         "Demand persists through cycle; 2020 COVID stress test showed janitor volumes resilient"],
        ["Recurring revenue", "Pass",
         "70-85% under multi-year framework agreements (typical 2-3yr auto-renewing); housing co-op contracts especially sticky"],
        ["Pricing power (real)", "Partial",
         "Indexed to KPI / wage index; real margin expansion via density, not price"],
        ["Fragmentation", "Pass (defining)",
         "Top 5 players <25% share; ~12,000 registered cleaning/FM entities; PHM <5%"],
        ["Capex intensity", "Pass",
         "Working capital + small fleet; <2% revenue maintenance capex"],
        ["Cash conversion", "Pass",
         "75-90% EBITDA-to-FCF after WC"],
        ["Cyclicality", "Partial",
         "Operationally insulated; new-contract origination exposed to construction cycle"],
        ["Regulatory shield", "Pass",
         "Allmenngjøringsloven creates a labour-cost floor that disadvantages informal sub-scale operators - i.e., consolidation tailwind"],
    ],
    widths_cm=[4.0, 2.5, 10.5])


h2(doc, "2.3  Why now")
bullet(doc,
    "Generational handover wave. Two-thirds of family-owned mid-sized Norwegian "
    "property service businesses (NOK 100-500m revenue) were founded 1985-2000 "
    "with founders now aged 60+. Sefbo's 2022 sale to PHM and Insider Group's 2024 "
    "Mitie carve-out are both signals that owner appetite is opening.")
bullet(doc,
    "PHM playbook proves the geography. Norvestor has demonstrated in real "
    "time that this segment absorbs a Nordic roll-up at scale. We are no longer "
    "asking whether it works - we are building #2 with the playbook understood.")
bullet(doc,
    "Bolt-on supply abundant. >200 named janitor / eiendomsdrift entities in the "
    "NOK 30-400m revenue band, the vast majority founder-owned and untouched.")
bullet(doc,
    "Multiples have rebased. 2022-2024 saw entry multiples on cleaning / FM assets "
    "compress from 11-13x to 7-9x. Bolt-on multiples 4-6x. Entry environment is constructive.")


# --------------------------------------------------------------------------------------
# 3. PHM AS THE PLAYBOOK
# --------------------------------------------------------------------------------------
h1(doc, "3.  PHM as the Playbook (and the White Space)")

h2(doc, "3.1  What Norvestor built")

bullet(doc,
    "Anchor: small Finnish janitor / property care platform, invested 2017 from Norvestor VIII.")
bullet(doc,
    "Trajectory: ~100 acquisitions across Finland, Norway, Sweden, Denmark. Group revenue "
    "estimated EUR 700-850m FY24 (not publicly disclosed; triangulated from Norwegian sub filings).")
bullet(doc,
    "Norway sub-platform: NOK 1.8-2.4bn revenue, ~30 acquisitions including Sefbo (2022, "
    "the prior largest Norwegian janitor roll-up).")
bullet(doc,
    "Sub-brand retention post-deal. Disciplined bolt-on pricing (4-6x EBITDA <NOK 50m, "
    "6-8x <NOK 200m). Local front-line continuity preserved.")
bullet(doc,
    "Norvestor VIII vintage 2017. Exit window open now. Likely process 2026-2027.")

h2(doc, "3.2  The white space PHM does not cover")

build_table(doc,
    headers=["Gap", "Why PHM has not filled it", "Implication for our platform"],
    rows=[
        ["Light commercial & public-sector janitor (~30-40% of TAM)",
         "PHM is ~70% residential; commercial customer relationships are different sales motion",
         "Our anchor and bolt-on focus tilts commercial - distinct from PHM's residential channel"],
        ["Technical service bundling (HVAC, electrical, plumbing visits)",
         "PHM does not have a technical service arm; bundling requires capability",
         "Nimlas adjacency: cross-referral and bundled-contract offering, 30-50 bps margin upside"],
        ["Damage / restoration handoff",
         "PHM has no restoration arm",
         "Ocab adjacency: water-damage referral partnership (revenue share, not equity)"],
        ["Sweden at scale",
         "PHM's Swedish operations are sub-scale relative to its Norway/Finland footprint",
         "Pan-Nordic platform construction via Swedish bolt-ons in years 3-5"],
        ["Mid-Norway / Vestland density",
         "PHM is Oslo / Akershus / Østfold concentrated",
         "Anchor in Trondheim or Bergen gives us a complementary regional base"],
    ],
    widths_cm=[4.5, 5.5, 7.0])

body(doc,
    "Strategic implication: do not build a residential-Oslo PHM lookalike that "
    "competes head-to-head for the same bolt-ons and the same customers. "
    "Build a commercial-and-Nordic platform with selective residential expansion. "
    "Same TAM, different route in.")


# --------------------------------------------------------------------------------------
# 4. THE ROLL-UP CONSTRUCTION
# --------------------------------------------------------------------------------------
h1(doc, "4.  Building the Platform - the Construction Plan")

h2(doc, "4.1  Anchor candidates")

body(doc,
    "Three credible anchor candidates that fit the commercial-tilt, non-PHM-overlap thesis. "
    "Each is founder-led, NOK 250-1,000m revenue, sale-ready or approachable within 12-18 months.")

build_table(doc,
    headers=["Anchor", "Revenue FY24E (NOKm)", "EBITDA margin (E)", "Owner", "Fit"],
    rows=[
        ["Vaktmesterkompaniet AS (Oslo)", "~1,000", "7-9%", "Founder-led",
         "Closest PHM analog - pure janitor model; founder ~60s; bolt-on track record (~10 deals). Risk: PHM presumed in dialogue"],
        ["Din Vaktmester AS (Trondheim)", "250-350", "8-10%", "Founder",
         "Pure janitor model in mid-Norway - geographically complementary to PHM. Sub-scale but right culture for an anchor build"],
        ["Insider Group AS (Høvik)", "1,000-1,100", "6-8%", "Family",
         "Mitie Norway carve-out (2024) proves M&A muscle; broader IFM mix - need to validate property-maintenance share of revenue"],
    ],
    widths_cm=[5.0, 3.0, 2.5, 2.5, 9.0])

h3(doc, "Anchor strategy view")
bullet(doc,
    "Preferred: Din Vaktmester (Trondheim). Why: smallest cheque (NOK 200-300m EV at 8-10x), "
    "purest property-maintenance fit, geographically complementary to PHM, lowest competing-bidder "
    "intensity. Use it as the anchor to drive Phase 1 bolt-ons in Bergen and Oslo.")
bullet(doc,
    "Acceptable: Vaktmesterkompaniet. Why: scale benefit (immediate Tier-1 platform). "
    "Downside: 9-11x entry, competing bidders likely include PHM and IFM strategics; "
    "may force a deal we cannot win on price.")
bullet(doc,
    "Tactical: Insider Group. Why: M&A muscle is genuine. Downside: cleaning content "
    "is >50% of mix - centre of gravity drifts from property maintenance towards general IFM.")
bullet(doc,
    "Alternative: dual-anchor (Din Vaktmester + Ability Gruppen Bergen, ~NOK 250-400m). "
    "Total invested capital NOK 350-500m EV; combined footprint Trondheim + Bergen + early Oslo. "
    "More integration work but cheaper entry and zero PHM competition.")

h2(doc, "4.2  Bolt-on pipeline - the engine of the roll-up")

body(doc,
    "The roll-up only works if the bolt-on supply is real. We have screened "
    "~75 Norwegian property service entities and identified ~50 in the relevant "
    "ownership / scale band. Tier-1 named bolt-ons (NOK 100-400m revenue):")

build_table(doc,
    headers=["Target", "HQ", "Revenue FY24E (NOKm)", "Owner", "Why interesting"],
    rows=[
        ["Bygårdsservice AS", "Oslo", "80-120", "Founder",
         "Pure PHM-style residential / housing-association janitor. Direct slot-in"],
        ["Aktiv Eiendomsdrift AS", "Roa (Innlandet)", "60-100", "Founder",
         "Property-maintenance-only operator; Innlandet density build"],
        ["Eiendomspartner 1 AS", "Bergen", "150-220", "Family", "Bergen scale; verify business model in Proff"],
        ["Ability FM AS / Ability Gruppen", "Bergen", "250-400", "Family",
         "Vestland platform candidate; can be re-cast from anchor to bolt-on depending on terms"],
        ["Toma Eiendomsdrift AS", "Oslo (Toma sub)", "200-300", "Carve-out from Tomagruppen",
         "Property-maintenance carve-out if family willing to divest non-core; long-shot, long-cultivation"],
        ["Facilitec AS (Insider sub)", "Lørenskog", "200-300", "Insider",
         "Carve-out from Insider if we anchor elsewhere; or comes as part of Insider deal"],
        ["BBL Daglig Drift AS", "Oslo", "100-150", "BBL co-op linked",
         "Housing co-op (borettslag) operator; cooperative ownership requires creative deal structure"],
        ["Conluo Facility Services AS", "Oslo", "150-250", "Founder",
         "Sub-scale platform we re-cast as bolt-on; Oslo density"],
        ["Hvass AS", "Oslo", "60-90", "Founder", "Pure janitor; Oslo East"],
        ["Bygård Vaktmesteren AS", "Oslo", "50-80", "Founder",
         "Direct PHM-comparable housing-association operator"],
        ["RSV Gruppen AS", "Bergen", "60-100", "Founder", "Vestland bolt-on"],
        ["Resolve AS", "Risør", "50-80", "Founder", "Southern Norway density"],
        ["Total Eiendomsservice AS", "Drammen", "50-90", "Founder", "Buskerud / Drammen corridor"],
        ["Northroads AS", "Mo i Rana", "80-130", "Founder", "Northern Norway niche; lower priority"],
        ["Allianse Service Partner", "Oslo", "100-160", "Asset-light franchise",
         "Different model - evaluate consolidation feasibility separately"],
    ],
    widths_cm=[4.5, 2.5, 2.5, 2.5, 7.0])

body(doc,
    "Below this set, the tail comprises 35-50 small operators in the NOK 30-100m "
    "revenue band - typical bolt-on entry at 3-5x EBITDA. Most are founder-owned, "
    "many are sale-ready as founders age out. This tail is the cheap M&A engine "
    "of years 2-4 once the anchor's integration capability is proven.")

h3(doc, "Bolt-on supply: realistic vs. theoretical")
bullet(doc,
    "Theoretical universe: ~200 named entities >NOK 30m revenue across our screen "
    "and Proff.no NACE 81.10 / 81.21 filter.")
bullet(doc,
    "Realistic 5-year pipeline (operator-willing, fit, not PHM-acquired): we model "
    "25-40 closed bolt-ons cumulating to NOK 600-900m acquired revenue.")
bullet(doc,
    "Risk: PHM's bolt-on pace has been 5-8 per year in Norway. If PHM accelerates "
    "in defence of a 2026-2027 exit narrative, our supply set compresses materially.")

h2(doc, "4.3  Five-year construction sequencing (illustrative)")

build_table(doc,
    headers=["Year", "Activity", "Revenue (NOKm)", "EBITDA margin", "EBITDA (NOKm)", "Cumulative bolt-on EV deployed (NOKm)"],
    rows=[
        ["Y0 (entry)", "Anchor acquisition + integration team", "300", "8.5%", "26", "0"],
        ["Y1", "3-5 small bolt-ons; route-density build; cross-brand integration", "550", "9.0%", "50", "120"],
        ["Y2", "5-8 bolt-ons incl. 1 mid-sized; first Swedish probe", "1,000", "9.5%", "95", "320"],
        ["Y3", "6-9 bolt-ons; Nimlas cross-sell pilot; recap optionality", "1,500", "10.5%", "158", "650"],
        ["Y4", "5-8 bolt-ons; Sweden scale-up", "1,950", "11.0%", "215", "950"],
        ["Y5 (exit)", "Tidying; sell-side preparation", "2,300", "11.5%", "265", "1,150"],
    ],
    widths_cm=[2.2, 6.5, 2.3, 2.0, 2.3, 3.7])

body(doc,
    "Anchor enters at 8.5-10.0x EBITDA. Bolt-ons blend at 4-6x EBITDA on entry. "
    "Weighted average acquisition multiple across the five-year program: 5.5-6.5x. "
    "Exit at 11-13x on a platform with NOK 250-280m EBITDA implies EV NOK 2.8-3.6bn.")

h2(doc, "4.4  Margin levers (route to 11-12% EBITDA at exit)")

build_table(doc,
    headers=["Lever", "Expected uplift", "Confidence", "Notes"],
    rows=[
        ["Route density (Oslo, Bergen, Trondheim corridors)", "+100-150 bps", "High",
         "Direct PHM precedent; mileage and idle time reduction on janitor routes"],
        ["Procurement consolidation (chemicals, equipment, fleet, insurance)", "+50-80 bps", "High",
         "Standard PE-roll-up lever; bolt-on accretion is mechanical"],
        ["Digital workforce scheduling and dispatch", "+30-60 bps", "Medium",
         "Implementation cost NOK 15-25m; payback Y2-Y3"],
        ["Cross-sell of light technical FM (Nimlas adjacency)", "+30-50 bps", "Medium",
         "Requires capability transfer; not synthetic but not free"],
        ["Bolt-on cost-synergy capture (back-office, office consolidation)", "+20-40 bps", "High",
         "Standard PMI playbook"],
        ["Total margin expansion (5-year)", "+230-380 bps", "",
         "Anchor enters ~8.5%; exit at 11-12% is mid-range outcome"],
    ],
    widths_cm=[6.0, 2.3, 2.0, 6.7])


# --------------------------------------------------------------------------------------
# 5. KLAR FRAMEWORKS
# --------------------------------------------------------------------------------------
h1(doc, "5.  KLAR Framework Application")

h2(doc, "5.1  4KQ - Four Key Questions")

build_table(doc,
    headers=["Question", "Assessment", "Verdict"],
    rows=[
        ["Q1.  Why this market?",
         "TAM NOK 15-25bn Norway, NOK 55-80bn Nordic. Fragmented (top-5 <25%), recurring (70-85%), regulation-shielded, recession-resilient. Growth 4-6% nominal with high visibility.",
         "Pass with conviction"],
        ["Q2.  Why this asset?",
         "Conditional on landing a founder-led anchor at <8.5x. Asset quality of the construction is strong (recurring, capex-light, scalable) but contingent on bolt-on supply executing.",
         "Conditional pass"],
        ["Q3.  Why now?",
         "Generational handover wave is live. PHM exit 2026-2027 reshapes the competitive landscape. Multiples have rebased to 7-9x at platform / 4-6x at bolt-on level. Window is 18-24 months.",
         "Pass"],
        ["Q4.  Why KLAR?",
         "Nimlas (technical) and Ocab (restoration) adjacencies are real, not synthetic - cross-referral and bundled-contract revenue 30-50 bps margin uplift. Ocab gives Swedish front door PHM lacks. Operational toolkit (route density, procurement, M&A integration) directly transferable.",
         "Pass"],
    ],
    widths_cm=[3.5, 11.0, 2.5])

body(doc, "Aggregate: 3.5 / 4. The gating question is asset access, which converts on Phase 2 origination.")

h2(doc, "5.2  Triple A")

build_table(doc,
    headers=["Dimension", "Score", "Rationale"],
    rows=[
        ["Attractive market", "A",
         "Fragmentation, recurrence, regulation, downside protection all present; moderate but visible growth"],
        ["Attractive business (post-construction)", "A-",
         "Strong on quality of revenue and cash conversion; ceiling on margin (labour services) and on real pricing power constrains the score"],
        ["Attractive entry / exit", "A-",
         "Entry environment constructive (7-9x platform, 4-6x bolt-on). Exit clear (next sponsor or strategic) but sponsor-to-sponsor universe is the realistic base case"],
    ],
    widths_cm=[5.0, 1.5, 10.5])

body(doc, "Triple A composite: A-. Investable; not a layup.")

h2(doc, "5.3  Key Characteristics Assessment")

build_table(doc,
    headers=["Characteristic", "Score (1-5)", "Read"],
    rows=[
        ["Mission-critical service", "5", "Snow, building safety, hygiene - legal requirement"],
        ["Non-discretionary demand", "5", "Resilient through cycle; proven in 2020"],
        ["Recurring revenue", "4", "70-85% multi-year framework"],
        ["Customer concentration", "5", "Typical largest customer <5%"],
        ["Supplier concentration", "5", "Labour + consumables; no chokepoints"],
        ["Pricing power (real, ex-indexation)", "2", "Labour-cost pass-through dominates; ceiling on real pricing"],
        ["Organic growth rate", "3", "2-4% real"],
        ["M&A optionality", "5", "200+ bolt-on targets; defining feature of the thesis"],
        ["EBITDA margin profile", "3", "6-10% entry, 10-12% at scale"],
        ["Cash conversion", "5", "75-90% post-WC"],
        ["ESG / regulatory positioning", "4", "Formal labour, training, energy-efficiency improvements net positive"],
        ["Defensible competitive position", "3", "Density and switching costs (keys, building knowledge); real but not deep"],
        ["Counter-cyclical attributes", "4", "Cleaning resilient; snow mandatory"],
        ["Founder / management dependency at anchor", "2", "Real transition risk - PMI playbook required"],
        ["Scalability of model", "4", "Linear ops with route-density gains"],
        ["Composite", "3.9 / 5.0",
         "B+ business, A- platform through aggressive consolidation. Margin ceiling is structural - this is labour services, not software"],
    ],
    widths_cm=[5.8, 1.7, 9.5])


# --------------------------------------------------------------------------------------
# 6. RETURNS
# --------------------------------------------------------------------------------------
h1(doc, "6.  Returns Sketch")

body(doc, "Illustrative, anchor + 25-40 bolt-on construction over 5 years.")

build_table(doc,
    headers=["Parameter", "Base case", "Downside", "Upside"],
    rows=[
        ["Anchor entry EV (NOKm)", "300", "300", "750"],
        ["Anchor EBITDA at entry", "26", "26", "65"],
        ["Anchor entry multiple", "11.5x", "11.5x", "11.5x"],
        ["Cumulative bolt-on EV deployed Y1-Y5 (NOKm)", "1,150", "700", "1,500"],
        ["Weighted-average acquisition multiple (blended)", "6.0x", "6.5x", "5.5x"],
        ["Year-5 platform revenue (NOKm)", "2,300", "1,400", "3,000"],
        ["Year-5 EBITDA margin", "11.5%", "10.0%", "12.5%"],
        ["Year-5 EBITDA (NOKm)", "265", "140", "375"],
        ["Exit multiple", "11.5x", "9.5x", "13.0x"],
        ["Exit EV (NOKm)", "3,050", "1,330", "4,880"],
        ["Total invested equity (NOKm)", "1,000", "900", "1,300"],
        ["Gross MoM", "2.7x", "1.3x", "3.4x"],
        ["Gross IRR", "22-25%", "5-10%", "30%+"],
    ],
    widths_cm=[6.0, 3.7, 3.7, 3.7])

body(doc,
    "Base case assumes leverage 4.5-5.5x at anchor; bolt-on accordion under term loan; "
    "dividend recap optionality in year 3 if integration is on plan. Downside case is "
    "the scenario where PHM compresses bolt-on supply, blended acquisition multiple "
    "drifts to 6.5x, and the platform exits to a strategic at 9.5x rather than to a "
    "sponsor at 11-12x.")


# --------------------------------------------------------------------------------------
# 7. RISKS
# --------------------------------------------------------------------------------------
h1(doc, "7.  Risks and Pushbacks")

h2(doc, "7.1  Risks that gate IC")

build_table(doc,
    headers=["Risk", "Severity", "Why this could kill the deal"],
    rows=[
        ["Anchor access fails on all three named candidates",
         "High",
         "Without a credible founder-led anchor at <8.5x within 12-18 months, the construction does not start. We must walk in that case"],
        ["PHM accelerates bolt-on pace pre-exit and compresses supply",
         "High",
         "If PHM closes 10-15 deals in 2026 in defence of process marketing, our pipeline thins and bolt-on prices rise"],
        ["Bolt-on multiples drift above 7x for sub-NOK 100m targets",
         "Medium-High",
         "Buy-vs-build spread shrinks from 4-7 turns to 2-4 turns; returns model fails the IC hurdle"],
        ["Labour cost step change (Renholdsoverenskomsten / minimum wage)",
         "Medium",
         "Symmetric across competitors but creates a one-off margin step-down of 100-200 bps"],
        ["Integration capacity at the anchor proves shallow",
         "Medium",
         "Roll-up depends on absorbing 5-8 deals per year; weak anchor management team is a real blocker"],
    ],
    widths_cm=[5.8, 2.2, 9.0])

h2(doc, "7.2  Pushbacks we expect at IC")

build_table(doc,
    headers=["Objection", "Counter"],
    rows=[
        ["PHM has a 9-year head start - they own the bolt-on relationships",
         "PHM has chosen residential / Oslo. Commercial property and mid-Norway / Vestland / Sweden are different acquirer sets. We are not buying the same bolt-ons"],
        ["This is a labour services play with no moat",
         "True at unit level. Moat is platform-level: route density + procurement + integration capability + M&A muscle. Same logic that built PHM's exit narrative"],
        ["Adjacencies to Nimlas / Ocab feel synthetic",
         "Cross-referral and bundled-contract revenue, not equity integration. Realistic uplift 30-50 bps EBITDA margin. We should not over-claim"],
        ["Sponsor-to-sponsor exit universe is thin",
         "6-10 credible bidders (Triton, IK, Adelis, Cinven, Norvestor under different fund vintage, Verdane, FSN) plus strategics (international FM majors, post-Norvestor PHM)"],
        ["Roll-up depends on cheap bolt-on multiples that won't last",
         "If supply tightens and bolt-on multiples drift, we slow the M&A pace and lean on organic and margin levers. Buy-vs-build spread of 3 turns is still investable"],
    ],
    widths_cm=[7.5, 9.5])


# --------------------------------------------------------------------------------------
# 8. CRITICAL PATH
# --------------------------------------------------------------------------------------
h1(doc, "8.  Critical Path to IC")

build_table(doc,
    headers=["Workstream", "Target timing", "Decision gate"],
    rows=[
        ["Management access - Din Vaktmester (preferred anchor)", "Q3 2026",
         "Founder posture on sale within 24 months; integration capacity assessment"],
        ["Management access - Vaktmesterkompaniet (acceptable anchor)", "Q3 2026",
         "Confirm PHM not in advanced dialogue; founder valuation expectation"],
        ["Management access - Insider Group (tactical anchor)", "Q4 2026",
         "Mitie integration trajectory; property-maintenance share of revenue"],
        ["Bolt-on long-list dilligence (top 15-20 named)", "Q4 2026",
         "Validate 5-year pipeline of NOK 600-900m acquired revenue"],
        ["Sector commercial diligence (third-party, scoped)", "Q4 2026",
         "TAM validation; PHM acquisition pace; bolt-on multiple discipline"],
        ["Sweden market scan (fastighetsskötsel)", "Q4 2026",
         "Cross-border optionality via Ocab introductions"],
        ["PHM exit timing intelligence", "Ongoing",
         "Norvestor process kick-off signal"],
        ["Regulatory diligence (allmenngjøring, wage trajectory)", "Q1 2027",
         "No step-change risk in next 24 months"],
    ],
    widths_cm=[6.5, 2.5, 8.0])

h3(doc, "Conditions for advancing to full IC")
bullet(doc, "Confirmed management access on at least one anchor candidate.")
bullet(doc, "Indicative entry multiple <8.5x on the anchor.")
bullet(doc, "Pre-cleared bolt-on pipeline of >NOK 250m revenue for Y1-Y2.")
bullet(doc, "Third-party validation of TAM and PHM positioning.")


# --------------------------------------------------------------------------------------
# 9. WHAT WE'D NEED TO VALIDATE
# --------------------------------------------------------------------------------------
h1(doc, "9.  What We'd Need to Validate Before IC")

build_table(doc,
    headers=["Item", "Why it matters", "Source"],
    rows=[
        ["Anchor candidate detailed financials FY22-24", "Sets entry valuation and integration baseline",
         "Proff.no, Brønnøysundregistrene, management approach"],
        ["PHM Norway segment P&L and bolt-on cadence", "Competitive intelligence on supply and pricing",
         "Norvestor exit prep marketing material (once process starts); banker intel"],
        ["Bolt-on pipeline scoring (top 20 named targets)", "Realism check on 5-year construction plan",
         "Per-name diligence; Proff.no filings; owner outreach"],
        ["Labour cost trajectory 2026-2030", "Margin guard rails",
         "Statistics Norway, NHO Service og Handel forecasts"],
        ["Customer-side reference checks", "Validate switching costs and tendering dynamics",
         "5-10 housing-association and commercial property managers"],
        ["Sweden fastighetsskötsel landscape", "Cross-border optionality",
         "Separate scan (~3 weeks of work); Ocab introductions"],
        ["Anchor management team integration capability", "Roll-up depends on this; weak team is fatal",
         "Reference checks, prior bolt-on track record review"],
    ],
    widths_cm=[5.5, 5.5, 6.0])


# --------------------------------------------------------------------------------------
# 10. CONCLUDING VIEW
# --------------------------------------------------------------------------------------
h1(doc, "10.  Concluding View")

body(doc,
    "This is a roll-up construction, not the acquisition of a finished #2 platform. "
    "The market thesis is correct: Nordic property maintenance is structurally "
    "attractive, fragmented, recurring, and regulation-shielded, with PHM at <5% share "
    "leaving the majority of the market open. Norvestor's PHM trajectory has de-risked "
    "the playbook.")

body(doc,
    "The honest weakness is execution dependence. The construction relies on (a) landing "
    "a founder-led anchor at disciplined entry, and (b) closing 25-40 bolt-ons across a "
    "supply set where PHM is the prime competing acquirer. The right-to-win narrative "
    "(commercial / mid-Norway / pan-Nordic tilt, Nimlas and Ocab adjacencies) is real, "
    "but it must be operationalised - not over-claimed in IC.")

body(doc,
    "Decision sought from screening committee: approval to dedicate one full-time "
    "origination resource and a modest sector commercial diligence budget "
    "(NOK 1.5-2.5m external) over Q3-Q4 2026. Mandatory re-screening checkpoint "
    "at end-Q4 2026 conditional on management access at one of the three named anchor "
    "candidates and validated bolt-on pipeline of NOK 250m+ revenue for years 1-2.")


# --------------------------------------------------------------------------------------
# APPENDICES
# --------------------------------------------------------------------------------------
h1(doc, "Appendix A - Screening Universe")

bullet(doc, "Total Norwegian entities screened: 75 verified names (post-trim from 140).")
bullet(doc, "Anchor candidates: 3 (Vaktmesterkompaniet, Din Vaktmester, Insider Group).")
bullet(doc, "Tier-1 named bolt-ons (NOK 100-400m revenue): 15.")
bullet(doc, "Tier-2 bolt-ons (NOK 30-100m revenue): 35-50 across the screen.")
bullet(doc, "Companion file: Norway_FM_Property_Services_PE_Screen.xlsx (4 tabs: Company overview, Financials, Fit vs PHM, PE angle).")

h1(doc, "Appendix B - Key Assumptions and Limitations")

bullet(doc,
    "Revenue and EBITDA figures for non-listed AS-level entities are largely analyst "
    "estimates triangulated from BoldData FY23 top-50 ranking, sector margin medians "
    "(6-9% cleaning, 8-10% janitor), and applied CAGR. Verification in Proff.no / "
    "Brønnøysundregistrene required before any IOI.")
bullet(doc,
    "TAM ranges reflect analyst triangulation against the brief's NOK 15-25bn baseline; "
    "not externally validated.")
bullet(doc,
    "PHM Norway revenue and EBITDA inferred from Norwegian subsidiary filings; PHM Group "
    "does not publicly disclose Norwegian segment results.")
bullet(doc,
    "Exit multiple ranges based on observed transactions in Nordic FM / services "
    "2021-2024 (PHM bolt-ons, 4Service to Compass, Cares to Equip); subject to market "
    "conditions.")
bullet(doc,
    "KLAR 4KQ, Triple A, and Key Characteristics Assessment framework applications "
    "are the deal team's interpretation pending partner review.")

h1(doc, "Appendix C - Glossary (Norwegian terms)")

bullet(doc, "Vaktmester - janitor / building caretaker.")
bullet(doc, "Eiendomsdrift - property operations.")
bullet(doc, "Fastighetsskötsel - (Swedish) property care / janitorial.")
bullet(doc, "Borettslag - housing cooperative; dominant residential ownership form in Norway.")
bullet(doc, "Sameie - condominium-style co-ownership.")
bullet(doc, "Allmenngjøring(sloven) - law extending collective wage agreement application to all workers in a sector; creates the labour-cost floor that disadvantages informal operators.")
bullet(doc, "Renholdsoverenskomsten - the collective wage agreement for cleaning services.")
bullet(doc, "NACE 81.10 / 81.21 - statistical sector codes for combined facilities support / general building cleaning.")


# --------------------------------------------------------------------------------------
# Save
# --------------------------------------------------------------------------------------
out_path = "/home/user/Pilot/output/Nordic_Property_Maintenance_Thesis.docx"
doc.save(out_path)
print(f"Saved: {out_path}")
