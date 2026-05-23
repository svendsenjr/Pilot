"""
Nordic Property Maintenance Roll-Up - Screening Memo.

Single thesis: PROCEED on the Norway-led #2 platform construction.
Clean structure (six sections), one returns model, one anchor view.
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


NAVY = RGBColor(0x1F, 0x3A, 0x5F)
DARK_GREY = RGBColor(0x33, 0x33, 0x33)
MID_GREY = RGBColor(0x66, 0x66, 0x66)
LIGHT_BG = "F2F4F7"
CALLOUT_BG = "E8F0FA"
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


def h1(doc, text):
    add_para(doc, text, size=15, bold=True, color=NAVY,
             space_before=14, space_after=4)


def h2(doc, text):
    add_para(doc, text, size=11.5, bold=True, color=DARK_GREY,
             space_before=8, space_after=2)


def body(doc, text):
    add_para(doc, text, size=10.5, color=DARK_GREY, space_after=4)


def bullet(doc, text):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.left_indent = Cm(0.6)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.font.size = Pt(10.5)
    r.font.color.rgb = DARK_GREY


def callout(doc, text):
    table = doc.add_table(rows=1, cols=1)
    cell = table.rows[0].cells[0]
    cell.text = ""
    p = cell.paragraphs[0]
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = NAVY
    set_cell_bg(cell, CALLOUT_BG)
    set_cell_borders(cell)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


def build_table(doc, headers, rows, widths_cm=None):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.autofit = False
    table.allow_autofit = False

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

    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return table


# ====================================================================================
doc = Document()

for section in doc.sections:
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.8)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)

style = doc.styles['Normal']
style.font.name = 'Calibri'
style.font.size = Pt(10.5)


# ====================================================================================
# COVER
# ====================================================================================
add_para(doc, "KLAR PARTNERS - SCREENING MEMO", size=9, bold=True,
         color=MID_GREY, space_after=2)
add_para(doc, "Nordic Property Maintenance Roll-Up",
         size=22, bold=True, color=NAVY, space_after=2)
add_para(doc, "The #2 platform behind PHM Group",
         size=13, italic=True, color=MID_GREY, space_after=10)

meta = doc.add_table(rows=4, cols=2)
meta_data = [
    ("Sector", "Property maintenance / janitorial (vaktmester, eiendomsdrift, fastighetsskötsel)"),
    ("Geography", "Norway anchor; Nordic platform optionality"),
    ("Reference", "PHM Group (Norvestor VIII, vintage 2017)"),
    ("Recommendation", "PROCEED to Phase 2 origination on three named anchor candidates"),
]
for i, (k, v) in enumerate(meta_data):
    c0, c1 = meta.rows[i].cells
    c0.text = ""; c1.text = ""
    r0 = c0.paragraphs[0].add_run(k)
    r0.bold = True; r0.font.size = Pt(10); r0.font.color.rgb = NAVY
    r1 = c1.paragraphs[0].add_run(v)
    r1.font.size = Pt(10); r1.font.color.rgb = DARK_GREY
    c0.width = Cm(3.5); c1.width = Cm(13.5)

doc.add_paragraph()


# ====================================================================================
# 1. RECOMMENDATION
# ====================================================================================
h1(doc, "1.  Recommendation")

callout(doc,
    "Proceed to Phase 2 origination. Anchor on a founder-led Norwegian janitor "
    "business (NOK 250-1,000m revenue) and execute 20-30 bolt-ons over five "
    "years to reach NOK 2.0-2.5bn platform revenue. Base case 2.5-2.8x MoM, "
    "22-25% gross IRR. Critical path: secure anchor at <9.5x within 18 months.")

body(doc,
    "The Norwegian property maintenance market is fragmented enough to support a "
    "second consolidator. PHM at <5% share leaves NOK 14-24bn of unconsolidated "
    "market behind them; the supply set of acquirable mid-market operators is "
    "200+ named entities; the M&A playbook is validated by PHM's own seven-year "
    "trajectory. The investment case is not built on out-running PHM head-to-head "
    "but on absorbing a different slice of the same fragmented market - "
    "commercial-tilt, complementary geography (Mid-Norway / Vestland), and "
    "selective Nordic expansion.")

body(doc,
    "The gating risk is anchor access at a disciplined multiple. With three "
    "founder-led candidates in the pipeline and category multiples rebased to "
    "7-9x (from 11-13x in 2022), the entry environment is constructive. "
    "Push-backs and conditions in Section 4.")


# ====================================================================================
# 2. WHY THIS MARKET
# ====================================================================================
h1(doc, "2.  Why This Market")

build_table(doc,
    headers=["Test", "Read"],
    rows=[
        ["TAM (Norway)", "NOK 15-25bn addressable; NOK 55-80bn Nordic"],
        ["Growth", "4-6% nominal, 1-2% real - moderate but visible"],
        ["Fragmentation", "Top 5 players <25% share; PHM <5%; ~12,000 registered FM entities"],
        ["Recurring revenue", "70-85% under multi-year framework agreements"],
        ["Customer concentration", "Low - typical largest customer <5%"],
        ["Capex intensity", "<2% of revenue; working capital and small fleet only"],
        ["Cash conversion", "75-90% EBITDA-to-FCF"],
        ["Regulatory shield", "Allmenngjøringsloven creates labour-cost floor; disadvantages informal operators"],
        ["Cyclicality", "Low operationally; demand resilient through cycle (proven 2020)"],
    ],
    widths_cm=[4.0, 13.0])

h2(doc, "Why now")

bullet(doc,
    "Generational handover. Two-thirds of Norwegian mid-market FM businesses (NOK 100-500m revenue) were founded 1985-2000; founders now aged 60+.")
bullet(doc,
    "Multiples have rebased. Platform multiples 7-9x (down from 11-13x in 2022); bolt-on multiples 4-6x for sub-NOK 100m targets. Constructive entry.")
bullet(doc,
    "PHM has proven the playbook works in this geography. We are not testing whether the strategy is viable; we are executing a validated strategy with our own positioning.")


# ====================================================================================
# 3. THE PLATFORM CONSTRUCTION
# ====================================================================================
h1(doc, "3.  The Platform Construction")

h2(doc, "3.1  Positioning vs. PHM")

body(doc,
    "PHM is residential-focused (~70% mix), Oslo-concentrated, pure janitor / "
    "property maintenance. Our platform tilts commercial property (offices, retail, "
    "public sector buildings) and geographically anchors outside PHM's core "
    "Oslo / Akershus density. Same TAM, different slice. We compete with PHM "
    "for some bolt-ons; we do not need to win every one.")

h2(doc, "3.2  Five-year construction plan")

build_table(doc,
    headers=["Year", "Activity", "Revenue (NOKm)", "EBITDA margin", "EBITDA (NOKm)"],
    rows=[
        ["Y0", "Anchor acquisition; integration team in place", "300", "8.5%", "26"],
        ["Y1", "3-5 bolt-ons; route density Oslo + anchor region", "550", "9.0%", "50"],
        ["Y2", "5-7 bolt-ons including one mid-sized; procurement consolidation", "950", "9.5%", "90"],
        ["Y3", "5-7 bolt-ons; Nimlas cross-sell pilot; recap optionality", "1,400", "10.5%", "147"],
        ["Y4", "4-6 bolt-ons; Sweden probe via Ocab introductions", "1,800", "11.0%", "198"],
        ["Y5", "Exit preparation; tidying", "2,200", "11.5%", "253"],
    ],
    widths_cm=[1.5, 7.0, 2.5, 2.5, 3.5])

body(doc,
    "Anchor enters at 8.5-10x EBITDA. Bolt-ons blend at 5-6x EBITDA. Weighted-"
    "average acquisition multiple ~6.5x. Exit at 11-12x on NOK 250m EBITDA "
    "implies EV NOK 2.8-3.0bn. The buy-vs-build spread of 4.5-5.5 turns is the "
    "core return engine; organic growth and margin expansion are secondary.")

h2(doc, "3.3  Margin expansion path (8.5% to 11.5%)")

build_table(doc,
    headers=["Lever", "Uplift", "Confidence"],
    rows=[
        ["Route density (Oslo, Bergen, Trondheim corridors)", "+100-150 bps", "High; direct PHM precedent"],
        ["Procurement consolidation (chemicals, fleet, insurance)", "+50-80 bps", "High; mechanical via bolt-ons"],
        ["Digital workforce scheduling and dispatch", "+30-50 bps", "Medium; implementation cost NOK 15-25m"],
        ["Nimlas-adjacent technical service cross-sell", "+20-40 bps", "Medium; requires capability transfer"],
        ["Back-office consolidation post-PMI", "+20-40 bps", "High; standard playbook"],
        ["Total margin expansion", "+220-360 bps", "Mid-range 280-300 bps is the underwrite"],
    ],
    widths_cm=[7.5, 3.0, 6.5])


# ====================================================================================
# 4. PUSHBACKS AND CONDITIONS
# ====================================================================================
h1(doc, "4.  Pushbacks and Conditions")

body(doc,
    "Not a layup. Five real pushbacks, each with our response and the test that "
    "would either confirm or kill the relevant assumption.")

build_table(doc,
    headers=["Pushback", "Response", "Test"],
    rows=[
        ["PHM has a nine-year head start - they own the deal flow",
         "True for residential / Oslo. Our positioning (commercial + Mid-Norway / Vestland) is structurally different. We will lose some deals to PHM; we do not need to win them all",
         "Track bolt-on competitive dynamics in Y1-Y2; if hit rate <40% on targeted deals, slow the M&A pace"],
        ["Norvestor exits PHM 2026-2027; new owner has more capital",
         "Plausible but not certain. New owner may also focus on integration over M&A in Y1-Y2 post-acquisition. And market is large enough for both to grow",
         "Monitor PHM M&A cadence Q3 2026 onwards; trigger re-pricing if PHM closes >10 deals/yr"],
        ["Bolt-on multiples drift up as competition intensifies",
         "Possible. Our model uses 5-6x blended; downside case at 7x compresses IRR by ~250 bps but stays above 18%",
         "Quarterly transaction comp tracking; renegotiate pipeline if multiples breach 7.5x"],
        ["Right to win is generic Nordic services playbook",
         "True at headline level. Specific advantages: Nimlas technical bundling (30-50 bps EBITDA), Ocab Swedish footprint for cross-Nordic intel and customer introductions, KLAR operating partner bench",
         "Articulated in Phase 2 management pitches; refined with anchor management team's input"],
        ["Margin expansion of 280-300 bps is aggressive for labour services",
         "Within range of PHM's own margin trajectory (estimated 6-7% at anchor in 2017 to 10-12% today). Standard PE roll-up arithmetic",
         "Operational diligence on density and procurement opportunity at the anchor"],
    ],
    widths_cm=[4.5, 7.5, 5.0])


# ====================================================================================
# 5. RETURNS
# ====================================================================================
h1(doc, "5.  Returns")

build_table(doc,
    headers=["Parameter", "Base case", "Upside", "Downside"],
    rows=[
        ["Anchor revenue at entry (NOKm)", "300", "750", "300"],
        ["Anchor entry multiple", "9.5x", "9.5x", "10.5x"],
        ["Anchor entry EV (NOKm)", "245", "615", "270"],
        ["Bolt-ons closed Y1-Y5", "22-26", "30+", "12-15"],
        ["Blended bolt-on multiple", "6.0x", "5.5x", "7.0x"],
        ["Bolt-on EV deployed Y1-Y5 (NOKm)", "950", "1,400", "560"],
        ["Y5 platform revenue (NOKm)", "2,200", "3,000", "1,400"],
        ["Y5 EBITDA margin", "11.5%", "12.5%", "10.0%"],
        ["Y5 EBITDA (NOKm)", "253", "375", "140"],
        ["Exit multiple", "11.5x", "13.0x", "9.5x"],
        ["Exit EV (NOKm)", "2,910", "4,875", "1,330"],
        ["Total invested equity (NOKm)", "900", "1,250", "750"],
        ["Gross MoM", "2.5 - 2.8x", "3.2 - 3.5x", "1.5 - 1.8x"],
        ["Gross IRR", "22 - 25%", "28 - 32%", "8 - 12%"],
    ],
    widths_cm=[7.0, 3.5, 3.5, 3.0])

body(doc,
    "Base case clears the platform-slot hurdle (>2.5x, >20% IRR). Downside is "
    "capital-preservation territory but not capital destruction - the asset's "
    "recurring-revenue base and cash conversion provide structural floor.")


# ====================================================================================
# 6. ANCHOR AND BOLT-ON PIPELINE
# ====================================================================================
h1(doc, "6.  Anchor and Bolt-On Pipeline")

h2(doc, "6.1  Anchor candidates")

build_table(doc,
    headers=["Anchor", "HQ", "Rev FY24E (NOKm)", "Owner", "Read"],
    rows=[
        ["Vaktmesterkompaniet AS", "Oslo", "~1,000", "Founder",
         "Closest PHM analog; pure janitor model; ~10 small bolt-ons completed. Risk: PHM likely in dialogue; multi-bidder process; entry at 9-11x"],
        ["Din Vaktmester AS", "Trondheim", "250-350", "Founder",
         "Preferred starting point for an asymmetric thesis. Mid-Norway base PHM has not built; pure culture; lower competing-bidder intensity; smaller cheque allows opportunistic mid-sized bolt-ons in Y1-Y2"],
        ["Insider Group AS", "Høvik", "1,000-1,100", "Family",
         "Recent Mitie Norway carve-out (2024) demonstrates M&A muscle. Caveat: mix is >50% cleaning rather than property maintenance; need to validate property-maintenance share before anchor decision"],
    ],
    widths_cm=[4.0, 2.5, 2.3, 1.7, 6.5])

h2(doc, "6.2  Tier-1 named bolt-ons")

build_table(doc,
    headers=["Target", "HQ", "Rev FY24E (NOKm)"],
    rows=[
        ["Bygårdsservice AS", "Oslo", "80-120"],
        ["Aktiv Eiendomsdrift AS", "Roa (Innlandet)", "60-100"],
        ["Eiendomspartner 1 AS", "Bergen", "150-220"],
        ["Ability FM / Ability Gruppen", "Bergen", "250-400"],
        ["Toma Eiendomsdrift AS (potential carve-out)", "Oslo", "200-300"],
        ["Facilitec AS (Insider sub - if not anchor)", "Lørenskog", "200-300"],
        ["BBL Daglig Drift AS", "Oslo", "100-150"],
        ["Conluo Facility Services AS", "Oslo", "150-250"],
        ["Hvass AS", "Oslo", "60-90"],
        ["Bygård Vaktmesteren AS", "Oslo", "50-80"],
        ["RSV Gruppen AS", "Bergen", "60-100"],
        ["Resolve AS", "Risør", "50-80"],
        ["Total Eiendomsservice AS", "Drammen", "50-90"],
        ["Northroads AS", "Mo i Rana", "80-130"],
        ["Allianse Service Partner AS", "Oslo", "100-160"],
    ],
    widths_cm=[6.0, 4.0, 3.0])

body(doc,
    "Tier-1 named bolt-ons aggregate to NOK 1.6-2.6bn revenue. The tail "
    "(35-50 small operators at NOK 30-100m revenue) adds another NOK 1.5-3.0bn. "
    "Net of PHM's parallel absorption pace (5-8 deals/yr), our realistic Y1-Y5 "
    "pipeline of 22-26 closed bolt-ons is well-covered.")


# ====================================================================================
# 7. KLAR FRAMEWORK SCORING
# ====================================================================================
h1(doc, "7.  KLAR Framework Scoring")

build_table(doc,
    headers=["Framework", "Score", "Rationale"],
    rows=[
        ["4KQ - Why this market?", "Pass",
         "Fragmented, recurring, regulation-shielded, downside-protected. Standard KLAR sector profile"],
        ["4KQ - Why this asset (construction)?", "Pass, conditional on anchor",
         "Three credible anchors; bolt-on supply ample. Gating risk is access at disciplined multiple"],
        ["4KQ - Why now?", "Pass",
         "Founder handover wave; multiples rebased; PHM exit reshapes (not removes) competitive dynamic"],
        ["4KQ - Why KLAR?", "Acceptable",
         "Nimlas technical bundling and Ocab Swedish footprint give a real, if modest, differentiation. Generic Nordic services playbook is the baseline"],
        ["Triple A - Attractive market", "A", "Sector profile clean"],
        ["Triple A - Attractive business", "A-", "Strong on recurring revenue and cash conversion; ceiling on real pricing power"],
        ["Triple A - Attractive entry / exit", "A-", "Entry environment constructive; sponsor-to-sponsor exit base case (6-10 credible buyers)"],
        ["KCA composite", "3.9 / 5.0",
         "B+ asset, A- platform via consolidation. Detail in Appendix A"],
    ],
    widths_cm=[5.0, 3.0, 9.0])


# ====================================================================================
# 8. PATH TO IC
# ====================================================================================
h1(doc, "8.  Path to IC")

build_table(doc,
    headers=["Workstream", "Owner", "Target", "Decision gate"],
    rows=[
        ["Anchor access - Din Vaktmester (preferred)", "[Origination]", "Q3 2026",
         "Founder posture on sale within 24 months; mgmt integration capability"],
        ["Anchor access - Vaktmesterkompaniet", "[Origination]", "Q3 2026",
         "Confirm whether PHM in advanced dialogue; valuation expectation"],
        ["Anchor access - Insider Group", "[Origination]", "Q4 2026",
         "Validate property-maintenance share of mix post-Mitie"],
        ["Sector commercial diligence (third party, scoped)", "[Deal team]", "Q4 2026",
         "TAM validation; PHM positioning; bolt-on multiple discipline"],
        ["Bolt-on pipeline diligence (top 15 named)", "[Deal team + analyst]", "Q4 2026",
         "Confirm Y1-Y2 acquired-revenue pipeline >NOK 250m"],
        ["Sweden optionality scan via Ocab", "[Stockholm]", "Q4 2026",
         "Establish Y3+ Nordic extension feasibility"],
        ["Regulatory diligence (allmenngjøring, wages)", "[Legal counsel]", "Q1 2027",
         "No 24-month step-change risk"],
    ],
    widths_cm=[6.0, 2.8, 2.2, 6.0])

h2(doc, "Conditions to advance to full IC")

bullet(doc, "Management access secured on at least one anchor candidate.")
bullet(doc, "Indicative anchor entry multiple <9.5x.")
bullet(doc, "Validated Y1-Y2 bolt-on pipeline of NOK 250m+ acquired revenue.")
bullet(doc, "Third-party TAM and competitive validation.")

h2(doc, "Resourcing ask")

body(doc,
    "One full-time origination resource for Q3-Q4 2026; external sector "
    "diligence budget NOK 1.5-2.5m. Re-screening checkpoint end-Q4 2026 on "
    "conditions above. No fund-slot commitment until full IC.")


# ====================================================================================
# APPENDIX
# ====================================================================================
h1(doc, "Appendix - Sources, Assumptions, Glossary")

h2(doc, "Sources")
bullet(doc,
    "Screening universe of 75 verified Norwegian property service / FM / cleaning "
    "entities compiled from BoldData FY23 top-50 ranking, PE press releases, "
    "Konkurransetilsynet filings, PHM Annual Report 2024, Proff.no / "
    "Brønnøysundregistrene. Full screen in Norway_FM_Property_Services_PE_Screen.xlsx.")
bullet(doc,
    "PHM Group financials (group revenue EUR 600-800m; Norway sub NOK 1.5-2.5bn) "
    "inferred from Norvestor commentary and Norwegian subsidiary filings.")
bullet(doc,
    "Bolt-on multiple ranges (4-6x sub-NOK 100m; 6-8x sub-NOK 200m) inferred from "
    "PHM acquisition press releases and Nordic FM transaction comps 2021-2024.")

h2(doc, "Assumptions")
bullet(doc,
    "TAM NOK 15-25bn Norway and NOK 55-80bn Nordic per the brief baseline, "
    "triangulated bottom-up against NACE 81.10/81.21 entity counts.")
bullet(doc,
    "Returns model is illustrative; per-deal underwriting requires anchor-specific assumptions.")
bullet(doc,
    "KLAR framework applications reflect the deal team's interpretation, pending partner review.")

h2(doc, "Glossary")
bullet(doc, "Vaktmester - janitor / building caretaker.")
bullet(doc, "Eiendomsdrift - property operations.")
bullet(doc, "Fastighetsskötsel - (Swedish) property care / janitorial.")
bullet(doc, "Borettslag - housing cooperative.")
bullet(doc, "Sameie - condominium-style co-ownership.")
bullet(doc, "Allmenngjøring(sloven) - law extending collective wage agreement application.")
bullet(doc, "Renholdsoverenskomsten - collective wage agreement for cleaning services.")
bullet(doc, "PMI - post-merger integration.")


# Save
out_path = "/home/user/Pilot/output/Nordic_Property_Maintenance_Thesis.docx"
doc.save(out_path)
print(f"Saved: {out_path}")
