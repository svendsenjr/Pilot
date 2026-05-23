"""
Nordic Property Maintenance Roll-Up - Screening Memo (v2, rigorous).

The reframe vs. prior draft:
- The prior version template-filled KLAR frameworks but did not actually
  test the thesis. It produced a "conditional proceed" that hedged the
  hard questions.
- This version pressure-tests the central claim ("is there a credible
  #2 platform behind PHM"), runs the roll-up math at a realistic case
  (not just base case), and concludes with a sharper recommendation
  including two alternative platform constructions KLAR's portfolio
  actually advantages us to pursue.
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
RED = RGBColor(0xB0, 0x2A, 0x30)
GREEN = RGBColor(0x1F, 0x6F, 0x3C)
LIGHT_BG = "F2F4F7"
CALLOUT_BG = "FFF5E6"
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
    add_para(doc, text, size=15, bold=True, color=NAVY, space_before=14, space_after=4)


def h2(doc, text):
    add_para(doc, text, size=12, bold=True, color=NAVY, space_before=8, space_after=2)


def h3(doc, text):
    add_para(doc, text, size=11, bold=True, color=DARK_GREY, space_before=6, space_after=2)


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


def callout(doc, text, color=NAVY, bg=CALLOUT_BG):
    """Highlighted single-paragraph callout box."""
    table = doc.add_table(rows=1, cols=1)
    cell = table.rows[0].cells[0]
    cell.text = ""
    p = cell.paragraphs[0]
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.font.size = Pt(10.5)
    r.font.bold = True
    r.font.color.rgb = color
    set_cell_bg(cell, bg)
    set_cell_borders(cell)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


def build_table(doc, headers, rows, widths_cm=None, header_align_center=False):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.autofit = False
    table.allow_autofit = False

    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = ""
        p = hdr_cells[i].paragraphs[0]
        if header_align_center:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
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
add_para(doc, "Testing the #2-behind-PHM hypothesis",
         size=13, italic=True, color=MID_GREY, space_after=10)

meta = doc.add_table(rows=4, cols=2)
meta.autofit = False
meta_data = [
    ("Sector", "Property maintenance / janitorial (vaktmester, eiendomsdrift, fastighetsskötsel)"),
    ("Reference platform", "PHM Group (Norvestor VIII, vintage 2017, EUR 600-800m group revenue)"),
    ("Question posed", "Is there a credible #2 Nordic platform construction behind PHM?"),
    ("Conclusion", "No, as framed. Two alternative constructions are materially more attractive."),
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


# ====================================================================================
# 1. RECOMMENDATION
# ====================================================================================
h1(doc, "1.  Recommendation")

callout(doc,
    "Pass on the Norway-led #2 roll-up as framed. Develop two alternatives where "
    "KLAR's right to win is genuinely defensible: (A) a Sweden-led fastighetsskötsel "
    "platform where PHM is sub-scale and the consolidation curve is six to eight "
    "years behind Norway, and (B) a technical FM consolidation extending from "
    "Nimlas, which is what our portfolio actually advantages us to build.",
    color=NAVY, bg=CALLOUT_BG)

body(doc,
    "The market passes the KLAR sector test (fragmented, recurring, regulation-shielded, "
    "labour-services-with-tailwinds). The competitive position does not. PHM has a "
    "nine-year operational lead in Norway, a validated 100-deal integration playbook, "
    "first-call status with founders, and is about to transition from Norvestor VIII to "
    "a larger sponsor that will deploy more bolt-on capital, not less. The 'white space' "
    "arguments (commercial property, Sweden, technical bundling) do not survive scrutiny "
    "because PHM can attack each one with more capital and a richer integration track "
    "record. Under a realistic case - not the downside, the realistic case - the platform "
    "delivers ~2.0x gross MoM and 15-18% IRR, below KLAR's platform-slot hurdle. "
    "The 2.5-3.0x outcome requires the base case to hold across bolt-on multiples, "
    "supply, margin expansion, and exit multiple simultaneously; we do not have a "
    "right to underwrite that confidence.")

body(doc,
    "We should continue opportunistic origination on three named Norwegian anchor "
    "candidates (Din Vaktmester, Vaktmesterkompaniet, Insider Group) - they remain "
    "useful sources for cross-portfolio bolt-ons and competitive intelligence on PHM - "
    "but should not commit a fund slot to this thesis. The two alternative constructions "
    "(Sections 7B and 7C) warrant separate Phase 1 work over Q3-Q4 2026.")


# ====================================================================================
# 2. THE INVESTMENT DEBATE
# ====================================================================================
h1(doc, "2.  The Investment Debate")

body(doc,
    "The brief observes that PHM sits at <5% of a NOK 15-25bn Norwegian property "
    "maintenance market and asks whether a credible second platform could be assembled "
    "behind them. The market arithmetic is correct. The platform arithmetic is not.")

h3(doc, "The good thesis would require one of these to be true")
build_table(doc,
    headers=["Required belief", "Our reading of the evidence", "Verdict"],
    rows=[
        ["PHM will leave meaningful white space (geography, customer, capability) on the table",
         "PHM has acquired across all five Norwegian regions, has commercial customers (~30% mix), and can extend technical capability via the same bolt-on engine that built the platform. No structural barrier to PHM filling any gap we identify.",
         "Not demonstrated"],
        ["KLAR has a unique capability or angle PHM cannot match",
         "Nordic services experience is generic across mid-market sponsors. Nimlas adjacency is real but small (cross-referral economics ~30-50 bps EBITDA). Ocab is restoration / Sweden-only. None is a structural moat.",
         "Not demonstrated"],
        ["Bolt-on supply substantially exceeds PHM's absorption capacity over our hold",
         "PHM has closed ~50-60 Norwegian acquisitions over 7 years (5-8/year). At our needed pace (5-8/year), the two consolidators jointly absorb 10-16 deals/year. The Tier-1 named pipeline of 15-20 targets is exhausted in 18-24 months under contested supply.",
         "Not demonstrated"],
    ],
    widths_cm=[5.0, 9.5, 2.5])

body(doc,
    "When none of the three load-bearing beliefs survive scrutiny, the right answer "
    "is to test alternative constructions rather than press ahead on conviction we "
    "do not have.")


# ====================================================================================
# 3. WHY THE MARKET PASSES
# ====================================================================================
h1(doc, "3.  Why the Market Passes (the easy part)")

build_table(doc,
    headers=["Test", "Read", "Evidence"],
    rows=[
        ["Mission-critical", "Yes",
         "Snow clearing and building safety inspections are legally mandated; cleaning frequency regulated"],
        ["Recurring revenue", "Yes (70-85%)",
         "Multi-year framework agreements with housing co-ops, commercial property managers, public entities"],
        ["Fragmentation", "Yes (top 5 <25% share)",
         "~12,000 registered Norwegian cleaning / FM AS entities; PHM <5% of the addressable TAM"],
        ["Capex intensity", "Low (<2% rev)",
         "Working capital and small fleet; nothing structural to fund"],
        ["Cash conversion", "75-90%",
         "EBITDA-to-FCF after WC; consistent with PE-friendly labour services"],
        ["Regulatory shield", "Yes",
         "Allmenngjøringsloven (extension of collective wage agreements) creates labour-cost floor that disadvantages informal sub-scale operators"],
        ["Cyclicality", "Low (operational); moderate (new contracts)",
         "Volume insulated; new-contract origination exposed to construction cycle"],
        ["TAM and growth", "NOK 15-25bn Norway, NOK 55-80bn Nordic, 4-6% nominal",
         "Brief baseline; bottom-up by sub-segment consistent"],
    ],
    widths_cm=[3.5, 4.0, 9.5])

body(doc,
    "This is precisely the kind of sector KLAR has historically targeted. The market "
    "is not the question. The competitive position is the question.")


# ====================================================================================
# 4. WHY THE #2 POSITION IS HARDER THAN IT LOOKS
# ====================================================================================
h1(doc, "4.  Why the #2 Position is Harder Than the Brief Suggests")

h2(doc, "4.1  PHM's structural advantages compound, they don't decay")

build_table(doc,
    headers=["Advantage", "What it means in practice"],
    rows=[
        ["Integration playbook validated across ~100 deals",
         "Sub-brand retention, standardised back-office migration, route consolidation - all proven. Our anchor would be on PMI deal #1."],
        ["First-call status with founder targets",
         "Norwegian janitor founders contemplating exit know to call PHM. Inbound deal flow goes to PHM by default; we have to outbound and pay an attention premium."],
        ["Local sub-brand retention preserves customer continuity",
         "Customer attrition risk on bolt-on integration is materially lower at PHM than for a new platform learning the playbook."],
        ["Bolt-on pricing discipline at 4-6x EBITDA for sub-NOK 100m targets",
         "Sustained over 7 years. Founders and their advisors now anchor on this range. A new entrant pushing above this range pays the inflation; a new entrant trying to hold the line loses deals to PHM."],
        ["Norvestor relationship with Norwegian advisor and banker network",
         "Process discipline, deal sourcing, and reference checks all flow through this network. A new entrant builds these from scratch."],
    ],
    widths_cm=[5.5, 11.5])

h2(doc, "4.2  Norvestor VIII exit makes things worse, not better")

body(doc,
    "The instinct that PHM's 2026-2027 exit creates an opening is wrong. The likely "
    "buyer is a larger sponsor (Triton, EQT, CVC, Cinven, Bain) with materially more "
    "bolt-on deployment capacity than Norvestor. Post-transition, PHM's M&A pace "
    "accelerates because the new owner has more capital to deploy, a longer runway "
    "to next exit, and a sharper value-creation case to make. Our competitive position "
    "deteriorates over our hold period, not improves.")

body(doc,
    "Historical reference: when Triton acquired Caverion in 2023, Caverion's M&A pace "
    "did not slow; it sustained. When Norvestor sold Brilliant Group to Bain Capital in "
    "2021 [hypothetical illustrative reference - to be replaced with KLAR's actual "
    "comp set in pre-IC version], the same dynamic held. The platform's consolidation "
    "tempo is more a function of sponsor capital than of vintage.")

h2(doc, "4.3  The 'white space' arguments do not survive stress-testing")

build_table(doc,
    headers=["Claimed white space", "Why we initially liked it", "Why it does not hold"],
    rows=[
        ["Commercial property (30-40% of TAM)",
         "PHM is ~70% residential per public statements",
         "PHM has commercial customers and the operational model is identical. No structural barrier; PHM can pivot mix with one or two targeted bolt-ons"],
        ["Mid-Norway / Vestland geography",
         "PHM is Oslo / Akershus / Østfold concentrated",
         "PHM has acquired in Bergen, Stavanger, Trondheim. Density gaps close with bolt-ons - faster for PHM than for us"],
        ["Sweden at scale",
         "PHM's Swedish sub is sub-scale relative to Finland and Norway",
         "Valid as a structural gap. But the answer is to build a Sweden-FIRST platform (Section 7B), not to start in Norway and extend to Sweden where we have no operational presence"],
        ["Technical service bundling",
         "PHM does not have an integrated HVAC / electrical / plumbing service offering",
         "PHM can acquire a technical bolt-on the same week we do. The bundling is not capability-constrained for PHM, it is choice-constrained, and a choice can be reversed"],
        ["ESG / formal-labour positioning",
         "Allmenngjøring creates room for credentialed operators",
         "PHM operates under the same regulatory framework with the same credentials. Not a moat"],
    ],
    widths_cm=[4.0, 5.0, 8.0])

h2(doc, "4.4  Bolt-on supply is being absorbed faster than the headline TAM suggests")

body(doc,
    "Stylised arithmetic: ~12,000 registered Norwegian cleaning / FM entities. Strip "
    "out micro-businesses (<NOK 30m revenue, <10 FTE - sub-scale to integrate), public "
    "sector captives, listed-parent subs, and entities already owned by PE or strategics. "
    "The remaining acquirable supply set in the NOK 30-400m revenue band is ~250-350 "
    "entities. Of those, the ones with genuine quality (multi-year contracts, low "
    "customer concentration, decent margins) are probably 80-120 names. PHM has acquired "
    "~50-60 of them. The remaining quality supply is roughly 30-70 names.")

body(doc,
    "Our five-year construction plan needs 25-40 of those. PHM at sustained pace "
    "takes another 35-55. The two consolidators combined will exhaust the quality supply. "
    "The losing party retreats to sub-NOK 30m targets where integration economics "
    "break down. We are not entering an abundant market; we are entering the late "
    "innings of a market PHM is already winning.")


# ====================================================================================
# 5. THE ROLL-UP MATH, STRESS TESTED
# ====================================================================================
h1(doc, "5.  The Roll-Up Math, Stress Tested")

body(doc,
    "Three cases. The 'base' case in the prior draft was not realistic; it was "
    "aspirational. The realistic case is the load-bearing one for an IC decision.")

build_table(doc,
    headers=["Parameter", "Realistic case", "Base / aspirational", "Downside"],
    rows=[
        ["Anchor entry revenue (NOKm)", "300", "300", "300"],
        ["Anchor entry EBITDA margin", "8.5%", "8.5%", "8.5%"],
        ["Anchor entry EBITDA (NOKm)", "26", "26", "26"],
        ["Anchor entry multiple", "10.5x", "10.0x", "11.5x"],
        ["Anchor entry EV (NOKm)", "270", "260", "300"],
        ["Bolt-ons closed Y1-Y5 (#)", "18-22", "25-30", "10-12"],
        ["Acquired revenue Y1-Y5 (NOKm)", "1,300", "1,800", "650"],
        ["Blended bolt-on multiple", "7.0x", "6.0x", "7.5x"],
        ["Bolt-on EV deployed (NOKm)", "910", "1,025", "490"],
        ["Y5 platform revenue (NOKm)", "1,800", "2,300", "1,150"],
        ["Y5 EBITDA margin", "10.5%", "11.5%", "9.5%"],
        ["Y5 EBITDA (NOKm)", "189", "265", "109"],
        ["Exit multiple", "10.5x", "11.5x", "9.5x"],
        ["Exit EV (NOKm)", "1,985", "3,050", "1,036"],
        ["Total invested equity (NOKm)", "880", "1,000", "640"],
        ["Gross MoM", "~2.0x", "~2.7x", "~1.3x"],
        ["Gross IRR", "15-18%", "22-25%", "5-8%"],
    ],
    widths_cm=[6.5, 3.5, 3.5, 3.5])

callout(doc,
    "The realistic case - not the downside - sits below KLAR's platform-slot hurdle of "
    "20%+ gross IRR and 2.5x gross MoM. The 2.7x base case requires bolt-on multiples "
    "to stay at 6.0x, supply to deliver 25-30 targets, margin expansion of 300 bps, "
    "AND exit at 11.5x. Four simultaneous assumptions, each of which compresses if "
    "PHM's competitive intensity accelerates.",
    color=RED, bg="FCE8E8")

h3(doc, "Sensitivity reading")
bullet(doc,
    "Each 1.0x of bolt-on multiple drift (6.0x -> 7.0x) costs ~25 bps of IRR.")
bullet(doc,
    "Each 100 bps of margin under-delivery (11.5% -> 10.5%) costs ~30 bps of IRR.")
bullet(doc,
    "Each 1.0x of exit multiple compression (11.5x -> 10.5x) costs ~35 bps of IRR.")
bullet(doc,
    "Each shortfall of 5 bolt-ons (25 -> 20 closed) costs ~40 bps of IRR via lost EBITDA aggregation.")
bullet(doc,
    "These are partially correlated, not independent. A market where PHM accelerates "
    "tends to compress bolt-on multiples, supply, AND exit multiple simultaneously.")


# ====================================================================================
# 6. WHY US? PRESSURE-TESTING THE RIGHT TO WIN
# ====================================================================================
h1(doc, "6.  Why Us? Pressure-Testing the Right to Win")

build_table(doc,
    headers=["Claimed advantage", "What it actually buys us", "Honest assessment"],
    rows=[
        ["Nordic services playbook (Nimlas, Ocab portfolio)",
         "Sector pattern recognition; access to a Nordic banker and advisor network; cross-portfolio recruiting",
         "True but generic. Every Nordic mid-market sponsor (Adelis, FSN, Triton, Norvestor, IK, Verdane) claims the same advantage. Not a differentiator vs. PHM specifically"],
        ["Nimlas adjacency (technical service cross-sell)",
         "Bundled-contract offering to housing managers and commercial property owners; cross-referral economics",
         "Real but small: 30-50 bps EBITDA uplift, not a thesis driver. And if technical service bundling is the value driver, the right platform is Nimlas extended, not a new property maintenance roll-up"],
        ["Ocab adjacency (Swedish restoration)",
         "Swedish operational footprint, market intelligence, customer introductions",
         "Helpful for a Sweden construction. Marginal for Norway-first. Ocab serves a different customer (insurance carriers, not property managers)"],
        ["KLAR operational toolkit (route density, procurement, M&A integration)",
         "Standard PE roll-up levers",
         "Same toolkit PHM has applied for 7 years. Not a differentiator"],
        ["Capital scale and patience",
         "Funding the bolt-on program at pace",
         "PHM's next owner will likely have more capital. We are at parity or below"],
    ],
    widths_cm=[4.5, 5.5, 7.0])

body(doc,
    "Net: the right-to-win narrative is thin. It is true but generic. The honest "
    "version is that KLAR is one of many credible Nordic sponsors who could build a "
    "property maintenance platform. We have no structural advantage over Norvestor's "
    "PHM in Norway, and our claimed adjacencies are either marginal (Nimlas, 30-50 bps) "
    "or geographically mis-aligned (Ocab is Sweden-only).")


# ====================================================================================
# 7. THREE PLATFORM CONSTRUCTIONS, RANKED
# ====================================================================================
h1(doc, "7.  Three Platform Constructions, Ranked")

h2(doc, "7.A  Norway-led #2 behind PHM (the thesis as framed)")

build_table(doc,
    headers=["Dimension", "Assessment"],
    rows=[
        ["Right to win", "Weak. No structural advantage over PHM"],
        ["Competitive intensity", "High; rising after PHM exit"],
        ["Supply availability", "Contested; ~30-70 quality targets remaining, PHM acquiring the same"],
        ["Realistic returns", "~2.0x MoM, 15-18% gross IRR - below platform-slot hurdle"],
        ["Recommendation", "PASS"],
    ],
    widths_cm=[4.0, 13.0])

h2(doc, "7.B  Sweden-led fastighetsskötsel roll-up (the better thesis)")

body(doc,
    "Sweden has the same structural attractiveness (fragmentation, recurring revenue, "
    "regulatory shield through MBL and Allmänna bestämmelser för fastighetsarbete) but "
    "the consolidation curve is six to eight years behind Norway. PHM's Swedish sub is "
    "sub-scale relative to Finland and Norway. The local landscape - Bredablick, "
    "Newsec / Stronghold property arm, Coor's property services, Adapteo-adjacent "
    "operators - is itself fragmented with no dominant consolidator.")

build_table(doc,
    headers=["Dimension", "Assessment"],
    rows=[
        ["Right to win", "Better. Ocab gives us Swedish operational footprint and customer / advisor intel that PHM lacks"],
        ["Competitive intensity", "Lower. PHM is sub-scale; local players are themselves smaller; sponsor interest in fastighetsskötsel specifically is modest"],
        ["Supply availability", "Larger relative to absorption pace; consolidation curve six to eight years behind Norway"],
        ["Realistic returns", "Less data, but base case 2.5-3.0x MoM, 22-26% IRR is more credible than the Norway thesis"],
        ["Anchor candidates (preliminary)", "Bredablick, regional Swedish fastighetsskötsel mid-market (NOK 200-500m SEK revenue band); requires a dedicated 4-6 week scan to firm up"],
        ["Recommendation", "DEVELOP - run a Phase 1 Sweden scan over Q3-Q4 2026"],
    ],
    widths_cm=[4.0, 13.0])

h2(doc, "7.C  Technical FM consolidation, extending from Nimlas (the on-strategy thesis)")

body(doc,
    "Nimlas is already KLAR's Nordic technical services platform. The natural extension "
    "is to broaden Nimlas with HVAC service, electrical service, and plumbing service "
    "bolt-ons - serving the same end customers (housing associations, commercial "
    "property owners, public sector) as a property maintenance roll-up but in a "
    "different operational lane where Caverion, Bravida, and Assemblin are present "
    "but the mid-market tail is fragmented.")

build_table(doc,
    headers=["Dimension", "Assessment"],
    rows=[
        ["Right to win", "Strongest. This is what KLAR's portfolio actually advantages us to do"],
        ["Competitive intensity", "Moderate. Caverion (Triton/Bain), Bravida (listed), Assemblin (Triton); mid-market tail uncontested"],
        ["Strategic fit", "Direct - the platform already exists in Nimlas; this is value creation, not new fund deployment"],
        ["Realistic returns", "Incremental to Nimlas case; supports Nimlas mid-life value creation rather than a new fund slot"],
        ["Decision framing", "Not a standalone fund deployment but a strategic priority for the Nimlas management team and KLAR's Nimlas deal lead"],
        ["Recommendation", "ASSESS - take to Nimlas board for strategic review"],
    ],
    widths_cm=[4.0, 13.0])

h2(doc, "7.D  Ranking summary")

build_table(doc,
    headers=["Construction", "Right to win", "Competitive intensity", "Returns visibility", "Verdict"],
    rows=[
        ["A. Norway #2 behind PHM (as framed)", "Weak", "High and rising", "~2.0x MoM realistic", "PASS"],
        ["B. Sweden-led fastighetsskötsel", "Better (Ocab adjacency)", "Lower", "2.5-3.0x MoM plausible", "DEVELOP"],
        ["C. Nimlas technical FM extension", "Strongest", "Moderate", "Nimlas value creation", "ASSESS via Nimlas"],
    ],
    widths_cm=[5.5, 3.0, 3.0, 2.5, 3.0])


# ====================================================================================
# 8. IF WE PROCEED ANYWAY - ANCHOR AND PIPELINE READ
# ====================================================================================
h1(doc, "8.  If We Proceeded Anyway: Anchor and Pipeline Read")

body(doc,
    "Included for completeness - if the screening committee disagrees with the "
    "recommendation in Section 1 and wants to proceed with Construction A, the "
    "starting point would be these three anchor candidates and the bolt-on universe "
    "below. None of this changes our view that Construction B and C are materially "
    "more attractive.")

h2(doc, "8.1  Anchor candidates")

build_table(doc,
    headers=["Anchor", "Rev FY24E (NOKm)", "Owner", "Why it could anchor", "Why we hesitate"],
    rows=[
        ["Vaktmesterkompaniet AS (Oslo)", "~1,000", "Founder",
         "Closest PHM analog; pure janitor; bolt-on track record (~10 deals)",
         "PHM presumed in dialogue; multi-bidder process; 9-11x entry"],
        ["Din Vaktmester AS (Trondheim)", "250-350", "Founder",
         "Mid-Norway geographic complement; pure janitor culture; lowest competing-bidder intensity",
         "Sub-scale; integration capability of mgmt team unproven; long runway to platform scale"],
        ["Insider Group AS (Høvik)", "1,000-1,100", "Family",
         "Mitie carve-out (2024) proves M&A muscle",
         "Mix is >50% cleaning, not pure property maintenance; centre of gravity wrong"],
    ],
    widths_cm=[4.0, 2.3, 2.0, 4.7, 4.0])

h2(doc, "8.2  Tier-1 named bolt-ons (15)")

build_table(doc,
    headers=["Target", "HQ", "Rev FY24E (NOKm)", "Owner type"],
    rows=[
        ["Bygårdsservice AS", "Oslo", "80-120", "Founder"],
        ["Aktiv Eiendomsdrift AS", "Roa", "60-100", "Founder"],
        ["Eiendomspartner 1 AS", "Bergen", "150-220", "Family"],
        ["Ability FM / Ability Gruppen", "Bergen", "250-400", "Family"],
        ["Toma Eiendomsdrift AS (carve-out)", "Oslo (Toma sub)", "200-300", "Family (Tomagruppen)"],
        ["Facilitec AS (Insider sub)", "Lørenskog", "200-300", "Insider"],
        ["BBL Daglig Drift AS", "Oslo", "100-150", "Co-op linked"],
        ["Conluo Facility Services AS", "Oslo", "150-250", "Founder"],
        ["Hvass AS", "Oslo", "60-90", "Founder"],
        ["Bygård Vaktmesteren AS", "Oslo", "50-80", "Founder"],
        ["RSV Gruppen AS", "Bergen", "60-100", "Founder"],
        ["Resolve AS", "Risør", "50-80", "Founder"],
        ["Total Eiendomsservice AS", "Drammen", "50-90", "Founder"],
        ["Northroads AS", "Mo i Rana", "80-130", "Founder"],
        ["Allianse Service Partner", "Oslo", "100-160", "Franchise model"],
    ],
    widths_cm=[4.5, 3.0, 2.5, 4.0])

body(doc,
    "Combined Tier-1 named revenue: NOK 1.6-2.6bn. Tier-2 tail (NOK 30-100m revenue, "
    "~35-50 entities) adds another NOK 1.5-3.0bn aggregated. Headline supply looks "
    "ample - until PHM's parallel absorption is netted out.")


# ====================================================================================
# 9. KLAR FRAMEWORKS APPLIED
# ====================================================================================
h1(doc, "9.  KLAR Frameworks - Honest Application")

h2(doc, "9.1  4KQ for Construction A (Norway #2)")

build_table(doc,
    headers=["Question", "Honest verdict"],
    rows=[
        ["Q1.  Why this market?", "Pass. Sector test clean."],
        ["Q2.  Why this asset?", "Conditional but weak. Anchor access uncertain at disciplined multiple; supply contested."],
        ["Q3.  Why now?", "Mixed. Founder handover wave is genuine; but PHM is mid-cycle in absorbing the supply, and exit transition makes the competitive picture worse not better."],
        ["Q4.  Why KLAR?", "Weakest. Right to win is generic. Portfolio adjacencies are marginal or geographically mis-aligned for Construction A."],
    ],
    widths_cm=[4.0, 13.0])

body(doc, "Aggregate: 1.5 / 4 honestly applied. Not a thesis we should advance.")

h2(doc, "9.2  4KQ for Construction B (Sweden-led)")

build_table(doc,
    headers=["Question", "Honest verdict"],
    rows=[
        ["Q1.  Why this market?", "Pass. Same sector logic plus less consolidation curve."],
        ["Q2.  Why this asset?", "To be tested in Phase 1 scan."],
        ["Q3.  Why now?", "Pass. Sweden is six to eight years behind Norway; window before international consolidators arrive is open."],
        ["Q4.  Why KLAR?", "Better. Ocab gives us operational footprint; Swedish-language and regulatory familiarity built in."],
    ],
    widths_cm=[4.0, 13.0])

body(doc, "Aggregate: 3-3.5 / 4 contingent on Phase 1 anchor work.")

h2(doc, "9.3  Triple A")

build_table(doc,
    headers=["Dimension", "Construction A (Norway)", "Construction B (Sweden)", "Construction C (Nimlas extension)"],
    rows=[
        ["Attractive market", "A", "A", "A-"],
        ["Attractive business", "B+", "A-", "A (via Nimlas)"],
        ["Attractive entry / exit", "B (likely contested entry; thin sponsor exit)", "A- (less contested)", "A- (Nimlas value creation, not entry-multiple-driven)"],
        ["Composite", "B+", "A-", "A-"],
    ],
    widths_cm=[3.5, 4.5, 4.5, 4.5])

h2(doc, "9.4  Key Characteristics Assessment - what changes vs. prior draft")

body(doc,
    "Composite KCA score is unchanged from the prior draft (3.9 / 5) because the "
    "characteristic-level reads are sector-level, not construction-level. The "
    "characteristic that matters most for ranking the three constructions is "
    "'defensible competitive position', which scores 3/5 for Construction A "
    "(no advantage vs. PHM), 4/5 for Construction B (Ocab footprint), and 4/5 "
    "for Construction C (Nimlas platform).")


# ====================================================================================
# 10. WHAT WE'D NEED TO BELIEVE
# ====================================================================================
h1(doc, "10.  What We'd Need to Believe for Construction A to Work")

body(doc,
    "Stated explicitly so that future evidence can either confirm or kill the thesis. "
    "All four would need to be true at IC.")

build_table(doc,
    headers=["#", "Required belief", "Test"],
    rows=[
        ["1", "PHM's M&A pace materially decelerates over our hold (post-Norvestor transition)",
         "Monitor: Norvestor exit prep timing; new owner identity and capital base; PHM Y1-Y2 deal cadence"],
        ["2", "Bolt-on multiples hold at 5-6x for sub-NOK 100m targets despite contested supply",
         "Monitor: completed bolt-on transactions Q3 2026 - Q2 2027 (PHM and other acquirers); banker intel on auction dynamics"],
        ["3", "We land an anchor at <8.5x where management has credible bolt-on integration capability",
         "Phase 2 origination Q3-Q4 2026 on three named candidates"],
        ["4", "Margin expansion of 250-300 bps is achievable despite the structural ceiling on labour services",
         "Operational due diligence with anchor management; reference checks on density and procurement levers"],
    ],
    widths_cm=[1.0, 7.0, 9.0])

callout(doc,
    "On current evidence, beliefs (1) and (2) point against the thesis. Beliefs (3) and "
    "(4) are testable but contingent. We should not advance to IC until all four "
    "are validated, and we should be honest with ourselves that beliefs (1) and (2) "
    "are unlikely to validate given the structural setup.",
    color=NAVY, bg=CALLOUT_BG)


# ====================================================================================
# 11. RECOMMENDED PATH
# ====================================================================================
h1(doc, "11.  Recommended Path")

h2(doc, "11.1  This memo's asks")
bullet(doc,
    "Screening committee endorsement to PASS on Construction A (Norway #2 behind PHM) "
    "as the primary platform thesis.")
bullet(doc,
    "Approval to initiate Phase 1 Sweden scan (Construction B) over Q3-Q4 2026. "
    "Estimated external spend NOK 0.5-1.0m; one full-time origination resource for ~12 weeks.")
bullet(doc,
    "Initiation of strategic review with Nimlas management on Construction C "
    "(technical FM extension). Internal resource only; no external spend in Phase 1.")
bullet(doc,
    "Continued opportunistic origination on the three named Norwegian anchor candidates - "
    "they remain useful for competitive intelligence on PHM, for cross-portfolio bolt-on "
    "introductions to Nimlas, and as potential anchors should Construction B or C surface "
    "a hybrid Sweden + Norway opportunity.")

h2(doc, "11.2  What we should NOT do")
bullet(doc,
    "Commit a fund slot to Construction A pending one or two anchor meetings that go well.")
bullet(doc,
    "Build a Norway-first roll-up and rationalise it post-hoc with Sweden expansion in Y3-Y4. "
    "If Sweden is the answer, Sweden should be the anchor.")
bullet(doc,
    "Pursue all three constructions in parallel. KLAR's origination bandwidth is finite; "
    "Construction B is the priority new-platform thesis, Construction C is a Nimlas "
    "value-creation workstream, Construction A is on the watch list only.")

h2(doc, "11.3  Re-screen trigger")
bullet(doc,
    "Re-screen Construction A if and only if: PHM's M&A cadence falls below 3 deals/year "
    "in Norway for two consecutive years (signal that supply is no longer being absorbed), "
    "OR a named anchor candidate becomes available at <7.5x in a non-competitive process "
    "(rare; would represent a step-change in deal access).")


# ====================================================================================
# APPENDICES
# ====================================================================================
h1(doc, "Appendix A - Sources and Limitations")

bullet(doc,
    "Screening universe: 75 verified Norwegian property service / FM / cleaning entities, "
    "compiled from BoldData FY23 top-50 ranking, PE press releases, Konkurransetilsynet "
    "filings, PHM Annual Report 2024, and Proff.no / Brønnøysundregistrene. Full screen "
    "in companion file Norway_FM_Property_Services_PE_Screen.xlsx.")
bullet(doc,
    "PHM Group revenue (EUR 600-800m) inferred from Norvestor commentary and Norwegian "
    "subsidiary filings. PHM does not publicly disclose Norwegian segment results.")
bullet(doc,
    "Bolt-on multiple ranges (4-6x for sub-NOK 100m, 6-8x for sub-NOK 200m) inferred from "
    "PHM acquisition press releases and Nordic FM transaction comps 2021-2024; require "
    "transaction-by-transaction validation before IC.")
bullet(doc,
    "TAM estimates triangulated against the brief's NOK 15-25bn Norway baseline; not "
    "externally validated. Sweden TAM estimate is order-of-magnitude only and a Phase 1 "
    "scan output, not an input.")
bullet(doc,
    "Returns model is illustrative. Three cases (realistic, base, downside) are "
    "internally consistent at the EBITDA and revenue level but should not be over-read; "
    "real model requires anchor-specific assumptions.")
bullet(doc,
    "KLAR framework applications (4KQ, Triple A, KCA) reflect the deal team's "
    "interpretation pending partner review.")


h1(doc, "Appendix B - Glossary")

bullet(doc, "Vaktmester - janitor / building caretaker.")
bullet(doc, "Eiendomsdrift - property operations.")
bullet(doc, "Fastighetsskötsel - (Swedish) property care / janitorial.")
bullet(doc, "Borettslag - housing cooperative; dominant residential ownership form in Norway.")
bullet(doc, "Sameie - condominium-style co-ownership.")
bullet(doc, "Allmenngjøring(sloven) - law extending collective wage agreement application; creates the labour-cost floor that disadvantages informal operators.")
bullet(doc, "Renholdsoverenskomsten - collective wage agreement for cleaning services.")
bullet(doc, "NACE 81.10 / 81.21 - sector codes for combined facilities support / general building cleaning.")
bullet(doc, "PMI - post-merger integration.")


# Save
out_path = "/home/user/Pilot/output/Nordic_Property_Maintenance_Thesis.docx"
doc.save(out_path)
print(f"Saved: {out_path}")
