"""
Nordic Property Maintenance Roll-Up - Screen memo.

Tight, opinionated, develop-recommendation with a 60-90 day workplan.
Targets 4-5 pages, not 8+. Frameworks inform prose; no scoring tables.
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


NAVY = RGBColor(0x1F, 0x3A, 0x5F)
DARK = RGBColor(0x2A, 0x2A, 0x2A)
MID = RGBColor(0x66, 0x66, 0x66)
LIGHT_BG = "F5F6F8"
CALLOUT_BG = "E8F0FA"
HEADER_BG = "1F3A5F"
HEADER_FG = RGBColor(0xFF, 0xFF, 0xFF)


def shd(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    s = OxmlElement('w:shd')
    s.set(qn('w:val'), 'clear')
    s.set(qn('w:color'), 'auto')
    s.set(qn('w:fill'), hex_color)
    tcPr.append(s)


def borders(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    tcb = OxmlElement('w:tcBorders')
    for e in ('top', 'left', 'bottom', 'right'):
        b = OxmlElement(f'w:{e}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '4')
        b.set(qn('w:color'), 'C8CCD2')
        tcb.append(b)
    tcPr.append(tcb)


def para(doc, text, size=10.5, bold=False, color=None, italic=False,
         before=0, after=4, align=None):
    p = doc.add_paragraph()
    if align:
        p.alignment = align
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.font.size = Pt(size)
    r.bold = bold
    r.italic = italic
    r.font.color.rgb = color or DARK
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    return p


def h1(doc, text):
    para(doc, text, size=14, bold=True, color=NAVY, before=12, after=3)


def h2(doc, text):
    para(doc, text, size=11, bold=True, color=DARK, before=6, after=2)


def body(doc, text):
    para(doc, text, size=10.5, after=4)


def bullet(doc, text):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.left_indent = Cm(0.55)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.font.size = Pt(10.5)
    r.font.color.rgb = DARK


def callout(doc, text):
    t = doc.add_table(rows=1, cols=1)
    c = t.rows[0].cells[0]
    c.text = ""
    p = c.paragraphs[0]
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = NAVY
    shd(c, CALLOUT_BG)
    borders(c)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


def tbl(doc, headers, rows, widths_cm=None):
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.autofit = False
    t.allow_autofit = False
    hdr = t.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = ""
        p = hdr[i].paragraphs[0]
        r = p.add_run(h)
        r.font.name = "Calibri"
        r.font.size = Pt(9.5)
        r.font.bold = True
        r.font.color.rgb = HEADER_FG
        shd(hdr[i], HEADER_BG)
        borders(hdr[i])
        hdr[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    for ri, row in enumerate(rows, start=1):
        cells = t.rows[ri].cells
        for ci, val in enumerate(row):
            cells[ci].text = ""
            p = cells[ci].paragraphs[0]
            r = p.add_run(str(val) if val is not None else "")
            r.font.name = "Calibri"
            r.font.size = Pt(9.5)
            r.font.color.rgb = DARK
            borders(cells[ci])
            cells[ci].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if ri % 2 == 0:
                shd(cells[ci], LIGHT_BG)
    if widths_cm:
        for i, w in enumerate(widths_cm):
            for row in t.rows:
                row.cells[i].width = Cm(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)
    return t


# ====================================================================================
doc = Document()
for s in doc.sections:
    s.top_margin = Cm(1.7); s.bottom_margin = Cm(1.7)
    s.left_margin = Cm(2.0); s.right_margin = Cm(2.0)
doc.styles['Normal'].font.name = 'Calibri'
doc.styles['Normal'].font.size = Pt(10.5)


# ====================================================================================
# COVER
# ====================================================================================
para(doc, "KLAR PARTNERS  |  SCREEN MEMO  |  CONFIDENTIAL", size=9, bold=True,
     color=MID, after=2)
para(doc, "Nordic Property Maintenance Roll-Up", size=21, bold=True,
     color=NAVY, after=1)
para(doc, "The #2 platform behind PHM Group", size=12, italic=True,
     color=MID, after=10)

m = doc.add_table(rows=4, cols=2)
md = [
    ("Sector",  "Property maintenance / janitorial (vaktmester, eiendomsdrift, fastighetsskötsel)"),
    ("Author",  "[Deal team]"),
    ("Date",    "May 2026"),
    ("Verdict", "DEVELOP - 12-week workplan to test a single binary question; convert to IC paper or pass"),
]
for i, (k, v) in enumerate(md):
    c0, c1 = m.rows[i].cells
    c0.text = ""; c1.text = ""
    r0 = c0.paragraphs[0].add_run(k); r0.bold = True
    r0.font.size = Pt(10); r0.font.color.rgb = NAVY
    r1 = c1.paragraphs[0].add_run(v)
    r1.font.size = Pt(10); r1.font.color.rgb = DARK
    c0.width = Cm(2.5); c1.width = Cm(14.5)
doc.add_paragraph()


# ====================================================================================
# 1. RECOMMENDATION
# ====================================================================================
h1(doc, "1.  Recommendation")

callout(doc,
    "Develop. Dedicate one senior originator for 12 weeks to test a single binary question: "
    "can we put Din Vaktmester (Trondheim) or Vaktmesterkompaniet (Oslo) under exclusivity "
    "at 8.5-9.5x FY24E EBITDA before PHM enters or completes its sell-side process? If yes, "
    "convert to IC. If not, pass without further work.")

body(doc,
    "The Norwegian property maintenance market is fragmented enough to absorb a "
    "second consolidator and PHM's seven-year track record removes execution risk on "
    "the playbook. The case turns entirely on anchor access at a disciplined entry "
    "multiple and on running ahead of, not into, PHM's sponsor transition. Both "
    "questions are testable inside a quarter for under NOK 1m of external spend.")

body(doc,
    "We do not recommend a fund-slot reservation. We do not recommend a full sector "
    "diligence yet. We recommend three management conversations and one banker round, "
    "to decide whether this becomes a real deal or stays a watching brief.")


# ====================================================================================
# 2. THE OPPORTUNITY
# ====================================================================================
h1(doc, "2.  The Opportunity")

h2(doc, "2.1  What PHM has demonstrated")

body(doc,
    "Norvestor VIII backed the formation of PHM Group in 2017 and has consolidated "
    "an estimated 100 Nordic property maintenance operators across Finland, Norway, "
    "Sweden and Denmark. Group revenue is approximately EUR 600-800m FY24 "
    "(not publicly disclosed; triangulated from Norwegian sub filings and Norvestor "
    "fund commentary). In Norway, PHM has closed roughly 30 acquisitions including "
    "the absorption of Sefbo in 2022, itself a 33-company roll-up. We estimate PHM "
    "Norway at NOK 1.5-2.5bn revenue, equating to 6-15% share of the NOK 15-25bn "
    "Norwegian addressable market depending on definitional cut.")

body(doc,
    "Two read-throughs from PHM matter for our case. First, the playbook works: "
    "Norvestor took a small Finnish anchor in 2017 and built a platform now likely "
    "valued at EUR 1.4-1.8bn EV (implied 11-12x FY24E EBITDA). Second, PHM cannot "
    "absorb the entire market alone. The Norwegian acquirable supply set in the "
    "NOK 30-400m revenue band is approximately 200-300 entities; PHM has taken "
    "30 of them. The remaining 170+ targets, plus organic-growth demand, are why "
    "a #2 platform construction is arithmetically possible.")

h2(doc, "2.2  Why now")

bullet(doc,
    "Sponsor transition window. Norvestor VIII vintage 2017 is at year nine. Sell-"
    "side process likely 2026-2027; group management will be process-distracted "
    "for 12-18 months. Founders considering exit are looking for alternative "
    "homes; we have a narrow window where PHM is a less-aggressive bidder.")
bullet(doc,
    "Generational handover. Norwegian janitor businesses founded 1985-2000 are now "
    "in founder-aged-60+ territory. Sefbo (2022), Insider's Mitie Norway carve-out "
    "(2024), and Cares' sale to Equip (2023) are the visible signal.")
bullet(doc,
    "Multiples have rebased. Nordic FM platform multiples were 11-13x in 2021-22 "
    "and have compressed to 7-9x as labour-cost services have de-rated. Bolt-on "
    "multiples are 4-6x for sub-NOK 100m targets, 6-8x for NOK 100-300m. The "
    "buy-vs-build spread to an 11-12x platform exit is currently 4-6 turns.")
bullet(doc,
    "Comparable exits print at scale. 4Service sold to Compass in Dec 2024 "
    "(catering-led, ~30 acquisitions over 8 years under Norvestor); Cares to Equip "
    "in 2023; both validate the sponsor-to-sponsor exit market for Nordic "
    "service-led roll-ups in the NOK 2-5bn EV range.")

h2(doc, "2.3  The customer base is genuinely sticky")

body(doc,
    "Norwegian residential property is dominated by borettslag and sameie cooperative "
    "structures; NBBL's umbrella covers 41 housing associations and roughly 525,000 "
    "units. OBOS Eiendomsforvaltning manages a meaningful share. Monthly common-area "
    "fees include a 8-12% property-services line; the contract is typically held by "
    "the housing manager or directly by the board, and switching incurs both "
    "operational disruption (keys, access codes, building knowledge) and political "
    "friction (board members defending their selection). Average customer churn in "
    "the segment is mid-single-digit annually. This is not a customer base that "
    "moves on price; it moves on relationship.")


# ====================================================================================
# 3. THE ROLL-UP ECONOMICS
# ====================================================================================
h1(doc, "3.  Roll-Up Economics")

body(doc,
    "Five-year plan: anchor at NOK 250-1,000m revenue, then 20-26 bolt-ons cumulating "
    "to NOK 1.2-1.6bn of acquired revenue, exiting at NOK 2.0-2.5bn revenue with "
    "11-12% EBITDA margin. The construction is identical in shape to PHM's first "
    "five years; what differs is positioning - commercial property tilt, complementary "
    "geography to PHM's Oslo / Akershus core, and selective Swedish probe via Ocab "
    "in years 3-5.")

tbl(doc,
    ["Year", "Activity", "Revenue (NOKm)", "EBITDA %", "EBITDA (NOKm)"],
    [
        ["Y0", "Anchor close; integration team built", "300", "8.5%", "26"],
        ["Y1", "3-5 bolt-ons; route density build", "550", "9.0%", "50"],
        ["Y2", "5-6 bolt-ons incl. one mid-sized; procurement consolidation", "950", "9.5%", "90"],
        ["Y3", "5-6 bolt-ons; Nimlas technical cross-sell pilot; recap option", "1,400", "10.5%", "147"],
        ["Y4", "4-5 bolt-ons; Swedish probe via Ocab", "1,800", "11.0%", "198"],
        ["Y5", "Exit prep; sell-side preparation", "2,200", "11.5%", "253"],
    ],
    [1.5, 6.5, 2.5, 2.0, 4.0])

body(doc,
    "Anchor enters at 8.5-9.5x. Blended bolt-on multiple ~6.0x. Weighted-average "
    "acquisition multiple across the program ~6.8x. Exit at 11.5x on NOK 253m "
    "EBITDA implies EV NOK 2.9bn. The arithmetic is conventional roll-up: "
    "multiple arbitrage of 4.7 turns provides ~60% of value creation; margin "
    "expansion (300 bps over five years through density and procurement) "
    "provides ~25%; organic growth (3-4% real) the remainder.")

h2(doc, "3.1  Returns")

tbl(doc,
    ["", "Base", "Upside", "Downside"],
    [
        ["Total invested equity (NOKm)",            "880",       "1,250",     "750"],
        ["Y5 platform EBITDA (NOKm)",               "253",       "375",       "140"],
        ["Exit multiple",                           "11.5x",     "13.0x",     "9.5x"],
        ["Exit EV (NOKm)",                          "2,910",     "4,875",     "1,330"],
        ["Gross MoM",                               "2.5 - 2.8x","3.2 - 3.5x","1.5 - 1.8x"],
        ["Gross IRR",                               "22-25%",    "28-32%",    "8-12%"],
    ],
    [7.0, 3.3, 3.3, 3.3])


# ====================================================================================
# 4. WHAT WE'D NEED TO BELIEVE
# ====================================================================================
h1(doc, "4.  What We'd Need to Believe")

body(doc,
    "Three load-bearing assumptions. Each is testable in the 12-week workplan in "
    "Section 7.")

bullet(doc,
    "Anchor access at 8.5-9.5x. Founder-led Norwegian janitor businesses are not "
    "auction-trained; many will transact in a managed bilateral process with the "
    "right relationship, particularly during PHM's sell-side distraction window. "
    "If the anchor only clears at 10-11x, returns compress from 2.7x to 2.2x MoM "
    "and the platform-slot bar is missed.")
bullet(doc,
    "Bolt-on multiples hold at 5-6x blended. PHM's bolt-on pricing discipline at "
    "4-6x for sub-NOK 100m targets has been visible across ~50 disclosed Norwegian "
    "acquisitions. Two consolidators do not necessarily inflate this band; PHM's "
    "and our cost-of-equity discipline anchor it. The risk is a third entrant "
    "(international FM major, strategic) coming in with synergy logic that pushes "
    "multiples to 7-8x.")
bullet(doc,
    "Margin expansion 280-300 bps over five years. PHM's margin trajectory (estimated "
    "6-7% in 2017 to 10-12% today) demonstrates the lever set works. Our anchor "
    "starts at 8-9%; getting to 11-12% requires route density across two-to-three "
    "regional clusters, full procurement consolidation, and back-office integration. "
    "This is mechanical, not speculative.")


# ====================================================================================
# 5. THE ANCHOR QUESTION
# ====================================================================================
h1(doc, "5.  The Anchor Question")

body(doc,
    "Three candidates. We have a clear preference and a clear backup; the third is a "
    "constructive option only if anchor #1 or #2 is structurally unavailable.")

tbl(doc,
    ["Candidate", "Rev FY24E", "Owner", "Read"],
    [
        ["Din Vaktmester AS (Trondheim) [preferred]",
         "NOK 250-350m",
         "Founder",
         "Mid-Norway base PHM has not built; pure janitor culture; smaller cheque (EV NOK 200-300m at 8.5-9.5x). Asymmetric thesis: high probability of bilateral process, lower competing-bidder intensity, allows opportunistic mid-sized bolt-on in Y1-Y2. Risk: smaller starting EBITDA means longer runway to platform scale; mgmt integration capability is the open question"],
        ["Vaktmesterkompaniet AS (Oslo) [backup]",
         "NOK ~1,000m",
         "Founder",
         "Closest PHM analog; scale advantage at entry; ~10 small bolt-ons closed. Risk: PHM is presumed in dialogue given the obvious fit; likely a multi-bidder process at 9-11x. If we land it, we have a Tier-1 platform on day one; if we lose, we have signalled our interest and trained the market"],
        ["Insider Group AS (Høvik) [contingent]",
         "NOK 1.0-1.1bn",
         "Family",
         "Mitie Norway carve-out (2024) is genuine M&A capability. Drawback: post-Mitie mix is >50% cleaning rather than property maintenance. The centre of gravity is wrong for our thesis. Only worth pursuing if both anchor #1 and #2 are unavailable and we restate the thesis as cleaning + janitor platform"],
    ],
    [4.2, 1.8, 1.7, 9.3])

body(doc,
    "Preferred path is Din Vaktmester for two reasons. First, the bilateral process "
    "is more likely; founder-led mid-Norway businesses do not typically run "
    "investment-banker-led auctions. Second, the platform we build off Din Vaktmester "
    "is structurally non-overlapping with PHM's Oslo / Akershus residential core, "
    "which makes the 'why us' narrative defensible at exit. Vaktmesterkompaniet is "
    "the high-quality fallback if Din Vaktmester is not sale-ready.")


# ====================================================================================
# 6. RISKS - REAL, NOT GENERIC
# ====================================================================================
h1(doc, "6.  Risks")

tbl(doc,
    ["Risk", "Real-world manifestation", "Severity"],
    [
        ["PHM defends its Norwegian footprint pre-exit",
         "Norvestor closes 5-10 Norwegian bolt-ons in 2026 to fatten the EBITDA being sold; our pipeline thins, multiples rise on contested targets. Mitigation: focus on Mid-Norway / Vestland targets PHM is less aggressive on; accept lower hit rate on Oslo",
         "Medium-High"],
        ["Anchor pricing breaks above 9.5x",
         "Vaktmesterkompaniet runs a Goldman / Carnegie / DC Advisory process; bidding clears at 10.5-11x. Our model breaks. Mitigation: walk - this is the single most binary test in the 12-week workplan",
         "High"],
        ["PHM's new owner accelerates M&A in 2027-28",
         "Triton / EQT / Cinven sized successor brings more bolt-on capital. Our competitive position deteriorates over our hold. Mitigation: front-load M&A in Y1-Y3 while PHM is in transition",
         "Medium"],
        ["Bolt-on supply migrates into auction process",
         "Founders engage advisors who run multi-party processes; our 4-6x bolt-on discipline cannot hold. Mitigation: source proprietary; build named-target outreach with founders 12-18 months pre-process",
         "Medium"],
        ["Allmenngjøring / wage cost step-change",
         "A new Renholdsoverenskomsten round delivers a step wage increase above price-index pass-through; sector EBITDA margin compresses 100-200 bps for 12-18 months. Symmetric across competitors but a one-off pain. Mitigation: contractual review at anchor; quarterly price re-baselining",
         "Low-Medium"],
        ["Customer concentration in OBOS / NBBL system",
         "If anchor or large bolt-on is over-indexed to OBOS-managed properties, single-customer concentration may exceed 15%. Mitigation: validate concentration in DD and discount entry valuation accordingly",
         "Low"],
    ],
    [4.5, 10.0, 2.5])


# ====================================================================================
# 7. THE 12-WEEK WORKPLAN
# ====================================================================================
h1(doc, "7.  Twelve-Week Workplan")

body(doc,
    "Scope and budget the work as a focused origination sprint, not a sector "
    "diligence engagement. Senior originator + analyst part-time; external spend "
    "capped at NOK 800k. Three decision gates.")

tbl(doc,
    ["Weeks", "Workstream", "Output", "Decision gate"],
    [
        ["1-2",
         "Banker round - DNB Carnegie, Carnegie, ABG, Pareto, Arctic. Confirm whether PHM sell-side mandate is live and target timing; confirm whether any of the three named anchors is under mandate",
         "Banker note; deal-clock map",
         "Go / no-go on full sprint based on PHM process timing and anchor mandate status"],
        ["3-6",
         "Anchor outreach: Din Vaktmester (founder direct or via advisor); Vaktmesterkompaniet (advisor route given likely process); Insider Group (family direct)",
         "Three management meetings; founder posture on sale; preliminary valuation expectation",
         "Bilateral exclusivity option on at least one anchor at <9.5x indicative"],
        ["7-9",
         "Top-15 bolt-on validation: management approach to 5-8 founders, Proff.no / Brønnøysund financial confirmation on remainder. Test 4-6x bolt-on multiple expectation",
         "Confirmed Y1-Y2 pipeline of NOK 250m+ acquired revenue at <6.5x",
         "Pipeline confirms / disconfirms the roll-up arithmetic"],
        ["10-12",
         "Sector commercial diligence (scoped 3-week engagement, NOK 600-800k). TAM validation; PHM positioning; customer-side reference checks (5-8 housing managers, 3-5 commercial property owners)",
         "Commercial DD deck; IC paper draft if anchor secured",
         "Convert to IC paper or pass"],
    ],
    [1.2, 6.3, 4.5, 5.0])

h2(doc, "What converts this to an IC paper")
bullet(doc, "Bilateral exclusivity (or first-look) on one anchor at indicative <9.5x.")
bullet(doc, "Validated 5-year bolt-on pipeline of NOK 1.0bn+ acquired revenue at <7x weighted.")
bullet(doc, "Banker network confirms PHM process timing leaves us 12-18 months of clear runway.")
bullet(doc, "No deal-killer regulatory / labour-cost surprise from sector DD.")

h2(doc, "What stops this here")
bullet(doc, "All three anchors are auction-bound, multi-bidder, and clearing >10x.")
bullet(doc, "PHM accelerates M&A pre-process, taking out 3-5 of our Tier-1 bolt-ons in 2026.")
bullet(doc, "Commercial DD reveals customer churn or contract-tendering dynamics we have underestimated.")


# ====================================================================================
# APPENDIX
# ====================================================================================
h1(doc, "Appendix - Sources and Caveats")

bullet(doc,
    "Screening universe of 75 verified Norwegian property service entities in "
    "companion file Norway_FM_Property_Services_PE_Screen.xlsx; compiled from "
    "BoldData FY23 top-50 ranking, PE press releases, Konkurransetilsynet "
    "filings, PHM Annual Report 2024, Proff.no / Brønnøysundregistrene.")
bullet(doc,
    "PHM Group revenue (EUR 600-800m group, NOK 1.5-2.5bn Norway sub) and "
    "acquisition count (~100 group, ~30 Norway) inferred; PHM does not publicly "
    "disclose Norwegian segment results.")
bullet(doc,
    "Transaction multiples cited (4Service to Compass 2024, Cares to Equip 2023, "
    "Sefbo to PHM 2022) are estimates from press coverage and market commentary; "
    "subject to verification in banker conversations.")
bullet(doc,
    "OBOS / NBBL ecosystem statistics (525,000 units, 41 BBLs, monthly fee "
    "composition) are from NBBL public disclosures; require validation for the "
    "specific share of fee allocated to property services.")
bullet(doc,
    "Returns model is illustrative; per-anchor underwriting requires deal-specific "
    "assumptions.")
bullet(doc,
    "KLAR framework applications inform the prose; the deal team has not used "
    "scoring tables in this memo because at screen stage, the analytical question "
    "is binary (proceed to spend 12 weeks of senior time, or not), not "
    "multi-dimensional scoring.")


out = "/home/user/Pilot/output/Nordic_Property_Maintenance_Thesis.docx"
doc.save(out)
print(f"Saved: {out}")
