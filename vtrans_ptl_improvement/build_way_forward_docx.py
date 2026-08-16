#!/usr/bin/env python3
"""Build V-Trans PTL Way Forward for Improvement Word document."""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

OUT = Path(__file__).resolve().parent / "VTrans_PTL_Way_Forward_for_Improvement.docx"
MD = Path(__file__).resolve().parent / "vtrans_ptl_way_forward.md"

NAVY = RGBColor(0x1B, 0x36, 0x5D)
TEAL = RGBColor(0x0F, 0x6C, 0x8C)
BLACK = RGBColor(0x1A, 0x23, 0x32)


def set_run(run, *, bold=False, size=11, color=BLACK, italic=False):
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.name = "Calibri"
    r = run._element
    r.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")


def add_heading(doc, text, level=1):
    p = doc.add_paragraph()
    if level == 0:
        run = p.add_run(text)
        set_run(run, bold=True, size=22, color=NAVY)
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_after = Pt(4)
    elif level == 1:
        run = p.add_run(text)
        set_run(run, bold=True, size=14, color=NAVY)
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(6)
    else:
        run = p.add_run(text)
        set_run(run, bold=True, size=12, color=TEAL)
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
    return p


def add_para(doc, text, *, bold=False, italic=False, size=11):
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_run(run, bold=bold, italic=italic, size=size)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.space_before = Pt(0)
    return p


def add_bullet(doc, text, *, bold_prefix=None):
    p = doc.add_paragraph(style="List Bullet")
    if bold_prefix:
        r1 = p.add_run(bold_prefix)
        set_run(r1, bold=True, size=11)
        r2 = p.add_run(text)
        set_run(r2, size=11)
    else:
        r = p.add_run(text)
        set_run(r, size=11)
    p.paragraph_format.space_after = Pt(3)
    return p


def shade_header(cell, fill="1B365D"):
    from docx.oxml import OxmlElement

    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def set_cell_text(cell, text, *, bold=False, white=False, size=10):
    cell.text = ""
    p = cell.paragraphs[0]
    run = p.add_run(text)
    set_run(run, bold=bold, size=size, color=RGBColor(0xFF, 0xFF, 0xFF) if white else BLACK)


def add_table(doc, headers, rows):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        set_cell_text(cell, h, bold=True, white=True, size=10)
        shade_header(cell)
    for r_idx, row in enumerate(rows, start=1):
        for c_idx, val in enumerate(row):
            set_cell_text(table.rows[r_idx].cells[c_idx], val, size=10)
    doc.add_paragraph()
    return table


def build():
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    section.left_margin = Inches(0.85)
    section.right_margin = Inches(0.85)

    add_heading(doc, "Way Forward for Improvement", 0)
    add_para(doc, "Unlocking V-Trans PTL Share in a Large Untapped Market", bold=True, size=13)
    add_para(
        doc,
        "Prepared from the diagnosis in “Large Untapped Market Opportunity.”",
        italic=True,
        size=10,
    )

    add_heading(doc, "1. Context from the opportunity diagnosis", 1)
    add_para(
        doc,
        "The organised PTL market is about 101,000 MT (excluding Goa & Konkan), with an estimated "
        "90,000 MT unorganised market (excluding intra-state movement). V-Trans currently moves about "
        "4,300 MT — roughly 4% organised share and ~2% of total intra-state PTL. Competitors such as "
        "ACPL, VRL, TCI, ARC and OM hold an edge through service quality and network depth, not price "
        "alone. Reach to customers is about one-sixth of ACPL and less than half of VRL.",
    )
    add_para(
        doc,
        "Growth is constrained by five reinforcing gaps: weak new-customer conversion; low share of "
        "wallet from existing accounts; sales time lost to operations; weak CRM/follow-up; and service "
        "performance (OTD ~55–60% vs desired 85%) amplified by warehouse and connectivity issues. In "
        "several branches/franchisees (e.g. Shiroli, Satara), FTL is used to bridge targets while PTL "
        "growth stays thin.",
    )

    add_heading(doc, "2. Strategic intent", 1)
    add_para(
        doc,
        "Win PTL share through reliability and reach—not discounting. Raise on-time delivery toward "
        "85%, expand customer reach and conversion, deepen share of wallet with existing accounts, and "
        "protect sales time for market development. Treat FTL as a supplement, not a substitute for "
        "PTL growth.",
        bold=True,
    )
    add_para(
        doc,
        "Outcome to aim for: a measurable lift in PTL tonnage and organised market share, driven by "
        "higher OTD, denser route/branch coverage on priority lanes, and a disciplined sales–CRM engine.",
    )

    add_heading(doc, "3. Diagnosis → way-forward map", 1)
    add_para(
        doc,
        "Every improvement action below is mapped to a specific gap called out in the opportunity note.",
        italic=True,
    )
    add_table(
        doc,
        ["Diagnosis input", "Way-forward response", "Primary owner"],
        [
            [
                "OTD 55–60% vs desired 85%; high TAT, holding, last-mile delay, weak connectivity",
                "Lane-level OTD ladder, cut-off/departure discipline, hub handoff SLAs, warehouse segregation",
                "Operations / Hub lead",
            ],
            [
                "Warehouse gaps: mixed material, loading delay, visibility, transshipment",
                "Branch ownership for segregation & loading readiness; single visibility trail for held shipments",
                "Branch ops + Hub",
            ],
            [
                "Competitor edge = service + network (ACPL, VRL, TCI, ARC, OM); price alone will not win",
                "Sell reliability & network fit; densify priority PTL lanes; publish service-promise pack to sales",
                "Network + Sales",
            ],
            [
                "Reach ~1/6th of ACPL and <½ of VRL; visits happen but conversion is low",
                "Named prospect universe + visit→trial→contract funnel; hand-held first shipments",
                "Sales / Branch manager",
            ],
            [
                "Existing customers give only 10–15% of freight",
                "Share-of-wallet mapping + top-account growth plans + structured BIC cadence",
                "Key account / BIC",
            ],
            [
                "BM & sales distracted by booking/loading/unloading",
                "Ring-fence sales hours; ops backfill at high-volume branches",
                "Regional head",
            ],
            [
                "Franchisees rely on relationship customers; salespersons weak on new logos",
                "Franchise rule: owners hold relationships; sales own new-logo & SOW targets",
                "Franchise + Sales",
            ],
            [
                "Weak CRM: delayed response, poor BIC, limited POD/LR sharing",
                "Mandatory CRM hygiene, response SLA, standard POD/LR & exception alerts",
                "CRM / Customer care",
            ],
            [
                "FTL used to bridge targets (Shiroli, Satara); PTL growth limited",
                "Separate PTL vs FTL scorecards; PTL growth visible even if overall MT is met via FTL",
                "Leadership / Region",
            ],
        ],
    )

    # Pillars
    pillars = [
        (
            "4. Pillar 1 — Fix service performance (non-negotiable growth enabler)",
            "Service failure is the largest barrier to retention and new-customer conversion. Price cannot compensate for low OTD, high TAT, holding, last-mile delay and weak connectivity.",
            [
                "Set a branch-level OTD target ladder toward 85% (e.g. 60% → 70% → 80% → 85%), with weekly review of OTD, TAT, holding and last-mile delay by lane.",
                "Stabilise priority PTL lanes first (highest volume / highest competitive intensity), with defined cut-offs, departure discipline and hub handoff SLAs.",
                "Close warehouse and handling leaks: enforce shipment segregation, stop mixed-material loading where it causes delay, and assign clear ownership for loading readiness.",
                "Improve transshipment coordination between origin, hub and destination with a single visibility trail for held/delayed shipments.",
                "Publish a simple “service promise” pack to sales (lane OTD, expected TAT) so market commitments match operating reality.",
            ],
            "OTD %, average TAT, shipments held >X hours, last-mile delay rate, delay-related complaints/claims.",
        ),
        (
            "5. Pillar 2 — Separate sales effort from branch operations",
            "Branch managers and salespersons currently lose selling time to booking, loading and unloading. That starves market visits and follow-up.",
            [
                "Ring-fence sales hours: define protected daily/weekly time for customer visits and follow-up; operations tasks must not consume that block by default.",
                "Create an ops backfill model at high-volume branches (dedicated booking/loading support or rostered ops cover) so sales is not the default ops buffer.",
                "Franchisee operating rule: franchise owners manage relationship accounts; salespersons own new-logo and share-of-wallet targets with visit and conversion dashboards.",
                "Stop using FTL as the quiet target filler. Track PTL and FTL separately; branch scorecards must show PTL growth even if overall tonnage is met via FTL.",
            ],
            "Productive sales hours / week, visits completed, new accounts opened, PTL vs FTL mix vs target.",
        ),
        (
            "6. Pillar 3 — Rebuild new-customer development and conversion",
            "Visits are happening (often 3–4 customers/day) but contract conversion remains low. Reach is far below ACPL/VRL. The gap is systematic prospecting, qualification and closing—not activity volume alone.",
            [
                "Build a named prospect universe by branch: organised + high-potential unorganised shippers on priority lanes, sized against competitor presence.",
                "Install a conversion funnel: visit → qualified opportunity → trial booking → contract → 90-day retention, with stage owners and weekly funnel review.",
                "Raise reach deliberately: set branch reach targets that close the gap vs ACPL/VRL over a defined horizon (pin codes / industrial clusters / dealer networks).",
                "Standardise the pitch around service reliability and network fit, not rate wars; use lane-level OTD evidence once Pillar 1 improves.",
                "Trial-to-contract playbook: first 3–5 shipments hand-held (POD, proactive status, issue escalation) to convert trials into stickiness.",
            ],
            "Prospects in pipeline, visit-to-trial %, trial-to-contract %, new PTL MT from new logos, reach vs competitor baseline.",
        ),
        (
            "7. Pillar 4 — Expand share of wallet with existing customers",
            "Many accounts give V-Trans only 10–15% of their freight. Winning incremental share from known customers is faster and cheaper than pure new acquisition.",
            [
                "Map share of wallet for top accounts per branch (V-Trans volume vs estimated total freight).",
                "Run account-growth plans for the top 20–30 customers/branch: missing lanes, service gaps, competitor lanes, and a 90-day volume-upside target.",
                "Introduce structured BIC / key-account cadence (weekly/fortnightly) covering performance, pending issues and next-lane opportunities.",
                "Fix relationship hygiene: timely responses, proper communication, consistent LR/POD sharing—stop avoidable churn on already-won accounts.",
                "Offer lane-fill and multi-destination packages where network allows, so customers consolidate more volume with V-Trans.",
            ],
            "Share of wallet %, repeat booking rate, revenue/MT from existing base, top-account retention, complaint closure time.",
        ),
        (
            "8. Pillar 5 — Strengthen network and customer engagement where competitors win",
            "Competitors win through stronger route coverage, branch connectivity and engagement. V-Trans cannot close the market-share gap with price alone.",
            [
                "Prioritise network densification on lanes where demand exists but V-Trans connectivity/TAT is weak relative to ACPL/VRL/TCI/ARC/OM.",
                "Improve branch-to-branch and hub connectivity for PTL density (frequency, cut-off clarity, return load balance).",
                "Local engagement engine: scheduled market meets, dealer/consignee outreach, and branch-level competitor win–loss notes after lost quotes.",
                "Franchise + company branch alignment on common service standards so the customer experience does not fragment by ownership model.",
            ],
            "Lane coverage vs plan, booking acceptance rate, lost-quote reasons, competitor win rate on contested lanes.",
        ),
        (
            "9. Pillar 6 — CRM, visibility and communication discipline",
            "New customers are acquired but follow-up is inconsistent. Delayed responses, weak BIC interaction, poor communication and limited POD/LR sharing drive low repeat business.",
            [
                "Mandatory CRM hygiene: every prospect/customer visit logged with next action and date; no open lead without an owner.",
                "SLA for customer response (e.g. same-day acknowledgement; issue update within agreed hours).",
                "Standardise POD/LR sharing and proactive exception alerts on delayed shipments.",
                "Post-delivery follow-up within 48–72 hours for new and at-risk accounts.",
                "Monthly churn review: accounts with falling volume or rising complaints get a recovery plan.",
            ],
            "Lead follow-up compliance, response SLA adherence, POD/LR turnaround, reactivation of silent accounts.",
        ),
    ]

    for title, intro, actions, kpis in pillars:
        add_heading(doc, title, 1)
        add_para(doc, intro)
        add_para(doc, "Actions", bold=True)
        for a in actions:
            add_bullet(doc, a)
        add_para(doc, "Immediate KPIs: " + kpis, italic=True)

    add_heading(doc, "10. 30 / 60 / 90-day execution plan", 1)
    add_table(
        doc,
        ["Horizon", "Must complete", "Exit criteria"],
        [
            [
                "0–30 days",
                "OTD/TAT baseline by priority lane; sales-hour ring-fence live; PTL vs FTL scorecard split; CRM visit logging mandatory; top-account SOW map started; POD/LR response SLA published",
                "Leadership can see OTD, PTL MT and sales hours weekly without debate",
            ],
            [
                "31–60 days",
                "Priority-lane cut-offs and hub SLAs enforced; warehouse segregation ownership named; top 20–30 accounts/branch have growth plans; BIC cadence running; trial-to-contract hand-hold live on new logos",
                "OTD moving up the ladder; SOW plans in motion; visit→trial conversion tracked",
            ],
            [
                "61–90 days",
                "Named prospect universe live at every branch; funnel review institutionalised; selective network densification on contested lanes; franchise sales/ops rule audited; reach gap vs ACPL/VRL measured",
                "PTL growth visible independent of FTL; conversion and reach improving on priority clusters",
            ],
        ],
    )
    add_para(
        doc,
        "Do not scale aggressive new acquisition on lanes where OTD remains far below promise; that creates churn and brand damage.",
        bold=True,
    )

    add_heading(doc, "11. Sequencing (order of attack)", 1)
    add_table(
        doc,
        ["Wave", "Focus", "Why"],
        [
            [
                "Near term",
                "Pillars 1, 2, 4, 6 — OTD/TAT, sales time protection, share-of-wallet, CRM hygiene",
                "Stops leakage from the current base and makes market promises believable",
            ],
            [
                "Next wave",
                "Pillar 3 — systematic new-customer funnel and reach expansion",
                "Conversion improves once service and follow-up are credible",
            ],
            [
                "Parallel / ongoing",
                "Pillar 5 — selective network densification on priority PTL lanes",
                "Builds the structural answer to competitor advantage",
            ],
        ],
    )

    add_heading(doc, "12. Operating rhythm", 1)
    add_bullet(doc, "OTD/holding exceptions; sales visit plan vs actual; open customer issues.", bold_prefix="Daily (branch): ")
    add_bullet(
        doc,
        "PTL vs FTL vs target; funnel conversion; top-account share-of-wallet moves; competitor losses.",
        bold_prefix="Weekly (cluster/region): ",
    )
    add_bullet(
        doc,
        "Organised market-share trajectory, reach gap vs ACPL/VRL, OTD vs 85% path, franchise compliance to sales/ops separation.",
        bold_prefix="Monthly (leadership): ",
    )

    add_heading(doc, "13. Leadership scorecard", 1)
    add_table(
        doc,
        ["Theme", "Metric", "Direction"],
        [
            ["Market", "PTL MT, organised MS %, reach vs ACPL/VRL", "Up"],
            ["Service", "OTD (toward 85%), TAT, holding, last-mile delay", "Improve"],
            ["Sales engine", "Protected sales hours, visits, visit→contract %", "Up"],
            ["Existing base", "Share of wallet (from 10–15% upward), retention", "Up"],
            ["Mix quality", "PTL growth independent of FTL target bridging", "Up"],
            ["CRM", "Follow-up compliance, POD/LR SLA, complaint closure", "Up"],
        ],
    )

    add_heading(doc, "14. Expected impact if executed with discipline", 1)
    for t in [
        "Service credibility rises as OTD moves toward 85%, unlocking conversion and retention.",
        "Existing-customer MT grows by reclaiming share of wallet beyond the current 10–15% band.",
        "New-logo conversion improves because visits sit inside a managed funnel, not activity theatre.",
        "Sales capacity increases as ops distraction is ring-fenced.",
        "Target quality improves: PTL growth is visible and not masked by FTL.",
        "Over time, organised share can move meaningfully above today’s ~4% as reach, network fit and reliability compound.",
    ]:
        add_bullet(doc, t)

    add_heading(doc, "15. One-line way forward", 1)
    add_para(
        doc,
        "Make PTL reliable, free salespeople to sell, convert more of what you already touch, deepen every existing account, and densify the network where competitors currently win—while scoring PTL on its own, not through FTL.",
        bold=True,
    )

    doc.save(OUT)
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    build()
