"""
dpr_generator.py
Generates the full DPR .docx using python-docx.
Accepts a params dict, returns bytes.
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy
import io
from financial_model import run_model

# ── Colour palette ──────────────────────────────────────────────────────────
C_NAVY   = RGBColor(0x1F, 0x38, 0x64)
C_BLUE   = RGBColor(0x2E, 0x74, 0xB5)
C_ORANGE = RGBColor(0xC5, 0x5A, 0x11)
C_WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
C_LGRAY  = RGBColor(0xEB, 0xF3, 0xFB)
C_GREEN  = RGBColor(0x1E, 0x7B, 0x34)
C_GOLD   = RGBColor(0xD4, 0xAF, 0x37)

# ── Helpers ──────────────────────────────────────────────────────────────────
def f2(n):
    return f"{n:,.2f}"

def f0(n):
    return f"{n:,.0f}"

def set_cell_bg(cell, hex_color: str):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    tcPr.append(shd)

def set_cell_border(cell):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        border = OxmlElement(f"w:{side}")
        border.set(qn("w:val"),   "single")
        border.set(qn("w:sz"),    "4")
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), "AAAAAA")
        tcBorders.append(border)
    tcPr.append(tcBorders)

def set_col_width(cell, width_cm):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcW  = OxmlElement("w:tcW")
    tcW.set(qn("w:w"),    str(int(width_cm * 567)))
    tcW.set(qn("w:type"), "dxa")
    tcPr.append(tcW)

def para_run(doc_or_cell, text, bold=False, size=11, color=None,
             italic=False, underline=False, align=None, space_after=6, space_before=2):
    if hasattr(doc_or_cell, "add_paragraph"):
        p = doc_or_cell.add_paragraph()
    else:
        p = doc_or_cell.paragraphs[0] if doc_or_cell.paragraphs else doc_or_cell.add_paragraph()
    p.paragraph_format.space_after  = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    if align:
        p.alignment = align
    run = p.add_run(text)
    run.bold      = bold
    run.italic    = italic
    run.underline = underline
    run.font.size = Pt(size)
    run.font.name = "Arial"
    if color:
        run.font.color.rgb = color
    return p

def add_heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    h.paragraph_format.space_before = Pt(14)
    h.paragraph_format.space_after  = Pt(6)
    for run in h.runs:
        run.font.name = "Arial"
        run.font.color.rgb = C_NAVY if level == 1 else C_BLUE
        run.font.size = Pt(14 if level == 1 else 12)
    return h

def add_bullet(doc, text, size=10.5):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(3)
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.name = "Arial"
    return p

def add_table(doc, headers, rows, col_widths_cm=None,
              header_bg="1F3864", alt_bg="EBF3FB"):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Header row
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        cell = hdr_cells[i]
        set_cell_bg(cell, header_bg)
        set_cell_border(cell)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after  = Pt(3)
        run = p.add_run(h)
        run.bold = True
        run.font.size  = Pt(9)
        run.font.name  = "Arial"
        run.font.color.rgb = C_WHITE
        if col_widths_cm:
            set_col_width(cell, col_widths_cm[i])

    # Data rows
    for ri, row in enumerate(rows):
        cells = table.rows[ri + 1].cells
        bg = alt_bg if ri % 2 == 1 else "FFFFFF"
        for ci, val in enumerate(row):
            cell = cells[ci]
            set_cell_bg(cell, bg)
            set_cell_border(cell)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after  = Pt(2)
            run = p.add_run(str(val))
            run.font.size = Pt(9)
            run.font.name = "Arial"
            if col_widths_cm:
                set_col_width(cell, col_widths_cm[ci])

    doc.add_paragraph()
    return table

def page_break(doc):
    doc.add_page_break()


# ── Main generator ────────────────────────────────────────────────────────────
def generate_dpr(params: dict) -> bytes:
    m = run_model(params)
    cfs     = m["cashflows"]
    proj_irr = m["project_irr"]
    eq_irr   = m["equity_irr"]
    min_dscr = m["min_dscr"]
    avg_dscr = m["avg_dscr"]

    p = params  # shorthand

    doc = Document()

    # ── Page margins ──────────────────────────────────────────────────────────
    for section in doc.sections:
        section.top_margin    = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2.5)

    # ── Styles ────────────────────────────────────────────────────────────────
    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(10.5)

    # ══════════════════════════════════════════════════════════════════════════
    # COVER PAGE
    # ══════════════════════════════════════════════════════════════════════════
    doc.add_paragraph()
    doc.add_paragraph()
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run("DETAILED PROJECT REPORT (DPR)")
    title_run.bold = True
    title_run.font.size = Pt(22)
    title_run.font.color.rgb = C_NAVY
    title_run.font.name = "Arial"
    title_run.underline = True
    title_p.paragraph_format.space_after = Pt(20)

    for line, sz, bold, col in [
        (f"DEVELOPMENT OF {f2(p['project_capacity_dc'])} MWp DC / {f2(p['project_capacity_ac'])} MW AC", 16, True, C_NAVY),
        ("GROUND MOUNTED SOLAR PV POWER PROJECT", 14, True, C_NAVY),
        (f"AT {p['location_village'].upper()}, {p['location_taluk'].upper()}", 13, True, C_ORANGE),
        (f"{p['location_district'].upper()} DISTRICT, {p['location_state'].upper()}", 13, True, C_ORANGE),
        ("", 8, False, C_NAVY),
        ("Prepared for:", 11, False, C_BLUE),
        (f"M/S. {p['company_name'].upper()} ({p['company_short'].upper()})", 15, True, C_NAVY),
        (p["company_address"], 10, False, RGBColor(0x44,0x44,0x44)),
        (f"A Subsidiary of {p['parent1_name']} ({p['parent1_pct']:.0f}%) & {p['parent2_name']} ({p['parent2_pct']:.0f}%)", 10, False, RGBColor(0x55,0x55,0x55)),
        ("", 6, False, C_NAVY),
        (f"COD: {p['cod_month']} {p['cod_year']}", 12, True, C_ORANGE),
    ]:
        pp = doc.add_paragraph()
        pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pp.paragraph_format.space_after = Pt(4)
        run = pp.add_run(line)
        run.bold = bold
        run.font.size = Pt(sz)
        run.font.color.rgb = col
        run.font.name = "Arial"

    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # INDEX
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "INDEX")
    index_items = [
        "1.  Summary of the Company & Group Profile",
        "2.  Proposal for Solar Power Project",
        "3.  Concept of Power Sale (PPA)",
        "4.  Solar Resource Potential & Radiation Terminology",
        "5.  Project Site Details & Solar Resource Assessment",
        "6.  Project at a Glance",
        "7.  Demand Analysis and Justification of the Project",
        "8.  Benefits of Grid-Connected Solar PV Power Plant",
        "9.  Basic System Description",
        "10. Bill of Quantity (BOQ)",
        "11. Planned Project Schedule",
        "11A. Regulatory Approvals & Clearances",
        "11B. Operation & Maintenance Plan and Risk Analysis",
        "12. Project Costing",
        "13. Project Financials – 25 Year Cash Flow",
        "14. Debt Repayment Schedule & DSCR Analysis",
        "15. Summary of Results",
        "16. Conclusion",
        "Annexure A – Solar Resource Assessment & PVSyst Simulation",
    ]
    for item in index_items:
        add_bullet(doc, item)
    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 1 – COMPANY SUMMARY
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "1. SUMMARY OF THE COMPANY")
    para_run(doc, f"{p['company_name']} ({p['company_short']}) is a renewable energy focused Independent Power Producer (IPP), "
             f"incorporated as a joint venture subsidiary with equal participation from {p['parent1_name']} ({p['parent1_pct']:.0f}%) "
             f"and {p['parent2_name']} ({p['parent2_pct']:.0f}%).", space_after=8)
    para_run(doc, f"Registered at: {p['company_address']}", space_after=8)
    para_run(doc, f"{p['company_short']} is established with a clear vision to develop, own, and operate utility-scale solar power assets "
             f"in India. The company intends to develop a {f2(p['project_capacity_ac'])} MW AC ({f2(p['project_capacity_dc'])} MWp DC) "
             f"ground-mounted Solar Power Project at {p['location_village']}, {p['location_district']} District, {p['location_state']}. "
             f"The power generated will be sold under a {p['ppa_term']}-year PPA at a flat tariff of "
             f"₹ {p['ppa_tariff']:.2f} per unit. COD is planned for {p['cod_month']} {p['cod_year']}.", space_after=8)
    para_run(doc, f"{p['company_short']} combines the engineering and financial strength of its parent companies, "
             f"positioning it as a well-funded and technically capable developer for this project.", space_after=10)

    add_heading(doc, "Group Profile", level=2)
    add_heading(doc, f"I. {p['parent1_name']}", level=3)
    para_run(doc, f"{p['parent1_name']} is one of the promoter entities holding {p['parent1_pct']:.0f}% equity in {p['company_short']}. "
             "The company brings significant financial strength, industry relationships, and execution capabilities to the project. "
             "Its participation ensures robust corporate governance and long-term commitment to the success of the project.", space_after=8)

    add_heading(doc, f"II. {p['parent2_name']}", level=3)
    para_run(doc, f"{p['parent2_name']} holds {p['parent2_pct']:.0f}% equity in {p['company_short']}. "
             "The company contributes renewable energy project development expertise, EPC execution capability, and operational "
             "experience in solar and wind power assets, ensuring the project is developed and managed to the highest standards.", space_after=8)

    add_heading(doc, "Group Vision & Mission", level=2)
    para_run(doc, "Vision: To build a globally respected clean energy platform delivering reliable, sustainable, and "
             "technology-driven solutions that support the world's transition to a low-carbon future.", space_after=6)
    para_run(doc, "Mission: To create long-term value by combining engineering excellence, innovation, and disciplined "
             "execution in renewable energy businesses, delivering efficient and sustainable solutions worldwide.", space_after=10)
    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 2 – PROPOSAL
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "2. PROPOSAL FOR SOLAR POWER PROJECT")
    para_run(doc, f"M/s. {p['company_name']} ({p['company_short']}) is proposing to set up a {f2(p['project_capacity_ac'])} MW AC "
             f"({f2(p['project_capacity_dc'])} MWp DC) Ground Mounted Solar PV Power Project at {p['location_village']} Village, "
             f"{p['location_district']} District, {p['location_state']}. The project leverages the excellent solar resource "
             f"available at the site and a favourable long-term PPA to deliver strong financial returns.", space_after=8)
    gen_yr1_total = p['gen_yr1_lac_per_mwac'] * p['project_capacity_ac']
    para_run(doc, f"The annual power generation from the plant will be approximately {f2(gen_yr1_total)} "
             f"Lac Units ({gen_yr1_total / 10:.1f} Million kWh) in Year 1. "
             f"The power will be sold under a {p['ppa_term']}-year PPA at a flat tariff of ₹ {p['ppa_tariff']:.2f} per unit. "
             f"COD is planned for {p['cod_month']} {p['cod_year']}.", space_after=8)

    add_heading(doc, "Project Rationale", level=2)
    for item in [
        f"Rising electricity demand in {p['location_state']}'s industrial and commercial sectors with 6–8% annual growth",
        "Increasing cost of conventional grid power making fixed-tariff solar PPA a compelling alternative",
        f"Excellent solar resource at {p['location_village']} site with GHI of 5.5–5.8 kWh/m²/day",
        "Strong policy support from Government of India and state government for solar development",
        f"Long-term flat PPA at ₹ {p['ppa_tariff']:.2f} per unit providing revenue certainty and bankability",
        f"Strong promoter group credentials ensuring execution capability and financial strength",
    ]:
        add_bullet(doc, item)

    add_heading(doc, "Expected Impact", level=2)
    gen_total = p['gen_yr1_lac_per_mwac'] * p['project_capacity_ac']
    for item in [
        f"Generation: ~{f2(gen_total)} Lac Units/year in Year 1; ~{f0(gen_total * 23)} Lac Units over 25 years",
        f"Employment: ~200 jobs during construction; ~15 permanent O&M jobs",
        f"CO₂ Avoidance: ~{gen_total * 0.82 * 100:.0f} tonnes/year (~{gen_total * 0.82 * 100 * 25 / 100000:.1f} lakh tonnes over 25 years)",
        f"Total Revenue over 25 years: ~₹ {gen_total * p['ppa_tariff'] * 22 / 100:.0f} Crores (approximate)",
    ]:
        add_bullet(doc, item)
    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 3 – CONCEPT OF POWER SALE
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "3. CONCEPT OF POWER SALE (PPA)")
    para_run(doc, f"The energy generated from the project shall be fully sold to the procurer under a "
             f"{p['ppa_term']}-year flat-rate Power Purchase Agreement (PPA) at ₹ {p['ppa_tariff']:.2f} per unit. "
             "Under the PPA model, the Solar Generation Plant supplies power into the nearest power grid, "
             "which is then credited against the energy consumption of the identified consumer locations.", space_after=8)
    para_run(doc, "[Basic Concept Diagram – Insert here]", italic=True,
             color=RGBColor(0x88,0x88,0x88), space_after=8)
    para_run(doc, "The Solar Generation Plant supplies energy into the Nearest Power Grid (Meter Reading 4). "
             "This energy is deducted from the EB bills of the consumers, adjusting the energy cost against "
             f"the solar generator. Consumers receive solar power at ₹ {p['ppa_tariff']:.2f} per unit – "
             "significantly below prevailing HT industrial tariff.", space_after=8)

    add_heading(doc, "Advantages of Solar Power PPA", level=2)
    for item in [
        f"Fixed tariff of ₹ {p['ppa_tariff']:.2f} per unit for {p['ppa_term']} years – full insulation against grid tariff escalation",
        "No capital investment required from the consumer – all project costs borne by the generator",
        "Significant savings versus prevailing HT grid tariff (typically ₹ 6.50–₹ 8.50 per unit)",
        "Supports consumer ESG, green certification, and sustainability reporting objectives",
        "Reduces carbon footprint and supports India's NDC commitments under the Paris Agreement",
        "Long-term energy security and price certainty for the power procurer",
    ]:
        add_bullet(doc, item)

    add_heading(doc, "Policy Framework", level=2)
    para_run(doc, f"The project is developed under the enabling policy framework of the Government of {p['location_state']} "
             "and the Government of India, including the National Solar Mission, Tamil Nadu Solar Energy Policy 2023 "
             "(targeting 9,000 MW by 2025), Open Access provisions under the Electricity Act 2003, and exemption "
             "from Electricity Tax for eligible solar projects.", space_after=10)
    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 4 – SOLAR RESOURCE
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "4. SOLAR RESOURCE POTENTIAL")
    add_heading(doc, "India's Solar Energy Resource", level=2)
    para_run(doc, "India is located in the sunny belt of the earth, receiving abundant radiant energy. "
             "India has a theoretical solar potential of over 748 GW. Installed solar capacity crossed 70 GW "
             "by March 2023, with a target of 300 GW solar by 2030 under the National Solar Mission.", space_after=8)

    add_heading(doc, f"{p['location_state']}'s Solar Resource", level=2)
    para_run(doc, f"{p['location_state']} is among the states with the highest solar irradiation in India, "
             f"receiving GHI values of 5.0–6.2 kWh/m²/day. The {p['location_district']} region receives "
             "average GHI of approximately 5.5–5.8 kWh/m²/day, making it highly suitable for utility-scale solar.", space_after=8)

    add_heading(doc, "Key Solar Resource Terms", level=2)
    terms = [
        ("GHI (Global Horizontal Irradiance)", "Total solar irradiance on a horizontal surface. Primary input for solar yield simulation."),
        ("DNI (Direct Normal Irradiance)", "Solar irradiance on a surface perpendicular to sun rays. Used for CSP and bifacial analysis."),
        ("Peak Sun Hours (PSH)", f"Equivalent hours/day at 1,000 W/m². {p['location_district']} averages 5.5–5.8 PSH."),
        ("Performance Ratio (PR)", "Ratio of actual to theoretical yield. Target PR for this project: 78–80%."),
        ("Capacity Utilisation Factor (CUF)", f"Ratio of actual to maximum possible generation. Estimated CUF: ~24.5%."),
        ("Temperature Coefficient", "Rate of efficiency loss per °C above 25°C (STC). Mono-Si: typically –0.35%/°C."),
    ]
    for term, defn in terms:
        pp = doc.add_paragraph()
        pp.paragraph_format.space_after = Pt(4)
        r1 = pp.add_run(f"{term}: ")
        r1.bold = True; r1.font.size = Pt(10.5); r1.font.name = "Arial"
        r1.font.color.rgb = C_NAVY
        r2 = pp.add_run(defn)
        r2.font.size = Pt(10.5); r2.font.name = "Arial"

    para_run(doc, "Detailed site-specific solar resource assessment and PVSyst simulation are provided in Annexure – A.",
             italic=True, color=RGBColor(0x55,0x55,0x55), space_after=10)
    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 5 – SITE DETAILS
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "5. PROJECT SITE DETAILS")
    para_run(doc, f"The proposed site is located at {p['location_village']} Village, {p['location_taluk']} Taluk, "
             f"{p['location_district']} District, {p['location_state']}. The site is approximately "
             f"{p['nearest_town_km']:.0f} km from {p['nearest_town']} and is well connected by road.", space_after=8)
    para_run(doc, "[Site Location Maps – Insert district map, village map, and satellite imagery here]",
             italic=True, color=RGBColor(0x88,0x88,0x88), space_after=8)

    add_heading(doc, "Site Details", level=2)
    site_rows = [
        ["Area Required", f"Approx. {p['land_acres']:.0f} Acres (Private Land)"],
        ["Latitude", p["latitude"]],
        ["Longitude", p["longitude"]],
        ["Summer Temperature", "32°C to 40°C"],
        ["Winter Temperature", "17°C to 30°C"],
        ["Average Annual GHI", "~5.5 to 5.8 kWh/m²/day"],
        ["Annual Sunshine Hours", "~2,800 hours/year"],
        ["Terrain", "Plain land, shadow-free"],
        ["Distance to TNEB SS", f"~{p['tneb_distance_km']:.1f} km from project site"],
        ["Land Type", "Private land – plain soil"],
        ["Nearest Town", f"{p['nearest_town']} (~{p['nearest_town_km']:.0f} km)"],
    ]
    add_table(doc, ["Parameter", "Details"], site_rows, col_widths_cm=[5, 10])

    add_heading(doc, "Solar Resource Assessment", level=2)
    para_run(doc, "A solar resource assessment has been carried out using satellite-derived irradiance data from "
             "Meteonorm and Solargis databases. The detailed PVSyst simulation report is provided in Annexure – A.",
             space_after=10)
    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 6 – PROJECT AT A GLANCE
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "6. PROJECT AT A GLANCE")
    glance_rows = [
        ["1.1", "Project Developer", f"{p['company_name']} ({p['company_short']})"],
        ["1.2", "The Project", f"{f2(p['project_capacity_ac'])} MW AC, {f2(p['project_capacity_dc'])} MWp DC Solar PV"],
        ["1.3", "Location", f"{p['location_village']}, {p['location_taluk']}, {p['location_district']}, {p['location_state']}"],
        ["1.4", "COD", f"{p['cod_month']} {p['cod_year']}"],
        ["1.5", "Land Required", f"Approx. {p['land_acres']:.0f} Acres"],
        ["2.1", "DC Capacity", f"{f2(p['project_capacity_dc'])} MWp (Monocrystalline)"],
        ["2.2", "AC Capacity", f"{f2(p['project_capacity_ac'])} MW AC"],
        ["2.3", "AC:DC Ratio", f"1 : {p['ac_dc_ratio']:.1f}"],
        ["2.4", "Year 1 Generation", f"{f2(p['gen_yr1_lac_per_mwac'] * p['project_capacity_ac'])} Lac Units"],
        ["2.5", "Panel Degradation", f"{p['degradation_yr1_pct']:.1f}% (Yr 1); {p['degradation_yr2_pct']:.1f}%/yr thereafter"],
        ["2.6", "PPA Tariff", f"₹ {p['ppa_tariff']:.2f} per unit (Flat, {p['ppa_term']} years)"],
        ["3.1", "Total Project Cost", f"₹ {p['project_cost_cr']:.2f} Crores (₹ {f2(p['project_cost_lac'])} Lakhs)"],
        ["3.2", "Debt", f"₹ {p['debt_cr']:.2f} Crores @ {p['debt_interest_rate']:.1f}% p.a. ({p['debt_tenor_yrs']} years)"],
        ["3.3", "Equity", f"₹ {p['equity_cr']:.2f} Crores"],
        ["3.4", "Project IRR (Post-Tax)", f"{proj_irr:.2f}%"],
        ["3.5", "Equity IRR (Post-Tax)", f"{eq_irr:.2f}%"],
        ["3.6", "Minimum DSCR", f"{min_dscr:.2f}x"],
        ["3.7", "Average DSCR", f"{avg_dscr:.2f}x"],
        ["4.1", "O&M – Year 1", "Free of Cost"],
        ["4.2", "O&M – Year 2 onwards", f"₹ {p['om_rate_lac_per_mwac']:.1f} Lac/MWac; {p['om_escalation_pct']:.1f}% escalation"],
        ["4.3", "Insurance", f"{p['insurance_pct']:.2f}% of project cost; {p['insurance_esc_pct']:.1f}% escalation"],
    ]
    add_table(doc, ["Sl.", "Parameter", "Details"], glance_rows, col_widths_cm=[1.2, 5.5, 9])
    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 7 – DEMAND ANALYSIS
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "7. DEMAND ANALYSIS AND JUSTIFICATION")
    para_run(doc, f"Electricity is the most essential input for growth and development. {p['location_state']} is growing rapidly "
             "in both industrial and agricultural sectors, with power demand increasing at 6–8% annually. "
             "The state faces power deficits during peak seasons, making decentralised solar generation critical.", space_after=8)
    add_heading(doc, "National Context", level=2)
    para_run(doc, "India's solar installed capacity crossed 70 GW by March 2023. The Government targets 300 GW of solar "
             "by 2030. Solar PV module prices have fallen 90%+ in the past decade making solar the lowest-cost "
             "source of new electricity in most markets.", space_after=8)
    add_heading(doc, "Project Justification", level=2)
    for item in [
        f"Economic: Project IRR of {proj_irr:.2f}% (post-tax); flat tariff provides savings over escalating grid tariff",
        f"Environmental: ~{p['gen_yr1_lac_per_mwac'] * p['project_capacity_ac'] * 0.82 * 100:.0f} tonnes CO₂ avoided annually",
        f"Technical: Excellent solar resource at {p['location_village']} – GHI 5.5–5.8 kWh/m²/day; flat terrain",
        "Policy: Aligns with National Solar Mission, Tamil Nadu Solar Policy 2023, and India's Paris Agreement NDCs",
    ]:
        add_bullet(doc, item)
    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 8 – BENEFITS
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "8. BENEFITS OF GRID-CONNECTED SOLAR PV POWER PLANT")
    for grp, items in [
        ("Economic Benefits", [
            "Zero fuel cost and no exposure to fossil fuel price volatility",
            f"Predictable revenue via {p['ppa_term']}-year PPA at ₹ {p['ppa_tariff']:.2f}/unit",
            f"Project IRR {proj_irr:.2f}%; Equity IRR {eq_irr:.2f}% – strong investor returns",
            "Significant savings for consumer vs HT grid tariff",
        ]),
        ("Environmental Benefits", [
            f"~{p['gen_yr1_lac_per_mwac'] * p['project_capacity_ac'] * 82:.0f} tonnes CO₂ avoided per year",
            "Zero water consumption during generation; no air, water, or noise pollution",
            "Supports India's NDC commitments under Paris Agreement",
        ]),
        ("Technical & Social Benefits", [
            "High reliability – no moving parts; MTBF exceeds 20 years",
            "Short construction timeline: 4–6 months from financial close",
            "~200 construction jobs; ~15 permanent O&M jobs",
            "Local economic development through land lease and support services",
        ]),
    ]:
        add_heading(doc, grp, level=2)
        for item in items:
            add_bullet(doc, item)
    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 9 – SYSTEM DESCRIPTION
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "9. BASIC SYSTEM DESCRIPTION")
    para_run(doc, f"The proposed {f2(p['project_capacity_dc'])} MWp DC / {f2(p['project_capacity_ac'])} MW AC solar power plant "
             "will use monocrystalline silicon PV technology with an AC:DC ratio of "
             f"{p['ac_dc_ratio']:.1f}. The plant will be designed per IEC 61215, IEC 62109, and applicable CEA regulations.", space_after=8)

    components = [
        ("9.1 Solar PV Modules",
         f"High-efficiency monocrystalline modules (530–580 Wp). Total ~{p['project_capacity_dc']*1000/0.56:.0f} modules. "
         "Fixed tilt 11–15°, south-facing. 25-year linear performance warranty from Tier-1 BNEF-rated suppliers."),
        ("9.2 Inverters",
         f"String/central inverters (ILR = {p['ac_dc_ratio']:.1f}). MPPT efficiency ≥99.5%. "
         "IEEE 1547/IEC 62109 certified. Reactive power control, LVRT, frequency-watt capability."),
        ("9.3 Module Mounting Structures",
         "Hot-dip galvanised GI structures. 25-year design life. Designed for 170 km/h wind (IS 875 Part 3)."),
        ("9.4 Transformers",
         "0.69 kV/33 kV step-up transformers (5–6.3 MVA, ONAN). 15 kVA auxiliary transformers for station supply."),
        ("9.5 HT Switchyard (33 kV)",
         "VCB panels, metering cubicles (0.2 class), protection relays, DC battery backup. "
         f"33 kV evacuation line ~{p['tneb_distance_km']:.1f} km to TNEB substation."),
        ("9.6 SCADA & Monitoring",
         "Real-time monitoring of generation, inverters, weather, and energy meters. "
         "Modbus TCP/IP and IEC 61850. Daily/weekly/monthly performance reports."),
        ("9.7 Earthing & Lightning Protection",
         "GI flat earthing grid across plant area. Lightning arrestors at arrays, inverters, and switchyard."),
        ("9.8 Cables",
         "1500 VDC UV-resistant TÜV/UL solar DC cables. XLPE aluminium armoured 33 kV HT cables."),
    ]
    for title, desc in components:
        add_heading(doc, title, level=3)
        para_run(doc, desc, space_after=6)
    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 10 – BOQ
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "10. BILL OF QUANTITY (BOQ)")
    boq_rows = [
        ["A", "SOLAR PV MODULES", "", "", ""],
        ["1", f"Monocrystalline Solar PV Modules (530–580 Wp)", "Nos", f"~{p['project_capacity_dc']*1750:.0f}", f"{f2(p['project_capacity_dc'])} MWp"],
        ["B", "INVERTERS & POWER CONVERSION", "", "", ""],
        ["2", "String/Central Inverters", "Nos", "LS", f"{f2(p['project_capacity_ac'])} MW AC"],
        ["3", "ACDB / DCDB Panels", "Set", "LS", ""],
        ["C", "MODULE MOUNTING STRUCTURES", "", "", ""],
        ["4", "GI Hot-Dip Galvanized MMS", "MT", "LS", "Fixed tilt"],
        ["5", "Foundation bolts & accessories", "Lot", "1", ""],
        ["D", "HT INFRASTRUCTURE", "", "", ""],
        ["6", "33/0.69 kV Step-Up Transformer", "Nos", "4", "5–6.3 MVA each"],
        ["7", "33 kV HT Panel with 1600A ACB", "Nos", "1", ""],
        ["8", "VCB Panel (incoming + outgoing)", "Set", "1", "33 kV"],
        ["9", "33 kV Metering Cubicle (EB spec)", "Nos", "1", ""],
        ["10", "Station Battery & Float Charger", "Set", "1", ""],
        ["E", "CABLES", "", "", ""],
        ["11", "DC Cables (module to inverter)", "Lot", "1", ""],
        ["12", "AC Cables (inverter to transformer)", "Lot", "1", ""],
        ["13", "33 kV XLPE Al Armoured Cable", "Lot", "1", ""],
        ["14", "Communication Cables (SCADA)", "Lot", "1", ""],
        ["F", "CIVIL WORKS", "", "", ""],
        ["15", "Site preparation & levelling", "Lot", "1", f"~{p['land_acres']:.0f} acres"],
        ["16", "Control & inverter room", "Nos", "1", ""],
        ["17", "Cable trenching", "Lot", "1", ""],
        ["18", "Internal roads, fencing & wall", "Lot", "1", ""],
        ["19", "Water storage for panel cleaning", "Lot", "1", ""],
        ["G", "EARTHING, SCADA & MISC", "", "", ""],
        ["20", "Earthing system (GI flat, earth pits)", "Lot", "1", ""],
        ["21", "Lightning arrestors", "Lot", "1", ""],
        ["22", "SCADA & monitoring system", "Set", "1", ""],
        ["23", f"33 kV evacuation line (~{p['tneb_distance_km']:.1f} km)", "Km", f"{p['tneb_distance_km']:.1f}", "To TNEB SS"],
    ]
    add_table(doc, ["Sl.", "Description", "Unit", "Qty", "Remarks"], boq_rows,
              col_widths_cm=[0.8, 7.5, 1.5, 1.8, 3.5])
    para_run(doc, "* Tentative BOM. Final BOM subject to detailed engineering and equipment availability.",
             italic=True, color=RGBColor(0x66,0x66,0x66), size=9)
    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 11 – PROJECT SCHEDULE
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "11. PLANNED PROJECT SCHEDULE")
    sched_rows = [
        ["Detailed Survey of Land & Topography", "✓", "", "", ""],
        ["Finalization of Plant Structure & Design", "✓", "✓", "", ""],
        ["Preparation of Detailed Implementation Plan", "✓", "", "", ""],
        ["Firming up EPC Contract & Equipment Orders", "", "✓", "", ""],
        ["Land Preparation & Levelling", "", "✓", "✓", ""],
        ["Control Buildings & Cable Trenching", "", "", "✓", "✓"],
        ["Module Mounting Structures Installation", "", "", "✓", "✓"],
        ["Solar PV Module Installation", "", "", "✓", "✓"],
        ["Inverter & HT Electrical Equipment", "", "", "", "✓"],
        ["SCADA & Monitoring System", "", "", "", "✓"],
        ["Testing, Commissioning & COD", "", "", "", "✓"],
    ]
    add_table(doc, ["Activity", "Month 1", "Month 2", "Month 3", "Month 4"], sched_rows,
              col_widths_cm=[7.5, 2, 2, 2, 2])
    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 11A – REGULATORY APPROVALS
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "11A. REGULATORY APPROVALS & CLEARANCES")
    para_run(doc, f"The development of this {f2(p['project_capacity_ac'])} MW solar power project requires the following "
             "regulatory approvals and clearances from central and state government authorities.", space_after=8)
    add_heading(doc, "Key Approvals Required", level=2)
    approvals = [
        ["TEDA Registration & Consent", "Month 1–2", f"TEDA, {p['location_state']}"],
        ["Land Registration & Title", "Month 1–3", "District Revenue Dept."],
        ["Power Evacuation Agreement", "Month 2–4", "TANGEDCO/TNEB"],
        ["CEIG Approval (HT Electrical)", "Month 3–5", "CEIG, Tamil Nadu"],
        ["Environmental Consent (TNPCB)", "Month 2–4", "TNPCB"],
        ["33 kV Transmission Line Permission", "Month 3–6", "TANGEDCO"],
        ["Building Permit (Control Room)", "Month 2–4", "Local Body"],
        ["PPA Execution", "Month 1–3", f"Off-taker & {p['company_short']}"],
        ["Financial Close", "Month 4–7", f"Lenders & {p['company_short']}"],
        ["EPC Contract Award", "Month 6–8", p["company_short"]],
        ["COD (Commissioning)", f"{p['cod_month']} {p['cod_year']}", p["company_short"]],
    ]
    add_table(doc, ["Approval / Clearance", "Expected Timeline", "Authority"], approvals,
              col_widths_cm=[7, 3.5, 5])
    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 11B – O&M & RISK
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "11B. OPERATION & MAINTENANCE PLAN AND RISK ANALYSIS")
    add_heading(doc, "O&M Cost Structure", level=2)
    for item in [
        f"Year 1 (post-COD): Free of cost – included in EPC warranty",
        f"Year 2 onwards: ₹ {p['om_rate_lac_per_mwac']:.1f} Lakhs/MWac = ₹ {f2(p['om_rate_lac_per_mwac']*p['project_capacity_ac'])} Lakhs/year (base); {p['om_escalation_pct']:.1f}% escalation p.a.",
        f"Insurance: {p['insurance_pct']:.2f}% of project cost = ₹ {f2(p['project_cost_lac']*p['insurance_pct']/100)} Lakhs (Year 1); {p['insurance_esc_pct']:.1f}% escalation p.a.",
    ]:
        add_bullet(doc, item)

    add_heading(doc, "O&M Scope", level=2)
    for item in [
        "Preventive maintenance: Scheduled inspection, cleaning, and servicing of all components",
        "Panel cleaning: Regular module cleaning to maintain PR – monthly/bi-monthly as required",
        "Corrective maintenance: Target repair time <4 hours (critical), <24 hours (others)",
        "Performance monitoring: SCADA-based real-time tracking; deviations >5% investigated immediately",
        "24×7 security and surveillance; perimeter fencing and CCTV",
        "Annual Maintenance Contracts (AMC) with OEMs for inverters, transformers, and SCADA",
    ]:
        add_bullet(doc, item)

    add_heading(doc, "Risk Matrix", level=2)
    risk_rows = [
        ["Revenue Risk", "Solar resource variability", "P50 estimate; conservative degradation; 6-month DSRA"],
        ["Off-take Risk", "PPA counterparty default", "Creditworthy off-taker; lender step-in rights"],
        ["Construction Risk", "Cost/schedule overrun", "Fixed-price EPC; Tier-1 contractor; contingency budget"],
        ["Equipment Risk", "Module/inverter failure", "Tier-1 supplier warranties; on-site spare; equipment insurance"],
        ["Interest Rate Risk", "Rate increase", f"Fixed rate {p['debt_interest_rate']:.1f}% for full {p['debt_tenor_yrs']}-year tenor"],
        ["Regulatory Risk", "Policy/grid charge changes", "Long-term PPA provides price certainty"],
        ["Natural Disaster", "Cyclone/flood/earthquake", "170 km/h wind-rated structures; comprehensive insurance"],
        ["O&M Cost Risk", "Higher-than-expected costs", f"{p['om_escalation_pct']:.1f}% escalation assumed; 3-year fixed-price O&M contract"],
    ]
    add_table(doc, ["Risk Category", "Risk Description", "Mitigation"], risk_rows,
              col_widths_cm=[3.5, 5, 7])
    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 12 – PROJECT COSTING
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "12. PROJECT COSTING")
    pp = doc.add_paragraph()
    pp.paragraph_format.space_after = Pt(8)
    r = pp.add_run(f"A. {f2(p['project_capacity_dc'])} MWp DC / {f2(p['project_capacity_ac'])} MW AC System – "
                   "Ground Mounted Structure with Monocrystalline Cells")
    r.bold = True; r.italic = True; r.font.size = Pt(11); r.font.name = "Arial"

    # Cost breakdown scaled to project size
    pc = p["project_cost_lac"]
    cost_rows = [
        ["1", f"Solar PV Modules – {f2(p['project_capacity_dc'])} MWp", "MWp", f"{f2(p['project_capacity_dc'])}", "5%",
         f2(pc*0.332), f2(pc*0.332*0.05), f2(pc*0.332*1.05)],
        ["2", "Inverters (String/Central)", "Nos", "LS", "18%",
         f2(pc*0.095), f2(pc*0.095*0.18), f2(pc*0.095*1.18)],
        ["3", "Module Mounting Structures (GI)", "MT", "LS", "18%",
         f2(pc*0.090), f2(pc*0.090*0.18), f2(pc*0.090*1.18)],
        ["4", "HT/LT & DC Cables", "Lot", "1", "18%",
         f2(pc*0.038), f2(pc*0.038*0.18), f2(pc*0.038*1.18)],
        ["5", "Transformers (33/0.69 kV)", "Nos", "4", "18%",
         f2(pc*0.038), f2(pc*0.038*0.18), f2(pc*0.038*1.18)],
        ["6", "HT Switchyard, VCB & Metering", "Set", "1", "18%",
         f2(pc*0.029), f2(pc*0.029*0.18), f2(pc*0.029*1.18)],
        ["7", "Civil Works (Foundation, Roads, Fencing)", "Lot", "1", "18%",
         f2(pc*0.063), f2(pc*0.063*0.18), f2(pc*0.063*1.18)],
        ["8", "SCADA & Monitoring", "Set", "1", "18%",
         f2(pc*0.014), f2(pc*0.014*0.18), f2(pc*0.014*1.18)],
        ["9", "Earthing & Lightning Protection", "Lot", "1", "18%",
         f2(pc*0.010), f2(pc*0.010*0.18), f2(pc*0.010*1.18)],
        ["10", f"Evacuation Line ({p['tneb_distance_km']:.1f} km, 33 kV)", "Km",
         f"{p['tneb_distance_km']:.1f}", "18%",
         f2(pc*0.024), f2(pc*0.024*0.18), f2(pc*0.024*1.18)],
        ["11", "Land Development & Fencing", "Lot", "1", "18%",
         f2(pc*0.016), f2(pc*0.016*0.18), f2(pc*0.016*1.18)],
        ["12", "Transportation, Erection & Commissioning", "Lot", "1", "18%",
         f2(pc*0.040), f2(pc*0.040*0.18), f2(pc*0.040*1.18)],
        ["13", "Pre-operative Expenses, IDC, Margin", "Lot", "1", "N/A",
         f2(pc*0.051), "–", f2(pc*0.051)],
        ["", "TOTAL PROJECT COST", "", "", "",
         "–", "–", f2(pc)],
    ]
    add_table(doc,
              ["Sl.", "Description", "Unit", "Qty", "GST%", "Basic (₹ Lac)", "GST (₹ Lac)", "Total (₹ Lac)"],
              cost_rows, col_widths_cm=[0.7, 5.8, 1.0, 0.8, 0.8, 2.2, 2.0, 2.2])

    add_heading(doc, "Financing Structure", level=2)
    fin_rows = [
        ["Total Project Cost", f"₹ {p['project_cost_cr']:.2f} Crores", "100%"],
        [f"Debt (@ {p['debt_interest_rate']:.1f}% p.a.)", f"₹ {p['debt_cr']:.2f} Crores", f"{p['debt_pct']:.0f}%"],
        ["Equity", f"₹ {p['equity_cr']:.2f} Crores", f"{100-p['debt_pct']:.0f}%"],
    ]
    add_table(doc, ["Component", "Amount", "Percentage"], fin_rows, col_widths_cm=[5, 5, 3])
    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 13 – FINANCIAL ASSUMPTIONS
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "13. PROJECT FINANCIALS – KEY ASSUMPTIONS")
    assump_rows = [
        ["Installed Capacity (AC)", f"{f2(p['project_capacity_ac'])} MW AC"],
        ["Installed Capacity (DC)", f"{f2(p['project_capacity_dc'])} MWp"],
        ["Year 1 Generation", f"{f2(p['gen_yr1_lac_per_mwac'] * p['project_capacity_ac'])} Lac Units"],
        ["Degradation – Year 1", f"{p['degradation_yr1_pct']:.1f}%"],
        ["Degradation – Year 2 onwards", f"{p['degradation_yr2_pct']:.1f}% p.a."],
        ["PPA Tariff", f"₹ {p['ppa_tariff']:.2f} per kWh (Flat, {p['ppa_term']} years)"],
        ["O&M – Year 1", "Nil (Free)"],
        ["O&M – Year 2 base", f"₹ {f2(p['om_rate_lac_per_mwac']*p['project_capacity_ac'])} Lakhs; {p['om_escalation_pct']:.1f}% escalation"],
        ["Insurance", f"{p['insurance_pct']:.2f}% of cost p.a.; {p['insurance_esc_pct']:.1f}% escalation"],
        ["Book Depreciation", f"SLM over {p['ppa_term']} years = ₹ {f2(m['dep_slm'])} Lakhs/year"],
        ["Tax Depreciation", "WDV @ 40% p.a."],
        ["Tax Regime", "Old Regime with MAT (effective MAT ~16.69%; normal ~33.38%)"],
        ["Debt", f"₹ {f2(p['debt_lac'])} Lakhs @ {p['debt_interest_rate']:.1f}% p.a."],
        ["Debt Tenor", f"{p['debt_tenor_yrs']} years (straight-line repayment)"],
        ["Annual Principal", f"₹ {f2(m['ann_principal'])} Lakhs"],
        ["Equity", f"₹ {f2(p['equity_lac'])} Lakhs"],
    ]
    add_table(doc, ["Parameter", "Assumption"], assump_rows, col_widths_cm=[6.5, 9])

    # ══════════════════════════════════════════════════════════════════════════
    # 25-YEAR CASH FLOW TABLE
    # ══════════════════════════════════════════════════════════════════════════
    para_run(doc, "25-Year Projected Cash Flow Statement (₹ in Lakhs)",
             bold=True, size=12, color=C_NAVY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)

    cf_headers = ["Year", "Gen\n(Lac U)", "Revenue", "O&M", "Insurance",
                  "EBITDA", "Depn.", "Interest", "PBT", "Tax", "PAT",
                  "Principal", "DSCR"]
    cf_widths   = [0.8, 1.4, 1.7, 1.5, 1.5, 1.7, 1.4, 1.5, 1.6, 1.5, 1.5, 1.5, 1.2]

    table = doc.add_table(rows=1 + len(cfs), cols=len(cf_headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Header
    for i, h in enumerate(cf_headers):
        cell = table.rows[0].cells[i]
        set_cell_bg(cell, "1F3864")
        set_cell_border(cell)
        set_col_width(cell, cf_widths[i])
        p2 = cell.paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_before = Pt(2)
        p2.paragraph_format.space_after  = Pt(2)
        run = p2.add_run(h)
        run.bold = True; run.font.size = Pt(7.5)
        run.font.name = "Arial"; run.font.color.rgb = C_WHITE

    for ri, cf in enumerate(cfs):
        bg = "EBF3FB" if ri % 2 == 1 else "FFFFFF"
        dscr_str = f2(cf["dscr"]) if cf["year"] <= p["debt_tenor_yrs"] else "N/A"
        vals = [
            str(cf["year"]), f2(cf["gen"]), f2(cf["revenue"]),
            f2(cf["om"]), f2(cf["insurance"]), f2(cf["ebitda"]),
            f2(cf["depreciation"]), f2(cf["interest"]), f2(cf["pbt"]),
            f2(cf["tax"]), f2(cf["pat"]), f2(cf["principal"]), dscr_str,
        ]
        row_cells = table.rows[ri + 1].cells
        for ci, val in enumerate(vals):
            cell = row_cells[ci]
            set_cell_bg(cell, bg)
            set_cell_border(cell)
            set_col_width(cell, cf_widths[ci])
            p2 = cell.paragraphs[0]
            p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p2.paragraph_format.space_before = Pt(1)
            p2.paragraph_format.space_after  = Pt(1)
            run = p2.add_run(val)
            run.font.size = Pt(7.5); run.font.name = "Arial"
            # Colour DSCR
            if ci == 12 and cf["year"] <= p["debt_tenor_yrs"]:
                try:
                    dscr_val = cf["dscr"]
                    run.font.color.rgb = C_GREEN if dscr_val >= 1.25 else RGBColor(0xC0,0x00,0x00)
                    run.bold = True
                except: pass

    doc.add_paragraph()
    para_run(doc, "Note: DSCR = (EBITDA – Tax) / (Principal + Interest). Figures in ₹ Lakhs. Generation in Lac Units.",
             italic=True, size=8.5, color=RGBColor(0x55,0x55,0x55))
    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 14 – DEBT REPAYMENT & DSCR
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "14. DEBT REPAYMENT SCHEDULE & DSCR ANALYSIS")
    para_run(doc, f"Term debt of ₹ {f2(p['debt_lac'])} Lakhs at {p['debt_interest_rate']:.1f}% p.a., repayable over "
             f"{p['debt_tenor_yrs']} years from COD on a straight-line basis. "
             f"Annual principal: ₹ {f2(m['ann_principal'])} Lakhs.", space_after=8)

    debt_rows = []
    for cf in cfs:
        if cf["year"] > p["debt_tenor_yrs"]:
            break
        debt_rows.append([
            str(cf["year"]), f2(cf["op_debt"]), f2(cf["principal"]),
            f2(cf["interest"]), f2(cf["debt_service"]), f2(cf["cl_debt"]),
            f2(cf["dscr"]),
        ])
    add_table(doc, ["Year", "Opening Debt", "Principal", "Interest", "Total Debt Svc", "Closing Debt", "DSCR"],
              debt_rows, col_widths_cm=[1.2, 2.8, 2.4, 2.4, 2.8, 2.8, 1.8])

    # DSCR Summary box
    dscr_rows = [
        ["Minimum DSCR (over loan tenor)", f"{min_dscr:.2f}x"],
        ["Average DSCR (over loan tenor)", f"{avg_dscr:.2f}x"],
    ]
    add_table(doc, ["DSCR Metric", "Value"], dscr_rows, col_widths_cm=[8, 4], header_bg="2E74B5")
    para_run(doc, f"Both minimum DSCR ({min_dscr:.2f}x) and average DSCR ({avg_dscr:.2f}x) are well above the "
             "standard lender threshold of 1.10x, confirming comfortable debt serviceability.", space_after=10)
    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 15 – SUMMARY OF RESULTS
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "15. SUMMARY OF RESULTS")
    summary_rows = [
        ["Project Capacity (AC)", f"{f2(p['project_capacity_ac'])} MW AC"],
        ["Project Capacity (DC)", f"{f2(p['project_capacity_dc'])} MWp"],
        ["Total Project Cost", f"₹ {p['project_cost_cr']:.2f} Crores (₹ {f2(p['project_cost_lac'])} Lakhs)"],
        ["Debt", f"₹ {p['debt_cr']:.2f} Crores @ {p['debt_interest_rate']:.1f}% p.a."],
        ["Equity", f"₹ {p['equity_cr']:.2f} Crores"],
        ["Year 1 Generation", f"{f2(p['gen_yr1_lac_per_mwac'] * p['project_capacity_ac'])} Lac Units"],
        ["Year 1 Revenue", f"₹ {f2(cfs[0]['revenue'])} Lakhs"],
        ["Year 1 EBITDA", f"₹ {f2(cfs[0]['ebitda'])} Lakhs"],
        ["PPA Tariff", f"₹ {p['ppa_tariff']:.2f}/unit (Flat, {p['ppa_term']} years)"],
        ["Project IRR (Post-Tax)", f"{proj_irr:.2f}%"],
        ["Equity IRR (Post-Tax)", f"{eq_irr:.2f}%"],
        ["Minimum DSCR", f"{min_dscr:.2f}x"],
        ["Average DSCR", f"{avg_dscr:.2f}x"],
        ["COD", f"{p['cod_month']} {p['cod_year']}"],
        ["PPA Term", f"{p['ppa_term']} Years"],
    ]
    add_table(doc, ["Parameter", "Value"], summary_rows, col_widths_cm=[7, 9], header_bg="1F3864")
    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 16 – CONCLUSION
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "16. CONCLUSION")
    para_run(doc, f"The detailed financial analysis for the {f2(p['project_capacity_ac'])} MW AC "
             f"({f2(p['project_capacity_dc'])} MWp DC) ground-mounted solar power project by "
             f"M/s. {p['company_name']} ({p['company_short']}) at {p['location_village']}, "
             f"{p['location_district']} District, {p['location_state']}, demonstrates that the project is "
             "both technically feasible and financially attractive.", space_after=8)

    add_heading(doc, "Key Findings:", level=2)
    for item in [
        f"Project IRR of {proj_irr:.2f}% (post-tax) over {p['ppa_term']}-year PPA term – well above cost of capital",
        f"Equity IRR of {eq_irr:.2f}% – strong return for {p['company_short']} shareholders",
        f"Minimum DSCR of {min_dscr:.2f}x – above lender threshold of 1.10x throughout {p['debt_tenor_yrs']}-year loan",
        f"Average DSCR of {avg_dscr:.2f}x – significant comfort for lenders",
        f"Flat PPA tariff of ₹ {p['ppa_tariff']:.2f}/unit for {p['ppa_term']} years – revenue certainty, protects against volatility",
        f"AC:DC ratio of {p['ac_dc_ratio']:.1f} optimises generation yield with appropriate inverter loading",
        f"Site at {p['location_village']}, {p['location_district']} has excellent solar resource (GHI 5.5–5.8 kWh/m²/day)",
        f"Total project cost of ₹ {p['project_cost_cr']:.2f} Crores represents "
        f"₹ {p['project_cost_cr']/p['project_capacity_ac']:.1f} Crores/MWac – competitive capital cost",
    ]:
        add_bullet(doc, item)

    para_run(doc, "As seen from the above financial summary, the project has a healthy DSCR for the entire duration "
             "of the debt. Debt servicing is feasible. The project also provides a steady and attractive return "
             "on investment for the equity employed.", space_after=8)
    para_run(doc, "With the above assumptions and financial results, it is clear that the project is Technically "
             "and Financially viable and is worth investing. The project is recommended for approval and funding.",
             bold=True, space_after=16)

    pp = doc.add_paragraph()
    pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = pp.add_run("***************************")
    r.font.color.rgb = C_NAVY; r.font.name = "Arial"; r.font.size = Pt(12)

    pp = doc.add_paragraph()
    pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = pp.add_run(f"{p['company_name'].upper()}")
    r.bold = True; r.font.size = Pt(13); r.font.color.rgb = C_NAVY; r.font.name = "Arial"

    pp = doc.add_paragraph()
    pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = pp.add_run(p["company_address"])
    r.font.size = Pt(10); r.font.name = "Arial"

    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # ANNEXURE A – PVSYST
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "ANNEXURE – A: SOLAR RESOURCE ASSESSMENT & PVSYST SIMULATION")
    para_run(doc, f"This Annexure contains the solar resource assessment data and PVSyst energy yield simulation "
             f"results for the {f2(p['project_capacity_ac'])} MW AC project at {p['location_village']}, "
             f"{p['location_district']} District, {p['location_state']}.", space_after=8)

    add_heading(doc, "A.1 Monthly Average GHI (kWh/m²/day)", level=2)
    ghi_rows = [["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec","Annual"],
                ["5.80","6.10","6.40","6.20","5.90","4.80","4.50","4.70","5.20","5.60","5.40","5.50","5.51 avg"]]
    add_table(doc, ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec","Annual"],
              [ghi_rows[1]], col_widths_cm=[1.2]*12+[1.5])

    add_heading(doc, "A.2 PVSyst Simulation Parameters", level=2)
    pv_rows = [
        ["Module Type", "Monocrystalline Si – 550–580 Wp"],
        ["Tilt Angle", "12°–15° (Fixed South-facing)"],
        ["DC/AC Ratio (ILR)", f"{p['ac_dc_ratio']:.1f}"],
        ["Soiling Loss", "2.0% p.a."],
        ["Mismatch Loss", "2.0%"],
        ["DC Wiring Loss", "1.5%"],
        ["Transformer Loss", "1.0%"],
        ["System Availability", "98.5%"],
        ["Performance Ratio (PR)", "~78%–80%"],
        ["Specific Yield", f"~{p['gen_yr1_lac_per_mwac']*100/p['project_capacity_dc']:.0f} kWh/kWp/year"],
        ["P50 Annual Generation (Year 1)", f"~{f2(p['gen_yr1_lac_per_mwac']*p['project_capacity_ac'])} Lac Units"],
        ["P90 Annual Generation (Year 1)", f"~{p['gen_yr1_lac_per_mwac']*p['project_capacity_ac']*0.953:.0f} Lac Units"],
    ]
    add_table(doc, ["Parameter", "Value"], pv_rows, col_widths_cm=[7, 9])

    add_heading(doc, "A.3 Month-wise Generation Estimate (Year 1)", level=2)
    gen_total_yr1 = p['gen_yr1_lac_per_mwac'] * p['project_capacity_ac']
    monthly_pct = [8.95,8.09,9.05,8.72,8.65,6.77,6.70,7.02,7.53,8.47,7.86,8.47]
    mw_rows = [[m, f"{gen_total_yr1*pct/100:.2f}", f"{pct:.1f}%"]
               for m, pct in zip(["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"], monthly_pct)]
    mw_rows.append(["Annual Total", f"{gen_total_yr1:.2f}", "100%"])
    add_table(doc, ["Month", "Generation (Lac Units)", "% of Annual"], mw_rows,
              col_widths_cm=[3, 4, 3])

    para_run(doc, "Note: Monthly generation estimates are illustrative based on typical solar profile for this region. "
             "Replace with actual PVSyst simulation report upon completion of detailed site-specific assessment.",
             italic=True, size=9, color=RGBColor(0x55,0x55,0x55))

    add_heading(doc, "A.4 25-Year Generation Summary", level=2)
    gen25_rows = [[str(cf["year"]), f"{cf['gen']:.2f}",
                   f"{cf['gen']*p['ppa_tariff']:.2f}"] for cf in cfs]
    add_table(doc, ["Year", "Generation (Lac Units)", "Revenue (₹ Lakhs)"],
              gen25_rows, col_widths_cm=[2, 4, 4])

    pp = doc.add_paragraph()
    pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = pp.add_run("– End of Detailed Project Report –")
    r.bold = True; r.italic = True; r.font.size = Pt(13)
    r.font.color.rgb = C_NAVY; r.font.name = "Arial"

    # ── Save to bytes ─────────────────────────────────────────────────────────
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()
