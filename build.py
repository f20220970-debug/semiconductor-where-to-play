"""Build semiconductor_where_to_play.xlsx: where an automation company should play in the chip value chain.

Segments are sized with SEMI / WSTS 2025 data, geographies with 2025 equipment billings,
then 11 segment-geography plays are scored on attractiveness vs right to win.
Blue = sourced, yellow = assumption or judgment, black = formula. Run: python build.py
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import ScatterChart, BarChart, Reference, Series
from openpyxl.chart.marker import Marker
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter as L

F = "Arial"
BLUE = Font(name=F, color="0000FF")
BLACK = Font(name=F)
GREEN = Font(name=F, color="008000")
BOLD = Font(name=F, bold=True)
WHITE_B = Font(name=F, bold=True, color="FFFFFF")
TITLE = Font(name=F, bold=True, size=14)
HDR = PatternFill("solid", fgColor="1F3A5F")
YEL = PatternFill("solid", fgColor="FFFF00")
thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(wrap_text=True, vertical="top")
BN = '#,##0.0'
PCT = '0.0%'
CR = '#,##0'

wb = Workbook()


def header(ws, row, labels, widths=None):
    for i, lab in enumerate(labels, 1):
        c = ws.cell(row=row, column=i, value=lab)
        c.font, c.fill, c.border = WHITE_B, HDR, BOX
        c.alignment = Alignment(wrap_text=True, vertical="center")
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[L(i)].width = w


def put(ws, ref, value, font=BLACK, fmt=None, fill=None):
    c = ws[ref]
    c.value, c.font, c.border = value, font, BOX
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    return c


# ---------------------------------------------------------------- README
ws = wb.active
ws.title = "README"
ws.column_dimensions["A"].width = 115
for i, (t, f) in enumerate([
    ("Semiconductor where-to-play map for an industrial-automation company", TITLE),
    ("Built by Pavitra Seth (BITS Pilani) as preparation for Honeywell PS-II. Public data only; not affiliated with Honeywell.", BLACK),
    ("", BLACK),
    ("Question: chips are a ~$800bn market and governments are paying for new fabs. Which parts of the value chain, in which countries, "
     "should an automation company (controls, safety, gas detection, building automation, software) go after, and how?", BOLD),
    ("", BLACK),
    ("Tabs: Capabilities (what Honeywell can bring, today's portfolio) > Segments (value chain sized with 2025 data) > Geographies "
     "(where fabs are being equipped) > Plays (11 segment x country plays scored) > Matrix (bubble chart) > India (12 approved projects and the "
     "facility-automation pool) > Sensitivity (do the top plays survive other weights?) > Sources.", BLACK),
    ("Colour code: blue = sourced; yellow = assumption or analyst score (change it); black = formula; green = link.", BLACK),
    ("Scores are 1-5 with a one-line evidence note each. They are judgment anchored on the data tabs, not measured facts.", BLACK),
], 1):
    c = ws.cell(row=i, column=1, value=t)
    c.font = f
    c.alignment = Alignment(wrap_text=True)

# ---------------------------------------------------------------- Capabilities
cap = wb.create_sheet("Capabilities")
cap["A1"] = "What an automation company like Honeywell Technologies can bring to a fab (check against the latest 10-K)"
cap["A1"].font = BOLD
header(cap, 3, ["Capability", "Still in Honeywell after the 2025-26 spin-offs?", "Relevance to fabs", "Evidence / note"], [36, 22, 40, 70])
caps = [
    ("Gas detection and toxic-gas monitoring", "Yes", "Fabs use toxic and pyrophoric gases (arsine, phosphine, silane); continuous monitoring is mandatory", "Honeywell gas detection portfolio (e.g. Midas, Vertex lines used in fabs) - verify current product names"),
    ("Building automation and cleanroom controls", "Yes", "Cleanrooms need tight temperature, humidity and pressure control 24/7", "Building Automation segment: $7.4bn FY2025, +8% organic"),
    ("Process control systems and safety systems", "Yes", "Utilities, ultrapure water, chemical and gas delivery, abatement", "Industrial Automation: Experion control and safety systems"),
    ("Industrial software and analytics", "Partly", "Energy optimisation and predictive maintenance for fab utilities", "Honeywell Forge; no fab MES product"),
    ("Specialty electronic materials (sputtering targets, etc.)", "No - moved with Solstice (Oct 2025)", "Wafer fab materials", "Solstice Advanced Materials spin-off, 30 Oct 2025"),
    ("Wafer fab equipment (tools)", "No", "Lithography, etch, deposition", "Not a Honeywell business"),
]
for i, row in enumerate(caps):
    for j, v in enumerate(row):
        c = cap.cell(row=4 + i, column=1 + j, value=v)
        c.font, c.alignment, c.border = (BOLD if j == 0 else BLACK), WRAP, BOX
    cap.row_dimensions[4 + i].height = 40

# ---------------------------------------------------------------- Segments
sg = wb.create_sheet("Segments")
sg["A1"] = "Value chain sized with 2025 data (US$ bn)"
sg["A1"].font = TITLE
header(sg, 3, ["Segment", "2025 size (US$ bn)", "2025 growth", "Type", "How sized / source"], [44, 14, 11, 12, 90])
put(sg, "A4", "Equipment share of fab capex", BOLD)
put(sg, "B4", 0.75, BLUE, PCT, YEL)
put(sg, "E4", "Construction Physics via Quartz: machines are 70-80% of a modern fab's cost; building and facility 20-30%")
put(sg, "A5", "Automation, controls and safety share of facility spend", BOLD)
put(sg, "B5", 0.10, BLUE, PCT, YEL)
put(sg, "E5", "Assumption: BMS, cleanroom controls, gas detection, safety systems and utility control within facility capex")
segs = [
    ("Total chip market (context)", 795.6, 0.262, "Source", "WSTS: global semiconductor sales $795.6bn in 2025, +26.2%"),
    ("Semiconductor equipment (all tools)", 135.1, 0.15, "Source", "SEMI: global equipment billings $135.1bn in 2025, +15%"),
    ("Wafer fab materials", 45.8, 0.054, "Source", "SEMI: $45.8bn in 2025, +5.4%"),
    ("Packaging materials", 27.4, 0.093, "Source", "SEMI: $27.4bn in 2025, +9.3%"),
    ("Fab building and facility systems", "=B9/$B$4*(1-$B$4)", "=C9", "Derived", "Equipment / equipment share x facility share; grows with equipment spend"),
    ("  of which: automation, controls and safety", "=B12*$B$5", "=C12", "Derived", "Facility x automation share: the pool closest to Honeywell"),
]
for i, (n, v, g, t, s) in enumerate(segs):
    rr = 8 + i
    put(sg, f"A{rr}", n, BOLD)
    put(sg, f"B{rr}", v, BLUE if t == "Source" else BLACK, BN)
    put(sg, f"C{rr}", g, BLUE if t == "Source" else BLACK, PCT)
    put(sg, f"D{rr}", t)
    put(sg, f"E{rr}", s).alignment = WRAP
header(sg, 7, ["Segment", "2025 size (US$ bn)", "2025 growth", "Type", "How sized / source"])
put(sg, "A15", "Automation pool as share of equipment market", BOLD)
put(sg, "B15", "=B13/B9", BLACK, PCT)
put(sg, "E15", "Small next to tools, but it is the part an automation company can actually win")

# ---------------------------------------------------------------- Geographies
ge = wb.create_sheet("Geographies")
ge["A1"] = "Where fabs are being equipped: 2025 equipment billings (SEMI) plus policy and access"
ge["A1"].font = TITLE
header(ge, 3, ["Geography", "2025 equipment billings (US$ bn)", "Growth 2025", "Share of world", "Implied automation pool (US$ bn)",
               "Policy tailwind (1-5)", "Access for a US company (1-5)", "Evidence"], [16, 14, 11, 11, 14, 11, 12, 70])
geos = [
    ("China", 49.3, -0.005, 3, 1, "Largest market but export controls and local suppliers limit a US company"),
    ("Taiwan", 31.5, 0.90, 4, 4, "Billings +90%; TSMC capacity build-out"),
    ("South Korea", 25.8, 0.26, 4, 3, "Memory and HBM expansion; strong local suppliers"),
    ("North America", 10.9, -0.20, 5, 5, "CHIPS Act: $30.9bn direct funding across 40 fab projects (CRS); billings fell 20% in 2025 as projects slipped"),
    ("Japan", 9.5, 0.22, 4, 4, "Subsidised fabs (e.g. TSMC Kumamoto, Rapidus)"),
    ("Europe", 2.9, -0.41, 3, 4, "Billings -41% in 2025; EU Chips Act projects delayed"),
    ("Rest of world (incl. India, SE Asia)", 5.2, 0.25, 5, 4, "India: 12 approved projects worth Rs 1.64 lakh crore; Malaysia and Singapore back-end"),
]
for i, (n, b, g, pol, acc, ev) in enumerate(geos):
    rr = 4 + i
    put(ge, f"A{rr}", n, BOLD)
    put(ge, f"B{rr}", b, BLUE, BN)
    put(ge, f"C{rr}", g, BLUE, PCT)
    put(ge, f"D{rr}", f"=B{rr}/SUM($B$4:$B$10)", BLACK, PCT)
    put(ge, f"E{rr}", f"=D{rr}*Segments!$B$13", BLACK, '0.00')
    put(ge, f"F{rr}", pol, BLUE, "0", YEL)
    put(ge, f"G{rr}", acc, BLUE, "0", YEL)
    put(ge, f"H{rr}", ev).alignment = WRAP
put(ge, "A11", "Total", BOLD)
put(ge, "B11", "=SUM(B4:B10)", BLACK, BN)
put(ge, "E11", "=SUM(E4:E10)", BLACK, '0.00')
put(ge, "H11", "Check: total should equal SEMI's $135.1bn")

# ---------------------------------------------------------------- Plays
pl = wb.create_sheet("Plays")
pl["A1"] = "11 plays scored (1-5). Attractiveness and right to win are weighted averages; weights in yellow."
pl["A1"].font = BOLD
pl.merge_cells("A1:P1")
wA = [("Pool size", 0.30), ("Growth", 0.25), ("Policy tailwind", 0.15), ("Access", 0.15), ("Profit pool", 0.15)]
wR = [("Portfolio fit", 0.40), ("Customer access", 0.25), ("Installed base", 0.20), ("Low competition", 0.15)]
put(pl, "A3", "Weights", BOLD)
for j, (n, w) in enumerate(wA + wR):
    put(pl, f"{L(3 + j)}3", w, BLUE, '0%', YEL)
labels = ["Play", "Segment / geography"] + [n for n, _ in wA] + [n for n, _ in wR] + ["Attractiveness", "Right to win", "Pool (US$ bn)", "Quadrant", "Evidence for scores"]
header(pl, 4, labels, [34, 26] + [10] * 9 + [13, 12, 11, 12, 70])
plays = [
    ("Fab facility automation & safety: US", "Facility / North America", [5, 2, 5, 5, 4], [5, 4, 4, 3], "=Geographies!E7", "CHIPS-funded fabs; Honeywell home market; billings fell in 2025 as projects slipped"),
    ("Fab facility automation & safety: Taiwan", "Facility / Taiwan", [5, 5, 4, 4, 4], [5, 3, 3, 3], "=Geographies!E5", "Fastest growth (+90%); strong local integrators"),
    ("Fab facility automation & safety: Japan", "Facility / Japan", [3, 4, 4, 4, 4], [5, 3, 3, 3], "=Geographies!E8", "Subsidised new fabs; demanding local suppliers"),
    ("Fab facility automation & safety: India", "Facility / India", [2, 5, 5, 4, 3], [5, 4, 3, 4], "=India!F18*10/88/1000", "Rs 1.64 lakh crore approved; few incumbents; Honeywell Automation India already local. Pool here is spread over the 2025-28 build, not one year"),
    ("Fab facility automation & safety: Korea", "Facility / Korea", [4, 4, 4, 3, 4], [5, 2, 3, 2], "=Geographies!E6", "Samsung and SK hynix favour local suppliers"),
    ("Fab facility automation & safety: China", "Facility / China", [5, 1, 3, 1, 3], [5, 2, 3, 1], "=Geographies!E4", "Biggest pool, flat growth, export-control and localisation risk"),
    ("Fab facility automation & safety: Europe", "Facility / Europe", [1, 1, 3, 4, 4], [5, 3, 3, 3], "=Geographies!E9", "Billings -41%; projects delayed"),
    ("ATMP / OSAT facility automation: India", "Back-end facility / India", [2, 5, 5, 4, 3], [4, 4, 3, 4], "=India!F19*10/88/1000", "9 of 12 Indian projects are packaging units"),
    ("ATMP / OSAT facility automation: SE Asia", "Back-end facility / SE Asia", [3, 4, 3, 4, 3], [4, 3, 3, 3], 0.2, "Malaysia and Singapore back-end clusters (pool is my estimate)"),
    ("Wafer fab materials: Taiwan", "Materials / Taiwan", [5, 2, 3, 3, 4], [1, 2, 1, 2], "=Segments!B10*0.3", "Materials business left with Solstice; Taiwan ~30% of materials demand"),
    ("Wafer fab equipment: Korea", "Equipment / Korea", [5, 4, 4, 3, 5], [1, 1, 1, 1], "=Geographies!B6", "No tool capability; incumbents dominate"),
]
for i, (name, seg, A, Rw, pool, ev) in enumerate(plays):
    rr = 5 + i
    put(pl, f"A{rr}", name, BOLD)
    put(pl, f"B{rr}", seg)
    for j, v in enumerate(A + Rw):
        put(pl, f"{L(3 + j)}{rr}", v, BLUE, "0", YEL)
    put(pl, f"L{rr}", f"=SUMPRODUCT(C{rr}:G{rr},$C$3:$G$3)/SUM($C$3:$G$3)", BLACK, '0.00')
    put(pl, f"M{rr}", f"=SUMPRODUCT(H{rr}:K{rr},$H$3:$K$3)/SUM($H$3:$K$3)", BLACK, '0.00')
    put(pl, f"N{rr}", pool, BLUE if isinstance(pool, float) else GREEN, '0.00', YEL if isinstance(pool, float) else None)
    put(pl, f"O{rr}", f'=IF(AND(L{rr}>=3.5,M{rr}>=3.5),"Invest",IF(L{rr}>=3.5,"Partner or buy",IF(M{rr}>=3.5,"Harvest","Avoid")))')
    put(pl, f"P{rr}", ev).alignment = WRAP
last = 5 + len(plays) - 1
pl.conditional_formatting.add(f"C5:K{last}", ColorScaleRule(start_type="num", start_value=1, start_color="F8696B", mid_type="num", mid_value=3, mid_color="FFFFFF", end_type="num", end_value=5, end_color="63BE7B"))
put(pl, f"A{last+2}", "Combined score (attractiveness x right to win)", BOLD)
header(pl, last + 3, ["Rank", "Play", "Combined", "Quadrant"], None)
for k in range(len(plays)):
    rr = last + 4 + k
    put(pl, f"A{rr}", k + 1, BLACK, "0")
    put(pl, f"C{rr}", f"=LARGE($Q$5:$Q${last},A{rr})", BLACK, '0.00')
    put(pl, f"B{rr}", f"=INDEX($A$5:$A${last},MATCH(C{rr},$Q$5:$Q${last},0))")
    put(pl, f"D{rr}", f"=INDEX($O$5:$O${last},MATCH(C{rr},$Q$5:$Q${last},0))")
for i in range(len(plays)):
    rr = 5 + i
    put(pl, f"Q{rr}", f"=L{rr}*M{rr}+ROW()/100000", BLACK, '0.00')
put(pl, "Q4", "Combined (helper)", BOLD)

# ---------------------------------------------------------------- Matrix chart
mx = wb.create_sheet("Matrix")
mx["A1"] = "Attractiveness vs right to win (each point = one play; labels in the table)"
mx["A1"].font = BOLD
header(mx, 3, ["Play", "Right to win", "Attractiveness"], [40, 12, 14])
for i in range(len(plays)):
    rr = 4 + i
    put(mx, f"A{rr}", f"=Plays!A{5+i}", GREEN)
    put(mx, f"B{rr}", f"=Plays!M{5+i}", GREEN, '0.00')
    put(mx, f"C{rr}", f"=Plays!L{5+i}", GREEN, '0.00')
sc = ScatterChart()
sc.title = "Where to play: Invest = top right"
sc.style = 13
sc.x_axis.title = "Right to win (1-5)"
sc.y_axis.title = "Attractiveness (1-5)"
sc.x_axis.scaling.min, sc.x_axis.scaling.max = 1, 5
sc.y_axis.scaling.min, sc.y_axis.scaling.max = 1, 5
palette = ["1F77B4", "FF7F0E", "2CA02C", "D62728", "9467BD", "8C564B", "E377C2", "7F7F7F", "BCBD22", "17BECF", "393B79"]
for i in range(len(plays)):
    rr = 4 + i
    s = Series(Reference(mx, min_col=3, min_row=rr), Reference(mx, min_col=2, min_row=rr), title_from_data=False, title=plays[i][0])
    s.marker = Marker(symbol="circle", size=11)
    s.marker.graphicalProperties.solidFill = palette[i]
    s.graphicalProperties.line.noFill = True
    sc.series.append(s)
sc.x_axis.delete = False
sc.y_axis.delete = False
sc.height, sc.width = 12, 22
mx.add_chart(sc, "E3")

# ---------------------------------------------------------------- India
ind = wb.create_sheet("India")
ind["A1"] = "India: 12 approved semiconductor projects and the facility-automation pool (INR crore)"
ind["A1"].font = TITLE
header(ind, 3, ["Project", "Type", "State", "Investment (INR cr)", "Facility share", "Automation pool (INR cr)", "Status (as reported)"], [36, 16, 14, 14, 11, 16, 50])
put(ind, "I3", "Automation share of facility", BOLD)
put(ind, "J3", "=Segments!B5", GREEN, PCT)
put(ind, "I4", "Facility share: fab", BOLD)
put(ind, "J4", "=1-Segments!B4", GREEN, PCT)
put(ind, "I5", "Facility share: packaging (ATMP/OSAT)", BOLD)
put(ind, "J5", 0.35, BLUE, PCT, YEL)
put(ind, "K5", "Assumption: back-end plants are less tool-heavy, so the building is a larger share")
ind.column_dimensions["I"].width = 30
proj = [
    ("Tata Electronics + PSMC", "Fab", "Gujarat", 91526, "Commissioning targeted 2028"),
    ("Micron Technology", "ATMP", "Gujarat", 22516, "Commercial production began 2026"),
    ("Tata Semiconductor Assembly & Test", "ATMP", "Assam", 27120, ""),
    ("CG Power + Renesas + Stars", "OSAT", "Gujarat", 7584, "Commercial production July 2026"),
    ("Kaynes Semicon", "OSAT", "Gujarat", 3307, "Inaugurated March 2026"),
    ("HCL-Foxconn JV (display driver ICs)", "OSAT", "Uttar Pradesh", 3706, ""),
    ("SiCSem (silicon carbide)", "Compound fab", "Odisha", 2066, ""),
    ("3D Glass Solutions", "ATMP", "Odisha", 1943, ""),
    ("CDIL (discrete power devices)", "OSAT", "Punjab", 117, ""),
    ("ASIP Technologies", "OSAT", "Andhra Pradesh", 480, ""),
    ("Crystal Matrix (GaN) + Suchi Semicon (combined)", "Compound fab / OSAT", "Gujarat", 3936, "Approved May 2026; combined figure"),
]
for i, (n, t, s, inv, st) in enumerate(proj):
    rr = 4 + i
    put(ind, f"A{rr}", n, BOLD)
    put(ind, f"B{rr}", t)
    put(ind, f"C{rr}", s)
    put(ind, f"D{rr}", inv, BLUE, CR)
    put(ind, f"E{rr}", f'=IF(ISNUMBER(SEARCH("fab",B{rr})),$J$4,$J$5)', BLACK, PCT)
    put(ind, f"F{rr}", f"=D{rr}*E{rr}*$J$3", BLACK, CR)
    put(ind, f"G{rr}", st)
le = 4 + len(proj) - 1  # 14
put(ind, f"A{le+1}", "Total", BOLD)
put(ind, f"D{le+1}", f"=SUM(D4:D{le})", BLACK, CR)
put(ind, f"F{le+1}", f"=SUM(F4:F{le})", BLACK, CR)
put(ind, f"G{le+1}", "Check: IANS reports a pipeline of about Rs 1.64 lakh crore across 12 projects")
put(ind, "A17", "Pool by type", BOLD)
put(ind, "A18", "Fabs (incl. compound)", BOLD)
put(ind, "F18", f'=SUMPRODUCT(ISNUMBER(SEARCH("fab",B4:B{le}))*F4:F{le})', BLACK, CR)
put(ind, "A19", "Packaging (ATMP / OSAT)", BOLD)
put(ind, "F19", f"=F{le+1}-F18", BLACK, CR)
put(ind, "A20", "Total pool in US$ m (INR 88/US$)", BOLD)
put(ind, "F20", f"=F{le+1}/88*10", BLACK, CR)
put(ind, "A21", "Share of pool in Gujarat", BOLD)
put(ind, "F21", f'=SUMIF(C4:C{le},"Gujarat",F4:F{le})/F{le+1}', BLACK, PCT)
put(ind, "A23", "First three customers to approach: Tata-PSMC Dholera fab, Micron Sanand ATMP, Tata Assam OSAT (largest pools; together most of the total).", BOLD)

# ---------------------------------------------------------------- Sensitivity
se = wb.create_sheet("Sensitivity")
se["A1"] = "Do the top plays survive different weights? (combined score under three weight sets)"
se["A1"].font = TITLE
header(se, 3, ["Play", "Base", "Growth-heavy", "Safety-first", "Rank base", "Rank growth", "Rank safety", "Spread"], [40, 10, 12, 12, 10, 11, 11, 9])
alts = {"Growth-heavy": ([0.15, 0.45, 0.15, 0.10, 0.15], [0.40, 0.25, 0.20, 0.15]),
        "Safety-first": ([0.25, 0.10, 0.15, 0.35, 0.15], [0.30, 0.30, 0.25, 0.15])}
for k, (nm, (a, b)) in enumerate(alts.items()):
    put(se, f"J{4 + k * 2}", nm, BOLD)
    for j, v in enumerate(a + b):
        put(se, f"{L(11 + j)}{4 + k * 2}", v, BLUE, '0%', YEL)
for i in range(len(plays)):
    rr = 4 + i
    pr = 5 + i
    put(se, f"A{rr}", f"=Plays!A{pr}", GREEN)
    put(se, f"B{rr}", f"=Plays!L{pr}*Plays!M{pr}", BLACK, '0.00')
    for col, wrow in (("C", 4), ("D", 6)):
        att = f"SUMPRODUCT(Plays!C{pr}:G{pr},$K${wrow}:$O${wrow})/SUM($K${wrow}:$O${wrow})"
        rtw = f"SUMPRODUCT(Plays!H{pr}:K{pr},$P${wrow}:$S${wrow})/SUM($P${wrow}:$S${wrow})"
        put(se, f"{col}{rr}", f"={att}*{rtw}", BLACK, '0.00')
    le2 = 4 + len(plays) - 1
    for col, src in (("E", "B"), ("F", "C"), ("G", "D")):
        put(se, f"{col}{rr}", f"=RANK({src}{rr},${src}$4:${src}${le2},0)", BLACK, "0")
    put(se, f"H{rr}", f"=MAX(E{rr}:G{rr})-MIN(E{rr}:G{rr})", BLACK, "0")
se.conditional_formatting.add(f"E4:G{4+len(plays)-1}", ColorScaleRule(start_type="num", start_value=1, start_color="63BE7B", end_type="num", end_value=11, end_color="F8696B"))

# ---------------------------------------------------------------- Sources
so = wb.create_sheet("Sources")
header(so, 1, ["Source", "Used for", "Link"], [50, 44, 100])
srcs = [
    ("WSTS: Global semiconductor market grows 26% in 2025 to $796B", "Chip market size and growth", "https://www.wsts.org/76/103/Global-Semiconductor-Market-grows-26-in-2025-to-796B"),
    ("SEMI: Global semiconductor equipment billings reached $135 billion in 2025", "Equipment billings by region", "https://www.semi.org/en/SEMI-Reports-Global-Semiconductor-Equipment-Billings-Reached-135-Billion-in-2025"),
    ("SEMI via AAP: Semiconductor materials market reached record $73.2 billion in 2025", "Wafer fab and packaging materials", "https://www.aap.com.au/aapreleases/cision20260512ae56606/"),
    ("Quartz (citing Construction Physics): semiconductor fab cost breakdown", "Equipment 70-80% vs building 20-30% of fab cost", "https://qz.com/semiconductor-fab-cost-breakdown-hidden-expenses-070626"),
    ("Congressional Research Service R49031: CHIPS-funded fabs", "$30.9bn direct funding, 40 fab projects, delays", "https://www.everycrsreport.com/reports/R49031.html"),
    ("IANS: India approves 12 chip projects with Rs 1.64 lakh crore pipeline", "India project count and total", "https://ianslive.in/india-approves-12-chip-manufacturing-projects-with-rs-164-lakh-crore-investment--20260628090925"),
    ("Adda247: Semiconductor plants in India 2026, complete list", "Project-level investment, type, state, status", "https://currentaffairs.adda247.com/semiconductor-plants-in-india-2026-complete-list-locations-and-manufacturing-details/"),
    ("India Semiconductor Mission", "Official programme page (verify project list)", "https://ism.gov.in"),
    ("Honeywell Q4 2025 results and spin-off releases", "Current portfolio (Building Automation, Industrial Automation; Solstice spun off)", "https://www.honeywell.com/us/en/press/2026/01/honeywell-reports-fourth-quarter-2025-results-with-adjusted-sales-and-adjusted-earnings-above-high-end-of-guidance-issues-2026-outlook"),
]
for i, s in enumerate(srcs):
    for j, v in enumerate(s):
        c = so.cell(row=2 + i, column=j + 1, value=v)
        c.font, c.alignment, c.border = BLACK, WRAP, BOX

for sh in wb.worksheets:
    sh.sheet_view.showGridLines = False
wb.save("semiconductor_where_to_play.xlsx")
print("saved semiconductor_where_to_play.xlsx")
