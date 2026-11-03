"""House Flipping Calculator – profit, ROI, 70% rule (MAO), rehab budget vs actual, comps-based ARV and a what-if table.

Deal (inputs + calculated results) <- Rehab Budget (line items by category, estimate vs actual) <- Comps (suggested ARV)
-> Dashboard (KPIs, profit breakdown, rule checks, rehab by category, sensitivity, charts).
"""
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter as L
from _common import Kit, F, B, MONEY, DATE, font, fill, style_range, input_cell, finish
import datetime as dt

MAIN = "C2410C"; SOFT = "FDEEE4"
K = Kit(MAIN, SOFT)
PCT = "0.0%"; PCT2 = "0.00%"
CATS = ["Demolition & cleanup", "Exterior & roof", "Structural & foundation", "Plumbing", "Electrical", "HVAC", "Kitchen",
        "Bathrooms", "Flooring", "Paint & drywall", "Windows & doors", "Landscaping", "Permits & fees", "Other"]
NR = 60; RB0 = 6; RB1 = RB0 + NR - 1
NCMP = 8; CP0 = 7; CP1 = CP0 + NCMP - 1
RB = "'Rehab Budget'"

# key, label, kind, fmt, help, sample, blank default
INP = [
    ("sec", "PROPERTY", None, None, None),
    ("addr", "Property address / name", None, "", "418 Willow Lane", None),
    ("sqft", "Living area (sq ft)", "#,##0", "Used with your comps to suggest an ARV.", 1450, None),
    ("sec", "PURCHASE & SALE PRICE", None, None, None),
    ("price", "Purchase price", MONEY, "What you pay for the property.", 146000, None),
    ("bclose", "Buying closing costs (% of price)", PCT, "Title, escrow, inspection, fees. Often 2–3%.", 0.025, 0.025),
    ("arv", "After repair value (ARV)", MONEY, "What it will sell for after the work. Leave empty to use the ARV suggested by your Comps tab.", 289000, None),
    ("sec", "FINANCING (leave loan % at 0% for a cash purchase)", None, None, None),
    ("ltv", "Loan: % of purchase price", PCT, "Hard-money lenders often lend 80–90% of the price.", 0.85, 0.0),
    ("rfin", "Loan: % of rehab financed", PCT, "Many flip lenders also fund 50–100% of the rehab.", 1.0, 0.0),
    ("rate", "Interest rate (yearly)", PCT2, "Hard money is often 10–13%. Interest-only payments assumed.", 0.11, 0.11),
    ("pts", "Points (% of loan)", PCT2, "Up-front lender fee. 1–3 points is common.", 0.02, 0.02),
    ("lfee", "Other lender fees", MONEY, "Underwriting, draw and doc fees.", 1250, None),
    ("sec", "HOLDING COSTS", None, None, None),
    ("months", "Months you hold the property", "0", "From purchase to sale: rehab time + time on market.", 5, 6),
    ("tax", "Property tax per year", MONEY, "", 3100, None),
    ("ins", "Insurance per year (builder's risk / vacant)", MONEY, "", 1800, None),
    ("util", "Utilities per month", MONEY, "Power, water, gas while you work on it.", 220, None),
    ("hoa", "HOA per month", MONEY, "", 0, None),
    ("ohold", "Other holding costs per month", MONEY, "Lawn care, security, dumpster rental…", 90, None),
    ("sec", "SELLING COSTS", None, None, None),
    ("comm", "Agent commission (% of sale price)", PCT, "Often 5–6% total.", 0.05, 0.05),
    ("sclose", "Seller closing costs (% of sale price)", PCT, "Title, transfer tax, concessions. Often 1–3%.", 0.015, 0.015),
    ("stage", "Staging, photos & marketing", MONEY, "", 2500, None),
    ("sec", "YOUR RULES", None, None, None),
    ("rule", "Max offer rule (% of ARV)", PCT, "The '70% rule': max offer = ARV × 70% – rehab. Use 65–75% for your market.", 0.70, 0.70),
    ("cont", "Rehab contingency", PCT, "Extra buffer for surprises. 10–15% is wise.", 0.10, 0.10),
    ("target", "Minimum profit you want", MONEY, "Your walk-away line.", 30000, 25000),
]
OUT = [
    ("sec", "RESULTS (CALCULATED)", None, None),
    ("arvs", "ARV suggested by your comps", MONEY, "Average price per sq ft of the comps marked 'Yes' × living area."),
    ("arvu", "ARV used", MONEY, "Your ARV above, or the comps ARV if you left it empty."),
    ("rest", "Rehab estimate (Rehab Budget tab)", MONEY, ""),
    ("rcon", "Contingency", MONEY, ""),
    ("rtot", "Rehab budget incl. contingency", MONEY, ""),
    ("ract", "Rehab actually spent so far", MONEY, ""),
    ("bcc", "Buying closing costs", MONEY, ""),
    ("loan", "Loan amount", MONEY, ""),
    ("ptsamt", "Points + lender fees", MONEY, ""),
    ("int", "Interest while holding", MONEY, "Interest-only, for the months you hold."),
    ("hold", "Holding costs", MONEY, "Tax, insurance, utilities, HOA and other costs for the months you hold."),
    ("sell", "Selling costs", MONEY, ""),
    ("total", "Total project cost", MONEY, "Everything: purchase + closing + rehab + financing + holding + selling."),
    ("profit", "Net profit", MONEY, "ARV – total project cost (before income tax)."),
    ("cash", "Cash you need (out of pocket)", MONEY, "All costs before the sale minus the loan."),
    ("roi", "ROI on total cost", PCT, "Profit ÷ total project cost."),
    ("coc", "Return on your cash", PCT, "Profit ÷ cash you need."),
    ("ann", "Annualized return on cash", PCT, "Return on cash scaled to a 12-month year."),
    ("margin", "Profit margin (% of ARV)", PCT, "Many flippers aim for 10–15%+."),
    ("mao", "Max offer (rule % of ARV – rehab)", MONEY, ""),
    ("vsmao", "Room under max offer", MONEY, "Max offer – purchase price. Negative = you pay more than your rule allows."),
    ("maot", "Max offer for your minimum profit (approx.)", MONEY, "Highest price that still leaves your minimum profit."),
    ("be", "Break-even sale price", MONEY, "Lowest sale price where you don't lose money."),
    ("fix", "Helper: fixed costs (not ARV- or month-based)", MONEY, "Used by the what-if table."),
    ("perm", "Helper: cost per month held", MONEY, "Used by the what-if table."),
    ("verdict", "Verdict", None, ""),
]
R = {}
r = 5
for row in INP + OUT:
    if row[0] != "sec": R[row[0]] = r
    r += 1


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); dl = wb.create_sheet("Deal"); rb = wb.create_sheet("Rehab Budget"); cp = wb.create_sheet("Comps")
    g = lambda k_: f"Deal!$B${R[k_]}"

    # ---------- Deal ----------
    K.title(dl, "Deal", "Fill in the yellow cells. Results calculate at the bottom and on the Dashboard.")
    dl["A3"] = "Currency label"; dl["A3"].font = font(bold=True); input_cell(dl["B3"], "USD")
    for c, w in zip("ABC", [44, 20, 90]): dl.column_dimensions[c].width = w
    r = 5
    for row in INP:
        if row[0] == "sec":
            dl[f"A{r}"] = row[1]; dl[f"A{r}"].font = Font(name=F, bold=True, color=MAIN)
            for c in "ABC": dl[f"{c}{r}"].fill = fill(SOFT)
        else:
            k_, lab, fmt, hlp, sv, bv = row
            dl[f"A{r}"] = lab; dl[f"A{r}"].font = font(); dl[f"A{r}"].border = B
            input_cell(dl[f"B{r}"], sv if sample else bv)
            if fmt: dl[f"B{r}"].number_format = fmt
            dl[f"C{r}"] = hlp or None; dl[f"C{r}"].font = font(italic=True, size=9)
        r += 1
    rest = f"SUM({RB}!$C${RB0}:$C${RB1})"
    calc = {
        "arvs": f"=IFERROR(AVERAGEIFS(Comps!$I${CP0}:$I${CP1},Comps!$H${CP0}:$H${CP1},\"Yes\",Comps!$I${CP0}:$I${CP1},\">0\")*{g('sqft')},0)",
        "arvu": f"=IF(N({g('arv')})>0,{g('arv')},{g('arvs')})",
        "rest": "=" + rest,
        "rcon": f"={g('rest')}*{g('cont')}",
        "rtot": f"={g('rest')}+{g('rcon')}",
        "ract": f"=SUM({RB}!$D${RB0}:$D${RB1})",
        "bcc": f"={g('price')}*{g('bclose')}",
        "loan": f"={g('price')}*{g('ltv')}+{g('rtot')}*{g('rfin')}",
        "ptsamt": f"=IF({g('loan')}>0,{g('loan')}*{g('pts')}+{g('lfee')},0)",
        "int": f"={g('loan')}*{g('rate')}/12*{g('months')}",
        "hold": f"=({g('tax')}/12+{g('ins')}/12+{g('util')}+{g('hoa')}+{g('ohold')})*{g('months')}",
        "sell": f"=IF({g('arvu')}>0,{g('arvu')}*({g('comm')}+{g('sclose')})+{g('stage')},0)",
        "total": f"={g('price')}+{g('bcc')}+{g('rtot')}+{g('ptsamt')}+{g('int')}+{g('hold')}+{g('sell')}",
        "profit": f"=IF({g('price')}=0,0,{g('arvu')}-{g('total')})",
        "cash": f"=MAX(0,{g('price')}+{g('bcc')}+{g('rtot')}+{g('ptsamt')}+{g('int')}+{g('hold')}-{g('loan')})",
        "roi": f"=IFERROR({g('profit')}/{g('total')},0)",
        "coc": f"=IFERROR({g('profit')}/{g('cash')},0)",
        "ann": f"=IF(OR({g('cash')}<=0,{g('months')}<=0,{g('coc')}<=-1),0,IFERROR((1+{g('coc')})^(12/{g('months')})-1,0))",
        "margin": f"=IFERROR({g('profit')}/{g('arvu')},0)",
        "mao": f"=IF({g('arvu')}>0,{g('arvu')}*{g('rule')}-{g('rtot')},0)",
        "vsmao": f"=IF({g('price')}=0,0,{g('mao')}-{g('price')})",
        "maot": (f"=IF({g('arvu')}=0,0,MAX(0,({g('arvu')}-{g('sell')}-{g('rtot')}-{g('hold')}-{g('target')}-{g('rtot')}*{g('rfin')}*({g('pts')}+{g('rate')}/12*{g('months')})"
                 f"-IF({g('ltv')}+{g('rfin')}>0,{g('lfee')},0))/(1+{g('bclose')}+{g('ltv')}*({g('pts')}+{g('rate')}/12*{g('months')}))))"),
        "be": f"=IFERROR(({g('total')}-{g('sell')}+{g('stage')})/(1-{g('comm')}-{g('sclose')}),0)",
        "fix": f"={g('price')}+{g('bcc')}+{g('rtot')}+{g('ptsamt')}+{g('stage')}",
        "perm": f"={g('loan')}*{g('rate')}/12+{g('tax')}/12+{g('ins')}/12+{g('util')}+{g('hoa')}+{g('ohold')}",
        "verdict": (f'=IF({g("price")}=0,"Fill in the purchase price",IF({g("profit")}<0,"Loses money – walk away or renegotiate",'
                    f'IF(AND({g("profit")}>={g("target")},{g("price")}<={g("mao")}),"Meets your rules – good deal",'
                    f'IF({g("profit")}>={g("target")},"Profit OK but above your max offer – negotiate","Profit below your minimum – negotiate"))))'),
    }
    for row in OUT:
        if row[0] == "sec":
            dl[f"A{r}"] = row[1]; dl[f"A{r}"].font = Font(name=F, bold=True, color=MAIN)
            for c in "ABC": dl[f"{c}{r}"].fill = fill(SOFT)
        else:
            k_, lab, fmt, hlp = row
            dl[f"A{r}"] = lab; dl[f"A{r}"].font = font(bold=k_ in ("profit", "cash", "total", "verdict"), size=9 if k_ in ("fix", "perm") else 10)
            dl[f"A{r}"].border = B
            c = dl[f"B{r}"]; c.value = calc[k_]; c.fill = fill("F3F4F6"); c.border = B; c.font = font(bold=k_ in ("profit", "verdict"))
            if fmt: c.number_format = fmt
            dl[f"C{r}"] = hlp or None; dl[f"C{r}"].font = font(italic=True, size=9)
        r += 1
    dl.conditional_formatting.add(f"B{R['profit']}", FormulaRule(formula=[f"B{R['profit']}<0"], font=Font(name=F, bold=True, color="B91C1C")))
    dl.conditional_formatting.add(f"B{R['vsmao']}", FormulaRule(formula=[f"B{R['vsmao']}<0"], font=Font(name=F, color="B91C1C")))
    dvp = DataValidation(type="decimal", operator="between", formula1="0", formula2="1", allow_blank=True,
                         error="Type a percentage, e.g. 10%", showErrorMessage=True)
    dl.add_data_validation(dvp)
    for row in INP:
        if row[0] != "sec" and row[2] in (PCT, PCT2): dvp.add(f"B{R[row[0]]}")

    # ---------- Rehab Budget ----------
    K.title(rb, "Rehab Budget", "One line per job. Type your estimate first, then the actual cost as bills come in.")
    rb["A3"] = "Common jobs are listed already – delete or change any line, and add your own in the empty rows."; rb["A3"].font = font(italic=True, size=9)
    K.hdr(rb, 5, ["Category", "Job / item", "Estimated cost", "Actual cost", "Over (+) / under (–)", "Status", "Contractor / store", "Notes"],
          [24, 36, 16, 16, 18, 14, 22, 30])
    style_range(rb, RB0, RB1, 1, 8, inputs=(1, 2, 3, 4, 6, 7, 8), fmt={3: MONEY, 4: MONEY, 5: '+#,##0.00;-#,##0.00;"-"'})
    items = [("Demolition & cleanup", "Dumpster & debris removal", 1800, 1950, "Done"),
             ("Demolition & cleanup", "Interior demo", 1200, 1100, "Done"),
             ("Exterior & roof", "Roof repair / replacement", 6500, 7100, "Done"),
             ("Exterior & roof", "Gutters & exterior paint", 2400, 0, "In progress"),
             ("Structural & foundation", "Foundation / framing repairs", 0, 0, ""),
             ("Plumbing", "Re-pipe, fixtures, water heater", 3600, 3780, "Done"),
             ("Electrical", "Panel upgrade, outlets, lighting", 3400, 3350, "Done"),
             ("HVAC", "New HVAC system", 5200, 5200, "Done"),
             ("Kitchen", "Cabinets & countertops", 6800, 0, "In progress"),
             ("Kitchen", "Appliances", 2800, 0, "Not started"),
             ("Bathrooms", "Full bath remodel", 5200, 0, "In progress"),
             ("Bathrooms", "Half bath refresh", 1500, 0, "Not started"),
             ("Flooring", "LVP flooring throughout", 5200, 0, "Not started"),
             ("Paint & drywall", "Drywall patch & interior paint", 3900, 0, "Not started"),
             ("Windows & doors", "Windows, front door, interior doors", 2800, 0, "Not started"),
             ("Landscaping", "Curb appeal, sod, mulch", 2400, 0, "Not started"),
             ("Permits & fees", "Building permits", 1300, 1300, "Done"),
             ("Other", "Final cleaning", 450, 0, "Not started")]
    for i, (cat, item, est, act, stt) in enumerate(items):
        rr = RB0 + i; rb[f"A{rr}"] = cat; rb[f"B{rr}"] = item
        if sample:
            if est: rb[f"C{rr}"] = est
            if act: rb[f"D{rr}"] = act
            if stt: rb[f"F{rr}"] = stt
            if item == "HVAC system" or cat == "HVAC": rb[f"G{rr}"] = "CoolAir Heating"
            if cat == "Exterior & roof" and est == 6500: rb[f"H{rr}"] = "Extra decking needed"
    for rr in range(RB0, RB1 + 1):
        rb[f"E{rr}"] = f'=IF(OR(C{rr}="",D{rr}=""),"",D{rr}-C{rr})'
    dvc = DataValidation(type="list", formula1=f"=Lists!$A$2:$A${1 + len(CATS)}", allow_blank=True); rb.add_data_validation(dvc); dvc.add(f"A{RB0}:A{RB1}")
    dvs = DataValidation(type="list", formula1='"Not started,In progress,Done"', allow_blank=True); rb.add_data_validation(dvs); dvs.add(f"F{RB0}:F{RB1}")
    TR = RB1 + 1
    rb[f"B{TR}"] = "Total"; rb[f"C{TR}"] = f"=SUM(C{RB0}:C{RB1})"; rb[f"D{TR}"] = f"=SUM(D{RB0}:D{RB1})"
    rb[f"E{TR}"] = f'=SUMIFS(D{RB0}:D{RB1},D{RB0}:D{RB1},"<>",C{RB0}:C{RB1},"<>")-SUMIFS(C{RB0}:C{RB1},D{RB0}:D{RB1},"<>",C{RB0}:C{RB1},"<>")'
    for c in "ABCDEFGH":
        cell = rb[f"{c}{TR}"]; cell.fill = fill(SOFT); cell.border = B; cell.font = font(bold=True)
    for c in "CD": rb[f"{c}{TR}"].number_format = MONEY
    rb[f"E{TR}"].number_format = '+#,##0.00;-#,##0.00;"-"'
    rb.conditional_formatting.add(f"E{RB0}:E{TR}", FormulaRule(formula=[f"AND(ISNUMBER(E{RB0}),E{RB0}>0)"], font=Font(name=F, bold=True, color="B91C1C")))
    rb.conditional_formatting.add(f"E{RB0}:E{TR}", FormulaRule(formula=[f"AND(ISNUMBER(E{RB0}),E{RB0}<0)"], font=Font(name=F, color="15803D")))
    rb.conditional_formatting.add(f"F{RB0}:F{RB1}", FormulaRule(formula=[f'F{RB0}="Done"'], fill=fill("DCFCE7")))
    rb.freeze_panes = f"A{RB0}"

    # ---------- Lists ----------
    ls = wb.create_sheet("Lists")
    ls["A1"] = "Rehab categories (rename if you like)"; ls["A1"].font = font(bold=True); ls.column_dimensions["A"].width = 34
    for i, cat in enumerate(CATS, 2):
        c = ls[f"A{i}"]; input_cell(c, cat)

    # ---------- Comps ----------
    K.title(cp, "Comps", "Recently SOLD homes near yours, similar in size and condition after your rehab. Mark the best ones 'Yes'.")
    cp["A3"] = "Your property's living area (sq ft)"; cp["A3"].font = font(bold=True); cp["D3"] = f"={g('sqft')}"; cp["D3"].number_format = "#,##0"
    cp["A4"] = "Suggested ARV (average $/sq ft of 'Yes' comps × your sq ft)"; cp["A4"].font = font(bold=True)
    cp["D4"] = f"={g('arvs')}"; cp["D4"].number_format = MONEY; cp["D4"].font = Font(name=F, bold=True, color=MAIN)
    for c in ("D3", "D4"): cp[c].fill = fill("F3F4F6"); cp[c].border = B
    K.hdr(cp, 6, ["Address", "Sale date", "Sale price", "Sq ft", "Beds", "Baths", "Distance (miles)", "Use? (Yes/No)", "Price per sq ft"],
          [30, 13, 16, 10, 8, 8, 14, 13, 15])
    style_range(cp, CP0, CP1, 1, 9, inputs=(1, 2, 3, 4, 5, 6, 7, 8), fmt={2: DATE, 3: MONEY, 4: "#,##0", 7: "0.0", 9: MONEY})
    comps = [("402 Willow Lane", dt.date(2026, 8, 14), 295000, 1480, 3, 2, 0.1, "Yes"),
             ("1130 Birch Ct", dt.date(2026, 7, 2), 279500, 1390, 3, 2, 0.4, "Yes"),
             ("77 Aspen Dr", dt.date(2026, 9, 21), 301000, 1525, 3, 2, 0.6, "Yes"),
             ("2290 Cedar Ave", dt.date(2026, 6, 9), 262000, 1410, 3, 1.5, 0.8, "No"),
             ("15 Larch St", dt.date(2026, 9, 3), 286000, 1440, 3, 2, 0.5, "Yes")]
    for rr in range(CP0, CP1 + 1):
        cp[f"I{rr}"] = f'=IF(OR(C{rr}="",N(D{rr})=0),"",C{rr}/D{rr})'
    if sample:
        for i, row in enumerate(comps):
            for j, v in enumerate(row): cp.cell(row=CP0 + i, column=1 + j, value=v)
    dvy = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True); cp.add_data_validation(dvy); dvy.add(f"H{CP0}:H{CP1}")
    cp[f"A{CP1 + 2}"] = ("Tip: use sales from the last 3–6 months, within about 1 mile, with a similar size, bedrooms and finish. "
                         "Mark 'No' for comps that are not a good match (needs work, much bigger, different area).")
    cp[f"A{CP1 + 2}"].font = font(italic=True, size=9)
    cp[f"A{CP1 + 3}"] = "Lowest / highest price per sq ft of 'Yes' comps"; cp[f"A{CP1 + 3}"].font = font(size=9)
    cp[f"D{CP1 + 3}"] = f'=IF(COUNTIFS(H{CP0}:H{CP1},"Yes",I{CP0}:I{CP1},">0")=0,0,_xlfn.MINIFS(I{CP0}:I{CP1},H{CP0}:H{CP1},"Yes",I{CP0}:I{CP1},">0"))'
    cp[f"E{CP1 + 3}"] = f'=IF(COUNTIFS(H{CP0}:H{CP1},"Yes",I{CP0}:I{CP1},">0")=0,0,_xlfn.MAXIFS(I{CP0}:I{CP1},H{CP0}:H{CP1},"Yes",I{CP0}:I{CP1},">0"))'
    for c in "DE": cp[f"{c}{CP1 + 3}"].number_format = MONEY; cp[f"{c}{CP1 + 3}"].font = font(size=9)

    # ---------- Dashboard ----------
    K.title(dash, "House Flip Dashboard", "Updates by itself from the Deal, Rehab Budget and Comps tabs.")
    dash["A3"] = f'=IF({g("addr")}="","",{g("addr")}&"  ·  ")&{g("verdict")}'
    dash["A3"].font = Font(name=F, bold=True, size=12, color=MAIN)
    dash.conditional_formatting.add("A3", FormulaRule(formula=['ISNUMBER(SEARCH("Loses",A3))'], font=Font(name=F, bold=True, size=12, color="B91C1C")))
    dash.conditional_formatting.add("A3", FormulaRule(formula=['ISNUMBER(SEARCH("good deal",A3))'], font=Font(name=F, bold=True, size=12, color="15803D")))
    kp = [("Net profit", "=" + g("profit"), MONEY), ("Return on your cash", "=" + g("coc"), PCT), ("Annualized return", "=" + g("ann"), PCT),
          ("Cash you need", "=" + g("cash"), MONEY), ("Max offer (rule)", "=" + g("mao"), MONEY), ("Profit margin of ARV", "=" + g("margin"), PCT)]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 3) * 3, 1 + (i % 3) * 2, lab, f, fmt)
        dash.cell(row=6 + (i // 3) * 3, column=1 + (i % 3) * 2).alignment = Alignment(horizontal="left")
    dash.conditional_formatting.add("A6", FormulaRule(formula=["A6<0"], font=Font(name=F, size=18, bold=True, color="B91C1C")))
    for c, w in zip("ABCDEFGH", [30, 16, 22, 30, 16, 16, 14, 14]): dash.column_dimensions[c].width = w

    K.section(dash, "A11", "Where the money goes")
    K.hdr(dash, 12, ["Item", "Amount"])
    wf = [("Sale price (ARV)", g("arvu")), ("Purchase price", g("price")), ("Buying closing costs", g("bcc")), ("Rehab + contingency", g("rtot")),
          ("Points, fees & interest", f"{g('ptsamt')}+{g('int')}"), ("Holding costs", g("hold")), ("Selling costs", g("sell")), ("Net profit", g("profit"))]
    for i, (lab, f) in enumerate(wf, 13):
        dash[f"A{i}"] = lab; dash[f"B{i}"] = "=" + f
    style_range(dash, 13, 20, 1, 2, fmt={2: MONEY})
    for c in "AB":
        for rr in (13, 20): dash[f"{c}{rr}"].font = font(bold=True); dash[f"{c}{rr}"].fill = fill(SOFT)
    dash.conditional_formatting.add("B20", FormulaRule(formula=["B20<0"], font=Font(name=F, bold=True, color="B91C1C")))

    K.section(dash, "D11", "Checks")
    checks = [("Profit ≥ your minimum", g("profit"), MONEY, f'IF({g("profit")}>={g("target")},"✔ Yes","✖ No")'),
              ("Price ≤ max offer (rule)", g("vsmao"), MONEY, f'IF({g("vsmao")}>=0,"✔ Yes","✖ No")'),
              ("Margin ≥ 10% of ARV", g("margin"), PCT, f'IF({g("margin")}>=0.1,"✔ Yes","✖ No")'),
              ("Rehab spent vs budget", f"IFERROR({g('ract')}/{g('rtot')},0)", PCT, f'IF({g("ract")}<={g("rtot")},"✔ Within","✖ Over")'),
              ("Max offer for your min. profit", g("maot"), MONEY, f'IF({g("price")}<={g("maot")},"✔ Price OK","✖ Too high")'),
              ("Break-even sale price", g("be"), MONEY, f'IF({g("be")}<={g("arvu")}*0.9,"✔ Safe margin","✖ Tight")')]
    K.hdr(dash, 12, ["Check", "Value", "Result"], col0=4)
    for i, (lab, f, fmt, res) in enumerate(checks, 13):
        dash[f"D{i}"] = lab; dash[f"E{i}"] = "=" + f; dash[f"F{i}"] = f'=IF({g("price")}=0,"",{res})'; dash[f"E{i}"].number_format = fmt
        for c in "DEF": dash[f"{c}{i}"].border = B; dash[f"{c}{i}"].fill = fill("F3F4F6"); dash[f"{c}{i}"].font = font()
    dash.conditional_formatting.add("F13:F18", FormulaRule(formula=['LEFT(F13,1)="✔"'], font=Font(name=F, bold=True, color="15803D")))
    dash.conditional_formatting.add("F13:F18", FormulaRule(formula=['LEFT(F13,1)="✖"'], font=Font(name=F, bold=True, color="B91C1C")))

    K.section(dash, "A23", "What if? Net profit if the sale price or the timeline changes")
    dash["A24"] = "Sale price ↓  /  extra months →"; dash["A24"].font = Font(name=F, bold=True, color="FFFFFF"); dash["A24"].fill = fill(MAIN)
    extra = [0, 1, 2, 3, 6]; chg = [-0.10, -0.05, 0, 0.05, 0.10]
    for j, e in enumerate(extra):
        c = dash.cell(row=24, column=2 + j, value=e); c.number_format = '"+"0" mo";"-"0" mo";"as planned"'
        c.font = Font(name=F, bold=True, color="FFFFFF"); c.fill = fill(MAIN); c.alignment = Alignment(horizontal="center"); c.border = B
    for i, d in enumerate(chg):
        rr = 25 + i; a = dash.cell(row=rr, column=1, value=d); a.number_format = '"ARV "+0%;"ARV "-0%;"ARV as planned"'
        a.font = font(bold=True); a.fill = fill(SOFT); a.border = B
        for j in range(len(extra)):
            cl = L(2 + j)
            dash[f"{cl}{rr}"] = (f"=IF({g('price')}=0,0,{g('arvu')}*(1+$A{rr})*(1-{g('comm')}-{g('sclose')})-{g('fix')}"
                                 f"-{g('perm')}*({g('months')}+{cl}$24))")
            dash[f"{cl}{rr}"].number_format = MONEY; dash[f"{cl}{rr}"].border = B; dash[f"{cl}{rr}"].font = font(); dash[f"{cl}{rr}"].fill = fill("F3F4F6")
    dash.conditional_formatting.add("B25:F29", FormulaRule(formula=["B25<0"], font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    dash.conditional_formatting.add("B25:F29", FormulaRule(formula=[f"B25>={g('target')}"], font=Font(name=F, color="15803D")))
    dash["A30"] = "Green = meets your minimum profit, red = loss. A slow sale costs interest and holding costs every extra month."
    dash["A30"].font = font(italic=True, size=9)

    K.section(dash, "A32", "Rehab by category")
    K.hdr(dash, 33, ["Category", "Estimated", "Actual so far", "Over (+) / under (–)", "Share of estimate"])
    for i in range(len(CATS)):
        rr = 34 + i
        dash[f"A{rr}"] = f"=Lists!A{2 + i}"
        dash[f"B{rr}"] = f"=SUMIFS({RB}!$C${RB0}:$C${RB1},{RB}!$A${RB0}:$A${RB1},A{rr})"
        dash[f"C{rr}"] = f"=SUMIFS({RB}!$D${RB0}:$D${RB1},{RB}!$A${RB0}:$A${RB1},A{rr})"
        dash[f"D{rr}"] = (f'=SUMIFS({RB}!$D${RB0}:$D${RB1},{RB}!$A${RB0}:$A${RB1},A{rr},{RB}!$C${RB0}:$C${RB1},"<>",{RB}!$D${RB0}:$D${RB1},"<>")'
                          f'-SUMIFS({RB}!$C${RB0}:$C${RB1},{RB}!$A${RB0}:$A${RB1},A{rr},{RB}!$C${RB0}:$C${RB1},"<>",{RB}!$D${RB0}:$D${RB1},"<>")')
        dash[f"E{rr}"] = f"=IFERROR(B{rr}/{g('rest')},0)"
    E1 = 33 + len(CATS)
    style_range(dash, 34, E1, 1, 5, fmt={2: MONEY, 3: MONEY, 4: '+#,##0.00;-#,##0.00;"-"', 5: PCT})
    dash.conditional_formatting.add(f"D34:D{E1}", FormulaRule(formula=["D34>0"], font=Font(name=F, bold=True, color="B91C1C")))
    dash[f"A{E1 + 1}"] = "Rehab done"; dash[f"A{E1 + 1}"].font = font(bold=True)
    dash[f"B{E1 + 1}"] = (f'=IFERROR(COUNTIFS({RB}!$F${RB0}:$F${RB1},"Done")/COUNTIFS({RB}!$C${RB0}:$C${RB1},">0"),0)')
    dash[f"B{E1 + 1}"].number_format = "0%"; dash[f"B{E1 + 1}"].font = Font(name=F, bold=True, color=MAIN)
    dash[f"C{E1 + 1}"] = "of jobs with an estimate are marked Done"; dash[f"C{E1 + 1}"].font = font(italic=True, size=9)

    ch = BarChart(); ch.type = "bar"; ch.title = "Where the money goes"; ch.height = 7.5; ch.width = 16; ch.legend = None
    ch.add_data(Reference(dash, min_col=2, min_row=12, max_row=20), titles_from_data=True)
    ch.set_categories(Reference(dash, min_col=1, min_row=13, max_row=20)); ch.y_axis.majorGridlines = None
    ch.series[0].graphicalProperties.solidFill = MAIN; ch.x_axis.scaling.orientation = "maxMin"
    dash.add_chart(ch, "H3")
    ch2 = BarChart(); ch2.type = "col"; ch2.grouping = "clustered"; ch2.title = "Rehab: estimate vs actual"; ch2.height = 8; ch2.width = 16
    ch2.add_data(Reference(dash, min_col=2, max_col=3, min_row=33, max_row=E1), titles_from_data=True)
    ch2.set_categories(Reference(dash, min_col=1, min_row=34, max_row=E1)); ch2.y_axis.majorGridlines = None
    for ser, col in zip(ch2.series, ("FDBA74", "C2410C")): ser.graphicalProperties.solidFill = col
    dash.add_chart(ch2, "H19")

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Comps tab: add 3–6 recently sold homes near yours that look like your house will after the work. Mark the good matches 'Yes' – you get a suggested ARV (after repair value).", False),
             ("2.  Rehab Budget tab: type an estimate for each job (common jobs are already listed). Later, add the actual cost as bills come in and mark jobs Done.", False),
             ("3.  Deal tab: fill in the purchase price, your ARV (or leave it empty to use the comps ARV), loan, holding months and selling costs.", False),
             ("4.  Dashboard: net profit, return on your cash, max offer (70% rule), checks, a what-if table and your rehab budget vs actual.", False),
             ("", False), ("Tips", True),
             ("•  Before you make an offer: look at 'Max offer for your minimum profit' on the Deal tab – that is your walk-away price.", False),
             ("•  Keep a contingency of 10–15%. Old houses almost always hide a surprise behind the walls.", False),
             ("•  Check the what-if table: if a 10% lower sale price or 3 extra months turns the deal red, the deal is risky.", False),
             ("•  Change the 70% rule to fit your market (for example 75% for cheaper markets with fast sales).", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  Interest-only loan payments are assumed (normal for hard-money and flip loans). Income tax on your profit is not included.", False),
             ("•  60 rehab lines, 8 comps, 14 rehab categories you can rename on the Lists tab. Any currency. No macros.", False),
             ("•  To use Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("The house, comps and rehab costs are made-up examples. Use the blank copy in your download for your own flip.", False)]
    K.start(st, "House Flipping Calculator", "Profit, ROI, 70% rule, rehab budget and comps – before you buy and while you build. Excel and Google Sheets.", lines)
    finish(wb, sample, path)


META = {
    "name": "House Flipping Calculator (Profit & ROI, 70% Rule, Rehab Budget, Comps ARV) – Excel & Google Sheets",
    "file_base": "House-Flipping-Calculator",
    "price": 24,
    "color": "#C2410C",
    "title": "House Flipping Calculator",
    "tagline": "Profit & ROI · 70% rule max offer · rehab budget vs actual · comps ARV",
    "intro": "Run the numbers on a fix-and-flip before you make an offer – and keep the rehab on budget while you work. Enter the price, rehab jobs, loan and holding time and see your net profit, return on cash, max offer (70% rule) and what happens if the house sells for less or takes longer.",
    "features": [
        "Net profit, ROI, return on your cash and annualized return – all costs included (closing, rehab, points, interest, holding, selling)",
        "Max allowable offer (70% rule, adjustable) and the max price that still leaves your minimum profit",
        "Rehab budget: 60 lines in 14 categories, estimate vs actual, status and contractor, plus a contingency buffer",
        "Comps tab: suggested ARV from recently sold homes (average price per sq ft)",
        "Hard-money financing: loan % of price and rehab, points, fees and interest-only payments",
        "What-if table: profit if the sale price drops 5–10% or the sale takes extra months",
        "Dashboard with a clear verdict, checks and charts; blank copy plus a sample copy",
    ],
    "tags": ["house flipping", "fix and flip", "flip calculator", "rehab budget", "real estate investing", "70 percent rule",
             "arv calculator", "real estate spreadsheet", "google sheets", "excel template", "property investor", "renovation budget", "flipping houses"],
}
