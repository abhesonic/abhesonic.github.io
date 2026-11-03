"""Rental Property Investment Analyzer – compare up to 5 deals side by side, with a 10-year projection and IRR for the selected deal.

Deals (inputs per deal column + calculated results) -> Dashboard (pick a deal: KPIs, rule-of-thumb checks, comparison, charts)
-> Projection (10 years for the selected deal: cash flow, loan balance, equity, total return, IRR).
"""
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.utils import get_column_letter as L
from _common import Kit, F, B, MONEY, font, fill, style_range, input_cell, finish

MAIN = "1D4ED8"; SOFT = "E8EEFC"
K = Kit(MAIN, SOFT)
ND = 5; D0 = 2; D1 = D0 + ND - 1      # deal columns B..F
PCT = "0.0%"; PCT2 = "0.00%"; X = '0.00"x"'; YRS = 10
DEALS = "Deals"

# key, label, kind (in = input, out = formula, sec = section header), format, help
ROWS = [
    ("sec", "THE PROPERTY", "sec", None, ""),
    ("name", "Property name", "in", None, "A short name you will recognise, e.g. '12 Oak St duplex'."),
    ("units", "Number of units", "in", "0", "1 for a single-family home, 2 for a duplex…"),
    ("sec", "PURCHASE", "sec", None, ""),
    ("price", "Purchase price", "in", MONEY, "Price you expect to pay."),
    ("close", "Closing costs", "in", MONEY, "Fees, title, inspection… typically 2–5% of the price."),
    ("repair", "Repairs before renting", "in", MONEY, "Work needed before the first tenant moves in."),
    ("sec", "FINANCING", "sec", None, ""),
    ("down", "Down payment %", "in", PCT, "Share of the price you pay in cash. 100% = no loan (cash purchase)."),
    ("rate", "Interest rate (yearly)", "in", PCT2, "Mortgage interest rate."),
    ("term", "Loan term (years)", "in", "0", "Usually 30 (or 15)."),
    ("sec", "INCOME (MONTHLY)", "sec", None, ""),
    ("rent", "Total rent per month (all units)", "in", MONEY, "Market rent – check similar listings nearby."),
    ("oinc", "Other income per month", "in", MONEY, "Parking, laundry, storage, pet fees…"),
    ("vac", "Vacancy %", "in", PCT, "Share of the year the unit is empty. 5–8% is common."),
    ("sec", "EXPENSES", "sec", None, ""),
    ("tax", "Property tax per year", "in", MONEY, ""),
    ("ins", "Insurance per year", "in", MONEY, ""),
    ("hoa", "HOA per month", "in", MONEY, ""),
    ("util", "Utilities you pay per month", "in", MONEY, "Water, trash, electricity you cover as the owner."),
    ("maint", "Repairs & maintenance % of rent", "in", PCT, "5–10% is a common estimate."),
    ("capex", "CapEx reserve % of rent", "in", PCT, "Saving for big items: roof, HVAC, water heater. 5–10%."),
    ("mgmt", "Property management % of rent", "in", PCT, "8–12% if you use a manager, 0% if you self-manage."),
    ("oexp", "Other expenses per month", "in", MONEY, "Lawn care, pest control, accounting…"),
    ("sec", "GROWTH (FOR THE 10-YEAR PROJECTION)", "sec", None, ""),
    ("rg", "Rent growth per year", "in", PCT, "2–3% is a cautious long-term guess."),
    ("eg", "Expense growth per year", "in", PCT, ""),
    ("ag", "Property value growth per year", "in", PCT, ""),
    ("sell", "Selling costs when you sell", "in", PCT, "Agent fees and closing costs on sale, about 6–8%."),
    ("sec", "RESULTS (CALCULATED)", "sec", None, ""),
    ("loan", "Loan amount", "out", MONEY, ""),
    ("pmt", "Mortgage payment per month (P&I)", "out", MONEY, ""),
    ("cash", "Total cash needed", "out", MONEY, "Down payment + closing costs + repairs."),
    ("gsi", "Gross income per year", "out", MONEY, ""),
    ("egi", "Effective income per year (after vacancy)", "out", MONEY, ""),
    ("opex", "Operating expenses per year", "out", MONEY, "All expenses except the mortgage."),
    ("noi", "Net operating income (NOI) per year", "out", MONEY, "Income after vacancy minus operating expenses."),
    ("ds", "Mortgage payments per year", "out", MONEY, ""),
    ("cf", "Cash flow per year", "out", MONEY, "NOI minus mortgage payments."),
    ("cfm", "Cash flow per month", "out", MONEY, ""),
    ("cfu", "Cash flow per unit per month", "out", MONEY, ""),
    ("cap", "Cap rate", "out", PCT2, "NOI ÷ purchase price. Compares properties without financing."),
    ("coc", "Cash-on-cash return", "out", PCT2, "Yearly cash flow ÷ total cash needed."),
    ("dscr", "Debt service coverage (DSCR)", "out", X, "NOI ÷ mortgage payments. Lenders like 1.25x or more."),
    ("grm", "Gross rent multiplier (GRM)", "out", "0.0", "Price ÷ yearly rent. Lower is better."),
    ("one", "Rent-to-price (1% rule)", "out", PCT2, "Monthly rent ÷ (price + repairs). 1% or more passes the 1% rule."),
    ("exr", "Operating expense ratio (50% rule)", "out", PCT, "Operating expenses ÷ income. Around 50% or less is healthy."),
    ("beo", "Break-even occupancy", "out", PCT, "How full the property must be to pay all costs and the mortgage."),
    ("irr", "10-year IRR (if sold after year 10)", "out", PCT2, "Yearly return including cash flow, loan paydown, value growth and sale."),
    ("verdict", "Quick verdict", "out", None, "Based on cash flow, cash-on-cash and DSCR. Always do your own checks."),
]
R = {}
r = 5
for k_, *_ in ROWS:
    if k_ != "sec": R[k_] = r
    r += 1
LAST = r - 1

SAMPLE = [  # name, units, price, close, repair, down, rate, term, rent, oinc, vac, tax, ins, hoa, util, maint, capex, mgmt, oexp, rg, eg, ag, sell
    ("Maple St duplex", 2, 259000, 7770, 6000, 0.25, 0.0675, 30, 3050, 60, 0.06, 3600, 1550, 0, 120, 0.06, 0.06, 0.08, 40, 0.03, 0.025, 0.03, 0.07),
    ("Oak Ave single-family", 1, 199000, 5970, 3500, 0.25, 0.07, 30, 1950, 0, 0.05, 2700, 1200, 0, 0, 0.06, 0.06, 0.0, 25, 0.03, 0.025, 0.035, 0.07),
    ("Riverside condo", 1, 179000, 5400, 0, 0.25, 0.069, 30, 1550, 0, 0.05, 2150, 650, 285, 0, 0.04, 0.04, 0.08, 0, 0.025, 0.03, 0.025, 0.07),
    ("Elm St fourplex", 4, 520000, 15600, 18000, 0.25, 0.0725, 30, 5600, 180, 0.07, 6900, 3100, 0, 380, 0.07, 0.07, 0.09, 120, 0.03, 0.025, 0.03, 0.07),
    ("Pine Rd (cash buy)", 1, 145000, 4350, 9000, 1.0, 0.0, 30, 1350, 0, 0.06, 1900, 1050, 0, 0, 0.07, 0.07, 0.0, 30, 0.03, 0.025, 0.03, 0.07),
]
KEYS_IN = ["name", "units", "price", "close", "repair", "down", "rate", "term", "rent", "oinc", "vac", "tax", "ins", "hoa", "util",
           "maint", "capex", "mgmt", "oexp", "rg", "eg", "ag", "sell"]
DEFAULTS = {"units": 1, "down": 0.25, "rate": 0.07, "term": 30, "vac": 0.06, "maint": 0.06, "capex": 0.06, "mgmt": 0.0,
            "rg": 0.03, "eg": 0.025, "ag": 0.03, "sell": 0.07}


def formulas(c):
    g = lambda k_: f"{c}{R[k_]}"
    has = f'{g("price")}>0'
    return {
        "loan": f"=IF({has},{g('price')}*MAX(0,1-{g('down')}),0)",
        "pmt": f"=IF({g('loan')}<=0,0,IF({g('rate')}=0,{g('loan')}/(MAX(1,{g('term')})*12),PMT({g('rate')}/12,MAX(1,{g('term')})*12,-{g('loan')})))",
        "cash": f"=IF({has},{g('price')}-{g('loan')}+{g('close')}+{g('repair')},0)",
        "gsi": f"=({g('rent')}+{g('oinc')})*12",
        "egi": f"={g('gsi')}*(1-{g('vac')})",
        "opex": f"={g('tax')}+{g('ins')}+({g('hoa')}+{g('util')}+{g('oexp')})*12+{g('rent')}*12*({g('maint')}+{g('capex')})+{g('egi')}*{g('mgmt')}",
        "noi": f"={g('egi')}-{g('opex')}",
        "ds": f"={g('pmt')}*12",
        "cf": f"={g('noi')}-{g('ds')}",
        "cfm": f"={g('cf')}/12",
        "cfu": f"=IFERROR({g('cfm')}/{g('units')},0)",
        "cap": f"=IFERROR({g('noi')}/{g('price')},0)",
        "coc": f"=IFERROR({g('cf')}/{g('cash')},0)",
        "dscr": f"=IFERROR({g('noi')}/{g('ds')},0)",
        "grm": f"=IFERROR({g('price')}/({g('rent')}*12),0)",
        "one": f"=IFERROR({g('rent')}/({g('price')}+{g('repair')}),0)",
        "exr": f"=IFERROR({g('opex')}/{g('egi')},0)",
        "beo": f"=IFERROR(({g('opex')}+{g('ds')})/{g('gsi')},0)",
        "irr": None,
        "verdict": (f'=IF(NOT({has}),"",IF({g("cf")}<0,"Loses money each month",IF(AND({g("coc")}>=0.08,OR({g("ds")}=0,{g("dscr")}>=1.25)),'
                    f'"Strong deal",IF({g("coc")}>=0.04,"OK – check the numbers","Weak return"))))'),
    }


def irr_block(ws, c, row0):
    """Hidden-ish helper table under the deals (rows row0..row0+11) with 11 yearly cash flows for each deal -> IRR."""
    g = lambda k_: f"{c}${R[k_]}"
    ws[f"{c}{row0}"] = f"=-{g('cash')}"
    for y in range(1, YRS + 1):
        cf = (f"(({g('rent')}+{g('oinc')})*12*(1+{g('rg')})^({y}-1)*(1-{g('vac')})"
              f"-({g('opex')})*(1+{g('eg')})^({y}-1)-{g('ds')}*IF({y}<={g('term')},1,0))")
        if y == YRS:
            bal = (f"IF({g('loan')}<=0,0,IF({YRS}>={g('term')},0,IF({g('rate')}=0,{g('loan')}-{g('pmt')}*12*{YRS},"
                   f"-FV({g('rate')}/12,12*{YRS},-{g('pmt')},{g('loan')}))))")
            cf += f"+{g('price')}*(1+{g('ag')})^{YRS}*(1-{g('sell')})-{bal}"
        ws[f"{c}{row0 + y}"] = "=" + cf
    return f"=IF({g('price')}>0,IFERROR(IRR({c}{row0}:{c}{row0 + YRS}),0),0)"


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); ds = wb.create_sheet(DEALS); pj = wb.create_sheet("Projection")

    # ---------- Deals ----------
    K.title(ds, "Deals", "One column per property. Yellow cells are yours to fill in – results calculate at the bottom.")
    ds["A3"] = "Currency label"; ds["A3"].font = font(bold=True); input_cell(ds["B3"], "USD")
    K.hdr(ds, 4, ["Item"] + [f"Deal {i + 1}" for i in range(ND)] + ["What it means"], [40] + [18] * ND + [70])
    r = 5
    for k_, lab, kind, fmt, hlp in ROWS:
        a = ds.cell(row=r, column=1, value=lab)
        if kind == "sec":
            a.font = Font(name=F, bold=True, color=MAIN)
            for col in range(1, D1 + 2): ds.cell(row=r, column=col).fill = fill(SOFT)
        else:
            a.font = font(bold=(kind == "out")); a.border = B
            h = ds.cell(row=r, column=D1 + 1, value=hlp or None); h.font = font(italic=True, size=9); h.alignment = Alignment(wrap_text=True)
        r += 1
    IR0 = LAST + 3
    for j in range(ND):
        col = D0 + j; c = L(col); fm = formulas(c)
        for k_, lab, kind, fmt, hlp in ROWS:
            if kind == "sec": continue
            cell = ds[f"{c}{R[k_]}"]
            if kind == "in":
                v = None
                if sample: v = SAMPLE[j][KEYS_IN.index(k_)]
                elif j == 0 and k_ in DEFAULTS: v = DEFAULTS[k_]
                input_cell(cell, v)
            else:
                cell.value = fm[k_] if k_ != "irr" else irr_block(ds, c, IR0)
                cell.fill = fill("F3F4F6"); cell.border = B; cell.font = font(bold=k_ in ("cf", "cfm", "coc", "cap", "verdict"))
            if fmt: cell.number_format = fmt
            cell.alignment = Alignment(horizontal="right" if k_ not in ("name", "verdict") else "left", wrap_text=k_ in ("name", "verdict"))
    ds[f"A{IR0 - 1}"] = "Helper: yearly cash flows used for the IRR (year 0 = cash in, year 10 includes the sale). No need to edit."
    ds[f"A{IR0 - 1}"].font = font(italic=True, size=9, color="6B7280")
    for y in range(YRS + 1):
        ds[f"A{IR0 + y}"] = f"Year {y}"; ds[f"A{IR0 + y}"].font = font(size=9, color="6B7280")
        for j in range(ND):
            cc = ds.cell(row=IR0 + y, column=D0 + j); cc.number_format = MONEY; cc.font = font(size=9, color="6B7280")
    dvp = DataValidation(type="decimal", operator="between", formula1="0", formula2="1", allow_blank=True,
                         error="Type a percentage, e.g. 25%", showErrorMessage=True)
    ds.add_data_validation(dvp)
    for k_, lab, kind, fmt, hlp in ROWS:
        if kind == "in" and fmt in (PCT, PCT2): dvp.add(f"{L(D0)}{R[k_]}:{L(D1)}{R[k_]}")
    ds.conditional_formatting.add(f"{L(D0)}{R['cf']}:{L(D1)}{R['cfu']}", FormulaRule(formula=[f"{L(D0)}{R['cf']}<0"], font=Font(name=F, bold=True, color="B91C1C")))
    ds.conditional_formatting.add(f"{L(D0)}{R['verdict']}:{L(D1)}{R['verdict']}", FormulaRule(formula=[f'{L(D0)}{R["verdict"]}="Strong deal"'], font=Font(name=F, bold=True, color="15803D")))
    ds.conditional_formatting.add(f"{L(D0)}{R['verdict']}:{L(D1)}{R['verdict']}", FormulaRule(formula=[f'{L(D0)}{R["verdict"]}="Loses money each month"'], font=Font(name=F, bold=True, color="B91C1C")))
    ds.freeze_panes = "B5"

    # ---------- Dashboard ----------
    K.title(dash, "Rental Property Dashboard", "Pick a deal in the yellow box – everything below updates. Edit numbers on the Deals tab.")
    dash["A3"] = "Deal to show (1–5)"; dash["A3"].font = font(bold=True)
    input_cell(dash["B3"], 1); dash["B3"].alignment = Alignment(horizontal="center")
    dv = DataValidation(type="list", formula1='"1,2,3,4,5"', allow_blank=False); dash.add_data_validation(dv); dv.add("B3")
    SEL = "Dashboard!$B$3"

    def pick(k_):
        return f"INDEX({DEALS}!$B${R[k_]}:${L(D1)}${R[k_]},{SEL})"
    dash["C3"] = f'=IF({pick("name")}="","(empty deal – fill it in on the Deals tab)",{pick("name")}&"  ·  "&{pick("verdict")})'
    dash["C3"].font = Font(name=F, bold=True, size=12, color=MAIN)
    kp = [("Cash flow per month", "=" + pick("cfm"), MONEY), ("Cash-on-cash return", "=" + pick("coc"), PCT2), ("Cap rate", "=" + pick("cap"), PCT2),
          ("Total cash needed", "=" + pick("cash"), MONEY), ("DSCR", "=" + pick("dscr"), X), ("10-year IRR", "=" + pick("irr"), PCT2)]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 3) * 3, 1 + (i % 3) * 2, lab, f, fmt)
        dash.cell(row=6 + (i // 3) * 3, column=1 + (i % 3) * 2).alignment = Alignment(horizontal="left")
    dash.conditional_formatting.add("A6", FormulaRule(formula=["A6<0"], font=Font(name=F, size=18, bold=True, color="B91C1C")))
    for c, w in zip("ABCDEFGH", [34, 22, 22, 30, 14, 16, 14, 14]): dash.column_dimensions[c].width = w

    K.section(dash, "A11", "Monthly breakdown")
    br = [("Rent + other income", f"=({pick('rent')}+{pick('oinc')})"), ("– Vacancy", f"=-({pick('gsi')}-{pick('egi')})/12"),
          ("– Operating expenses", f"=-{pick('opex')}/12"), ("→ Net operating income", f"={pick('noi')}/12"),
          ("– Mortgage (P&I)", f"=-{pick('pmt')}"), ("→ Cash flow", f"={pick('cfm')}")]
    for i, (lab, f) in enumerate(br, 12):
        dash[f"A{i}"] = lab; dash[f"B{i}"] = f
        for c in "AB":
            dash[f"{c}{i}"].border = B; dash[f"{c}{i}"].fill = fill(SOFT if lab.startswith("→") else "F3F4F6")
            dash[f"{c}{i}"].font = font(bold=lab.startswith("→"))
        dash[f"B{i}"].number_format = MONEY

    K.section(dash, "D11", "Rule-of-thumb checks")
    checks = [("1% rule (rent ÷ price)", pick("one"), PCT2, f'IF({pick("one")}>=0.01,"✔ Pass","✖ Below 1%")'),
              ("50% rule (expense ratio)", pick("exr"), PCT, f'IF({pick("exr")}<=0.5,"✔ Pass","✖ Above 50%")'),
              ("DSCR 1.25x or more", pick("dscr"), X, f'IF({pick("ds")}=0,"✔ No loan",IF({pick("dscr")}>=1.25,"✔ Pass","✖ Below 1.25x"))'),
              ("Positive cash flow", pick("cfm"), MONEY, f'IF({pick("cfm")}>0,"✔ Pass","✖ Negative")'),
              ("Cash-on-cash 8% or more", pick("coc"), PCT2, f'IF({pick("coc")}>=0.08,"✔ Pass","✖ Below 8%")'),
              ("Break-even occupancy", pick("beo"), PCT, f'IF({pick("beo")}<=0.85,"✔ Safe margin","✖ Tight")')]
    for i, (lab, f, fmt, res) in enumerate(checks, 12):
        dash[f"D{i}"] = lab; dash[f"E{i}"] = "=" + f; dash[f"F{i}"] = f'=IF({pick("price")}=0,"",{res})'
        dash[f"E{i}"].number_format = fmt
        for c in "DEF": dash[f"{c}{i}"].border = B; dash[f"{c}{i}"].fill = fill("F3F4F6"); dash[f"{c}{i}"].font = font()
    dash.conditional_formatting.add("F12:F17", FormulaRule(formula=['LEFT(F12,1)="✔"'], font=Font(name=F, bold=True, color="15803D")))
    dash.conditional_formatting.add("F12:F17", FormulaRule(formula=['LEFT(F12,1)="✖"'], font=Font(name=F, bold=True, color="B91C1C")))

    K.section(dash, "A20", "All deals side by side")
    cols = [("Deal", None), ("Property", "name"), ("Price", "price"), ("Cash needed", "cash"), ("Cash flow / month", "cfm"),
            ("Cash-on-cash", "coc"), ("Cap rate", "cap"), ("10-yr IRR", "irr")]
    K.hdr(dash, 21, [c[0] for c in cols])
    for j in range(ND):
        rr = 22 + j; c = L(D0 + j)
        dash[f"A{rr}"] = j + 1; dash[f"A{rr}"].alignment = Alignment(horizontal="center")
        for i, (lab, k_) in enumerate(cols[1:], 2):
            dash.cell(row=rr, column=i, value=f'=IF({DEALS}!{c}${R["price"]}=0,"",{DEALS}!{c}{R[k_]})' if k_ != "name"
                      else f'=IF({DEALS}!{c}{R["name"]}="","Deal {j + 1}",{DEALS}!{c}{R["name"]})')
    style_range(dash, 22, 21 + ND, 1, 8, fmt={3: MONEY, 4: MONEY, 5: MONEY, 6: PCT2, 7: PCT2, 8: PCT2})
    dash.conditional_formatting.add("A22:H26", FormulaRule(formula=[f"$A22={SEL}"], fill=fill("FDE68A")))
    dash.conditional_formatting.add("E22:E26", FormulaRule(formula=["AND(ISNUMBER(E22),E22<0)"], font=Font(name=F, color="B91C1C")))

    ch = BarChart(); ch.type = "col"; ch.title = "Cash flow per month by deal"; ch.height = 7; ch.width = 16; ch.legend = None
    ch.add_data(Reference(dash, min_col=5, min_row=21, max_row=26), titles_from_data=True)
    ch.set_categories(Reference(dash, min_col=2, min_row=22, max_row=26)); ch.y_axis.majorGridlines = None
    ch.series[0].graphicalProperties.solidFill = MAIN; dash.add_chart(ch, "J3")
    ch2 = BarChart(); ch2.type = "col"; ch2.grouping = "clustered"; ch2.title = "Returns by deal"; ch2.height = 7; ch2.width = 16
    ch2.add_data(Reference(dash, min_col=6, max_col=8, min_row=21, max_row=26), titles_from_data=True)
    ch2.set_categories(Reference(dash, min_col=2, min_row=22, max_row=26)); ch2.y_axis.majorGridlines = None; ch2.y_axis.number_format = "0%"
    for ser, col in zip(ch2.series, ("1D4ED8", "60A5FA", "F59E0B")): ser.graphicalProperties.solidFill = col
    dash.add_chart(ch2, "J18")

    # ---------- Projection ----------
    K.title(pj, "10-Year Projection", "For the deal chosen on the Dashboard. Uses the growth rates on the Deals tab.")
    pj["A3"] = "=Dashboard!C3"; pj["A3"].font = Font(name=F, bold=True, color=MAIN)
    hd = ["Year", "Rent + other income", "Vacancy", "Operating expenses", "NOI", "Mortgage payments", "Cash flow", "Cumulative cash flow",
          "Property value", "Loan balance", "Equity", "Total profit if sold", "Return on cash"]
    K.hdr(pj, 5, hd, [8, 18, 14, 18, 16, 18, 16, 18, 18, 16, 16, 18, 14])
    for y in range(1, YRS + 1):
        r = 5 + y
        pj[f"A{r}"] = y
        pj[f"B{r}"] = f"=({pick('rent')}+{pick('oinc')})*12*(1+{pick('rg')})^(A{r}-1)"
        pj[f"C{r}"] = f"=-B{r}*{pick('vac')}"
        pj[f"D{r}"] = f"=-{pick('opex')}*(1+{pick('eg')})^(A{r}-1)"
        pj[f"E{r}"] = f"=B{r}+C{r}+D{r}"
        pj[f"F{r}"] = f"=-{pick('ds')}*IF(A{r}<={pick('term')},1,0)"
        pj[f"G{r}"] = f"=E{r}+F{r}"
        pj[f"H{r}"] = f"=G{r}" if y == 1 else f"=H{r - 1}+G{r}"
        pj[f"I{r}"] = f"={pick('price')}*(1+{pick('ag')})^A{r}"
        pj[f"J{r}"] = (f"=IF({pick('loan')}<=0,0,IF(A{r}>={pick('term')},0,IF({pick('rate')}=0,MAX(0,{pick('loan')}-{pick('pmt')}*12*A{r}),"
                       f"-FV({pick('rate')}/12,12*A{r},-{pick('pmt')},{pick('loan')}))))")
        pj[f"K{r}"] = f"=I{r}-J{r}"
        pj[f"L{r}"] = f"=H{r}+I{r}*(1-{pick('sell')})-J{r}-{pick('cash')}"
        pj[f"M{r}"] = f"=IFERROR(L{r}/{pick('cash')},0)"
    style_range(pj, 6, 5 + YRS, 1, 13, fmt={**{c: MONEY for c in range(2, 13)}, 13: PCT})
    pj.conditional_formatting.add("C6:L15", FormulaRule(formula=["C6<0"], font=Font(name=F, color="B91C1C")))
    pj["A17"] = ("Total profit if sold = cumulative cash flow + sale price after selling costs – loan balance – your cash invested. "
                 "Return on cash = total profit ÷ cash invested. Taxes and depreciation are not included.")
    pj["A17"].font = font(italic=True, size=9)
    lc = LineChart(); lc.title = "Value, loan balance and equity"; lc.height = 8; lc.width = 18
    lc.add_data(Reference(pj, min_col=9, max_col=11, min_row=5, max_row=15), titles_from_data=True)
    lc.set_categories(Reference(pj, min_col=1, min_row=6, max_row=15)); lc.y_axis.majorGridlines = None
    for ser, col in zip(lc.series, ("1D4ED8", "F87171", "16A34A")): ser.graphicalProperties.line.solidFill = col
    pj.add_chart(lc, "A19")
    bc = BarChart(); bc.type = "col"; bc.title = "Cash flow per year"; bc.height = 8; bc.width = 14; bc.legend = None
    bc.add_data(Reference(pj, min_col=7, min_row=5, max_row=15), titles_from_data=True)
    bc.set_categories(Reference(pj, min_col=1, min_row=6, max_row=15)); bc.y_axis.majorGridlines = None
    bc.series[0].graphicalProperties.solidFill = "60A5FA"; pj.add_chart(bc, "H19")

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Deals tab: fill in the yellow cells for a property you are looking at – price, loan, rent and expenses. You can enter up to 5 deals side by side.", False),
             ("2.  The results at the bottom of each column calculate by themselves: cash flow, cap rate, cash-on-cash return, DSCR, 1% and 50% rules, IRR and a quick verdict.", False),
             ("3.  Dashboard: type the deal number (1–5) in the yellow box to see its key numbers, monthly breakdown and rule-of-thumb checks, plus all deals compared.", False),
             ("4.  Projection tab: 10 years of cash flow, loan paydown, property value, equity and total profit for the chosen deal.", False),
             ("", False), ("Tips", True),
             ("•  Be cautious: use real market rent from similar listings, and don't skip vacancy, maintenance and CapEx – they are real costs, even if not every month.", False),
             ("•  For a cash purchase set the down payment to 100% – the loan becomes 0.", False),
             ("•  Copy a deal column to the next one to test 'what if' versions (lower price, higher rent, different rate).", False),
             ("•  The verdict is a simple guide (cash flow, cash-on-cash ≥ 8% and DSCR ≥ 1.25x = strong). Every market is different.", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  This is an estimate tool, not financial advice. Taxes, depreciation and refinancing are not included.", False),
             ("•  Works with any currency (change the label on the Deals tab). No macros.", False),
             ("•  To use Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("The 5 properties and all numbers are made-up examples. Use the blank copy in your download for your own deals.", False)]
    K.start(st, "Rental Property Investment Analyzer", "Cash flow, cap rate, cash-on-cash, DSCR and a 10-year projection – compare 5 deals. Excel and Google Sheets.", lines)
    finish(wb, sample, path)


META = {
    "name": "Rental Property Investment Analyzer (Cash Flow, Cap Rate, Cash-on-Cash, 10-Year IRR) – Excel & Google Sheets",
    "file_base": "Rental-Property-Investment-Analyzer",
    "price": 24,
    "color": "#1D4ED8",
    "title": "Rental Property Analyzer",
    "tagline": "Cash flow · cap rate · cash-on-cash · DSCR · 10-year IRR — compare 5 deals",
    "intro": "Know if a rental property is a good deal before you buy it. Type the price, loan, rent and expenses and get cash flow, cap rate, cash-on-cash return, DSCR, the 1% and 50% rules and a 10-year projection with IRR – for up to 5 properties side by side.",
    "features": [
        "Compare up to 5 properties side by side (single-family, condo, duplex, small multifamily)",
        "Mortgage payment, total cash needed, NOI and monthly cash flow (also per unit) calculated for you",
        "Cap rate, cash-on-cash return, DSCR, gross rent multiplier, break-even occupancy",
        "Rule-of-thumb checks: 1% rule, 50% rule, DSCR 1.25x, positive cash flow – with a quick verdict",
        "10-year projection: rent and expense growth, loan paydown, property value, equity, total profit if sold",
        "10-year IRR including the sale (after selling costs)",
        "Dashboard with a deal picker and charts; works for cash purchases too; blank copy plus a sample copy",
    ],
    "tags": ["rental property analysis", "real estate investing", "cash flow calculator", "cap rate calculator", "cash on cash return",
             "rental property spreadsheet", "real estate calculator", "landlord template", "google sheets", "excel template",
             "investment property", "deal analyzer", "brrrr"],
}
