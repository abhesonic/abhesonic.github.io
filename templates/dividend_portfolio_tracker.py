"""Dividend Income & Stock Portfolio Tracker – holdings, yield, yield on cost, income calendar, dividends received (sample + blank).

Settings (currency, as-of, monthly income goal, sectors) -> Holdings (shares, cost, price, dividend, frequency, pay months) ->
Income Calendar (projected income per month per holding) -> Dividend Log (what actually arrived, tax withheld, DRIP) ->
Dashboard (value, gain, projected income, yield, YOC, received this year vs projected, sectors, goal, charts).
"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.utils import get_column_letter as L
from _common import Kit, F, B, MONEY, DATE, font, fill, style_range, input_cell, finish

MAIN = "1D4ED8"; SOFT = "E8EEFC"
K = Kit(MAIN, SOFT)
NH = 50; H0 = 5; H1 = H0 + NH - 1
NL = 500; L0 = 5; L1 = L0 + NL - 1
FREQS = [("Monthly", 12), ("Quarterly", 4), ("Semi-annual", 2), ("Annual", 1)]
SECTORS = ["Consumer staples", "Healthcare", "Financials", "Utilities", "Energy", "Real estate (REIT)", "Industrials",
           "Technology", "Telecom", "Consumer discretionary", "Materials", "ETF / fund", "Other"]
NS = len(SECTORS); S0 = 5; S1 = S0 + NS - 1
HD = "Holdings"; CAL = "'Income Calendar'"; LOG = "'Dividend Log'"
ASOF = "Settings!$B$5"
PCT = "0.00%"; SIGNED = '+#,##0.00;-#,##0.00;0.00'
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
BAR = 20


def bar(ref):
    return f'REPT("█",ROUND(MIN(1,MAX(0,N({ref})))*{BAR},0))&REPT("░",{BAR}-ROUND(MIN(1,MAX(0,N({ref})))*{BAR},0))'


# ticker, name, sector, shares, avg cost, price, annual dividend/share, frequency, first pay month
SAMPLE = [("KO", "Coca-Cola", "Consumer staples", 60, 52.40, 69.80, 2.04, "Quarterly", 4),
          ("PG", "Procter & Gamble", "Consumer staples", 25, 138.00, 158.20, 4.23, "Quarterly", 2),
          ("JNJ", "Johnson & Johnson", "Healthcare", 30, 149.50, 171.00, 5.20, "Quarterly", 3),
          ("O", "Realty Income", "Real estate (REIT)", 120, 54.10, 57.30, 3.23, "Monthly", 1),
          ("VZ", "Verizon", "Telecom", 80, 38.90, 41.10, 2.73, "Quarterly", 2),
          ("XOM", "Exxon Mobil", "Energy", 35, 98.00, 116.40, 4.00, "Quarterly", 3),
          ("JPM", "JPMorgan Chase", "Financials", 20, 151.00, 284.00, 6.00, "Quarterly", 1),
          ("NEE", "NextEra Energy", "Utilities", 40, 66.30, 76.90, 2.27, "Quarterly", 3),
          ("MSFT", "Microsoft", "Technology", 10, 320.00, 505.00, 3.64, "Quarterly", 3),
          ("SCHD", "Schwab US Dividend Equity ETF", "ETF / fund", 220, 25.10, 27.40, 1.03, "Quarterly", 3),
          ("MAIN", "Main Street Capital", "Financials", 50, 41.20, 59.60, 3.06, "Monthly", 1)]


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); hd = wb.create_sheet(HD); cal = wb.create_sheet("Income Calendar")
    lg = wb.create_sheet("Dividend Log"); s = wb.create_sheet("Settings")

    # ---------- Settings ----------
    K.title(s, "Settings", "Yellow cells are yours to edit.")
    rows = [("Currency label", "USD", None),
            ("Today / as-of date", dt.date(2026, 10, 30) if sample else "=TODAY()", DATE),
            ("Monthly dividend income goal", 1000 if sample else 500, MONEY)]
    for i, (k_, v, fmt) in enumerate(rows, 4):
        s.cell(row=i, column=1, value=k_).font = font(bold=True)
        c = s.cell(row=i, column=2); input_cell(c, v)
        if fmt: c.number_format = fmt
    notes = {5: "Leave as =TODAY(). The year of this date is 'this year' on the dashboard.",
             6: "The passive income you are working towards, per month."}
    for r, t in notes.items(): s[f"C{r}"] = t; s[f"C{r}"].font = font(italic=True, size=9)
    s["H4"] = "Pay frequency"; s["I4"] = "Payments / year"; s["K4"] = "Sectors"
    for c in "HIK": s[f"{c}4"].font = font(bold=True)
    for i, (fq, n) in enumerate(FREQS): s.cell(row=5 + i, column=8, value=fq).font = font(); s.cell(row=5 + i, column=9, value=n).font = font()
    for i, sec in enumerate(SECTORS):
        c = s.cell(row=S0 + i, column=11, value=sec); input_cell(c)
    s[f"K{S1 + 2}"] = "Rename or add your own sectors here."; s[f"K{S1 + 2}"].font = font(italic=True, size=9)
    for c, w in zip("ABCDEFGHIJK", [30, 16, 70, 3, 3, 3, 3, 14, 15, 3, 26]): s.column_dimensions[c].width = w
    FQ = "Settings!$H$5:$H$8"; FN = "Settings!$I$5:$I$8"; SEC = f"Settings!$K${S0}:$K${S1}"

    # ---------- Holdings ----------
    K.title(hd, "Holdings", "One row per stock, ETF or fund. Update the price column whenever you like (once a month is plenty).")
    cols = ["Ticker", "Name", "Sector", "Shares", "Avg cost / share", "Current price", "Annual dividend / share", "Pay frequency",
            "First pay month (1-12)", "Cost basis", "Market value", "Gain / loss", "Gain %", "Dividend yield", "Yield on cost",
            "Annual income", "Monthly average", "Weight", "Received this year"]
    K.hdr(hd, 4, cols, [9, 24, 18, 9, 12, 11, 12, 12, 11, 13, 13, 13, 9, 9, 9, 12, 11, 8, 12])
    style_range(hd, H0, H1, 1, 19, inputs=range(1, 10),
                fmt={4: "#,##0.####", 5: MONEY, 6: MONEY, 7: "#,##0.0000", 10: MONEY, 11: MONEY, 12: SIGNED, 13: "0.0%",
                     14: PCT, 15: PCT, 16: MONEY, 17: MONEY, 18: "0.0%", 19: MONEY})
    if sample:
        for i, row in enumerate(SAMPLE):
            for j, v in enumerate(row, 1): hd.cell(row=H0 + i, column=j, value=v)
    TOTV = f"$K${H1 + 1}"
    for r in range(H0, H1 + 1):
        e = f'A{r}=""'
        hd[f"J{r}"] = f'=IF({e},"",N(D{r})*N(E{r}))'
        hd[f"K{r}"] = f'=IF({e},"",N(D{r})*N(F{r}))'
        hd[f"L{r}"] = f'=IF({e},"",K{r}-J{r})'
        hd[f"M{r}"] = f'=IF(OR({e},N(J{r})=0),"",L{r}/J{r})'
        hd[f"N{r}"] = f'=IF(OR({e},N(F{r})=0),"",N(G{r})/F{r})'
        hd[f"O{r}"] = f'=IF(OR({e},N(E{r})=0),"",N(G{r})/E{r})'
        hd[f"P{r}"] = f'=IF({e},"",N(D{r})*N(G{r}))'
        hd[f"Q{r}"] = f'=IF({e},"",P{r}/12)'
        hd[f"R{r}"] = f'=IF(OR({e},N({TOTV})=0),"",K{r}/{TOTV})'
        hd[f"S{r}"] = (f'=IF({e},"",SUMIFS({LOG}!$F${L0}:$F${L1},{LOG}!$B${L0}:$B${L1},A{r},'
                       f'{LOG}!$A${L0}:$A${L1},">="&DATE(YEAR({ASOF}),1,1),{LOG}!$A${L0}:$A${L1},"<"&DATE(YEAR({ASOF})+1,1,1)))')
    hd[f"A{H1 + 1}"] = "Total"
    for c in "JKLPQS": hd[f"{c}{H1 + 1}"] = f"=SUM({c}{H0}:{c}{H1})"
    hd[f"M{H1 + 1}"] = f"=IFERROR(L{H1 + 1}/J{H1 + 1},0)"; hd[f"N{H1 + 1}"] = f"=IFERROR(P{H1 + 1}/K{H1 + 1},0)"
    hd[f"O{H1 + 1}"] = f"=IFERROR(P{H1 + 1}/J{H1 + 1},0)"
    for c in "AJKLMNOPQS":
        x = hd[f"{c}{H1 + 1}"]; x.font = font(bold=True); x.border = B; x.fill = fill(SOFT)
        x.number_format = {"M": "0.0%", "N": PCT, "O": PCT, "L": SIGNED}.get(c, MONEY)
    for r in range(H0, H1 + 1):
        hd[f"A{r}"].font = Font(name=F, bold=True, color="0000FF"); hd[f"H{r}"].alignment = Alignment(horizontal="center")
        hd[f"I{r}"].alignment = Alignment(horizontal="center")
    hd.conditional_formatting.add(f"L{H0}:M{H1}", FormulaRule(formula=[f'AND(ISNUMBER(L{H0}),L{H0}<0)'], font=Font(name=F, color="B91C1C")))
    hd.conditional_formatting.add(f"L{H0}:M{H1}", FormulaRule(formula=[f'AND(ISNUMBER(L{H0}),L{H0}>0)'], font=Font(name=F, color="15803D")))
    for rng, f1 in ((f"C{H0}:C{H1}", f"={SEC}"), (f"H{H0}:H{H1}", f"={FQ}")):
        dv = DataValidation(type="list", formula1=f1, allow_blank=True); hd.add_data_validation(dv); dv.add(rng)
    dv = DataValidation(type="whole", operator="between", formula1="1", formula2="12", allow_blank=True); hd.add_data_validation(dv); dv.add(f"I{H0}:I{H1}")
    hd[f"A{H1 + 3}"] = ("Annual dividend / share = the yearly total per share (e.g. 4 quarterly payments of 0.51 = 2.04). "
                        "First pay month = the first month of the year it pays (e.g. 3 for Mar, Jun, Sep, Dec).")
    hd[f"A{H1 + 3}"].font = font(italic=True, size=9)
    hd.freeze_panes = "B5"

    # ---------- Income Calendar ----------
    K.title(cal, "Income Calendar", "Projected dividends per month, from your Holdings. Nothing to type here.")
    K.hdr(cal, 4, ["Ticker"] + MONTHS + ["Total / year"], [10] + [10] * 12 + [13])
    style_range(cal, H0, H1, 1, 14, fmt={c: MONEY for c in range(2, 15)})
    for i in range(NH):
        r = H0 + i; hr = H0 + i
        cal[f"A{r}"] = f'=IF({HD}!A{hr}="","",{HD}!A{hr})'
        npay = f"IFERROR(INDEX({FN},MATCH({HD}!$H{hr},{FQ},0)),4)"
        for m in range(12):
            cl = L(2 + m)
            cal[f"{cl}{r}"] = (f'=IF($A{r}="","",IF(MOD({m + 1}-IF(N({HD}!$I{hr})=0,1,{HD}!$I{hr}),12/{npay})=0,'
                               f'N({HD}!$P{hr})/{npay},0))')
        cal[f"N{r}"] = f'=IF(A{r}="","",SUM(B{r}:M{r}))'
        cal[f"N{r}"].font = font(bold=True)
    cal[f"A{H1 + 1}"] = "Total"
    for m in range(13):
        cl = L(2 + m); cal[f"{cl}{H1 + 1}"] = f"=SUM({cl}{H0}:{cl}{H1})"
    for m in range(14):
        x = cal.cell(row=H1 + 1, column=1 + m); x.font = font(bold=True); x.border = B; x.fill = fill(SOFT)
        if m: x.number_format = MONEY
    cal.conditional_formatting.add(f"B{H0}:M{H1}", FormulaRule(formula=[f'AND(ISNUMBER(B{H0}),B{H0}>0)'], fill=fill("BFDBFE")))
    cal.freeze_panes = "B5"

    # ---------- Dividend Log ----------
    K.title(lg, "Dividend Log", "Each dividend that lands in your account. Choose the ticker from the list.")
    K.hdr(lg, 4, ["Pay date", "Ticker", "Gross amount", "Tax withheld", "Reinvested? (x)", "Net received", "Note"], [12, 10, 13, 12, 12, 13, 28])
    style_range(lg, L0, L1, 1, 7, inputs=(1, 2, 3, 4, 5, 7), fmt={1: DATE, 3: MONEY, 4: MONEY, 6: MONEY})
    for r in range(L0, L1 + 1):
        lg[f"F{r}"] = f'=IF(C{r}="","",N(C{r})-N(D{r}))'; lg[f"E{r}"].alignment = Alignment(horizontal="center")
    if sample:
        rows_ = []
        for (tk, nm, sec, sh, cost, px, div, fq, first) in SAMPLE:
            n = dict(FREQS)[fq]; step = 12 // n
            for m in range(1, 11):
                if (m - first) % step == 0:
                    gross = round(sh * div / n, 2)
                    day = 15 if n == 12 else (1 + (len(tk) * 7) % 27)
                    rows_.append((dt.date(2026, m, day), tk, gross, round(gross * 0.0, 2), "x" if tk in ("SCHD", "O") else None))
        rows_.sort()
        for i, (d, tk, g, t, x) in enumerate(rows_):
            r = L0 + i; lg[f"A{r}"] = d; lg[f"B{r}"] = tk; lg[f"C{r}"] = g; lg[f"D{r}"] = t or None; lg[f"E{r}"] = x
    dv = DataValidation(type="list", formula1=f"={HD}!$A${H0}:$A${H1}", allow_blank=True); lg.add_data_validation(dv); dv.add(f"B{L0}:B{L1}")
    dv = DataValidation(type="list", formula1='"x"', allow_blank=True); lg.add_data_validation(dv); dv.add(f"E{L0}:E{L1}")
    lg[f"I4"] = "Tax withheld: leave empty if none (for example in an IRA, Roth or ISA). Net received = gross − tax."
    lg["I4"].font = font(italic=True, size=9)
    lg.freeze_panes = "A5"

    # ---------- Dashboard ----------
    K.title(dash, "Dividend Income Dashboard", "Updates by itself from Holdings and the Dividend Log.")
    dash["A3"] = f'="As of "&TEXT({ASOF},"mmm d, yyyy")&"  ·  "&COUNTIF({HD}!A{H0}:A{H1},"?*")&" holdings  ·  Amounts in "&Settings!B4'
    dash["A3"].font = font(bold=True)
    T = H1 + 1
    kp = [("Portfolio value", f"={HD}!K{T}", MONEY),
          ("Unrealized gain / loss", f"={HD}!L{T}", SIGNED),
          ("Projected annual income", f"={HD}!P{T}", MONEY),
          ("Portfolio yield", f"={HD}!N{T}", PCT),
          ("Yield on cost", f"={HD}!O{T}", PCT),
          ("Received this year (net)", f"={HD}!S{T}", MONEY)]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 3) * 3, 1 + (i % 3) * 2, lab, f, fmt)
        dash.cell(row=6 + (i // 3) * 3, column=1 + (i % 3) * 2).alignment = Alignment(horizontal="left")
    for c, w in zip("ABCDEFGH", [34, 16, 26, 16, 26, 16, 4, 16]): dash.column_dimensions[c].width = w

    def block(r0, title, rows_):
        K.section(dash, f"A{r0}", title)
        for i, (lab, f, fmt) in enumerate(rows_, r0 + 1):
            dash[f"A{i}"] = lab; dash[f"A{i}"].font = font(bold=True); dash[f"A{i}"].border = B; dash[f"A{i}"].fill = fill(SOFT)
            c = dash[f"B{i}"]; c.value = f; c.number_format = fmt; c.border = B; c.font = Font(name=F, bold=True, color=MAIN)
            c.alignment = Alignment(horizontal="right")
    block(11, "Income goal", [
        ("Monthly income goal", "=Settings!B6", MONEY),
        ("Projected monthly income (average)", f"={HD}!Q{T}", MONEY),
        ("Progress to goal", "=IFERROR(B13/B12,0)", "0.0%"),
        ("Portfolio needed at today's yield", "=IFERROR(B12*12/A9,0)", MONEY),
        ("Still to invest at today's yield", "=MAX(0,B15-A6)", MONEY),
        ("Cost basis (money invested)", f"={HD}!J{T}", MONEY),
        ("Reinvested this year (DRIP, net)", (f'=SUMIFS({LOG}!$F${L0}:$F${L1},{LOG}!$E${L0}:$E${L1},"x",{LOG}!$A${L0}:$A${L1},">="&DATE(YEAR({ASOF}),1,1),'
                                                f'{LOG}!$A${L0}:$A${L1},"<"&DATE(YEAR({ASOF})+1,1,1))'), MONEY)])
    dash["C14"] = f'={bar("B14")}'; dash["C14"].font = Font(name=F, color=MAIN)

    # month table: projected vs received this year
    K.section(dash, "A20", "Month by month (this year)")
    K.hdr(dash, 21, ["Month", "Projected", "Received (net)", "Difference"])
    for m in range(12):
        r = 22 + m; cl = L(2 + m)
        dash[f"A{r}"] = f"=DATE(YEAR({ASOF}),{m + 1},1)"
        dash[f"B{r}"] = f"={CAL}!{cl}{T}"
        dash[f"C{r}"] = (f'=SUMIFS({LOG}!$F${L0}:$F${L1},{LOG}!$A${L0}:$A${L1},">="&A{r},{LOG}!$A${L0}:$A${L1},"<"&DATE(YEAR(A{r}),MONTH(A{r})+1,1))')
        dash[f"D{r}"] = f'=IF(A{r}>{ASOF},"",ROUND(C{r}-B{r},2))'
    dash["A34"] = "Total"; dash["B34"] = "=SUM(B22:B33)"; dash["C34"] = "=SUM(C22:C33)"
    style_range(dash, 22, 34, 1, 4, fmt={1: "mmm yyyy", 2: MONEY, 3: MONEY, 4: SIGNED})
    for c in "ABC": dash[f"{c}34"].font = font(bold=True); dash[f"{c}34"].fill = fill(SOFT)
    dash.conditional_formatting.add("D22:D33", FormulaRule(formula=['AND(ISNUMBER(D22),D22<0)'], font=Font(name=F, color="B91C1C")))
    dash["A35"] = "Difference < 0 can just mean a payment date moved, or tax was withheld."; dash["A35"].font = font(italic=True, size=9)

    # by sector
    K.section(dash, "A37", "By sector")
    K.hdr(dash, 38, ["Sector", "Market value", "Weight", "Annual income", "Share of income"])
    for i in range(NS):
        r = 39 + i; k = S0 + i
        dash[f"A{r}"] = f'=IF(Settings!K{k}="","",Settings!K{k})'
        dash[f"B{r}"] = f'=IF(A{r}="",0,SUMIFS({HD}!$K${H0}:$K${H1},{HD}!$C${H0}:$C${H1},A{r}))'
        dash[f"C{r}"] = f"=IFERROR(B{r}/{HD}!$K${T},0)"
        dash[f"D{r}"] = f'=IF(A{r}="",0,SUMIFS({HD}!$P${H0}:$P${H1},{HD}!$C${H0}:$C${H1},A{r}))'
        dash[f"E{r}"] = f"=IFERROR(D{r}/{HD}!$P${T},0)"
    style_range(dash, 39, 38 + NS, 1, 5, fmt={2: MONEY, 3: "0.0%", 4: MONEY, 5: "0.0%"})

    ch = BarChart(); ch.type = "col"; ch.grouping = "clustered"; ch.title = "Dividends per month: projected vs received"
    ch.height = 8; ch.width = 18; ch.y_axis.majorGridlines = None
    ch.add_data(Reference(dash, min_col=2, max_col=3, min_row=21, max_row=33), titles_from_data=True)
    ch.set_categories(Reference(dash, min_col=1, min_row=22, max_row=33)); ch.x_axis.number_format = "mmm"
    for ser, col in zip(ch.series, ("BFDBFE", MAIN)): ser.graphicalProperties.solidFill = col
    dash.add_chart(ch, "H5")
    pie = PieChart(); pie.title = "Income by sector"; pie.height = 8; pie.width = 13
    pie.add_data(Reference(dash, min_col=4, min_row=39, max_row=38 + NS)); pie.set_categories(Reference(dash, min_col=1, min_row=39, max_row=38 + NS))
    dash.add_chart(pie, "H22")

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Settings tab: set your currency label and your monthly dividend income goal.", False),
             ("2.  Holdings tab: one row per stock, ETF or fund – ticker, shares, average cost, current price, annual dividend per share, how often it pays and the first month of the year it pays.", False),
             ("3.  Income Calendar: fills itself – you see which months bring how much income, per holding.", False),
             ("4.  Dividend Log: when a dividend arrives, add the date, ticker and amount (and tax withheld, if any). Type x if it was reinvested (DRIP).", False),
             ("5.  Dashboard: portfolio value, gain/loss, projected yearly income, yield, yield on cost, received this year vs projected, sectors and your goal.", False),
             ("", False), ("Tips", True),
             ("•  Prices: update them once a month. In Google Sheets you can type =GOOGLEFINANCE(\"KO\",\"price\") in a price cell to get it live.", False),
             ("•  Annual dividend per share: look up the 'forward annual dividend' on any finance site, or multiply the last payment by the payments per year.", False),
             ("•  Yield on cost = income ÷ what you paid. It grows as companies raise their dividends – a great number to watch.", False),
             ("•  Pay months are an estimate (first pay month + frequency). The Dividend Log shows what really arrived.", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  50 holdings, 500 log rows, any currency. This is a tracking tool, not investment advice.", False),
             ("•  To use Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("The holdings, prices and dividends are made-up examples (as-of date Oct 30, 2026) – not current market data and not a recommendation. "
                   "Use the blank copy in your download for your own portfolio.", False)]
    K.start(st, "Dividend Income Tracker", "Stock portfolio, dividend yield, income calendar and dividends received. Works in Excel and Google Sheets.", lines)
    finish(wb, sample, path)


META = {
    "name": "Dividend Income & Stock Portfolio Tracker (Yield, Income Calendar) – Excel & Google Sheets",
    "file_base": "Dividend-Income-Portfolio-Tracker",
    "price": 17,
    "color": "#1D4ED8",
    "title": "Dividend Income Tracker",
    "tagline": "Portfolio · yield · yield on cost · income calendar — Excel & Google Sheets",
    "intro": "Watch your passive income grow. Track your dividend stocks and ETFs, see what each month will pay you, log the dividends that actually arrive and follow your progress to your monthly income goal.",
    "features": [
        "Holdings for up to 50 stocks, ETFs or funds: cost basis, market value, gain/loss, weight",
        "Dividend yield, yield on cost, annual and monthly income per holding and for the whole portfolio",
        "Income calendar: projected dividends per month for every holding (monthly, quarterly, semi-annual, annual payers)",
        "Dividend log with tax withheld and DRIP (reinvested) tracking – received this year per holding",
        "Dashboard: projected vs received per month, income by sector, value and income charts",
        "Income goal: progress bar and how much more you need to invest at today's yield",
        "Any currency, no macros, blank copy plus a sample copy (prices update by hand or with GOOGLEFINANCE in Sheets)",
    ],
    "tags": ["dividend tracker", "dividend spreadsheet", "stock portfolio tracker", "passive income", "dividend income", "investment tracker",
             "google sheets", "excel template", "yield on cost", "fire", "personal finance", "dividend calendar", "stock tracker"],
}
