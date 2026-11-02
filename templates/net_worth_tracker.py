"""Net Worth Tracker – assets, debts and monthly snapshots for 24 months, FI number and goal (sample + blank versions).

Settings (start month, goal, yearly spending, categories) -> Balances (one row per account, one column per month) ->
Dashboard (latest net worth, change, by category, debt ratio, goal + FI progress, months to goal, charts).
"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.utils import get_column_letter as L
from _common import Kit, F, B, MONEY, DATE, font, fill, style_range, input_cell, finish

MAIN = "0F766E"; SOFT = "E6F4F1"
K = Kit(MAIN, SOFT)
NM = 24; C0 = 5; C1 = C0 + NM - 1            # month columns E..AB
NA = 40; R0 = 6; R1 = R0 + NA - 1            # account rows 6..45
TA, TL, TN, TC, TH = R1 + 2, R1 + 3, R1 + 4, R1 + 5, R1 + 6   # totals rows
CATS = [("Cash & savings", "Asset"), ("Investments", "Asset"), ("Retirement accounts", "Asset"), ("Real estate", "Asset"),
        ("Vehicles", "Asset"), ("Business", "Asset"), ("Other assets", "Asset"),
        ("Mortgage", "Debt"), ("Student loans", "Debt"), ("Car loans", "Debt"), ("Credit cards", "Debt"),
        ("Personal loans", "Debt"), ("Other debt", "Debt")]
NC = len(CATS); K0 = 5; K1 = K0 + NC - 1     # category list in Settings H5:I17
BAL = "Balances"
PCT = "0.0%"
SIGNED = '+#,##0.00;-#,##0.00;0.00'
BAR = 20


def bar(ref):
    return f'REPT("█",ROUND(MIN(1,MAX(0,N({ref})))*{BAR},0))&REPT("░",{BAR}-ROUND(MIN(1,MAX(0,N({ref})))*{BAR},0))'


SAMPLE = [  # name, category, note, start balance, monthly change
    ("Checking account", "Cash & savings", "Main bank", 3200, 60),
    ("High-yield savings", "Cash & savings", "Emergency fund", 9500, 420),
    ("Brokerage account", "Investments", "Index funds", 18400, 650),
    ("401(k)", "Retirement accounts", "Employer plan", 42300, 980),
    ("Roth IRA", "Retirement accounts", "", 15800, 560),
    ("Home", "Real estate", "Estimated value", 385000, 600),
    ("Car", "Vehicles", "Resale value", 21000, -260),
    ("Mortgage", "Mortgage", "30-year fixed", 296000, -640),
    ("Student loan", "Student loans", "", 23800, -410),
    ("Car loan", "Car loans", "", 14200, -390),
    ("Credit card", "Credit cards", "Paid in full most months", 1850, -60),
]


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); bal = wb.create_sheet(BAL); s = wb.create_sheet("Settings")

    # ---------- Settings ----------
    K.title(s, "Settings", "Yellow cells are yours to edit.")
    rows = [("Currency label", "USD", None),
            ("First month", dt.date(2026, 1, 1) if sample else "=DATE(YEAR(TODAY()),MONTH(TODAY()),1)", "mmm yyyy"),
            ("Net worth goal", 250000 if sample else 100000, MONEY),
            ("Yearly spending (for FI number)", 48000 if sample else 40000, MONEY),
            ("Safe withdrawal rate", 0.04, "0.0%")]
    for i, (k_, v, fmt) in enumerate(rows, 4):
        s.cell(row=i, column=1, value=k_).font = font(bold=True)
        c = s.cell(row=i, column=2); input_cell(c, v)
        if fmt: c.number_format = fmt
    notes = {5: "The first month column on the Balances tab. Type a date like 1/1/2026 (any day in the month works).",
             6: "The net worth you are working towards.",
             7: "What you spend in a year. Used to work out your FI (financial independence) number.",
             8: "4% is the classic rule of thumb: FI number = yearly spending ÷ 4% (= 25 × spending)."}
    for r, t in notes.items(): s[f"C{r}"] = t; s[f"C{r}"].font = font(italic=True, size=9)
    s["A10"] = "FI number"; s["A10"].font = font(italic=True)
    s["B10"] = "=IFERROR(B7/B8,0)"; s["B10"].number_format = MONEY; s["B10"].border = B; s["B10"].fill = fill("F3F4F6"); s["B10"].font = font()
    s["H4"] = "Category"; s["I4"] = "Type"
    for c in "HI": s[f"{c}4"].font = font(bold=True)
    for i, (cat, typ) in enumerate(CATS):
        r = K0 + i; s[f"H{r}"] = cat; s[f"I{r}"] = typ
        for c in "HI": s[f"{c}{r}"].font = Font(name=F, color="0000FF"); s[f"{c}{r}"].fill = fill("FFF8DB"); s[f"{c}{r}"].border = B
    s[f"H{K1 + 2}"] = "Rename a category if you like. Type must be Asset or Debt."; s[f"H{K1 + 2}"].font = font(italic=True, size=9)
    dv = DataValidation(type="list", formula1='"Asset,Debt"', allow_blank=False); s.add_data_validation(dv); dv.add(f"I{K0}:I{K1}")
    for c, w in zip("ABCDEFGHI", [32, 16, 86, 3, 3, 3, 3, 22, 10]): s.column_dimensions[c].width = w

    # ---------- Balances ----------
    K.title(bal, "Balances", "One row per account. Once a month, type each balance in that month's column. Debts as positive numbers.")
    bal["A3"] = "Tip: copy last month's column to the next one, then just change what moved."; bal["A3"].font = font(italic=True, size=9)
    K.hdr(bal, 5, ["Account", "Category", "Type", "Note"], [24, 20, 8, 20])
    for j in range(NM):
        col = C0 + j; c = bal.cell(row=5, column=col, value=f"=DATE(YEAR(Settings!$B$5),MONTH(Settings!$B$5)+{j},1)")
        c.number_format = "mmm yy"; c.font = Font(name=F, bold=True, color="FFFFFF"); c.fill = fill(MAIN); c.border = B
        c.alignment = Alignment(horizontal="center", vertical="center"); bal.column_dimensions[L(col)].width = 12
        bal.cell(row=4, column=col, value=j + 1).font = font(size=8, color="9CA3AF")
    bal["D4"] = "Month #"; bal["D4"].font = font(size=8, color="9CA3AF"); bal["D4"].alignment = Alignment(horizontal="right")
    style_range(bal, R0, R1, 1, C1, inputs=(1, 2, 4) + tuple(range(C0, C1 + 1)), fmt={c: MONEY for c in range(C0, C1 + 1)})
    cats = f"Settings!$H${K0}:$H${K1}"; types = f"Settings!$I${K0}:$I${K1}"
    for r in range(R0, R1 + 1):
        bal[f"C{r}"] = f'=IF(B{r}="","",IFERROR(INDEX({types},MATCH(B{r},{cats},0)),"?"))'
        bal[f"C{r}"].alignment = Alignment(horizontal="center")
    if sample:
        for i, (nm, cat, note, b0, d) in enumerate(SAMPLE):
            r = R0 + i; bal[f"A{r}"] = nm; bal[f"B{r}"] = cat; bal[f"D{r}"] = note or None
            for m in range(10):
                v = b0 + d * m + ((m * 37 + i * 11) % 7 - 3) * abs(d) * 0.35
                if nm == "Brokerage account" and m == 3: v -= 1400        # a market dip
                bal.cell(row=r, column=C0 + m, value=round(max(0, v), 2))
    dv = DataValidation(type="list", formula1=f"={cats}", allow_blank=True); bal.add_data_validation(dv); dv.add(f"B{R0}:B{R1}")
    labels = {TA: "Total assets", TL: "Total debts", TN: "Net worth", TC: "Change vs previous month", TH: "Month has data? (month #)"}
    for r, t in labels.items():
        bal[f"A{r}"] = t; bal[f"A{r}"].font = font(bold=True)
        for col in range(1, C1 + 1):
            c = bal.cell(row=r, column=col); c.border = B; c.fill = fill(SOFT if r != TH else "F3F4F6")
    for j in range(NM):
        cl = L(C0 + j); rng = f"{cl}{R0}:{cl}{R1}"
        bal[f"{cl}{TA}"] = f'=SUMIFS({rng},$C${R0}:$C${R1},"Asset")'
        bal[f"{cl}{TL}"] = f'=SUMIFS({rng},$C${R0}:$C${R1},"Debt")'
        bal[f"{cl}{TN}"] = f'=IF({cl}{TH}=0,"",{cl}{TA}-{cl}{TL})'
        bal[f"{cl}{TC}"] = "" if j == 0 else f'=IF(OR({cl}{TH}=0,{L(C0 + j - 1)}{TH}=0),"",{cl}{TN}-{L(C0 + j - 1)}{TN})'
        bal[f"{cl}{TH}"] = f"=IF(COUNT({rng})>0,{cl}4,0)"
        for r, fmt in ((TA, MONEY), (TL, MONEY), (TN, MONEY), (TC, SIGNED), (TH, "0")):
            bal[f"{cl}{r}"].number_format = fmt; bal[f"{cl}{r}"].font = font(bold=(r == TN), color=MAIN if r == TN else "1F2A37")
    bal.conditional_formatting.add(f"E{TC}:{L(C1)}{TC}", FormulaRule(formula=[f'AND(ISNUMBER(E{TC}),E{TC}<0)'], font=Font(name=F, color="B91C1C")))
    bal.conditional_formatting.add(f"C{R0}:C{R1}", FormulaRule(formula=[f'C{R0}="Debt"'], font=Font(name=F, bold=True, color="B91C1C")))
    bal.freeze_panes = f"{L(C0)}{R0}"

    # ---------- Dashboard ----------
    LAST = "Dashboard!$H$3"; FIRST = "Dashboard!$J$3"
    K.title(dash, "Net Worth Dashboard", "Updates by itself from the Balances tab.")
    dash["G3"] = "Latest month #"; dash["H3"] = f"=MAX({BAL}!$E${TH}:${L(C1)}${TH})"
    dash["I3"] = "First month #"; dash["J3"] = f'=IFERROR(MATCH(TRUE,INDEX({BAL}!$E${TH}:${L(C1)}${TH}>0,0),0),0)'
    for c in "GHIJ": dash[f"{c}3"].font = font(size=8, color="FFFFFF")   # helper cells, hidden

    def at(row, idx):
        return f"IF({idx}=0,0,INDEX({BAL}!$E${row}:${L(C1)}${row},{idx}))"
    dash["A3"] = (f'=IF(H3=0,"Add your balances on the Balances tab to get started.","Latest month: "&TEXT(INDEX({BAL}!$E$5:${L(C1)}$5,H3),"mmmm yyyy")'
                  f'&"  ·  "&H3-J3+1&" month(s) tracked  ·  Amounts in "&Settings!B4)')
    dash["A3"].font = font(bold=True)
    kp = [("Net worth", "=" + at(TN, "H3"), MONEY),
          ("Change vs last month", f"=IF(H3<=1,0,{at(TN, 'H3')}-{at(TN, 'H3-1')})", SIGNED),
          ("Change since you started", f"={at(TN, 'H3')}-{at(TN, 'J3')}", SIGNED),
          ("Total assets", "=" + at(TA, "H3"), MONEY),
          ("Total debts", "=" + at(TL, "H3"), MONEY),
          ("Debt-to-asset ratio", "=IFERROR(" + at(TL, "H3") + "/" + at(TA, "H3") + ",0)", PCT)]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 3) * 3, 1 + (i % 3) * 2, lab, f, fmt)
        dash.cell(row=6 + (i // 3) * 3, column=1 + (i % 3) * 2).alignment = Alignment(horizontal="left")
    for c, w in zip("ABCDEFGHIJ", [36, 18, 26, 18, 26, 18, 14, 8, 14, 8]): dash.column_dimensions[c].width = w

    def block(r0, title, rows_):
        K.section(dash, f"A{r0}", title)
        for i, (lab, f, fmt) in enumerate(rows_, r0 + 1):
            dash[f"A{i}"] = lab; dash[f"A{i}"].font = font(bold=True); dash[f"A{i}"].border = B; dash[f"A{i}"].fill = fill(SOFT)
            c = dash[f"B{i}"]; c.value = f; c.number_format = fmt; c.border = B; c.font = Font(name=F, bold=True, color=MAIN)
            c.alignment = Alignment(horizontal="right")
    block(11, "Goals", [
        ("Net worth goal", "=Settings!B6", MONEY),
        ("Progress to goal", "=IFERROR(MAX(0,A6)/B12,0)", PCT),
        ("FI number (spending ÷ withdrawal rate)", "=Settings!B10", MONEY),
        ("Progress to FI", "=IFERROR(MAX(0,A6)/B14,0)", PCT),
        ("Average change per month", '=IF(H3-J3<1,0,(' + at(TN, "H3") + "-" + at(TN, "J3") + ")/(H3-J3))", SIGNED),
        ("Months to goal at this pace", '=IF(A6>=B12,"Reached!",IF(B16<=0,"–",ROUNDUP((B12-A6)/B16,0)))', "0"),
        ("Goal reached around", '=IF(ISNUMBER(B17),DATE(YEAR(INDEX(' + BAL + '!$E$5:$' + L(C1) + '$5,MAX(1,H3))),MONTH(INDEX(' + BAL + '!$E$5:$' + L(C1) + '$5,MAX(1,H3)))+B17,1),B17)', "mmm yyyy")])
    dash["C13"] = f'={bar("B13")}'; dash["C15"] = f'={bar("B15")}'
    for c in ("C13", "C15"): dash[c].font = Font(name=F, color=MAIN)

    # by category (latest month)
    K.section(dash, "A20", "By category (latest month)")
    K.hdr(dash, 21, ["Category", "Type", "Amount", "Share of its type"])
    for i in range(NC):
        r = 22 + i; k = K0 + i
        dash[f"A{r}"] = f"=Settings!H{k}"; dash[f"B{r}"] = f"=Settings!I{k}"
        dash[f"C{r}"] = (f'=IF($H$3=0,0,SUMIFS(INDEX({BAL}!$E${R0}:${L(C1)}${R1},0,$H$3),{BAL}!$B${R0}:$B${R1},A{r}))')
        dash[f"D{r}"] = f'=IFERROR(C{r}/IF(B{r}="Debt",{at(TL, "$H$3")},{at(TA, "$H$3")}),0)'
    style_range(dash, 22, 21 + NC, 1, 4, fmt={3: MONEY, 4: PCT})
    dash.conditional_formatting.add(f"A22:D{21 + NC}", FormulaRule(formula=['$B22="Debt"'], font=Font(name=F, color="B91C1C")))

    # monthly history (for charts)
    HR = 22 + NC + 2
    K.section(dash, f"A{HR}", "Month by month")
    K.hdr(dash, HR + 1, ["Month", "Assets", "Debts", "Net worth", "Change"])
    for j in range(NM):
        r = HR + 2 + j; cl = L(C0 + j)
        dash[f"A{r}"] = f"={BAL}!{cl}5"
        dash[f"B{r}"] = f'=IF({BAL}!{cl}{TH}=0,"",{BAL}!{cl}{TA})'
        dash[f"C{r}"] = f'=IF({BAL}!{cl}{TH}=0,"",{BAL}!{cl}{TL})'
        dash[f"D{r}"] = f'=IF({BAL}!{cl}{TH}=0,"",{BAL}!{cl}{TN})'
        dash[f"E{r}"] = f'=IF({BAL}!{cl}{TH}=0,"",{BAL}!{cl}{TC})' if j else '=""'
    H1 = HR + 1 + NM
    style_range(dash, HR + 2, H1, 1, 5, fmt={1: "mmm yyyy", 2: MONEY, 3: MONEY, 4: MONEY, 5: SIGNED})
    dash.conditional_formatting.add(f"E{HR + 2}:E{H1}", FormulaRule(formula=[f'AND(ISNUMBER(E{HR + 2}),E{HR + 2}<0)'], font=Font(name=F, color="B91C1C")))

    ch = BarChart(); ch.type = "col"; ch.title = "Net worth by month"; ch.height = 7.5; ch.width = 17; ch.legend = None
    ch.add_data(Reference(dash, min_col=4, min_row=HR + 1, max_row=H1), titles_from_data=True)
    ch.set_categories(Reference(dash, min_col=1, min_row=HR + 2, max_row=H1)); ch.y_axis.majorGridlines = None
    ch.series[0].graphicalProperties.solidFill = MAIN; ch.x_axis.number_format = "mmm yy"
    dash.add_chart(ch, "G5")
    ch2 = BarChart(); ch2.type = "col"; ch2.grouping = "clustered"; ch2.title = "Assets vs debts"; ch2.height = 7.5; ch2.width = 17
    ch2.add_data(Reference(dash, min_col=2, max_col=3, min_row=HR + 1, max_row=H1), titles_from_data=True)
    ch2.set_categories(Reference(dash, min_col=1, min_row=HR + 2, max_row=H1)); ch2.y_axis.majorGridlines = None
    for ser, col in zip(ch2.series, ("14B8A6", "F87171")): ser.graphicalProperties.solidFill = col
    dash.add_chart(ch2, "G21")
    pie = PieChart(); pie.title = "Assets by category"; pie.height = 7.5; pie.width = 12
    pie.add_data(Reference(dash, min_col=3, min_row=22, max_row=28)); pie.set_categories(Reference(dash, min_col=1, min_row=22, max_row=28))
    dash.add_chart(pie, f"G{HR}")

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Settings tab: choose your first month, your net worth goal and your yearly spending (for your FI number).", False),
             ("2.  Balances tab: list every account – bank, savings, investments, retirement, home, car… and every debt (mortgage, loans, credit cards). Choose a category for each one.", False),
             ("3.  Once a month (for example on the 1st), type the balance of each account in that month's column. Debts are typed as positive numbers.", False),
             ("4.  Dashboard: your net worth, the change since last month and since you started, debt-to-asset ratio, goal and FI progress, and charts.", False),
             ("", False), ("Tips", True),
             ("•  Use realistic values for your home and car (what you could sell them for today), or leave them out if you prefer a 'liquid' net worth.", False),
             ("•  Fastest monthly update: copy last month's column into the new month, then change only what moved.", False),
             ("•  A falling month is normal (markets go down sometimes). Look at the trend over 6–12 months.", False),
             ("•  Rename the categories on the Settings tab if you like (for example 'Crypto' or 'Pension'). Type must be Asset or Debt.", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  Tracks 40 accounts for 24 months. Works with any currency (change the label on the Settings tab). Year 3? Save a copy and set the new first month.", False),
             ("•  To use Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("The accounts and balances (January – October 2026) are made-up examples. Use the blank copy in your download for your own numbers.", False)]
    K.start(st, "Net Worth Tracker", "Assets, debts and your net worth month by month – with goal and FI progress. Works in Excel and Google Sheets.", lines)
    finish(wb, sample, path)


META = {
    "name": "Net Worth Tracker (Assets, Debts, Monthly Snapshots, FI Goal) – Excel & Google Sheets",
    "file_base": "Net-Worth-Tracker",
    "price": 14,
    "color": "#0F766E",
    "title": "Net Worth Tracker",
    "tagline": "Assets · debts · monthly net worth · goal & FI progress — Excel & Google Sheets",
    "intro": "See your whole financial picture on one page. Type your account balances once a month and watch your net worth grow – with charts, your debt-to-asset ratio, and how far you are from your goal and financial independence.",
    "features": [
        "Track up to 40 accounts (cash, investments, retirement, real estate, vehicles, mortgage, loans, credit cards) for 24 months",
        "Net worth, change vs last month and since you started – calculated for you",
        "Debt-to-asset ratio and totals by category (latest month)",
        "Net worth goal and FI number (yearly spending ÷ 4%) with progress bars",
        "Average monthly growth, months to reach your goal and the expected date",
        "Charts: net worth by month, assets vs debts, assets by category",
        "Editable categories, any currency, no macros, blank copy plus a sample copy",
    ],
    "tags": ["net worth tracker", "net worth spreadsheet", "fire tracker", "financial independence", "asset tracker", "debt tracker",
             "google sheets", "excel template", "personal finance", "wealth tracker", "money tracker", "financial planner", "budget template"],
}
