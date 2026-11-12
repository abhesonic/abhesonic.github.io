"""Sinking Funds Tracker – savings goals with target dates, monthly & per-paycheck amounts, deposit/withdrawal log, 12-month view.

Settings -> Funds -> Transactions -> 12-Month View -> Dashboard.
"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, DataBarRule
from openpyxl.chart import BarChart, Reference
from _common import Kit, F, B, MONEY, DATE, font, fill, style_range, input_cell, finish

MAIN = "0F766E"; SOFT = "CCFBF1"
K = Kit(MAIN, SOFT)
NF = 40; U0 = 6; U1 = U0 + NF - 1          # funds
NT = 1000; T0 = 6; T1 = T0 + NT - 1        # transactions
CATS = ["Home", "Car", "Holidays & gifts", "Travel", "Health & dental", "Kids & school", "Pets", "Insurance & taxes",
        "Tech & phone", "Clothing", "Emergency", "Fun & hobbies", "Other"]
NCAT = 15; C0 = 6; C1 = C0 + NCAT - 1
PCT = "0%"


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); fu = wb.create_sheet("Funds"); tr = wb.create_sheet("Transactions")
    mv = wb.create_sheet("12-Month View"); s = wb.create_sheet("Settings")

    # ---------- Settings ----------
    K.title(s, "Settings", "Yellow cells are yours to edit.")
    SET = [("Currency (label only)", "USD"), ("Paychecks per month (e.g. 1, 2 or 4.33 for weekly)", 2),
           ("12-month view starts (blank = this month)", None)]
    for i, (lab, v) in enumerate(SET, 5):
        s[f"A{i}"] = lab; s[f"A{i}"].font = font(bold=True); s[f"A{i}"].border = B; input_cell(s[f"C{i}"], v)
    s["C7"].number_format = "mmm yyyy"
    s["A8"] = "12-month view starts"; s["A8"].font = font(italic=True)
    s["C8"] = '=IF(ISNUMBER(C7),DATE(YEAR(C7),MONTH(C7),1),DATE(YEAR(TODAY()),MONTH(TODAY()),1))'; s["C8"].number_format = "mmm yyyy"
    K.hdr(s, 5, ["Categories"], col0=6)
    for i in range(NCAT): input_cell(s[f"F{C0 + i}"], CATS[i] if i < len(CATS) else None)
    for c, w in zip("ABCDEF", [46, 3, 14, 3, 3, 24]): s.column_dimensions[c].width = w
    PAY = "Settings!$C$6"

    # ---------- Funds ----------
    K.title(fu, "Funds", "One row per sinking fund. Yellow = you type, grey = automatic. Deposits and withdrawals come from the Transactions tab.")
    cols = ["Fund", "Category", "Target amount", "Target date", "Starting balance", "I plan to save / month", "Priority", "Notes",
            "Deposits", "Withdrawals", "Saved now", "Still to save", "% funded", "Months left", "Needed / month", "Needed / paycheck",
            "Status", "Fully funded by (at your plan)"]
    K.hdr(fu, 5, cols, [24, 17, 12, 12, 12, 12, 10, 26, 11, 11, 12, 12, 9, 8, 12, 12, 26, 14])
    style_range(fu, U0, U1, 1, 18, inputs=range(1, 9),
                fmt={3: MONEY, 4: DATE, 5: MONEY, 6: MONEY, 9: MONEY, 10: MONEY, 11: MONEY, 12: MONEY, 13: PCT, 14: "0",
                     15: MONEY, 16: MONEY, 18: "mmm yyyy"})
    TR = lambda c: f"Transactions!${c}${T0}:${c}${T1}"
    for r in range(U0, U1 + 1):
        fu[f"I{r}"] = f'=IF(A{r}="","",SUMIFS({TR("D")},{TR("B")},A{r},{TR("C")},"Deposit"))'
        fu[f"J{r}"] = f'=IF(A{r}="","",SUMIFS({TR("D")},{TR("B")},A{r},{TR("C")},"Withdrawal"))'
        fu[f"K{r}"] = f'=IF(A{r}="","",N(E{r})+I{r}-J{r})'
        fu[f"L{r}"] = f'=IF(OR(A{r}="",NOT(ISNUMBER(C{r}))),"",MAX(0,C{r}-K{r}))'
        fu[f"M{r}"] = f'=IF(OR(A{r}="",N(C{r})<=0),"",MIN(1,MAX(0,K{r}/C{r})))'
        fu[f"N{r}"] = f'=IF(OR(A{r}="",NOT(ISNUMBER(D{r}))),"",MAX(1,(YEAR(D{r})-YEAR(TODAY()))*12+MONTH(D{r})-MONTH(TODAY())))'
        fu[f"O{r}"] = f'=IF(OR(L{r}="",N{r}=""),"",L{r}/N{r})'
        fu[f"P{r}"] = f'=IF(O{r}="","",O{r}/MAX(1,N({PAY})))'
        fu[f"Q{r}"] = (f'=IF(A{r}="","",IF(NOT(ISNUMBER(C{r})),"Type a target amount",IF(K{r}>=C{r},"Fully funded",'
                       f'IF(NOT(ISNUMBER(D{r})),"Type a target date",IF(D{r}<TODAY(),"Target date passed",'
                       f'IF(N(F{r})>=O{r}-0.005,"On track","Behind – add "&TEXT(O{r}-N(F{r}),"#,##0.00")&" / month"))))))')
        fu[f"R{r}"] = (f'=IF(OR(L{r}="",A{r}=""),"",IF(L{r}=0,TODAY(),IF(N(F{r})<=0,"no plan yet",'
                       f'EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),ROUNDUP(L{r}/F{r},0)))))')
        fu[f"Q{r}"].font = font(bold=True)
    for col, lst in (("B", f"=Settings!$F${C0}:$F${C1}"), ("G", '"High,Medium,Low"')):
        dv = DataValidation(type="list", formula1=lst, allow_blank=True); fu.add_data_validation(dv); dv.add(f"{col}{U0}:{col}{U1}")
    fu.conditional_formatting.add(f"Q{U0}:Q{U1}", FormulaRule(formula=[f'OR(LEFT(Q{U0},6)="Behind",LEFT(Q{U0},6)="Target")'],
                                                             font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    fu.conditional_formatting.add(f"Q{U0}:Q{U1}", FormulaRule(formula=[f'OR(Q{U0}="On track",Q{U0}="Fully funded")'],
                                                             font=Font(name=F, bold=True, color="166534"), fill=fill("DCFCE7")))
    fu.conditional_formatting.add(f"M{U0}:M{U1}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="2DD4BF"))
    fu.freeze_panes = f"B{U0}"; fu.auto_filter.ref = f"A5:R{U1}"

    # ---------- Transactions ----------
    K.title(tr, "Transactions", "Every time you move money into a fund (Deposit) or spend from it (Withdrawal), add one row.")
    K.hdr(tr, 5, ["Date", "Fund", "Type", "Amount", "Note"], [12, 24, 12, 12, 40])
    style_range(tr, T0, T1, 1, 5, inputs=range(1, 6), fmt={1: DATE, 4: MONEY})
    dv = DataValidation(type="list", formula1=f"=Funds!$A${U0}:$A${U1}", allow_blank=True); tr.add_data_validation(dv); dv.add(f"B{T0}:B{T1}")
    dv = DataValidation(type="list", formula1='"Deposit,Withdrawal"', allow_blank=True); tr.add_data_validation(dv); dv.add(f"C{T0}:C{T1}")
    tr.conditional_formatting.add(f"C{T0}:D{T1}", FormulaRule(formula=[f'$C{T0}="Withdrawal"'], font=Font(name=F, color="B91C1C")))
    tr.freeze_panes = f"A{T0}"; tr.auto_filter.ref = f"A5:E{T1}"

    # ---------- 12-Month View ----------
    K.title(mv, "12-Month View", "Net amount saved in each fund per month (deposits minus withdrawals).")
    K.hdr(mv, 5, ["Fund"] + [""] * 12 + ["12-month total"], [24] + [10] * 12 + [13])
    for j in range(12):
        c = mv.cell(row=5, column=2 + j, value=f"=EDATE(Settings!$C$8,{j})"); c.number_format = "mmm yy"
        c.font = Font(name=F, bold=True, color="FFFFFF")
    mv["A6"] = "Total per month"
    for j in range(13):
        Lc = CL(2 + j); mv[f"{Lc}6"] = f"=SUM({Lc}7:{Lc}{6 + NF})"
    for c in range(1, 15):
        cell = mv.cell(row=6, column=c); cell.font = Font(name=F, bold=True, color=MAIN); cell.fill = fill(SOFT); cell.border = B
        if c >= 2: cell.number_format = MONEY
    for k in range(NF):
        r = 7 + k; fr = U0 + k
        mv[f"A{r}"] = f'=IF(Funds!$A${fr}="","",Funds!$A${fr})'
        for j in range(12):
            Lc = CL(2 + j); ms = f"{Lc}$5"; me = f"EOMONTH({Lc}$5,0)"
            base = f'{TR("D")},{TR("B")},$A{r},{TR("A")},">="&{ms},{TR("A")},"<="&{me}'
            mv[f"{Lc}{r}"] = f'=IF($A{r}="",0,SUMIFS({base},{TR("C")},"Deposit")-SUMIFS({base},{TR("C")},"Withdrawal"))'
        mv[f"N{r}"] = f"=SUM(B{r}:M{r})"
    style_range(mv, 7, 6 + NF, 1, 14, fmt={c: MONEY for c in range(2, 15)})
    for r in range(7, 7 + NF): mv[f"N{r}"].font = font(bold=True)
    mv.freeze_panes = "B7"

    # ---------- sample ----------
    if sample:
        funds = [("Car repairs & tires", "Car", 1200, dt.date(2027, 6, 30), 300, 110, "High", "Brakes due next spring"),
                 ("Christmas & birthdays", "Holidays & gifts", 900, dt.date(2026, 12, 15), 100, 100, "High", "List of 14 people"),
                 ("Summer vacation", "Travel", 3000, dt.date(2027, 7, 1), 600, 250, "Medium", "Beach week with the kids"),
                 ("Home maintenance", "Home", 2000, dt.date(2027, 12, 31), 400, 110, "Medium", "Gutters, roof check"),
                 ("Car insurance (yearly)", "Insurance & taxes", 1140, dt.date(2027, 3, 1), 250, 180, "High", "Pay in full to save 8%"),
                 ("Dentist & glasses", "Health & dental", 650, dt.date(2027, 4, 30), 120, 55, "Medium", ""),
                 ("Back to school", "Kids & school", 500, dt.date(2027, 8, 15), 100, 45, "Low", "Clothes, supplies, laptop bag"),
                 ("Vet & pet insurance", "Pets", 600, dt.date(2027, 5, 31), 150, 60, "Medium", "Annual check-up + vaccines"),
                 ("New phone", "Tech & phone", 900, dt.date(2027, 9, 30), 100, 0, "Low", "Waiting for the new model"),
                 ("Emergency fund", "Emergency", 6000, dt.date(2028, 6, 30), 2500, 200, "High", "3 months of expenses"),
                 ("Property tax", "Insurance & taxes", 1800, dt.date(2026, 11, 30), 900, 150, "High", "Second instalment"),
                 ("Winter clothes", "Clothing", 350, dt.date(2026, 11, 10), 350, 0, "Low", "Already bought most")]
        for k, row in enumerate(funds):
            for c, v in zip("ABCDEFGH", row):
                if v not in (None, ""): fu[f"{c}{U0 + k}"] = v
        txs = []
        months = [dt.date(2026, 6, 1), dt.date(2026, 7, 1), dt.date(2026, 8, 1), dt.date(2026, 9, 1), dt.date(2026, 10, 1), dt.date(2026, 11, 1)]
        plan = {f[0]: f[5] for f in funds}
        for m in months:
            for name, amt in plan.items():
                if amt and not (name == "Back to school" and m.month < 9):
                    for half, day in ((0.5, 1), (0.5, 15)):
                        if m.month == 11 and day == 15: continue
                        txs.append((m.replace(day=day), name, "Deposit", round(amt * half, 2), "Paycheck transfer"))
        txs += [(dt.date(2026, 7, 22), "Car repairs & tires", "Withdrawal", 385.40, "Two new tires"),
                (dt.date(2026, 8, 9), "Vet & pet insurance", "Withdrawal", 142.00, "Vaccines"),
                (dt.date(2026, 9, 3), "Back to school", "Withdrawal", 95.60, "School supplies"),
                (dt.date(2026, 10, 18), "Home maintenance", "Withdrawal", 260.00, "Plumber – kitchen sink"),
                (dt.date(2026, 10, 25), "Christmas & birthdays", "Withdrawal", 74.99, "Mom's birthday"),
                (dt.date(2026, 11, 5), "Winter clothes", "Withdrawal", 186.30, "Coats for the kids"),
                (dt.date(2026, 11, 8), "Emergency fund", "Deposit", 300, "Tax refund"),
                (dt.date(2026, 6, 20), "Summer vacation", "Deposit", 150, "Sold old bike")]
        txs.sort(key=lambda t: t[0])
        for k, t in enumerate(txs):
            for c, v in zip("ABCDE", t): tr[f"{c}{T0 + k}"] = v
        s["C7"] = dt.date(2026, 6, 1)

    # ---------- Dashboard ----------
    FR = lambda c: f"Funds!${c}${U0}:${c}${U1}"
    K.title(dash, "Sinking Funds Dashboard", "Updates by itself from the Funds and Transactions tabs.")
    dash["A3"] = '="Amounts in "&Settings!C5&"  ·  today: "&TEXT(TODAY(),"mmm d, yyyy")'; dash["A3"].font = font(bold=True)
    kp = [("Total of all targets", f"=SUM({FR('C')})", MONEY), ("Saved so far", f"=SUM({FR('K')})", MONEY),
          ("Still to save", f"=SUM({FR('L')})", MONEY),
          ("Overall % funded", f"=IF(SUM({FR('C')})=0,0,MIN(1,SUM({FR('K')})/SUM({FR('C')})))", PCT),
          ("Needed per month (all funds)", f"=SUM({FR('O')})", MONEY), ("Needed per paycheck", f"=SUM({FR('P')})", MONEY),
          ("You plan per month", f"=SUM({FR('F')})", MONEY),
          ("Funds behind / past date", f'=COUNTIFS({FR("Q")},"Behind*")+COUNTIFS({FR("Q")},"Target date passed")', "0"),
          ("Fully funded", f'=COUNTIFS({FR("Q")},"Fully funded")', "0")]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 3) * 3, 1 + (i % 3) * 2, lab, f, fmt)
        dash.cell(row=6 + (i // 3) * 3, column=1 + (i % 3) * 2).alignment = Alignment(horizontal="left")
    for c, w in zip("ABCDEFGHI", [26, 13, 24, 13, 26, 13, 13, 31, 10]): dash.column_dimensions[c].width = w
    H = 15
    K.section(dash, f"A{H - 1}", "All funds")
    K.hdr(dash, H, ["Fund", "Saved", "Target", "Still to save", "% funded", "Needed / month", "Target date", "Status", "Priority"])
    D0 = H + 1; D1 = H + NF
    for k in range(NF):
        r = D0 + k; fr = U0 + k
        dash[f"A{r}"] = f'=IF(Funds!A{fr}="","",Funds!A{fr})'
        for col, src in zip("BCDEFGHI", "KCLMODQG"):
            dash[f"{col}{r}"] = f'=IF(A{r}="","",IF(Funds!{src}{fr}="","",Funds!{src}{fr}))'
    style_range(dash, D0, D1, 1, 9, fmt={2: MONEY, 3: MONEY, 4: MONEY, 5: PCT, 6: MONEY, 7: "mmm d, yyyy"})
    dash.conditional_formatting.add(f"E{D0}:E{D1}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="2DD4BF"))
    dash.conditional_formatting.add(f"H{D0}:H{D1}", FormulaRule(formula=[f'OR(LEFT(H{D0},6)="Behind",LEFT(H{D0},6)="Target")'],
                                                               font=Font(name=F, bold=True, color="B91C1C")))
    dash.conditional_formatting.add(f"H{D0}:H{D1}", FormulaRule(formula=[f'OR(H{D0}="On track",H{D0}="Fully funded")'],
                                                               font=Font(name=F, bold=True, color="166534")))
    # helper for chart: saved capped at target + remaining
    for k in range(NF):
        r = D0 + k
        dash[f"K{r}"] = f'=IF(A{r}="","",MIN(N(B{r}),N(C{r})))'; dash[f"L{r}"] = f'=IF(A{r}="","",N(D{r}))'
        dash[f"K{r}"].font = font(color="FFFFFF", size=8); dash[f"L{r}"].font = font(color="FFFFFF", size=8)
    dash[f"K{H}"] = "Saved"; dash[f"L{H}"] = "Still to save"
    dash[f"K{H}"].font = font(color="FFFFFF", size=8); dash[f"L{H}"].font = font(color="FFFFFF", size=8)
    dash.column_dimensions["J"].width = 3
    ch = BarChart(); ch.type = "bar"; ch.grouping = "stacked"; ch.overlap = 100; ch.title = "Saved vs still to save"
    ch.height = 9; ch.width = 16
    ch.add_data(Reference(dash, min_col=11, max_col=12, min_row=H, max_row=H + 12), titles_from_data=True)
    ch.set_categories(Reference(dash, min_col=1, min_row=D0, max_row=H + 12))
    ch.x_axis.scaling.orientation = "maxMin"; ch.y_axis.majorGridlines = None
    ch.series[0].graphicalProperties.solidFill = MAIN; ch.series[1].graphicalProperties.solidFill = "CBD5E1"
    dash.add_chart(ch, "K4")
    ch2 = BarChart(); ch2.type = "col"; ch2.title = "Saved per month (net)"; ch2.height = 7.5; ch2.width = 16; ch2.legend = None
    ch2.add_data(Reference(mv, min_col=2, max_col=13, min_row=6, max_row=6), from_rows=True, titles_from_data=False)
    ch2.set_categories(Reference(mv, min_col=2, max_col=13, min_row=5, max_row=5)); ch2.y_axis.majorGridlines = None
    ch2.series[0].graphicalProperties.solidFill = "5EEAD4"; ch2.gapWidth = 60
    dash.add_chart(ch2, "K24")

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Settings tab: how many paychecks you get per month (1, 2, or 4.33 if weekly) and your currency label. Edit the categories if you like.", False),
             ("2.  Funds tab: one row per sinking fund – name, category, target amount, target date, what is already saved (starting balance) and what you plan to save per month.", False),
             ("     Months left, needed per month / per paycheck, % funded, status and 'fully funded by' are calculated for you.", False),
             ("3.  Transactions tab: each time you put money into a fund, add a 'Deposit' row. Each time you spend from it, add a 'Withdrawal' row.", False),
             ("4.  12-Month View: how much went into each fund every month.", False),
             ("5.  Dashboard: total saved, still to save, what you need per month and per paycheck, funds that are behind, a progress bar per fund and charts.", False),
             ("", False), ("Tips", True),
             ("•  A sinking fund is money you set aside bit by bit for a cost you know is coming. Start with 3–5 funds for costs that surprised you last year.", False),
             ("•  'Behind' means your planned amount per month is lower than what is needed to reach the target in time – raise the plan, move the date or lower the target.", False),
             ("•  Keep all funds in one savings account if you like – this sheet keeps track of how much belongs to each fund.", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  40 funds, 1,000 transactions, 15 categories. No macros. Works in Excel 2010+ and Google Sheets.", False),
             ("•  Months left and needed amounts update with today's date every time you open the file.", False),
             ("•  To use Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("The funds and amounts are made-up examples. Use the blank copy in your download for your own funds.", False)]
    K.start(st, "Sinking Funds Tracker", "Save for every planned expense – per month and per paycheck – Excel and Google Sheets.", lines)
    finish(wb, sample, path)


META = {
    "name": "Sinking Funds Tracker (Savings Goals, Per-Paycheck Amounts, Progress Bars, Deposit Log) – Excel & Google Sheets",
    "file_base": "Sinking-Funds-Tracker",
    "price": 9,
    "color": "#0F766E",
    "title": "Sinking Funds Tracker",
    "tagline": "needed per month & per paycheck · progress bars · on-track status",
    "intro": "Stop being surprised by car repairs, gifts, insurance and holidays. Give each planned expense its own sinking fund with a target amount and date – the tracker tells you exactly how much to set aside per month and per paycheck, shows a progress bar for every fund and warns you when a fund is falling behind.",
    "features": [
        "40 sinking funds: category, target amount, target date, starting balance, planned amount per month, priority and notes",
        "Automatic months left, amount needed per month and per paycheck (monthly, twice a month or weekly pay)",
        "Status per fund: On track, Behind (with the extra amount needed), Fully funded or Target date passed",
        "'Fully funded by' month at your current plan",
        "Deposit / withdrawal log (1,000 rows) with drop-down lists – balances update by themselves",
        "12-month view of how much went into each fund every month",
        "Dashboard: total saved, still to save, % funded, needed per month and per paycheck, progress bars and charts",
        "Blank copy plus a sample copy filled with example funds",
    ],
    "tags": ["sinking funds", "savings tracker", "savings goals", "budget template", "google sheets",
             "excel template", "personal finance", "money planner", "paycheck budget", "savings planner", "budget planner", "money goals"],
    "summary": "Instant download: Excel + Google Sheets template, blank copy + sample copy.",
}
