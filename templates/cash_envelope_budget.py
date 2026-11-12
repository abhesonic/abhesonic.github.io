"""Cash Envelope Budget – envelopes, fill/spend log, monthly dashboard, cash withdrawal planner (bill breakdown), year summary.

Settings -> Envelopes -> Transactions -> Cash Withdrawal -> Year Summary -> Dashboard.
"""
import datetime as dt, random
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, DataBarRule
from openpyxl.chart import BarChart, Reference
from _common import Kit, F, B, MONEY, DATE, font, fill, style_range, input_cell, finish

MAIN = "B45309"; SOFT = "FEF3C7"
K = Kit(MAIN, SOFT)
NE = 30; E0 = 6; E1 = E0 + NE - 1          # envelopes
NT = 1500; T0 = 6; T1 = T0 + NT - 1        # transactions
DEN = [100, 50, 20, 10, 5, 1]; ND = len(DEN); D0 = 6; D1 = D0 + ND - 1
PCT = "0%"
TYPES = '"Fill,Spend,Move in,Move out"'


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); en = wb.create_sheet("Envelopes"); tr = wb.create_sheet("Transactions")
    cw = wb.create_sheet("Cash Withdrawal"); ys = wb.create_sheet("Year Summary"); s = wb.create_sheet("Settings")

    # ---------- Settings ----------
    K.title(s, "Settings", "Yellow cells are yours to edit.")
    SET = [("Currency (label only)", "USD"), ("Paychecks per month (1, 2, or 4.33 if weekly)", 2),
           ("Year Summary starts (blank = January this year)", None)]
    for i, (lab, v) in enumerate(SET, 5):
        s[f"A{i}"] = lab; s[f"A{i}"].font = font(bold=True); s[f"A{i}"].border = B; input_cell(s[f"C{i}"], v)
    s["C7"].number_format = "mmm yyyy"
    s["A8"] = "Year Summary starts"; s["A8"].font = font(italic=True)
    s["C8"] = '=IF(ISNUMBER(C7),DATE(YEAR(C7),MONTH(C7),1),DATE(YEAR(TODAY()),1,1))'; s["C8"].number_format = "mmm yyyy"
    K.hdr(s, 5, ["Bill values (largest first)"], col0=6)
    for i in range(ND): input_cell(s[f"F{D0 + i}"], DEN[i])
    s[f"F{D1 + 1}"] = "Change to the bills your bank gives out. Use 0 to skip one."; s[f"F{D1 + 1}"].font = font(italic=True, size=9)
    for c, w in zip("ABCDEF", [46, 3, 14, 3, 3, 26]): s.column_dimensions[c].width = w
    PAY = "Settings!$C$6"

    # ---------- Envelopes ----------
    K.title(en, "Envelopes", "One row per envelope. Cash envelopes get a bill breakdown on the Cash Withdrawal tab.")
    K.hdr(en, 5, ["Envelope", "Type", "Monthly budget", "Budget per paycheck", "Cash per paycheck (optional – your own round number)",
                  "Cash to withdraw per paycheck", "Notes"], [24, 10, 13, 13, 20, 14, 34])
    style_range(en, E0, E1, 1, 7, inputs=(1, 2, 3, 5, 7), fmt={3: MONEY, 4: MONEY, 5: MONEY, 6: MONEY})
    for r in range(E0, E1 + 1):
        en[f"D{r}"] = f'=IF(OR(A{r}="",NOT(ISNUMBER(C{r}))),"",C{r}/MAX(1,N({PAY})))'
        en[f"F{r}"] = f'=IF(A{r}="","",IF(B{r}<>"Cash",0,IF(ISNUMBER(E{r}),E{r},IF(D{r}="",0,ROUND(D{r},0)))))'
    dv = DataValidation(type="list", formula1='"Cash,Digital"', allow_blank=True); en.add_data_validation(dv); dv.add(f"B{E0}:B{E1}")
    en.row_dimensions[5].height = 44
    r = E1 + 1
    en[f"A{r}"] = "Total"; en[f"C{r}"] = f"=SUM(C{E0}:C{E1})"; en[f"D{r}"] = f"=SUM(D{E0}:D{E1})"; en[f"F{r}"] = f"=SUM(F{E0}:F{E1})"
    for c in "ACDF":
        en[f"{c}{r}"].font = font(bold=True); en[f"{c}{r}"].fill = fill(SOFT); en[f"{c}{r}"].border = B
        if c != "A": en[f"{c}{r}"].number_format = MONEY
    en.freeze_panes = f"B{E0}"

    # ---------- Transactions ----------
    K.title(tr, "Transactions", "Fill = cash put in the envelope · Spend = money spent from it · Move in / Move out = moving money between envelopes (two rows).")
    K.hdr(tr, 5, ["Date", "Envelope", "Type", "Amount", "Where / what"], [12, 24, 12, 12, 40])
    style_range(tr, T0, T1, 1, 5, inputs=range(1, 6), fmt={1: DATE, 4: MONEY})
    for r in range(T0, T1 + 1):
        tr[f"F{r}"] = f'=IF(OR(B{r}="",NOT(ISNUMBER(D{r}))),0,IF(OR(C{r}="Fill",C{r}="Move in"),D{r},-D{r}))'
    tr.column_dimensions["F"].hidden = True
    dv = DataValidation(type="list", formula1=f"=Envelopes!$A${E0}:$A${E1}", allow_blank=True); tr.add_data_validation(dv); dv.add(f"B{T0}:B{T1}")
    dv = DataValidation(type="list", formula1=TYPES, allow_blank=True); tr.add_data_validation(dv); dv.add(f"C{T0}:C{T1}")
    tr.conditional_formatting.add(f"C{T0}:D{T1}", FormulaRule(formula=[f'$C{T0}="Spend"'], font=Font(name=F, color="B91C1C")))
    tr.conditional_formatting.add(f"C{T0}:D{T1}", FormulaRule(formula=[f'$C{T0}="Fill"'], font=Font(name=F, color="166534")))
    tr.freeze_panes = f"A{T0}"; tr.auto_filter.ref = f"A5:E{T1}"
    TR = lambda c: f"Transactions!${c}${T0}:${c}${T1}"

    # ---------- Cash Withdrawal ----------
    K.title(cw, "Cash Withdrawal Planner", "How many of each bill to put in every cash envelope on payday – and how many to ask the bank for.")
    hd = ["Envelope", "Cash per paycheck"] + [f'=IF(N(Settings!$F${D0 + i})=0,"-","Bills of "&Settings!$F${D0 + i})' for i in range(ND)] + ["Coins / left over"]
    K.hdr(cw, 5, hd, [24, 14] + [11] * ND + [12])
    W0 = 6; W1 = W0 + NE - 1
    for k in range(NE):
        r = W0 + k; er = E0 + k
        cw[f"A{r}"] = f'=IF(OR(Envelopes!A{er}="",Envelopes!B{er}<>"Cash"),"",Envelopes!A{er})'
        cw[f"B{r}"] = f'=IF(A{r}="",0,N(Envelopes!F{er}))'
        prev = f"B{r}"
        for i in range(ND):
            c = CL(3 + i); h = CL(13 + i); d = f"Settings!$F${D0 + i}"
            cw[f"{c}{r}"] = f'=IF(N({d})<=0,0,INT(ROUND({prev},2)/{d}))'
            cw[f"{h}{r}"] = f"=ROUND({prev}-{c}{r}*N({d}),2)"
            prev = f"{h}{r}"
        cw[f"{CL(3 + ND)}{r}"] = f"={prev}"
    for i in range(ND): cw.column_dimensions[CL(13 + i)].hidden = True
    style_range(cw, W0, W1, 1, 3 + ND, fmt={2: MONEY, **{3 + i: '0;-0;"-"' for i in range(ND)}, 3 + ND: MONEY})
    tot = W1 + 1
    cw[f"A{tot}"] = "Ask the bank for (per paycheck)"
    cw[f"B{tot}"] = f"=SUM(B{W0}:B{W1})"
    for i in range(ND + 1):
        c = CL(3 + i); cw[f"{c}{tot}"] = f"=SUM({c}{W0}:{c}{W1})"
    for c in range(1, 4 + ND):
        cell = cw.cell(row=tot, column=c); cell.font = font(bold=True, color=MAIN); cell.fill = fill(SOFT); cell.border = B
        cell.number_format = MONEY if c in (2, 3 + ND) else '0;-0;"-"'
    cw[f"A{tot + 1}"] = "Cash per month"; cw[f"B{tot + 1}"] = f"=B{tot}*N({PAY})"; cw[f"B{tot + 1}"].number_format = MONEY
    cw[f"A{tot + 1}"].font = font(bold=True); cw[f"B{tot + 1}"].font = font(bold=True)
    cw[f"A{tot + 2}"] = "Check: bills add up to"; cw[f"A{tot + 2}"].font = font(italic=True)
    cw[f"B{tot + 2}"] = "=" + "+".join(f"{CL(3 + i)}{tot}*N(Settings!$F${D0 + i})" for i in range(ND)) + f"+{CL(3 + ND)}{tot}"
    cw[f"B{tot + 2}"].number_format = MONEY
    cw.freeze_panes = f"C{W0}"

    # ---------- Year Summary ----------
    K.title(ys, "Year Summary", "Money spent from each envelope per month, compared with the budget.")
    K.hdr(ys, 5, ["Envelope"] + [""] * 12 + ["Spent (12 months)", "Budget (12 months)", "Difference"], [24] + [10] * 12 + [13, 13, 12])
    for j in range(12):
        c = ys.cell(row=5, column=2 + j, value=f"=EDATE(Settings!$C$8,{j})"); c.number_format = "mmm yy"
        c.font = Font(name=F, bold=True, color="FFFFFF")
    ys["A6"] = "Total spent per month"
    for j in range(15):
        Lc = CL(2 + j); ys[f"{Lc}6"] = f"=SUM({Lc}7:{Lc}{6 + NE})"
    for c in range(1, 17):
        cell = ys.cell(row=6, column=c); cell.font = Font(name=F, bold=True, color=MAIN); cell.fill = fill(SOFT); cell.border = B
        if c >= 2: cell.number_format = MONEY
    for k in range(NE):
        r = 7 + k; er = E0 + k
        ys[f"A{r}"] = f'=IF(Envelopes!$A${er}="","",Envelopes!$A${er})'
        for j in range(12):
            Lc = CL(2 + j)
            ys[f"{Lc}{r}"] = (f'=IF($A{r}="",0,SUMIFS({TR("D")},{TR("B")},$A{r},{TR("C")},"Spend",'
                              f'{TR("A")},">="&{Lc}$5,{TR("A")},"<="&EOMONTH({Lc}$5,0)))')
        ys[f"N{r}"] = f"=SUM(B{r}:M{r})"
        ys[f"O{r}"] = f"=N(Envelopes!$C${er})*12"
        ys[f"P{r}"] = f"=O{r}-N{r}"
    style_range(ys, 7, 6 + NE, 1, 16, fmt={c: MONEY for c in range(2, 17)})
    ys.freeze_panes = "B7"

    # ---------- sample ----------
    if sample:
        envs = [("Groceries", "Cash", 600, None, "Weekly shop + market"), ("Gas", "Cash", 200, None, ""),
                ("Dining out", "Cash", 150, None, "Incl. coffee"), ("Entertainment", "Cash", 80, None, ""),
                ("Personal care", "Cash", 60, None, "Haircuts, toiletries"), ("Clothing", "Cash", 75, None, ""),
                ("Household", "Cash", 70, None, "Cleaning, small repairs"), ("Gifts", "Cash", 50, None, ""),
                ("Kids", "Cash", 120, None, "School trips, activities"), ("Pets", "Cash", 60, None, "Food + treats"),
                ("Fun money", "Cash", 100, 50, "Spend on anything"), ("Rent", "Digital", 1400, None, "Auto-pay on the 1st"),
                ("Utilities", "Digital", 220, None, "Power, water, internet"), ("Phone", "Digital", 70, None, ""),
                ("Savings", "Digital", 300, None, "Transfer to savings account")]
        for k, row in enumerate(envs):
            for c, v in zip("ABCEG", row):
                if v not in (None, ""): en[f"{c}{E0 + k}"] = v
        rnd = random.Random(7)
        where = {"Groceries": ["FreshMart", "Farmers market", "Corner store"], "Gas": ["Fuel station"], "Dining out": ["Pizza night", "Cafe", "Burger place"],
                 "Entertainment": ["Movie tickets", "Bowling", "Streaming rental"], "Personal care": ["Haircut", "Pharmacy"],
                 "Clothing": ["Shoes", "T-shirts"], "Household": ["Cleaning supplies", "Hardware store"], "Gifts": ["Birthday gift"],
                 "Kids": ["School trip", "Swim class"], "Pets": ["Pet store"], "Fun money": ["Books", "Games", "Hobby supplies"]}
        share = {"Groceries": 0.97, "Gas": 0.93, "Dining out": 1.08, "Entertainment": 0.8, "Personal care": 0.7, "Clothing": 0.6,
                 "Household": 0.85, "Gifts": 0.5, "Kids": 0.9, "Pets": 0.95, "Fun money": 0.75}
        txs = []
        for mth in range(6, 12):
            last = 11 if mth == 11 else (31 if mth in (7, 8, 10) else 30)
            for name, typ, bud, _, _ in envs:
                if typ == "Digital":
                    txs.append((dt.date(2026, mth, 1), name, "Fill", bud, "From paycheck"))
                    if mth != 11 or name in ("Rent", "Savings"):
                        txs.append((dt.date(2026, mth, 1 if name in ("Rent", "Savings") else 18), name, "Spend", bud if name != "Utilities" else round(bud * rnd.uniform(0.85, 1.0), 2), "Auto-pay"))
                    continue
                for day in (1, 15):
                    if day <= last: txs.append((dt.date(2026, mth, day), name, "Fill", round(bud / 2, 2), "Payday cash"))
                target = bud * share[name] * rnd.uniform(0.85, 1.1) * (last / 30.0)
                n = max(1, round(target / (bud / 4 if bud >= 150 else bud / 2)))
                for i in range(n):
                    d = dt.date(2026, mth, min(last, 2 + int((last - 2) * (i + rnd.random()) / n)))
                    txs.append((d, name, "Spend", round(target / n * rnd.uniform(0.8, 1.2), 2), rnd.choice(where[name])))
        txs += [(dt.date(2026, 10, 20), "Fun money", "Move out", 30, "Top up dining out"), (dt.date(2026, 10, 20), "Dining out", "Move in", 30, "From fun money")]
        txs.sort(key=lambda t: (t[0], t[2] != "Fill"))
        for k, t in enumerate(txs):
            for c, v in zip("ABCDE", t): tr[f"{c}{T0 + k}"] = v
        s["C7"] = dt.date(2026, 6, 1)

    # ---------- Dashboard ----------
    K.title(dash, "Envelope Dashboard", "Pick a month in the yellow cell – everything below updates by itself.")
    dash["A3"] = "Month to show (blank = this month):"; dash["A3"].font = font(bold=True)
    input_cell(dash["C3"], dt.date(2026, 11, 1) if sample else None); dash["C3"].number_format = "mmm yyyy"
    dash["L1"] = '=IF(ISNUMBER(C3),DATE(YEAR(C3),MONTH(C3),1),DATE(YEAR(TODAY()),MONTH(TODAY()),1))'; dash["L1"].number_format = DATE
    dash["L1"].font = font(color="FFFFFF", size=8)
    dash["E3"] = '="Showing "&TEXT($L$1,"mmmm yyyy")&"  ·  amounts in "&Settings!C5'; dash["E3"].font = font(bold=True, color=MAIN)
    MS = "$L$1"; ME = "EOMONTH($L$1,0)"
    H = 14; R0 = H + 1; R1 = H + NE
    RR = lambda c: f"${c}${R0}:${c}${R1}"
    kp = [("Budget this month", f"=SUM({RR('C')})", MONEY), ("Filled this month", f"=SUM({RR('D')})", MONEY),
          ("Spent this month", f"=SUM({RR('E')})", MONEY), ("Left to spend (budget − spent)", f"=SUM({RR('F')})", MONEY),
          ("Money in all envelopes now", f"=SUM({RR('H')})", MONEY), ("Envelopes over budget", f'=COUNTIFS({RR("I")},"Over budget")', "0")]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 3) * 3, 1 + (i % 3) * 2, lab, f, fmt)
        dash.cell(row=6 + (i // 3) * 3, column=1 + (i % 3) * 2).alignment = Alignment(horizontal="left")
    for c, w in zip("ABCDEFGHI", [24, 10, 26, 13, 26, 13, 10, 15, 20]): dash.column_dimensions[c].width = w
    K.section(dash, f"A{H - 1}", "Envelopes this month")
    K.hdr(dash, H, ["Envelope", "Type", "Budget", "Filled", "Spent", "Left of budget", "% used", "Left in envelope", "Status"])
    rng = f'{TR("A")},">="&{MS},{TR("A")},"<="&{ME}'
    for k in range(NE):
        r = R0 + k; er = E0 + k
        dash[f"A{r}"] = f'=IF(Envelopes!A{er}="","",Envelopes!A{er})'
        dash[f"B{r}"] = f'=IF(A{r}="","",Envelopes!B{er})'
        dash[f"C{r}"] = f'=IF(A{r}="","",N(Envelopes!C{er}))'
        dash[f"D{r}"] = (f'=IF(A{r}="","",SUMIFS({TR("D")},{TR("B")},A{r},{TR("C")},"Fill",{rng})'
                         f'+SUMIFS({TR("D")},{TR("B")},A{r},{TR("C")},"Move in",{rng}))')
        dash[f"E{r}"] = f'=IF(A{r}="","",SUMIFS({TR("D")},{TR("B")},A{r},{TR("C")},"Spend",{rng}))'
        dash[f"F{r}"] = f'=IF(A{r}="","",C{r}-E{r})'
        dash[f"G{r}"] = f'=IF(OR(A{r}="",N(C{r})<=0),"",E{r}/C{r})'
        dash[f"H{r}"] = f'=IF(A{r}="","",SUMIFS({TR("F")},{TR("B")},A{r},{TR("A")},"<="&{ME}))'
        dash[f"I{r}"] = (f'=IF(A{r}="","",IF(H{r}<0,"Envelope below zero",IF(E{r}>C{r},"Over budget",IF(AND(C{r}>0,E{r}=C{r}),"Fully used",'
                         f'IF(AND(G{r}<>"",N(G{r})>=0.9),"Almost used up","OK")))))')
    style_range(dash, R0, R1, 1, 9, fmt={3: MONEY, 4: MONEY, 5: MONEY, 6: MONEY, 7: PCT, 8: MONEY})
    dash.conditional_formatting.add(f"G{R0}:G{R1}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="F59E0B"))
    dash.conditional_formatting.add(f"I{R0}:I{R1}", FormulaRule(formula=[f'OR(I{R0}="Over budget",I{R0}="Envelope below zero")'],
                                                               font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    dash.conditional_formatting.add(f"I{R0}:I{R1}", FormulaRule(formula=[f'I{R0}="Almost used up"'], font=Font(name=F, bold=True, color="B45309")))
    dash.conditional_formatting.add(f"I{R0}:I{R1}", FormulaRule(formula=[f'OR(I{R0}="OK",I{R0}="Fully used")'], font=Font(name=F, bold=True, color="166534")))
    dash.conditional_formatting.add(f"F{R0}:F{R1}", FormulaRule(formula=[f'AND(ISNUMBER(F{R0}),F{R0}<0)'], font=Font(name=F, bold=True, color="B91C1C")))
    ch = BarChart(); ch.type = "bar"; ch.title = "Budget vs spent"; ch.height = 9.5; ch.width = 15
    ch.add_data(Reference(dash, min_col=3, min_row=H, max_row=H + 15), titles_from_data=True)
    ch.add_data(Reference(dash, min_col=5, min_row=H, max_row=H + 15), titles_from_data=True)
    ch.set_categories(Reference(dash, min_col=1, min_row=R0, max_row=H + 15))
    ch.x_axis.scaling.orientation = "maxMin"; ch.y_axis.majorGridlines = None
    ch.series[0].graphicalProperties.solidFill = "FCD34D"; ch.series[1].graphicalProperties.solidFill = MAIN
    dash.add_chart(ch, "K3")
    ch2 = BarChart(); ch2.type = "col"; ch2.title = "Spent per month"; ch2.height = 7.5; ch2.width = 15; ch2.legend = None
    ch2.add_data(Reference(ys, min_col=2, max_col=13, min_row=6, max_row=6), from_rows=True, titles_from_data=False)
    ch2.set_categories(Reference(ys, min_col=2, max_col=13, min_row=5, max_row=5)); ch2.y_axis.majorGridlines = None
    ch2.series[0].graphicalProperties.solidFill = "F59E0B"; ch2.gapWidth = 60
    dash.add_chart(ch2, "K23")

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Settings tab: paychecks per month (1, 2 or 4.33 if weekly), currency label and the bills your bank gives out.", False),
             ("2.  Envelopes tab: one row per envelope – name, Cash or Digital, monthly budget. Budget per paycheck is calculated for you.", False),
             ("3.  Cash Withdrawal tab: how many 100s, 50s, 20s, 10s, 5s and 1s go in each envelope, and the total to ask the bank for on payday.", False),
             ("4.  Transactions tab: 'Fill' when you put money in an envelope, 'Spend' when you pay with it.", False),
             ("     Moving money between envelopes? Add two rows: 'Move out' from one and 'Move in' to the other.", False),
             ("5.  Dashboard: pick a month – budget, filled, spent, left to spend, what is still in each envelope and a status per envelope.", False),
             ("6.  Year Summary: spending per envelope per month, compared with 12 months of budget.", False),
             ("", False), ("Tips", True),
             ("•  Cash envelopes work best for groceries, gas, eating out and fun money – the spending that is easy to overdo with a card.", False),
             ("•  Money left in an envelope stays there (it rolls over). Use it next month or move it to Savings.", False),
             ("•  Prefer smaller bills? On the Settings tab set the 100 (or 50) to 0 and the planner uses 20s, 10s and 5s instead.", False),
             ("•  Not using real cash? Mark envelopes as Digital – it works the same, just without the bill breakdown.", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  30 envelopes, 1,500 transactions. No macros. Works in Excel 2010+ and Google Sheets.", False),
             ("•  To use Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("The envelopes and amounts are made-up examples (June – November 2026). Use the blank copy for your own budget.", False)]
    K.start(st, "Cash Envelope Budget", "Cash stuffing made simple – envelopes, spending log and bill breakdown – Excel and Google Sheets.", lines)
    finish(wb, sample, path)


META = {
    "name": "Cash Envelope Budget System (Cash Stuffing, Bill Breakdown, Envelope Tracker, Monthly Dashboard) – Excel & Google Sheets",
    "file_base": "Cash-Envelope-Budget",
    "price": 12,
    "color": "#B45309",
    "title": "Cash Envelope Budget",
    "tagline": "cash stuffing · bill breakdown for payday · envelope balances",
    "intro": "Run the cash envelope method without the guesswork. Set a monthly budget for each envelope – the planner splits it per paycheck, tells you exactly how many 100s, 50s, 20s, 10s, 5s and 1s to put in every envelope, tracks what you spend and shows what is left in each envelope, month by month.",
    "features": [
        "30 envelopes (cash or digital) with monthly budget and automatic budget per paycheck – monthly, twice a month or weekly pay",
        "Cash withdrawal planner: number of bills per envelope and the total to ask the bank for on payday (bill values are editable)",
        "Spending log with Fill, Spend and Move in / Move out between envelopes (1,500 rows, drop-down lists)",
        "Monthly dashboard: budget, filled, spent, left to spend and money still in every envelope",
        "Status per envelope: OK, Almost used up, Fully used, Over budget, Envelope below zero",
        "Year summary: spending per envelope per month vs 12 months of budget, with chart",
        "Blank copy plus a sample copy filled with six months of example data",
    ],
    "tags": ["cash envelopes", "cash stuffing", "envelope budget", "budget template", "budget spreadsheet", "google sheets",
             "excel template", "personal finance", "zero based budget", "paycheck budget", "budget planner", "money planner", "cash budget"],
    "summary": "Instant download: Excel + Google Sheets template, blank copy + sample copy.",
}
