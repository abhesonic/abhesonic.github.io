"""Bill Payment Tracker & Calendar – due dates, autopay, payment log, monthly calendar, 12-month paid/late grid.

Settings -> Bills -> Payments -> Bill Calendar / Year Overview -> Dashboard.
"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import BarChart, Reference
from _common import Kit, F, B, MONEY, DATE, font, fill, style_range, input_cell, finish

MAIN = "0E7490"; SOFT = "CFFAFE"
K = Kit(MAIN, SOFT)
NB = 60; B0 = 6; B1 = B0 + NB - 1
NP = 600; P0 = 6; P1 = P0 + NP - 1
FREQ = [("Monthly", 1), ("Every 2 months", 2), ("Quarterly", 3), ("Twice a year", 6), ("Yearly", 12), ("One time", 0)]
FQ0 = 6; FQ1 = FQ0 + len(FREQ) - 1
CATS = ["Housing", "Utilities", "Phone & internet", "Insurance", "Car & transport", "Loans & credit cards", "Subscriptions",
        "Kids & school", "Health", "Taxes & fees", "Savings transfer", "Other"]
NCAT = 15; CA0 = 6; CA1 = CA0 + NCAT - 1
ACCS = ["Checking", "Credit card", "Savings", "PayPal", "Cash"]
NA = 8; A0 = 6; A1 = A0 + NA - 1
Y0 = 7; Y1 = Y0 + NB - 1  # Year Overview rows


def due_parts(ms, sr):
    """(is-due condition, due date) for bill row sr in the month starting at cell/expr ms."""
    E = f"Bills!$E${sr}"; Q = f"Bills!$Q${sr}"; R = f"Bills!$R${sr}"
    diff = f"((YEAR({ms})-YEAR({E}))*12+MONTH({ms})-MONTH({E}))"
    cond = f"AND({R}=1,{diff}>=0,IF({Q}=0,{diff}=0,MOD({diff},MAX({Q},1))=0))"
    return cond, f"EDATE({E},{diff})"


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); bl = wb.create_sheet("Bills"); pay = wb.create_sheet("Payments")
    cal = wb.create_sheet("Bill Calendar"); yo = wb.create_sheet("Year Overview"); s = wb.create_sheet("Settings")

    # ---------- Settings ----------
    K.title(s, "Settings", "Yellow cells are yours to edit.")
    K.hdr(s, 5, ["How often", "Months between"])
    for i, (n, q) in enumerate(FREQ):
        s.cell(row=FQ0 + i, column=1, value=n); s.cell(row=FQ0 + i, column=2, value=q)
    style_range(s, FQ0, FQ1, 1, 2)
    s[f"A{FQ1 + 1}"] = "Fixed list – used by the formulas."; s[f"A{FQ1 + 1}"].font = font(italic=True, size=9)
    K.hdr(s, 5, ["Categories"], col0=6)
    for i in range(NCAT): input_cell(s[f"F{CA0 + i}"], CATS[i] if i < len(CATS) else None)
    K.hdr(s, 5, ["Pay from (accounts)"], col0=8)
    for i in range(NA): input_cell(s[f"H{A0 + i}"], ACCS[i] if i < len(ACCS) else None)
    SET = [(14, "Currency (label only)", "USD"), (15, "Warn me this many days before a due date", 5),
           (16, "Calendar month (blank = this month)", None), (18, "Year for the Year Overview (blank = this year)", None)]
    for r, lab, v in SET:
        s[f"A{r}"] = lab; s[f"A{r}"].font = font(bold=True); s[f"A{r}"].border = B; input_cell(s[f"C{r}"], v)
    s["C16"].number_format = "mmm yyyy"
    s["A17"] = "Calendar shows"; s["A17"].font = font(italic=True)
    s["C17"] = '=IF(ISNUMBER(C16),DATE(YEAR(C16),MONTH(C16),1),DATE(YEAR(TODAY()),MONTH(TODAY()),1))'; s["C17"].number_format = "mmm yyyy"
    s["A19"] = "Year Overview shows"; s["A19"].font = font(italic=True)
    s["C19"] = '=IF(ISNUMBER(C18),C18,YEAR(TODAY()))'; s["C19"].number_format = "0"
    for c, w in zip("ABCDEFGH", [22, 16, 16, 3, 3, 24, 3, 22]): s.column_dimensions[c].width = w
    WARN = "Settings!$C$15"; FRQ = f"Settings!$A${FQ0}:$A${FQ1}"; CALM = "Settings!$C$17"
    THISM = "DATE(YEAR(TODAY()),MONTH(TODAY()),1)"

    # ---------- Bills ----------
    K.title(bl, "Bills", "One row per bill. Yellow = you type, grey = automatic. 'First due date' sets the day of the month it is due.")
    cols = ["Bill", "Category", "Usual amount", "How often", "First due date", "Autopay?", "Pay from", "Status", "Website / account no. / notes",
            "Monthly average", "Yearly cost", "Next due", "Days until", "This month", "Alert"]
    K.hdr(bl, 5, cols, [24, 18, 12, 15, 13, 10, 14, 10, 30, 12, 12, 12, 9, 17, 24])
    style_range(bl, B0, B1, 1, 15, inputs=range(1, 10), fmt={3: MONEY, 5: DATE, 10: MONEY, 11: MONEY, 12: DATE, 13: "0"})
    for r in range(B0, B1 + 1):
        bl[f"Q{r}"] = f'=IFERROR(INDEX(Settings!$B${FQ0}:$B${FQ1},MATCH(D{r},{FRQ},0)),-1)'
        bl[f"R{r}"] = f'=IF(AND(A{r}<>"",ISNUMBER(C{r}),ISNUMBER(E{r}),Q{r}>=0,OR(H{r}="Active",H{r}="")),1,0)'
        c1, d1 = due_parts(CALM, r); bl[f"S{r}"] = f'=IF({c1},{d1},"")'
        c2, d2 = due_parts(THISM, r); bl[f"V{r}"] = f'=IF({c2},{d2},"")'
        bl[f"W{r}"] = f'=IF(A{r}="",0,SUMIFS(Payments!$C${P0}:$C${P1},Payments!$B${P0}:$B${P1},A{r},Payments!$H${P0}:$H${P1},{THISM}))'
        bl[f"J{r}"] = f'=IF(A{r}="","",IF(OR(R{r}<>1,Q{r}<=0),0,C{r}/Q{r}))'
        bl[f"K{r}"] = f'=IF(J{r}="","",J{r}*12)'
        m0 = f"((YEAR(TODAY())-YEAR(E{r}))*12+MONTH(TODAY())-MONTH(E{r}))"
        n = f"CEILING(MAX(0,{m0})/Q{r},1)*Q{r}"
        base = f"IF(EDATE(E{r},{n})>=TODAY(),EDATE(E{r},{n}),EDATE(E{r},{n}+Q{r}))"
        bl[f"L{r}"] = (f'=IF(R{r}<>1,"",IF(Q{r}=0,IF(AND(E{r}>=TODAY(),W{r}=0),E{r},""),'
                       f'IF(ISNUMBER(V{r}),IF(W{r}>0,EDATE(E{r},{m0}+Q{r}),IF(V{r}<TODAY(),V{r},{base})),{base})))')
        bl[f"M{r}"] = f'=IF(L{r}="","",L{r}-TODAY())'
        bl[f"N{r}"] = (f'=IF(R{r}<>1,IF(A{r}="","","Inactive"),IF(W{r}>0,"Paid",IF(NOT(ISNUMBER(V{r})),"Not due this month",'
                       f'IF(V{r}<TODAY(),"LATE",IF(F{r}="Yes","Autopay "&TEXT(V{r},"mmm d"),"Due "&TEXT(V{r},"mmm d"))))))')
        bl[f"O{r}"] = (f'=IF(A{r}="","",IF(N{r}="LATE","Overdue – pay now!",IF(AND(ISNUMBER(M{r}),F{r}<>"Yes"),'
                       f'IF(M{r}<={WARN},"Due in "&M{r}&" day(s)",""),"")))')
        bl[f"T{r}"] = f'=IF(ISNUMBER(S{r}),S{r}+ROW()/1000000,99999999)'
        bl[f"U{r}"] = f'=IF(ISNUMBER(L{r}),L{r}+ROW()/1000000,99999999)'
        bl[f"N{r}"].font = font(bold=True); bl[f"O{r}"].font = font(bold=True)
    for c in "QRSTUVW": bl.column_dimensions[c].hidden = True
    for col, lst in (("D", f"={FRQ}"), ("B", f"=Settings!$F${CA0}:$F${CA1}"), ("G", f"=Settings!$H${A0}:$H${A1}"),
                     ("F", '"Yes,No"'), ("H", '"Active,Paused,Ended"')):
        dv = DataValidation(type="list", formula1=lst, allow_blank=True); bl.add_data_validation(dv); dv.add(f"{col}{B0}:{col}{B1}")
    red = dict(font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2"))
    green = dict(font=Font(name=F, bold=True, color="15803D"), fill=fill("DCFCE7"))
    amber = dict(font=Font(name=F, bold=True, color="B45309"), fill=fill("FEF3C7"))
    bl.conditional_formatting.add(f"N{B0}:N{B1}", FormulaRule(formula=[f'N{B0}="LATE"'], **red))
    bl.conditional_formatting.add(f"N{B0}:N{B1}", FormulaRule(formula=[f'N{B0}="Paid"'], **green))
    bl.conditional_formatting.add(f"O{B0}:O{B1}", FormulaRule(formula=[f'LEFT(O{B0},7)="Overdue"'], **red))
    bl.conditional_formatting.add(f"O{B0}:O{B1}", FormulaRule(formula=[f'LEFT(O{B0},3)="Due"'], **amber))
    bl.conditional_formatting.add(f"A{B0}:O{B1}", FormulaRule(formula=[f'OR($H{B0}="Paused",$H{B0}="Ended")'], font=Font(name=F, color="9CA3AF")))
    bl.freeze_panes = f"B{B0}"; bl.auto_filter.ref = f"A5:O{B1}"
    BR = lambda c: f"Bills!${c}${B0}:${c}${B1}"

    # ---------- Payments ----------
    K.title(pay, "Payments", "Log every payment. 'For month' is optional – leave blank and the payment counts for the month you paid it.")
    K.hdr(pay, 5, ["Date paid", "Bill", "Amount paid", "For month (optional)", "Paid from", "Confirmation #", "Notes", "Counts for"],
          [12, 24, 12, 14, 14, 16, 28, 12])
    style_range(pay, P0, P1, 1, 8, inputs=range(1, 8), fmt={1: DATE, 3: MONEY, 4: "mmm yyyy", 8: "mmm yyyy"})
    for r in range(P0, P1 + 1):
        pay[f"H{r}"] = (f'=IF(B{r}="","",IF(ISNUMBER(D{r}),DATE(YEAR(D{r}),MONTH(D{r}),1),'
                        f'IF(ISNUMBER(A{r}),DATE(YEAR(A{r}),MONTH(A{r}),1),"")))')
    for col, lst in (("B", f"={BR('A')}"), ("E", f"=Settings!$H${A0}:$H${A1}")):
        dv = DataValidation(type="list", formula1=lst, allow_blank=True); pay.add_data_validation(dv); dv.add(f"{col}{P0}:{col}{P1}")
    pay.freeze_panes = f"A{P0}"; pay.auto_filter.ref = f"A5:H{P1}"
    PR = lambda c: f"Payments!${c}${P0}:${c}${P1}"

    # ---------- Year Overview ----------
    K.title(yo, "Year Overview", "Every bill, every month: ✓ Paid, Due, or LATE. Pick the year in Settings (blank = this year).")
    K.hdr(yo, 5, ["Bill", "Category", "Usual amount"] + [""] * 12, [24, 16, 11] + [10] * 12)
    for j in range(12):
        c = yo.cell(row=5, column=4 + j, value=f"=DATE(Settings!$C$19,{j + 1},1)"); c.number_format = "mmm yy"
        c.font = Font(name=F, bold=True, color="FFFFFF")
        yo.cell(row=5, column=17 + j, value=f"={CL(4 + j)}5").number_format = "mmm"
        yo.cell(row=5, column=29 + j, value=f"={CL(4 + j)}5").number_format = "mmm"
    yo["Q4"] = "Due amounts (helper)"; yo["AC4"] = "Paid amounts (helper)"
    for k in range(NB):
        r = Y0 + k; sr = B0 + k
        yo[f"A{r}"] = f'=IF(Bills!A{sr}="","",Bills!A{sr})'
        yo[f"B{r}"] = f'=IF(A{r}="","",Bills!B{sr})'
        yo[f"C{r}"] = f'=IF(A{r}="","",Bills!C{sr})'
        for j in range(12):
            ms = f"{CL(4 + j)}$5"; cond, dd = due_parts(ms, sr)
            DUE = f"{CL(17 + j)}{r}"; PD = f"{CL(29 + j)}{r}"
            yo[DUE] = f"=IF({cond},Bills!$C${sr},0)"
            yo[PD] = f'=IF(A{r}="",0,SUMIFS({PR("C")},{PR("B")},$A{r},{PR("H")},{ms}))'
            yo[f"{CL(4 + j)}{r}"] = (f'=IF(A{r}="","",IF({PD}>0,"✓ Paid",IF({DUE}>0,IF({dd}<TODAY(),"LATE","Due "&TEXT({dd},"d")),"")))')
    style_range(yo, Y0, Y1, 1, 15, fmt={3: MONEY})
    for r in range(Y0, Y1 + 1):
        for j in range(12): yo.cell(row=r, column=4 + j).alignment = Alignment(horizontal="center")
    for c in range(17, 41): yo.column_dimensions[CL(c)].hidden = True
    TR = Y1 + 2
    for i, (lab, src) in enumerate((("Total due", 17), ("Total paid", 29))):
        yo[f"A{TR + i}"] = lab
        for j in range(12):
            L = CL(src + j); yo.cell(row=TR + i, column=4 + j, value=f"=SUM({L}{Y0}:{L}{Y1})").number_format = '#,##0'
    yo[f"A{TR + 2}"] = "Late payments"
    for j in range(12):
        L = CL(4 + j); yo.cell(row=TR + 2, column=4 + j, value=f'=COUNTIF({L}{Y0}:{L}{Y1},"LATE")')
    for rr in range(TR, TR + 3):
        for c in range(1, 16):
            cell = yo.cell(row=rr, column=c); cell.font = Font(name=F, bold=True, color=MAIN); cell.fill = fill(SOFT); cell.border = B
            if c >= 4: cell.alignment = Alignment(horizontal="center")
    rng = f"D{Y0}:O{Y1}"
    yo.conditional_formatting.add(rng, FormulaRule(formula=[f'D{Y0}="LATE"'], **red))
    yo.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT(D{Y0},1)="✓"'], **green))
    yo.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT(D{Y0},3)="Due"'], **amber))
    yo.freeze_panes = f"D{Y0}"

    # ---------- Bill Calendar ----------
    K.title(cal, "Bill Calendar", "Bills due each day of the month. Change the month in Settings (blank = this month).")
    cal["A3"] = f'=TEXT({CALM},"mmmm yyyy")'; cal["A3"].font = Font(name=F, size=16, bold=True, color=MAIN)
    for i, d in enumerate(["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]):
        c = cal.cell(row=5, column=1 + i, value=d); c.font = Font(name=F, bold=True, color="FFFFFF"); c.fill = fill(MAIN)
        c.alignment = Alignment(horizontal="center"); c.border = B; cal.column_dimensions[CL(1 + i)].width = 17
    G0 = f"({CALM}-WEEKDAY({CALM},1)+1)"
    for w in range(6):
        r = 6 + w * 3
        for d in range(7):
            col = CL(1 + d); D = f"{col}{r}"
            cal[D] = f"={G0}+{w * 7 + d}"; cal[D].number_format = "d"
            cnt = f"COUNTIF({BR('S')},{D})"
            cal[f"{col}{r + 1}"] = (f'=IF(MONTH({D})<>MONTH({CALM}),"",IF({cnt}=0,"",TEXT(SUMIFS({BR("C")},{BR("S")},{D}),"#,##0")&" · "&{cnt}&" bill"&IF({cnt}>1,"s","")))')
            cal[f"{col}{r + 2}"] = (f'=IF(MONTH({D})<>MONTH({CALM}),"",IFERROR(INDEX({BR("A")},MATCH({D},{BR("S")},0))&IF({cnt}>1," +"&({cnt}-1),""),""))')
            for rr, sz, b in ((r, 12, True), (r + 1, 9, True), (r + 2, 9, False)):
                cell = cal[f"{col}{rr}"]; cell.font = Font(name=F, size=sz, bold=b, color="1F2A37")
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            for rr in (r, r + 1, r + 2):
                cell = cal[f"{col}{rr}"]; cell.fill = fill("FFFFFF")
            ln = Side(style="thin", color="D7DEE3")
            cal[f"{col}{r}"].border = Border(left=ln, right=ln, top=ln)
            cal[f"{col}{r + 1}"].border = Border(left=ln, right=ln)
            cal[f"{col}{r + 2}"].border = Border(left=ln, right=ln, bottom=ln)
        cal.row_dimensions[r + 2].height = 26
    GE = 6 + 18 - 1
    for w in range(6):
        r = 6 + w * 3
        cal.conditional_formatting.add(f"A{r}:G{r + 2}", FormulaRule(formula=[f'MONTH(A${r})<>MONTH({CALM})'],
                                                                    font=Font(name=F, color="C0C7CF"), fill=fill("F8FAFC")))
        cal.conditional_formatting.add(f"A{r}:G{r + 2}", FormulaRule(formula=[f'A${r}=TODAY()'], fill=fill("FEF3C7")))
        cal.conditional_formatting.add(f"A{r + 1}:G{r + 1}", FormulaRule(formula=[f'A{r + 1}<>""'], font=Font(name=F, size=9, bold=True, color=MAIN)))
    # list of bills due this month
    cal.column_dimensions["H"].width = 3
    K.section(cal, "I4", "Due this month")
    K.hdr(cal, 5, ["Bill", "Due", "Amount", "Autopay", "Status"], [22, 9, 11, 9, 11], col0=9)
    key = BR("T")
    for k in range(25):
        r = 6 + k; pos = f"MATCH(SMALL({key},{k + 1}),{key},0)"
        cal[f"I{r}"] = f'=IFERROR(IF(SMALL({key},{k + 1})>=99999999,"",INDEX({BR("A")},{pos})),"")'
        cal[f"J{r}"] = f'=IF(I{r}="","",INDEX({BR("S")},{pos}))'
        cal[f"K{r}"] = f'=IF(I{r}="","",INDEX({BR("C")},{pos}))'
        cal[f"L{r}"] = f'=IF(I{r}="","",INDEX({BR("F")},{pos}))'
        cal[f"M{r}"] = (f'=IF(I{r}="","",IF(SUMIFS({PR("C")},{PR("B")},I{r},{PR("H")},{CALM})>0,"Paid",IF(J{r}<TODAY(),"LATE","Unpaid")))')
    style_range(cal, 6, 30, 9, 13, fmt={10: "mmm d", 11: MONEY})
    cal.conditional_formatting.add("M6:M30", FormulaRule(formula=['M6="LATE"'], **red))
    cal.conditional_formatting.add("M6:M30", FormulaRule(formula=['M6="Paid"'], **green))
    for i, (lab, f) in enumerate((("Due this month", f"=SUM(K6:K30)"), ("Paid", f'=SUMIFS(K6:K30,M6:M30,"Paid")'),
                                  ("Still to pay", "=K32-K33"))):
        r = 32 + i; cal[f"I{r}"] = lab; cal[f"K{r}"] = f; cal[f"K{r}"].number_format = MONEY
        for c in "IJK":
            cal[f"{c}{r}"].font = Font(name=F, bold=True, color=MAIN); cal[f"{c}{r}"].fill = fill(SOFT)

    # ---------- sample ----------
    if sample:
        bills = [("Rent", "Housing", 1650, "Monthly", dt.date(2025, 1, 1), "No", "Checking", "Active", "Pay via landlord portal"),
                 ("Electricity", "Utilities", 118, "Monthly", dt.date(2025, 1, 18), "Yes", "Checking", "Active", "Account 4471-22"),
                 ("Water & sewer", "Utilities", 64, "Every 2 months", dt.date(2025, 2, 10), "No", "Checking", "Active", None),
                 ("Gas / heating", "Utilities", 72, "Monthly", dt.date(2025, 1, 22), "Yes", "Checking", "Active", None),
                 ("Internet", "Phone & internet", 65, "Monthly", dt.date(2025, 1, 12), "Yes", "Credit card", "Active", None),
                 ("Mobile phones", "Phone & internet", 95, "Monthly", dt.date(2025, 1, 25), "Yes", "Credit card", "Active", "Family plan, 3 lines"),
                 ("Car insurance", "Insurance", 612, "Twice a year", dt.date(2025, 3, 15), "No", "Checking", "Active", "Policy renews in March"),
                 ("Renters insurance", "Insurance", 168, "Yearly", dt.date(2025, 9, 5), "No", "Credit card", "Active", None),
                 ("Life insurance", "Insurance", 38, "Monthly", dt.date(2025, 1, 3), "Yes", "Checking", "Active", None),
                 ("Car loan", "Car & transport", 389, "Monthly", dt.date(2025, 1, 20), "No", "Checking", "Active", "Ends Jan 2028"),
                 ("Credit card (Visa)", "Loans & credit cards", 300, "Monthly", dt.date(2025, 1, 14), "No", "Checking", "Active", "Pay full statement"),
                 ("Student loan", "Loans & credit cards", 210, "Monthly", dt.date(2025, 1, 28), "Yes", "Checking", "Active", None),
                 ("Streaming bundle", "Subscriptions", 23, "Monthly", dt.date(2025, 1, 9), "Yes", "Credit card", "Active", None),
                 ("Gym", "Health", 35, "Monthly", dt.date(2025, 1, 5), "Yes", "Credit card", "Active", None),
                 ("Daycare", "Kids & school", 920, "Monthly", dt.date(2025, 1, 1), "No", "Checking", "Active", None),
                 ("Car registration", "Taxes & fees", 145, "Yearly", dt.date(2025, 6, 30), "No", "Credit card", "Active", None),
                 ("Property tax (storage unit)", "Taxes & fees", 240, "Quarterly", dt.date(2025, 1, 31), "No", "Checking", "Active", None),
                 ("Emergency fund transfer", "Savings transfer", 200, "Monthly", dt.date(2025, 1, 2), "Yes", "Savings", "Active", None),
                 ("Dentist payment plan", "Health", 90, "Monthly", dt.date(2026, 7, 15), "No", "Credit card", "Ended", "Paid off"),
                 ("Holiday trip deposit", "Other", 500, "One time", dt.date(2026, 11, 20), "No", "Credit card", "Active", None)]
        for k, row in enumerate(bills):
            for c, v in zip("ABCDEFGHI", row):
                if v is not None: bl[f"{c}{B0 + k}"] = v
        today = dt.date(2026, 11, 9); var = {"Electricity": [1.32, 1.18, 1.0, .86, .8, .96, 1.24, 1.31, 1.08, .9, .94],
                                             "Gas / heating": [1.6, 1.5, 1.2, .8, .5, .4, .35, .35, .5, .9, 1.2],
                                             "Credit card (Visa)": [1.4, .8, 1.1, .9, 1.25, 1.0, .7, 1.3, .95, 1.15, 1.05]}
        rows = []
        for (name, cat, amt, freq, first, auto, acc, status, _n) in bills:
            q = dict(FREQ)[freq]
            for m in range(1, 12):
                if status == "Ended" and m > 10: continue
                diff = (2026 - first.year) * 12 + m - first.month
                if diff < 0 or (q == 0 and diff != 0) or (q and diff % q): continue
                day = min(first.day, 28 if m == 2 else 30 if m in (4, 6, 9, 11) else 31); due = dt.date(2026, m, day)
                if name == "Dentist payment plan" and m < 7: continue
                if m == 11 and (due > today or name in ("Credit card (Visa)", "Daycare")):  # unpaid / late this month
                    continue
                if name == "Water & sewer" and m == 8: continue  # missed payment -> LATE in Year Overview
                paid_on = due if auto == "Yes" else max(dt.date(2026, m, 1), due - dt.timedelta(days=2 if m % 3 else 0))
                if name == "Rent" and m == 6: paid_on = dt.date(2026, 6, 4)
                a = round(amt * var.get(name, [1] * 11)[m - 1], 2)
                rows.append((paid_on, name, a, None, acc, f"CF{m:02d}{len(rows) + 101}" if auto == "No" else None,
                             "Late fee waived" if name == "Rent" and m == 6 else None))
        # one late-paid bill counted for the month it belonged to
        rows.append((dt.date(2026, 11, 3), "Car loan", 389, dt.date(2026, 10, 1), "Checking", "CL-7781", "Paid late for October"))
        rows = [r for r in rows if not (r[1] == "Car loan" and r[0].month == 10)]
        rows.sort(key=lambda x: x[0])
        for k, row in enumerate(rows):
            for c, v in zip("ABCDEFG", row):
                if v is not None: pay[f"{c}{P0 + k}"] = v

    # ---------- Dashboard ----------
    K.title(dash, "Bill Dashboard", "Updates by itself from the Bills and Payments tabs.")
    dash["A3"] = '="Amounts in "&Settings!C14&"  ·  today: "&TEXT(TODAY(),"mmm d, yyyy")'; dash["A3"].font = font(bold=True)
    kp = [("Due this month", f'=SUMIFS({BR("C")},{BR("V")},">0")', MONEY),
          ("Paid this month", f'=SUMIFS({PR("C")},{PR("H")},{THISM})', MONEY),
          ("Still to pay this month", f'=SUMIFS({BR("C")},{BR("V")},">0",{BR("W")},0)', MONEY),
          ("Overdue bills", f'=COUNTIF({BR("N")},"LATE")', "0"),
          ("Due in the next 7 days", f'=SUMIFS({BR("C")},{BR("M")},">=0",{BR("M")},"<=7")', MONEY),
          ("Average bills per month", f"=SUM({BR('J')})", MONEY)]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 3) * 3, 1 + (i % 3) * 2, lab, f, fmt)
        dash.cell(row=6 + (i // 3) * 3, column=1 + (i % 3) * 2).alignment = Alignment(horizontal="left")
    dash.conditional_formatting.add("A9", FormulaRule(formula=["A9>0"], font=Font(name=F, size=18, bold=True, color="B91C1C")))
    for c, w in zip("ABCDEFGHIJ", [24, 13, 24, 13, 24, 13, 3, 24, 12, 11]): dash.column_dimensions[c].width = w
    K.section(dash, "A11", "Bills by category")
    K.hdr(dash, 12, ["Category", "Per month", "Per year", "Bills"])
    for i in range(NCAT):
        r = 13 + i
        dash[f"A{r}"] = f'=IF(Settings!F{CA0 + i}="","",Settings!F{CA0 + i})'
        dash[f"B{r}"] = f'=IF(A{r}="","",SUMIFS({BR("J")},{BR("B")},A{r}))'
        dash[f"C{r}"] = f'=IF(A{r}="","",B{r}*12)'
        dash[f"D{r}"] = f'=IF(A{r}="","",COUNTIFS({BR("B")},A{r},{BR("R")},1))'
    CE = 12 + NCAT
    style_range(dash, 13, CE, 1, 4, fmt={2: MONEY, 3: MONEY, 4: "0"})
    K.section(dash, "F11", "Coming up next")
    K.hdr(dash, 12, ["Bill", "Due", "In days", "Amount", "Autopay"], col0=6)
    dash.column_dimensions["F"].width = 24; dash.column_dimensions["G"].width = 10; dash.column_dimensions["H"].width = 9
    key = BR("U")
    for k in range(NCAT):
        r = 13 + k; pos = f"MATCH(SMALL({key},{k + 1}),{key},0)"
        dash[f"F{r}"] = f'=IFERROR(IF(SMALL({key},{k + 1})>=99999999,"",INDEX({BR("A")},{pos})),"")'
        for col, src in zip("GHIJ", "LMCF"):
            dash[f"{col}{r}"] = f'=IF(F{r}="","",INDEX({BR(src)},{pos}))'
    style_range(dash, 13, CE, 6, 10, fmt={7: "mmm d", 8: "0", 9: MONEY})
    dash.conditional_formatting.add(f"H13:H{CE}", FormulaRule(formula=[f"AND(ISNUMBER(H13),H13<={WARN})"], font=Font(name=F, bold=True, color="B91C1C")))
    FR = CE + 3
    K.section(dash, f"A{FR}", "This year – due vs paid")
    K.hdr(dash, FR + 1, ["Month", "Due", "Paid", "Late"])
    for j in range(12):
        r = FR + 2 + j; L = CL(4 + j)
        dash[f"A{r}"] = f"='Year Overview'!{L}5"; dash[f"B{r}"] = f"='Year Overview'!{L}{TR}"
        dash[f"C{r}"] = f"='Year Overview'!{L}{TR + 1}"; dash[f"D{r}"] = f"='Year Overview'!{L}{TR + 2}"
    style_range(dash, FR + 2, FR + 13, 1, 4, fmt={1: "mmm yyyy", 2: MONEY, 3: MONEY, 4: "0"})
    dash[f"A{FR + 14}"] = "Total"
    for c in "BCD": dash[f"{c}{FR + 14}"] = f"=SUM({c}{FR + 2}:{c}{FR + 13})"; dash[f"{c}{FR + 14}"].number_format = MONEY if c != "D" else "0"
    for c in "ABCD": dash[f"{c}{FR + 14}"].font = font(bold=True)
    K.section(dash, f"F{FR}", "Alerts")
    K.hdr(dash, FR + 1, ["Bill", "Alert"], col0=6)
    for r in range(B0, B1 + 1): bl[f"X{r}"] = f'=IF(O{r}="",99999999,ROW())'
    bl.column_dimensions["X"].hidden = True
    akey = BR("X")
    for k in range(12):
        r = FR + 2 + k; pos = f"MATCH(SMALL({akey},{k + 1}),{akey},0)"
        dash[f"F{r}"] = f'=IFERROR(IF(SMALL({akey},{k + 1})>=99999999,"",INDEX({BR("A")},{pos})),"")'
        dash[f"G{r}"] = f'=IF(F{r}="","",INDEX({BR("O")},{pos}))'
        dash.merge_cells(f"G{r}:J{r}")
    style_range(dash, FR + 2, FR + 13, 6, 10)
    dash.conditional_formatting.add(f"G{FR + 2}:G{FR + 13}", FormulaRule(formula=[f'LEFT(G{FR + 2},7)="Overdue"'], **red))
    ch = BarChart(); ch.type = "bar"; ch.title = "Monthly bills by category"; ch.height = 8; ch.width = 14; ch.legend = None
    ch.add_data(Reference(dash, min_col=2, min_row=12, max_row=12 + len(CATS)), titles_from_data=True)
    ch.set_categories(Reference(dash, min_col=1, min_row=13, max_row=12 + len(CATS)))
    ch.x_axis.scaling.orientation = "maxMin"; ch.y_axis.majorGridlines = None; ch.series[0].graphicalProperties.solidFill = MAIN
    dash.add_chart(ch, "L5")
    ch2 = BarChart(); ch2.type = "col"; ch2.title = "Due vs paid by month"; ch2.height = 7.5; ch2.width = 14
    ch2.add_data(Reference(dash, min_col=2, max_col=3, min_row=FR + 1, max_row=FR + 13), titles_from_data=True)
    ch2.set_categories(Reference(dash, min_col=1, min_row=FR + 2, max_row=FR + 13)); ch2.y_axis.majorGridlines = None
    ch2.series[0].graphicalProperties.solidFill = "94A3B8"; ch2.series[1].graphicalProperties.solidFill = MAIN; ch2.gapWidth = 60
    dash.add_chart(ch2, "L22")

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Settings tab: check the categories and type the accounts you pay from. Optional: currency label and how many days' warning you want.", False),
             ("2.  Bills tab: one row per bill – name, category, usual amount, how often (monthly, every 2 months, quarterly, twice a year, yearly, one time)", False),
             ("     and the first due date (any past due date works – it sets the day of the month). Mark Autopay Yes or No.", False),
             ("3.  The next due date, days until it, monthly average, yearly cost, this month's status and alerts are calculated for you.", False),
             ("4.  Payments tab: each time you pay, add a row – date, bill (pick from the list) and amount. The bill turns green 'Paid'.", False),
             ("     Paid a bill late, in the next month? Type the month it belonged to in 'For month'.", False),
             ("5.  Bill Calendar: a month view with what is due each day plus a checklist of this month's bills (Paid / Unpaid / LATE).", False),
             ("6.  Year Overview: every bill and every month of the year at a glance – ✓ Paid, Due, or LATE in red.", False),
             ("7.  Dashboard: due, paid and still to pay this month, overdue bills, next 7 days, bills by category, upcoming bills and alerts.", False),
             ("", False), ("Tips", True),
             ("•  Variable bills (electricity, credit card): type your usual amount; log what you really paid in Payments.", False),
             ("•  Set Status to 'Paused' or 'Ended' when a bill stops – it stays in your history but no longer counts.", False),
             ("•  Print the Bill Calendar each month and stick it on the fridge.", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  60 bills, 600 payments, 15 categories, 8 accounts. No macros. Works in Excel 2010+ and Google Sheets.", False),
             ("•  Dates update with today's date every time you open the file.", False),
             ("•  To use Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("The bills and amounts are made-up examples for 2026. Use the blank copy in your download for your own bills.", False)]
    K.start(st, "Bill Payment Tracker & Calendar", "Never miss a due date again – Excel and Google Sheets.", lines)
    finish(wb, sample, path)


META = {
    "name": "Bill Payment Tracker & Calendar (Due Dates, Autopay, Late Alerts, Monthly Calendar, Yearly Paid Checklist) – Excel & Google Sheets",
    "file_base": "Bill-Payment-Tracker-Calendar",
    "price": 12,
    "color": "#0E7490",
    "title": "Bill Payment Tracker",
    "tagline": "due dates · monthly bill calendar · paid / late checklist · alerts",
    "intro": "Never miss a bill or pay a late fee again. List your bills once – amount, how often and the due date – and the tracker shows what is due this month, what you already paid and what is late. Log a payment in one row and the bill turns green. A printable monthly bill calendar and a 12-month paid/late checklist give you the full picture at a glance.",
    "features": [
        "Track 60 bills: category, usual amount, how often (monthly to yearly or one time), due date, autopay, pay-from account and notes",
        "Automatic next due date, days until it, monthly average and yearly cost",
        "This month's status for every bill: Paid, Due, Autopay or LATE – plus alerts a few days before the due date",
        "Payment log for 600 payments with confirmation numbers; late payments can be counted for the month they belong to",
        "Monthly bill calendar showing what is due each day, with a checklist of this month's bills",
        "Year overview: every bill × every month with ✓ Paid, Due or LATE colours",
        "Dashboard: due, paid and still to pay this month, overdue bills, next 7 days, bills by category, charts",
        "Blank copy plus a sample copy filled with example bills and payments",
    ],
    "tags": ["bill tracker", "bill payment tracker", "bill calendar", "budget template", "google sheets", "excel template",
             "monthly bills", "bill organizer", "personal finance", "due date tracker", "household budget", "bill checklist", "expense tracker"],
    "summary": "Instant download: Excel + Google Sheets template, blank copy + sample copy.",
}
