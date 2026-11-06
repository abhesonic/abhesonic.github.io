"""Payroll & Timesheet Calculator – bi-weekly timesheet with clock in/out, daily + weekly overtime, PTO, gross-to-net and pay stubs.

Settings (business, period start, overtime rules, employer tax %) -> Employees (rate, withholding %, deductions)
-> Timesheet (one row per person per day: in, out, unpaid break, PTO, tips) -> Payroll (hours, OT, gross, deductions, net)
-> Pay Stub (pick a name, print) -> Dashboard (totals, overtime cost, employer cost, charts).
"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import BarChart, Reference
from _common import Kit, F, B, MONEY, DATE, font, fill, style_range, input_cell, finish

MAIN = "0F766E"; SOFT = "CCFBF1"
K = Kit(MAIN, SOFT)
HRS = '0.00;-0.00;"-"'; TIME = "hh:mm"; PCT = "0.0%"
NE = 30; E0 = 6; E1 = E0 + NE - 1            # Employees rows
NT = 600; T0 = 6; T1 = T0 + NT - 1           # Timesheet rows
P0 = 7; P1 = P0 + NE - 1                     # Payroll rows

EMP = [("Daniel Reyes", "Crew lead", 24.00, 0.12, 40, 0), ("Hannah Cole", "Technician", 21.50, 0.10, 25, 0),
       ("Marcus Hill", "Technician", 20.00, 0.10, 0, 15), ("Julia Novak", "Office admin", 19.00, 0.09, 25, 0),
       ("Kevin Osei", "Apprentice", 17.00, 0.08, 0, 0), ("Rosa Mendez", "Technician", 21.00, 0.10, 25, 0)]
# per employee: list of (in, out, break) per working day pattern, 14 days
PAT = {0: "7:00-17:30/30 7:00-16:30/30 7:00-17:00/30 7:00-16:00/30 7:00-15:30/30 - - 7:00-17:00/30 7:00-16:30/30 7:00-18:00/30 7:00-16:00/30 7:00-14:00/30 8:00-12:00/0 -",
       1: "8:00-16:30/30 8:00-16:30/30 8:00-16:30/30 8:00-16:30/30 8:00-16:30/30 - - 8:00-16:30/30 8:00-16:30/30 8:00-16:30/30 8:00-16:30/30 8:00-16:30/30 - -",
       2: "7:30-18:30/60 7:30-17:30/30 7:30-18:00/30 7:30-16:00/30 7:30-16:00/30 8:00-13:00/0 - 7:30-16:00/30 7:30-16:00/30 7:30-19:00/30 7:30-16:00/30 7:30-16:00/30 - -",
       3: "9:00-15:00/30 9:00-15:00/30 9:00-15:00/30 9:00-15:00/30 - - - 9:00-15:00/30 9:00-15:00/30 9:00-15:00/30 9:00-15:00/30 - - -",
       4: "8:00-16:30/30 8:00-16:30/30 8:00-16:30/30 8:00-16:30/30 8:00-16:30/30 - - 8:00-16:30/30 8:00-16:30/30 PTO 8:00-16:30/30 8:00-16:30/30 - -",
       5: "- 8:00-17:00/30 8:00-17:00/30 8:00-17:00/30 8:00-17:00/30 8:00-14:00/30 - - 8:00-17:00/30 8:00-17:00/30 8:00-17:00/30 8:00-17:00/30 9:00-13:00/0 -"}
TIPS = {(1, 4): 40, (5, 12): 25, (2, 9): 60}


def tm(s):
    h, m = s.split(":"); return dt.time(int(h), int(m))


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); s = wb.create_sheet("Settings"); em = wb.create_sheet("Employees")
    ts = wb.create_sheet("Timesheet"); pr = wb.create_sheet("Payroll"); ps = wb.create_sheet("Pay Stub")
    p0 = dt.date(2026, 11, 9)

    # ---------- Settings ----------
    K.title(s, "Settings", "Yellow cells are yours to edit.")
    rows = [("Business name", "Brightline Plumbing & Heating" if sample else "My Business", None,
             None),
            ("Pay period start date (day 1)", p0 if sample else "=TODAY()-WEEKDAY(TODAY(),3)", DATE,
             "The pay period is 14 days (2 weeks) from this date. Change it every pay period."),
            ("Currency label", "USD", None, None),
            ("Weekly overtime after (hours)", 40, "0.0", "US federal rule (FLSA): overtime after 40 hours in a work week."),
            ("Daily overtime after (hours, 0 = off)", 0, "0.0", "Some states (e.g. California) also pay overtime after 8 hours in a day. 0 = no daily overtime."),
            ("Overtime pay multiplier", 1.5, "0.00", "1.5 = time and a half."),
            ("Employer payroll tax % (estimate)", 0.0765, PCT, "7.65% = US Social Security + Medicare employer share. Add state/federal unemployment if you want."),
            ("Pay date", "=B5+17", DATE, "Formula: 4 days after the period ends. Overwrite with your own date if you like."),
            ("Pay period end (calculated)", "=B5+13", DATE, None),
            ("Week 2 starts (calculated)", "=B5+7", DATE, None)]
    for i, (k_, v, fmt, note) in enumerate(rows, 4):
        s.cell(row=i, column=1, value=k_).font = font(bold=True)
        c = s.cell(row=i, column=2)
        if i >= 12:
            c.value = v; c.border = B; c.fill = fill("F3F4F6"); c.font = font()
        else:
            input_cell(c, v)
        if fmt: c.number_format = fmt
        if note: s[f"C{i}"] = note; s[f"C{i}"].font = font(italic=True, size=9)
    s["A16"] = "Overtime rules differ by country and state. These settings are a calculator, not legal or tax advice."
    s["A16"].font = font(italic=True, size=9)
    for c, w in zip("ABC", [38, 30, 90]): s.column_dimensions[c].width = w

    # ---------- Employees ----------
    K.title(em, "Employees", "One row per person. Withholding % is your own estimate of income tax withheld (check your local tax tables).")
    K.hdr(em, 5, ["Name", "Job title", "Hourly rate", "Tax withholding %", "Pre-tax deduction per period (401k, health…)",
                  "After-tax deduction per period", "Employee ID", "Email", "Notes"], [22, 16, 12, 13, 18, 16, 12, 28, 26])
    style_range(em, E0, E1, 1, 9, inputs=tuple(range(1, 10)), fmt={3: MONEY, 4: PCT, 5: MONEY, 6: MONEY})
    if sample:
        for i, (n, job, rate, wh, pre, post) in enumerate(EMP):
            r = E0 + i; em[f"A{r}"] = n; em[f"B{r}"] = job; em[f"C{r}"] = rate; em[f"D{r}"] = wh; em[f"E{r}"] = pre; em[f"F{r}"] = post
            em[f"G{r}"] = f"E-{101 + i}"; em[f"H{r}"] = n.split()[0].lower() + "@brightline.example"
    em.freeze_panes = f"B{E0}"
    EN = f"Employees!$A${E0}:$A${E1}"

    # ---------- Timesheet ----------
    K.title(ts, "Timesheet", "One row per person per day. Type clock-in and clock-out (e.g. 7:30 and 16:00). Overnight shifts work.")
    heads = ["Date", "Employee", "Clock in", "Clock out", "Unpaid break (min)", "PTO / holiday hours", "Tips / bonus", "Notes",
             "Hours worked", "Daily OT hours", "Week", "In this period?", "Check"]
    K.hdr(ts, 5, heads, [12, 22, 10, 10, 11, 11, 11, 24, 10, 10, 7, 10, 26])
    style_range(ts, T0, T1, 1, 13, inputs=(1, 2, 3, 4, 5, 6, 7, 8), fmt={1: DATE, 3: TIME, 4: TIME, 5: "0", 6: HRS, 7: MONEY, 9: HRS, 10: HRS, 11: "0"})
    for r in range(T0, T1 + 1):
        ts[f"I{r}"] = f'=IF(OR(C{r}="",D{r}=""),0,MAX(0,ROUND(MOD(D{r}-C{r},1)*24-N(E{r})/60,2)))'
        ts[f"J{r}"] = f'=IF(Settings!$B$8>0,MAX(0,I{r}-Settings!$B$8),0)'
        ts[f"K{r}"] = f'=IF(A{r}="","",IF(A{r}<Settings!$B$5,"",IF(A{r}<Settings!$B$13,1,IF(A{r}<=Settings!$B$12,2,""))))'
        ts[f"L{r}"] = f'=IF(A{r}="","",IF(K{r}="","No – outside","Yes"))'
        ts[f"M{r}"] = (f'=IF(A{r}="","",TRIM(IF(B{r}="","No employee  ",IF(ISNA(MATCH(B{r},{EN},0)),"Not on Employees list  ",""))'
                       f'&IF(AND(C{r}<>"",D{r}=""),"Missing clock-out  ","")&IF(I{r}>16,"Over 16 h – check times  ","")'
                       f'&IF(COUNTIFS($A${T0}:$A${T1},A{r},$B${T0}:$B${T1},B{r})>1,"Duplicate day","")))')
    if sample:
        r = T0
        for i in range(len(EMP)):
            days = PAT[i].split()
            for d, cell in enumerate(days):
                if cell == "-": continue
                ts[f"A{r}"] = p0 + dt.timedelta(days=d); ts[f"B{r}"] = EMP[i][0]
                if cell == "PTO":
                    ts[f"F{r}"] = 8; ts[f"H{r}"] = "Paid time off"
                else:
                    t, b = cell.split("/"); a, z = t.split("-")
                    ts[f"C{r}"] = tm(a); ts[f"D{r}"] = tm(z); ts[f"E{r}"] = int(b)
                if (i, d) in TIPS: ts[f"G{r}"] = TIPS[(i, d)]; ts[f"H{r}"] = "Customer tip / job bonus"
                r += 1
    dv = DataValidation(type="list", formula1=f"={EN}", allow_blank=True); ts.add_data_validation(dv); dv.add(f"B{T0}:B{T1}")
    ts.conditional_formatting.add(f"M{T0}:M{T1}", FormulaRule(formula=[f'M{T0}<>""'], font=Font(name=F, bold=True, color="B91C1C")))
    ts.conditional_formatting.add(f"L{T0}:L{T1}", FormulaRule(formula=[f'LEFT(L{T0},2)="No"'], font=Font(name=F, color="9CA3AF")))
    ts.conditional_formatting.add(f"J{T0}:J{T1}", FormulaRule(formula=[f'J{T0}>0'], font=Font(name=F, bold=True, color="B45309")))
    ts.freeze_panes = f"C{T0}"

    # ---------- Payroll ----------
    K.title(pr, "Payroll", "Calculated for the current pay period from the Timesheet and Employees tabs – nothing to type here.")
    pr["A3"] = '=Settings!B4&"  ·  "&TEXT(Settings!B5,"d mmm yyyy")&" – "&TEXT(Settings!B12,"d mmm yyyy")&"  ·  pay date "&TEXT(Settings!B11,"d mmm yyyy")'
    pr["A3"].font = font(bold=True)
    ph = ["Employee", "Hourly rate", "Week 1 hours", "Week 2 hours", "Daily OT hours", "Weekly OT hours", "Regular hours", "Overtime hours",
          "PTO hours", "Regular pay", "Overtime pay", "PTO pay", "Tips / bonus", "Gross pay", "Pre-tax deductions", "Taxable pay",
          "Tax withheld (est.)", "After-tax deductions", "Net pay", "Employer tax (est.)", "Total cost to you"]
    K.hdr(pr, 6, ph, [22] + [11] * 20)
    style_range(pr, P0, P1, 1, 21, fmt={2: MONEY, 3: HRS, 4: HRS, 5: HRS, 6: HRS, 7: HRS, 8: HRS, 9: HRS,
                                        **{c: MONEY for c in range(10, 22)}})
    TA = f"Timesheet!$B${T0}:$B${T1}"; TK = f"Timesheet!$K${T0}:$K${T1}"
    TI = f"Timesheet!$I${T0}:$I${T1}"; TJ = f"Timesheet!$J${T0}:$J${T1}"; TF = f"Timesheet!$F${T0}:$F${T1}"; TG = f"Timesheet!$G${T0}:$G${T1}"
    LK = lambda col, r: f"IFERROR(INDEX(Employees!${col}${E0}:${col}${E1},MATCH($A{r},{EN},0)),0)"
    for i in range(NE):
        r = P0 + i; e = E0 + i
        pr[f"A{r}"] = f'=IF(Employees!A{e}="","",Employees!A{e})'
        g = f'IF($A{r}="",0,'
        pr[f"B{r}"] = f"={g}N(Employees!C{e}))"
        pr[f"C{r}"] = f"={g}SUMIFS({TI},{TA},$A{r},{TK},1))"
        pr[f"D{r}"] = f"={g}SUMIFS({TI},{TA},$A{r},{TK},2))"
        pr[f"E{r}"] = f"={g}SUMIFS({TJ},{TA},$A{r},{TK},1)+SUMIFS({TJ},{TA},$A{r},{TK},2))"
        pr[f"F{r}"] = (f"={g}MAX(0,C{r}-SUMIFS({TJ},{TA},$A{r},{TK},1)-Settings!$B$7)"
                       f"+MAX(0,D{r}-SUMIFS({TJ},{TA},$A{r},{TK},2)-Settings!$B$7))")
        pr[f"G{r}"] = f"=C{r}+D{r}-H{r}"
        pr[f"H{r}"] = f"=E{r}+F{r}"
        pr[f"I{r}"] = f'={g}SUMIFS({TF},{TA},$A{r},{TK},1)+SUMIFS({TF},{TA},$A{r},{TK},2))'
        pr[f"J{r}"] = f"=G{r}*B{r}"
        pr[f"K{r}"] = f"=H{r}*B{r}*Settings!$B$9"
        pr[f"L{r}"] = f"=I{r}*B{r}"
        pr[f"M{r}"] = f'={g}SUMIFS({TG},{TA},$A{r},{TK},1)+SUMIFS({TG},{TA},$A{r},{TK},2))'
        pr[f"N{r}"] = f"=J{r}+K{r}+L{r}+M{r}"
        pr[f"O{r}"] = f"=IF(N{r}=0,0,MIN(N{r},N(Employees!E{e})))"
        pr[f"P{r}"] = f"=N{r}-O{r}"
        pr[f"Q{r}"] = f"=ROUND(P{r}*N(Employees!D{e}),2)"
        pr[f"R{r}"] = f"=IF(N{r}=0,0,MIN(P{r}-Q{r},N(Employees!F{e})))"
        pr[f"S{r}"] = f"=N{r}-O{r}-Q{r}-R{r}"
        pr[f"T{r}"] = f"=ROUND(N{r}*Settings!$B$10,2)"
        pr[f"U{r}"] = f"=N{r}+T{r}"
    TR = P1 + 1
    pr[f"A{TR}"] = "TOTAL"
    for c in range(1, 22):
        cell = pr.cell(row=TR, column=c); cell.font = Font(name=F, bold=True, color="FFFFFF"); cell.fill = fill(MAIN); cell.border = B
        if c >= 3:
            col = cell.column_letter; cell.value = f"=SUM({col}{P0}:{col}{P1})"; cell.number_format = HRS if c <= 9 else MONEY
    pr[f"A{TR + 2}"] = ("Overtime: daily OT (if switched on) first, then weekly hours above the limit for each week separately. "
                        "Tax withheld uses the % on the Employees tab – an estimate, not an official tax calculation.")
    pr[f"A{TR + 2}"].font = font(italic=True, size=9)
    pr.conditional_formatting.add(f"H{P0}:H{P1}", FormulaRule(formula=[f'H{P0}>0'], font=Font(name=F, bold=True, color="B45309")))
    pr.freeze_panes = f"B{P0}"

    # ---------- Pay Stub ----------
    K.title(ps, "Pay Stub", "Pick an employee in the yellow cell, then print this tab (File > Print) or save it as PDF.")
    ps["A4"] = "Employee"; ps["A4"].font = font(bold=True)
    input_cell(ps["B4"], EMP[0][0] if sample else None)
    dv = DataValidation(type="list", formula1=f"={EN}", allow_blank=True); ps.add_data_validation(dv); dv.add("B4")
    ps["A6"] = "=Settings!B4"; ps["A6"].font = Font(name=F, size=14, bold=True, color=MAIN)
    ps["A7"] = "EARNINGS STATEMENT"; ps["A7"].font = font(bold=True, size=11)
    PR = lambda col: f'IFERROR(INDEX(Payroll!${col}${P0}:${col}${P1},MATCH($B$4,Payroll!$A${P0}:$A${P1},0)),0)'
    info = [("Employee", "=IF(B4=\"\",\"(pick a name above)\",B4)", None),
            ("Employee ID", f'=IFERROR(INDEX(Employees!$G${E0}:$G${E1},MATCH($B$4,{EN},0))&"","")', None),
            ("Pay period", '=TEXT(Settings!B5,"d mmm yyyy")&" – "&TEXT(Settings!B12,"d mmm yyyy")', None),
            ("Pay date", "=Settings!B11", DATE), ("Hourly rate", f"={PR('B')}", MONEY)]
    for i, (lab, f, fmt) in enumerate(info, 9):
        ps[f"A{i}"] = lab; ps[f"A{i}"].font = font(bold=True)
        ps[f"B{i}"] = f; ps[f"B{i}"].font = font(); ps[f"B{i}"].alignment = Alignment(horizontal="left")
        if fmt: ps[f"B{i}"].number_format = fmt
    K.hdr(ps, 15, ["Earnings", "Hours", "Rate", "Amount"])
    earn = [("Regular", "G", "=B13", "J"), ("Overtime", "H", "=B13*Settings!B9", "K"), ("PTO / holiday", "I", "=B13", "L"), ("Tips / bonus", None, None, "M")]
    for i, (lab, h, rate, amt) in enumerate(earn, 16):
        ps[f"A{i}"] = lab
        ps[f"B{i}"] = f"={PR(h)}" if h else None
        ps[f"C{i}"] = rate
        ps[f"D{i}"] = f"={PR(amt)}"
    ps["A20"] = "Gross pay"; ps["B20"] = "=SUM(B16:B18)"; ps["D20"] = f"={PR('N')}"
    style_range(ps, 16, 20, 1, 4, fmt={2: HRS, 3: MONEY, 4: MONEY})
    for c in "ABCD": ps[f"{c}20"].font = font(bold=True)
    K.hdr(ps, 22, ["Deductions", "", "", "Amount"])
    ded = [("Pre-tax deductions", "O"), ("Tax withheld (estimate)", "Q"), ("After-tax deductions", "R")]
    for i, (lab, col) in enumerate(ded, 23):
        ps[f"A{i}"] = lab; ps[f"D{i}"] = f"={PR(col)}"
    ps["A26"] = "Total deductions"; ps["D26"] = "=SUM(D23:D25)"
    style_range(ps, 23, 26, 1, 4, fmt={4: MONEY})
    ps["A26"].font = font(bold=True); ps["D26"].font = font(bold=True)
    ps["A28"] = "NET PAY"; ps["A28"].font = Font(name=F, size=14, bold=True, color=MAIN)
    ps["D28"] = f"={PR('S')}"; ps["D28"].font = Font(name=F, size=14, bold=True, color=MAIN); ps["D28"].number_format = MONEY
    for c in "ABCD": ps[f"{c}28"].fill = fill(SOFT)
    ps["A30"] = "Check: gross – deductions = net"; ps["A30"].font = font(italic=True, size=9)
    ps["D30"] = '=IF(ROUND(D20-D26-D28,2)=0,"OK","Check")'; ps["D30"].font = font(italic=True, size=9)
    for c, w in zip("ABCD", [30, 26, 14, 16]): ps.column_dimensions[c].width = w

    # ---------- Dashboard ----------
    K.title(dash, "Payroll Dashboard", "Updates by itself from the Timesheet, Employees and Settings tabs.")
    dash["A3"] = '=Settings!B4&"  ·  "&TEXT(Settings!B5,"d mmm yyyy")&" – "&TEXT(Settings!B12,"d mmm yyyy")&"  ·  amounts in "&Settings!B6'
    dash["A3"].font = font(bold=True)
    T = lambda col: f"Payroll!${col}${TR}"
    kp = [("Gross payroll", f"={T('N')}", MONEY), ("Net pay to employees", f"={T('S')}", MONEY), ("Total cost to you", f"={T('U')}", MONEY),
          ("Hours worked", f"={T('C')}+{T('D')}", HRS), ("Overtime hours", f"={T('H')}", HRS),
          ("Overtime % of hours", f'=IFERROR({T("H")}/({T("C")}+{T("D")}),0)', PCT)]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 3) * 3, 1 + (i % 3) * 2, lab, f, fmt)
        dash.cell(row=6 + (i // 3) * 3, column=1 + (i % 3) * 2).alignment = Alignment(horizontal="left")
    for c, w in zip("ABCDEFG", [34, 14, 22, 14, 22, 14, 3]): dash.column_dimensions[c].width = w
    K.section(dash, "A11", "Watch list")
    wl = [("Extra cost of overtime premium", f"=SUMPRODUCT(Payroll!$H${P0}:$H${P1},Payroll!$B${P0}:$B${P1})*(Settings!B9-1)", MONEY),
          ("People with overtime", f'=COUNTIF(Payroll!$H${P0}:$H${P1},">0")', "0"),
          ("Timesheet rows with a problem", f'=COUNTIF(Timesheet!$M${T0}:$M${T1},"?*")', "0"),
          ("Timesheet rows outside this period", f'=COUNTIF(Timesheet!$L${T0}:$L${T1},"No*")', "0"),
          ("Employer payroll tax (est.)", f"={T('T')}", MONEY),
          ("Tax withheld to pay over (est.)", f"={T('Q')}", MONEY),
          ("Average cost per hour worked", f'=IFERROR({T("U")}/({T("C")}+{T("D")}),0)', MONEY)]
    for i, (lab, f, fmt) in enumerate(wl, 12):
        dash[f"A{i}"] = lab; dash[f"A{i}"].font = font(bold=True); dash[f"A{i}"].border = B; dash[f"A{i}"].fill = fill(SOFT)
        c = dash[f"B{i}"]; c.value = f; c.number_format = fmt; c.border = B; c.font = Font(name=F, bold=True, color=MAIN)
    RR = 21
    K.section(dash, f"A{RR}", "By employee")
    K.hdr(dash, RR + 1, ["Employee", "Hours", "Overtime hrs", "Gross pay", "Net pay"])
    for i in range(NE):
        r = RR + 2 + i; p = P0 + i
        dash[f"A{r}"] = f"=Payroll!A{p}"
        dash[f"B{r}"] = f'=IF(A{r}="","",Payroll!C{p}+Payroll!D{p})'
        dash[f"C{r}"] = f'=IF(A{r}="","",Payroll!H{p})'
        dash[f"D{r}"] = f'=IF(A{r}="","",Payroll!N{p})'
        dash[f"E{r}"] = f'=IF(A{r}="","",Payroll!S{p})'
    RE = RR + 1 + NE
    style_range(dash, RR + 2, RE, 1, 5, fmt={2: HRS, 3: HRS, 4: MONEY, 5: MONEY})
    ch = BarChart(); ch.type = "bar"; ch.title = "Gross pay by employee"; ch.height = 9; ch.width = 16
    nshow = len(EMP) if sample else 10
    ch.add_data(Reference(dash, min_col=4, min_row=RR + 1, max_row=RR + 1 + max(nshow, 10)), titles_from_data=True)
    ch.set_categories(Reference(dash, min_col=1, min_row=RR + 2, max_row=RR + 1 + max(nshow, 10)))
    ch.legend = None; ch.x_axis.scaling.orientation = "maxMin"; ch.y_axis.majorGridlines = None
    ch.series[0].graphicalProperties.solidFill = MAIN
    dash.add_chart(ch, "H5")
    # hours per day (14 days)
    DR = RE + 2
    K.section(dash, f"A{DR}", "Hours per day")
    K.hdr(dash, DR + 1, ["Day", "Hours worked", "People working"])
    for d in range(14):
        r = DR + 2 + d
        dash[f"A{r}"] = f'=TEXT(Settings!$B$5+{d},"ddd d mmm")'
        dash[f"B{r}"] = f"=SUMIFS({TI},Timesheet!$A${T0}:$A${T1},Settings!$B$5+{d})"
        dash[f"C{r}"] = f'=COUNTIFS(Timesheet!$A${T0}:$A${T1},Settings!$B$5+{d},{TI},">0")'
    style_range(dash, DR + 2, DR + 15, 1, 3, fmt={2: HRS, 3: "0"})
    ch2 = BarChart(); ch2.type = "col"; ch2.title = "Hours worked per day"; ch2.height = 8; ch2.width = 16
    ch2.add_data(Reference(dash, min_col=2, min_row=DR + 1, max_row=DR + 15), titles_from_data=True)
    ch2.set_categories(Reference(dash, min_col=1, min_row=DR + 2, max_row=DR + 15)); ch2.legend = None; ch2.y_axis.majorGridlines = None
    ch2.series[0].graphicalProperties.solidFill = "14B8A6"
    dash.add_chart(ch2, "H24")

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Settings tab: business name, first day of the 2-week pay period, overtime rules (weekly and optional daily) and employer tax %.", False),
             ("2.  Employees tab: name, hourly rate, tax withholding % and any deductions per pay period.", False),
             ("3.  Timesheet tab: one row per person per day – date, name, clock in, clock out and unpaid break. Add PTO/holiday hours and tips if any.", False),
             ("4.  Payroll tab: regular hours, overtime, gross pay, deductions and net pay for every employee are calculated for you.", False),
             ("5.  Pay Stub tab: pick a name and print or save as PDF to give each employee a pay stub.", False),
             ("6.  Dashboard: total payroll, net pay, cost to you, overtime and charts. The 'Check' column on the Timesheet warns about typos.", False),
             ("", False), ("Tips", True),
             ("•  Type times like 7:30 or 16:00 (24-hour) – or 4:00 PM. Overnight shifts (e.g. 22:00 → 06:00) calculate correctly.", False),
             ("•  Rows outside the pay period are ignored, so you can keep the whole year on the Timesheet and just change the start date.", False),
             ("•  Overtime is counted per work week (week 1 and week 2 separately). Turn on daily overtime (e.g. 8 h) if your state requires it.", False),
             ("•  Next pay period: change the start date on Settings. Everything moves by itself.", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  30 employees, 600 timesheet rows. No macros. Works in Excel 2010+ and Google Sheets.", False),
             ("•  Tax withheld uses a flat % you choose – it is an estimate, not an official tax table. Check with your accountant or payroll provider.", False),
             ("•  To use Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("The company, staff, hours and rates are made-up examples. Use the blank copy in your download for your own team.", False)]
    K.start(st, "Payroll & Timesheet Calculator", "Clock in/out, overtime, PTO, gross-to-net pay and pay stubs – Excel and Google Sheets.", lines)
    finish(wb, sample, path)


META = {
    "name": "Payroll & Timesheet Calculator (Overtime, PTO, Gross to Net, Pay Stubs) – Excel & Google Sheets",
    "file_base": "Payroll-Timesheet-Calculator",
    "price": 19,
    "color": "#0F766E",
    "title": "Payroll & Timesheet Calculator",
    "tagline": "clock in/out · weekly + daily overtime · PTO · gross to net · pay stubs",
    "intro": "Run payroll for your small team without expensive software: type clock-in and clock-out times, and the calculator works out hours, overtime, PTO, gross pay, deductions and net pay for every employee in the two-week pay period – then prints a pay stub for each person. Made for small businesses, contractors, cafés, shops, cleaning and trade companies.",
    "features": [
        "Bi-weekly timesheet: date, clock in/out, unpaid break, PTO/holiday hours and tips – overnight shifts calculated correctly",
        "Weekly overtime (40 h, 1.5x by default) plus optional daily overtime (e.g. after 8 h for California)",
        "Payroll for up to 30 employees: regular, overtime and PTO pay, gross pay, pre-tax and after-tax deductions, net pay",
        "Estimated tax withholding % per employee and employer payroll tax (7.65% by default) – total cost to you",
        "Printable pay stub: pick a name and print or save as PDF",
        "Check column flags missing clock-outs, duplicate days, unknown names and very long shifts",
        "Dashboard with payroll totals, overtime cost and charts; blank copy plus a sample copy (6-person company)",
    ],
    "tags": ["payroll template", "timesheet calculator", "overtime calculator", "pay stub template", "employee hours", "biweekly payroll",
             "small business payroll", "time card", "gross to net", "google sheets", "excel template", "work hours tracker", "pto tracker"],
    "summary": "Instant download: Excel + Google Sheets template, blank copy + sample copy.",
}
