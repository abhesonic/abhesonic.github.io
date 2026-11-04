"""Employee Schedule & Shift Planner – 2-week (bi-weekly) rota with shift codes, hours, overtime and labor cost.

Settings (business, period start, overtime rule, shift codes with start/end/break, roles) -> Employees (rate, max hours)
-> Schedule (14-day grid: type a shift code per day; hours, overtime, cost, alerts; staff needed & sales per day)
-> Coverage (how many people on each shift each day) -> Time Off (requests, approved days flagged on the grid)
-> Dashboard (hours, labor cost, overtime, labor % of sales, under-staffed days, cost by role, daily chart).
"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter as L
from _common import Kit, F, B, MONEY, DATE, font, fill, style_range, input_cell, finish

MAIN = "4338CA"; SOFT = "E0E7FF"
K = Kit(MAIN, SOFT)
PCT = "0.0%"; HRS = '0.0;-0.0;"-"'; TIME = "hh:mm"
NS = 20; S0 = 15; S1 = S0 + NS - 1           # shift codes on Settings A15:F34
NR = 12; R0 = 15; R1 = R0 + NR - 1           # roles on Settings H15:H26
NEMP = 40; M0 = 6; M1 = M0 + NEMP - 1        # Employees rows
GR0 = 7; GR1 = GR0 + NEMP - 1                # Schedule grid rows
D0C = 3; NDAY = 14                           # day columns C..P
NT = 80; T0 = 6; T1 = T0 + NT - 1            # Time Off rows
DAYS = [L(D0C + i) for i in range(NDAY)]     # C..P
CODES = f"Settings!$A${S0}:$A${S1}"; HRSR = f"Settings!$F${S0}:$F${S1}"
SHIFTS = [("M", "Morning", "07:00", "15:00", 30), ("D", "Day", "09:00", "17:00", 30), ("E", "Evening", "15:00", "23:00", 30),
          ("N", "Night", "23:00", "07:00", 30), ("H", "Half day", "09:00", "13:00", 0), ("L", "Long day", "08:00", "20:00", 60),
          ("OFF", "Day off", None, None, None), ("V", "Vacation / leave", None, None, None)]
ROLES = ["Manager", "Shift lead", "Barista", "Cashier", "Cook", "Kitchen helper", "Server", "Cleaner"]
COLORS = {"M": "FEF3C7", "D": "DCFCE7", "E": "DBEAFE", "N": "EDE9FE", "H": "FCE7F3", "L": "FFEDD5"}

EMP = [("Maria Lopez", "Manager", 26, 45), ("James Carter", "Shift lead", 21, 40), ("Priya Nair", "Shift lead", 21, 40),
       ("Tom Becker", "Barista", 16.5, 40), ("Aisha Khan", "Barista", 16.5, 32), ("Leo Rossi", "Barista", 16, 30),
       ("Chen Wei", "Cook", 19, 40), ("Sofia Silva", "Cook", 18.5, 40), ("Noah Smith", "Kitchen helper", 15, 35),
       ("Emma Brown", "Cashier", 15.5, 32), ("Omar Ali", "Cashier", 15.5, 25), ("Grace Lee", "Cleaner", 15, 25)]
PAT = ["M M M M M OFF OFF L M M M M OFF OFF", "D D OFF E E E E OFF D D E E E OFF", "E E E OFF OFF D D E E E OFF OFF D D",
       "M M M OFF M M OFF M M M OFF M M OFF", "OFF E E E E OFF D OFF E E E E D OFF", "OFF OFF H H OFF L L OFF OFF H H OFF L L",
       "M M M M OFF OFF M M M M M OFF OFF M", "OFF E E E E E OFF E OFF E E E E OFF", "H H OFF H H D D OFF H H H OFF D D",
       "OFF OFF E E E OFF D OFF OFF E E E OFF D", "D D OFF OFF OFF H H D OFF OFF OFF V V V", "N OFF N OFF N OFF OFF N OFF N OFF N OFF OFF"]


def hrs(cell):
    return f"SUMIFS({HRSR},{CODES},{cell})"


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); s = wb.create_sheet("Settings"); em = wb.create_sheet("Employees")
    sc = wb.create_sheet("Schedule"); cv = wb.create_sheet("Coverage"); to = wb.create_sheet("Time Off")
    p0 = dt.date(2026, 11, 9)

    # ---------- Settings ----------
    K.title(s, "Settings", "Yellow cells are yours to edit.")
    rows = [("Business name", "Bean & Bread Café" if sample else "My Business", None),
            ("Pay period start date (day 1)", p0 if sample else "=TODAY()-WEEKDAY(TODAY(),3)", DATE),
            ("Currency label", "USD", None),
            ("Overtime after (hours per week)", 40, "0.0"),
            ("Overtime pay multiplier", 1.5, "0.00"),
            ("Pay period end (calculated)", "=B5+13", DATE)]
    for i, (k_, v, fmt) in enumerate(rows, 4):
        s.cell(row=i, column=1, value=k_).font = font(bold=True)
        c = s.cell(row=i, column=2)
        if i == 9:
            c.value = v; c.border = B; c.fill = fill("F3F4F6"); c.font = font()
        else:
            input_cell(c, v)
        if fmt: c.number_format = fmt
    notes = {5: "The schedule shows 14 days (2 weeks) from this date. Change it for the next period.",
             7: "US federal rule: overtime after 40 hours in a work week. Check your local law.",
             8: "1.5 = time and a half."}
    for r, t in notes.items(): s[f"C{r}"] = t; s[f"C{r}"].font = font(italic=True, size=9)
    K.hdr(s, 14, ["Shift code", "Shift name", "Start", "End", "Unpaid break (min)", "Paid hours"])
    for i in range(NS):
        r = S0 + i
        sh = SHIFTS[i] if i < len(SHIFTS) else (None,) * 5
        code, name, a, b, brk = sh
        input_cell(s[f"A{r}"], code); input_cell(s[f"B{r}"], name)
        input_cell(s[f"C{r}"], dt.time(*map(int, a.split(":"))) if a else None)
        input_cell(s[f"D{r}"], dt.time(*map(int, b.split(":"))) if b else None)
        input_cell(s[f"E{r}"], brk)
        s[f"C{r}"].number_format = TIME; s[f"D{r}"].number_format = TIME
        c = s[f"F{r}"]; c.value = f'=IF(OR(C{r}="",D{r}=""),0,MAX(0,ROUND(MOD(D{r}-C{r},1)*24-N(E{r})/60,2)))'
        c.number_format = HRS; c.border = B; c.fill = fill("F3F4F6"); c.font = font()
    s[f"A{S1 + 1}"] = "Overnight shifts work (e.g. 23:00 → 07:00). Codes without times (OFF, V) count 0 hours. Keep codes short – you type them in the schedule."
    s[f"A{S1 + 1}"].font = font(italic=True, size=9)
    K.hdr(s, 14, ["Roles / positions"], col0=8)
    for i in range(NR): input_cell(s[f"H{R0 + i}"], ROLES[i] if i < len(ROLES) else None)
    for c, w in zip("ABCDEFGH", [32, 20, 10, 10, 12, 11, 3, 22]): s.column_dimensions[c].width = w

    # ---------- Employees ----------
    K.title(em, "Employees", "One row per person. Hourly rate and max hours are used for cost and alerts.")
    K.hdr(em, 5, ["Name", "Role", "Hourly rate", "Max hours / week", "Phone", "Email", "Start date", "Notes"], [24, 18, 12, 12, 14, 28, 12, 28])
    style_range(em, M0, M1, 1, 8, inputs=(1, 2, 3, 4, 5, 6, 7, 8), fmt={3: MONEY, 4: "0", 7: DATE})
    if sample:
        for i, (n, r_, rate, mx) in enumerate(EMP):
            r = M0 + i; em[f"A{r}"] = n; em[f"B{r}"] = r_; em[f"C{r}"] = rate; em[f"D{r}"] = mx
            em[f"E{r}"] = f"555-01{10 + i}"; em[f"F{r}"] = n.split()[0].lower() + "@beanbread.example"
            em[f"G{r}"] = dt.date(2023 + i % 3, 1 + i % 12, 1 + i)
    dv = DataValidation(type="list", formula1=f"=Settings!$H${R0}:$H${R1}", allow_blank=True); em.add_data_validation(dv); dv.add(f"B{M0}:B{M1}")
    em.freeze_panes = f"B{M0}"
    ENAME = f"Employees!$A${M0}:$A${M1}"

    # ---------- Schedule ----------
    K.title(sc, "Schedule", "Pick a person, then type a shift code for each day (codes are on Settings). Everything else calculates.")
    sc["A3"] = '=Settings!B4&"  ·  "&TEXT(Settings!B5,"d mmm yyyy")&" – "&TEXT(Settings!B9,"d mmm yyyy")'
    sc["A3"].font = font(bold=True)
    sc["A4"] = "Week"; sc["A5"] = "Date"
    for i, col in enumerate(DAYS):
        sc[f"{col}4"] = "Week 1" if i < 7 else "Week 2"
        sc[f"{col}5"] = "=Settings!$B$5" if i == 0 else f"={DAYS[i - 1]}5+1"
        sc[f"{col}5"].number_format = "d mmm"
        sc[f"{col}6"] = f'=TEXT({col}5,"ddd")'
        for rr in (4, 5): sc[f"{col}{rr}"].font = font(bold=True, size=9); sc[f"{col}{rr}"].alignment = Alignment(horizontal="center")
        sc.column_dimensions[col].width = 6.5
    for rr in (4, 5): sc[f"A{rr}"].font = font(bold=True, size=9)
    heads = ["Employee", "Role"] + [None] * NDAY + ["Week 1 hrs", "Week 2 hrs", "Total hrs", "Regular hrs", "Overtime hrs", "Hourly rate",
                                                     "Labor cost", "Time-off clashes", "Alerts"]
    for j, h in enumerate(heads, 1):
        c = sc.cell(row=6, column=j)
        if h is not None: c.value = h
        c.font = Font(name=F, bold=True, color="FFFFFF"); c.fill = fill(MAIN); c.border = B
        c.alignment = Alignment(horizontal="center" if h is None else "left", vertical="center", wrap_text=True)
    sc.row_dimensions[6].height = 30
    Q, R, S_, T, U, V, W, X, Y = [L(D0C + NDAY + k) for k in range(9)]
    for col, w in zip([Q, R, S_, T, U, V, W, X, Y], [9, 9, 9, 9, 10, 9, 12, 10, 30]): sc.column_dimensions[col].width = w
    sc.column_dimensions["A"].width = 20; sc.column_dimensions["B"].width = 14
    style_range(sc, GR0, GR1, 1, D0C + NDAY + 8, inputs=tuple([1] + list(range(D0C, D0C + NDAY))),
                fmt={D0C + NDAY: HRS, D0C + NDAY + 1: HRS, D0C + NDAY + 2: HRS, D0C + NDAY + 3: HRS, D0C + NDAY + 4: HRS,
                     D0C + NDAY + 5: MONEY, D0C + NDAY + 6: MONEY, D0C + NDAY + 7: "0"})
    TOA = f"'Time Off'!$A${T0}:$A${T1}"; TOB = f"'Time Off'!$B${T0}:$B${T1}"; TOC = f"'Time Off'!$C${T0}:$C${T1}"; TOE = f"'Time Off'!$E${T0}:$E${T1}"
    for r in range(GR0, GR1 + 1):
        for col in DAYS: sc[f"{col}{r}"].alignment = Alignment(horizontal="center")
        sc[f"B{r}"] = f'=IF(A{r}="","",IFERROR(INDEX(Employees!$B${M0}:$B${M1},MATCH(A{r},{ENAME},0)),"?"))'
        sc[f"{Q}{r}"] = f"=SUMPRODUCT({hrs(f'C{r}:I{r}')})"
        sc[f"{R}{r}"] = f"=SUMPRODUCT({hrs(f'J{r}:P{r}')})"
        sc[f"{S_}{r}"] = f"={Q}{r}+{R}{r}"
        sc[f"{T}{r}"] = f"=MIN({Q}{r},Settings!$B$7)+MIN({R}{r},Settings!$B$7)"
        sc[f"{U}{r}"] = f"={S_}{r}-{T}{r}"
        sc[f"{V}{r}"] = f'=IF(A{r}="",0,IFERROR(INDEX(Employees!$C${M0}:$C${M1},MATCH(A{r},{ENAME},0)),0))'
        sc[f"{W}{r}"] = f"={T}{r}*{V}{r}+{U}{r}*{V}{r}*Settings!$B$8"
        terms = [f'({hrs(f"{c}{r}")}>0)*(COUNTIFS({TOA},$A{r},{TOB},"<="&{c}$5,{TOC},">="&{c}$5,{TOE},"Approved")>0)' for c in DAYS]
        sc[f"{X}{r}"] = f'=IF(A{r}="",0,' + "+".join(terms) + ")"
        mx = f"IFERROR(INDEX(Employees!$D${M0}:$D${M1},MATCH(A{r},{ENAME},0)),0)"
        sc[f"{Y}{r}"] = (f'=IF(A{r}="","",TRIM(IF(B{r}="?","Not on Employees list  ","")&IF(AND({mx}>0,MAX({Q}{r},{R}{r})>{mx}),"Over max hours  ","")'
                         f'&IF({U}{r}>0,"Overtime  ","")&IF({X}{r}>0,"Works on approved time off","")))')
    if sample:
        for i, (n, *_r) in enumerate(EMP):
            r = GR0 + i; sc[f"A{r}"] = n
            for col, code in zip(DAYS, PAT[i].split()): sc[f"{col}{r}"] = code
        sc[f"G{GR0 + 3}"] = "M"  # Tom works on his approved day off -> clash example
    dv = DataValidation(type="list", formula1=f"={ENAME}", allow_blank=True); sc.add_data_validation(dv); dv.add(f"A{GR0}:A{GR1}")
    dv = DataValidation(type="list", formula1=f"={CODES}", allow_blank=True, showErrorMessage=False); sc.add_data_validation(dv)
    dv.add(f"C{GR0}:P{GR1}")
    grid = f"C{GR0}:P{GR1}"
    for code, col in COLORS.items():
        sc.conditional_formatting.add(grid, FormulaRule(formula=[f'C{GR0}="{code}"'], fill=fill(col)))
    sc.conditional_formatting.add(grid, FormulaRule(formula=[f'OR(C{GR0}="OFF",C{GR0}="V")'], font=Font(name=F, color="9CA3AF")))
    sc.conditional_formatting.add(grid, FormulaRule(formula=[f'AND(C{GR0}<>"",ISNA(MATCH(C{GR0},{CODES},0)))'], font=Font(name=F, bold=True, color="B91C1C")))
    sc.conditional_formatting.add(f"{Y}{GR0}:{Y}{GR1}", FormulaRule(formula=[f'{Y}{GR0}<>""'], font=Font(name=F, bold=True, color="B91C1C")))
    sc.conditional_formatting.add(f"{U}{GR0}:{U}{GR1}", FormulaRule(formula=[f'{U}{GR0}>0'], font=Font(name=F, bold=True, color="B45309")))
    # totals per day
    TR = GR1 + 2
    labels = ["People working", "Hours scheduled", "Labor cost (straight time)", "People needed", "Gap (+ extra / – short)", "Sales forecast", "Labor % of sales"]
    for k, lab in enumerate(labels):
        r = TR + k
        sc[f"A{r}"] = lab; sc[f"A{r}"].font = font(bold=True); sc[f"A{r}"].fill = fill(SOFT); sc[f"B{r}"].fill = fill(SOFT)
        for col in DAYS:
            c = sc[f"{col}{r}"]; c.border = B; c.font = font(size=9); c.alignment = Alignment(horizontal="center")
            rng = f"{col}{GR0}:{col}{GR1}"
            if k == 0: c.value = f"=SUMPRODUCT(--({hrs(rng)}>0))"; c.fill = fill("F3F4F6"); c.number_format = "0"
            if k == 1: c.value = f"=SUMPRODUCT({hrs(rng)})"; c.fill = fill("F3F4F6"); c.number_format = HRS
            if k == 2: c.value = f"=SUMPRODUCT({hrs(rng)}*${V}${GR0}:${V}${GR1})"; c.fill = fill("F3F4F6"); c.number_format = "#,##0"
            if k == 3: input_cell(c, ([5, 5, 5, 5, 6, 8, 5] * 2)[DAYS.index(col)] if sample else None); c.font = Font(name=F, size=9, color="0000FF")
            if k == 4: c.value = f'=IF({col}{TR + 3}="","",{col}{TR}-{col}{TR + 3})'; c.fill = fill("F3F4F6"); c.number_format = '+0;-0;0'
            if k == 5:
                input_cell(c, ([3600, 3500, 3700, 3900, 4800, 6400, 5200] * 2)[DAYS.index(col)] if sample else None)
                c.number_format = "#,##0"; c.font = Font(name=F, size=9, color="0000FF")
            if k == 6: c.value = f'=IF(N({col}{TR + 5})=0,"",{col}{TR + 2}/{col}{TR + 5})'; c.fill = fill("F3F4F6"); c.number_format = "0%"
    sc.conditional_formatting.add(f"C{TR + 4}:P{TR + 4}", FormulaRule(formula=[f'AND(ISNUMBER(C{TR + 4}),C{TR + 4}<0)'], font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    sc[f"A{TR + 7}"] = ("Totals per employee include overtime pay; the daily labor cost is straight time only. "
                        "Red code = not on the Settings list. Red cell in 'Gap' = not enough people that day.")
    sc[f"A{TR + 7}"].font = font(italic=True, size=9)
    sc.freeze_panes = f"C{GR0}"
    sc.page_setup.orientation = "landscape"

    # ---------- Coverage ----------
    K.title(cv, "Coverage", "How many people are on each shift, each day. Calculated from the Schedule – nothing to type.")
    cv["A4"] = "Shift"; cv["B4"] = "Hours"
    for i, col in enumerate(DAYS):
        cv[f"{col}4"] = f"=Schedule!{col}5"; cv[f"{col}4"].number_format = "ddd d"
    K.hdr(cv, 4, ["Shift", "Name"]); cv["A4"].value = "Shift"; cv["B4"].value = "Name"
    for col in DAYS:
        c = cv[f"{col}4"]; c.font = Font(name=F, bold=True, color="FFFFFF", size=9); c.fill = fill(MAIN); c.border = B
        c.alignment = Alignment(horizontal="center", wrap_text=True); cv.column_dimensions[col].width = 8
    cv.column_dimensions["A"].width = 10; cv.column_dimensions["B"].width = 20
    for i in range(NS):
        r = 5 + i; k = S0 + i
        cv[f"A{r}"] = f'=IF(Settings!A{k}="","",Settings!A{k})'
        cv[f"B{r}"] = f'=IF(Settings!A{k}="","",Settings!B{k}&IF(Settings!C{k}="",""," "&TEXT(Settings!C{k},"hh:mm")&"–"&TEXT(Settings!D{k},"hh:mm")))'
        for col in DAYS: cv[f"{col}{r}"] = f'=IF($A{r}="","",COUNTIF(Schedule!{col}${GR0}:{col}${GR1},$A{r}))'
    style_range(cv, 5, 4 + NS, 1, D0C + NDAY - 1)
    for r in range(5, 5 + NS):
        for col in DAYS: cv[f"{col}{r}"].alignment = Alignment(horizontal="center")
    cv.conditional_formatting.add(f"C5:P{4 + NS}", FormulaRule(formula=['AND(ISNUMBER(C5),C5>0)'], font=Font(name=F, bold=True, color=MAIN), fill=fill(SOFT)))
    cv.freeze_panes = "C5"

    # ---------- Time Off ----------
    K.title(to, "Time Off", "Requests and approved leave. Approved days that clash with a shift are flagged on the Schedule.")
    K.hdr(to, 5, ["Employee", "From", "To", "Type", "Status", "Days", "Notes"], [24, 13, 13, 14, 13, 8, 34])
    style_range(to, T0, T1, 1, 7, inputs=(1, 2, 3, 4, 5, 7), fmt={2: DATE, 3: DATE, 6: "0"})
    for r in range(T0, T1 + 1):
        to[f"F{r}"] = f'=IF(OR(B{r}="",C{r}=""),"",C{r}-B{r}+1)'
    if sample:
        req = [("Omar Ali", 12, 13, "Vacation", "Approved", "Family trip"), ("Tom Becker", 4, 4, "Personal", "Approved", "Moving house"),
               ("Aisha Khan", 18, 18, "Personal", "Requested", "Exam"), ("Chen Wei", 20, 21, "Vacation", "Requested", ""),
               ("Leo Rossi", 9, 9, "Sick", "Approved", "")]
        for i, (n, a, b, t, stt, note) in enumerate(req):
            r = T0 + i; to[f"A{r}"] = n; to[f"B{r}"] = p0 + dt.timedelta(days=a); to[f"C{r}"] = p0 + dt.timedelta(days=b)
            to[f"D{r}"] = t; to[f"E{r}"] = stt; to[f"G{r}"] = note or None
    for rng, f1 in ((f"A{T0}:A{T1}", f"={ENAME}"), (f"D{T0}:D{T1}", '"Vacation,Sick,Personal,Unpaid,Other"'),
                    (f"E{T0}:E{T1}", '"Requested,Approved,Denied"')):
        dv = DataValidation(type="list", formula1=f1, allow_blank=True); to.add_data_validation(dv); dv.add(rng)
    for txt, col in (("Approved", "15803D"), ("Requested", "B45309"), ("Denied", "9CA3AF")):
        to.conditional_formatting.add(f"E{T0}:E{T1}", FormulaRule(formula=[f'E{T0}="{txt}"'], font=Font(name=F, bold=True, color=col)))
    to.freeze_panes = f"A{T0}"

    # ---------- Dashboard ----------
    K.title(dash, "Schedule Dashboard", "Updates by itself from the Schedule, Employees and Time Off tabs.")
    dash["A3"] = '=Settings!B4&"  ·  "&TEXT(Settings!B5,"d mmm yyyy")&" – "&TEXT(Settings!B9,"d mmm yyyy")&"  ·  amounts in "&Settings!B6'
    dash["A3"].font = font(bold=True)
    G = lambda col: f"Schedule!${col}${GR0}:${col}${GR1}"
    kp = [("Hours scheduled", f"=SUM({G(S_)})", HRS), ("Labor cost", f"=SUM({G(W)})", MONEY), ("Overtime hours", f"=SUM({G(U)})", HRS),
          ("People on the schedule", f'=COUNTIF({G(S_)},">0")', "0"),
          ("Short-staffed days", f'=COUNTIF(Schedule!$C${TR + 4}:$P${TR + 4},"<0")', "0"),
          ("Labor % of sales", f'=IFERROR(SUM({G(W)})/SUM(Schedule!$C${TR + 5}:$P${TR + 5}),"–")', PCT)]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 3) * 3, 1 + (i % 3) * 2, lab, f, fmt)
        dash.cell(row=6 + (i // 3) * 3, column=1 + (i % 3) * 2).alignment = Alignment(horizontal="left")
    for c, w in zip("ABCDEFG", [36, 14, 19, 14, 17, 17, 3]): dash.column_dimensions[c].width = w

    def block(r0, title, rows_):
        K.section(dash, f"A{r0}", title)
        for i, (lab, f, fmt) in enumerate(rows_, r0 + 1):
            dash[f"A{i}"] = lab; dash[f"A{i}"].font = font(bold=True); dash[f"A{i}"].border = B; dash[f"A{i}"].fill = fill(SOFT)
            c = dash[f"B{i}"]; c.value = f; c.number_format = fmt; c.border = B; c.font = Font(name=F, bold=True, color=MAIN)
            c.alignment = Alignment(horizontal="right")
    block(11, "Watch list", [
        ("Extra cost of overtime", f"=SUMPRODUCT({G(U)},{G(V)})*(Settings!B8-1)", MONEY),
        ("People with overtime", f'=COUNTIF({G(U)},">0")', "0"),
        ("Over their max hours", f'=COUNTIF({G(Y)},"*Over max*")', "0"),
        ("Shifts on approved time off", f"=SUM({G(X)})", "0"),
        ("Time-off requests waiting", f"=COUNTIF('Time Off'!$E${T0}:$E${T1},\"Requested\")", "0"),
        ("Average hours per person", "=IFERROR(A6/A9,0)", HRS),
        ("Average cost per hour", "=IFERROR(C6/A6,0)", MONEY)])

    RR = 20
    K.section(dash, f"A{RR}", "Hours and cost by role")
    K.hdr(dash, RR + 1, ["Role", "People", "Hours", "Overtime hrs", "Labor cost"])
    for i in range(NR):
        r = RR + 2 + i; k = R0 + i
        dash[f"A{r}"] = f'=IF(Settings!H{k}="","",Settings!H{k})'
        dash[f"B{r}"] = f'=IF(A{r}="","",COUNTIFS({G("B")},A{r},{G(S_)},">0"))'
        dash[f"C{r}"] = f'=IF(A{r}="","",SUMIFS({G(S_)},{G("B")},A{r}))'
        dash[f"D{r}"] = f'=IF(A{r}="","",SUMIFS({G(U)},{G("B")},A{r}))'
        dash[f"E{r}"] = f'=IF(A{r}="","",SUMIFS({G(W)},{G("B")},A{r}))'
    RE = RR + 1 + NR
    style_range(dash, RR + 2, RE, 1, 5, fmt={2: "0", 3: HRS, 4: HRS, 5: MONEY})

    DR = RE + 2
    K.section(dash, f"A{DR}", "Day by day")
    K.hdr(dash, DR + 1, ["Day", "People", "Hours", "Needed", "Gap", "Labor % of sales"])
    for i, col in enumerate(DAYS):
        r = DR + 2 + i
        dash[f"A{r}"] = f'=TEXT(Schedule!{col}5,"ddd d mmm")'
        dash[f"B{r}"] = f"=Schedule!{col}{TR}"
        dash[f"C{r}"] = f"=Schedule!{col}{TR + 1}"
        dash[f"D{r}"] = f'=IF(Schedule!{col}{TR + 3}="","",Schedule!{col}{TR + 3})'
        dash[f"E{r}"] = f"=Schedule!{col}{TR + 4}"
        dash[f"F{r}"] = f"=Schedule!{col}{TR + 6}"
    DE = DR + 1 + NDAY
    style_range(dash, DR + 2, DE, 1, 6, fmt={2: "0", 3: HRS, 4: "0", 5: '+0;-0;0', 6: "0%"})
    dash.conditional_formatting.add(f"E{DR + 2}:E{DE}", FormulaRule(formula=[f'AND(ISNUMBER(E{DR + 2}),E{DR + 2}<0)'], font=Font(name=F, bold=True, color="B91C1C")))

    ch = BarChart(); ch.type = "col"; ch.title = "Hours scheduled per day"; ch.height = 8; ch.width = 17
    ch.add_data(Reference(dash, min_col=3, min_row=DR + 1, max_row=DE), titles_from_data=True)
    ch.set_categories(Reference(dash, min_col=1, min_row=DR + 2, max_row=DE)); ch.legend = None; ch.y_axis.majorGridlines = None
    ch.series[0].graphicalProperties.solidFill = MAIN
    dash.add_chart(ch, "H5")
    ch2 = BarChart(); ch2.type = "bar"; ch2.title = "Labor cost by role"; ch2.height = 8; ch2.width = 17
    ch2.add_data(Reference(dash, min_col=5, min_row=RR + 1, max_row=RE), titles_from_data=True)
    ch2.set_categories(Reference(dash, min_col=1, min_row=RR + 2, max_row=RE)); ch2.legend = None; ch2.y_axis.majorGridlines = None
    ch2.x_axis.scaling.orientation = "maxMin"; ch2.series[0].graphicalProperties.solidFill = "818CF8"
    dash.add_chart(ch2, "H22")

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Settings tab: business name, first day of the 2-week period, overtime rule, and your shift codes (code, start, end, unpaid break). Paid hours calculate themselves – overnight shifts work too.", False),
             ("2.  Employees tab: name, role, hourly rate and the most hours each person should work per week.", False),
             ("3.  Schedule tab: pick a person in column A, then type (or pick) a shift code for each day: M, D, E, N, OFF… Shifts are colour-coded.", False),
             ("4.  Under the grid: type how many people you need each day and (optional) your sales forecast. Short-staffed days turn red.", False),
             ("5.  Time Off tab: log requests. When a request is Approved and the person still has a shift that day, the Schedule warns you.", False),
             ("6.  Coverage tab: how many people are on each shift each day. Dashboard: hours, labor cost, overtime, labor % of sales and charts.", False),
             ("", False), ("Tips", True),
             ("•  Next period: save a copy of the file, change the start date on Settings and adjust the codes – the dates move by themselves.", False),
             ("•  Overtime is counted per week (week 1 and week 2 separately), as US federal law does. Change the hours and multiplier on Settings for your rules.", False),
             ("•  Restaurants and cafés often aim for labor at 25–35% of sales – watch the 'Labor % of sales' line.", False),
             ("•  To share the rota: print the Schedule tab (it fits on one landscape page) or share the Google Sheet as view-only.", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  40 employees, 20 shift codes, 12 roles, 80 time-off requests, 14 days per period. No macros.", False),
             ("•  Labor cost is an estimate from hourly rates – it does not include payroll taxes or benefits.", False),
             ("•  To use Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("The café, staff, rates and sales are made-up examples. Use the blank copy in your download for your own team.", False)]
    K.start(st, "Employee Schedule & Shift Planner", "Two-week rota with shift codes, hours, overtime, labor cost and time off – Excel and Google Sheets.", lines)
    finish(wb, sample, path)


META = {
    "name": "Employee Schedule & Shift Planner (2-Week Rota, Overtime, Labor Cost, Time Off) – Excel & Google Sheets",
    "file_base": "Employee-Shift-Schedule-Planner",
    "price": 19,
    "color": "#4338CA",
    "title": "Employee Shift Schedule Planner",
    "tagline": "2-week rota · shift codes · overtime · labor cost · time off · coverage",
    "intro": "Build your staff rota in minutes: type a shift code (M, E, N, OFF…) for each person and day, and the planner adds up hours, overtime and labor cost for the two-week pay period. It warns you about short-staffed days, people over their max hours and shifts that clash with approved time off. Made for cafés, restaurants, retail shops, salons, clinics and small teams.",
    "features": [
        "14-day (bi-weekly) schedule grid for up to 40 employees with colour-coded shift codes",
        "Your own shift codes with start, end and unpaid break – overnight shifts calculated correctly",
        "Weekly overtime (40 h by default, 1.5x pay) and labor cost per person, per role and per day",
        "Staff needed vs scheduled per day, with short-staffed days in red; optional labor % of sales",
        "Time-off requests with approval status; clashes with scheduled shifts are flagged",
        "Coverage view: how many people are on each shift every day",
        "Dashboard with alerts and charts; blank copy plus a sample copy (café with 12 staff)",
    ],
    "tags": ["employee schedule", "shift planner", "work schedule template", "staff rota", "shift schedule", "labor cost",
             "overtime calculator", "small business", "restaurant schedule", "google sheets", "excel template", "time off tracker", "biweekly schedule"],
}
