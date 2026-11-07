"""Teacher Gradebook & Attendance Tracker – weighted categories, letter grades, missing work, term filter, attendance, at-risk list.

Settings (class, terms, categories + weights, grade scale) -> Students -> Gradebook (students x assignments; blank = not graded,
EX = excused, M = missing counts 0) -> Attendance (students x school days; P / A / T / E) -> Report (per student) -> Dashboard.
"""
import datetime as dt, random
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, ColorScaleRule
from openpyxl.chart import BarChart, Reference
from _common import Kit, F, B, DATE, font, fill, style_range, input_cell, finish, INPUT

MAIN = "4338CA"; SOFT = "E0E7FF"
K = Kit(MAIN, SOFT)
PCT = "0.0%"
NS = 40; S0 = 6; S1 = S0 + NS - 1               # Students rows
NA = 80; AC0 = 4; AC1 = AC0 + NA - 1              # Gradebook assignment columns D..
GR0 = 11; GR1 = GR0 + NS - 1                     # Gradebook student rows
ND = 200; DC0 = 3; DC1 = DC0 + ND - 1             # Attendance day columns C..
AR0 = 7; AR1 = AR0 + NS - 1                      # Attendance student rows
NCAT = 8; CT0 = 6; CT1 = CT0 + NCAT - 1           # categories Settings A6:B13
SCALE = [("F", 0), ("D-", 0.60), ("D", 0.63), ("D+", 0.67), ("C-", 0.70), ("C", 0.73), ("C+", 0.77), ("B-", 0.80), ("B", 0.83),
         ("B+", 0.87), ("A-", 0.90), ("A", 0.93), ("A+", 0.97)]
GS0 = 6; GS1 = GS0 + len(SCALE) - 1               # grade scale Settings E6:F18
CATS = [("Homework", 0.20), ("Classwork", 0.10), ("Quizzes", 0.20), ("Tests", 0.35), ("Projects", 0.15)]
TERMS = ["Term 1", "Term 2", "Term 3", "Term 4"]


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); s = wb.create_sheet("Settings"); su = wb.create_sheet("Students")
    gb = wb.create_sheet("Gradebook"); at = wb.create_sheet("Attendance"); rp = wb.create_sheet("Report")

    # ---------- Settings ----------
    K.title(s, "Settings", "Yellow cells are yours to edit. Weights should add up to 100%.")
    SET = [("Teacher", "Ms. Rivera" if sample else "My name"), ("Class / subject", "Grade 7 Science – Period 3" if sample else "My class"),
           ("School year", "2026–2027"), ("Term to show (All, Term 1…)", "All"), ("Passing grade", 0.60), ("Attendance warning below", 0.90)]
    for i, (lab, v) in enumerate(SET, 3):
        s[f"H{i}"] = lab; s[f"H{i}"].font = font(bold=True); s[f"H{i}"].border = B; input_cell(s[f"I{i}"], v)
    s["I7"].number_format = PCT; s["I8"].number_format = PCT
    dv = DataValidation(type="list", formula1='"All,Term 1,Term 2,Term 3,Term 4"', allow_blank=False); s.add_data_validation(dv); dv.add("I6")
    K.hdr(s, 5, ["Category", "Weight"])
    for i in range(NCAT):
        r = CT0 + i; c = CATS[i] if i < len(CATS) else (None, None)
        input_cell(s[f"A{r}"], c[0]); input_cell(s[f"B{r}"], c[1]); s[f"B{r}"].number_format = PCT
    s[f"A{CT1 + 1}"] = "Total weight"; s[f"A{CT1 + 1}"].font = font(bold=True)
    s[f"B{CT1 + 1}"] = f"=SUM(B{CT0}:B{CT1})"; s[f"B{CT1 + 1}"].number_format = PCT; s[f"B{CT1 + 1}"].font = font(bold=True)
    s.conditional_formatting.add(f"B{CT1 + 1}", FormulaRule(formula=[f"ROUND(B{CT1 + 1},4)<>1"], font=Font(name=F, bold=True, color="B91C1C")))
    s[f"A{CT1 + 2}"] = "Weights are re-balanced automatically if a category has no grades yet."; s[f"A{CT1 + 2}"].font = font(italic=True, size=9)
    K.hdr(s, 5, ["Letter", "From %"], col0=5)
    for i, (L, v) in enumerate(SCALE):
        r = GS0 + i; input_cell(s[f"E{r}"], L); input_cell(s[f"F{r}"], v); s[f"F{r}"].number_format = PCT
    s[f"E{GS1 + 1}"] = "Keep the list sorted from lowest to highest."; s[f"E{GS1 + 1}"].font = font(italic=True, size=9)
    for c, w in zip("ABCDEFGHI", [20, 10, 3, 3, 9, 10, 3, 30, 28]): s.column_dimensions[c].width = w
    CAT = f"Settings!$A${CT0}:$A${CT1}"; TERM = "Settings!$I$6"; PASS = "Settings!$I$7"; AWARN = "Settings!$I$8"
    LET = f"Settings!$E${GS0}:$E${GS1}"; MINP = f"Settings!$F${GS0}:$F${GS1}"

    # ---------- Students ----------
    K.title(su, "Students", "One row per student. The order here is used on every other tab.")
    K.hdr(su, 5, ["#", "Student name", "Student ID", "Parent / guardian", "Guardian email", "Guardian phone", "Notes (IEP, seating…)"],
          [5, 26, 12, 22, 28, 15, 30])
    style_range(su, S0, S1, 1, 7, inputs=(2, 3, 4, 5, 6, 7))
    for r in range(S0, S1 + 1): su[f"A{r}"] = r - S0 + 1
    su.freeze_panes = f"C{S0}"

    # ---------- Gradebook ----------
    K.title(gb, "Gradebook", "Rows 5–9: describe each assignment. Then type scores. Blank = not graded yet, EX = excused, M = missing (counts as 0).")
    labels = [(5, "Assignment"), (6, "Category"), (7, "Date"), (8, "Max points"), (9, "Term"), (10, "Class average")]
    for r, lab in labels:
        c = gb.cell(row=r, column=3, value=lab); c.font = Font(name=F, bold=True, color="FFFFFF"); c.fill = fill(MAIN); c.border = B
    for c in range(1, 3):
        for r in range(5, 11): gb.cell(row=r, column=c).fill = fill(MAIN)
    gb["A10"] = "Student"; gb["A10"].font = Font(name=F, bold=True, color="FFFFFF"); gb["B10"] = "Overall"; gb["B10"].font = Font(name=F, bold=True, color="FFFFFF")
    gb.column_dimensions["A"].width = 24; gb.column_dimensions["B"].width = 9; gb.column_dimensions["C"].width = 13
    for j in range(NA):
        col = AC0 + j; L = CL(col); gb.column_dimensions[L].width = 11
        for r in range(5, 10):
            c = gb.cell(row=r, column=col); c.fill = fill(INPUT); c.border = B; c.font = Font(name=F, color="0000FF", bold=(r == 5))
            c.alignment = Alignment(wrap_text=True, horizontal="center")
        gb.cell(row=7, column=col).number_format = "mmm d"
        avg = gb.cell(row=10, column=col)
        avg.value = (f'=IF(OR({L}$8="",COUNT({L}{GR0}:{L}{GR1})+COUNTIF({L}{GR0}:{L}{GR1},"M")=0),"",'
                     f'SUM({L}{GR0}:{L}{GR1})/((COUNT({L}{GR0}:{L}{GR1})+COUNTIF({L}{GR0}:{L}{GR1},"M"))*{L}$8))')
        avg.number_format = "0%"; avg.font = Font(name=F, bold=True, color=MAIN); avg.fill = fill(SOFT); avg.border = B
        avg.alignment = Alignment(horizontal="center")
    gb.row_dimensions[5].height = 32
    for k in range(NS):
        r = GR0 + k
        gb[f"A{r}"] = f'=IF(Students!B{S0 + k}="","",Students!B{S0 + k})'
        gb[f"B{r}"] = f'=IF(A{r}="","",Report!J{6 + k})'
        gb[f"C{r}"] = f'=IF(A{r}="","",Report!K{6 + k})'
        for c, fmt in (("A", None), ("B", "0.0%"), ("C", None)):
            cell = gb[f"{c}{r}"]; cell.border = B; cell.fill = fill("F3F4F6"); cell.font = font(bold=(c != "A"))
            if fmt: cell.number_format = fmt
        for j in range(NA):
            c = gb.cell(row=r, column=AC0 + j); c.fill = fill(INPUT); c.border = B; c.font = Font(name=F, color="0000FF")
            c.alignment = Alignment(horizontal="center")
    AL = CL(AC0); AR = CL(AC1)
    dv = DataValidation(type="list", formula1=f"={CAT}", allow_blank=True); gb.add_data_validation(dv); dv.add(f"{AL}6:{AR}6")
    dv = DataValidation(type="list", formula1='"Term 1,Term 2,Term 3,Term 4"', allow_blank=True); gb.add_data_validation(dv); dv.add(f"{AL}9:{AR}9")
    gb.conditional_formatting.add(f"{AL}{GR0}:{AR}{GR1}", FormulaRule(formula=[f'UPPER({AL}{GR0})="M"'], font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    gb.conditional_formatting.add(f"{AL}{GR0}:{AR}{GR1}", FormulaRule(formula=[f'AND(ISNUMBER({AL}{GR0}),{AL}$8>0,{AL}{GR0}/{AL}$8<{PASS})'],
                                                                      font=Font(name=F, bold=True, color="B45309")))
    gb.conditional_formatting.add(f"B{GR0}:B{GR1}", FormulaRule(formula=[f'AND(ISNUMBER(B{GR0}),B{GR0}<{PASS})'], font=Font(name=F, bold=True, color="B91C1C")))
    gb.freeze_panes = f"{AL}{GR0}"

    # ---------- Attendance ----------
    K.title(at, "Attendance", "Row 5: type each school day's date (or fill the first and add 1). Then P = present, A = absent, T = tardy, E = excused absence.")
    at["A5"] = "Date"; at["A6"] = "Student"
    for c in ("A5", "A6", "B5", "B6"):
        at[c].font = Font(name=F, bold=True, color="FFFFFF"); at[c].fill = fill(MAIN); at[c].border = B
    at["B6"] = "Attendance %"
    at.column_dimensions["A"].width = 24; at.column_dimensions["B"].width = 12
    DL = CL(DC0); DR = CL(DC1)
    for j in range(ND):
        col = DC0 + j; L = CL(col); at.column_dimensions[L].width = 6.5
        c = at.cell(row=5, column=col); c.fill = fill(INPUT); c.border = B; c.font = Font(name=F, color="0000FF", size=8); c.number_format = "m/d"
        c.alignment = Alignment(horizontal="center", text_rotation=90)
        w = at.cell(row=6, column=col, value=f'=IF({L}5="","",TEXT({L}5,"ddd"))'); w.font = Font(name=F, size=8, color="FFFFFF"); w.fill = fill(MAIN)
        w.alignment = Alignment(horizontal="center")
    at.row_dimensions[5].height = 42
    for k in range(NS):
        r = AR0 + k
        at[f"A{r}"] = f'=IF(Students!B{S0 + k}="","",Students!B{S0 + k})'
        at[f"B{r}"] = f'=IF(A{r}="","",Report!O{6 + k})'
        for c, fmt in (("A", None), ("B", "0.0%")):
            cell = at[f"{c}{r}"]; cell.border = B; cell.fill = fill("F3F4F6"); cell.font = font(bold=(c == "B"))
            if fmt: cell.number_format = fmt
        for j in range(ND):
            c = at.cell(row=r, column=DC0 + j); c.fill = fill(INPUT); c.border = B; c.font = Font(name=F, color="0000FF", size=9)
            c.alignment = Alignment(horizontal="center")
    dv = DataValidation(type="list", formula1='"P,A,T,E"', allow_blank=True); at.add_data_validation(dv); dv.add(f"{DL}{AR0}:{DR}{AR1}")
    at.conditional_formatting.add(f"{DL}{AR0}:{DR}{AR1}", FormulaRule(formula=[f'{DL}{AR0}="A"'], font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    at.conditional_formatting.add(f"{DL}{AR0}:{DR}{AR1}", FormulaRule(formula=[f'{DL}{AR0}="T"'], font=Font(name=F, bold=True, color="B45309"), fill=fill("FEF3C7")))
    at.conditional_formatting.add(f"{DL}{AR0}:{DR}{AR1}", FormulaRule(formula=[f'{DL}{AR0}="E"'], font=Font(name=F, color="6B7280"), fill=fill("F3F4F6")))
    at.conditional_formatting.add(f"B{AR0}:B{AR1}", FormulaRule(formula=[f'AND(ISNUMBER(B{AR0}),B{AR0}<{AWARN})'], font=Font(name=F, bold=True, color="B91C1C")))
    at.freeze_panes = f"{DL}{AR0}"

    # ---------- sample data ----------
    if sample:
        rnd = random.Random(5)
        studs = ["Aiden Brooks", "Bella Cruz", "Caleb Dunn", "Daisy Evans", "Elijah Ford", "Fiona Grant", "Gabriel Hayes", "Hannah Ito",
                 "Isaac Jensen", "Jasmine Khan", "Kai Lopez", "Luna Mendes", "Mason Nash", "Nina Ortiz", "Oscar Price", "Piper Quinn",
                 "Quentin Reyes", "Rosa Silva", "Samuel Tran", "Tessa Underwood", "Uriel Vance", "Violet White", "Wyatt Xu", "Yara Young",
                 "Zane Abbott", "Amara Bell"]
        for k, n in enumerate(studs):
            r = S0 + k
            for col, v in zip("BCDEF", [n, f"S{2400 + k}", f"{['Mr.', 'Ms.'][k % 2]} {n.split()[1]}", f"{n.split()[1].lower()}.family@mail.example",
                                        f"(555) 03{k % 10}-{5100 + 7 * k}"]):
                su[f"{col}{r}"] = v
        su[f"G{S0 + 4}"] = "IEP – extra time on tests"; su[f"G{S0 + 11}"] = "Front row seat"
        skill = [rnd.uniform(0.70, 0.99) for _ in studs]; skill[6] = 0.52; skill[14] = 0.58
        plan = []
        d = dt.date(2026, 8, 24)
        hw = q = t = p = cw = 0
        for wk in range(11):
            base = d + dt.timedelta(days=7 * wk)
            hw += 1; plan.append((f"HW {hw}", "Homework", base + dt.timedelta(days=1), 10))
            cw += 1; plan.append((f"Lab {cw}", "Classwork", base + dt.timedelta(days=2), 20))
            if wk % 2 == 1: q += 1; plan.append((f"Quiz {q}", "Quizzes", base + dt.timedelta(days=4), 25))
            if wk in (4, 9): t += 1; plan.append((f"Unit {t} Test", "Tests", base + dt.timedelta(days=3), 100))
            if wk == 7: p += 1; plan.append(("Ecosystem Project", "Projects", base + dt.timedelta(days=4), 50))
        for j, (nm, cat, dd, mx) in enumerate(plan):
            col = AC0 + j
            for r, v in ((5, nm), (6, cat), (7, dd), (8, mx), (9, "Term 1")): gb.cell(row=r, column=col, value=v)
            for k in range(len(studs)):
                x = rnd.random()
                if x < 0.035 or (k in (6, 14) and cat == "Homework" and x < 0.25): v = "M"
                elif x < 0.045: v = "EX"
                else: v = max(0, min(mx, round(mx * min(1, rnd.gauss(skill[k], 0.08)) * 2) / 2))
                gb.cell(row=GR0 + k, column=col, value=v)
        days = []; dd = d
        while dd <= dt.date(2026, 11, 6):
            if dd.weekday() < 5 and dd not in (dt.date(2026, 9, 7), dt.date(2026, 10, 12)): days.append(dd)
            dd += dt.timedelta(days=1)
        for j, day in enumerate(days):
            at.cell(row=5, column=DC0 + j, value=day)
            for k in range(len(studs)):
                x = rnd.random(); lim = 0.12 if k in (6, 19) else 0.03
                v = "A" if x < lim else ("T" if x < lim + 0.03 else ("E" if x < lim + 0.045 else "P"))
                at.cell(row=AR0 + k, column=DC0 + j, value=v)

    # ---------- Report ----------
    K.title(rp, "Student Report", "Calculated for the term chosen on Settings (cell I6). Copy a row into a parent email or report card.")
    rp["A3"] = '="Showing: "&Settings!I6&"  ·  "&Settings!I4'; rp["A3"].font = font(bold=True)
    K.hdr(rp, 5, ["#", "Student"] + [""] * 6 + ["", "Overall %", "Letter", "Missing", "Excused", "Graded", "Attendance %", "Absent", "Tardy", "Status"],
          [5, 24, 11, 11, 11, 11, 11, 11, 3, 10, 8, 8, 8, 8, 12, 8, 8, 20])
    for i in range(6):
        c = rp.cell(row=5, column=3 + i, value=f'=IF(Settings!A{CT0 + i}="","",Settings!A{CT0 + i})')
        c.font = Font(name=F, bold=True, color="FFFFFF")
    R0 = 6
    GCOLS = f"Gradebook!${AL}$"; rngrow = lambda row: f"Gradebook!${AL}${row}:${AR}${row}"
    TERMOK = f'((Gradebook!${AL}$9:${AR}$9={TERM})+({TERM}="All")>0)'
    for k in range(NS):
        r = R0 + k; g = GR0 + k; sc = rngrow(g); a = AR0 + k
        rp[f"A{r}"] = k + 1
        rp[f"B{r}"] = f'=IF(Students!B{S0 + k}="","",Students!B{S0 + k})'
        for i in range(6):
            col = CL(3 + i); cat = f"Settings!$A${CT0 + i}"
            cats = f"Gradebook!${AL}$6:${AR}$6"
            earned = f'IF({TERM}="All",SUMIFS({sc},{cats},{cat}),SUMIFS({sc},{cats},{cat},Gradebook!${AL}$9:${AR}$9,{TERM}))'
            possible = (f'SUMPRODUCT((ISNUMBER({sc})+(UPPER({sc})="M"))*({cats}={cat})*{TERMOK}*{rngrow(8)})')
            rp[f"{col}{r}"] = f'=IF(OR(B{r}="",{cat}=""),"",IF({possible}=0,"",{earned}/{possible}))'
        wsum = "+".join(f"IF(ISNUMBER({CL(3 + i)}{r}),{CL(3 + i)}{r}*Settings!$B${CT0 + i},0)" for i in range(6))
        wtot = "+".join(f"IF(ISNUMBER({CL(3 + i)}{r}),Settings!$B${CT0 + i},0)" for i in range(6))
        rp[f"J{r}"] = f'=IF(B{r}="","",IF(({wtot})=0,"",({wsum})/({wtot})))'
        rp[f"K{r}"] = f'=IF(OR(B{r}="",J{r}=""),"",INDEX({LET},MATCH(J{r}+0.0000001,{MINP},1)))'
        rp[f"L{r}"] = f'=IF(B{r}="","",SUMPRODUCT((UPPER({sc})="M")*{TERMOK}))'
        rp[f"M{r}"] = f'=IF(B{r}="","",SUMPRODUCT((UPPER({sc})="EX")*{TERMOK}))'
        rp[f"N{r}"] = f'=IF(B{r}="","",SUMPRODUCT(ISNUMBER({sc})*{TERMOK}))'
        arow = f"Attendance!${DL}${a}:${DR}${a}"
        rp[f"P{r}"] = f'=IF(B{r}="","",COUNTIF({arow},"A")+COUNTIF({arow},"E"))'
        rp[f"Q{r}"] = f'=IF(B{r}="","",COUNTIF({arow},"T"))'
        rp[f"O{r}"] = f'=IF(B{r}="","",IF(COUNTIF({arow},"?")=0,"",(COUNTIF({arow},"P")+COUNTIF({arow},"T"))/COUNTIF({arow},"?")))'
        rp[f"S{r}"] = f'=IF(B{r}="",9,IF(ISNUMBER(J{r}),J{r},2))+ROW()/1000000'
        rp[f"R{r}"] = (f'=IF(B{r}="","",IF(AND(ISNUMBER(J{r}),J{r}<{PASS}),"At risk – failing",IF(N(L{r})>=3,"Missing work",'
                       f'IF(AND(ISNUMBER(O{r}),O{r}<{AWARN}),"Attendance",IF(J{r}="","","On track")))))')
    RE = R0 + NS - 1
    style_range(rp, R0, RE, 1, 18, fmt={3: PCT, 4: PCT, 5: PCT, 6: PCT, 7: PCT, 8: PCT, 10: PCT, 12: "0", 13: "0", 14: "0", 15: PCT, 16: "0", 17: "0"})
    for r in range(R0, RE + 1): rp[f"J{r}"].font = font(bold=True); rp[f"K{r}"].font = Font(name=F, bold=True, color=MAIN)
    rp.conditional_formatting.add(f"R{R0}:R{RE}", FormulaRule(formula=[f'LEFT(R{R0},2)="At"'], font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    rp.conditional_formatting.add(f"R{R0}:R{RE}", FormulaRule(formula=[f'OR(R{R0}="Missing work",R{R0}="Attendance")'], font=Font(name=F, bold=True, color="B45309"), fill=fill("FEF3C7")))
    rp.conditional_formatting.add(f"R{R0}:R{RE}", FormulaRule(formula=[f'R{R0}="On track"'], font=Font(name=F, color="15803D")))
    rp.conditional_formatting.add(f"J{R0}:J{RE}", ColorScaleRule(start_type="num", start_value=0.5, start_color="FCA5A5", mid_type="num", mid_value=0.75,
                                                                mid_color="FEF08A", end_type="num", end_value=1, end_color="86EFAC"))
    rp.column_dimensions["S"].hidden = True
    rp.freeze_panes = f"C{R0}"
    RR = lambda c: f"Report!${c}${R0}:${c}${RE}"

    # ---------- Dashboard ----------
    K.title(dash, "Class Dashboard", "Updates by itself from the Gradebook and Attendance tabs.")
    dash["A3"] = '=Settings!I4&"  ·  "&Settings!I3&"  ·  "&Settings!I5&"  ·  showing: "&Settings!I6'; dash["A3"].font = font(bold=True)
    kp = [("Students", f'=COUNTIF({RR("B")},"?*")', "0"), ("Class average", f'=IFERROR(AVERAGE({RR("J")}),0)', PCT),
          ("Average attendance", f'=IFERROR(AVERAGE({RR("O")}),0)', PCT),
          ("Students at risk", f'=COUNTIF({RR("R")},"At risk*")', "0"), ("Missing assignments", f'=SUM({RR("L")})', "0"),
          ("Assignments entered", f'=COUNTIF(Gradebook!${AL}$5:${AR}$5,"?*")', "0")]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 3) * 3, 1 + (i % 3) * 2, lab, f, fmt)
        dash.cell(row=6 + (i // 3) * 3, column=1 + (i % 3) * 2).alignment = Alignment(horizontal="left")
    for c, w in zip("ABCDEFG", [26, 14, 22, 14, 22, 14, 3]): dash.column_dimensions[c].width = w
    K.section(dash, "A11", "Grade distribution")
    K.hdr(dash, 12, ["Grade", "Students", "Share"])
    for i, L in enumerate("ABCDF"):
        r = 13 + i; dash[f"A{r}"] = L
        dash[f"B{r}"] = f'=COUNTIF({RR("K")},"{L}*")'
        dash[f"C{r}"] = f'=IFERROR(B{r}/SUM($B$13:$B$17),0)'
    style_range(dash, 13, 17, 1, 3, fmt={2: "0", 3: PCT})
    K.section(dash, "E11", "Category averages")
    K.hdr(dash, 12, ["Category", "Class avg", "Weight"], col0=5)
    for i in range(6):
        r = 13 + i
        dash[f"E{r}"] = f'=IF(Settings!A{CT0 + i}="","",Settings!A{CT0 + i})'
        dash[f"F{r}"] = f'=IF(E{r}="","",IFERROR(AVERAGE({RR(CL(3 + i))}),""))'
        dash[f"G{r}"] = f'=IF(E{r}="","",Settings!B{CT0 + i})'
    style_range(dash, 13, 18, 5, 7, fmt={6: PCT, 7: PCT})
    dash.column_dimensions["G"].width = 9
    TR = 21
    K.section(dash, f"A{TR}", "Students who need attention (lowest overall first)")
    K.hdr(dash, TR + 1, ["Student", "Overall", "Letter", "Missing", "Attendance", "Status"])
    # rank key: overall % (or 2 if blank) + row/1e6 -> SMALL
    key = RR("S")
    for k in range(10):
        r = TR + 2 + k
        pos = f'MATCH(SMALL({key},{k + 1}),{key},0)'
        dash[f"A{r}"] = f'=IFERROR(IF(SMALL({key},{k + 1})>=2,"",INDEX({RR("B")},{pos})),"")'
        for col, src in zip("BCDEF", "JKLOR"):
            dash[f"{col}{r}"] = f'=IF(A{r}="","",INDEX({RR(src)},MATCH(A{r},{RR("B")},0)))'
    style_range(dash, TR + 2, TR + 11, 1, 6, fmt={2: PCT, 4: "0", 5: PCT})
    dash.conditional_formatting.add(f"F{TR + 2}:F{TR + 11}", FormulaRule(formula=[f'LEFT(F{TR + 2},2)="At"'], font=Font(name=F, bold=True, color="B91C1C")))
    dash.column_dimensions["F"].width = 20
    ch = BarChart(); ch.type = "col"; ch.title = "Grade distribution"; ch.height = 7.5; ch.width = 14; ch.legend = None
    ch.add_data(Reference(dash, min_col=2, min_row=12, max_row=17), titles_from_data=True)
    ch.set_categories(Reference(dash, min_col=1, min_row=13, max_row=17)); ch.y_axis.majorGridlines = None
    ch.series[0].graphicalProperties.solidFill = MAIN
    dash.add_chart(ch, "I5")
    ch2 = BarChart(); ch2.type = "bar"; ch2.title = "Class average by category"; ch2.height = 7.5; ch2.width = 14; ch2.legend = None
    ch2.add_data(Reference(dash, min_col=6, min_row=12, max_row=12 + len(CATS)), titles_from_data=True)
    ch2.set_categories(Reference(dash, min_col=5, min_row=13, max_row=12 + len(CATS)))
    ch2.x_axis.scaling.orientation = "maxMin"; ch2.y_axis.majorGridlines = None; ch2.series[0].graphicalProperties.solidFill = "818CF8"
    ch2.y_axis.scaling.min = 0; ch2.y_axis.scaling.max = 1; ch2.y_axis.number_format = "0%"
    dash.add_chart(ch2, "I21")

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Settings tab: your name, class, school year, grading categories with weights (e.g. Tests 35%) and your letter-grade scale.", False),
             ("2.  Students tab: type your students' names (up to 40) and guardian contacts. They appear on every other tab automatically.", False),
             ("3.  Gradebook tab: for each assignment fill rows 5–9 (name, category, date, max points, term), then type each student's score.", False),
             ("     Leave a score blank if not graded yet, type EX for excused (ignored) or M for missing (counts as zero).", False),
             ("4.  Attendance tab: type the school dates in row 5, then P (present), A (absent), T (tardy) or E (excused) for each student.", False),
             ("5.  Report tab: category %, weighted overall %, letter grade, missing work, attendance % and a status (On track, Missing work, Attendance, At risk).", False),
             ("6.  Dashboard: class average, grade distribution, category averages and the students who need attention – all automatic.", False),
             ("", False), ("Tips", True),
             ("•  Change 'Term to show' on Settings to see grades for one term only (Term 1, Term 2…) or for the whole year (All).", False),
             ("•  Points-based grading? Put everything in one category with weight 100%.", False),
             ("•  Teach several classes? Make one copy of the file per class (File > Save as / Make a copy).", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  40 students, 80 assignments, 200 school days, 8 categories, 13-step letter scale. No macros. Works in Excel 2010+ and Google Sheets.", False),
             ("•  To use Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("The students, scores and attendance are made-up examples. Use the blank copy in your download for your own class.", False)]
    K.start(st, "Teacher Gradebook & Attendance Tracker", "Weighted grades, letter grades, missing work and attendance – Excel and Google Sheets.", lines)
    finish(wb, sample, path)


META = {
    "name": "Teacher Gradebook & Attendance Tracker (Weighted Categories, Letter Grades, Missing Work, At-Risk Students) – Excel & Google Sheets",
    "file_base": "Teacher-Gradebook-Attendance-Tracker",
    "price": 14,
    "color": "#4338CA",
    "title": "Teacher Gradebook & Attendance",
    "tagline": "weighted grades · letter grades · missing work · attendance · at-risk list",
    "intro": "A simple gradebook that does the math for you. Type scores, and each student's weighted average, letter grade and missing work update instantly. Track daily attendance with P / A / T / E codes, filter by term, and see at a glance which students are failing, missing work or missing class – ready for parent emails and report cards.",
    "features": [
        "Gradebook for 40 students and 80 assignments with categories, dates, max points and terms",
        "Weighted categories (homework, quizzes, tests…) that re-balance if a category has no grades yet",
        "Custom letter-grade scale (A+ to F), EX for excused and M for missing (counts as zero)",
        "Attendance for 200 school days: present, absent, tardy, excused, attendance % with warnings",
        "Student report: category %, overall %, letter, missing count, attendance and an On track / At risk status",
        "Term filter: show Term 1, Term 2… or the whole year",
        "Dashboard: class average, grade distribution, category averages, students who need attention and charts; blank copy plus a sample copy",
    ],
    "tags": ["teacher gradebook", "gradebook template", "attendance tracker", "teacher planner", "grade tracker", "weighted grades",
             "google sheets", "excel template", "classroom", "student tracker", "homeschool gradebook", "teacher spreadsheet", "grade calculator"],
    "summary": "Instant download: Excel + Google Sheets template, blank copy + sample copy.",
}
