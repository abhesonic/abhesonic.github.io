"""Student Assignment & Exam Planner – assignments, due dates, exam countdown, weighted grades, GPA, class timetable, final-exam calculator.

Settings (semester, courses, grade categories & weights, grading scale) -> Assignments + Class Schedule -> Grades / This Week / Timetable -> Dashboard.
"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.formatting.rule import FormulaRule, DataBarRule
from openpyxl.chart import BarChart, Reference
from _common import Kit, F, B, DATE, font, fill, style_range, input_cell, finish

MAIN = "4338CA"; SOFT = "E0E7FF"
K = Kit(MAIN, SOFT)
NC = 10; C0 = 6; C1 = C0 + NC - 1            # courses A6:A15 (cols A-F)
NK = 6                                         # grade categories (H6:I11)
KC0 = 6; KC1 = KC0 + NK - 1
CATS = [("Homework", "No"), ("Quizzes", "No"), ("Projects / papers", "No"), ("Midterm exam", "Yes"), ("Final exam", "Yes"), ("Participation", "No")]
SCALE = [(0, "F", 0.0), (60, "D-", 0.7), (63, "D", 1.0), (67, "D+", 1.3), (70, "C-", 1.7), (73, "C", 2.0), (77, "C+", 2.3),
         (80, "B-", 2.7), (83, "B", 3.0), (87, "B+", 3.3), (90, "A-", 3.7), (93, "A", 4.0)]
G0 = 6; G1 = G0 + len(SCALE) - 1               # K6:M17
W0 = 22                                        # weights grid rows 22..31, cols B..G
NA = 500; A0 = 6; A1 = A0 + NA - 1
NS = 40; S0 = 6; S1 = S0 + NS - 1
COURSES = [("Calculus I", "MATH 151", "Dr. Patel", 4, 85), ("Intro to Psychology", "PSY 101", "Prof. Rivera", 3, 90),
           ("English Composition", "ENG 110", "Ms. Chen", 3, 88), ("Biology I + Lab", "BIO 120", "Dr. Okafor", 4, 85),
           ("Spanish II", "SPAN 102", "Sra. Lopez", 3, 90)]
WEIGHTS = {"Calculus I": [15, 15, 0, 30, 35, 5], "Intro to Psychology": [10, 20, 20, 20, 25, 5],
           "English Composition": [20, 0, 50, 0, 20, 10], "Biology I + Lab": [10, 10, 25, 25, 30, 0], "Spanish II": [20, 20, 10, 20, 20, 10]}
DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
H0, H1 = 7, 21                                 # timetable hours 7:00 .. 21:00


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); asg = wb.create_sheet("Assignments"); wk = wb.create_sheet("This Week")
    gr = wb.create_sheet("Grades"); cs = wb.create_sheet("Class Schedule"); tt = wb.create_sheet("Timetable"); s = wb.create_sheet("Settings")

    # ---------- Settings ----------
    K.title(s, "Settings", "Yellow cells are yours to edit. Start with your semester dates and your courses.")
    s["A3"] = "Semester starts"; s["A3"].font = font(bold=True); input_cell(s["B3"], dt.date(2026, 8, 24)); s["B3"].number_format = DATE
    s["C3"] = "Semester ends"; s["C3"].font = font(bold=True); input_cell(s["D3"], dt.date(2026, 12, 18)); s["D3"].number_format = DATE
    K.hdr(s, 5, ["Course", "Code", "Instructor", "Credits", "Target grade %", "Notes"])
    for i in range(NC):
        r = C0 + i
        vals = COURSES[i] if sample and i < len(COURSES) else ((f"Course {i + 1}", None, None, 3, 85) if not sample and i < 4 else (None,) * 5)
        for c, v in zip("ABCDE", vals): input_cell(s[f"{c}{r}"], v)
        input_cell(s[f"F{r}"])
        s[f"D{r}"].number_format = "0"; s[f"E{r}"].number_format = "0"
    K.hdr(s, 5, ["Grade category", "Exam?"], col0=8)
    for i, (n, e) in enumerate(CATS):
        input_cell(s[f"H{KC0 + i}"], n); input_cell(s[f"I{KC0 + i}"], e)
    dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True); s.add_data_validation(dv); dv.add(f"I{KC0}:I{KC1}")
    K.hdr(s, 5, ["Score from %", "Letter", "GPA points"], col0=11)
    for i, (lo, le, gp) in enumerate(SCALE):
        input_cell(s[f"K{G0 + i}"], lo); input_cell(s[f"L{G0 + i}"], le); input_cell(s[f"M{G0 + i}"], gp); s[f"M{G0 + i}"].number_format = "0.0"
    s[f"K{G1 + 1}"] = "Lowest score first. Change it to match your school."; s[f"K{G1 + 1}"].font = font(italic=True, size=9)
    K.section(s, f"A{W0 - 2}", "Grade weights per course (% of the final grade – each row should add up to 100)")
    K.hdr(s, W0 - 1, ["Course"] + [""] * NK + ["Total"])
    for k in range(NK):
        c = s.cell(row=W0 - 1, column=2 + k, value=f'=IF(H{KC0 + k}="","–",H{KC0 + k})'); c.font = Font(name=F, bold=True, color="FFFFFF")
    for i in range(NC):
        r = W0 + i; s[f"A{r}"] = f'=IF(A{C0 + i}="","",A{C0 + i})'; s[f"A{r}"].border = B; s[f"A{r}"].fill = fill("F3F4F6"); s[f"A{r}"].font = font()
        name = COURSES[i][0] if sample and i < len(COURSES) else None
        for k in range(NK):
            v = WEIGHTS[name][k] if name else ([20, 20, 10, 20, 30, 0][k] if not sample and i < 4 else None)
            input_cell(s.cell(row=r, column=2 + k), v)
        s[f"H{r}"] = f'=IF(A{r}="","",SUM(B{r}:G{r}))'; s[f"H{r}"].border = B; s[f"H{r}"].font = font(bold=True)
    s.conditional_formatting.add(f"H{W0}:H{W0 + NC - 1}", FormulaRule(formula=[f'AND(ISNUMBER(H{W0}),H{W0}<>100)'], font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    s[f"A{W0 + NC}"] = "Red total = the weights of that course do not add up to 100. Use 0 for categories a course does not have."
    s[f"A{W0 + NC}"].font = font(italic=True, size=9)
    for c, w in zip("ABCDEFGHIJKLM", [24, 11, 14, 9, 13, 16, 11, 20, 8, 3, 12, 8, 10]): s.column_dimensions[c].width = w
    CRS = f"Settings!$A${C0}:$A${C1}"; CAT = f"Settings!$H${KC0}:$H${KC1}"
    LET = lambda x: f'IF(NOT(ISNUMBER({x})),"",INDEX(Settings!$L${G0}:$L${G1},MATCH({x}*100,Settings!$K${G0}:$K${G1},1)))'
    GPA = lambda x: f'IF(NOT(ISNUMBER({x})),"",INDEX(Settings!$M${G0}:$M${G1},MATCH({x}*100,Settings!$K${G0}:$K${G1},1)))'

    # ---------- Assignments ----------
    K.title(asg, "Assignments & Exams", "One row per assignment, quiz, paper or exam. Mark it Done when you hand it in; type the points when you get it back.")
    cols = ["Course", "Assignment / exam", "Category", "Due date", "Priority", "Status", "Points earned", "Points possible", "Notes / link",
            "Days left", "Score %", "Alert", "Exam?", "Week of"]
    K.hdr(asg, 5, cols, [20, 30, 17, 12, 9, 12, 10, 10, 24, 9, 9, 11, 7, 11])
    style_range(asg, A0, A1, 1, 14, inputs=range(1, 10), fmt={4: DATE, 7: "0.0", 8: "0.0", 10: "0", 11: "0.0%", 14: DATE})
    for r in range(A0, A1 + 1):
        asg[f"J{r}"] = f'=IF(OR(NOT(ISNUMBER(D{r})),F{r}="Done"),"",D{r}-TODAY())'
        asg[f"K{r}"] = f'=IF(AND(ISNUMBER(G{r}),N(H{r})>0),G{r}/H{r},"")'
        asg[f"L{r}"] = f'=IF(OR(B{r}="",NOT(ISNUMBER(D{r}))),"",IF(F{r}="Done","Done",IF(D{r}<TODAY(),"OVERDUE",IF(D{r}=TODAY(),"Today",IF(D{r}-TODAY()<=3,"Soon","")))))'
        asg[f"M{r}"] = f'=IF(C{r}="","",IFERROR(INDEX(Settings!$I${KC0}:$I${KC1},MATCH(C{r},{CAT},0)),"No"))'
        asg[f"N{r}"] = f'=IF(ISNUMBER(D{r}),D{r}-WEEKDAY(D{r},2)+1,"")'
        asg[f"O{r}"] = f'=IF(AND(B{r}<>"",ISNUMBER(D{r}),F{r}<>"Done"),D{r}+ROW()/100000,99999999)'
        asg[f"P{r}"] = f'=IF(AND(B{r}<>"",ISNUMBER(D{r}),D{r}>=Week_Start,D{r}<Week_Start+7),D{r}+ROW()/100000,99999999)'
        asg[f"Q{r}"] = f'=IF(AND(B{r}<>"",ISNUMBER(D{r}),F{r}<>"Done",M{r}="Yes",D{r}>=TODAY()),D{r}+ROW()/100000,99999999)'
    for c in "OPQ": asg.column_dimensions[c].hidden = True
    for col, lst in (("A", f"={CRS}"), ("C", f"={CAT}"), ("E", '"High,Medium,Low"'), ("F", '"Not started,In progress,Done"')):
        dv = DataValidation(type="list", formula1=lst, allow_blank=True); asg.add_data_validation(dv); dv.add(f"{col}{A0}:{col}{A1}")
    asg.conditional_formatting.add(f"L{A0}:L{A1}", FormulaRule(formula=[f'L{A0}="OVERDUE"'], font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    asg.conditional_formatting.add(f"L{A0}:L{A1}", FormulaRule(formula=[f'OR(L{A0}="Today",L{A0}="Soon")'], font=Font(name=F, bold=True, color="B45309"), fill=fill("FEF3C7")))
    asg.conditional_formatting.add(f"L{A0}:L{A1}", FormulaRule(formula=[f'L{A0}="Done"'], font=Font(name=F, bold=True, color="15803D"), fill=fill("DCFCE7")))
    asg.conditional_formatting.add(f"E{A0}:E{A1}", FormulaRule(formula=[f'E{A0}="High"'], font=Font(name=F, bold=True, color="B91C1C")))
    asg.conditional_formatting.add(f"A{A0}:I{A1}", FormulaRule(formula=[f'$F{A0}="Done"'], font=Font(name=F, color="9CA3AF")))
    asg.conditional_formatting.add(f"K{A0}:K{A1}", FormulaRule(formula=[f'AND(ISNUMBER(K{A0}),K{A0}<0.7)'], font=Font(name=F, bold=True, color="B91C1C")))
    asg.freeze_panes = f"C{A0}"; asg.auto_filter.ref = f"A5:N{A1}"
    AR = lambda c: f"Assignments!${c}${A0}:${c}${A1}"

    # ---------- This Week ----------
    K.title(wk, "This Week", "Everything due in one week, in date order. Type any date below to look at another week (blank = this week).")
    wk["A3"] = "Any date in the week"; wk["A3"].font = font(bold=True); input_cell(wk["C3"]); wk["C3"].number_format = DATE
    wk["E3"] = "Week of"; wk["E3"].font = font(bold=True)
    wk["F3"] = "=IF(ISNUMBER(C3),C3-WEEKDAY(C3,2)+1,TODAY()-WEEKDAY(TODAY(),2)+1)"; wk["F3"].number_format = "mmm d, yyyy"; wk["F3"].font = Font(name=F, bold=True, color=MAIN)
    wb.defined_names["Week_Start"] = DefinedName("Week_Start", attr_text="'This Week'!$F$3")
    K.hdr(wk, 5, ["Day", "Due", "Course", "Assignment / exam", "Category", "Priority", "Status", "✓"], [6, 10, 20, 34, 17, 9, 12, 5])
    key = AR("P"); NW = 40
    for k in range(NW):
        r = 6 + k; pos = f"MATCH(SMALL({key},{k + 1}),{key},0)"
        wk[f"B{r}"] = f'=IFERROR(IF(SMALL({key},{k + 1})>=99999999,"",INDEX({AR("D")},{pos})),"")'
        wk[f"A{r}"] = f'=IF(B{r}="","",TEXT(B{r},"ddd"))'
        for col, src in zip("CDEFG", "ABCEF"):
            wk[f"{col}{r}"] = f'=IF(B{r}="","",IF(INDEX({AR(src)},{pos})="","",INDEX({AR(src)},{pos})))'
    style_range(wk, 6, 5 + NW, 1, 8, fmt={2: "mmm d"})
    wk.conditional_formatting.add(f"G6:G{5 + NW}", FormulaRule(formula=['G6="Done"'], font=Font(name=F, bold=True, color="15803D"), fill=fill("DCFCE7")))
    wk.conditional_formatting.add(f"A6:H{5 + NW}", FormulaRule(formula=['AND(ISNUMBER($B6),$B6<TODAY(),$G6<>"Done")'], font=Font(name=F, bold=True, color="B91C1C")))
    wk[f"A{7 + NW}"] = "Red = past due and not Done. Print this page and tick ✓ as you go."; wk[f"A{7 + NW}"].font = font(italic=True, size=9)

    # ---------- Grades ----------
    K.title(gr, "Grades & GPA", "Your current grade per course, weighted by category. Only assignments with points count.")
    K.hdr(gr, 5, ["Course"] + [""] * NK + ["Grade %", "Letter", "GPA pts", "Credits", "Target %", "vs target", "Graded items"])
    for k in range(NK):
        c = gr.cell(row=5, column=2 + k, value=f'=IF(Settings!H{KC0 + k}="","–",Settings!H{KC0 + k})'); c.font = Font(name=F, bold=True, color="FFFFFF")
    for i in range(NC):
        r = 6 + i; sc = f"Settings!A{C0 + i}"
        gr[f"A{r}"] = f'=IF({sc}="","",{sc})'
        terms_w, terms_wa = [], []
        for k in range(NK):
            L = CL(2 + k); hc = f"{L}$5"
            gr[f"{L}{r}"] = (f'=IF(OR($A{r}="",{hc}="–"),"",IFERROR(SUMIFS({AR("G")},{AR("A")},$A{r},{AR("C")},{hc},{AR("G")},">=0")/'
                             f'SUMIFS({AR("H")},{AR("A")},$A{r},{AR("C")},{hc},{AR("G")},">=0"),""))')
            wcell = f"N(Settings!{L}{W0 + i})"
            terms_wa.append(f"IF(ISNUMBER({L}{r}),{wcell}*{L}{r},0)"); terms_w.append(f"IF(ISNUMBER({L}{r}),{wcell},0)")
        gr[f"H{r}"] = f'=IF(A{r}="","",IFERROR(({"+".join(terms_wa)})/({"+".join(terms_w)}),""))'
        gr[f"I{r}"] = "=" + LET(f"H{r}")
        gr[f"J{r}"] = "=" + GPA(f"H{r}")
        gr[f"K{r}"] = f'=IF(A{r}="","",N(Settings!D{C0 + i}))'
        gr[f"L{r}"] = f'=IF(OR(A{r}="",Settings!E{C0 + i}=""),"",Settings!E{C0 + i}/100)'
        gr[f"M{r}"] = f'=IF(OR(NOT(ISNUMBER(H{r})),NOT(ISNUMBER(L{r}))),"",H{r}-L{r})'
        gr[f"N{r}"] = f'=IF(A{r}="","",COUNTIFS({AR("A")},A{r},{AR("G")},">=0"))'
        gr[f"O{r}"] = f'=IF(ISNUMBER(J{r}),K{r},0)'; gr[f"P{r}"] = f'=IF(ISNUMBER(J{r}),J{r}*K{r},0)'
    gr.column_dimensions["O"].hidden = True; gr.column_dimensions["P"].hidden = True
    fm = {c: "0.0%" for c in range(2, 9)}; fm.update({10: "0.00", 11: "0", 12: "0%", 13: '+0.0%;-0.0%;0.0%', 14: "0"})
    style_range(gr, 6, 5 + NC, 1, 14, fmt=fm)
    for r in range(6, 6 + NC):
        gr[f"H{r}"].font = font(bold=True); gr[f"I{r}"].font = font(bold=True); gr[f"I{r}"].alignment = Alignment(horizontal="center")
    gr.conditional_formatting.add(f"M6:M{5 + NC}", FormulaRule(formula=['AND(ISNUMBER(M6),M6<0)'], font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    gr.conditional_formatting.add(f"M6:M{5 + NC}", FormulaRule(formula=['AND(ISNUMBER(M6),M6>=0)'], font=Font(name=F, bold=True, color="15803D"), fill=fill("DCFCE7")))
    T = 7 + NC
    gr[f"A{T}"] = "Semester GPA"; gr[f"A{T}"].font = font(bold=True, size=12)
    gr[f"B{T}"] = f'=IF(SUM(O6:O{5 + NC})=0,"",SUM(P6:P{5 + NC})/SUM(O6:O{5 + NC}))'; gr[f"B{T}"].number_format = "0.00"
    gr[f"B{T}"].font = Font(name=F, bold=True, size=14, color=MAIN)
    gr[f"D{T}"] = "Credits graded"; gr[f"D{T}"].font = font(bold=True); gr[f"F{T}"] = f"=SUM(O6:O{5 + NC})"; gr[f"F{T}"].number_format = "0"
    for c in ("A", "B", "D", "E", "F"): gr[f"{c}{T}"].fill = fill(SOFT)
    # final-exam calculator
    Q = T + 3
    K.section(gr, f"A{Q}", "What do I need on the final?")
    gr[f"A{Q + 1}"] = "Pick a course, the grade you want and how much the final exam is worth. Use it before the final is graded."
    gr[f"A{Q + 1}"].font = font(italic=True, size=9)
    rows = [("Course", "Calculus I" if sample else None, "General", True), ("Grade I want (%)", 85, "0", True),
            ("Final exam weight (%)", 35 if sample else 30, "0", True),
            ("Current grade (without the final)", f'=IF(B{Q + 2}="","",IFERROR(INDEX($H$6:$H${5 + NC},MATCH(B{Q + 2},$A$6:$A${5 + NC},0)),""))', "0.0%", False),
            ("You need on the final", f'=IF(OR(NOT(ISNUMBER(B{Q + 5})),N(B{Q + 4})<=0),"",(B{Q + 3}/100-B{Q + 5}*(1-B{Q + 4}/100))/(B{Q + 4}/100))', "0.0%", False),
            ("", f'=IF(NOT(ISNUMBER(B{Q + 6})),"",IF(B{Q + 6}>1,"Above 100% – not reachable with the final alone",IF(B{Q + 6}<=0,"Already safe – even 0% on the final works","Reachable – go for it!")))', "General", False)]
    for j, (lab, v, fmt, inp) in enumerate(rows):
        r = Q + 2 + j; gr[f"A{r}"] = lab; gr[f"A{r}"].font = font(bold=True); gr[f"A{r}"].border = B
        if inp: input_cell(gr[f"B{r}"], v)
        else: gr[f"B{r}"] = v; gr[f"B{r}"].fill = fill(SOFT); gr[f"B{r}"].border = B; gr[f"B{r}"].font = font(bold=True)
        gr[f"B{r}"].number_format = fmt
    gr[f"B{Q + 6}"].font = Font(name=F, bold=True, size=14, color=MAIN)
    dv = DataValidation(type="list", formula1=f"={CRS}", allow_blank=True); gr.add_data_validation(dv); dv.add(f"B{Q + 2}")
    gr[f"A{Q + 8}"] = "Tip: the current grade leaves out categories with no points yet – enter all other graded work first."
    gr[f"A{Q + 8}"].font = font(italic=True, size=9)
    for c, w in zip("ABCDEFGHIJKLMN", [30, 15, 11, 13, 12, 11, 13, 10, 8, 9, 9, 10, 10, 10]): gr.column_dimensions[c].width = w
    gr.freeze_panes = "B6"

    # ---------- Class Schedule ----------
    K.title(cs, "Class Schedule", "One row per class meeting: a course that meets Mon, Wed and Fri needs three rows. Times like 9:30 or 14:00.")
    K.hdr(cs, 5, ["Course", "Day", "Start", "End", "Room / link", "Type"], [22, 8, 9, 9, 22, 12])
    style_range(cs, S0, S1, 1, 6, inputs=range(1, 7), fmt={3: "h:mm", 4: "h:mm"})
    for col, lst in (("A", f"={CRS}"), ("B", '"' + ",".join(DAYS) + '"'), ("F", '"Lecture,Lab,Seminar,Tutorial,Office hours,Online,Study group,Work"')):
        dv = DataValidation(type="list", formula1=lst, allow_blank=True); cs.add_data_validation(dv); dv.add(f"{col}{S0}:{col}{S1}")
    cs.freeze_panes = f"A{S0}"
    CSR = lambda c: f"'Class Schedule'!${c}${S0}:${c}${S1}"

    # ---------- Timetable ----------
    K.title(tt, "Weekly Timetable", "Builds itself from the Class Schedule tab. Each cell shows the class that runs during that hour.")
    K.hdr(tt, 5, ["Time"] + DAYS, [9] + [17] * 7)
    for j, h in enumerate(range(H0, H1 + 1)):
        r = 6 + j; tt[f"A{r}"] = dt.time(h, 0); tt[f"A{r}"].number_format = "h:mm"
        for d in range(7):
            c = CL(2 + d); dn = f"{c}$5"
            n = (f"SUMPRODUCT(MAX(({CSR('B')}={dn})*({CSR('C')}<$A{r}+1/24-0.0001)*({CSR('D')}>$A{r}+0.0001)*({CSR('A')}<>\"\")"
                 f"*(ROW({CSR('A')})-{S0 - 1})))")
            tt[f"{c}{r}"] = f'=IF({n}=0,"",INDEX({CSR("A")},{n}))'
    style_range(tt, 6, 6 + H1 - H0, 1, 8, fmt={1: "h:mm"})
    for r in range(6, 7 + H1 - H0):
        tt[f"A{r}"].font = font(bold=True)
        for d in range(7): tt[f"{CL(2 + d)}{r}"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        tt.row_dimensions[r].height = 22
    tt.conditional_formatting.add(f"B6:H{6 + H1 - H0}", FormulaRule(formula=['B6<>""'], font=Font(name=F, bold=True, color="FFFFFF"), fill=fill(MAIN)))
    tt[f"A{8 + H1 - H0}"] = "If two classes overlap, the one lower in the Class Schedule list is shown."; tt[f"A{8 + H1 - H0}"].font = font(italic=True, size=9)

    # ---------- sample ----------
    if sample:
        today = dt.date(2026, 11, 10); start = dt.date(2026, 8, 24)
        meet = {"Calculus I": [("Mon", 9, 0, 10, 15, "Lecture"), ("Wed", 9, 0, 10, 15, "Lecture"), ("Fri", 9, 0, 9, 50, "Tutorial")],
                "Intro to Psychology": [("Tue", 11, 0, 12, 15, "Lecture"), ("Thu", 11, 0, 12, 15, "Lecture")],
                "English Composition": [("Mon", 13, 0, 14, 15, "Seminar"), ("Wed", 13, 0, 14, 15, "Seminar")],
                "Biology I + Lab": [("Tue", 8, 0, 9, 15, "Lecture"), ("Thu", 8, 0, 9, 15, "Lecture"), ("Thu", 14, 0, 16, 50, "Lab")],
                "Spanish II": [("Mon", 15, 0, 15, 50, "Lecture"), ("Wed", 15, 0, 15, 50, "Lecture"), ("Fri", 11, 0, 11, 50, "Lecture")]}
        rooms = {"Calculus I": "Science Hall 204", "Intro to Psychology": "Auditorium B", "English Composition": "Library 3.12",
                 "Biology I + Lab": "Bio Building 110", "Spanish II": "Humanities 18"}
        r = S0
        for c, ms in meet.items():
            for (d, h1, m1, h2, m2, ty) in ms:
                for col, v in zip("ABCDEF", (c, d, dt.time(h1, m1), dt.time(h2, m2), rooms[c] if ty != "Lab" else "Bio Lab 2", ty)): cs[f"{col}{r}"] = v
                r += 1
        items = []
        def add(course, title, cat, due, poss, prio="Medium", sc=None):
            items.append([course, title, cat, due, prio, poss, sc])
        wk0 = lambda w, d: start + dt.timedelta(days=7 * w + d)
        base = {"Calculus I": 0.84, "Intro to Psychology": 0.91, "English Composition": 0.88, "Biology I + Lab": 0.80, "Spanish II": 0.93}
        for w in range(16):
            add("Calculus I", f"Problem set {w + 1}", "Homework", wk0(w, 4), 20)
            if w % 2 == 1: add("Calculus I", f"Quiz {w // 2 + 1}", "Quizzes", wk0(w, 2), 10, "High")
            add("Spanish II", f"Workbook unit {w + 1}", "Homework", wk0(w, 0) + dt.timedelta(days=7), 10, "Low")
            if w % 3 == 2: add("Spanish II", f"Vocab quiz {w // 3 + 1}", "Quizzes", wk0(w, 4), 20)
            if w % 2 == 0: add("Intro to Psychology", f"Chapter {w + 1}-{w + 2} quiz", "Quizzes", wk0(w, 3), 15)
            if w % 2 == 0: add("Biology I + Lab", f"Lab report {w // 2 + 1}", "Projects / papers", wk0(w, 3) + dt.timedelta(days=7), 25)
            if w % 3 == 0: add("Biology I + Lab", f"Reading questions ch. {w // 3 + 1}", "Homework", wk0(w, 1), 10, "Low")
        add("Calculus I", "Midterm exam", "Midterm exam", dt.date(2026, 10, 14), 100, "High")
        add("Calculus I", "Final exam", "Final exam", dt.date(2026, 12, 14), 150, "High")
        add("Intro to Psychology", "Midterm exam", "Midterm exam", dt.date(2026, 10, 15), 80, "High")
        add("Intro to Psychology", "Research paper: memory", "Projects / papers", dt.date(2026, 11, 20), 100, "High")
        add("Intro to Psychology", "Final exam", "Final exam", dt.date(2026, 12, 15), 100, "High")
        add("English Composition", "Essay 1: narrative", "Projects / papers", dt.date(2026, 9, 18), 100)
        add("English Composition", "Essay 2: argument", "Projects / papers", dt.date(2026, 10, 16), 100, "High")
        add("English Composition", "Essay 3: research draft", "Projects / papers", dt.date(2026, 11, 13), 100, "High")
        add("English Composition", "Essay 3: final + portfolio", "Final exam", dt.date(2026, 12, 11), 100, "High")
        for w in range(0, 16, 2): add("English Composition", f"Peer review {w // 2 + 1}", "Homework", wk0(w, 2) + dt.timedelta(days=7), 10, "Low")
        add("Biology I + Lab", "Midterm exam", "Midterm exam", dt.date(2026, 10, 13), 100, "High")
        add("Biology I + Lab", "Final exam", "Final exam", dt.date(2026, 12, 16), 120, "High")
        add("Spanish II", "Oral midterm", "Midterm exam", dt.date(2026, 10, 9), 50, "High")
        add("Spanish II", "Final exam (written + oral)", "Final exam", dt.date(2026, 12, 17), 100, "High")
        add("Biology I + Lab", "Practical exam (lab)", "Midterm exam", dt.date(2026, 11, 19), 50, "High")
        items.sort(key=lambda x: x[3])
        n = 0
        for it in items:
            course, title, cat, due, prio, poss, _ = it
            n += 1
            b = base[course] + (((n * 37) % 15) - 7) / 100
            if due < today:
                status = "Done"; earned = round(min(1, b) * poss * 2) / 2
            elif due <= today + dt.timedelta(days=4):
                status = "In progress" if n % 2 else "Not started"; earned = None
            else:
                status = "Not started"; earned = None
            if due < today and n % 29 == 0: status = "In progress"; earned = None
            note = "Ch. 3–5, bring calculator" if title == "Midterm exam" and course == "Calculus I" else ("Use APA style" if "paper" in title else None)
            row = A0 + n - 1
            for col, v in zip("ABCDEFGHI", (course, title, cat, due, prio, status, earned, poss, note)):
                if v is not None: asg[f"{col}{row}"] = v

    # ---------- Dashboard ----------
    K.title(dash, "Student Dashboard", "Updates by itself from the Assignments, Grades and Settings tabs.")
    dash["A3"] = ('="Today: "&TEXT(TODAY(),"dddd, mmm d, yyyy")&"   ·   week "&MAX(1,INT((TODAY()-Settings!B3)/7)+1)&" of "&'
                  'ROUNDUP((Settings!D3-Settings!B3+1)/7,0)&"   ·   "&MAX(0,Settings!D3-TODAY())&" days left in the semester"')
    dash["A3"].font = font(bold=True)
    kp = [("Due in the next 7 days", f'=COUNTIFS({AR("D")},">="&TODAY(),{AR("D")},"<"&(TODAY()+7),{AR("F")},"<>Done",{AR("B")},"<>")', "0"),
          ("Overdue", f'=COUNTIFS({AR("D")},"<"&TODAY(),{AR("F")},"<>Done",{AR("B")},"<>")', "0"),
          ("Exams in the next 30 days", f'=COUNTIFS({AR("D")},">="&TODAY(),{AR("D")},"<"&(TODAY()+30),{AR("M")},"Yes",{AR("F")},"<>Done")', "0"),
          ("Next exam in (days)", f'=IFERROR(IF(SMALL({AR("Q")},1)>=99999999,"–",INT(SMALL({AR("Q")},1))-TODAY()),"–")', "0"),
          ("Assignments done", f'=IFERROR(COUNTIFS({AR("F")},"Done",{AR("B")},"<>")/COUNTIF({AR("B")},"?*"),0)', "0%"),
          ("Semester GPA", f"=IF(Grades!B{T}=\"\",\"–\",Grades!B{T})", "0.00")]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 3) * 3, 1 + (i % 3) * 3, lab, f, fmt)
        dash.cell(row=6 + (i // 3) * 3, column=1 + (i % 3) * 3).alignment = Alignment(horizontal="left")
    q1 = f'SMALL({AR("Q")},1)'; qp = f'MATCH({q1},{AR("Q")},0)'
    dash["A10"] = f'=IFERROR(IF({q1}>=99999999,"","Next: "&INDEX({AR("A")},{qp})&" – "&INDEX({AR("B")},{qp})&" ("&TEXT(INDEX({AR("D")},{qp}),"mmm d")&")"),"")'
    dash["A10"].font = font(size=9, italic=True)
    K.section(dash, "A12", "Next deadlines (overdue first)")
    K.hdr(dash, 13, ["Due", "Days left", "Course", "Assignment / exam", "Category", "Priority", "Status"])
    key = AR("O")
    for k in range(12):
        r = 14 + k; pos = f"MATCH(SMALL({key},{k + 1}),{key},0)"
        dash[f"A{r}"] = f'=IFERROR(IF(SMALL({key},{k + 1})>=99999999,"",INDEX({AR("D")},{pos})),"")'
        dash[f"B{r}"] = f'=IF(A{r}="","",A{r}-TODAY())'
        for col, src in zip("CDEFG", "ABCEF"):
            dash[f"{col}{r}"] = f'=IF(A{r}="","",IF(INDEX({AR(src)},{pos})="","",INDEX({AR(src)},{pos})))'
    style_range(dash, 14, 25, 1, 7, fmt={1: "ddd mmm d", 2: "0"})
    dash.conditional_formatting.add("A14:G25", FormulaRule(formula=['AND(ISNUMBER($B14),$B14<0)'], font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    dash.conditional_formatting.add("A14:G25", FormulaRule(formula=['AND(ISNUMBER($B14),$B14<=3)'], font=Font(name=F, bold=True, color="B45309"), fill=fill("FEF3C7")))
    K.section(dash, "A28", "Courses")
    K.hdr(dash, 29, ["Course", "Grade %", "Letter", "Target %", "Open tasks", "Overdue", "Next due"])
    for i in range(NC):
        r = 30 + i; cn = f"A{r}"
        dash[cn] = f"=Grades!A{6 + i}"
        dash[f"B{r}"] = f"=Grades!H{6 + i}"; dash[f"C{r}"] = f"=Grades!I{6 + i}"; dash[f"D{r}"] = f"=Grades!L{6 + i}"
        dash[f"E{r}"] = f'=IF({cn}="","",COUNTIFS({AR("A")},{cn},{AR("F")},"<>Done",{AR("B")},"<>"))'
        dash[f"F{r}"] = f'=IF({cn}="","",COUNTIFS({AR("A")},{cn},{AR("F")},"<>Done",{AR("D")},"<"&TODAY(),{AR("B")},"<>"))'
        v = (f'SUMPRODUCT(MIN(({AR("A")}={cn})*({AR("F")}<>"Done")*({AR("D")}>=TODAY())*{AR("D")}'
             f'+((({AR("A")}<>{cn})+({AR("F")}="Done")+({AR("D")}<TODAY()))>0)*999999))')
        dash[f"G{r}"] = f'=IF({cn}="","",IF({v}>=999999,"",{v}))'
    style_range(dash, 30, 29 + NC, 1, 7, fmt={2: "0.0%", 4: "0%", 5: "0", 6: "0", 7: "mmm d"})
    dash.conditional_formatting.add(f"F30:F{29 + NC}", FormulaRule(formula=['AND(ISNUMBER(F30),F30>0)'], font=Font(name=F, bold=True, color="B91C1C")))
    dash.conditional_formatting.add(f"B30:B{29 + NC}", DataBarRule(start_type="num", start_value=0.5, end_type="num", end_value=1, color="A5B4FC"))
    for c, w in zip("ABCDEFGHI", [20, 10, 18, 32, 16, 9, 15, 3, 3]): dash.column_dimensions[c].width = w
    ch = BarChart(); ch.type = "bar"; ch.title = "Current grade by course"; ch.height = 7.5; ch.width = 14; ch.legend = None
    ch.add_data(Reference(dash, min_col=2, min_row=29, max_row=29 + max(len(COURSES), 4)), titles_from_data=True)
    ch.set_categories(Reference(dash, min_col=1, min_row=30, max_row=29 + max(len(COURSES), 4)))
    ch.x_axis.scaling.orientation = "maxMin"; ch.y_axis.majorGridlines = None; ch.y_axis.scaling.min = 0; ch.y_axis.scaling.max = 1
    ch.y_axis.number_format = "0%"; ch.series[0].graphicalProperties.solidFill = MAIN
    dash.add_chart(ch, "J5")
    # tasks due per week (next 8 weeks)
    K.section(dash, "J28", "Workload: tasks due per week")
    K.hdr(dash, 29, ["Week of", "Tasks", "Exams"], col0=10)
    wk0f = "TODAY()-WEEKDAY(TODAY(),2)+1"
    for k in range(8):
        r = 30 + k; dash[f"J{r}"] = f"={wk0f}+{k * 7}"
        dash[f"K{r}"] = f'=COUNTIFS({AR("N")},J{r},{AR("B")},"<>")'
        dash[f"L{r}"] = f'=COUNTIFS({AR("N")},J{r},{AR("M")},"Yes")'
    style_range(dash, 30, 37, 10, 12, fmt={10: "mmm d", 11: "0", 12: "0"})
    for c, w in zip("JKL", [11, 8, 8]): dash.column_dimensions[c].width = w
    ch2 = BarChart(); ch2.type = "col"; ch2.grouping = "stacked"; ch2.overlap = 100; ch2.title = "Due per week"; ch2.height = 6.5; ch2.width = 10
    ch2.add_data(Reference(dash, min_col=11, max_col=12, min_row=29, max_row=37), titles_from_data=True)
    ch2.set_categories(Reference(dash, min_col=10, min_row=30, max_row=37)); ch2.y_axis.majorGridlines = None; ch2.gapWidth = 40
    ch2.series[0].graphicalProperties.solidFill = "A5B4FC"; ch2.series[1].graphicalProperties.solidFill = "B91C1C"
    dash.add_chart(ch2, "N28")

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Settings: type the first and last day of your semester and your courses (up to 10) with credits and the grade you aim for.", False),
             ("     Then fill in the grade weights from each syllabus (e.g. Homework 15%, Midterm 30%, Final 35%). Each row should total 100.", False),
             ("2.  Assignments: one row per homework, quiz, paper, project or exam – pick the course and category, type the due date.", False),
             ("     Set Status to Done when you hand it in. When the grade comes back, type Points earned and Points possible.", False),
             ("3.  This Week: everything due this week in date order. Type any date to look ahead. Print it for your desk.", False),
             ("4.  Grades: your weighted grade per course, letter grade, GPA and a 'What do I need on the final?' calculator.", False),
             ("5.  Class Schedule + Timetable: type each class meeting once – the weekly timetable fills itself in.", False),
             ("6.  Dashboard: due in the next 7 days, overdue, days until the next exam, % done, GPA, next deadlines and workload per week.", False),
             ("", False), ("Tips", True),
             ("•  Copy every deadline from all your syllabi on day one – then nothing surprises you in week 10.", False),
             ("•  Exams are the categories marked 'Exam? = Yes' in Settings – they get their own countdown.", False),
             ("•  'Soon' = due within 3 days. Red rows on the Dashboard are overdue.", False),
             ("•  Your school uses another grading scale or no +/-? Change the scale table in Settings.", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  10 courses, 6 grade categories, 500 assignments, 40 class meetings. No macros. Works in Excel 2010+ and Google Sheets.", False),
             ("•  Grades here are an estimate from your own numbers – your instructor's gradebook is the official one.", False),
             ("•  To use Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("A made-up first-year student with five courses, fall semester 2026. Use the blank copy in your download for your own classes.", False)]
    K.start(st, "Student Assignment & Exam Planner", "Deadlines, exams, grades and GPA for the whole semester – Excel and Google Sheets.", lines)
    finish(wb, sample, path)


META = {
    "name": "Student Assignment & Exam Planner (Due Dates, Exam Countdown, Weighted Grade & GPA Calculator, Class Timetable) – Excel & Google Sheets",
    "file_base": "Student-Assignment-Exam-Planner",
    "price": 12,
    "color": "#4338CA",
    "title": "Student Assignment & Exam Planner",
    "tagline": "deadlines · exam countdown · weighted grades · GPA · class timetable",
    "intro": "Keep every deadline, exam and grade of your semester in one spreadsheet. Type your assignments once from your syllabi and the planner shows what is due this week, what is overdue and how many days are left until your next exam. Enter your scores and it calculates your weighted grade per course, your letter grade and your semester GPA – and tells you what you need on the final.",
    "features": [
        "Assignment tracker for 500 tasks: course, category, due date, priority, status, points, notes – with Overdue / Today / Soon alerts",
        "This Week view: everything due in any week, in date order, ready to print",
        "Exam countdown: exams in the next 30 days and days until your next exam",
        "Weighted grade calculator per course using the weights from your syllabus (homework, quizzes, papers, midterm, final, participation)",
        "Letter grade and GPA points from an editable grading scale (A to F with +/-), semester GPA weighted by credits",
        "'What do I need on the final?' calculator",
        "Class schedule that builds a weekly timetable (7:00 to 21:00, Monday to Sunday) by itself",
        "Dashboard: due in 7 days, overdue, % done, GPA, next 12 deadlines, grade per course vs target, workload per week chart",
        "Up to 10 courses; blank copy plus a sample copy with a full example semester",
    ],
    "tags": ["assignment tracker", "student planner", "exam planner", "gpa calculator", "grade tracker", "college planner",
             "homework tracker", "study planner", "google sheets", "excel template", "semester planner", "class schedule", "university planner"],
    "summary": "Instant download: Excel + Google Sheets template, blank copy + sample copy.",
}
