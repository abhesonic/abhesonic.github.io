"""Homeschool Planner & Lesson Tracker – lesson log, weekly plan, attendance (days & hours), subject progress, grades, report card.

Settings (students, subjects, school year, grading scale) -> Lesson Log + Attendance -> Weekly Plan / Progress / Report Card -> Dashboard.
"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, DataBarRule
from openpyxl.chart import BarChart, Reference
from _common import Kit, F, B, DATE, font, fill, style_range, input_cell, finish

MAIN = "15803D"; SOFT = "DCFCE7"
K = Kit(MAIN, SOFT)
NS = 6; ST0 = 6; ST1 = ST0 + NS - 1          # students A6:A11
NSUB = 15; SU0 = 6; SU1 = SU0 + NSUB - 1      # subjects D6:D20
GR = [(0, "F"), (60, "D"), (70, "C"), (80, "B"), (90, "A")]; G0 = 6; G1 = G0 + len(GR) - 1   # G6:H10
NL = 1200; L0 = 6; L1 = L0 + NL - 1
NA = 260; A0 = 6; A1 = A0 + NA - 1
STUDENTS = [("Emma", "Grade 5"), ("Noah", "Grade 3"), ("Lily", "Grade 1")]
SUBJ = [("Math", 4), ("Reading", 3), ("Writing", 2), ("Science", 2), ("History", 2), ("Spelling", 1), ("Art", 1), ("Music", 1),
        ("Physical Education", 2)]


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); wp = wb.create_sheet("Weekly Plan"); lg = wb.create_sheet("Lesson Log")
    at = wb.create_sheet("Attendance"); pr = wb.create_sheet("Progress"); rc = wb.create_sheet("Report Card"); s = wb.create_sheet("Settings")

    # ---------- Settings ----------
    K.title(s, "Settings", "Yellow cells are yours to edit. Type your students and subjects first.")
    K.hdr(s, 5, ["Student", "Grade / note"])
    K.hdr(s, 5, ["Subject", "Target hours / week"], col0=4)
    K.hdr(s, 5, ["Score from", "Letter"], col0=7)
    for i in range(NS):
        nm, gr = STUDENTS[i] if sample and i < len(STUDENTS) else ((f"Student {i + 1}", None) if not sample and i < 2 else (None, None))
        input_cell(s[f"A{ST0 + i}"], nm); input_cell(s[f"B{ST0 + i}"], gr)
    for i in range(NSUB):
        nm, h = SUBJ[i] if i < len(SUBJ) else (None, None)
        input_cell(s[f"D{SU0 + i}"], nm); input_cell(s[f"E{SU0 + i}"], h)
    for i, (lo, le) in enumerate(GR):
        input_cell(s[f"G{G0 + i}"], lo); input_cell(s[f"H{G0 + i}"], le)
    s[f"G{G1 + 1}"] = "Lowest score first. Scores are 0–100."; s[f"G{G1 + 1}"].font = font(italic=True, size=9)
    SET = [(23, "School year starts", dt.date(2026, 8, 17), DATE), (24, "School year ends", dt.date(2027, 5, 28), DATE),
           (25, "Required school days per year", 180, "0"), (26, "Required hours per year (0 = none)", 1000, "0")]
    for r, lab, v, fmt in SET:
        s[f"A{r}"] = lab; s[f"A{r}"].font = font(bold=True); s[f"A{r}"].border = B; input_cell(s[f"C{r}"], v); s[f"C{r}"].number_format = fmt
    s["A27"] = "Weeks of school so far"; s["A27"].font = font(italic=True)
    s["C27"] = "=MAX(0,(MIN(TODAY(),C24)-C23+1)/7)"; s["C27"].number_format = "0.0"
    s["A28"] = "Weeks in the school year"; s["A28"].font = font(italic=True)
    s["C28"] = "=MAX(1,(C24-C23+1)/7)"; s["C28"].number_format = "0.0"
    s["A29"] = "Rules vary by state / country – check what your local law requires and type it above."; s["A29"].font = font(italic=True, size=9)
    for c, w in zip("ABCDEFGH", [30, 14, 13, 22, 13, 3, 11, 8]): s.column_dimensions[c].width = w
    STU = f"Settings!$A${ST0}:$A${ST1}"; SUB = f"Settings!$D${SU0}:$D${SU1}"
    LET = lambda x: f'IF({x}="","",INDEX(Settings!$H${G0}:$H${G1},MATCH({x},Settings!$G${G0}:$G${G1},1)))'

    # ---------- Lesson Log ----------
    K.title(lg, "Lesson Log", "Plan and record every lesson. Status 'Done' counts the minutes; type a score (0–100) for graded work.")
    cols = ["Date", "Student", "Subject", "Lesson / topic", "Curriculum / resource", "Minutes", "Status", "Score (0–100)", "Notes",
            "Hours", "Week of", "Letter"]
    K.hdr(lg, 5, cols, [12, 12, 17, 32, 22, 9, 10, 10, 26, 8, 11, 7])
    style_range(lg, L0, L1, 1, 12, inputs=range(1, 10), fmt={1: DATE, 6: "0", 8: "0", 10: "0.00", 11: DATE})
    for r in range(L0, L1 + 1):
        lg[f"J{r}"] = f'=IF(AND(G{r}="Done",ISNUMBER(F{r})),F{r}/60,0)'
        lg[f"K{r}"] = f'=IF(ISNUMBER(A{r}),A{r}-WEEKDAY(A{r},2)+1,"")'
        lg[f"L{r}"] = "=" + LET(f"H{r}")
        lg[f"M{r}"] = f'=IF(AND(ISNUMBER(A{r}),A{r}>=Weekly_Start,A{r}<Weekly_Start+7,B{r}<>""),A{r}+ROW()/100000,99999999)'
    lg.column_dimensions["M"].hidden = True
    for col, lst in (("B", f"={STU}"), ("C", f"={SUB}"), ("G", '"Planned,Done,Skipped"')):
        dv = DataValidation(type="list", formula1=lst, allow_blank=True); lg.add_data_validation(dv); dv.add(f"{col}{L0}:{col}{L1}")
    dv = DataValidation(type="decimal", operator="between", formula1="0", formula2="100", allow_blank=True,
                        error="Type a score from 0 to 100", showErrorMessage=True); lg.add_data_validation(dv); dv.add(f"H{L0}:H{L1}")
    lg.conditional_formatting.add(f"G{L0}:G{L1}", FormulaRule(formula=[f'G{L0}="Done"'], font=Font(name=F, bold=True, color="15803D"), fill=fill(SOFT)))
    lg.conditional_formatting.add(f"G{L0}:G{L1}", FormulaRule(formula=[f'G{L0}="Skipped"'], font=Font(name=F, color="9CA3AF")))
    lg.conditional_formatting.add(f"G{L0}:G{L1}", FormulaRule(formula=[f'AND(G{L0}="Planned",ISNUMBER(A{L0}),A{L0}<TODAY())'],
                                                              font=Font(name=F, bold=True, color="B45309"), fill=fill("FEF3C7")))
    lg.freeze_panes = f"A{L0}"; lg.auto_filter.ref = f"A5:L{L1}"
    LR = lambda c: f"'Lesson Log'!${c}${L0}:${c}${L1}"

    # ---------- Attendance ----------
    K.title(at, "Attendance", "School days fill in by themselves (Mon–Fri). Type P = present, F = field trip / outing, A = absent, H = holiday / no school.")
    K.hdr(at, 5, ["Date", "Day"] + [""] * NS + ["Note"], [12, 6] + [11] * NS + [28])
    for i in range(NS):
        c = at.cell(row=5, column=3 + i, value=f'=IF(Settings!A{ST0 + i}="","–",Settings!A{ST0 + i})'); c.font = Font(name=F, bold=True, color="FFFFFF")
    style_range(at, A0, A1, 1, 3 + NS, inputs=range(3, 4 + NS), fmt={1: DATE})
    at[f"A{A0}"] = "=WORKDAY(Settings!$C$23-1,1)"
    for r in range(A0, A1 + 1):
        if r > A0: at[f"A{r}"] = f"=WORKDAY(A{r - 1},1)"
        at[f"B{r}"] = f'=TEXT(A{r},"ddd")'
    dv = DataValidation(type="list", formula1='"P,F,A,H"', allow_blank=True); at.add_data_validation(dv)
    dv.add(f"C{A0}:{CL(2 + NS)}{A1}")
    ar = f"C{A0}:{CL(2 + NS)}{A1}"
    at.conditional_formatting.add(ar, FormulaRule(formula=[f'OR(C{A0}="P",C{A0}="F")'], font=Font(name=F, bold=True, color="15803D"), fill=fill(SOFT)))
    at.conditional_formatting.add(ar, FormulaRule(formula=[f'C{A0}="A"'], font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    at.conditional_formatting.add(f"A{A0}:{CL(3 + NS)}{A1}", FormulaRule(formula=[f'$A{A0}=TODAY()'], fill=fill("FEF3C7")))
    at.conditional_formatting.add(f"A{A0}:B{A1}", FormulaRule(formula=[f'$A{A0}>Settings!$C$24'], font=Font(name=F, color="C0C7CF")))
    at.freeze_panes = f"C{A0}"
    AT = lambda i: f"Attendance!${CL(3 + i)}${A0}:${CL(3 + i)}${A1}"
    DAYS = lambda i: f'(COUNTIF({AT(i)},"P")+COUNTIF({AT(i)},"F"))'

    # ---------- Weekly Plan ----------
    K.title(wp, "Weekly Plan", "Lessons for one week, in date order – print it for the fridge. Change the week below (blank = this week).")
    wp["A3"] = "Any date in the week"; wp["A3"].font = font(bold=True); input_cell(wp["C3"]); wp["C3"].number_format = DATE
    wp["E3"] = "Week of"; wp["E3"].font = font(bold=True)
    wp["F3"] = "=IF(ISNUMBER(C3),C3-WEEKDAY(C3,2)+1,TODAY()-WEEKDAY(TODAY(),2)+1)"; wp["F3"].number_format = "mmm d, yyyy"
    wp["F3"].font = Font(name=F, bold=True, color=MAIN)
    from openpyxl.workbook.defined_name import DefinedName
    wb.defined_names["Weekly_Start"] = DefinedName("Weekly_Start", attr_text="'Weekly Plan'!$F$3")
    K.hdr(wp, 5, ["Day", "Date", "Student", "Subject", "Lesson / topic", "Resource", "Minutes", "Status", "✓"], [6, 10, 12, 17, 34, 22, 9, 10, 5])
    key = LR("M"); NW = 60
    for k in range(NW):
        r = 6 + k; pos = f"MATCH(SMALL({key},{k + 1}),{key},0)"
        wp[f"B{r}"] = f'=IFERROR(IF(SMALL({key},{k + 1})>=99999999,"",INDEX({LR("A")},{pos})),"")'
        wp[f"A{r}"] = f'=IF(B{r}="","",TEXT(B{r},"ddd"))'
        for col, src in zip("CDEFGH", "BCDEFG"):
            wp[f"{col}{r}"] = f'=IF(B{r}="","",IF(INDEX({LR(src)},{pos})="","",INDEX({LR(src)},{pos})))'
    style_range(wp, 6, 5 + NW, 1, 9, fmt={2: "mmm d", 7: "0"})
    wp.conditional_formatting.add(f"A6:I{5 + NW}", FormulaRule(formula=['AND($B6<>"",$B6<>$B5,ROW()>6)'], border=None,
                                                             font=Font(name=F, bold=True)))
    wp.conditional_formatting.add(f"H6:H{5 + NW}", FormulaRule(formula=['H6="Done"'], font=Font(name=F, bold=True, color="15803D"), fill=fill(SOFT)))
    wp[f"A{7 + NW}"] = "Showing up to 60 lessons. Tick ✓ on paper, then set Status to Done in the Lesson Log."; wp[f"A{7 + NW}"].font = font(italic=True, size=9)

    # ---------- Progress ----------
    K.title(pr, "Progress", "Hours done per subject and student compared with your weekly target, plus average scores.")
    SBLK = [("Hours done", "hours"), ("Average score", "score"), ("Lessons done", "count")]
    row = 4
    for title, kind in SBLK:
        K.section(pr, f"A{row}", title + (" (target to date = target hours per week × weeks so far)" if kind == "hours" else ""))
        hdrs = ["Subject"] + ([ "Target / week", "Target to date"] if kind == "hours" else ["", ""])
        K.hdr(pr, row + 1, hdrs + [""] * NS)
        for i in range(NS):
            c = pr.cell(row=row + 1, column=4 + i, value=f'=IF(Settings!A{ST0 + i}="","–",Settings!A{ST0 + i})'); c.font = Font(name=F, bold=True, color="FFFFFF")
        for j in range(NSUB):
            r = row + 2 + j; pr[f"A{r}"] = f'=IF(Settings!D{SU0 + j}="","",Settings!D{SU0 + j})'
            if kind == "hours":
                pr[f"B{r}"] = f'=IF(A{r}="","",N(Settings!E{SU0 + j}))'; pr[f"C{r}"] = f'=IF(A{r}="","",B{r}*Settings!$C$27)'
            for i in range(NS):
                cc = f"{CL(4 + i)}{r}"; sn = f"{CL(4 + i)}${row + 1}"
                if kind == "hours":
                    pr[cc] = f'=IF(OR($A{r}="",{sn}="–"),"",SUMIFS({LR("J")},{LR("B")},{sn},{LR("C")},$A{r}))'
                elif kind == "score":
                    pr[cc] = f'=IF(OR($A{r}="",{sn}="–"),"",IFERROR(AVERAGEIFS({LR("H")},{LR("B")},{sn},{LR("C")},$A{r},{LR("G")},"Done"),""))'
                else:
                    pr[cc] = f'=IF(OR($A{r}="",{sn}="–"),"",COUNTIFS({LR("B")},{sn},{LR("C")},$A{r},{LR("G")},"Done"))'
        fmt = {c: ("0.0" if kind != "count" else "0") for c in range(2, 4 + NS)}
        style_range(pr, row + 2, row + 1 + NSUB, 1, 3 + NS, fmt=fmt)
        if kind == "hours":
            rg = f"D{row + 2}:{CL(3 + NS)}{row + 1 + NSUB}"
            pr.conditional_formatting.add(rg, FormulaRule(formula=[f'AND(ISNUMBER(D{row + 2}),N($C{row + 2})>0,D{row + 2}<0.8*$C{row + 2})'],
                                                          font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
            pr.conditional_formatting.add(rg, FormulaRule(formula=[f'AND(ISNUMBER(D{row + 2}),N($C{row + 2})>0,D{row + 2}>=$C{row + 2})'],
                                                          font=Font(name=F, bold=True, color="15803D"), fill=fill(SOFT)))
            tr = row + 2 + NSUB; pr[f"A{tr}"] = "Total"
            for c in range(2, 4 + NS):
                L = CL(c); pr[f"{L}{tr}"] = f"=SUM({L}{row + 2}:{L}{row + 1 + NSUB})"; pr[f"{L}{tr}"].number_format = "0.0"
                pr[f"{L}{tr}"].font = font(bold=True)
            pr[f"A{tr}"].font = font(bold=True)
            pr[f"A{tr + 1}"] = "Red = less than 80% of the target so far. Green = on target."; pr[f"A{tr + 1}"].font = font(italic=True, size=9)
            HTOT = tr
        if kind == "score":
            rg = f"D{row + 2}:{CL(3 + NS)}{row + 1 + NSUB}"
            pr.conditional_formatting.add(rg, FormulaRule(formula=[f'AND(ISNUMBER(D{row + 2}),D{row + 2}<70)'], font=Font(name=F, bold=True, color="B91C1C")))
        row += NSUB + 5
    for c, w in zip("ABC", [20, 13, 14]): pr.column_dimensions[c].width = w
    for i in range(NS): pr.column_dimensions[CL(4 + i)].width = 11
    pr.freeze_panes = "B4"

    # ---------- Report Card ----------
    K.title(rc, "Report Card", "Pick a student – the report fills in by itself. Ready to print or save as PDF for your records / portfolio.")
    rc["A3"] = "Student"; rc["A3"].font = font(bold=True); input_cell(rc["B3"], "Emma" if sample else None)
    dv = DataValidation(type="list", formula1=f"={STU}", allow_blank=True); rc.add_data_validation(dv); dv.add("B3")
    rc["C3"] = f'=IF(B3="","",IFERROR(INDEX(Settings!$B${ST0}:$B${ST1},MATCH(B3,{STU},0)),""))'
    rc["D3"] = '="School year "&TEXT(Settings!C23,"mmm yyyy")&" – "&TEXT(Settings!C24,"mmm yyyy")&"   ·   as of "&TEXT(TODAY(),"mmm d, yyyy")'
    rc["D3"].font = font(italic=True)
    K.hdr(rc, 5, ["Subject", "Lessons done", "Hours", "Target to date", "% of target", "Average score", "Grade", "Teacher comment"],
          [20, 12, 9, 13, 11, 13, 8, 40])
    for j in range(NSUB):
        r = 6 + j; sj = f"Settings!D{SU0 + j}"
        rc[f"A{r}"] = f'=IF({sj}="","",{sj})'
        rc[f"B{r}"] = f'=IF(OR(A{r}="",$B$3=""),"",COUNTIFS({LR("B")},$B$3,{LR("C")},A{r},{LR("G")},"Done"))'
        rc[f"C{r}"] = f'=IF(OR(A{r}="",$B$3=""),"",SUMIFS({LR("J")},{LR("B")},$B$3,{LR("C")},A{r}))'
        rc[f"D{r}"] = f'=IF(A{r}="","",N(Settings!E{SU0 + j})*Settings!$C$27)'
        rc[f"E{r}"] = f'=IF(OR(C{r}="",N(D{r})=0),"",C{r}/D{r})'
        rc[f"F{r}"] = f'=IF(OR(A{r}="",$B$3=""),"",IFERROR(AVERAGEIFS({LR("H")},{LR("B")},$B$3,{LR("C")},A{r},{LR("G")},"Done"),""))'
        rc[f"G{r}"] = "=" + LET(f"F{r}")
    style_range(rc, 6, 5 + NSUB, 1, 8, inputs=(8,), fmt={2: "0", 3: "0.0", 4: "0.0", 5: "0%", 6: "0.0"})
    for r in range(6, 6 + NSUB): rc[f"G{r}"].font = font(bold=True); rc[f"G{r}"].alignment = Alignment(horizontal="center")
    rc.conditional_formatting.add(f"E6:E{5 + NSUB}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="86EFAC"))
    T = 7 + NSUB
    K.section(rc, f"A{T}", "Attendance & totals")
    si = f"MATCH($B$3,{STU},0)"
    att = [("School days attended", f'=IF($B$3="","",IFERROR(INDEX(Dashboard!$B$14:$B$19,{si}),""))', "0"),
           ("Required school days", "=Settings!C25", "0"),
           ("Days absent", f'=IF($B$3="","",IFERROR(INDEX(Dashboard!$D$14:$D$19,{si}),""))', "0"),
           ("Total hours", f"=SUM(C6:C{5 + NSUB})", "0.0"),
           ("Required hours", "=Settings!C26", "0"),
           ("Overall average score", f'=IF($B$3="","",IFERROR(AVERAGEIFS({LR("H")},{LR("B")},$B$3,{LR("G")},"Done"),""))', "0.0"),
           ("Overall grade", "=" + LET(f"B{T + 6}"), "General")]
    for i, (lab, f, fmt) in enumerate(att):
        r = T + 1 + i; rc[f"A{r}"] = lab; rc[f"B{r}"] = f
        rc[f"A{r}"].font = font(bold=True); rc[f"B{r}"].number_format = fmt
        for c in "AB": rc[f"{c}{r}"].border = B; rc[f"{c}{r}"].fill = fill(SOFT)
    rc[f"A{T + 10}"] = "Parent / teacher signature: ______________________________      Date: ______________"; rc[f"A{T + 10}"].font = font()
    if sample:
        rc["H6"] = "Mastered long division; working on fractions."; rc["H7"] = "Reads chapter books independently."
        rc["H9"] = "Loved the plant cell model project."

    # ---------- sample ----------
    if sample:
        topics = {"Math": ["Place value", "Long division", "Fractions intro", "Equivalent fractions", "Adding fractions", "Decimals", "Word problems", "Geometry: angles"],
                  "Reading": ["Charlotte's Web ch.", "Main idea", "Vocabulary", "Read-aloud", "Book report prep", "Poetry"],
                  "Writing": ["Paragraph structure", "Narrative draft", "Editing", "Letter writing", "Journal"],
                  "Science": ["Plant cells", "Water cycle", "Weather log", "Simple machines", "Nature walk notes"],
                  "History": ["Ancient Egypt", "Map skills", "Timeline project", "Explorers"],
                  "Spelling": ["Weekly list", "Spelling test"], "Art": ["Watercolor", "Clay", "Drawing animals"],
                  "Music": ["Recorder practice", "Rhythm games"], "Physical Education": ["Park run", "Swimming", "Bike ride"]}
        res = {"Math": "Math workbook", "Reading": "Library books", "Writing": "Writing journal", "Science": "Science kit",
               "History": "Story of the World", "Spelling": "Spelling list", "Art": "Art supplies", "Music": "Recorder", "Physical Education": "Outdoors"}
        plan = {0: ["Math", "Reading", "Writing", "Science"], 1: ["Math", "Reading", "Spelling", "History"],
                2: ["Math", "Reading", "Writing", "Art"], 3: ["Math", "Reading", "Science", "Music"], 4: ["Math", "Spelling", "History", "Physical Education"]}
        mins = {"Emma": {"Math": 60, "Reading": 45, "Writing": 40, "Science": 50, "History": 45, "Spelling": 20, "Art": 45, "Music": 30, "Physical Education": 60},
                "Noah": {"Math": 45, "Reading": 40, "Writing": 30, "Science": 40, "History": 30, "Spelling": 20, "Art": 40, "Music": 30, "Physical Education": 60},
                "Lily": {"Math": 30, "Reading": 30, "Writing": 20, "Science": 30, "History": 20, "Spelling": 15, "Art": 40, "Music": 30, "Physical Education": 45}}
        base = {"Emma": 91, "Noah": 84, "Lily": 88}
        today = dt.date(2026, 11, 9); d = dt.date(2026, 8, 17); end = dt.date(2026, 11, 13)
        absent = {(dt.date(2026, 10, 2), "Noah"), (dt.date(2026, 10, 19), "Lily")}
        holiday = {dt.date(2026, 9, 7), dt.date(2026, 11, 26)}
        trips = {dt.date(2026, 9, 25), dt.date(2026, 10, 23)}
        rows = []; att = {}; n = 0
        while d <= end:
            wd = d.weekday()
            if wd < 5:
                for si, (stu, _g) in enumerate(STUDENTS):
                    code = "H" if d in holiday else "A" if (d, stu) in absent else "F" if d in trips else "P"
                    if d <= today: att[(d, si)] = code
                    if code in ("H", "A"): continue
                    subs = plan[wd] if code == "P" else ["Science", "History"]
                    for sj in subs:
                        n += 1; tl = topics[sj]; t = tl[(n // 7) % len(tl)]
                        if sj == "Reading" and t.endswith("ch."): t = t + f" {(n // 5) % 22 + 1}"
                        status = "Done" if d < today else "Planned"
                        if status == "Done" and n % 37 == 0: status = "Skipped"
                        sc = None
                        if status == "Done" and sj in ("Math", "Spelling", "Science", "History", "Writing") and n % 3 == 0:
                            sc = max(55, min(100, base[stu] + ((n * 7) % 17) - 8 - (6 if sj == "Spelling" and stu == "Noah" else 0)))
                        note = "Field trip: science museum" if code == "F" and sj == "Science" else None
                        rows.append((d, stu, sj, t, res[sj], mins[stu][sj], status, sc, note))
            d += dt.timedelta(days=1)
        for k, row in enumerate(rows[:NL]):
            for c, v in zip("ABCDEFGHI", row):
                if v is not None: lg[f"{c}{L0 + k}"] = v
        wd0 = dt.date(2026, 8, 17); k = 0; day = wd0
        while day <= today:
            if day.weekday() < 5:
                for si in range(len(STUDENTS)):
                    if (day, si) in att: at[f"{CL(3 + si)}{A0 + k}"] = att[(day, si)]
                k += 1
            day += dt.timedelta(days=1)
        at[f"{CL(3 + NS)}{A0 + 15}"] = "Labor Day"; at[f"{CL(3 + NS)}{A0 + 29}"] = "Science museum trip"

    # ---------- Dashboard ----------
    K.title(dash, "Homeschool Dashboard", "Updates by itself from the Lesson Log and Attendance tabs.")
    dash["A3"] = '="Today: "&TEXT(TODAY(),"dddd, mmm d, yyyy")&"   ·   week "&INT(Settings!C27)+1&" of "&ROUND(Settings!C28,0)'
    dash["A3"].font = font(bold=True)
    wk = "TODAY()-WEEKDAY(TODAY(),2)+1"
    kp = [("Lessons done", f'=COUNTIF({LR("G")},"Done")', "0"),
          ("Hours logged", f"=SUM({LR('J')})", "0.0"),
          ("Average score", f'=IFERROR(AVERAGEIFS({LR("H")},{LR("G")},"Done"),"–")', "0.0"),
          ("Lessons planned this week", f'=COUNTIFS({LR("A")},">="&{wk},{LR("A")},"<"&({wk}+7),{LR("G")},"Planned")', "0"),
          ("Overdue planned lessons", f'=COUNTIFS({LR("A")},"<"&TODAY(),{LR("G")},"Planned")', "0"),
          ("School days left this year", f'=MAX(0,NETWORKDAYS(MAX(TODAY(),Settings!C23),Settings!C24))', "0")]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 3) * 3, 1 + (i % 3) * 2, lab, f, fmt)
        dash.cell(row=6 + (i // 3) * 3, column=1 + (i % 3) * 2).alignment = Alignment(horizontal="left")
    K.section(dash, "A12", "Students")
    K.hdr(dash, 13, ["Student", "Days attended", "% of required", "Days absent", "Hours", "% of req. hours", "Avg score", "Grade"])
    for i in range(NS):
        r = 14 + i; sn = f"A{r}"
        dash[sn] = f'=IF(Settings!A{ST0 + i}="","",Settings!A{ST0 + i})'
        dash[f"B{r}"] = f'=IF({sn}="","",{DAYS(i)})'
        dash[f"C{r}"] = f'=IF(OR({sn}="",N(Settings!$C$25)=0),"",B{r}/Settings!$C$25)'
        dash[f"D{r}"] = f'=IF({sn}="","",COUNTIF({AT(i)},"A"))'
        dash[f"E{r}"] = f'=IF({sn}="","",SUMIFS({LR("J")},{LR("B")},{sn}))'
        dash[f"F{r}"] = f'=IF(OR({sn}="",N(Settings!$C$26)=0),"",E{r}/Settings!$C$26)'
        dash[f"G{r}"] = f'=IF({sn}="","",IFERROR(AVERAGEIFS({LR("H")},{LR("B")},{sn},{LR("G")},"Done"),""))'
        dash[f"H{r}"] = "=" + LET(f"G{r}")
    style_range(dash, 14, 13 + NS, 1, 8, fmt={2: "0", 3: "0%", 4: "0", 5: "0.0", 6: "0%", 7: "0.0"})
    for c in ("C", "F"):
        dash.conditional_formatting.add(f"{c}14:{c}{13 + NS}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="86EFAC"))
    K.section(dash, "A22", "Hours by subject (all students)")
    K.hdr(dash, 23, ["Subject", "Hours", "Target to date", "Lessons"])
    for j in range(NSUB):
        r = 24 + j; sj = f"Settings!D{SU0 + j}"
        dash[f"A{r}"] = f'=IF({sj}="","",{sj})'
        dash[f"B{r}"] = f'=IF(A{r}="","",SUMIFS({LR("J")},{LR("C")},A{r}))'
        dash[f"C{r}"] = f'=IF(A{r}="","",N(Settings!E{SU0 + j})*Settings!$C$27*COUNTIF({STU},"?*"))'
        dash[f"D{r}"] = f'=IF(A{r}="","",COUNTIFS({LR("C")},A{r},{LR("G")},"Done"))'
    style_range(dash, 24, 23 + NSUB, 1, 4, fmt={2: "0.0", 3: "0.0", 4: "0"})
    K.section(dash, "F22", "Last 12 weeks – hours")
    K.hdr(dash, 23, ["Week of", "Hours", "Lessons"], col0=6)
    for k in range(12):
        r = 24 + k; dash[f"F{r}"] = f"={wk}-{(11 - k) * 7}"
        dash[f"G{r}"] = f'=SUMIFS({LR("J")},{LR("K")},F{r})'
        dash[f"H{r}"] = f'=COUNTIFS({LR("K")},F{r},{LR("G")},"Done")'
    style_range(dash, 24, 35, 6, 8, fmt={6: "mmm d", 7: "0.0", 8: "0"})
    for c, w in zip("ABCDEFGH", [20, 14, 14, 12, 12, 14, 10, 9]): dash.column_dimensions[c].width = w
    ch = BarChart(); ch.type = "bar"; ch.title = "Hours by subject"; ch.height = 8; ch.width = 14; ch.legend = None
    ch.add_data(Reference(dash, min_col=2, min_row=23, max_row=23 + len(SUBJ)), titles_from_data=True)
    ch.set_categories(Reference(dash, min_col=1, min_row=24, max_row=23 + len(SUBJ)))
    ch.x_axis.scaling.orientation = "maxMin"; ch.y_axis.majorGridlines = None; ch.series[0].graphicalProperties.solidFill = MAIN
    dash.add_chart(ch, "J5")
    ch2 = BarChart(); ch2.type = "col"; ch2.title = "Hours per week"; ch2.height = 7.5; ch2.width = 14; ch2.legend = None
    ch2.add_data(Reference(dash, min_col=7, min_row=23, max_row=35), titles_from_data=True)
    ch2.set_categories(Reference(dash, min_col=6, min_row=24, max_row=35)); ch2.y_axis.majorGridlines = None
    ch2.series[0].graphicalProperties.solidFill = "4ADE80"; ch2.gapWidth = 60
    dash.add_chart(ch2, "J22")

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Settings tab: type your students (up to 6), your subjects (up to 15) with a target number of hours per week,", False),
             ("     the first and last day of your school year and what your state / country requires (school days and hours).", False),
             ("2.  Lesson Log: plan lessons ahead with Status 'Planned'. When a lesson is finished, set it to 'Done' and type the minutes.", False),
             ("     Graded work? Type a score from 0 to 100 – the letter grade appears by itself (change the scale in Settings).", False),
             ("3.  Weekly Plan: shows this week's lessons in date order for all children. Type any date to see another week. Print it!", False),
             ("4.  Attendance: school days (Mon–Fri) fill in by themselves. Type P (present), F (field trip), A (absent) or H (holiday).", False),
             ("5.  Progress: hours per subject and child against your target (red = behind, green = on track), average scores, lessons done.", False),
             ("6.  Report Card: pick a child – subjects, hours, scores, grades, attendance and space for comments. Print or save as PDF.", False),
             ("7.  Dashboard: lessons done, hours, average score, overdue planned lessons, days attended vs required, charts.", False),
             ("", False), ("Tips", True),
             ("•  Copy a week of 'Planned' rows and change the dates – planning the next week takes two minutes.", False),
             ("•  Keep the Report Card PDFs and the Attendance tab as your homeschool records / portfolio.", False),
             ("•  Days with F (field trip) count as school days.", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  6 students, 15 subjects, 1,200 lessons, 260 school days. No macros. Works in Excel 2010+ and Google Sheets.", False),
             ("•  Homeschool rules differ by state and country – this planner helps you keep records, it is not legal advice.", False),
             ("•  To use Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("Three made-up students (Emma, Noah, Lily) from August to November 2026. Use the blank copy in your download for your family.", False)]
    K.start(st, "Homeschool Planner & Lesson Tracker", "Plan lessons, log hours and attendance, track grades – Excel and Google Sheets.", lines)
    finish(wb, sample, path)


META = {
    "name": "Homeschool Planner & Lesson Tracker (Multiple Kids, Attendance Days & Hours, Grades, Weekly Plan, Report Card) – Excel & Google Sheets",
    "file_base": "Homeschool-Planner-Lesson-Tracker",
    "price": 17,
    "color": "#15803D",
    "title": "Homeschool Planner",
    "tagline": "lesson log · weekly plan · attendance days & hours · grades · report card",
    "intro": "Plan, record and prove your homeschool year in one spreadsheet. Plan lessons for up to 6 children, mark them done, and the planner adds up hours per subject, school days attended and average grades – compared with your own targets and your state's required days and hours. Print a weekly plan for the fridge and a report card for each child.",
    "features": [
        "Lesson log for 1,200 lessons: date, child, subject, topic, curriculum, minutes, status (Planned / Done / Skipped), score and notes",
        "Weekly Plan: all lessons of any week in date order, ready to print",
        "Attendance for up to 6 children – school days fill in automatically; present, field trip, absent or holiday",
        "Required days and hours tracker for your state or country, with % complete per child",
        "Progress by subject and child: hours vs your weekly target (red when behind), average scores and lessons done",
        "Report card per child with hours, grades, attendance and teacher comments – print or save as PDF for your portfolio",
        "Dashboard with lessons done, hours, average score, overdue lessons, days left and charts",
        "Custom grading scale, up to 15 subjects; blank copy plus a sample copy with three example students",
    ],
    "tags": ["homeschool planner", "homeschool tracker", "lesson planner", "attendance tracker", "homeschool hours", "report card template",
             "google sheets", "excel template", "homeschool record", "lesson plan template", "grade tracker", "school planner", "homeschool"],
    "summary": "Instant download: Excel + Google Sheets template, blank copy + sample copy.",
}
