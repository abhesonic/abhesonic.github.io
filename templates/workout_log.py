"""Fitness & Workout Log with progress charts (sample + blank versions).

Workout Log (one row per set: date, workout, exercise, weight, reps, or duration/distance for cardio)
-> estimated 1-rep max (Epley), volume, automatic PR flags, workout-day counter.
Exercises tab: library (60+ prefilled) with sets logged, best weight, best est. 1RM, last done.
Progress tab: pick an exercise -> 16-week table + chart. Body Stats: weight & measurements.
Dashboard: this week / month, PRs, volume, workouts per week, sets per muscle group, top exercises.
Excel-2010-safe: conditional max via SUMPRODUCT(MAX(...)) instead of MAXIFS.
"""
import random, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, ColorScaleRule
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.utils import get_column_letter as L
from _common import Kit, F, INK, B, DATE, font, fill, style_range, input_cell, finish

MAIN = "EA580C"; SOFT = "FFF1E7"; ACC = "FDBA74"
K = Kit(MAIN, SOFT)
NL = 1500; R0 = 6; R1 = R0 + NL - 1          # log rows
NE = 80; E0 = 6; E1 = E0 + NE - 1            # exercises rows
NB = 200; B0 = 6; B1 = B0 + NB - 1           # body stats rows
WL = "'Workout Log'"; BS = "'Body Stats'"
def lg(c): return f"{WL}!${c}${R0}:${c}${R1}"
AS = "Settings!$B$7"; WK = "Settings!$B$10"; MS = "Settings!$B$11"
W1 = '#,##0.0;-#,##0.0;""'; W0 = '#,##0;-#,##0;""'
GROUPS = ["Chest", "Back", "Shoulders", "Legs", "Glutes", "Biceps", "Triceps", "Core", "Full body", "Cardio"]
LIB = [("Bench Press", "Chest"), ("Incline Dumbbell Press", "Chest"), ("Dumbbell Fly", "Chest"), ("Push-up", "Chest"), ("Chest Dip", "Chest"),
       ("Deadlift", "Back"), ("Barbell Row", "Back"), ("Lat Pulldown", "Back"), ("Pull-up", "Back"), ("Seated Cable Row", "Back"),
       ("One-arm Dumbbell Row", "Back"), ("Overhead Press", "Shoulders"), ("Dumbbell Shoulder Press", "Shoulders"),
       ("Lateral Raise", "Shoulders"), ("Face Pull", "Shoulders"), ("Rear Delt Fly", "Shoulders"), ("Back Squat", "Legs"),
       ("Front Squat", "Legs"), ("Leg Press", "Legs"), ("Romanian Deadlift", "Legs"), ("Walking Lunge", "Legs"),
       ("Bulgarian Split Squat", "Legs"), ("Leg Extension", "Legs"), ("Leg Curl", "Legs"), ("Calf Raise", "Legs"),
       ("Hip Thrust", "Glutes"), ("Glute Bridge", "Glutes"), ("Cable Kickback", "Glutes"), ("Barbell Curl", "Biceps"),
       ("Dumbbell Curl", "Biceps"), ("Hammer Curl", "Biceps"), ("Triceps Pushdown", "Triceps"), ("Skull Crusher", "Triceps"),
       ("Overhead Triceps Extension", "Triceps"), ("Plank", "Core"), ("Hanging Leg Raise", "Core"), ("Cable Crunch", "Core"),
       ("Russian Twist", "Core"), ("Kettlebell Swing", "Full body"), ("Clean and Press", "Full body"), ("Burpee", "Full body"),
       ("Running", "Cardio"), ("Cycling", "Cardio"), ("Rowing Machine", "Cardio"), ("Walking", "Cardio"), ("Swimming", "Cardio"),
       ("Jump Rope", "Cardio"), ("Elliptical", "Cardio"), ("Stair Climber", "Cardio"), ("HIIT", "Cardio")]


def sample_rows():
    random.seed(11)
    plan = {0: ("Push Day", [("Bench Press", 60, 8, 1.25), ("Overhead Press", 35, 8, 0.6), ("Incline Dumbbell Press", 20, 10, 0.5), ("Triceps Pushdown", 25, 12, 0.5)]),
            2: ("Pull Day", [("Deadlift", 90, 5, 2.5), ("Barbell Row", 50, 8, 1.0), ("Lat Pulldown", 45, 10, 0.8), ("Barbell Curl", 22.5, 10, 0.4)]),
            4: ("Leg Day", [("Back Squat", 70, 6, 2.0), ("Romanian Deadlift", 60, 8, 1.25), ("Leg Press", 120, 10, 2.5), ("Plank", 0, 0, 0)])}
    rows = []; start = dt.date(2026, 8, 3); end = dt.date(2026, 10, 30)
    d = start
    while d <= end:
        wk = (d - start).days // 7; wd = d.weekday()
        if wd in plan and random.random() > 0.08:
            name, exs = plan[wd]
            for ex, w0, reps, inc in exs:
                if ex == "Plank":
                    for sn in range(1, 4): rows.append((d, name, ex, sn, None, None, round(1 + wk * 0.08, 2), None, 7, None))
                    continue
                w = round((w0 + inc * wk) / 2.5) * 2.5 if w0 >= 40 else round((w0 + inc * wk) * 2) / 2
                for sn in range(1, 4):
                    r = reps - (1 if sn == 3 and random.random() < 0.5 else 0) + (1 if random.random() < 0.2 else 0)
                    rows.append((d, name, ex, sn, w, r, None, None, 7 + (sn == 3) + (random.random() < 0.3), "Felt strong" if sn == 1 and random.random() < 0.06 else None))
        if wd == 5 and random.random() > 0.25:
            dist = round(4 + wk * 0.12 + random.random(), 1)
            rows.append((d, "Cardio", "Running", 1, None, None, round(dist * (6.3 - wk * 0.04)), dist, 6, "Easy pace"))
        d += dt.timedelta(days=1)
    return rows


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); lw = wb.create_sheet("Workout Log"); ex = wb.create_sheet("Exercises")
    pg = wb.create_sheet("Progress"); bs = wb.create_sheet("Body Stats"); cc = wb.create_sheet("Calc"); s = wb.create_sheet("Settings")

    # ---------- Settings ----------
    K.title(s, "Settings", "Yellow cells are yours to edit.")
    rows = [("Weight unit", "kg", None), ("Distance unit", "km", None), ("Workouts per week goal", 3 if sample else 3, "0"),
            ("Today / as-of date", dt.date(2026, 10, 30) if sample else "=TODAY()", DATE)]
    for i, (k_, v, fmt) in enumerate(rows, 4):
        s.cell(row=i, column=1, value=k_).font = font(bold=True)
        c = s.cell(row=i, column=2); input_cell(c, v)
        if fmt: c.number_format = fmt
    s["C7"] = "Keep =TODAY() so the dashboard always shows this week." + (" (The sample uses a fixed date.)" if sample else "")
    s["C7"].font = font(italic=True, size=9)
    calc = [("Week starts (Monday)", "=B7-WEEKDAY(B7,2)+1", DATE), ("Month starts", "=DATE(YEAR(B7),MONTH(B7),1)", DATE),
            ("Last 30 days from", "=B7-29", DATE)]
    for i, (k_, f, fmt) in enumerate(calc, 10):
        s.cell(row=i, column=1, value=k_).font = font(italic=True)
        c = s.cell(row=i, column=2, value=f); c.number_format = fmt; c.border = B; c.fill = fill("F3F4F6"); c.font = font()
    K.hdr(s, 4, ["Muscle groups"], col0=8)
    for i in range(12):
        c = s.cell(row=5 + i, column=8, value=GROUPS[i] if i < len(GROUPS) else None); input_cell(c)
    K.hdr(s, 4, ["Units"], col0=10)
    for i, u in enumerate(["kg", "lb", "km", "mi"]):
        c = s.cell(row=5 + i, column=10, value=u); c.border = B; c.font = font(size=9)
    for c, w in zip("ABCDEFGHIJ", [26, 14, 3, 3, 3, 3, 3, 16, 3, 8]): s.column_dimensions[c].width = w
    dv = DataValidation(type="list", formula1="=Settings!$J$5:$J$6", allow_blank=False); s.add_data_validation(dv); dv.add("B4")
    dv = DataValidation(type="list", formula1="=Settings!$J$7:$J$8", allow_blank=False); s.add_data_validation(dv); dv.add("B5")
    s["A15"] = "Estimated 1-rep max (est. 1RM) uses the Epley formula: weight × (1 + reps ÷ 30), for sets of 1 to 15 reps."
    s["A15"].font = font(italic=True, size=9)

    # ---------- Workout Log ----------
    K.title(lw, "Workout Log", "One row per set. Strength: weight + reps. Cardio: minutes + distance. Add new rows at the bottom, in date order.")
    lw["A3"] = '="Weights in "&Settings!B4&" · distance in "&Settings!B5'; lw["A3"].font = font(bold=True)
    heads = ["Date", "Workout (e.g. Push Day)", "Exercise", "Set #", "Weight", "Reps", "Minutes", "Distance", "Effort (RPE 1-10)", "Notes",
             "Muscle group", "Volume (weight × reps)", "Est. 1RM", "Week of", "Workout day", "PR", "Exercise day"]
    K.hdr(lw, 5, heads, [11, 16, 24, 6, 9, 6, 8, 9, 8, 18, 12, 11, 9, 11, 8, 8, 8])
    style_range(lw, R0, R1, 1, 17, inputs=range(1, 11), fmt={1: DATE, 5: "0.0#", 7: "0.0#", 8: "0.0#", 12: W0, 13: W1, 14: DATE, 15: '0;-0;""', 17: '0;-0;""'})
    data = sample_rows() if sample else []
    for i, row in enumerate(data):
        for j, v in enumerate(row, 1):
            if v is not None: lw.cell(row=R0 + i, column=j, value=v)
    for r in range(R0, R1 + 1):
        lw[f"K{r}"] = f'=IF(C{r}="","",IFERROR(INDEX(Exercises!$B${E0}:$B${E1},MATCH(C{r},Exercises!$A${E0}:$A${E1},0)),"Other"))'
        lw[f"L{r}"] = f"=IF(AND(N(E{r})>0,N(F{r})>0),E{r}*F{r},0)"
        lw[f"M{r}"] = f"=IF(AND(N(E{r})>0,N(F{r})>0,N(F{r})<=15),ROUND(E{r}*(1+F{r}/30),1),0)"
        lw[f"N{r}"] = f'=IF(N(A{r})=0,0,A{r}-WEEKDAY(A{r},2)+1)'
        lw[f"O{r}"] = f'=IF(N(A{r})=0,0,IF(COUNTIF($A${R0}:A{r},A{r})=1,1,0))'
        cA = f"$A${R0}:$A${R1}"; cC = f"$C${R0}:$C${R1}"; cM = f"$M${R0}:$M${R1}"
        lw[f"P{r}"] = (f'=IF(OR(M{r}=0,N(A{r})=0),"",IF(COUNTIFS({cC},C{r},{cA},"<"&A{r})=0,"",'
                       f'IF(AND(M{r}>SUMPRODUCT(MAX(({cC}=C{r})*({cA}<A{r})*{cM})),M{r}=SUMPRODUCT(MAX(({cC}=C{r})*({cA}=A{r})*{cM})),'
                       f'COUNTIFS($C${R0}:C{r},C{r},$A${R0}:A{r},A{r},$M${R0}:M{r},M{r})=1),"PR ★","")))')
        lw[f"Q{r}"] = f'=IF(OR(C{r}="",N(A{r})=0),0,IF(COUNTIFS($C${R0}:C{r},C{r},$A${R0}:A{r},A{r})=1,1,0))'
        for c in "DFIOP": lw[f"{c}{r}"].alignment = Alignment(horizontal="center")
    dv = DataValidation(type="list", formula1=f"=Exercises!$A${E0}:$A${E1}", allow_blank=True, showErrorMessage=False); lw.add_data_validation(dv); dv.add(f"C{R0}:C{R1}")
    dv = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True, error="Type a number", showErrorMessage=True)
    lw.add_data_validation(dv); dv.add(f"E{R0}:H{R1}")
    dv = DataValidation(type="whole", operator="between", formula1="1", formula2="10", allow_blank=True, error="Type 1 to 10", showErrorMessage=True)
    lw.add_data_validation(dv); dv.add(f"I{R0}:I{R1}")
    lw.conditional_formatting.add(f"P{R0}:P{R1}", FormulaRule(formula=[f'P{R0}="PR ★"'], fill=fill("FDE68A"), font=Font(name=F, bold=True, color="92400E")))
    lw.conditional_formatting.add(f"A{R0}:C{R1}", FormulaRule(formula=[f'AND($A{R0}<>"",$A{R0}<>$A{R0-1},$O{R0}=1)'], font=Font(name=F, bold=True, color=INK)))
    lw.freeze_panes = "D6"

    # ---------- Exercises ----------
    K.title(ex, "Exercises", "Your exercise library. Add your own in the empty yellow rows. Grey columns show your records.")
    ex["A3"] = '="Weights in "&Settings!B4'; ex["A3"].font = font(bold=True)
    K.hdr(ex, 5, ["Exercise", "Muscle group", "Note / cue", "Sets logged", "Last done", "Best weight", "Best est. 1RM",
                  "Total volume", "Cardio minutes", "Cardio distance", "Rank key"], [26, 13, 24, 8, 11, 9, 9, 11, 9, 9, 8])
    style_range(ex, E0, E1, 1, 11, inputs=(1, 2, 3), fmt={5: DATE, 6: W1, 7: W1, 8: W0, 9: W0, 10: W1, 11: "0.000"})
    for i, (n, g) in enumerate(LIB):
        ex[f"A{E0+i}"] = n; ex[f"B{E0+i}"] = g
    for r in range(E0, E1 + 1):
        on = f'A{r}=""'
        ex[f"D{r}"] = f'=IF({on},0,COUNTIF({lg("C")},A{r}))'
        ex[f"E{r}"] = f'=IF(D{r}=0,"",SUMPRODUCT(MAX(({lg("C")}=A{r})*{lg("A")})))'
        ex[f"F{r}"] = f'=IF(D{r}=0,0,SUMPRODUCT(MAX(({lg("C")}=A{r})*{lg("E")})))'
        ex[f"G{r}"] = f'=IF(D{r}=0,0,SUMPRODUCT(MAX(({lg("C")}=A{r})*{lg("M")})))'
        ex[f"H{r}"] = f'=IF(D{r}=0,0,SUMIFS({lg("L")},{lg("C")},A{r}))'
        ex[f"I{r}"] = f'=IF(D{r}=0,0,SUMIFS({lg("G")},{lg("C")},A{r}))'
        ex[f"J{r}"] = f'=IF(D{r}=0,0,SUMIFS({lg("H")},{lg("C")},A{r}))'
        ex[f"K{r}"] = f"=IF(D{r}>0,D{r}+MIN(G{r},9999)/100000+ROW()/1000000000,0)"
    dv = DataValidation(type="list", formula1="=Settings!$H$5:$H$16", allow_blank=True); ex.add_data_validation(dv); dv.add(f"B{E0}:B{E1}")
    ex.freeze_panes = "B6"

    # ---------- Body Stats ----------
    K.title(bs, "Body Stats", "Weigh in and measure as often as you like (weekly is plenty). One row per check-in, in date order.")
    bs["A3"] = '="Body weight in "&Settings!B4&" · measurements in your own unit (cm or in)"'; bs["A3"].font = font(bold=True)
    K.hdr(bs, 5, ["Date", "Body weight", "Body fat %", "Waist", "Chest", "Arms", "Hips", "Thighs", "Notes", "Weight change since start", "Waist change since start"],
          [11, 10, 9, 8, 8, 8, 8, 8, 22, 12, 12])
    style_range(bs, B0, B1, 1, 11, inputs=range(1, 10), fmt={1: DATE, 2: "0.0", 3: "0.0%", 10: '+0.0;-0.0;0.0', 11: '+0.0;-0.0;0.0'})
    if sample:
        random.seed(3); d = dt.date(2026, 8, 1); w = 82.4; waist = 89.0; i = 0
        while d <= dt.date(2026, 10, 30):
            r = B0 + i; bs[f"A{r}"] = d; bs[f"B{r}"] = round(w, 1)
            if i % 2 == 0:
                bs[f"C{r}"] = round(0.21 - i * 0.0018, 3); bs[f"D{r}"] = round(waist, 1); bs[f"E{r}"] = round(102 + i * 0.1, 1)
                bs[f"F{r}"] = round(36 + i * 0.07, 1); bs[f"G{r}"] = 98.0 - i * 0.05; bs[f"H{r}"] = round(57 + i * 0.05, 1)
            w += random.uniform(-0.55, 0.25); waist -= random.uniform(0.0, 0.35); d += dt.timedelta(days=7); i += 1
    for r in range(B0, B1 + 1):
        bs[f"J{r}"] = f'=IF(OR(N(B{r})=0,N($B${B0})=0),"",B{r}-$B${B0})'
        bs[f"K{r}"] = f'=IF(OR(N(D{r})=0,N($D${B0})=0),"",D{r}-$D${B0})'
    bs.freeze_panes = "B6"
    BA, BW = f"{BS}!$A${B0}:$A${B1}", f"{BS}!$B${B0}:$B${B1}"

    # ---------- Calc ----------
    K.title(cc, "Calc", "Numbers for the charts. Calculated – nothing to type.")
    K.hdr(cc, 4, ["Week of", "Label", "Workouts", "Sets", "Volume", "Body weight", "Goal"], [11, 9, 9, 8, 11, 11, 7])
    for i in range(12):
        r = 5 + i
        cc[f"A{r}"] = f"={WK}-7*{11 - i}"
        cc[f"B{r}"] = f'=TEXT(A{r},"d mmm")'
        cc[f"C{r}"] = f'=SUMIFS({lg("O")},{lg("N")},A{r})'
        cc[f"D{r}"] = f'=COUNTIFS({lg("N")},A{r},{lg("C")},"<>")'
        cc[f"E{r}"] = f'=SUMIFS({lg("L")},{lg("N")},A{r})'
        prev = f"F{r-1}" if i else f'IF(SUMPRODUCT(MAX(({BA}<A{r})*({BW}>0)*{BA}))=0,N({BS}!$B${B0}),AVERAGEIFS({BW},{BA},SUMPRODUCT(MAX(({BA}<A{r})*({BW}>0)*{BA}))))'
        cc[f"F{r}"] = f'=IF(COUNTIFS({BA},">="&A{r},{BA},"<"&A{r}+7,{BW},">0")>0,AVERAGEIFS({BW},{BA},">="&A{r},{BA},"<"&A{r}+7,{BW},">0"),{prev})'
        cc[f"G{r}"] = "=Settings!$B$6"
    style_range(cc, 5, 16, 1, 7, fmt={1: DATE, 5: W0, 6: "0.0"})
    K.hdr(cc, 19, ["Muscle group", "Sets (last 30 days)"], [11, 9])
    for i in range(12):
        r = 20 + i
        cc[f"A{r}"] = f'=IF(Settings!$H${5+i}="","",Settings!$H${5+i})'
        cc[f"B{r}"] = f'=IF(A{r}="",0,COUNTIFS({lg("K")},A{r},{lg("A")},">="&Settings!$B$12,{lg("A")},"<="&{AS}))'
    style_range(cc, 20, 31, 1, 2)
    cc.column_dimensions["A"].width = 14

    # ---------- Progress ----------
    K.title(pg, "Exercise Progress", "Pick an exercise in the yellow cell to see the last 16 weeks.")
    pg["A3"] = "Exercise:"; pg["A3"].font = font(bold=True)
    top1 = f'=IFERROR(INDEX(Exercises!$A${E0}:$A${E1},MATCH(LARGE(Exercises!$K${E0}:$K${E1},1),Exercises!$K${E0}:$K${E1},0)),"Bench Press")'
    input_cell(pg["B3"], "Bench Press" if sample else top1)
    dv = DataValidation(type="list", formula1=f"=Exercises!$A${E0}:$A${E1}", allow_blank=False); pg.add_data_validation(dv); dv.add("B3")
    pg.merge_cells("B3:C3")
    cA, cC = lg("A"), lg("C")
    stats = [("Best est. 1RM", f"=IFERROR(INDEX(Exercises!$G${E0}:$G${E1},MATCH(B3,Exercises!$A${E0}:$A${E1},0)),0)", W1),
             ("Best weight", f"=IFERROR(INDEX(Exercises!$F${E0}:$F${E1},MATCH(B3,Exercises!$A${E0}:$A${E1},0)),0)", W1),
             ("Sets logged", f'=COUNTIF({cC},B3)', "0"),
             ("PRs", f'=COUNTIFS({cC},B3,{lg("P")},"PR ★")', "0")]
    for i, (lab, f, fmt) in enumerate(stats):
        K.kpi(pg, 5, 1 + i * 2, lab, f, fmt)
    for c, w in zip("ABCDEFGHI", [12, 11, 11, 11, 11, 11, 12, 12, 11]): pg.column_dimensions[c].width = w
    K.hdr(pg, 9, ["Week of", "Sessions", "Sets", "Best weight", "Top reps at best weight", "Best est. 1RM", "Volume", "Est. 1RM trend", "Minutes"])
    for i in range(16):
        r = 10 + i
        pg[f"A{r}"] = f"={WK}-7*{15 - i}"
        wkc = f'({lg("N")}=A{r})'; exc = f"({cC}=$B$3)"
        pg[f"B{r}"] = f'=SUMIFS({lg("Q")},{cC},$B$3,{lg("N")},A{r})'
        pg[f"C{r}"] = f'=COUNTIFS({cC},$B$3,{lg("N")},A{r})'
        pg[f"D{r}"] = f"=SUMPRODUCT(MAX({exc}*{wkc}*{lg('E')}))"
        pg[f"E{r}"] = f'=IF(D{r}=0,0,SUMPRODUCT(MAX({exc}*{wkc}*({lg("E")}=D{r})*{lg("F")})))'
        pg[f"F{r}"] = f"=SUMPRODUCT(MAX({exc}*{wkc}*{lg('M')}))"
        pg[f"G{r}"] = f'=SUMIFS({lg("L")},{cC},$B$3,{lg("N")},A{r})'
        prev = f"H{r-1}" if i else f"SUMPRODUCT(MAX({exc}*({cA}<A{r})*({cA}>0)*{lg('M')}))"
        pg[f"H{r}"] = f"=IF(F{r}>0,F{r},{prev})"
        pg[f"I{r}"] = f'=SUMIFS({lg("G")},{cC},$B$3,{lg("N")},A{r})'
    style_range(pg, 10, 25, 1, 9, fmt={1: "d mmm yy", 2: '0;-0;""', 3: '0;-0;""', 4: W1, 5: '0;-0;""', 6: W1, 7: W0, 8: W1, 9: W0})
    ch = LineChart(); ch.title = "Estimated 1-rep max (weekly best, carried forward)"; ch.height = 7.5; ch.width = 20; ch.legend = None
    ch.add_data(Reference(pg, min_col=8, min_row=9, max_row=25), titles_from_data=True)
    ch.set_categories(Reference(pg, min_col=1, min_row=10, max_row=25))
    ch.series[0].graphicalProperties.line.solidFill = MAIN; ch.series[0].graphicalProperties.line.width = 28000; ch.series[0].smooth = False
    ch.x_axis.number_format = "d mmm"; ch.y_axis.majorGridlines = None
    pg.add_chart(ch, "A28")
    pg["A27"] = "Weeks without this exercise keep the last value, so the line shows your strength trend."; pg["A27"].font = font(italic=True, size=9)

    # ---------- Dashboard ----------
    K.title(dash, "Workout Dashboard", "Updates by itself as you log sets on the Workout Log tab.")
    dash["A3"] = f'="As of "&TEXT({AS},"dd mmm yyyy")&"  ·  weights in "&Settings!B4&"  ·  goal "&Settings!B6&" workouts a week"'
    dash["A3"].font = font(bold=True)
    inrange = lambda a, b: f'{lg("A")},">="&{a},{lg("A")},"<="&{b}'
    lastw = f'IFERROR(INDEX({BW},MATCH(9.99E+307,{BW})),0)'
    kp = [("Workouts this week", f'=SUMIFS({lg("O")},{inrange(WK, AS)})&" of "&Settings!B6', "@"),
          ("Workouts this month", f'=SUMIFS({lg("O")},{inrange(MS, AS)})', "0"),
          ("Sets this month", f'=COUNTIFS({inrange(MS, AS)},{lg("C")},"<>")', "0"),
          ("Volume this month", f'=SUMIFS({lg("L")},{inrange(MS, AS)})', '#,##0'),
          ("New PRs this month", f'=COUNTIFS({inrange(MS, AS)},{lg("P")},"PR ★")', "0"),
          ("Cardio minutes this month", f'=SUMIFS({lg("G")},{inrange(MS, AS)})', "0"),
          ("Body weight now", f'={lastw}', '0.0;-0.0;"—"'),
          ("Change since start", f'=IF(OR({lastw}=0,N({BS}!$B${B0})=0),"—",{lastw}-{BS}!$B${B0})', '+0.0;-0.0;0.0'),
          ("Workout days in total", f'=SUM({lg("O")})', "0")]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 3) * 3, 1 + (i % 3) * 2, lab, f, fmt)
        dash.cell(row=6 + (i // 3) * 3, column=1 + (i % 3) * 2).alignment = Alignment(horizontal="left")
    for c, w in zip("ABCDEFGHIJ", [24, 13, 13, 13, 13, 13, 4, 13, 13, 13]): dash.column_dimensions[c].width = w
    K.section(dash, "A15", "Your top exercises (most sets logged)")
    K.hdr(dash, 16, ["Exercise", "Muscle group", "Sets logged", "Best weight", "Best est. 1RM", "Last done"])
    for n in range(10):
        r = 17 + n
        m = f"MATCH(LARGE(Exercises!$K${E0}:$K${E1},{n+1}),Exercises!$K${E0}:$K${E1},0)"
        dash[f"A{r}"] = f'=IF(LARGE(Exercises!$K${E0}:$K${E1},{n+1})=0,"",INDEX(Exercises!$A${E0}:$A${E1},{m}))'
        for c, src in zip("BCDEF", "BDFGE"):
            dash[f"{c}{r}"] = f'=IF($A{r}="","",INDEX(Exercises!${src}${E0}:${src}${E1},{m}))'
    style_range(dash, 17, 26, 1, 6, fmt={4: W1, 5: W1, 6: "dd mmm yyyy"})
    for r in range(17, 27):
        for c in "CDEF": dash[f"{c}{r}"].alignment = Alignment(horizontal="center")
    dash.conditional_formatting.add("E17:E26", ColorScaleRule(start_type="min", start_color="FFFFFF", end_type="max", end_color=ACC))
    K.section(dash, "H15", "Sets per muscle group (30 days)")
    K.hdr(dash, 16, ["Muscle group", "Sets"], col0=8)
    for i in range(10):
        dash[f"H{17+i}"] = f"=Calc!A{20+i}"; dash[f"I{17+i}"] = f'=IF(H{17+i}="","",Calc!B{20+i})'
    style_range(dash, 17, 26, 8, 9)
    dash.conditional_formatting.add("I17:I26", ColorScaleRule(start_type="num", start_value=0, start_color="FFFFFF", end_type="max", end_color=ACC))
    c1 = BarChart(); c1.type = "col"; c1.title = "Workouts per week (last 12 weeks)"; c1.height = 7; c1.width = 15
    c1.add_data(Reference(cc, min_col=3, min_row=4, max_row=16), titles_from_data=True)
    c1.set_categories(Reference(cc, min_col=2, min_row=5, max_row=16))
    c1.series[0].graphicalProperties.solidFill = MAIN; c1.legend = None; c1.y_axis.majorGridlines = None; c1.y_axis.scaling.min = 0
    dash.add_chart(c1, "A29")
    c2 = BarChart(); c2.type = "col"; c2.title = "Training volume per week"; c2.height = 7; c2.width = 15
    c2.add_data(Reference(cc, min_col=5, min_row=4, max_row=16), titles_from_data=True)
    c2.set_categories(Reference(cc, min_col=2, min_row=5, max_row=16))
    c2.series[0].graphicalProperties.solidFill = ACC; c2.legend = None; c2.y_axis.majorGridlines = None
    dash.add_chart(c2, "E29")
    c3 = LineChart(); c3.title = "Body weight (weekly average)"; c3.height = 7; c3.width = 15; c3.legend = None
    c3.add_data(Reference(cc, min_col=6, min_row=4, max_row=16), titles_from_data=True)
    c3.set_categories(Reference(cc, min_col=2, min_row=5, max_row=16))
    c3.series[0].graphicalProperties.line.solidFill = MAIN; c3.series[0].graphicalProperties.line.width = 28000; c3.series[0].smooth = False
    c3.y_axis.majorGridlines = None
    dash.add_chart(c3, "A44")
    c4 = BarChart(); c4.type = "bar"; c4.title = "Sets per muscle group (last 30 days)"; c4.height = 7; c4.width = 15; c4.legend = None
    c4.add_data(Reference(cc, min_col=2, min_row=19, max_row=29), titles_from_data=True)
    c4.set_categories(Reference(cc, min_col=1, min_row=20, max_row=29))
    c4.series[0].graphicalProperties.solidFill = MAIN; c4.x_axis.scaling.orientation = "maxMin"; c4.y_axis.majorGridlines = None
    dash.add_chart(c4, "E44")

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Settings tab: choose kg or lb, km or mi, and how many workouts a week you aim for.", False),
             ("2.  Workout Log tab: one row per set – date, workout name, exercise (pick from the list), set number, weight and reps. For cardio, type minutes and distance instead.", False),
             ("3.  Dashboard: workouts this week vs your goal, sets, volume, new personal records, body weight, top exercises and charts.", False),
             ("4.  Progress tab: pick an exercise to see your best weight, estimated 1-rep max and volume week by week, with a chart.", False),
             ("5.  Body Stats tab: log your weight and measurements (waist, chest, arms…) to see your change since you started.", False),
             ("", False), ("How it works", True),
             ("Volume = weight × reps. Est. 1RM = weight × (1 + reps ÷ 30) (Epley formula) – a safe way to see strength gains without testing a real max.", False),
             ("A set is marked PR ★ when its est. 1RM beats every earlier day for that exercise (one PR per exercise per day).", False),
             ("Exercises tab: 50 common exercises are already there, grouped by muscle. Add your own in the empty yellow rows.", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  Log in date order (new rows at the bottom). Room for 1,500 sets – about a year of training at 4 workouts a week.", False),
             ("•  Google Sheets on your phone works well at the gym: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False),
             ("•  This log is a tracking tool, not medical or training advice.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("A made-up 3-day push / pull / legs program from August to October 2026, shown as of 30 Oct 2026. Use the blank copy in your download for your own training.", False)]
    K.start(st, "Fitness & Workout Log", "Log every set, see your PRs, strength trend, volume and body stats. Works in Excel and Google Sheets.", lines)
    finish(wb, sample, path)


META = {
    "name": "Fitness & Workout Log with Progress Charts – Excel & Google Sheets",
    "file_base": "Workout-Log-Progress-Tracker",
    "price": 15,
    "color": "#EA580C",
    "title": "Workout Log & Progress Tracker",
    "tagline": "Sets & reps · automatic PRs · strength charts · body stats — Excel & Google Sheets",
    "intro": "See your strength go up. Log each set in seconds and the workout log finds your personal records, tracks your estimated 1-rep max, volume and body weight, and charts your progress week by week.",
    "features": [
        "Dashboard: workouts this week vs your goal, sets, volume, new PRs, cardio minutes and body weight change",
        "Automatic PR ★ flags and estimated 1-rep max (Epley) for every set",
        "Progress tab: pick any exercise to see 16 weeks of best weight, est. 1RM, volume and a trend chart",
        "Charts: workouts per week, training volume, body weight and sets per muscle group",
        "50 exercises pre-loaded by muscle group, add your own; strength and cardio (minutes, distance)",
        "Body Stats: weight, body fat %, waist, chest, arms, hips and thighs with change since start",
        "kg or lb, km or mi; room for 1,500 sets; works great on your phone in Google Sheets",
    ],
    "tags": ["workout log", "workout tracker", "fitness tracker", "gym log", "google sheets", "excel template", "strength training",
             "progressive overload", "workout planner", "lifting log", "fitness planner", "body measurement tracker", "1rm calculator"],
}
