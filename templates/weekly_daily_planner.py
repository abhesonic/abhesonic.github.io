"""Weekly & Daily Planner – tasks with priorities, appointments, weekly goals, habit check-offs,
a 7-day week view (Dashboard) and a printable Daily Page.

Settings -> Tasks / Appointments / Weekly Goals / Habit Log -> Dashboard (week view) + Daily Page.
"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, DataBarRule
from openpyxl.chart import BarChart, Reference
from _common import Kit, F, B, DATE, font, fill, style_range, input_cell, finish

MAIN = "0F766E"; SOFT = "CCFBF1"
K = Kit(MAIN, SOFT)
NT = 1000; T0 = 6; T1 = T0 + NT - 1
NA = 500; A0 = 6; A1 = A0 + NA - 1
NG = 300; G0 = 6; G1 = G0 + NG - 1
NH = 3000; H0 = 6; H1 = H0 + NH - 1
NS = 12; S0 = 6; S1 = S0 + NS - 1
PCT = "0%"; TIME = "h:mm AM/PM"
CATS = ["Work", "Home", "Family", "Health", "Errands", "Money", "Learning", "Personal"]
GREYF = dict(color="9CA3AF", size=7)


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); dp = wb.create_sheet("Daily Page"); tk = wb.create_sheet("Tasks")
    ap = wb.create_sheet("Appointments"); wg = wb.create_sheet("Weekly Goals"); hl = wb.create_sheet("Habit Log")
    s = wb.create_sheet("Settings")

    # ---------- Settings ----------
    K.title(s, "Settings", "Yellow cells are yours to edit.")
    for i, (lab, v) in enumerate([("Your name (shown on the planner)", None), ("Week starts on (Monday or Sunday)", "Monday")], 5):
        s[f"A{i}"] = lab; s[f"A{i}"].font = font(bold=True); s[f"A{i}"].border = B; input_cell(s[f"C{i}"], v)
    dv = DataValidation(type="list", formula1='"Monday,Sunday"', allow_blank=False); s.add_data_validation(dv); dv.add("C6")
    s["A8"] = "Weekday type (automatic)"; s["A8"].font = font(italic=True)
    s["C8"] = '=IF(C6="Sunday",1,2)'; s["C8"].fill = fill("F3F4F6"); s["C8"].border = B; s["C8"].font = font()
    K.hdr(s, 5, ["Categories"], col0=6); K.hdr(s, 5, ["Habits", "Days per week (target)"], col0=8)
    habits = ["Drink 8 glasses of water", "Exercise 30 min", "Read 20 pages", "No phone after 10 pm", "Plan tomorrow"]
    targets = [7, 4, 5, 6, 5]
    for i in range(NS):
        input_cell(s[f"F{S0 + i}"], CATS[i] if i < len(CATS) else None)
        input_cell(s[f"H{S0 + i}"], habits[i] if sample and i < len(habits) else None)
        input_cell(s[f"I{S0 + i}"], targets[i] if sample and i < len(habits) else None)
    for c, w in zip("ABCDEFGHI", [40, 3, 14, 3, 3, 20, 3, 30, 12]): s.column_dimensions[c].width = w
    WT = "Settings!$C$8"

    # week start shown on the dashboard
    WS = "Dashboard!$C$4"; DD = "'Daily Page'!$C$4"
    TK = lambda c: f"Tasks!${c}${T0}:${c}${T1}"
    AP = lambda c: f"Appointments!${c}${A0}:${c}${A1}"
    WG = lambda c: f"'Weekly Goals'!${c}${G0}:${c}${G1}"
    HL = lambda c: f"'Habit Log'!${c}${H0}:${c}${H1}"
    DAYC = "JKLMNOP"  # day-key helper columns (Tasks and Appointments)

    # ---------- Tasks ----------
    K.title(tk, "Tasks", "One row per to-do. Give it a date to put it on that day of the week view; leave the date empty for 'someday'. Choose Yes under Done when finished.")
    cols = ["Task", "Date", "Priority", "Category", "Est. minutes", "Done?", "Notes", "Status", "Week starts"] + \
           [f"(day {d + 1})" for d in range(7)] + ["(daily)", "(overdue)", "(prio)"]
    K.hdr(tk, 5, cols, [42, 12, 10, 13, 9, 7, 30, 11, 11] + [5] * 10)
    style_range(tk, T0, T1, 1, 19, inputs=range(1, 8), fmt={2: DATE, 5: "0", 9: DATE})
    for r in range(T0, T1 + 1):
        tk[f"H{r}"] = (f'=IF(A{r}="","",IF(F{r}="Yes","Done",IF(NOT(ISNUMBER(B{r})),"Someday",'
                       f'IF(B{r}<TODAY(),"Overdue",IF(B{r}=TODAY(),"Today","Planned")))))')
        tk[f"I{r}"] = f'=IF(OR(A{r}="",NOT(ISNUMBER(B{r}))),"",B{r}-WEEKDAY(B{r},{WT})+1)'
        tk[f"S{r}"] = f'=IF(C{r}="High",1,IF(C{r}="Low",3,2))'
        for d, c in enumerate(DAYC):
            tk[f"{c}{r}"] = f'=IF(AND(A{r}<>"",ISNUMBER(B{r})),IF(B{r}={WS}+{d},S{r}*100000+ROW(),""),"")'
        tk[f"Q{r}"] = f'=IF(AND(A{r}<>"",ISNUMBER(B{r})),IF(B{r}={DD},S{r}*100000+ROW(),""),"")'
        tk[f"R{r}"] = f'=IF(H{r}="Overdue",B{r}+ROW()/10000000,"")'
        for c in DAYC + "QRS": tk[f"{c}{r}"].font = font(**GREYF)
        tk[f"H{r}"].font = font(bold=True)
    dv = DataValidation(type="list", formula1='"High,Medium,Low"', allow_blank=True); tk.add_data_validation(dv); dv.add(f"C{T0}:C{T1}")
    dv = DataValidation(type="list", formula1=f"=Settings!$F${S0}:$F${S1}", allow_blank=True); tk.add_data_validation(dv); dv.add(f"D{T0}:D{T1}")
    dv = DataValidation(type="list", formula1='"Yes"', allow_blank=True); tk.add_data_validation(dv); dv.add(f"F{T0}:F{T1}")
    tk.conditional_formatting.add(f"H{T0}:H{T1}", FormulaRule(formula=[f'H{T0}="Overdue"'], font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    tk.conditional_formatting.add(f"H{T0}:H{T1}", FormulaRule(formula=[f'H{T0}="Today"'], font=Font(name=F, bold=True, color="92400E"), fill=fill("FEF3C7")))
    tk.conditional_formatting.add(f"C{T0}:C{T1}", FormulaRule(formula=[f'C{T0}="High"'], font=Font(name=F, bold=True, color="B91C1C")))
    tk.conditional_formatting.add(f"A{T0}:I{T1}", FormulaRule(formula=[f'$F{T0}="Yes"'], font=Font(name=F, color="9CA3AF", strike=True)))
    tk.freeze_panes = f"B{T0}"; tk.auto_filter.ref = f"A5:I{T1}"

    # ---------- Appointments ----------
    K.title(ap, "Appointments", "Meetings, classes, doctor visits, birthdays – anything at a set time. They appear on the week view and the Daily Page.")
    K.hdr(ap, 5, ["Date", "Start time", "End time", "What", "Where / who", "Notes", "Day", "", ""] + [f"(day {d + 1})" for d in range(7)] + ["(daily)"],
          [12, 11, 11, 34, 22, 28, 8, 2, 2] + [5] * 8)
    style_range(ap, A0, A1, 1, 7, inputs=range(1, 7), fmt={1: DATE, 2: TIME, 3: TIME})
    for r in range(A0, A1 + 1):
        ap[f"G{r}"] = f'=IF(ISNUMBER(A{r}),TEXT(A{r},"ddd"),"")'
        for d, c in enumerate(DAYC):
            ap[f"{c}{r}"] = f'=IF(AND(D{r}<>"",ISNUMBER(A{r})),IF(A{r}={WS}+{d},N(B{r})+ROW()/10000000,""),"")'
        ap[f"Q{r}"] = f'=IF(AND(D{r}<>"",ISNUMBER(A{r})),IF(A{r}={DD},N(B{r})+ROW()/10000000,""),"")'
        for c in DAYC + "Q": ap[f"{c}{r}"].font = font(**GREYF)
    ap.freeze_panes = f"A{A0}"; ap.auto_filter.ref = f"A5:G{A1}"

    # ---------- Weekly Goals ----------
    K.title(wg, "Weekly Goals", "Your 1–5 big things for a week. Type any date in that week; the week view shows the goals for the week you are looking at.")
    K.hdr(wg, 5, ["Any date in the week", "Goal", "Done?", "Week starts", "(key)"], [14, 50, 8, 12, 6])
    style_range(wg, G0, G1, 1, 5, inputs=(1, 2, 3), fmt={1: DATE, 4: DATE})
    for r in range(G0, G1 + 1):
        wg[f"D{r}"] = f'=IF(OR(B{r}="",NOT(ISNUMBER(A{r}))),"",A{r}-WEEKDAY(A{r},{WT})+1)'
        wg[f"E{r}"] = f'=IF(D{r}="","",IF(D{r}={WS},ROW(),""))'; wg[f"E{r}"].font = font(**GREYF)
    dv = DataValidation(type="list", formula1='"Yes"', allow_blank=True); wg.add_data_validation(dv); dv.add(f"C{G0}:C{G1}")
    wg.conditional_formatting.add(f"A{G0}:D{G1}", FormulaRule(formula=[f'$C{G0}="Yes"'], font=Font(name=F, color="9CA3AF", strike=True)))
    wg.freeze_panes = f"A{G0}"

    # ---------- Habit Log ----------
    K.title(hl, "Habit Log", "One row each time you do a habit (date + habit). The habit grids on the Dashboard and Daily Page tick themselves.")
    K.hdr(hl, 5, ["Date", "Habit", "Note"], [12, 32, 40])
    style_range(hl, H0, H1, 1, 3, inputs=(1, 2, 3), fmt={1: DATE})
    dv = DataValidation(type="list", formula1=f"=Settings!$H${S0}:$H${S1}", allow_blank=True); hl.add_data_validation(dv); dv.add(f"B{H0}:B{H1}")
    hl.freeze_panes = f"A{H0}"

    # ---------- sample ----------
    D = dt.date
    if sample:
        s["C5"] = "Alex"
        mon = D(2026, 11, 16)
        tasks = [  # (task, day offset or None, prio, cat, minutes, done, notes)
            ("Finish Q4 report draft", 0, "High", "Work", 120, "Yes", ""), ("Reply to client emails", 0, "Medium", "Work", 30, "Yes", ""),
            ("Book dentist appointment", 0, "Low", "Health", 10, "Yes", ""), ("Grocery shopping", 0, "Medium", "Errands", 60, "Yes", "List on fridge"),
            ("Team meeting prep", 1, "High", "Work", 45, "Yes", ""), ("Pay electricity bill", 1, "High", "Money", 10, "Yes", ""),
            ("Laundry", 1, "Low", "Home", 40, None, ""), ("Call Mom", 1, "Medium", "Family", 20, "Yes", ""),
            ("Send Q4 report to manager", 2, "High", "Work", 20, None, ""), ("Update budget spreadsheet", 2, "Medium", "Money", 30, None, ""),
            ("Online course – lesson 4", 2, "Low", "Learning", 45, None, ""), ("Pick up dry cleaning", 2, "Low", "Errands", 15, None, ""),
            ("Plan Thanksgiving menu", 2, "Medium", "Family", 30, None, "8 people"),
            ("1:1 with Jordan", 3, "High", "Work", 30, None, ""), ("Order Thanksgiving turkey", 3, "High", "Errands", 15, None, "Pick up Wed"),
            ("Clean out the fridge", 3, "Low", "Home", 30, None, ""), ("Gym – leg day", 3, "Medium", "Health", 60, None, ""),
            ("Submit expense report", 4, "High", "Work", 20, None, ""), ("Kids' school forms", 4, "Medium", "Family", 15, None, ""),
            ("Vacuum + mop", 4, "Low", "Home", 45, None, ""), ("Movie night", 4, "Low", "Family", 120, None, ""),
            ("Farmers market", 5, "Medium", "Errands", 60, None, ""), ("Long run 8 km", 5, "Medium", "Health", 50, None, ""),
            ("Clean the garage", 5, "Low", "Home", 120, None, ""), ("Meal prep for the week", 6, "Medium", "Home", 90, None, ""),
            ("Weekly review + plan next week", 6, "High", "Personal", 30, None, ""), ("Read book club novel", 6, "Low", "Learning", 60, None, ""),
            ("Renew car registration", -3, "High", "Money", 20, None, "Due end of month"), ("Fix leaky faucet", -5, "Medium", "Home", 40, None, ""),
            ("Write thank-you cards", -2, "Low", "Family", 30, "Yes", ""),
            ("Christmas gift list", 8, "Medium", "Family", 30, None, ""), ("Black Friday – buy new laptop?", 11, "Low", "Money", 30, None, "Compare prices first"),
            ("Learn Spanish basics", None, "Low", "Learning", None, None, "Someday"), ("Repaint the bedroom", None, "Low", "Home", None, None, "Spring?")]
        for k, (t, off, p, c, m, dn, n) in enumerate(tasks):
            r = T0 + k; tk[f"A{r}"] = t
            if off is not None: tk[f"B{r}"] = mon + dt.timedelta(days=off)
            tk[f"C{r}"] = p; tk[f"D{r}"] = c
            if m: tk[f"E{r}"] = m
            if dn: tk[f"F{r}"] = dn
            if n: tk[f"G{r}"] = n
        T = lambda h, mi=0: dt.time(h, mi)
        appts = [(0, T(9), T(9, 30), "Daily stand-up", "Zoom", ""), (0, T(18, 30), T(19, 30), "Yoga class", "Studio 5", ""),
                 (1, T(9), T(9, 30), "Daily stand-up", "Zoom", ""), (1, T(14), T(15), "Team meeting", "Room B", ""),
                 (2, T(9), T(9, 30), "Daily stand-up", "Zoom", ""), (2, T(12, 30), T(13, 30), "Lunch with Priya", "Café Rosa", ""),
                 (2, T(16), T(16, 45), "Parent-teacher meeting", "School", "Bring report card"),
                 (3, T(9), T(9, 30), "Daily stand-up", "Zoom", ""), (3, T(15), T(15, 30), "1:1 with Jordan", "Office", ""),
                 (4, T(8, 15), T(9), "Dentist", "Dr. Lee", ""), (4, T(19), T(21), "Movie night", "Home", ""),
                 (5, T(10), T(11, 30), "Soccer game (Max)", "City park", ""), (6, T(17), T(20), "Dinner at Sam & Jo's", "Their place", "Bring dessert"),
                 (9, T(11), T(12), "Haircut", "Salon", "")]
        for k, (off, a, b, w, wh, n) in enumerate(appts):
            r = A0 + k; ap[f"A{r}"] = mon + dt.timedelta(days=off); ap[f"B{r}"] = a; ap[f"C{r}"] = b; ap[f"D{r}"] = w; ap[f"E{r}"] = wh
            if n: ap[f"F{r}"] = n
        goals = [(mon, "Send the Q4 report", None), (mon, "Exercise 4 times", None), (mon, "Get Thanksgiving shopping done", None),
                 (mon - dt.timedelta(days=7), "Finish onboarding new hire", "Yes"), (mon - dt.timedelta(days=7), "Declutter the closet", "Yes")]
        for k, (d, g, dn) in enumerate(goals):
            wg[f"A{G0 + k}"] = d; wg[f"B{G0 + k}"] = g
            if dn: wg[f"C{G0 + k}"] = dn
        logs = []
        done = {0: [0, 1, 2, 4], 1: [0, 2, 3, 4], 2: [0, 1, 3], 3: [0, 2], -1: [0, 1, 2, 3, 4], -2: [0, 2, 3], -3: [0, 1, 3, 4],
                -4: [0, 2, 4], -5: [0, 1, 2, 3], -6: [0, 3], -7: [0, 1, 2, 3, 4]}
        for off in sorted(done):
            for h in done[off]: logs.append((mon + dt.timedelta(days=off), habits[h]))
        for k, (d, h) in enumerate(logs):
            hl[f"A{H0 + k}"] = d; hl[f"B{H0 + k}"] = h

    # ---------- Dashboard (week view) ----------
    K.title(dash, "Weekly Planner", "Your week at a glance – updates by itself from Tasks, Appointments, Weekly Goals and the Habit Log.")
    dash["A3"] = "Week containing  ▸"; dash["A3"].font = font(bold=True); dash["A3"].alignment = Alignment(horizontal="right")
    input_cell(dash["B3"], D(2026, 11, 16) if sample else None); dash["B3"].number_format = DATE
    dash["C3"] = "(any date – leave empty for this week)"; dash["C3"].font = font(italic=True, size=9, color="5B6770")
    dash["A4"] = "Week starting"; dash["A4"].font = font(size=9, color="5B6770"); dash["A4"].alignment = Alignment(horizontal="right")
    dash["B4"] = '=IF(ISNUMBER(B3),B3,TODAY())'; dash["B4"].number_format = DATE; dash["B4"].font = font(color="FFFFFF", size=6)
    dash["C4"] = f'=B4-WEEKDAY(B4,{WT})+1'; dash["C4"].number_format = 'ddd, mmm d, yyyy'; dash["C4"].font = font(bold=True, color=MAIN)
    dash["E3"] = '=IF(Settings!$C$5="","",Settings!$C$5&"\'s week")'; dash["E3"].font = font(bold=True, size=12, color=MAIN)
    wk = f'{TK("I")},{WS},{TK("A")},"?*"'
    kp = [("Tasks this week", f'=COUNTIFS({wk})', "0"),
          ("Done", f'=COUNTIFS({wk},{TK("F")},"Yes")', "0"),
          ("Completed", f'=IF(B6=0,0,D6/B6)', PCT),
          ("Overdue (all weeks)", f'=COUNTIFS({TK("H")},"Overdue")', "0"),
          ("High priority still open", f'=COUNTIFS({wk},{TK("C")},"High",{TK("F")},"<>Yes")', "0"),
          ("Planned hours", f'=SUMIFS({TK("E")},{TK("I")},{WS})/60', "0.0"),
          ("Appointments", f'=COUNTIFS({AP("A")},">="&{WS},{AP("A")},"<"&({WS}+7),{AP("D")},"?*")', "0"),
          ("Habit check-offs", f'=COUNTIFS({HL("A")},">="&{WS},{HL("A")},"<"&({WS}+7),{HL("B")},"?*")', "0")]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 4) * 3, 2 + (i % 4) * 2, lab, f, fmt)
        dash.cell(row=6 + (i // 4) * 3, column=2 + (i % 4) * 2).alignment = Alignment(horizontal="left")
    dash.conditional_formatting.add("H6", FormulaRule(formula=["H6>0"], font=Font(name=F, size=18, bold=True, color="B91C1C")))
    dash.conditional_formatting.add("B9", FormulaRule(formula=["B9>0"], font=Font(name=F, size=18, bold=True, color="B45309")))
    # goals
    K.section(dash, "A11", "Goals for this week")
    GK = WG("E")
    for k in range(5):
        r = 12 + k
        m = f"MATCH(SMALL({GK},{k + 1}),{GK},0)"
        dash[f"B{r}"] = f'=IF(COUNT({GK})<{k + 1},"",IF(INDEX({WG("C")},{m})="Yes","✓  ","○  ")&INDEX({WG("B")},{m}))'
        dash[f"B{r}"].font = font(size=11)
    dash.conditional_formatting.add("B12:B16", FormulaRule(formula=['LEFT(B12,1)="✓"'], font=Font(name=F, color="9CA3AF", strike=True)))
    dash["A12"] = f'=IF(COUNT({GK})=0,"(add goals on the Weekly Goals tab)","")'; dash["A12"].font = font(italic=True, size=9, color="5B6770")
    # day grid
    R = 19
    dash[f"A{R}"] = ""; K.hdr(dash, R, [""] + [""] * 7)
    for d in range(7):
        c = dash.cell(row=R, column=2 + d, value=f'=TEXT({WS}+{d},"ddd, mmm d")')
        c.alignment = Alignment(horizontal="center", vertical="center")
    K.section(dash, f"A{R - 1}", "The week")
    AR = R + 1; NAP = 4; TR = AR + NAP; NTK = 10
    for d in range(7):
        col = chr(ord("B") + d); key = TK(DAYC[d]); akey = AP(DAYC[d])
        for k in range(NAP):
            r = AR + k; m = f"MATCH(SMALL({akey},{k + 1}),{akey},0)"
            dash[f"{col}{r}"] = (f'=IF(COUNT({akey})<{k + 1},"",IF(ISNUMBER(INDEX({AP("B")},{m})),TEXT(INDEX({AP("B")},{m}),"h:mm")&" ","")'
                                 f'&INDEX({AP("D")},{m}))')
        for k in range(NTK):
            r = TR + k; m = f"MATCH(SMALL({key},{k + 1}),{key},0)"
            dash[f"{col}{r}"] = (f'=IF(COUNT({key})<{k + 1},"",IF(INDEX({TK("F")},{m})="Yes","✓ ","○ ")&INDEX({TK("A")},{m})'
                                 f'&IF(INDEX({TK("C")},{m})="High"," !",""))')
        dash[f"{col}{TR + NTK}"] = f'=IF(COUNT({key})>{NTK},"+ "&(COUNT({key})-{NTK})&" more","")'
        dash[f"{col}{TR + NTK + 1}"] = f'=COUNTIFS({TK("B")},{WS}+{d},{TK("A")},"?*",{TK("F")},"Yes")'
        dash[f"{col}{TR + NTK + 2}"] = f'=COUNTIFS({TK("B")},{WS}+{d},{TK("A")},"?*")-{col}{TR + NTK + 1}'
    style_range(dash, AR, TR + NTK + 2, 2, 8)
    for r in range(AR, TR + NTK + 1):
        for d in range(7):
            c = dash.cell(row=r, column=2 + d); c.font = font(size=9); c.alignment = Alignment(wrap_text=False, shrink_to_fit=True)
            if r < TR: c.fill = fill("EEF6FF")
    for r, lab in ((AR, "Appointments"), (TR, "Tasks  (✓ done, ! high)"), (TR + NTK + 1, "Done"), (TR + NTK + 2, "Still open")):
        dash[f"A{r}"] = lab; dash[f"A{r}"].font = font(bold=True, size=9, color=MAIN); dash[f"A{r}"].alignment = Alignment(horizontal="right", vertical="top")
    for r in (TR + NTK + 1, TR + NTK + 2):
        for d in range(7): dash.cell(row=r, column=2 + d).alignment = Alignment(horizontal="center"); dash.cell(row=r, column=2 + d).number_format = "0"
    dash.conditional_formatting.add(f"B{TR}:H{TR + NTK - 1}", FormulaRule(formula=[f'LEFT(B{TR},1)="✓"'], font=Font(name=F, size=9, color="9CA3AF", strike=True)))
    dash.conditional_formatting.add(f"B{TR}:H{TR + NTK - 1}", FormulaRule(formula=[f'RIGHT(B{TR},1)="!"'], font=Font(name=F, size=9, bold=True, color="B91C1C")))
    for d in range(7):
        col = chr(ord("B") + d)
        dash.conditional_formatting.add(f"{col}{R}", FormulaRule(formula=[f"{WS}+{d}=TODAY()"], fill=fill("F59E0B")))
    # habits
    HR = TR + NTK + 5
    K.section(dash, f"A{HR - 1}", "Habits this week")
    K.hdr(dash, HR, ["Habit"] + [""] * 7 + ["Done", "Target"])
    for d in range(7):
        dash.cell(row=HR, column=2 + d, value=f'=TEXT({WS}+{d},"ddd")').alignment = Alignment(horizontal="center")
    for k in range(NS):
        r = HR + 1 + k
        dash[f"A{r}"] = f'=IF(Settings!$H${S0 + k}="","",Settings!$H${S0 + k})'
        for d in range(7):
            col = chr(ord("B") + d)
            dash[f"{col}{r}"] = f'=IF($A{r}="","",IF(COUNTIFS({HL("A")},{WS}+{d},{HL("B")},$A{r})>0,"✓",""))'
            dash[f"{col}{r}"].alignment = Alignment(horizontal="center")
        dash[f"I{r}"] = f'=IF($A{r}="","",COUNTIF(B{r}:H{r},"✓"))'
        dash[f"J{r}"] = f'=IF($A{r}="","",IF(ISNUMBER(Settings!$I${S0 + k}),Settings!$I${S0 + k},7))'
    style_range(dash, HR + 1, HR + NS, 1, 10)
    for r in range(HR + 1, HR + NS + 1):
        for c in range(2, 9): dash.cell(row=r, column=c).alignment = Alignment(horizontal="center"); dash.cell(row=r, column=c).font = font(bold=True, color=MAIN)
        dash.cell(row=r, column=9).alignment = Alignment(horizontal="center"); dash.cell(row=r, column=10).alignment = Alignment(horizontal="center")
    dash.conditional_formatting.add(f"B{HR + 1}:H{HR + NS}", FormulaRule(formula=[f'B{HR + 1}="✓"'], fill=fill(SOFT)))
    dash.conditional_formatting.add(f"I{HR + 1}:I{HR + NS}", FormulaRule(formula=[f'AND(ISNUMBER(I{HR + 1}),I{HR + 1}>=J{HR + 1})'], font=Font(name=F, bold=True, color="166534"), fill=fill("DCFCE7")))
    dash.column_dimensions["A"].width = 26
    for d in range(7): dash.column_dimensions[chr(ord("B") + d)].width = 24
    dash.column_dimensions["I"].width = 8; dash.column_dimensions["J"].width = 8
    # chart
    ch = BarChart(); ch.type = "col"; ch.grouping = "stacked"; ch.overlap = 100; ch.title = "Tasks per day"; ch.height = 7; ch.width = 15
    ch.add_data(Reference(dash, min_col=1, max_col=8, min_row=TR + NTK + 1), titles_from_data=True, from_rows=True)
    ch.add_data(Reference(dash, min_col=1, max_col=8, min_row=TR + NTK + 2), titles_from_data=True, from_rows=True)
    ch.set_categories(Reference(dash, min_col=2, max_col=8, min_row=R))
    ch.series[0].graphicalProperties.solidFill = MAIN; ch.series[1].graphicalProperties.solidFill = "99F6E4"
    ch.y_axis.majorGridlines = None; ch.legend.position = "b"; ch.legend.overlay = False; ch.title.overlay = False
    ch.roundedCorners = False; ch.x_axis.delete = False; ch.y_axis.delete = False; ch.gapWidth = 60
    dash.add_chart(ch, "J3")

    # ---------- Daily Page ----------
    K.title(dp, "Daily Page", "Pick a day (leave empty for today). Print it or keep it open – your tasks, appointments, overdue items and habits for that day.")
    dp["A3"] = "Day  ▸"; dp["A3"].font = font(bold=True); dp["A3"].alignment = Alignment(horizontal="right")
    input_cell(dp["B3"], D(2026, 11, 18) if sample else None); dp["B3"].number_format = DATE
    dp["C3"] = "(leave empty for today)"; dp["C3"].font = font(italic=True, size=9, color="5B6770")
    dp["B4"] = '=IF(ISNUMBER(B3),B3,TODAY())'; dp["B4"].number_format = DATE; dp["B4"].font = font(color="FFFFFF", size=6)
    dp["C4"] = "=B4"; dp["C4"].number_format = 'dddd, mmmm d, yyyy'; dp["C4"].font = font(bold=True, size=14, color=MAIN)
    for c, w in zip("ABCDEFG", [16, 14, 44, 11, 13, 10, 3]): dp.column_dimensions[c].width = w
    # appointments
    K.section(dp, "A6", "Schedule")
    K.hdr(dp, 7, ["Start", "End", "What", "Where / who"], col0=1)
    QA = AP("Q")
    for k in range(8):
        r = 8 + k; m = f"MATCH(SMALL({QA},{k + 1}),{QA},0)"
        for col, src in zip("ABCD", "BCDE"):
            dp[f"{col}{r}"] = f'=IF(COUNT({QA})<{k + 1},"",IF(INDEX({AP(src)},{m})="","",INDEX({AP(src)},{m})))'
    style_range(dp, 8, 15, 1, 4, fmt={1: TIME, 2: TIME})
    # tasks
    K.section(dp, "A17", "To-do (high priority first)")
    K.hdr(dp, 18, ["Done", "Priority", "Task", "Category", "Est. minutes", "Status"], col0=1)
    QT = TK("Q")
    for k in range(15):
        r = 19 + k; m = f"MATCH(SMALL({QT},{k + 1}),{QT},0)"
        dp[f"A{r}"] = f'=IF(COUNT({QT})<{k + 1},"",IF(INDEX({TK("F")},{m})="Yes","✓","○"))'
        for col, src in zip("BCDEF", "CADEH"):
            dp[f"{col}{r}"] = f'=IF(COUNT({QT})<{k + 1},"",IF(INDEX({TK(src)},{m})="","",INDEX({TK(src)},{m})))'
    style_range(dp, 19, 33, 1, 6, fmt={5: "0"})
    for r in range(19, 34): dp[f"A{r}"].alignment = Alignment(horizontal="center")
    dp.conditional_formatting.add("A19:F33", FormulaRule(formula=['$A19="✓"'], font=Font(name=F, color="9CA3AF", strike=True)))
    dp.conditional_formatting.add("B19:B33", FormulaRule(formula=['B19="High"'], font=Font(name=F, bold=True, color="B91C1C")))
    dp["A34"] = f'="Planned: "&TEXT(SUMIFS({TK("E")},{TK("B")},$B$4)/60,"0.0")&" h  ·  done "&COUNTIFS({TK("B")},$B$4,{TK("A")},"?*",{TK("F")},"Yes")&" of "&COUNTIFS({TK("B")},$B$4,{TK("A")},"?*")'
    dp["A34"].font = font(bold=True, color=MAIN)
    # overdue
    K.section(dp, "A36", "Overdue – reschedule or do today")
    K.hdr(dp, 37, ["Was due", "Priority", "Task", "Category"], col0=1)
    QO = TK("R")
    for k in range(8):
        r = 38 + k; m = f"MATCH(SMALL({QO},{k + 1}),{QO},0)"
        for col, src in zip("ABCD", "BCAD"):
            dp[f"{col}{r}"] = f'=IF(COUNT({QO})<{k + 1},"",IF(INDEX({TK(src)},{m})="","",INDEX({TK(src)},{m})))'
    style_range(dp, 38, 45, 1, 4, fmt={1: DATE})
    dp.conditional_formatting.add("A38:D45", FormulaRule(formula=['$A38<>""'], font=Font(name=F, color="B91C1C")))
    # habits today
    K.section(dp, "A47", "Habits today")
    for k in range(NS):
        r = 48 + k
        dp[f"A{r}"] = f'=IF(Settings!$H${S0 + k}="","",IF(COUNTIFS({HL("A")},$B$4,{HL("B")},Settings!$H${S0 + k})>0,"✓","○"))'
        dp[f"B{r}"] = f'=IF(Settings!$H${S0 + k}="","",Settings!$H${S0 + k})'
        dp[f"A{r}"].alignment = Alignment(horizontal="center"); dp[f"A{r}"].font = font(bold=True, color=MAIN); dp[f"B{r}"].font = font()
    # notes area
    K.section(dp, "E47", "Notes")
    for r in range(48, 56):
        for c in "EF": dp[f"{c}{r}"].border = B

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Settings: your name, whether your week starts on Monday or Sunday, your task categories and the habits you want to build.", False),
             ("2.  Tasks: type each to-do with a date and priority (High / Medium / Low). No date = a 'someday' task. Choose Yes under Done when it is finished.", False),
             ("3.  Appointments: anything at a set time (meetings, classes, doctor, kids' games). They show on the week view in time order.", False),
             ("4.  Weekly Goals: your 1–5 big things for the week. Type any date in that week.", False),
             ("5.  Habit Log: add a row (date + habit) each time you do a habit – the habit grid ticks itself.", False),
             ("6.  Dashboard: the whole week – goals, appointments and tasks for each day, habits and progress. Type a date at the top to look at another week.", False),
             ("7.  Daily Page: one day in detail – schedule, to-do list (high priority first), overdue tasks and habits. Great for printing.", False),
             ("", False), ("Tips", True),
             ("•  Plan on Sunday: add next week's tasks and goals, then check the Dashboard.", False),
             ("•  Overdue tasks turn red everywhere. Change the date to move a task to another day.", False),
             ("•  Up to 10 tasks and 4 appointments show per day on the week view ('+ more' if there are more); the Daily Page shows 15.", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  Room for 1,000 tasks, 500 appointments, 300 weekly goals and 3,000 habit check-offs. No macros. Works in Excel 2010+ and Google Sheets.", False),
             ("•  To use Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("A made-up busy week (Nov 16–22, 2026) with tasks, appointments, goals and habits. Use the blank copy in your download for your own planning.", False)]
    K.start(st, "Weekly & Daily Planner", "Tasks, priorities, appointments, weekly goals and habits – a week view and a daily page, Excel and Google Sheets.", lines)
    finish(wb, sample, path)


META = {
    "name": "Weekly & Daily Planner (To-Do List, Priorities, Appointments, Weekly Goals, Habit Tracker) – Excel & Google Sheets",
    "file_base": "Weekly-Daily-Planner",
    "price": 12,
    "color": "#0F766E",
    "title": "Weekly & Daily Planner",
    "tagline": "to-dos · priorities · appointments · weekly goals · habits",
    "intro": "Plan your whole week in one spreadsheet. Type your to-dos with a date and priority, add appointments and your goals for the week, and the planner lays everything out in a clean 7-day view – high-priority tasks first, overdue tasks in red, done tasks crossed out. A Daily Page shows one day in detail for printing, and a small habit tracker keeps your routines on track.",
    "features": [
        "7-day week view: appointments in time order and up to 10 tasks per day, today highlighted",
        "To-do list with date, priority (High / Medium / Low), category, time estimate and done check",
        "Automatic status: Today, Planned, Overdue, Done, Someday",
        "Weekly goals (1–5 per week) shown at the top of the week view",
        "Daily Page: schedule, to-do list sorted by priority, overdue tasks to reschedule, habits and notes – print-ready",
        "Habit tracker: tick habits per day and compare with your weekly target",
        "Weekly stats: tasks done, % completed, planned hours, high-priority items still open, chart of tasks per day",
        "Week starts on Monday or Sunday; look at any week by typing a date",
        "Blank copy plus a sample copy with a full example week",
    ],
    "tags": ["weekly planner", "daily planner", "to do list", "digital planner", "google sheets planner", "excel planner",
             "productivity", "habit tracker", "weekly goals", "time management", "2027 planner", "task tracker"],
    "summary": "Instant download: Excel + Google Sheets template, blank copy + sample copy.",
}
