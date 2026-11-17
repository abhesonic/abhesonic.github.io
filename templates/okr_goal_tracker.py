"""Goal Planner & OKR Tracker – objectives, measurable key results, weekly check-ins with confidence, actions,
progress vs expected pace, dashboard. Works for personal goals and small teams.

Settings -> Objectives -> Key Results -> Check-ins -> Actions -> Dashboard.
"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, DataBarRule
from openpyxl.chart import BarChart, Reference
from _common import Kit, F, B, DATE, font, fill, style_range, input_cell, finish

MAIN = "4338CA"; SOFT = "E0E7FF"
K = Kit(MAIN, SOFT)
NO = 40; O0 = 6; O1 = O0 + NO - 1
NK = 160; K0 = 6; K1 = K0 + NK - 1
NC = 1500; C0 = 6; C1 = C0 + NC - 1
NA = 300; A0 = 6; A1 = A0 + NA - 1
NS = 12; S0 = 6; S1 = S0 + NS - 1     # areas / owners lists
PCT = "0%"; NUM = '#,##0.##'
PERIODS = "Q1,Q2,Q3,Q4,Year"
AREAS = ["Career", "Business", "Health & fitness", "Money", "Learning", "Relationships", "Home", "Personal growth", "Team"]


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); ob = wb.create_sheet("Objectives"); kr = wb.create_sheet("Key Results")
    ci = wb.create_sheet("Check-ins"); ac = wb.create_sheet("Actions"); s = wb.create_sheet("Settings")

    # ---------- Settings ----------
    K.title(s, "Settings", "Yellow cells are yours to edit.")
    SET = [("Year", None), ("Show on dashboard (All, Q1–Q4 or Year)", "All"), ("'On track' if progress is at least this close to expected", 0.10),
           ("'At risk' if within (otherwise 'Behind')", 0.25), ("Check-in reminder: flag key results not updated for (days)", 14)]
    for i, (lab, v) in enumerate(SET, 5):
        s[f"A{i}"] = lab; s[f"A{i}"].font = font(bold=True); s[f"A{i}"].border = B; input_cell(s[f"C{i}"], v)
    s["C7"].number_format = PCT; s["C8"].number_format = PCT; s["C5"].number_format = "0"
    dv = DataValidation(type="list", formula1=f'"All,{PERIODS}"', allow_blank=False); s.add_data_validation(dv); dv.add("C6")
    s["A11"] = "Year used (automatic)"; s["A11"].font = font(italic=True)
    s["C11"] = "=IF(ISNUMBER(C5),C5,YEAR(TODAY()))"; s["C11"].number_format = "0"
    s["C11"].fill = fill("F3F4F6"); s["C11"].border = B; s["C11"].font = font()
    K.hdr(s, 5, ["Life / work areas"], col0=6); K.hdr(s, 5, ["People / owners"], col0=8)
    for i in range(NS):
        input_cell(s[f"F{S0 + i}"], AREAS[i] if i < len(AREAS) else None)
        input_cell(s[f"H{S0 + i}"], ("Me" if i == 0 else None) if not sample else None)
    for c, w in zip("ABCDEFGH", [52, 3, 10, 3, 3, 22, 3, 22]): s.column_dimensions[c].width = w
    YR = "Settings!$C$11"; SHOW = "Settings!$C$6"; TOL = "Settings!$C$7"; RISK = "Settings!$C$8"; STALE = "Settings!$C$9"

    OB = lambda c: f"Objectives!${c}${O0}:${c}${O1}"
    KR = lambda c: f"'Key Results'!${c}${K0}:${c}${K1}"
    CI = lambda c: f"'Check-ins'!${c}${C0}:${c}${C1}"
    AC = lambda c: f"Actions!${c}${A0}:${c}${A1}"

    # ---------- Objectives ----------
    K.title(ob, "Objectives", "What you want to achieve – short and inspiring. 3–5 per quarter is plenty. Progress comes from the key results.")
    cols = ["ID", "Objective", "Area", "Owner", "Period", "Year", "Why it matters", "Start date", "End date", "Key results",
            "Progress", "Expected by today", "Status", "Shown on dashboard?", "(sort key)"]
    K.hdr(ob, 5, cols, [6, 40, 16, 13, 8, 7, 34, 11, 11, 9, 10, 10, 13, 10, 6])
    style_range(ob, O0, O1, 1, 15, inputs=range(2, 8), fmt={6: "0", 8: DATE, 9: DATE, 10: "0", 11: PCT, 12: PCT})
    for r in range(O0, O1 + 1):
        ob[f"A{r}"] = f'=IF(B{r}="","","O"&COUNTIF($B${O0}:B{r},"?*"))'
        yr = f"IF(ISNUMBER(F{r}),F{r},{YR})"
        ob[f"H{r}"] = f'=IF(B{r}="","",IF(E{r}="Q2",DATE({yr},4,1),IF(E{r}="Q3",DATE({yr},7,1),IF(E{r}="Q4",DATE({yr},10,1),DATE({yr},1,1)))))'
        ob[f"I{r}"] = f'=IF(B{r}="","",IF(E{r}="Q1",DATE({yr},3,31),IF(E{r}="Q2",DATE({yr},6,30),IF(E{r}="Q3",DATE({yr},9,30),DATE({yr},12,31)))))'
        ob[f"J{r}"] = f'=IF(B{r}="","",COUNTIFS({KR("A")},A{r},{KR("C")},"?*"))'
        ob[f"K{r}"] = f'=IF(OR(B{r}="",N(J{r})=0),"",SUMIFS({KR("W")},{KR("A")},A{r})/SUMIFS({KR("V")},{KR("A")},A{r}))'
        ob[f"L{r}"] = f'=IF(OR(B{r}="",N(J{r})=0),"",SUMIFS({KR("X")},{KR("A")},A{r})/SUMIFS({KR("V")},{KR("A")},A{r}))'
        ob[f"M{r}"] = (f'=IF(B{r}="","",IF(N(J{r})=0,"Add key results",IF(K{r}>=0.9999,"Done",IF(TODAY()>I{r},"Missed",'
                       f'IF(K{r}>=L{r}-{TOL},"On track",IF(K{r}>=L{r}-{RISK},"At risk","Behind"))))))')
        ob[f"N{r}"] = f'=IF(B{r}="","",IF(AND({yr}={YR},OR({SHOW}="All",{SHOW}=E{r})),"Yes","No"))'
        ob[f"O{r}"] = f'=IF(N{r}="Yes",ROW(),"")'; ob[f"O{r}"].font = font(color="9CA3AF", size=7)
        ob[f"M{r}"].font = font(bold=True)
    for col, f1 in (("C", f"=Settings!$F${S0}:$F${S1}"), ("D", f"=Settings!$H${S0}:$H${S1}"), ("E", f'"{PERIODS}"')):
        dv = DataValidation(type="list", formula1=f1, allow_blank=True); ob.add_data_validation(dv); dv.add(f"{col}{O0}:{col}{O1}")
    ob.conditional_formatting.add(f"K{O0}:K{O1}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="818CF8"))
    ST = {"Done": ("166534", "DCFCE7"), "On track": ("166534", None), "At risk": ("92400E", "FEF3C7"), "Behind": ("B91C1C", "FEE2E2"), "Missed": ("6B7280", "F3F4F6")}

    def status_cf(ws, rng, first):
        for k_, (fg, bg) in ST.items():
            kw = dict(font=Font(name=F, bold=True, color=fg))
            if bg: kw["fill"] = fill(bg)
            ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{first}="{k_}"'], **kw))
    status_cf(ob, f"M{O0}:M{O1}", f"M{O0}")
    ob.freeze_panes = f"C{O0}"

    # ---------- Key Results ----------
    K.title(kr, "Key Results", "2–4 measurable results per objective: a number you can check. For yes/no results use start 0, target 1.")
    cols = ["Objective ID", "KR ID", "Key result", "Owner", "Unit", "Start value", "Target value", "Weight (1 = normal)", "Own due date (optional)",
            "Objective", "Current value", "Progress", "Due", "Expected by today", "Status", "Confidence (latest, 1–10)",
            "Last check-in", "Days since check-in", "Open actions", "Needs update?", "(start)", "(weight)", "(w×progress)", "(w×expected)", "(rank key)", "(shown)"]
    K.hdr(kr, 5, cols, [10, 7, 40, 12, 9, 10, 10, 9, 11, 30, 10, 10, 11, 10, 11, 11, 11, 9, 8, 10, 6, 6, 6, 6, 6, 6])
    style_range(kr, K0, K1, 1, 26, inputs=(1, 3, 4, 5, 6, 7, 8, 9),
                fmt={6: NUM, 7: NUM, 8: "0.0", 9: DATE, 11: NUM, 12: PCT, 13: DATE, 14: PCT, 16: "0", 17: DATE, 18: "0", 19: "0"})
    for r in range(K0, K1 + 1):
        m = f"MATCH(A{r},{OB('A')},0)"
        kr[f"B{r}"] = f'=IF(OR(A{r}="",C{r}=""),"",A{r}&"."&COUNTIFS($A${K0}:A{r},A{r},$C${K0}:C{r},"?*"))'
        kr[f"J{r}"] = f'=IF(A{r}="","",IFERROR(INDEX({OB("B")},{m}),"(unknown objective ID)"))'
        kr[f"K{r}"] = (f'=IF(B{r}="","",IF(COUNTIFS({CI("B")},B{r},{CI("C")},"<>")=0,F{r},'
                       f'LOOKUP(2,1/(({CI("B")}=B{r})*ISNUMBER({CI("C")})),{CI("C")})))')
        kr[f"L{r}"] = (f'=IF(OR(B{r}="",NOT(ISNUMBER(K{r}))),"",IF(N(G{r})=N(F{r}),IF(N(K{r})=N(G{r}),1,0),'
                       f'MAX(0,MIN(1,(N(K{r})-N(F{r}))/(N(G{r})-N(F{r}))))))')
        kr[f"U{r}"] = f'=IF(B{r}="","",IFERROR(INDEX({OB("H")},{m}),""))'
        kr[f"M{r}"] = f'=IF(B{r}="","",IF(ISNUMBER(I{r}),I{r},IFERROR(INDEX({OB("I")},{m}),"")))'
        kr[f"N{r}"] = f'=IF(OR(B{r}="",NOT(ISNUMBER(U{r})),NOT(ISNUMBER(M{r}))),"",MAX(0,MIN(1,(TODAY()-U{r})/MAX(1,M{r}-U{r}))))'
        kr[f"O{r}"] = (f'=IF(B{r}="","",IF(L{r}>=0.9999,"Done",IF(AND(ISNUMBER(M{r}),TODAY()>M{r}),"Missed",IF(NOT(ISNUMBER(N{r})),"",'
                       f'IF(L{r}>=N{r}-{TOL},"On track",IF(L{r}>=N{r}-{RISK},"At risk","Behind"))))))')
        kr[f"P{r}"] = (f'=IF(B{r}="","",IF(COUNTIFS({CI("B")},B{r},{CI("D")},">0")=0,"",'
                       f'LOOKUP(2,1/(({CI("B")}=B{r})*ISNUMBER({CI("D")})),{CI("D")})))')
        kr[f"Q{r}"] = f'=IF(B{r}="","",IF(COUNTIFS({CI("B")},B{r})=0,"",SUMPRODUCT(MAX(({CI("B")}=B{r})*{CI("H")}))))'
        kr[f"R{r}"] = f'=IF(OR(B{r}="",NOT(ISNUMBER(Q{r}))),"",TODAY()-Q{r})'
        kr[f"S{r}"] = f'=IF(B{r}="","",COUNTIFS({AC("B")},B{r},{AC("E")},"<>Yes",{AC("A")},"?*"))'
        kr[f"T{r}"] = (f'=IF(OR(B{r}="",O{r}="Done",O{r}="Missed"),"",IF(NOT(ISNUMBER(Q{r})),"No check-in yet",'
                       f'IF(R{r}>{STALE},"Update ("&R{r}&" days)","")))')
        kr[f"V{r}"] = f'=IF(B{r}="",0,IF(ISNUMBER(H{r}),H{r},1))'
        kr[f"W{r}"] = f'=IF(B{r}="",0,V{r}*N(L{r}))'
        kr[f"X{r}"] = f'=IF(B{r}="",0,V{r}*N(N{r}))'
        kr[f"Z{r}"] = f'=IF(B{r}="","",IFERROR(INDEX({OB("N")},{m}),"No"))'
        kr[f"Y{r}"] = f'=IF(AND(Z{r}="Yes",OR(O{r}="Behind",O{r}="At risk",O{r}="Missed")),N(L{r})-N(N{r})+ROW()/1000000,"")'
        for c in "UVWXYZ": kr[f"{c}{r}"].font = font(color="9CA3AF", size=7)
        kr[f"O{r}"].font = font(bold=True)
    dv = DataValidation(type="list", formula1=f"=Objectives!$A${O0}:$A${O1}", allow_blank=True); kr.add_data_validation(dv); dv.add(f"A{K0}:A{K1}")
    dv = DataValidation(type="list", formula1=f"=Settings!$H${S0}:$H${S1}", allow_blank=True); kr.add_data_validation(dv); dv.add(f"D{K0}:D{K1}")
    kr.conditional_formatting.add(f"L{K0}:L{K1}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="818CF8"))
    status_cf(kr, f"O{K0}:O{K1}", f"O{K0}")
    kr.conditional_formatting.add(f"T{K0}:T{K1}", FormulaRule(formula=[f'T{K0}<>""'], font=Font(name=F, bold=True, color="92400E"), fill=fill("FEF3C7")))
    kr.conditional_formatting.add(f"P{K0}:P{K1}", FormulaRule(formula=[f'AND(ISNUMBER(P{K0}),P{K0}<=4)'], font=Font(name=F, bold=True, color="B91C1C")))
    kr.freeze_panes = f"D{K0}"; kr.auto_filter.ref = f"A5:T{K1}"

    # ---------- Check-ins ----------
    K.title(ci, "Check-ins", "Once a week: the new value of each key result and how confident you are (1–10) to hit the target. Add new rows at the bottom.")
    K.hdr(ci, 5, ["Date", "KR ID", "New value", "Confidence (1–10)", "What happened / what's next", "Key result", "Progress at this check-in", "(date)"],
          [12, 8, 11, 11, 46, 40, 12, 6])
    style_range(ci, C0, C1, 1, 8, inputs=(1, 2, 3, 4, 5), fmt={1: DATE, 3: NUM, 4: "0", 7: PCT, 8: "0"})
    for r in range(C0, C1 + 1):
        m = f"MATCH(B{r},{KR('B')},0)"
        ci[f"F{r}"] = f'=IF(B{r}="","",IFERROR(INDEX({KR("C")},{m}),"(unknown KR ID)"))'
        ci[f"G{r}"] = (f'=IF(OR(B{r}="",NOT(ISNUMBER(C{r}))),"",IFERROR(IF(INDEX({KR("G")},{m})=INDEX({KR("F")},{m}),IF(C{r}=INDEX({KR("G")},{m}),1,0),'
                       f'MAX(0,MIN(1,(C{r}-INDEX({KR("F")},{m}))/(INDEX({KR("G")},{m})-INDEX({KR("F")},{m}))))),""))')
        ci[f"H{r}"] = f'=IF(AND(B{r}<>"",ISNUMBER(A{r})),A{r},0)'; ci[f"H{r}"].font = font(color="9CA3AF", size=7)
    dv = DataValidation(type="list", formula1=f"='Key Results'!$B${K0}:$B${K1}", allow_blank=True); ci.add_data_validation(dv); dv.add(f"B{C0}:B{C1}")
    dv = DataValidation(type="whole", operator="between", formula1="1", formula2="10", allow_blank=True); ci.add_data_validation(dv); dv.add(f"D{C0}:D{C1}")
    ci.freeze_panes = f"A{C0}"; ci.auto_filter.ref = f"A5:G{C1}"

    # ---------- Actions ----------
    K.title(ac, "Actions", "The to-dos and projects that move a key result. Choose 'Yes' under Done when finished.")
    K.hdr(ac, 5, ["Action", "KR ID", "Owner", "Due", "Done?", "Notes", "Key result", "Status"], [40, 8, 12, 11, 7, 30, 38, 16])
    style_range(ac, A0, A1, 1, 8, inputs=range(1, 7), fmt={4: DATE})
    for r in range(A0, A1 + 1):
        ac[f"G{r}"] = f'=IF(B{r}="","",IFERROR(INDEX({KR("C")},MATCH(B{r},{KR("B")},0)),"(unknown KR ID)"))'
        ac[f"H{r}"] = (f'=IF(A{r}="","",IF(E{r}="Yes","Done",IF(NOT(ISNUMBER(D{r})),"Open",IF(D{r}<TODAY(),"Overdue",'
                       f'IF(D{r}-TODAY()<=7,"Due this week","Open")))))')
        ac[f"H{r}"].font = font(bold=True)
    dv = DataValidation(type="list", formula1=f"='Key Results'!$B${K0}:$B${K1}", allow_blank=True); ac.add_data_validation(dv); dv.add(f"B{A0}:B{A1}")
    dv = DataValidation(type="list", formula1=f"=Settings!$H${S0}:$H${S1}", allow_blank=True); ac.add_data_validation(dv); dv.add(f"C{A0}:C{A1}")
    dv = DataValidation(type="list", formula1='"Yes"', allow_blank=True); ac.add_data_validation(dv); dv.add(f"E{A0}:E{A1}")
    ac.conditional_formatting.add(f"H{A0}:H{A1}", FormulaRule(formula=[f'H{A0}="Overdue"'], font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    ac.conditional_formatting.add(f"H{A0}:H{A1}", FormulaRule(formula=[f'H{A0}="Due this week"'], font=Font(name=F, bold=True, color="92400E"), fill=fill("FEF3C7")))
    ac.conditional_formatting.add(f"A{A0}:H{A1}", FormulaRule(formula=[f'$E{A0}="Yes"'], font=Font(name=F, color="9CA3AF", strike=True)))
    ac.freeze_panes = f"A{A0}"; ac.auto_filter.ref = f"A5:H{A1}"

    # ---------- sample ----------
    if sample:
        s["C5"] = 2026
        for i, n in enumerate(["Me", "Sam (partner)", "Coach Rita"]): s[f"H{S0 + i}"] = n
        objs = [("Launch my online store and get the first customers", "Business", "Me", "Q4", None, "Side income so I can work 4 days a week by 2028"),
                ("Get fit and feel strong for the new year", "Health & fitness", "Me", "Q4", None, "More energy for the kids"),
                ("Become debt-free (credit card)", "Money", "Me", "Year", None, "Stop paying 24% interest"),
                ("Hold a basic conversation in Spanish (A2)", "Learning", "Me", "Year", None, "Family trip to Mexico next summer"),
                ("Read more and switch off screens at night", "Personal growth", "Me", "Q3", None, "Better sleep"),
                ("Make our home calm and organised", "Home", "Sam (partner)", "Q4", None, "Less stress in the mornings")]
        for k, row in enumerate(objs):
            for c, v in zip("BCDEFG", row):
                if v: ob[f"{c}{O0 + k}"] = v
        krs = [("O1", "Products listed in the store", "Me", "products", 0, 20, 1, None),
               ("O1", "Sales", "Me", "orders", 0, 50, 2, None),
               ("O1", "Email subscribers", "Me", "people", 300, 1000, 1, None),
               ("O2", "Body weight", "Me", "lb", 185, 175, 1, None),
               ("O2", "5k run time", "Coach Rita", "minutes", 31, 28, 1, None),
               ("O2", "Workouts completed", "Me", "workouts", 0, 36, 1, None),
               ("O3", "Credit card balance", "Me", "USD", 6000, 0, 1, None),
               ("O4", "Duolingo lessons", "Me", "lessons", 0, 300, 1, None),
               ("O4", "Conversation sessions with a tutor", "Me", "sessions", 0, 24, 2, None),
               ("O5", "Books finished", "Me", "books", 0, 6, 1, None),
               ("O5", "Nights with no phone after 10 pm", "Me", "nights", 0, 60, 1, None),
               ("O6", "Rooms decluttered", "Sam (partner)", "rooms", 0, 6, 1, None),
               ("O6", "Weekly meal plan made", "Sam (partner)", "weeks", 0, 12, 1, None),
               ("O6", "Garage sale done", "Me", "yes = 1", 0, 1, 1, dt.date(2026, 11, 30))]
        for k, row in enumerate(krs):
            for c, v in zip("ACDEFGHI", row):
                if v is not None: kr[f"{c}{K0 + k}"] = v
        D = dt.date
        cks = []
        weekly = {"O1.1": [3, 6, 8, 11, 12, 14], "O1.2": [0, 2, 5, 9, 13, 17], "O1.3": [340, 410, 480, 540, 600, 655],
                  "O2.1": [184, 183, 182.4, 181, 180.6, 180.2], "O2.2": [30.6, 30.4, 30.1, 30.0, 29.8, 29.6], "O2.3": [3, 6, 9, 11, 14, 16],
                  "O3.1": [3400, 3100, 2850, 2600, 2400, 2150], "O4.1": [200, 212, 224, 236, 247, 258], "O4.2": [8, 9, 10, 11, 12, 12],
                  "O6.1": [0, 1, 1, 2, 2, 2], "O6.2": [1, 2, 3, 4, 4, 5]}
        conf = {"O1.1": [7, 8, 8, 8, 8, 8], "O1.2": [5, 5, 6, 6, 6, 6], "O1.3": [6, 6, 7, 7, 7, 7], "O2.1": [6, 6, 6, 6, 5, 5],
                "O2.2": [5, 5, 5, 6, 6, 6], "O2.3": [7, 7, 7, 6, 6, 6], "O3.1": [8, 8, 8, 8, 9, 9], "O4.1": [8, 8, 8, 8, 8, 8],
                "O4.2": [6, 6, 6, 6, 5, 5], "O6.1": [6, 5, 5, 4, 4, 3], "O6.2": [8, 8, 7, 7, 6, 6]}
        notes = {"O1.2": "First sale from Pinterest!", "O2.1": "Holiday parties coming – stick to the plan", "O6.1": "Busy weeks, garage still to do",
                 "O1.3": "Freebie checklist is working", "O4.2": "Missed one session (sick)"}
        dates = [D(2026, 10, 12) + dt.timedelta(days=7 * i) for i in range(6)]
        for i, d in enumerate(dates):
            for kid, vals in weekly.items():
                if kid == "O6.1" and i == 5: continue
                cks.append((d, kid, vals[i], conf[kid][i], notes.get(kid, "") if i == 5 else ""))
        cks += [(D(2026, 7, 31), "O5.1", 2, 7, "Summer reading"), (D(2026, 8, 31), "O5.1", 4, 8, ""), (D(2026, 9, 30), "O5.1", 6, 10, "Done!"),
                (D(2026, 7, 31), "O5.2", 15, 6, ""), (D(2026, 8, 31), "O5.2", 33, 6, ""), (D(2026, 9, 30), "O5.2", 48, 5, "Missed it – try again in Q1"),
                (D(2026, 9, 30), "O3.1", 3700, 8, "Paid bonus towards the card")]
        cks.sort(key=lambda c: c[0])
        for k, row in enumerate(cks):
            for c, v in zip("ABCDE", row):
                if v not in (None, ""): ci[f"{c}{C0 + k}"] = v
        acts = [("Photograph the last 6 products", "O1.1", "Me", D(2026, 11, 20), None, ""),
                ("Write descriptions for new listings", "O1.1", "Me", D(2026, 11, 22), None, ""),
                ("Black Friday discount code + email", "O1.2", "Me", D(2026, 11, 26), None, "20% off, 3 days"),
                ("Post 3 Pinterest pins a day", "O1.2", "Me", D(2026, 11, 30), None, ""),
                ("Set up welcome email sequence", "O1.3", "Me", D(2026, 10, 20), "Yes", ""),
                ("Book 6 sessions with Coach Rita", "O2.2", "Me", D(2026, 10, 10), "Yes", ""),
                ("Meal-prep Sundays", "O2.1", "Sam (partner)", D(2026, 11, 16), None, ""),
                ("Set up auto-payment of $250/week to the card", "O3.1", "Me", D(2026, 10, 3), "Yes", ""),
                ("Book 4 more tutor sessions", "O4.2", "Me", D(2026, 11, 14), None, ""),
                ("Declutter the garage", "O6.1", "Sam (partner)", D(2026, 11, 15), None, ""),
                ("Price items for garage sale", "O6.3", "Me", D(2026, 11, 25), None, "")]
        for k, row in enumerate(acts):
            for c, v in zip("ABCDEF", row):
                if v: ac[f"{c}{A0 + k}"] = v

    # ---------- Dashboard ----------
    K.title(dash, "Goals & OKR Dashboard", "Updates by itself from Objectives, Key Results, Check-ins and Actions.")
    dash["A3"] = '="Showing: "&IF(Settings!$C$6="All","all periods",Settings!$C$6)&" "&Settings!$C$11&"  ·  today: "&TEXT(TODAY(),"mmm d, yyyy")&"  (change in Settings)"'
    dash["A3"].font = font(bold=True)
    sh = f'{KR("Z")},"Yes"'
    kp = [("Objectives", f'=COUNTIFS({OB("N")},"Yes")', "0"),
          ("Average progress", f'=IF(SUMIFS({KR("V")},{KR("Z")},"Yes")=0,0,SUMIFS({KR("W")},{KR("Z")},"Yes")/SUMIFS({KR("V")},{KR("Z")},"Yes"))', PCT),
          ("Expected by today", f'=IF(SUMIFS({KR("V")},{KR("Z")},"Yes")=0,0,SUMIFS({KR("X")},{KR("Z")},"Yes")/SUMIFS({KR("V")},{KR("Z")},"Yes"))', PCT),
          ("Key results on track or done", f'=COUNTIFS({sh},{KR("O")},"On track")+COUNTIFS({sh},{KR("O")},"Done")', "0"),
          ("Key results at risk", f'=COUNTIFS({sh},{KR("O")},"At risk")', "0"),
          ("Key results behind or missed", f'=COUNTIFS({sh},{KR("O")},"Behind")+COUNTIFS({sh},{KR("O")},"Missed")', "0"),
          ("Check-ins in the last 7 days", f'=COUNTIFS({CI("A")},">="&(TODAY()-7),{CI("A")},"<="&TODAY())', "0"),
          ("Actions overdue", f'=COUNTIFS({AC("H")},"Overdue")', "0"),
          ("Key results needing an update", f'=COUNTIFS({sh},{KR("T")},"?*")', "0")]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 3) * 3, 1 + (i % 3) * 2, lab, f, fmt)
        dash.cell(row=6 + (i // 3) * 3, column=1 + (i % 3) * 2).alignment = Alignment(horizontal="left")
    for c, w in zip("ABCDEFGH", [50, 12, 26, 14, 26, 14, 3, 3]): dash.column_dimensions[c].width = w
    for ref in ("C12", "E12"):
        dash.conditional_formatting.add(ref, FormulaRule(formula=[f"{ref}>0"], font=Font(name=F, size=18, bold=True, color="B91C1C")))
    dash.conditional_formatting.add("E9", FormulaRule(formula=["E9>0"], font=Font(name=F, size=18, bold=True, color="B91C1C")))
    dash.conditional_formatting.add("C9", FormulaRule(formula=["C9>0"], font=Font(name=F, size=18, bold=True, color="B45309")))
    # objectives table
    H = 15
    K.section(dash, f"A{H - 1}", "Objectives")
    K.hdr(dash, H, ["Objective", "Period", "Owner", "Progress", "Status", "Expected"])
    OKEY = OB("O")
    for k in range(10):
        r = H + 1 + k
        dash[f"H{r}"] = f'=IF(COUNT({OKEY})<{k + 1},"",SMALL({OKEY},{k + 1}))'; dash[f"H{r}"].font = font(color="FFFFFF", size=6)
        m = f"MATCH($H{r},{OKEY},0)"
        dash[f"A{r}"] = f'=IF($H{r}="","",INDEX({OB("A")},{m})&"  "&INDEX({OB("B")},{m}))'
        dash[f"B{r}"] = f'=IF($H{r}="","",INDEX({OB("E")},{m}))'
        dash[f"C{r}"] = f'=IF($H{r}="","",INDEX({OB("D")},{m})&"")'
        dash[f"D{r}"] = f'=IF($H{r}="","",N(INDEX({OB("K")},{m})))'
        dash[f"E{r}"] = f'=IF($H{r}="","",INDEX({OB("M")},{m}))'
        dash[f"F{r}"] = f'=IF($H{r}="","",N(INDEX({OB("L")},{m})))'
    style_range(dash, H + 1, H + 10, 1, 6, fmt={4: PCT, 6: PCT})
    for r in range(H + 1, H + 11): dash[f"E{r}"].font = font(bold=True)
    dash.conditional_formatting.add(f"D{H + 1}:D{H + 10}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="818CF8"))
    status_cf(dash, f"E{H + 1}:E{H + 10}", f"E{H + 1}")
    # attention list
    T = H + 14
    K.section(dash, f"A{T - 1}", "Key results that need attention (furthest behind first)")
    K.hdr(dash, T, ["Key result", "KR ID", "Owner", "Progress", "Status", "Expected", "Confidence"])
    RK = KR("Y")
    for k in range(8):
        r = T + 1 + k
        dash[f"H{r}"] = f'=IF(COUNT({RK})<{k + 1},"",SMALL({RK},{k + 1}))'; dash[f"H{r}"].font = font(color="FFFFFF", size=6)
        m = f"MATCH($H{r},{RK},0)"
        for col, src in zip("ABCDEFG", "CBDLONP"):
            dash[f"{col}{r}"] = f'=IF($H{r}="","",INDEX({KR(src)},{m}))'
    style_range(dash, T + 1, T + 8, 1, 7, fmt={4: PCT, 6: PCT, 7: "0"})
    for r in range(T + 1, T + 9): dash[f"E{r}"].font = font(bold=True)
    dash.conditional_formatting.add(f"D{T + 1}:D{T + 8}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="818CF8"))
    status_cf(dash, f"E{T + 1}:E{T + 8}", f"E{T + 1}")
    dash.column_dimensions["G"].width = 13
    # chart
    ch = BarChart(); ch.type = "bar"; ch.title = "Progress vs expected by objective"; ch.height = 9; ch.width = 17
    ch.add_data(Reference(dash, min_col=4, min_row=H, max_row=H + 6), titles_from_data=True)
    ch.add_data(Reference(dash, min_col=6, min_row=H, max_row=H + 6), titles_from_data=True)
    ch.set_categories(Reference(dash, min_col=1, min_row=H + 1, max_row=H + 6)); ch.y_axis.majorGridlines = None
    ch.x_axis.scaling.orientation = "maxMin"; ch.y_axis.scaling.min = 0; ch.y_axis.scaling.max = 1; ch.y_axis.number_format = "0%"
    ch.series[0].graphicalProperties.solidFill = MAIN; ch.series[1].graphicalProperties.solidFill = "C7D2FE"
    ch.gapWidth = 40; ch.legend.position = "b"; ch.legend.overlay = False
    ch.roundedCorners = False; ch.x_axis.delete = False; ch.y_axis.delete = False
    dash.add_chart(ch, "I4")

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Settings: the year, your life/work areas and the people who own goals (just 'Me' is fine).", False),
             ("2.  Objectives: what you want to achieve – pick the area, owner and period (Q1–Q4 or the whole Year). IDs (O1, O2…) and dates fill in by themselves.", False),
             ("3.  Key Results: 2–4 numbers per objective that prove you got there. Pick the objective ID, then type the start value and target.", False),
             ("     It works both ways: 'grow subscribers 300 → 1,000' and 'cut weight 185 → 175 lb' both show 0–100% progress. For yes/no results use start 0, target 1.", False),
             ("4.  Check-ins (weekly): the new value of each key result and your confidence (1–10). Current value, progress and status update automatically.", False),
             ("5.  Actions: the to-dos that move each key result, with due dates. Overdue ones turn red.", False),
             ("6.  Dashboard: overall progress vs where you should be by today, every objective, and the key results that need attention.", False),
             ("", False), ("How the status works", True),
             ("Expected = how much of the period has passed. On track = progress within 10% of expected (change in Settings). At risk = within 25%. Behind = further back. Missed = due date passed. Done = 100%.", False),
             ("", False), ("Tips", True),
             ("•  Objectives are qualitative ('Launch my online store'); key results are numbers ('50 sales'). 70% on an ambitious goal is a good result.", False),
             ("•  Use 'Weight' to make one key result count more than the others inside an objective.", False),
             ("•  Pick a period in Settings > 'Show on dashboard' to focus on one quarter.", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  40 objectives, 160 key results, 1,500 check-ins, 300 actions. No macros. Works in Excel 2010+ and Google Sheets.", False),
             ("•  To use Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("Six made-up personal goals for 2026 with key results, weekly check-ins and actions. Use the blank copy in your download for your own goals.", False)]
    K.start(st, "Goal Planner & OKR Tracker", "Objectives, key results, weekly check-ins and actions – personal goals or a small team, Excel and Google Sheets.", lines)
    _orig_save = wb.save

    def _save(p):
        st.page_setup.fitToHeight = 1
        _orig_save(p)
    wb.save = _save
    finish(wb, sample, path)


META = {
    "name": "Goal Planner & OKR Tracker (Objectives, Key Results, Weekly Check-ins, Progress Dashboard) – Excel & Google Sheets",
    "file_base": "Goal-Planner-OKR-Tracker",
    "price": 15,
    "color": "#4338CA",
    "title": "Goal Planner & OKR Tracker",
    "tagline": "objectives · key results · weekly check-ins · progress vs pace",
    "intro": "Turn big goals into numbers you can track. Write your objectives, add 2–4 measurable key results to each, and do a two-minute check-in every week – the tracker shows your progress, how far along you should be by today, and which goals are on track, at risk or falling behind. Built for personal goals, side businesses and small teams.",
    "features": [
        "Objectives by area (career, health, money, learning…), owner and period (quarter or full year) – up to 40",
        "Key results with start value, target and weight – works for 'increase to' and 'reduce to' goals and yes/no milestones",
        "Weekly check-ins with value, confidence (1–10) and notes; current value and progress update automatically",
        "Expected-by-today pace and automatic status: Done, On track, At risk, Behind, Missed",
        "Reminder when a key result hasn't been updated for 2 weeks",
        "Actions list linked to key results, with overdue and due-this-week flags",
        "Dashboard: average progress vs expected, objective list with progress bars, key results that need attention, chart",
        "Filter the dashboard by quarter or year",
        "Blank copy plus a sample copy with six example goals",
    ],
    "tags": ["okr template", "goal tracker", "goal planner", "okr tracker", "key results", "2027 goals", "quarterly goals",
             "google sheets", "excel template", "new year goals", "goal setting spreadsheet", "team okr"],
    "summary": "Instant download: Excel + Google Sheets template, blank copy + sample copy.",
}
