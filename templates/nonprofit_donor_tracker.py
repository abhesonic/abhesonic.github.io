"""Nonprofit Donor & Donation Tracker – donors, every gift, campaigns with goals, pledges, thank-you/receipt tracking,
donor retention (new / retained / LYBUNT / lapsed), year-end giving statements and a dashboard.

Settings -> Donors -> Donations -> Campaigns -> Pledges -> Giving Statement -> Dashboard.
"""
import datetime as dt
import random
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, DataBarRule
from openpyxl.chart import BarChart, Reference
from _common import Kit, F, B, DATE, MONEY, font, fill, style_range, input_cell, finish

MAIN = "BE185D"; SOFT = "FCE7F3"
K = Kit(MAIN, SOFT)
ND = 600; D0 = 6; D1 = D0 + ND - 1
NG = 3000; G0 = 6; G1 = G0 + NG - 1
NC = 40; C0 = 6; C1 = C0 + NC - 1
NP = 200; P0 = 6; P1 = P0 + NP - 1
NS = 12; S0 = 15; S1 = S0 + NS - 1
PCT = "0%"; USD = '$#,##0;($#,##0);"-"'; USD2 = '$#,##0.00;($#,##0.00);"-"'
TYPES = ["Individual", "Family", "Business", "Foundation", "Church / group", "Government"]
METHODS = ["Card (online)", "Check", "Cash", "Bank transfer", "PayPal", "Stock", "In-kind goods"]
FUNDS = ["General (unrestricted)", "Building fund", "Scholarships", "Food program", "Youth program", "Emergency relief"]
GF = dict(color="9CA3AF", size=7)


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); gv = wb.create_sheet("Donations"); dn = wb.create_sheet("Donors")
    cp = wb.create_sheet("Campaigns"); pl = wb.create_sheet("Pledges"); gs = wb.create_sheet("Giving Statement")
    s = wb.create_sheet("Settings")

    # ---------- Settings ----------
    K.title(s, "Settings", "Yellow cells are yours to edit.")
    SET = [("Organization name", None), ("Year to report on (empty = this year)", None), ("Major donor: gave at least this much in a year ($)", 1000),
           ("Address / tax ID line for statements", None),
           ("Statement note", "No goods or services were provided in exchange for these contributions, except where noted. Please keep this statement for your tax records.")]
    for i, (lab, v) in enumerate(SET, 5):
        s[f"A{i}"] = lab; s[f"A{i}"].font = font(bold=True); s[f"A{i}"].border = B; input_cell(s[f"C{i}"], v)
    s["C6"].number_format = "0"; s["C7"].number_format = USD
    s["C9"].alignment = Alignment(wrap_text=True, vertical="top"); s.row_dimensions[9].height = 42
    s["A11"] = "Year used (automatic)"; s["A11"].font = font(italic=True)
    s["C11"] = "=IF(ISNUMBER(C6),C6,YEAR(TODAY()))"; s["C11"].number_format = "0"; s["C11"].fill = fill("F3F4F6"); s["C11"].border = B
    K.hdr(s, S0 - 1, ["Donor types"], col0=1); K.hdr(s, S0 - 1, ["Payment methods"], col0=3); K.hdr(s, S0 - 1, ["Funds / designations"], col0=5)
    for i in range(NS):
        input_cell(s[f"A{S0 + i}"], TYPES[i] if i < len(TYPES) else None)
        input_cell(s[f"C{S0 + i}"], METHODS[i] if i < len(METHODS) else None)
        input_cell(s[f"E{S0 + i}"], FUNDS[i] if i < len(FUNDS) else None)
    for c, w in zip("ABCDE", [44, 3, 60, 3, 28]): s.column_dimensions[c].width = w
    YR = "Settings!$C$11"; MAJ = "Settings!$C$7"

    DN = lambda c: f"Donors!${c}${D0}:${c}${D1}"
    GV = lambda c: f"Donations!${c}${G0}:${c}${G1}"
    CP = lambda c: f"Campaigns!${c}${C0}:${c}${C1}"
    PL = lambda c: f"Pledges!${c}${P0}:${c}${P1}"
    SEL = "'Giving Statement'!$C$4"; SELY = "'Giving Statement'!$C$7"

    # ---------- Donors ----------
    K.title(dn, "Donors", "One row per donor (person, family, business or foundation). Totals, status and level fill in from the Donations tab.")
    cols = ["Donor name", "Type", "Email", "Phone", "Street address", "City, state, ZIP", "Notes", "First gift", "Last gift", "Gifts",
            "Lifetime total", "This year", "Last year", "Largest gift", "Status", "Level", "Thank-yous to send", "(rank)", "(win-back)"]
    K.hdr(dn, 5, cols, [26, 13, 26, 14, 24, 22, 24, 11, 11, 7, 12, 12, 12, 11, 20, 9, 10, 5, 5])
    style_range(dn, D0, D1, 1, 19, inputs=range(1, 8), fmt={8: DATE, 9: DATE, 10: "0", 11: USD, 12: USD, 13: USD, 14: USD, 17: "0"})
    for r in range(D0, D1 + 1):
        nm = f'{GV("B")},A{r}'
        dn[f"J{r}"] = f'=IF(A{r}="","",COUNTIFS({nm}))'
        dn[f"H{r}"] = f'=IF(N(J{r})=0,"",_xlfn.MINIFS({GV("A")},{nm}))'
        dn[f"I{r}"] = f'=IF(N(J{r})=0,"",_xlfn.MAXIFS({GV("A")},{nm}))'
        dn[f"K{r}"] = f'=IF(A{r}="","",SUMIFS({GV("C")},{nm}))'
        dn[f"L{r}"] = f'=IF(A{r}="","",SUMIFS({GV("C")},{nm},{GV("L")},{YR}))'
        dn[f"M{r}"] = f'=IF(A{r}="","",SUMIFS({GV("C")},{nm},{GV("L")},{YR}-1))'
        dn[f"N{r}"] = f'=IF(N(J{r})=0,"",_xlfn.MAXIFS({GV("C")},{nm}))'
        dn[f"O{r}"] = (f'=IF(A{r}="","",IF(N(J{r})=0,"Prospect",IF(AND(L{r}>0,M{r}>0),"Retained",IF(AND(L{r}>0,YEAR(H{r})={YR}),"New",'
                       f'IF(L{r}>0,"Came back",IF(M{r}>0,"Gave last year – not yet","Lapsed"))))))')
        dn[f"P{r}"] = f'=IF(A{r}="","",IF(MAX(N(L{r}),N(M{r}))>={MAJ},"Major",IF(N(J{r})>0,"Donor","")))'
        dn[f"Q{r}"] = f'=IF(A{r}="","",COUNTIFS({nm},{GV("H")},"<>Yes",{GV("C")},">0"))'
        dn[f"R{r}"] = f'=IF(N(L{r})>0,L{r}+ROW()/1000000,"")'
        dn[f"S{r}"] = f'=IF(O{r}="Gave last year – not yet",M{r}+ROW()/1000000,"")'
        for c in "RS": dn[f"{c}{r}"].font = font(**GF)
        dn[f"O{r}"].font = font(bold=True)
    dv = DataValidation(type="list", formula1=f"=Settings!$A${S0}:$A${S1}", allow_blank=True); dn.add_data_validation(dv); dv.add(f"B{D0}:B{D1}")
    for txt, fg, bg in (("Retained", "166534", "DCFCE7"), ("New", "1D4ED8", "DBEAFE"), ("Came back", "166534", None),
                        ("Gave last year – not yet", "92400E", "FEF3C7"), ("Lapsed", "6B7280", None)):
        kw = dict(font=Font(name=F, bold=True, color=fg))
        if bg: kw["fill"] = fill(bg)
        dn.conditional_formatting.add(f"O{D0}:O{D1}", FormulaRule(formula=[f'O{D0}="{txt}"'], **kw))
    dn.conditional_formatting.add(f"P{D0}:P{D1}", FormulaRule(formula=[f'P{D0}="Major"'], font=Font(name=F, bold=True, color=MAIN), fill=fill(SOFT)))
    dn.conditional_formatting.add(f"Q{D0}:Q{D1}", FormulaRule(formula=[f'N(Q{D0})>0'], font=Font(name=F, bold=True, color="B91C1C")))
    dn.freeze_panes = f"B{D0}"; dn.auto_filter.ref = f"A5:Q{D1}"

    # ---------- Donations ----------
    K.title(gv, "Donations", "One row per gift. Pick the donor from the list (add new donors on the Donors tab first). In-kind gifts: enter the fair value.")
    cols = ["Date", "Donor", "Amount ($)", "Payment method", "Campaign / appeal", "Fund", "Receipt sent?", "Thank-you sent?", "Notes",
            "Donor type", "Month", "Year", "To do", "(statement)"]
    K.hdr(gv, 5, cols, [12, 26, 12, 15, 24, 22, 10, 11, 28, 13, 8, 7, 18, 6])
    style_range(gv, G0, G1, 1, 14, inputs=range(1, 10), fmt={1: DATE, 3: USD2, 11: "0", 12: "0"})
    for r in range(G0, G1 + 1):
        gv[f"J{r}"] = f'=IF(B{r}="","",IFERROR(INDEX({DN("B")},MATCH(B{r},{DN("A")},0))&"","(not on Donors tab)"))'
        gv[f"K{r}"] = f'=IF(ISNUMBER(A{r}),MONTH(A{r}),"")'
        gv[f"L{r}"] = f'=IF(ISNUMBER(A{r}),YEAR(A{r}),"")'
        gv[f"M{r}"] = (f'=IF(OR(B{r}="",N(C{r})=0),"",IF(AND(G{r}="Yes",H{r}="Yes"),"",IF(H{r}<>"Yes",IF(G{r}<>"Yes","Thank-you + receipt","Send thank-you"),"Send receipt")))')
        gv[f"N{r}"] = f'=IF(AND(B{r}<>"",B{r}={SEL},L{r}={SELY}),A{r}+ROW()/10000000,"")'; gv[f"N{r}"].font = font(**GF)
    for col, f1 in (("B", f"=Donors!$A${D0}:$A${D1}"), ("D", f"=Settings!$C${S0}:$C${S1}"), ("E", f"=Campaigns!$A${C0}:$A${C1}"),
                    ("F", f"=Settings!$E${S0}:$E${S1}"), ("G", '"Yes"'), ("H", '"Yes"')):
        dv = DataValidation(type="list", formula1=f1, allow_blank=True); gv.add_data_validation(dv); dv.add(f"{col}{G0}:{col}{G1}")
    gv.conditional_formatting.add(f"M{G0}:M{G1}", FormulaRule(formula=[f'M{G0}<>""'], font=Font(name=F, bold=True, color="92400E"), fill=fill("FEF3C7")))
    gv.conditional_formatting.add(f"J{G0}:J{G1}", FormulaRule(formula=[f'J{G0}="(not on Donors tab)"'], font=Font(name=F, bold=True, color="B91C1C")))
    gv.freeze_panes = f"C{G0}"; gv.auto_filter.ref = f"A5:M{G1}"

    # ---------- Campaigns ----------
    K.title(cp, "Campaigns", "Fundraising appeals and events with a goal: year-end appeal, gala, Giving Tuesday, capital campaign…")
    cols = ["Campaign / appeal", "Goal ($)", "Start", "End", "Notes", "Raised", "Gifts", "Pledged, not yet paid", "% of goal", "Status"]
    K.hdr(cp, 5, cols, [30, 12, 11, 11, 28, 12, 7, 12, 10, 14])
    style_range(cp, C0, C1, 1, 10, inputs=range(1, 6), fmt={2: USD, 3: DATE, 4: DATE, 6: USD, 7: "0", 8: USD, 9: PCT})
    for r in range(C0, C1 + 1):
        cp[f"F{r}"] = f'=IF(A{r}="","",SUMIFS({GV("C")},{GV("E")},A{r}))'
        cp[f"G{r}"] = f'=IF(A{r}="","",COUNTIFS({GV("E")},A{r},{GV("C")},">0"))'
        cp[f"H{r}"] = f'=IF(A{r}="","",SUMIFS({PL("H")},{PL("B")},A{r}))'
        cp[f"I{r}"] = f'=IF(OR(A{r}="",N(B{r})=0),"",F{r}/B{r})'
        cp[f"J{r}"] = (f'=IF(A{r}="","",IF(AND(N(B{r})>0,F{r}>=B{r}),"Goal reached",IF(AND(ISNUMBER(C{r}),TODAY()<C{r}),"Upcoming",'
                       f'IF(AND(ISNUMBER(D{r}),TODAY()>D{r}),"Ended","Active"))))')
        cp[f"J{r}"].font = font(bold=True)
    cp.conditional_formatting.add(f"I{C0}:I{C1}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="F472B6"))
    cp.conditional_formatting.add(f"J{C0}:J{C1}", FormulaRule(formula=[f'J{C0}="Goal reached"'], font=Font(name=F, bold=True, color="166534"), fill=fill("DCFCE7")))
    cp.conditional_formatting.add(f"J{C0}:J{C1}", FormulaRule(formula=[f'J{C0}="Active"'], font=Font(name=F, bold=True, color="1D4ED8")))
    cp.freeze_panes = f"B{C0}"

    # ---------- Pledges ----------
    K.title(pl, "Pledges", "Promised gifts. Payments are matched automatically: donations from the same donor to the same campaign on or after the pledge date.")
    cols = ["Donor", "Campaign / appeal", "Pledged ($)", "Pledge date", "Due by", "Notes", "Paid so far", "Balance", "Status"]
    K.hdr(pl, 5, cols, [26, 26, 12, 11, 11, 26, 12, 12, 14])
    style_range(pl, P0, P1, 1, 9, inputs=range(1, 7), fmt={3: USD, 4: DATE, 5: DATE, 7: USD, 8: USD})
    for r in range(P0, P1 + 1):
        pl[f"G{r}"] = f'=IF(A{r}="","",MIN(N(C{r}),SUMIFS({GV("C")},{GV("B")},A{r},{GV("E")},B{r},{GV("A")},">="&N(D{r}))))'
        pl[f"H{r}"] = f'=IF(A{r}="","",MAX(0,N(C{r})-G{r}))'
        pl[f"I{r}"] = (f'=IF(A{r}="","",IF(H{r}=0,"Paid in full",IF(AND(ISNUMBER(E{r}),TODAY()>E{r}),"Overdue",IF(G{r}>0,"Partly paid","Open"))))')
        pl[f"I{r}"].font = font(bold=True)
    dv = DataValidation(type="list", formula1=f"=Donors!$A${D0}:$A${D1}", allow_blank=True); pl.add_data_validation(dv); dv.add(f"A{P0}:A{P1}")
    dv = DataValidation(type="list", formula1=f"=Campaigns!$A${C0}:$A${C1}", allow_blank=True); pl.add_data_validation(dv); dv.add(f"B{P0}:B{P1}")
    pl.conditional_formatting.add(f"I{P0}:I{P1}", FormulaRule(formula=[f'I{P0}="Overdue"'], font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    pl.conditional_formatting.add(f"I{P0}:I{P1}", FormulaRule(formula=[f'I{P0}="Paid in full"'], font=Font(name=F, bold=True, color="166534"), fill=fill("DCFCE7")))
    pl.freeze_panes = f"B{P0}"

    # ---------- sample ----------
    if sample:
        rnd = random.Random(7)
        s["C5"] = "Brightside Community Center"; s["C6"] = 2026
        s["C8"] = "1200 Maple Avenue, Springfield · Tax ID 00-0000000 (sample)"
        donors = [("Johnson Family", "Family", "Springfield, IL 62701"), ("Maria Gonzalez", "Individual", "Springfield, IL 62702"),
                  ("Green Valley Bakery", "Business", "Springfield, IL 62704"), ("The Harper Foundation", "Foundation", "Chicago, IL 60601"),
                  ("David Chen", "Individual", "Springfield, IL 62703"), ("St. Mark's Church", "Church / group", "Springfield, IL 62701"),
                  ("Linda & Tom Brooks", "Family", "Chatham, IL 62629"), ("Aisha Rahman", "Individual", "Springfield, IL 62711"),
                  ("Peterson Auto Repair", "Business", "Springfield, IL 62702"), ("Robert Miller", "Individual", "Rochester, IL 62563"),
                  ("Emily Nguyen", "Individual", "Springfield, IL 62704"), ("Rotary Club of Springfield", "Church / group", "Springfield, IL 62701"),
                  ("James O'Connor", "Individual", "Springfield, IL 62707"), ("Sarah Williams", "Individual", "Sherman, IL 62684"),
                  ("Lakeside Dental", "Business", "Springfield, IL 62703"), ("Patel Family", "Family", "Springfield, IL 62712"),
                  ("Karen Lewis", "Individual", "Springfield, IL 62702"), ("Michael Brown", "Individual", "Springfield, IL 62704"),
                  ("City Community Grant", "Government", "Springfield, IL 62701"), ("Olivia Martin", "Individual", "Springfield, IL 62703"),
                  ("Daniel Kim", "Individual", "Springfield, IL 62711"), ("Hannah Scott", "Individual", "Chatham, IL 62629")]
        for k, (n, t, c) in enumerate(donors):
            r = D0 + k; dn[f"A{r}"] = n; dn[f"B{r}"] = t; dn[f"F{r}"] = c
            first = n.split()[0].lower().replace("'", "")
            if t in ("Individual", "Family"):
                dn[f"C{r}"] = f"{first}@example.com"; dn[f"D{r}"] = f"(217) 555-{1000 + k * 37:04d}"
                dn[f"E{r}"] = f"{100 + k * 23} {['Oak', 'Elm', 'Pine', 'Cedar', 'Birch'][k % 5]} St"
            else:
                dn[f"C{r}"] = f"info@{first}.example.org"; dn[f"E{r}"] = f"{20 + k * 11} Main St"
        dn[f"G{D0 + 3}"] = "Grant report due Jan 31"; dn[f"G{D0 + 1}"] = "Monthly donor since 2024"; dn[f"G{D0 + 20}"] = "Met at the gala – follow up"
        camps = [("Spring Gala 2025", 15000, dt.date(2025, 3, 1), dt.date(2025, 4, 30)), ("Year-End Appeal 2025", 20000, dt.date(2025, 11, 1), dt.date(2025, 12, 31)),
                 ("Spring Gala 2026", 18000, dt.date(2026, 3, 1), dt.date(2026, 4, 30)), ("Summer Food Drive 2026", 6000, dt.date(2026, 6, 1), dt.date(2026, 8, 31)),
                 ("Giving Tuesday 2026", 5000, dt.date(2026, 12, 1), dt.date(2026, 12, 1)), ("Year-End Appeal 2026", 25000, dt.date(2026, 11, 1), dt.date(2026, 12, 31)),
                 ("New Roof Fund", 40000, dt.date(2026, 1, 1), dt.date(2027, 6, 30)), ("Monthly giving", None, None, None)]
        for k, (n, g, a, b) in enumerate(camps):
            r = C0 + k; cp[f"A{r}"] = n
            if g: cp[f"B{r}"] = g
            if a: cp[f"C{r}"] = a; cp[f"D{r}"] = b
        gifts = []
        def add(d, who, amt, meth, camp, fund, note=""):
            gifts.append((d, who, amt, meth, camp, fund, note))
        for y in (2025, 2026):
            last_m = 12 if y == 2025 else 10
            for m in range(1, last_m + 1): add(dt.date(y, m, 5), "Maria Gonzalez", 50, "Card (online)", "Monthly giving", FUNDS[0])
            for m in range(1, last_m + 1, 3): add(dt.date(y, m, 15), "Emily Nguyen", 75, "PayPal", "Monthly giving", FUNDS[0])
        add(dt.date(2025, 4, 12), "The Harper Foundation", 10000, "Bank transfer", "Spring Gala 2025", "Youth program", "Program grant")
        add(dt.date(2026, 2, 20), "The Harper Foundation", 12500, "Bank transfer", "New Roof Fund", "Building fund", "Year 1 of 2")
        add(dt.date(2026, 7, 1), "City Community Grant", 8000, "Bank transfer", "Summer Food Drive 2026", "Food program", "Restricted grant")
        add(dt.date(2025, 7, 10), "City Community Grant", 7500, "Bank transfer", "", "Food program", "")
        gala25 = ["Johnson Family", "Green Valley Bakery", "David Chen", "Linda & Tom Brooks", "Peterson Auto Repair", "Robert Miller",
                  "Rotary Club of Springfield", "Sarah Williams", "Karen Lewis", "Michael Brown", "James O'Connor"]
        for i, who in enumerate(gala25):
            add(dt.date(2025, 4, 5 + i % 20), who, rnd.choice([100, 150, 250, 500, 1000]), rnd.choice(METHODS[:5]), "Spring Gala 2025", FUNDS[0])
        ye25 = ["Johnson Family", "David Chen", "St. Mark's Church", "Linda & Tom Brooks", "Aisha Rahman", "Robert Miller", "Patel Family",
                "Karen Lewis", "Sarah Williams", "Lakeside Dental", "James O'Connor", "Michael Brown"]
        for i, who in enumerate(ye25):
            add(dt.date(2025, 11 + i % 2, 3 + i * 2), who, rnd.choice([100, 200, 250, 500, 750, 1500]), rnd.choice(METHODS[:5]), "Year-End Appeal 2025", rnd.choice(FUNDS[:3]))
        gala26 = ["Johnson Family", "Green Valley Bakery", "David Chen", "Linda & Tom Brooks", "Peterson Auto Repair", "Rotary Club of Springfield",
                  "Olivia Martin", "Daniel Kim", "Patel Family", "Aisha Rahman", "Lakeside Dental"]
        for i, who in enumerate(gala26):
            add(dt.date(2026, 4, 3 + i % 22), who, rnd.choice([150, 250, 500, 1000, 1200]), rnd.choice(METHODS[:5]), "Spring Gala 2026", FUNDS[0])
        add(dt.date(2026, 4, 18), "Green Valley Bakery", 600, "In-kind goods", "Spring Gala 2026", FUNDS[0], "Desserts for 120 guests (fair value)")
        for i, who in enumerate(["St. Mark's Church", "Hannah Scott", "Johnson Family", "Patel Family", "Rotary Club of Springfield"]):
            add(dt.date(2026, 6, 1) + dt.timedelta(days=10 + i * 12), who, rnd.choice([100, 250, 400]), rnd.choice(METHODS[:4]), "Summer Food Drive 2026", "Food program")
        for i, who in enumerate(["David Chen", "Linda & Tom Brooks", "Johnson Family", "Daniel Kim"]):
            add(dt.date(2026, 3 + i * 2, 12), who, rnd.choice([500, 1000, 2000]), "Check", "New Roof Fund", "Building fund")
        for i, who in enumerate(["Aisha Rahman", "Olivia Martin", "Hannah Scott", "Johnson Family", "Patel Family", "St. Mark's Church"]):
            add(dt.date(2026, 11, 2 + i * 2), who, rnd.choice([100, 150, 250, 500]), rnd.choice(METHODS[:5]), "Year-End Appeal 2026", rnd.choice(FUNDS[:3]))
        gifts.sort(key=lambda g: g[0])
        today_ref = dt.date(2026, 11, 1)
        for k, (d, who, amt, meth, camp, fund, note) in enumerate(gifts):
            r = G0 + k
            gv[f"A{r}"] = d; gv[f"B{r}"] = who; gv[f"C{r}"] = amt; gv[f"D{r}"] = meth
            if camp: gv[f"E{r}"] = camp
            gv[f"F{r}"] = fund
            if d < today_ref or (k % 3 == 0): gv[f"G{r}"] = "Yes"
            if d < today_ref: gv[f"H{r}"] = "Yes"
            if note: gv[f"I{r}"] = note
        pledges = [("The Harper Foundation", "New Roof Fund", 25000, dt.date(2026, 1, 15), dt.date(2027, 3, 31), "2 payments"),
                   ("Rotary Club of Springfield", "Year-End Appeal 2026", 2000, dt.date(2026, 10, 20), dt.date(2026, 12, 31), "Pledged at October meeting"),
                   ("Peterson Auto Repair", "New Roof Fund", 1500, dt.date(2026, 3, 1), dt.date(2026, 9, 30), ""),
                   ("David Chen", "New Roof Fund", 3000, dt.date(2026, 2, 1), dt.date(2026, 12, 31), "")]
        for k, row in enumerate(pledges):
            for c, v in zip("ABCDEF", row):
                if v: pl[f"{c}{P0 + k}"] = v

    # ---------- Giving Statement ----------
    K.title(gs, "Giving Statement", "Year-end statement / tax receipt for one donor. Pick the donor and year, then print or save as PDF.")
    gs["B4"] = "Donor  ▸"; gs["B4"].font = font(bold=True); gs["B4"].alignment = Alignment(horizontal="right")
    input_cell(gs["C4"], "Johnson Family" if sample else None)
    dv = DataValidation(type="list", formula1=f"=Donors!$A${D0}:$A${D1}", allow_blank=True); gs.add_data_validation(dv); dv.add("C4")
    gs["B5"] = "Year  ▸"; gs["B5"].font = font(bold=True); gs["B5"].alignment = Alignment(horizontal="right")
    input_cell(gs["C5"]); gs["C5"].number_format = "0"
    gs["D5"] = "(empty = the year in Settings)"; gs["D5"].font = font(italic=True, size=9, color="5B6770")
    gs["C7"] = f"=IF(ISNUMBER(C5),C5,{YR})"; gs["C7"].number_format = "0"; gs["C7"].font = font(color="FFFFFF", size=6)
    gs["B9"] = '=IF(Settings!$C$5="","Your organization (set the name in Settings)",Settings!$C$5)'; gs["B9"].font = font(bold=True, size=16, color=MAIN)
    gs["B10"] = '=IF(Settings!$C$8="","",Settings!$C$8)'; gs["B10"].font = font(size=9, color="5B6770")
    gs["B12"] = '="Giving statement for "&C7'; gs["B12"].font = font(bold=True, size=13)
    gs["B14"] = '=IF(C4="","(pick a donor above)",C4)'; gs["B14"].font = font(bold=True, size=11)
    dm = f"MATCH($C$4,{DN('A')},0)"
    gs["B15"] = f'=IF($C$4="","",IFERROR(INDEX({DN("E")},{dm})&"",""))'
    gs["B16"] = f'=IF($C$4="","",IFERROR(INDEX({DN("F")},{dm})&"",""))'
    for ref in ("B15", "B16"): gs[ref].font = font()
    K.hdr(gs, 18, ["Date", "Amount", "Payment method", "Campaign / appeal", "Fund", "Notes"], col0=2)
    NK = GV("N"); ROWS = 40
    for k in range(ROWS):
        r = 19 + k; m = f"MATCH(SMALL({NK},{k + 1}),{NK},0)"
        for col, src in zip("BCDEFG", "ACDEFI"):
            gs[f"{col}{r}"] = f'=IF(COUNT({NK})<{k + 1},"",IF(INDEX({GV(src)},{m})="","",INDEX({GV(src)},{m})))'
    style_range(gs, 19, 18 + ROWS, 2, 7, fmt={2: DATE, 3: USD2})
    TRW = 19 + ROWS
    gs[f"B{TRW}"] = "Total"; gs[f"B{TRW}"].font = font(bold=True, size=12)
    gs[f"C{TRW}"] = f"=SUM(C19:C{18 + ROWS})"; gs[f"C{TRW}"].number_format = USD2; gs[f"C{TRW}"].font = font(bold=True, size=12, color=MAIN)
    gs[f"D{TRW}"] = f'=IF(COUNT({NK})>{ROWS},"Note: more than {ROWS} gifts – see the Donations tab",COUNT({NK})&" gift(s)")'; gs[f"D{TRW}"].font = font(italic=True, size=9)
    gs[f"B{TRW + 1}"] = f'=IF(SUMIFS({GV("C")},{GV("B")},$C$4,{GV("L")},$C$7,{GV("D")},"In-kind goods")>0,"Includes in-kind gifts valued at "&TEXT(SUMIFS({GV("C")},{GV("B")},$C$4,{GV("L")},$C$7,{GV("D")},"In-kind goods"),"$#,##0.00")&" (fair value as stated by the donor).","")'
    gs[f"B{TRW + 1}"].font = font(size=9, italic=True)
    gs[f"B{TRW + 3}"] = '=Settings!$C$9&""'; gs[f"B{TRW + 3}"].font = font(size=10); gs[f"B{TRW + 3}"].alignment = Alignment(wrap_text=True, vertical="top")
    gs.merge_cells(f"B{TRW + 3}:G{TRW + 4}"); gs.row_dimensions[TRW + 3].height = 30
    gs[f"B{TRW + 6}"] = "Thank you for your generous support!"; gs[f"B{TRW + 6}"].font = font(bold=True, color=MAIN, size=12)
    for c, w in zip("ABCDEFG", [3, 13, 13, 16, 26, 22, 30]): gs.column_dimensions[c].width = w

    # ---------- Dashboard ----------
    K.title(dash, "Donor & Fundraising Dashboard", "Updates by itself from Donations, Donors, Campaigns and Pledges.")
    dash["A3"] = '=IF(Settings!$C$5="","",Settings!$C$5&"  ·  ")&"Year: "&Settings!$C$11&" compared with "&(Settings!$C$11-1)&"  (change the year in Settings)"'
    dash["A3"].font = font(bold=True)
    yr = f'{GV("L")},{YR}'
    kp = [("Raised this year", f'=SUMIFS({GV("C")},{yr})', USD),
          ("Raised last year", f'=SUMIFS({GV("C")},{GV("L")},{YR}-1)', USD),
          ("Change vs all of last year", f'=IF(C6=0,"",A6/C6-1)', '+0%;-0%;0%'),
          ("Gifts this year", f'=COUNTIFS({yr},{GV("C")},">0")', "0"),
          ("Donors this year", f'=COUNTIFS({DN("L")},">0")', "0"),
          ("New donors", f'=COUNTIFS({DN("O")},"New")', "0"),
          ("Donor retention", f'=IF(COUNTIFS({DN("M")},">0")=0,"",COUNTIFS({DN("O")},"Retained")/COUNTIFS({DN("M")},">0"))', PCT),
          ("Average gift this year", f'=IF(G6=0,0,A6/G6)', USD),
          ("Major donors", f'=COUNTIFS({DN("P")},"Major",{DN("L")},">0")', "0"),
          ("Thank-yous / receipts to send", f'=COUNTIFS({GV("M")},"?*")', "0"),
          ("Pledges still to collect", f'=SUMIFS({PL("H")},{PL("A")},"?*")', USD),
          ("Donors to win back", f'=COUNTIFS({DN("O")},"Gave last year – not yet")', "0")]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 4) * 3, 1 + (i % 4) * 2, lab, f, fmt)
        dash.cell(row=6 + (i // 4) * 3, column=1 + (i % 4) * 2).alignment = Alignment(horizontal="left")
    dash.conditional_formatting.add("C12", FormulaRule(formula=["C12>0"], font=Font(name=F, size=18, bold=True, color="B45309")))
    dash.conditional_formatting.add("E6", FormulaRule(formula=['AND(ISNUMBER(E6),E6<0)'], font=Font(name=F, size=18, bold=True, color="B91C1C")))
    dash.conditional_formatting.add("E6", FormulaRule(formula=['AND(ISNUMBER(E6),E6>0)'], font=Font(name=F, size=18, bold=True, color="166534")))
    for c, w in zip("ABCDEFGHIJ", [26, 13, 24, 13, 22, 13, 24, 13, 3, 3]): dash.column_dimensions[c].width = w
    # monthly table
    MR = 16
    K.section(dash, f"A{MR - 1}", "Donations by month")
    K.hdr(dash, MR, ["Month", "This year", "Last year"])
    for k in range(12):
        r = MR + 1 + k
        dash[f"A{r}"] = dt.date(2026, k + 1, 1).strftime("%b")
        dash[f"B{r}"] = f'=SUMIFS({GV("C")},{GV("L")},{YR},{GV("K")},{k + 1})'
        dash[f"C{r}"] = f'=SUMIFS({GV("C")},{GV("L")},{YR}-1,{GV("K")},{k + 1})'
    dash[f"B{MR}"] = f"={YR}&\"\""; dash[f"C{MR}"] = f"=({YR}-1)&\"\""
    style_range(dash, MR + 1, MR + 12, 1, 3, fmt={2: USD, 3: USD})
    # by fund
    K.section(dash, f"E{MR - 1}", "This year by fund")
    K.hdr(dash, MR, ["Fund", "Raised", "Share"], col0=5)
    for k in range(8):
        r = MR + 1 + k
        dash[f"E{r}"] = f'=IF(Settings!$E${S0 + k}="","",Settings!$E${S0 + k})'
        dash[f"F{r}"] = f'=IF(E{r}="","",SUMIFS({GV("C")},{yr},{GV("F")},E{r}))'
        dash[f"G{r}"] = f'=IF(OR(E{r}="",$A$6=0),"",F{r}/$A$6)'
    style_range(dash, MR + 1, MR + 8, 5, 7, fmt={6: USD, 7: PCT})
    dash.conditional_formatting.add(f"G{MR + 1}:G{MR + 8}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="F472B6"))
    # campaigns
    CR = MR + 15
    K.section(dash, f"A{CR - 1}", "Campaigns")
    K.hdr(dash, CR, ["Campaign", "Goal", "Raised", "% of goal", "Status", "Pledged, unpaid"])
    for k in range(10):
        r = CR + 1 + k; src = C0 + k
        dash[f"A{r}"] = f'=IF(Campaigns!$A${src}="","",Campaigns!$A${src})'
        for col, sc in zip("BCDEF", "BFIJH"):
            dash[f"{col}{r}"] = f'=IF($A{r}="","",Campaigns!${sc}${src})'
    style_range(dash, CR + 1, CR + 10, 1, 6, fmt={2: USD, 3: USD, 4: PCT, 6: USD})
    dash.conditional_formatting.add(f"D{CR + 1}:D{CR + 10}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="F472B6"))
    dash.conditional_formatting.add(f"E{CR + 1}:E{CR + 10}", FormulaRule(formula=[f'E{CR + 1}="Goal reached"'], font=Font(name=F, bold=True, color="166534")))
    # top donors + win back
    TR = CR + 13
    K.section(dash, f"A{TR - 1}", "Top donors this year")
    K.hdr(dash, TR, ["Donor", "This year", "Last year", "Status"])
    RK = DN("R")
    for k in range(10):
        r = TR + 1 + k; m = f"MATCH(LARGE({RK},{k + 1}),{RK},0)"
        for col, src in zip("ABCD", "ALMO"):
            dash[f"{col}{r}"] = f'=IF(COUNT({RK})<{k + 1},"",INDEX({DN(src)},{m}))'
    style_range(dash, TR + 1, TR + 10, 1, 4, fmt={2: USD, 3: USD})
    K.section(dash, f"E{TR - 1}", "Win back: gave last year, not yet this year")
    K.hdr(dash, TR, ["Donor", "Last year", "Last gift", "Email"], col0=5)
    WK = DN("S")
    for k in range(10):
        r = TR + 1 + k; m = f"MATCH(LARGE({WK},{k + 1}),{WK},0)"
        for col, src in zip("EFGH", "AMIC"):
            dash[f"{col}{r}"] = f'=IF(COUNT({WK})<{k + 1},"",INDEX({DN(src)},{m}))'
    style_range(dash, TR + 1, TR + 10, 5, 8, fmt={6: USD, 7: DATE})
    # chart
    ch = BarChart(); ch.type = "col"; ch.title = "Donations by month"; ch.height = 8; ch.width = 18
    ch.add_data(Reference(dash, min_col=2, max_col=3, min_row=MR, max_row=MR + 12), titles_from_data=True)
    ch.set_categories(Reference(dash, min_col=1, min_row=MR + 1, max_row=MR + 12))
    ch.series[0].graphicalProperties.solidFill = MAIN; ch.series[1].graphicalProperties.solidFill = "F9A8D4"
    ch.y_axis.majorGridlines = None; ch.y_axis.number_format = '$#,##0'; ch.legend.position = "b"; ch.legend.overlay = False; ch.title.overlay = False
    ch.roundedCorners = False; ch.x_axis.delete = False; ch.y_axis.delete = False; ch.gapWidth = 50
    dash.add_chart(ch, "K4")

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Settings: your organization's name, the year to report on, your major-donor level and the address / tax ID line for statements. Edit the lists of donor types, payment methods and funds.", False),
             ("2.  Donors: one row per person, family, business or foundation. Totals, first/last gift, status and level fill in by themselves.", False),
             ("3.  Campaigns: your appeals and events with a goal (year-end appeal, gala, Giving Tuesday, building fund…).", False),
             ("4.  Donations: one row per gift – date, donor (from the list), amount, method, campaign and fund. Tick 'Receipt sent' and 'Thank-you sent' when done.", False),
             ("5.  Pledges: promised gifts. Payments from the same donor to the same campaign are matched automatically and the balance updates.", False),
             ("6.  Giving Statement: pick a donor and year to get a ready-to-print year-end statement listing every gift and the total.", False),
             ("7.  Dashboard: raised this year vs last year, donor retention, new donors, campaign progress, top donors and who to win back.", False),
             ("", False), ("Donor status explained", True),
             ("New = first gift this year.  Retained = gave last year and this year.  Came back = gave this year after a gap.  "
              "Gave last year – not yet = gave last year but not yet this year (your best people to contact).  Lapsed = no gift last year or this year.  Prospect = no gifts yet.", False),
             ("Donor retention = retained donors ÷ donors who gave last year.", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  Room for 600 donors, 3,000 gifts, 40 campaigns and 200 pledges. No macros. Works in Excel 2010+ and Google Sheets.", False),
             ("•  In-kind gifts: choose 'In-kind goods' as the method and enter the fair value; the statement notes them separately.", False),
             ("•  This template helps you keep records; check your country's rules for what a tax receipt must say.", False),
             ("•  To use Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("A made-up community center with 22 donors, two years of gifts, campaigns and pledges. Use the blank copy in your download for your own organization.", False)]
    K.start(st, "Nonprofit Donor & Donation Tracker", "Donors, gifts, campaigns, pledges, thank-yous and year-end giving statements – Excel and Google Sheets.", lines)
    finish(wb, sample, path)


META = {
    "name": "Nonprofit Donor & Donation Tracker (Donor Database, Campaign Goals, Pledges, Year-End Giving Statements) – Excel & Google Sheets",
    "file_base": "Nonprofit-Donor-Donation-Tracker",
    "price": 19,
    "color": "#BE185D",
    "title": "Nonprofit Donor & Donation Tracker",
    "tagline": "donors · gifts · campaigns · pledges · giving statements",
    "intro": "A simple donor database for small nonprofits, churches, schools and clubs – no expensive CRM needed. Log every gift once and the tracker totals each donor, shows who is new, retained or slipping away, tracks campaign goals and pledges, reminds you which thank-you notes and receipts are still to send, and prints a year-end giving statement for any donor in one click.",
    "features": [
        "Donor list with contact details, first and last gift, lifetime total, this year vs last year and largest gift",
        "Automatic donor status: New, Retained, Came back, Gave last year – not yet, Lapsed, Prospect – plus major-donor flag",
        "Donations log with payment method, campaign, fund/designation, receipt and thank-you tracking",
        "Campaigns with goals, amount raised, % of goal and status",
        "Pledges with automatic payment matching, balance and overdue flag",
        "Year-end giving statement / donation receipt for any donor and year – print or save as PDF",
        "Dashboard: raised this year vs last, donor retention rate, average gift, monthly chart, giving by fund, campaign progress, top donors and a win-back list",
        "Works for churches, schools, PTAs, clubs, shelters and charities – in any currency",
        "Blank copy plus a sample copy with two years of example data",
    ],
    "tags": ["donor tracker", "donation tracker", "nonprofit spreadsheet", "church donation", "donor database", "fundraising tracker",
             "giving statement", "donation receipt", "pledge tracker", "google sheets", "excel template", "charity"],
    "summary": "Instant download: Excel + Google Sheets template, blank copy + sample copy.",
}
