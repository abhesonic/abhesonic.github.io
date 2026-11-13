"""Wedding Guest List, RSVP & Seating Planner – guest list with households, RSVP and meal tracking, tables with capacity,
automatic seating chart, unseated list, gifts & thank-yous, dashboard.

Settings -> Guest List -> Tables -> Seating Chart -> Dashboard.
"""
import datetime as dt, random
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, DataBarRule
from openpyxl.chart import BarChart, Reference
from _common import Kit, F, B, MONEY, DATE, font, fill, style_range, input_cell, finish

MAIN = "6D28D9"; SOFT = "EDE9FE"
K = Kit(MAIN, SOFT)
NG = 500; G0 = 6; G1 = G0 + NG - 1           # guests
NT = 40; T0 = 6; T1 = T0 + NT - 1            # tables
NS = 14                                      # max seats shown per table
NL = 12; S0 = 6; S1 = S0 + NL - 1            # list rows in Settings
NUNS = 60                                    # unseated list rows
PCT = "0%"; INT = '0;-0;"-"'
RSVP = ["Attending", "Declined", "Pending"]
SIDES = ["Bride's side", "Groom's side", "Both / mutual"]
GROUPS = ["Family", "Friends", "Work", "Neighbours", "Plus-ones", "Wedding party"]
MEALS = ["Chicken", "Beef", "Fish", "Vegetarian", "Vegan", "Kids meal"]


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); gl = wb.create_sheet("Guest List"); tb = wb.create_sheet("Tables")
    sc = wb.create_sheet("Seating Chart"); s = wb.create_sheet("Settings")

    # ---------- Settings ----------
    K.title(s, "Settings", "Yellow cells are yours to edit. The three lists feed the drop-downs on the Guest List.")
    SET = [("Couple (label only)", "Your names", None), ("Wedding date", None, DATE), ("RSVP deadline", None, DATE),
           ("Catering cost per adult", None, MONEY), ("Catering cost per child", None, MONEY),
           ("Default seats per table", 10, "0"), ("Currency (label only)", "USD", None)]
    for i, (lab, v, fmt) in enumerate(SET, 5):
        s[f"A{i}"] = lab; s[f"A{i}"].font = font(bold=True); s[f"A{i}"].border = B; input_cell(s[f"C{i}"], v)
        if fmt: s[f"C{i}"].number_format = fmt
    K.hdr(s, 5, ["Sides", "Groups", "Meal options"], col0=6)
    style_range(s, S0, S1, 6, 8, inputs=(6, 7, 8))
    for col, vals in zip("FGH", (SIDES, GROUPS, MEALS)):
        for i, v in enumerate(vals): s[f"{col}{S0 + i}"] = v
    for r in range(S0, S1 + 1):
        for c in "FGH": s[f"{c}{r}"].font = Font(name=F, color="0000FF")
    s[f"F{S1 + 1}"] = "Rename freely (e.g. \"Emma's side\"), add up to 12 per list."; s[f"F{S1 + 1}"].font = font(italic=True, size=9)
    for c, w in zip("ABCDEFGH", [30, 3, 16, 3, 3, 18, 16, 16]): s.column_dimensions[c].width = w
    SIDE = f"Settings!$F${S0}:$F${S1}"; GRP = f"Settings!$G${S0}:$G${S1}"; MEAL = f"Settings!$H${S0}:$H${S1}"

    # ---------- Guest List ----------
    K.title(gl, "Guest List", "One row per person (children too). Guests of the same household share a 'Party' name – that is who the invitation goes to.")
    hd = ["Guest name", "Party / household", "Side", "Group", "A or B list", "Adult / Child", "Email / phone", "Mailing address",
          "Save-the-date sent", "Invitation sent", "RSVP", "Meal choice", "Dietary needs / allergies", "Table #",
          "Gift received", "Thank-you sent", "Notes", "Status", "To do"]
    K.hdr(gl, 5, hd, [22, 20, 13, 12, 7, 8, 22, 28, 9, 9, 11, 12, 18, 7, 18, 9, 18, 13, 16])
    gl.row_dimensions[5].height = 44
    style_range(gl, G0, G1, 1, 19, inputs=range(1, 18), fmt={14: "0"})
    for r in range(G0, G1 + 1):
        gl[f"R{r}"] = (f'=IF(A{r}="","",IF(K{r}="Attending","Attending",IF(K{r}="Declined","Declined",'
                       f'IF(J{r}="Yes","Awaiting reply","Not invited yet"))))')
        gl[f"S{r}"] = (f'=IF(A{r}="","",IF(R{r}="Attending",IF(L{r}="","Ask meal choice",IF(NOT(ISNUMBER(N{r})),"Assign a table",'
                       f'IF(AND(O{r}<>"",P{r}<>"Yes"),"Send thank-you",""))),IF(R{r}="Declined",IF(AND(O{r}<>"",P{r}<>"Yes"),"Send thank-you",""),'
                       f'IF(R{r}="Awaiting reply","Follow up RSVP",IF(E{r}="B","","Send invitation")))))')
        # seat key: table-number & position among non-declined guests at that table
        gl[f"T{r}"] = (f'=IF(OR(A{r}="",NOT(ISNUMBER(N{r})),R{r}="Declined"),"",'
                       f'N{r}&"-"&COUNTIFS($N${G0}:N{r},N{r},$R${G0}:R{r},"<>Declined"))')
        gl[f"U{r}"] = f'=IF(AND(R{r}="Attending",NOT(ISNUMBER(N{r}))),COUNTIFS($R${G0}:R{r},"Attending",$N${G0}:N{r},""),"")'
        gl[f"V{r}"] = f'=IF(A{r}="",0,IF(F{r}="Child",N(Settings!$C$9),N(Settings!$C$8)))'
    for c in "TUV": gl.column_dimensions[c].hidden = True
    dvs = [(f"={SIDE}", "C"), (f"={GRP}", "D"), ('"A,B"', "E"), ('"Adult,Child"', "F"), ('"Yes,No"', "I"), ('"Yes,No"', "J"),
           ('"' + ",".join(RSVP) + '"', "K"), (f"={MEAL}", "L"), ('"Yes,No"', "P")]
    for f1, col in dvs:
        dv = DataValidation(type="list", formula1=f1, allow_blank=True); gl.add_data_validation(dv); dv.add(f"{col}{G0}:{col}{G1}")
    dv = DataValidation(type="whole", operator="between", formula1="1", formula2=str(NT), allow_blank=True,
                        error="Use a table number from the Tables tab (1–40).", showErrorMessage=True)
    gl.add_data_validation(dv); dv.add(f"N{G0}:N{G1}")
    gl.conditional_formatting.add(f"R{G0}:R{G1}", FormulaRule(formula=[f'R{G0}="Attending"'], font=Font(name=F, bold=True, color="166534"), fill=fill("DCFCE7")))
    gl.conditional_formatting.add(f"R{G0}:R{G1}", FormulaRule(formula=[f'R{G0}="Declined"'], font=Font(name=F, color="9CA3AF")))
    gl.conditional_formatting.add(f"R{G0}:R{G1}", FormulaRule(formula=[f'R{G0}="Awaiting reply"'], font=Font(name=F, bold=True, color="B45309"), fill=fill("FEF3C7")))
    gl.conditional_formatting.add(f"S{G0}:S{G1}", FormulaRule(formula=[f'S{G0}<>""'], font=Font(name=F, bold=True, color="B91C1C")))
    gl.freeze_panes = f"B{G0}"; gl.auto_filter.ref = f"A5:S{G1}"
    GR = lambda c: f"'Guest List'!${c}${G0}:${c}${G1}"

    # ---------- Tables ----------
    K.title(tb, "Tables", "Your tables and how many seats each has. Seats used counts every guest with that table number who has not declined.")
    K.hdr(tb, 5, ["Table #", "Table name (optional)", "Seats (blank = default)", "Seats used", "Free seats", "Status", "Notes"],
          [8, 22, 11, 9, 9, 16, 30])
    style_range(tb, T0, T1, 1, 7, inputs=(2, 3, 7), fmt={1: "0", 3: "0", 4: INT, 5: '0;-0;"0"'})
    for k in range(NT):
        r = T0 + k
        tb[f"A{r}"] = k + 1; tb[f"A{r}"].font = font(bold=True)
        tb[f"H{r}"] = f'=IF(ISNUMBER(C{r}),C{r},N(Settings!$C$10))'
        tb[f"D{r}"] = f'=COUNTIFS({GR("N")},A{r},{GR("R")},"<>Declined",{GR("A")},"<>")'
        tb[f"E{r}"] = f'=IF(AND(D{r}=0,NOT(ISNUMBER(C{r})),B{r}=""),"",H{r}-D{r})'
        tb[f"F{r}"] = f'=IF(E{r}="","Not used",IF(E{r}<0,"Over capacity",IF(E{r}=0,"Full",E{r}&" seat(s) free")))'
    tb.column_dimensions["H"].hidden = True
    tb.conditional_formatting.add(f"F{T0}:F{T1}", FormulaRule(formula=[f'F{T0}="Over capacity"'], font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    tb.conditional_formatting.add(f"F{T0}:F{T1}", FormulaRule(formula=[f'F{T0}="Full"'], font=Font(name=F, bold=True, color="166534")))
    tb.conditional_formatting.add(f"F{T0}:F{T1}", FormulaRule(formula=[f'F{T0}="Not used"'], font=Font(name=F, color="9CA3AF")))
    tb.freeze_panes = f"A{T0}"
    r = T1 + 1
    tb[f"A{r}"] = "Total"; tb[f"C{r}"] = f'=SUMPRODUCT((F{T0}:F{T1}<>"Not used")*H{T0}:H{T1})'; tb[f"D{r}"] = f"=SUM(D{T0}:D{T1})"
    tb[f"E{r}"] = f"=C{r}-D{r}"
    for c in "ABCDE":
        tb[f"{c}{r}"].font = font(bold=True, color=MAIN); tb[f"{c}{r}"].fill = fill(SOFT); tb[f"{c}{r}"].border = B
    tb[f"B{r}"].fill = fill(SOFT)

    # ---------- Seating Chart ----------
    K.title(sc, "Seating Chart", "Builds itself from the Table # column on the Guest List. Print it for your venue or planner.")
    K.hdr(sc, 5, ["Table", "Name", "Guests"] + [f"Seat {i}" for i in range(1, NS + 1)], [7, 16, 8] + [15] * NS)
    for k in range(NT):
        r = 6 + k; tr = T0 + k
        sc[f"A{r}"] = f"=Tables!A{tr}"
        sc[f"B{r}"] = f'=IF(Tables!B{tr}="","",Tables!B{tr})'
        sc[f"C{r}"] = f"=Tables!D{tr}"
        for i in range(NS):
            c = CL(4 + i)
            sc[f"{c}{r}"] = f'=IFERROR(INDEX({GR("A")},MATCH($A{r}&"-{i + 1}",{GR("T")},0)),"")'
    style_range(sc, 6, 5 + NT, 1, 3 + NS, fmt={1: "0", 3: INT})
    for k in range(NT): sc[f"A{6 + k}"].font = font(bold=True, color=MAIN)
    sc.conditional_formatting.add(f"A6:{CL(3 + NS)}{5 + NT}", FormulaRule(formula=["$C6=0"], font=Font(name=F, color="D1D5DB")))
    sc.freeze_panes = "D6"
    u = 6 + NT + 2
    K.section(sc, f"A{u}", "Attending guests without a table")
    K.hdr(sc, u + 1, ["#", "Guest", "", "Party / household", "", "Side", "Group"], col0=1)
    for k in range(NUNS):
        r = u + 2 + k; m = f"MATCH({k + 1},{GR('U')},0)"
        sc[f"A{r}"] = k + 1
        sc[f"B{r}"] = f'=IFERROR(INDEX({GR("A")},{m}),"")'
        sc[f"D{r}"] = f'=IFERROR(INDEX({GR("B")},{m})&"","")'
        sc[f"F{r}"] = f'=IFERROR(INDEX({GR("C")},{m})&"","")'
        sc[f"G{r}"] = f'=IFERROR(INDEX({GR("D")},{m})&"","")'
    style_range(sc, u + 2, u + 1 + NUNS, 1, 7)
    sc.conditional_formatting.add(f"A{u + 2}:A{u + 1 + NUNS}", FormulaRule(formula=[f'$B{u + 2}=""'], font=Font(name=F, color="D1D5DB")))

    # ---------- sample ----------
    if sample:
        s["C5"] = "Emma & James"; s["C6"] = dt.date(2027, 6, 12); s["C7"] = dt.date(2027, 5, 1); s["C8"] = 85; s["C9"] = 35
        s["F6"] = "Emma's side"; s["F7"] = "James's side"
        rnd = random.Random(5)
        first = ["Olivia", "Liam", "Sophia", "Noah", "Ava", "Ethan", "Isabella", "Mason", "Mia", "Lucas", "Charlotte", "Logan", "Amelia",
                 "Jack", "Harper", "Aiden", "Ella", "Owen", "Grace", "Leo", "Chloe", "Henry", "Lily", "Sam", "Zoe", "Daniel", "Nora",
                 "Ryan", "Hannah", "Max", "Ruby", "Oscar", "Emily", "Theo", "Alice", "Adam", "Sarah", "David", "Laura", "Mark"]
        kids = ["Finn", "Ivy", "Milo", "Rosie", "Leo Jr.", "Poppy", "Arlo", "Daisy"]
        last = ["Smith", "Johnson", "Brown", "Taylor", "Wilson", "Davies", "Evans", "Thomas", "Walker", "Wright", "Robinson", "Thompson",
                "White", "Hughes", "Edwards", "Green", "Hall", "Wood", "Harris", "Martin", "Jackson", "Clarke", "Turner", "Hill",
                "Moore", "Cooper", "Ward", "King", "Baker", "Lewis", "Young", "Allen", "Scott", "Adams", "Mitchell", "Carter"]
        cities = ["12 Oak St, Austin TX 78701", "45 Maple Ave, Denver CO 80203", "8 Pine Rd, Portland OR 97205", "220 Elm St, Chicago IL 60614",
                  "17 Birch Ln, Seattle WA 98101", "3 Cedar Ct, Boston MA 02116", "91 Lake Dr, Madison WI 53703", "64 Hill St, Nashville TN 37203"]
        guests = []; li = 0
        for h in range(62):
            side = "Emma's side" if h % 2 == 0 else "James's side"
            if h % 9 == 0: side = "Both / mutual"
            grp = rnd.choice(["Family", "Family", "Friends", "Friends", "Friends", "Work", "Neighbours"])
            ln_ = last[li % len(last)]; li += 1
            size = rnd.choice([1, 2, 2, 2, 2, 3, 4]) if grp == "Family" else rnd.choice([1, 2, 2])
            lst = "B" if h >= 56 else "A"
            fn = rnd.sample(first, 2)
            people = [(f"{fn[0]} {ln_}", "Adult")]
            party = f"{fn[0]} {ln_}"
            if size >= 2: people.append((f"{fn[1]} {ln_}", "Adult")); party = f"The {ln_}s" if size > 2 else f"{fn[0]} & {fn[1]} {ln_}"
            for k in range(size - 2): people.append((f"{rnd.choice(kids)} {ln_}", "Child"))
            addr = rnd.choice(cities)
            sent = lst == "A"
            if sent:
                x = rnd.random(); rs = "Attending" if x < 0.62 else ("Declined" if x < 0.74 else None)
            else: rs = None
            for i, (nm, ac) in enumerate(people):
                g = dict(A=nm, B=party, C=side, D=grp, E=lst, F=ac, H=addr, I="Yes" if lst == "A" else "No", J="Yes" if sent else "No")
                if i == 0: g["G"] = f"{nm.split()[0].lower()}.{ln_.lower()}@email.com"
                if rs: g["K"] = rs
                if rs == "Attending":
                    g["L"] = "Kids meal" if ac == "Child" else (rnd.choice(MEALS[:5]) if rnd.random() < 0.9 else None)
                    if rnd.random() < 0.12: g["M"] = rnd.choice(["Gluten free", "Nut allergy", "No shellfish", "Lactose free"])
                if rs and rnd.random() < 0.35:
                    g["O"] = rnd.choice(["Cash card", "Registry: stand mixer", "Registry: towels", "Photo frame", "Gift card"])
                    if rnd.random() < 0.5: g["P"] = "Yes"
                guests.append(g)
        # table assignment for attending households: fill tables of 10 in order, keep parties together
        tables = {t: 0 for t in range(1, 13)}; cap = {1: 8}
        cur_party = None; cur_t = 1
        att = [g for g in guests if g.get("K") == "Attending"]
        parties = []
        for g in att:
            if not parties or parties[-1][0] != g["B"]: parties.append([g["B"], []])
            parties[-1][1].append(g)
        for k_, (pn, mem) in enumerate(parties[:-4]):          # leave a few unseated to show the list
            for t in range(1, 13):
                if tables[t] + len(mem) <= cap.get(t, 10):
                    for m in mem: m["N"] = t
                    tables[t] += len(mem); break
        for k, g in enumerate(guests):
            r = G0 + k
            for c, v in g.items():
                if v is not None: gl[f"{c}{r}"] = v
        names = {1: "Head table", 2: "Emma's family", 3: "James's family", 4: "University friends", 5: "Work friends", 6: "Neighbours"}
        for t in range(1, 13):
            if t in names: tb[f"B{T0 + t - 1}"] = names[t]
        tb[f"C{T0}"] = 8

    # ---------- Dashboard ----------
    K.title(dash, "Wedding Guest Dashboard", '=IF(OR(Settings!C5="",Settings!C5="Your names"),"Guest list, RSVPs, meals and seating at a glance.",Settings!C5&" – guest list, RSVPs, meals and seating at a glance.")')
    dash["A3"] = ('=IF(ISNUMBER(Settings!C6),"Wedding day: "&TEXT(Settings!C6,"dddd d mmmm yyyy")&IF(Settings!C6>=TODAY(),"  ·  "&(Settings!C6-TODAY())&" days to go",""),'
                  '"Add your wedding date and RSVP deadline on the Settings tab.")')
    dash["A3"].font = font(bold=True, color=MAIN)
    att = f'{GR("R")},"Attending"'
    kp = [("Guests on the list", f'=COUNTIFS({GR("A")},"<>")', "0"),
          ("Invitations sent (households)", f'=SUMPRODUCT(({GR("B")}<>"")*({GR("J")}="Yes")/(COUNTIFS({GR("B")},{GR("B")}&"",{GR("J")},"Yes")+({GR("B")}="")+({GR("J")}<>"Yes")))', "0"),
          ("Attending", f"=COUNTIFS({att})", "0"),
          ("Declined", f'=COUNTIFS({GR("R")},"Declined")', "0"),
          ("Awaiting reply", f'=COUNTIFS({GR("R")},"Awaiting reply")', "0"),
          ("Replies received", f'=IF(COUNTIFS({GR("J")},"Yes",{GR("A")},"<>")=0,"",(E6+A9)/COUNTIFS({GR("J")},"Yes",{GR("A")},"<>"))', PCT),
          ("Attending without a table", f'=COUNTIFS({att},{GR("N")},"")', "0"),
          ("Estimated catering cost", f"=SUMIFS({GR('V')},{att})", MONEY),
          ('=IF(ISNUMBER(Settings!C7),IF(Settings!C7>=TODAY(),"Days until RSVP deadline","RSVP deadline passed"),"RSVP deadline (Settings)")',
           '=IF(ISNUMBER(Settings!C7),IF(Settings!C7>=TODAY(),Settings!C7-TODAY(),TODAY()-Settings!C7),"")', "0")]
    # KPI grid 3x3; positions: row 5/8/11, cols A/C/E
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 3) * 3, 1 + (i % 3) * 2, lab, f, fmt)
        dash.cell(row=6 + (i // 3) * 3, column=1 + (i % 3) * 2).alignment = Alignment(horizontal="left")
    for c, w in zip("ABCDEFGHI", [26, 9, 26, 10, 26, 10, 10, 12, 10]): dash.column_dimensions[c].width = w
    # By side / group
    H = 16
    K.section(dash, f"A{H - 1}", "Guests by side and group")
    K.hdr(dash, H, ["Side / group", "On list", "Invited", "Attending", "Declined", "Awaiting", "Children", "Thank-yous to send"])
    rows = [(f"Settings!F{S0 + i}", "C") for i in range(6)] + [(f"Settings!G{S0 + i}", "D") for i in range(8)]
    for k, (ref, col) in enumerate(rows):
        r = H + 1 + k; c = GR(col)
        dash[f"A{r}"] = f'=IF({ref}="","",{ref})'
        dash[f"B{r}"] = f'=IF(A{r}="","",COUNTIFS({c},A{r},{GR("A")},"<>"))'
        dash[f"C{r}"] = f'=IF(A{r}="","",COUNTIFS({c},A{r},{GR("J")},"Yes",{GR("A")},"<>"))'
        dash[f"D{r}"] = f'=IF(A{r}="","",COUNTIFS({c},A{r},{att}))'
        dash[f"E{r}"] = f'=IF(A{r}="","",COUNTIFS({c},A{r},{GR("R")},"Declined"))'
        dash[f"F{r}"] = f'=IF(A{r}="","",COUNTIFS({c},A{r},{GR("R")},"Awaiting reply"))'
        dash[f"G{r}"] = f'=IF(A{r}="","",COUNTIFS({c},A{r},{att},{GR("F")},"Child"))'
        dash[f"H{r}"] = f'=IF(A{r}="","",COUNTIFS({c},A{r},{GR("S")},"Send thank-you"))'
    style_range(dash, H + 1, H + len(rows), 1, 8, fmt={i: INT for i in range(2, 9)})
    for k in range(6): dash[f"A{H + 1 + k}"].font = font(bold=True)
    # Meals
    M = H + len(rows) + 3
    K.section(dash, f"A{M - 1}", "Meal count for the caterer (attending guests)")
    K.hdr(dash, M, ["Meal", "Guests", "Dietary notes"])
    for i in range(NL):
        r = M + 1 + i
        dash[f"A{r}"] = f'=IF(Settings!H{S0 + i}="","",Settings!H{S0 + i})'
        dash[f"B{r}"] = f'=IF(A{r}="","",COUNTIFS({GR("L")},A{r},{att}))'
        dash[f"C{r}"] = f'=IF(A{r}="","",COUNTIFS({GR("L")},A{r},{att},{GR("M")},"<>"))'
    r = M + 1 + NL
    dash[f"A{r}"] = "No meal chosen yet"; dash[f"B{r}"] = f'=COUNTIFS({GR("L")},"",{att})'
    style_range(dash, M + 1, r, 1, 3, fmt={2: INT, 3: INT})
    dash.conditional_formatting.add(f"B{r}", FormulaRule(formula=[f"B{r}>0"], font=Font(name=F, bold=True, color="B91C1C")))
    dash.conditional_formatting.add(f"B{M + 1}:B{M + NL}", DataBarRule(start_type="num", start_value=0, end_type="max", color="A78BFA"))
    # charts
    dash["K40"] = "Attending"; dash["K41"] = "Declined"; dash["K42"] = "Awaiting reply"; dash["K43"] = "Not invited yet"
    for i in range(4):
        dash[f"L{40 + i}"] = f'=COUNTIFS({GR("R")},K{40 + i})'
    for rr in range(40, 44):
        for c in "KL": dash[f"{c}{rr}"].font = font(color="FFFFFF", size=8)
    ch = BarChart(); ch.type = "col"; ch.title = "RSVP status"; ch.height = 7.5; ch.width = 14; ch.legend = None
    ch.add_data(Reference(dash, min_col=12, min_row=40, max_row=43), titles_from_data=False)
    ch.set_categories(Reference(dash, min_col=11, min_row=40, max_row=43)); ch.y_axis.majorGridlines = None; ch.gapWidth = 50
    ch.series[0].graphicalProperties.solidFill = MAIN
    ch.roundedCorners = False; ch.x_axis.delete = False; ch.y_axis.delete = False
    dash.add_chart(ch, "K3")
    ch2 = BarChart(); ch2.type = "bar"; ch2.title = "Meals"; ch2.height = 7.5; ch2.width = 14; ch2.legend = None
    ch2.add_data(Reference(dash, min_col=2, min_row=M + 1, max_row=M + 7), titles_from_data=False)
    ch2.set_categories(Reference(dash, min_col=1, min_row=M + 1, max_row=M + 7)); ch2.y_axis.majorGridlines = None
    ch2.x_axis.scaling.orientation = "maxMin"; ch2.series[0].graphicalProperties.solidFill = "A78BFA"; ch2.gapWidth = 50
    ch2.roundedCorners = False; ch2.x_axis.delete = False; ch2.y_axis.delete = False
    dash.add_chart(ch2, "K19")

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Settings tab: your names, wedding date, RSVP deadline, catering cost per guest, seats per table – and rename the sides, groups and meals.", False),
             ("2.  Guest List tab: one row per person, including children. Give everyone in a household the same 'Party' name (e.g. 'The Smiths').", False),
             ("     A list = invite now, B list = invite if A-list guests decline.", False),
             ("3.  As you go: mark Save-the-date sent and Invitation sent, then fill in RSVP and meal choice when replies come in.", False),
             ("     The Status and To do columns tell you what is left for every guest (send invitation, follow up, ask meal, assign table, send thank-you).", False),
             ("4.  Tables tab: name your tables and set the number of seats (blank = the default from Settings).", False),
             ("5.  Guest List: type a table number for each attending guest. The Seating Chart tab fills itself and lists who still needs a seat.", False),
             ("6.  Dashboard: head count, replies, meal count for the caterer, guests by side and group, estimated catering cost.", False),
             ("", False), ("Tips", True),
             ("•  Sort or filter the Guest List by Party to keep families together when you seat them.", False),
             ("•  Tables tab turns red when a table is over capacity.", False),
             ("•  Log gifts in the 'Gift received' column – the To do column reminds you to send a thank-you note.", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  500 guests, 40 tables, up to 14 seats per table on the chart. No macros. Works in Excel 2010+ and Google Sheets.", False),
             ("•  To use Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("All names, addresses and replies are made up. Use the blank copy for your own wedding.", False)]
    K.start(st, "Wedding Guest List, RSVP & Seating Planner", "Guests, replies, meals and tables in one place – Excel and Google Sheets.", lines)
    finish(wb, sample, path)


META = {
    "name": "Wedding Guest List, RSVP & Seating Chart Planner (Meal Count, Table Planner, Thank-You Tracker) – Excel & Google Sheets",
    "file_base": "Wedding-Guest-List-RSVP-Seating-Planner",
    "price": 14,
    "color": "#6D28D9",
    "title": "Wedding Guest List & Seating",
    "tagline": "RSVPs · meal count · seating chart that builds itself · thank-yous",
    "intro": "Keep every guest, reply, meal choice and seat in one calm place. Add your guests by household, track save-the-dates, invitations and RSVPs, get the meal count your caterer asks for, type a table number and watch the seating chart build itself – with a to-do for every guest so nothing slips.",
    "features": [
        "Guest list for up to 500 people with households (party), side, group, A/B list, adult/child, contact and mailing address",
        "Save-the-date, invitation and RSVP tracking with automatic status: Attending, Declined, Awaiting reply, Not invited yet",
        "To-do column per guest: send invitation, follow up RSVP, ask meal choice, assign a table, send thank-you",
        "Meal count for the caterer with dietary notes, plus estimated catering cost (adult and child prices)",
        "40 tables with seat limits and an over-capacity warning",
        "Seating chart that fills itself from table numbers (up to 14 seats per table) and a list of attending guests without a seat",
        "Dashboard: head count, households invited, reply rate, days to the RSVP deadline, guests by side and group, charts",
        "Gift and thank-you note tracker",
        "Blank copy plus a sample copy with a 118-guest example wedding",
    ],
    "tags": ["wedding guest list", "wedding planner", "seating chart", "rsvp tracker", "wedding spreadsheet", "google sheets",
             "excel template", "wedding seating", "guest tracker", "meal count", "wedding budget", "bride to be", "wedding organizer"],
    "summary": "Instant download: Excel + Google Sheets template, blank copy + sample copy.",
}
