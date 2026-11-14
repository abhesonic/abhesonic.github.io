"""Photography Business Tracker – clients & bookings, packages, deposits & balances, gallery delivery deadlines,
mileage, expenses, monthly profit, tax set-aside and a dashboard.

Settings (business, packages, shoot types, lead sources, expense categories) -> Bookings (one row per shoot)
-> Expenses -> Monthly -> Dashboard.
"""
import datetime as dt, random
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import BarChart, Reference
from _common import Kit, F, B, MONEY, DATE, font, fill, style_range, input_cell, finish

MAIN = "7C3AED"; SOFT = "EDE9FE"
K = Kit(MAIN, SOFT)
PCT = "0.0%"; HRS = '0.0;-0.0;"-"'
NP = 20; P0 = 6; P1 = P0 + NP - 1            # packages Settings A6:F25
L0 = 30; L1 = 41                              # lists rows 30..41 (shoot types A, lead sources C, expense cats E)
NB = 600; B0 = 6; B1 = B0 + NB - 1            # bookings
NX = 400; X0 = 6; X1 = X0 + NX - 1            # expenses
STATUS = ["Inquiry", "Booked", "Shot", "Editing", "Delivered", "Cancelled", "Lost"]
TYPES = ["Wedding", "Engagement", "Family", "Newborn", "Portrait", "Headshots", "Event", "Brand / commercial", "Real estate", "Mini session"]
SOURCES = ["Instagram", "Google search", "Referral", "Past client", "Website", "Wedding directory", "Facebook", "Other"]
EXP_CATS = ["Camera gear", "Lenses & lighting", "Software & editing", "Website & gallery hosting", "Insurance", "Marketing & ads",
            "Props & backdrops", "Prints & albums (cost)", "Education", "Studio rent", "Other"]
PACKAGES = [("Wedding – Full day", "Wedding", 3200, 8, 0.30, 42), ("Wedding – Half day", "Wedding", 2100, 5, 0.30, 35),
            ("Elopement", "Wedding", 1200, 3, 0.30, 28), ("Engagement session", "Engagement", 350, 1.5, 0.5, 14),
            ("Family session", "Family", 425, 1, 0.5, 14), ("Newborn studio", "Newborn", 550, 3, 0.5, 21),
            ("Portrait – 1 hour", "Portrait", 275, 1, 0.5, 10), ("Headshots – mini", "Headshots", 180, 0.5, 1, 5),
            ("Event – per 3 hours", "Event", 750, 3, 0.5, 14), ("Brand session", "Brand / commercial", 950, 3, 0.5, 14),
            ("Real estate listing", "Real estate", 225, 1.5, 0, 2), ("Holiday mini session", "Mini session", 175, 0.33, 1, 7)]
CLIENTS = ["Emma & Jake Collins", "The Rivera Family", "Sophie Turner", "Marcus Lee", "Hannah & Owen Park", "Bright Bakery Co.",
           "The Nguyen Family", "Olivia Brooks", "Daniel Ross", "Chloe & Sam Patel", "Harbor Realty", "Mia Thompson", "The Garcia Family",
           "Liam Walker", "Ava & Noah Kim", "Summit Tech Inc.", "Grace Miller", "The Johnson Family", "Isla Bennett", "Ethan Wright",
           "Ruby & Max Allen", "Lakeside Yoga", "Zoe Carter", "The Hughes Family", "Leo Martin", "Nora & Ben Scott", "Coastal Homes",
           "Lily Evans", "Jack Foster", "The Murphy Family", "Ella Price", "Henry Cole", "Aria & Luke Ward", "The Diaz Family", "Maple Cafe"]


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); s = wb.create_sheet("Settings"); bk = wb.create_sheet("Bookings")
    ex = wb.create_sheet("Expenses"); mo = wb.create_sheet("Monthly")

    # ---------- Settings ----------
    K.title(s, "Settings", "Yellow cells are yours to edit.")
    opts = [("Business name", "Golden Hour Photography" if sample else "My Photography Business", None),
            ("Year for Dashboard / Monthly", 2026 if sample else "=YEAR(TODAY())", "0"),
            ("Currency label", "USD", None),
            ("Default gallery turnaround (days)", 21, "0"),
            ("Mileage rate per mile / km", 0.70, MONEY),
            ("Set aside for tax (% of profit)", 0.25, PCT),
            ("Alert: balance due within (days of shoot)", 7, "0")]
    for i, (lab, v, fmt) in enumerate(opts, 3):
        s[f"I{i}"] = lab; s[f"I{i}"].font = font(bold=True); input_cell(s[f"J{i}"], v)
        if fmt: s[f"J{i}"].number_format = fmt
    K.hdr(s, 5, ["Package", "Shoot type", "Price", "Hours", "Deposit %", "Turnaround days"])
    for i in range(NP):
        r = P0 + i; p = PACKAGES[i] if i < len(PACKAGES) else (None,) * 6
        for col, v in zip("ABCDEF", p): input_cell(s[f"{col}{r}"], v)
        s[f"C{r}"].number_format = MONEY; s[f"D{r}"].number_format = HRS; s[f"E{r}"].number_format = "0%"; s[f"F{r}"].number_format = "0"
    K.hdr(s, 29, ["Shoot types"]); K.hdr(s, 29, ["Lead sources"], col0=3); K.hdr(s, 29, ["Expense categories"], col0=5)
    for i in range(L1 - L0 + 1):
        r = L0 + i
        input_cell(s[f"A{r}"], TYPES[i] if i < len(TYPES) else None)
        input_cell(s[f"C{r}"], SOURCES[i] if i < len(SOURCES) else None)
        input_cell(s[f"E{r}"], EXP_CATS[i] if i < len(EXP_CATS) else None)
    for c, w in zip("ABCDEFGHIJ", [24, 18, 18, 8, 22, 14, 3, 3, 40, 26]): s.column_dimensions[c].width = w
    PK = f"Settings!$A${P0}:$A${P1}"; TY = f"Settings!$A${L0}:$A${L1}"; SO = f"Settings!$C${L0}:$C${L1}"
    YR = "Settings!$J$4"
    PKC = lambda c: f"Settings!${c}${P0}:${c}${P1}"

    # ---------- Bookings ----------
    K.title(bk, "Bookings", "One row per shoot. Yellow = you type, grey = calculated. Leave 'Price' empty to use the package price.")
    hd = ["Client", "Phone / email", "Shoot type", "Package", "Shoot date", "Location", "Status", "Lead source", "Price (optional)",
          "Extras (prints, hours…)", "Discount", "Deposit paid", "Other payments", "Shoot costs (2nd shooter, rental…)", "Miles / km (round trip)",
          "Gallery delivered on", "Notes",
          "Package price", "Booking total", "Paid so far", "Balance due", "Mileage cost", "Profit", "Gallery due", "Alert", "Month", "Year", "Hours", "rank"]
    K.hdr(bk, 5, hd, [22, 20, 15, 22, 12, 18, 11, 15, 11, 11, 10, 11, 11, 13, 10, 12, 22, 11, 12, 11, 11, 10, 11, 12, 22, 7, 7, 7, 6])
    style_range(bk, B0, B1, 1, 29, inputs=tuple(range(1, 18)),
                fmt={5: DATE, 9: MONEY, 10: MONEY, 11: MONEY, 12: MONEY, 13: MONEY, 14: MONEY, 15: "#,##0", 16: DATE,
                     18: MONEY, 19: MONEY, 20: MONEY, 21: MONEY, 22: MONEY, 23: MONEY, 24: DATE, 26: "0", 27: "0", 28: HRS, 29: "0.00000"})
    for r in range(B0, B1 + 1):
        bk[f"R{r}"] = f'=IF(A{r}="",0,IF(I{r}<>"",I{r},IFERROR(INDEX({PKC("C")},MATCH(D{r},{PK},0)),0)))'
        bk[f"S{r}"] = (f'=IF(A{r}="",0,IF(OR(G{r}="Inquiry",G{r}="Lost",G{r}=""),0,IF(G{r}="Cancelled",T{r},R{r}+N(J{r})-N(K{r}))))')
        bk[f"T{r}"] = f'=N(L{r})+N(M{r})'
        bk[f"U{r}"] = f'=IF(A{r}="",0,MAX(0,S{r}-T{r}))'
        bk[f"V{r}"] = f'=IF(OR(S{r}=0,G{r}="Cancelled"),0,N(O{r})*Settings!$J$7)'
        bk[f"W{r}"] = f'=IF(S{r}=0,0,S{r}-N(N{r})-V{r})'
        bk[f"X{r}"] = (f'=IF(OR(A{r}="",E{r}=""),"",E{r}+IF(N(IFERROR(INDEX({PKC("F")},MATCH(D{r},{PK},0)),0))>0,'
                       f'INDEX({PKC("F")},MATCH(D{r},{PK},0)),Settings!$J$6))')
        bk[f"Y{r}"] = (f'=IF(OR(A{r}="",G{r}="Cancelled",G{r}="Lost"),"",'
                       f'IF(G{r}="Inquiry","Open inquiry – follow up",'
                       f'IF(AND(G{r}="Booked",N(L{r})=0),"No deposit yet",'
                       f'IF(AND(E{r}<>"",E{r}>=TODAY(),E{r}-TODAY()<=Settings!$J$9,U{r}>0),"Balance due before shoot",'
                       f'IF(OR(G{r}="Delivered",Q{r}<>"",X{r}=""),IF(U{r}>0,"Balance unpaid",""),'
                       f'IF(E{r}>TODAY(),"",IF(X{r}<TODAY(),"Gallery overdue",IF(X{r}-TODAY()<=5,"Gallery due soon",""))))))))')
        bk[f"Z{r}"] = f'=IF(E{r}="","",MONTH(E{r}))'
        bk[f"AA{r}"] = f'=IF(E{r}="","",YEAR(E{r}))'
        bk[f"AB{r}"] = f'=IF(A{r}="",0,IFERROR(INDEX({PKC("D")},MATCH(D{r},{PK},0)),0))'
        bk[f"AC{r}"] = (f'=IF(AND(A{r}<>"",E{r}<>"",OR(G{r}="Booked",G{r}="Inquiry")),IF(E{r}>=TODAY(),E{r}+ROW()/100000,""),"")')
    for rng, f1 in ((f"C{B0}:C{B1}", f"={TY}"), (f"D{B0}:D{B1}", f"={PK}"), (f"G{B0}:G{B1}", '"' + ",".join(STATUS) + '"'),
                    (f"H{B0}:H{B1}", f"={SO}")):
        dv = DataValidation(type="list", formula1=f1, allow_blank=True); bk.add_data_validation(dv); dv.add(rng)
    bk.conditional_formatting.add(f"Y{B0}:Y{B1}", FormulaRule(formula=[f'OR(Y{B0}="Gallery overdue",Y{B0}="Balance due before shoot")'],
                                                             font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    bk.conditional_formatting.add(f"Y{B0}:Y{B1}", FormulaRule(formula=[f'AND(Y{B0}<>"",Y{B0}<>"Gallery overdue",Y{B0}<>"Balance due before shoot")'],
                                                             font=Font(name=F, bold=True, color="B45309"), fill=fill("FEF3C7")))
    bk.conditional_formatting.add(f"G{B0}:G{B1}", FormulaRule(formula=[f'OR(G{B0}="Cancelled",G{B0}="Lost")'], font=Font(name=F, color="9CA3AF")))
    bk.conditional_formatting.add(f"G{B0}:G{B1}", FormulaRule(formula=[f'G{B0}="Delivered"'], font=Font(name=F, bold=True, color="15803D")))
    bk.conditional_formatting.add(f"U{B0}:U{B1}", FormulaRule(formula=[f'U{B0}>0'], font=Font(name=F, bold=True, color="B91C1C")))
    bk.column_dimensions["AC"].hidden = True
    bk.freeze_panes = f"B{B0}"
    BK = lambda c: f"Bookings!${c}${B0}:${c}${B1}"

    # ---------- Expenses ----------
    K.title(ex, "Expenses", "Business costs not tied to one shoot: gear, software, insurance, ads, education…")
    K.hdr(ex, 5, ["Date", "Category", "Description", "Amount", "Month", "Year"], [12, 24, 34, 12, 8, 8])
    style_range(ex, X0, X1, 1, 6, inputs=(1, 2, 3, 4), fmt={1: DATE, 4: MONEY, 5: "0", 6: "0"})
    for r in range(X0, X1 + 1):
        ex[f"E{r}"] = f'=IF(A{r}="","",MONTH(A{r}))'; ex[f"F{r}"] = f'=IF(A{r}="","",YEAR(A{r}))'
    dv = DataValidation(type="list", formula1=f"=Settings!$E${L0}:$E${L1}", allow_blank=True); ex.add_data_validation(dv); dv.add(f"B{X0}:B{X1}")
    ex.freeze_panes = f"A{X0}"
    XP = lambda c: f"Expenses!${c}${X0}:${c}${X1}"

    # ---------- sample data ----------
    if sample:
        rnd = random.Random(11)
        today = dt.date(2026, 11, 14)
        pk = {p[0]: p for p in PACKAGES}
        plan = []
        # spread shoots over the year (weddings in season), plus upcoming ones
        wed = ["Wedding – Full day", "Wedding – Half day", "Elopement"]
        other = ["Engagement session", "Family session", "Newborn studio", "Portrait – 1 hour", "Headshots – mini", "Event – per 3 hours",
                 "Brand session", "Real estate listing"]
        d = dt.date(2026, 1, 10); k = 0
        while d <= dt.date(2027, 1, 20):
            m = d.month
            if 5 <= m <= 10 and k % 3 == 0: pname = wed[k % 3 if k % 2 else 0]
            elif m == 11 and d.day < 25 and k % 2 == 0: pname = "Holiday mini session"
            else: pname = other[(k * 5 + m) % len(other)]
            plan.append((d, pname)); k += 1
            d += dt.timedelta(days=rnd.choice([4, 5, 6, 7, 8]))
        rows = []
        for i, (d, pname) in enumerate(plan):
            p = pk[pname]; client = CLIENTS[i % len(CLIENTS)] if p[1] != "Mini session" else ["Wood family", "Reed family", "Hall family", "Young family", "King family"][i % 5]
            if d < today - dt.timedelta(days=30): stt = "Delivered"
            elif d < today - dt.timedelta(days=4): stt = "Editing"
            elif d < today: stt = "Shot"
            else: stt = "Booked"
            if rnd.random() < 0.04 and d < today: stt = "Cancelled"
            dep = round(p[2] * p[4], 0)
            if stt == "Booked" and i % 7 == 0: dep = 0
            extras = rnd.choice([0, 0, 0, 0, 150, 250]) if p[1] in ("Wedding", "Family", "Newborn") else 0
            disc = 50 if i % 11 == 0 else 0
            total = p[2] + extras - disc
            other_pay = 0 if stt in ("Booked", "Cancelled") else (total - dep if (stt == "Delivered" or i % 4) else 0)
            if stt == "Delivered" and i % 17 == 0: other_pay = max(0, total - dep - 200)
            costs = 400 if pname == "Wedding – Full day" else (60 if p[1] == "Newborn" else (rnd.choice([0, 0, 25]) ))
            miles = rnd.choice([12, 18, 24, 36, 48, 60]) if p[1] not in ("Newborn", "Headshots", "Mini session") else 0
            deliv = d + dt.timedelta(days=p[5] - rnd.choice([1, 2, 3, 4])) if stt == "Delivered" else None
            rows.append([client, f"(555) 02{10 + i % 80}-{4000 + 13 * i}", p[1], pname, d,
                         ["Riverside Park", "Studio", "Client home", "Old Mill Venue", "Downtown", "Beach at Point Lane", "Client office"][i % 7],
                         stt, SOURCES[(i * 3) % len(SOURCES)], None, extras or None, disc or None, dep or None, other_pay or None, costs or None,
                         miles or None, deliv, None])
        # a gallery overdue, an inquiry or two, a lost lead
        for rw in rows:
            if rw[6] == "Editing": rw[4] = rw[4] - dt.timedelta(days=0); break
        rows.append(["Penny & Tom Grey", "penny@mail.example", "Wedding", "Wedding – Full day", dt.date(2027, 6, 12), "Old Mill Venue", "Inquiry",
                     "Wedding directory", None, None, None, None, None, None, None, None, "Asked for album prices"])
        rows.append(["Northside Dental", "office@mail.example", "Headshots", "Headshots – mini", dt.date(2026, 12, 3), "Client office", "Inquiry",
                     "Google search", 1200, None, None, None, None, None, 20, None, "Team of 8 – quoted 1,200"])
        rows.append(["Kate Owens", "kate@mail.example", "Family", "Family session", dt.date(2026, 10, 2), "Riverside Park", "Lost", "Instagram",
                     None, None, None, None, None, None, None, None, "Went with cheaper option"])
        for j, (nm, pn, dd, src) in enumerate([("Amy Fisher", "Newborn studio", dt.date(2026, 4, 9), "Facebook"),
                                                ("Josh & Kim Hale", "Wedding – Half day", dt.date(2026, 8, 22), "Wedding directory"),
                                                ("Rachel Stone", "Portrait – 1 hour", dt.date(2026, 6, 14), "Instagram"),
                                                ("Peak Fitness", "Brand session", dt.date(2026, 7, 7), "Google search"),
                                                ("Tara Quinn", "Family session", dt.date(2026, 9, 19), "Instagram")]):
            rows.append([nm, f"{nm.split()[0].lower()}@mail.example", pk[pn][1], pn, dd, None, "Lost", src] + [None] * 8 + ["No reply after quote"])
        rows.append(["Bella & Finn Lowe", "bella@mail.example", "Wedding", "Wedding – Full day", dt.date(2026, 9, 26), "Old Mill Venue", "Editing",
                     "Referral", None, 300, None, 960, 2540, 400, 54, None, "Album upgrade"])
        rows.sort(key=lambda r: r[4])
        assert len(rows) <= NB
        for r, rw in enumerate(rows, B0):
            for c, v in enumerate(rw, 1):
                if v is not None: bk.cell(row=r, column=c, value=v)
        xs = []
        for m in range(1, 12):
            xs += [(dt.date(2026, m, 2), "Software & editing", "Lightroom + Photoshop plan", 20),
                   (dt.date(2026, m, 5), "Website & gallery hosting", "Website + online galleries", 38),
                   (dt.date(2026, m, 12), "Marketing & ads", "Instagram ads", 60 + 20 * (m % 3))]
            if m in (1, 7): xs.append((dt.date(2026, m, 18), "Insurance", "Liability + gear insurance (6 months)", 310))
        xs += [(dt.date(2026, 3, 9), "Lenses & lighting", "85mm f/1.8 lens", 549), (dt.date(2026, 4, 21), "Camera gear", "Backup camera body", 1199),
               (dt.date(2026, 2, 14), "Education", "Posing workshop", 199), (dt.date(2026, 10, 20), "Props & backdrops", "Holiday mini set", 145),
               (dt.date(2026, 8, 4), "Prints & albums (cost)", "Sample albums", 260)]
        xs.sort()
        for r, x in enumerate(xs, X0):
            for col, v in zip("ABCD", x): ex[f"{col}{r}"] = v

    # ---------- Monthly ----------
    K.title(mo, "Monthly Summary", "By shoot date, for the year on the Settings tab. Cancelled shoots count only the money you kept.")
    mo["A3"] = '="Year "&Settings!J4&"  ·  amounts in "&Settings!J5'; mo["A3"].font = font(bold=True)
    K.hdr(mo, 5, ["Month", "Shoots", "Revenue", "Shoot costs", "Mileage", "Other expenses", "Profit", "Tax set-aside", "Take-home", "Balance unpaid"],
          [12, 9, 13, 12, 11, 14, 13, 13, 13, 13])
    for m in range(1, 13):
        r = 5 + m
        cnd = f'{BK("Z")},{m},{BK("AA")},{YR}'
        mo[f"A{r}"] = f'=TEXT(DATE({YR},{m},1),"mmm yyyy")'
        mo[f"B{r}"] = (f'=COUNTIFS({cnd},{BK("G")},"Booked")+COUNTIFS({cnd},{BK("G")},"Shot")+COUNTIFS({cnd},{BK("G")},"Editing")'
                       f'+COUNTIFS({cnd},{BK("G")},"Delivered")')
        mo[f"C{r}"] = f"=SUMIFS({BK('S')},{cnd})"
        mo[f"D{r}"] = f'=SUMIFS({BK("N")},{cnd},{BK("S")},">0")'
        mo[f"E{r}"] = f"=SUMIFS({BK('V')},{cnd})"
        mo[f"F{r}"] = f"=SUMIFS({XP('D')},{XP('E')},{m},{XP('F')},{YR})"
        mo[f"G{r}"] = f"=C{r}-D{r}-E{r}-F{r}"
        mo[f"H{r}"] = f"=MAX(0,G{r})*Settings!$J$8"
        mo[f"I{r}"] = f"=G{r}-H{r}"
        mo[f"J{r}"] = f"=SUMIFS({BK('U')},{cnd})"
    style_range(mo, 6, 17, 1, 10, fmt={2: "0", 3: MONEY, 4: MONEY, 5: MONEY, 6: MONEY, 7: MONEY, 8: MONEY, 9: MONEY, 10: MONEY})
    mo["A18"] = "TOTAL"
    for c in "BCDEFGHIJ": mo[f"{c}18"] = f"=SUM({c}6:{c}17)"
    for c, fmt in zip("ABCDEFGHIJ", [None, "0"] + [MONEY] * 8):
        cell = mo[f"{c}18"]; cell.font = Font(name=F, bold=True, color="FFFFFF"); cell.fill = fill(MAIN); cell.border = B
        if fmt: cell.number_format = fmt
    mo.conditional_formatting.add("G6:G17", FormulaRule(formula=["G6<0"], font=Font(name=F, bold=True, color="B91C1C")))

    # ---------- Dashboard ----------
    K.title(dash, "Photography Business Dashboard", "Updates by itself from the Bookings and Expenses tabs.")
    dash["A3"] = '=Settings!J3&"  ·  year "&Settings!J4&"  ·  amounts in "&Settings!J5'; dash["A3"].font = font(bold=True)
    won = f'(COUNTIF({BK("G")},"Booked")+COUNTIF({BK("G")},"Shot")+COUNTIF({BK("G")},"Editing")+COUNTIF({BK("G")},"Delivered")+COUNTIF({BK("G")},"Cancelled"))'
    kp = [("Revenue this year", "=Monthly!C18", MONEY), ("Profit this year", "=Monthly!G18", MONEY), ("Shoots this year", "=Monthly!B18", "0"),
          ("Average booking", "=IFERROR(Monthly!C18/Monthly!B18,0)", MONEY), ("Balance still to collect", f"=SUM({BK('U')})", MONEY),
          ("Inquiry → booking rate", f'=IFERROR({won}/({won}+COUNTIF({BK("G")},"Lost")),0)', "0%")]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 3) * 3, 1 + (i % 3) * 2, lab, f, fmt)
        dash.cell(row=6 + (i // 3) * 3, column=1 + (i % 3) * 2).alignment = Alignment(horizontal="left")
    for c, w in zip("ABCDEFG", [32, 14, 24, 14, 26, 14, 3]): dash.column_dimensions[c].width = w
    K.section(dash, "A11", "To do")
    todo = [("Shoots in the next 14 days", f'=COUNTIFS({BK("G")},"Booked",{BK("E")},">="&TODAY(),{BK("E")},"<"&TODAY()+14)', "0"),
            ("Galleries overdue", f'=COUNTIF({BK("Y")},"Gallery overdue")', "0"),
            ("Galleries due within 5 days", f'=COUNTIF({BK("Y")},"Gallery due soon")', "0"),
            ("Balances due before a shoot", f'=COUNTIF({BK("Y")},"Balance due before shoot")', "0"),
            ("Booked without a deposit", f'=COUNTIF({BK("Y")},"No deposit yet")', "0"),
            ("Open inquiries to follow up", f'=COUNTIF({BK("G")},"Inquiry")', "0"),
            ("Tax to set aside this year", "=Monthly!H18", MONEY),
            ("Revenue per shooting hour", f'=IFERROR(Monthly!C18/SUMPRODUCT(({BK("AA")}={YR})*({BK("S")}>0)*({BK("G")}<>"Cancelled")*{BK("AB")}),0)', MONEY)]
    for i, (lab, f, fmt) in enumerate(todo, 12):
        dash[f"A{i}"] = lab; dash[f"A{i}"].font = font(bold=True); dash[f"A{i}"].border = B; dash[f"A{i}"].fill = fill(SOFT)
        c = dash[f"B{i}"]; c.value = f; c.number_format = fmt; c.border = B; c.font = Font(name=F, bold=True, color=MAIN)
    dash.conditional_formatting.add("B13", FormulaRule(formula=["B13>0"], font=Font(name=F, bold=True, color="B91C1C")))
    UR = 21
    K.section(dash, f"A{UR}", "Next 10 shoots")
    K.hdr(dash, UR + 1, ["Client", "Date", "Package", "Balance due", "Alert"])
    RK = BK("AC")
    for k in range(10):
        r = UR + 2 + k
        dash[f"F{r}"] = f'=IFERROR(MATCH(SMALL({RK},{k + 1}),{RK},0),"")'
        dash[f"A{r}"] = f'=IF(F{r}="","",INDEX({BK("A")},F{r}))'
        dash[f"B{r}"] = f'=IF(F{r}="","",INDEX({BK("E")},F{r}))'
        dash[f"C{r}"] = f'=IF(F{r}="","",INDEX({BK("D")},F{r})&"")'
        dash[f"D{r}"] = f'=IF(F{r}="","",INDEX({BK("U")},F{r}))'
        dash[f"E{r}"] = f'=IF(F{r}="","",INDEX({BK("Y")},F{r})&"")'
        dash[f"F{r}"].font = Font(name=F, color="FFFFFF", size=8)
    style_range(dash, UR + 2, UR + 11, 1, 5, fmt={2: DATE, 4: MONEY})
    TR = UR + 13
    K.section(dash, f"A{TR}", "Revenue by shoot type (this year)")
    K.hdr(dash, TR + 1, ["Shoot type", "Shoots", "Revenue", "Avg booking"])
    NT = L1 - L0 + 1
    for i in range(NT):
        r = TR + 2 + i; k = L0 + i
        dash[f"A{r}"] = f'=IF(Settings!A{k}="","",Settings!A{k})'
        dash[f"B{r}"] = f'=IF(A{r}="","",COUNTIFS({BK("C")},A{r},{BK("AA")},{YR},{BK("S")},">0"))'
        dash[f"C{r}"] = f'=IF(A{r}="","",SUMIFS({BK("S")},{BK("C")},A{r},{BK("AA")},{YR}))'
        dash[f"D{r}"] = f'=IF(OR(A{r}="",N(B{r})=0),"",C{r}/B{r})'
    style_range(dash, TR + 2, TR + 1 + NT, 1, 4, fmt={2: "0", 3: MONEY, 4: MONEY})
    SR = TR + 3 + NT
    K.section(dash, f"A{SR}", "Where clients come from (all time)")
    K.hdr(dash, SR + 1, ["Lead source", "Inquiries", "Booked", "Revenue"])
    for i in range(NT):
        r = SR + 2 + i; k = L0 + i
        dash[f"A{r}"] = f'=IF(Settings!C{k}="","",Settings!C{k})'
        dash[f"B{r}"] = f'=IF(A{r}="","",COUNTIF({BK("H")},A{r}))'
        dash[f"C{r}"] = f'=IF(A{r}="","",COUNTIFS({BK("H")},A{r},{BK("S")},">0"))'
        dash[f"D{r}"] = f'=IF(A{r}="","",SUMIFS({BK("S")},{BK("H")},A{r}))'
    style_range(dash, SR + 2, SR + 1 + NT, 1, 4, fmt={2: "0", 3: "0", 4: MONEY})
    ch = BarChart(); ch.type = "col"; ch.title = "Revenue and profit by month"; ch.height = 8; ch.width = 17
    ch.add_data(Reference(mo, min_col=3, min_row=5, max_row=17), titles_from_data=True)
    ch.add_data(Reference(mo, min_col=7, min_row=5, max_row=17), titles_from_data=True)
    ch.set_categories(Reference(mo, min_col=1, min_row=6, max_row=17)); ch.y_axis.majorGridlines = None
    ch.series[0].graphicalProperties.solidFill = MAIN; ch.series[1].graphicalProperties.solidFill = "C4B5FD"
    ch.roundedCorners = False; ch.legend.overlay = False; ch.x_axis.delete = False; ch.y_axis.delete = False
    dash.add_chart(ch, "H5")
    ch2 = BarChart(); ch2.type = "bar"; ch2.title = "Revenue by shoot type"; ch2.height = 8; ch2.width = 17
    ch2.add_data(Reference(dash, min_col=3, min_row=TR + 1, max_row=TR + 1 + len(TYPES)), titles_from_data=True)
    ch2.set_categories(Reference(dash, min_col=1, min_row=TR + 2, max_row=TR + 1 + len(TYPES))); ch2.legend = None
    ch2.x_axis.scaling.orientation = "maxMin"; ch2.y_axis.majorGridlines = None; ch2.series[0].graphicalProperties.solidFill = MAIN
    ch2.roundedCorners = False; ch2.x_axis.delete = False; ch2.y_axis.delete = False
    dash.add_chart(ch2, "H22")

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Settings tab: business name, year, your packages (price, hours, deposit %, gallery turnaround), shoot types, lead sources and expense categories.", False),
             ("2.  Bookings tab: one row per shoot – client, shoot type, package, date, location and status (Inquiry, Booked, Shot, Editing, Delivered, Cancelled, Lost).", False),
             ("3.  Type the deposit and other payments as they come in. The balance due, mileage cost and profit per shoot calculate by themselves.", False),
             ("4.  'Gallery due' = shoot date + turnaround days. Type the date in 'Gallery delivered on' and set the status to Delivered when you send it.", False),
             ("5.  The Alert column warns you: Gallery overdue, Gallery due soon, Balance due before shoot, No deposit yet, Open inquiry – follow up.", False),
             ("6.  Expenses tab: gear, software, insurance, ads. Monthly and Dashboard then show profit, tax to set aside (a savings guide, not tax advice), next 10 shoots and best lead sources.", False),
             ("", False), ("Tips", True),
             ("•  Leave 'Price' empty to use the package price; type a price for custom quotes. Use 'Extras' for prints, albums or extra hours.", False),
             ("•  Log inquiries too (status Inquiry). Change them to Booked or Lost – the Dashboard shows your booking rate and which lead sources pay.", False),
             ("•  If a client cancels, set Cancelled: only the money you already received (e.g. a non-refundable retainer) counts as revenue.", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  600 bookings, 400 expenses, 20 packages. Works in miles or km. No macros. Works in Excel 2010+ and Google Sheets.", False),
             ("•  To use Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("The business, clients, shoots and prices are made-up examples. Use the blank copy in your download for your own business.", False)]
    K.start(st, "Photography Business Tracker", "Clients, bookings, deposits, gallery deadlines, mileage, expenses and profit – Excel and Google Sheets.", lines)
    finish(wb, sample, path)


META = {
    "name": "Photography Business Tracker (Clients, Bookings, Deposits, Gallery Deadlines, Income & Expenses) – Excel & Google Sheets",
    "file_base": "Photography-Business-Tracker",
    "price": 19,
    "color": "#7C3AED",
    "title": "Photography Business Tracker",
    "tagline": "bookings · deposits · gallery deadlines · mileage · expenses · profit",
    "intro": "Run your photography business from one simple spreadsheet: log every inquiry and booking, see deposits and balances still to collect, never miss a gallery delivery deadline, and know your real profit after shoot costs, mileage and expenses. Works for wedding, family, newborn, portrait, brand and real-estate photographers.",
    "features": [
        "Bookings log for 600 shoots: client, shoot type, package, date, location, status from Inquiry to Delivered",
        "Package list with price, hours, deposit % and gallery turnaround – the price fills in by itself",
        "Deposits, payments and balance due per client, plus alerts: balance due before shoot, no deposit yet",
        "Gallery due dates with Overdue / Due soon warnings so every client gets their photos on time",
        "Mileage cost, shoot costs (second shooter, rentals) and profit per shoot",
        "Monthly summary with expenses, profit, tax set-aside and take-home pay",
        "Dashboard: next 10 shoots, revenue by shoot type, best lead sources, booking rate and charts; blank copy plus a sample copy",
    ],
    "tags": ["photography business", "photographer spreadsheet", "photography client tracker", "photography bookings", "wedding photographer",
             "photography income tracker", "photographer expenses", "small business", "google sheets", "excel template", "client tracker",
             "photography planner", "photography finance"],
    "summary": "Instant download: Excel + Google Sheets template, blank copy + sample copy.",
}
