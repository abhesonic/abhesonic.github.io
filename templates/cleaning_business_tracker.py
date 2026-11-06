"""Cleaning Business Client & Job Tracker – clients, recurring cleans, jobs, payments, cleaner pay, expenses, quote calculator.

Settings (business, year, services + prices, cleaners + pay, quote rates) -> Clients (frequency, price, access notes;
jobs, revenue, unpaid, last/next clean, overdue) -> Jobs (date, client, service, cleaner, hours, price, paid?)
-> Expenses -> Monthly (revenue, costs, profit) -> Quote (price a new home) -> Dashboard.
"""
import datetime as dt, random
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import BarChart, PieChart, Reference
from _common import Kit, F, B, MONEY, DATE, font, fill, style_range, input_cell, finish

MAIN = "0369A1"; SOFT = "E0F2FE"
K = Kit(MAIN, SOFT)
PCT = "0.0%"; HRS = '0.0;-0.0;"-"'
NSV = 12; SV0 = 6; SV1 = SV0 + NSV - 1        # services Settings A6:D17
NST = 10; ST0 = 6; ST1 = ST0 + NST - 1        # cleaners Settings F6:G15
NC = 100; C0 = 6; C1 = C0 + NC - 1            # Clients
NJ = 1000; J0 = 6; J1 = J0 + NJ - 1           # Jobs
NX = 300; X0 = 6; X1 = X0 + NX - 1            # Expenses
FREQ = '"Weekly,Bi-weekly,Every 3 weeks,Monthly,One-time"'
SERVICES = [("Standard clean", 140, 3), ("Deep clean", 280, 6), ("Move-in / move-out", 350, 7), ("Office clean", 180, 4),
            ("Airbnb turnover", 110, 2.5), ("Post-construction", 450, 9), ("Add-on only", 40, 1)]
STAFF = [("Owner (you)", 0), ("Lucy Grant", 20), ("Ana Pereira", 19), ("Tasha Moore", 18)]
EXP_CATS = ["Supplies", "Equipment", "Fuel / mileage", "Insurance", "Marketing", "Software / phone", "Uniforms", "Other"]
CLIENTS = [("Johnson Family", "Bi-weekly", "Standard clean", 150, "Key under mat", "1 dog", dt.date(2026, 1, 8)),
           ("Patel Residence", "Weekly", "Standard clean", 125, "Door code 4512", "", dt.date(2026, 2, 3)),
           ("Greenway Dental", "Weekly", "Office clean", 190, "Alarm code in app", "", dt.date(2026, 1, 12)),
           ("Seaside Airbnb #1", "Weekly", "Airbnb turnover", 115, "Lockbox 2290", "", dt.date(2026, 4, 4)),
           ("Miller Home", "Monthly", "Deep clean", 260, "Owner home", "2 cats", dt.date(2026, 3, 16)),
           ("Carter Apartment", "Bi-weekly", "Standard clean", 135, "Concierge has key", "", dt.date(2026, 5, 7)),
           ("Nguyen Family", "Every 3 weeks", "Standard clean", 160, "Garage code 7781", "1 dog", dt.date(2026, 6, 2)),
           ("Brooks Realty (move-outs)", "One-time", "Move-in / move-out", 380, "Agent meets you", "", dt.date(2026, 7, 21)),
           ("Lopez Townhouse", "Bi-weekly", "Standard clean", 140, "Key in lockbox", "", dt.date(2026, 8, 11)),
           ("Harbor Yoga Studio", "Weekly", "Office clean", 170, "Back door", "", dt.date(2026, 9, 1))]
DAYS = {"Weekly": 7, "Bi-weekly": 14, "Every 3 weeks": 21, "Monthly": 30}


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); s = wb.create_sheet("Settings"); cl = wb.create_sheet("Clients")
    jb = wb.create_sheet("Jobs"); ex = wb.create_sheet("Expenses"); mo = wb.create_sheet("Monthly"); qt = wb.create_sheet("Quote")

    # ---------- Settings ----------
    K.title(s, "Settings", "Yellow cells are yours to edit.")
    s["I3"] = "Business name"; s["I3"].font = font(bold=True); input_cell(s["J3"], "Sparkle & Shine Cleaning" if sample else "My Cleaning Business")
    s["I4"] = "Year for Dashboard / Monthly"; s["I4"].font = font(bold=True); input_cell(s["J4"], 2026 if sample else "=YEAR(TODAY())")
    s["I5"] = "Currency label"; s["I5"].font = font(bold=True); input_cell(s["J5"], "USD")
    s["I6"] = "Remind me when a clean is due within (days)"; s["I6"].font = font(bold=True); input_cell(s["J6"], 3)
    K.hdr(s, 5, ["Service", "Default price", "Est. hours", "Notes"])
    for i in range(NSV):
        r = SV0 + i; sv = SERVICES[i] if i < len(SERVICES) else (None, None, None)
        input_cell(s[f"A{r}"], sv[0]); input_cell(s[f"B{r}"], sv[1]); input_cell(s[f"C{r}"], sv[2]); input_cell(s[f"D{r}"])
        s[f"B{r}"].number_format = MONEY; s[f"C{r}"].number_format = HRS
    K.hdr(s, 5, ["Cleaner", "Pay per hour"], col0=6)
    for i in range(NST):
        r = ST0 + i; sf = STAFF[i] if i < len(STAFF) else (None, None)
        input_cell(s[f"F{r}"], sf[0]); input_cell(s[f"G{r}"], sf[1]); s[f"G{r}"].number_format = MONEY
    s[f"F{ST1 + 1}"] = "Pay 0 for yourself – your profit is what's left."; s[f"F{ST1 + 1}"].font = font(italic=True, size=9)
    K.hdr(s, 20, ["Expense categories"])
    for i in range(12): input_cell(s[f"A{21 + i}"], EXP_CATS[i] if i < len(EXP_CATS) else None)
    # quote rates
    K.hdr(s, 20, ["Quote calculator rates", "Amount"], col0=6)
    QR = [("Base fee per visit", 60, MONEY), ("Per bedroom", 20, MONEY), ("Per bathroom", 25, MONEY), ("Per 100 sq ft over 1,000", 4, MONEY),
          ("Deep clean multiplier", 1.8, "0.00"), ("Move-in / move-out multiplier", 2.2, "0.00"),
          ("Discount weekly", 0.15, PCT), ("Discount bi-weekly", 0.10, PCT), ("Discount every 3 weeks / monthly", 0.05, PCT),
          ("Inside oven", 35, MONEY), ("Inside fridge", 30, MONEY), ("Interior windows (each)", 6, MONEY), ("Laundry / bed change (per load)", 15, MONEY),
          ("Minutes of work per $ of price (for hours estimate)", 1.3, "0.00")]
    for i, (lab, v, fmt) in enumerate(QR, 21):
        s[f"F{i}"] = lab; s[f"F{i}"].font = font(); s[f"F{i}"].border = B; input_cell(s[f"G{i}"], v); s[f"G{i}"].number_format = fmt
    for c, w in zip("ABCDEFGHIJ", [24, 12, 10, 22, 3, 40, 12, 3, 36, 26]): s.column_dimensions[c].width = w
    SVC = f"Settings!$A${SV0}:$A${SV1}"; STF = f"Settings!$F${ST0}:$F${ST1}"
    YR = "Settings!$J$4"

    # ---------- Clients ----------
    K.title(cl, "Clients", "One row per client. Yellow = you type, grey = calculated from the Jobs tab.")
    ch_ = ["Client", "Phone", "Email", "Address", "Frequency", "Usual service", "Agreed price", "Access / key notes", "Pets / notes",
           "Start date", "Status", "Cleans done", "Revenue (all time)", "Unpaid", "Last clean", "Next clean due", "Reminder", "rank"]
    K.hdr(cl, 5, ch_, [24, 13, 24, 26, 13, 18, 11, 20, 14, 12, 10, 9, 13, 11, 12, 13, 16, 6])
    style_range(cl, C0, C1, 1, 18, inputs=tuple(range(1, 12)), fmt={7: MONEY, 10: DATE, 12: "0", 13: MONEY, 14: MONEY, 15: DATE, 16: DATE, 18: "0.000000"})
    JB = lambda c: f"Jobs!${c}${J0}:${c}${J1}"
    for r in range(C0, C1 + 1):
        cl[f"L{r}"] = f'=IF(A{r}="","",COUNTIFS({JB("B")},A{r},{JB("J")},"Done"))'
        cl[f"M{r}"] = f'=IF(A{r}="","",SUMIFS({JB("O")},{JB("B")},A{r}))'
        cl[f"N{r}"] = f'=IF(A{r}="","",SUMIFS({JB("R")},{JB("B")},A{r}))'
        cl[f"O{r}"] = f'=IF(OR(A{r}="",N(L{r})=0),"",SUMPRODUCT(MAX(({JB("B")}=A{r})*({JB("J")}="Done")*{JB("A")})))'
        cl[f"P{r}"] = (f'=IF(OR(A{r}="",K{r}<>"Active",E{r}="One-time",E{r}=""),"",'
                       f'IF(O{r}="",J{r},O{r}+IF(E{r}="Weekly",7,IF(E{r}="Bi-weekly",14,IF(E{r}="Every 3 weeks",21,30)))))')
        cl[f"Q{r}"] = (f'=IF(OR(P{r}="",NOT(ISNUMBER(P{r}))),"",IF(COUNTIFS({JB("B")},A{r},{JB("J")},"Scheduled",{JB("A")},">="&TODAY())>0,"Booked",'
                       f'IF(P{r}<TODAY(),"Overdue",IF(P{r}-TODAY()<=Settings!$J$6,"Due soon – book",""))))')
        cl[f"R{r}"] = f'=IF(A{r}="","",N(M{r})+ROW()/1000000)'
    for rng, f1 in ((f"E{C0}:E{C1}", FREQ), (f"F{C0}:F{C1}", f"={SVC}"), (f"K{C0}:K{C1}", '"Active,Paused,Lost"')):
        dv = DataValidation(type="list", formula1=f1, allow_blank=True); cl.add_data_validation(dv); dv.add(rng)
    cl.conditional_formatting.add(f"Q{C0}:Q{C1}", FormulaRule(formula=[f'Q{C0}="Overdue"'], font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    cl.conditional_formatting.add(f"Q{C0}:Q{C1}", FormulaRule(formula=[f'LEFT(Q{C0},3)="Due"'], font=Font(name=F, bold=True, color="B45309"), fill=fill("FEF3C7")))
    cl.conditional_formatting.add(f"Q{C0}:Q{C1}", FormulaRule(formula=[f'Q{C0}="Booked"'], font=Font(name=F, color="15803D")))
    cl.conditional_formatting.add(f"N{C0}:N{C1}", FormulaRule(formula=[f'AND(ISNUMBER(N{C0}),N{C0}>0)'], font=Font(name=F, bold=True, color="B91C1C")))
    cl.column_dimensions["R"].hidden = True
    cl.freeze_panes = f"B{C0}"
    CN = f"Clients!$A${C0}:$A${C1}"

    # ---------- Jobs ----------
    K.title(jb, "Jobs", "One row per clean. Leave Price empty to use the client's agreed price (or the service's default price).")
    jh = ["Date", "Client", "Service", "Cleaner", "Hours", "Price (optional)", "Extras", "Discount", "Supplies cost", "Status", "Paid?",
          "Payment method", "Notes", "Base price", "Total charge", "Cleaner pay", "Profit", "Unpaid", "Month", "Year"]
    K.hdr(jb, 5, jh, [12, 24, 18, 15, 8, 11, 9, 9, 10, 11, 8, 13, 22, 11, 12, 11, 11, 10, 7, 7])
    style_range(jb, J0, J1, 1, 20, inputs=tuple(range(1, 14)), fmt={1: DATE, 5: HRS, 6: MONEY, 7: MONEY, 8: MONEY, 9: MONEY, 14: MONEY,
                                                                     15: MONEY, 16: MONEY, 17: MONEY, 18: MONEY, 19: "0", 20: "0"})
    for r in range(J0, J1 + 1):
        jb[f"N{r}"] = (f'=IF(B{r}="",0,IF(F{r}<>"",F{r},IF(AND(C{r}<>"",C{r}<>IFERROR(INDEX(Clients!$F${C0}:$F${C1},MATCH(B{r},{CN},0)),"")),'
                       f'IFERROR(INDEX(Settings!$B${SV0}:$B${SV1},MATCH(C{r},{SVC},0)),0),'
                       f'IFERROR(INDEX(Clients!$G${C0}:$G${C1},MATCH(B{r},{CN},0)),0))))')
        jb[f"O{r}"] = f'=IF(J{r}="Done",N{r}+N(G{r})-N(H{r}),0)'
        jb[f"P{r}"] = f'=IF(J{r}="Done",N(E{r})*IFERROR(INDEX(Settings!$G${ST0}:$G${ST1},MATCH(D{r},{STF},0)),0),0)'
        jb[f"Q{r}"] = f'=IF(J{r}="Done",O{r}-P{r}-N(I{r}),0)'
        jb[f"R{r}"] = f'=IF(AND(J{r}="Done",K{r}<>"Yes"),O{r},0)'
        jb[f"S{r}"] = f'=IF(A{r}="","",MONTH(A{r}))'
        jb[f"T{r}"] = f'=IF(A{r}="","",YEAR(A{r}))'
    for rng, f1 in ((f"B{J0}:B{J1}", f"={CN}"), (f"C{J0}:C{J1}", f"={SVC}"), (f"D{J0}:D{J1}", f"={STF}"),
                    (f"J{J0}:J{J1}", '"Scheduled,Done,Cancelled,No-show"'), (f"K{J0}:K{J1}", '"Yes,No"'),
                    (f"L{J0}:L{J1}", '"Cash,Card,Bank transfer,Check,Zelle / Venmo / PayPal,Other"')):
        dv = DataValidation(type="list", formula1=f1, allow_blank=True); jb.add_data_validation(dv); dv.add(rng)
    jb.conditional_formatting.add(f"J{J0}:J{J1}", FormulaRule(formula=[f'J{J0}="Scheduled"'], font=Font(name=F, bold=True, color=MAIN)))
    jb.conditional_formatting.add(f"J{J0}:J{J1}", FormulaRule(formula=[f'OR(J{J0}="Cancelled",J{J0}="No-show")'], font=Font(name=F, color="9CA3AF")))
    jb.conditional_formatting.add(f"R{J0}:R{J1}", FormulaRule(formula=[f'R{J0}>0'], font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    jb.freeze_panes = f"C{J0}"

    # ---------- Expenses ----------
    K.title(ex, "Expenses", "Business costs that are not tied to one job: supplies stock, fuel, insurance, ads…")
    K.hdr(ex, 5, ["Date", "Category", "Description", "Amount", "Month", "Year"], [12, 18, 34, 12, 8, 8])
    style_range(ex, X0, X1, 1, 6, inputs=(1, 2, 3, 4), fmt={1: DATE, 4: MONEY, 5: "0", 6: "0"})
    for r in range(X0, X1 + 1):
        ex[f"E{r}"] = f'=IF(A{r}="","",MONTH(A{r}))'; ex[f"F{r}"] = f'=IF(A{r}="","",YEAR(A{r}))'
    dv = DataValidation(type="list", formula1="=Settings!$A$21:$A$32", allow_blank=True); ex.add_data_validation(dv); dv.add(f"B{X0}:B{X1}")
    ex.freeze_panes = f"A{X0}"

    # ---------- sample data ----------
    if sample:
        rnd = random.Random(7)
        for i, (n, fr, sv, pr, acc, pet, sd) in enumerate(CLIENTS):
            r = C0 + i
            for col, v in zip("ABCDEFGHIJK", [n, f"(555) 01{20 + i}-{3000 + 17 * i}", n.split()[0].lower().strip("(#") + "@mail.example",
                                              f"{100 + 37 * i} {['Oak', 'Maple', 'Harbor', 'Pine', 'Cedar'][i % 5]} St", fr, sv, pr, acc, pet or None, sd,
                                              "Active"]):
                cl[f"{col}{r}"] = v
        jobs = []; today = dt.date(2026, 11, 6); end = dt.date(2026, 11, 20)
        cleaners = ["Lucy Grant", "Ana Pereira", "Tasha Moore", "Owner (you)"]
        for i, (n, fr, sv, pr, acc, pet, sd) in enumerate(CLIENTS):
            hrs = dict((a, c) for a, b, c in SERVICES)[sv]
            if fr == "One-time":
                for d in (sd, dt.date(2026, 9, 14), dt.date(2026, 10, 27)):
                    jobs.append((d, n, sv, "Lucy Grant", 7, pr if d == sd else 395, 0, 0, 22, "Done", "Yes" if d != dt.date(2026, 10, 27) else "No"))
                continue
            d = sd; k = 0
            stop = dt.date(2026, 10, 10) if i == 6 else end      # Nguyen overdue
            while d <= stop:
                stt = "Done" if d <= today else "Scheduled"
                if stt == "Done" and rnd.random() < 0.04: stt = "Cancelled"
                paid = "Yes" if (stt == "Done" and (today - d).days > 10 or rnd.random() < 0.8) else "No"
                extra = 30 if rnd.random() < 0.12 else 0
                svc = sv if not (fr == "Monthly" or (k % 8 == 7 and sv == "Standard clean")) else "Deep clean"
                h = hrs + (2 if svc != sv else 0) + rnd.choice([-0.5, 0, 0, 0.5])
                jobs.append((d, n, svc, cleaners[(i + k) % 3] if i % 4 else "Owner (you)", h, None, extra, 0, rnd.choice([6, 8, 10]),
                             stt, paid if stt == "Done" else None))
                d += dt.timedelta(days=DAYS[fr]); k += 1
        if True:  # a deep-clean upsell and a discount example
            jobs.append((dt.date(2026, 11, 12), "Johnson Family", "Deep clean", "Ana Pereira", 6, None, 0, 25, 12, "Scheduled", None))
        jobs.sort(key=lambda j: (j[0], j[1]))
        for r, j in enumerate(jobs, J0):
            for col, v in zip("ABCDEFGHIJK", j): jb[f"{col}{r}"] = v
            if j[9] == "Done" and j[10] == "Yes": jb[f"L{r}"] = ["Card", "Bank transfer", "Cash", "Zelle / Venmo / PayPal"][r % 4]
        assert len(jobs) <= NJ, len(jobs)
        xs = []
        for m in range(1, 12):
            xs += [(dt.date(2026, m, 3), "Supplies", "Cleaning products restock", 85 + 5 * (m % 4)),
                   (dt.date(2026, m, 10), "Fuel / mileage", "Gas", 120 + 10 * (m % 3)),
                   (dt.date(2026, m, 15), "Software / phone", "Phone + booking app", 45)]
            if m in (1, 7): xs.append((dt.date(2026, m, 20), "Insurance", "Liability insurance (6 months)", 420))
            if m in (3, 9): xs.append((dt.date(2026, m, 22), "Marketing", "Flyers + Facebook ads", 150))
        xs.append((dt.date(2026, 2, 18), "Equipment", "New vacuum", 320))
        for r, x in enumerate(xs, X0):
            for col, v in zip("ABCD", x): ex[f"{col}{r}"] = v

    # ---------- Monthly ----------
    K.title(mo, "Monthly Summary", "Calculated from Jobs and Expenses for the year on the Settings tab.")
    mo["A3"] = '="Year "&Settings!J4&"  ·  amounts in "&Settings!J5'; mo["A3"].font = font(bold=True)
    K.hdr(mo, 5, ["Month", "Cleans done", "Revenue", "Cleaner pay", "Job supplies", "Other expenses", "Profit", "Profit margin", "Unpaid", "Cancelled / no-show"],
          [12, 11, 13, 13, 13, 14, 13, 11, 12, 13])
    for m in range(1, 13):
        r = 5 + m
        mo[f"A{r}"] = f'=TEXT(DATE({YR},{m},1),"mmm yyyy")'
        cnd = f'{JB("S")},{m},{JB("T")},{YR}'
        mo[f"B{r}"] = f'=COUNTIFS({cnd},{JB("J")},"Done")'
        mo[f"C{r}"] = f"=SUMIFS({JB('O')},{cnd})"
        mo[f"D{r}"] = f"=SUMIFS({JB('P')},{cnd})"
        mo[f"E{r}"] = f'=SUMIFS({JB("I")},{cnd},{JB("J")},"Done")'
        mo[f"F{r}"] = f"=SUMIFS(Expenses!$D${X0}:$D${X1},Expenses!$E${X0}:$E${X1},{m},Expenses!$F${X0}:$F${X1},{YR})"
        mo[f"G{r}"] = f"=C{r}-D{r}-E{r}-F{r}"
        mo[f"H{r}"] = f'=IF(C{r}=0,"",G{r}/C{r})'
        mo[f"I{r}"] = f"=SUMIFS({JB('R')},{cnd})"
        mo[f"J{r}"] = f'=COUNTIFS({cnd},{JB("J")},"Cancelled")+COUNTIFS({cnd},{JB("J")},"No-show")'
    style_range(mo, 6, 17, 1, 10, fmt={2: "0", 3: MONEY, 4: MONEY, 5: MONEY, 6: MONEY, 7: MONEY, 8: PCT, 9: MONEY, 10: "0"})
    mo["A18"] = "TOTAL"
    for c in "BCDEFGIJ": mo[f"{c}18"] = f"=SUM({c}6:{c}17)"
    mo["H18"] = '=IF(C18=0,"",G18/C18)'
    for c, fmt in zip("ABCDEFGHIJ", [None, "0", MONEY, MONEY, MONEY, MONEY, MONEY, PCT, MONEY, "0"]):
        cell = mo[f"{c}18"]; cell.font = Font(name=F, bold=True, color="FFFFFF"); cell.fill = fill(MAIN); cell.border = B
        if fmt: cell.number_format = fmt
    mo.conditional_formatting.add("G6:G17", FormulaRule(formula=["G6<0"], font=Font(name=F, bold=True, color="B91C1C")))

    # ---------- Quote ----------
    K.title(qt, "Quote Calculator", "Price a new home in seconds. Type in the yellow cells; rates are on the Settings tab.")
    qin = [("Client / address", "New lead – 12 Birch Lane" if sample else None, None), ("Bedrooms", 3, "0"), ("Bathrooms", 2, "0"),
           ("Square feet", 1800, "#,##0"), ("Type of clean", "Standard", None), ("How often", "Bi-weekly", None),
           ("Inside oven (0/1)", 0, "0"), ("Inside fridge (0/1)", 0, "0"), ("Interior windows (number)", 0, "0"), ("Laundry loads", 0, "0")]
    for i, (lab, v, fmt) in enumerate(qin, 4):
        qt[f"A{i}"] = lab; qt[f"A{i}"].font = font(bold=True); input_cell(qt[f"B{i}"], v)
        if fmt: qt[f"B{i}"].number_format = fmt
    for rng, f1 in (("B8", '"Standard,Deep clean,Move-in / move-out"'), ("B9", FREQ)):
        dv = DataValidation(type="list", formula1=f1, allow_blank=False); qt.add_data_validation(dv); dv.add(rng)
    G = lambda n: f"Settings!$G${20 + n}"
    qout = [("Base price (rooms + size)", f"={G(1)}+B5*{G(2)}+B6*{G(3)}+MAX(0,B7-1000)/100*{G(4)}", MONEY),
            ("Type multiplier", f'=IF(B8="Deep clean",{G(5)},IF(B8="Move-in / move-out",{G(6)},1))', "0.00"),
            ("Extras", f"=B10*{G(10)}+B11*{G(11)}+B12*{G(12)}+B13*{G(13)}", MONEY),
            ("Price before discount", "=ROUND(B15*B16+B17,0)", MONEY),
            ("Recurring discount", f'=IF(B8<>"Standard",0,IF(B9="Weekly",{G(7)},IF(B9="Bi-weekly",{G(8)},IF(OR(B9="Every 3 weeks",B9="Monthly"),{G(9)},0))))', PCT),
            ("QUOTE PER VISIT", "=ROUND(B18*(1-B19)/5,0)*5", MONEY),
            ("Estimated hours (1 cleaner)", f"=ROUND(B20*{G(14)}/60*2,0)/2", HRS),
            ("Estimated per month", f'=IF(B9="Weekly",B20*52/12,IF(B9="Bi-weekly",B20*26/12,IF(B9="Every 3 weeks",B20*52/3/12,IF(B9="Monthly",B20,0))))', MONEY)]
    for i, (lab, f, fmt) in enumerate(qout, 15):
        qt[f"A{i}"] = lab; qt[f"A{i}"].font = font(bold=True); c = qt[f"B{i}"]; c.value = f; c.number_format = fmt
        c.border = B; c.fill = fill("F3F4F6"); c.font = font()
    for c in "AB":
        qt[f"{c}20"].font = Font(name=F, size=14, bold=True, color=MAIN); qt[f"{c}20"].fill = fill(SOFT)
    qt["A24"] = "Rounded to the nearest 5. Recurring discounts apply to standard cleans only. Change all rates on the Settings tab."
    qt["A24"].font = font(italic=True, size=9)
    qt.column_dimensions["A"].width = 32; qt.column_dimensions["B"].width = 28

    # ---------- Dashboard ----------
    K.title(dash, "Cleaning Business Dashboard", "Updates by itself from the Clients, Jobs and Expenses tabs.")
    dash["A3"] = '=Settings!J3&"  ·  year "&Settings!J4&"  ·  amounts in "&Settings!J5'; dash["A3"].font = font(bold=True)
    kp = [("Revenue this year", "=Monthly!C18", MONEY), ("Profit this year", "=Monthly!G18", MONEY), ("Cleans done", "=Monthly!B18", "0"),
          ("Average price per clean", "=IFERROR(Monthly!C18/Monthly!B18,0)", MONEY),
          ("Unpaid (all time)", f"=SUM({JB('R')})", MONEY), ("Active clients", f'=COUNTIF(Clients!$K${C0}:$K${C1},"Active")', "0")]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 3) * 3, 1 + (i % 3) * 2, lab, f, fmt)
        dash.cell(row=6 + (i // 3) * 3, column=1 + (i % 3) * 2).alignment = Alignment(horizontal="left")
    for c, w in zip("ABCDEFG", [32, 14, 22, 14, 22, 14, 3]): dash.column_dimensions[c].width = w
    K.section(dash, "A11", "To do")
    todo = [("Clients overdue for a clean", f'=COUNTIF(Clients!$Q${C0}:$Q${C1},"Overdue")', "0"),
            ("Clients due soon – not booked", f'=COUNTIF(Clients!$Q${C0}:$Q${C1},"Due*")', "0"),
            ("Cleans booked next 7 days", f'=COUNTIFS({JB("J")},"Scheduled",{JB("A")},">="&TODAY(),{JB("A")},"<"&TODAY()+7)', "0"),
            ("Expected income next 7 days", f'=SUMPRODUCT(({JB("J")}="Scheduled")*({JB("A")}>=TODAY())*({JB("A")}<TODAY()+7)*({JB("N")}+{JB("G")}-{JB("H")}))', MONEY),
            ("Jobs done but not paid", f'=COUNTIF({JB("R")},">0")', "0"),
            ("Cancelled / no-show this year", "=Monthly!J18", "0"),
            ("Profit margin this year", '=IFERROR(Monthly!G18/Monthly!C18,0)', PCT)]
    for i, (lab, f, fmt) in enumerate(todo, 12):
        dash[f"A{i}"] = lab; dash[f"A{i}"].font = font(bold=True); dash[f"A{i}"].border = B; dash[f"A{i}"].fill = fill(SOFT)
        c = dash[f"B{i}"]; c.value = f; c.number_format = fmt; c.border = B; c.font = Font(name=F, bold=True, color=MAIN)
    dash.conditional_formatting.add("B12", FormulaRule(formula=["B12>0"], font=Font(name=F, bold=True, color="B91C1C")))
    TR = 21
    K.section(dash, f"A{TR}", "Top 10 clients (all-time revenue)")
    K.hdr(dash, TR + 1, ["Client", "Revenue", "Cleans", "Frequency", "Reminder"])
    RK = f"Clients!$R${C0}:$R${C1}"
    for k in range(10):
        r = TR + 2 + k
        dash[f"A{r}"] = f'=IFERROR(INDEX({CN},MATCH(LARGE({RK},{k + 1}),{RK},0)),"")'
        dash[f"B{r}"] = f'=IF(A{r}="","",INDEX(Clients!$M${C0}:$M${C1},MATCH(A{r},{CN},0)))'
        dash[f"C{r}"] = f'=IF(A{r}="","",INDEX(Clients!$L${C0}:$L${C1},MATCH(A{r},{CN},0)))'
        dash[f"D{r}"] = f'=IF(A{r}="","",INDEX(Clients!$E${C0}:$E${C1},MATCH(A{r},{CN},0))&"")'
        dash[f"E{r}"] = f'=IF(A{r}="","",INDEX(Clients!$Q${C0}:$Q${C1},MATCH(A{r},{CN},0))&"")'
    style_range(dash, TR + 2, TR + 11, 1, 5, fmt={2: MONEY, 3: "0"})
    SR = TR + 13
    K.section(dash, f"A{SR}", "Revenue by service (this year)")
    K.hdr(dash, SR + 1, ["Service", "Cleans", "Revenue", "Avg price"])
    for i in range(NSV):
        r = SR + 2 + i; k = SV0 + i
        dash[f"A{r}"] = f'=IF(Settings!A{k}="","",Settings!A{k})'
        dash[f"B{r}"] = f'=IF(A{r}="","",COUNTIFS({JB("C")},A{r},{JB("J")},"Done",{JB("T")},{YR}))'
        dash[f"C{r}"] = f'=IF(A{r}="","",SUMIFS({JB("O")},{JB("C")},A{r},{JB("T")},{YR}))'
        dash[f"D{r}"] = f'=IF(OR(A{r}="",N(B{r})=0),"",C{r}/B{r})'
    SE = SR + 1 + NSV
    style_range(dash, SR + 2, SE, 1, 4, fmt={2: "0", 3: MONEY, 4: MONEY})
    ch = BarChart(); ch.type = "col"; ch.title = "Revenue and profit by month"; ch.height = 8; ch.width = 17
    ch.add_data(Reference(mo, min_col=3, min_row=5, max_row=17), titles_from_data=True)
    ch.add_data(Reference(mo, min_col=7, min_row=5, max_row=17), titles_from_data=True)
    ch.set_categories(Reference(mo, min_col=1, min_row=6, max_row=17)); ch.y_axis.majorGridlines = None
    ch.series[0].graphicalProperties.solidFill = MAIN; ch.series[1].graphicalProperties.solidFill = "38BDF8"
    dash.add_chart(ch, "H5")
    ch2 = BarChart(); ch2.type = "bar"; ch2.title = "Revenue by service"; ch2.height = 8; ch2.width = 17
    ch2.add_data(Reference(dash, min_col=3, min_row=SR + 1, max_row=SR + 1 + len(SERVICES)), titles_from_data=True)
    ch2.set_categories(Reference(dash, min_col=1, min_row=SR + 2, max_row=SR + 1 + len(SERVICES))); ch2.legend = None
    ch2.x_axis.scaling.orientation = "maxMin"; ch2.y_axis.majorGridlines = None; ch2.series[0].graphicalProperties.solidFill = MAIN
    dash.add_chart(ch2, "H22")

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Settings tab: business name, year, your services with default prices, your cleaners with their hourly pay, and the quote rates.", False),
             ("2.  Clients tab: one row per client – how often you clean (weekly, bi-weekly…), usual service, agreed price, key/alarm notes and pets.", False),
             ("3.  Jobs tab: one row per clean – date, client, cleaner, hours and status (Scheduled, Done, Cancelled, No-show). Mark 'Paid?' Yes when paid. Leave 'Price' empty to use the client's agreed price.", False),
             ("4.  Expenses tab: supplies stock, fuel, insurance, ads and other costs.", False),
             ("5.  Clients tab shows last clean, next clean due and a reminder: Overdue (red), Due soon – book (orange) or Booked (green).", False),
             ("6.  Quote tab: type bedrooms, bathrooms, square feet, type and frequency to price a new home in seconds.", False),
             ("7.  Monthly and Dashboard: revenue, cleaner pay, profit, unpaid invoices, top clients and charts – all automatic.", False),
             ("", False), ("Tips", True),
             ("•  Book the next visit as a 'Scheduled' row. On the day, change it to 'Done' – revenue and cleaner pay count only when a job is Done.", False),
             ("•  Use 'Extras' for add-ons (oven, fridge, windows) and 'Discount' for coupons or referral rewards.", False),
             ("•  'Supplies cost' on a job is for products used on that job; buy-in-bulk supplies go on the Expenses tab.", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  100 clients, 1,000 jobs, 300 expenses, 12 services, 10 cleaners. No macros. Works in Excel 2010+ and Google Sheets.", False),
             ("•  To use Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("The business, clients, jobs and prices are made-up examples. Use the blank copy in your download for your own business.", False)]
    K.start(st, "Cleaning Business Client & Job Tracker", "Clients, recurring cleans, payments, cleaner pay, profit and a quote calculator – Excel and Google Sheets.", lines)
    finish(wb, sample, path)


META = {
    "name": "Cleaning Business Client & Job Tracker (Recurring Cleans, Payments, Profit, Quote Calculator) – Excel & Google Sheets",
    "file_base": "Cleaning-Business-Client-Job-Tracker",
    "price": 19,
    "color": "#0369A1",
    "title": "Cleaning Business Tracker",
    "tagline": "clients · recurring cleans · payments · cleaner pay · profit · quotes",
    "intro": "Run your house-cleaning or janitorial business from one simple spreadsheet: keep every client's frequency, price and key notes, log each clean, see who hasn't paid, what you paid your cleaners and what profit is left. Recurring clients get a 'next clean due' date and turn red when they are overdue, and a quote calculator prices a new home in seconds.",
    "features": [
        "Client list for up to 100 clients: frequency (weekly, bi-weekly, every 3 weeks, monthly), usual service, agreed price, access/key and pet notes",
        "Automatic last clean, next clean due and reminders: Overdue, Due soon – book, Booked",
        "Job log for 1,000 cleans: client price fills in by itself, extras, discounts, supplies, status and paid/unpaid",
        "Cleaner pay per job from hourly rates, and profit per job, per month and per year",
        "Quote calculator: bedrooms, bathrooms, square feet, deep/move-out multipliers, extras and recurring discounts",
        "Monthly summary and expenses tracker (supplies, fuel, insurance, ads)",
        "Dashboard with top 10 clients, revenue by service, unpaid total and charts; blank copy plus a sample copy",
    ],
    "tags": ["cleaning business", "cleaning client tracker", "house cleaning", "cleaning business planner", "janitorial", "client tracker",
             "cleaning quote calculator", "small business", "job tracker", "google sheets", "excel template", "cleaning schedule", "maid service"],
    "summary": "Instant download: Excel + Google Sheets template, blank copy + sample copy.",
}
