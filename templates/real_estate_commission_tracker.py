"""Real Estate Agent Commission & Lead Tracker – transactions, GCI, brokerage split with annual cap, referral/team/transaction fees,
net income, lead follow-ups, lead-source ROI, expenses, monthly summary and a GCI goal planner.

Settings -> Transactions (price, commission %, side, source, status -> GCI, cap-aware brokerage split, fees, net) -> Leads (stage, follow-up)
-> Expenses -> Monthly -> Goal Planner -> Dashboard.
"""
import datetime as dt, random
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, DataBarRule
from openpyxl.chart import BarChart, Reference
from _common import Kit, F, B, MONEY, DATE, font, fill, style_range, input_cell, finish

MAIN = "7C2D12"; SOFT = "FFEDD5"
K = Kit(MAIN, SOFT)
PCT = "0.0%"; USD0 = '#,##0;(#,##0);"-"'
NT = 200; T0 = 6; T1 = T0 + NT - 1      # Transactions
NL = 500; L0 = 6; L1 = L0 + NL - 1      # Leads
NX = 300; X0 = 6; X1 = X0 + NX - 1      # Expenses
SOURCES = ["Sphere / past client", "Referral", "Open house", "Zillow / portals", "Facebook / Instagram ads", "Website / SEO",
           "Sign call", "Door knocking", "Agent referral", "Other"]
NS = 12; S0 = 21; S1 = S0 + NS - 1       # lead sources Settings A21:A32
EXP_CATS = ["MLS & board dues", "Brokerage desk fee", "Marketing & ads", "Lead generation", "Photography & staging", "Signs & lockboxes",
            "Car & mileage", "Phone & software", "Client gifts", "Education & licensing", "E&O insurance", "Other"]
STAGES = '"New,Contacted,Appointment set,Agreement signed,Under contract,Closed,Nurture,Lost"'


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); s = wb.create_sheet("Settings"); tr = wb.create_sheet("Transactions")
    ld = wb.create_sheet("Leads"); ex = wb.create_sheet("Expenses"); mo = wb.create_sheet("Monthly"); gp = wb.create_sheet("Goal Planner")

    # ---------- Settings ----------
    K.title(s, "Settings", "Yellow cells are yours to edit. Your brokerage plan lives here.")
    SET = [("Agent / team name", "Jordan Ellis – Coastal Realty" if sample else "My Real Estate Business", None),
           ("Year for Dashboard / Monthly", 2026 if sample else "=YEAR(TODAY())", "0"),
           ("Currency label", "USD", None),
           ("Your split (share of commission you keep before cap)", 0.70, PCT),
           ("Annual cap (most you pay the brokerage per year, 0 = no cap)", 18000, USD0),
           ("Cap year starts on (month number, 1 = January)", 1, "0"),
           ("Franchise / royalty fee (% of commission)", 0.06, PCT),
           ("Franchise fee cap per year (0 = no cap)", 3000, USD0),
           ("Transaction / admin fee per closing", 395, MONEY),
           ("Tax set-aside (% of net income)", 0.25, PCT),
           ("Default commission rate (used when left empty)", 0.025, "0.00%"),
           ("Follow-up reminder: days before due", 2, "0")]
    for i, (lab, v, fmt) in enumerate(SET, 4):
        s[f"A{i}"] = lab; s[f"A{i}"].font = font(bold=True); s[f"A{i}"].border = B
        input_cell(s[f"B{i}"], v)
        if fmt: s[f"B{i}"].number_format = fmt
    K.hdr(s, 20, ["Lead sources"])
    for i in range(NS): input_cell(s[f"A{S0 + i}"], SOURCES[i] if i < len(SOURCES) else None)
    K.hdr(s, 20, ["Expense categories"], col0=4)
    for i in range(14): input_cell(s[f"D{21 + i}"], EXP_CATS[i] if i < len(EXP_CATS) else None)
    s["A35"] = "Split example: 70% means you keep 70% and the brokerage gets 30% of each commission until the cap is reached; after that you keep 100%."
    s["A35"].font = font(italic=True, size=9)
    for c, w in zip("ABCD", [58, 30, 3, 28]): s.column_dimensions[c].width = w
    YR = "Settings!$B$5"; SPLIT = "Settings!$B$7"; CAP = "Settings!$B$8"; CAPM = "Settings!$B$9"
    FR = "Settings!$B$10"; FRCAP = "Settings!$B$11"; TXF = "Settings!$B$12"; TAX = "Settings!$B$13"; DEFR = "Settings!$B$14"
    SRC = f"Settings!$A${S0}:$A${S1}"

    # ---------- Transactions ----------
    K.title(tr, "Transactions", "One row per deal. Yellow = you type. Grey = calculated. Money counts in income only when Status is Closed.")
    th = ["Client", "Property address", "Side", "Lead source", "Status", "Contract date", "Close date", "Sale price", "Commission %",
          "Referral fee paid %", "Team lead split %", "Bonus / other income", "Notes",
          "GCI", "Referral paid", "Adjusted GCI", "Brokerage split", "Franchise fee", "Transaction fee", "Team split", "Net to you",
          "Tax set-aside", "Cap year", "Month", "Year", "raw split", "rank"]
    K.hdr(tr, 5, th, [20, 26, 9, 20, 12, 12, 12, 13, 10, 10, 10, 11, 18, 12, 11, 12, 12, 11, 11, 11, 12, 11, 8, 7, 7, 8, 6])
    style_range(tr, T0, T1, 1, 27, inputs=tuple(range(1, 14)),
                fmt={6: DATE, 7: DATE, 8: USD0, 9: "0.00%", 10: PCT, 11: PCT, 12: MONEY, 14: MONEY, 15: MONEY, 16: MONEY, 17: MONEY,
                     18: MONEY, 19: MONEY, 20: MONEY, 21: MONEY, 22: MONEY, 23: "0", 24: "0", 25: "0", 26: MONEY, 27: "0.000000"})
    TT = lambda c: f"Transactions!${c}${T0}:${c}${T1}"
    for r in range(T0, T1 + 1):
        cl = f'E{r}="Closed"'
        tr[f"N{r}"] = f'=IF(OR(A{r}="",H{r}=""),0,H{r}*IF(I{r}="",{DEFR},I{r}))'
        tr[f"O{r}"] = f"=N{r}*N(J{r})"
        tr[f"P{r}"] = f"=N{r}-O{r}"
        # brokerage split before cap (only closed deals)
        tr[f"Z{r}"] = f"=IF({cl},P{r}*(1-{SPLIT}),0)"
        # cap year: year in which this cap period starts
        tr[f"W{r}"] = f'=IF(G{r}="","",YEAR(G{r})-IF(MONTH(G{r})<{CAPM},1,0))'
        prior = (f'SUMIFS($Z${T0}:$Z${T1},$W${T0}:$W${T1},W{r},$G${T0}:$G${T1},"<"&G{r})'
                 f'+SUMIFS($Z${T0}:$Z{r},$W${T0}:$W{r},W{r},$G${T0}:$G{r},G{r})-Z{r}')
        tr[f"Q{r}"] = f'=IF(NOT({cl}),0,IF({CAP}<=0,Z{r},MIN({CAP},{prior}+Z{r})-MIN({CAP},{prior})))'
        priorf = (f'SUMIFS($N${T0}:$N${T1},$W${T0}:$W${T1},W{r},$G${T0}:$G${T1},"<"&G{r},$E${T0}:$E${T1},"Closed")'
                  f'+SUMIFS($N${T0}:$N{r},$W${T0}:$W{r},W{r},$G${T0}:$G{r},G{r},$E${T0}:$E{r},"Closed")-N{r}')
        tr[f"R{r}"] = (f'=IF(NOT({cl}),0,IF({FRCAP}<=0,P{r}*{FR},MIN({FRCAP},({priorf})*{FR}+P{r}*{FR})-MIN({FRCAP},({priorf})*{FR})))')
        tr[f"S{r}"] = f"=IF({cl},{TXF},0)"
        tr[f"T{r}"] = f"=IF({cl},(P{r}-Q{r}-R{r})*N(K{r}),0)"
        tr[f"U{r}"] = f"=IF({cl},P{r}-Q{r}-R{r}-S{r}-T{r}+N(L{r}),0)"
        tr[f"V{r}"] = f"=U{r}*{TAX}"
        tr[f"X{r}"] = f'=IF(G{r}="","",MONTH(G{r}))'
        tr[f"Y{r}"] = f'=IF(G{r}="","",YEAR(G{r}))'
        tr[f"AA{r}"] = f'=IF(A{r}="","",N(U{r})+ROW()/1000000)'
    for rng, f1 in ((f"C{T0}:C{T1}", '"Buyer,Seller,Dual,Lease,Referral out"'), (f"D{T0}:D{T1}", f"={SRC}"),
                    (f"E{T0}:E{T1}", '"Active listing,Under contract,Pending,Closed,Fell through"')):
        dv = DataValidation(type="list", formula1=f1, allow_blank=True); tr.add_data_validation(dv); dv.add(rng)
    tr.conditional_formatting.add(f"E{T0}:E{T1}", FormulaRule(formula=[f'E{T0}="Closed"'], font=Font(name=F, bold=True, color="15803D")))
    tr.conditional_formatting.add(f"E{T0}:E{T1}", FormulaRule(formula=[f'E{T0}="Fell through"'], font=Font(name=F, color="9CA3AF")))
    tr.conditional_formatting.add(f"E{T0}:E{T1}", FormulaRule(formula=[f'OR(E{T0}="Pending",E{T0}="Under contract")'], font=Font(name=F, bold=True, color="B45309")))
    for c in ("W", "X", "Y", "Z", "AA"): tr.column_dimensions[c].hidden = True
    tr.freeze_panes = f"C{T0}"

    # ---------- Leads ----------
    K.title(ld, "Leads", "Every lead in one place. Set a next follow-up date – overdue follow-ups turn red.")
    lh = ["Date added", "Name", "Phone", "Email", "Lead source", "Buyer / Seller", "Price range", "Stage", "Last contact", "Next follow-up",
          "Notes", "Follow-up", "Year"]
    K.hdr(ld, 5, lh, [12, 22, 14, 24, 20, 12, 14, 16, 12, 13, 30, 14, 7])
    style_range(ld, L0, L1, 1, 13, inputs=tuple(range(1, 12)), fmt={1: DATE, 9: DATE, 10: DATE, 13: "0"})
    for r in range(L0, L1 + 1):
        ld[f"L{r}"] = (f'=IF(OR(B{r}="",J{r}="",H{r}="Closed",H{r}="Lost"),"",IF(J{r}<TODAY(),"Overdue",'
                       f'IF(J{r}=TODAY(),"Today",IF(J{r}-TODAY()<=Settings!$B$15,"Soon",""))))')
        ld[f"M{r}"] = f'=IF(A{r}="","",YEAR(A{r}))'
    for rng, f1 in ((f"E{L0}:E{L1}", f"={SRC}"), (f"F{L0}:F{L1}", '"Buyer,Seller,Both,Renter,Investor"'), (f"H{L0}:H{L1}", STAGES)):
        dv = DataValidation(type="list", formula1=f1, allow_blank=True); ld.add_data_validation(dv); dv.add(rng)
    ld.conditional_formatting.add(f"L{L0}:L{L1}", FormulaRule(formula=[f'L{L0}="Overdue"'], font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    ld.conditional_formatting.add(f"L{L0}:L{L1}", FormulaRule(formula=[f'OR(L{L0}="Today",L{L0}="Soon")'], font=Font(name=F, bold=True, color="B45309"), fill=fill("FEF3C7")))
    ld.column_dimensions["M"].hidden = True
    ld.freeze_panes = f"C{L0}"
    LL = lambda c: f"Leads!${c}${L0}:${c}${L1}"

    # ---------- Expenses ----------
    K.title(ex, "Expenses", "Business costs: dues, desk fees, marketing, lead generation, mileage…")
    K.hdr(ex, 5, ["Date", "Category", "Description", "Amount", "Lead source (if lead gen)", "Month", "Year"], [12, 24, 32, 12, 22, 8, 8])
    style_range(ex, X0, X1, 1, 7, inputs=(1, 2, 3, 4, 5), fmt={1: DATE, 4: MONEY, 6: "0", 7: "0"})
    for r in range(X0, X1 + 1):
        ex[f"F{r}"] = f'=IF(A{r}="","",MONTH(A{r}))'; ex[f"G{r}"] = f'=IF(A{r}="","",YEAR(A{r}))'
    for rng, f1 in ((f"B{X0}:B{X1}", "=Settings!$D$21:$D$34"), (f"E{X0}:E{X1}", f"={SRC}")):
        dv = DataValidation(type="list", formula1=f1, allow_blank=True); ex.add_data_validation(dv); dv.add(rng)
    ex.freeze_panes = f"A{X0}"
    XX = lambda c: f"Expenses!${c}${X0}:${c}${X1}"

    # ---------- sample data ----------
    if sample:
        rnd = random.Random(11)
        names = ["Garcia", "Thompson", "Nakamura", "Okafor", "Bennett", "Rossi", "Kim", "Walsh", "Hughes", "Patel", "Novak", "Ramirez",
                 "Fischer", "Coleman", "Adeyemi", "Lindqvist", "Morales", "Chen", "Doyle", "Sullivan"]
        streets = ["Bayview Dr", "Elm St", "Harbor Ln", "Sunset Blvd", "Oak Ridge Ct", "Pelican Way", "Marina Ave", "Cedar Pl"]
        deals = []
        for k in range(13):
            m = 1 + int(k * 10.5 / 13); d = dt.date(2026, m, 3 + (k * 7) % 24)
            price = rnd.choice([265000, 298000, 342000, 365000, 389000, 415000, 459000, 525000, 610000])
            side = rnd.choice(["Buyer", "Seller", "Seller", "Buyer", "Buyer", "Seller", "Buyer", "Seller", "Buyer", "Dual"])
            rate = 0.025 if side != "Dual" else 0.05
            src = rnd.choice(SOURCES[:7])
            ref = 0.25 if src == "Zillow / portals" else (0.0 if src != "Agent referral" else 0.25)
            deals.append((f"{names[k]} {'family' if k % 3 == 0 else ''}".strip(), f"{100 + 23 * k} {streets[k % 8]}", side, src, "Closed",
                          d - dt.timedelta(days=35), d, price, rate, ref or None, None, None, None))
        deals.append(("Hughes", "77 Harbor Ln", "Seller", "Sphere / past client", "Pending", dt.date(2026, 10, 20), dt.date(2026, 11, 24),
                      498000, 0.025, None, None, None, "Inspection done"))
        deals.append(("Chen", "12 Marina Ave", "Buyer", "Open house", "Under contract", dt.date(2026, 10, 28), dt.date(2026, 12, 5),
                      372000, 0.025, None, None, None, None))
        deals.append(("Doyle", "450 Sunset Blvd", "Seller", "Website / SEO", "Active listing", None, None, 649000, 0.025, None, None, None, "Listed Nov 1"))
        deals.append(("Sullivan", "8 Elm St", "Buyer", "Facebook / Instagram ads", "Fell through", dt.date(2026, 8, 2), None, 310000, 0.025, None, None, None, "Financing denied"))
        deals.sort(key=lambda x: (x[6] or dt.date(2030, 1, 1)))
        for r, dl in enumerate(deals, T0):
            for col, v in zip("ABCDEFGHIJKLM", dl): tr[f"{col}{r}"] = v
        stages = ["New", "Contacted", "Appointment set", "Nurture", "Lost", "Agreement signed", "Closed", "Contacted", "Nurture"]
        first = ["Ava", "Liam", "Mia", "Noah", "Zoe", "Ethan", "Ella", "Lucas", "Ivy", "Owen", "Ruby", "Leo", "Nora", "Jack", "Lily", "Max"]
        today = dt.date(2026, 11, 7)
        for k in range(150):
            r = L0 + k; added = dt.date(2026, 1, 5) + dt.timedelta(days=int(k * 2.08))
            stg = stages[k % len(stages)] if added < dt.date(2026, 10, 1) else rnd.choice(["New", "Contacted", "Appointment set"])
            vals = [added, f"{first[k % 16]} {names[(k * 7) % 20]}", f"(555) 02{k % 10}-{4100 + 13 * k}", f"{first[k % 16].lower()}{k}@mail.example",
                    SOURCES[(k * 3) % 8], rnd.choice(["Buyer", "Buyer", "Seller", "Investor"]), rnd.choice(["$250–350k", "$350–450k", "$450–600k", "$600k+"]),
                    stg, added + dt.timedelta(days=rnd.randint(3, 20)),
                    (today + dt.timedelta(days=rnd.randint(-2, 24))) if stg not in ("Closed", "Lost") else None, None]
            for col, v in zip("ABCDEFGHIJK", vals): ld[f"{col}{r}"] = v
        xs = []
        for m in range(1, 12):
            xs += [(dt.date(2026, m, 1), "Brokerage desk fee", "Monthly desk fee", 150, None),
                   (dt.date(2026, m, 5), "Phone & software", "CRM + phone", 89, None),
                   (dt.date(2026, m, 12), "Lead generation", "Portal leads subscription", 300, "Zillow / portals"),
                   (dt.date(2026, m, 18), "Car & mileage", "Gas and mileage", 160 + 10 * (m % 3), None)]
            if m % 2 == 0: xs.append((dt.date(2026, m, 22), "Marketing & ads", "Facebook / Instagram ads", 250, "Facebook / Instagram ads"))
            if m % 3 == 1: xs.append((dt.date(2026, m, 25), "Photography & staging", "Listing photos", 225, None))
        xs += [(dt.date(2026, 1, 15), "MLS & board dues", "Annual MLS + board dues", 980, None),
               (dt.date(2026, 3, 10), "E&O insurance", "Errors & omissions insurance", 650, None),
               (dt.date(2026, 6, 20), "Client gifts", "Closing gifts", 340, None),
               (dt.date(2026, 9, 9), "Education & licensing", "CE course", 199, None)]
        xs.sort(key=lambda x: x[0])
        for r, x in enumerate(xs, X0):
            for col, v in zip("ABCDE", x): ex[f"{col}{r}"] = v

    # ---------- Monthly ----------
    K.title(mo, "Monthly Summary", "Closed deals by close date, for the year on the Settings tab.")
    mo["A3"] = '="Year "&Settings!B5&"  ·  amounts in "&Settings!B6'; mo["A3"].font = font(bold=True)
    K.hdr(mo, 5, ["Month", "Closings", "Sales volume", "GCI", "Brokerage + fees", "Net commission", "Expenses", "Profit", "Tax set-aside",
                  "Cap paid so far"], [12, 9, 14, 13, 14, 14, 12, 13, 12, 13])
    for m in range(1, 13):
        r = 5 + m; cnd = f'{TT("X")},{m},{TT("Y")},{YR},{TT("E")},"Closed"'
        mo[f"A{r}"] = f'=TEXT(DATE({YR},{m},1),"mmm yyyy")'
        mo[f"B{r}"] = f"=COUNTIFS({cnd})"
        mo[f"C{r}"] = f"=SUMIFS({TT('H')},{cnd})"
        mo[f"D{r}"] = f"=SUMIFS({TT('N')},{cnd})"
        mo[f"F{r}"] = f"=SUMIFS({TT('U')},{cnd})"
        mo[f"E{r}"] = f"=D{r}-F{r}+SUMIFS({TT('L')},{cnd})"
        mo[f"G{r}"] = f"=SUMIFS({XX('D')},{XX('F')},{m},{XX('G')},{YR})"
        mo[f"H{r}"] = f"=F{r}-G{r}"
        mo[f"I{r}"] = f"=MAX(0,H{r})*{TAX}"
        mo[f"J{r}"] = f'=SUMIFS({TT("Q")},{TT("W")},{YR}-IF({m}<{CAPM},1,0),{TT("G")},"<"&DATE({YR},{m}+1,1))'
    style_range(mo, 6, 17, 1, 10, fmt={2: "0", 3: USD0, 4: MONEY, 5: MONEY, 6: MONEY, 7: MONEY, 8: MONEY, 9: MONEY, 10: MONEY})
    mo["A18"] = "TOTAL"
    for c in "BCDEFGHI": mo[f"{c}18"] = f"=SUM({c}6:{c}17)"
    mo["J18"] = "=J17"
    for c, fmt in zip("ABCDEFGHIJ", [None, "0", USD0, MONEY, MONEY, MONEY, MONEY, MONEY, MONEY, MONEY]):
        cell = mo[f"{c}18"]; cell.font = Font(name=F, bold=True, color="FFFFFF"); cell.fill = fill(MAIN); cell.border = B
        if fmt: cell.number_format = fmt
    mo.conditional_formatting.add("H6:H17", FormulaRule(formula=["H6<0"], font=Font(name=F, bold=True, color="B91C1C")))
    mo["A20"] = "Brokerage + fees = GCI minus what you kept (brokerage split, franchise, transaction fee, team split, referral fees)."
    mo["A20"].font = font(italic=True, size=9)

    # ---------- Goal Planner ----------
    K.title(gp, "Goal Planner", "Work backwards from your income goal: how many deals and leads you need.")
    gin = [("Net income goal for the year", 180000, MONEY), ("Average sale price", 425000, USD0), ("Average commission rate", 0.025, "0.00%"),
           ("Lead-to-closing conversion rate", 0.03, PCT), ("Months left to work this year", 12, "0")]
    for i, (lab, v, fmt) in enumerate(gin, 4):
        gp[f"A{i}"] = lab; gp[f"A{i}"].font = font(bold=True); input_cell(gp[f"B{i}"], v); gp[f"B{i}"].number_format = fmt
    gout = [("GCI per average deal", "=B5*B6", MONEY),
            ("Your average keep % (from your closed deals, else your split)", f'=IFERROR(Monthly!F18/Monthly!D18,{SPLIT})', PCT),
            ("Net per average deal", "=B11*B12", MONEY),
            ("Deals needed for the year", "=IFERROR(ROUNDUP(B4/B13,0),0)", "0"),
            ("Deals already closed this year", "=Monthly!B18", "0"),
            ("Deals still needed", "=MAX(0,B14-B15)", "0"),
            ("Leads needed (at your conversion rate)", "=IFERROR(ROUNDUP(B16/B7,0),0)", "#,##0"),
            ("Leads needed per month", "=IFERROR(ROUNDUP(B17/B8,0),0)", "#,##0"),
            ("Leads added this year so far", f'=COUNTIF({LL("M")},{YR})', "#,##0"),
            ("Goal progress (net income)", "=IFERROR(Monthly!F18/B4,0)", PCT)]
    for i, (lab, f, fmt) in enumerate(gout, 11):
        gp[f"A{i}"] = lab; gp[f"A{i}"].font = font(bold=True); c = gp[f"B{i}"]; c.value = f; c.number_format = fmt
        c.border = B; c.fill = fill("F3F4F6"); c.font = font()
    for c in "AB": gp[f"{c}14"].font = Font(name=F, size=13, bold=True, color=MAIN); gp[f"{c}14"].fill = fill(SOFT)
    gp["A22"] = "Tip: your real conversion rate = closed deals ÷ leads on the Dashboard (Lead source table)."
    gp["A22"].font = font(italic=True, size=9)
    gp.column_dimensions["A"].width = 58; gp.column_dimensions["B"].width = 18

    # ---------- Dashboard ----------
    K.title(dash, "Real Estate Commission Dashboard", "Updates by itself from Transactions, Leads and Expenses.")
    dash["A3"] = '=Settings!B4&"  ·  year "&Settings!B5&"  ·  amounts in "&Settings!B6'; dash["A3"].font = font(bold=True)
    kp = [("GCI this year", "=Monthly!D18", MONEY), ("Net commission", "=Monthly!F18", MONEY), ("Profit after expenses", "=Monthly!H18", MONEY),
          ("Closed deals", "=Monthly!B18", "0"), ("Sales volume", "=Monthly!C18", USD0),
          ("Pending GCI (under contract)", f'=SUMIFS({TT("N")},{TT("E")},"Pending")+SUMIFS({TT("N")},{TT("E")},"Under contract")', MONEY)]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 3) * 3, 1 + (i % 3) * 2, lab, f, fmt)
        dash.cell(row=6 + (i // 3) * 3, column=1 + (i % 3) * 2).alignment = Alignment(horizontal="left")
    for c, w in zip("ABCDEFG", [34, 14, 22, 14, 22, 14, 3]): dash.column_dimensions[c].width = w
    K.section(dash, "A11", "Cap & to do")
    todo = [("Brokerage cap paid this cap year", f'=SUMIFS({TT("Q")},{TT("W")},IF(YEAR(TODAY())={YR},YEAR(TODAY())-IF(MONTH(TODAY())<{CAPM},1,0),{YR}))', MONEY),
            ("Cap progress", f'=IF({CAP}<=0,"no cap",MIN(1,B12/{CAP}))', PCT),
            ("Left until you cap", f'=IF({CAP}<=0,"no cap",MAX(0,{CAP}-B12))', MONEY),
            ("Average GCI per closing", "=IFERROR(Monthly!D18/Monthly!B18,0)", MONEY),
            ("Lead follow-ups overdue", f'=COUNTIF({LL("L")},"Overdue")', "0"),
            ("Follow-ups today / soon", f'=COUNTIF({LL("L")},"Today")+COUNTIF({LL("L")},"Soon")', "0"),
            ("Tax set-aside this year", "=Monthly!I18", MONEY),
            ("Goal progress (Goal Planner)", "='Goal Planner'!B20", PCT)]
    for i, (lab, f, fmt) in enumerate(todo, 12):
        dash[f"A{i}"] = lab; dash[f"A{i}"].font = font(bold=True); dash[f"A{i}"].border = B; dash[f"A{i}"].fill = fill(SOFT)
        c = dash[f"B{i}"]; c.value = f; c.number_format = fmt; c.border = B; c.font = Font(name=F, bold=True, color=MAIN)
    dash.conditional_formatting.add("B16", FormulaRule(formula=["B16>0"], font=Font(name=F, bold=True, color="B91C1C")))
    dash.conditional_formatting.add("B13", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="F59E0B"))
    SR = 22
    K.section(dash, f"A{SR}", "Lead source ROI (this year)")
    K.hdr(dash, SR + 1, ["Lead source", "Leads", "Closings", "Conversion", "GCI", "Spend", "GCI per $1 spent"])
    for i in range(NS):
        r = SR + 2 + i; k = S0 + i
        dash[f"A{r}"] = f'=IF(Settings!A{k}="","",Settings!A{k})'
        dash[f"B{r}"] = f'=IF(A{r}="","",COUNTIFS({LL("E")},A{r},{LL("M")},{YR}))'
        dash[f"C{r}"] = f'=IF(A{r}="","",COUNTIFS({TT("D")},A{r},{TT("E")},"Closed",{TT("Y")},{YR}))'
        dash[f"D{r}"] = f'=IF(OR(A{r}="",N(B{r})=0),"",C{r}/B{r})'
        dash[f"E{r}"] = f'=IF(A{r}="","",SUMIFS({TT("N")},{TT("D")},A{r},{TT("E")},"Closed",{TT("Y")},{YR}))'
        dash[f"F{r}"] = f'=IF(A{r}="","",SUMIFS({XX("D")},{XX("E")},A{r},{XX("G")},{YR}))'
        dash[f"G{r}"] = f'=IF(OR(A{r}="",N(F{r})=0),"",E{r}/F{r})'
    style_range(dash, SR + 2, SR + 1 + NS, 1, 7, fmt={2: "0", 3: "0", 4: PCT, 5: MONEY, 6: MONEY, 7: '0.0"x"'})
    dash.column_dimensions["G"].width = 14
    TR = SR + NS + 3
    K.section(dash, f"A{TR}", "Top 5 deals by net commission")
    K.hdr(dash, TR + 1, ["Client", "Close date", "Sale price", "GCI", "Net to you"])
    RK = TT("AA"); CN = TT("A")
    for k in range(5):
        r = TR + 2 + k
        dash[f"A{r}"] = f'=IFERROR(INDEX({CN},MATCH(LARGE({RK},{k + 1}),{RK},0)),"")'
        for col, src in zip("BCDE", "GHNU"):
            dash[f"{col}{r}"] = f'=IF(A{r}="","",INDEX({TT(src)},MATCH(LARGE({RK},{k + 1}),{RK},0)))'
    style_range(dash, TR + 2, TR + 6, 1, 5, fmt={2: DATE, 3: USD0, 4: MONEY, 5: MONEY})
    ch = BarChart(); ch.type = "col"; ch.title = "GCI and net commission by month"; ch.height = 8; ch.width = 17
    ch.add_data(Reference(mo, min_col=4, min_row=5, max_row=17), titles_from_data=True)
    ch.add_data(Reference(mo, min_col=6, min_row=5, max_row=17), titles_from_data=True)
    ch.set_categories(Reference(mo, min_col=1, min_row=6, max_row=17)); ch.y_axis.majorGridlines = None
    ch.series[0].graphicalProperties.solidFill = MAIN; ch.series[1].graphicalProperties.solidFill = "F59E0B"
    dash.add_chart(ch, "I5")
    ch2 = BarChart(); ch2.type = "bar"; ch2.title = "GCI by lead source"; ch2.height = 8; ch2.width = 17
    ch2.add_data(Reference(dash, min_col=5, min_row=SR + 1, max_row=SR + 1 + len(SOURCES)), titles_from_data=True)
    ch2.set_categories(Reference(dash, min_col=1, min_row=SR + 2, max_row=SR + 1 + len(SOURCES))); ch2.legend = None
    ch2.x_axis.scaling.orientation = "maxMin"; ch2.y_axis.majorGridlines = None; ch2.series[0].graphicalProperties.solidFill = MAIN
    dash.add_chart(ch2, "I22")

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Settings tab: your name, the year, your brokerage split (e.g. 70%), annual cap, franchise fee, transaction fee and tax set-aside %.", False),
             ("2.  Transactions tab: one row per deal – client, side, lead source, status, close date, sale price and commission %. Add referral % or team split % if any.", False),
             ("3.  When a deal closes, set Status to Closed. GCI, brokerage split (stops when you hit your cap), fees, net to you and tax set-aside calculate themselves.", False),
             ("4.  Leads tab: every new lead with source, stage and next follow-up date. Overdue follow-ups turn red.", False),
             ("5.  Expenses tab: dues, desk fees, ads, lead generation… Tag lead-gen costs with a lead source to see your ROI per source.", False),
             ("6.  Goal Planner: type your income goal and see how many deals and leads you need.", False),
             ("7.  Monthly and Dashboard: GCI, net, profit, cap progress, lead source ROI, top deals and charts – all automatic.", False),
             ("", False), ("Tips", True),
             ("•  Leave Commission % empty to use the default rate on Settings. Use 5% (or both sides) for a dual-agency deal.", False),
             ("•  The cap counts closings by close date inside your cap year (set the start month on Settings if your cap year isn't January).", False),
             ("•  On a team? Put the team lead's share in 'Team lead split %'. Paying a referral fee? Put it in 'Referral fee paid %'.", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  200 transactions, 500 leads, 300 expenses, 12 lead sources. No macros. Works in Excel 2010+ and Google Sheets.", False),
             ("•  To use Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False),
             ("•  This is a planning tool, not tax advice. Check your brokerage agreement for exact split and cap rules.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("The agent, clients, deals and leads are made-up examples. Use the blank copy in your download for your own business.", False)]
    K.start(st, "Real Estate Agent Commission & Lead Tracker", "GCI, brokerage split with cap, net income, leads and follow-ups – Excel and Google Sheets.", lines)
    finish(wb, sample, path)


META = {
    "name": "Real Estate Agent Commission & Lead Tracker (GCI, Split & Cap, Net Income, Lead Follow-up, Goal Planner) – Excel & Google Sheets",
    "file_base": "Real-Estate-Agent-Commission-Tracker",
    "price": 19,
    "color": "#7C2D12",
    "title": "Real Estate Commission Tracker",
    "tagline": "GCI · split & cap · net income · leads · follow-ups · goal planner",
    "intro": "Know exactly what you earn from every deal. Log each transaction once and this spreadsheet works out your GCI, the brokerage split (it stops charging when you hit your annual cap), franchise, referral, team and transaction fees, your net commission and a tax set-aside. Track every lead with a follow-up date, see which lead sources actually pay off, and plan how many deals you need to hit your income goal.",
    "features": [
        "Transaction log for 200 deals: buyer/seller/dual side, status, sale price, commission %, referral fee and team split",
        "Automatic GCI, brokerage split with annual cap, franchise fee (with its own cap), transaction fee, net to you and tax set-aside",
        "Lead tracker for 500 leads: source, stage, next follow-up date with Overdue / Today / Soon reminders",
        "Lead source ROI: leads, closings, conversion rate, GCI and spend per source",
        "Goal planner: income goal to deals needed and leads needed per month",
        "Expenses tracker and monthly summary: GCI, net, expenses, profit, cap paid so far",
        "Dashboard with cap progress, pending GCI, top deals and charts; blank copy plus a sample copy",
    ],
    "tags": ["real estate agent", "commission tracker", "realtor spreadsheet", "GCI tracker", "real estate business plan", "lead tracker",
             "realtor commission", "google sheets", "excel template", "real estate planner", "brokerage cap", "realtor business", "sales tracker"],
    "summary": "Instant download: Excel + Google Sheets template, blank copy + sample copy.",
}
