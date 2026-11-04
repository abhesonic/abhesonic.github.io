"""Travel Planner & Trip Budget – itinerary, multi-currency expenses, group split, packing list.

Settings (trip, dates, travellers, home currency, budget per category, exchange rates) -> Itinerary (day-by-day plan & bookings)
-> Expenses (any currency, converted to home currency, paid by / shared or personal, planned vs paid) -> Packing list
-> Dashboard (countdown, budget used, cost per day & per person, category table + chart, who owes whom).
"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import BarChart, Reference
from _common import Kit, F, B, MONEY, DATE, font, fill, style_range, input_cell, finish

MAIN = "0F766E"; SOFT = "CCFBF1"
K = Kit(MAIN, SOFT)
PCT = "0.0%"; SIGNED = '#,##0.00;[Red]-#,##0.00'
CATS = ["Flights", "Accommodation", "Local transport", "Car rental & fuel", "Food & drinks", "Activities & tours",
        "Shopping & gifts", "Insurance, visas & fees", "Phone & internet", "Tips", "Other"]
NC = len(CATS); C0 = 26; C1 = C0 + NC - 1           # Settings category rows
NP = 8; P0 = 15; P1 = P0 + NP - 1                    # travellers (Settings A15:A22)
NX = 12; X0 = 15; X1 = X0 + NX - 1                   # currencies (Settings D15:E26)
NI = 150; I0 = 6; I1 = I0 + NI - 1                   # itinerary
NE = 250; E0 = 6; E1 = E0 + NE - 1                   # expenses
NK = 120; K0 = 6; K1 = K0 + NK - 1                   # packing
TYPES = ["Flight", "Train / bus", "Hotel check-in", "Hotel check-out", "Sightseeing", "Tour / activity", "Food", "Transport", "Free time", "Other"]
STATUS = ["Paid", "Booked – pay later", "Planned"]
BAR = 20

PACK = [("Passport / ID", "Documents"), ("Visa / entry documents", "Documents"), ("Travel insurance details", "Documents"),
        ("Flight & hotel confirmations", "Documents"), ("Driver's licence", "Documents"), ("Credit / debit cards", "Money"),
        ("Some local cash", "Money"), ("Phone + charger", "Electronics"), ("Power bank", "Electronics"), ("Plug adapter", "Electronics"),
        ("Headphones", "Electronics"), ("Camera", "Electronics"), ("Underwear & socks", "Clothes"), ("T-shirts / tops", "Clothes"),
        ("Trousers / shorts", "Clothes"), ("Light jacket / sweater", "Clothes"), ("Pyjamas", "Clothes"), ("Comfortable walking shoes", "Clothes"),
        ("Swimwear", "Clothes"), ("Toothbrush & toothpaste", "Toiletries"), ("Shampoo & soap", "Toiletries"), ("Deodorant", "Toiletries"),
        ("Sunscreen", "Toiletries"), ("Medicines & prescriptions", "Health"), ("First-aid basics", "Health"), ("Sunglasses", "Accessories"),
        ("Reusable water bottle", "Accessories"), ("Day bag / backpack", "Accessories"), ("Travel pillow", "Accessories"),
        ("Snacks for the journey", "Other"), ("Copy of documents (photo / cloud)", "Documents"), ("House keys", "Other")]


def bar(ref):
    return f'REPT("█",ROUND(MIN(1,MAX(0,N({ref})))*{BAR},0))&REPT("░",{BAR}-ROUND(MIN(1,MAX(0,N({ref})))*{BAR},0))'


ITIN = [  # day offset, time, plan, place, type, ref, booked
    (0, "07:40", "Flight New York JFK → Lisbon (TP 210)", "JFK Terminal 1", "Flight", "TP7XQ2", "Yes"),
    (0, "20:30", "Check in – Casa Alfama apartment", "Rua dos Remédios 12, Lisbon", "Hotel check-in", "BK-558201", "Yes"),
    (1, "10:00", "Walking tour of Alfama & Castelo de São Jorge", "Meet at Largo do Chafariz", "Tour / activity", "GYG-88412", "Yes"),
    (1, "19:30", "Dinner + fado show", "Clube de Fado", "Food", "", "No"),
    (2, "09:30", "Day trip to Sintra (Pena Palace)", "Train from Rossio station", "Sightseeing", "PENA-2291", "Yes"),
    (3, "11:00", "Belém: tower, monastery, pastéis de nata", "Belém", "Sightseeing", "", "No"),
    (4, "10:00", "Check out, train to Porto", "Santa Apolónia station", "Train / bus", "CP-1043", "Yes"),
    (4, "14:00", "Check in – Ribeira Guesthouse", "Cais da Ribeira, Porto", "Hotel check-in", "BK-771034", "Yes"),
    (5, "15:00", "Port wine cellar tour", "Vila Nova de Gaia", "Tour / activity", "", "No"),
    (6, "09:00", "Douro Valley boat & winery day", "Pick-up at hotel", "Tour / activity", "DV-5520", "Yes"),
    (7, "Free", "Beach day – Matosinhos", "Matosinhos", "Free time", "", ""),
    (8, "11:00", "Check out, flight Porto → New York", "Porto airport", "Flight", "TP9KL4", "Yes")]

EXP = [  # day offset, category, description, amount, currency, paid by, split, status
    (-60, "Flights", "Return flights x2", 1180, "USD", "Sam", "Shared", "Paid"),
    (-45, "Accommodation", "Lisbon apartment, 4 nights", 520, "EUR", "Alex", "Shared", "Paid"),
    (-45, "Accommodation", "Porto guesthouse, 4 nights", 440, "EUR", "Sam", "Shared", "Booked – pay later"),
    (-30, "Insurance, visas & fees", "Travel insurance", 96, "USD", "Alex", "Shared", "Paid"),
    (-20, "Activities & tours", "Alfama walking tour x2", 50, "EUR", "Alex", "Shared", "Paid"),
    (-20, "Activities & tours", "Pena Palace tickets x2", 40, "EUR", "Sam", "Shared", "Paid"),
    (-20, "Activities & tours", "Douro Valley day trip x2", 190, "EUR", "Alex", "Shared", "Paid"),
    (-5, "Phone & internet", "eSIM 10 GB", 15, "USD", "Sam", "Personal", "Paid"),
    (0, "Local transport", "Airport taxi", 22, "EUR", "Alex", "Shared", "Paid"),
    (0, "Food & drinks", "Groceries for apartment", 38, "EUR", "Sam", "Shared", "Paid"),
    (1, "Food & drinks", "Dinner + fado show", 96, "EUR", "Alex", "Shared", "Paid"),
    (1, "Local transport", "Viva Viagem metro cards", 20, "EUR", "Sam", "Shared", "Paid"),
    (2, "Local transport", "Train Lisbon–Sintra return x2", 9.2, "EUR", "Alex", "Shared", "Paid"),
    (2, "Food & drinks", "Lunch in Sintra", 34, "EUR", "Sam", "Shared", "Paid"),
    (3, "Food & drinks", "Pastéis de Belém + coffee", 12, "EUR", "Alex", "Shared", "Paid"),
    (3, "Shopping & gifts", "Tiles for mum", 45, "EUR", "Alex", "Personal", "Paid"),
    (4, "Local transport", "Train Lisbon → Porto x2", 62, "EUR", "Sam", "Shared", "Paid"),
    (5, "Activities & tours", "Port cellar tour", 40, "EUR", "Sam", "Shared", "Planned"),
    (5, "Food & drinks", "Dinner at Ribeira", 58, "EUR", "Alex", "Shared", "Planned"),
    (7, "Food & drinks", "Beach lunch", 30, "EUR", "Sam", "Shared", "Planned"),
    (8, "Tips", "Tips for guides & drivers", 25, "EUR", "Alex", "Shared", "Planned")]


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); s = wb.create_sheet("Settings"); it = wb.create_sheet("Itinerary")
    ex = wb.create_sheet("Expenses"); pk = wb.create_sheet("Packing List")
    d0 = dt.date(2026, 12, 18)

    # ---------- Settings ----------
    K.title(s, "Settings", "Yellow cells are yours to edit.")
    rows = [("Trip name", "Portugal – Lisbon & Porto" if sample else "My Trip", None),
            ("First day of trip", d0 if sample else "=TODAY()+60", DATE),
            ("Last day of trip", d0 + dt.timedelta(days=8) if sample else "=TODAY()+67", DATE),
            ("Home currency (code)", "USD", None),
            ("Total trip budget (home currency)", 5200 if sample else 3000, MONEY),
            ("Today (for countdown)", "=TODAY()", DATE),
            ("Number of travellers", f'=MAX(1,COUNTA(A{P0}:A{P1}))', "0")]
    for i, (k_, v, fmt) in enumerate(rows, 4):
        s.cell(row=i, column=1, value=k_).font = font(bold=True)
        c = s.cell(row=i, column=2)
        if i in (9, 10):
            c.value = v; c.border = B; c.fill = fill("F3F4F6"); c.font = font()
        else:
            input_cell(c, v)
        if fmt: c.number_format = fmt
    notes = {7: "All totals on the Dashboard are shown in this currency.", 8: "Your spending limit for the whole trip, all travellers together.",
             9: "Calculated – nothing to type.", 10: "Counted from the names below."}
    for r, t in notes.items(): s[f"C{r}"] = t; s[f"C{r}"].font = font(italic=True, size=9)

    K.hdr(s, 14, ["Travellers (names)"]); s["A14"].value = "Travellers (names)"
    for i in range(NP):
        input_cell(s[f"A{P0 + i}"], (["Alex", "Sam"][i] if sample and i < 2 else None))
    s[f"A{P1 + 1}"] = "Up to 8 people. Used for 'Paid by' and the who-owes-whom table."; s[f"A{P1 + 1}"].font = font(italic=True, size=9)

    K.hdr(s, 14, ["Currency code", "1 unit = ? home currency"], col0=4)
    s["D14"].value = "Currency code"; s["E14"].value = "1 unit = ? home currency"
    rates = [("USD", 1), ("EUR", 1.08), ("GBP", 1.27)] if sample else [("USD", 1)]
    for i in range(NX):
        r = X0 + i
        code, rate = rates[i] if i < len(rates) else (None, None)
        input_cell(s[f"D{r}"], code); input_cell(s[f"E{r}"], rate); s[f"E{r}"].number_format = "0.0000"
    s[f"D{X1 + 1}"] = "Put your home currency first with rate 1. Update rates from xe.com or your bank."
    s[f"D{X1 + 1}"].font = font(italic=True, size=9)
    s[f"D{X1 + 2}"] = "Example: home = USD, 1 EUR = 1.08 USD → type EUR and 1.08."; s[f"D{X1 + 2}"].font = font(italic=True, size=9)

    K.hdr(s, C0 - 1, ["Budget category", "Planned budget (home currency)", "Share of total"])
    shares = [0.24, 0.20, 0.05, 0.03, 0.18, 0.12, 0.05, 0.04, 0.01, 0.02, 0.06]
    for i, cat in enumerate(CATS):
        r = C0 + i
        input_cell(s[f"A{r}"], cat); input_cell(s[f"B{r}"], round(5200 * shares[i], -1) if sample else None)
        s[f"B{r}"].number_format = MONEY
        c = s[f"C{r}"]; c.value = f"=IFERROR(B{r}/$B$8,0)"; c.number_format = PCT; c.border = B; c.fill = fill("F3F4F6"); c.font = font()
    s[f"A{C1 + 1}"] = "Total of categories"; s[f"A{C1 + 1}"].font = font(bold=True)
    s[f"B{C1 + 1}"] = f"=SUM(B{C0}:B{C1})"; s[f"B{C1 + 1}"].number_format = MONEY; s[f"B{C1 + 1}"].font = font(bold=True)
    s[f"C{C1 + 1}"] = (f'=IF(ABS(B{C1 + 1}-B8)<0.5,"✓ matches total budget",IF(B{C1 + 1}>B8,"Over total budget by "&TEXT(B{C1 + 1}-B8,"#,##0"),'
                       f'"Not yet planned: "&TEXT(B8-B{C1 + 1},"#,##0")))')
    s[f"C{C1 + 1}"].font = font(italic=True, size=9)
    for c, w in zip("ABCDE", [34, 22, 24, 16, 24]): s.column_dimensions[c].width = w
    cats = f"Settings!$A${C0}:$A${C1}"; people = f"Settings!$A${P0}:$A${P1}"
    codes = f"Settings!$D${X0}:$D${X1}"; rts = f"Settings!$E${X0}:$E${X1}"

    # ---------- Itinerary ----------
    K.title(it, "Itinerary", "Your day-by-day plan: travel, stays, tours and reservations. Sort by date any time.")
    K.hdr(it, 5, ["Date", "Day #", "Weekday", "Time", "Plan / activity", "Place / address", "Type", "Booking ref", "Booked?", "Notes"],
          [13, 8, 11, 9, 44, 30, 16, 14, 10, 26])
    style_range(it, I0, I1, 1, 10, inputs=(1, 4, 5, 6, 7, 8, 9, 10), fmt={1: DATE, 2: "0"})
    for r in range(I0, I1 + 1):
        it[f"B{r}"] = f'=IF(A{r}="","",A{r}-Settings!$B$5+1)'
        it[f"C{r}"] = f'=IF(A{r}="","",TEXT(A{r},"ddd"))'
    if sample:
        for i, (off, tm, plan, place, typ, ref, bk) in enumerate(ITIN):
            r = I0 + i; it[f"A{r}"] = d0 + dt.timedelta(days=off); it[f"D{r}"] = tm; it[f"E{r}"] = plan
            it[f"F{r}"] = place; it[f"G{r}"] = typ; it[f"H{r}"] = ref or None; it[f"I{r}"] = bk or None
    it["L5"] = "Types"; it["L5"].font = font(bold=True)
    for i, t in enumerate(TYPES): it[f"L{6 + i}"] = t; it[f"L{6 + i}"].font = font(size=9)
    it.column_dimensions["L"].width = 18
    dv = DataValidation(type="list", formula1=f"=$L$6:$L${5 + len(TYPES)}", allow_blank=True); it.add_data_validation(dv); dv.add(f"G{I0}:G{I1}")
    dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True); it.add_data_validation(dv); dv.add(f"I{I0}:I{I1}")
    it.conditional_formatting.add(f"I{I0}:I{I1}", FormulaRule(formula=[f'I{I0}="No"'], font=Font(name=F, bold=True, color="B91C1C")))
    it.conditional_formatting.add(f"A{I0}:J{I1}", FormulaRule(formula=[f'AND($A{I0}<>"",MOD($B{I0},2)=0)'], fill=fill("ECFDF5")))
    it["A3"] = "Day # counts from the first day on Settings. Booked? = No shows in red so you don't forget to reserve."; it["A3"].font = font(italic=True, size=9)
    it.freeze_panes = f"E{I0}"

    # ---------- Expenses ----------
    K.title(ex, "Expenses", "Every cost, in any currency. It is converted to your home currency with the rates on Settings.")
    K.hdr(ex, 5, ["Date", "Category", "Description", "Amount", "Currency", "Rate", "Amount (home)", "Paid by", "Shared or personal", "Status", "Notes"],
          [13, 22, 32, 12, 10, 9, 15, 13, 14, 19, 22])
    style_range(ex, E0, E1, 1, 11, inputs=(1, 2, 3, 4, 5, 8, 9, 10, 11), fmt={1: DATE, 4: MONEY, 6: "0.0000", 7: MONEY})
    for r in range(E0, E1 + 1):
        ex[f"F{r}"] = f'=IF(D{r}="","",IF(OR(E{r}="",E{r}=Settings!$B$7),1,IFERROR(INDEX({rts},MATCH(E{r},{codes},0)),"rate?")))'
        ex[f"G{r}"] = f'=IF(D{r}="","",IF(ISNUMBER(F{r}),D{r}*F{r},0))'
    if sample:
        for i, (off, cat, desc, amt, cur, who, sp, stt) in enumerate(EXP):
            r = E0 + i
            for col, v in zip("ABCDEHIJ", (d0 + dt.timedelta(days=off), cat, desc, amt, cur, who, sp, stt)): ex[f"{col}{r}"] = v
    for rng, f1 in ((f"B{E0}:B{E1}", f"={cats}"), (f"H{E0}:H{E1}", f"={people}"), (f"E{E0}:E{E1}", f"={codes}"),
                    (f"I{E0}:I{E1}", '"Shared,Personal"'), (f"J{E0}:J{E1}", '"' + ",".join(STATUS) + '"')):
        dv = DataValidation(type="list", formula1=f1, allow_blank=True); ex.add_data_validation(dv); dv.add(rng)
    ex.conditional_formatting.add(f"F{E0}:F{E1}", FormulaRule(formula=[f'F{E0}="rate?"'], font=Font(name=F, bold=True, color="B91C1C")))
    ex.conditional_formatting.add(f"J{E0}:J{E1}", FormulaRule(formula=[f'J{E0}="Planned"'], font=Font(name=F, italic=True, color="9CA3AF")))
    ex.conditional_formatting.add(f"J{E0}:J{E1}", FormulaRule(formula=[f'J{E0}="Booked – pay later"'], font=Font(name=F, bold=True, color="B45309")))
    ex["A3"] = ("Status: Paid = money spent · Booked – pay later = reserved, pay at the hotel/desk · Planned = an estimate. "
                "Empty currency = home currency. 'rate?' in red = add that currency on Settings.")
    ex["A3"].font = font(italic=True, size=9)
    tr = E1 + 1; ex[f"A{tr}"] = "TOTAL"; ex[f"G{tr}"] = f"=SUM(G{E0}:G{E1})"; ex[f"G{tr}"].number_format = MONEY
    for c in "ABCDEFGHIJK": ex[f"{c}{tr}"].font = font(bold=True); ex[f"{c}{tr}"].fill = fill(SOFT); ex[f"{c}{tr}"].border = B
    ex.freeze_panes = f"D{E0}"

    # ---------- Packing ----------
    K.title(pk, "Packing list", "A ready-made list – delete what you don't need, add your own items.")
    K.hdr(pk, 5, ["Item", "Category", "For whom", "Packed?", "Notes"], [36, 16, 14, 10, 28])
    style_range(pk, K0, K1, 1, 5, inputs=(1, 2, 3, 4, 5))
    for i, (item, cat) in enumerate(PACK):
        r = K0 + i; pk[f"A{r}"] = item; pk[f"B{r}"] = cat
        if sample: pk[f"D{r}"] = "Yes" if i % 3 != 2 else "No"; pk[f"C{r}"] = "Everyone" if cat == "Documents" else None
    dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True); pk.add_data_validation(dv); dv.add(f"D{K0}:D{K1}")
    pk.conditional_formatting.add(f"A{K0}:E{K1}", FormulaRule(formula=[f'$D{K0}="Yes"'], font=Font(name=F, color="9CA3AF", strike=True)))
    pk.freeze_panes = f"A{K0}"

    # ---------- Dashboard ----------
    K.title(dash, "Trip Dashboard", "Updates by itself from the other tabs. All amounts in your home currency.")
    dash["A3"] = ('=Settings!B4&"  ·  "&TEXT(Settings!B5,"d mmm yyyy")&" – "&TEXT(Settings!B6,"d mmm yyyy")&"  ·  "&Settings!B10&" traveller(s)  ·  "&Settings!B7')
    dash["A3"].font = font(bold=True)
    G = f"Expenses!$G${E0}:$G${E1}"; J = f"Expenses!$J${E0}:$J${E1}"; H = f"Expenses!$H${E0}:$H${E1}"; Isp = f"Expenses!$I${E0}:$I${E1}"
    total = f"SUM({G})"; spent = f'SUMIFS({G},{J},"Paid")'
    kp = [("Days to go", '=IF(Settings!B5-Settings!B9>0,Settings!B5-Settings!B9,IF(Settings!B6>=Settings!B9,"On the trip!","Home again"))', "0"),
          ("Trip length (days)", "=Settings!B6-Settings!B5+1", "0"),
          ("Total budget", "=Settings!B8", MONEY),
          ("Total trip cost", f"={total}", MONEY),
          ("Paid so far", f"={spent}", MONEY),
          ("Left in budget", f"=Settings!B8-{total}", SIGNED)]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 3) * 3, 1 + (i % 3) * 2, lab, f, fmt)
        dash.cell(row=6 + (i // 3) * 3, column=1 + (i % 3) * 2).alignment = Alignment(horizontal="left")
    for c, w in zip("ABCDEFG", [40, 16, 15, 15, 15, 15, 3]): dash.column_dimensions[c].width = w

    def block(r0, title, rows_):
        K.section(dash, f"A{r0}", title)
        for i, (lab, f, fmt) in enumerate(rows_, r0 + 1):
            dash[f"A{i}"] = lab; dash[f"A{i}"].font = font(bold=True); dash[f"A{i}"].border = B; dash[f"A{i}"].fill = fill(SOFT)
            c = dash[f"B{i}"]; c.value = f; c.number_format = fmt; c.border = B; c.font = Font(name=F, bold=True, color=MAIN)
            c.alignment = Alignment(horizontal="right")
    block(11, "Money", [
        ("Budget used (total cost ÷ budget)", f"=IFERROR({total}/Settings!B8,0)", PCT),
        ("Booked – still to pay", f'=SUMIFS({G},{J},"Booked – pay later")', MONEY),
        ("Planned (estimates)", f'=SUMIFS({G},{J},"Planned")', MONEY),
        ("Cost per day", "=IFERROR(A9/C6,0)", MONEY),
        ("Cost per person", "=IFERROR(A9/Settings!B10,0)", MONEY),
        ("Budget left per day (from today)", f'=IFERROR(E9/MAX(1,MIN(C6,Settings!B6-MAX(Settings!B9,Settings!B5)+1)),0)', SIGNED),
        ("Expenses missing an exchange rate", f'=COUNTIF(Expenses!$F${E0}:$F${E1},"rate?")', "0")])
    dash["C12"] = f'={bar("B12")}'; dash["C12"].font = Font(name=F, color=MAIN)
    block(20, "Plans & packing", [
        ("Itinerary items", f'=COUNTA(Itinerary!$E${I0}:$E${I1})', "0"),
        ("Not booked yet", f'=COUNTIF(Itinerary!$I${I0}:$I${I1},"No")', "0"),
        ("Packing list done", f'=IFERROR(COUNTIF(\'Packing List\'!$D${K0}:$D${K1},"Yes")/COUNTA(\'Packing List\'!$A${K0}:$A${K1}),0)', PCT),
        ("Items still to pack", f'=COUNTA(\'Packing List\'!$A${K0}:$A${K1})-COUNTIF(\'Packing List\'!$D${K0}:$D${K1},"Yes")', "0")])
    dash["C23"] = f'={bar("B23")}'; dash["C23"].font = Font(name=F, color=MAIN)

    # who owes whom
    WR = 27
    K.section(dash, f"A{WR}", "Who owes whom (shared costs split equally, status Paid only)")
    K.hdr(dash, WR + 1, ["Traveller", "Paid", "Fair share", "Balance", "What to do"])
    shared = f'SUMIFS({G},{J},"Paid",{Isp},"<>Personal")'
    for i in range(NP):
        r = WR + 2 + i; k = P0 + i
        dash[f"A{r}"] = f'=IF(Settings!A{k}="","",Settings!A{k})'
        dash[f"B{r}"] = f'=IF(A{r}="","",SUMIFS({G},{H},A{r},{J},"Paid"))'
        dash[f"C{r}"] = f'=IF(A{r}="","",{shared}/Settings!$B$10+SUMIFS({G},{H},A{r},{J},"Paid",{Isp},"Personal"))'
        dash[f"D{r}"] = f'=IF(A{r}="","",B{r}-C{r})'
        dash[f"E{r}"] = f'=IF(A{r}="","",IF(ABS(D{r})<0.01,"All square",IF(D{r}>0,"Gets back "&TEXT(D{r},"#,##0.00"),"Pays in "&TEXT(-D{r},"#,##0.00"))))'
    WE = WR + 1 + NP
    style_range(dash, WR + 2, WE, 1, 5, fmt={2: MONEY, 3: MONEY, 4: SIGNED})
    dash[f"A{WE + 1}"] = "Costs paid without a name in 'Paid by' are not counted here. Personal costs only count for the person who had them."
    dash[f"A{WE + 1}"].font = font(italic=True, size=9)

    CR = WE + 3
    K.section(dash, f"A{CR}", "Budget by category")
    K.hdr(dash, CR + 1, ["Category", "Planned budget", "Total cost", "Paid", "Left", "Used"])
    Bc = f"Expenses!$B${E0}:$B${E1}"
    for i in range(NC):
        r = CR + 2 + i; k = C0 + i
        dash[f"A{r}"] = f'=IF(Settings!A{k}="","",Settings!A{k})'
        dash[f"B{r}"] = f"=N(Settings!B{k})"
        dash[f"C{r}"] = f"=SUMIFS({G},{Bc},Settings!A{k})"
        dash[f"D{r}"] = f'=SUMIFS({G},{Bc},Settings!A{k},{J},"Paid")'
        dash[f"E{r}"] = f"=B{r}-C{r}"
        dash[f"F{r}"] = f'=IF(B{r}=0,"",C{r}/B{r})'
    CE = CR + 1 + NC
    style_range(dash, CR + 2, CE, 1, 6, fmt={2: MONEY, 3: MONEY, 4: MONEY, 5: SIGNED, 6: PCT})
    dash[f"A{CE + 1}"] = "Total"
    for c in "BCDE": dash[f"{c}{CE + 1}"] = f"=SUM({c}{CR + 2}:{c}{CE})"; dash[f"{c}{CE + 1}"].number_format = MONEY if c != "E" else SIGNED
    for c in "ABCDEF": dash[f"{c}{CE + 1}"].font = font(bold=True); dash[f"{c}{CE + 1}"].fill = fill(SOFT); dash[f"{c}{CE + 1}"].border = B
    dash.conditional_formatting.add(f"F{CR + 2}:F{CE}", FormulaRule(formula=[f'AND(ISNUMBER(F{CR + 2}),F{CR + 2}>1)'], font=Font(name=F, bold=True, color="B91C1C")))

    ch = BarChart(); ch.type = "bar"; ch.grouping = "clustered"; ch.title = "Budget vs cost by category"; ch.height = 11; ch.width = 16
    ch.add_data(Reference(dash, min_col=2, min_row=CR + 1, max_row=CE), titles_from_data=True)
    ch.add_data(Reference(dash, min_col=3, min_row=CR + 1, max_row=CE), titles_from_data=True)
    ch.set_categories(Reference(dash, min_col=1, min_row=CR + 2, max_row=CE)); ch.y_axis.majorGridlines = None
    ch.x_axis.scaling.orientation = "maxMin"; ch.legend.position = "b"; ch.x_axis.tickLblSkip = 1
    for ser, col in zip(ch.series, ("99F6E4", MAIN)): ser.graphicalProperties.solidFill = col
    dash.add_chart(ch, "H5")

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Settings tab: trip name, first and last day, your home currency, total budget and the names of everyone travelling. Add the exchange rates for the currencies you will spend.", False),
             ("2.  Split the budget over the categories (flights, stays, food…). Rename categories if you like.", False),
             ("3.  Itinerary tab: one row per plan – flights, check-ins, tours, restaurants. Day # and weekday fill in by themselves; 'Booked? No' shows in red.", False),
             ("4.  Expenses tab: every cost in the currency you paid. Pick who paid, Shared or Personal, and the status (Paid / Booked – pay later / Planned).", False),
             ("5.  Packing List: tick Yes as you pack – packed items are crossed out.", False),
             ("6.  Dashboard: countdown, budget used, cost per day and per person, budget by category and who owes whom at the end of the trip.", False),
             ("", False), ("Tips", True),
             ("•  Before the trip, add your estimates as 'Planned' – you see the full trip cost before you book anything.", False),
             ("•  Travelling with friends? Mark group costs as Shared. The Dashboard tells each person how much to pay in or get back.", False),
             ("•  Update exchange rates before you leave; every amount is recalculated automatically.", False),
             ("•  Google Sheets on your phone: open the file in the Sheets app and add expenses on the go.", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  150 itinerary rows, 250 expenses, 12 currencies, 8 travellers, 120 packing items. No macros.", False),
             ("•  To use Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("The Portugal trip, bookings and amounts are made-up examples. Use the blank copy in your download for your own trip.", False)]
    K.start(st, "Travel Planner & Trip Budget", "Itinerary, multi-currency expenses, group cost split and packing list – Excel and Google Sheets.", lines)
    finish(wb, sample, path)


META = {
    "name": "Travel Planner & Trip Budget (Itinerary, Multi-Currency Expenses, Group Split, Packing List) – Excel & Google Sheets",
    "file_base": "Travel-Planner-Trip-Budget",
    "price": 15,
    "color": "#0F766E",
    "title": "Travel Planner & Trip Budget",
    "tagline": "Itinerary · multi-currency expenses · who owes whom · packing list",
    "intro": "Plan your whole trip in one file: a day-by-day itinerary, every booking, a budget by category and expenses in any currency, converted to your home currency automatically. Travelling with friends or family? It splits shared costs and tells everyone how much to pay in or get back.",
    "features": [
        "Trip budget by category with budget used, cost per day and cost per person",
        "Expenses in up to 12 currencies, converted to your home currency with your own exchange rates",
        "Status per cost: Paid, Booked – pay later or Planned, so you see the full trip cost before you book",
        "Group split for up to 8 travellers: shared vs personal costs and who owes whom",
        "Day-by-day itinerary with day numbers, booking references and 'not booked yet' alerts",
        "Ready-made packing list (32 items) with progress bar",
        "Dashboard with countdown and chart; blank copy plus a sample copy (Lisbon & Porto trip)",
    ],
    "tags": ["travel planner", "trip budget", "travel itinerary", "vacation planner", "travel spreadsheet", "expense splitter",
             "packing list", "google sheets", "excel template", "holiday budget", "road trip planner", "group travel", "travel budget tracker"],
}
