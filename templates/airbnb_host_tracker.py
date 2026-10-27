"""Builds the Airbnb & Short-Term Rental Host Tracker (sample + blank versions)."""
import random, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.utils import get_column_letter as L
from _common import Kit, F, INK, B, MONEY, DATE, font, fill, style_range, input_cell, finish

MAIN = "BE123C"; SOFT = "FDECEF"
K = Kit(MAIN, SOFT)
BR = 800  # booking rows
ER = 600  # expense rows
LISTN = 10

def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); bk = wb.create_sheet("Bookings"); cal = wb.create_sheet("Calendar")
    exp = wb.create_sheet("Expenses"); lp = wb.create_sheet("Listing P&L"); s = wb.create_sheet("Settings")

    # ---------- Settings ----------
    K.title(s, "Settings", "Set these once. Yellow cells are yours to edit.")
    for i, (k, v) in enumerate([("Currency label", "USD"), ("Report year", 2026), ("Calendar month (1-12)", 10)], 4):
        s.cell(row=i, column=1, value=k).font = font(bold=True); input_cell(s.cell(row=i, column=2), v)
    s["A7"] = "Calendar month start"; s["B7"] = "=DATE(B5,B6,1)"; s["B7"].number_format = DATE
    s["A8"] = "Days in report year"; s["B8"] = "=DATE(B5,12,31)-DATE(B5,1,1)+1"
    for a in ("A7", "A8"): s[a].font = font(italic=True)
    K.section(s, "A10", "Listings (up to 10)")
    K.hdr(s, 11, ["Listing name", "Address / notes", "Bedrooms", "Active from (optional)"])
    listings = [("Beach Loft", "21 Shore Dr", 1, None), ("Mountain Cabin", "4 Pine Ridge", 3, None), ("City Studio", "77 Main St #5", 0, dt.date(2026, 3, 1))] if sample else []
    for i, row in enumerate(listings):
        for j, v in enumerate(row, 1): s.cell(row=12 + i, column=j, value=v)
    style_range(s, 12, 11 + LISTN, 1, 4, inputs=(1, 2, 3, 4), fmt={4: DATE})
    K.section(s, "F10", "Booking channels & host fee %")
    K.hdr(s, 11, ["Channel", "Host fee %"], col0=6)
    chans = [("Airbnb", 0.155), ("Vrbo", 0.08), ("Booking.com", 0.15), ("Direct", 0.0), ("Other", 0.0)]
    for i in range(8):
        a = s.cell(row=12 + i, column=6, value=chans[i][0] if i < len(chans) else None)
        b = s.cell(row=12 + i, column=7, value=chans[i][1] if i < len(chans) else None)
        input_cell(a); input_cell(b); b.number_format = '0.0%'
    s["F21"] = "Fee % is used only when you leave 'Platform fee' blank on a booking. Check your own fee in your host dashboard."; s["F21"].font = font(italic=True, size=9)
    K.section(s, "I10", "Expense categories")
    K.hdr(s, 11, ["Category"], col0=9)
    cats = ["Cleaning", "Laundry & linens", "Supplies & toiletries", "Utilities", "Internet & TV", "Repairs & maintenance", "Furnishing",
            "Mortgage / rent", "Insurance", "Property tax", "Software & pricing tools", "Co-host / management", "Other"]
    for i in range(15): input_cell(s.cell(row=12 + i, column=9, value=cats[i] if i < len(cats) else None))
    K.section(s, "K10", "Booking status")
    for i, v in enumerate(["Confirmed", "Cancelled"]): s.cell(row=11 + i, column=11, value=v).font = font()
    for col, w in zip("ABCDEFGHIJK", [24, 26, 10, 14, 3, 16, 11, 3, 24, 3, 14]): s.column_dimensions[col].width = w

    LST = f"Settings!$A$12:$A${11+LISTN}"; CH = "Settings!$F$12:$F$19"; CATS = "Settings!$I$12:$I$26"; YR = "Settings!$B$5"
    N = 4 + BR
    rng = lambda c: f"Bookings!${c}$5:${c}${N}"

    # ---------- Bookings ----------
    K.title(bk, "Bookings", "One row per stay. Yellow columns are yours; grey columns calculate themselves. Leave 'Platform fee' blank to use the channel's fee %.")
    cols = ["Listing", "Channel", "Guest name", "Check-in", "Check-out", "Guests", "Nightly rate", "Cleaning fee charged", "Other income",
            "Platform fee (optional)", "Cleaner cost", "Status", "Rating (1-5)",
            "Nights", "Gross revenue", "Platform fee", "Payout", "Net after cleaning", "Year", "Month", "Days to check-in"]
    K.hdr(bk, 4, cols, [18, 13, 20, 12, 12, 8, 11, 11, 10, 12, 11, 12, 9, 8, 13, 12, 13, 13, 7, 7, 10])
    style_range(bk, 5, N, 1, 21, inputs=range(1, 14), fmt={4: DATE, 5: DATE, 7: MONEY, 8: MONEY, 9: MONEY, 10: MONEY, 11: MONEY, 15: MONEY, 16: MONEY, 17: MONEY, 18: MONEY})
    for r in range(5, N + 1):
        live = f'AND(A{r}<>"",D{r}<>"",E{r}<>"",L{r}<>"Cancelled")'
        bk[f"N{r}"] = f'=IF(OR(D{r}="",E{r}=""),"",MAX(0,E{r}-D{r}))'
        bk[f"O{r}"] = f'=IF(N{r}="","",N{r}*G{r}+H{r}+I{r})'
        bk[f"P{r}"] = f'=IF(O{r}="","",IF(J{r}<>"",J{r},ROUND(O{r}*IFERROR(INDEX(Settings!$G$12:$G$19,MATCH(B{r},{CH},0)),0),2)))'
        bk[f"Q{r}"] = f'=IF(O{r}="","",O{r}-P{r})'
        bk[f"R{r}"] = f'=IF(Q{r}="","",Q{r}-K{r})'
        bk[f"S{r}"] = f'=IF(D{r}="","",YEAR(D{r}))'
        bk[f"T{r}"] = f'=IF(D{r}="","",MONTH(D{r}))'
        bk[f"U{r}"] = f'=IF(OR(D{r}="",L{r}="Cancelled"),"",IF(D{r}>=TODAY(),D{r}-TODAY(),""))'
    for ref, f1 in ((f"A5:A{N}", f"={LST}"), (f"B5:B{N}", f"={CH}"), (f"L5:L{N}", "=Settings!$K$11:$K$12")):
        dv = DataValidation(type="list", formula1=f1, allow_blank=True); bk.add_data_validation(dv); dv.add(ref)
    dv = DataValidation(type="decimal", operator="between", formula1="1", formula2="5", allow_blank=True); bk.add_data_validation(dv); dv.add(f"M5:M{N}")
    dv = DataValidation(type="custom", formula1="E5>D5", allow_blank=True, error="Check-out must be after check-in", showErrorMessage=True); bk.add_data_validation(dv); dv.add(f"E5:E{N}")
    bk.conditional_formatting.add(f"A5:U{N}", FormulaRule(formula=['$L5="Cancelled"'], font=Font(name=F, color="9CA3AF", strike=True)))
    bk.conditional_formatting.add(f"U5:U{N}", FormulaRule(formula=['AND(ISNUMBER(U5),U5<=7)'], fill=fill("FEF3C7"), font=Font(name=F, bold=True, color="92400E")))
    bk.freeze_panes = "D5"

    # ---------- Expenses ----------
    K.title(exp, "Expenses", "Log every cost. Pick a listing, or leave Listing blank for costs shared by all listings.")
    K.hdr(exp, 4, ["Date", "Listing (blank = shared)", "Category", "Description", "Amount", "Vendor / payee", "Year", "Month"], [12, 22, 22, 30, 12, 20, 8, 8])
    style_range(exp, 5, 4 + ER, 1, 8, inputs=range(1, 7), fmt={1: DATE, 5: MONEY})
    for r in range(5, 5 + ER):
        exp[f"G{r}"] = f'=IF(A{r}="","",YEAR(A{r}))'; exp[f"H{r}"] = f'=IF(A{r}="","",MONTH(A{r}))'
    for ref, f1 in ((f"B5:B{4+ER}", f"={LST}"), (f"C5:C{4+ER}", f"={CATS}")):
        dv = DataValidation(type="list", formula1=f1, allow_blank=True); exp.add_data_validation(dv); dv.add(ref)
    exp.freeze_panes = "A5"

    if sample:
        random.seed(11)
        guests = ["Olivia M.", "Liam K.", "Noah P.", "Emma W.", "Ava R.", "Lucas T.", "Mia S.", "Ethan B.", "Sophia L.", "James H.", "Chloe D.", "Ben A."]
        prof = {"Beach Loft": (145, 75, 55), "Mountain Cabin": (230, 120, 90), "City Studio": (98, 50, 40)}
        r = 5
        for name, (rate, cfee, ccost) in prof.items():
            d = dt.date(2026, 3, 3) if name == "City Studio" else dt.date(2026, 1, 3)
            while d < dt.date(2026, 12, 20):
                nights = random.choice([2, 2, 3, 3, 4, 5, 7])
                ch = random.choice(["Airbnb", "Airbnb", "Airbnb", "Vrbo", "Booking.com", "Direct"])
                summer = 1.25 if d.month in (6, 7, 8) else 1.0
                bk[f"A{r}"] = name; bk[f"B{r}"] = ch; bk[f"C{r}"] = random.choice(guests)
                bk[f"D{r}"] = d; bk[f"E{r}"] = d + dt.timedelta(days=nights); bk[f"F{r}"] = random.randint(1, 4)
                bk[f"G{r}"] = round(rate * summer); bk[f"H{r}"] = cfee; bk[f"K{r}"] = ccost
                cancelled = random.random() < 0.05
                bk[f"L{r}"] = "Cancelled" if cancelled else "Confirmed"
                if not cancelled and d < dt.date(2026, 10, 27): bk[f"M{r}"] = random.choice([5, 5, 5, 4.5, 4, 5, 4.8])
                r += 1
                d = d + dt.timedelta(days=nights + random.choice([0, 1, 1, 2, 3, 5, 8]))
        er = 5
        for m in range(1, 11):
            for (lst, cat, desc, amt, ven) in [("", "Internet & TV", "Internet bundle", 120, "FiberNet"), ("", "Software & pricing tools", "Dynamic pricing tool", 39, "PriceBot"),
                                                ("Beach Loft", "Utilities", "Electricity & water", 140, "City Power"), ("Mountain Cabin", "Utilities", "Electricity & propane", 210, "Valley Energy"),
                                                ("Mountain Cabin", "Mortgage / rent", "Mortgage payment", 1350, "First Bank"), ("City Studio", "Mortgage / rent", "Lease payment", 1100, "Main St Realty")]:
                if lst == "City Studio" and m < 3: continue
                exp[f"A{er}"] = dt.date(2026, m, random.randint(1, 8)); exp[f"B{er}"] = lst or None; exp[f"C{er}"] = cat
                exp[f"D{er}"] = desc; exp[f"E{er}"] = amt; exp[f"F{er}"] = ven; er += 1
            for _ in range(2):
                lst = random.choice(list(prof)); cat, desc, amt = random.choice([("Supplies & toiletries", "Restock coffee, soap, paper", 65), ("Laundry & linens", "New towel set", 90),
                                                                                  ("Repairs & maintenance", "Fix dishwasher", 180), ("Furnishing", "Replace lamp", 45)])
                exp[f"A{er}"] = dt.date(2026, m, random.randint(10, 27)); exp[f"B{er}"] = lst; exp[f"C{er}"] = cat; exp[f"D{er}"] = desc; exp[f"E{er}"] = amt; er += 1
        exp[f"A{er}"] = dt.date(2026, 2, 14); exp[f"C{er}"] = "Insurance"; exp[f"D{er}"] = "Short-term rental insurance (all listings)"; exp[f"E{er}"] = 960; exp[f"F{er}"] = "HostShield"

    # night-overlap formula for a period [S, E) where S, E are cell refs to dates; optional listing criterion
    def nights_in(S, E, listing=None):
        ci, co, st_, ls = rng("D"), rng("E"), rng("L"), rng("A")
        min_end = f"({co}*({co}<{E})+{E}*({co}>={E}))"
        max_st = f"({ci}*({ci}>{S})+{S}*({ci}<={S}))"
        cond = f'({st_}<>"Cancelled")*({ls}<>"")'
        if listing: cond += f"*({ls}={listing})"
        return f"SUMPRODUCT({cond}*({min_end}-{max_st})*({min_end}>{max_st}))"
    def sumb(col, *crit):
        return f"SUMIFS({rng(col)},{rng('L')},\"<>Cancelled\"" + "".join(f",{rng(c)},{v}" for c, v in crit) + ")"
    # listing-days available in [S,E): sum over listings of days active
    def avail(S, E, single=None):
        if single:
            return f"IF({single[0]}=\"\",0,MAX(0,{E}-MAX({S},IF({single[1]}=\"\",{S},{single[1]}))))"
        a = f"Settings!$A$12:$A${11+LISTN}"; d = f"Settings!$D$12:$D${11+LISTN}"
        start = f"(({d}=\"\")*{S}+({d}<>\"\")*(({d}>{S})*{d}+({d}<={S})*{S}))"
        return f"SUMPRODUCT(({a}<>\"\")*({E}-{start})*({E}>{start}))"

    # ---------- Dashboard ----------
    K.title(dash, "Short-Term Rental Dashboard", "Everything here calculates from your other tabs. Change the report year on the Settings tab.")
    dash["A3"] = '="Year "&Settings!B5&"  ·  Amounts in "&Settings!B4&"  ·  Cancelled bookings are excluded"'; dash["A3"].font = font(bold=True)
    for c, w in zip("ABCDEFGHIJ", [17, 13, 14, 13, 15, 14, 13, 14, 13, 12]): dash.column_dimensions[c].width = w
    # month table first (rows 13-25) so KPIs can reference it
    K.section(dash, "A12", "Month by month (revenue counted in the check-in month)")
    K.hdr(dash, 13, ["Month", "Nights booked", "Occupancy", "Gross revenue", "Payout", "Cleaning costs", "Expenses", "Net profit", "ADR"])
    for m in range(1, 13):
        r = 13 + m
        dash[f"K{r}"] = f"=DATE({YR},{m},1)"; dash[f"L{r}"] = f"=DATE({YR},{m+1},1)" if m < 12 else f"=DATE({YR}+1,1,1)"
        dash[f"A{r}"] = dt.date(2026, m, 1).strftime("%b")
        dash[f"B{r}"] = "=" + nights_in(f"$K{r}", f"$L{r}")
        dash[f"C{r}"] = f"=IFERROR(B{r}/{avail(f'$K{r}', f'$L{r}')},0)"
        dash[f"D{r}"] = "=" + sumb("O", ("S", YR), ("T", m))
        dash[f"E{r}"] = "=" + sumb("Q", ("S", YR), ("T", m))
        dash[f"F{r}"] = "=" + sumb("K", ("S", YR), ("T", m))
        dash[f"G{r}"] = f"=SUMIFS(Expenses!$E$5:$E${4+ER},Expenses!$G$5:$G${4+ER},{YR},Expenses!$H$5:$H${4+ER},{m})"
        dash[f"H{r}"] = f"=E{r}-F{r}-G{r}"
        dash[f"I{r}"] = f"=IFERROR(({sumb('O', ('S', YR), ('T', m))}-{sumb('H', ('S', YR), ('T', m))}-{sumb('I', ('S', YR), ('T', m))})/{sumb('N', ('S', YR), ('T', m))},0)"
        for c in "ABCDEFGHI":
            cell = dash[f"{c}{r}"]; cell.border = B; cell.font = font()
            cell.number_format = {"B": "0", "C": "0%", "A": "General"}.get(c, MONEY)
        for c in "KL": dash[f"{c}{r}"].number_format = DATE; dash[f"{c}{r}"].font = Font(name=F, color="FFFFFF", size=8)
    dash["A26"] = "Total"
    dash["B26"] = "=SUM(B14:B25)"; dash["C26"] = f"=IFERROR(B26/{avail('$K$14', '$L$25')},0)"
    for c in "DEFGH": dash[f"{c}26"] = f"=SUM({c}14:{c}25)"
    dash["I26"] = f"=IFERROR(({sumb('O', ('S', YR))}-{sumb('H', ('S', YR))}-{sumb('I', ('S', YR))})/{sumb('N', ('S', YR))},0)"
    for c in "ABCDEFGHI":
        cell = dash[f"{c}26"]; cell.font = font(bold=True); cell.border = B; cell.fill = fill(SOFT)
        cell.number_format = {"B": "0", "C": "0%", "A": "General"}.get(c, MONEY)
    K0 = '#,##0;(#,##0);"-"'
    kpis = [("Gross revenue", "=D26", K0), ("Payout after platform fees", "=E26", K0), ("Net profit", "=H26", K0),
            ("Nights booked", "=B26", "0"), ("Occupancy", "=C26", "0%"), ("Average daily rate (ADR)", "=I26", MONEY),
            ("Bookings", f'=COUNTIFS({rng("S")},{YR},{rng("L")},"<>Cancelled")', "0"),
            ("Average rating", f'=IFERROR(AVERAGEIFS({rng("M")},{rng("S")},{YR},{rng("L")},"<>Cancelled"),0)', "0.00"),
            ("Check-ins next 7 days", f'=COUNTIFS({rng("U")},">=0",{rng("U")},"<=7")', "0")]
    pos = [(4, 1), (4, 3), (4, 5), (6, 1), (6, 3), (6, 5), (8, 1), (8, 3), (8, 5)]
    for (row, col), (lab, f, fmt) in zip(pos, kpis): K.kpi(dash, row, col, lab, f, fmt)
    dash["G4"] = "Platform fees paid"; dash["G5"] = "=D26-E26"
    dash["G6"] = "Cancelled bookings"; dash["G7"] = f'=COUNTIFS({rng("S")},{YR},{rng("L")},"Cancelled")'
    dash["G8"] = "Profit margin"; dash["G9"] = "=IFERROR(H26/D26,0)"
    for a, fmt in (("G5", MONEY), ("G7", "0"), ("G9", "0%")):
        dash[a].font = Font(name=F, size=14, bold=True, color=INK); dash[a].number_format = fmt
    for a in ("G4", "G6", "G8"): dash[a].font = Font(name=F, size=9, color="5B6770", bold=True)
    ch = BarChart(); ch.type = "col"; ch.title = "Payout vs costs by month"; ch.height = 7.5; ch.width = 17
    ch.add_data(Reference(dash, min_col=5, min_row=13, max_row=25), titles_from_data=True)
    ch.add_data(Reference(dash, min_col=7, min_row=13, max_row=25), titles_from_data=True)
    ch.set_categories(Reference(dash, min_col=1, min_row=14, max_row=25)); ch.y_axis.majorGridlines = None
    for ser, col in zip(ch.series, (MAIN, "F59E0B")):
        ser.graphicalProperties = GraphicalProperties(solidFill=col); ser.graphicalProperties.line.solidFill = col
    dash.add_chart(ch, "A45")
    lc = LineChart(); lc.title = "Occupancy by month"; lc.height = 7.5; lc.width = 15
    lc.add_data(Reference(dash, min_col=3, min_row=13, max_row=25), titles_from_data=True)
    lc.set_categories(Reference(dash, min_col=1, min_row=14, max_row=25)); lc.y_axis.number_format = "0%"; lc.y_axis.majorGridlines = None
    lc.series[0].graphicalProperties.line.solidFill = MAIN; lc.series[0].graphicalProperties.line.width = 28000
    dash.add_chart(lc, "F45")
    # by listing
    K.section(dash, "A28", "By listing (report year)")
    K.hdr(dash, 29, ["Listing", "Nights booked", "Occupancy", "Gross revenue", "Payout", "Cleaning costs", "Expenses", "Net profit", "ADR", "Avg rating"])
    dash["K29"] = "Year start"; dash["L29"] = "Next year"
    dash["K30"] = f"=DATE({YR},1,1)"; dash["L30"] = f"=DATE({YR}+1,1,1)"
    for a in ("K29", "L29", "K30", "L30"): dash[a].font = Font(name=F, color="FFFFFF", size=8)
    for i in range(LISTN):
        r = 30 + i; sr = 12 + i
        dash[f"A{r}"] = f'=IF(Settings!A{sr}="","",Settings!A{sr})'
        dash[f"B{r}"] = f'=IF(A{r}="","",{nights_in("$K$30", "$L$30", f"$A{r}")})'
        dash[f"C{r}"] = f'=IF(A{r}="","",IFERROR(B{r}/{avail("$K$30", "$L$30", (f"Settings!A{sr}", f"Settings!D{sr}"))},0))'
        dash[f"D{r}"] = f'=IF(A{r}="","",{sumb("O", ("S", YR), ("A", f"$A{r}"))})'
        dash[f"E{r}"] = f'=IF(A{r}="","",{sumb("Q", ("S", YR), ("A", f"$A{r}"))})'
        dash[f"F{r}"] = f'=IF(A{r}="","",{sumb("K", ("S", YR), ("A", f"$A{r}"))})'
        dash[f"G{r}"] = f'=IF(A{r}="","",SUMIFS(Expenses!$E$5:$E${4+ER},Expenses!$B$5:$B${4+ER},$A{r},Expenses!$G$5:$G${4+ER},{YR}))'
        dash[f"H{r}"] = f'=IF(A{r}="","",E{r}-F{r}-G{r})'
        dash[f"I{r}"] = f'=IF(A{r}="","",IFERROR(({sumb("O", ("S", YR), ("A", f"$A{r}"))}-{sumb("H", ("S", YR), ("A", f"$A{r}"))}-{sumb("I", ("S", YR), ("A", f"$A{r}"))})/{sumb("N", ("S", YR), ("A", f"$A{r}"))},0))'
        dash[f"J{r}"] = f'=IF(A{r}="","",IFERROR(AVERAGEIFS({rng("M")},{rng("S")},{YR},{rng("A")},$A{r},{rng("L")},"<>Cancelled"),"-"))'
        for j in range(1, 11):
            c = dash.cell(row=r, column=j); c.border = B; c.font = font()
            c.number_format = {2: "0", 3: "0%", 10: "0.00"}.get(j, MONEY if j > 1 else "General")
    dash["A40"] = "Shared costs (no listing)"; dash["G40"] = f'=SUMIFS(Expenses!$E$5:$E${4+ER},Expenses!$B$5:$B${4+ER},"",Expenses!$G$5:$G${4+ER},{YR},Expenses!$A$5:$A${4+ER},"<>")'
    dash["H40"] = "=-G40"
    for c in "AGH": dash[f"{c}40"].font = font(italic=True); dash[f"{c}40"].border = B
    for c in "GH": dash[f"{c}40"].number_format = MONEY
    # by channel
    K.section(dash, "A42", "Channel mix: see the table at the right (column N)")
    K.hdr(dash, 29, ["Channel", "Bookings", "Gross revenue", "Platform fees", "Payout"], [14, 10, 14, 13, 14], col0=14)
    for i in range(8):
        r = 30 + i; sr = 12 + i
        dash[f"N{r}"] = f'=IF(Settings!F{sr}="","",Settings!F{sr})'
        dash[f"O{r}"] = f'=IF(N{r}="","",COUNTIFS({rng("B")},N{r},{rng("S")},{YR},{rng("L")},"<>Cancelled"))'
        dash[f"P{r}"] = f'=IF(N{r}="","",{sumb("O", ("S", YR), ("B", f"N{r}"))})'
        dash[f"Q{r}"] = f'=IF(N{r}="","",{sumb("P", ("S", YR), ("B", f"N{r}"))})'
        dash[f"R{r}"] = f'=IF(N{r}="","",{sumb("Q", ("S", YR), ("B", f"N{r}"))})'
        for j in range(14, 19):
            c = dash.cell(row=r, column=j); c.border = B; c.font = font(); c.number_format = "0" if j == 15 else (MONEY if j > 15 else "General")
    dash["A42"] = None
    dash.column_dimensions["M"].width = 3

    # ---------- Calendar ----------
    K.title(cal, "Booking Calendar", "Nights booked per listing for the month set on the Settings tab (Calendar month). A filled cell = booked night.")
    cal["A3"] = '=TEXT(Settings!B7,"mmmm yyyy")'; cal["A3"].font = font(bold=True, size=12)
    cal["A5"] = "Listing"; cal["A5"].font = Font(name=F, bold=True, color="FFFFFF"); cal["A5"].fill = fill(MAIN); cal.column_dimensions["A"].width = 20
    for d in range(1, 32):
        col = 1 + d
        c = cal.cell(row=4, column=col, value=f'=IF({d}>DAY(EOMONTH(Settings!$B$7,0)),"",TEXT(Settings!$B$7+{d-1},"ddd"))'); c.font = font(size=8); c.alignment = Alignment(horizontal="center")
        c = cal.cell(row=5, column=col, value=f'=IF({d}>DAY(EOMONTH(Settings!$B$7,0)),"",{d})')
        c.font = Font(name=F, bold=True, color="FFFFFF"); c.fill = fill(MAIN); c.alignment = Alignment(horizontal="center")
        cal.column_dimensions[L(col)].width = 4.2
    cal.cell(row=5, column=33, value="Nights").font = Font(name=F, bold=True, color="FFFFFF"); cal.cell(row=5, column=33).fill = fill(MAIN)
    cal.cell(row=5, column=34, value="Occupancy").font = Font(name=F, bold=True, color="FFFFFF"); cal.cell(row=5, column=34).fill = fill(MAIN)
    cal.column_dimensions["AG"].width = 8; cal.column_dimensions["AH"].width = 10
    for i in range(LISTN):
        r = 6 + i
        cal[f"A{r}"] = f'=IF(Settings!A{12+i}="","",Settings!A{12+i})'; cal[f"A{r}"].font = font(bold=True); cal[f"A{r}"].border = B
        for d in range(1, 32):
            col = L(1 + d)
            cal[f"{col}{r}"] = (f'=IF(OR($A{r}="",{col}$5=""),"",COUNTIFS({rng("A")},$A{r},{rng("L")},"<>Cancelled",'
                                f'{rng("D")},"<="&(Settings!$B$7+{d-1}),{rng("E")},">"&(Settings!$B$7+{d-1})))')
            cell = cal[f"{col}{r}"]; cell.border = B; cell.alignment = Alignment(horizontal="center"); cell.font = Font(name=F, size=8, color="FFFFFF")
        cal[f"AG{r}"] = f'=IF($A{r}="","",SUM(B{r}:AF{r}))'
        cal[f"AH{r}"] = f'=IF($A{r}="","",AG{r}/DAY(EOMONTH(Settings!$B$7,0)))'; cal[f"AH{r}"].number_format = "0%"
        for c in ("AG", "AH"): cal[f"{c}{r}"].border = B; cal[f"{c}{r}"].font = font(bold=True)
    cal.conditional_formatting.add(f"B6:AF{5+LISTN}", CellIsRule(operator="greaterThanOrEqual", formula=["1"], fill=fill(MAIN), font=Font(name=F, color=MAIN)))
    cal.conditional_formatting.add(f"B6:AF{5+LISTN}", CellIsRule(operator="equal", formula=["0"], fill=fill("F9FAFB"), font=Font(name=F, color="F9FAFB")))
    cal["A18"] = "Tip: a 2 in a cell means two bookings overlap on that night - check for a double booking."; cal["A18"].font = font(italic=True, size=9)
    cal.conditional_formatting.add(f"B6:AF{5+LISTN}", CellIsRule(operator="greaterThanOrEqual", formula=["2"], fill=fill("111827"), font=Font(name=F, bold=True, color="FFFFFF")))

    # ---------- Listing P&L ----------
    K.title(lp, "Listing Profit & Loss", "Pick a listing in the yellow cell to see its 12-month P&L for the report year.")
    lp["A4"] = "Listing"; lp["A4"].font = font(bold=True)
    input_cell(lp["B4"], listings[0][0] if listings else None); lp["B4"].font = Font(name=F, color="0000FF", bold=True)
    dv = DataValidation(type="list", formula1=f"={LST}", allow_blank=True); lp.add_data_validation(dv); dv.add("B4")
    lp["D4"] = '="Year "&Settings!B5'; lp["D4"].font = font(bold=True)
    months = [dt.date(2026, m, 1).strftime("%b") for m in range(1, 13)]
    K.hdr(lp, 6, ["Line"] + months + ["Total"], [26] + [10] * 12 + [12])
    lines = [("Nights booked", "N"), ("Gross revenue", "O"), ("Platform fees", "P"), ("Payout", "Q"), ("Cleaning costs", "K")]
    for i, (lab, col) in enumerate(lines):
        r = 7 + i; lp[f"A{r}"] = lab
        for m in range(1, 13):
            lp.cell(row=r, column=1 + m, value=f"=IF($B$4=\"\",0,{sumb(col, ('S', YR), ('T', m), ('A', '$B$4'))})")
    lp["A12"] = "Expenses by category"; lp["A12"].font = Font(name=F, bold=True, color=MAIN)
    for i in range(15):
        r = 13 + i
        lp[f"A{r}"] = f'=IF(Settings!I{12+i}="","",Settings!I{12+i})'
        for m in range(1, 13):
            lp.cell(row=r, column=1 + m, value=f'=IF(OR($A{r}="",$B$4=""),"",SUMIFS(Expenses!$E$5:$E${4+ER},Expenses!$B$5:$B${4+ER},$B$4,Expenses!$C$5:$C${4+ER},$A{r},Expenses!$G$5:$G${4+ER},{YR},Expenses!$H$5:$H${4+ER},{m}))')
    lp["A28"] = "Total expenses"; lp["A29"] = "Net profit (payout - cleaning - expenses)"
    for m in range(1, 13):
        c = L(1 + m); lp[f"{c}28"] = f"=SUM({c}13:{c}27)"; lp[f"{c}29"] = f"={c}10-{c}11-{c}28"
    for r in list(range(7, 12)) + list(range(13, 30)):
        lp[f"N{r}"] = f'=IF(A{r}="","",SUM(B{r}:M{r}))'
        for c in range(1, 15):
            cell = lp.cell(row=r, column=c); cell.border = B
            cell.font = font(bold=(r in (8, 10, 28, 29) or c == 14))
            if c > 1: cell.number_format = "0" if r == 7 else MONEY
    for r in (28, 29):
        for c in range(1, 15): lp.cell(row=r, column=c).fill = fill(SOFT)
    lp.column_dimensions["A"].width = 38

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Settings tab: type your currency label, report year and your listings. Check the host fee % for each booking channel.", False),
             ("2.  Bookings tab: add one row per stay - listing, channel, guest, check-in/out, nightly rate, cleaning fee, and what you pay your cleaner.", False),
             ("     Platform fee: type the exact fee from your payout report, or leave it blank and the channel's fee % is used.", False),
             ("3.  Expenses tab: log every cost. Leave Listing blank for costs shared by all listings (insurance, software...).", False),
             ("4.  Dashboard, Calendar and Listing P&L update on their own.", False),
             ("", False),
             ("What you'll see", True),
             ("•  Dashboard: revenue, payout, net profit, nights booked, occupancy, ADR, average rating, upcoming check-ins, month-by-month table and charts, results per listing and per channel.", False),
             ("•  Calendar: a month view of booked nights for every listing - double bookings show up dark. Change the month on the Settings tab.", False),
             ("•  Listing P&L: a 12-month profit and loss for one listing, with expenses by category.", False),
             ("", False),
             ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("Cancelled bookings turn grey and are left out of every total. Check-ins within 7 days are flagged amber.", False),
             ("", False),
             ("Good to know", True),
             ("•  Occupancy counts every night inside the month, even when a stay crosses months. Revenue is counted in the check-in month.", False),
             ("•  ADR = nightly revenue (gross minus cleaning fee and other income) divided by nights.", False),
             ("•  Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False),
             ("•  Room for 10 listings, 800 bookings and 600 expenses. Copy a row down if you need more.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("The listings, bookings and expenses are made-up examples. Clear the yellow cells on Bookings and Expenses (and the listings on Settings) before entering your own, or use the blank copy in your download.", False)]
    K.start(st, "Airbnb & Short-Term Rental Host Tracker", "Bookings, occupancy, payouts, cleaning costs and profit per listing. Works in Excel and Google Sheets.", lines)
    finish(wb, sample, path)

META = {
    "name": "Airbnb Host Tracker – Short-Term Rental Spreadsheet (Excel & Google Sheets)",
    "file_base": "Airbnb-Host-Tracker",
    "price": 24,
    "color": "#BE123C",
    "title": "Airbnb & Short-Term Rental Host Tracker",
    "tagline": "Bookings · occupancy · payouts · cleaning · profit per listing — Excel & Google Sheets",
    "intro": "Know exactly what each of your Airbnb, Vrbo or Booking.com listings earns. Log your stays and costs once and the tracker works out occupancy, ADR, platform fees, payouts and net profit for you.",
    "features": [
        "Dashboard: gross revenue, payout, net profit, nights booked, occupancy, ADR, average rating and check-ins in the next 7 days",
        "Month-by-month table with payout vs cost and occupancy charts",
        "Results per listing (up to 10) and per channel (Airbnb, Vrbo, Booking.com, Direct...)",
        "Bookings log: nights, gross, platform fee (exact or automatic by channel %), payout and net after cleaning calculated for you",
        "Booking calendar: month view of booked nights per listing that flags double bookings",
        "Listing P&L: 12-month profit and loss for each listing with 13 ready expense categories",
        "Cancelled bookings are excluded automatically; any currency; room for 800 bookings",
    ],
    "tags": ["airbnb spreadsheet", "airbnb host", "short term rental", "vrbo tracker", "rental income", "occupancy tracker", "google sheets",
             "excel template", "airbnb template", "vacation rental", "host income tracker", "airbnb expenses", "superhost"],
}
