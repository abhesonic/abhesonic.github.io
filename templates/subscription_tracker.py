"""Subscription Tracker – recurring payments, renewal dates, free-trial alerts, monthly/yearly cost, 12-month forecast, cancel list.

Settings (cycles, categories, payment methods) -> Subscriptions -> 12-Month Forecast -> Dashboard.
"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import BarChart, Reference
from _common import Kit, F, B, MONEY, DATE, font, fill, style_range, input_cell, finish

MAIN = "7C3AED"; SOFT = "EDE9FE"
K = Kit(MAIN, SOFT)
NS = 150; S0 = 6; S1 = S0 + NS - 1
F0 = 7; F1 = F0 + NS - 1
CYCLES = [("Weekly", 52, 0, 7), ("Every 2 weeks", 26, 0, 14), ("Monthly", 12, 1, 0), ("Every 3 months", 4, 3, 0),
          ("Every 6 months", 2, 6, 0), ("Yearly", 1, 12, 0)]
CY0 = 6; CY1 = CY0 + len(CYCLES) - 1
CATS = ["Streaming TV", "Music & audio", "Software & apps", "Cloud storage", "News & magazines", "Gaming", "Fitness & health",
        "Food & delivery", "Shopping memberships", "Education", "Phone & internet", "Business tools", "Other"]
NCAT = 15; CA0 = 6; CA1 = CA0 + NCAT - 1
PAYS = ["Visa ••1234", "Mastercard ••5678", "PayPal", "Bank debit", "Apple / Google Pay"]
NP = 8; P0 = 6; P1 = P0 + NP - 1


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); su = wb.create_sheet("Subscriptions"); fc = wb.create_sheet("12-Month Forecast")
    s = wb.create_sheet("Settings")

    # ---------- Settings ----------
    K.title(s, "Settings", "Yellow cells are yours to edit.")
    K.hdr(s, 5, ["Billing cycle", "Charges / year", "Months per charge", "Days per charge"])
    for i, row in enumerate(CYCLES):
        for j, v in enumerate(row): s.cell(row=CY0 + i, column=1 + j, value=v)
    style_range(s, CY0, CY1, 1, 4)
    s[f"A{CY1 + 1}"] = "Fixed list – used by the formulas."; s[f"A{CY1 + 1}"].font = font(italic=True, size=9)
    K.hdr(s, 5, ["Categories"], col0=6)
    for i in range(NCAT): input_cell(s[f"F{CA0 + i}"], CATS[i] if i < len(CATS) else None)
    K.hdr(s, 5, ["Payment methods"], col0=8)
    for i in range(NP): input_cell(s[f"H{P0 + i}"], (PAYS[i] if i < len(PAYS) else None) if sample else (["Card 1", "Card 2", "PayPal", "Bank debit"] + [None] * 4)[i])
    SET = [("Currency (label only)", "USD"), ("Warn me this many days before a charge", 7),
           ("Forecast start month (blank = this month)", None)]
    for i, (lab, v) in enumerate(SET, 14):
        s[f"A{i}"] = lab; s[f"A{i}"].font = font(bold=True); s[f"A{i}"].border = B; input_cell(s[f"C{i}"], v)
    s["C16"].number_format = "mmm yyyy"
    s["A17"] = "Forecast starts"; s["A17"].font = font(italic=True)
    s["C17"] = '=IF(ISNUMBER(C16),DATE(YEAR(C16),MONTH(C16),1),DATE(YEAR(TODAY()),MONTH(TODAY()),1))'; s["C17"].number_format = "mmm yyyy"
    for c, w in zip("ABCDEFGH", [20, 14, 16, 15, 3, 24, 3, 22]): s.column_dimensions[c].width = w
    WARN = "Settings!$C$15"; CYC = f"Settings!$A${CY0}:$A${CY1}"
    CUR = "Settings!$C$14"

    # ---------- Subscriptions ----------
    K.title(su, "Subscriptions", "One row per subscription. Yellow = you type, grey = automatic. Status 'Cancelled' or 'Paused' stops the cost.")
    cols = ["Subscription", "Category", "Cost per charge", "Billing cycle", "First charge date", "Payment method", "Status", "Free trial ends",
            "How often I use it", "Verdict", "Notes / cancel link", "Monthly cost", "Yearly cost", "Next charge", "Days until", "Alert"]
    K.hdr(su, 5, cols, [24, 18, 11, 14, 12, 18, 11, 12, 11, 9, 30, 11, 11, 12, 8, 28])
    style_range(su, S0, S1, 1, 16, inputs=range(1, 12), fmt={3: MONEY, 5: DATE, 8: DATE, 12: MONEY, 13: MONEY, 14: DATE, 15: "0"})
    for r in range(S0, S1 + 1):
        su[f"Q{r}"] = f'=IFERROR(INDEX(Settings!$C${CY0}:$C${CY1},MATCH(D{r},{CYC},0)),0)'
        su[f"R{r}"] = f'=IFERROR(INDEX(Settings!$D${CY0}:$D${CY1},MATCH(D{r},{CYC},0)),0)'
        su[f"S{r}"] = f'=IF(AND(G{r}="Free trial",ISNUMBER(H{r})),H{r},IF(ISNUMBER(E{r}),E{r},""))'
        su[f"T{r}"] = (f'=IF(AND(A{r}<>"",ISNUMBER(C{r}),ISNUMBER(S{r}),Q{r}+R{r}>0,OR(G{r}="Active",G{r}="Free trial",G{r}="")),1,0)')
        su[f"L{r}"] = f'=IF(OR(A{r}="",NOT(ISNUMBER(C{r}))),"",IF(OR(G{r}="Paused",G{r}="Cancelled"),0,C{r}*IFERROR(INDEX(Settings!$B${CY0}:$B${CY1},MATCH(D{r},{CYC},0)),0)/12))'
        su[f"M{r}"] = f'=IF(L{r}="","",L{r}*12)'
        kk = f"((YEAR(TODAY())-YEAR(S{r}))*12+MONTH(TODAY())-MONTH(S{r}))"
        c1 = f"EDATE(S{r},CEILING({kk}/Q{r},1)*Q{r})"
        su[f"N{r}"] = (f'=IF(T{r}<>1,"",IF(S{r}>=TODAY(),S{r},IF(R{r}>0,S{r}+CEILING((TODAY()-S{r})/R{r},1)*R{r},'
                       f'IF({c1}>=TODAY(),{c1},EDATE(S{r},(CEILING({kk}/Q{r},1)+1)*Q{r})))))')
        su[f"O{r}"] = f'=IF(N{r}="","",N{r}-TODAY())'
        su[f"P{r}"] = (f'=IF(A{r}="","",IF(AND(G{r}="Free trial",ISNUMBER(O{r}),O{r}<={WARN}),"Trial ends soon – cancel or keep?",'
                       f'IF(AND(J{r}="Cancel",G{r}<>"Cancelled"),"Cancel to save "&TEXT(N(M{r}),"#,##0")&" a year",'
                       f'IF(AND(ISNUMBER(O{r}),O{r}<={WARN}),"Charges in "&O{r}&" day(s)",""))))')
        su[f"U{r}"] = f'=IF(ISNUMBER(N{r}),N{r}+ROW()/1000000,99999999)'
        su[f"P{r}"].font = font(bold=True)
    for c in "QRSTU": su.column_dimensions[c].hidden = True
    for col, lst in (("D", f"={CYC}"), ("B", f"=Settings!$F${CA0}:$F${CA1}"), ("F", f"=Settings!$H${P0}:$H${P1}"),
                     ("G", '"Active,Free trial,Paused,Cancelled"'), ("I", '"Daily,Weekly,Monthly,Rarely,Never"'), ("J", '"Keep,Review,Cancel"')):
        dv = DataValidation(type="list", formula1=lst, allow_blank=True); su.add_data_validation(dv); dv.add(f"{col}{S0}:{col}{S1}")
    su.conditional_formatting.add(f"P{S0}:P{S1}", FormulaRule(formula=[f'OR(LEFT(P{S0},5)="Trial",LEFT(P{S0},6)="Cancel")'],
                                                             font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    su.conditional_formatting.add(f"P{S0}:P{S1}", FormulaRule(formula=[f'LEFT(P{S0},7)="Charges"'], font=Font(name=F, bold=True, color="B45309"), fill=fill("FEF3C7")))
    su.conditional_formatting.add(f"A{S0}:P{S1}", FormulaRule(formula=[f'OR($G{S0}="Cancelled",$G{S0}="Paused")'], font=Font(name=F, color="9CA3AF")))
    su.conditional_formatting.add(f"I{S0}:I{S1}", FormulaRule(formula=[f'OR(I{S0}="Rarely",I{S0}="Never")'], font=Font(name=F, bold=True, color="B91C1C")))
    su.freeze_panes = f"B{S0}"; su.auto_filter.ref = f"A5:P{S1}"

    # ---------- 12-Month Forecast ----------
    K.title(fc, "12-Month Forecast", "What will actually leave your account each month (active subscriptions and free trials after they end).")
    K.hdr(fc, 5, ["Subscription", "Billing cycle", "Cost"] + [""] * 12 + ["12-month total"], [24, 14, 10] + [10] * 12 + [13])
    for j in range(12):
        c = fc.cell(row=5, column=4 + j, value=f"=EDATE(Settings!$C$17,{j})"); c.number_format = "mmm yy"
        c.font = Font(name=F, bold=True, color="FFFFFF")
    fc["A6"] = "Total per month"
    for j in range(13):
        L = CL(4 + j); fc[f"{L}6"] = f"=SUM({L}{F0}:{L}{F1})"
    for c in range(1, 17):
        cell = fc.cell(row=6, column=c); cell.font = Font(name=F, bold=True, color=MAIN); cell.fill = fill(SOFT); cell.border = B
        if c >= 4: cell.number_format = MONEY
    for k in range(NS):
        r = F0 + k; sr = S0 + k; X = lambda c: f"Subscriptions!${c}${sr}"
        fc[f"A{r}"] = f'=IF({X("A")}="","",{X("A")})'
        fc[f"B{r}"] = f'=IF({X("A")}="","",{X("D")})'
        fc[f"C{r}"] = f'=IF({X("A")}="","",{X("C")})'
        for j in range(12):
            L = CL(4 + j); ms = f"{L}$5"; me = f"EOMONTH({L}$5,0)"; S = X("S"); R = X("R"); Q = X("Q")
            diff = f"((YEAR({ms})-YEAR({S}))*12+MONTH({ms})-MONTH({S}))"
            byday = f"IF({me}<{S},0,INT(({me}-{S})/{R})+1-IF({ms}>{S},INT(({ms}-1-{S})/{R})+1,0))"
            bymon = f"IF(AND({diff}>=0,MOD({diff},{Q})=0),1,0)"
            fc[f"{L}{r}"] = f"=IF({X('T')}<>1,0,{X('C')}*IF({R}>0,{byday},{bymon}))"
        fc[f"P{r}"] = f"=SUM(D{r}:O{r})"
    style_range(fc, F0, F1, 1, 16, fmt={3: MONEY, **{c: MONEY for c in range(4, 17)}})
    for r in range(F0, F1 + 1): fc[f"P{r}"].font = font(bold=True)
    fc.freeze_panes = f"D{F0}"

    # ---------- sample ----------
    if sample:
        rows = [("StreamFlix Standard", "Streaming TV", 15.49, "Monthly", dt.date(2024, 3, 12), "Visa ••1234", "Active", None, "Daily", "Keep"),
                ("ScreenPlus Premium", "Streaming TV", 13.99, "Monthly", dt.date(2025, 6, 2), "Visa ••1234", "Active", None, "Rarely", "Cancel"),
                ("SportsLive Pass", "Streaming TV", 74.99, "Every 6 months", dt.date(2026, 2, 20), "PayPal", "Active", None, "Monthly", "Review"),
                ("TuneBox Family", "Music & audio", 16.99, "Monthly", dt.date(2023, 9, 5), "Mastercard ••5678", "Active", None, "Daily", "Keep"),
                ("AudioReads", "Music & audio", 14.95, "Monthly", dt.date(2025, 1, 18), "Visa ••1234", "Active", None, "Weekly", "Keep"),
                ("PodPro", "Music & audio", 4.99, "Monthly", dt.date(2026, 4, 9), "Apple / Google Pay", "Active", None, "Never", "Cancel"),
                ("OfficeSuite Personal", "Software & apps", 99.99, "Yearly", dt.date(2025, 2, 14), "Visa ••1234", "Active", None, "Daily", "Keep"),
                ("PhotoEdit Pro", "Software & apps", 9.99, "Monthly", dt.date(2025, 11, 1), "Mastercard ••5678", "Active", None, "Rarely", "Review"),
                ("VPNshield", "Software & apps", 59.88, "Yearly", dt.date(2025, 12, 1), "PayPal", "Active", None, "Weekly", "Keep"),
                ("PassVault", "Software & apps", 35.88, "Yearly", dt.date(2026, 5, 22), "Visa ••1234", "Active", None, "Daily", "Keep"),
                ("CloudDrive 2 TB", "Cloud storage", 9.99, "Monthly", dt.date(2024, 7, 28), "Apple / Google Pay", "Active", None, "Daily", "Keep"),
                ("PhotoCloud 200 GB", "Cloud storage", 2.99, "Monthly", dt.date(2025, 8, 3), "Apple / Google Pay", "Active", None, "Rarely", "Cancel"),
                ("Daily Ledger digital", "News & magazines", 4.0, "Weekly", dt.date(2026, 1, 6), "Visa ••1234", "Active", None, "Weekly", "Review"),
                ("GameClub Online", "Gaming", 79.99, "Yearly", dt.date(2025, 11, 25), "Mastercard ••5678", "Active", None, "Monthly", "Keep"),
                ("FitStudio App", "Fitness & health", 12.99, "Monthly", dt.date(2026, 10, 30), "Visa ••1234", "Free trial", dt.date(2026, 11, 13), "Weekly", "Review"),
                ("City Gym", "Fitness & health", 39.0, "Monthly", dt.date(2024, 1, 15), "Bank debit", "Active", None, "Rarely", "Cancel"),
                ("MealKit Box", "Food & delivery", 69.92, "Every 2 weeks", dt.date(2026, 6, 4), "Visa ••1234", "Paused", None, "Monthly", "Review"),
                ("DeliveryPass", "Food & delivery", 9.99, "Monthly", dt.date(2025, 10, 8), "PayPal", "Active", None, "Weekly", "Keep"),
                ("ShopPrime", "Shopping memberships", 139.0, "Yearly", dt.date(2023, 12, 3), "Visa ••1234", "Active", None, "Weekly", "Keep"),
                ("LearnLingo", "Education", 59.99, "Every 3 months", dt.date(2026, 3, 1), "PayPal", "Active", None, "Rarely", "Review"),
                ("Mobile plan", "Phone & internet", 45.0, "Monthly", dt.date(2022, 5, 20), "Bank debit", "Active", None, "Daily", "Keep"),
                ("Home internet", "Phone & internet", 60.0, "Monthly", dt.date(2022, 5, 25), "Bank debit", "Active", None, "Daily", "Keep"),
                ("Design Assets Plus", "Business tools", 16.5, "Monthly", dt.date(2025, 4, 11), "Mastercard ••5678", "Cancelled", None, "Never", "Cancel")]
        for k, row in enumerate(rows):
            for c, v in zip("ABCDEFGHIJ", row):
                if v is not None: su[f"{c}{S0 + k}"] = v
        su[f"K{S0 + 1}"] = "Cancel in Account > Membership"; su[f"K{S0 + 14}"] = "Trial – cancel before Nov 13 if not using"

    # ---------- Dashboard ----------
    SR = lambda c: f"Subscriptions!${c}${S0}:${c}${S1}"
    K.title(dash, "Subscription Dashboard", "Updates by itself from the Subscriptions tab.")
    dash["A3"] = '="Amounts in "&Settings!C14&"  ·  today: "&TEXT(TODAY(),"mmm d, yyyy")'; dash["A3"].font = font(bold=True)
    kp = [("Monthly cost", f"=SUM({SR('L')})", MONEY), ("Yearly cost", f"=SUM({SR('M')})", MONEY), ("Active subscriptions", f"=SUM({SR('T')})", "0"),
          ("Charges in next 30 days", f'=SUMIFS({SR("C")},{SR("O")},">=0",{SR("O")},"<=30")', MONEY),
          ("Rarely / never used (per year)", f'=SUMIFS({SR("M")},{SR("I")},"Rarely")+SUMIFS({SR("M")},{SR("I")},"Never")', MONEY),
          ("Save by cancelling (per year)", f'=SUMIFS({SR("M")},{SR("J")},"Cancel")', MONEY)]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 3) * 3, 1 + (i % 3) * 2, lab, f, fmt)
        dash.cell(row=6 + (i // 3) * 3, column=1 + (i % 3) * 2).alignment = Alignment(horizontal="left")
    for c, w in zip("ABCDEFGHI", [26, 13, 24, 13, 26, 13, 12, 14, 18]): dash.column_dimensions[c].width = w
    K.section(dash, "A11", "Cost by category")
    K.hdr(dash, 12, ["Category", "Monthly", "Yearly", "Active"])
    for i in range(NCAT):
        r = 13 + i
        dash[f"A{r}"] = f'=IF(Settings!F{CA0 + i}="","",Settings!F{CA0 + i})'
        dash[f"B{r}"] = f'=IF(A{r}="","",SUMIFS({SR("L")},{SR("B")},A{r}))'
        dash[f"C{r}"] = f'=IF(A{r}="","",B{r}*12)'
        dash[f"D{r}"] = f'=IF(A{r}="","",COUNTIFS({SR("B")},A{r},{SR("T")},1))'
    CE = 12 + NCAT
    style_range(dash, 13, CE, 1, 4, fmt={2: MONEY, 3: MONEY, 4: "0"})
    K.section(dash, "F11", "Next charges")
    K.hdr(dash, 12, ["Subscription", "Date", "In days", "Amount", "Paid with"], col0=6)
    dash.column_dimensions["F"].width = 24
    key = SR("U")
    for k in range(NCAT):
        r = 13 + k; pos = f"MATCH(SMALL({key},{k + 1}),{key},0)"
        dash[f"F{r}"] = f'=IFERROR(IF(SMALL({key},{k + 1})>=99999999,"",INDEX({SR("A")},{pos})),"")'
        for col, src in zip("GHIJ", "NOCF"):
            dash[f"{col}{r}"] = f'=IF(F{r}="","",INDEX({SR(src)},{pos}))'
    style_range(dash, 13, CE, 6, 10, fmt={7: "mmm d", 8: "0", 9: MONEY})
    dash.column_dimensions["J"].width = 18
    dash.conditional_formatting.add(f"H13:H{CE}", FormulaRule(formula=[f"AND(ISNUMBER(H13),H13<={WARN})"], font=Font(name=F, bold=True, color="B91C1C")))
    FR = CE + 3
    K.section(dash, f"A{FR}", "Next 12 months – money out")
    K.hdr(dash, FR + 1, ["Month", "Amount"])
    for j in range(12):
        r = FR + 2 + j; L = CL(4 + j)
        dash[f"A{r}"] = f"='12-Month Forecast'!{L}5"; dash[f"B{r}"] = f"='12-Month Forecast'!{L}6"
    style_range(dash, FR + 2, FR + 13, 1, 2, fmt={1: "mmm yyyy", 2: MONEY})
    dash[f"A{FR + 14}"] = "Total"; dash[f"B{FR + 14}"] = f"=SUM(B{FR + 2}:B{FR + 13})"
    for c in ("A", "B"): dash[f"{c}{FR + 14}"].font = font(bold=True)
    dash[f"B{FR + 14}"].number_format = MONEY
    K.section(dash, f"F{FR}", "Alerts")
    K.hdr(dash, FR + 1, ["Subscription", "Alert"], col0=6)
    akey = SR("V")
    for r in range(S0, S1 + 1): su[f"V{r}"] = f'=IF(P{r}="",99999999,ROW())'
    su.column_dimensions["V"].hidden = True
    for k in range(12):
        r = FR + 2 + k; pos = f"MATCH(SMALL({akey},{k + 1}),{akey},0)"
        dash[f"F{r}"] = f'=IFERROR(IF(SMALL({akey},{k + 1})>=99999999,"",INDEX({SR("A")},{pos})),"")'
        dash[f"G{r}"] = f'=IF(F{r}="","",INDEX({SR("P")},{pos}))'
    style_range(dash, FR + 2, FR + 13, 6, 7)
    ch = BarChart(); ch.type = "bar"; ch.title = "Monthly cost by category"; ch.height = 8; ch.width = 14; ch.legend = None
    ch.add_data(Reference(dash, min_col=2, min_row=12, max_row=12 + len(CATS)), titles_from_data=True)
    ch.set_categories(Reference(dash, min_col=1, min_row=13, max_row=12 + len(CATS)))
    ch.x_axis.scaling.orientation = "maxMin"; ch.y_axis.majorGridlines = None; ch.series[0].graphicalProperties.solidFill = MAIN
    dash.add_chart(ch, "L5")
    ch2 = BarChart(); ch2.type = "col"; ch2.title = "Next 12 months"; ch2.height = 7.5; ch2.width = 14; ch2.legend = None
    ch2.add_data(Reference(dash, min_col=2, min_row=FR + 1, max_row=FR + 13), titles_from_data=True)
    ch2.set_categories(Reference(dash, min_col=1, min_row=FR + 2, max_row=FR + 13)); ch2.y_axis.majorGridlines = None
    ch2.series[0].graphicalProperties.solidFill = "A78BFA"; ch2.gapWidth = 60
    dash.add_chart(ch2, "L22")

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Settings tab: check the categories and type your cards / payment methods. Optional: currency label and how many days' warning you want.", False),
             ("2.  Subscriptions tab: one row per subscription – name, category, cost per charge, billing cycle and the date of the FIRST charge.", False),
             ("     The next charge date, days until it, monthly and yearly cost are calculated for you (weekly, every 2 weeks, monthly, 3 / 6 months, yearly).", False),
             ("3.  Free trial? Set Status to 'Free trial' and type the date it ends – you get a warning before it turns into a paid plan.", False),
             ("4.  Be honest in 'How often I use it' and pick a Verdict (Keep, Review, Cancel). The Dashboard shows how much you would save.", False),
             ("5.  Cancelled something? Set Status to 'Cancelled' (or 'Paused') – it stops counting but stays in your history.", False),
             ("6.  12-Month Forecast tab: what will leave your account each month, per subscription.", False),
             ("7.  Dashboard: monthly and yearly total, charges in the next 30 days, cost by category, next charges, alerts and charts – all automatic.", False),
             ("", False), ("Tips", True),
             ("•  Go through your bank and card statements for the last 3 months – most people find 2–3 subscriptions they forgot about.", False),
             ("•  Yearly plans are often 15–20% cheaper than monthly – compare the 'Yearly cost' column before you switch.", False),
             ("•  Paste the cancel link or steps in Notes so cancelling takes one minute.", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  150 subscriptions, 15 categories, 8 payment methods. No macros. Works in Excel 2010+ and Google Sheets.", False),
             ("•  Dates update with today's date every time you open the file.", False),
             ("•  To use Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("The services and prices are made-up examples. Use the blank copy in your download for your own subscriptions.", False)]
    K.start(st, "Subscription Tracker", "Renewal dates, free-trial alerts and what you really spend – Excel and Google Sheets.", lines)
    finish(wb, sample, path)


META = {
    "name": "Subscription Tracker (Renewal Dates, Free Trial Alerts, Monthly & Yearly Cost, 12-Month Forecast) – Excel & Google Sheets",
    "file_base": "Subscription-Tracker",
    "price": 9,
    "color": "#7C3AED",
    "title": "Subscription Tracker",
    "tagline": "next charge dates · free-trial alerts · yearly cost · cancel list",
    "intro": "See every subscription in one place and stop paying for things you don't use. Type the cost, billing cycle and first charge date – the tracker works out the next charge date, warns you before renewals and free trials end, converts everything to monthly and yearly cost, and shows how much you would save by cancelling the ones you rarely use.",
    "features": [
        "Track 150 subscriptions: category, cost, billing cycle, payment method, status, free-trial end date and notes",
        "Automatic next charge date and days until it – weekly, every 2 weeks, monthly, every 3 or 6 months, yearly",
        "Alerts: free trial ends soon, charge coming in X days, cancel to save X a year",
        "'How often I use it' and Keep / Review / Cancel verdict to find money you are wasting",
        "12-month forecast of what leaves your account each month, per subscription",
        "Dashboard: monthly and yearly total, next 30 days, cost by category, next charges, alerts and charts",
        "Blank copy plus a sample copy filled with example subscriptions",
    ],
    "tags": ["subscription tracker", "subscription spreadsheet", "recurring payments", "bill tracker", "budget template", "free trial tracker",
             "google sheets", "excel template", "personal finance", "expense tracker", "renewal tracker", "money saving", "budget planner"],
    "summary": "Instant download: Excel + Google Sheets template, blank copy + sample copy.",
}
