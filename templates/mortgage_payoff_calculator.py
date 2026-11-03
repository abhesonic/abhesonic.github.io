"""Mortgage Payoff & Extra Payment Calculator – full amortization with extra monthly, yearly, bi-weekly and lump-sum payments.

Loan (inputs) + Lump Sums -> Schedule (480 months, with extras; baseline balance by closed form) -> Yearly (summary)
-> Dashboard (payoff date, months & interest saved, PMI drop-off, what-if table, charts).
"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.utils import get_column_letter as L
from _common import Kit, F, B, MONEY, DATE, font, fill, style_range, input_cell, finish

MAIN = "1E40AF"; SOFT = "E8EEFB"
K = Kit(MAIN, SOFT)
PCT = "0.0%"; PCT3 = "0.000%"
NM = 480; S0 = 6; S1 = S0 + NM - 1          # schedule rows
NY = 40; Y0 = 6; Y1 = Y0 + NY - 1           # yearly rows
NL = 30; LS0 = 6; LS1 = LS0 + NL - 1        # lump sum rows
SCH = "Schedule"; LN = "Loan"; LSS = "'Lump Sums'"

# key, label, fmt, help, sample, blank
INP = [
    ("sec", "YOUR LOAN"),
    ("amt", "Loan amount (or current balance)", MONEY, "New loan: the amount you borrowed. Existing loan: today's balance from your statement.", 320000, 250000),
    ("rate", "Interest rate (yearly)", PCT3, "From your loan papers, e.g. 6.5%.", 0.0649, 0.065),
    ("yrs", "Loan term in years (or years left)", "0", "Usually 30 or 15. For an existing loan: the years still left.", 30, 30),
    ("start", "First payment month", "mmm yyyy", "The month of the first payment in this schedule (any day in that month works).", dt.date(2026, 1, 1), "=DATE(YEAR(TODAY()),MONTH(TODAY()),1)"),
    ("sec", "EXTRA PAYMENTS (leave at 0 to see your normal schedule)"),
    ("xm", "Extra every month", MONEY, "Added to the principal each month.", 200, 0),
    ("xms", "…starting with payment #", "0", "1 = from the first payment. Use a later number if you start extra payments later.", 1, 1),
    ("xy", "Extra once a year", MONEY, "For example part of a tax refund or a yearly bonus.", 1500, 0),
    ("xym", "…in month (1–12)", "0", "1 = January, 12 = December.", 4, 12),
    ("bw", "Bi-weekly payments? (Yes/No)", None, "Paying half every 2 weeks = 13 full payments a year. Adds 1/12 of a payment each month.", "No", "No"),
    ("sec", "HOUSING COSTS & PMI (optional)"),
    ("home", "Home value (when you bought it)", MONEY, "Used for PMI and your loan-to-value. Leave empty if you have no PMI.", 360000, None),
    ("pmi", "PMI per month", MONEY, "Private mortgage insurance. Stops when the balance reaches the PMI-free % below.", 145, 0),
    ("pmil", "PMI stops at loan-to-value of", PCT, "Usually 80% (lender request) or 78% (automatic).", 0.78, 0.78),
    ("tax", "Property tax per year", MONEY, "", 4200, None),
    ("ins", "Home insurance per year", MONEY, "", 1650, None),
    ("hoa", "HOA per month", MONEY, "", 0, None),
]
OUT = [
    ("sec", "RESULTS (CALCULATED)"),
    ("r", "Monthly interest rate", "0.0000%", "Yearly rate ÷ 12."),
    ("n", "Number of payments (normal)", "0", "Years × 12."),
    ("pi", "Monthly payment (principal + interest)", MONEY, ""),
    ("bwx", "Extra from bi-weekly payments (per month)", MONEY, "1/12 of a monthly payment, when bi-weekly is Yes."),
    ("all", "Total monthly housing cost", MONEY, "Payment + extra + tax + insurance + HOA + PMI (first month)."),
]


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); ln = wb.create_sheet(LN); lsum = wb.create_sheet("Lump Sums")
    sch = wb.create_sheet(SCH); yr = wb.create_sheet("Yearly")
    ref = {}

    # ---------- Loan ----------
    K.title(ln, "Loan details", "Yellow cells are yours to edit. Grey cells calculate themselves.")
    r = 4
    for row in INP + OUT:
        if row[0] == "sec":
            r += 1 if r > 4 else 0; K.section(ln, f"A{r}", row[1]); r += 1; continue
        key, lab, fmt, hlp = row[:4]
        ln[f"A{r}"] = lab; ln[f"A{r}"].font = font(bold=True)
        c = ln[f"B{r}"]
        if len(row) > 4:
            v = row[4] if sample else row[5]
            input_cell(c, v)
        else:
            c.border = B; c.fill = fill("F3F4F6"); c.font = font(bold=True)
        if fmt: c.number_format = fmt
        if hlp: ln[f"C{r}"] = hlp; ln[f"C{r}"].font = font(italic=True, size=9)
        ref[key] = f"{LN}!$B${r}"; r += 1
    loc = {k: v.replace(f"{LN}!", "") for k, v in ref.items()}
    ln[loc["r"].replace("$", "")] = f"={loc['rate']}/12"
    ln[loc["n"].replace("$", "")] = f"=ROUND({loc['yrs']}*12,0)"
    ln[loc["pi"].replace("$", "")] = f"=IFERROR(ROUND(IF({loc['r']}=0,{loc['amt']}/{loc['n']},PMT({loc['r']},{loc['n']},-{loc['amt']})),2),0)"
    ln[loc["bwx"].replace("$", "")] = f'=IF({loc["bw"]}="Yes",ROUND({loc["pi"]}/12,2),0)'
    ln[loc["all"].replace("$", "")] = (f"={loc['pi']}+{loc['xm']}*({loc['xms']}<=1)+{loc['bwx']}+({loc['tax']}+{loc['ins']})/12+{loc['hoa']}"
                                       f"+IF(AND({loc['home']}>0,{loc['amt']}/MAX(1,{loc['home']})>{loc['pmil']}),{loc['pmi']},0)")
    dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=False); ln.add_data_validation(dv); dv.add(loc["bw"].replace("$", ""))
    for c, w in zip("ABC", [40, 18, 92]): ln.column_dimensions[c].width = w

    # ---------- Lump Sums ----------
    K.title(lsum, "Lump-sum payments", "One-time extra payments (bonus, inheritance, sale of a car…). Type the month and the amount.")
    K.hdr(lsum, 5, ["Month paid", "Amount", "Note", "Payment #"], [16, 16, 34, 12])
    style_range(lsum, LS0, LS1, 1, 4, inputs=(1, 2, 3), fmt={1: "mmm yyyy", 2: MONEY, 4: "0"})
    for i in range(NL):
        rr = LS0 + i
        lsum[f"D{rr}"] = (f'=IF(OR(A{rr}="",B{rr}=""),"",IF((YEAR(A{rr})-YEAR({ref["start"]}))*12+MONTH(A{rr})-MONTH({ref["start"]})+1<1,"before start",'
                          f'(YEAR(A{rr})-YEAR({ref["start"]}))*12+MONTH(A{rr})-MONTH({ref["start"]})+1))')
        lsum[f"D{rr}"].alignment = Alignment(horizontal="center")
    if sample:
        for i, (d, a, n) in enumerate([(dt.date(2027, 3, 1), 5000, "Tax refund + bonus"), (dt.date(2029, 6, 1), 10000, "Sold second car"),
                                       (dt.date(2032, 12, 1), 8000, "Year-end bonus")]):
            lsum[f"A{LS0 + i}"] = d; lsum[f"B{LS0 + i}"] = a; lsum[f"C{LS0 + i}"] = n
    lsum[f"A{LS1 + 2}"] = "Total lump sums"; lsum[f"A{LS1 + 2}"].font = font(bold=True)
    lsum[f"B{LS1 + 2}"] = f"=SUM(B{LS0}:B{LS1})"; lsum[f"B{LS1 + 2}"].number_format = MONEY; lsum[f"B{LS1 + 2}"].font = font(bold=True)
    lsum["F5"] = "Tip: the payment # shows which month of your schedule the lump sum lands in."; lsum["F5"].font = font(italic=True, size=9)

    # ---------- Schedule ----------
    K.title(sch, "Amortization schedule", "Calculated from the Loan and Lump Sums tabs. Nothing to type here.")
    cols = ["#", "Month", "Start balance", "Interest", "Regular principal", "Extra payment", "Total paid", "End balance",
            "Total interest so far", "PMI", "Normal balance (no extras)", "Loan-to-value"]
    K.hdr(sch, 5, cols, [6, 11, 15, 12, 14, 13, 13, 15, 15, 9, 17, 11])
    R = ref
    for i in range(NM):
        rr = S0 + i; k = i + 1
        sch[f"A{rr}"] = k
        sch[f"B{rr}"] = f"=DATE(YEAR({R['start']}),MONTH({R['start']})+{i},1)"
        sch[f"C{rr}"] = f"={R['amt']}" if i == 0 else f"=H{rr - 1}"
        sch[f"D{rr}"] = f"=IF(C{rr}<=0.005,0,ROUND(C{rr}*{R['r']},2))"
        sch[f"E{rr}"] = f"=IF(C{rr}<=0.005,0,IF(A{rr}>={R['n']},C{rr},MIN({R['pi']}-D{rr},C{rr})))"
        extra = (f"IF(A{rr}>={R['xms']},{R['xm']},0)+IF(MONTH(B{rr})={R['xym']},{R['xy']},0)+{R['bwx']}"
                 f"+SUMIFS({LSS}!$B${LS0}:$B${LS1},{LSS}!$A${LS0}:$A${LS1},\">=\"&B{rr},{LSS}!$A${LS0}:$A${LS1},\"<\"&DATE(YEAR(B{rr}),MONTH(B{rr})+1,1))")
        sch[f"F{rr}"] = f"=IF(C{rr}<=0.005,0,MAX(0,MIN({extra},C{rr}-E{rr})))"
        sch[f"G{rr}"] = f"=D{rr}+E{rr}+F{rr}"
        sch[f"H{rr}"] = f"=MAX(0,C{rr}-E{rr}-F{rr})"
        sch[f"I{rr}"] = f"=D{rr}" if i == 0 else f"=I{rr - 1}+D{rr}"
        sch[f"J{rr}"] = f"=IF(AND(C{rr}>0.005,{R['home']}>0,C{rr}/MAX(1,{R['home']})>{R['pmil']}),{R['pmi']},0)"
        sch[f"K{rr}"] = (f"=IF(A{rr}>={R['n']},0,MAX(0,IF({R['r']}=0,{R['amt']}-{R['pi']}*A{rr},"
                         f"{R['amt']}*(1+{R['r']})^A{rr}-{R['pi']}*((1+{R['r']})^A{rr}-1)/{R['r']})))")
        sch[f"L{rr}"] = f"=IF({R['home']}>0,H{rr}/{R['home']},0)"
    style_range(sch, S0, S1, 1, 12, fmt={2: "mmm yyyy", **{c: MONEY for c in range(3, 12)}, 12: PCT})
    sch.conditional_formatting.add(f"A{S0}:L{S1}", FormulaRule(formula=[f"$C{S0}<=0.005"], font=Font(name=F, color="C0C6CC")))
    sch.conditional_formatting.add(f"F{S0}:F{S1}", FormulaRule(formula=[f"$F{S0}>0"], font=Font(name=F, bold=True, color="15803D")))
    sch.freeze_panes = f"C{S0}"

    # ---------- Yearly ----------
    K.title(yr, "Year by year", "Summary of the schedule per loan year (12 payments).")
    K.hdr(yr, 5, ["Loan year", "Calendar year (end)", "Interest paid", "Principal paid", "Extra paid", "Your plan balance", "Normal balance (no extras)", "Ahead of schedule by"],
          [10, 18, 14, 15, 13, 15, 22, 18])
    for i in range(NY):
        rr = Y0 + i; a = S0 + i * 12; b = a + 11
        yr[f"A{rr}"] = i + 1
        yr[f"B{rr}"] = f"=YEAR({SCH}!B{b})"
        yr[f"C{rr}"] = f"=SUM({SCH}!D{a}:D{b})"
        yr[f"D{rr}"] = f"=SUM({SCH}!E{a}:E{b})+SUM({SCH}!F{a}:F{b})"
        yr[f"E{rr}"] = f"=SUM({SCH}!F{a}:F{b})"
        yr[f"F{rr}"] = f"={SCH}!H{b}"
        yr[f"G{rr}"] = f"={SCH}!K{b}"
        yr[f"H{rr}"] = f"=G{rr}-F{rr}"
    style_range(yr, Y0, Y1, 1, 8, fmt={2: "0", **{c: MONEY for c in range(3, 9)}})
    yr.conditional_formatting.add(f"A{Y0}:H{Y1}", FormulaRule(formula=[f"AND($F{Y0}=0,$G{Y0}=0)"], font=Font(name=F, color="C0C6CC")))
    yr.freeze_panes = f"A{Y0}"

    # ---------- Dashboard ----------
    K.title(dash, "Mortgage Payoff Dashboard", "Updates by itself from the Loan and Lump Sums tabs.")
    SC = f"{SCH}!$C${S0}:$C${S1}"
    dash["A3"] = (f'=IF({R["amt"]}<=0,"Type your loan on the Loan tab to get started.",TEXT({R["amt"]},"#,##0")&" at "&TEXT({R["rate"]},"0.00%")&" over "&{R["yrs"]}&" years  ·  first payment "&TEXT({R["start"]},"mmm yyyy"))')
    dash["A3"].font = font(bold=True)
    # helper cells (hidden in white): months with extras, normal months
    dash["H3"] = "m"; dash["I3"] = f'=COUNTIF({SC},">0.005")'
    dash["J3"] = "n"; dash["K3"] = f"={R['n']}"
    for c in "HIJK": dash[f"{c}3"].font = font(size=8, color="FFFFFF")
    payoff_new = f"DATE(YEAR({R['start']}),MONTH({R['start']})+$I$3-1,1)"
    payoff_old = f"DATE(YEAR({R['start']}),MONTH({R['start']})+$K$3-1,1)"
    int_new = f"SUM({SCH}!$D${S0}:$D${S1})"
    int_old = f"MAX(0,{R['pi']}*{R['n']}-{R['amt']})"
    kp = [("Monthly payment (P&I)", f"={R['pi']}", MONEY),
          ("New payoff date", f"=IF($I$3=0,\"–\",{payoff_new})", "mmm yyyy"),
          ("Time saved", '=IF($I$3=0,"–",INT(MAX(0,$K$3-$I$3)/12)&" yr "&MOD(MAX(0,$K$3-$I$3),12)&" mo")', "@"),
          ("Total interest (with extras)", f"={int_new}", MONEY),
          ("Interest saved", f"=MAX(0,{int_old}-{int_new})", MONEY),
          ("Paid off so far (today)", f"=IFERROR(1-INDEX({SC},MAX(1,MIN({NM},(YEAR(TODAY())-YEAR({R['start']}))*12+MONTH(TODAY())-MONTH({R['start']})+1)))/{R['amt']},0)", PCT)]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 3) * 3, 1 + (i % 3) * 2, lab, f, fmt)
        dash.cell(row=6 + (i // 3) * 3, column=1 + (i % 3) * 2).alignment = Alignment(horizontal="left")
    for c, w in zip("ABCDEFGHIJK", [40, 18, 26, 22, 26, 18, 4, 4, 4, 4, 4]): dash.column_dimensions[c].width = w

    K.section(dash, "A11", "Normal schedule vs your plan")
    K.hdr(dash, 12, ["", "Normal", "With your extra payments", "Difference"])
    cmp_rows = [("Payoff date", f"={payoff_old}", f"=IF($I$3=0,\"–\",{payoff_new})", '=IF($I$3=0,"–",MAX(0,$K$3-$I$3)&" months sooner")', "mmm yyyy"),
                ("Number of payments", "=$K$3", "=$I$3", "=B14-C14", "0"),
                ("Total interest", f"={int_old}", f"={int_new}", "=MAX(0,B15-C15)", MONEY),
                ("Total paid (principal + interest)", f"={R['amt']}+B15", f"={R['amt']}+C15", "=MAX(0,B16-C16)", MONEY),
                ("Total extra payments", "=0", f"=SUM({SCH}!$F${S0}:$F${S1})", "=C17-B17", MONEY),
                ("Interest saved per $1 of extra", "", "=IFERROR(D15/C17,0)", "", "0.00")]
    for i, (lab, b, c_, d, fmt) in enumerate(cmp_rows):
        rr = 13 + i
        dash[f"A{rr}"] = lab; dash[f"B{rr}"] = b or None; dash[f"C{rr}"] = c_; dash[f"D{rr}"] = d or None
    style_range(dash, 13, 18, 1, 4)
    for rr, fmt in zip(range(13, 19), [r_[4] for r_ in cmp_rows]):
        for c in "BCD": dash[f"{c}{rr}"].number_format = fmt
        dash[f"A{rr}"].font = font(bold=True); dash[f"D{rr}"].font = Font(name=F, bold=True, color="15803D")
    dash["D13"].number_format = "@"

    K.section(dash, "A20", "Monthly cost & PMI")
    mc = [("Payment + extras (first month)", f"={SCH}!G{S0}", MONEY),
          ("Tax + insurance + HOA per month", f"=({R['tax']}+{R['ins']})/12+{R['hoa']}", MONEY),
          ("Total monthly housing cost", f"={R['all']}", MONEY),
          ("PMI stops (your plan)", f'=IF(OR({R["pmi"]}<=0,{R["home"]}<=0),"No PMI",IFERROR(INDEX({SCH}!$B${S0}:$B${S1},MATCH(0,{SCH}!$J${S0}:$J${S1},0)),"–"))', "mmm yyyy"),
          ("Total PMI paid (your plan)", f"=SUM({SCH}!$J${S0}:$J${S1})", MONEY),
          ("Loan-to-value today", f"=IFERROR(INDEX({SCH}!$C${S0}:$C${S1},MAX(1,MIN({NM},(YEAR(TODAY())-YEAR({R['start']}))*12+MONTH(TODAY())-MONTH({R['start']})+1)))/{R['home']},0)", PCT)]
    for i, (lab, f, fmt) in enumerate(mc):
        rr = 21 + i
        dash[f"A{rr}"] = lab; dash[f"B{rr}"] = f
    style_range(dash, 21, 26, 1, 2)
    for i, (_, _, fmt) in enumerate(mc):
        dash[f"B{21 + i}"].number_format = fmt; dash[f"A{21 + i}"].font = font(bold=True); dash[f"B{21 + i}"].font = Font(name=F, bold=True, color=MAIN)

    K.section(dash, "A28", "What if you paid a fixed extra amount every month from the start?")
    K.hdr(dash, 29, ["Extra per month", "Payoff in (years)", "Total interest", "Interest saved", "Payoff month"])
    WI = 30
    for i, x in enumerate([0, 50, 100, 200, 300, 500, 750, 1000]):
        rr = WI + i
        input_cell(dash[f"A{rr}"], x); dash[f"A{rr}"].number_format = MONEY
        nper = f"IF({R['r']}=0,{R['amt']}/({R['pi']}+A{rr}),NPER({R['r']},-({R['pi']}+A{rr}),{R['amt']}))"
        dash[f"B{rr}"] = f"=IFERROR({nper}/12,0)"
        dash[f"C{rr}"] = f"=IFERROR(MAX(0,({R['pi']}+A{rr})*{nper}-{R['amt']}),0)"
        dash[f"D{rr}"] = f"=MAX(0,$C${WI}-C{rr})"
        dash[f"E{rr}"] = f"=IFERROR(DATE(YEAR({R['start']}),MONTH({R['start']})+ROUNDUP(ROUND({nper},2),0)-1,1),\"–\")"
        for c, fmt in zip("BCDE", ["0.0", MONEY, MONEY, "mmm yyyy"]):
            cell = dash[f"{c}{rr}"]; cell.number_format = fmt; cell.border = B; cell.fill = fill("F3F4F6"); cell.font = font()
    dash[f"A{WI + 9}"] = "Yellow amounts can be changed. Estimates assume no lump sums, no yearly extra and no bi-weekly payments."
    dash[f"A{WI + 9}"].font = font(italic=True, size=9)

    # charts
    ch = LineChart(); ch.title = "Loan balance: normal vs your plan"; ch.height = 8; ch.width = 18
    ch.add_data(Reference(yr, min_col=6, min_row=5, max_row=Y0 + 34), titles_from_data=True)
    ch.add_data(Reference(yr, min_col=7, min_row=5, max_row=Y0 + 34), titles_from_data=True)
    ch.set_categories(Reference(yr, min_col=1, min_row=Y0, max_row=Y0 + 34))
    ch.series[0].graphicalProperties.line.solidFill = MAIN; ch.series[0].graphicalProperties.line.width = 28000
    ch.series[1].graphicalProperties.line.solidFill = "9CA3AF"; ch.series[1].graphicalProperties.line.dashStyle = "dash"
    ch.y_axis.majorGridlines = None; ch.x_axis.title = "Loan year"; ch.legend.position = "b"
    dash.add_chart(ch, "F11")
    ch2 = BarChart(); ch2.type = "col"; ch2.grouping = "stacked"; ch2.overlap = 100; ch2.title = "Interest vs principal per year (your plan)"
    ch2.height = 8; ch2.width = 18
    ch2.add_data(Reference(yr, min_col=3, max_col=4, min_row=5, max_row=Y0 + 34), titles_from_data=True)
    ch2.set_categories(Reference(yr, min_col=1, min_row=Y0, max_row=Y0 + 34))
    for ser, col in zip(ch2.series, ("F59E0B", MAIN)): ser.graphicalProperties.solidFill = col
    ch2.y_axis.majorGridlines = None; ch2.legend.position = "b"
    dash.add_chart(ch2, "F28")

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Loan tab: type your loan amount, interest rate, term and the month of the first payment. Your monthly payment is calculated for you.", False),
             ("2.  Still in the Loan tab: try an extra amount every month, a yearly extra (tax refund, bonus) or switch bi-weekly payments to Yes.", False),
             ("3.  Lump Sums tab (optional): add one-time extra payments with the month you plan to pay them.", False),
             ("4.  Dashboard: see your new payoff date, how many years you save and how much interest you keep in your pocket.", False),
             ("5.  Schedule and Yearly tabs: the full month-by-month amortization table and a summary per year.", False),
             ("", False), ("Already paying a mortgage?", True),
             ("Type today's balance as the loan amount, the years left as the term, and next month as the first payment month. Your payment may differ by a few cents from your lender's.", False),
             ("", False), ("Tips", True),
             ("•  Extra payments early in the loan save the most interest – look at 'Interest saved per $1 of extra' on the Dashboard.", False),
             ("•  Ask your lender to apply extra money to the principal, and check there is no prepayment penalty.", False),
             ("•  The what-if table on the Dashboard compares 8 extra amounts at once. Change the yellow amounts to test your own.", False),
             ("•  PMI: type your home value and PMI to see when PMI stops and what you pay in total.", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  Handles loans up to 40 years, any currency and fixed interest rates. Estimates for planning only – your lender's statement is the final word.", False),
             ("•  To use Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("The loan, extra payments and lump sums are made-up examples. Use the blank copy in your download for your own loan.", False)]
    K.start(st, "Mortgage Payoff & Extra Payment Calculator", "Pay your home off years sooner – see exactly how much time and interest every extra payment saves.", lines)
    finish(wb, sample, path)


META = {
    "name": "Mortgage Payoff Calculator (Extra Payments, Lump Sums, Bi-Weekly, Amortization Schedule) – Excel & Google Sheets",
    "file_base": "Mortgage-Payoff-Calculator",
    "price": 17,
    "color": "#1E40AF",
    "title": "Mortgage Payoff Calculator",
    "tagline": "Extra payments · lump sums · bi-weekly · interest saved · full amortization",
    "intro": "See how fast you can be mortgage-free. Type your loan once, then try extra monthly payments, a yearly bonus payment, bi-weekly payments or one-time lump sums – the calculator shows your new payoff date, the years you save and the interest you keep.",
    "features": [
        "Full amortization schedule (up to 40 years) with extra monthly, yearly, bi-weekly and lump-sum payments",
        "New payoff date, time saved (years + months) and total interest saved – calculated for you",
        "Side-by-side: normal schedule vs your plan, plus 'interest saved per $1 of extra'",
        "What-if table: 8 extra amounts compared at once (payoff years, interest, payoff month)",
        "PMI drop-off date, loan-to-value and total monthly housing cost (tax, insurance, HOA)",
        "Year-by-year summary and charts: balance normal vs your plan, interest vs principal",
        "Works for a new or an existing mortgage, any currency, no macros, blank copy plus a sample copy",
    ],
    "tags": ["mortgage calculator", "mortgage payoff", "extra payment calculator", "amortization schedule", "pay off mortgage early",
             "biweekly mortgage", "loan calculator", "google sheets", "excel template", "debt free", "home loan", "personal finance", "mortgage tracker"],
}
