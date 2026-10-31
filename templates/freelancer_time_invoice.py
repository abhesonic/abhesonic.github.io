"""Freelancer Time Tracker & Invoice Generator (sample + blank versions).

Clients (rate, payment terms) + Projects (optional rate, budget hours) + Time Log (start/end or hours, billable, invoice #)
-> hours and amount per entry, unbilled work, project budget used.
Invoices list (issue date, extra items, discount, paid date) -> totals with tax, due date, status (Sent / Overdue / Paid).
Printable Invoice tab: pick an invoice number and it fills itself from the time log.
Dashboard: hours, billable %, earned, unbilled, outstanding, overdue, paid this year, effective hourly rate, charts.
"""
import random, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.chart import BarChart, LineChart, Reference
from _common import Kit, F, B, DATE, font, fill, style_range, input_cell, finish

MAIN = "0F766E"; SOFT = "E6F4F1"; ACC = "5EEAD4"
K = Kit(MAIN, SOFT)
NC = 100; C0 = 6; C1 = C0 + NC - 1
NJ = 150; J0 = 6; J1 = J0 + NJ - 1
NT = 3000; T0 = 6; T1 = T0 + NT - 1
NI = 300; I0 = 6; I1 = I0 + NI - 1
NL = 30  # invoice lines
CUR = '#,##0.00;-#,##0.00;"-"'; HRS = '#,##0.00;-#,##0.00;"-"'; PCT = '0%;-0%;"-"'; TIME = "hh:mm"
TL = "'Time Log'"
def tl(c): return f"{TL}!${c}${T0}:${c}${T1}"
def cl(c): return f"Clients!${c}${C0}:${c}${C1}"
def pj(c): return f"Projects!${c}${J0}:${c}${J1}"
def iv(c): return f"Invoices!${c}${I0}:${c}${I1}"
AS = "Settings!$B$5"; TAXL = "Settings!$B$11"; TAX = "Settings!$B$12"; TERMS = "Settings!$B$13"; DRATE = "Settings!$B$14"; RND = "Settings!$B$15"
YS = "Settings!$B$18"; MS = "Settings!$B$19"; WS = "Settings!$B$20"

CLIENTS = [("Northwind Coffee Co.", "Hannah Lee", "hannah@northwind.example", "18 Harbor St, Portland, OR 97201", 65, 14),
           ("Bright Path Coaching", "Marcus Diaz", "marcus@brightpath.example", "402 Elm Ave, Austin, TX 78701", 60, 30),
           ("Kettle & Crumb Bakery", "Olivia Brown", "olivia@kettlecrumb.example", "9 Mill Lane, Bristol BS1 4DJ, UK", 55, 14),
           ("Atlas Fitness Studio", "Ryan Patel", "ryan@atlasfit.example", "77 King St W, Toronto, ON M5H 1J9", 70, 30),
           ("Greenleaf Realty", "Sofia Rossi", "sofia@greenleaf.example", "250 Market St, Denver, CO 80202", 75, 30)]
PROJECTS = [("Brand refresh", 0, None, 40, "Done"), ("Monthly social graphics", 0, 55, None, "Active"),
            ("Course workbook design", 1, None, 30, "Active"), ("Website landing page", 1, 80, 20, "Done"),
            ("Menu & signage", 2, None, 18, "Done"), ("Packaging labels", 2, None, 24, "Active"),
            ("Class schedule posters", 3, None, None, "Active"), ("Listing brochure template", 4, None, 25, "Active")]
TASKS = ["Concept sketches", "Design round 1", "Client call", "Revisions", "Final files & export", "Research", "Layout", "Feedback changes"]


def simulate():
    random.seed(11)
    logs = []; d = dt.date(2026, 3, 2)
    while d <= dt.date(2026, 10, 30):
        if d.weekday() < 5:
            for _ in range(random.choice([1, 2, 2, 3])):
                pi = random.choice([1, 1, 2, 3, 5, 6, 7, 0] if d.month < 9 else [1, 1, 2, 5, 6, 7, 7])
                if PROJECTS[pi][4] == "Done" and d.month >= 9: continue
                start = random.choice([8.5, 9, 9.5, 10, 13, 13.5, 14, 15])
                dur = random.choice([0.75, 1, 1.5, 2, 2.5, 3])
                brk = 15 if dur >= 2.5 else 0
                bill = "No" if random.random() < 0.08 else "Yes"
                logs.append([d, CLIENTS[PROJECTS[pi][1]][0], PROJECTS[pi][0], random.choice(TASKS),
                             dt.time(int(start), int((start % 1) * 60)),
                             dt.time(int(start + dur + brk / 60), int(((start + dur + brk / 60) % 1) * 60)), brk or None, None, bill, None])
            if random.random() < 0.1:
                logs.append([d, "", "", "Admin, email & bookkeeping", None, None, None, 1, "No", None])
        d += dt.timedelta(days=1)
    # invoices: one per client per month for Jul–Sep
    invs = []; n = 1
    for m in range(3, 10):
        for ci, c in enumerate(CLIENTS):
            rows = [l for l in logs if l[1] == c[0] and l[0].month == m and l[8] == "Yes"]
            if not rows: continue
            no = f"INV-2026-{n:03d}"; n += 1
            for l in rows: l[9] = no
            issue = dt.date(2026, m + 1, 1) + dt.timedelta(days=1)
            paid = None
            if m <= 8 and not (ci == 4 and m == 8): paid = issue + dt.timedelta(days=random.randint(5, c[5] + 3))
            elif ci in (0, 2): paid = issue + dt.timedelta(days=random.randint(6, 12))
            extra = 45.0 if (ci == 2 and m == 8) else None
            disc = 50.0 if (ci == 3 and m == 7) else None
            note = "Stock photo licence" if extra else ("Loyalty discount" if disc else None)
            invs.append((no, c[0], issue, None, extra, disc, paid, note))
    return logs, invs


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); tsh = wb.create_sheet("Time Log"); ish = wb.create_sheet("Invoices"); inv = wb.create_sheet("Invoice")
    csh = wb.create_sheet("Clients"); jsh = wb.create_sheet("Projects"); cc = wb.create_sheet("Calc"); s = wb.create_sheet("Settings")
    logs, invs = simulate() if sample else ([], [])

    # ---------- Settings ----------
    K.title(s, "Settings", "Yellow cells are yours to edit. Your details appear on every invoice.")
    rows = [("Currency", "USD", None, "Shown on invoices and the dashboard."),
            ("Today / as-of date", dt.date(2026, 10, 30) if sample else "=TODAY()", DATE, "Keep =TODAY() so 'this week / month' and overdue invoices stay current."),
            ("Your name / business", "Alex Rivera Design" if sample else None, None, ""),
            ("Address", "120 Pine Street, Suite 4, Seattle, WA 98101" if sample else None, None, ""),
            ("Email", "hello@alexrivera.example" if sample else None, None, ""),
            ("Phone / website", "+1 555 0123 · alexrivera.example" if sample else None, None, ""),
            ("Payment details", "Bank transfer: Example Bank · Account 000123456 · Routing 000000000 · or PayPal hello@alexrivera.example" if sample else None, None, "How clients should pay you."),
            ("Tax name", "Tax", None, "e.g. VAT, GST, Sales tax. Leave the rate at 0% if you don't charge tax."),
            ("Tax rate", 0, PCT, ""),
            ("Default payment terms (days)", 30, "0", "Used when a client has no terms of their own."),
            ("Default hourly rate", 60, CUR, "Used when the client and project have no rate."),
            ("Round time up to (minutes)", 0, "0", "0 = exact. 6, 15 or 30 rounds every entry up."),
            ("Tax ID / registration (optional)", None, None, "")]
    for i, (k_, v, fmt, note) in enumerate(rows, 4):
        s.cell(row=i, column=1, value=k_).font = font(bold=True)
        c = s.cell(row=i, column=2); input_cell(c, v)
        if fmt: c.number_format = fmt
        s.cell(row=i, column=3, value=note).font = font(italic=True, size=9)
    for i, (k_, f) in enumerate([("Year starts", f"=DATE(YEAR({AS}),1,1)"), ("This month starts", f"=DATE(YEAR({AS}),MONTH({AS}),1)"),
                                 ("This week starts (Monday)", f"={AS}-WEEKDAY({AS},2)+1")], 18):
        s.cell(row=i, column=1, value=k_).font = font(italic=True)
        c = s.cell(row=i, column=2, value=f); c.number_format = DATE; c.border = B; c.fill = fill("F3F4F6"); c.font = font()
    for c, w in zip("ABC", [34, 60, 60]): s.column_dimensions[c].width = w

    # ---------- Clients ----------
    K.title(csh, "Clients", "One row per client. Grey columns calculate themselves.")
    K.hdr(csh, 5, ["Client", "Contact person", "Email", "Billing address", "Hourly rate", "Payment terms (days)", "Notes",
                   "Hours (all)", "Invoiced", "Paid", "Outstanding", "Unbilled work", "Billable this year", "Rank key"],
          [26, 16, 26, 34, 9, 9, 16, 9, 11, 11, 11, 11, 11, 7])
    style_range(csh, C0, C1, 1, 14, inputs=range(1, 8), fmt={5: CUR, 6: "0", 8: HRS, 9: CUR, 10: CUR, 11: CUR, 12: CUR, 13: CUR, 14: "0.000"})
    if sample:
        for i, c in enumerate(CLIENTS):
            for j, v in enumerate(c, 1): csh.cell(row=C0 + i, column=j, value=v)
    for r in range(C0, C1 + 1):
        on = f'A{r}=""'
        csh[f"H{r}"] = f'=IF({on},0,SUMIFS({tl("K")},{tl("B")},A{r}))'
        csh[f"I{r}"] = f'=IF({on},0,SUMIFS({iv("M")},{iv("B")},A{r}))'
        csh[f"J{r}"] = f'=IF({on},0,SUMIFS({iv("M")},{iv("B")},A{r},{iv("O")},"Paid"))'
        csh[f"K{r}"] = f"=I{r}-J{r}"
        csh[f"L{r}"] = f'=IF({on},0,SUMIFS({tl("M")},{tl("B")},A{r},{tl("O")},"Unbilled"))'
        csh[f"M{r}"] = f'=IF({on},0,SUMIFS({tl("M")},{tl("B")},A{r},{tl("A")},">="&{YS},{tl("A")},"<="&{AS}))'
        csh[f"N{r}"] = f"=IF(M{r}<=0,0,M{r}+ROW()/100000)"
    csh.conditional_formatting.add(f"A{C0}:A{C1}", FormulaRule(formula=[f'AND(A{C0}<>"",COUNTIF($A${C0}:$A${C1},A{C0})>1)'], fill=fill("FCA5A5")))
    csh.freeze_panes = "B6"

    # ---------- Projects ----------
    K.title(jsh, "Projects", "Optional. A project rate overrides the client rate. Budget hours shows how much of the estimate is used.")
    K.hdr(jsh, 5, ["Project", "Client", "Hourly rate (optional)", "Budget hours (optional)", "Status", "Notes",
                   "Hours logged", "Billable amount", "Budget used", "Hours left"], [28, 26, 10, 10, 10, 20, 9, 11, 9, 9])
    style_range(jsh, J0, J1, 1, 10, inputs=range(1, 7), fmt={3: CUR, 4: HRS, 7: HRS, 8: CUR, 9: PCT, 10: HRS})
    if sample:
        for i, p in enumerate(PROJECTS):
            for j, v in enumerate((p[0], CLIENTS[p[1]][0], p[2], p[3], p[4]), 1):
                if v is not None: jsh.cell(row=J0 + i, column=j, value=v)
    for r in range(J0, J1 + 1):
        on = f'A{r}=""'
        jsh[f"G{r}"] = f'=IF({on},0,SUMIFS({tl("K")},{tl("C")},A{r},{tl("B")},B{r}))'
        jsh[f"H{r}"] = f'=IF({on},0,SUMIFS({tl("M")},{tl("C")},A{r},{tl("B")},B{r}))'
        jsh[f"I{r}"] = f'=IF(OR({on},N(D{r})=0),"",G{r}/D{r})'
        jsh[f"J{r}"] = f'=IF(OR({on},N(D{r})=0),"",D{r}-G{r})'
    dv = DataValidation(type="list", formula1=f"={cl('A')}", allow_blank=True); jsh.add_data_validation(dv); dv.add(f"B{J0}:B{J1}")
    dv = DataValidation(type="list", formula1='"Active,On hold,Done"', allow_blank=True); jsh.add_data_validation(dv); dv.add(f"E{J0}:E{J1}")
    jsh.conditional_formatting.add(f"I{J0}:I{J1}", CellIsRule(operator="greaterThan", formula=["1"], fill=fill("FEE2E2"), font=Font(name=F, bold=True, color="B91C1C")))
    jsh.conditional_formatting.add(f"I{J0}:I{J1}", CellIsRule(operator="between", formula=["0.8", "1"], fill=fill("FEF3C7"), font=Font(name=F, bold=True, color="92400E")))
    jsh.freeze_panes = "B6"

    # ---------- Time Log ----------
    K.title(tsh, "Time Log", "One row per work session. Type start and end time (e.g. 9:30 and 11:15) OR just the hours.")
    tsh["A3"] = "Invoice #: leave empty until you bill the work – then type the invoice number (it must match a row on the Invoices tab)."
    tsh["A3"].font = font(bold=True, size=9)
    heads = ["Date", "Client", "Project", "Task / description", "Start", "End", "Break (min)", "Hours (instead of start/end)", "Billable?",
             "Invoice #", "Hours", "Rate", "Amount", "Month", "Status", "Invoice line key"]
    K.hdr(tsh, 5, heads, [11, 24, 24, 30, 7, 7, 7, 10, 8, 13, 8, 8, 10, 9, 11, 7])
    style_range(tsh, T0, T1, 1, 16, inputs=range(1, 11), fmt={1: DATE, 5: TIME, 6: TIME, 7: "0", 8: HRS, 11: HRS, 12: CUR, 13: CUR, 14: "mmm yy;;", 16: "0"})
    for i, row in enumerate(logs):
        for j, v in enumerate(row, 1):
            if v not in (None, ""): tsh.cell(row=T0 + i, column=j, value=v)
    for r in range(T0, T1 + 1):
        raw = f"IF(N(H{r})>0,H{r},IF(AND(E{r}<>\"\",F{r}<>\"\"),MAX(0,MOD(F{r}-E{r},1)*24-N(G{r})/60),0))"
        tsh[f"K{r}"] = f'=IF(N({RND})>0,ROUNDUP({raw}*60/{RND},0)*{RND}/60,{raw})'
        prate = f"IFERROR(INDEX({pj('C')},MATCH(C{r},{pj('A')},0)),0)"
        crate = f"IFERROR(INDEX({cl('E')},MATCH(B{r},{cl('A')},0)),0)"
        tsh[f"L{r}"] = f'=IF(K{r}=0,"",IF(N({prate})>0,{prate},IF(N({crate})>0,{crate},{DRATE})))'
        tsh[f"M{r}"] = f'=IF(OR(K{r}=0,I{r}="No"),0,K{r}*L{r})'
        tsh[f"N{r}"] = f'=IF(N(A{r})=0,0,DATE(YEAR(A{r}),MONTH(A{r}),1))'
        tsh[f"O{r}"] = f'=IF(K{r}=0,"",IF(I{r}="No","Non-billable",IF(J{r}="","Unbilled","Invoiced")))'
        tsh[f"P{r}"] = f'=IF(AND(J{r}<>"",J{r}=Invoice!$C$4,K{r}>0),ROW(),0)'
    dv = DataValidation(type="list", formula1=f"={cl('A')}", allow_blank=True); tsh.add_data_validation(dv); dv.add(f"B{T0}:B{T1}")
    dv = DataValidation(type="list", formula1=f"={pj('A')}", allow_blank=True); tsh.add_data_validation(dv); dv.add(f"C{T0}:C{T1}")
    dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True); tsh.add_data_validation(dv); dv.add(f"I{T0}:I{T1}")
    dv = DataValidation(type="list", formula1=f"={iv('A')}", allow_blank=True); tsh.add_data_validation(dv); dv.add(f"J{T0}:J{T1}")
    tsh.conditional_formatting.add(f"O{T0}:O{T1}", CellIsRule(operator="equal", formula=['"Unbilled"'], fill=fill("FEF3C7"), font=Font(name=F, bold=True, color="92400E")))
    tsh.conditional_formatting.add(f"O{T0}:O{T1}", CellIsRule(operator="equal", formula=['"Invoiced"'], fill=fill("D1FAE5"), font=Font(name=F, color="166534")))
    tsh.freeze_panes = "C6"

    # ---------- Invoices ----------
    K.title(ish, "Invoices", "One row per invoice. The amount comes from the Time Log rows with the same invoice number.")
    ish["A3"] = "Type the paid date when the money arrives – the status changes to Paid. Due date is optional (default: issue date + client's payment terms)."
    ish["A3"].font = font(bold=True, size=9)
    heads = ["Invoice #", "Client", "Issue date", "Due date (optional)", "Extra items / expenses", "Discount", "Paid date", "Notes",
             "Hours", "Time amount", "Subtotal", "Tax", "Total", "Due date", "Status", "Days overdue", "Month issued", "Open key"]
    K.hdr(ish, 5, heads, [13, 24, 11, 11, 10, 9, 11, 18, 8, 11, 11, 9, 11, 11, 10, 8, 9, 7])
    style_range(ish, I0, I1, 1, 18, inputs=range(1, 9), fmt={3: DATE, 4: DATE, 5: CUR, 6: CUR, 7: DATE, 9: HRS, 10: CUR, 11: CUR, 12: CUR,
                                                            13: CUR, 14: DATE, 16: '0;-0;""', 17: "mmm yy;;", 18: "0.000"})
    for i, row in enumerate(invs):
        for j, v in enumerate(row, 1):
            if v is not None: ish.cell(row=I0 + i, column=j, value=v)
    for r in range(I0, I1 + 1):
        on = f'A{r}=""'
        ish[f"I{r}"] = f'=IF({on},0,SUMIFS({tl("K")},{tl("J")},A{r},{tl("I")},"<>No"))'
        ish[f"J{r}"] = f'=IF({on},0,SUMIFS({tl("M")},{tl("J")},A{r}))'
        ish[f"K{r}"] = f"=J{r}+N(E{r})-N(F{r})"
        ish[f"L{r}"] = f"=K{r}*{TAX}"
        ish[f"M{r}"] = f"=K{r}+L{r}"
        terms = f"IFERROR(INDEX({cl('F')},MATCH(B{r},{cl('A')},0)),0)"
        ish[f"N{r}"] = f'=IF(OR({on},N(C{r})=0),"",IF(N(D{r})>0,D{r},C{r}+IF(N({terms})>0,{terms},{TERMS})))'
        ish[f"O{r}"] = f'=IF({on},"",IF(N(G{r})>0,"Paid",IF(N(C{r})=0,"Draft",IF({AS}>N{r},"Overdue","Sent"))))'
        ish[f"P{r}"] = f'=IF(O{r}="Overdue",{AS}-N{r},"")'
        ish[f"Q{r}"] = f'=IF(N(C{r})=0,0,DATE(YEAR(C{r}),MONTH(C{r}),1))'
        ish[f"R{r}"] = f'=IF(OR(O{r}="Sent",O{r}="Overdue"),N{r}+ROW()/100000,0)'
    dv = DataValidation(type="list", formula1=f"={cl('A')}", allow_blank=True); ish.add_data_validation(dv); dv.add(f"B{I0}:B{I1}")
    ish.conditional_formatting.add(f"O{I0}:O{I1}", CellIsRule(operator="equal", formula=['"Overdue"'], fill=fill("FEE2E2"), font=Font(name=F, bold=True, color="B91C1C")))
    ish.conditional_formatting.add(f"O{I0}:O{I1}", CellIsRule(operator="equal", formula=['"Sent"'], fill=fill("FEF3C7"), font=Font(name=F, color="92400E")))
    ish.conditional_formatting.add(f"O{I0}:O{I1}", CellIsRule(operator="equal", formula=['"Paid"'], fill=fill("D1FAE5"), font=Font(name=F, color="166534")))
    ish.conditional_formatting.add(f"A{I0}:A{I1}", FormulaRule(formula=[f'AND(A{I0}<>"",COUNTIF($A${I0}:$A${I1},A{I0})>1)'], fill=fill("FCA5A5")))
    ish.freeze_panes = "B6"

    # ---------- Invoice (printable) ----------
    inv.sheet_view.showGridLines = False
    for c, w in zip("ABCDEFG", [3, 12, 44, 10, 11, 14, 3]): inv.column_dimensions[c].width = w
    inv["B2"] = "Choose invoice #  ▸"; inv["B2"].font = font(bold=True, size=9)
    inv["B4"] = "Invoice #"; inv["B4"].font = font(bold=True)
    input_cell(inv["C4"], invs[-1][0] if invs else None); inv["C4"].font = Font(name=F, bold=True, size=12, color="0000FF")
    inv["C2"] = "Pick an invoice number in the yellow cell – everything below fills in. Print or File > Save as PDF (only this tab)."
    inv["C2"].font = font(italic=True, size=9)
    dv = DataValidation(type="list", formula1=f"={iv('A')}", allow_blank=True); inv.add_data_validation(dv); dv.add("C4")
    ir = f"MATCH($C$4,{iv('A')},0)"
    def ivx(col): return f"IFERROR(INDEX({iv(col)},{ir}),\"\")"
    inv["E4"] = "INVOICE"; inv["E4"].font = Font(name=F, size=24, bold=True, color=MAIN)
    inv["B6"] = "=IF(Settings!B6=\"\",\"Your name / business\",Settings!B6)"; inv["B6"].font = Font(name=F, size=14, bold=True, color=MAIN)
    for i, ref in enumerate(["B7", "B8", "B9"], 7):
        inv[f"B{i}"] = f'=IF(Settings!{ref}="","",Settings!{ref})'.replace(ref, f"B{i}") if False else f'=IF(Settings!B{i}="","",Settings!B{i})'
        inv[f"B{i}"].font = font(size=9)
    inv["B10"] = '=IF(Settings!B16="","",Settings!B16)'; inv["B10"].font = font(size=9)
    for i, (lab, f, fmt) in enumerate([("Invoice date", f"={ivx('C')}", DATE), ("Due date", f"={ivx('N')}", DATE), ("Status", f"={ivx('O')}", None)], 6):
        inv[f"E{i}"] = lab; inv[f"E{i}"].font = font(bold=True, size=9)
        inv[f"F{i}"] = f; inv[f"F{i}"].font = font(); inv[f"F{i}"].alignment = Alignment(horizontal="right")
        if fmt: inv[f"F{i}"].number_format = fmt
    inv["B12"] = "BILL TO"; inv["B12"].font = Font(name=F, bold=True, size=9, color=MAIN)
    cr = f"MATCH({ivx('B')},{cl('A')},0)"
    inv["B13"] = f"={ivx('B')}"; inv["B13"].font = font(bold=True)
    for i, col in enumerate("BCD", 14):
        inv[f"B{i}"] = f'=IFERROR(INDEX({cl(col)},{cr})&"","")'; inv[f"B{i}"].font = font(size=9)
    K.hdr(inv, 18, ["Date", "Description", "Hours", "Rate", "Amount"], col0=2)
    L0 = 19
    for n in range(NL):
        r = L0 + n
        rw = f"SMALL({tl('P')},COUNTIF({tl('P')},0)+{n + 1})"
        inv[f"B{r}"] = f'=IFERROR(INDEX({TL}!$A:$A,{rw}),"")'
        inv[f"C{r}"] = (f'=IFERROR(IF(INDEX({TL}!$C:$C,{rw})="","",INDEX({TL}!$C:$C,{rw})&" – ")&INDEX({TL}!$D:$D,{rw})'
                        f'&IF(INDEX({TL}!$I:$I,{rw})="No"," (no charge)",""),"")')
        inv[f"D{r}"] = f'=IFERROR(INDEX({TL}!$K:$K,{rw}),"")'
        inv[f"E{r}"] = f'=IFERROR(INDEX({TL}!$L:$L,{rw}),"")'
        inv[f"F{r}"] = f'=IFERROR(INDEX({TL}!$M:$M,{rw}),"")'
        for c in "BCDEF":
            cell = inv[f"{c}{r}"]; cell.font = font(size=9); cell.border = Border(bottom=Side(style="thin", color="E5E7EB"))
        inv[f"B{r}"].number_format = DATE; inv[f"D{r}"].number_format = HRS; inv[f"E{r}"].number_format = CUR; inv[f"F{r}"].number_format = CUR
    e = L0 + NL
    inv[f"B{e}"] = f'=IF(COUNTIF({tl("P")},">0")>{NL},"Note: "&COUNTIF({tl("P")},">0")&" time entries – only the first {NL} are listed; the totals include all of them.","")'
    inv[f"B{e}"].font = Font(name=F, italic=True, size=8, color="B91C1C")
    tots = [("Total hours", f"={ivx('I')}", HRS), ("Time", f"={ivx('J')}", CUR), ("Extra items / expenses", f"=N({ivx('E')})", CUR),
            ("Discount", f"=-N({ivx('F')})", CUR), ("Subtotal", f"={ivx('K')}", CUR),
            (f'=Settings!B11&" ("&TEXT(Settings!B12,"0.0%")&")"', f"={ivx('L')}", CUR), ("TOTAL DUE", f"={ivx('M')}", CUR)]
    for i, (lab, f, fmt) in enumerate(tots, e + 1):
        inv[f"D{i}"] = lab; inv[f"D{i}"].font = font(bold=True, size=9)
        inv[f"F{i}"] = f; inv[f"F{i}"].number_format = fmt; inv[f"F{i}"].font = font()
    last = e + len(tots)
    for c in "DEF": inv[f"{c}{last}"].fill = fill(SOFT)
    inv[f"D{last}"].font = Font(name=F, bold=True, size=11, color=MAIN)
    inv[f"F{last}"] = f'=IF($C$4="","",Settings!$B$4&" "&TEXT({ivx("M")},"#,##0.00"))'
    inv[f"F{last}"].font = Font(name=F, bold=True, size=11, color=MAIN); inv[f"F{last}"].alignment = Alignment(horizontal="right")
    inv[f"B{last + 2}"] = "PAYMENT DETAILS"; inv[f"B{last + 2}"].font = Font(name=F, bold=True, size=9, color=MAIN)
    inv[f"B{last + 3}"] = '=IF(Settings!B10="","",Settings!B10)'; inv[f"B{last + 3}"].font = font(size=9)
    inv[f"B{last + 4}"] = f'=IF($C$4="","","Please pay by "&TEXT({ivx("N")},"dd mmm yyyy")&" and mention invoice "&$C$4&". Thank you!")'
    inv[f"B{last + 4}"].font = font(size=9, italic=True)
    inv[f"B{last + 5}"] = f'=IF({ivx("H")}="","","Note: "&{ivx("H")})'; inv[f"B{last + 5}"].font = font(size=9)
    inv.print_area = f"A1:G{last + 6}"

    # ---------- Calc ----------
    K.title(cc, "Calc", "Numbers for the dashboard charts. Calculated – nothing to type.")
    K.hdr(cc, 4, ["Month", "Label", "Hours", "Billable hours", "Earned (work done)", "Invoiced", "Paid"], [11, 9, 9, 9, 12, 12, 12])
    for i in range(12):
        r = 5 + i
        cc[f"A{r}"] = f"=DATE(YEAR({AS}),MONTH({AS})-{11 - i},1)"
        cc[f"B{r}"] = f'=TEXT(A{r},"mmm yy")'
        cc[f"C{r}"] = f'=SUMIFS({tl("K")},{tl("N")},A{r})'
        cc[f"D{r}"] = f'=SUMIFS({tl("K")},{tl("N")},A{r},{tl("I")},"<>No")'
        cc[f"E{r}"] = f'=SUMIFS({tl("M")},{tl("N")},A{r})'
        cc[f"F{r}"] = f'=SUMIFS({iv("M")},{iv("Q")},A{r})'
        cc[f"G{r}"] = f'=SUMIFS({iv("M")},{iv("G")},">="&A{r},{iv("G")},"<"&DATE(YEAR(A{r}),MONTH(A{r})+1,1))'
    style_range(cc, 5, 16, 1, 7, fmt={1: "mmm yyyy", 3: HRS, 4: HRS, 5: CUR, 6: CUR, 7: CUR})

    # ---------- Dashboard ----------
    K.title(dash, "Freelance Dashboard", "Updates by itself from the Time Log and Invoices tabs.")
    dash["A3"] = f'=IF(Settings!B6="","",Settings!B6&"  ·  ")&"As of "&TEXT({AS},"dd mmm yyyy")&"  ·  amounts in "&Settings!B4'
    dash["A3"].font = font(bold=True)
    inm = f'{tl("A")},">="&{MS},{tl("A")},"<="&{AS}'; iny = f'{tl("A")},">="&{YS},{tl("A")},"<="&{AS}'
    kp = [("Hours this week", f'=SUMIFS({tl("K")},{tl("A")},">="&{WS},{tl("A")},"<="&{AS})', "#,##0.0"),
          ("Hours this month", f'=SUMIFS({tl("K")},{inm})', "#,##0.0"),
          ("Billable % this month", f'=IFERROR(SUMIFS({tl("K")},{inm},{tl("I")},"<>No")/C6,0)', "0%"),
          ("Earned this month (work done)", f'=SUMIFS({tl("M")},{inm})', "#,##0.00"),
          ("Unbilled work (not invoiced yet)", f'=SUMIFS({tl("M")},{tl("O")},"Unbilled")', "#,##0.00"),
          ("Outstanding invoices", f'=SUMIFS({iv("M")},{iv("O")},"Sent")+SUMIFS({iv("M")},{iv("O")},"Overdue")', "#,##0.00"),
          ("Overdue", f'=SUMIFS({iv("M")},{iv("O")},"Overdue")', "#,##0.00"),
          ("Paid this year", f'=SUMIFS({iv("M")},{iv("G")},">="&{YS},{iv("G")},"<="&{AS})', "#,##0.00"),
          ("Effective hourly rate (year)", f'=IFERROR(SUMIFS({tl("M")},{iny})/SUMIFS({tl("K")},{iny}),0)', "#,##0.00")]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 3) * 3, 1 + (i % 3) * 2, lab, f, fmt)
        dash.cell(row=6 + (i // 3) * 3, column=1 + (i % 3) * 2).alignment = Alignment(horizontal="left")
    dash["A14"] = "Effective hourly rate = billable amount ÷ ALL hours you logged (including non-billable admin time)."
    dash["A14"].font = font(italic=True, size=8)
    for c, w in zip("ABCDEFGHIJKL", [28, 14, 22, 14, 26, 12, 3, 14, 24, 11, 11, 11]): dash.column_dimensions[c].width = w
    K.section(dash, "A16", "Clients (billable this year)")
    K.hdr(dash, 17, ["Client", "Billable", "Unbilled", "Outstanding"])
    for n in range(10):
        r = 18 + n
        m = f"MATCH(LARGE({cl('N')},{n + 1}),{cl('N')},0)"
        dash[f"A{r}"] = f'=IF(LARGE({cl("N")},{n + 1})=0,"",INDEX({cl("A")},{m}))'
        for c, src in zip("BCD", "MLK"):
            dash[f"{c}{r}"] = f'=IF($A{r}="","",INDEX({cl(src)},{m}))'
    style_range(dash, 18, 27, 1, 4, fmt={2: CUR, 3: CUR, 4: CUR})
    K.section(dash, "H16", "Unpaid invoices (oldest due first)")
    K.hdr(dash, 17, ["Invoice #", "Client", "Due date", "Total", "Status"], col0=8)
    for n in range(10):
        r = 18 + n
        m = f"MATCH(SMALL({iv('R')},COUNTIF({iv('R')},0)+{n + 1}),{iv('R')},0)"
        dash[f"H{r}"] = f'=IFERROR(INDEX({iv("A")},{m}),"")'
        for c, src in zip("IJKL", "BNMO"):
            dash[f"{c}{r}"] = f'=IF($H{r}="","",INDEX({iv(src)},{m}))'
    style_range(dash, 18, 27, 8, 12, fmt={10: DATE, 11: CUR})
    dash.conditional_formatting.add("L18:L27", CellIsRule(operator="equal", formula=['"Overdue"'], fill=fill("FEE2E2"), font=Font(name=F, bold=True, color="B91C1C")))
    c1 = BarChart(); c1.type = "col"; c1.title = "Earned vs paid by month"; c1.height = 7.5; c1.width = 17
    c1.add_data(Reference(cc, min_col=5, min_row=4, max_row=16), titles_from_data=True)
    c1.add_data(Reference(cc, min_col=7, min_row=4, max_row=16), titles_from_data=True)
    c1.set_categories(Reference(cc, min_col=2, min_row=5, max_row=16))
    for ser, col in zip(c1.series, (MAIN, ACC)): ser.graphicalProperties.solidFill = col
    c1.y_axis.majorGridlines = None; c1.y_axis.number_format = "#,##0"; c1.y_axis.scaling.min = 0
    dash.add_chart(c1, "A30")
    c2 = BarChart(); c2.type = "col"; c2.grouping = "clustered"; c2.title = "Hours per month"; c2.height = 7.5; c2.width = 14
    c2.add_data(Reference(cc, min_col=3, max_col=4, min_row=4, max_row=16), titles_from_data=True)
    c2.set_categories(Reference(cc, min_col=2, min_row=5, max_row=16))
    for ser, col in zip(c2.series, ("94A3B8", MAIN)): ser.graphicalProperties.solidFill = col
    c2.y_axis.majorGridlines = None; c2.y_axis.scaling.min = 0
    dash.add_chart(c2, "H30")

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Settings tab: your name, address, email, payment details, currency, tax and default hourly rate (they appear on your invoices).", False),
             ("2.  Clients tab: one row per client with their hourly rate and payment terms. Projects tab (optional): a different rate or a budget in hours.", False),
             ("3.  Time Log tab: one row per work session – start and end time, or just the number of hours. Mark non-billable time 'No'.", False),
             ("4.  Ready to bill? Add a row on the Invoices tab (invoice #, client, issue date) and type that invoice # on the Time Log rows it covers.", False),
             ("5.  Invoice tab: choose the invoice number – the invoice fills itself in. Print it or save it as PDF and send it.", False),
             ("6.  When the client pays, type the paid date on the Invoices tab. The Dashboard shows unbilled work, unpaid and overdue invoices.", False),
             ("", False), ("How it works", True),
             ("Hours = end − start − break (works past midnight), or the hours you type. Optional rounding up to 6, 15 or 30 minutes (Settings).", False),
             ("Rate = project rate if set, otherwise the client's rate, otherwise your default rate.", False),
             ("Invoice total = time on the invoice + extra items − discount + tax. Due date = issue date + the client's payment terms.", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  Room for 100 clients, 150 projects, 3,000 time entries and 300 invoices. Each invoice lists up to 30 time entries (the total always includes all).", False),
             ("•  Save as PDF: select the Invoice tab, then File > Print (or Download > PDF in Google Sheets) and print only the current sheet.", False),
             ("•  Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("A made-up freelance designer with 5 clients, 8 projects and 8 months of time entries and invoices, shown as of 30 Oct 2026. Use the blank copy for your own work.", False)]
    K.start(st, "Freelancer Time Tracker & Invoice", "Track your hours, send invoices in one click and see who still owes you. Works in Excel and Google Sheets.", lines)
    finish(wb, sample, path)


META = {
    "name": "Freelancer Time Tracker & Invoice Generator – Excel & Google Sheets",
    "file_base": "Freelancer-Time-Tracker-Invoice",
    "price": 19,
    "color": "#0F766E",
    "title": "Freelancer Time Tracker & Invoice",
    "tagline": "Hours · invoices · unpaid & overdue — Excel & Google Sheets",
    "intro": "Log your hours, turn them into a professional invoice in one click and always know who still owes you. Built for freelancers, consultants, designers, writers, developers and virtual assistants who bill by the hour.",
    "features": [
        "Time log with start/end times or plain hours, breaks, optional rounding (6/15/30 min) and billable / non-billable time",
        "Automatic rates: per client, per project or your default rate",
        "Printable invoice that fills itself in when you pick an invoice number – save as PDF and send",
        "Invoices list with tax, extra items, discounts, due dates and status: Sent, Overdue or Paid",
        "Project budgets: hours logged vs. estimate, with warnings at 80% and 100%",
        "Dashboard: hours this week and month, billable %, unbilled work, outstanding and overdue invoices, paid this year and your effective hourly rate",
        "Charts of earned vs. paid and hours per month",
    ],
    "tags": ["freelancer", "time tracker", "invoice template", "timesheet", "hourly billing", "google sheets", "excel template",
             "freelance business", "invoice generator", "client tracker", "consultant", "virtual assistant", "small business"],
}
