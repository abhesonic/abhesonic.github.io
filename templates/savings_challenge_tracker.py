"""Savings Challenge Tracker – 52-week challenge, 100-envelope challenge and savings goals (sample + blank versions).

Settings (start date, 52-week style + amount, envelope multiplier) -> 52-Week Challenge (planned deposit per week, tick when
saved, running total, on-track status) -> 100 Envelopes (list + 10x10 grid that turns green) -> Savings Goals (targets,
deadlines, needed per month/week, fed by the Deposits log) -> Dashboard (total saved, progress bars, ahead/behind, chart).
"""
import random, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.chart import LineChart, Reference
from openpyxl.utils import get_column_letter as L
from _common import Kit, F, INK, B, MONEY, DATE, font, fill, style_range, input_cell, finish

MAIN = "B45309"; SOFT = "FEF6E7"; GREEN = "BBF7D0"
K = Kit(MAIN, SOFT)
W0 = 5; W1 = W0 + 51          # 52 weeks: rows 5..56
E0 = 5; E1 = E0 + 99          # 100 envelopes: rows 5..104
NG = 10; G0 = 5; G1 = G0 + NG - 1
ND = 400; D0 = 5; D1 = D0 + ND - 1
STYLES = ["Increasing – week 1 is the smallest", "Decreasing – start with the biggest week", "Same amount every week"]
WK = "'52-Week Challenge'"; ENV = "'100 Envelopes'"; GL = "'Savings Goals'"
START, STEP, MULT, ASOF, STY = "Settings!$B$5", "Settings!$B$7", "Settings!$B$8", "Settings!$B$9", "Settings!$B$11"
PCT = '0%'
BAR = 20


def bar(ref):
    return f'REPT("█",ROUND(MIN(1,MAX(0,N({ref})))*{BAR},0))&REPT("░",{BAR}-ROUND(MIN(1,MAX(0,N({ref})))*{BAR},0))'


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); wk = wb.create_sheet("52-Week Challenge"); env = wb.create_sheet("100 Envelopes")
    gl = wb.create_sheet("Savings Goals"); dp = wb.create_sheet("Deposits"); s = wb.create_sheet("Settings")
    random.seed(11)

    # ---------- Settings ----------
    K.title(s, "Settings", "Yellow cells are yours to edit.")
    rows = [("Currency label", "USD", None),
            ("Challenge start date", dt.date(2026, 1, 5) if sample else "=TODAY()-WEEKDAY(TODAY(),3)", DATE),
            ("52-week style", STYLES[0], None),
            ("52-week amount", 5 if sample else 1, MONEY),
            ("Envelope multiplier", 1, "0.##"),
            ("Today / as-of date", dt.date(2026, 10, 30) if sample else "=TODAY()", DATE)]
    for i, (k_, v, fmt) in enumerate(rows, 4):
        s.cell(row=i, column=1, value=k_).font = font(bold=True)
        c = s.cell(row=i, column=2); input_cell(c, v)
        if fmt: c.number_format = fmt
    notes = {5: "Week 1 starts on this date. It starts as this week's Monday.",
             6: "Choose from the list.",
             7: "Increasing/decreasing: week n saves n × this amount (1 → 1,378 a year · 5 → 6,890 · 10 → 13,780). Same amount: saved every week.",
             8: "Envelope n holds n × this amount (1 → 5,050 in total · 2 → 10,100 · 0.5 → 2,525).",
             9: "Leave as =TODAY() – used for 'this week' and 'on track' checks."}
    for r, t in notes.items(): s[f"C{r}"] = t; s[f"C{r}"].font = font(italic=True, size=9)
    calc = [("Style number (1-3)", "=IFERROR(MATCH(B6,$H$5:$H$7,0),1)", "0"),
            ("Current challenge week", f"=MIN(53,MAX(0,INT((B9-B5)/7)+1))", "0")]
    for i, (k_, f, fmt) in enumerate(calc, 11):
        s.cell(row=i, column=1, value=k_).font = font(italic=True)
        c = s.cell(row=i, column=2, value=f); c.number_format = fmt; c.border = B; c.fill = fill("F3F4F6"); c.font = font()
    s["H4"] = "52-week styles"; s["H4"].font = font(bold=True, size=9)
    for i, t in enumerate(STYLES): s.cell(row=5 + i, column=8, value=t).font = font(size=9)
    dv = DataValidation(type="list", formula1="=$H$5:$H$7", allow_blank=False); s.add_data_validation(dv); dv.add("B6")
    for c, w in zip("ABCDEFGH", [28, 16, 92, 3, 3, 3, 3, 38]): s.column_dimensions[c].width = w

    # ---------- 52-Week Challenge ----------
    K.title(wk, "52-Week Savings Challenge", "Type x in the 'Saved?' column when you put the money away. Saved a different amount? Type it in 'Actual amount'.")
    K.hdr(wk, 4, ["Week", "Week starting", "Planned deposit", "Saved? (x)", "Actual amount (optional)", "Saved", "Running total",
                  "Planned running total", "Status"], [8, 14, 14, 10, 15, 13, 15, 15, 16])
    style_range(wk, W0, W1, 1, 9, inputs=(4, 5), fmt={2: DATE, 3: MONEY, 5: MONEY, 6: MONEY, 7: MONEY, 8: MONEY})
    missed = {17, 30}
    for i in range(52):
        r = W0 + i; n = i + 1
        wk[f"A{r}"] = n
        wk[f"B{r}"] = f"={START}+7*(A{r}-1)"
        wk[f"C{r}"] = f"=CHOOSE({STY},{STEP}*A{r},{STEP}*(53-A{r}),{STEP})"
        wk[f"F{r}"] = f'=IF(E{r}<>"",N(E{r}),IF(D{r}<>"",C{r},0))'
        wk[f"G{r}"] = f"=SUM($F${W0}:F{r})"
        wk[f"H{r}"] = f"=SUM($C${W0}:C{r})"
        wk[f"I{r}"] = (f'=IF(F{r}>0,IF(F{r}>=C{r},"Done","Part saved"),IF(A{r}=Settings!$B$12,"This week",'
                       f'IF(A{r}<Settings!$B$12,"Missed – catch up","Coming up")))')
        wk[f"A{r}"].alignment = Alignment(horizontal="center"); wk[f"D{r}"].alignment = Alignment(horizontal="center")
        if sample and n <= 42 and n not in missed:
            wk[f"D{r}"] = "x"
            if n == 25: wk[f"E{r}"] = 100
    wk[f"A{W1 + 1}"] = "Total"; wk[f"C{W1 + 1}"] = f"=SUM(C{W0}:C{W1})"; wk[f"F{W1 + 1}"] = f"=SUM(F{W0}:F{W1})"
    for c in "ACF":
        x = wk[f"{c}{W1 + 1}"]; x.font = font(bold=True); x.border = B; x.fill = fill(SOFT)
        if c != "A": x.number_format = MONEY
    wk.conditional_formatting.add(f"A{W0}:I{W1}", FormulaRule(formula=[f'$I{W0}="Done"'], fill=fill(GREEN)))
    wk.conditional_formatting.add(f"A{W0}:I{W1}", FormulaRule(formula=[f'$I{W0}="This week"'], fill=fill("FDE68A"), font=Font(name=F, bold=True, color="92400E")))
    wk.conditional_formatting.add(f"I{W0}:I{W1}", FormulaRule(formula=[f'LEFT($I{W0},6)="Missed"'], font=Font(name=F, bold=True, color="B91C1C")))
    dv = DataValidation(type="list", formula1='"x"', allow_blank=True); wk.add_data_validation(dv); dv.add(f"D{W0}:D{W1}")
    wk.freeze_panes = "A5"

    # ---------- 100 Envelopes ----------
    K.title(env, "100 Envelope Challenge", "Pick any envelope, save its amount, then type x next to it. The grid turns green as you go.")
    K.hdr(env, 4, ["Envelope", "Amount", "Saved? (x)", "Date (optional)", "Saved"], [10, 12, 10, 14, 12])
    style_range(env, E0, E1, 1, 5, inputs=(3, 4), fmt={2: MONEY, 4: DATE, 5: MONEY})
    done_env = set(random.sample(range(1, 101), 58)) if sample else set()
    for i in range(100):
        r = E0 + i; n = i + 1
        env[f"A{r}"] = n; env[f"A{r}"].alignment = Alignment(horizontal="center")
        env[f"B{r}"] = f"=A{r}*{MULT}"
        env[f"E{r}"] = f'=IF(C{r}<>"",B{r},0)'
        env[f"C{r}"].alignment = Alignment(horizontal="center")
        if n in done_env:
            env[f"C{r}"] = "x"; env[f"D{r}"] = dt.date(2026, 1, 5) + dt.timedelta(days=int(random.random() * 295))
    env[f"A{E1 + 1}"] = "Total"; env[f"B{E1 + 1}"] = f"=SUM(B{E0}:B{E1})"; env[f"E{E1 + 1}"] = f"=SUM(E{E0}:E{E1})"
    for c in "ABE":
        x = env[f"{c}{E1 + 1}"]; x.font = font(bold=True); x.border = B; x.fill = fill(SOFT)
        if c != "A": x.number_format = MONEY
    env.conditional_formatting.add(f"A{E0}:E{E1}", FormulaRule(formula=[f'$C{E0}<>""'], fill=fill(GREEN)))
    dv = DataValidation(type="list", formula1='"x"', allow_blank=True); env.add_data_validation(dv); dv.add(f"C{E0}:C{E1}")
    # grid
    env["G3"] = "Envelope grid (number = envelope, green = saved)"; env["G3"].font = Font(name=F, bold=True, size=12, color=MAIN)
    for c in range(10):
        env.column_dimensions[L(7 + c)].width = 8
        for rr in range(10):
            n = rr * 10 + c + 1; cell = env.cell(row=5 + rr, column=7 + c, value=n)
            cell.alignment = Alignment(horizontal="center", vertical="center"); cell.border = B; cell.font = font(bold=True)
            cell.fill = fill("F3F4F6")
        for rr in range(10): env.row_dimensions[5 + rr].height = 24
    env.column_dimensions["F"].width = 3
    env.conditional_formatting.add("G5:P14", FormulaRule(formula=[f'INDEX($C${E0}:$C${E1},G5)<>""'], fill=fill("22C55E"), font=Font(name=F, bold=True, color="FFFFFF")))
    env["G16"] = "Envelopes saved"; env["J16"] = f'=COUNTIF(C{E0}:C{E1},"<>")'
    env["G17"] = "Saved so far"; env["J17"] = f"=E{E1 + 1}"; env["J17"].number_format = MONEY
    env["G18"] = "Still to save"; env["J18"] = f"=B{E1 + 1}-E{E1 + 1}"; env["J18"].number_format = MONEY
    env["G19"] = "Smallest open envelope"; env["J19"] = f'=IFERROR(INDEX(A{E0}:A{E1},MATCH(0,E{E0}:E{E1},0)),"All done!")'
    for rr in range(16, 20):
        env[f"G{rr}"].font = font(bold=True); env[f"J{rr}"].font = Font(name=F, bold=True, color=MAIN)
    env.freeze_panes = "A5"

    # ---------- Deposits ----------
    K.title(dp, "Deposits", "Every amount you put towards a savings goal. Choose the goal from the list.")
    K.hdr(dp, 4, ["Date", "Goal", "Amount", "Note"], [12, 26, 12, 30])
    style_range(dp, D0, D1, 1, 4, inputs=(1, 2, 3, 4), fmt={1: DATE, 3: MONEY})
    if sample:
        deps = []
        plan = [("Emergency fund", 250, 1), ("Vacation", 150, 2), ("New laptop", 120, 1), ("Holiday gifts", 60, 1)]
        for m in range(1, 11):
            for g, a, every in plan:
                if m % every == 0 and not (g == "Holiday gifts" and m < 6):
                    deps.append((dt.date(2026, m, 1 + (len(deps) * 3) % 26), g, a, "Monthly transfer"))
        deps.append((dt.date(2026, 4, 18), "Emergency fund", 600, "Tax refund"))
        deps.sort()
        for i, (d, g, a, note) in enumerate(deps):
            r = D0 + i; dp[f"A{r}"] = d; dp[f"B{r}"] = g; dp[f"C{r}"] = a; dp[f"D{r}"] = note
    dv = DataValidation(type="list", formula1=f"={GL}!$A${G0}:$A${G1}", allow_blank=True); dp.add_data_validation(dv); dv.add(f"B{D0}:B{D1}")
    dp.freeze_panes = "A5"

    # ---------- Savings Goals ----------
    K.title(gl, "Savings Goals", "Up to 10 goals. Add deposits on the Deposits tab – progress and 'needed per month' update by themselves.")
    K.hdr(gl, 4, ["Goal", "Target amount", "Target date", "Already saved before", "Saved", "Still needed", "Progress", "Months left",
                  "Needed per month", "Needed per week", "Progress bar"], [24, 13, 12, 14, 13, 13, 10, 9, 13, 13, 26])
    style_range(gl, G0, G1, 1, 11, inputs=(1, 2, 3, 4), fmt={2: MONEY, 3: DATE, 4: MONEY, 5: MONEY, 6: MONEY, 7: PCT, 9: MONEY, 10: MONEY})
    if sample:
        goals = [("Emergency fund", 6000, dt.date(2027, 6, 30), 1200), ("Vacation", 2500, dt.date(2027, 7, 31), 0),
                 ("New laptop", 1400, dt.date(2027, 2, 28), 0), ("Holiday gifts", 600, dt.date(2026, 12, 15), 0)]
        for i, row in enumerate(goals):
            for j, v in enumerate(row, 1): gl.cell(row=G0 + i, column=j, value=v)
    for r in range(G0, G1 + 1):
        gl[f"E{r}"] = f'=IF(A{r}="","",N(D{r})+SUMIFS(Deposits!$C${D0}:$C${D1},Deposits!$B${D0}:$B${D1},A{r}))'
        gl[f"F{r}"] = f'=IF(A{r}="","",MAX(0,N(B{r})-E{r}))'
        gl[f"G{r}"] = f'=IF(OR(A{r}="",N(B{r})=0),"",MIN(1,E{r}/B{r}))'
        gl[f"H{r}"] = f'=IF(OR(A{r}="",C{r}=""),"",MAX(0,(YEAR(C{r})-YEAR({ASOF}))*12+MONTH(C{r})-MONTH({ASOF})+IF(DAY(C{r})>=DAY({ASOF}),0,-1)))'
        gl[f"I{r}"] = f'=IF(OR(A{r}="",C{r}=""),"",IF(F{r}=0,0,IF(C{r}<={ASOF},F{r},F{r}/MAX(1,H{r}))))'
        gl[f"J{r}"] = f'=IF(OR(A{r}="",C{r}=""),"",IF(F{r}=0,0,F{r}/MAX(1,ROUNDUP((C{r}-{ASOF})/7,0))))'
        gl[f"K{r}"] = f'=IF(G{r}="","",{bar(f"G{r}")})'
        gl[f"K{r}"].font = Font(name=F, color=MAIN, size=9)
        gl[f"H{r}"].alignment = Alignment(horizontal="center")
    gl[f"A{G1 + 1}"] = "Total"
    for c in "BEF": gl[f"{c}{G1 + 1}"] = f"=SUM({c}{G0}:{c}{G1})"
    for c in "ABEF":
        x = gl[f"{c}{G1 + 1}"]; x.font = font(bold=True); x.border = B; x.fill = fill(SOFT)
        if c != "A": x.number_format = MONEY
    gl.conditional_formatting.add(f"A{G0}:K{G1}", FormulaRule(formula=[f'AND($A{G0}<>"",N($G{G0})>=1)'], fill=fill(GREEN)))
    gl[f"A{G1 + 3}"] = "'Already saved before' = money you had for this goal before you started tracking. Months left counts whole months until the target date."
    gl[f"A{G1 + 3}"].font = font(italic=True, size=9)
    gl.freeze_panes = "B5"

    # ---------- Dashboard ----------
    K.title(dash, "Savings Challenge Dashboard", "Updates by itself. Tick your weeks and envelopes and add deposits to your goals.")
    dash["A3"] = '="As of "&TEXT(Settings!B9,"mmm d, yyyy")&"  ·  Challenge week "&Settings!B12&" of 52  ·  Amounts in "&Settings!B4'
    dash["A3"].font = font(bold=True)
    kp = [("Total saved (all challenges)", f"=B15+B22+{GL}!E{G1 + 1}", MONEY),
          ("52-week challenge saved", f"={WK}!F{W1 + 1}", MONEY),
          ("100 envelopes saved", f"={ENV}!E{E1 + 1}", MONEY),
          ("Weeks completed", f'=COUNTIF({WK}!I{W0}:I{W1},"Done")&" / 52"', None),
          ("Envelopes completed", f'={ENV}!J16&" / 100"', None),
          ("Savings goals saved", f"={GL}!E{G1 + 1}", MONEY)]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 3) * 3, 1 + (i % 3) * 2, lab, f, fmt or "General")
        dash.cell(row=6 + (i // 3) * 3, column=1 + (i % 3) * 2).alignment = Alignment(horizontal="left")
    for c, w in zip("ABCDEFGH", [30, 16, 26, 16, 26, 16, 14, 26]): dash.column_dimensions[c].width = w

    def block(r0, title, rows_):
        K.section(dash, f"A{r0}", title)
        for i, (lab, f, fmt) in enumerate(rows_, r0 + 1):
            dash[f"A{i}"] = lab; dash[f"A{i}"].font = font(bold=True); dash[f"A{i}"].border = B; dash[f"A{i}"].fill = fill(SOFT)
            c = dash[f"B{i}"]; c.value = f; c.number_format = fmt; c.border = B; c.font = Font(name=F, bold=True, color=MAIN)
            c.alignment = Alignment(horizontal="right")
    ws_ = Settings_week = "Settings!$B$12"
    block(11, "52-Week Challenge", [
        ("Goal for the year", f"={WK}!C{W1 + 1}", MONEY),
        ("Should be saved by last week", f'=IF({ws_}<=1,0,INDEX({WK}!$H${W0}:$H${W1},MIN(52,{ws_}-1)))', MONEY),
        ("This week's deposit", f'=IF(OR({ws_}=0,{ws_}>52),0,INDEX({WK}!$C${W0}:$C${W1},{ws_}))', MONEY),
        ("Saved so far", f"={WK}!F{W1 + 1}", MONEY),
        ("Ahead (+) / behind (−)", "=B15-B13", '+#,##0.00;-#,##0.00;0.00'),
        ("Progress", "=IFERROR(B15/B12,0)", PCT)])
    dash["C17"] = f'={bar("B17")}'; dash["C17"].font = Font(name=F, color=MAIN)
    dash["C16"] = '=IF(B16>=0,"On track – nice work!","Missed weeks: catch up when you can")'; dash["C16"].font = font(italic=True, size=9)
    dash["C15"] = f'="Missed weeks: "&COUNTIF({WK}!I{W0}:I{W1},"Missed*")'; dash["C15"].font = font(italic=True, size=9)
    block(19, "100 Envelopes", [
        ("Total of all envelopes", f"={ENV}!B{E1 + 1}", MONEY),
        ("Envelopes saved", f"={ENV}!J16", "0"),
        ("Saved so far", f"={ENV}!J17", MONEY),
        ("Still to save", f"={ENV}!J18", MONEY),
        ("Progress", "=IFERROR(B22/B20,0)", PCT)])
    dash["C24"] = f'={bar("B24")}'; dash["C24"].font = Font(name=F, color=MAIN)
    K.section(dash, "A26", "Savings goals")
    K.hdr(dash, 27, ["Goal", "Target", "Progress", "Saved", "Still needed", "Needed per month", "Target date"])
    dash["H27"] = None
    for i in range(NG):
        r = 28 + i; g = G0 + i
        dash[f"A{r}"] = f'=IF({GL}!A{g}="","",{GL}!A{g})'
        dash[f"B{r}"] = f'=IF(A{r}="","",{GL}!B{g})'
        dash[f"C{r}"] = f'=IF(A{r}="","",{GL}!K{g})'
        dash[f"D{r}"] = f'=IF(A{r}="","",{GL}!E{g})'
        dash[f"E{r}"] = f'=IF(A{r}="","",{GL}!F{g})'
        dash[f"F{r}"] = f'=IF(A{r}="","",{GL}!I{g})'
        dash[f"G{r}"] = f'=IF(A{r}="","",{GL}!C{g})'
    style_range(dash, 28, 27 + NG, 1, 7, fmt={2: MONEY, 4: MONEY, 5: MONEY, 6: MONEY, 7: "mmm d, yyyy"})
    for r in range(28, 28 + NG): dash[f"C{r}"].font = Font(name=F, color=MAIN, size=9)
    ch = LineChart(); ch.title = "52-week challenge: planned vs saved"; ch.height = 8; ch.width = 18
    ch.add_data(Reference(wk, min_col=7, min_row=4, max_col=8, max_row=W1), titles_from_data=True)
    ch.set_categories(Reference(wk, min_col=1, min_row=W0, max_row=W1))
    ch.y_axis.majorGridlines = None
    for ser, col in zip(ch.series, (MAIN, "9CA3AF")):
        ser.graphicalProperties.line.solidFill = col; ser.graphicalProperties.line.width = 28000; ser.smooth = False
    dash.add_chart(ch, "A40")

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Settings tab: set your start date and choose your 52-week style (increasing, decreasing or the same amount every week) and amount. Classic challenge = increasing, amount 1 (1,378 in a year). Amount 5 = 6,890.", False),
             ("2.  52-Week Challenge: each week, save the planned deposit and type x in the 'Saved?' column. Saved a different amount? Type it in 'Actual amount'.", False),
             ("3.  100 Envelopes: pick any open envelope, save its amount and type x next to it. The grid turns green. Set the multiplier to 2 for the 10,000 version.", False),
             ("4.  Savings Goals: add your goals (emergency fund, vacation…) with a target and date, then log deposits on the Deposits tab. You see what you need to save per month and per week.", False),
             ("5.  Dashboard: total saved, progress bars, whether you are ahead or behind schedule, and a chart.", False),
             ("", False), ("Tips", True),
             ("•  Tight week? The decreasing style puts the big amounts in January, so the end of the year (holidays) is easy.", False),
             ("•  Missed a week? It shows in red as 'Missed – catch up'. Save it later and type x – it turns green.", False),
             ("•  Do one challenge, two or all three: the dashboard adds up whatever you use.", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  Works with any currency (change the label on the Settings tab).", False),
             ("•  To use Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("The ticks, envelopes and deposits are made-up examples (as-of date Oct 30, 2026). Use the blank copy in your download for your own challenges.", False)]
    K.start(st, "Savings Challenge Tracker", "52-week challenge, 100 envelopes and your own savings goals in one file. Works in Excel and Google Sheets.", lines)
    finish(wb, sample, path)


META = {
    "name": "Savings Challenge Tracker (52 Week, 100 Envelopes, Goals) – Excel & Google Sheets",
    "file_base": "Savings-Challenge-Tracker",
    "price": 12,
    "color": "#B45309",
    "title": "Savings Challenge Tracker",
    "tagline": "52-week challenge · 100 envelopes · savings goals — Excel & Google Sheets",
    "intro": "Make saving a game you actually finish. Track the 52-week money challenge, the 100 envelope challenge and your own savings goals in one file – tick a box and watch your progress bars fill up.",
    "features": [
        "52-week challenge: increasing, decreasing (reverse) or same-amount style, any weekly amount (1,378 / 6,890 / 13,780 and more)",
        "Tick each week when saved, log a different amount, see missed weeks and whether you are ahead or behind",
        "100 envelope challenge with a 10x10 grid that turns green, plus a multiplier for the 10,000 version",
        "Savings goals: up to 10 goals with target and date – shows what you need per month and per week",
        "Deposits log feeds the goals automatically",
        "Dashboard: total saved across all challenges, progress bars and a planned vs saved chart",
        "Any currency, no macros, blank copy plus a sample copy",
    ],
    "tags": ["savings challenge", "52 week challenge", "100 envelope challenge", "savings tracker", "money saving challenge", "google sheets",
             "excel template", "savings goal", "budget template", "sinking funds", "personal finance", "10k challenge", "money tracker"],
}
