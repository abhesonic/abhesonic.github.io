"""Baby Tracker – feeding, sleep, diapers, pumping, growth and appointments.

Settings -> Log -> Daily Summary -> Growth -> Health & Appointments -> Dashboard.
"""
import datetime as dt, random
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, ColorScaleRule
from openpyxl.chart import BarChart, LineChart, Reference
from _common import Kit, F, B, DATE, font, fill, style_range, input_cell, finish

MAIN = "7C3AED"; SOFT = "EDE9FE"
K = Kit(MAIN, SOFT)
NL = 3000; L0 = 6; L1 = L0 + NL - 1          # log rows
ND = 180; S0 = 6; S1 = S0 + ND - 1           # daily summary rows
NG = 60; G0 = 6; G1 = G0 + NG - 1            # growth rows
NH = 100; H0 = 6; H1 = H0 + NH - 1           # health rows
ACTS = ["Feed: breast left", "Feed: breast right", "Feed: bottle", "Feed: solids", "Sleep: nap", "Sleep: night",
        "Diaper: wet", "Diaper: dirty", "Diaper: wet + dirty", "Pump", "Tummy time", "Medicine", "Temperature", "Bath", "Other"]
TIME = "h:mm AM/PM"; NUM = '#,##0;-#,##0;"-"'; N1 = '0.0;-0.0;"-"'


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); lg = wb.create_sheet("Log"); ds = wb.create_sheet("Daily Summary")
    gr = wb.create_sheet("Growth"); he = wb.create_sheet("Health & Appointments"); s = wb.create_sheet("Settings")

    # ---------- Settings ----------
    K.title(s, "Settings", "Yellow cells are yours to edit.")
    SET = [("Baby's name", None), ("Date of birth", None), ("Feeding amount unit (ml or oz)", "ml"),
           ("Weight unit (kg or lb)", "kg"), ("Length unit (cm or in)", "cm"),
           ("Daily Summary starts on (blank = date of birth)", None), ("Dashboard day (blank = latest day in the Log)", None)]
    for i, (lab, v) in enumerate(SET, 5):
        s[f"A{i}"] = lab; s[f"A{i}"].font = font(bold=True); s[f"A{i}"].border = B; input_cell(s[f"C{i}"], v)
    for r in (6, 10, 11): s[f"C{r}"].number_format = DATE
    s["A13"] = "Used by the sheets (automatic)"; s["A13"].font = font(bold=True, color=MAIN)
    s["A14"] = "Daily Summary start"; s["A15"] = "Dashboard day"; s["A16"] = "Name shown"
    LD = f"Log!$A${L0}:$A${L1}"
    s["C14"] = f'=IF(ISNUMBER(C10),C10,IF(ISNUMBER(C6),C6,IF(COUNT({LD})>0,MIN({LD}),TODAY())))'
    s["C15"] = f'=IF(ISNUMBER(C11),C11,IF(COUNT({LD})>0,MAX({LD}),TODAY()))'
    s["C16"] = '=IF(C5="","Baby",C5)'
    for r in (14, 15, 16):
        s[f"A{r}"].font = font(italic=True); s[f"C{r}"].fill = fill("F3F4F6"); s[f"C{r}"].border = B; s[f"C{r}"].font = font()
    s["C14"].number_format = DATE; s["C15"].number_format = DATE
    K.hdr(s, 5, ["Activity list (used in the Log drop-down)"], col0=6)
    for i, a in enumerate(ACTS): s[f"F{6 + i}"] = a; s[f"F{6 + i}"].font = font(); s[f"F{6 + i}"].border = B
    s["F22"] = "Keep the words before ':' (Feed, Sleep, Diaper) – the totals use them."; s["F22"].font = font(italic=True, size=9)
    for c, w in zip("ABCDEF", [46, 3, 14, 3, 3, 34]): s.column_dimensions[c].width = w
    BIRTH = "Settings!$C$6"; AU = "Settings!$C$7"; WU = "Settings!$C$8"; LU = "Settings!$C$9"
    DAY = "Settings!$C$15"; NAME = "Settings!$C$16"

    # ---------- Log ----------
    K.title(lg, "Log", "One row per feed, nap, diaper, pump… Type the date, start time, end time (for feeds, naps, pumping) and pick the activity.")
    K.hdr(lg, 5, ["Date", "Start time", "End time", "Activity", "Amount (ml / oz)", "Note", "Minutes", "Start (date + time)"],
          [12, 11, 11, 20, 11, 36, 9, 17])
    style_range(lg, L0, L1, 1, 8, inputs=range(1, 7), fmt={1: DATE, 2: TIME, 3: TIME, 5: N1, 7: NUM, 8: 'yyyy-mm-dd h:mm;;'})
    for r in range(L0, L1 + 1):
        lg[f"G{r}"] = f'=IF(OR(A{r}="",NOT(ISNUMBER(B{r})),NOT(ISNUMBER(C{r}))),"",ROUND(MOD(C{r}-B{r},1)*1440,0))'
        lg[f"H{r}"] = f'=IF(OR(NOT(ISNUMBER(A{r})),NOT(ISNUMBER(B{r}))),0,A{r}+B{r})'
    dv = DataValidation(type="list", formula1="=Settings!$F$6:$F$20", allow_blank=True); lg.add_data_validation(dv); dv.add(f"D{L0}:D{L1}")
    for pre, col in (("Feed", "DBEAFE"), ("Sleep", "EDE9FE"), ("Diaper", "FEF3C7")):
        lg.conditional_formatting.add(f"D{L0}:D{L1}", FormulaRule(formula=[f'LEFT($D{L0},{len(pre)})="{pre}"'], fill=fill(col)))
    lg.freeze_panes = f"A{L0}"; lg.auto_filter.ref = f"A5:H{L1}"

    # ---------- Daily Summary ----------
    LGc = lambda c: f"Log!${c}${L0}:${c}${L1}"
    K.title(ds, "Daily Summary", "One line per day, filled in automatically from the Log. Sleep that starts before midnight counts for the day it started.")
    ds["A3"] = '="Amounts in "&Settings!$C$7'; ds["A3"].font = font(italic=True, size=9)
    cols = ["Date", "Day of life", "Feeds", "Bottle amount", "Solids feeds", "Breastfeeding minutes", "Sleep hours",
            "Naps", "Wet diapers", "Dirty diapers", "Pumped amount", "Tummy time (min)", "Medicine doses"]
    K.hdr(ds, 5, cols, [12, 9, 8, 11, 9, 13, 9, 7, 9, 9, 11, 11, 10])
    style_range(ds, S0, S1, 1, 13, fmt={1: "ddd yyyy-mm-dd", 2: NUM, 3: NUM, 4: NUM, 5: NUM, 6: NUM, 7: N1, 8: NUM,
                                         9: NUM, 10: NUM, 11: NUM, 12: NUM, 13: NUM})
    ds.column_dimensions["A"].width = 16
    for k in range(ND):
        r = S0 + k; d = f"$A{r}"; base = f'{LGc("A")},{d}'
        ds[f"A{r}"] = f"=Settings!$C$14+{k}"
        ds[f"B{r}"] = f'=IF(ISNUMBER({BIRTH}),A{r}-{BIRTH}+1,"")'
        ds[f"C{r}"] = f'=COUNTIFS({LGc("A")},{d},{LGc("D")},"Feed:*")'
        ds[f"D{r}"] = f'=SUMIFS({LGc("E")},{base},{LGc("D")},"Feed: bottle")'
        ds[f"E{r}"] = f'=COUNTIFS({base},{LGc("D")},"Feed: solids")'
        ds[f"F{r}"] = f'=SUMIFS({LGc("G")},{base},{LGc("D")},"Feed: breast*")'
        ds[f"G{r}"] = f'=SUMIFS({LGc("G")},{base},{LGc("D")},"Sleep:*")/60'
        ds[f"H{r}"] = f'=COUNTIFS({base},{LGc("D")},"Sleep: nap")'
        ds[f"I{r}"] = f'=COUNTIFS({base},{LGc("D")},"Diaper: wet*")'
        ds[f"J{r}"] = f'=COUNTIFS({base},{LGc("D")},"Diaper: dirty")+COUNTIFS({base},{LGc("D")},"Diaper: wet + dirty")'
        ds[f"K{r}"] = f'=SUMIFS({LGc("E")},{base},{LGc("D")},"Pump")'
        ds[f"L{r}"] = f'=SUMIFS({LGc("G")},{base},{LGc("D")},"Tummy time")'
        ds[f"M{r}"] = f'=COUNTIFS({base},{LGc("D")},"Medicine")'
    ds.conditional_formatting.add(f"A{S0}:M{S1}", FormulaRule(formula=[f'$A{S0}={DAY}'], font=Font(name=F, bold=True, color=MAIN), fill=fill(SOFT)))
    ds.conditional_formatting.add(f"G{S0}:G{S1}", ColorScaleRule(start_type="num", start_value=0, start_color="FFFFFF", end_type="num", end_value=18, end_color="C4B5FD"))
    ds.freeze_panes = f"B{S0}"

    # ---------- Growth ----------
    K.title(gr, "Growth", "Weight, length and head size from check-ups or home weigh-ins. Age and change are calculated.")
    K.hdr(gr, 5, ["Date", "Weight", "Length", "Head circumference", "Measured by / note", "Age (weeks)", "Weight change since last", "Change per week"],
          [12, 10, 10, 12, 30, 10, 13, 12])
    style_range(gr, G0, G1, 1, 8, inputs=range(1, 6), fmt={1: DATE, 2: "0.00", 3: "0.0", 4: "0.0", 6: "0.0", 7: '+0.00;-0.00;0.00', 8: '+0.00;-0.00;0.00'})
    for r in range(G0, G1 + 1):
        gr[f"F{r}"] = f'=IF(OR(NOT(ISNUMBER(A{r})),NOT(ISNUMBER({BIRTH}))),"",(A{r}-{BIRTH})/7)'
        if r == G0:
            gr[f"G{r}"] = '=""'; gr[f"H{r}"] = '=""'
        else:
            p = r - 1
            gr[f"G{r}"] = f'=IF(OR(NOT(ISNUMBER(B{r})),NOT(ISNUMBER(B{p}))),"",B{r}-B{p})'
            gr[f"H{r}"] = f'=IF(OR(G{r}="",NOT(ISNUMBER(A{r})),NOT(ISNUMBER(A{p})),A{r}<=A{p}),"",G{r}/((A{r}-A{p})/7))'
    gr["J5"] = '="Units: weight "&Settings!$C$8&", length "&Settings!$C$9'; gr["J5"].font = font(italic=True, size=9)
    gr["J6"] = "Add rows in date order (oldest first)."; gr["J6"].font = font(italic=True, size=9)
    gr.freeze_panes = f"A{G0}"

    # ---------- Health & Appointments ----------
    K.title(he, "Health & Appointments", "Check-ups, vaccines, sick visits and medicine. Copy dates from your doctor's schedule; mark Done = Yes afterwards.")
    K.hdr(he, 5, ["Date", "Time", "Type", "What", "Doctor / place", "Done?", "Notes", "Status", "(helper)"], [12, 10, 13, 30, 22, 8, 30, 26, 8])
    style_range(he, H0, H1, 1, 9, inputs=range(1, 8), fmt={1: DATE, 2: TIME})
    for r in range(H0, H1 + 1):
        he[f"H{r}"] = (f'=IF(OR(D{r}="",NOT(ISNUMBER(A{r}))),"",IF(F{r}="Yes","Done",IF(A{r}<TODAY(),"Past – not marked done",'
                       f'IF(A{r}=TODAY(),"Today",IF(A{r}-TODAY()<=14,"In "&(A{r}-TODAY())&" days","Planned")))))')
        he[f"I{r}"] = f'=IF(AND(ISNUMBER(A{r}),D{r}<>"",F{r}<>"Yes",N(A{r})>=TODAY()),A{r}+ROW()/100000,"")'
        he[f"H{r}"].font = font(bold=True); he[f"I{r}"].font = font(color="9CA3AF", size=8)
    for col, lst in (("C", '"Check-up,Vaccine,Sick visit,Medicine,Dentist,Other"'), ("F", '"Yes,No"')):
        dv = DataValidation(type="list", formula1=lst, allow_blank=True); he.add_data_validation(dv); dv.add(f"{col}{H0}:{col}{H1}")
    he.conditional_formatting.add(f"H{H0}:H{H1}", FormulaRule(formula=[f'LEFT(H{H0},4)="Past"'], font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    he.conditional_formatting.add(f"H{H0}:H{H1}", FormulaRule(formula=[f'OR(H{H0}="Today",LEFT(H{H0},3)="In ")'], font=Font(name=F, bold=True, color="92400E"), fill=fill("FEF3C7")))
    he.conditional_formatting.add(f"H{H0}:H{H1}", FormulaRule(formula=[f'H{H0}="Done"'], font=Font(name=F, bold=True, color="166534")))
    he.freeze_panes = f"A{H0}"; he.auto_filter.ref = f"A5:H{H1}"

    # ---------- sample ----------
    if sample:
        rnd = random.Random(7)
        s["C5"] = "Lily"; s["C6"] = dt.date(2026, 9, 20); s["C10"] = dt.date(2026, 11, 1)
        rows = []
        T = lambda h, m: dt.time(h % 24, m % 60)
        for dd in range(14):
            day = dt.date(2026, 11, 1) + dt.timedelta(days=dd)
            # feeds every ~3 h
            t = 6 * 60 + rnd.randint(0, 30)
            side = 0
            while t < 23 * 60:
                h, m = divmod(t, 60)
                if t // 60 in (12, 18):
                    st_ = T(h, m); dur = rnd.randint(12, 20); e = divmod(t + dur, 60)
                    rows.append((day, st_, T(*e), "Feed: bottle", rnd.choice([90, 100, 110, 120]), "Expressed milk" if t // 60 == 12 else ""))
                else:
                    dur = rnd.randint(12, 22); e = divmod(t + dur, 60)
                    rows.append((day, T(h, m), T(*e), "Feed: breast left" if side % 2 == 0 else "Feed: breast right", None, "")); side += 1
                # diaper after most feeds
                d2 = divmod(t + 25, 60)
                kind = rnd.choice(["Diaper: wet", "Diaper: wet", "Diaper: wet", "Diaper: wet + dirty", "Diaper: dirty"])
                rows.append((day, T(*d2), None, kind, None, ""))
                # nap after daytime feeds
                if 8 * 60 <= t <= 18 * 60:
                    ns = t + 40; nd = ns + rnd.randint(35, 100)
                    rows.append((day, T(*divmod(ns, 60)), T(*divmod(nd, 60)), "Sleep: nap", None, ""))
                t += rnd.randint(140, 215)
            for _ in range(rnd.choice([0, 1, 1, 2])):
                cf = rnd.randint(14 * 60, 20 * 60)
                rows.append((day, T(*divmod(cf, 60)), T(*divmod(cf + rnd.randint(6, 12), 60)), "Feed: breast right", None, "Cluster feed"))
            # night sleep (crosses midnight) + one night feed
            ns = 21 * 60 + rnd.randint(0, 40); ne = ns + rnd.randint(250, 330)
            rows.append((day, T(*divmod(ns, 60)), T(*divmod(ne, 60)), "Sleep: night", None, "Longest stretch" if dd == 12 else ""))
            nf = ne + 5
            rows.append((day, T(*divmod(nf, 60)), T(*divmod(nf + 18, 60)), "Feed: breast left", None, "Night feed"))
            ns2 = nf + 25; ne2 = ns2 + rnd.randint(150, 200)
            rows.append((day, T(*divmod(ns2, 60)), T(*divmod(ne2, 60)), "Sleep: night", None, ""))
            pt = 9 * 60 + rnd.randint(0, 40)
            rows.append((day, T(*divmod(pt, 60)), T(*divmod(pt + 15, 60)), "Pump", rnd.choice([80, 90, 100, 120]), ""))
            tt = 11 * 60 + rnd.randint(0, 60)
            rows.append((day, T(*divmod(tt, 60)), T(*divmod(tt + rnd.randint(5, 12), 60)), "Tummy time", None, ""))
            if dd % 2 == 0: rows.append((day, T(19, 0), None, "Bath", None, ""))
            if dd in (5, 6): rows.append((day, T(20, 15), None, "Medicine", 2.5, "Infant paracetamol (doctor's advice)"))
            if dd == 5: rows.append((day, T(20, 10), None, "Temperature", 37.9, "Day after vaccines"))
        rows.sort(key=lambda x: (x[0], x[1]))
        for k, row in enumerate(rows):
            for c, v in zip("ABCDEF", row):
                if v not in (None, ""): lg[f"{c}{L0 + k}"] = v
        growth = [(dt.date(2026, 9, 20), 3.35, 50.0, 34.5, "Hospital – birth"), (dt.date(2026, 9, 24), 3.18, None, None, "Day 5 midwife"),
                  (dt.date(2026, 10, 4), 3.52, 51.5, 35.4, "2-week check"), (dt.date(2026, 10, 20), 4.05, 53.5, 36.6, "Home scale"),
                  (dt.date(2026, 11, 5), 4.62, 55.8, 37.8, "Pediatrician"), (dt.date(2026, 11, 13), 4.88, None, None, "Home scale")]
        for k, row in enumerate(growth):
            for c, v in zip("ABCDE", row):
                if v is not None: gr[f"{c}{G0 + k}"] = v
        health = [(dt.date(2026, 9, 22), dt.time(10, 0), "Check-up", "Newborn check", "Dr. Patel, Maple Pediatrics", "Yes", ""),
                  (dt.date(2026, 10, 4), dt.time(9, 30), "Check-up", "2-week check-up", "Dr. Patel", "Yes", "Back to birth weight"),
                  (dt.date(2026, 11, 5), dt.time(11, 0), "Vaccine", "2-month vaccines", "Dr. Patel", "Yes", "Mild fever next evening"),
                  (dt.date(2026, 11, 5), dt.time(11, 0), "Check-up", "2-month check-up", "Dr. Patel", "Yes", "Weight on track"),
                  (dt.date(2026, 11, 24), dt.time(14, 0), "Other", "Hearing re-test", "City Hospital", "No", "Bring red book"),
                  (dt.date(2027, 1, 14), dt.time(10, 30), "Check-up", "4-month check-up", "Dr. Patel", "No", ""),
                  (dt.date(2027, 1, 14), dt.time(10, 30), "Vaccine", "4-month vaccines", "Dr. Patel", "No", "Ask about dose timing")]
        for k, row in enumerate(health):
            for c, v in zip("ABCDEFG", row): he[f"{c}{H0 + k}"] = v

    # ---------- Dashboard ----------
    K.title(dash, "Baby Tracker Dashboard", "Updates by itself from the Log, Growth and Health tabs.")
    dash["A3"] = (f'={NAME}&IF(ISNUMBER({BIRTH}),"  ·  "&INT((TODAY()-{BIRTH})/7)&" weeks old ("&DATEDIF({BIRTH},TODAY(),"m")&" months "'
                  f'&DATEDIF({BIRTH},TODAY(),"md")&" days)","")&"  ·  day shown: "&TEXT({DAY},"ddd mmm d, yyyy")')
    dash["A3"].font = font(bold=True, size=11)
    SD = lambda c: f"'Daily Summary'!${c}${S0}:${c}${S1}"
    one = lambda c: f"SUMIFS({SD(c)},{SD('A')},{DAY})"
    wk = lambda c: f"SUMIFS({SD(c)},{SD('A')},\">=\"&({DAY}-6),{SD('A')},\"<=\"&{DAY})/7"
    lastf = f'SUMPRODUCT(MAX((LEFT({LGc("D")},5)="Feed:")*{LGc("H")}))'
    lastd = f'SUMPRODUCT(MAX((LEFT({LGc("D")},7)="Diaper:")*{LGc("H")}))'
    GW = f"Growth!$B${G0}:$B${G1}"; GA = f"Growth!$A${G0}:$A${G1}"
    lastw = f'IF(COUNT({GW})=0,"",LOOKUP(2,1/ISNUMBER({GW}),{GW}))'
    kp = [("Feeds (day shown)", f"={one('C')}", NUM), ("Sleep hours (day shown)", f"={one('G')}", N1),
          ("Wet / dirty diapers (day shown)", f'={one("I")}&" / "&{one("J")}', "@"),
          ("Avg feeds / day (last 7 days)", f"={wk('C')}", N1), ("Avg sleep hours / day (7 days)", f"={wk('G')}", N1),
          ("Avg diapers / day (7 days)", f"=({wk('I')})+({wk('J')})", N1),
          ("Bottle amount (day shown)", f'=TEXT({one("D")},"#,##0")&" "&{AU}', "@"),
          ("Last feed", f'=IF({lastf}=0,"",{lastf})', "ddd h:mm AM/PM"),
          ("Last diaper", f'=IF({lastd}=0,"",{lastd})', "ddd h:mm AM/PM"),
          ("Latest weight", f'=IF(COUNT({GW})=0,"",TEXT({lastw},"0.00")&" "&{WU})', "@"),
          ("Weight change since first", f'=IF(COUNT({GW})<2,"",TEXT({lastw}-INDEX({GW},MATCH(TRUE,INDEX(ISNUMBER({GW}),0),0)),"+0.00;-0.00")&" "&{WU})', "@"),
          ("Next appointment", f"=IF(MIN('Health & Appointments'!$I${H0}:$I${H1})=0,\"none planned\",TEXT(INT(MIN('Health & Appointments'!$I${H0}:$I${H1})),\"mmm d\")&\": \"&INDEX('Health & Appointments'!$D${H0}:$D${H1},MATCH(MIN('Health & Appointments'!$I${H0}:$I${H1}),'Health & Appointments'!$I${H0}:$I${H1},0)))", "@")]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 3) * 3, 1 + (i % 3) * 2, lab, f, fmt)
        c = dash.cell(row=6 + (i // 3) * 3, column=1 + (i % 3) * 2); c.alignment = Alignment(horizontal="left")
        if i in (2, 6, 9, 10, 11): c.font = Font(name=F, size=14, bold=True, color=MAIN)
    dash["A17"] = '="Time since last feed (right now): "&IF(' + lastf + '=0,"–",TEXT(NOW()-' + lastf + ',"[h]:mm")&" h")'
    dash["A17"].font = font(italic=True, size=9)
    for c, w in zip("ABCDEFGH", [26, 12, 26, 12, 26, 12, 3, 3]): dash.column_dimensions[c].width = w
    H = 20
    K.section(dash, f"A{H - 1}", "Last 14 days")
    K.hdr(dash, H, ["Date", "Feeds", "Sleep hours", "Wet", "Dirty", "Bottle amount"])
    for k in range(14):
        r = H + 1 + k
        dash[f"A{r}"] = f"={DAY}-{13 - k}"
        dash[f"B{r}"] = f"=SUMIFS({SD('C')},{SD('A')},A{r})"; dash[f"C{r}"] = f"=SUMIFS({SD('G')},{SD('A')},A{r})"
        dash[f"D{r}"] = f"=SUMIFS({SD('I')},{SD('A')},A{r})"; dash[f"E{r}"] = f"=SUMIFS({SD('J')},{SD('A')},A{r})"
        dash[f"F{r}"] = f"=SUMIFS({SD('D')},{SD('A')},A{r})"
    style_range(dash, H + 1, H + 14, 1, 6, fmt={1: "ddd mmm d", 2: NUM, 3: N1, 4: NUM, 5: NUM, 6: NUM})
    dash.conditional_formatting.add(f"A{H + 1}:F{H + 14}", FormulaRule(formula=[f'$A{H + 1}={DAY}'], font=Font(name=F, bold=True, color=MAIN), fill=fill(SOFT)))
    ch = BarChart(); ch.type = "col"; ch.title = "Feeds per day"; ch.height = 7; ch.width = 15; ch.legend = None
    ch.add_data(Reference(dash, min_col=2, min_row=H, max_row=H + 14), titles_from_data=True)
    ch.set_categories(Reference(dash, min_col=1, min_row=H + 1, max_row=H + 14)); ch.y_axis.majorGridlines = None
    ch.series[0].graphicalProperties.solidFill = "A78BFA"; ch.gapWidth = 50; ch.x_axis.number_format = "mmm d"
    dash.add_chart(ch, "I4")
    ch2 = BarChart(); ch2.type = "col"; ch2.title = "Sleep hours per day"; ch2.height = 7; ch2.width = 15; ch2.legend = None
    ch2.add_data(Reference(dash, min_col=3, min_row=H, max_row=H + 14), titles_from_data=True)
    ch2.set_categories(Reference(dash, min_col=1, min_row=H + 1, max_row=H + 14)); ch2.y_axis.majorGridlines = None
    ch2.y_axis.scaling.min = 0; ch.y_axis.scaling.min = 0; ch2.series[0].graphicalProperties.solidFill = MAIN; ch2.gapWidth = 50; ch2.x_axis.number_format = "mmm d"
    dash.add_chart(ch2, "I19")
    ch3 = LineChart(); ch3.title = "Weight"; ch3.height = 7; ch3.width = 15; ch3.legend = None
    ch3.add_data(Reference(gr, min_col=2, min_row=5, max_row=G0 + 19), titles_from_data=True)
    ch3.set_categories(Reference(gr, min_col=1, min_row=G0, max_row=G0 + 19)); ch3.y_axis.majorGridlines = None
    ch3.series[0].graphicalProperties.line.solidFill = "DB2777"; ch3.series[0].graphicalProperties.line.width = 28000
    ch3.x_axis.number_format = "mmm d"; ch3.series[0].smooth = False
    dash.add_chart(ch3, "I34")
    for c in (ch, ch2, ch3):
        c.roundedCorners = False; c.x_axis.delete = False; c.y_axis.delete = False
        if c.title: pass

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Settings tab: your baby's name, date of birth and the units you use (ml or oz, kg or lb, cm or in).", False),
             ("2.  Log tab: add one row for every feed, nap, night sleep, diaper, pumping session, tummy time, medicine dose or bath.", False),
             ("     Type the date and start time. For feeds, sleep, pumping and tummy time also type the end time – minutes are calculated. Bottle and pumping: type the amount.", False),
             ("     Tip: in Excel press Ctrl+;  for today's date and Ctrl+Shift+;  for the current time. In Google Sheets: Ctrl+;  and Ctrl+Shift+;  as well.", False),
             ("3.  Daily Summary: feeds, bottle amount, breastfeeding minutes, sleep hours, naps, wet and dirty diapers, pumping – one line per day.", False),
             ("4.  Growth tab: weight, length and head size. Age in weeks and weight change per week are calculated.", False),
             ("5.  Health & Appointments: check-ups, vaccines and doctor visits from your doctor's schedule. Status shows what is coming up.", False),
             ("6.  Dashboard: today at a glance, 7-day averages, last feed and diaper, latest weight, next appointment and charts.", False),
             ("", False), ("Tips", True),
             ("•  The Dashboard shows the latest day in the Log. To look at another day, type it in Settings > 'Dashboard day'.", False),
             ("•  Night sleep that starts before midnight and ends after it counts for the day it started.", False),
             ("•  Print the Daily Summary or Growth tab before a doctor's visit – it answers most of their questions.", False),
             ("•  This sheet is a diary, not medical advice. Always follow your doctor or nurse.", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  3,000 log rows, 180 days of summary, 60 growth measurements, 100 appointments. No macros. Works in Excel 2010+ and Google Sheets.", False),
             ("•  To use Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("Two weeks of made-up entries for a baby called Lily. Use the blank copy in your download for your own baby.", False)]
    K.start(st, "Baby Tracker", "Feeding, sleep, diapers, pumping, growth and appointments – Excel and Google Sheets.", lines)
    finish(wb, sample, path)


META = {
    "name": "Baby Tracker (Newborn Feeding, Sleep, Diaper, Pumping & Growth Log) – Excel & Google Sheets",
    "file_base": "Baby-Tracker",
    "price": 12,
    "color": "#7C3AED",
    "title": "Baby Tracker",
    "tagline": "feeding · sleep · diapers · pumping · growth · appointments",
    "intro": "Keep every feed, nap, diaper and pumping session in one simple log – no app, no subscription, no ads. The tracker turns your entries into a daily summary, 7-day averages and charts, shows when the last feed and diaper were, follows your baby's weight and lists upcoming check-ups. Ideal for new parents, partners taking turns, nannies and grandparents.",
    "features": [
        "One easy log for breastfeeding (left/right), bottle, solids, naps, night sleep, wet and dirty diapers, pumping, tummy time, medicine, temperature and baths",
        "Type start and end time – minutes are calculated, night sleep across midnight handled",
        "Daily Summary: feeds, bottle amount, breastfeeding minutes, sleep hours, naps, wet and dirty diapers, pumped amount per day",
        "Dashboard: day at a glance, 7-day averages, last feed and last diaper, latest weight, next appointment",
        "Charts: feeds per day, sleep hours per day (last 14 days) and weight curve",
        "Growth log: weight, length, head size, age in weeks and weight change per week",
        "Health & appointments list with Done / Today / In X days / Past status",
        "Works in ml or oz, kg or lb, cm or in",
        "Blank copy plus a sample copy with two weeks of example entries",
    ],
    "tags": ["baby tracker", "newborn tracker", "feeding log", "baby sleep tracker", "diaper log", "breastfeeding tracker",
             "google sheets", "excel template", "new mom gift", "baby schedule", "pumping log", "baby growth chart"],
    "summary": "Instant download: Excel + Google Sheets template, blank copy + sample copy.",
}
