"""Pet Care & Expense Tracker – pet profiles, expenses vs budget, vaccine/flea/vet reminders, weight log.

Settings -> Pets -> Expenses -> Care Schedule -> Weight Log -> Monthly View -> Dashboard.
"""
import datetime as dt, random
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, DataBarRule
from openpyxl.chart import BarChart, Reference
from _common import Kit, F, B, MONEY, DATE, font, fill, style_range, input_cell, finish

MAIN = "C2410C"; SOFT = "FFEDD5"
K = Kit(MAIN, SOFT)
NP = 20; P0 = 6; P1 = P0 + NP - 1
NE = 1500; E0 = 6; E1 = E0 + NE - 1
NC = 150; C0 = 6; C1 = C0 + NC - 1
NW = 300; W0 = 6; W1 = W0 + NW - 1
NK = 15; K0 = 6; K1 = K0 + NK - 1
CATS = [("Food & treats", 120), ("Vet visits", 60), ("Medication & preventives", 45), ("Insurance", 55), ("Grooming", 35),
        ("Toys & supplies", 25), ("Boarding & pet sitting", 30), ("Training & classes", 15), ("License & registration", 3), ("Other", 10)]
CARE = ["Vaccine", "Flea & tick", "Deworming", "Heartworm", "Check-up", "Dental", "Grooming", "Nail trim", "Medication refill",
        "License renewal", "Other"]
SPECIES = '"Dog,Cat,Bird,Rabbit,Guinea pig,Hamster,Fish,Reptile,Horse,Other"'
NUM = '0'


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); pe = wb.create_sheet("Pets"); ex = wb.create_sheet("Expenses")
    ca = wb.create_sheet("Care Schedule"); wl = wb.create_sheet("Weight Log"); mv = wb.create_sheet("Monthly View")
    s = wb.create_sheet("Settings")

    # ---------- Settings ----------
    K.title(s, "Settings", "Yellow cells are yours to edit.")
    SET = [("Currency (label only)", "USD"), ("Year to show (blank = this year)", None), ("Weight unit (kg or lb)", "lb")]
    for i, (lab, v) in enumerate(SET, 5):
        s[f"A{i}"] = lab; s[f"A{i}"].font = font(bold=True); s[f"A{i}"].border = B; input_cell(s[f"C{i}"], v)
    s["A9"] = "Year shown (automatic)"; s["A9"].font = font(italic=True)
    s["C9"] = "=IF(ISNUMBER(C6),C6,YEAR(TODAY()))"; s["C9"].number_format = "0"
    s["A10"] = "Months counted for averages"; s["A10"].font = font(italic=True)
    s["C10"] = "=IF(C9<YEAR(TODAY()),12,IF(C9>YEAR(TODAY()),1,MONTH(TODAY())))"
    s["A11"] = "Remind me this many days before due"; s["A11"].font = font(bold=True); s["A11"].border = B; input_cell(s["C11"], 30)
    for r in (9, 10): s[f"C{r}"].fill = fill("F3F4F6"); s[f"C{r}"].border = B; s[f"C{r}"].font = font()
    K.hdr(s, 5, ["Expense category", "Budget / month"], col0=6)
    for i in range(NK):
        n, b = CATS[i] if i < len(CATS) else (None, None)
        input_cell(s[f"F{K0 + i}"], n); input_cell(s[f"G{K0 + i}"], b); s[f"G{K0 + i}"].number_format = MONEY
    s[f"F{K1 + 1}"] = "Total / month"; s[f"F{K1 + 1}"].font = font(bold=True)
    s[f"G{K1 + 1}"] = f"=SUM(G{K0}:G{K1})"; s[f"G{K1 + 1}"].font = font(bold=True); s[f"G{K1 + 1}"].number_format = MONEY
    for c, w in zip("ABCDEFG", [40, 3, 12, 3, 3, 28, 14]): s.column_dimensions[c].width = w
    YR = "Settings!$C$9"; MC = "Settings!$C$10"; WU = "Settings!$C$7"; LEAD = "Settings!$C$11"

    EX = lambda c: f"Expenses!${c}${E0}:${c}${E1}"
    CS = lambda c: f"'Care Schedule'!${c}${C0}:${c}${C1}"
    WL = lambda c: f"'Weight Log'!${c}${W0}:${c}${W1}"

    # ---------- Pets ----------
    K.title(pe, "Pets", "One row per pet. Yellow = you type, grey = automatic.")
    cols = ["Name", "Species", "Breed", "Birthday", "Sex", "Microchip number", "Vet clinic", "Vet phone", "Insurance (company & policy)",
            "Food & portion", "Allergies / notes", "Age", "Latest weight", "Spent this year", "Spent all time",
            "Care items overdue", "Next care due", "Next care item"]
    K.hdr(pe, 5, cols, [14, 10, 18, 11, 8, 17, 20, 14, 24, 24, 24, 14, 10, 13, 13, 10, 12, 24])
    style_range(pe, P0, P1, 1, 18, inputs=range(1, 12), fmt={4: DATE, 13: "0.0", 14: MONEY, 15: MONEY, 16: NUM, 17: DATE})
    for r in range(P0, P1 + 1):
        pe[f"L{r}"] = (f'=IF(OR(A{r}="",NOT(ISNUMBER(D{r}))),"",IF(DATEDIF(D{r},TODAY(),"y")>0,DATEDIF(D{r},TODAY(),"y")&" y ","")'
                       f'&DATEDIF(D{r},TODAY(),"ym")&" m")')
        pe[f"M{r}"] = f'=IF(A{r}="","",IF(COUNTIFS({WL("B")},A{r},{WL("C")},">0")=0,"",LOOKUP(2,1/(({WL("B")}=A{r})*ISNUMBER({WL("C")})),{WL("C")})))'
        pe[f"N{r}"] = f'=IF(A{r}="","",SUMIFS({EX("F")},{EX("B")},A{r},{EX("G")},{YR}))'
        pe[f"O{r}"] = f'=IF(A{r}="","",SUMIFS({EX("F")},{EX("B")},A{r}))'
        pe[f"P{r}"] = f'=IF(A{r}="","",COUNTIFS({CS("A")},A{r},{CS("J")},"Overdue*"))'
        pe[f"Q{r}"] = f'=IF(A{r}="","",IF(SUMPRODUCT(MIN(({CS("A")}=A{r})*{CS("L")}+({CS("A")}<>A{r})*99999))>=99999,"",SUMPRODUCT(MIN(({CS("A")}=A{r})*{CS("L")}+({CS("A")}<>A{r})*99999))))'
        pe[f"R{r}"] = (f'=IF(Q{r}="","",INDEX({CS("C")},MATCH(1,INDEX(({CS("A")}=A{r})*({CS("H")}=Q{r}),0),0))&'
                       f'" ("&INDEX({CS("B")},MATCH(1,INDEX(({CS("A")}=A{r})*({CS("H")}=Q{r}),0),0))&")")')
    for col, lst in (("B", SPECIES), ("E", '"Male,Female,Male (neutered),Female (spayed)"')):
        dv = DataValidation(type="list", formula1=lst, allow_blank=True); pe.add_data_validation(dv); dv.add(f"{col}{P0}:{col}{P1}")
    pe.conditional_formatting.add(f"P{P0}:P{P1}", FormulaRule(formula=[f'N(P{P0})>0'], font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    pe.conditional_formatting.add(f"Q{P0}:Q{P1}", FormulaRule(formula=[f'AND(ISNUMBER(Q{P0}),Q{P0}<TODAY())'], font=Font(name=F, bold=True, color="B91C1C")))
    pe.freeze_panes = f"B{P0}"

    PETLIST = f"=Pets!$A${P0}:$A${P1}"

    # ---------- Expenses ----------
    K.title(ex, "Expenses", "One row per purchase or bill. Pick the pet and category from the lists. For shared costs choose 'Shared'.")
    K.hdr(ex, 5, ["Date", "Pet", "Category", "What", "Shop / vet", "Amount", "Year", "Month"], [12, 14, 24, 32, 22, 12, 7, 7])
    style_range(ex, E0, E1, 1, 8, inputs=range(1, 7), fmt={1: DATE, 6: MONEY, 7: "0", 8: "0"})
    for r in range(E0, E1 + 1):
        ex[f"G{r}"] = f'=IF(ISNUMBER(A{r}),YEAR(A{r}),"")'; ex[f"H{r}"] = f'=IF(ISNUMBER(A{r}),MONTH(A{r}),"")'
    ex["J5"] = "Pet list (+ Shared)"; ex["J5"].font = font(bold=True, size=8, color="9CA3AF")
    for i in range(NP): ex[f"J{6 + i}"] = f'=IF(Pets!A{P0 + i}="","",Pets!A{P0 + i})'; ex[f"J{6 + i}"].font = font(size=8, color="9CA3AF")
    ex[f"J{6 + NP}"] = "Shared"; ex[f"J{6 + NP}"].font = font(size=8, color="9CA3AF")
    dv = DataValidation(type="list", formula1=f"=$J$6:$J${6 + NP}", allow_blank=True); ex.add_data_validation(dv); dv.add(f"B{E0}:B{E1}")
    dv = DataValidation(type="list", formula1=f"=Settings!$F${K0}:$F${K1}", allow_blank=True); ex.add_data_validation(dv); dv.add(f"C{E0}:C{E1}")
    ex.freeze_panes = f"A{E0}"; ex.auto_filter.ref = f"A5:H{E1}"

    # ---------- Care Schedule ----------
    K.title(ca, "Care Schedule", "Vaccines, flea & tick, deworming, check-ups, grooming… Type when it was last done and how often it repeats – the next due date is calculated.")
    K.hdr(ca, 5, ["Pet", "Type", "What", "Last done", "Repeat every (months)", "Notes", "Or type next due date", "Next due",
                  "Days left", "Status", "(helper)"], [14, 17, 28, 12, 11, 28, 13, 12, 9, 22, 8])
    style_range(ca, C0, C1, 1, 11, inputs=range(1, 8), fmt={4: DATE, 5: "0", 7: DATE, 8: DATE, 9: '0;-0;0'})
    for r in range(C0, C1 + 1):
        ca[f"H{r}"] = f'=IF(OR(A{r}="",C{r}=""),"",IF(ISNUMBER(G{r}),G{r},IF(AND(ISNUMBER(D{r}),N(E{r})>0),EDATE(D{r},E{r}),"")))'
        ca[f"I{r}"] = f'=IF(ISNUMBER(H{r}),H{r}-TODAY(),"")'
        ca[f"J{r}"] = (f'=IF(OR(A{r}="",C{r}=""),"",IF(NOT(ISNUMBER(H{r})),IF(ISNUMBER(D{r}),"Done – no repeat","Type a date"),'
                       f'IF(I{r}<0,"Overdue by "&-I{r}&IF(I{r}=-1," day"," days"),IF(I{r}=0,"Due today",IF(I{r}<={LEAD},"Due in "&I{r}&IF(I{r}=1," day"," days"),"OK")))))')
        ca[f"K{r}"] = f'=IF(ISNUMBER(H{r}),H{r}+ROW()/100000,"")'
        ca[f"L{r}"] = f'=IF(ISNUMBER(H{r}),H{r},99999)'; ca[f"L{r}"].font = font(color="FFFFFF", size=6)
        ca[f"J{r}"].font = font(bold=True); ca[f"K{r}"].font = font(color="9CA3AF", size=8)
    dv = DataValidation(type="list", formula1=PETLIST, allow_blank=True); ca.add_data_validation(dv); dv.add(f"A{C0}:A{C1}")
    dv = DataValidation(type="list", formula1='"' + ",".join(CARE) + '"', allow_blank=True); ca.add_data_validation(dv); dv.add(f"B{C0}:B{C1}")
    ca.conditional_formatting.add(f"J{C0}:J{C1}", FormulaRule(formula=[f'LEFT(J{C0},7)="Overdue"'], font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    ca.conditional_formatting.add(f"J{C0}:J{C1}", FormulaRule(formula=[f'LEFT(J{C0},3)="Due"'], font=Font(name=F, bold=True, color="92400E"), fill=fill("FEF3C7")))
    ca.conditional_formatting.add(f"J{C0}:J{C1}", FormulaRule(formula=[f'J{C0}="OK"'], font=Font(name=F, bold=True, color="166534")))
    ca.freeze_panes = f"A{C0}"; ca.auto_filter.ref = f"A5:J{C1}"

    # ---------- Weight Log ----------
    K.title(wl, "Weight Log", "Weigh-ins at home or at the vet. Add new rows at the bottom (oldest first).")
    K.hdr(wl, 5, ["Date", "Pet", "Weight", "Note", "Change since this pet's previous weigh-in"], [12, 14, 10, 30, 18])
    style_range(wl, W0, W1, 1, 5, inputs=range(1, 5), fmt={1: DATE, 3: "0.0", 5: '+0.0;-0.0;0.0'})
    for r in range(W0, W1 + 1):
        if r == W0: wl[f"E{r}"] = '=""'; continue
        rng = lambda c: f"${c}${W0}:${c}{r - 1}"
        wl[f"E{r}"] = (f'=IF(OR(B{r}="",NOT(ISNUMBER(C{r}))),"",IF(COUNTIFS({rng("B")},B{r},{rng("C")},">0")=0,"",'
                       f'C{r}-LOOKUP(2,1/(({rng("B")}=B{r})*ISNUMBER({rng("C")})),{rng("C")})))')
    wl["G5"] = '="Unit: "&Settings!$C$7'; wl["G5"].font = font(italic=True, size=9)
    dv = DataValidation(type="list", formula1=PETLIST, allow_blank=True); wl.add_data_validation(dv); dv.add(f"B{W0}:B{W1}")
    wl.freeze_panes = f"A{W0}"

    # ---------- Monthly View ----------
    K.title(mv, "Monthly View", "Spending per category and per pet for each month of the year chosen in Settings.")
    mv["A3"] = '="Year "&Settings!$C$9&"  ·  amounts in "&Settings!$C$5'; mv["A3"].font = font(bold=True)
    MN = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    K.hdr(mv, 5, ["Category"] + MN + ["Year total", "Budget / year", "Left (+) / over (-)"], [26] + [9] * 12 + [12, 12, 14])
    for k in range(NK):
        r = 6 + k; kr = K0 + k
        mv[f"A{r}"] = f'=IF(Settings!$F${kr}="","",Settings!$F${kr})'
        for j in range(12):
            mv[f"{CL(2 + j)}{r}"] = f'=IF($A{r}="",0,SUMIFS({EX("F")},{EX("C")},$A{r},{EX("G")},{YR},{EX("H")},{j + 1}))'
        mv[f"N{r}"] = f"=SUM(B{r}:M{r})"
        mv[f"O{r}"] = f'=IF($A{r}="","",N(Settings!$G${kr})*12)'
        mv[f"P{r}"] = f'=IF($A{r}="","",O{r}-N{r})'
    TR = 6 + NK
    mv[f"A{TR}"] = "Total"
    for j in range(15): mv[f"{CL(2 + j)}{TR}"] = f"=SUM({CL(2 + j)}6:{CL(2 + j)}{TR - 1})"
    style_range(mv, 6, TR, 1, 16, fmt={c: MONEY for c in range(2, 17)})
    for c in range(1, 17):
        cell = mv.cell(row=TR, column=c); cell.font = Font(name=F, bold=True, color=MAIN); cell.fill = fill(SOFT)
    mv.conditional_formatting.add(f"P6:P{TR}", FormulaRule(formula=['AND(ISNUMBER(P6),P6<0)'], font=Font(name=F, bold=True, color="B91C1C")))
    mv[f"A{TR + 1}"] = "Budget / month"
    for j in range(12): mv[f"{CL(2 + j)}{TR + 1}"] = f"=Settings!$G${K1 + 1}"
    style_range(mv, TR + 1, TR + 1, 1, 13, fmt={c: MONEY for c in range(2, 14)})
    PR0 = TR + 4
    K.hdr(mv, PR0 - 1, ["Pet"] + MN + ["Year total"])
    for k in range(NP + 1):
        r = PR0 + k
        mv[f"A{r}"] = f'=IF(Pets!$A${P0 + k}="","",Pets!$A${P0 + k})' if k < NP else "Shared"
        for j in range(12):
            mv[f"{CL(2 + j)}{r}"] = f'=IF($A{r}="",0,SUMIFS({EX("F")},{EX("B")},$A{r},{EX("G")},{YR},{EX("H")},{j + 1}))'
        mv[f"N{r}"] = f"=SUM(B{r}:M{r})"
    style_range(mv, PR0, PR0 + NP, 1, 14, fmt={c: MONEY for c in range(2, 15)})
    mv.freeze_panes = "B6"

    # ---------- sample ----------
    if sample:
        rnd = random.Random(11)
        s["C6"] = 2026
        pets = [("Max", "Dog", "Golden Retriever", dt.date(2021, 4, 12), "Male (neutered)", "985 112 004 567 321", "Riverside Animal Clinic",
                 "(555) 201-3344", "Trupanion – policy TP-448812", "Large-breed kibble, 2 × 1.5 cups", "Sensitive skin – oatmeal shampoo"),
                ("Luna", "Cat", "Domestic shorthair", dt.date(2023, 8, 1), "Female (spayed)", "985 112 008 990 145", "Riverside Animal Clinic",
                 "(555) 201-3344", "", "Wet food 2 × 3 oz + dry 1/4 cup", "Indoor only"),
                ("Kiwi", "Bird", "Budgie", dt.date(2024, 5, 20), "Male", "", "Avian & Exotic Vets", "(555) 330-7781", "",
                 "Seed mix + fresh greens daily", "")]
        for k, row in enumerate(pets):
            for c, v in zip("ABCDEFGHIJK", row):
                if v: pe[f"{c}{P0 + k}"] = v
        exps = []
        for m in range(1, 12):
            exps.append((dt.date(2026, m, 3), "Max", "Food & treats", "Kibble 30 lb bag", "Chewy", 68.99))
            exps.append((dt.date(2026, m, 3), "Luna", "Food & treats", "Wet food 24-pack + dry food", "Chewy", round(rnd.uniform(38, 46), 2)))
            exps.append((dt.date(2026, m, 1), "Max", "Insurance", "Monthly premium", "Trupanion", 54.20))
            exps.append((dt.date(2026, m, 8), "Max", "Medication & preventives", "Flea, tick & heartworm chew", "Riverside Animal Clinic", 31.50))
            exps.append((dt.date(2026, m, 12), "Shared", "Toys & supplies", rnd.choice(["Poop bags & litter", "Cat litter 40 lb", "Treats & chews", "Toys"]), "PetSmart", round(rnd.uniform(14, 36), 2)))
            if m % 2 == 0: exps.append((dt.date(2026, m, 15), "Kiwi", "Food & treats", "Seed mix & millet", "Local pet shop", round(rnd.uniform(9, 14), 2)))
            if m % 3 == 1: exps.append((dt.date(2026, m, 20), "Max", "Grooming", "Bath & nail trim", "Happy Paws Grooming", 65.00))
        exps += [(dt.date(2026, 2, 17), "Max", "Vet visits", "Annual check-up + vaccines (DHPP, rabies)", "Riverside Animal Clinic", 214.00),
                 (dt.date(2026, 3, 9), "Luna", "Vet visits", "Annual check-up + FVRCP", "Riverside Animal Clinic", 148.00),
                 (dt.date(2026, 5, 2), "Max", "License & registration", "City dog license", "City Hall", 25.00),
                 (dt.date(2026, 6, 22), "Max", "Vet visits", "Ear infection", "Riverside Animal Clinic", 186.40),
                 (dt.date(2026, 6, 22), "Max", "Medication & preventives", "Ear drops", "Riverside Animal Clinic", 32.75),
                 (dt.date(2026, 7, 10), "Shared", "Boarding & pet sitting", "Pet sitter – 6 days vacation", "Sarah (sitter)", 270.00),
                 (dt.date(2026, 4, 5), "Max", "Training & classes", "Intermediate obedience class", "Good Dog Academy", 149.00),
                 (dt.date(2026, 8, 30), "Kiwi", "Vet visits", "Wellness check & beak trim", "Avian & Exotic Vets", 95.00),
                 (dt.date(2026, 9, 14), "Luna", "Medication & preventives", "Flea treatment 6-pack", "Chewy", 58.99),
                 (dt.date(2026, 10, 18), "Kiwi", "Toys & supplies", "New cage perches & toys", "Local pet shop", 27.50),
                 (dt.date(2026, 11, 6), "Luna", "Vet visits", "Dental cleaning", "Riverside Animal Clinic", 389.00),
                 (dt.date(2025, 12, 20), "Max", "Toys & supplies", "Christmas toys", "PetSmart", 42.00)]
        exps.sort(key=lambda e: e[0])
        for k, e in enumerate(exps):
            for c, v in zip("ABCDEF", e): ex[f"{c}{E0 + k}"] = v
        care = [("Max", "Vaccine", "Rabies (3-year)", dt.date(2026, 2, 17), 36, "", None),
                ("Max", "Vaccine", "DHPP booster", dt.date(2026, 2, 17), 12, "", None),
                ("Max", "Vaccine", "Bordetella (kennel cough)", dt.date(2025, 11, 2), 12, "Needed for boarding", None),
                ("Max", "Heartworm", "Heartworm chew", dt.date(2026, 11, 8), 1, "Give with breakfast", None),
                ("Max", "Grooming", "Bath & nail trim", dt.date(2026, 10, 20), 3, "Happy Paws", None),
                ("Max", "Check-up", "Annual wellness exam", dt.date(2026, 2, 17), 12, "", None),
                ("Luna", "Vaccine", "FVRCP", dt.date(2026, 3, 9), 12, "", None),
                ("Luna", "Vaccine", "Rabies (1-year)", dt.date(2025, 11, 20), 12, "", None),
                ("Luna", "Flea & tick", "Topical flea treatment", dt.date(2026, 10, 14), 1, "", None),
                ("Luna", "Dental", "Dental cleaning", dt.date(2026, 11, 6), 12, "Two teeth polished", None),
                ("Luna", "Deworming", "Dewormer", dt.date(2026, 8, 1), 3, "", None),
                ("Kiwi", "Check-up", "Avian wellness check", dt.date(2026, 8, 30), 12, "", None),
                ("Kiwi", "Nail trim", "Nail & beak trim", dt.date(2026, 8, 30), 4, "", None),
                ("Max", "License renewal", "City dog license", None, None, "Renew online", dt.date(2027, 4, 30))]
        for k, row in enumerate(care):
            for c, v in zip("ABCDEFG", row):
                if v not in (None, ""): ca[f"{c}{C0 + k}"] = v
        weights = [(dt.date(2026, 2, 17), "Max", 72.4, "Vet"), (dt.date(2026, 3, 9), "Luna", 9.1, "Vet"),
                   (dt.date(2026, 5, 1), "Max", 74.0, "Home scale"), (dt.date(2026, 6, 22), "Max", 73.2, "Vet"),
                   (dt.date(2026, 8, 30), "Kiwi", 0.08, "Vet – 36 g"), (dt.date(2026, 9, 1), "Max", 71.8, "Home scale – on diet"),
                   (dt.date(2026, 11, 6), "Luna", 9.6, "Vet – watch weight"), (dt.date(2026, 11, 10), "Max", 70.9, "Home scale")]
        for k, row in enumerate(weights):
            for c, v in zip("ABCD", row): wl[f"{c}{W0 + k}"] = v

    # ---------- Dashboard ----------
    K.title(dash, "Pet Care Dashboard", "Updates by itself from the Pets, Expenses and Care Schedule tabs.")
    dash["A3"] = '="Year "&Settings!$C$9&"  ·  amounts in "&Settings!$C$5&"  ·  today: "&TEXT(TODAY(),"mmm d, yyyy")'; dash["A3"].font = font(bold=True)
    TOT = f"'Monthly View'!$N${TR}"
    thism = f"INDEX('Monthly View'!$B${TR}:$M${TR},1,{MC})"
    kp = [("Spent this year", f"={TOT}", MONEY), ("Spent in month " + "(latest)", f"={thism}", MONEY),
          ("Average per month", f"={TOT}/MAX(1,{MC})", MONEY),
          ("Budget for the year", f"='Monthly View'!$O${TR}", MONEY), ("Left in budget (+) / over (-)", f"='Monthly View'!$O${TR}-{TOT}", MONEY),
          ("Spent all time (all years)", f"=SUM({EX('F')})", MONEY),
          ("Care items overdue", f'=COUNTIFS({CS("J")},"Overdue*")', "0"),
          ("Due soon (within reminder days)", f'=COUNTIFS({CS("J")},"Due*")', "0"),
          ("Pets", f'=COUNTA(Pets!$A${P0}:$A${P1})', "0")]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 3) * 3, 1 + (i % 3) * 2, lab, f, fmt)
        dash.cell(row=6 + (i // 3) * 3, column=1 + (i % 3) * 2).alignment = Alignment(horizontal="left")
    dash["C5"] = f'="Spent in "&TEXT(DATE({YR},{MC},1),"mmmm")'
    for c, w in zip("ABCDEFGH", [26, 13, 26, 13, 28, 13, 3, 3]): dash.column_dimensions[c].width = w
    dash.conditional_formatting.add("C12", FormulaRule(formula=["C12>0"], font=Font(name=F, size=18, bold=True, color="B91C1C")))
    dash.conditional_formatting.add("E9", FormulaRule(formula=["E9<0"], font=Font(name=F, size=18, bold=True, color="B91C1C")))
    # upcoming care
    H = 15
    K.section(dash, f"A{H - 1}", "Coming up next (sorted by due date)")
    K.hdr(dash, H, ["Due", "Pet", "What", "Type", "Status", "Days left"])
    KEY = CS("K")
    for k in range(10):
        r = H + 1 + k
        dash[f"H{r}"] = f'=IF(COUNT({KEY})<{k + 1},"",SMALL({KEY},{k + 1}))'; dash[f"H{r}"].font = font(color="FFFFFF", size=6)
        m = f"MATCH($H{r},{KEY},0)"
        dash[f"A{r}"] = f'=IF($H{r}="","",INT($H{r}))'
        dash[f"B{r}"] = f'=IF($H{r}="","",INDEX({CS("A")},{m}))'
        dash[f"C{r}"] = f'=IF($H{r}="","",INDEX({CS("C")},{m}))'
        dash[f"D{r}"] = f'=IF($H{r}="","",INDEX({CS("B")},{m}))'
        dash[f"E{r}"] = f'=IF($H{r}="","",INDEX({CS("J")},{m}))'
        dash[f"F{r}"] = f'=IF($H{r}="","",INDEX({CS("I")},{m}))'
    style_range(dash, H + 1, H + 10, 1, 6, fmt={1: "mmm d, yyyy", 6: "0"})
    dash.conditional_formatting.add(f"E{H + 1}:E{H + 10}", FormulaRule(formula=[f'LEFT(E{H + 1},7)="Overdue"'], font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    dash.conditional_formatting.add(f"E{H + 1}:E{H + 10}", FormulaRule(formula=[f'LEFT(E{H + 1},3)="Due"'], font=Font(name=F, bold=True, color="92400E"), fill=fill("FEF3C7")))
    # per pet
    P = H + 13
    K.section(dash, f"A{P - 1}", "Per pet")
    K.hdr(dash, P, ["Pet", "Spent this year", "Species", "Age", "Next care due", "Overdue items"])
    for k in range(NP):
        r = P + 1 + k; pr = P0 + k
        dash[f"A{r}"] = f'=IF(Pets!A{pr}="","",Pets!A{pr})'
        for col, src in zip("BCDEF", "NBLQP"):
            dash[f"{col}{r}"] = f'=IF(A{r}="","",Pets!{src}{pr})'
    style_range(dash, P + 1, P + NP, 1, 6, fmt={2: MONEY, 5: "mmm d, yyyy", 6: "0"})
    dash.conditional_formatting.add(f"B{P + 1}:B{P + NP}", DataBarRule(start_type="num", start_value=0, end_type="max", color="FDBA74"))
    # charts
    ch = BarChart(); ch.type = "col"; ch.title = "Spending per month"; ch.height = 7.5; ch.width = 16; ch.legend = None
    ch.add_data(Reference(mv, min_col=2, max_col=13, min_row=TR, max_row=TR), from_rows=True, titles_from_data=False)
    ch.set_categories(Reference(mv, min_col=2, max_col=13, min_row=5, max_row=5)); ch.y_axis.majorGridlines = None
    ch.series[0].graphicalProperties.solidFill = "FB923C"; ch.gapWidth = 50; ch.y_axis.scaling.min = 0
    dash.add_chart(ch, "I4")
    ch2 = BarChart(); ch2.type = "bar"; ch2.title = "Spent this year by category"; ch2.height = 8.5; ch2.width = 16; ch2.legend = None
    ch2.add_data(Reference(mv, min_col=14, min_row=6, max_row=6 + len(CATS) - 1), titles_from_data=False)
    ch2.set_categories(Reference(mv, min_col=1, min_row=6, max_row=6 + len(CATS) - 1)); ch2.y_axis.majorGridlines = None
    ch2.x_axis.scaling.orientation = "maxMin"; ch2.series[0].graphicalProperties.solidFill = MAIN; ch2.gapWidth = 40
    dash.add_chart(ch2, "I20")
    for c in (ch, ch2): c.roundedCorners = False; c.x_axis.delete = False; c.y_axis.delete = False

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Settings tab: your currency label, weight unit, and a monthly budget for each expense category (edit the category names if you like).", False),
             ("2.  Pets tab: one row per pet – species, breed, birthday, microchip number, vet, insurance, food and notes. Age, latest weight and spending fill in by themselves.", False),
             ("3.  Expenses tab: every purchase or bill – date, pet (or 'Shared'), category, what, shop and amount.", False),
             ("4.  Care Schedule: vaccines, flea & tick, heartworm, deworming, check-ups, grooming, licenses. Type when it was last done and how many months until it repeats.", False),
             ("     The next due date, days left and status (OK / Due in X days / Overdue) are calculated. For a one-off date, type it in 'Or type next due date'.", False),
             ("5.  Weight Log: weigh-ins at home or at the vet – the change since the last weigh-in is shown.", False),
             ("6.  Monthly View: spending per category (with budget) and per pet for every month.", False),
             ("7.  Dashboard: spent this year and this month, budget left, what is coming up next, overdue care and charts.", False),
             ("", False), ("Tips", True),
             ("•  Ask your vet for your pet's vaccine schedule and copy the dates into the Care Schedule – your pet's own needs come first.", False),
             ("•  Keep the microchip number, vet phone and insurance policy here so anyone looking after your pet can find them.", False),
             ("•  Change 'Remind me this many days before due' in Settings to see items earlier or later.", False),
             ("•  To see an earlier year, type it in Settings > 'Year to show'.", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  20 pets, 1,500 expenses, 150 care items, 300 weigh-ins, 15 categories. No macros. Works in Excel 2010+ and Google Sheets.", False),
             ("•  To use Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("Three made-up pets (a dog, a cat and a budgie) with a year of expenses. Use the blank copy in your download for your own pets.", False)]
    K.start(st, "Pet Care & Expense Tracker", "Pet profiles, expenses vs budget, vaccine & vet reminders, weight log – Excel and Google Sheets.", lines)
    _orig_save = wb.save

    def _save(p):  # keep the Start Here tab on one printed page
        st.page_setup.fitToHeight = 1
        _orig_save(p)
    wb.save = _save
    finish(wb, sample, path)


META = {
    "name": "Pet Care & Expense Tracker (Vet, Vaccine & Flea Reminders, Pet Budget, Weight Log) – Excel & Google Sheets",
    "file_base": "Pet-Care-Expense-Tracker",
    "price": 12,
    "color": "#C2410C",
    "title": "Pet Care & Expense Tracker",
    "tagline": "pet budget · vaccine & vet reminders · weight log · pet profiles",
    "intro": "Know exactly what your dogs, cats and other pets cost – and never miss a vaccine, flea treatment or check-up again. Keep each pet's details in one place, log every expense against a monthly budget, and let the tracker work out when the next vaccine, heartworm dose or grooming is due.",
    "features": [
        "Pet profiles for up to 20 pets: species, breed, birthday (age calculated), microchip, vet, insurance, food and allergies",
        "Expense log (1,500 rows) with drop-down pets and categories, including shared costs",
        "Monthly budget per category with year total and amount left or over",
        "Care schedule: vaccines, flea & tick, heartworm, deworming, check-ups, dental, grooming, licenses – next due date, days left and Overdue / Due soon status",
        "Dashboard: spent this year and this month, average per month, budget left, next 10 care items sorted by date, per-pet spending",
        "Monthly view per category and per pet, plus charts",
        "Weight log with change since the previous weigh-in",
        "Blank copy plus a sample copy with a dog, a cat and a budgie",
    ],
    "tags": ["pet tracker", "pet expense tracker", "dog vaccine record", "pet health record", "pet budget", "cat care",
             "google sheets", "excel template", "dog care planner", "vet records", "pet planner", "pet spreadsheet"],
    "summary": "Instant download: Excel + Google Sheets template, blank copy + sample copy.",
}
