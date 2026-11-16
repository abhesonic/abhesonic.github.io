"""Home Maintenance Schedule & Renovation Budget Planner – recurring home tasks with reminders, maintenance log,
renovation projects with line-item budget (estimate vs actual vs paid), appliance & warranty inventory.

Settings -> Maintenance Schedule -> Maintenance Log -> Projects -> Project Costs -> Home Inventory -> Monthly View -> Dashboard.
"""
import datetime as dt, random
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, DataBarRule
from openpyxl.chart import BarChart, Reference
from _common import Kit, F, B, MONEY, DATE, font, fill, style_range, input_cell, finish

MAIN = "0F766E"; SOFT = "CCFBF1"
K = Kit(MAIN, SOFT)
NM = 150; M0 = 6; M1 = M0 + NM - 1      # schedule
NL = 600; L0 = 6; L1 = L0 + NL - 1      # log
NP = 30; P0 = 6; P1 = P0 + NP - 1       # projects
NC = 800; C0 = 6; C1 = C0 + NC - 1      # project costs
NI = 100; I0 = 6; I1 = I0 + NI - 1      # inventory
NA = 20; A0 = 6; A1 = A0 + NA - 1       # areas (Settings F)
NK = 18; K0 = 6; K1 = K0 + NK - 1       # cost categories (Settings H)
AREAS = ["Kitchen", "Bathroom", "Bedroom", "Living room", "Basement", "Garage", "Laundry", "Exterior", "Roof", "Yard & garden",
         "HVAC", "Plumbing", "Electrical", "Whole house"]
CATS = ["Materials", "Labor / contractor", "Fixtures", "Appliances", "Flooring", "Paint & finishes", "Cabinets & countertops",
        "Lighting & electrical", "Plumbing", "Permits & fees", "Tools & rental", "Furniture & decor", "Disposal & cleanup",
        "Design & planning", "Other"]
TASKS = [("Replace HVAC air filter", "HVAC", 3, "Every season", "DIY", 15),
         ("Test smoke & CO detectors", "Whole house", 6, "Spring & fall", "DIY", 0),
         ("Replace smoke detector batteries", "Whole house", 12, "Fall", "DIY", 20),
         ("Furnace / heating service", "HVAC", 12, "Early fall", "Pro", 150),
         ("AC service & clean condenser", "HVAC", 12, "Spring", "Pro", 150),
         ("Clean gutters & downspouts", "Exterior", 6, "Spring & fall", "Pro", 180),
         ("Inspect roof & flashing", "Roof", 12, "Spring", "DIY", 0),
         ("Flush water heater", "Plumbing", 12, "Any", "DIY", 0),
         ("Test water heater pressure valve", "Plumbing", 12, "Any", "DIY", 0),
         ("Clean dryer vent duct", "Laundry", 12, "Any", "Pro", 120),
         ("Clean refrigerator coils", "Kitchen", 12, "Any", "DIY", 0),
         ("Clean range hood filter", "Kitchen", 3, "Any", "DIY", 0),
         ("Clean dishwasher filter", "Kitchen", 1, "Any", "DIY", 0),
         ("Replace fridge water filter", "Kitchen", 6, "Any", "DIY", 45),
         ("Check caulk & grout in bathrooms", "Bathroom", 12, "Any", "DIY", 15),
         ("Clean shower heads & faucet aerators", "Bathroom", 6, "Any", "DIY", 0),
         ("Test GFCI outlets", "Electrical", 6, "Any", "DIY", 0),
         ("Check fire extinguisher", "Whole house", 12, "Any", "DIY", 0),
         ("Pest control treatment", "Whole house", 3, "Any", "Pro", 110),
         ("Shut off & drain outdoor faucets", "Exterior", 12, "Late fall", "DIY", 0),
         ("Service lawn mower", "Yard & garden", 12, "Spring", "DIY", 40),
         ("Seal deck / fence", "Exterior", 24, "Summer", "DIY", 120),
         ("Power-wash siding & walkways", "Exterior", 12, "Spring", "DIY", 0),
         ("Chimney sweep & inspection", "Living room", 12, "Late summer", "Pro", 250),
         ("Test sump pump", "Basement", 6, "Spring & fall", "DIY", 0),
         ("Check weatherstripping & window tracks", "Whole house", 12, "Fall", "DIY", 20),
         ("Septic tank pump-out (if you have one)", "Exterior", 36, "Any", "Pro", 450)]


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); ms = wb.create_sheet("Maintenance Schedule"); ml = wb.create_sheet("Maintenance Log")
    pj = wb.create_sheet("Projects"); pc = wb.create_sheet("Project Costs"); hi = wb.create_sheet("Home Inventory")
    mv = wb.create_sheet("Monthly View"); s = wb.create_sheet("Settings")

    # ---------- Settings ----------
    K.title(s, "Settings", "Yellow cells are yours to edit.")
    for r, lab, v in ((5, "Currency (label only)", "USD"), (6, "Year to show (blank = this year)", None),
                      (10, "Remind me this many days before due", 30), (11, "Safety margin for renovations (contingency %)", 0.10)):
        s[f"A{r}"] = lab; s[f"A{r}"].font = font(bold=True); s[f"A{r}"].border = B; input_cell(s[f"C{r}"], v)
    s["C11"].number_format = "0%"
    s["A8"] = "Year shown (automatic)"; s["C8"] = "=IF(ISNUMBER(C6),C6,YEAR(TODAY()))"; s["C8"].number_format = "0"
    s["A9"] = "Months counted for averages"; s["C9"] = "=IF(C8<YEAR(TODAY()),12,IF(C8>YEAR(TODAY()),1,MONTH(TODAY())))"
    for r in (8, 9):
        s[f"A{r}"].font = font(italic=True); s[f"C{r}"].fill = fill("F3F4F6"); s[f"C{r}"].border = B; s[f"C{r}"].font = font()
    K.hdr(s, 5, ["Areas / rooms"], col0=6); K.hdr(s, 5, ["Renovation cost categories"], col0=8)
    for i in range(NA): input_cell(s[f"F{A0 + i}"], AREAS[i] if i < len(AREAS) else None)
    for i in range(NK): input_cell(s[f"H{K0 + i}"], CATS[i] if i < len(CATS) else None)
    for c, w in zip("ABCDEFGH", [44, 3, 12, 3, 3, 22, 3, 26]): s.column_dimensions[c].width = w
    YR = "Settings!$C$8"; MC = "Settings!$C$9"; LEAD = "Settings!$C$10"; CONT = "Settings!$C$11"
    AREA = f"=Settings!$F${A0}:$F${A1}"

    MSr = lambda c: f"'Maintenance Schedule'!${c}${M0}:${c}${M1}"
    MLr = lambda c: f"'Maintenance Log'!${c}${L0}:${c}${L1}"
    PJr = lambda c: f"Projects!${c}${P0}:${c}${P1}"
    PCr = lambda c: f"'Project Costs'!${c}${C0}:${c}${C1}"
    HIr = lambda c: f"'Home Inventory'!${c}${I0}:${c}${I1}"

    # ---------- Maintenance Schedule ----------
    K.title(ms, "Maintenance Schedule", "Recurring jobs that keep your home in shape. A starter list is filled in – edit, delete or add your own. "
                                        "Leave 'Last done' blank and it is read from the Maintenance Log.")
    K.hdr(ms, 5, ["Task", "Area", "Every (months)", "Best time", "DIY / Pro", "Typical cost", "Last done (blank = from log)",
                  "Or type next due date", "Notes", "Last done (used)", "Next due", "Days left", "Status", "(sort key)"],
          [36, 14, 9, 14, 8, 11, 14, 13, 24, 12, 12, 9, 15, 9])
    style_range(ms, M0, M1, 1, 14, inputs=range(1, 10), fmt={3: "0", 6: MONEY, 7: DATE, 8: DATE, 10: DATE, 11: DATE, 12: '0;-0;0'})
    for r in range(M0, M1 + 1):
        mx = f'SUMPRODUCT(MAX(({MLr("B")}=A{r})*{MLr("A")}))'
        ms[f"J{r}"] = f'=IF(A{r}="","",IF(ISNUMBER(G{r}),MAX(G{r},{mx}),IF({mx}>0,{mx},"")))'
        ms[f"K{r}"] = (f'=IF(A{r}="","",IF(AND(ISNUMBER(H{r}),OR(NOT(ISNUMBER(J{r})),N(H{r})>N(J{r}))),H{r},'
                       f'IF(AND(ISNUMBER(J{r}),N(C{r})>0),EDATE(J{r},C{r}),"")))')
        ms[f"L{r}"] = f'=IF(ISNUMBER(K{r}),K{r}-TODAY(),"")'
        ms[f"M{r}"] = (f'=IF(A{r}="","",IF(NOT(ISNUMBER(K{r})),IF(ISNUMBER(J{r}),"Done","Set last done"),'
                       f'IF(L{r}<0,"Overdue",IF(L{r}<={LEAD},"Due soon","OK"))))')
        ms[f"N{r}"] = f'=IF(ISNUMBER(K{r}),K{r}+ROW()/100000,"")'
        ms[f"M{r}"].font = font(bold=True); ms[f"N{r}"].font = font(color="9CA3AF", size=8)
    for i, (t, a, m, when, who, cost) in enumerate(TASKS):
        r = M0 + i
        for c, v in zip("ABCDEF", (t, a, m, when, who, cost)): ms[f"{c}{r}"] = v
    dv = DataValidation(type="list", formula1=AREA, allow_blank=True); ms.add_data_validation(dv); dv.add(f"B{M0}:B{M1}")
    dv = DataValidation(type="list", formula1='"DIY,Pro"', allow_blank=True); ms.add_data_validation(dv); dv.add(f"E{M0}:E{M1}")
    ms.conditional_formatting.add(f"M{M0}:M{M1}", FormulaRule(formula=[f'M{M0}="Overdue"'], font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    ms.conditional_formatting.add(f"M{M0}:M{M1}", FormulaRule(formula=[f'M{M0}="Due soon"'], font=Font(name=F, bold=True, color="92400E"), fill=fill("FEF3C7")))
    ms.conditional_formatting.add(f"M{M0}:M{M1}", FormulaRule(formula=[f'M{M0}="OK"'], font=Font(name=F, bold=True, color="166534")))
    ms.freeze_panes = f"B{M0}"; ms.auto_filter.ref = f"A5:M{M1}"

    # ---------- Maintenance Log ----------
    K.title(ml, "Maintenance Log", "Every job done – scheduled tasks and one-off repairs. Pick the task from the list (or type a new one) so reminders reset by themselves.")
    K.hdr(ml, 5, ["Date", "Task", "Area", "Done by (DIY or company)", "Cost", "Notes", "Year", "Month"], [12, 36, 14, 24, 11, 30, 7, 7])
    style_range(ml, L0, L1, 1, 8, inputs=range(1, 7), fmt={1: DATE, 5: MONEY, 7: "0", 8: "0"})
    for r in range(L0, L1 + 1):
        ml[f"G{r}"] = f'=IF(ISNUMBER(A{r}),YEAR(A{r}),"")'; ml[f"H{r}"] = f'=IF(ISNUMBER(A{r}),MONTH(A{r}),"")'
    dv = DataValidation(type="list", formula1=f"='Maintenance Schedule'!$A${M0}:$A${M1}", allow_blank=True, showErrorMessage=False)
    ml.add_data_validation(dv); dv.add(f"B{L0}:B{L1}")
    dv = DataValidation(type="list", formula1=AREA, allow_blank=True); ml.add_data_validation(dv); dv.add(f"C{L0}:C{L1}")
    ml.freeze_panes = f"A{L0}"; ml.auto_filter.ref = f"A5:H{L1}"

    # ---------- Projects ----------
    K.title(pj, "Projects", "One row per renovation or improvement project. Costs come from the Project Costs tab.")
    K.hdr(pj, 5, ["Project", "Area", "Priority", "Status", "Start", "Target finish", "Budget", "Notes",
                  "Estimated (quotes)", "Actual cost so far", "Paid so far", "Still owed", "Expected total", "Left (+) / over (-)",
                  "% of budget", "Safety margin", "Cost lines"],
          [28, 14, 9, 12, 11, 11, 12, 26, 12, 12, 12, 11, 12, 12, 9, 11, 7])
    style_range(pj, P0, P1, 1, 17, inputs=range(1, 9),
                fmt={5: DATE, 6: DATE, 7: MONEY, 9: MONEY, 10: MONEY, 11: MONEY, 12: MONEY, 13: MONEY, 14: MONEY, 15: "0%", 16: MONEY, 17: "0"})
    for r in range(P0, P1 + 1):
        e = f'=IF(A{r}="","",SUMIFS({{}},{PCr("A")},A{r}))'
        pj[f"I{r}"] = e.format(PCr("K")); pj[f"J{r}"] = e.format(PCr("G")); pj[f"K{r}"] = e.format(PCr("H"))
        pj[f"L{r}"] = e.format(PCr("M")); pj[f"M{r}"] = e.format(PCr("L"))
        pj[f"N{r}"] = f'=IF(A{r}="","",N(G{r})-M{r})'
        pj[f"O{r}"] = f'=IF(OR(A{r}="",N(G{r})=0),"",M{r}/G{r})'
        pj[f"P{r}"] = f'=IF(A{r}="","",N(G{r})*{CONT})'
        pj[f"Q{r}"] = f'=IF(A{r}="","",COUNTIFS({PCr("A")},A{r}))'
    dv = DataValidation(type="list", formula1=AREA, allow_blank=True); pj.add_data_validation(dv); dv.add(f"B{P0}:B{P1}")
    dv = DataValidation(type="list", formula1='"High,Medium,Low"', allow_blank=True); pj.add_data_validation(dv); dv.add(f"C{P0}:C{P1}")
    dv = DataValidation(type="list", formula1='"Idea,Planning,Getting quotes,In progress,On hold,Done"', allow_blank=True)
    pj.add_data_validation(dv); dv.add(f"D{P0}:D{P1}")
    pj.conditional_formatting.add(f"N{P0}:N{P1}", FormulaRule(formula=[f'AND(ISNUMBER(N{P0}),N{P0}<0)'], font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    pj.conditional_formatting.add(f"N{P0}:N{P1}", FormulaRule(formula=[f'AND(ISNUMBER(N{P0}),N{P0}<P{P0})'], font=Font(name=F, bold=True, color="92400E"), fill=fill("FEF3C7")))
    pj.conditional_formatting.add(f"O{P0}:O{P1}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1.2, color="5EEAD4"))
    pj.freeze_panes = f"B{P0}"

    # ---------- Project Costs ----------
    K.title(pc, "Project Costs", "Every item, quote and invoice. Type the estimate first; add the actual cost when you know it. 'Expected' uses the actual if there is one.")
    K.hdr(pc, 5, ["Project", "Category", "Item / description", "Supplier / contractor", "Qty (blank = 1)", "Estimated unit cost",
                  "Actual cost (total)", "Paid so far", "Date paid", "Status", "Estimate (total)", "Expected", "Still owed", "Year paid", "Month paid"],
          [24, 20, 30, 22, 8, 12, 12, 12, 12, 12, 12, 12, 11, 7, 7])
    style_range(pc, C0, C1, 1, 15, inputs=range(1, 11),
                fmt={5: "0.##", 6: MONEY, 7: MONEY, 8: MONEY, 9: DATE, 11: MONEY, 12: MONEY, 13: MONEY, 14: "0", 15: "0"})
    for r in range(C0, C1 + 1):
        pc[f"K{r}"] = f'=IF(AND(C{r}="",N(F{r})=0),"",IF(ISNUMBER(E{r}),E{r},1)*N(F{r}))'
        pc[f"L{r}"] = f'=IF(A{r}&C{r}="","",IF(ISNUMBER(G{r}),G{r},N(K{r})))'
        pc[f"M{r}"] = f'=IF(ISNUMBER(G{r}),MAX(0,G{r}-N(H{r})),"")'
        pc[f"N{r}"] = f'=IF(ISNUMBER(I{r}),YEAR(I{r}),"")'; pc[f"O{r}"] = f'=IF(ISNUMBER(I{r}),MONTH(I{r}),"")'
    dv = DataValidation(type="list", formula1=f"=Projects!$A${P0}:$A${P1}", allow_blank=True); pc.add_data_validation(dv); dv.add(f"A{C0}:A{C1}")
    dv = DataValidation(type="list", formula1=f"=Settings!$H${K0}:$H${K1}", allow_blank=True); pc.add_data_validation(dv); dv.add(f"B{C0}:B{C1}")
    dv = DataValidation(type="list", formula1='"Idea,Quote,To buy,Ordered,Bought,Installed / done"', allow_blank=True)
    pc.add_data_validation(dv); dv.add(f"J{C0}:J{C1}")
    pc.conditional_formatting.add(f"M{C0}:M{C1}", FormulaRule(formula=[f'N(M{C0})>0'], font=Font(name=F, bold=True, color="B45309")))
    pc.freeze_panes = f"A{C0}"; pc.auto_filter.ref = f"A5:O{C1}"

    # ---------- Home Inventory ----------
    K.title(hi, "Home Inventory", "Appliances and systems: model and serial numbers, purchase date, warranty – handy for repairs and insurance claims.")
    K.hdr(hi, 5, ["Item", "Area", "Brand", "Model number", "Serial number", "Purchase date", "Price", "Warranty ends",
                  "Bought from / manual link", "Notes", "Age (years)", "Warranty"],
          [24, 14, 14, 16, 18, 12, 11, 12, 28, 24, 9, 18])
    style_range(hi, I0, I1, 1, 12, inputs=range(1, 11), fmt={6: DATE, 7: MONEY, 8: DATE, 11: "0.0"})
    for r in range(I0, I1 + 1):
        hi[f"K{r}"] = f'=IF(OR(A{r}="",NOT(ISNUMBER(F{r}))),"",(TODAY()-F{r})/365.25)'
        hi[f"L{r}"] = (f'=IF(OR(A{r}="",NOT(ISNUMBER(H{r}))),"",IF(H{r}<TODAY(),"Expired",'
                       f'IF(H{r}-TODAY()<={LEAD}*3,"Ends in "&(H{r}-TODAY())&" days","Active")))')
    dv = DataValidation(type="list", formula1=AREA, allow_blank=True); hi.add_data_validation(dv); dv.add(f"B{I0}:B{I1}")
    hi.conditional_formatting.add(f"L{I0}:L{I1}", FormulaRule(formula=[f'LEFT(L{I0},4)="Ends"'], font=Font(name=F, bold=True, color="92400E"), fill=fill("FEF3C7")))
    hi.conditional_formatting.add(f"L{I0}:L{I1}", FormulaRule(formula=[f'L{I0}="Active"'], font=Font(name=F, bold=True, color="166534")))
    hi.conditional_formatting.add(f"L{I0}:L{I1}", FormulaRule(formula=[f'L{I0}="Expired"'], font=Font(name=F, color="6B7280")))
    hi.freeze_panes = f"B{I0}"

    # ---------- Monthly View ----------
    K.title(mv, "Monthly View", "Money spent per month: maintenance (by date done) and renovation payments (by date paid).")
    mv["A3"] = '="Year "&Settings!$C$8&"  ·  amounts in "&Settings!$C$5'; mv["A3"].font = font(bold=True)
    MN = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    K.hdr(mv, 5, ["Spending"] + MN + ["Year total"], [30] + [10] * 12 + [12])
    mv["A6"] = "Maintenance & repairs"; mv["A7"] = "Renovation payments"; mv["A8"] = "Total"
    for j in range(12):
        c = CL(2 + j)
        mv[f"{c}6"] = f'=SUMIFS({MLr("E")},{MLr("G")},{YR},{MLr("H")},{j + 1})'
        mv[f"{c}7"] = f'=SUMIFS({PCr("H")},{PCr("N")},{YR},{PCr("O")},{j + 1})'
        mv[f"{c}8"] = f"={c}6+{c}7"
    for r in (6, 7, 8): mv[f"N{r}"] = f"=SUM(B{r}:M{r})"
    style_range(mv, 6, 8, 1, 14, fmt={c: MONEY for c in range(2, 15)})
    for c in range(1, 15):
        cell = mv.cell(row=8, column=c); cell.font = Font(name=F, bold=True, color=MAIN); cell.fill = fill(SOFT)
    AR0 = 12
    K.hdr(mv, AR0 - 1, ["Maintenance by area"] + MN + ["Year total"])
    for k in range(NA):
        r = AR0 + k
        mv[f"A{r}"] = f'=IF(Settings!$F${A0 + k}="","",Settings!$F${A0 + k})'
        for j in range(12):
            mv[f"{CL(2 + j)}{r}"] = f'=IF($A{r}="",0,SUMIFS({MLr("E")},{MLr("C")},$A{r},{MLr("G")},{YR},{MLr("H")},{j + 1}))'
        mv[f"N{r}"] = f"=SUM(B{r}:M{r})"
    style_range(mv, AR0, AR0 + NA - 1, 1, 14, fmt={c: MONEY for c in range(2, 15)})
    PR0 = AR0 + NA + 3
    K.hdr(mv, PR0 - 1, ["Renovation paid by project"] + MN + ["Year total"])
    for k in range(NP):
        r = PR0 + k
        mv[f"A{r}"] = f'=IF(Projects!$A${P0 + k}="","",Projects!$A${P0 + k})'
        for j in range(12):
            mv[f"{CL(2 + j)}{r}"] = f'=IF($A{r}="",0,SUMIFS({PCr("H")},{PCr("A")},$A{r},{PCr("N")},{YR},{PCr("O")},{j + 1}))'
        mv[f"N{r}"] = f"=SUM(B{r}:M{r})"
    style_range(mv, PR0, PR0 + NP - 1, 1, 14, fmt={c: MONEY for c in range(2, 15)})
    mv.freeze_panes = "B6"

    # ---------- sample ----------
    if sample:
        rnd = random.Random(5)
        s["C6"] = 2026
        logs = [(dt.date(2026, 1, 8), "Replace HVAC air filter", "HVAC", "DIY", 14.99, "MERV 11, 16x25x1"),
                (dt.date(2026, 1, 15), "Clean dishwasher filter", "Kitchen", "DIY", 0, ""),
                (dt.date(2026, 2, 21), "Pest control treatment", "Whole house", "BugFree Pest Co.", 110, "Quarterly plan"),
                (dt.date(2026, 3, 28), "Test smoke & CO detectors", "Whole house", "DIY", 0, "All 6 OK"),
                (dt.date(2026, 4, 9), "Replace HVAC air filter", "HVAC", "DIY", 14.99, ""),
                (dt.date(2026, 4, 18), "Clean gutters & downspouts", "Exterior", "Clear Gutter Pros", 180, ""),
                (dt.date(2026, 4, 25), "AC service & clean condenser", "HVAC", "Comfort Air HVAC", 149, "Refrigerant OK"),
                (dt.date(2026, 5, 2), "Service lawn mower", "Yard & garden", "DIY", 38.50, "Oil, blade, spark plug"),
                (dt.date(2026, 5, 2), "Power-wash siding & walkways", "Exterior", "DIY", 0, "Rented washer: see Tools"),
                (dt.date(2026, 5, 22), "Pest control treatment", "Whole house", "BugFree Pest Co.", 110, ""),
                (dt.date(2026, 6, 14), "Leaking kitchen faucet cartridge", "Plumbing", "DIY", 27.80, "One-off repair"),
                (dt.date(2026, 7, 10), "Replace HVAC air filter", "HVAC", "DIY", 14.99, ""),
                (dt.date(2026, 7, 19), "Replace fridge water filter", "Kitchen", "DIY", 44.99, ""),
                (dt.date(2026, 8, 22), "Pest control treatment", "Whole house", "BugFree Pest Co.", 110, ""),
                (dt.date(2026, 9, 6), "Clean dryer vent duct", "Laundry", "DuctMasters", 119, ""),
                (dt.date(2026, 9, 27), "Furnace / heating service", "HVAC", "Comfort Air HVAC", 149, ""),
                (dt.date(2026, 10, 4), "Test smoke & CO detectors", "Whole house", "DIY", 0, ""),
                (dt.date(2026, 10, 4), "Replace smoke detector batteries", "Whole house", "DIY", 19.99, "9V x6"),
                (dt.date(2026, 10, 10), "Clean range hood filter", "Kitchen", "DIY", 0, ""),
                (dt.date(2026, 10, 25), "Clean gutters & downspouts", "Exterior", "Clear Gutter Pros", 180, ""),
                (dt.date(2026, 11, 1), "Clean dishwasher filter", "Kitchen", "DIY", 0, ""),
                (dt.date(2026, 11, 8), "Shut off & drain outdoor faucets", "Exterior", "DIY", 0, "")]
        for k, row in enumerate(logs):
            for c, v in zip("ABCDEF", row):
                if v != "": ml[f"{c}{L0 + k}"] = v
        ms[f"G{M0 + 7}"] = dt.date(2025, 9, 14)          # water heater flush
        ms[f"G{M0 + 10}"] = dt.date(2025, 12, 1)         # fridge coils
        ms[f"G{M0 + 23}"] = dt.date(2025, 9, 2)          # chimney
        ms[f"G{M0 + 24}"] = dt.date(2026, 5, 10)         # sump pump
        ms[f"G{M0 + 6}"] = dt.date(2025, 4, 20)          # roof
        projects = [("Kitchen refresh", "Kitchen", "High", "In progress", dt.date(2026, 9, 1), dt.date(2026, 12, 15), 18000,
                     "New counters, backsplash, paint, lighting"),
                    ("Primary bathroom remodel", "Bathroom", "Medium", "Getting quotes", dt.date(2027, 3, 1), dt.date(2027, 5, 1), 12000,
                     "Walk-in shower, new vanity"),
                    ("Backyard deck & lights", "Yard & garden", "Low", "Done", dt.date(2026, 5, 15), dt.date(2026, 6, 15), 2500,
                     "Stain deck, string lights")]
        for k, row in enumerate(projects):
            for c, v in zip("ABCDEFGH", row): pj[f"{c}{P0 + k}"] = v
        KR, BR, DK = projects[0][0], projects[1][0], projects[2][0]
        lines = [(KR, "Cabinets & countertops", "Quartz countertops (42 sq ft) incl. install", "StoneWorks", 42, 85, 3780, 1890, dt.date(2026, 9, 12), "Ordered"),
                 (KR, "Cabinets & countertops", "Cabinet refacing & soft-close hinges", "CabinetCare", None, 5200, 5650, 5650, dt.date(2026, 10, 2), "Installed / done"),
                 (KR, "Materials", "Subway tile backsplash 30 sq ft", "Home Depot", 30, 9.5, 312.40, 312.40, dt.date(2026, 10, 9), "Bought"),
                 (KR, "Labor / contractor", "Backsplash install", "Mike's Tile", None, 900, None, 300, dt.date(2026, 10, 20), "To buy"),
                 (KR, "Plumbing", "Undermount sink + faucet", "Ferguson", None, 780, 742.15, 742.15, dt.date(2026, 10, 9), "Bought"),
                 (KR, "Labor / contractor", "Plumber: sink & dishwasher hookup", "Rapid Plumbing", None, 450, None, None, None, "Quote"),
                 (KR, "Lighting & electrical", "Under-cabinet LED + 4 pendants", "Lowe's", None, 650, 598.30, 598.30, dt.date(2026, 10, 16), "Bought"),
                 (KR, "Labor / contractor", "Electrician", "BrightSpark Electric", None, 800, None, None, None, "Quote"),
                 (KR, "Paint & finishes", "Paint & supplies (walls + ceiling)", "Sherwin-Williams", None, 260, 284.75, 284.75, dt.date(2026, 9, 20), "Bought"),
                 (KR, "Appliances", "Range hood", "Best Buy", None, 450, None, None, None, "To buy"),
                 (KR, "Permits & fees", "Electrical permit", "City", None, 120, 120, 120, dt.date(2026, 9, 8), "Installed / done"),
                 (KR, "Disposal & cleanup", "Dumpster bag", "Waste Management", None, 180, 179, 179, dt.date(2026, 9, 5), "Installed / done"),
                 (BR, "Labor / contractor", "Contractor quote A – full remodel labor", "Hudson Renovations", None, 6800, None, None, None, "Quote"),
                 (BR, "Fixtures", "Shower glass door", "Build.com", None, 1100, None, None, None, "Idea"),
                 (BR, "Fixtures", "Vanity 48\" with top", "Wayfair", None, 950, None, None, None, "Idea"),
                 (BR, "Flooring", "Porcelain tile floor + shower walls", "Floor & Decor", 160, 6.5, None, None, None, "Idea"),
                 (BR, "Plumbing", "Thermostatic shower valve & head", "Ferguson", None, 620, None, None, None, "Idea"),
                 (BR, "Design & planning", "Design consult", "Studio Nine", None, 300, 300, 300, dt.date(2026, 10, 28), "Installed / done"),
                 (DK, "Materials", "Deck stain (5 gal) + brushes", "Home Depot", None, 280, 264.90, 264.90, dt.date(2026, 5, 16), "Bought"),
                 (DK, "Tools & rental", "Pressure washer rental (1 day)", "Home Depot", None, 90, 89, 89, dt.date(2026, 5, 2), "Installed / done"),
                 (DK, "Lighting & electrical", "Outdoor string lights + posts", "Amazon", None, 320, 356.40, 356.40, dt.date(2026, 5, 30), "Installed / done"),
                 (DK, "Furniture & decor", "Outdoor dining set", "Costco", None, 1200, 1149.99, 1149.99, dt.date(2026, 6, 6), "Installed / done"),
                 (DK, "Labor / contractor", "Helper for staining (2 days)", "Local handyman", None, 500, 560, 560, dt.date(2026, 6, 12), "Installed / done")]
        for k, row in enumerate(lines):
            for c, v in zip("ABCDEFGHIJ", row):
                if v is not None: pc[f"{c}{C0 + k}"] = v
        inv = [("Refrigerator", "Kitchen", "Samsung", "RF28T5001SR", "0B4K4BAR500123", dt.date(2022, 8, 14), 1899, dt.date(2027, 8, 14), "Best Buy – 5-yr plan", ""),
               ("Dishwasher", "Kitchen", "Bosch", "SHE53C85N", "FD 0012 0345", dt.date(2024, 2, 3), 849, dt.date(2026, 2, 3), "Lowe's", ""),
               ("Range", "Kitchen", "GE", "JGB735SPSS", "ZH512345K", dt.date(2019, 6, 1), 1099, dt.date(2020, 6, 1), "Home Depot", ""),
               ("Washer", "Laundry", "LG", "WM4000HWA", "103KWYT9X123", dt.date(2023, 11, 24), 899, dt.date(2026, 12, 24), "Costco", "Black Friday"),
               ("Dryer", "Laundry", "LG", "DLE7400WE", "103KWJK5Y456", dt.date(2023, 11, 24), 849, dt.date(2026, 12, 24), "Costco", ""),
               ("Furnace", "HVAC", "Carrier", "59SC5A080E17", "2918A12345", dt.date(2018, 10, 10), 4200, dt.date(2028, 10, 10), "Comfort Air HVAC", "10-yr parts"),
               ("Water heater (50 gal)", "Plumbing", "Rheem", "XG50T06EC36U1", "Q251812345", dt.date(2020, 3, 5), 1350, dt.date(2026, 3, 5), "Rapid Plumbing", ""),
               ("Garage door opener", "Garage", "Chamberlain", "B4545T", "", dt.date(2021, 7, 18), 329, dt.date(2026, 7, 18), "Home Depot", "")]
        for k, row in enumerate(inv):
            for c, v in zip("ABCDEFGHIJ", row):
                if v != "": hi[f"{c}{I0 + k}"] = v

    # ---------- Dashboard ----------
    K.title(dash, "Home Dashboard", "Updates by itself from the other tabs.")
    dash["A3"] = '="Year "&Settings!$C$8&"  ·  amounts in "&Settings!$C$5&"  ·  today: "&TEXT(TODAY(),"mmm d, yyyy")'; dash["A3"].font = font(bold=True)
    kp = [("Renovation budget (all projects)", f"=SUM({PJr('G')})", MONEY), ("Expected total cost", f"=SUM({PJr('M')})", MONEY),
          ("Left (+) / over (-)", f"=SUM({PJr('G')})-SUM({PJr('M')})", MONEY),
          ("Renovation paid so far", f"=SUM({PJr('K')})", MONEY), ("Still owed to suppliers", f"=SUM({PJr('L')})", MONEY),
          ("Maintenance this year", "='Monthly View'!$N$6", MONEY),
          ("Tasks overdue", f'=COUNTIFS({MSr("M")},"Overdue")', "0"), ("Tasks due soon", f'=COUNTIFS({MSr("M")},"Due soon")', "0"),
          ("Warranties ending soon", f'=COUNTIFS({HIr("L")},"Ends*")', "0")]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 3) * 3, 1 + (i % 3) * 2, lab, f, fmt)
        dash.cell(row=6 + (i // 3) * 3, column=1 + (i % 3) * 2).alignment = Alignment(horizontal="left")
    for c, w in zip("ABCDEFGH", [30, 22, 26, 14, 26, 14, 3, 3]): dash.column_dimensions[c].width = w
    dash.conditional_formatting.add("E6", FormulaRule(formula=["E6<0"], font=Font(name=F, size=18, bold=True, color="B91C1C")))
    dash.conditional_formatting.add("A12", FormulaRule(formula=["A12>0"], font=Font(name=F, size=18, bold=True, color="B91C1C")))
    dash.conditional_formatting.add("C12", FormulaRule(formula=["C12>0"], font=Font(name=F, size=18, bold=True, color="B45309")))
    H = 15
    K.section(dash, f"A{H - 1}", "Maintenance coming up (sorted by due date)")
    K.hdr(dash, H, ["Task", "Area", "Due", "Days left", "Status", "Typical cost"])
    KEY = MSr("N")
    for k in range(10):
        r = H + 1 + k
        dash[f"H{r}"] = f'=IF(COUNT({KEY})<{k + 1},"",SMALL({KEY},{k + 1}))'; dash[f"H{r}"].font = font(color="FFFFFF", size=6)
        m = f"MATCH($H{r},{KEY},0)"
        for col, src in zip("ABCDEF", "ABKLMF"):
            dash[f"{col}{r}"] = f'=IF($H{r}="","",INDEX({MSr(src)},{m}))'
    style_range(dash, H + 1, H + 10, 1, 6, fmt={3: "mmm d, yyyy", 4: "0", 6: MONEY})
    dash.conditional_formatting.add(f"E{H + 1}:E{H + 10}", FormulaRule(formula=[f'E{H + 1}="Overdue"'], font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    dash.conditional_formatting.add(f"E{H + 1}:E{H + 10}", FormulaRule(formula=[f'E{H + 1}="Due soon"'], font=Font(name=F, bold=True, color="92400E"), fill=fill("FEF3C7")))
    dash.conditional_formatting.add(f"E{H + 1}:E{H + 10}", FormulaRule(formula=[f'E{H + 1}="OK"'], font=Font(name=F, bold=True, color="166534")))
    P = H + 13; NPD = 10
    K.section(dash, f"A{P - 1}", "Renovation projects")
    K.hdr(dash, P, ["Project", "Status", "Budget", "Expected total", "Left (+) / over (-)", "% of budget"])
    for k in range(NPD):
        r = P + 1 + k; pr = P0 + k
        dash[f"A{r}"] = f'=IF(Projects!A{pr}="","",Projects!A{pr})'
        for col, src in zip("BCDEF", "DGMNO"):
            dash[f"{col}{r}"] = f'=IF(A{r}="","",IF(Projects!{src}{pr}="","",Projects!{src}{pr}))'
    style_range(dash, P + 1, P + NPD, 1, 6, fmt={3: MONEY, 4: MONEY, 5: MONEY, 6: "0%"})
    dash.conditional_formatting.add(f"E{P + 1}:E{P + NPD}", FormulaRule(formula=[f'AND(ISNUMBER(E{P + 1}),E{P + 1}<0)'], font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    dash.conditional_formatting.add(f"F{P + 1}:F{P + NPD}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1.2, color="5EEAD4"))
    ch = BarChart(); ch.type = "col"; ch.title = "Budget vs expected cost"; ch.height = 7.5; ch.width = 17
    for col, colr in ((3, "99F6E4"), (4, MAIN)):
        ch.add_data(Reference(dash, min_col=col, min_row=P, max_row=P + 5), titles_from_data=True)
        ch.series[-1].graphicalProperties.solidFill = colr
    ch.set_categories(Reference(dash, min_col=1, min_row=P + 1, max_row=P + 5)); ch.y_axis.majorGridlines = None
    ch.gapWidth = 60; ch.y_axis.scaling.min = 0; ch.legend.position = "b"; ch.legend.overlay = False
    dash.add_chart(ch, "I4")
    ch2 = BarChart(); ch2.type = "col"; ch2.grouping = "stacked"; ch2.overlap = 100; ch2.title = "Home spending per month"; ch2.height = 7.5; ch2.width = 17
    for row, colr in ((6, "F59E0B"), (7, MAIN)):
        ch2.add_data(Reference(mv, min_col=1, max_col=13, min_row=row, max_row=row), from_rows=True, titles_from_data=True)
        ch2.series[-1].graphicalProperties.solidFill = colr; ch2.series[-1].graphicalProperties.line.noFill = True
    ch2.set_categories(Reference(mv, min_col=2, max_col=13, min_row=5, max_row=5)); ch2.y_axis.majorGridlines = None
    ch2.gapWidth = 50; ch2.y_axis.scaling.min = 0; ch2.legend.position = "b"; ch2.legend.overlay = False
    dash.add_chart(ch2, "I20")
    for c in (ch, ch2): c.roundedCorners = False; c.x_axis.delete = False; c.y_axis.delete = False

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Settings: currency label, how many days ahead you want reminders, your renovation safety margin (10% is common) and your rooms/areas.", False),
             ("2.  Maintenance Schedule: a starter list of 27 common home jobs is already there (HVAC filter, gutters, smoke detectors, water heater…).", False),
             ("     Change how often (months), delete what you don't need, add your own. Type when each was last done – then Next due and Status fill in.", False),
             ("3.  Maintenance Log: write down every job when you do it (pick the task from the list). The schedule resets itself – no need to edit 'Last done'.", False),
             ("4.  Projects: one row per renovation – area, priority, status, dates and budget.", False),
             ("5.  Project Costs: every item, quote and invoice – estimate first, then the actual cost and what you have paid.", False),
             ("     Each project shows estimate, actual, paid, still owed, expected total and how much is left or over budget.", False),
             ("6.  Home Inventory: appliances and systems with model/serial numbers and warranty end dates.", False),
             ("7.  Dashboard: maintenance coming up, overdue jobs, renovation budget vs expected cost, money still owed and charts.", False),
             ("", False), ("Tips", True),
             ("•  Follow the maker's manual for your appliances and systems – the starter intervals are common guidelines only.", False),
             ("•  'Left / over' turns amber when the money left is smaller than your safety margin, and red when you are over budget.", False),
             ("•  Get at least two quotes for big jobs: add each as a line with status 'Quote', then delete the one you don't choose.", False),
             ("•  To see an earlier year in the Monthly View, type it in Settings > 'Year to show'.", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  150 maintenance tasks, 600 log entries, 30 projects, 800 cost lines, 100 inventory items. No macros. Excel 2010+ and Google Sheets.", False),
             ("•  To use Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("A made-up home with a year of maintenance, three renovation projects and eight appliances. Use the blank copy for your own home.", False)]
    K.start(st, "Home Maintenance & Renovation Planner", "Maintenance schedule with reminders, renovation budget, appliance & warranty inventory – Excel and Google Sheets.", lines)
    _orig_save = wb.save

    def _save(p):
        st.page_setup.fitToHeight = 1
        _orig_save(p)
    wb.save = _save
    finish(wb, sample, path)


META = {
    "name": "Home Maintenance Schedule & Renovation Budget Planner (Home Repair Tracker, Appliance Warranty Log) – Excel & Google Sheets",
    "file_base": "Home-Maintenance-Renovation-Budget-Planner",
    "price": 17,
    "color": "#0F766E",
    "title": "Home Maintenance & Renovation Planner",
    "tagline": "maintenance reminders · renovation budget · warranties",
    "intro": "Keep your home in shape and your renovation on budget – in one spreadsheet. A ready-made maintenance schedule reminds you when the HVAC filter, gutters, smoke detectors or water heater are due, and the renovation budget tracks every quote, invoice and payment so you always know what a project will really cost.",
    "features": [
        "Maintenance schedule with 27 common home jobs already filled in (HVAC filter, gutters, smoke & CO detectors, water heater, dryer vent, roof…) – edit or add your own",
        "Next due date, days left and Overdue / Due soon / OK status for every job; reminders reset when you log the job",
        "Maintenance log for scheduled jobs and one-off repairs, with cost and who did it",
        "Up to 30 renovation projects: area, priority, status, dates and budget",
        "Project cost lines (800): estimate, actual, paid, still owed – expected total uses the actual cost when you have it",
        "Left / over budget per project with a safety-margin (contingency) warning",
        "Home inventory: appliances and systems with model and serial numbers, purchase price and warranty status",
        "Dashboard: upcoming maintenance, overdue jobs, budget vs expected cost, money still owed, warranties ending soon, charts",
        "Monthly view of maintenance and renovation spending, by area and by project",
        "Blank copy plus a sample copy with a year of data and three projects",
    ],
    "tags": ["home maintenance schedule", "home maintenance checklist", "renovation budget", "home renovation planner", "remodel budget",
             "house project tracker", "google sheets", "excel template", "homeowner planner", "appliance warranty tracker",
             "home repair log", "new homeowner"],
    "summary": "Instant download: Excel + Google Sheets template, blank copy + sample copy.",
}
