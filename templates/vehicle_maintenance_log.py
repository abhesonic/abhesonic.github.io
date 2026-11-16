"""Vehicle Maintenance & Fuel Log – fuel economy (MPG or L/100km), service history, service reminders by date AND mileage, running costs.

Settings -> Vehicles -> Fuel Log -> Service Log -> Service Schedule -> Other Costs -> Monthly View -> Dashboard.
"""
import datetime as dt, random
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, DataBarRule
from openpyxl.chart import BarChart, Reference
from _common import Kit, F, B, MONEY, DATE, font, fill, style_range, input_cell, finish

MAIN = "1D4ED8"; SOFT = "DBEAFE"
K = Kit(MAIN, SOFT)
NV = 10; V0 = 6; V1 = V0 + NV - 1
NF = 1000; F0 = 6; F1 = F0 + NF - 1
NS = 600; S0 = 6; S1 = S0 + NS - 1
NR = 120; R0 = 6; R1 = R0 + NR - 1
NO = 600; O0 = 6; O1 = O0 + NO - 1
NT = 20; T0 = 6; T1 = T0 + NT - 1          # service types (Settings F)
NK = 12; K0 = 6; K1 = K0 + NK - 1          # other cost categories (Settings H)
STYPES = ["Oil & filter change", "Tire rotation", "New tires", "Brakes", "Battery", "Engine air filter", "Cabin air filter",
          "Wiper blades", "Coolant flush", "Transmission fluid", "Brake fluid", "Spark plugs", "Wheel alignment",
          "Inspection / emissions", "Timing belt", "Repair", "Other"]
OCATS = ["Insurance", "Registration & taxes", "Loan / lease payment", "Parking", "Tolls", "Car wash & detailing",
         "Accessories", "Fines & tickets", "Roadside assistance", "Other"]
ODO = '#,##0'; VOL = '0.00'; ECO = '0.0'


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); ve = wb.create_sheet("Vehicles"); fl = wb.create_sheet("Fuel Log")
    sl = wb.create_sheet("Service Log"); sc = wb.create_sheet("Service Schedule"); oc = wb.create_sheet("Other Costs")
    mv = wb.create_sheet("Monthly View"); s = wb.create_sheet("Settings")

    # ---------- Settings ----------
    K.title(s, "Settings", "Yellow cells are yours to edit.")
    SET = [(5, "Currency (label only)", "USD"), (6, "Distance unit (mi or km)", "mi"), (7, "Fuel unit (gal or L)", "gal"),
           (8, "Year to show (blank = this year)", None), (12, "Remind me this many days before a date", 30),
           (13, "Remind me this many miles/km before", 500)]
    for r, lab, v in SET:
        s[f"A{r}"] = lab; s[f"A{r}"].font = font(bold=True); s[f"A{r}"].border = B; input_cell(s[f"C{r}"], v)
    for col, lst, r in (("C", '"mi,km"', 6), ("C", '"gal,L"', 7)):
        dv = DataValidation(type="list", formula1=lst, allow_blank=False); s.add_data_validation(dv); dv.add(f"{col}{r}")
    s["A10"] = "Year shown (automatic)"; s["C10"] = "=IF(ISNUMBER(C8),C8,YEAR(TODAY()))"; s["C10"].number_format = "0"
    s["A11"] = "Months counted for averages"; s["C11"] = "=IF(C10<YEAR(TODAY()),12,IF(C10>YEAR(TODAY()),1,MONTH(TODAY())))"
    s["A14"] = "Fuel economy shown as (automatic)"; s["C14"] = '=IF(C6="km","L/100km",IF(C7="L","mi per L","MPG"))'
    for r in (10, 11, 14):
        s[f"A{r}"].font = font(italic=True); s[f"C{r}"].fill = fill("F3F4F6"); s[f"C{r}"].border = B; s[f"C{r}"].font = font()
    K.hdr(s, 5, ["Service types"], col0=6); K.hdr(s, 5, ["Other cost categories"], col0=8)
    for i in range(NT): input_cell(s[f"F{T0 + i}"], STYPES[i] if i < len(STYPES) else None)
    for i in range(NK): input_cell(s[f"H{K0 + i}"], OCATS[i] if i < len(OCATS) else None)
    for c, w in zip("ABCDEFGH", [40, 3, 12, 3, 3, 26, 3, 26]): s.column_dimensions[c].width = w
    YR = "Settings!$C$10"; MC = "Settings!$C$11"; DU = "Settings!$C$6"; LEAD = "Settings!$C$12"; LEADD = "Settings!$C$13"
    ECON = lambda d, v: f'IF({DU}="km",({v})/({d})*100,({d})/({v}))'

    FLr = lambda c: f"'Fuel Log'!${c}${F0}:${c}${F1}"
    SLr = lambda c: f"'Service Log'!${c}${S0}:${c}${S1}"
    SCr = lambda c: f"'Service Schedule'!${c}${R0}:${c}${R1}"
    OCr = lambda c: f"'Other Costs'!${c}${O0}:${c}${O1}"
    VEH = f"=Vehicles!$A${V0}:$A${V1}"

    # ---------- Vehicles ----------
    K.title(ve, "Vehicles", "One row per car, truck or motorbike. Yellow = you type, grey = automatic.")
    cols = ["Nickname", "Make", "Model", "Year", "Plate", "VIN", "Purchase date", "Purchase price", "Odometer when you started this log",
            "Insurance (company & policy)", "Insurance renews", "Registration renews", "Notes",
            "Latest odometer", "Distance tracked", "Avg fuel economy", "Fuel this year", "Service this year",
            "Other costs this year", "Total this year", "Cost per mile/km (all time)", "Service items due / overdue"]
    K.hdr(ve, 5, cols, [14, 11, 12, 7, 10, 19, 12, 12, 11, 24, 12, 12, 22, 11, 11, 11, 12, 12, 12, 12, 12, 11])
    ve["P5"] = '="Avg fuel economy ("&Settings!$C$14&")"'; ve["U5"] = '="Cost per "&Settings!$C$6&" (all time)"'
    style_range(ve, V0, V1, 1, 22, inputs=range(1, 14),
                fmt={4: "0", 7: DATE, 8: MONEY, 9: ODO, 11: DATE, 12: DATE, 14: ODO, 15: ODO, 16: ECO,
                     17: MONEY, 18: MONEY, 19: MONEY, 20: MONEY, 21: '0.00', 22: "0"})
    for r in range(V0, V1 + 1):
        ve[f"N{r}"] = (f'=IF(A{r}="","",MAX(N(I{r}),SUMPRODUCT(MAX(({FLr("B")}=A{r})*{FLr("C")})),'
                       f'SUMPRODUCT(MAX(({SLr("B")}=A{r})*{SLr("C")}))))')
        ve[f"O{r}"] = (f'=IF(A{r}="","",N{r}-IF(N(I{r})>0,I{r},IF(COUNTIFS({FLr("B")},A{r},{FLr("C")},">0")=0,N{r},'
                       f'SUMPRODUCT(MIN(({FLr("B")}<>A{r})*10^9+({FLr("C")}=0)*10^9+{FLr("C")})))))')
        ve[f"P{r}"] = (f'=IF(A{r}="","",IF(OR(SUMIFS({FLr("N")},{FLr("B")},A{r})<=0,SUMIFS({FLr("O")},{FLr("B")},A{r})<=0),"",'
                       f'{ECON(f"SUMIFS({FLr("N")},{FLr("B")},A{r})", f"SUMIFS({FLr("O")},{FLr("B")},A{r})")}))')
        ve[f"Q{r}"] = f'=IF(A{r}="","",SUMIFS({FLr("E")},{FLr("B")},A{r},{FLr("L")},{YR}))'
        ve[f"R{r}"] = f'=IF(A{r}="","",SUMIFS({SLr("I")},{SLr("B")},A{r},{SLr("J")},{YR}))'
        ve[f"S{r}"] = f'=IF(A{r}="","",SUMIFS({OCr("E")},{OCr("B")},A{r},{OCr("F")},{YR}))'
        ve[f"T{r}"] = f'=IF(A{r}="","",Q{r}+R{r}+S{r})'
        ve[f"U{r}"] = (f'=IF(A{r}="","",IF(N(O{r})<=0,"",(SUMIFS({FLr("E")},{FLr("B")},A{r})+SUMIFS({SLr("I")},{SLr("B")},A{r})'
                       f'+SUMIFS({OCr("E")},{OCr("B")},A{r}))/O{r}))')
        ve[f"V{r}"] = f'=IF(A{r}="","",COUNTIFS({SCr("A")},A{r},{SCr("N")},"Overdue")+COUNTIFS({SCr("A")},A{r},{SCr("N")},"Due soon"))'
    for col in "KL":
        ve.conditional_formatting.add(f"{col}{V0}:{col}{V1}", FormulaRule(formula=[f'AND(ISNUMBER({col}{V0}),{col}{V0}<TODAY())'],
                                      font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
        ve.conditional_formatting.add(f"{col}{V0}:{col}{V1}", FormulaRule(formula=[f'AND(ISNUMBER({col}{V0}),{col}{V0}-TODAY()<={LEAD})'],
                                      font=Font(name=F, bold=True, color="92400E"), fill=fill("FEF3C7")))
    ve.conditional_formatting.add(f"V{V0}:V{V1}", FormulaRule(formula=[f'N(V{V0})>0'], font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    ve.freeze_panes = f"B{V0}"

    # ---------- Fuel Log ----------
    K.title(fl, "Fuel Log", "One row per fill-up, oldest first. Fill the tank to full when you can: economy is worked out from full tank to full tank.")
    K.hdr(fl, 5, ["Date", "Vehicle", "Odometer", "Fuel added", "Total paid", "Full tank? (blank = Yes)", "Station", "Notes",
                  "Price per unit", "Distance since last fill", "Fuel economy", "Year", "Month", "(helper) distance", "(helper) fuel"],
          [12, 14, 11, 10, 11, 11, 18, 22, 10, 11, 11, 7, 7, 9, 9])
    fl["D5"] = '="Fuel added ("&Settings!$C$7&")"'; fl["K5"] = '="Fuel economy ("&Settings!$C$14&")"'
    fl["J5"] = '="Distance since last fill ("&Settings!$C$6&")"'
    style_range(fl, F0, F1, 1, 15, inputs=range(1, 9),
                fmt={1: DATE, 3: ODO, 4: VOL, 5: MONEY, 9: '0.000', 10: ODO, 11: ECO, 12: "0", 13: "0", 14: ODO, 15: VOL})
    for r in range(F0, F1 + 1):
        fl[f"I{r}"] = f'=IF(AND(N(D{r})>0,N(E{r})>0),E{r}/D{r},"")'
        fl[f"L{r}"] = f'=IF(ISNUMBER(A{r}),YEAR(A{r}),"")'; fl[f"M{r}"] = f'=IF(ISNUMBER(A{r}),MONTH(A{r}),"")'
        fl[f"N{r}"].font = font(color="9CA3AF", size=8); fl[f"O{r}"].font = font(color="9CA3AF", size=8)
        if r == F0:
            fl[f"J{r}"] = '=""'; fl[f"N{r}"] = 0; fl[f"O{r}"] = 0; fl[f"K{r}"] = '=""'; continue
        rg = lambda c: f"${c}${F0}:${c}{r - 1}"
        fl[f"J{r}"] = (f'=IF(OR(B{r}="",N(C{r})<=0),"",IF(COUNTIFS({rg("B")},B{r},{rg("C")},">0")=0,"",'
                       f'C{r}-LOOKUP(2,1/(({rg("B")}=B{r})*({rg("C")}>0)),{rg("C")})))')
        prevfull = f'LOOKUP(2,1/(({rg("B")}=B{r})*({rg("C")}>0)*({rg("F")}<>"No")),{rg("C")})'
        fl[f"N{r}"] = (f'=IF(OR(B{r}="",N(C{r})<=0,F{r}="No"),0,IF(COUNTIFS({rg("B")},B{r},{rg("C")},">0",{rg("F")},"<>No")=0,0,'
                       f'MAX(0,C{r}-{prevfull})))')
        fl[f"O{r}"] = (f'=IF(N{r}<=0,0,SUMIFS($D${F0}:$D{r},$B${F0}:$B{r},B{r},$C${F0}:$C{r},">"&(C{r}-N{r}),$C${F0}:$C{r},"<="&C{r}))')
        fl[f"K{r}"] = f'=IF(OR(N{r}<=0,O{r}<=0),"",{ECON(f"N{r}", f"O{r}")})'
    dv = DataValidation(type="list", formula1=VEH, allow_blank=True); fl.add_data_validation(dv); dv.add(f"B{F0}:B{F1}")
    dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True); fl.add_data_validation(dv); dv.add(f"F{F0}:F{F1}")
    fl.freeze_panes = f"A{F0}"; fl.auto_filter.ref = f"A5:M{F1}"

    # ---------- Service Log ----------
    K.title(sl, "Service Log", "Every service, repair or part – oil changes, tires, brakes, inspections. Pick the service type from the list.")
    K.hdr(sl, 5, ["Date", "Vehicle", "Odometer", "Service type", "What was done", "Shop / DIY", "Parts", "Labor", "Total", "Year", "Month"],
          [12, 14, 11, 20, 34, 20, 11, 11, 12, 7, 7])
    style_range(sl, S0, S1, 1, 11, inputs=range(1, 9), fmt={1: DATE, 3: ODO, 7: MONEY, 8: MONEY, 9: MONEY, 10: "0", 11: "0"})
    for r in range(S0, S1 + 1):
        sl[f"I{r}"] = f'=IF(AND(G{r}="",H{r}=""),"",N(G{r})+N(H{r}))'
        sl[f"J{r}"] = f'=IF(ISNUMBER(A{r}),YEAR(A{r}),"")'; sl[f"K{r}"] = f'=IF(ISNUMBER(A{r}),MONTH(A{r}),"")'
    dv = DataValidation(type="list", formula1=VEH, allow_blank=True); sl.add_data_validation(dv); dv.add(f"B{S0}:B{S1}")
    dv = DataValidation(type="list", formula1=f"=Settings!$F${T0}:$F${T1}", allow_blank=True); sl.add_data_validation(dv); dv.add(f"D{S0}:D{S1}")
    sl.freeze_panes = f"A{S0}"; sl.auto_filter.ref = f"A5:K{S1}"

    # ---------- Service Schedule ----------
    K.title(sc, "Service Schedule", "What to do and how often – by distance, by months, or both. Leave 'Last done' blank and it is read from the Service Log.")
    K.hdr(sc, 5, ["Vehicle", "Service item", "Every (distance)", "Every (months)", "Last done date (blank = from log)",
                  "Last done odometer (blank = from log)", "Notes", "Last done (used)", "Odometer then", "Next due odometer",
                  "Next due date", "Distance left", "Days left", "Status", "(sort key)"],
          [14, 22, 10, 9, 14, 14, 22, 12, 11, 11, 12, 10, 9, 14, 8])
    style_range(sc, R0, R1, 1, 15, inputs=range(1, 8),
                fmt={3: ODO, 4: "0", 5: DATE, 6: ODO, 8: DATE, 9: ODO, 10: ODO, 11: DATE, 12: '#,##0;-#,##0;0', 13: '0;-0;0'})
    for r in range(R0, R1 + 1):
        ok = f'OR(A{r}="",B{r}="")'
        mxd = f'SUMPRODUCT(MAX(({SLr("B")}=A{r})*({SLr("D")}=B{r})*{SLr("A")}))'
        mxo = f'SUMPRODUCT(MAX(({SLr("B")}=A{r})*({SLr("D")}=B{r})*{SLr("C")}))'
        sc[f"H{r}"] = f'=IF({ok},"",IF(ISNUMBER(E{r}),E{r},IF({mxd}>0,{mxd},"")))'
        sc[f"I{r}"] = f'=IF({ok},"",IF(ISNUMBER(F{r}),F{r},IF({mxo}>0,{mxo},"")))'
        sc[f"J{r}"] = f'=IF(AND(ISNUMBER(I{r}),N(C{r})>0),I{r}+C{r},"")'
        sc[f"K{r}"] = f'=IF(AND(ISNUMBER(H{r}),N(D{r})>0),EDATE(H{r},D{r}),"")'
        sc[f"L{r}"] = f'=IF(ISNUMBER(J{r}),J{r}-IFERROR(INDEX(Vehicles!$N${V0}:$N${V1},MATCH(A{r},Vehicles!$A${V0}:$A${V1},0)),0),"")'
        sc[f"M{r}"] = f'=IF(ISNUMBER(K{r}),K{r}-TODAY(),"")'
        sc[f"N{r}"] = (f'=IF({ok},"",IF(AND(NOT(ISNUMBER(L{r})),NOT(ISNUMBER(M{r}))),"Type last done",'
                       f'IF(OR(AND(ISNUMBER(L{r}),N(L{r})<0),AND(ISNUMBER(M{r}),N(M{r})<0)),"Overdue",'
                       f'IF(OR(AND(ISNUMBER(L{r}),N(L{r})<={LEADD}),AND(ISNUMBER(M{r}),N(M{r})<={LEAD})),"Due soon","OK"))))')
        sc[f"O{r}"] = (f'=IF({ok},"",IF(N{r}="Overdue",1,IF(N{r}="Due soon",2,IF(N{r}="OK",3,4)))*1000000'
                       f'+IF(ISNUMBER(M{r}),MAX(0,MIN(M{r}+100000,900000)),IF(ISNUMBER(L{r}),MAX(0,MIN(L{r}/10+100000,900000)),950000))+ROW()/10000)')
        sc[f"N{r}"].font = font(bold=True); sc[f"O{r}"].font = font(color="9CA3AF", size=8)
    dv = DataValidation(type="list", formula1=VEH, allow_blank=True); sc.add_data_validation(dv); dv.add(f"A{R0}:A{R1}")
    dv = DataValidation(type="list", formula1=f"=Settings!$F${T0}:$F${T1}", allow_blank=True); sc.add_data_validation(dv); dv.add(f"B{R0}:B{R1}")
    sc.conditional_formatting.add(f"N{R0}:N{R1}", FormulaRule(formula=[f'N{R0}="Overdue"'], font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    sc.conditional_formatting.add(f"N{R0}:N{R1}", FormulaRule(formula=[f'N{R0}="Due soon"'], font=Font(name=F, bold=True, color="92400E"), fill=fill("FEF3C7")))
    sc.conditional_formatting.add(f"N{R0}:N{R1}", FormulaRule(formula=[f'N{R0}="OK"'], font=Font(name=F, bold=True, color="166534")))
    sc.freeze_panes = f"C{R0}"; sc.auto_filter.ref = f"A5:N{R1}"

    # ---------- Other Costs ----------
    K.title(oc, "Other Costs", "Insurance, registration, loan or lease payments, parking, tolls, car wash… anything that is not fuel or service.")
    K.hdr(oc, 5, ["Date", "Vehicle", "Category", "What", "Amount", "Year", "Month"], [12, 14, 24, 34, 12, 7, 7])
    style_range(oc, O0, O1, 1, 7, inputs=range(1, 6), fmt={1: DATE, 5: MONEY, 6: "0", 7: "0"})
    for r in range(O0, O1 + 1):
        oc[f"F{r}"] = f'=IF(ISNUMBER(A{r}),YEAR(A{r}),"")'; oc[f"G{r}"] = f'=IF(ISNUMBER(A{r}),MONTH(A{r}),"")'
    dv = DataValidation(type="list", formula1=VEH, allow_blank=True); oc.add_data_validation(dv); dv.add(f"B{O0}:B{O1}")
    dv = DataValidation(type="list", formula1=f"=Settings!$H${K0}:$H${K1}", allow_blank=True); oc.add_data_validation(dv); dv.add(f"C{O0}:C{O1}")
    oc.freeze_panes = f"A{O0}"; oc.auto_filter.ref = f"A5:G{O1}"

    # ---------- Monthly View ----------
    K.title(mv, "Monthly View", "Running costs per month for the year chosen in Settings.")
    mv["A3"] = '="Year "&Settings!$C$10&"  ·  amounts in "&Settings!$C$5'; mv["A3"].font = font(bold=True)
    MN = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    K.hdr(mv, 5, ["Cost"] + MN + ["Year total"], [28] + [10] * 12 + [12])
    mv["A6"] = "Fuel"; mv["A7"] = "Service & repairs"
    for j in range(12):
        c = CL(2 + j)
        mv[f"{c}6"] = f'=SUMIFS({FLr("E")},{FLr("L")},{YR},{FLr("M")},{j + 1})'
        mv[f"{c}7"] = f'=SUMIFS({SLr("I")},{SLr("J")},{YR},{SLr("K")},{j + 1})'
    for k in range(NK):
        r = 8 + k; kr = K0 + k
        mv[f"A{r}"] = f'=IF(Settings!$H${kr}="","",Settings!$H${kr})'
        for j in range(12):
            mv[f"{CL(2 + j)}{r}"] = f'=IF($A{r}="",0,SUMIFS({OCr("E")},{OCr("C")},$A{r},{OCr("F")},{YR},{OCr("G")},{j + 1}))'
    TR = 8 + NK; OR_ = TR + 1; VR = TR + 2
    mv[f"A{TR}"] = "Total"; mv[f"A{OR_}"] = "Other costs (all categories)"; mv[f"A{VR}"] = '="Fuel added ("&Settings!$C$7&")"'
    for j in range(13):
        c = CL(2 + j)
        if j < 12:
            mv[f"{c}{TR}"] = f"=SUM({c}6:{c}{TR - 1})"; mv[f"{c}{OR_}"] = f"=SUM({c}8:{c}{TR - 1})"
            mv[f"{c}{VR}"] = f'=SUMIFS({FLr("D")},{FLr("L")},{YR},{FLr("M")},{j + 1})'
        else:
            for r in [6, 7] + list(range(8, TR)) + [TR, OR_, VR]: mv[f"N{r}"] = f"=SUM(B{r}:M{r})"
    style_range(mv, 6, VR, 1, 14, fmt={c: MONEY for c in range(2, 15)})
    for c in range(2, 15): mv.cell(row=VR, column=c).number_format = VOL
    for c in range(1, 15):
        cell = mv.cell(row=TR, column=c); cell.font = Font(name=F, bold=True, color=MAIN); cell.fill = fill(SOFT)
    PR0 = VR + 3
    K.hdr(mv, PR0 - 1, ["Vehicle (all costs)"] + MN + ["Year total"])
    for k in range(NV):
        r = PR0 + k; vr = V0 + k
        mv[f"A{r}"] = f'=IF(Vehicles!$A${vr}="","",Vehicles!$A${vr})'
        for j in range(12):
            m = j + 1
            mv[f"{CL(2 + j)}{r}"] = (f'=IF($A{r}="",0,SUMIFS({FLr("E")},{FLr("B")},$A{r},{FLr("L")},{YR},{FLr("M")},{m})'
                                    f'+SUMIFS({SLr("I")},{SLr("B")},$A{r},{SLr("J")},{YR},{SLr("K")},{m})'
                                    f'+SUMIFS({OCr("E")},{OCr("B")},$A{r},{OCr("F")},{YR},{OCr("G")},{m}))')
        mv[f"N{r}"] = f"=SUM(B{r}:M{r})"
    style_range(mv, PR0, PR0 + NV - 1, 1, 14, fmt={c: MONEY for c in range(2, 15)})
    mv.freeze_panes = "B6"

    # ---------- sample ----------
    if sample:
        rnd = random.Random(7)
        s["C8"] = 2026
        cars = [("Family SUV", "Toyota", "RAV4", 2019, "8KXR214", "2T3W1RFV5KW012345", dt.date(2021, 6, 15), 24900, 57900,
                 "GEICO – policy 4471-22-918", dt.date(2027, 3, 1), dt.date(2027, 6, 30), "Regular gas, 0W-20 oil"),
                ("Commuter", "Honda", "Civic", 2017, "7MBN552", "2HGFC2F59HH500123", dt.date(2020, 2, 3), 15400, 96200,
                 "GEICO – policy 4471-22-918", dt.date(2027, 3, 1), dt.date(2026, 12, 10), "Tires: 215/55R16")]
        for k, row in enumerate(cars):
            for c, v in zip("ABCDEFGHIJKLM", row): ve[f"{c}{V0 + k}"] = v
        fills = []
        for name, odo0, mpg, tank, gap in (("Family SUV", 58100, 28.5, 12.5, 11), ("Commuter", 96400, 34.0, 11.5, 8)):
            d = dt.date(2026, 1, 4 if name == "Commuter" else 6); odo = odo0
            while d <= dt.date(2026, 11, 12):
                partial = rnd.random() < 0.08
                gal = round(tank * rnd.uniform(0.75, 0.95) * (0.5 if partial else 1), 2)
                odo += int(gal * mpg * rnd.uniform(0.93, 1.06))
                price = round(3.05 + 0.25 * (1 if 4 <= d.month <= 8 else 0) + rnd.uniform(-0.12, 0.15), 3)
                fills.append((d, name, odo, gal, round(gal * price, 2), "No" if partial else None,
                              rnd.choice(["Shell", "Costco", "Chevron", "Exxon"]), "Road trip" if name == "Family SUV" and d.month == 7 else None))
                d += dt.timedelta(days=gap + rnd.randint(-2, 3))
        fills.sort(key=lambda x: x[0])
        for k, row in enumerate(fills):
            for c, v in zip("ABCDEFGH", row):
                if v is not None: fl[f"{c}{F0 + k}"] = v
        def odo_at(name, day):
            best = None
            for f in fills:
                if f[1] == name and f[0] <= day: best = f[2]
            return (best or (58100 if name == "Family SUV" else 96400)) + 12
        svc = [(dt.date(2026, 1, 20), "Family SUV", "Oil & filter change", "Synthetic 0W-20 + filter", "Jiffy Lube", 38, 30),
               (dt.date(2026, 1, 20), "Family SUV", "Tire rotation", "Rotate & balance", "Jiffy Lube", 0, 25),
               (dt.date(2026, 2, 11), "Commuter", "Oil & filter change", "Synthetic 0W-20 + filter", "DIY", 42, 0),
               (dt.date(2026, 3, 14), "Commuter", "Wiper blades", "Front wiper blades", "DIY", 28.98, 0),
               (dt.date(2026, 4, 2), "Commuter", "Brakes", "Front pads & rotors", "Midas", 186, 210),
               (dt.date(2026, 5, 9), "Commuter", "Inspection / emissions", "State smog check", "Quick Smog", 0, 49.75),
               (dt.date(2026, 5, 23), "Commuter", "Oil & filter change", "Synthetic 0W-20 + filter", "DIY", 42, 0),
               (dt.date(2026, 6, 18), "Family SUV", "Oil & filter change", "Synthetic 0W-20 + filter, pre road trip", "Toyota dealer", 45, 45),
               (dt.date(2026, 6, 18), "Family SUV", "Engine air filter", "Engine air filter", "Toyota dealer", 32, 15),
               (dt.date(2026, 6, 18), "Family SUV", "Tire rotation", "Rotate & balance", "Toyota dealer", 0, 0),
               (dt.date(2026, 8, 6), "Family SUV", "Repair", "AC recharge", "Toyota dealer", 60, 120),
               (dt.date(2026, 8, 29), "Commuter", "Oil & filter change", "Synthetic 0W-20 + filter", "DIY", 42, 0),
               (dt.date(2026, 9, 14), "Commuter", "Battery", "New battery (old one 5 yrs)", "AutoZone", 189.99, 0),
               (dt.date(2026, 10, 3), "Family SUV", "Wheel alignment", "4-wheel alignment", "Discount Tire", 0, 99.99)]
        for k, (d, n, t, w, shop, parts, lab) in enumerate(svc):
            r = S0 + k
            for c, v in zip("ABCDEF", (d, n, odo_at(n, d), t, w, shop)): sl[f"{c}{r}"] = v
            if parts: sl[f"G{r}"] = parts
            if lab: sl[f"H{r}"] = lab
        sched = [("Family SUV", "Oil & filter change", 5000, 6), ("Family SUV", "Tire rotation", 5000, 6),
                 ("Family SUV", "Engine air filter", 30000, 36), ("Family SUV", "Cabin air filter", 15000, 12),
                 ("Family SUV", "Brake fluid", None, 36), ("Family SUV", "Wheel alignment", None, 12),
                 ("Commuter", "Oil & filter change", 5000, 6), ("Commuter", "Tire rotation", 7500, None),
                 ("Commuter", "Inspection / emissions", None, 24), ("Commuter", "Wiper blades", None, 12),
                 ("Commuter", "Coolant flush", 60000, 60), ("Commuter", "Spark plugs", 100000, None)]
        for k, (n, item, dist, mon) in enumerate(sched):
            r = R0 + k; sc[f"A{r}"] = n; sc[f"B{r}"] = item
            if dist: sc[f"C{r}"] = dist
            if mon: sc[f"D{r}"] = mon
        r = R0 + 3; sc[f"E{r}"] = dt.date(2025, 11, 22); sc[f"F{r}"] = 56900; sc[f"G{r}"] = "Done at dealer last year"
        r = R0 + 7; sc[f"E{r}"] = dt.date(2026, 1, 30); sc[f"F{r}"] = 97350
        r = R0 + 10; sc[f"E{r}"] = dt.date(2022, 4, 1); sc[f"F{r}"] = 52000; sc[f"G{r}"] = "Honda Type 2 coolant"
        r = R0 + 11; sc[f"F{r}"] = 10; sc[f"G{r}"] = "Original plugs (bought new)"
        oth = []
        for m in range(1, 12):
            oth.append((dt.date(2026, m, 1), "Family SUV", "Loan / lease payment", "Auto loan – Chase", 389.00))
            oth.append((dt.date(2026, m, 15), "Commuter", "Parking", "Monthly garage pass", 95.00))
            oth.append((dt.date(2026, m, rnd.randint(5, 25)), "Commuter", "Tolls", "E-ZPass top-up", 35.00))
            if m % 2 == 0: oth.append((dt.date(2026, m, 20), "Family SUV", "Car wash & detailing", "Car wash", 14.00))
        oth += [(dt.date(2026, 3, 1), "Family SUV", "Insurance", "6-month premium", 642.00),
                (dt.date(2026, 3, 1), "Commuter", "Insurance", "6-month premium", 498.00),
                (dt.date(2026, 9, 1), "Family SUV", "Insurance", "6-month premium", 642.00),
                (dt.date(2026, 9, 1), "Commuter", "Insurance", "6-month premium", 498.00),
                (dt.date(2026, 6, 20), "Family SUV", "Registration & taxes", "Annual registration", 187.00),
                (dt.date(2026, 7, 9), "Family SUV", "Accessories", "Roof cargo box", 249.00),
                (dt.date(2026, 4, 17), "Commuter", "Fines & tickets", "Parking ticket", 45.00),
                (dt.date(2026, 1, 2), "Family SUV", "Roadside assistance", "AAA membership", 124.00)]
        oth.sort(key=lambda x: x[0])
        for k, row in enumerate(oth):
            for c, v in zip("ABCDE", row): oc[f"{c}{O0 + k}"] = v

    # ---------- Dashboard ----------
    K.title(dash, "Vehicle Dashboard", "Updates by itself from the other tabs.")
    dash["A3"] = '="Year "&Settings!$C$10&"  ·  amounts in "&Settings!$C$5&"  ·  today: "&TEXT(TODAY(),"mmm d, yyyy")'; dash["A3"].font = font(bold=True)
    TOT = f"'Monthly View'!$N${TR}"
    thism = f"INDEX('Monthly View'!$B${TR}:$M${TR},1,{MC})"
    yd = f'SUMIFS({FLr("N")},{FLr("L")},{YR})'; yf = f'SUMIFS({FLr("O")},{FLr("L")},{YR})'
    kp = [("Total cost this year", f"={TOT}", MONEY), ("Spent this month", f"={thism}", MONEY),
          ("Average per month", f"={TOT}/MAX(1,{MC})", MONEY),
          ("Fuel this year", "='Monthly View'!$N$6", MONEY), ("Service & repairs this year", "='Monthly View'!$N$7", MONEY),
          ("Other costs this year", f"='Monthly View'!$N${OR_}", MONEY),
          ("Avg fuel economy this year", f'=IF(OR({yd}<=0,{yf}<=0),"",{ECON(yd, yf)})', ECO),
          ("Service items overdue", f'=COUNTIFS({SCr("N")},"Overdue")', "0"),
          ("Service items due soon", f'=COUNTIFS({SCr("N")},"Due soon")', "0")]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 3) * 3, 1 + (i % 3) * 2, lab, f, fmt)
        dash.cell(row=6 + (i // 3) * 3, column=1 + (i % 3) * 2).alignment = Alignment(horizontal="left")
    dash["C5"] = f'="Spent in "&TEXT(DATE({YR},{MC},1),"mmmm")'
    dash["A11"] = '="Avg fuel economy this year ("&Settings!$C$14&")"'
    for c, w in zip("ABCDEFGH", [26, 20, 26, 14, 26, 14, 3, 3]): dash.column_dimensions[c].width = w
    dash.conditional_formatting.add("C12", FormulaRule(formula=["C12>0"], font=Font(name=F, size=18, bold=True, color="B91C1C")))
    dash.conditional_formatting.add("E12", FormulaRule(formula=["E12>0"], font=Font(name=F, size=18, bold=True, color="B45309")))
    H = 15
    K.section(dash, f"A{H - 1}", "Service reminders (overdue first, then due soon)")
    K.hdr(dash, H, ["Vehicle", "Service item", "Status", "Next due date", "Days left", "Distance left"])
    KEY = SCr("O")
    for k in range(10):
        r = H + 1 + k
        dash[f"H{r}"] = f'=IF(COUNT({KEY})<{k + 1},"",SMALL({KEY},{k + 1}))'; dash[f"H{r}"].font = font(color="FFFFFF", size=6)
        m = f"MATCH($H{r},{KEY},0)"
        for col, src in zip("ABCDEF", "ABNKML"):
            dash[f"{col}{r}"] = f'=IF($H{r}="","",INDEX({SCr(src)},{m}))'
    style_range(dash, H + 1, H + 10, 1, 6, fmt={4: "mmm d, yyyy", 5: "0", 6: '#,##0'})
    dash.conditional_formatting.add(f"C{H + 1}:C{H + 10}", FormulaRule(formula=[f'C{H + 1}="Overdue"'], font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    dash.conditional_formatting.add(f"C{H + 1}:C{H + 10}", FormulaRule(formula=[f'C{H + 1}="Due soon"'], font=Font(name=F, bold=True, color="92400E"), fill=fill("FEF3C7")))
    dash.conditional_formatting.add(f"C{H + 1}:C{H + 10}", FormulaRule(formula=[f'C{H + 1}="OK"'], font=Font(name=F, bold=True, color="166534")))
    P = H + 13
    K.section(dash, f"A{P - 1}", "Per vehicle")
    K.hdr(dash, P, ["Vehicle", "Latest odometer", "Avg fuel economy", "Total this year", "Cost per mile/km", "Insurance renews", "Registration renews"])
    dash[f"E{P}"] = '="Cost per "&Settings!$C$6'
    dash.column_dimensions["G"].width = 14
    for k in range(NV):
        r = P + 1 + k; vr = V0 + k
        dash[f"A{r}"] = f'=IF(Vehicles!A{vr}="","",Vehicles!A{vr})'
        for col, src in zip("BCDEFG", "NPTUKL"):
            dash[f"{col}{r}"] = f'=IF(A{r}="","",IF(Vehicles!{src}{vr}="","",Vehicles!{src}{vr}))'
    style_range(dash, P + 1, P + NV, 1, 7, fmt={2: ODO, 3: ECO, 4: MONEY, 5: '0.00', 6: "mmm d, yyyy", 7: "mmm d, yyyy"})
    for col in "FG":
        dash.conditional_formatting.add(f"{col}{P + 1}:{col}{P + NV}", FormulaRule(formula=[f'AND(ISNUMBER({col}{P + 1}),{col}{P + 1}-TODAY()<={LEAD})'],
                                        font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    dash.conditional_formatting.add(f"D{P + 1}:D{P + NV}", DataBarRule(start_type="num", start_value=0, end_type="max", color="93C5FD"))
    ch = BarChart(); ch.type = "col"; ch.grouping = "stacked"; ch.overlap = 100; ch.title = "Costs per month"; ch.height = 7.5; ch.width = 17
    for row, colr in ((6, "3B82F6"), (7, "F59E0B"), (OR_, "94A3B8")):
        ch.add_data(Reference(mv, min_col=1, max_col=13, min_row=row, max_row=row), from_rows=True, titles_from_data=True)
        ch.series[-1].graphicalProperties.solidFill = colr; ch.series[-1].graphicalProperties.line.noFill = True
    ch.set_categories(Reference(mv, min_col=2, max_col=13, min_row=5, max_row=5)); ch.y_axis.majorGridlines = None
    ch.gapWidth = 50; ch.y_axis.scaling.min = 0; ch.legend.position = "b"; ch.legend.overlay = False
    dash.add_chart(ch, "I4")
    ch2 = BarChart(); ch2.type = "bar"; ch2.title = "Fuel added per month"; ch2.height = 7.5; ch2.width = 17; ch2.legend = None
    ch2.type = "col"
    ch2.add_data(Reference(mv, min_col=2, max_col=13, min_row=VR, max_row=VR), from_rows=True, titles_from_data=False)
    ch2.set_categories(Reference(mv, min_col=2, max_col=13, min_row=5, max_row=5)); ch2.y_axis.majorGridlines = None
    ch2.series[0].graphicalProperties.solidFill = MAIN; ch2.gapWidth = 50; ch2.y_axis.scaling.min = 0
    dash.add_chart(ch2, "I20")
    for c in (ch, ch2): c.roundedCorners = False; c.x_axis.delete = False; c.y_axis.delete = False

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Settings: currency label, miles or km, gallons or liters, and how early you want reminders (days and distance).", False),
             ("2.  Vehicles: one row per vehicle – make, model, plate, VIN, the odometer when you start, insurance and registration renewal dates.", False),
             ("3.  Fuel Log: every fill-up – date, vehicle, odometer, fuel added and total paid. Keep rows oldest first.", False),
             ("     Fuel economy (MPG, or L/100km if you choose km) is worked out from one full tank to the next. For a partial fill choose 'No' in 'Full tank?'.", False),
             ("4.  Service Log: oil changes, tires, brakes, battery, inspections and repairs, with parts and labor cost.", False),
             ("5.  Service Schedule: what to do and how often – every X miles/km, every X months, or both (whichever comes first).", False),
             ("     Leave 'Last done' blank and the latest matching entry in the Service Log is used. Status shows OK / Due soon / Overdue.", False),
             ("6.  Other Costs: insurance, registration, loan or lease, parking, tolls, car wash…", False),
             ("7.  Dashboard: total running cost, fuel economy, cost per mile/km, service reminders and renewal dates for every vehicle.", False),
             ("", False), ("Tips", True),
             ("•  Use the intervals in your owner's manual for the Service Schedule – your manual comes first.", False),
             ("•  Change 'Remind me…' in Settings to see reminders earlier or later.", False),
             ("•  Edit the service types and cost categories in Settings to match your vehicles (motorbike, RV, truck, EV charging…).", False),
             ("•  To see an earlier year, type it in Settings > 'Year to show'.", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  10 vehicles, 1,000 fill-ups, 600 services, 120 schedule items, 600 other costs. No macros. Works in Excel 2010+ and Google Sheets.", False),
             ("•  To use Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("Two made-up cars with a year of fill-ups, services and costs. Use the blank copy in your download for your own vehicles.", False)]
    K.start(st, "Vehicle Maintenance & Fuel Log", "Fuel economy, service history, service reminders by date and mileage, running costs – Excel and Google Sheets.", lines)
    _orig_save = wb.save

    def _save(p):
        st.page_setup.fitToHeight = 1
        _orig_save(p)
    wb.save = _save
    finish(wb, sample, path)


META = {
    "name": "Vehicle Maintenance & Fuel Log (MPG Tracker, Car Service Reminders, Car Expense Tracker) – Excel & Google Sheets",
    "file_base": "Vehicle-Maintenance-Fuel-Log",
    "price": 14,
    "color": "#1D4ED8",
    "title": "Vehicle Maintenance & Fuel Log",
    "tagline": "fuel economy · service reminders · car costs · up to 10 vehicles",
    "intro": "Know what your car really costs and never miss an oil change again. Log every fill-up and the tracker works out your fuel economy (MPG or L/100km), log every service and repair, and set reminders by mileage, by date or both – the dashboard shows what is overdue, what is due soon and your cost per mile for every vehicle.",
    "features": [
        "Up to 10 vehicles: make, model, plate, VIN, starting odometer, insurance and registration renewal dates (highlighted when they come up)",
        "Fuel log (1,000 fill-ups): price per gallon/liter, distance since last fill and fuel economy from full tank to full tank – partial fills handled",
        "Miles or kilometers, gallons or liters: MPG or L/100km chosen automatically",
        "Service log with parts and labor cost – oil changes, tires, brakes, battery, inspections, repairs",
        "Service schedule with reminders every X miles/km and/or every X months: next due odometer, next due date, Overdue / Due soon / OK",
        "'Last done' is read from the service log automatically, so reminders reset by themselves",
        "Other costs: insurance, registration, loan or lease, parking, tolls, car wash",
        "Dashboard: total running cost, this month, average per month, fuel economy, cost per mile/km, service reminders, renewal dates, charts",
        "Monthly view per cost type and per vehicle",
        "Blank copy plus a sample copy with two cars and a year of data",
    ],
    "tags": ["vehicle maintenance log", "car maintenance tracker", "fuel log", "mpg tracker", "car expense tracker", "mileage log",
             "oil change tracker", "google sheets", "excel template", "car service record", "auto maintenance", "vehicle expense"],
    "summary": "Instant download: Excel + Google Sheets template, blank copy + sample copy.",
}
