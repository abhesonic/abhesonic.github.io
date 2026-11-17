"""Garden Planner & Planting Calendar – frost-date based sowing/transplant/harvest dates, week-by-week calendar,
bed space, harvest log with value, garden expenses.

Settings -> Plant Library -> Plantings -> Calendar -> Harvest Log -> Expenses -> Season Summary -> Dashboard.
"""
import datetime as dt, random
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, DataBarRule
from openpyxl.chart import BarChart, Reference
from _common import Kit, F, B, MONEY, DATE, font, fill, style_range, input_cell, finish

MAIN = "15803D"; SOFT = "DCFCE7"
K = Kit(MAIN, SOFT)
NL = 70; L0 = 6; L1 = L0 + NL - 1          # plant library
NP = 150; P0 = 6; P1 = P0 + NP - 1         # plantings
NH = 1000; H0 = 6; H1 = H0 + NH - 1        # harvest log
NE = 400; E0 = 6; E1 = E0 + NE - 1         # expenses
NB = 20; B0 = 6; B1 = B0 + NB - 1          # beds
NW = 52
ECATS = ["Seeds", "Plants & seedlings", "Soil & compost", "Fertilizer", "Mulch", "Tools & equipment", "Pots & raised beds",
         "Pest control", "Water", "Other"]
UNITS = "lb,kg,oz,g,count,bunch,head"
METHODS = "Start indoors,Direct sow,Buy transplants"
# crop, family, season, based on, start indoors (wks), transplant (wks), direct sow (wks), days to harvest, harvest weeks, spacing in, notes
LIB = [
    ("Tomato", "Nightshade", "Warm", "Last spring frost", -7, 1, None, 75, 10, 24, "Days counted from transplant. Stake or cage."),
    ("Cherry tomato", "Nightshade", "Warm", "Last spring frost", -7, 1, None, 65, 12, 24, "Very productive; pick often."),
    ("Pepper (sweet)", "Nightshade", "Warm", "Last spring frost", -9, 2, None, 75, 10, 18, "Needs warm soil (65°F+)."),
    ("Pepper (hot)", "Nightshade", "Warm", "Last spring frost", -9, 2, None, 80, 10, 18, ""),
    ("Eggplant", "Nightshade", "Warm", "Last spring frost", -9, 2, None, 75, 8, 24, ""),
    ("Potato", "Nightshade", "Cool", "Last spring frost", None, None, -2, 90, 3, 12, "Plant seed potatoes; hill as they grow."),
    ("Broccoli", "Brassica", "Cool", "Last spring frost", -7, -2, None, 65, 3, 18, "Side shoots after main head."),
    ("Cabbage", "Brassica", "Cool", "Last spring frost", -7, -3, None, 70, 3, 18, ""),
    ("Cauliflower", "Brassica", "Cool", "Last spring frost", -7, -2, None, 60, 2, 18, ""),
    ("Kale", "Brassica", "Cool", "Last spring frost", -6, -3, -4, 55, 16, 12, "Frost improves flavour."),
    ("Brussels sprouts", "Brassica", "Cool", "Last spring frost", -6, 2, None, 100, 6, 24, "Harvest after first fall frosts."),
    ("Radish", "Brassica", "Cool", "Last spring frost", None, None, -4, 28, 2, 2, "Sow every 2 weeks."),
    ("Turnip", "Brassica", "Cool", "Last spring frost", None, None, -3, 50, 3, 4, ""),
    ("Lettuce", "Daisy", "Cool", "Last spring frost", -6, -2, -4, 45, 4, 8, "Sow every 2–3 weeks for a steady supply."),
    ("Spinach", "Amaranth", "Cool", "Last spring frost", None, None, -6, 40, 4, 4, "Bolts in heat."),
    ("Swiss chard", "Amaranth", "Cool", "Last spring frost", -5, -1, -2, 55, 16, 8, "Cut-and-come-again."),
    ("Beet", "Amaranth", "Cool", "Last spring frost", None, None, -3, 55, 4, 4, ""),
    ("Carrot", "Carrot", "Cool", "Last spring frost", None, None, -3, 70, 6, 2, "Keep soil moist until germination."),
    ("Parsnip", "Carrot", "Cool", "Last spring frost", None, None, -2, 110, 8, 4, ""),
    ("Peas", "Legume", "Cool", "Last spring frost", None, None, -6, 60, 4, 2, "Give them a trellis."),
    ("Bush beans", "Legume", "Warm", "Last spring frost", None, None, 1, 55, 4, 4, ""),
    ("Pole beans", "Legume", "Warm", "Last spring frost", None, None, 1, 65, 8, 6, "Needs a trellis."),
    ("Cucumber", "Cucurbit", "Warm", "Last spring frost", -3, 2, 1, 55, 8, 12, "Trellis to save space."),
    ("Zucchini / summer squash", "Cucurbit", "Warm", "Last spring frost", -3, 2, 1, 50, 10, 24, "Pick small and often."),
    ("Winter squash", "Cucurbit", "Warm", "Last spring frost", -3, 2, 2, 95, 4, 36, ""),
    ("Pumpkin", "Cucurbit", "Warm", "Last spring frost", -3, 2, 2, 100, 4, 48, ""),
    ("Melon / cantaloupe", "Cucurbit", "Warm", "Last spring frost", -3, 2, None, 85, 4, 24, ""),
    ("Watermelon", "Cucurbit", "Warm", "Last spring frost", -3, 2, None, 85, 4, 36, ""),
    ("Sweet corn", "Grass", "Warm", "Last spring frost", None, None, 1, 75, 2, 12, "Plant in blocks for pollination."),
    ("Onion (from seed)", "Allium", "Cool", "Last spring frost", -10, -4, None, 100, 4, 4, ""),
    ("Leek", "Allium", "Cool", "Last spring frost", -10, -2, None, 110, 10, 6, ""),
    ("Garlic", "Allium", "Cool", "First fall frost", None, None, 3, 240, 2, 6, "Plant cloves in fall, harvest next summer."),
    ("Basil", "Herb", "Warm", "Last spring frost", -6, 1, 2, 60, 12, 10, "Pinch flowers off."),
    ("Parsley", "Herb", "Cool", "Last spring frost", -9, -2, None, 75, 20, 8, ""),
    ("Cilantro", "Herb", "Cool", "Last spring frost", None, None, -2, 50, 3, 6, "Sow every 3 weeks."),
    ("Dill", "Herb", "Warm", "Last spring frost", None, None, 0, 50, 6, 10, ""),
    ("Strawberry", "Fruit", "Cool", "Last spring frost", None, -4, None, 90, 4, 12, "Bare-root plants."),
    ("Sunflower", "Flower", "Warm", "Last spring frost", None, None, 1, 80, 3, 12, ""),
    ("Zinnia", "Flower", "Warm", "Last spring frost", -5, 1, 1, 75, 12, 9, "Cut flowers all summer."),
    ("Marigold", "Flower", "Warm", "Last spring frost", -6, 1, None, 50, 14, 8, "Companion plant for tomatoes."),
    ("Fall lettuce", "Daisy", "Cool", "First fall frost", -12, -8, -8, 45, 6, 8, "Fall crop – dates count back from first fall frost."),
    ("Fall spinach", "Amaranth", "Cool", "First fall frost", None, None, -6, 40, 8, 4, "Survives light frost."),
    ("Fall kale", "Brassica", "Cool", "First fall frost", -14, -10, None, 55, 12, 12, "Fall crop."),
]


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); pl = wb.create_sheet("Plantings"); ca = wb.create_sheet("Calendar")
    hv = wb.create_sheet("Harvest Log"); ex = wb.create_sheet("Expenses"); su = wb.create_sheet("Season Summary")
    lb = wb.create_sheet("Plant Library"); s = wb.create_sheet("Settings")

    # ---------- Settings ----------
    K.title(s, "Settings", "Yellow cells are yours to edit. Find your frost dates from your local extension office or a frost-date website (search: 'frost dates' + your ZIP / town).")
    SET = [("Last spring frost date", None, DATE), ("First fall frost date", None, DATE), ("Currency (label only)", "USD", None),
           ("Harvest weight unit (lb or kg)", "lb", None), ("Growing zone / notes (optional)", None, None),
           ("Show tasks due within (days)", 14, "0")]
    for i, (lab, v, fmt) in enumerate(SET, 5):
        s[f"A{i}"] = lab; s[f"A{i}"].font = font(bold=True); s[f"A{i}"].border = B; input_cell(s[f"C{i}"], v)
        if fmt: s[f"C{i}"].number_format = fmt
    dv = DataValidation(type="list", formula1='"lb,kg"', allow_blank=True); s.add_data_validation(dv); dv.add("C8")
    s["A12"] = "Season year (automatic)"; s["A12"].font = font(italic=True)
    s["C12"] = "=IF(ISNUMBER(C5),YEAR(C5),YEAR(TODAY()))"; s["C12"].number_format = "0"
    s["A13"] = "Frost-free days (automatic)"; s["A13"].font = font(italic=True)
    s["C13"] = '=IF(AND(ISNUMBER(C5),ISNUMBER(C6)),C6-C5,"")'; s["C13"].number_format = "0"
    for r in (12, 13): s[f"C{r}"].fill = fill("F3F4F6"); s[f"C{r}"].border = B; s[f"C{r}"].font = font()
    K.hdr(s, 5, ["Bed / area name", "Length (ft)", "Width (ft)", "Area (sq ft)", "Planned (sq ft)", "Free (sq ft)"], col0=6)
    style_range(s, B0, B1, 6, 11, inputs=(6, 7, 8), fmt={7: "0.0", 8: "0.0", 9: "0.0", 10: "0.0", 11: '0.0;[Red]-0.0;0.0'})
    for r in range(B0, B1 + 1):
        s[f"I{r}"] = f'=IF(F{r}="","",N(G{r})*N(H{r}))'
        s[f"J{r}"] = f"=IF(F{r}=\"\",\"\",SUMIFS(Plantings!$S${P0}:$S${P1},Plantings!$C${P0}:$C${P1},F{r}))"
        s[f"K{r}"] = f'=IF(F{r}="","",I{r}-J{r})'
    s[f"F{B1 + 1}"] = "Total"; s[f"F{B1 + 1}"].font = font(bold=True)
    for c in "IJK":
        s[f"{c}{B1 + 1}"] = f"=SUM({c}{B0}:{c}{B1})"; s[f"{c}{B1 + 1}"].font = font(bold=True); s[f"{c}{B1 + 1}"].number_format = "0.0"
    for c, w in zip("ABCDEFGHIJK", [36, 3, 14, 3, 3, 24, 11, 11, 12, 13, 12]): s.column_dimensions[c].width = w
    SPR = "Settings!$C$5"; FAL = "Settings!$C$6"; YR = "Settings!$C$12"; WU = "Settings!$C$8"; SOON = "Settings!$C$10"

    # ---------- Plant Library ----------
    K.title(lb, "Plant Library", "Timing in weeks relative to your frost date (-6 = six weeks BEFORE, 2 = two weeks AFTER). Blank = not usual for this crop. Edit for your climate or add your own crops.")
    K.hdr(lb, 5, ["Crop", "Family", "Season", "Timing based on", "Start indoors (wks)", "Transplant out (wks)", "Direct sow (wks)",
                  "Days to harvest", "Harvest lasts (wks)", "Spacing (in)", "Notes"], [24, 12, 9, 17, 11, 11, 10, 10, 10, 9, 44])
    style_range(lb, L0, L1, 1, 11, inputs=range(1, 12), fmt={5: "0", 6: "0", 7: "0", 8: "0", 9: "0", 10: "0"})
    for k, row in enumerate(LIB):
        for c, v in enumerate(row, 1):
            if v not in (None, ""): lb.cell(row=L0 + k, column=c, value=v)
    for col, lst in (("C", '"Cool,Warm"'), ("D", '"Last spring frost,First fall frost"')):
        dv = DataValidation(type="list", formula1=lst, allow_blank=True); lb.add_data_validation(dv); dv.add(f"{col}{L0}:{col}{L1}")
    lb.freeze_panes = f"B{L0}"
    LB = lambda c: f"'Plant Library'!${c}${L0}:${c}${L1}"

    # ---------- Plantings ----------
    K.title(pl, "Plantings", "Your plan: one row per crop (or per sowing for succession crops). Yellow = you type, grey = calculated from your frost dates and the Plant Library.")
    cols = ["Crop", "Variety", "Bed / area", "Method", "Plants (qty)", "Shift (wks, +later / -earlier)", "Or my own plant-out / sow date",
            "Seeds started indoors on", "Planted out / sown on", "First harvest on", "Finished?", "Notes",
            "Timing based on", "Frost date used", "Start indoors", "Plant out / sow", "Harvest from", "Harvest until",
            "Space (sq ft)", "Status", "Next task", "Task date", "Days left", "When", "(sort key)", "(lib row)", "(eff. start)", "(eff. plant)"]
    K.hdr(pl, 5, cols, [20, 16, 14, 15, 8, 10, 13, 12, 12, 12, 9, 22, 15, 12, 12, 12, 12, 12, 9, 17, 19, 12, 8, 16, 7, 5, 6, 6])
    D2 = "mmm d"
    style_range(pl, P0, P1, 1, 28, inputs=range(1, 13),
                fmt={5: "0", 6: "0", 7: DATE, 8: DATE, 9: DATE, 10: DATE, 14: DATE, 15: D2, 16: D2, 17: D2, 18: D2, 19: "0.0", 22: DATE, 23: '0;-0;0'})
    for r in range(P0, P1 + 1):
        Z = f"$Z{r}"
        I = lambda c: f"INDEX({LB(c)},{Z})"
        pl[f"Z{r}"] = f'=IF(A{r}="","",IFERROR(MATCH(A{r},{LB("A")},0),""))'
        pl[f"M{r}"] = f'=IF(OR(A{r}="",Z{r}=""),"",IF({I("D")}="First fall frost","First fall frost","Last spring frost"))'
        pl[f"N{r}"] = f'=IF(M{r}="","",IF(M{r}="First fall frost",IF(ISNUMBER({FAL}),{FAL},""),IF(ISNUMBER({SPR}),{SPR},"")))'
        pl[f"P{r}"] = (f'=IF(OR(A{r}="",Z{r}=""),"",IF(ISNUMBER(G{r}),G{r},IF(NOT(ISNUMBER(N{r})),"",'
                       f'IF(D{r}="Direct sow",IF(ISNUMBER({I("G")}),N{r}+({I("G")}+N(F{r}))*7,""),'
                       f'IF(ISNUMBER({I("F")}),N{r}+({I("F")}+N(F{r}))*7,"")))))')
        pl[f"O{r}"] = (f'=IF(OR(A{r}="",Z{r}="",D{r}<>"Start indoors"),"",IF(NOT(ISNUMBER({I("E")})),"",'
                       f'IF(ISNUMBER(G{r}),IF(ISNUMBER({I("F")}),G{r}-({I("F")}-{I("E")})*7,""),'
                       f'IF(ISNUMBER(N{r}),N{r}+({I("E")}+N(F{r}))*7,""))))')
        pl[f"AA{r}"] = f'=IF(ISNUMBER(H{r}),H{r},O{r})'
        pl[f"AB{r}"] = f'=IF(ISNUMBER(I{r}),I{r},P{r})'
        pl[f"Q{r}"] = f'=IF(OR(A{r}="",Z{r}=""),"",IF(ISNUMBER(AB{r}),AB{r}+N({I("H")}),""))'
        pl[f"R{r}"] = f'=IF(ISNUMBER(Q{r}),IF(ISNUMBER(J{r}),MAX(J{r},Q{r}),Q{r})+N({I("I")})*7,"")'
        pl[f"S{r}"] = f'=IF(OR(A{r}="",Z{r}="",N(E{r})=0),"",E{r}*(N({I("J")})/12)^2)'
        pl[f"T{r}"] = (f'=IF(A{r}="","",IF(Z{r}="","Add crop to Plant Library",IF(K{r}="Yes","Finished",IF(ISNUMBER(J{r}),"Harvesting",'
                       f'IF(ISNUMBER(I{r}),"Growing",IF(ISNUMBER(H{r}),"Seedlings indoors",IF(NOT(ISNUMBER(N{r})),"Add frost dates in Settings",'
                       f'IF(NOT(ISNUMBER(P{r})),"Pick another method","Planned"))))))))')
        pl[f"U{r}"] = (f'=IF(OR(A{r}="",Z{r}="",K{r}="Yes",NOT(ISNUMBER(AB{r}))),"",IF(AND(D{r}="Start indoors",ISNUMBER(O{r}),NOT(ISNUMBER(H{r})),NOT(ISNUMBER(I{r}))),"Start seeds indoors",'
                       f'IF(NOT(ISNUMBER(I{r})),IF(D{r}="Direct sow","Sow outside","Plant out"),IF(NOT(ISNUMBER(J{r})),"First harvest","Harvest ends – clear bed"))))')
        pl[f"V{r}"] = (f'=IF(U{r}="","",IF(U{r}="Start seeds indoors",O{r},IF(OR(U{r}="Sow outside",U{r}="Plant out"),P{r},IF(U{r}="First harvest",Q{r},R{r}))))')
        pl[f"W{r}"] = f'=IF(ISNUMBER(V{r}),V{r}-TODAY(),"")'
        pl[f"X{r}"] = (f'=IF(NOT(ISNUMBER(W{r})),"",IF(W{r}<0,"Late by "&-W{r}&IF(W{r}=-1," day"," days"),IF(W{r}=0,"Today",'
                       f'IF(W{r}<={SOON},"In "&W{r}&IF(W{r}=1," day"," days"),"Later"))))')
        pl[f"Y{r}"] = f'=IF(ISNUMBER(V{r}),V{r}+ROW()/100000,"")'
        for c in ("Y", "Z", "AA", "AB"): pl[f"{c}{r}"].font = font(color="9CA3AF", size=7)
        pl[f"AA{r}"].number_format = D2; pl[f"AB{r}"].number_format = D2
        pl[f"T{r}"].font = font(bold=True)
    dv = DataValidation(type="list", formula1=f"={LB('A')}", allow_blank=True); pl.add_data_validation(dv); dv.add(f"A{P0}:A{P1}")
    dv = DataValidation(type="list", formula1=f"=Settings!$F${B0}:$F${B1}", allow_blank=True); pl.add_data_validation(dv); dv.add(f"C{P0}:C{P1}")
    dv = DataValidation(type="list", formula1=f'"{METHODS}"', allow_blank=True); pl.add_data_validation(dv); dv.add(f"D{P0}:D{P1}")
    dv = DataValidation(type="list", formula1='"Yes"', allow_blank=True); pl.add_data_validation(dv); dv.add(f"K{P0}:K{P1}")
    red = dict(font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2"))
    amb = dict(font=Font(name=F, bold=True, color="92400E"), fill=fill("FEF3C7"))
    pl.conditional_formatting.add(f"X{P0}:X{P1}", FormulaRule(formula=[f'LEFT(X{P0},7)="Late by"'], **red))
    pl.conditional_formatting.add(f"X{P0}:X{P1}", FormulaRule(formula=[f'OR(X{P0}="Today",LEFT(X{P0},2)="In")'], **amb))
    pl.conditional_formatting.add(f"T{P0}:T{P1}", FormulaRule(formula=[f'OR(LEFT(T{P0},3)="Add",LEFT(T{P0},4)="Pick")'], **red))
    pl.conditional_formatting.add(f"T{P0}:T{P1}", FormulaRule(formula=[f'T{P0}="Harvesting"'], font=Font(name=F, bold=True, color="166534"), fill=fill("DCFCE7")))
    pl.conditional_formatting.add(f"T{P0}:T{P1}", FormulaRule(formula=[f'T{P0}="Finished"'], font=Font(name=F, color="9CA3AF")))
    pl.freeze_panes = f"C{P0}"; pl.auto_filter.ref = f"A5:X{P1}"
    PL = lambda c: f"Plantings!${c}${P0}:${c}${P1}"

    # ---------- Calendar ----------
    K.title(ca, "Planting Calendar", "Week by week for the season year.  S = start seeds indoors · T = plant out · D = direct sow · H = harvest.  Light shading = seedlings indoors / growing.")
    ca["A3"] = '="Season "&Settings!$C$12&"   ·   last spring frost: "&IF(ISNUMBER(Settings!$C$5),TEXT(Settings!$C$5,"mmm d"),"(set in Settings)")&"   ·   first fall frost: "&IF(ISNUMBER(Settings!$C$6),TEXT(Settings!$C$6,"mmm d"),"(set in Settings)")'
    ca["A3"].font = font(bold=True)
    K.hdr(ca, 5, ["Crop & variety", "Bed"], [26, 11])
    ca.row_dimensions[5].height = 42
    for j in range(NW):
        c = ca.cell(row=5, column=3 + j, value=f"=DATE({YR},1,1)-WEEKDAY(DATE({YR},1,1),3)+{7 * j}")
        c.number_format = "mmm d"; c.font = Font(name=F, bold=True, color="FFFFFF", size=7); c.fill = fill(MAIN)
        c.alignment = Alignment(text_rotation=90, horizontal="center", vertical="center"); c.border = B
        ca.column_dimensions[CL(3 + j)].width = 3.2
    for k in range(NP):
        r = 6 + k; pr = P0 + k
        ca[f"A{r}"] = f'=IF(Plantings!A{pr}="","",Plantings!A{pr}&IF(Plantings!B{pr}="",""," – "&Plantings!B{pr}))'
        ca[f"B{r}"] = f'=IF(Plantings!C{pr}="","",Plantings!C{pr})'
        S_ = f"Plantings!$AA{pr}"; P_ = f"Plantings!$AB{pr}"; Q_ = f"Plantings!$Q{pr}"; R_ = f"Plantings!$R{pr}"; M_ = f"Plantings!$D{pr}"
        for j in range(NW):
            w = f"{CL(3 + j)}$5"
            ca.cell(row=r, column=3 + j, value=(
                f'=IF($A{r}="","",IF(AND(ISNUMBER({S_}),{S_}>={w},{S_}<{w}+7),"S",IF(AND(ISNUMBER({P_}),{P_}>={w},{P_}<{w}+7),IF({M_}="Direct sow","D","T"),'
                f'IF(AND(ISNUMBER({Q_}),{Q_}<{w}+7,{R_}>={w}),"H",IF(AND(ISNUMBER({S_}),ISNUMBER({P_}),{S_}<{w},{P_}>={w}+7),"·",'
                f'IF(AND(ISNUMBER({P_}),ISNUMBER({Q_}),{P_}<{w},{Q_}>={w}+7),"·",""))))))'))
    last = CL(2 + NW)
    for rr in range(6, 6 + NP):
        for cc in range(1, 3 + NW):
            cell = ca.cell(row=rr, column=cc); cell.border = B; cell.font = font(size=8 if cc > 2 else 9, bold=cc > 2)
            if cc > 2: cell.alignment = Alignment(horizontal="center")
    rng = f"C6:{last}{5 + NP}"
    for letter, bg, fg in (("S", "C4B5FD", "4C1D95"), ("T", "86EFAC", "14532D"), ("D", "FDE68A", "78350F"), ("H", "FB923C", "7C2D12"), ("·", "F0FDF4", "86EFAC")):
        ca.conditional_formatting.add(rng, FormulaRule(formula=[f'C6="{letter}"'], fill=fill(bg), font=Font(name=F, bold=True, color=fg, size=8)))
    ca.conditional_formatting.add(f"C5:{last}5", FormulaRule(formula=[f'AND(TODAY()>=C$5,TODAY()<C$5+7)'], fill=fill("F59E0B")))
    ca.conditional_formatting.add(rng, FormulaRule(formula=[f'AND(ISNUMBER(Settings!$C$5),Settings!$C$5>=C$5,Settings!$C$5<C$5+7,C6="")'], fill=fill("DBEAFE")))
    ca.conditional_formatting.add(rng, FormulaRule(formula=[f'AND(ISNUMBER(Settings!$C$6),Settings!$C$6>=C$5,Settings!$C$6<C$5+7,C6="")'], fill=fill("DBEAFE")))
    ca.freeze_panes = "C6"
    ca["A4"] = "Blue column = frost-date week · orange header = this week"; ca["A4"].font = font(size=8, italic=True, color="5B6770")

    # ---------- Harvest Log ----------
    K.title(hv, "Harvest Log", "Every picking: date, crop, amount. Add what it would cost in a shop to see what your garden saves you.")
    K.hdr(hv, 5, ["Date", "Crop", "Bed / area", "Amount", "Unit", "Shop price per unit", "Value", "Note", "Year", "Month"],
          [12, 22, 14, 9, 8, 12, 11, 26, 7, 7])
    style_range(hv, H0, H1, 1, 10, inputs=range(1, 9), fmt={1: DATE, 4: "0.0", 6: MONEY, 7: MONEY, 9: "0", 10: "0"})
    for r in range(H0, H1 + 1):
        hv[f"G{r}"] = f'=IF(A{r}="","",N(D{r})*N(F{r}))'
        hv[f"I{r}"] = f'=IF(ISNUMBER(A{r}),YEAR(A{r}),"")'; hv[f"J{r}"] = f'=IF(ISNUMBER(A{r}),MONTH(A{r}),"")'
    dv = DataValidation(type="list", formula1=f"={LB('A')}", allow_blank=True); hv.add_data_validation(dv); dv.add(f"B{H0}:B{H1}")
    dv = DataValidation(type="list", formula1=f"=Settings!$F${B0}:$F${B1}", allow_blank=True); hv.add_data_validation(dv); dv.add(f"C{H0}:C{H1}")
    dv = DataValidation(type="list", formula1=f'"{UNITS}"', allow_blank=True); hv.add_data_validation(dv); dv.add(f"E{H0}:E{H1}")
    hv.freeze_panes = f"A{H0}"; hv.auto_filter.ref = f"A5:J{H1}"
    HV = lambda c: f"'Harvest Log'!${c}${H0}:${c}${H1}"

    # ---------- Expenses ----------
    K.title(ex, "Expenses", "Seeds, plants, soil, tools… everything you spend on the garden.")
    K.hdr(ex, 5, ["Date", "Category", "What", "Shop", "Amount", "Year", "Month"], [12, 20, 32, 20, 12, 7, 7])
    style_range(ex, E0, E1, 1, 7, inputs=range(1, 6), fmt={1: DATE, 5: MONEY, 6: "0", 7: "0"})
    for r in range(E0, E1 + 1):
        ex[f"F{r}"] = f'=IF(ISNUMBER(A{r}),YEAR(A{r}),"")'; ex[f"G{r}"] = f'=IF(ISNUMBER(A{r}),MONTH(A{r}),"")'
    dv = DataValidation(type="list", formula1='"' + ",".join(ECATS) + '"', allow_blank=True); ex.add_data_validation(dv); dv.add(f"B{E0}:B{E1}")
    ex.freeze_panes = f"A{E0}"; ex.auto_filter.ref = f"A5:G{E1}"
    EX = lambda c: f"Expenses!${c}${E0}:${c}${E1}"

    # ---------- Season Summary ----------
    K.title(su, "Season Summary", "Harvest and spending for the season year (from Settings).")
    su["A3"] = '="Season "&Settings!$C$12&"  ·  amounts in "&Settings!$C$7&"  ·  weight in "&Settings!$C$8'; su["A3"].font = font(bold=True)
    MN = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    K.hdr(su, 5, ["Month", "Harvest value", "Harvest weight", "Pickings", "Spent", "Value minus spent"], [24, 13, 13, 10, 12, 15])
    for j in range(12):
        r = 6 + j; su[f"A{r}"] = MN[j]
        su[f"B{r}"] = f'=SUMIFS({HV("G")},{HV("I")},{YR},{HV("J")},{j + 1})'
        su[f"C{r}"] = f'=SUMIFS({HV("D")},{HV("I")},{YR},{HV("J")},{j + 1},{HV("E")},{WU})'
        su[f"D{r}"] = f'=COUNTIFS({HV("I")},{YR},{HV("J")},{j + 1})'
        su[f"E{r}"] = f'=SUMIFS({EX("E")},{EX("F")},{YR},{EX("G")},{j + 1})'
        su[f"F{r}"] = f"=B{r}-E{r}"
    su["A18"] = "Season total"
    for c in "BCDEF": su[f"{c}18"] = f"=SUM({c}6:{c}17)"
    style_range(su, 6, 18, 1, 6, fmt={2: MONEY, 3: "0.0", 4: "0", 5: MONEY, 6: MONEY})
    for c in range(1, 7): su.cell(row=18, column=c).font = Font(name=F, bold=True, color=MAIN); su.cell(row=18, column=c).fill = fill(SOFT)
    C0 = 22
    K.section(su, f"A{C0 - 2}", "Per crop (all crops in the Plant Library)")
    K.hdr(su, C0 - 1, ["Crop", "Harvest value", "Harvest weight", "Pickings", "Plantings", "(rank key)"])
    for k in range(NL):
        r = C0 + k; lr = L0 + k
        su[f"A{r}"] = f"=IF('Plant Library'!A{lr}=\"\",\"\",'Plant Library'!A{lr})"
        su[f"B{r}"] = f'=IF(A{r}="",0,SUMIFS({HV("G")},{HV("B")},A{r},{HV("I")},{YR}))'
        su[f"C{r}"] = f'=IF(A{r}="",0,SUMIFS({HV("D")},{HV("B")},A{r},{HV("I")},{YR},{HV("E")},{WU}))'
        su[f"D{r}"] = f'=IF(A{r}="",0,COUNTIFS({HV("B")},A{r},{HV("I")},{YR}))'
        su[f"E{r}"] = f'=IF(A{r}="",0,COUNTIFS({PL("A")},A{r}))'
        su[f"F{r}"] = f'=IF(OR(A{r}="",B{r}+C{r}=0),"",B{r}+C{r}/1000+ROW()/10000000)'; su[f"F{r}"].font = font(color="9CA3AF", size=7)
    style_range(su, C0, C0 + NL - 1, 1, 5, fmt={2: MONEY, 3: "0.0", 4: "0", 5: "0"})
    su.conditional_formatting.add(f"B{C0}:B{C0 + NL - 1}", DataBarRule(start_type="num", start_value=0, end_type="max", color="86EFAC"))
    su.freeze_panes = "B6"
    KEYC = f"'Season Summary'!$F${C0}:$F${C0 + NL - 1}"

    # ---------- sample ----------
    if sample:
        rnd = random.Random(7)
        s["C5"] = dt.date(2026, 4, 20); s["C6"] = dt.date(2026, 10, 15); s["C9"] = "Zone 6b (sample)"
        beds = [("Bed 1", 8, 4), ("Bed 2", 8, 4), ("Bed 3", 8, 4), ("Bed 4", 8, 4), ("Herb bed", 6, 3), ("Squash patch", 10, 6), ("Pots on patio", 4, 2)]
        for k, (n, l, w) in enumerate(beds):
            s[f"F{B0 + k}"] = n; s[f"G{B0 + k}"] = l; s[f"H{B0 + k}"] = w
        D = dt.date
        # crop, variety, bed, method, qty, shift, own date, sown indoors, planted, first harvest, finished, notes
        plan = [
            ("Onion (from seed)", "Walla Walla", "Bed 1", "Start indoors", 24, None, None, D(2026, 2, 10), D(2026, 3, 25), D(2026, 7, 10), "Yes", ""),
            ("Peas", "Sugar Snap", "Bed 1", "Direct sow", 30, None, None, None, D(2026, 3, 12), D(2026, 5, 15), "Yes", "Trellis on north side"),
            ("Lettuce", "Buttercrunch", "Bed 2", "Start indoors", 8, None, None, D(2026, 3, 9), D(2026, 4, 7), D(2026, 5, 20), "Yes", ""),
            ("Lettuce", "Buttercrunch", "Bed 2", "Direct sow", 8, 3, None, None, D(2026, 4, 15), D(2026, 6, 1), "Yes", "Succession 2"),
            ("Radish", "French Breakfast", "Bed 2", "Direct sow", 30, None, None, None, D(2026, 3, 24), D(2026, 4, 22), "Yes", ""),
            ("Spinach", "Bloomsdale", "Bed 2", "Direct sow", 20, None, None, None, D(2026, 3, 10), D(2026, 4, 22), "Yes", ""),
            ("Broccoli", "Calabrese", "Bed 3", "Start indoors", 6, None, None, D(2026, 3, 3), D(2026, 4, 8), D(2026, 6, 12), "Yes", ""),
            ("Carrot", "Danvers", "Bed 3", "Direct sow", 60, None, None, None, D(2026, 3, 31), D(2026, 6, 15), "Yes", ""),
            ("Tomato", "Sungold", "Bed 4", "Start indoors", 2, None, None, D(2026, 3, 2), D(2026, 4, 30), D(2026, 7, 6), "Yes", ""),
            ("Tomato", "Brandywine", "Bed 4", "Start indoors", 3, None, None, D(2026, 3, 2), D(2026, 4, 30), D(2026, 7, 25), "Yes", "Stake"),
            ("Pepper (sweet)", "California Wonder", "Bed 4", "Start indoors", 4, None, None, D(2026, 2, 17), D(2026, 5, 6), D(2026, 7, 22), "Yes", ""),
            ("Basil", "Genovese", "Herb bed", "Start indoors", 6, None, None, D(2026, 3, 9), D(2026, 4, 30), D(2026, 6, 25), "Yes", ""),
            ("Parsley", "Italian flat-leaf", "Herb bed", "Buy transplants", 3, None, None, None, D(2026, 4, 10), D(2026, 6, 1), None, ""),
            ("Bush beans", "Provider", "Bed 1", "Direct sow", 24, None, None, None, D(2026, 7, 15), D(2026, 9, 5), "Yes", "After onions"),
            ("Zucchini / summer squash", "Black Beauty", "Squash patch", "Direct sow", 3, None, None, None, D(2026, 5, 1), D(2026, 6, 22), "Yes", ""),
            ("Winter squash", "Butternut", "Squash patch", "Direct sow", 3, None, None, None, D(2026, 5, 6), D(2026, 9, 10), "Yes", ""),
            ("Cucumber", "Marketmore", "Bed 3", "Direct sow", 4, None, None, None, D(2026, 6, 20), D(2026, 8, 10), "Yes", "After broccoli"),
            ("Fall kale", "Lacinato", "Bed 2", "Start indoors", 6, None, None, D(2026, 7, 8), D(2026, 8, 5), D(2026, 9, 30), None, ""),
            ("Fall spinach", "Space", "Bed 2", "Direct sow", 30, None, None, None, D(2026, 9, 3), D(2026, 10, 15), None, "Cover with row cover"),
            ("Garlic", "Music (hardneck)", "Bed 4", "Direct sow", 32, None, None, None, D(2026, 11, 5), None, None, "Plant cloves pointy end up"),
            ("Strawberry", "Albion", "Pots on patio", "Buy transplants", 4, None, None, None, D(2026, 4, 1), D(2026, 6, 10), "Yes", ""),
        ]
        for k, row in enumerate(plan):
            for c, v in zip("ABCDEFGHIJKL", row):
                if v not in (None, ""): pl[f"{c}{P0 + k}"] = v
        harv = []
        def add(crop, bed, d0, d1, n, amt, unit, price):
            for i in range(n):
                d = d0 + dt.timedelta(days=int((d1 - d0).days * i / max(1, n - 1)))
                harv.append((d, crop, bed, round(amt * rnd.uniform(0.7, 1.3), 1) if unit in ("lb", "kg") else max(1, round(amt * rnd.uniform(0.7, 1.3))), unit, price))
        add("Radish", "Bed 2", D(2026, 4, 22), D(2026, 5, 6), 3, 12, "count", 0.15)
        add("Spinach", "Bed 2", D(2026, 4, 22), D(2026, 5, 20), 4, 0.5, "lb", 6.00)
        add("Peas", "Bed 1", D(2026, 5, 15), D(2026, 6, 20), 6, 0.8, "lb", 5.00)
        add("Lettuce", "Bed 2", D(2026, 5, 20), D(2026, 6, 30), 6, 2, "head", 2.50)
        add("Broccoli", "Bed 3", D(2026, 6, 12), D(2026, 7, 5), 4, 1.0, "lb", 3.00)
        add("Carrot", "Bed 3", D(2026, 6, 15), D(2026, 7, 20), 3, 1.5, "lb", 1.80)
        add("Zucchini / summer squash", "Squash patch", D(2026, 6, 22), D(2026, 8, 31), 12, 2.0, "lb", 2.00)
        add("Basil", "Herb bed", D(2026, 6, 25), D(2026, 9, 15), 7, 2, "bunch", 3.00)
        add("Tomato", "Bed 4", D(2026, 7, 6), D(2026, 9, 25), 14, 3.0, "lb", 3.50)
        add("Onion (from seed)", "Bed 1", D(2026, 7, 10), D(2026, 7, 20), 2, 6, "lb", 1.50)
        add("Pepper (sweet)", "Bed 4", D(2026, 7, 22), D(2026, 9, 30), 7, 1.2, "lb", 4.00)
        add("Cucumber", "Bed 3", D(2026, 8, 10), D(2026, 9, 15), 6, 1.5, "lb", 2.00)
        add("Bush beans", "Bed 1", D(2026, 9, 5), D(2026, 9, 30), 4, 1.0, "lb", 4.00)
        add("Winter squash", "Squash patch", D(2026, 9, 25), D(2026, 9, 25), 1, 14, "lb", 1.50)
        add("Strawberry", "Pots on patio", D(2026, 6, 10), D(2026, 9, 20), 8, 0.4, "lb", 5.00)
        add("Fall kale", "Bed 2", D(2026, 9, 30), D(2026, 11, 12), 6, 2, "bunch", 3.00)
        add("Fall spinach", "Bed 2", D(2026, 10, 15), D(2026, 11, 10), 3, 0.5, "lb", 6.00)
        harv.sort(key=lambda h: h[0])
        for k, h in enumerate(harv):
            for c, v in zip("ABCDEF", h): hv[f"{c}{H0 + k}"] = v
        exps = [(D(2026, 1, 18), "Seeds", "Seed order – 14 packets", "Johnny's Seeds", 52.40),
                (D(2026, 2, 5), "Tools & equipment", "Grow light + seed trays", "Amazon", 64.99),
                (D(2026, 2, 5), "Soil & compost", "Seed-starting mix", "Garden center", 12.99),
                (D(2026, 3, 14), "Soil & compost", "Compost, 10 bags", "Garden center", 59.90),
                (D(2026, 3, 28), "Plants & seedlings", "Parsley plants + strawberries", "Garden center", 26.50),
                (D(2026, 4, 2), "Pots & raised beds", "Patio pots", "Home Depot", 34.00),
                (D(2026, 5, 2), "Mulch", "Straw bale", "Farm store", 9.99),
                (D(2026, 5, 20), "Fertilizer", "Organic tomato feed", "Garden center", 16.49),
                (D(2026, 6, 15), "Pest control", "Row cover + neem oil", "Amazon", 27.80),
                (D(2026, 7, 1), "Water", "Extra water bill (estimate)", "City water", 18.00),
                (D(2026, 8, 1), "Water", "Extra water bill (estimate)", "City water", 22.00),
                (D(2026, 8, 20), "Seeds", "Fall seeds", "Local nursery", 11.50),
                (D(2026, 10, 2), "Seeds", "Seed garlic, 1 lb", "Filaree Garlic Farm", 24.00)]
        for k, e in enumerate(exps):
            for c, v in zip("ABCDE", e): ex[f"{c}{E0 + k}"] = v

    # ---------- Dashboard ----------
    K.title(dash, "Garden Dashboard", "Updates by itself from Settings, Plantings, Harvest Log and Expenses.")
    dash["A3"] = '="Season "&Settings!$C$12&"  ·  amounts in "&Settings!$C$7&"  ·  today: "&TEXT(TODAY(),"mmm d, yyyy")'; dash["A3"].font = font(bold=True)
    kp = [("Last spring frost", f'=IF(ISNUMBER({SPR}),{SPR},"set in Settings")', "mmm d, yyyy"),
          ("Days until last spring frost", f'=IF(ISNUMBER({SPR}),IF({SPR}>=TODAY(),{SPR}-TODAY(),"passed"),"")', "0"),
          ("First fall frost", f'=IF(ISNUMBER({FAL}),{FAL},"set in Settings")', "mmm d, yyyy"),
          ("Plantings in plan", f'=COUNTA({PL("A")})', "0"),
          ("Tasks late", f'=COUNTIFS({PL("X")},"Late by*")', "0"),
          ("Tasks due soon", f'=COUNTIFS({PL("X")},"Today")+COUNTIFS({PL("X")},"In *")', "0"),
          ("Harvest value this season", "='Season Summary'!$B$18", MONEY),
          ("Spent this season", "='Season Summary'!$E$18", MONEY),
          ("Garden value minus spending", "='Season Summary'!$F$18", MONEY)]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 3) * 3, 1 + (i % 3) * 2, lab, f, fmt)
        dash.cell(row=6 + (i // 3) * 3, column=1 + (i % 3) * 2).alignment = Alignment(horizontal="left")
    K.kpi(dash, 14, 1, "Harvest weight", "='Season Summary'!$C$18", "0.0"); dash["A14"] = '="Harvest weight ("&Settings!$C$8&")"'
    K.kpi(dash, 14, 3, "Pickings logged", "='Season Summary'!$D$18", "0")
    for c in ("A15", "C15"): dash[c].alignment = Alignment(horizontal="left")
    for c, w in zip("ABCDEFGH", [26, 24, 24, 15, 22, 15, 3, 3]): dash.column_dimensions[c].width = w
    dash.conditional_formatting.add("C9", FormulaRule(formula=["C9>0"], font=Font(name=F, size=18, bold=True, color="B91C1C")))
    dash.conditional_formatting.add("E12", FormulaRule(formula=["E12<0"], font=Font(name=F, size=18, bold=True, color="B91C1C")))
    # to-do list
    H = 18
    K.section(dash, f"A{H - 1}", "Garden to-do list (next tasks, by date)")
    K.hdr(dash, H, ["Date", "Task", "Crop", "Variety", "Bed / area", "When"])
    KEY = PL("Y")
    for k in range(12):
        r = H + 1 + k
        dash[f"H{r}"] = f'=IF(COUNT({KEY})<{k + 1},"",SMALL({KEY},{k + 1}))'; dash[f"H{r}"].font = font(color="FFFFFF", size=6)
        m = f"MATCH($H{r},{KEY},0)"
        dash[f"A{r}"] = f'=IF($H{r}="","",INT($H{r}))'
        for col, src in zip("BCDEF", "UABCX"):
            dash[f"{col}{r}"] = f'=IF($H{r}="","",INDEX({PL(src)},{m})&"")'
    style_range(dash, H + 1, H + 12, 1, 6, fmt={1: "mmm d, yyyy"})
    dash.conditional_formatting.add(f"F{H + 1}:F{H + 12}", FormulaRule(formula=[f'LEFT(F{H + 1},7)="Late by"'], **red))
    dash.conditional_formatting.add(f"F{H + 1}:F{H + 12}", FormulaRule(formula=[f'OR(F{H + 1}="Today",LEFT(F{H + 1},2)="In")'], **amb))
    # top crops
    T = H + 16
    K.section(dash, f"A{T - 1}", "Top crops this season (by harvest value)")
    K.hdr(dash, T, ["Crop", "Harvest value", '="Weight ("&Settings!$C$8&")"', "Pickings"])
    for k in range(10):
        r = T + 1 + k
        dash[f"H{r}"] = f'=IF(COUNT({KEYC})<{k + 1},"",LARGE({KEYC},{k + 1}))'; dash[f"H{r}"].font = font(color="FFFFFF", size=6)
        m = f"MATCH($H{r},{KEYC},0)"
        dash[f"A{r}"] = f"=IF($H{r}=\"\",\"\",INDEX('Season Summary'!$A${C0}:$A${C0 + NL - 1},{m}))"
        for col, src in zip("BCD", "BCD"):
            dash[f"{col}{r}"] = f"=IF($H{r}=\"\",\"\",INDEX('Season Summary'!${src}${C0}:${src}${C0 + NL - 1},{m}))"
    style_range(dash, T + 1, T + 10, 1, 4, fmt={2: MONEY, 3: "0.0", 4: "0"})
    dash.conditional_formatting.add(f"B{T + 1}:B{T + 10}", DataBarRule(start_type="num", start_value=0, end_type="max", color="86EFAC"))
    # charts
    ch = BarChart(); ch.type = "col"; ch.title = "Harvest value vs spending by month"; ch.height = 7.5; ch.width = 17
    ch.add_data(Reference(su, min_col=2, max_col=2, min_row=5, max_row=17), titles_from_data=True)
    ch.add_data(Reference(su, min_col=5, max_col=5, min_row=5, max_row=17), titles_from_data=True)
    ch.set_categories(Reference(su, min_col=1, min_row=6, max_row=17)); ch.y_axis.majorGridlines = None
    ch.series[0].graphicalProperties.solidFill = "22C55E"; ch.series[1].graphicalProperties.solidFill = "F59E0B"
    ch.gapWidth = 50; ch.y_axis.scaling.min = 0; ch.legend.position = "b"; ch.legend.overlay = False
    dash.add_chart(ch, "I4")
    ch2 = BarChart(); ch2.type = "bar"; ch2.title = "Bed space: planned vs size (sq ft)"; ch2.height = 8; ch2.width = 17
    ch2.add_data(Reference(s, min_col=9, max_col=10, min_row=5, max_row=B0 + 7), titles_from_data=True)
    ch2.set_categories(Reference(s, min_col=6, min_row=B0, max_row=B0 + 7)); ch2.y_axis.majorGridlines = None
    ch2.x_axis.scaling.orientation = "maxMin"; ch2.series[0].graphicalProperties.solidFill = "BBF7D0"; ch2.series[1].graphicalProperties.solidFill = MAIN
    ch2.gapWidth = 40; ch2.legend.position = "b"; ch2.legend.overlay = False
    dash.add_chart(ch2, "I21")
    for c in (ch, ch2): c.roundedCorners = False; c.x_axis.delete = False; c.y_axis.delete = False

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Settings: type your last spring frost date and first fall frost date (search 'frost dates' + your ZIP code or town). Add your beds or areas with their size.", False),
             ("2.  Plant Library: 43 common vegetables, herbs and flowers are ready, with timing in weeks before (-) or after (+) your frost date, days to harvest and spacing. Adjust anything for your climate or add your own crops.", False),
             ("3.  Plantings: one row per crop – pick the crop, bed and method (Start indoors / Direct sow / Buy transplants) and how many plants.", False),
             ("     The start-indoors date, plant-out or sowing date, harvest window and space needed are calculated. Use 'Shift' for succession sowings (e.g. 2, 4, 6 weeks later) or type your own date.", False),
             ("4.  As you go, type the real dates (seeds started, planted out, first harvest) and choose 'Yes' under Finished when the crop is done. The status and next task update.", False),
             ("5.  Calendar: a week-by-week picture of the whole season – S start indoors, T plant out, D direct sow, H harvest.", False),
             ("6.  Harvest Log and Expenses: record pickings (with a shop price to see the value) and garden spending.", False),
             ("7.  Dashboard: days to frost, your to-do list sorted by date with late tasks in red, harvest value vs spending, top crops and bed space.", False),
             ("", False), ("Tips", True),
             ("•  Frost dates are averages – keep an eye on the weather forecast before planting out tender crops like tomatoes, peppers and squash.", False),
             ("•  Timings in the library are general guidelines; seed packets and your local extension office know your area best.", False),
             ("•  For a new season, change the frost dates in Settings, clear the yellow dates in Plantings and keep your plan.", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  150 plantings, 70 library crops, 20 beds, 1,000 harvest entries, 400 expenses. Spacing in inches, bed sizes in feet. No macros. Works in Excel 2010+ and Google Sheets.", False),
             ("•  To use Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("A made-up 2026 season in zone 6b with 7 beds, 21 plantings, harvests and expenses. Use the blank copy in your download for your own garden.", False)]
    K.start(st, "Garden Planner & Planting Calendar", "Frost-date planting schedule, week-by-week calendar, harvest log & garden budget – Excel and Google Sheets.", lines)
    _orig_save = wb.save

    def _save(p):
        st.page_setup.fitToHeight = 1
        _orig_save(p)
    wb.save = _save
    finish(wb, sample, path)


META = {
    "name": "Garden Planner & Planting Calendar (Frost Date Sowing Schedule, Harvest Log, Garden Budget) – Excel & Google Sheets",
    "file_base": "Garden-Planner-Planting-Calendar",
    "price": 14,
    "color": "#15803D",
    "title": "Garden Planner & Planting Calendar",
    "tagline": "frost-date sowing schedule · week-by-week calendar · harvest log · budget",
    "intro": "Plan your vegetable garden around your own frost dates. Type your last spring and first fall frost date once, pick your crops, and the planner works out when to start seeds indoors, when to plant out or sow, and when to expect a harvest – then shows the whole season on a week-by-week calendar and keeps a to-do list of what is due next.",
    "features": [
        "Plant library with 43 vegetables, herbs and flowers: start-indoors, transplant and direct-sow timing, days to harvest, spacing – fully editable",
        "Plantings plan (150 rows): crop, variety, bed, method and quantity → start date, plant-out/sow date, harvest window and space needed",
        "Succession sowing: shift any planting by a few weeks, or type your own date",
        "Week-by-week planting calendar for the whole year with colour codes (start indoors, plant out, direct sow, harvest)",
        "To-do list on the dashboard sorted by date, with late tasks in red",
        "Bed planner: size of each bed vs space planned",
        "Harvest log with shop-price value, plus a garden expense tracker – see what your garden really saves you",
        "Season summary per month and per crop, top crops and charts",
        "Blank copy plus a sample copy with a full season",
    ],
    "tags": ["garden planner", "planting calendar", "vegetable garden planner", "seed starting schedule", "frost date calculator",
             "garden spreadsheet", "harvest tracker", "google sheets", "excel template", "garden journal", "homestead planner", "succession planting"],
    "summary": "Instant download: Excel + Google Sheets template, blank copy + sample copy.",
}
