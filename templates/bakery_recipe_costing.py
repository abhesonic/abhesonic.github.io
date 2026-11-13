"""Bakery Recipe Cost & Pricing Calculator – ingredient costs with unit conversion, recipe costing, suggested retail &
wholesale prices, recipe card scaler, order tracker and profit dashboard.

Settings (+ units) -> Ingredients -> Recipes -> Recipe Lines -> Recipe Card -> Orders -> 12 Months -> Dashboard.
"""
import datetime as dt, random
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, DataBarRule
from openpyxl.chart import BarChart, Reference
from _common import Kit, F, B, MONEY, DATE, font, fill, style_range, input_cell, finish

MAIN = "9D174D"; SOFT = "FCE7F3"
K = Kit(MAIN, SOFT)
NU = 25; U0 = 6; U1 = U0 + NU - 1            # units (Settings F:H)
NI = 150; I0 = 6; I1 = I0 + NI - 1           # ingredients
NR = 60; R0_ = 6; R1_ = R0_ + NR - 1         # recipes
NL = 1000; L0 = 6; L1 = L0 + NL - 1          # recipe lines
NO = 1000; O0 = 6; O1 = O0 + NO - 1          # orders
NC = 40                                      # recipe card lines
PCT = "0%"; QTY = '#,##0.##;-#,##0.##;"-"'; MONEY4 = '#,##0.0000;-#,##0.0000;"-"'
STATUS = ["Quote", "Confirmed", "Paid", "Delivered", "Cancelled"]
UNITS = [("g", "Mass", 1), ("kg", "Mass", 1000), ("oz", "Mass", 28.3495), ("lb", "Mass", 453.592),
         ("ml", "Volume", 1), ("l", "Volume", 1000), ("tsp", "Volume", 4.92892), ("tbsp", "Volume", 14.7868),
         ("cup", "Volume", 236.588), ("fl oz", "Volume", 29.5735), ("pint", "Volume", 473.176), ("quart", "Volume", 946.353),
         ("gallon", "Volume", 3785.41), ("each", "Count", 1), ("dozen", "Count", 12)]


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); ing = wb.create_sheet("Ingredients"); rec = wb.create_sheet("Recipes")
    ln = wb.create_sheet("Recipe Lines"); card = wb.create_sheet("Recipe Card"); od = wb.create_sheet("Orders")
    ym = wb.create_sheet("12 Months"); s = wb.create_sheet("Settings")

    # ---------- Settings ----------
    K.title(s, "Settings", "Yellow cells are yours to edit. These numbers drive every price in the workbook.")
    SET = [("Currency (label only)", "USD", None), ("Your hourly wage (pay yourself!)", 18, MONEY),
           ("Monthly overhead (kitchen, power, insurance, website, fees)", 250, MONEY),
           ("Batches you bake per month", 40, "0"), ("Overhead per batch", None, MONEY),
           ("Target profit margin (share of the price that is profit)", 0.35, PCT),
           ("Wholesale price as % of your retail price", 0.6, PCT), ("Round prices up to the nearest", 0.25, MONEY),
           ("12 Months tab starts (blank = January this year)", None, "mmm yyyy")]
    for i, (lab, v, fmt) in enumerate(SET, 5):
        s[f"A{i}"] = lab; s[f"A{i}"].font = font(bold=True); s[f"A{i}"].border = B
        if lab == "Overhead per batch":
            s[f"C{i}"] = "=IF(N(C8)<=0,0,N(C7)/C8)"; s[f"C{i}"].fill = fill("F3F4F6"); s[f"C{i}"].border = B; s[f"C{i}"].font = font()
        else:
            input_cell(s[f"C{i}"], v)
        if fmt: s[f"C{i}"].number_format = fmt
    s["A14"] = "12 Months starts"; s["A14"].font = font(italic=True)
    s["C14"] = '=IF(ISNUMBER(C13),DATE(YEAR(C13),MONTH(C13),1),DATE(YEAR(TODAY()),1,1))'; s["C14"].number_format = "mmm yyyy"
    s["A16"] = "Tip: margin 35% means a $10 price holds $3.50 profit after ingredients, packaging, your wage and overhead."
    s["A16"].font = font(italic=True, size=9)
    K.hdr(s, 5, ["Unit", "Type", "Size in g / ml / each"], col0=6)
    style_range(s, U0, U1, 6, 8, inputs=(6, 7, 8), fmt={8: '#,##0.#####'})
    for i, (u, t, f) in enumerate(UNITS):
        for c, v in zip("FGH", (u, t, f)): s[f"{c}{U0 + i}"] = v
    for r in range(U0, U1 + 1):
        for c in "FGH": s[f"{c}{r}"].font = Font(name=F, color="0000FF")
    dv = DataValidation(type="list", formula1='"Mass,Volume,Count"', allow_blank=True); s.add_data_validation(dv); dv.add(f"G{U0}:G{U1}")
    s[f"F{U1 + 1}"] = "Add your own units in the empty rows (e.g. 'stick butter', Mass, 113)."; s[f"F{U1 + 1}"].font = font(italic=True, size=9)
    for c, w in zip("ABCDEFGH", [56, 3, 14, 3, 3, 14, 10, 18]): s.column_dimensions[c].width = w
    RATE = "Settings!$C$6"; OHB = "Settings!$C$9"; MARG = "Settings!$C$10"; WPCT = "Settings!$C$11"; RND = "Settings!$C$12"
    UN = f"Settings!$F${U0}:$F${U1}"; UT = f"Settings!$G${U0}:$G${U1}"; UF = f"Settings!$H${U0}:$H${U1}"
    rnd = lambda x: f"IF(N({RND})<=0,{x},CEILING({x},{RND}))"

    # ---------- Ingredients ----------
    K.title(ing, "Ingredients", "What you pay at the store. Cost per g / ml / each is calculated for you.")
    K.hdr(ing, 5, ["Ingredient", "Store / supplier", "Package size", "Package unit", "Package price", "Grams per cup (optional)",
                   "Cost per base unit", "Base unit", "Notes"], [24, 16, 10, 10, 11, 14, 12, 9, 30])
    style_range(ing, I0, I1, 1, 9, inputs=(1, 2, 3, 4, 5, 6, 9), fmt={3: QTY, 5: MONEY, 6: QTY, 7: MONEY4})
    for r in range(I0, I1 + 1):
        m = f"MATCH(D{r},{UN},0)"
        ing[f"G{r}"] = f'=IF(OR(A{r}="",N(C{r})<=0,N(E{r})<=0,ISNA({m})),"",E{r}/(C{r}*INDEX({UF},{m})))'
        ing[f"H{r}"] = f'=IF(OR(A{r}="",ISNA({m})),"",IF(INDEX({UT},{m})="Mass","g",IF(INDEX({UT},{m})="Volume","ml","each")))'
        ing[f"J{r}"] = f'=IF(OR(A{r}="",ISNA({m})),"",INDEX({UT},{m}))'
    ing.column_dimensions["J"].hidden = True
    dv = DataValidation(type="list", formula1=f"={UN}", allow_blank=True); ing.add_data_validation(dv); dv.add(f"D{I0}:D{I1}")
    ing.freeze_panes = f"B{I0}"; ing.auto_filter.ref = f"A5:I{I1}"
    IN = f"Ingredients!$A${I0}:$A${I1}"; IC = f"Ingredients!$G${I0}:$G${I1}"; IT = f"Ingredients!$J${I0}:$J${I1}"; IG = f"Ingredients!$F${I0}:$F${I1}"

    # ---------- Recipe Lines ----------
    K.title(ln, "Recipe Lines", "One row per ingredient in a recipe. Quantity is for ONE batch (the yield on the Recipes tab).")
    K.hdr(ln, 5, ["Recipe", "Ingredient", "Quantity", "Unit", "Cost", "Check", "Note"], [24, 24, 10, 9, 11, 34, 24])
    style_range(ln, L0, L1, 1, 7, inputs=(1, 2, 3, 4, 7), fmt={3: QTY, 5: MONEY})
    for r in range(L0, L1 + 1):
        mi = f"MATCH(B{r},{IN},0)"; mu = f"MATCH(D{r},{UN},0)"
        ln[f"H{r}"] = f'=IF(OR(B{r}="",ISNA({mu})),"",INDEX({UT},{mu}))'                       # line unit type
        ln[f"I{r}"] = f'=IF(OR(B{r}="",ISNA({mi})),"",INDEX({IT},{mi}))'                       # ingredient type
        ln[f"J{r}"] = f'=IF(OR(B{r}="",ISNA({mi})),0,N(INDEX({IG},{mi})))'                     # grams per cup
        ln[f"K{r}"] = (f'=IF(OR(H{r}="",I{r}="",NOT(ISNUMBER(C{r}))),"",IF(H{r}=I{r},C{r}*INDEX({UF},{mu}),'
                       f'IF(AND(I{r}="Mass",H{r}="Volume",J{r}>0),C{r}*INDEX({UF},{mu})/236.588*J{r},'
                       f'IF(AND(I{r}="Volume",H{r}="Mass",J{r}>0),C{r}*INDEX({UF},{mu})/J{r}*236.588,""))))')
        ln[f"E{r}"] = f'=IF(OR(K{r}="",B{r}=""),0,IFERROR(K{r}*INDEX({IC},{mi}),0))'
        ln[f"F{r}"] = (f'=IF(AND(A{r}="",B{r}=""),"",IF(ISNA({mi}),"Ingredient not on Ingredients tab",IF(H{r}="","Pick a unit",'
                       f'IF(NOT(ISNUMBER(C{r})),"Enter a quantity",IF(K{r}="","Weight vs cups: add grams per cup on Ingredients",'
                       f'IF(N(E{r})=0,"Ingredient has no price yet","OK"))))))')
        ln[f"L{r}"] = f"=IF(AND(A{r}<>\"\",A{r}='Recipe Card'!$C$4),COUNTIF($A${L0}:A{r},'Recipe Card'!$C$4),\"\")"
    for c in "HIJKL": ln.column_dimensions[c].hidden = True
    dv = DataValidation(type="list", formula1=f"=Recipes!$A${R0_}:$A${R1_}", allow_blank=True); ln.add_data_validation(dv); dv.add(f"A{L0}:A{L1}")
    dv = DataValidation(type="list", formula1=f"={IN}", allow_blank=True); ln.add_data_validation(dv); dv.add(f"B{L0}:B{L1}")
    dv = DataValidation(type="list", formula1=f"={UN}", allow_blank=True); ln.add_data_validation(dv); dv.add(f"D{L0}:D{L1}")
    ln.conditional_formatting.add(f"F{L0}:F{L1}", FormulaRule(formula=[f'AND(F{L0}<>"",F{L0}<>"OK")'], font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    ln.conditional_formatting.add(f"F{L0}:F{L1}", FormulaRule(formula=[f'F{L0}="OK"'], font=Font(name=F, color="166534")))
    ln.freeze_panes = f"A{L0}"; ln.auto_filter.ref = f"A5:G{L1}"
    LR = lambda c: f"'Recipe Lines'!${c}${L0}:${c}${L1}"

    # ---------- Recipes ----------
    K.title(rec, "Recipes", "Yellow: your inputs per batch. Grey: true cost per piece and the price you should charge.")
    K.hdr(rec, 5, ["Recipe", "Category", "Yield (pieces per batch)", "Sold as (cookie, slice…)", "Hands-on time per batch (min)",
                   "Packaging per piece", "Your price per piece (optional)", "Ingredients (batch)", "Labour (batch)", "Overhead (batch)",
                   "Packaging (batch)", "Total batch cost", "Cost per piece", "Suggested price per piece", "Wholesale price per piece",
                   "Price used", "Profit per piece", "Margin", "Lines to fix"],
          [24, 12, 9, 12, 11, 10, 11, 11, 10, 10, 10, 11, 10, 11, 11, 10, 10, 8, 9])
    rec.row_dimensions[5].height = 58
    style_range(rec, R0_, R1_, 1, 19, inputs=(1, 2, 3, 4, 5, 6, 7),
                fmt={**{c: MONEY for c in (6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17)}, 18: PCT, 3: QTY, 5: "0", 19: '0;-0;"-"'})
    for r in range(R0_, R1_ + 1):
        e = f'IF(A{r}=""'
        rec[f"H{r}"] = f'={e},"",SUMIFS({LR("E")},{LR("A")},A{r}))'
        rec[f"I{r}"] = f'={e},"",N(E{r})/60*N({RATE}))'
        rec[f"J{r}"] = f'={e},"",N({OHB}))'
        rec[f"K{r}"] = f'={e},"",N(F{r})*N(C{r}))'
        rec[f"L{r}"] = f'={e},"",H{r}+I{r}+J{r}+K{r})'
        rec[f"M{r}"] = f'=IF(OR(A{r}="",N(C{r})<=0),"",L{r}/C{r})'
        rec[f"N{r}"] = f'=IF(M{r}="","",{rnd(f"M{r}/MAX(0.01,1-N({MARG}))")})'
        rec[f"O{r}"] = f'=IF(N{r}="","",{rnd(f"N{r}*N({WPCT})")})'
        rec[f"P{r}"] = f'=IF(M{r}="","",IF(ISNUMBER(G{r}),G{r},N{r}))'
        rec[f"Q{r}"] = f'=IF(P{r}="","",P{r}-M{r})'
        rec[f"R{r}"] = f'=IF(OR(P{r}="",N(P{r})<=0),"",Q{r}/P{r})'
        rec[f"S{r}"] = f'={e},"",COUNTIFS({LR("A")},A{r},{LR("F")},"<>OK"))'
    dv = DataValidation(type="list", formula1='"Cookies,Cakes,Cupcakes,Bread,Pastry,Pies,Bars,Other"', allow_blank=True)
    rec.add_data_validation(dv); dv.add(f"B{R0_}:B{R1_}")
    rec.conditional_formatting.add(f"Q{R0_}:R{R1_}", FormulaRule(formula=[f'AND(ISNUMBER(Q{R0_}),Q{R0_}<0)'], font=Font(name=F, bold=True, color="B91C1C")))
    rec.conditional_formatting.add(f"S{R0_}:S{R1_}", FormulaRule(formula=[f'N(S{R0_})>0'], font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    rec.conditional_formatting.add(f"N{R0_}:N{R1_}", FormulaRule(formula=[f'ISNUMBER(N{R0_})'], font=Font(name=F, bold=True, color=MAIN)))
    rec.freeze_panes = f"B{R0_}"
    RC = lambda c: f"Recipes!${c}${R0_}:${c}${R1_}"

    # ---------- Recipe Card (scaler) ----------
    K.title(card, "Recipe Card & Scaler", "Pick a recipe and how many pieces you need – quantities and cost scale automatically.")
    card["A4"] = "Recipe"; card["A5"] = "Pieces you need (blank = one batch)"
    for a in ("A4", "A5"): card[a].font = font(bold=True)
    input_cell(card["C4"]); input_cell(card["C5"])
    dv = DataValidation(type="list", formula1=f"={RC('A')}", allow_blank=True); card.add_data_validation(dv); dv.add("C4")
    mr = f"MATCH($C$4,{RC('A')},0)"
    info = [("Yield of one batch", f'=IFERROR(INDEX({RC("C")},{mr}),"")', QTY),
            ("Scale factor", '=IF(OR(F4="",N(F4)<=0),"",IF(ISNUMBER(C5),C5/F4,1))', "0.00x"),
            ("Cost per piece", f'=IFERROR(INDEX({RC("M")},{mr}),"")', MONEY),
            ("Suggested price per piece", f'=IFERROR(INDEX({RC("N")},{mr}),"")', MONEY),
            ("Total cost for this order", '=IF(OR(F6="",F5=""),"",F6*F4*F5)', MONEY),
            ("Total at suggested price", '=IF(OR(F7="",F5=""),"",F7*F4*F5)', MONEY)]
    for i, (lab, f, fmt) in enumerate(info):
        r = 4 + i; card[f"E{r}"] = lab; card[f"E{r}"].font = font(bold=True)
        card[f"F{r}"] = f; card[f"F{r}"].number_format = fmt; card[f"F{r}"].font = font(bold=True, color=MAIN); card[f"F{r}"].fill = fill(SOFT); card[f"F{r}"].border = B
    H0 = 12
    K.hdr(card, H0 - 1, ["#", "Ingredient", "Qty (1 batch)", "Unit", "Qty (scaled)", "Cost (scaled)", "Note"])
    for c, w in zip("ABCDEFG", [5, 26, 13, 9, 30, 16, 24]): card.column_dimensions[c].width = w
    for k in range(NC):
        r = H0 + k; m = f"MATCH({k + 1},{LR('L')},0)"
        card[f"A{r}"] = k + 1
        card[f"H{r}"] = f'=IFERROR({m},"")'
        card[f"B{r}"] = f'=IF(H{r}="","",INDEX({LR("B")},H{r}))'
        card[f"C{r}"] = f'=IF(H{r}="","",INDEX({LR("C")},H{r}))'
        card[f"D{r}"] = f'=IF(H{r}="","",INDEX({LR("D")},H{r}))'
        card[f"E{r}"] = f'=IF(OR(H{r}="",$F$5=""),"",C{r}*$F$5)'
        card[f"F{r}"] = f'=IF(OR(H{r}="",$F$5=""),"",INDEX({LR("E")},H{r})*$F$5)'
        card[f"G{r}"] = f'=IF(H{r}="","",INDEX({LR("G")},H{r})&"")'
    card.column_dimensions["H"].hidden = True
    style_range(card, H0, H0 + NC - 1, 1, 7, fmt={3: QTY, 5: QTY, 6: MONEY})
    card.conditional_formatting.add(f"A{H0}:A{H0 + NC - 1}", FormulaRule(formula=[f'$B{H0}=""'], font=Font(name=F, color="D1D5DB")))
    card[f"A{H0 + NC + 1}"] = "Shows the first 40 ingredients of the recipe. Print this tab for the kitchen."; card[f"A{H0 + NC + 1}"].font = font(italic=True, size=9)

    # ---------- Orders ----------
    K.title(od, "Orders", "One row per order line. Leave the price blank to use the recipe's price. Quote and Cancelled are not counted as sales.")
    K.hdr(od, 5, ["Order date", "Customer", "Recipe", "Pieces", "Price per piece (blank = recipe price)", "Delivery / extra fee",
                  "Discount", "Status", "Pickup / delivery date", "Notes", "Price used", "Revenue", "Cost", "Profit"],
          [12, 18, 22, 8, 13, 11, 10, 11, 13, 22, 10, 11, 11, 11])
    od.row_dimensions[5].height = 44
    style_range(od, O0, O1, 1, 14, inputs=range(1, 11), fmt={1: DATE, 4: QTY, 5: MONEY, 6: MONEY, 7: MONEY, 9: DATE, 11: MONEY, 12: MONEY, 13: MONEY, 14: MONEY})
    for r in range(O0, O1 + 1):
        mr = f"MATCH(C{r},{RC('A')},0)"
        od[f"K{r}"] = f'=IF(OR(C{r}="",NOT(ISNUMBER(D{r}))),"",IF(ISNUMBER(E{r}),E{r},IFERROR(N(INDEX({RC("P")},{mr})),0)))'
        od[f"L{r}"] = f'=IF(K{r}="",0,D{r}*K{r}+N(F{r})-N(G{r}))'
        od[f"M{r}"] = f'=IF(K{r}="",0,D{r}*IFERROR(N(INDEX({RC("M")},{mr})),0))'
        od[f"N{r}"] = f"=L{r}-M{r}"
        od[f"O{r}"] = f'=IF(OR(H{r}="Confirmed",H{r}="Paid",H{r}="Delivered"),1,0)'
    od.column_dimensions["O"].hidden = True
    dv = DataValidation(type="list", formula1=f"={RC('A')}", allow_blank=True); od.add_data_validation(dv); dv.add(f"C{O0}:C{O1}")
    dv = DataValidation(type="list", formula1='"' + ",".join(STATUS) + '"', allow_blank=True); od.add_data_validation(dv); dv.add(f"H{O0}:H{O1}")
    od.conditional_formatting.add(f"A{O0}:N{O1}", FormulaRule(formula=[f'$H{O0}="Cancelled"'], font=Font(name=F, color="9CA3AF", strike=True)))
    od.conditional_formatting.add(f"H{O0}:H{O1}", FormulaRule(formula=[f'$H{O0}="Confirmed"'], font=Font(name=F, bold=True, color="B45309")))
    od.conditional_formatting.add(f"N{O0}:N{O1}", FormulaRule(formula=[f'$N{O0}<0'], font=Font(name=F, bold=True, color="B91C1C")))
    od.freeze_panes = f"B{O0}"; od.auto_filter.ref = f"A5:N{O1}"
    OR_ = lambda c: f"Orders!${c}${O0}:${c}${O1}"

    # ---------- 12 Months ----------
    K.title(ym, "12 Months", "Sales per month (Confirmed, Paid and Delivered orders, by order date).")
    K.hdr(ym, 5, ["Month", "Orders", "Pieces", "Revenue", "Cost", "Profit", "Margin"], [12, 10, 10, 13, 13, 13, 9])
    for j in range(12):
        r = 6 + j
        ym[f"A{r}"] = f"=EDATE(Settings!$C$14,{j})"
        cond = f'{OR_("O")},1,{OR_("A")},">="&$A{r},{OR_("A")},"<="&EOMONTH($A{r},0)'
        ym[f"B{r}"] = f"=COUNTIFS({cond})"
        ym[f"C{r}"] = f"=SUMIFS({OR_('D')},{cond})"
        ym[f"D{r}"] = f"=SUMIFS({OR_('L')},{cond})"
        ym[f"E{r}"] = f"=SUMIFS({OR_('M')},{cond})"
        ym[f"F{r}"] = f"=D{r}-E{r}"
        ym[f"G{r}"] = f'=IF(D{r}<=0,"",F{r}/D{r})'
    style_range(ym, 6, 17, 1, 7, fmt={1: "mmm yyyy", 2: "0", 3: QTY, 4: MONEY, 5: MONEY, 6: MONEY, 7: PCT})
    ym["A18"] = "Total"
    for c in "BCDEF": ym[f"{c}18"] = f"=SUM({c}6:{c}17)"
    ym["G18"] = '=IF(D18<=0,"",F18/D18)'
    for c, fmt in zip("ABCDEFG", [None, "0", QTY, MONEY, MONEY, MONEY, PCT]):
        cell = ym[f"{c}18"]; cell.font = font(bold=True, color=MAIN); cell.fill = fill(SOFT); cell.border = B
        if fmt: cell.number_format = fmt

    # ---------- sample ----------
    if sample:
        ings = [("All-purpose flour", "Costco", 25, "lb", 11.99, 125), ("Granulated sugar", "Costco", 10, "lb", 8.49, 200),
                ("Brown sugar", "Walmart", 2, "lb", 2.98, 213), ("Powdered sugar", "Walmart", 2, "lb", 2.64, 120),
                ("Butter, unsalted", "Costco", 4, "lb", 14.99, 227), ("Eggs, large", "Costco", 5, "dozen", 15.99, None),
                ("Whole milk", "Walmart", 1, "gallon", 3.78, None), ("Heavy cream", "Walmart", 1, "quart", 5.97, None),
                ("Cream cheese", "Walmart", 3, "lb", 9.98, 232), ("Vanilla extract", "Costco", 16, "fl oz", 17.99, None),
                ("Baking soda", "Walmart", 1, "lb", 1.12, 220), ("Baking powder", "Walmart", 10, "oz", 2.48, 192),
                ("Salt", "Walmart", 26, "oz", 1.24, 292), ("Chocolate chips", "Costco", 72, "oz", 13.99, 170),
                ("Cocoa powder", "Walmart", 8, "oz", 3.54, 85), ("Instant yeast", "Costco", 2, "lb", 5.49, 136),
                ("Bread flour", "Walmart", 5, "lb", 4.27, 127), ("Pecans", "Costco", 2, "lb", 14.99, 109),
                ("Cinnamon", "Walmart", 2.37, "oz", 1.48, 125), ("Lemons", "Market", 2, "lb", 3.49, None),
                ("Cookie bag + sticker", "Uline", 100, "each", 14.50, None), ("Cake box 10in + board", "Uline", 25, "each", 31.25, None),
                ("Cupcake box (6)", "Amazon", 50, "each", 27.99, None), ("Food coloring gel", "Amazon", 12, "oz", 9.99, None)]
        ings[19] = ("Lemons", "Market", 10, "each", 4.99, None)
        for k, row in enumerate(ings):
            for c, v in zip("ABCDEF", row):
                if v is not None: ing[f"{c}{I0 + k}"] = v
        recipes = [("Chocolate chip cookies", "Cookies", 24, "cookie", 50, 0.15, 2.50),
                   ("Vanilla cupcakes (buttercream)", "Cupcakes", 24, "cupcake", 75, 0.09, 3.50),
                   ("Red velvet cake 8in", "Cakes", 1, "cake", 150, 1.25, 85),
                   ("Cinnamon rolls", "Pastry", 12, "roll", 90, 0.12, None),
                   ("Sourdough-style sandwich loaf", "Bread", 2, "loaf", 45, 0.40, 9),
                   ("Pecan pie bars", "Bars", 16, "bar", 55, 0.10, 3.25),
                   ("Lemon loaf cake", "Cakes", 2, "loaf", 35, 0.40, 14),
                   ("Fudge brownies", "Bars", 16, "brownie", 40, 0.10, 2.75)]
        for k, row in enumerate(recipes):
            for c, v in zip("ABCDEFG", row):
                if v is not None: rec[f"{c}{R0_ + k}"] = v
        lines = {
            "Chocolate chip cookies": [("All-purpose flour", 2.25, "cup"), ("Butter, unsalted", 1, "cup"), ("Granulated sugar", 0.75, "cup"),
                                       ("Brown sugar", 0.75, "cup"), ("Eggs, large", 2, "each"), ("Vanilla extract", 2, "tsp"),
                                       ("Baking soda", 1, "tsp"), ("Salt", 1, "tsp"), ("Chocolate chips", 340, "g"), ("Cookie bag + sticker", 0, "each")],
            "Vanilla cupcakes (buttercream)": [("All-purpose flour", 3, "cup"), ("Granulated sugar", 2, "cup"), ("Butter, unsalted", 2, "cup"),
                                               ("Eggs, large", 4, "each"), ("Whole milk", 1, "cup"), ("Baking powder", 1, "tbsp"),
                                               ("Vanilla extract", 1, "tbsp"), ("Salt", 0.5, "tsp"), ("Powdered sugar", 4, "cup"), ("Heavy cream", 3, "tbsp")],
            "Red velvet cake 8in": [("All-purpose flour", 2.5, "cup"), ("Granulated sugar", 1.5, "cup"), ("Eggs, large", 2, "each"),
                                    ("Cocoa powder", 2, "tbsp"), ("Whole milk", 1, "cup"), ("Butter, unsalted", 1, "cup"),
                                    ("Cream cheese", 450, "g"), ("Powdered sugar", 4, "cup"), ("Food coloring gel", 0.5, "oz"),
                                    ("Vanilla extract", 2, "tsp"), ("Baking soda", 1, "tsp")],
            "Cinnamon rolls": [("Bread flour", 4.5, "cup"), ("Whole milk", 1, "cup"), ("Instant yeast", 2.25, "tsp"), ("Eggs, large", 2, "each"),
                               ("Butter, unsalted", 0.75, "cup"), ("Brown sugar", 1, "cup"), ("Cinnamon", 2, "tbsp"),
                               ("Cream cheese", 115, "g"), ("Powdered sugar", 1.5, "cup"), ("Salt", 1, "tsp")],
            "Sourdough-style sandwich loaf": [("Bread flour", 1000, "g"), ("Instant yeast", 1, "tsp"), ("Salt", 20, "g"),
                                              ("Whole milk", 300, "ml"), ("Butter, unsalted", 60, "g")],
            "Pecan pie bars": [("All-purpose flour", 2, "cup"), ("Butter, unsalted", 1, "cup"), ("Brown sugar", 1.5, "cup"),
                               ("Eggs, large", 4, "each"), ("Pecans", 2, "cup"), ("Vanilla extract", 2, "tsp"), ("Salt", 0.5, "tsp")],
            "Lemon loaf cake": [("All-purpose flour", 3, "cup"), ("Granulated sugar", 2, "cup"), ("Butter, unsalted", 1, "cup"),
                                ("Eggs, large", 6, "each"), ("Lemons", 4, "each"), ("Whole milk", 1, "cup"), ("Powdered sugar", 2, "cup"),
                                ("Baking powder", 2, "tsp")],
            "Fudge brownies": [("Butter, unsalted", 1, "cup"), ("Granulated sugar", 2, "cup"), ("Eggs, large", 4, "each"),
                               ("Cocoa powder", 0.75, "cup"), ("All-purpose flour", 1, "cup"), ("Chocolate chips", 1, "cup"),
                               ("Vanilla extract", 2, "tsp"), ("Salt", 0.5, "tsp")],
        }
        r = L0
        for rn, items in lines.items():
            for (i_, q, u) in items:
                if q == 0: continue
                ln[f"A{r}"] = rn; ln[f"B{r}"] = i_; ln[f"C{r}"] = q; ln[f"D{r}"] = u; r += 1
        card["C4"] = "Chocolate chip cookies"; card["C5"] = 60
        rr = random.Random(11)
        custs = ["Emma R.", "Coffee Corner Cafe", "Jacob M.", "Olivia S.", "Farmers market", "Sophia L.", "St. Mary's school", "Liam K.",
                 "Ava T.", "Office party – Noah", "Mia W.", "Book club", "Lucas P.", "Harper G."]
        qty = {"Chocolate chip cookies": (12, 48), "Vanilla cupcakes (buttercream)": (6, 36), "Red velvet cake 8in": (1, 2),
               "Cinnamon rolls": (6, 24), "Sourdough-style sandwich loaf": (1, 4), "Pecan pie bars": (8, 24),
               "Lemon loaf cake": (1, 3), "Fudge brownies": (8, 32)}
        orders = []
        for mth, n in zip(range(6, 12), (8, 9, 10, 11, 13, 12)):
            for i in range(n):
                d = dt.date(2026, mth, min(28, 1 + int(27 * (i + rr.random()) / n)))
                rn = rr.choice(list(qty)); q = rr.randint(*qty[rn])
                if rn in ("Chocolate chip cookies", "Vanilla cupcakes (buttercream)", "Cinnamon rolls"): q = max(6, q // 6 * 6)
                cust = rr.choice(custs)
                price = None
                if cust in ("Coffee Corner Cafe", "Farmers market"): price = "WH"
                status = "Delivered"
                if mth == 11 and d.day > 10: status = rr.choice(["Confirmed", "Paid", "Quote"])
                orders.append([d, cust, rn, q, price, rr.choice([0, 0, 0, 5, 8]), rr.choice([0, 0, 0, 0, 5]), status, d + dt.timedelta(days=rr.randint(2, 6)), ""])
        orders.append([dt.date(2026, 10, 9), "Liam K.", "Red velvet cake 8in", 1, None, 0, 0, "Cancelled", dt.date(2026, 10, 12), "Party moved"])
        orders.sort(key=lambda o: o[0])
        for k, o in enumerate(orders):
            r = O0 + k
            for c, v in zip("ABCDEFGHIJ", o):
                if c == "E" and v == "WH":
                    od[f"E{r}"] = f'=IFERROR(INDEX({RC("O")},MATCH(C{r},{RC("A")},0)),"")'
                    od[f"J{r}"] = "Wholesale price"; continue
                if v not in (None, "", 0): od[f"{c}{r}"] = v
        s["C13"] = dt.date(2026, 6, 1)

    # ---------- Dashboard ----------
    K.title(dash, "Bakery Profit Dashboard", "Pick a month in the yellow cell – sales, profit and what every recipe really earns.")
    dash["A3"] = "Month to show (blank = this month):"; dash["A3"].font = font(bold=True)
    input_cell(dash["C3"], dt.date(2026, 11, 1) if sample else None); dash["C3"].number_format = "mmm yyyy"
    dash["L1"] = '=IF(ISNUMBER(C3),DATE(YEAR(C3),MONTH(C3),1),DATE(YEAR(TODAY()),MONTH(TODAY()),1))'; dash["L1"].number_format = DATE
    dash["L1"].font = font(color="FFFFFF", size=8)
    dash["E3"] = '="Showing "&TEXT($L$1,"mmmm yyyy")&"  ·  amounts in "&Settings!C5'; dash["E3"].font = font(bold=True, color=MAIN)
    cond = f'{OR_("O")},1,{OR_("A")},">="&$L$1,{OR_("A")},"<="&EOMONTH($L$1,0)'
    kp = [("Orders this month", f"=COUNTIFS({cond})", "0"), ("Revenue this month", f"=SUMIFS({OR_('L')},{cond})", MONEY),
          ("Profit this month", f"=SUMIFS({OR_('N')},{cond})", MONEY), ("Pieces sold", f"=SUMIFS({OR_('D')},{cond})", QTY),
          ("Confirmed, not paid yet", f'=SUMIFS({OR_("L")},{OR_("H")},"Confirmed")', MONEY),
          ("Due in the next 7 days", f'=COUNTIFS({OR_("H")},"Confirmed",{OR_("I")},">="&TODAY(),{OR_("I")},"<="&TODAY()+7)'
                                     f'+COUNTIFS({OR_("H")},"Paid",{OR_("I")},">="&TODAY(),{OR_("I")},"<="&TODAY()+7)', "0")]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 3) * 3, 1 + (i % 3) * 2, lab, f, fmt)
        dash.cell(row=6 + (i // 3) * 3, column=1 + (i % 3) * 2).alignment = Alignment(horizontal="left")
    for c, w in zip("ABCDEFGHI", [30, 11, 22, 12, 22, 12, 10, 12, 13]): dash.column_dimensions[c].width = w
    H = 14; D0 = H + 1; D1 = H + 30
    K.section(dash, f"A{H - 1}", "What each recipe earns")
    K.hdr(dash, H, ["Recipe", "Cost per piece", "Suggested price", "Price used", "Profit per piece", "Margin", "Lines to fix",
                    "Pieces sold (month)", "Profit (month)"])
    for k in range(30):
        r = D0 + k; rr_ = R0_ + k
        dash[f"A{r}"] = f'=IF(Recipes!A{rr_}="","",Recipes!A{rr_})'
        for c, src in zip("BCDEFG", "MNPQRS"):
            dash[f"{c}{r}"] = f'=IF(A{r}="","",Recipes!{src}{rr_})'
        dash[f"H{r}"] = f'=IF(A{r}="","",SUMIFS({OR_("D")},{OR_("C")},A{r},{cond}))'
        dash[f"I{r}"] = f'=IF(A{r}="","",SUMIFS({OR_("N")},{OR_("C")},A{r},{cond}))'
    style_range(dash, D0, D1, 1, 9, fmt={2: MONEY, 3: MONEY, 4: MONEY, 5: MONEY, 6: PCT, 7: '0;-0;"-"', 8: QTY, 9: MONEY})
    dash.conditional_formatting.add(f"F{D0}:F{D1}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="EC4899"))
    dash.conditional_formatting.add(f"E{D0}:F{D1}", FormulaRule(formula=[f'AND(ISNUMBER(E{D0}),E{D0}<0)'], font=Font(name=F, bold=True, color="B91C1C")))
    dash.conditional_formatting.add(f"G{D0}:G{D1}", FormulaRule(formula=[f'N(G{D0})>0'], font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    ch = BarChart(); ch.type = "bar"; ch.title = "Cost vs price per piece"; ch.height = 9.5; ch.width = 15
    ch.add_data(Reference(dash, min_col=2, min_row=H, max_row=H + 10), titles_from_data=True)
    ch.add_data(Reference(dash, min_col=4, min_row=H, max_row=H + 10), titles_from_data=True)
    ch.set_categories(Reference(dash, min_col=1, min_row=D0, max_row=H + 10))
    ch.x_axis.scaling.orientation = "maxMin"; ch.y_axis.majorGridlines = None
    ch.series[0].graphicalProperties.solidFill = "F9A8D4"; ch.series[1].graphicalProperties.solidFill = MAIN
    ch.roundedCorners = False; ch.legend.overlay = False; ch.x_axis.delete = False; ch.y_axis.delete = False
    dash.add_chart(ch, "K3")
    ch2 = BarChart(); ch2.type = "col"; ch2.title = "Revenue and profit per month"; ch2.height = 7.5; ch2.width = 15
    ch2.add_data(Reference(ym, min_col=4, min_row=5, max_row=17), titles_from_data=True)
    ch2.add_data(Reference(ym, min_col=6, min_row=5, max_row=17), titles_from_data=True)
    ch2.set_categories(Reference(ym, min_col=1, min_row=6, max_row=17)); ch2.y_axis.majorGridlines = None
    ch2.series[0].graphicalProperties.solidFill = "F9A8D4"; ch2.series[1].graphicalProperties.solidFill = MAIN; ch2.gapWidth = 60
    ch2.x_axis.number_format = "mmm"; ch2.roundedCorners = False; ch2.legend.overlay = False; ch2.x_axis.delete = False; ch2.y_axis.delete = False
    dash.add_chart(ch2, "K23")

    # ---------- Start Here ----------
    lines_ = [("How to start", True),
              ("1.  Settings tab: your hourly wage, monthly overhead, batches per month, target profit margin and wholesale %.", False),
              ("2.  Ingredients tab: what you pay – package size, unit (lb, oz, kg, gallon, dozen…) and price. Cost per gram / ml / each is calculated.", False),
              ("3.  Recipes tab: recipe name, pieces per batch, hands-on minutes and packaging per piece.", False),
              ("4.  Recipe Lines tab: the ingredients of each recipe for ONE batch, in any unit you like (cups, tbsp, g, each).", False),
              ("     The Check column says OK – or tells you what to fix. Buy by weight but measure in cups? Add 'grams per cup' on the Ingredients tab.", False),
              ("5.  Recipes tab now shows: true cost per piece, suggested retail price, wholesale price, profit per piece and margin.", False),
              ("6.  Recipe Card tab: pick a recipe and how many pieces you need – scaled quantities and the cost of that order.", False),
              ("7.  Orders tab: log every order. Dashboard: pick a month – orders, revenue, profit and what each recipe earns.", False),
              ("", False), ("Tips", True),
              ("•  Pay yourself: the price includes your hourly wage, so even a small order covers your time, not just the ingredients.", False),
              ("•  Type your own price in 'Your price per piece' to see if it is profitable. Red profit = you lose money on every piece.", False),
              ("•  Update ingredient prices when you shop – every recipe and price updates by itself.", False),
              ("•  Overhead per batch = monthly overhead ÷ batches per month. Bake more, and each batch carries less overhead.", False),
              ("", False), ("Colour key", True),
              ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
              ("", False), ("Good to know", True),
              ("•  150 ingredients, 60 recipes, 1,000 recipe lines, 1,000 orders. 15 units built in, add your own. No macros. Excel 2010+ and Google Sheets.", False),
              ("•  To use Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False)]
    if sample:
        lines_ += [("", False), ("This copy contains sample data", True),
                   ("Ingredients, prices, recipes and orders are made-up examples (June – November 2026). Use the blank copy for your own bakery.", False)]
    K.start(st, "Bakery Recipe Cost & Pricing Calculator", "Know your true cost per cookie, cake and loaf – and what to charge. Excel and Google Sheets.", lines_)
    finish(wb, sample, path)


META = {
    "name": "Bakery Recipe Cost & Pricing Calculator (Cost per Piece, Cake & Cookie Pricing, Order Tracker, Profit Dashboard) – Excel & Google Sheets",
    "file_base": "Bakery-Recipe-Cost-Pricing-Calculator",
    "price": 17,
    "color": "#9D174D",
    "title": "Bakery Recipe Cost & Pricing",
    "tagline": "true cost per piece · suggested retail & wholesale price · orders & profit",
    "intro": "Stop guessing what to charge for your cookies, cakes and bread. Enter what you pay for ingredients and what goes into each recipe – in cups, spoons, grams or pieces – and the calculator works out the true cost per piece (ingredients, packaging, your time and overhead), a suggested retail price for your profit margin, a wholesale price, and tracks your orders and profit month by month.",
    "features": [
        "Ingredient price list with automatic unit conversion: buy in lb, oz, kg, gallon or dozen – use cups, tbsp, tsp, grams or pieces",
        "Grams-per-cup column so flour, sugar and butter bought by weight can be measured by the cup",
        "Recipe costing: ingredients, packaging, your hourly wage and overhead per batch → true cost per piece",
        "Suggested retail price from your target profit margin, wholesale price, profit per piece and margin (prices rounded to your step)",
        "Check column that flags missing prices, wrong units and typos in every recipe line",
        "Recipe card & scaler: pick a recipe and the number of pieces – scaled quantities and the cost of that order",
        "Order tracker (Quote, Confirmed, Paid, Delivered, Cancelled) with revenue, cost and profit per order",
        "Monthly dashboard: orders, revenue, profit, unpaid orders, orders due this week and what each recipe earns, with charts",
        "12-month sales summary, 150 ingredients, 60 recipes, 1,000 order rows",
        "Blank copy plus a sample copy with 8 recipes and six months of example orders",
    ],
    "tags": ["bakery pricing", "recipe cost", "cake pricing", "cookie pricing", "home bakery", "bakery business", "food cost",
             "pricing calculator", "excel template", "google sheets", "small business", "cottage bakery", "order tracker"],
    "summary": "Instant download: Excel + Google Sheets template, blank copy + sample copy.",
}
