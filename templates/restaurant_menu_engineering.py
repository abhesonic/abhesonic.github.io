"""Restaurant Menu Engineering & Food Cost Calculator – ingredient costs with yield, recipe/plate costing, food cost %,
contribution margin, menu-engineering matrix (Star / Plowhorse / Puzzle / Dog), price suggestions, price tester, dashboard.

Settings -> Ingredients -> Recipe Lines -> Menu Items (analysis) -> Price Tester -> Dashboard.
"""
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import BarChart, Reference
from _common import Kit, F, B, MONEY, DATE, font, fill, style_range, input_cell, finish

MAIN = "B45309"; SOFT = "FEF3C7"
K = Kit(MAIN, SOFT)
PCT = "0.0%"; QTY = '#,##0.##;-#,##0.##;"-"'; MONEY4 = '#,##0.0000;-#,##0.0000;"-"'
NCAT = 15; C0 = 6; C1 = C0 + NCAT - 1          # categories Settings A6:A20
NI = 300; I0 = 6; I1 = I0 + NI - 1             # ingredients
NL = 2000; L0 = 6; L1 = L0 + NL - 1            # recipe lines
NM = 150; M0 = 6; M1 = M0 + NM - 1             # menu items
CATS = ["Starters", "Mains", "Burgers & Sandwiches", "Salads", "Desserts", "Drinks"]
CLASSES = ["Star", "Plowhorse", "Puzzle", "Dog"]
ACTIONS = {"Star": "Keep it: feature it, keep quality and portion; test a small price rise.",
           "Plowhorse": "Popular but low margin: raise price a little or cut plate cost.",
           "Puzzle": "Profitable but slow: promote it, rename it, move it to a better spot on the menu.",
           "Dog": "Low sales and low margin: remove, replace or rework it."}
# name, supplier, purchase unit, pack size, recipe unit, pack price, yield
INGS = [("Beef mince 80/20", "Prime Meats", "5 kg case", 5000, "g", 47.50, 1), ("Chicken breast", "Prime Meats", "5 kg case", 5000, "g", 42.00, 0.95),
        ("Salmon fillet", "Ocean Fresh", "2 kg pack", 2000, "g", 46.00, 0.92), ("Ribeye steak", "Prime Meats", "4 kg strip", 4000, "g", 98.00, 0.88),
        ("Bacon", "Prime Meats", "2 kg pack", 2000, "g", 19.80, 1), ("Shrimp 16/20", "Ocean Fresh", "1 kg bag", 1000, "g", 18.50, 0.9),
        ("Brioche bun", "City Bakery", "case of 48", 48, "each", 21.60, 1), ("Sourdough loaf", "City Bakery", "each", 16, "each", 5.20, 1),
        ("Cheddar slices", "Dairy Direct", "pack of 120", 120, "each", 14.40, 1), ("Mozzarella", "Dairy Direct", "2.5 kg block", 2500, "g", 22.50, 1),
        ("Parmesan", "Dairy Direct", "1 kg wedge", 1000, "g", 21.00, 0.95), ("Butter", "Dairy Direct", "1 kg block", 1000, "g", 8.40, 1),
        ("Heavy cream", "Dairy Direct", "2 l carton", 2000, "ml", 7.80, 1), ("Eggs", "Dairy Direct", "tray of 30", 30, "each", 6.90, 1),
        ("Milk", "Dairy Direct", "4 l jug", 4000, "ml", 4.20, 1), ("Potatoes (fries)", "Green Farm", "10 kg sack", 10000, "g", 9.50, 0.8),
        ("Frozen fries", "Food Supply Co", "6 x 2 kg", 12000, "g", 27.00, 1), ("Romaine lettuce", "Green Farm", "case of 24", 24, "each", 26.40, 0.75),
        ("Mixed greens", "Green Farm", "1.5 kg box", 1500, "g", 13.50, 0.95), ("Tomatoes", "Green Farm", "5 kg box", 5000, "g", 14.00, 0.9),
        ("Red onion", "Green Farm", "5 kg sack", 5000, "g", 6.50, 0.88), ("Avocado", "Green Farm", "case of 48", 48, "each", 52.80, 0.7),
        ("Lemon", "Green Farm", "case of 100", 100, "each", 28.00, 1), ("Garlic", "Green Farm", "1 kg bag", 1000, "g", 7.20, 0.85),
        ("Spaghetti", "Food Supply Co", "5 kg case", 5000, "g", 11.50, 1), ("Arborio rice", "Food Supply Co", "5 kg bag", 5000, "g", 16.00, 1),
        ("Mushrooms", "Green Farm", "2.5 kg box", 2500, "g", 15.00, 0.92), ("Olive oil", "Food Supply Co", "5 l tin", 5000, "ml", 38.00, 1),
        ("Canola oil (fryer)", "Food Supply Co", "20 l drum", 20000, "ml", 42.00, 1), ("Tomato sauce", "Food Supply Co", "3 kg can", 3000, "g", 7.50, 1),
        ("Burger sauce", "Food Supply Co", "2 kg tub", 2000, "g", 12.00, 1), ("Caesar dressing", "Food Supply Co", "4 l jug", 4000, "ml", 26.00, 1),
        ("Chicken wings", "Prime Meats", "5 kg case", 5000, "g", 31.00, 1), ("Buffalo sauce", "Food Supply Co", "4 l jug", 4000, "ml", 22.00, 1),
        ("Flour tortilla 12in", "City Bakery", "pack of 72", 72, "each", 18.00, 1), ("Dark chocolate", "Sweet Supply", "2.5 kg bag", 2500, "g", 34.00, 1),
        ("Sugar", "Food Supply Co", "10 kg bag", 10000, "g", 9.80, 1), ("Vanilla ice cream", "Dairy Direct", "5 l tub", 5000, "ml", 19.50, 1),
        ("Cream cheese", "Dairy Direct", "1.36 kg tub", 1360, "g", 11.90, 1), ("Graham crumbs", "Sweet Supply", "1 kg bag", 1000, "g", 6.20, 1),
        ("Coffee beans", "Roast House", "1 kg bag", 1000, "g", 18.00, 1), ("Lemonade base", "Drink Co", "case of 24 x 330 ml", 24, "each", 14.40, 1),
        ("Craft beer keg", "Local Brewery", "30 l keg", 30000, "ml", 165.00, 0.95), ("House wine", "Wine Co", "case 12 x 750 ml", 9000, "ml", 78.00, 1),
        ("Fresh herbs", "Green Farm", "bunch of 100 g", 100, "g", 2.40, 0.8), ("Croutons", "Food Supply Co", "1 kg bag", 1000, "g", 8.50, 1)]
# item, category, price, units sold, other cost (garnish/packaging), recipe [(ingredient, qty)]
MENU = [("Garlic Bread", "Starters", 7.00, 410, 0.10, [("Sourdough loaf", 0.25), ("Butter", 25), ("Garlic", 8), ("Parmesan", 10), ("Fresh herbs", 2)]),
        ("Buffalo Wings", "Starters", 13.00, 520, 0.15, [("Chicken wings", 300), ("Buffalo sauce", 60), ("Canola oil (fryer)", 30), ("Butter", 10)]),
        ("Garlic Shrimp", "Starters", 15.00, 140, 0.10, [("Shrimp 16/20", 150), ("Butter", 25), ("Garlic", 10), ("Sourdough loaf", 0.125), ("Lemon", 0.25)]),
        ("Loaded Fries", "Starters", 11.00, 380, 0.10, [("Frozen fries", 250), ("Bacon", 40), ("Cheddar slices", 2), ("Burger sauce", 30), ("Canola oil (fryer)", 30)]),
        ("Avocado Toast", "Starters", 12.00, 90, 0.10, [("Sourdough loaf", 0.125), ("Avocado", 1), ("Eggs", 1), ("Tomatoes", 40), ("Olive oil", 10), ("Lemon", 0.25)]),
        ("Classic Cheeseburger", "Burgers & Sandwiches", 16.00, 760, 0.25, [("Beef mince 80/20", 180), ("Brioche bun", 1), ("Cheddar slices", 2), ("Burger sauce", 30), ("Tomatoes", 30), ("Romaine lettuce", 0.06), ("Red onion", 15), ("Frozen fries", 180), ("Canola oil (fryer)", 25)]),
        ("Bacon Double Burger", "Burgers & Sandwiches", 21.00, 300, 0.25, [("Beef mince 80/20", 300), ("Brioche bun", 1), ("Cheddar slices", 3), ("Bacon", 50), ("Burger sauce", 40), ("Frozen fries", 180), ("Canola oil (fryer)", 25)]),
        ("Crispy Chicken Sandwich", "Burgers & Sandwiches", 17.00, 450, 0.25, [("Chicken breast", 180), ("Flour tortilla 12in", 0), ("Brioche bun", 1), ("Burger sauce", 30), ("Romaine lettuce", 0.06), ("Canola oil (fryer)", 40), ("Frozen fries", 180), ("Eggs", 0.5)]),
        ("Chicken Wrap", "Burgers & Sandwiches", 14.00, 130, 0.20, [("Chicken breast", 150), ("Flour tortilla 12in", 1), ("Romaine lettuce", 0.08), ("Tomatoes", 40), ("Caesar dressing", 30), ("Parmesan", 10)]),
        ("Ribeye Steak Frites", "Mains", 34.00, 260, 0.30, [("Ribeye steak", 300), ("Potatoes (fries)", 300), ("Canola oil (fryer)", 40), ("Butter", 20), ("Garlic", 5), ("Fresh herbs", 3), ("Mixed greens", 30)]),
        ("Grilled Salmon", "Mains", 28.00, 210, 0.30, [("Salmon fillet", 200), ("Arborio rice", 0), ("Mixed greens", 50), ("Lemon", 0.5), ("Butter", 20), ("Potatoes (fries)", 200), ("Olive oil", 10)]),
        ("Spaghetti Pomodoro", "Mains", 18.00, 330, 0.15, [("Spaghetti", 140), ("Tomato sauce", 150), ("Olive oil", 15), ("Garlic", 8), ("Parmesan", 15), ("Fresh herbs", 3)]),
        ("Mushroom Risotto", "Mains", 22.00, 120, 0.15, [("Arborio rice", 110), ("Mushrooms", 120), ("Butter", 30), ("Parmesan", 25), ("Heavy cream", 40), ("Garlic", 6), ("Olive oil", 10)]),
        ("Shrimp Pasta", "Mains", 26.00, 95, 0.15, [("Spaghetti", 140), ("Shrimp 16/20", 180), ("Heavy cream", 100), ("Garlic", 10), ("Parmesan", 20), ("Butter", 15)]),
        ("Caesar Salad", "Salads", 13.00, 290, 0.10, [("Romaine lettuce", 0.5), ("Caesar dressing", 60), ("Parmesan", 20), ("Croutons", 30), ("Lemon", 0.25)]),
        ("Chicken Caesar", "Salads", 17.00, 340, 0.10, [("Romaine lettuce", 0.5), ("Caesar dressing", 60), ("Parmesan", 20), ("Croutons", 30), ("Chicken breast", 150), ("Olive oil", 5)]),
        ("Salmon Avocado Bowl", "Salads", 21.00, 85, 0.20, [("Salmon fillet", 150), ("Avocado", 1), ("Mixed greens", 80), ("Arborio rice", 80), ("Lemon", 0.5), ("Olive oil", 10)]),
        ("Chocolate Brownie", "Desserts", 9.00, 310, 0.10, [("Dark chocolate", 50), ("Butter", 40), ("Sugar", 50), ("Eggs", 0.75), ("Vanilla ice cream", 80)]),
        ("NY Cheesecake", "Desserts", 10.00, 150, 0.10, [("Cream cheese", 120), ("Sugar", 35), ("Eggs", 0.5), ("Graham crumbs", 30), ("Butter", 15), ("Heavy cream", 20)]),
        ("Affogato", "Desserts", 7.00, 60, 0.05, [("Vanilla ice cream", 120), ("Coffee beans", 18)]),
        ("Ice Cream Sundae", "Desserts", 8.00, 70, 0.10, [("Vanilla ice cream", 200), ("Dark chocolate", 30), ("Heavy cream", 40)]),
        ("Craft Beer (pint)", "Drinks", 8.00, 980, 0, [("Craft beer keg", 473)]),
        ("House Wine (glass)", "Drinks", 10.00, 520, 0, [("House wine", 175)]),
        ("Cappuccino", "Drinks", 4.50, 610, 0.08, [("Coffee beans", 18), ("Milk", 180)]),
        ("Homemade Lemonade", "Drinks", 5.00, 170, 0.08, [("Lemonade base", 1), ("Lemon", 0.5), ("Sugar", 15)]),
        ("Milkshake", "Drinks", 8.50, 120, 0.10, [("Vanilla ice cream", 250), ("Milk", 150), ("Heavy cream", 30)])]


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); s = wb.create_sheet("Settings"); ing = wb.create_sheet("Ingredients")
    ln = wb.create_sheet("Recipe Lines"); mi = wb.create_sheet("Menu Items"); pt = wb.create_sheet("Price Tester")

    # ---------- Settings ----------
    K.title(s, "Settings", "Yellow cells are yours to edit.")
    opts = [("Restaurant name", "Harbor Street Bistro" if sample else "My Restaurant", None), ("Currency label", "USD", None),
            ("Sales period (label)", "October 2026" if sample else "This month", None), ("Target food cost %", 0.30, PCT),
            ("Compare items within", "Category", None), ("Popularity factor (70% rule)", 0.70, "0%")]
    for i, (lab, v, fmt) in enumerate(opts, 3):
        s[f"D{i}"] = lab; s[f"D{i}"].font = font(bold=True); input_cell(s[f"E{i}"], v)
        if fmt: s[f"E{i}"].number_format = fmt
    dv = DataValidation(type="list", formula1='"Category,Whole menu"', allow_blank=False); s.add_data_validation(dv); dv.add("E7")
    s["D10"] = "How the matrix works"; s["D10"].font = font(bold=True)
    notes = ["• Popular = the item's share of sales is at least 70% of a fair share (1 ÷ number of items in its group).",
             "• High margin = the item's contribution margin (price − plate cost) is at least the group's average, weighted by units sold.",
             "• Star = popular + high margin · Plowhorse = popular + low margin · Puzzle = slow + high margin · Dog = slow + low margin.",
             "• 'Compare items within' Category compares starters with starters, mains with mains. 'Whole menu' compares everything together."]
    for i, t in enumerate(notes, 11):
        s[f"D{i}"] = t; s[f"D{i}"].font = font(size=9); s[f"D{i}"].alignment = Alignment(wrap_text=True, vertical="top"); s.row_dimensions[i].height = 26
    K.hdr(s, 5, ["Menu categories"])
    for i in range(NCAT): input_cell(s[f"A{C0 + i}"], CATS[i] if i < len(CATS) else None)
    s.column_dimensions["A"].width = 26; s.column_dimensions["B"].width = 3; s.column_dimensions["C"].width = 3
    s.column_dimensions["D"].width = 70; s.column_dimensions["E"].width = 24
    TGT = "Settings!$E$6"; CAT = f"Settings!$A${C0}:$A${C1}"

    # ---------- Ingredients ----------
    K.title(ing, "Ingredients", "Type what you pay. Pack size must be in the recipe unit you use (g, ml, each…). Yield = usable share after trimming/cooking.")
    K.hdr(ing, 5, ["Ingredient", "Supplier", "Purchase unit", "Pack size", "Recipe unit", "Pack price", "Yield %", "Price updated", "Cost per recipe unit", "Used in items"],
          [24, 16, 18, 10, 10, 11, 9, 12, 13, 10])
    style_range(ing, I0, I1, 1, 10, inputs=tuple(range(1, 9)), fmt={4: QTY, 6: MONEY, 7: "0%", 8: DATE, 9: MONEY4, 10: "0"})
    for r in range(I0, I1 + 1):
        ing[f"I{r}"] = f'=IF(OR(A{r}="",N(D{r})=0),0,F{r}/(D{r}*IF(N(G{r})=0,1,G{r})))'
        ing[f"J{r}"] = f'=IF(A{r}="","",COUNTIFS(\'Recipe Lines\'!$B${L0}:$B${L1},A{r},\'Recipe Lines\'!$C${L0}:$C${L1},">0"))'
    dv = DataValidation(type="list", formula1='"g,kg,ml,l,oz,lb,fl oz,each,portion"', allow_blank=True); ing.add_data_validation(dv); dv.add(f"E{I0}:E{I1}")
    ing.freeze_panes = f"B{I0}"
    IN = f"Ingredients!$A${I0}:$A${I1}"

    # ---------- Recipe Lines ----------
    K.title(ln, "Recipe Lines", "One row per ingredient in a menu item. Quantity is in the ingredient's recipe unit (e.g. 180 g, 1 each).")
    K.hdr(ln, 5, ["Menu item", "Ingredient", "Quantity", "Unit", "Cost per unit", "Line cost"], [28, 24, 10, 9, 13, 11])
    style_range(ln, L0, L1, 1, 6, inputs=(1, 2, 3), fmt={3: QTY, 5: MONEY4, 6: MONEY})
    for r in range(L0, L1 + 1):
        ln[f"D{r}"] = f'=IF(B{r}="","",IFERROR(INDEX(Ingredients!$E${I0}:$E${I1},MATCH(B{r},{IN},0))&"","?"))'
        ln[f"E{r}"] = f'=IF(B{r}="",0,IFERROR(INDEX(Ingredients!$I${I0}:$I${I1},MATCH(B{r},{IN},0)),0))'
        ln[f"F{r}"] = f'=N(C{r})*E{r}'
    for rng, f1 in ((f"A{L0}:A{L1}", f"='Menu Items'!$A${M0}:$A${M1}"), (f"B{L0}:B{L1}", f"={IN}")):
        dv = DataValidation(type="list", formula1=f1, allow_blank=True); ln.add_data_validation(dv); dv.add(rng)
    ln.conditional_formatting.add(f"D{L0}:D{L1}", FormulaRule(formula=[f'D{L0}="?"'], font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    ln.freeze_panes = f"A{L0}"

    # ---------- Menu Items ----------
    K.title(mi, "Menu Items", "One row per dish or drink. Type price and units sold for the period – the rest calculates from your recipes.")
    hd = ["Menu item", "Category", "Selling price", "Units sold", "Other cost (garnish, packaging)", "Plate cost", "Food cost %",
          "Contribution margin", "Sales", "Food cost total", "Total margin", "Menu mix %", "Popularity", "Margin", "Class", "What to do",
          "Price for target food cost", "Price gap", "group", "rank margin", "rank fc"]
    K.hdr(mi, 5, hd, [26, 20, 10, 9, 12, 10, 9, 12, 11, 11, 12, 9, 10, 9, 11, 46, 12, 10, 6, 6, 6])
    style_range(mi, M0, M1, 1, 21, inputs=(1, 2, 3, 4, 5),
                fmt={3: MONEY, 4: "#,##0", 5: MONEY, 6: MONEY, 7: PCT, 8: MONEY, 9: MONEY, 10: MONEY, 11: MONEY, 12: PCT, 17: MONEY, 18: MONEY,
                     20: "0.000000", 21: "0.0000000"})
    LN = lambda c: f"'Recipe Lines'!${c}${L0}:${c}${L1}"
    MI = lambda c: f"$" + c + f"${M0}:${c}${M1}"
    for r in range(M0, M1 + 1):
        mi[f"F{r}"] = f'=IF(A{r}="",0,SUMIFS({LN("F")},{LN("A")},A{r})+N(E{r}))'
        mi[f"G{r}"] = f'=IF(OR(A{r}="",N(C{r})=0),"",F{r}/C{r})'
        mi[f"H{r}"] = f'=IF(A{r}="",0,N(C{r})-F{r})'
        mi[f"I{r}"] = f'=N(C{r})*N(D{r})'
        mi[f"J{r}"] = f'=F{r}*N(D{r})'
        mi[f"K{r}"] = f'=H{r}*N(D{r})'
        mi[f"S{r}"] = f'=IF(A{r}="","",IF(Settings!$E$7="Whole menu","All",IF(B{r}="","(none)",B{r})))'
        mi[f"L{r}"] = f'=IF(A{r}="","",IFERROR(N(D{r})/SUMIFS({MI("D")},{MI("S")},S{r}),0))'
        mi[f"M{r}"] = f'=IF(A{r}="","",IF(L{r}>=Settings!$E$8/COUNTIF({MI("S")},S{r}),"High","Low"))'
        mi[f"N{r}"] = f'=IF(A{r}="","",IF(H{r}>=IFERROR(SUMIFS({MI("K")},{MI("S")},S{r})/SUMIFS({MI("D")},{MI("S")},S{r}),0),"High","Low"))'
        mi[f"O{r}"] = f'=IF(A{r}="","",IF(M{r}="High",IF(N{r}="High","Star","Plowhorse"),IF(N{r}="High","Puzzle","Dog")))'
        mi[f"P{r}"] = (f'=IF(O{r}="","",IF(O{r}="Star","{ACTIONS["Star"]}",IF(O{r}="Plowhorse","{ACTIONS["Plowhorse"]}",'
                       f'IF(O{r}="Puzzle","{ACTIONS["Puzzle"]}","{ACTIONS["Dog"]}"))))')
        mi[f"Q{r}"] = f'=IF(OR(A{r}="",N({TGT})=0),"",ROUND(F{r}/{TGT}*4,0)/4)'
        mi[f"R{r}"] = f'=IF(Q{r}="","",Q{r}-N(C{r}))'
        mi[f"T{r}"] = f'=IF(A{r}="","",K{r}+ROW()/1000000)'
        mi[f"U{r}"] = f'=IF(OR(A{r}="",G{r}=""),"",G{r}+ROW()/10000000)'
    dv = DataValidation(type="list", formula1=f"={CAT}", allow_blank=True); mi.add_data_validation(dv); dv.add(f"B{M0}:B{M1}")
    for cls, fc, bg in (("Star", "15803D", "DCFCE7"), ("Plowhorse", "1D4ED8", "DBEAFE"), ("Puzzle", "7C3AED", "EDE9FE"), ("Dog", "B91C1C", "FEE2E2")):
        mi.conditional_formatting.add(f"O{M0}:O{M1}", FormulaRule(formula=[f'O{M0}="{cls}"'], font=Font(name=F, bold=True, color=fc), fill=fill(bg)))
    mi.conditional_formatting.add(f"G{M0}:G{M1}", FormulaRule(formula=[f'AND(ISNUMBER(G{M0}),G{M0}>{TGT})'], font=Font(name=F, bold=True, color="B91C1C")))
    mi.conditional_formatting.add(f"R{M0}:R{M1}", FormulaRule(formula=[f'AND(ISNUMBER(R{M0}),R{M0}>0)'], font=Font(name=F, bold=True, color="B91C1C")))
    for c in "STU": mi.column_dimensions[c].hidden = True
    mi.freeze_panes = f"B{M0}"
    MN = f"'Menu Items'!$A${M0}:$A${M1}"
    MC = lambda c: f"'Menu Items'!${c}${M0}:${c}${M1}"

    # ---------- sample data ----------
    if sample:
        for r, x in enumerate(INGS, I0):
            for col, v in zip("ABCDEFG", x): ing[f"{col}{r}"] = round(v * 1.35, 2) if col == "F" else v
            ing[f"H{r}"] = None
        import datetime as dt
        for r in range(I0, I0 + len(INGS)): ing[f"H{r}"] = dt.date(2026, 10, 1 + (r % 20))
        r = L0
        for i, (name, cat, price, units, oth, rec) in enumerate(MENU):
            mr = M0 + i
            for col, v in zip("ABCDE", [name, cat, price, units, oth or None]): mi[f"{col}{mr}"] = v
            for ingn, q in rec:
                if q == 0: continue
                ln[f"A{r}"] = name; ln[f"B{r}"] = ingn; ln[f"C{r}"] = q; r += 1
        assert r - L0 <= NL

    # ---------- Price Tester ----------
    K.title(pt, "Price Tester", "What happens to profit if you change a price? Pick an item, type a new price and how sales might change.")
    pin = [("Menu item", MENU[5][0] if sample else None, None), ("New selling price", 17.5 if sample else None, MONEY),
           ("Expected change in units sold", -0.03 if sample else 0, "0%"), ("New plate cost (optional)", None, MONEY)]
    for i, (lab, v, fmt) in enumerate(pin, 4):
        pt[f"A{i}"] = lab; pt[f"A{i}"].font = font(bold=True); input_cell(pt[f"B{i}"], v)
        if fmt: pt[f"B{i}"].number_format = fmt
    dv = DataValidation(type="list", formula1=f"={MN}", allow_blank=True); pt.add_data_validation(dv); dv.add("B4")
    K.hdr(pt, 9, ["", "Now", "With the change", "Difference"])
    LK = lambda c: f'IFERROR(INDEX({MC(c)},MATCH($B$4,{MN},0)),0)'
    rows = [("Selling price", f"={LK('C')}", '=IF(N(B5)=0,B10,B5)', MONEY),
            ("Plate cost", f"={LK('F')}", '=IF(B7="",B11,B7)', MONEY),
            ("Food cost %", '=IFERROR(B11/B10,0)', '=IFERROR(C11/C10,0)', PCT),
            ("Margin per plate", "=B10-B11", "=C10-C11", MONEY),
            ("Units sold", f"={LK('D')}", "=ROUND(B14*(1+N(B6)),0)", "#,##0"),
            ("Sales", "=B10*B14", "=C10*C14", MONEY),
            ("Total margin (profit before labour & overheads)", "=B13*B14", "=C13*C14", MONEY)]
    for i, (lab, a, b, fmt) in enumerate(rows, 10):
        pt[f"A{i}"] = lab; pt[f"A{i}"].font = font(bold=True)
        pt[f"B{i}"] = a; pt[f"C{i}"] = b; pt[f"D{i}"] = f"=C{i}-B{i}"
        for c in "BCD":
            pt[f"{c}{i}"].number_format = fmt; pt[f"{c}{i}"].border = B; pt[f"{c}{i}"].fill = fill("F3F4F6"); pt[f"{c}{i}"].font = font()
    for c in "ABCD":
        pt[f"{c}16"].font = Font(name=F, bold=True, size=12, color=MAIN); pt[f"{c}16"].fill = fill(SOFT)
    pt["A18"] = "Units you need to sell at the new price to earn the same total margin"; pt["A18"].font = font(bold=True)
    pt["C18"] = '=IF(C13<=0,"",B16/C13)'; pt["C18"].number_format = "#,##0"
    pt["A19"] = "So sales can drop by up to"; pt["A19"].font = font(bold=True)
    pt["C19"] = '=IF(OR(C18="",B14=0),"",1-C18/B14)'; pt["C19"].number_format = "0.0%"
    pt["A20"] = "Price that hits your target food cost"; pt["A20"].font = font(bold=True)
    pt["C20"] = f'=IF(N({TGT})=0,"",ROUND(C11/{TGT}*4,0)/4)'; pt["C20"].number_format = MONEY
    for c in ("C18", "C19", "C20"):
        pt[c].border = B; pt[c].fill = fill("F3F4F6"); pt[c].font = Font(name=F, bold=True, color=MAIN)
    pt.column_dimensions["A"].width = 50
    for c in "BCD": pt.column_dimensions[c].width = 16

    # ---------- Dashboard ----------
    K.title(dash, "Menu Engineering Dashboard", "Updates by itself from Ingredients, Recipe Lines and Menu Items.")
    dash["A3"] = '=Settings!E3&"  ·  "&Settings!E5&"  ·  amounts in "&Settings!E4'; dash["A3"].font = font(bold=True)
    kp = [("Sales", f"=SUM({MC('I')})", MONEY), ("Food cost %", f"=IFERROR(SUM({MC('J')})/SUM({MC('I')}),0)", PCT),
          ("Total margin", f"=SUM({MC('K')})", MONEY), ("Items sold", f"=SUM({MC('D')})", "#,##0"),
          ("Avg margin per item sold", f"=IFERROR(SUM({MC('K')})/SUM({MC('D')}),0)", MONEY),
          ("Items over target food cost", f'=SUMPRODUCT(({MC("A")}<>"")*({MC("C")}>0)*({MC("F")}>{MC("C")}*{TGT}))', "0")]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 3) * 3, 1 + (i % 3) * 2, lab, f, fmt)
        dash.cell(row=6 + (i // 3) * 3, column=1 + (i % 3) * 2).alignment = Alignment(horizontal="left")
    dash.conditional_formatting.add("C6", FormulaRule(formula=[f"C6>{TGT}"], font=Font(name=F, bold=True, size=18, color="B91C1C")))
    for c, w in zip("ABCDEFG", [28, 13, 22, 13, 22, 13, 3]): dash.column_dimensions[c].width = w
    K.section(dash, "A11", "Menu matrix")
    K.hdr(dash, 12, ["Class", "Items", "Units sold", "Sales", "Total margin", "What to do"])
    for i, cls in enumerate(CLASSES):
        r = 13 + i
        dash[f"A{r}"] = cls
        dash[f"B{r}"] = f'=COUNTIF({MC("O")},A{r})'
        dash[f"C{r}"] = f'=SUMIFS({MC("D")},{MC("O")},A{r})'
        dash[f"D{r}"] = f'=SUMIFS({MC("I")},{MC("O")},A{r})'
        dash[f"E{r}"] = f'=SUMIFS({MC("K")},{MC("O")},A{r})'
        dash[f"F{r}"] = {"Star": "Keep & feature", "Plowhorse": "Raise price / cut cost", "Puzzle": "Promote / reposition", "Dog": "Remove or rework"}[cls]
    style_range(dash, 13, 16, 1, 6, fmt={2: "0", 3: "#,##0", 4: MONEY, 5: MONEY})
    for i, (cls, fc, bg) in enumerate((("Star", "15803D", "DCFCE7"), ("Plowhorse", "1D4ED8", "DBEAFE"), ("Puzzle", "7C3AED", "EDE9FE"), ("Dog", "B91C1C", "FEE2E2"))):
        c = dash[f"A{13 + i}"]; c.font = Font(name=F, bold=True, color=fc); c.fill = fill(bg)
    TR = 19
    K.section(dash, f"A{TR}", "Top 10 items by total margin")
    K.hdr(dash, TR + 1, ["Menu item", "Units sold", "Margin per plate", "Total margin", "Class", "Food cost %"])
    RK = MC("T"); RF = MC("U")
    for k in range(10):
        r = TR + 2 + k
        dash[f"H{r}"] = f'=IFERROR(MATCH(LARGE({RK},{k + 1}),{RK},0),"")'
        dash[f"A{r}"] = f'=IF(H{r}="","",INDEX({MC("A")},H{r}))'
        dash[f"B{r}"] = f'=IF(H{r}="","",INDEX({MC("D")},H{r}))'
        dash[f"C{r}"] = f'=IF(H{r}="","",INDEX({MC("H")},H{r}))'
        dash[f"D{r}"] = f'=IF(H{r}="","",INDEX({MC("K")},H{r}))'
        dash[f"E{r}"] = f'=IF(H{r}="","",INDEX({MC("O")},H{r})&"")'
        dash[f"F{r}"] = f'=IF(H{r}="","",INDEX({MC("G")},H{r}))'
        dash[f"H{r}"].font = Font(name=F, color="FFFFFF", size=8)
    style_range(dash, TR + 2, TR + 11, 1, 6, fmt={2: "#,##0", 3: MONEY, 4: MONEY, 6: PCT})
    FR = TR + 13
    K.section(dash, f"A{FR}", "Highest food cost % – check portions and prices first")
    K.hdr(dash, FR + 1, ["Menu item", "Price", "Plate cost", "Food cost %", "Price for target", "Class"])
    for k in range(5):
        r = FR + 2 + k
        dash[f"H{r}"] = f'=IFERROR(MATCH(LARGE({RF},{k + 1}),{RF},0),"")'
        dash[f"A{r}"] = f'=IF(H{r}="","",INDEX({MC("A")},H{r}))'
        dash[f"B{r}"] = f'=IF(H{r}="","",INDEX({MC("C")},H{r}))'
        dash[f"C{r}"] = f'=IF(H{r}="","",INDEX({MC("F")},H{r}))'
        dash[f"D{r}"] = f'=IF(H{r}="","",INDEX({MC("G")},H{r}))'
        dash[f"E{r}"] = f'=IF(H{r}="","",INDEX({MC("Q")},H{r}))'
        dash[f"F{r}"] = f'=IF(H{r}="","",INDEX({MC("O")},H{r})&"")'
        dash[f"H{r}"].font = Font(name=F, color="FFFFFF", size=8)
    style_range(dash, FR + 2, FR + 6, 1, 6, fmt={2: MONEY, 3: MONEY, 4: PCT, 5: MONEY})
    CR = FR + 8
    K.section(dash, f"A{CR}", "By category")
    K.hdr(dash, CR + 1, ["Category", "Items", "Units sold", "Sales", "Food cost %", "Total margin"])
    for i in range(NCAT):
        r = CR + 2 + i; k = C0 + i
        dash[f"A{r}"] = f'=IF(Settings!A{k}="","",Settings!A{k})'
        dash[f"B{r}"] = f'=IF(A{r}="","",COUNTIF({MC("B")},A{r}))'
        dash[f"C{r}"] = f'=IF(A{r}="","",SUMIFS({MC("D")},{MC("B")},A{r}))'
        dash[f"D{r}"] = f'=IF(A{r}="","",SUMIFS({MC("I")},{MC("B")},A{r}))'
        dash[f"E{r}"] = f'=IF(OR(A{r}="",N(D{r})=0),"",SUMIFS({MC("J")},{MC("B")},A{r})/D{r})'
        dash[f"F{r}"] = f'=IF(A{r}="","",SUMIFS({MC("K")},{MC("B")},A{r}))'
    style_range(dash, CR + 2, CR + 1 + NCAT, 1, 6, fmt={2: "0", 3: "#,##0", 4: MONEY, 5: PCT, 6: MONEY})
    ch = BarChart(); ch.type = "bar"; ch.title = "Total margin by category"; ch.height = 8; ch.width = 17
    ch.add_data(Reference(dash, min_col=6, min_row=CR + 1, max_row=CR + 1 + len(CATS)), titles_from_data=True)
    ch.set_categories(Reference(dash, min_col=1, min_row=CR + 2, max_row=CR + 1 + len(CATS))); ch.legend = None
    ch.x_axis.scaling.orientation = "maxMin"; ch.y_axis.majorGridlines = None; ch.series[0].graphicalProperties.solidFill = MAIN
    ch.roundedCorners = False; ch.x_axis.delete = False; ch.y_axis.delete = False
    dash.add_chart(ch, "J5")
    ch2 = BarChart(); ch2.type = "col"; ch2.title = "Total margin by class"; ch2.height = 8; ch2.width = 17
    ch2.add_data(Reference(dash, min_col=5, min_row=12, max_row=16), titles_from_data=True)
    ch2.set_categories(Reference(dash, min_col=1, min_row=13, max_row=16)); ch2.legend = None; ch2.y_axis.majorGridlines = None
    ch2.series[0].graphicalProperties.solidFill = "F59E0B"
    ch2.roundedCorners = False; ch2.x_axis.delete = False; ch2.y_axis.delete = False
    dash.add_chart(ch2, "J22")

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Settings tab: restaurant name, sales period, your target food cost % (e.g. 28–32%) and your menu categories.", False),
             ("2.  Ingredients tab: what you pay for each ingredient – pack size in the unit your recipes use (g, ml, each), pack price and yield % after trimming.", False),
             ("3.  Menu Items tab: each dish or drink with its category, selling price and how many you sold in the period (from your POS report).", False),
             ("4.  Recipe Lines tab: one row per ingredient in each item (e.g. Cheeseburger · Beef mince · 180). Plate cost and food cost % appear by themselves.", False),
             ("5.  Menu Items now shows contribution margin, menu mix and the class of every item: Star, Plowhorse, Puzzle or Dog – plus what to do and the price that hits your target.", False),
             ("6.  Price Tester: try a new price and see how many sales you could lose and still earn the same.", False),
             ("7.  Dashboard: sales, food cost %, total margin, the menu matrix, top items, worst food cost items and categories.", False),
             ("", False), ("Tips", True),
             ("•  When a supplier price changes, update only the Ingredients tab – every plate cost updates.", False),
             ("•  Use 'Other cost' for garnish, sauces you don't track, napkins or take-away packaging.", False),
             ("•  Do this every month with fresh sales numbers – menu engineering works best when you compare periods.", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  300 ingredients, 150 menu items, 2,000 recipe lines. No macros. Works in Excel 2010+ and Google Sheets.", False),
             ("•  To use Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("The restaurant, ingredients, prices and sales are made-up examples. Use the blank copy in your download for your own menu.", False)]
    K.start(st, "Restaurant Menu Engineering & Food Cost Calculator", "Plate costs, food cost %, contribution margin and the Star / Plowhorse / Puzzle / Dog matrix – Excel and Google Sheets.", lines)
    finish(wb, sample, path)


META = {
    "name": "Restaurant Menu Engineering & Food Cost Calculator (Plate Cost, Food Cost %, Menu Matrix, Price Tester) – Excel & Google Sheets",
    "file_base": "Restaurant-Menu-Engineering-Food-Cost-Calculator",
    "price": 24,
    "color": "#B45309",
    "title": "Menu Engineering & Food Cost",
    "tagline": "plate cost · food cost % · margin · Stars, Plowhorses, Puzzles & Dogs",
    "intro": "Find out which dishes really make your restaurant money. Enter what you pay for ingredients and your recipes once, add your selling prices and sales from your POS, and the spreadsheet costs every plate, shows food cost % and contribution margin, and sorts your whole menu into Stars, Plowhorses, Puzzles and Dogs with a clear action for each. For restaurants, cafes, bars, food trucks and caterers.",
    "features": [
        "Ingredient cost list for 300 items with pack size, pack price and yield % (trim and cooking loss)",
        "Recipe costing: plate cost per dish from up to 2,000 recipe lines, plus garnish/packaging cost",
        "Food cost %, contribution margin, sales and total margin for up to 150 menu items",
        "Menu-engineering matrix (Kasavana & Smith method): Star, Plowhorse, Puzzle or Dog with what to do – by category or whole menu",
        "Suggested price for your target food cost and the price gap for every item",
        "Price Tester: see how many sales you can lose at a new price and still earn the same",
        "Dashboard: food cost %, matrix summary, top 10 items, worst food cost items, categories and charts; blank copy plus a sample copy",
    ],
    "tags": ["menu engineering", "food cost calculator", "restaurant spreadsheet", "recipe costing", "plate cost", "food cost percentage",
             "restaurant template", "menu pricing", "cafe business", "food truck", "google sheets", "excel template", "restaurant profit"],
    "summary": "Instant download: Excel + Google Sheets template, blank copy + sample copy.",
}
