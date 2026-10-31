"""Meal Planner & Grocery List with macro totals (sample + blank versions).

Ingredients (nutrition per 100 g/ml or per piece, pack size & price, what you have at home)
-> Recipes + Recipe Ingredients (calories, protein, carbs, fat and cost per serving, calculated)
-> Weekly Plan (7 days x breakfast/lunch/dinner/snack, portions to cook) with daily macro totals vs targets
-> Grocery List (only what you still need, grouped by store section, packs to buy, estimated cost)
-> Dashboard.
"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter as L
from _common import Kit, F, INK, B, MONEY, DATE, font, fill, style_range, input_cell, finish

MAIN = "0F766E"; SOFT = "E6F6F4"; ACC = "5EEAD4"
K = Kit(MAIN, SOFT)
NI = 200; I0 = 6; I1 = I0 + NI - 1      # ingredients
NR = 80; R0 = 6; R1 = R0 + NR - 1       # recipes
NL = 800; L0 = 6; L1 = L0 + NL - 1      # recipe ingredient lines
P0 = 7; P1 = 13                          # plan days
NG = 100; G0 = 7                         # grocery rows
SECTIONS = ["Produce", "Meat & fish", "Dairy & eggs", "Refrigerated", "Bakery", "Pantry", "Frozen", "Spices & sauces", "Drinks", "Household", "Other", ""]
MEALTYPES = ["Breakfast", "Lunch", "Dinner", "Snack", "Any"]
RI = "'Recipe Ingredients'"; WP = "'Weekly Plan'"
G = "0.#"; CUR = '#,##0.00;-#,##0.00;""'

# name, section, unit, per, kcal, protein, carbs, fat, pack size, pack price   (approximate values, USDA-style)
ING = [("Rolled oats", "Pantry", "g", 100, 379, 13, 68, 6.5, 1000, 4.5), ("Milk (2%)", "Dairy & eggs", "ml", 100, 50, 3.3, 4.8, 2, 1000, 1.5),
       ("Greek yogurt (plain, 2%)", "Dairy & eggs", "g", 100, 73, 10, 3.9, 2, 900, 5), ("Banana", "Produce", "pcs", 1, 105, 1.3, 27, 0.4, 1, 0.3),
       ("Blueberries", "Produce", "g", 100, 57, 0.7, 14.5, 0.3, 170, 3.5), ("Eggs", "Dairy & eggs", "pcs", 1, 72, 6.3, 0.4, 4.8, 12, 4),
       ("Whole wheat bread (slice)", "Bakery", "pcs", 1, 100, 4, 17, 1.5, 20, 3.5), ("Peanut butter", "Pantry", "g", 100, 588, 25, 20, 50, 450, 3.5),
       ("Chicken breast", "Meat & fish", "g", 100, 120, 22.5, 0, 2.6, 1000, 9), ("Brown rice (dry)", "Pantry", "g", 100, 367, 7.5, 76, 3, 1000, 3),
       ("Broccoli", "Produce", "g", 100, 34, 2.8, 7, 0.4, 500, 2.5), ("Olive oil", "Pantry", "ml", 100, 884, 0, 0, 100, 500, 7),
       ("Salmon fillet", "Meat & fish", "g", 100, 208, 20, 0, 13, 500, 11), ("Sweet potato", "Produce", "g", 100, 86, 1.6, 20, 0.1, 1000, 2.5),
       ("Spinach", "Produce", "g", 100, 23, 2.9, 3.6, 0.4, 250, 3), ("Ground beef (90% lean)", "Meat & fish", "g", 100, 176, 20, 0, 10, 500, 6.5),
       ("Whole wheat pasta (dry)", "Pantry", "g", 100, 348, 14, 72, 1.4, 500, 2), ("Canned tomatoes", "Pantry", "g", 100, 32, 1.6, 7, 0.3, 400, 1.2),
       ("Onion", "Produce", "pcs", 1, 44, 1.2, 10, 0.1, 1, 0.5), ("Garlic (clove)", "Produce", "pcs", 1, 4.5, 0.2, 1, 0, 10, 0.6),
       ("Bell pepper", "Produce", "pcs", 1, 31, 1, 7, 0.3, 1, 1), ("Black beans (canned)", "Pantry", "g", 100, 132, 8.9, 23.7, 0.5, 425, 1.1),
       ("Whole wheat tortilla", "Bakery", "pcs", 1, 130, 4, 22, 3.5, 8, 3.5), ("Cheddar cheese", "Dairy & eggs", "g", 100, 403, 25, 1.3, 33, 225, 4),
       ("Avocado", "Produce", "pcs", 1, 240, 3, 13, 22, 1, 1.3), ("Mixed salad greens", "Produce", "g", 100, 15, 1.4, 2.9, 0.2, 150, 3.5),
       ("Tuna (canned in water)", "Pantry", "g", 100, 116, 26, 0, 1, 120, 1.5), ("Cucumber", "Produce", "pcs", 1, 45, 2, 11, 0.3, 1, 0.8),
       ("Firm tofu", "Refrigerated", "g", 100, 144, 15.6, 2.8, 8.7, 400, 2.5), ("Soy sauce", "Spices & sauces", "ml", 100, 53, 8, 5, 0.6, 250, 2.5),
       ("Frozen mixed vegetables", "Frozen", "g", 100, 65, 2.9, 13, 0.5, 500, 2), ("Apple", "Produce", "pcs", 1, 95, 0.5, 25, 0.3, 1, 0.6),
       ("Almonds", "Pantry", "g", 100, 579, 21, 22, 50, 300, 6), ("Whey protein powder", "Pantry", "g", 100, 400, 80, 8, 6, 900, 30),
       ("Honey", "Pantry", "g", 100, 304, 0.3, 82, 0, 340, 5), ("Lemon", "Produce", "pcs", 1, 17, 0.6, 5.4, 0.2, 1, 0.6),
       ("Quinoa (dry)", "Pantry", "g", 100, 368, 14, 64, 6, 500, 5), ("Feta cheese", "Dairy & eggs", "g", 100, 264, 14, 4, 21, 200, 4),
       ("Cherry tomatoes", "Produce", "g", 100, 18, 0.9, 3.9, 0.2, 300, 3), ("Parmesan", "Dairy & eggs", "g", 100, 431, 38, 4, 29, 150, 5),
       ("Butter", "Dairy & eggs", "g", 100, 717, 0.9, 0.1, 81, 250, 4), ("Potatoes", "Produce", "g", 100, 77, 2, 17, 0.1, 2000, 3.5),
       ("Carrots", "Produce", "g", 100, 41, 0.9, 10, 0.2, 1000, 1.5), ("Hummus", "Refrigerated", "g", 100, 166, 7.9, 14, 9.6, 280, 3.5),
       ("Turkey breast slices", "Meat & fish", "g", 100, 104, 17, 4, 2, 200, 4.5), ("Chia seeds", "Pantry", "g", 100, 486, 17, 42, 31, 300, 6),
       ("Granola", "Pantry", "g", 100, 471, 10, 64, 20, 400, 4.5), ("Cottage cheese", "Dairy & eggs", "g", 100, 98, 11, 3.4, 4.3, 450, 3.5),
       ("Lentils (dry)", "Pantry", "g", 100, 352, 25, 63, 1, 500, 2.5), ("Coconut milk (canned)", "Pantry", "ml", 100, 197, 2, 3, 21, 400, 2),
       ("Curry paste", "Spices & sauces", "g", 100, 120, 2, 12, 7, 120, 3), ("Shrimp", "Frozen", "g", 100, 99, 24, 0.2, 0.3, 450, 9)]

# recipe: (name, meal type, servings, prep min, note, [(ingredient, qty)])
RECIPES = [("Overnight Oats with Berries", "Breakfast", 1, 5, "Mix the night before, keep in the fridge.",
            [("Rolled oats", 50), ("Milk (2%)", 150), ("Greek yogurt (plain, 2%)", 75), ("Blueberries", 75), ("Chia seeds", 10), ("Honey", 10)]),
           ("Veggie Scrambled Eggs on Toast", "Breakfast", 1, 10, "",
            [("Eggs", 3), ("Spinach", 40), ("Bell pepper", 0.5), ("Whole wheat bread (slice)", 2), ("Butter", 5)]),
           ("Greek Yogurt Parfait", "Breakfast", 1, 3, "", [("Greek yogurt (plain, 2%)", 200), ("Granola", 40), ("Blueberries", 75), ("Honey", 10)]),
           ("Peanut Butter Banana Toast", "Breakfast", 1, 5, "", [("Whole wheat bread (slice)", 2), ("Peanut butter", 30), ("Banana", 1)]),
           ("Chicken, Rice & Broccoli Bowls", "Lunch", 4, 35, "Meal prep: 4 boxes for the week.",
            [("Chicken breast", 600), ("Brown rice (dry)", 280), ("Broccoli", 500), ("Olive oil", 30), ("Soy sauce", 40), ("Garlic (clove)", 3)]),
           ("Turkey & Hummus Wrap", "Lunch", 1, 5, "", [("Whole wheat tortilla", 1), ("Turkey breast slices", 80), ("Hummus", 40), ("Mixed salad greens", 30), ("Cucumber", 0.5)]),
           ("Tuna Salad Bowl", "Lunch", 2, 10, "",
            [("Tuna (canned in water)", 240), ("Mixed salad greens", 150), ("Cherry tomatoes", 150), ("Cucumber", 1), ("Avocado", 1), ("Olive oil", 15), ("Lemon", 1)]),
           ("Quinoa Feta Salad", "Lunch", 3, 20, "Keeps 3 days in the fridge.",
            [("Quinoa (dry)", 180), ("Feta cheese", 100), ("Cherry tomatoes", 200), ("Cucumber", 1), ("Spinach", 60), ("Olive oil", 30), ("Lemon", 1)]),
           ("Baked Salmon with Sweet Potato", "Dinner", 2, 30, "200°C / 400°F for 15-18 min.",
            [("Salmon fillet", 300), ("Sweet potato", 400), ("Broccoli", 250), ("Olive oil", 15), ("Lemon", 1)]),
           ("Beef & Bean Chili", "Dinner", 6, 45, "Freezes well.",
            [("Ground beef (90% lean)", 500), ("Black beans (canned)", 850), ("Canned tomatoes", 800), ("Onion", 1), ("Bell pepper", 2), ("Garlic (clove)", 3), ("Olive oil", 15)]),
           ("Spaghetti Bolognese", "Dinner", 4, 40, "",
            [("Whole wheat pasta (dry)", 400), ("Ground beef (90% lean)", 500), ("Canned tomatoes", 800), ("Onion", 1), ("Carrots", 150), ("Garlic (clove)", 3), ("Parmesan", 40), ("Olive oil", 15)]),
           ("Tofu Veggie Stir-fry", "Dinner", 3, 25, "",
            [("Firm tofu", 400), ("Frozen mixed vegetables", 500), ("Brown rice (dry)", 210), ("Soy sauce", 45), ("Garlic (clove)", 3), ("Olive oil", 20)]),
           ("Chicken Fajitas", "Dinner", 4, 25, "",
            [("Chicken breast", 600), ("Bell pepper", 3), ("Onion", 1), ("Whole wheat tortilla", 8), ("Cheddar cheese", 80), ("Avocado", 1), ("Olive oil", 20)]),
           ("Red Lentil Coconut Curry", "Dinner", 4, 35, "Vegan.",
            [("Lentils (dry)", 250), ("Coconut milk (canned)", 400), ("Canned tomatoes", 400), ("Curry paste", 40), ("Onion", 1), ("Spinach", 100), ("Brown rice (dry)", 280)]),
           ("Apple & Almonds", "Snack", 1, 1, "", [("Apple", 1), ("Almonds", 25)]),
           ("Protein Shake", "Snack", 1, 2, "", [("Whey protein powder", 30), ("Milk (2%)", 300), ("Banana", 1)]),
           ("Hummus & Carrot Sticks", "Snack", 2, 5, "", [("Hummus", 120), ("Carrots", 200)]),
           ("Cottage Cheese & Berries", "Snack", 1, 2, "", [("Cottage cheese", 200), ("Blueberries", 75)])]

PLAN = [("Overnight Oats with Berries", "Chicken, Rice & Broccoli Bowls", "Beef & Bean Chili", "Apple & Almonds"),
        ("Veggie Scrambled Eggs on Toast", "Chicken, Rice & Broccoli Bowls", "Baked Salmon with Sweet Potato", "Protein Shake"),
        ("Overnight Oats with Berries", "Beef & Bean Chili", "Tofu Veggie Stir-fry", "Hummus & Carrot Sticks"),
        ("Greek Yogurt Parfait", "Chicken, Rice & Broccoli Bowls", "Spaghetti Bolognese", "Apple & Almonds"),
        ("Overnight Oats with Berries", "Chicken, Rice & Broccoli Bowls", "Chicken Fajitas", "Cottage Cheese & Berries"),
        ("Peanut Butter Banana Toast", "Tuna Salad Bowl", "Red Lentil Coconut Curry", "Protein Shake"),
        ("Veggie Scrambled Eggs on Toast", "Quinoa Feta Salad", "Spaghetti Bolognese", "Hummus & Carrot Sticks")]


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); wp = wb.create_sheet("Weekly Plan"); gl = wb.create_sheet("Grocery List")
    rc = wb.create_sheet("Recipes"); ri = wb.create_sheet("Recipe Ingredients"); ig = wb.create_sheet("Ingredients"); s = wb.create_sheet("Settings")

    # ---------- Settings ----------
    K.title(s, "Settings", "Yellow cells are yours to edit.")
    rows = [("Currency label", "USD", None), ("Daily calorie target (kcal)", 2000, "#,##0"), ("Daily protein target (g)", 140, "0"),
            ("Daily carbs target (g)", 200, "0"), ("Daily fat target (g)", 70, "0")]
    for i, (k_, v, fmt) in enumerate(rows, 4):
        s.cell(row=i, column=1, value=k_).font = font(bold=True)
        c = s.cell(row=i, column=2); input_cell(c, v)
        if fmt: c.number_format = fmt
    s["C5"] = "Targets are per person per day. Not sure? A dietitian or an online TDEE calculator can help."; s["C5"].font = font(italic=True, size=9)
    K.hdr(s, 4, ["Store sections (grocery order)"], col0=8)
    for i in range(12):
        c = s.cell(row=5 + i, column=8, value=SECTIONS[i] or None); input_cell(c)
    K.hdr(s, 4, ["Meal types"], col0=10)
    for i, m in enumerate(MEALTYPES):
        c = s.cell(row=5 + i, column=10, value=m); c.font = font(size=9); c.border = B
    K.hdr(s, 4, ["Units"], col0=12)
    for i, u in enumerate(["g", "ml", "pcs"]):
        c = s.cell(row=5 + i, column=12, value=u); c.font = font(size=9); c.border = B
    for c, w in zip("ABCDEFGHIJKL", [30, 12, 3, 3, 3, 3, 3, 26, 3, 12, 3, 8]): s.column_dimensions[c].width = w
    s["A11"] = "Reorder the store sections to match the aisles of your shop – the Grocery List follows this order."; s["A11"].font = font(italic=True, size=9)

    # ---------- Ingredients ----------
    K.title(ig, "Ingredients", "Your ingredient list with nutrition and prices. 50 common foods are already filled in – edit or add your own.")
    ig["A3"] = '="Prices in "&Settings!B4&" · nutrition per 100 g, per 100 ml or per 1 piece (column D)"'; ig["A3"].font = font(bold=True)
    heads = ["Ingredient", "Store section", "Unit", "Nutrition per (100 or 1)", "Calories", "Protein (g)", "Carbs (g)", "Fat (g)",
             "Pack size (in unit)", "Pack price", "Have at home (in unit)", "Price per unit", "Needed this week", "To buy", "Packs to buy", "Est. cost", "Sort key"]
    K.hdr(ig, 5, heads, [26, 14, 6, 10, 8, 8, 8, 8, 9, 9, 10, 9, 10, 9, 8, 9, 8])
    style_range(ig, I0, I1, 1, 17, inputs=range(1, 12), fmt={5: G, 6: G, 7: G, 8: G, 9: G, 10: MONEY, 11: G, 12: "0.000", 13: '#,##0.#;-#,##0.#;""', 14: '#,##0.#;-#,##0.#;""', 15: '0;-0;""', 16: CUR})
    for i, row in enumerate(ING):
        for j, v in enumerate(row, 1): ig.cell(row=I0 + i, column=j, value=v)
    if sample:
        for nm, have in (("Olive oil", 400), ("Soy sauce", 150), ("Honey", 200), ("Garlic (clove)", 4), ("Rolled oats", 300), ("Brown rice (dry)", 500), ("Whey protein powder", 600), ("Almonds", 200), ("Peanut butter", 300), ("Chia seeds", 150), ("Granola", 200)):
            ig.cell(row=I0 + [x[0] for x in ING].index(nm), column=11, value=have)
    for r in range(I0, I1 + 1):
        ig[f"L{r}"] = f'=IF(OR(N(I{r})=0,N(J{r})=0),0,J{r}/I{r})'
        ig[f"M{r}"] = f'=IF(A{r}="",0,SUMIFS({RI}!$K${L0}:$K${L1},{RI}!$B${L0}:$B${L1},A{r}))'
        ig[f"N{r}"] = f"=MAX(0,ROUND(M{r}-N(K{r}),1))"
        ig[f"O{r}"] = f"=IF(N{r}=0,0,IF(N(I{r})=0,1,ROUNDUP(N{r}/I{r}-1E-9,0)))"
        ig[f"P{r}"] = f"=O{r}*N(J{r})"
        ig[f"Q{r}"] = f'=IF(N{r}=0,"",IFERROR(MATCH(B{r},Settings!$H$5:$H$16,0),12)*1000+ROW()-{I0 - 1})'
    for ref, f1 in ((f"B{I0}:B{I1}", "=Settings!$H$5:$H$16"), (f"C{I0}:C{I1}", "=Settings!$L$5:$L$7")):
        dv = DataValidation(type="list", formula1=f1, allow_blank=True); ig.add_data_validation(dv); dv.add(ref)
    ig.freeze_panes = "B6"
    IA = f"Ingredients!$A${I0}:$A${I1}"
    def icol(c): return f"Ingredients!${c}${I0}:${c}${I1}"

    # ---------- Recipes ----------
    K.title(rc, "Recipes", "One row per recipe. Then list its ingredients on the Recipe Ingredients tab. Grey columns calculate themselves.")
    heads = ["Recipe", "Meal type", "Servings (recipe makes)", "Prep time (min)", "Notes / link", "Calories per serving", "Protein (g)",
             "Carbs (g)", "Fat (g)", "Cost per serving", "Portions planned this week", "Ingredient lines"]
    K.hdr(rc, 5, heads, [32, 11, 10, 9, 34, 10, 9, 9, 9, 10, 11, 9])
    style_range(rc, R0, R1, 1, 12, inputs=(1, 2, 3, 4, 5), fmt={6: '#,##0;-#,##0;""', 7: G, 8: G, 9: G, 10: CUR, 11: '0;-0;""', 12: '0;-0;""'})
    for i, (nm, mt, sv, pm, note, _) in enumerate(RECIPES):
        r = R0 + i
        rc[f"A{r}"] = nm; rc[f"B{r}"] = mt; rc[f"C{r}"] = sv; rc[f"D{r}"] = pm
        if note: rc[f"E{r}"] = note
    for r in range(R0, R1 + 1):
        for c, src in zip("FGHIJ", "FGHIJ"):
            rc[f"{c}{r}"] = f'=IF(OR(A{r}="",N(C{r})=0),"",SUMIFS({RI}!${src}${L0}:${src}${L1},{RI}!$A${L0}:$A${L1},A{r})/C{r})'
        rc[f"K{r}"] = (f'=IF(A{r}="",0,SUMIFS({WP}!$R${P0}:$R${P1},{WP}!$C${P0}:$C${P1},A{r})+SUMIFS({WP}!$S${P0}:$S${P1},{WP}!$E${P0}:$E${P1},A{r})'
                       f'+SUMIFS({WP}!$T${P0}:$T${P1},{WP}!$G${P0}:$G${P1},A{r})+SUMIFS({WP}!$U${P0}:$U${P1},{WP}!$I${P0}:$I${P1},A{r}))')
        rc[f"L{r}"] = f'=IF(A{r}="",0,COUNTIF({RI}!$A${L0}:$A${L1},A{r}))'
    dv = DataValidation(type="list", formula1="=Settings!$J$5:$J$9", allow_blank=True); rc.add_data_validation(dv); dv.add(f"B{R0}:B{R1}")
    rc.conditional_formatting.add(f"L{R0}:L{R1}", FormulaRule(formula=[f'AND($A{R0}<>"",L{R0}=0)'], fill=fill("FEE2E2")))
    rc["A87"] = "Red 'Ingredient lines' = this recipe has no ingredients yet. Calories, cost and the grocery list need them."
    rc["A87"].font = font(italic=True, size=9)
    rc.freeze_panes = "B6"

    # ---------- Recipe Ingredients ----------
    K.title(ri, "Recipe Ingredients", "One row per ingredient of a recipe. Quantity is for the WHOLE recipe, in the ingredient's unit (g, ml or pcs).")
    heads = ["Recipe", "Ingredient", "Quantity", "Note (chopped, cooked…)", "Unit", "Calories", "Protein (g)", "Carbs (g)", "Fat (g)", "Cost",
             "Needed this week", "Ingredient #", "Week factor"]
    K.hdr(ri, 5, heads, [32, 26, 9, 20, 6, 9, 8, 8, 8, 8, 10, 8, 8])
    style_range(ri, L0, L1, 1, 13, inputs=(1, 2, 3, 4), fmt={3: G, 6: '#,##0;-#,##0;""', 7: G, 8: G, 9: G, 10: CUR, 11: '#,##0.#;-#,##0.#;""', 13: "0.00"})
    n = 0
    for nm, _, _, _, _, items in RECIPES:
        for ing, q in items:
            r = L0 + n; ri[f"A{r}"] = nm; ri[f"B{r}"] = ing; ri[f"C{r}"] = q; n += 1
    for r in range(L0, L1 + 1):
        ri[f"L{r}"] = f'=IF(B{r}="",0,IFERROR(MATCH(B{r},{IA},0),0))'
        ri[f"E{r}"] = f'=IF(L{r}=0,IF(B{r}="","","not in list"),INDEX({icol("C")},L{r}))'
        for c, src in zip("FGHI", "EFGH"):
            ri[f"{c}{r}"] = f'=IF(OR(L{r}=0,N(C{r})=0),0,IFERROR(C{r}/INDEX({icol("D")},L{r})*N(INDEX({icol(src)},L{r})),0))'
        ri[f"J{r}"] = f'=IF(OR(L{r}=0,N(C{r})=0),0,C{r}*INDEX({icol("L")},L{r}))'
        ri[f"M{r}"] = (f'=IF(A{r}="",0,IFERROR(INDEX(Recipes!$K${R0}:$K${R1},MATCH(A{r},Recipes!$A${R0}:$A${R1},0))/'
                       f'INDEX(Recipes!$C${R0}:$C${R1},MATCH(A{r},Recipes!$A${R0}:$A${R1},0)),0))')
        ri[f"K{r}"] = f"=N(C{r})*M{r}"
    dv = DataValidation(type="list", formula1=f"=Recipes!$A${R0}:$A${R1}", allow_blank=True); ri.add_data_validation(dv); dv.add(f"A{L0}:A{L1}")
    dv = DataValidation(type="list", formula1=f"={IA}", allow_blank=True); ri.add_data_validation(dv); dv.add(f"B{L0}:B{L1}")
    ri.conditional_formatting.add(f"E{L0}:E{L1}", FormulaRule(formula=[f'E{L0}="not in list"'], fill=fill("FEE2E2"), font=Font(name=F, bold=True, color="B91C1C")))
    ri.freeze_panes = "C6"

    # ---------- Weekly Plan ----------
    K.title(wp, "Weekly Plan", "Pick a recipe for each meal. Portions = how many servings you cook for that meal (blank = 1). Totals are per person.")
    wp["A3"] = "Week starting:"; wp["A3"].font = font(bold=True)
    input_cell(wp["C3"], dt.date(2026, 11, 2) if sample else "=TODAY()-WEEKDAY(TODAY(),2)+1"); wp["C3"].number_format = "ddd dd mmm yyyy"
    wp["E3"] = "Tip: for a new week, change this date and pick new meals (or copy this tab first to keep the old plan)."; wp["E3"].font = font(italic=True, size=9)
    heads = ["Day", "Date", "Breakfast", "Portions", "Lunch", "Portions", "Dinner", "Portions", "Snack", "Portions",
             "Calories", "Protein (g)", "Carbs (g)", "Fat (g)", "Food cost", "vs calorie target", "", "B port.", "L port.", "D port.", "S port."]
    K.hdr(wp, 6, heads, [8, 9, 28, 8, 28, 8, 28, 8, 24, 8, 9, 9, 9, 9, 9, 11, 2, 6, 6, 6, 6])
    style_range(wp, P0, P1, 1, 16, inputs=(3, 4, 5, 6, 7, 8, 9, 10), fmt={2: "dd mmm", 11: "#,##0", 12: "0", 13: "0", 14: "0", 15: CUR, 16: '+0%;-0%;0%'})
    style_range(wp, P0, P1, 18, 21)
    meals = [("C", "D", "R"), ("E", "F", "S"), ("G", "H", "T"), ("I", "J", "U")]
    RA = f"Recipes!$A${R0}:$A${R1}"
    for i in range(7):
        r = P0 + i
        wp[f"A{r}"] = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"][i]
        wp[f"B{r}"] = f"=$C$3+{i}"
        if sample:
            for (mc, pc, _), rec in zip(meals, PLAN[i]):
                wp[f"{mc}{r}"] = rec
            if i == 0: wp[f"H{r}"] = 2      # cook chili for dinner + next day's lunch
        for mc, pc, hc in meals:
            wp[f"{hc}{r}"] = f'=IF({mc}{r}="",0,IF(N({pc}{r})>0,N({pc}{r}),1))'
        for c, src in zip("KLMNO", "FGHIJ"):
            wp[f"{c}{r}"] = "=" + "+".join(f"IFERROR(N(INDEX(Recipes!${src}${R0}:${src}${R1},MATCH({mc}{r},{RA},0))),0)" for mc, _, _ in meals)
        wp[f"P{r}"] = f'=IF(OR(K{r}=0,N(Settings!$B$5)=0),"",K{r}/Settings!$B$5-1)'
        for mc, pc, _ in meals: wp[f"{pc}{r}"].alignment = Alignment(horizontal="center")
    wp["A14"] = "Daily average (planned days)"; wp["A15"] = "Your daily target"
    for c in "KLMNO":
        wp[f"{c}14"] = f'=IF(COUNTIF($K${P0}:$K${P1},">0")=0,0,SUMIF($K${P0}:$K${P1},">0",{c}{P0}:{c}{P1})/COUNTIF($K${P0}:$K${P1},">0"))'
    for c, sref in zip("KLMN", ("B5", "B6", "B7", "B8")): wp[f"{c}15"] = f"=Settings!{sref}"
    for r in (14, 15):
        for c in "AKLMNO":
            x = wp[f"{c}{r}"]; x.font = font(bold=True); x.fill = fill(SOFT); x.border = B
            x.number_format = {"K": "#,##0", "O": CUR}.get(c, "0")
    for mc, _, _ in meals:
        dv = DataValidation(type="list", formula1=f"={RA}", allow_blank=True); wp.add_data_validation(dv); dv.add(f"{mc}{P0}:{mc}{P1}")
    dv = DataValidation(type="decimal", operator="between", formula1="0", formula2="50", allow_blank=True); wp.add_data_validation(dv)
    for _, pc, _ in meals: dv.add(f"{pc}{P0}:{pc}{P1}")
    wp.conditional_formatting.add(f"P{P0}:P{P1}", FormulaRule(formula=[f'AND(P{P0}<>"",ABS(P{P0})<=0.1)'], fill=fill("D1FAE5"), font=Font(name=F, color="166534")))
    wp.conditional_formatting.add(f"P{P0}:P{P1}", FormulaRule(formula=[f'AND(P{P0}<>"",ABS(P{P0})>0.1)'], fill=fill("FEF3C7"), font=Font(name=F, color="92400E")))
    wp["A17"] = "Food cost = cost of one serving of each meal. Grocery quantities follow the portions: 2 portions of a 6-serving recipe buys 1/3 of its ingredients."
    wp["A18"] = "Green = within 10% of your calorie target, amber = more than 10% above or below."
    for a in ("A17", "A18"): wp[a].font = font(italic=True, size=9)
    wp.freeze_panes = "C7"

    # ---------- Grocery List ----------
    K.title(gl, "Grocery List", "Built from your Weekly Plan, minus what you have at home (Ingredients tab). Grouped by store section.")
    gl["A3"] = f'="Week of "&TEXT({WP}!$C$3,"dd mmm yyyy")&"  ·  estimated total: "&Settings!$B$4&" "&TEXT(H3,"#,##0.00")'; gl["A3"].font = font(bold=True)
    gl["H3"] = f"=SUM(H{G0}:H{G0 + NG - 1})+SUM(M{G0}:M{G0 + 19})"; gl["H3"].number_format = CUR; gl["H3"].font = Font(name=F, bold=True, color=MAIN, size=12)
    K.hdr(gl, 6, ["✓", "Section", "Item", "Amount needed", "Unit", "Packs to buy", "Pack size", "Est. cost"], [5, 14, 28, 10, 6, 8, 8, 10])
    keys = icol("Q")
    for k in range(NG):
        r = G0 + k
        kk = f"SMALL({keys},{k + 1})"
        idx = f"MOD({kk},1000)"
        gl[f"B{r}"] = f'=IFERROR(IF(INDEX(Settings!$H$5:$H$16,INT({kk}/1000))="","Other",INDEX(Settings!$H$5:$H$16,INT({kk}/1000))),"")'
        gl[f"C{r}"] = f'=IFERROR(INDEX({icol("A")},{idx}),"")'
        for c, src in zip("DEFGH", "NCOIP"):
            gl[f"{c}{r}"] = f'=IF($C{r}="","",INDEX({icol(src)},{idx}))'
    style_range(gl, G0, G0 + NG - 1, 1, 8, inputs=(1,), fmt={4: '#,##0.#', 6: "0", 7: G, 8: CUR})
    for r in range(G0, G0 + NG):
        gl[f"A{r}"].alignment = Alignment(horizontal="center"); gl[f"F{r}"].alignment = Alignment(horizontal="center")
    gl.conditional_formatting.add(f"A{G0}:H{G0 + NG - 1}", FormulaRule(formula=[f'$A{G0}<>""'], font=Font(name=F, color="9CA3AF", strike=True)))
    gl.conditional_formatting.add(f"B{G0}:B{G0 + NG - 1}", FormulaRule(formula=[f'AND($B{G0}<>"",$B{G0}=$B{G0 - 1})'], font=Font(name=F, color="FFFFFF")))
    K.section(gl, "J5", "Extra items (not in a recipe)")
    K.hdr(gl, 6, ["✓", "Item", "Qty", "Est. cost"], [5, 24, 6, 10], col0=10)
    style_range(gl, G0, G0 + 19, 10, 13, inputs=(10, 11, 12, 13), fmt={13: CUR})
    if sample:
        for i, (it, q, c) in enumerate([("Paper towels", 1, 4.5), ("Dish soap", 1, 3), ("Coffee", 1, 8.5)]):
            gl[f"K{G0 + i}"] = it; gl[f"L{G0 + i}"] = q; gl[f"M{G0 + i}"] = c
    gl.conditional_formatting.add(f"J{G0}:M{G0 + 19}", FormulaRule(formula=[f'$J{G0}<>""'], font=Font(name=F, color="9CA3AF", strike=True)))
    gl["J28"] = "Type x in the ✓ column when an item is in your cart. Clear the ✓ column for a new week."
    gl["J29"] = "Packs are rounded up; amounts are in the ingredient's unit."
    for a in ("J28", "J29"): gl[a].font = font(italic=True, size=9)
    gl.freeze_panes = "A7"

    # ---------- Dashboard ----------
    K.title(dash, "Meal Plan Dashboard", "Your week at a glance. Updates by itself from the Weekly Plan.")
    dash["A3"] = f'="Week of "&TEXT({WP}!$C$3,"dd mmm yyyy")&"  ·  amounts in "&Settings!B4&"  ·  targets: "&TEXT(Settings!B5,"#,##0")&" kcal, "&Settings!B6&" g protein"'
    dash["A3"].font = font(bold=True)
    meal_cells = ",".join(f"{WP}!${c}${P0}:${c}${P1}" for c in "CEGI")
    kp = [("Meals planned", f'=COUNTA({meal_cells})&" of 28"', "@"),
          ("Avg calories / day", f"={WP}!K14", "#,##0"),
          ("Avg protein / day (g)", f"={WP}!L14", "0"),
          ("Grocery list total", "='Grocery List'!H3", '#,##0.00'),
          ("Items to buy", f'=COUNTIF({icol("N")},">0")', "0"),
          ("Avg food cost / day", f"={WP}!O14", '#,##0.00')]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 3) * 3, 1 + (i % 3) * 2, lab, f, fmt)
        dash.cell(row=6 + (i // 3) * 3, column=1 + (i % 3) * 2).alignment = Alignment(horizontal="left")
    for c, w in zip("ABCDEFGHIJ", [14, 31, 31, 31, 26, 10, 9, 9, 9, 9]): dash.column_dimensions[c].width = w
    K.section(dash, "A11", "This week")
    K.hdr(dash, 12, ["Day", "Breakfast", "Lunch", "Dinner", "Snack", "Calories", "Protein", "Carbs", "Fat", "vs target"])
    for i in range(7):
        r = 13 + i; pr = P0 + i
        dash[f"A{r}"] = f'={WP}!A{pr}&" "&TEXT({WP}!B{pr},"d mmm")'
        for c, src in zip("BCDE", "CEGI"): dash[f"{c}{r}"] = f'=IF({WP}!{src}{pr}="","—",{WP}!{src}{pr})'
        for c, src in zip("FGHIJ", "KLMNP"): dash[f"{c}{r}"] = f"={WP}!{src}{pr}"
    dash["A20"] = "Daily average"; dash["A21"] = "Target"
    for c, src in zip("FGHI", "KLMN"):
        dash[f"{c}20"] = f"={WP}!{src}14"; dash[f"{c}21"] = f"={WP}!{src}15"
    style_range(dash, 13, 21, 1, 10, fmt={6: "#,##0", 7: "0", 8: "0", 9: "0", 10: '+0%;-0%;0%'})
    for r in (20, 21):
        for c in range(1, 11): dash.cell(row=r, column=c).font = font(bold=True); dash.cell(row=r, column=c).fill = fill(SOFT)
    dash.conditional_formatting.add("J13:J19", FormulaRule(formula=['AND(J13<>"",ABS(J13)<=0.1)'], fill=fill("D1FAE5"), font=Font(name=F, color="166534")))
    dash.conditional_formatting.add("J13:J19", FormulaRule(formula=['AND(J13<>"",ABS(J13)>0.1)'], fill=fill("FEF3C7"), font=Font(name=F, color="92400E")))
    c1 = BarChart(); c1.type = "col"; c1.title = "Calories per day"; c1.height = 7.5; c1.width = 15; c1.legend = None
    c1.add_data(Reference(dash, min_col=6, min_row=12, max_row=19), titles_from_data=True)
    c1.set_categories(Reference(dash, min_col=1, min_row=13, max_row=19))
    c1.series[0].graphicalProperties.solidFill = MAIN; c1.y_axis.majorGridlines = None; c1.y_axis.scaling.min = 0
    dash.add_chart(c1, "A24")
    c2 = BarChart(); c2.type = "col"; c2.grouping = "stacked"; c2.overlap = 100; c2.title = "Protein, carbs & fat per day (g)"; c2.height = 7.5; c2.width = 15
    c2.add_data(Reference(dash, min_col=7, max_col=9, min_row=12, max_row=19), titles_from_data=True)
    c2.set_categories(Reference(dash, min_col=1, min_row=13, max_row=19))
    for ser, col in zip(c2.series, (MAIN, ACC, "FBBF24")): ser.graphicalProperties.solidFill = col
    c2.y_axis.majorGridlines = None
    dash.add_chart(c2, "D24")

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Settings tab: set your currency and daily targets for calories, protein, carbs and fat.", False),
             ("2.  Ingredients tab: 50 common foods with nutrition and typical prices are ready. Edit prices to match your shop and add your own foods.", False),
             ("3.  Recipes tab: name each recipe and how many servings it makes. Recipe Ingredients tab: list what goes in it (whole-recipe amounts).", False),
             ("4.  Weekly Plan tab: pick a recipe for breakfast, lunch, dinner and snack each day. Portions = servings you will cook (blank = 1).", False),
             ("5.  Grocery List tab: everything you need for the week, grouped by store section, minus what you have at home. Tick items off as you shop.", False),
             ("6.  Dashboard: calories and macros per day vs your targets, grocery total and food cost.", False),
             ("", False), ("How the numbers work", True),
             ("Nutrition is per 100 g, per 100 ml or per piece (column D on the Ingredients tab). Recipes add up their ingredients and divide by servings.", False),
             ("Grocery amounts scale each recipe to the portions you planned, then subtract 'Have at home'. Packs are rounded up.", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  The nutrition values that come with the template are typical averages – check your product labels for exact numbers.", False),
             ("•  Room for 200 ingredients, 80 recipes and 800 recipe lines.", False),
             ("•  On your phone: open drive.google.com, upload this file, then choose Open with > Google Sheets – handy at the supermarket.", False),
             ("•  This planner is a tool to organise meals, not medical or dietary advice.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("18 example recipes and a full week of meals (week of 2 Nov 2026). Use the blank copy in your download for your own plan – it already includes the 50 ingredients.", False)]
    K.start(st, "Meal Planner & Grocery List", "Plan your week, see calories and macros per day, and get your grocery list automatically. Works in Excel and Google Sheets.", lines)
    finish(wb, sample, path)


META = {
    "name": "Meal Planner & Grocery List with Macro Tracker – Excel & Google Sheets",
    "file_base": "Meal-Planner-Grocery-List",
    "price": 14,
    "color": "#0F766E",
    "title": "Meal Planner & Grocery List",
    "tagline": "Weekly menu · calories & macros · automatic grocery list — Excel & Google Sheets",
    "intro": "Plan a week of meals in minutes. Pick recipes for each day and the planner adds up calories, protein, carbs and fat, compares them with your targets and builds your grocery list automatically – grouped by store section.",
    "features": [
        "Weekly Plan: breakfast, lunch, dinner and snack for 7 days, with portions for meal prep and families",
        "Daily calories, protein, carbs and fat vs your targets, plus food cost per day",
        "Automatic grocery list: only what you still need, minus what you have at home, grouped by store section, with packs and cost",
        "Recipe book: calories, macros and cost per serving calculated from the ingredients",
        "50 common ingredients pre-loaded with nutrition and typical prices – add your own",
        "Dashboard with charts of calories and macros per day",
        "Works for weight loss, muscle gain, budgeting or family meal planning; phone-friendly in Google Sheets",
    ],
    "tags": ["meal planner", "grocery list", "macro tracker", "meal prep", "google sheets", "excel template", "weekly meal plan",
             "calorie tracker", "recipe organizer", "budget meal planner", "nutrition tracker", "shopping list", "family meal planner"],
}
