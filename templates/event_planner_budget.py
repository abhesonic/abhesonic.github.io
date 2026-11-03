"""Event Planner & Budget – parties, conferences, fundraisers, corporate events.

Settings (event, date, guests, budget per category) -> Budget (expense lines, estimate/actual/paid, due dates)
-> Income (tickets, sponsors) -> Vendors -> Guests (RSVP) -> Tasks (checklist counted back from the event date)
-> Dashboard (countdown, budget vs actual, profit/loss, cost per guest, RSVP, tasks, payments due, charts).
"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.utils import get_column_letter as L
from _common import Kit, F, B, MONEY, DATE, font, fill, style_range, input_cell, finish

MAIN = "BE185D"; SOFT = "FCE7F1"
K = Kit(MAIN, SOFT)
PCT = "0.0%"
CATS = ["Venue", "Catering & food", "Drinks & bar", "Decor & flowers", "Entertainment & speakers", "AV, tech & lighting",
        "Marketing & invitations", "Staff & security", "Printing & signage", "Gifts, swag & prizes", "Transport & parking",
        "Rentals (tables, chairs, tents)", "Photo & video", "Other"]
NC = len(CATS); C0 = 14; C1 = C0 + NC - 1           # Settings rows for categories
NB = 100; B0 = 6; B1 = B0 + NB - 1                  # Budget rows
NI = 30; I0 = 6; I1 = I0 + NI - 1                   # Income rows
NV = 40; V0 = 6; V1 = V0 + NV - 1                   # Vendor rows
NG = 300; G0 = 6; G1 = G0 + NG - 1                  # Guest rows
NT = 80; T0 = 6; T1 = T0 + NT - 1                   # Task rows
ED = "Settings!$B$5"
SRC = ["Tickets", "Sponsors", "Donations", "Vendor / exhibitor fees", "Raffle & auction", "Merchandise", "Other"]
BAR = 20


def bar(ref):
    return f'REPT("█",ROUND(MIN(1,MAX(0,N({ref})))*{BAR},0))&REPT("░",{BAR}-ROUND(MIN(1,MAX(0,N({ref})))*{BAR},0))'


TASKS = [  # task, category, weeks before
    ("Set the goal, date, guest count and total budget", "Planning", 16), ("Choose a theme / format and agenda outline", "Planning", 15),
    ("Shortlist and visit venues", "Venue", 14), ("Book the venue and sign the contract", "Venue", 12),
    ("Open a budget and track every quote here", "Planning", 12), ("Book caterer and plan the menu", "Catering & food", 10),
    ("Book speakers / entertainment", "Entertainment & speakers", 10), ("Find sponsors and send sponsor packages", "Sponsors", 10),
    ("Create the guest list", "Guests", 9), ("Book photographer / videographer", "Photo & video", 8),
    ("Book AV, sound and lighting", "AV, tech & lighting", 8), ("Order rentals (tables, chairs, tents, linens)", "Rentals", 7),
    ("Launch ticket sales / registration page", "Marketing", 7), ("Send invitations or save-the-dates", "Guests", 6),
    ("Plan decor and order flowers", "Decor & flowers", 6), ("Promote on social media and email", "Marketing", 5),
    ("Order printing: signs, badges, programs", "Printing & signage", 4), ("Order gifts, swag or prizes", "Gifts & swag", 4),
    ("Arrange transport and parking", "Transport", 4), ("Book staff, volunteers and security", "Staff", 3),
    ("Confirm final menu and dietary needs", "Catering & food", 2), ("Chase RSVPs and give final numbers to caterer", "Guests", 2),
    ("Write the run-of-show (minute-by-minute)", "Planning", 2), ("Confirm every vendor: time, place, contact", "Vendors", 1),
    ("Pay remaining vendor balances", "Budget", 1), ("Prepare seating / name badges", "Guests", 1),
    ("Pack the event-day kit (tape, chargers, first aid, pens)", "Planning", 0), ("Set up the venue and test AV", "Venue", 0),
    ("Send thank-you notes to guests, sponsors and vendors", "After", -1), ("Share photos and a short survey", "After", -1),
    ("Close the budget: final actuals and lessons learned", "After", -2)]

BUD = [  # category, item, vendor, estimate, actual, paid, due days before event
    ("Venue", "Hall rental (6 hours)", "Riverside Event Hall", 2400, 2400, 2400, 60),
    ("Venue", "Cleaning fee", "Riverside Event Hall", 250, 250, 0, 7),
    ("Catering & food", "Dinner buffet – 120 guests", "Fresh Table Catering", 4200, 4380, 1500, 10),
    ("Catering & food", "Desserts & cake", "Sweet Spot Bakery", 450, 420, 420, 5),
    ("Drinks & bar", "Open bar (3 hours)", "Fresh Table Catering", 1800, 1950, 0, 3),
    ("Decor & flowers", "Centrepieces & stage flowers", "Bloom & Co", 900, 840, 300, 14),
    ("Decor & flowers", "Balloons & backdrop", "Party Supply Hub", 300, 285, 285, 20),
    ("Entertainment & speakers", "DJ (4 hours)", "DJ Marco", 800, 800, 200, 30),
    ("Entertainment & speakers", "Keynote speaker fee", "Dr. A. Lewis", 1500, 1500, 750, 21),
    ("AV, tech & lighting", "Projector, mics, sound", "SoundPro Rentals", 650, 720, 0, 2),
    ("Marketing & invitations", "Online ads", "Social ads", 400, 365, 365, 45),
    ("Marketing & invitations", "Ticketing platform fees", "Ticketing site", 300, 288, 288, 0),
    ("Printing & signage", "Banners, programs, badges", "PrintFast", 380, 0, 0, 10),
    ("Gifts, swag & prizes", "Raffle prizes & goodie bags", "Various", 600, 0, 0, 7),
    ("Staff & security", "2 security guards", "SafeGuard", 480, 0, 0, 1),
    ("Rentals (tables, chairs, tents)", "Chairs & linens", "Party Supply Hub", 520, 520, 260, 14),
    ("Photo & video", "Photographer (4 hours)", "Lens & Light Studio", 700, 700, 200, 30),
    ("Transport & parking", "Shuttle bus", "City Shuttles", 350, 0, 0, 3)]

INC = [("Tickets", "Early-bird tickets", 60, 95, 60), ("Tickets", "Regular tickets", 50, 120, 31), ("Sponsors", "Gold sponsor – Bright Bank", 1, 3000, 1),
       ("Sponsors", "Silver sponsor – Nova Tech", 1, 1500, 0), ("Raffle & auction", "Raffle tickets", 150, 10, 0), ("Donations", "Online donations", 1, 800, 1)]

VEND = [("Riverside Event Hall", "Venue", "Sara Kim", "555-0142", "events@riverside.example", "Yes"),
        ("Fresh Table Catering", "Catering & bar", "Tom Reyes", "555-0187", "tom@freshtable.example", "Yes"),
        ("Bloom & Co", "Flowers", "Ivy Chen", "555-0110", "hello@bloom.example", "Yes"),
        ("DJ Marco", "Music", "Marco D.", "555-0199", "", "Yes"),
        ("SoundPro Rentals", "AV", "Dev Patel", "555-0121", "rent@soundpro.example", "No"),
        ("Party Supply Hub", "Rentals & decor", "", "555-0160", "", "Yes"),
        ("Lens & Light Studio", "Photography", "Nina B.", "555-0133", "", "Yes"),
        ("PrintFast", "Printing", "", "555-0175", "", "No"),
        ("SafeGuard", "Security", "", "555-0108", "", "No"),
        ("City Shuttles", "Transport", "", "555-0150", "", "No"),
        ("Sweet Spot Bakery", "Desserts", "", "555-0115", "", "Yes"),
        ("Dr. A. Lewis", "Keynote speaker", "Assistant: Jo", "555-0191", "", "Yes")]

FIRST = ["Alex", "Maria", "James", "Priya", "Chen", "Olivia", "Noah", "Fatima", "Liam", "Sofia", "Ethan", "Aisha", "Lucas", "Emma", "Mateo",
         "Grace", "Daniel", "Hana", "Ryan", "Zoe", "Omar", "Chloe", "Leo", "Mia", "Samuel", "Nora", "Jack", "Lily", "Ben", "Ava"]
LAST = ["Smith", "Garcia", "Brown", "Patel", "Wang", "Johnson", "Lee", "Khan", "Martin", "Rossi", "Clark", "Ali", "Silva", "Taylor", "Lopez"]


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); s = wb.create_sheet("Settings"); bud = wb.create_sheet("Budget")
    inc = wb.create_sheet("Income"); ven = wb.create_sheet("Vendors"); gst = wb.create_sheet("Guests"); tsk = wb.create_sheet("Tasks")
    edate = dt.date(2026, 12, 12)

    # ---------- Settings ----------
    K.title(s, "Settings", "Yellow cells are yours to edit.")
    rows = [("Event name", "Winter Charity Gala 2026" if sample else "My Event", None),
            ("Event date", edate if sample else "=TODAY()+90", DATE),
            ("Venue / place", "Riverside Event Hall" if sample else None, None),
            ("Expected guests / attendees", 120 if sample else 50, "0"),
            ("Total budget", 17500 if sample else 5000, MONEY),
            ("Currency label", "USD", None),
            ("Contingency (buffer) %", 0.10, PCT),
            ("Today (for countdown)", "=TODAY()", DATE)]
    for i, (k_, v, fmt) in enumerate(rows, 4):
        s.cell(row=i, column=1, value=k_).font = font(bold=True)
        c = s.cell(row=i, column=2)
        if i == 11:
            c.value = v; c.border = B; c.fill = fill("F3F4F6"); c.font = font()
        else:
            input_cell(c, v)
        if fmt: c.number_format = fmt
    notes = {5: "Due dates on the Tasks tab are counted back from this date.", 8: "Your spending limit for the whole event.",
             10: "Planners keep 5–15% aside for surprises. Shown on the Dashboard.", 11: "Calculated – nothing to type."}
    for r, t in notes.items(): s[f"C{r}"] = t; s[f"C{r}"].font = font(italic=True, size=9)
    K.section(s, "A13", "Budget per category")
    s["A13"].value = None
    s.cell(row=13, column=1, value="Category").font = font(bold=True); s.cell(row=13, column=2, value="Planned budget").font = font(bold=True)
    K.hdr(s, 13, ["Category", "Planned budget", "Share of total"])
    shares = [0.14, 0.26, 0.10, 0.06, 0.13, 0.04, 0.04, 0.03, 0.02, 0.04, 0.02, 0.03, 0.04, 0.05]
    for i, cat in enumerate(CATS):
        r = C0 + i
        input_cell(s[f"A{r}"], cat); input_cell(s[f"B{r}"], round(17500 * shares[i], -1) if sample else None)
        s[f"B{r}"].number_format = MONEY
        s[f"C{r}"] = f"=IFERROR(B{r}/$B$8,0)"; s[f"C{r}"].number_format = PCT; s[f"C{r}"].border = B; s[f"C{r}"].fill = fill("F3F4F6"); s[f"C{r}"].font = font()
    s[f"A{C1 + 1}"] = "Total of categories"; s[f"A{C1 + 1}"].font = font(bold=True)
    s[f"B{C1 + 1}"] = f"=SUM(B{C0}:B{C1})"; s[f"B{C1 + 1}"].number_format = MONEY; s[f"B{C1 + 1}"].font = font(bold=True)
    s[f"C{C1 + 1}"] = f'=IF(ABS(B{C1 + 1}-B8)<0.5,"✓ matches total budget",IF(B{C1 + 1}>B8,"Over total budget by "&TEXT(B{C1 + 1}-B8,"#,##0"),"Not yet planned: "&TEXT(B8-B{C1 + 1},"#,##0")))'
    s[f"C{C1 + 1}"].font = font(italic=True, size=9)
    s[f"A{C1 + 3}"] = "Rename categories to fit your event (e.g. 'Exhibitor booths'). Leave the budget empty if you don't need a category."
    s[f"A{C1 + 3}"].font = font(italic=True, size=9)
    for c, w in zip("ABC", [34, 26, 70]): s.column_dimensions[c].width = w
    cats = f"Settings!$A${C0}:$A${C1}"

    # ---------- Budget ----------
    K.title(bud, "Budget & payments", "One row per cost. Estimate first, then type the actual cost and what you have paid.")
    K.hdr(bud, 5, ["Category", "Item", "Vendor", "Estimate", "Actual cost", "Paid so far", "Still to pay", "Payment due", "Status", "Notes"],
          [26, 30, 22, 13, 13, 13, 13, 13, 15, 24])
    style_range(bud, B0, B1, 1, 10, inputs=(1, 2, 3, 4, 5, 6, 8, 10), fmt={4: MONEY, 5: MONEY, 6: MONEY, 7: MONEY, 8: DATE})
    for r in range(B0, B1 + 1):
        bud[f"G{r}"] = f'=IF(AND(D{r}="",E{r}=""),"",MAX(0,IF(E{r}="",D{r},E{r})-N(F{r})))'
        bud[f"I{r}"] = (f'=IF(AND(D{r}="",E{r}=""),"",IF(G{r}=0,"Paid",IF(AND(H{r}<>"",H{r}<Settings!$B$11),"OVERDUE",'
                        f'IF(AND(H{r}<>"",H{r}-Settings!$B$11<=14),"Due soon",IF(N(F{r})>0,"Deposit paid","Not paid")))))')
    if sample:
        for i, (cat, item, vend, est, act, paid, days) in enumerate(BUD):
            r = B0 + i
            bud[f"A{r}"] = cat; bud[f"B{r}"] = item; bud[f"C{r}"] = vend; bud[f"D{r}"] = est
            bud[f"E{r}"] = act or None; bud[f"F{r}"] = paid or None; bud[f"H{r}"] = edate - dt.timedelta(days=days)
    dv = DataValidation(type="list", formula1=f"={cats}", allow_blank=True); bud.add_data_validation(dv); dv.add(f"A{B0}:A{B1}")
    for txt, col in (("OVERDUE", "B91C1C"), ("Due soon", "B45309"), ("Paid", "15803D")):
        bud.conditional_formatting.add(f"I{B0}:I{B1}", FormulaRule(formula=[f'I{B0}="{txt}"'], font=Font(name=F, bold=True, color=col)))
    bud.conditional_formatting.add(f"E{B0}:E{B1}", FormulaRule(formula=[f'AND(ISNUMBER(E{B0}),ISNUMBER(D{B0}),E{B0}>D{B0})'], font=Font(name=F, color="B91C1C")))
    bud["A3"] = "Tip: leave 'Actual cost' empty until you have the final invoice – the estimate is used until then."; bud["A3"].font = font(italic=True, size=9)
    tr = B1 + 1
    bud[f"A{tr}"] = "TOTAL"
    for c in "DEFG": bud[f"{c}{tr}"] = f"=SUM({c}{B0}:{c}{B1})"; bud[f"{c}{tr}"].number_format = MONEY
    for c in "ABCDEFGHIJ": bud[f"{c}{tr}"].font = font(bold=True); bud[f"{c}{tr}"].fill = fill(SOFT); bud[f"{c}{tr}"].border = B
    bud.freeze_panes = f"C{B0}"

    # ---------- Income ----------
    K.title(inc, "Income", "Tickets, sponsors, donations… Leave empty for a private party.")
    K.hdr(inc, 5, ["Source", "Description", "Quantity", "Price each", "Expected", "Qty sold / received", "Received", "Still to come"],
          [24, 34, 11, 13, 14, 18, 14, 14])
    style_range(inc, I0, I1, 1, 8, inputs=(1, 2, 3, 4, 6), fmt={3: "#,##0", 4: MONEY, 5: MONEY, 6: "#,##0", 7: MONEY, 8: MONEY})
    for r in range(I0, I1 + 1):
        inc[f"E{r}"] = f'=IF(OR(C{r}="",D{r}=""),"",C{r}*D{r})'
        inc[f"G{r}"] = f'=IF(OR(F{r}="",D{r}=""),"",F{r}*D{r})'
        inc[f"H{r}"] = f'=IF(E{r}="","",MAX(0,E{r}-N(G{r})))'
    if sample:
        for i, (src, d, q, p, sold) in enumerate(INC):
            r = I0 + i; inc[f"A{r}"] = src; inc[f"B{r}"] = d; inc[f"C{r}"] = q; inc[f"D{r}"] = p; inc[f"F{r}"] = sold
    inc["J5"] = "Sources"; inc["J5"].font = font(bold=True)
    for i, t in enumerate(SRC): inc[f"J{6 + i}"] = t; inc[f"J{6 + i}"].font = font(size=9)
    dv = DataValidation(type="list", formula1=f"=$J$6:$J${5 + len(SRC)}", allow_blank=True); inc.add_data_validation(dv); dv.add(f"A{I0}:A{I1}")
    tr = I1 + 1; inc[f"A{tr}"] = "TOTAL"
    for c in "EGH": inc[f"{c}{tr}"] = f"=SUM({c}{I0}:{c}{I1})"; inc[f"{c}{tr}"].number_format = MONEY
    for c in "ABCDEFGH": inc[f"{c}{tr}"].font = font(bold=True); inc[f"{c}{tr}"].fill = fill(SOFT); inc[f"{c}{tr}"].border = B
    inc["A3"] = "Quantity = how many you plan to sell (or 1 for a sponsor). Qty sold / received = how many are actually paid."; inc["A3"].font = font(italic=True, size=9)
    inc.column_dimensions["J"].width = 22

    # ---------- Vendors ----------
    K.title(ven, "Vendors", "Contacts and contracts. Money columns add up automatically from the Budget tab (vendor name must match).")
    K.hdr(ven, 5, ["Vendor", "Service", "Contact person", "Phone", "Email", "Contract signed?", "Total cost", "Paid", "Still to pay", "Next payment due", "Notes"],
          [24, 18, 18, 13, 28, 12, 13, 13, 13, 14, 24])
    style_range(ven, V0, V1, 1, 11, inputs=(1, 2, 3, 4, 5, 6, 11), fmt={7: MONEY, 8: MONEY, 9: MONEY, 10: DATE})
    BR = f"Budget!$C${B0}:$C${B1}"
    for r in range(V0, V1 + 1):
        ven[f"G{r}"] = f'=IF(A{r}="","",SUMIFS(Budget!$E${B0}:$E${B1},{BR},A{r})+SUMIFS(Budget!$D${B0}:$D${B1},{BR},A{r},Budget!$E${B0}:$E${B1},""))'
        ven[f"H{r}"] = f'=IF(A{r}="","",SUMIFS(Budget!$F${B0}:$F${B1},{BR},A{r}))'
        ven[f"I{r}"] = f'=IF(A{r}="","",SUMIFS(Budget!$G${B0}:$G${B1},{BR},A{r}))'
        cond = f"({BR}=A{r})*ISNUMBER(Budget!$G${B0}:$G${B1})*(Budget!$G${B0}:$G${B1}>0)*ISNUMBER(Budget!$H${B0}:$H${B1})"
        ven[f"J{r}"] = f'=IF(A{r}="","",IF(SUMPRODUCT({cond})=0,"",SUMPRODUCT(MIN({cond}*Budget!$H${B0}:$H${B1}+(1-{cond})*1000000))))'
    if sample:
        for i, row in enumerate(VEND):
            for j, v in enumerate(row): ven.cell(row=V0 + i, column=1 + j, value=v or None)
    dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True); ven.add_data_validation(dv); dv.add(f"F{V0}:F{V1}")
    ven.conditional_formatting.add(f"F{V0}:F{V1}", FormulaRule(formula=[f'F{V0}="No"'], font=Font(name=F, bold=True, color="B91C1C")))

    # ---------- Guests ----------
    K.title(gst, "Guests & RSVP", "One row per invitation. 'Party size' = this person plus anyone they bring.")
    K.hdr(gst, 5, ["Name", "Group", "Email / phone", "Invited?", "RSVP", "Party size", "Coming (people)", "Dietary needs", "Table", "Ticket / paid?", "Notes"],
          [24, 16, 26, 10, 11, 10, 13, 18, 8, 12, 22])
    style_range(gst, G0, G1, 1, 11, inputs=(1, 2, 3, 4, 5, 6, 8, 9, 10, 11), fmt={6: "0", 7: "0"})
    for r in range(G0, G1 + 1):
        gst[f"G{r}"] = f'=IF(E{r}="Yes",MAX(1,N(F{r})),0)'
    for col, lst in (("D", '"Yes,No"'), ("E", '"Yes,No,Maybe,Pending"'), ("J", '"Yes,No,Free"')):
        dv = DataValidation(type="list", formula1=lst, allow_blank=True); gst.add_data_validation(dv); dv.add(f"{col}{G0}:{col}{G1}")
    if sample:
        groups = ["Board", "Donors", "Staff", "Sponsors", "Community", "Press"]
        diets = ["", "", "", "Vegetarian", "", "Gluten-free", "", "Vegan", "", "Nut allergy"]
        rsvp = ["Yes", "Yes", "Yes", "No", "Yes", "Maybe", "Pending", "Yes", "Yes", "Pending"]
        for i in range(64):
            r = G0 + i
            gst[f"A{r}"] = f"{FIRST[i % 30]} {LAST[(i * 7) % 15]}"; gst[f"B{r}"] = groups[i % 6]
            gst[f"C{r}"] = f"{FIRST[i % 30].lower()}@mail.example"; gst[f"D{r}"] = "Yes"; gst[f"E{r}"] = rsvp[i % 10]
            gst[f"F{r}"] = 2 if i % 3 == 0 else 1; gst[f"H{r}"] = diets[i % 10] or None
            gst[f"I{r}"] = (i // 8) + 1 if rsvp[i % 10] == "Yes" else None; gst[f"J{r}"] = "Yes" if rsvp[i % 10] == "Yes" else None
    for txt, col in (("Yes", "15803D"), ("No", "9CA3AF"), ("Maybe", "B45309"), ("Pending", "B45309")):
        gst.conditional_formatting.add(f"E{G0}:E{G1}", FormulaRule(formula=[f'E{G0}="{txt}"'], font=Font(name=F, bold=True, color=col)))
    gst.freeze_panes = f"B{G0}"

    # ---------- Tasks ----------
    K.title(tsk, "Tasks & timeline", "A ready-made checklist counted back from your event date. Edit, delete or add tasks.")
    K.hdr(tsk, 5, ["Task", "Area", "Owner", "Weeks before event", "Due date", "Done?", "Status", "Notes"], [50, 22, 14, 12, 13, 9, 14, 26])
    style_range(tsk, T0, T1, 1, 8, inputs=(1, 2, 3, 4, 6, 8), fmt={4: "0", 5: DATE})
    for r in range(T0, T1 + 1):
        tsk[f"E{r}"] = f'=IF(OR(A{r}="",D{r}=""),"",{ED}-D{r}*7)'
        tsk[f"G{r}"] = (f'=IF(A{r}="","",IF(F{r}="Yes","Done",IF(E{r}="","No date",IF(E{r}<Settings!$B$11,"OVERDUE",'
                        f'IF(E{r}-Settings!$B$11<=7,"This week","Upcoming")))))')
    for i, (t, a, w) in enumerate(TASKS):
        r = T0 + i; tsk[f"A{r}"] = t; tsk[f"B{r}"] = a; tsk[f"D{r}"] = w
        if sample:
            due = edate - dt.timedelta(weeks=w)
            if due < dt.date(2026, 11, 3): tsk[f"F{r}"] = "Yes" if i not in (7, 12) else None
            tsk[f"C{r}"] = ["Anna", "Ben", "Carla"][i % 3]
    dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True); tsk.add_data_validation(dv); dv.add(f"F{T0}:F{T1}")
    for txt, col in (("OVERDUE", "B91C1C"), ("This week", "B45309"), ("Done", "15803D")):
        tsk.conditional_formatting.add(f"G{T0}:G{T1}", FormulaRule(formula=[f'G{T0}="{txt}"'], font=Font(name=F, bold=True, color=col)))
    tsk.conditional_formatting.add(f"A{T0}:F{T1}", FormulaRule(formula=[f'$F{T0}="Yes"'], font=Font(name=F, color="9CA3AF")))
    tsk["A3"] = "Negative weeks = after the event (thank-yous, wrap-up). Type 0 for the event day itself."; tsk["A3"].font = font(italic=True, size=9)
    tsk.freeze_panes = f"B{T0}"

    # ---------- Dashboard ----------
    K.title(dash, "Event Dashboard", "Updates by itself from the other tabs.")
    dash["A3"] = '=Settings!B4&"  ·  "&TEXT(Settings!B5,"dddd d mmmm yyyy")&IF(Settings!B6="","","  ·  "&Settings!B6)&"  ·  amounts in "&Settings!B9'
    dash["A3"].font = font(bold=True)
    BT = B1 + 1; IT = I1 + 1
    cost = f"(Budget!$E${BT}+SUMIFS(Budget!$D${B0}:$D${B1},Budget!$E${B0}:$E${B1},\"\"))"
    kp = [("Days to go", '=IF(Settings!B5-Settings!B11<0,"Done!",Settings!B5-Settings!B11)', "0"),
          ("Total budget", "=Settings!B8", MONEY),
          ("Expected cost", f"={cost}", MONEY),
          ("Left in budget", f"=Settings!B8-{cost}", '#,##0.00;[Red]-#,##0.00'),
          ("Guests coming", f"=SUM(Guests!$G${G0}:$G${G1})", "0"),
          ("Tasks done", f'=IFERROR(COUNTIF(Tasks!$G${T0}:$G${T1},"Done")/COUNTA(Tasks!$A${T0}:$A${T1}),0)', PCT)]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 3) * 3, 1 + (i % 3) * 2, lab, f, fmt)
        dash.cell(row=6 + (i // 3) * 3, column=1 + (i % 3) * 2).alignment = Alignment(horizontal="left")
    for c, w in zip("ABCDEFG", [40, 16, 15, 15, 15, 15, 3]): dash.column_dimensions[c].width = w

    def block(r0, title, rows_):
        K.section(dash, f"A{r0}", title)
        for i, (lab, f, fmt) in enumerate(rows_, r0 + 1):
            dash[f"A{i}"] = lab; dash[f"A{i}"].font = font(bold=True); dash[f"A{i}"].border = B; dash[f"A{i}"].fill = fill(SOFT)
            c = dash[f"B{i}"]; c.value = f; c.number_format = fmt; c.border = B; c.font = Font(name=F, bold=True, color=MAIN)
            c.alignment = Alignment(horizontal="right")
    block(11, "Money", [
        ("Budget used (expected cost ÷ budget)", f"=IFERROR({cost}/Settings!B8,0)", PCT),
        ("Paid so far", f"=Budget!$F${BT}", MONEY),
        ("Still to pay vendors", f"=Budget!$G${BT}", MONEY),
        ("Overdue payments", f'=COUNTIF(Budget!$I${B0}:$I${B1},"OVERDUE")&" ("&TEXT(SUMIFS(Budget!$G${B0}:$G${B1},Budget!$I${B0}:$I${B1},"OVERDUE"),"#,##0")&")"', "@"),
        ("Due in the next 14 days", f'=COUNTIF(Budget!$I${B0}:$I${B1},"Due soon")&" ("&TEXT(SUMIFS(Budget!$G${B0}:$G${B1},Budget!$I${B0}:$I${B1},"Due soon"),"#,##0")&")"', "@"),
        ("Contingency buffer (Settings %)", "=Settings!B8*Settings!B10", MONEY),
        ("Income expected", f"=Income!$E${IT}", MONEY),
        ("Income received", f"=Income!$G${IT}", MONEY),
        ("Profit / loss (expected income – cost)", f"=Income!$E${IT}-{cost}", '#,##0.00;[Red]-#,##0.00'),
        ("Cost per guest", f"=IFERROR({cost}/IF(C9>0,C9,Settings!B7),0)", MONEY)])
    dash["C12"] = f'={bar("B12")}'; dash["C12"].font = Font(name=F, color=MAIN)
    block(23, "Guests & tasks", [
        ("Invitations", f'=COUNTIF(Guests!$D${G0}:$D${G1},"Yes")', "0"),
        ("RSVP yes (people)", f"=SUM(Guests!$G${G0}:$G${G1})", "0"),
        ("Still waiting (pending + maybe)", f'=COUNTIF(Guests!$E${G0}:$E${G1},"Pending")+COUNTIF(Guests!$E${G0}:$E${G1},"Maybe")', "0"),
        ("Declined", f'=COUNTIF(Guests!$E${G0}:$E${G1},"No")', "0"),
        ("Coming vs expected", f"=IFERROR(B25/Settings!B7,0)", PCT),
        ("Special dietary needs", f'=COUNTIFS(Guests!$H${G0}:$H${G1},"<>",Guests!$E${G0}:$E${G1},"Yes")', "0"),
        ("Tasks overdue", f'=COUNTIF(Tasks!$G${T0}:$G${T1},"OVERDUE")', "0"),
        ("Tasks due this week", f'=COUNTIF(Tasks!$G${T0}:$G${T1},"This week")', "0"),
        ("Vendors without a signed contract", f'=COUNTIF(Vendors!$F${V0}:$F${V1},"No")', "0")])
    dash["C28"] = f'={bar("B28")}'; dash["C28"].font = Font(name=F, color=MAIN)

    # by category
    CR = 34
    K.section(dash, f"A{CR}", "Budget by category")
    K.hdr(dash, CR + 1, ["Category", "Planned budget", "Estimate", "Actual", "Expected cost", "Over / under"])
    for i in range(NC):
        r = CR + 2 + i; k = C0 + i
        dash[f"A{r}"] = f'=IF(Settings!A{k}="","",Settings!A{k})'
        dash[f"B{r}"] = f"=N(Settings!B{k})"
        dash[f"C{r}"] = f"=SUMIFS(Budget!$D${B0}:$D${B1},Budget!$A${B0}:$A${B1},Settings!A{k})"
        dash[f"D{r}"] = f"=SUMIFS(Budget!$E${B0}:$E${B1},Budget!$A${B0}:$A${B1},Settings!A{k})"
        dash[f"E{r}"] = f'=D{r}+SUMIFS(Budget!$D${B0}:$D${B1},Budget!$A${B0}:$A${B1},Settings!A{k},Budget!$E${B0}:$E${B1},"")'
        dash[f"F{r}"] = f"=B{r}-E{r}"
    CE = CR + 1 + NC
    style_range(dash, CR + 2, CE, 1, 6, fmt={2: MONEY, 3: MONEY, 4: MONEY, 5: MONEY, 6: '#,##0.00;[Red]-#,##0.00'})
    dash[f"A{CE + 1}"] = "Total"
    for c in "BCDEF": dash[f"{c}{CE + 1}"] = f"=SUM({c}{CR + 2}:{c}{CE})"; dash[f"{c}{CE + 1}"].number_format = MONEY if c != "F" else '#,##0.00;[Red]-#,##0.00'
    for c in "ABCDEF": dash[f"{c}{CE + 1}"].font = font(bold=True); dash[f"{c}{CE + 1}"].fill = fill(SOFT); dash[f"{c}{CE + 1}"].border = B
    dash[f"A{CE + 2}"] = "Over / under = planned budget – expected cost. Red = over budget."; dash[f"A{CE + 2}"].font = font(italic=True, size=9)

    ch = BarChart(); ch.type = "bar"; ch.grouping = "clustered"; ch.title = "Planned vs expected cost"; ch.height = 9; ch.width = 16
    ch.add_data(Reference(dash, min_col=2, min_row=CR + 1, max_row=CE), titles_from_data=True)
    ch.add_data(Reference(dash, min_col=5, min_row=CR + 1, max_row=CE), titles_from_data=True)
    ch.set_categories(Reference(dash, min_col=1, min_row=CR + 2, max_row=CE)); ch.y_axis.majorGridlines = None
    ch.x_axis.scaling.orientation = "maxMin"; ch.legend.position = "b"; ch.x_axis.tickLblSkip = 1; ch.height = 11
    for ser, col in zip(ch.series, ("F9A8D4", MAIN)): ser.graphicalProperties.solidFill = col
    dash.add_chart(ch, "H11")
    # RSVP pie helper
    dash["K3"] = "Yes"; dash["L3"] = f'=COUNTIF(Guests!$E${G0}:$E${G1},"Yes")'
    dash["K4"] = "Pending / maybe"; dash["L4"] = "=B26"
    dash["K5"] = "No"; dash["L5"] = "=B27"
    for c in ("K3", "K4", "K5", "L3", "L4", "L5"): dash[c].font = font(size=8, color="FFFFFF")
    pie = PieChart(); pie.title = "RSVPs (invitations)"; pie.height = 7; pie.width = 10
    pie.add_data(Reference(dash, min_col=12, min_row=3, max_row=5)); pie.set_categories(Reference(dash, min_col=11, min_row=3, max_row=5))
    from openpyxl.chart.series import DataPoint
    for idx, col in enumerate(("15803D", "F59E0B", "9CA3AF")):
        pt = DataPoint(idx=idx); pt.graphicalProperties.solidFill = col; pie.series[0].dPt.append(pt)
    dash.add_chart(pie, "H34")

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Settings tab: type the event name, date, expected guests and total budget. Split the budget over the categories (rename them to fit your event).", False),
             ("2.  Budget tab: one row per cost. Type the estimate, later the actual cost, what you paid and when the rest is due. Status shows Paid / Due soon / OVERDUE.", False),
             ("3.  Income tab (optional): tickets, sponsors, donations, raffle… See your profit or loss on the Dashboard.", False),
             ("4.  Vendors tab: contacts and contracts. Cost, paid and next payment date fill in from the Budget tab when the vendor name matches.", False),
             ("5.  Guests tab: invitations, RSVP, party size, dietary needs and table numbers.", False),
             ("6.  Tasks tab: a 31-step checklist is ready, with due dates counted back from your event date. Mark tasks Done as you go.", False),
             ("7.  Dashboard: countdown, budget used, money still to pay, overdue items, guests coming, tasks done and charts.", False),
             ("", False), ("Tips", True),
             ("•  Keep 5–15% of the budget as a buffer – most events have surprise costs (extra chairs, overtime, delivery).", False),
             ("•  Catering is usually the biggest cost: give the caterer final numbers from the Guests tab 1–2 weeks before.", False),
             ("•  Works for birthdays, corporate events, conferences, fundraisers, reunions, workshops and launches.", False),
             ("•  Planning several events? Save one copy of this file per event.", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  100 budget lines, 30 income lines, 40 vendors, 300 guests, 80 tasks. Any currency (change the label on Settings).", False),
             ("•  To use Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("The gala, vendors, guests and amounts are made-up examples. Use the blank copy in your download for your own event.", False)]
    K.start(st, "Event Planner & Budget", "Budget, payments, vendors, guests and a timeline checklist for any event – Excel and Google Sheets.", lines)
    finish(wb, sample, path)


META = {
    "name": "Event Planner & Budget (Vendors, Guest RSVP, Tasks Timeline, Ticket Income) – Excel & Google Sheets",
    "file_base": "Event-Planner-Budget",
    "price": 17,
    "color": "#BE185D",
    "title": "Event Planner & Budget",
    "tagline": "Budget & payments · vendors · guest RSVP · timeline checklist · ticket income",
    "intro": "Plan any event in one file – a party, conference, fundraiser, workshop or company event. Track every cost and payment, your vendors, guests and RSVPs, and a ready-made checklist with due dates counted back from your event day. The dashboard shows what is left in the budget, what is overdue and how many people are coming.",
    "features": [
        "Budget by category: planned vs estimate vs actual, with over/under in red",
        "Payments: paid so far, still to pay, due date and automatic status (Paid, Due soon, OVERDUE)",
        "Income tab for tickets, sponsors, donations and raffles – see your profit or loss and cost per guest",
        "Vendor list with contacts, contracts and totals pulled from the budget",
        "Guest list with RSVP, party size, dietary needs and table numbers (300 rows)",
        "31-step timeline checklist with due dates counted back from your event date",
        "Dashboard with countdown, budget used, overdue alerts and charts; blank copy plus a sample copy",
    ],
    "tags": ["event planner", "event budget", "party planner", "conference planner", "event checklist", "guest list template",
             "rsvp tracker", "vendor tracker", "google sheets", "excel template", "fundraiser planner", "event management", "budget spreadsheet"],
}
