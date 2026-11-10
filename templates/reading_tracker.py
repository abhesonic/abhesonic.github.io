"""Reading Log & Book Tracker – library with ratings, reading challenge (yearly goal), reading journal with streaks, calendar heatmap, stats.

Settings (year, goals, genres, formats, sources) -> Library + Reading Journal -> Monthly / Stats / Reading Calendar -> Dashboard.
"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, DataBarRule, ColorScaleRule
from openpyxl.chart import BarChart, LineChart, Reference
from _common import Kit, F, B, MONEY, DATE, font, fill, style_range, input_cell, finish

MAIN = "9D174D"; SOFT = "FCE7F3"
K = Kit(MAIN, SOFT)
GENRES = ["Classics", "Literary fiction", "Fantasy", "Science fiction", "Mystery & thriller", "Romance", "Historical fiction", "Horror",
          "Adventure", "Young adult", "Non-fiction", "Biography & memoir", "Self-help", "Philosophy", "Poetry"]
NG = 20; GE0 = 6; GE1 = GE0 + NG - 1
FORMATS = ["Paperback", "Hardcover", "eBook", "Audiobook"]
NF = 6; FO0 = 6; FO1 = FO0 + NF - 1
SOURCES = ["Owned", "Library", "Borrowed", "Subscription", "Gift", "Second-hand"]
NSO = 8; SO0 = 6; SO1 = SO0 + NSO - 1
STATUS = ["Want to read", "Reading", "Finished", "Did not finish"]
NL = 600; L0 = 6; L1 = L0 + NL - 1
NJ = 2000; J0 = 6; J1 = J0 + NJ - 1
D0 = 6; D1 = D0 + 365                         # daily helper rows (366 days) in Reading Journal H:J
MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

BOOKS = [  # title, author, genre, format, pages, rating, source, price
    ("Pride and Prejudice", "Jane Austen", "Classics", "Paperback", 432, 5, "Owned", 8.99),
    ("Frankenstein", "Mary Shelley", "Horror", "eBook", 280, 4, "Subscription", None),
    ("The Hound of the Baskervilles", "Arthur Conan Doyle", "Mystery & thriller", "Paperback", 256, 4.5, "Library", None),
    ("Meditations", "Marcus Aurelius", "Philosophy", "Hardcover", 304, 4.5, "Owned", 14.50),
    ("The Time Machine", "H. G. Wells", "Science fiction", "Audiobook", 128, 3.5, "Subscription", None),
    ("Jane Eyre", "Charlotte Brontë", "Classics", "Paperback", 532, 4.5, "Second-hand", 3.00),
    ("Treasure Island", "Robert Louis Stevenson", "Adventure", "eBook", 292, 4, "Subscription", None),
    ("Dracula", "Bram Stoker", "Horror", "Paperback", 488, 4, "Library", None),
    ("Walden", "Henry David Thoreau", "Non-fiction", "Paperback", 352, 3, "Owned", 9.99),
    ("Little Women", "Louisa May Alcott", "Classics", "Hardcover", 759, 4.5, "Gift", None),
    ("The War of the Worlds", "H. G. Wells", "Science fiction", "Audiobook", 192, 4, "Subscription", None),
    ("Persuasion", "Jane Austen", "Romance", "Paperback", 272, 5, "Owned", 7.99),
    ("Around the World in Eighty Days", "Jules Verne", "Adventure", "eBook", 256, 4, "Subscription", None),
    ("The Picture of Dorian Gray", "Oscar Wilde", "Literary fiction", "Paperback", 254, 4.5, "Library", None),
    ("The Art of War", "Sun Tzu", "Philosophy", "eBook", 112, 3, "Subscription", None),
    ("Wuthering Heights", "Emily Brontë", "Romance", "Paperback", 416, 3.5, "Second-hand", 2.50),
    ("The Call of the Wild", "Jack London", "Adventure", "Audiobook", 172, 4, "Subscription", None),
    ("Great Expectations", "Charles Dickens", "Classics", "Paperback", 544, 4, "Library", None),
    ("The Secret Garden", "Frances Hodgson Burnett", "Young adult", "Hardcover", 331, 4.5, "Gift", None),
    ("Twenty Thousand Leagues Under the Seas", "Jules Verne", "Science fiction", "eBook", 448, 3.5, "Subscription", None),
    ("Mrs Dalloway", "Virginia Woolf", "Literary fiction", "Paperback", 194, 3.5, "Owned", 10.99),
    ("The Murder of Roger Ackroyd", "Agatha Christie", "Mystery & thriller", "Paperback", 312, 5, "Owned", 9.50),
    ("Anne of Green Gables", "L. M. Montgomery", "Young adult", "eBook", 320, 4.5, "Subscription", None),
    ("Heart of Darkness", "Joseph Conrad", "Literary fiction", "Audiobook", 96, 3, "Subscription", None),
    ("The Count of Monte Cristo", "Alexandre Dumas", "Adventure", "Paperback", 1276, 5, "Owned", 15.99),
    ("Siddhartha", "Hermann Hesse", "Philosophy", "Paperback", 152, 4, "Library", None),
    ("The Great Gatsby", "F. Scott Fitzgerald", "Literary fiction", "Hardcover", 180, 4, "Owned", 12.00),
    ("Sense and Sensibility", "Jane Austen", "Romance", "Audiobook", 409, 4, "Subscription", None),
    ("A Study in Scarlet", "Arthur Conan Doyle", "Mystery & thriller", "eBook", 176, 4, "Subscription", None),
    ("The Wind in the Willows", "Kenneth Grahame", "Classics", "Hardcover", 224, 3.5, "Gift", None),
    ("Leaves of Grass", "Walt Whitman", "Poetry", "Paperback", 368, 3, "Second-hand", 4.00),
    ("Narrative of the Life of Frederick Douglass", "Frederick Douglass", "Biography & memoir", "eBook", 144, 5, "Subscription", None),
    ("The Phantom of the Opera", "Gaston Leroux", "Horror", "Paperback", 360, 3.5, "Library", None),
    ("Peter Pan", "J. M. Barrie", "Young adult", "Audiobook", 192, 4, "Subscription", None),
    ("The Autobiography of Benjamin Franklin", "Benjamin Franklin", "Biography & memoir", "Paperback", 256, 3.5, "Library", None),
    ("The Three Musketeers", "Alexandre Dumas", "Historical fiction", "eBook", 704, 4.5, "Subscription", None),
]
READING = [("Middlemarch", "George Eliot", "Classics", "Paperback", 880, 412, "Owned", 11.99),
           ("The Hunchback of Notre-Dame", "Victor Hugo", "Historical fiction", "Audiobook", 512, 140, "Subscription", None)]
TBR = [("Moby-Dick", "Herman Melville", "Adventure"), ("Anna Karenina", "Leo Tolstoy", "Classics"), ("Emma", "Jane Austen", "Romance"),
       ("Crime and Punishment", "Fyodor Dostoevsky", "Literary fiction"), ("The Iliad", "Homer", "Poetry"),
       ("A Tale of Two Cities", "Charles Dickens", "Historical fiction"), ("The Jungle Book", "Rudyard Kipling", "Young adult"),
       ("Gulliver's Travels", "Jonathan Swift", "Adventure"), ("The Prophet", "Kahlil Gibran", "Poetry"), ("Don Quixote", "Miguel de Cervantes", "Classics")]
DNF = ("Ulysses", "James Joyce", "Literary fiction", "Paperback", 730, 120, "Library")


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); lib = wb.create_sheet("Library"); jr = wb.create_sheet("Reading Journal")
    mo = wb.create_sheet("Monthly"); ss = wb.create_sheet("Stats"); cal = wb.create_sheet("Reading Calendar"); s = wb.create_sheet("Settings")

    # ---------- Settings ----------
    K.title(s, "Settings", "Yellow cells are yours to edit. Set your reading year and goals first.")
    SET = [(3, "Reading year", 2026, "0"), (4, "Goal: books this year", 40, "0"), (5, "Goal: pages this year", 12000, "#,##0"),
           (6, "Goal: minutes per day", 30, "0")]
    for r, lab, v, fmt in SET:
        s[f"A{r}"] = lab; s[f"A{r}"].font = font(bold=True); s[f"A{r}"].border = B; input_cell(s[f"B{r}"], v); s[f"B{r}"].number_format = fmt
    s["A8"] = "Currency symbol is set by the number format – change it in Excel if you like."; s["A8"].font = font(italic=True, size=9)
    K.hdr(s, 5, ["Genres"], col0=4)
    for i in range(NG): input_cell(s[f"D{GE0 + i}"], GENRES[i] if i < len(GENRES) else None)
    K.hdr(s, 5, ["Formats"], col0=6)
    for i in range(NF): input_cell(s[f"F{FO0 + i}"], FORMATS[i] if i < len(FORMATS) else None)
    K.hdr(s, 5, ["Sources"], col0=8)
    for i in range(NSO): input_cell(s[f"H{SO0 + i}"], SOURCES[i] if i < len(SOURCES) else None)
    for c, w in zip("ABCDEFGH", [26, 10, 3, 20, 3, 14, 3, 14]): s.column_dimensions[c].width = w
    YR = "Settings!$B$3"
    GEN = f"Settings!$D${GE0}:$D${GE1}"; FMT = f"Settings!$F${FO0}:$F${FO1}"; SRC = f"Settings!$H${SO0}:$H${SO1}"

    # ---------- Library ----------
    K.title(lib, "Library", "Every book: want to read, reading, finished or did not finish. Ratings from 0.5 to 5 (half stars allowed).")
    cols = ["Title", "Author", "Genre", "Format", "Pages", "Status", "Started", "Finished", "Rating (0.5–5)", "Current page",
            "Source", "Price", "Notes / favourite quote", "Progress", "Stars", "Days to read", "Year read", "Month"]
    K.hdr(lib, 5, cols, [32, 22, 17, 12, 8, 13, 11, 11, 9, 9, 12, 9, 30, 10, 10, 8, 7, 7])
    style_range(lib, L0, L1, 1, 18, inputs=range(1, 14), fmt={5: "#,##0", 7: DATE, 8: DATE, 9: "0.0", 10: "#,##0", 12: MONEY, 14: "0%", 16: "0"})
    for r in range(L0, L1 + 1):
        lib[f"N{r}"] = f'=IF(OR(A{r}="",N(E{r})=0),"",IF(F{r}="Finished",1,MIN(1,N(J{r})/E{r})))'
        lib[f"O{r}"] = f'=IF(NOT(ISNUMBER(I{r})),"",REPT("★",INT(I{r}))&IF(I{r}-INT(I{r})>=0.5,"½","")&REPT("☆",5-INT(I{r})-IF(I{r}-INT(I{r})>=0.5,1,0)))'
        lib[f"P{r}"] = f'=IF(AND(ISNUMBER(G{r}),ISNUMBER(H{r})),H{r}-G{r}+1,"")'
        lib[f"Q{r}"] = f'=IF(AND(F{r}="Finished",ISNUMBER(H{r})),YEAR(H{r}),"")'
        lib[f"R{r}"] = f'=IF(Q{r}="","",MONTH(H{r}))'
        lib[f"S{r}"] = f'=IF(AND(A{r}<>"",F{r}="Reading"),ROW(),99999999)'
        lib[f"T{r}"] = f'=IF(AND(Q{r}={YR},ISNUMBER(I{r})),I{r}+ROW()/1000000,-1)'
        lib[f"U{r}"] = f'=IF(AND(A{r}<>"",F{r}="Want to read"),ROW(),99999999)'
    for c in "STU": lib.column_dimensions[c].hidden = True
    for col, lst in (("C", f"={GEN}"), ("D", f"={FMT}"), ("F", '"' + ",".join(STATUS) + '"'), ("K", f"={SRC}")):
        dv = DataValidation(type="list", formula1=lst, allow_blank=True); lib.add_data_validation(dv); dv.add(f"{col}{L0}:{col}{L1}")
    dv = DataValidation(type="decimal", operator="between", formula1="0.5", formula2="5", allow_blank=True, showErrorMessage=True,
                        error="Rating from 0.5 to 5"); lib.add_data_validation(dv); dv.add(f"I{L0}:I{L1}")
    lib.conditional_formatting.add(f"F{L0}:F{L1}", FormulaRule(formula=[f'F{L0}="Finished"'], font=Font(name=F, bold=True, color="15803D"), fill=fill("DCFCE7")))
    lib.conditional_formatting.add(f"F{L0}:F{L1}", FormulaRule(formula=[f'F{L0}="Reading"'], font=Font(name=F, bold=True, color=MAIN), fill=fill(SOFT)))
    lib.conditional_formatting.add(f"F{L0}:F{L1}", FormulaRule(formula=[f'F{L0}="Did not finish"'], font=Font(name=F, color="9CA3AF")))
    lib.conditional_formatting.add(f"N{L0}:N{L1}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="F9A8D4"))
    for r in range(L0, L1 + 1): lib[f"O{r}"].font = Font(name=F, color="D97706")
    lib.freeze_panes = f"B{L0}"; lib.auto_filter.ref = f"A5:R{L1}"
    LR = lambda c: f"Library!${c}${L0}:${c}${L1}"

    # ---------- Reading Journal ----------
    K.title(jr, "Reading Journal", "Log each reading session: date, book, pages and minutes. Builds your streak and the Reading Calendar.")
    K.hdr(jr, 5, ["Date", "Book", "Pages read", "Minutes", "Thoughts"], [12, 32, 10, 9, 36])
    style_range(jr, J0, J1, 1, 5, inputs=range(1, 6), fmt={1: DATE, 3: "0", 4: "0"})
    dv = DataValidation(type="list", formula1=f"={LR('A')}", allow_blank=True); jr.add_data_validation(dv); dv.add(f"B{J0}:B{J1}")
    jr.freeze_panes = f"A{J0}"
    JR = lambda c: f"'Reading Journal'!${c}${J0}:${c}${J1}"
    # daily helper (hidden)
    jr["H5"] = "Day"; jr["I5"] = "Minutes"; jr["J5"] = "Streak"
    for k in range(366):
        r = D0 + k
        jr[f"H{r}"] = f"=DATE({YR},1,1)+{k}"; jr[f"H{r}"].number_format = DATE
        jr[f"I{r}"] = f'=SUMIFS($D${J0}:$D${J1},$A${J0}:$A${J1},H{r})+COUNTIFS($A${J0}:$A${J1},H{r},$D${J0}:$D${J1},"")*0.001'
        jr[f"J{r}"] = f"=IF(I{r}>0,{'J' + str(r - 1) + '+1' if k else '1'},0)"
    for c in "HIJ": jr.column_dimensions[c].hidden = True
    DH = f"'Reading Journal'!$H${D0}:$H${D1}"; DI = f"'Reading Journal'!$I${D0}:$I${D1}"; DJ = f"'Reading Journal'!$J${D0}:$J${D1}"
    IDX = f"(TODAY()-DATE({YR},1,1)+1)"

    # ---------- Monthly ----------
    K.title(mo, "Monthly Reading Challenge", "Books and pages per month for the year in Settings, compared with the pace you need for your goal.")
    K.hdr(mo, 5, ["Month", "Books finished", "Pages (finished)", "Avg rating", "Minutes read", "Pages logged", "Books so far", "Goal pace", "Ahead / behind"],
          [10, 12, 13, 10, 12, 12, 12, 10, 13])
    for m in range(12):
        r = 6 + m; mo[f"A{r}"] = MON[m]
        ms = f"DATE({YR},{m + 1},1)"
        mo[f"B{r}"] = f'=COUNTIFS({LR("Q")},{YR},{LR("R")},{m + 1})'
        mo[f"C{r}"] = f'=SUMIFS({LR("E")},{LR("Q")},{YR},{LR("R")},{m + 1})'
        mo[f"D{r}"] = f'=IFERROR(AVERAGEIFS({LR("I")},{LR("Q")},{YR},{LR("R")},{m + 1}),"")'
        mo[f"E{r}"] = f'=SUMIFS({JR("D")},{JR("A")},">="&{ms},{JR("A")},"<"&EDATE({ms},1))'
        mo[f"F{r}"] = f'=SUMIFS({JR("C")},{JR("A")},">="&{ms},{JR("A")},"<"&EDATE({ms},1))'
        mo[f"G{r}"] = f"=SUM($B$6:B{r})"
        mo[f"H{r}"] = f"=Settings!$B$4*{m + 1}/12"
        mo[f"I{r}"] = f'=IF({ms}>TODAY(),"",G{r}-H{r})'
    style_range(mo, 6, 17, 1, 9, fmt={2: "0", 3: "#,##0", 4: "0.0", 5: "#,##0", 6: "#,##0", 7: "0", 8: "0.0", 9: '+0.0;-0.0;0.0'})
    mo["A18"] = "Total"; mo["B18"] = "=SUM(B6:B17)"; mo["C18"] = "=SUM(C6:C17)"; mo["D18"] = f'=IFERROR(AVERAGEIFS({LR("I")},{LR("Q")},{YR}),"")'
    mo["E18"] = "=SUM(E6:E17)"; mo["F18"] = "=SUM(F6:F17)"
    for c, f_ in zip("BCDEF", ["0", "#,##0", "0.0", "#,##0", "#,##0"]):
        mo[f"{c}18"].number_format = f_; mo[f"{c}18"].font = font(bold=True); mo[f"{c}18"].fill = fill(SOFT)
    mo["A18"].font = font(bold=True); mo["A18"].fill = fill(SOFT)
    mo.conditional_formatting.add("I6:I17", FormulaRule(formula=['AND(ISNUMBER(I6),I6<0)'], font=Font(name=F, bold=True, color="B91C1C")))
    mo.conditional_formatting.add("I6:I17", FormulaRule(formula=['AND(ISNUMBER(I6),I6>=0)'], font=Font(name=F, bold=True, color="15803D")))
    mo.conditional_formatting.add("B6:B17", DataBarRule(start_type="num", start_value=0, end_type="max", color="F9A8D4"))
    mo["A20"] = "Goal pace = your yearly book goal spread evenly over 12 months."; mo["A20"].font = font(italic=True, size=9)
    lc = LineChart(); lc.title = "Books so far vs goal"; lc.height = 7.5; lc.width = 15; lc.y_axis.majorGridlines = None
    lc.add_data(Reference(mo, min_col=7, max_col=8, min_row=5, max_row=17), titles_from_data=True)
    lc.set_categories(Reference(mo, min_col=1, min_row=6, max_row=17))
    lc.series[0].graphicalProperties.line.solidFill = MAIN; lc.series[0].graphicalProperties.line.width = 32000
    lc.series[1].graphicalProperties.line.solidFill = "9CA3AF"; lc.series[1].graphicalProperties.line.dashStyle = "dash"
    mo.add_chart(lc, "K5")

    # ---------- Stats ----------
    K.title(ss, "Reading Stats", "Books finished in the reading year, by genre, format and source – plus your rating spread and top-rated books.")
    K.hdr(ss, 5, ["Genre", "Books", "Pages", "Avg rating", "% of books"], [20, 8, 10, 10, 10])
    for i in range(NG):
        r = 6 + i; g = f"Settings!D{GE0 + i}"
        ss[f"A{r}"] = f'=IF({g}="","",{g})'
        ss[f"B{r}"] = f'=IF(A{r}="","",COUNTIFS({LR("C")},A{r},{LR("Q")},{YR}))'
        ss[f"C{r}"] = f'=IF(A{r}="","",SUMIFS({LR("E")},{LR("C")},A{r},{LR("Q")},{YR}))'
        ss[f"D{r}"] = f'=IF(A{r}="","",IFERROR(AVERAGEIFS({LR("I")},{LR("C")},A{r},{LR("Q")},{YR}),""))'
        ss[f"E{r}"] = f'=IF(OR(A{r}="",Monthly!$B$18=0),"",B{r}/Monthly!$B$18)'
    style_range(ss, 6, 5 + NG, 1, 5, fmt={2: "0", 3: "#,##0", 4: "0.0", 5: "0%"})
    ss.conditional_formatting.add(f"B6:B{5 + NG}", DataBarRule(start_type="num", start_value=0, end_type="max", color="F9A8D4"))
    K.hdr(ss, 5, ["Format", "Books", "Pages"], [14, 8, 10], col0=7)
    for i in range(NF):
        r = 6 + i; g = f"Settings!F{FO0 + i}"
        ss[f"G{r}"] = f'=IF({g}="","",{g})'
        ss[f"H{r}"] = f'=IF(G{r}="","",COUNTIFS({LR("D")},G{r},{LR("Q")},{YR}))'
        ss[f"I{r}"] = f'=IF(G{r}="","",SUMIFS({LR("E")},{LR("D")},G{r},{LR("Q")},{YR}))'
    style_range(ss, 6, 5 + NF, 7, 9, fmt={8: "0", 9: "#,##0"})
    K.hdr(ss, 14, ["Source", "Books", "Spent"], col0=7)
    for i in range(NSO):
        r = 15 + i; g = f"Settings!H{SO0 + i}"
        ss[f"G{r}"] = f'=IF({g}="","",{g})'
        ss[f"H{r}"] = f'=IF(G{r}="","",COUNTIFS({LR("K")},G{r},{LR("Q")},{YR}))'
        ss[f"I{r}"] = f'=IF(G{r}="","",SUMIFS({LR("L")},{LR("K")},G{r},{LR("Q")},{YR}))'
    style_range(ss, 15, 14 + NSO, 7, 9, fmt={8: "0", 9: MONEY})
    K.hdr(ss, 5, ["Rating", "Books"], [10, 8], col0=11)
    for j, st_ in enumerate([5, 4, 3, 2, 1]):
        r = 6 + j; ss[f"K{r}"] = "★" * st_
        ss[f"L{r}"] = f'=COUNTIFS({LR("Q")},{YR},{LR("I")},">="&({st_}-0.5),{LR("I")},"<"&({st_}+0.5))'
        ss[f"K{r}"].font = Font(name=F, color="D97706")
    style_range(ss, 6, 10, 12, 12, fmt={12: "0"})
    for r in range(6, 11): ss[f"K{r}"].border = B
    ss["K11"] = "4.5 counts as 5★, 3.5 as 4★ …"; ss["K11"].font = font(italic=True, size=8)
    K.section(ss, "K13", "Top-rated books this year")
    K.hdr(ss, 14, ["#", "Title", "Author", "Rating"], [4, 30, 20, 8], col0=11)
    key = LR("T")
    for k in range(10):
        r = 15 + k; pos = f"MATCH(LARGE({key},{k + 1}),{key},0)"
        ss[f"K{r}"] = k + 1
        ss[f"L{r}"] = f'=IFERROR(IF(LARGE({key},{k + 1})<0,"",INDEX({LR("A")},{pos})),"")'
        ss[f"M{r}"] = f'=IF(L{r}="","",INDEX({LR("B")},{pos}))'
        ss[f"N{r}"] = f'=IF(L{r}="","",INDEX({LR("I")},{pos}))'
    style_range(ss, 15, 24, 11, 14, fmt={14: "0.0"})
    ss.column_dimensions["F"].width = 3; ss.column_dimensions["J"].width = 3

    # ---------- Reading Calendar ----------
    K.title(cal, "Reading Calendar", "Minutes read per day from the Reading Journal – the darker, the more you read. Year from Settings.")
    cal["A3"] = '="Year "&Settings!B3'; cal["A3"].font = font(bold=True)
    K.hdr(cal, 5, ["Month"] + [str(d) for d in range(1, 32)] + ["Days read", "Minutes"])
    cal.column_dimensions["A"].width = 8
    for d in range(31): cal.column_dimensions[CL(2 + d)].width = 3.6
    cal.column_dimensions["AG"].width = 9; cal.column_dimensions["AH"].width = 9
    for m in range(12):
        r = 6 + m; cal[f"A{r}"] = MON[m]; cal[f"A{r}"].font = font(bold=True)
        for d in range(31):
            c = CL(2 + d)
            cal[f"{c}{r}"] = (f'=IF({d + 1}>DAY(EOMONTH(DATE({YR},{m + 1},1),0)),"",'
                              f'SUMIFS({JR("D")},{JR("A")},DATE({YR},{m + 1},{d + 1})))')
            cal[f"{c}{r}"].number_format = '0;-0;""'; cal[f"{c}{r}"].font = font(size=7); cal[f"{c}{r}"].border = B
            cal[f"{c}{r}"].alignment = Alignment(horizontal="center")
        cal[f"AG{r}"] = f'=COUNTIF(B{r}:AF{r},">0")'; cal[f"AH{r}"] = f"=SUM(B{r}:AF{r})"
        for c in ("AG", "AH"): cal[f"{c}{r}"].border = B; cal[f"{c}{r}"].fill = fill(SOFT); cal[f"{c}{r}"].font = font(bold=True)
        cal.row_dimensions[r].height = 20
    cal.conditional_formatting.add("B6:AF17", FormulaRule(formula=['B6=""'], fill=fill("FFFFFF")))
    cal.conditional_formatting.add("B6:AF17", ColorScaleRule(start_type="num", start_value=0, start_color="F9FAFB",
                                                             mid_type="num", mid_value=30, mid_color="F9A8D4", end_type="num", end_value=90, end_color="9D174D"))
    cal.conditional_formatting.add("B6:AF17", FormulaRule(formula=[f'AND(ISNUMBER(B6),DATE({YR},ROW()-5,COLUMN()-1)=TODAY())'], border=None,
                                                          font=Font(name=F, bold=True, size=7, color="000000", underline="single")))
    cal["A19"] = "Colour: white = no reading, pink = around 30 minutes, dark = 90 minutes or more. Today's number is underlined."
    cal["A19"].font = font(italic=True, size=9)

    # ---------- sample ----------
    if sample:
        d = dt.date(2026, 1, 3); r = L0; n = 0; jrows = []
        for (t, a, g, f_, p, rt, src, pr) in BOOKS:
            n += 1
            speed = 90 if f_ == "Audiobook" else 45
            days = max(3, round(p / (speed * 1.1))) + (n % 3)
            fin = d + dt.timedelta(days=days)
            vals = (t, a, g, f_, p, "Finished", d, fin, rt, None, src, pr,
                    "A favourite – would reread." if rt == 5 else ("Slow start, great ending." if n % 5 == 0 else None))
            for c, v in zip("ABCDEFGHIJKLM", vals):
                if v is not None: lib[f"{c}{r}"] = v
            # journal: sessions on most days
            day = d; left = p
            while day <= fin and left > 0:
                if (day.toordinal() * 7) % 10 != 0:
                    mins = 20 + (day.toordinal() * 13) % 45
                    pages = min(left, int(mins * (0.9 if f_ != "Audiobook" else 1.5)))
                    left -= pages; jrows.append((day, t, pages, mins))
                day += dt.timedelta(days=1)
            r += 1; d = fin + dt.timedelta(days=n % 2)
        start_read = [min(d, dt.date(2026, 10, 30)), dt.date(2026, 11, 2)]
        for (t, a, g, f_, p, cur, src, pr), sd in zip(READING, start_read):
            for c, v in zip("ABCDEFGJKL", (t, a, g, f_, p, "Reading", sd, cur, src, pr)):
                if v is not None: lib[f"{c}{r}"] = v
            day = sd
            while day <= dt.date(2026, 11, 9):
                jrows.append((day, t, 25 if f_ != "Audiobook" else 30, 35 if f_ != "Audiobook" else 40))
                day += dt.timedelta(days=1)
            r += 1
        t, a, g, f_, p, cur, src = DNF
        for c, v in zip("ABCDEFGJKM", (t, a, g, f_, p, "Did not finish", dt.date(2026, 6, 1), cur, src, "Not for me right now – maybe later.")):
            lib[f"{c}{r}"] = v
        r += 1
        for (t, a, g) in TBR:
            lib[f"A{r}"] = t; lib[f"B{r}"] = a; lib[f"C{r}"] = g; lib[f"F{r}"] = "Want to read"; r += 1
        jrows.sort(key=lambda x: x[0])
        # keep the most recent days continuous for a nice streak
        for k, (dd, t, pg, mi) in enumerate(jrows[:NJ]):
            jr[f"A{J0 + k}"] = dd; jr[f"B{J0 + k}"] = t; jr[f"C{J0 + k}"] = pg; jr[f"D{J0 + k}"] = mi
        jr[f"E{J0 + 2}"] = "Lizzy and Darcy – the first-ball scene is perfect."

    # ---------- Dashboard ----------
    K.title(dash, "Reading Dashboard", "Updates by itself from the Library and the Reading Journal.")
    dash["A3"] = '="Reading year "&Settings!B3&"   ·   today: "&TEXT(TODAY(),"mmm d, yyyy")'; dash["A3"].font = font(bold=True)
    cur_streak = (f'IFERROR(IF(INDEX({DI},{IDX})>0,INDEX({DJ},{IDX}),INDEX({DJ},{IDX}-1)),0)')
    kp = [("Books read this year", f"=Monthly!B18", "0"),
          ("Reading goal", f'=IF(N(Settings!B4)=0,"–",Monthly!B18/Settings!B4)', "0%"),
          ("Pages read", "=Monthly!C18", "#,##0"),
          ("Average rating", f'=IF(Monthly!D18="","–",Monthly!D18)', "0.0"),
          ("Current streak (days)", "=" + cur_streak, "0"),
          ("Longest streak (days)", f"=MAX({DJ})", "0"),
          ("Minutes this month", f'=SUMIFS({JR("D")},{JR("A")},">="&(TODAY()-DAY(TODAY())+1),{JR("A")},"<="&TODAY())', "#,##0"),
          ("Want to read", f'=COUNTIF({LR("F")},"Want to read")', "0")]
    for i, (lab, f_, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 4) * 3, 1 + (i % 4) * 2, lab, f_, fmt)
        dash.cell(row=6 + (i // 4) * 3, column=1 + (i % 4) * 2).alignment = Alignment(horizontal="left")
    dash["C7"] = '=Monthly!B18&" of "&Settings!B4&" books"'; dash["C7"].font = font(size=8, italic=True)
    dash["C10"] = '="goal "&Settings!B6&" min/day"'
    dash["C10"] = '=IFERROR("today: "&ROUND(INDEX(' + DI + ',' + IDX + '),0)&" of "&Settings!B6&" min",0)'
    dash["C10"].font = font(size=8, italic=True)
    dash["E10"] = f'=IFERROR("avg "&ROUND(SUMIFS({JR("D")},{JR("A")},">="&(TODAY()-DAY(TODAY())+1),{JR("A")},"<="&TODAY())/DAY(TODAY()),0)&" min/day","")'
    dash["E10"].font = font(size=8, italic=True)
    K.section(dash, "A12", "Currently reading")
    K.hdr(dash, 13, ["Title", "Author", "Format", "Started", "Page", "of", "Progress", "Pages/day"])
    key = LR("S")
    for k in range(6):
        r = 14 + k; pos = f"MATCH(SMALL({key},{k + 1}),{key},0)"
        dash[f"A{r}"] = f'=IFERROR(IF(SMALL({key},{k + 1})>=99999999,"",INDEX({LR("A")},{pos})),"")'
        for col, src in zip("BCDEFG", "BDGJEN"):
            dash[f"{col}{r}"] = f'=IF(A{r}="","",IF(INDEX({LR(src)},{pos})="","",INDEX({LR(src)},{pos})))'
        dash[f"H{r}"] = f'=IF(OR(A{r}="",NOT(ISNUMBER(D{r})),NOT(ISNUMBER(E{r}))),"",E{r}/MAX(1,TODAY()-D{r}+1))'
    style_range(dash, 14, 19, 1, 8, fmt={4: "mmm d", 5: "#,##0", 6: "#,##0", 7: "0%", 8: "0"})
    dash.conditional_formatting.add("G14:G19", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="F9A8D4"))
    K.section(dash, "A22", "Books per month")
    K.hdr(dash, 23, ["Month", "Books", "Goal pace", "Pages", "Minutes"])
    for m in range(12):
        r = 24 + m; dash[f"A{r}"] = MON[m]
        dash[f"B{r}"] = f"=Monthly!B{6 + m}"; dash[f"C{r}"] = f"=Monthly!$B$4/12" if False else f"=Settings!$B$4/12"
        dash[f"D{r}"] = f"=Monthly!C{6 + m}"; dash[f"E{r}"] = f"=Monthly!E{6 + m}"
    style_range(dash, 24, 35, 1, 5, fmt={2: "0", 3: "0.0", 4: "#,##0", 5: "#,##0"})
    for c, w in zip("ABCDEFGH", [30, 20, 12, 13, 12, 12, 10, 10]): dash.column_dimensions[c].width = w
    ch = BarChart(); ch.type = "col"; ch.title = "Books per month"; ch.height = 7.5; ch.width = 15; ch.y_axis.majorGridlines = None
    ch.add_data(Reference(dash, min_col=2, min_row=23, max_row=35), titles_from_data=True)
    ch.set_categories(Reference(dash, min_col=1, min_row=24, max_row=35)); ch.series[0].graphicalProperties.solidFill = MAIN; ch.gapWidth = 50
    pace = LineChart(); pace.add_data(Reference(dash, min_col=3, min_row=23, max_row=35), titles_from_data=True)
    pace.series[0].graphicalProperties.line.solidFill = "9CA3AF"; pace.series[0].graphicalProperties.line.dashStyle = "dash"
    ch += pace; ch.legend.position = "b"
    dash.add_chart(ch, "J5")
    ch2 = BarChart(); ch2.type = "bar"; ch2.title = "Books by genre"; ch2.height = 8; ch2.width = 15; ch2.legend = None
    ch2.add_data(Reference(ss, min_col=2, min_row=5, max_row=5 + len(GENRES)), titles_from_data=True)
    ch2.set_categories(Reference(ss, min_col=1, min_row=6, max_row=5 + len(GENRES)))
    ch2.x_axis.scaling.orientation = "maxMin"; ch2.y_axis.majorGridlines = None; ch2.series[0].graphicalProperties.solidFill = "EC4899"
    dash.add_chart(ch2, "J22")

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Settings: type your reading year and your goals (books, pages, minutes per day). Genres, formats and sources are editable.", False),
             ("2.  Library: add every book you want to read with Status 'Want to read' – that is your TBR list.", False),
             ("     When you start a book set 'Reading' + the start date; update 'Current page' now and then to see your progress.", False),
             ("     When you finish: set 'Finished', type the finish date and your rating (0.5 to 5 – half stars allowed).", False),
             ("3.  Reading Journal (optional but fun): log each session – date, book, pages, minutes. This builds your streak.", False),
             ("4.  Monthly: books and pages per month and whether you are ahead or behind your yearly goal.", False),
             ("5.  Stats: books by genre, format and source, money spent, your rating spread and your top-rated books of the year.", False),
             ("6.  Reading Calendar: a 12-month heatmap of the minutes you read every day.", False),
             ("7.  Dashboard: books read vs goal, pages, average rating, current and longest streak, currently reading, charts.", False),
             ("", False), ("Tips", True),
             ("•  Only books with Status 'Finished' and a finish date in your reading year count for the goal.", False),
             ("•  Audiobooks count too – type the page count of the print edition so pages stay comparable.", False),
             ("•  New year? Change the year in Settings – all stats switch to that year and your library stays.", False),
             ("•  A streak counts days with at least one journal entry; today counts once you log it.", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  600 books, 2,000 journal entries, 20 genres. No macros. Works in Excel 2010+ and Google Sheets.", False),
             ("•  To use Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("A made-up reader's 2026 with classic books. Use the blank copy in your download for your own library.", False)]
    K.start(st, "Reading Log & Book Tracker", "Your library, reading goal, streaks and stats – Excel and Google Sheets.", lines)
    finish(wb, sample, path)


META = {
    "name": "Reading Log & Book Tracker (Reading Challenge, TBR List, Ratings, Reading Streak, Stats Dashboard) – Excel & Google Sheets",
    "file_base": "Reading-Log-Book-Tracker",
    "price": 9,
    "color": "#9D174D",
    "title": "Reading Log & Book Tracker",
    "tagline": "library & TBR · yearly reading goal · ratings · streaks · stats",
    "intro": "Track every book you read, want to read or gave up on – and watch your year in books add up. Set a yearly reading goal and see if you are ahead or behind each month, rate books with half stars, log reading sessions to build a streak, and get automatic stats by genre, format and source plus a reading calendar heatmap.",
    "features": [
        "Library for 600 books: title, author, genre, format, pages, status (Want to read / Reading / Finished / Did not finish), dates, rating, source, price, notes",
        "Star ratings from 0.5 to 5 shown as ★★★★½, progress bars for books you are reading",
        "Reading challenge: books and pages per month vs the pace you need for your yearly goal",
        "Reading journal with current and longest reading streak and minutes per day",
        "Reading Calendar heatmap: 12 months × 31 days coloured by minutes read",
        "Stats: books by genre, format and source, money spent, rating spread and your top-10 books of the year",
        "Dashboard: books vs goal, pages, average rating, streaks, currently reading, books-per-month and genre charts",
        "Editable genres, formats and sources; blank copy plus a sample copy with a full example year",
    ],
    "tags": ["reading log", "book tracker", "reading tracker", "reading journal", "tbr list", "book log", "reading challenge",
             "google sheets", "excel template", "book lover gift", "reading goal", "bookish", "library tracker"],
    "summary": "Instant download: Excel + Google Sheets template, blank copy + sample copy.",
}
