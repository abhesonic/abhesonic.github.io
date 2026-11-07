"""Job Application Tracker – applications, status pipeline, interviews, networking contacts, follow-up alerts, weekly goal, source analysis.

Settings -> Applications (one row per job) -> Interviews (linked by application #) -> Contacts -> Dashboard.
"""
import datetime as dt, random
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import BarChart, Reference
from _common import Kit, F, B, MONEY, DATE, font, fill, style_range, input_cell, finish

MAIN = "0F766E"; SOFT = "CCFBF1"
K = Kit(MAIN, SOFT)
PCT = "0.0%"
NA = 300; A0 = 6; A1 = A0 + NA - 1
NI = 200; I0 = 6; I1 = I0 + NI - 1
NC = 100; C0 = 6; C1 = C0 + NC - 1
STATUSES = ["Wishlist", "Applied", "Screening", "Interviewing", "Offer", "Accepted", "Rejected", "Ghosted", "Withdrawn"]
SOURCES = ["LinkedIn", "Indeed", "Company website", "Referral", "Recruiter", "Glassdoor", "Job fair", "Other"]
NSRC = 12; SR0 = 6; SR1 = SR0 + NSRC - 1
NW = 16


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); ap = wb.create_sheet("Applications"); iv = wb.create_sheet("Interviews")
    ct = wb.create_sheet("Contacts"); s = wb.create_sheet("Settings")

    # ---------- Settings ----------
    K.title(s, "Settings", "Yellow cells are yours to edit.")
    SET = [("Your name", "Jordan Lee" if sample else "My name"), ("Target role", "Marketing Manager" if sample else "My target role"),
           ("Search start date (a Monday)", dt.date(2026, 8, 3) if sample else dt.date(2026, 11, 9)),
           ("Weekly application goal", 4 if sample else 10), ("Days with no reply before a nudge", 14)]
    for i, (lab, v) in enumerate(SET, 4):
        s[f"D{i}"] = lab; s[f"D{i}"].font = font(bold=True); s[f"D{i}"].border = B; input_cell(s[f"E{i}"], v)
    s["E6"].number_format = DATE
    K.hdr(s, 5, ["Job sources"])
    for i in range(NSRC):
        input_cell(s[f"A{SR0 + i}"], SOURCES[i] if i < len(SOURCES) else None)
    s[f"A{SR1 + 2}"] = "Add or rename sources – they appear in the drop-down on Applications."; s[f"A{SR1 + 2}"].font = font(italic=True, size=9)
    K.hdr(s, 11, ["Status (fixed list)", "Meaning"], col0=4)
    MEAN = ["Saved, not applied yet", "Application sent, no reply", "Recruiter / HR call", "In interviews", "Offer received",
            "Offer accepted", "Turned down", "No reply after the nudge period", "You stepped out"]
    for i, (a, b) in enumerate(zip(STATUSES, MEAN)):
        s.cell(row=12 + i, column=4, value=a); s.cell(row=12 + i, column=5, value=b)
    style_range(s, 12, 11 + len(STATUSES), 4, 5)
    for c, w in zip("ABCDE", [22, 3, 3, 32, 32]): s.column_dimensions[c].width = w
    START = "Settings!$E$6"; GOAL = "Settings!$E$7"; NUDGE = "Settings!$E$8"
    SRC = f"Settings!$A${SR0}:$A${SR1}"

    # ---------- Applications ----------
    K.title(ap, "Applications", "One row per job. Yellow = you type, grey = automatic. Update Status and Last update as things move.")
    cols = ["#", "Date applied", "Company", "Job title", "Location", "Work type", "Source", "Job link", "Salary min", "Salary max",
            "Status", "Last update", "Next follow-up", "Contact person", "Interest (1-5)", "Notes",
            "Interviews", "Days since applied", "Next action", "Got interview", "Got reply", "Offer", "Applied"]
    K.hdr(ap, 5, cols, [5, 12, 20, 24, 15, 10, 15, 22, 11, 11, 13, 12, 13, 18, 9, 28, 10, 10, 30, 8, 8, 7, 8])
    style_range(ap, A0, A1, 1, 23, inputs=range(2, 17), fmt={2: DATE, 9: MONEY, 10: MONEY, 12: DATE, 13: DATE, 17: "0", 18: "0"})
    for r in range(A0, A1 + 1):
        ap[f"A{r}"] = r - A0 + 1
        ap[f"Q{r}"] = f'=IF(C{r}="","",COUNTIF(Interviews!$A${I0}:$A${I1},A{r}))'
        ap[f"R{r}"] = f'=IF(OR(C{r}="",B{r}=""),"",TODAY()-B{r})'
        ap[f"T{r}"] = f'=IF(C{r}="","",IF(OR(N(Q{r})>0,K{r}="Interviewing",K{r}="Offer",K{r}="Accepted"),1,0))'
        ap[f"U{r}"] = f'=IF(C{r}="","",IF(OR(T{r}=1,K{r}="Screening",K{r}="Rejected"),1,0))'
        ap[f"V{r}"] = f'=IF(C{r}="","",IF(OR(K{r}="Offer",K{r}="Accepted"),1,0))'
        ap[f"W{r}"] = f'=IF(OR(C{r}="",B{r}="",K{r}="Wishlist"),0,1)'
        ap[f"S{r}"] = (f'=IF(C{r}="","",IF(OR(K{r}="Rejected",K{r}="Withdrawn",K{r}="Accepted",K{r}="Ghosted"),"Closed",'
                       f'IF(OR(K{r}="Wishlist",K{r}=""),"Apply",IF(K{r}="Offer","Decide on the offer",'
                       f'IF(AND(ISNUMBER(M{r}),M{r}<=TODAY()),"Follow up now",'
                       f'IF(AND(K{r}="Applied",ISNUMBER(R{r}),TODAY()-MAX(B{r},N(L{r}))>={NUDGE}),"No reply – nudge or mark Ghosted","Waiting"))))))')
    for r in range(A0, A1 + 1): ap[f"S{r}"].font = font(bold=True)
    dv = DataValidation(type="list", formula1=f'"{",".join(STATUSES)}"', allow_blank=True); ap.add_data_validation(dv); dv.add(f"K{A0}:K{A1}")
    dv = DataValidation(type="list", formula1=f"={SRC}", allow_blank=True); ap.add_data_validation(dv); dv.add(f"G{A0}:G{A1}")
    dv = DataValidation(type="list", formula1='"Remote,Hybrid,On-site"', allow_blank=True); ap.add_data_validation(dv); dv.add(f"F{A0}:F{A1}")
    dv = DataValidation(type="whole", operator="between", formula1="1", formula2="5", allow_blank=True); ap.add_data_validation(dv); dv.add(f"O{A0}:O{A1}")
    rng = f"A{A0}:W{A1}"
    ap.conditional_formatting.add(f"S{A0}:S{A1}", FormulaRule(formula=[f'OR(S{A0}="Follow up now",LEFT(S{A0},8)="No reply")'],
                                                             font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    ap.conditional_formatting.add(f"S{A0}:S{A1}", FormulaRule(formula=[f'S{A0}="Decide on the offer"'], font=Font(name=F, bold=True, color="15803D"), fill=fill("DCFCE7")))
    ap.conditional_formatting.add(f"K{A0}:K{A1}", FormulaRule(formula=[f'OR(K{A0}="Offer",K{A0}="Accepted")'], font=Font(name=F, bold=True, color="15803D")))
    ap.conditional_formatting.add(f"K{A0}:K{A1}", FormulaRule(formula=[f'OR(K{A0}="Rejected",K{A0}="Ghosted",K{A0}="Withdrawn")'], font=Font(name=F, color="9CA3AF")))
    ap.conditional_formatting.add(f"K{A0}:K{A1}", FormulaRule(formula=[f'K{A0}="Interviewing"'], font=Font(name=F, bold=True, color="1D4ED8")))
    ap.freeze_panes = f"D{A0}"; ap.auto_filter.ref = f"A5:W{A1}"

    # ---------- Interviews ----------
    K.title(iv, "Interviews", "Type the application # from the Applications tab – company and job fill in by themselves.")
    K.hdr(iv, 5, ["App #", "Company", "Job title", "Date", "Time", "Round", "Format", "Interviewer(s)", "Outcome", "Thank-you sent", "Questions asked / notes"],
          [7, 20, 24, 12, 8, 14, 10, 22, 11, 11, 40])
    style_range(iv, I0, I1, 1, 11, inputs=(1, 4, 5, 6, 7, 8, 9, 10, 11), fmt={4: DATE, 5: "hh:mm"})
    for r in range(I0, I1 + 1):
        iv[f"B{r}"] = f'=IF(A{r}="","",IFERROR(INDEX(Applications!$C${A0}:$C${A1},MATCH(A{r},Applications!$A${A0}:$A${A1},0)),"?"))'
        iv[f"C{r}"] = f'=IF(A{r}="","",IFERROR(INDEX(Applications!$D${A0}:$D${A1},MATCH(A{r},Applications!$A${A0}:$A${A1},0)),"?"))'
    for col, lst in (("F", "Phone screen,1st round,2nd round,Technical,Take-home task,Final round"), ("G", "Phone,Video,On-site"),
                     ("I", "Pending,Passed,Rejected,Offer"), ("J", "Yes,No")):
        dv = DataValidation(type="list", formula1=f'"{lst}"', allow_blank=True); iv.add_data_validation(dv); dv.add(f"{col}{I0}:{col}{I1}")
    iv.conditional_formatting.add(f"J{I0}:J{I1}", FormulaRule(formula=[f'AND(A{I0}<>"",ISNUMBER(D{I0}),D{I0}<TODAY(),J{I0}<>"Yes")'],
                                                             font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    iv.conditional_formatting.add(f"D{I0}:D{I1}", FormulaRule(formula=[f'AND(ISNUMBER(D{I0}),D{I0}>=TODAY(),D{I0}-TODAY()<=7)'],
                                                             font=Font(name=F, bold=True, color="1D4ED8"), fill=fill("DBEAFE")))
    iv.freeze_panes = f"B{I0}"

    # ---------- Contacts ----------
    K.title(ct, "Networking Contacts", "Recruiters, referrals and people you met. Set a follow-up date – it turns red when it is due.")
    K.hdr(ct, 5, ["Name", "Company", "Role", "Email / LinkedIn", "How we met", "Last contact", "Next follow-up", "Notes", "Due?"],
          [20, 20, 20, 28, 18, 12, 13, 36, 12])
    style_range(ct, C0, C1, 1, 9, inputs=range(1, 9), fmt={6: DATE, 7: DATE})
    for r in range(C0, C1 + 1):
        ct[f"I{r}"] = f'=IF(OR(A{r}="",NOT(ISNUMBER(G{r}))),"",IF(G{r}<=TODAY(),"Reach out",""))'
    ct.conditional_formatting.add(f"I{C0}:I{C1}", FormulaRule(formula=[f'I{C0}="Reach out"'], font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    ct.freeze_panes = f"B{C0}"

    # ---------- sample ----------
    if sample:
        rnd = random.Random(11)
        comps = ["Northwind Foods", "Brightline Media", "Kestrel Labs", "Harbor Health", "Summit Outdoor", "Bluepeak Software", "Orchid Retail",
                 "Atlas Logistics", "Cobalt Bank", "Greenfield Energy", "Lumen Analytics", "Maple & Co.", "Nova Travel", "Pinecrest Hotels",
                 "Quartz Apparel", "Riverstone Group", "Sparrow Apps", "Tidewater Insurance", "Urban Nest", "Vertex Robotics", "Willow Pet Co.",
                 "Zenith Fitness", "Copperleaf Games", "Driftwood Coffee", "Evergreen Schools", "Falcon Freight", "Granite Legal", "Horizon Telecom",
                 "Ironclad Security", "Juniper Beauty", "Keystone Realty", "Lakeside Clinic", "Meridian Cars", "Nimbus Cloud", "Oakridge Foods",
                 "Pioneer Media", "Redwood Bank", "Silverline Tours", "Trident Marine", "Unity Charity", "Vivid Studio", "Wavelength Audio"]
        titles = ["Marketing Manager", "Senior Marketing Manager", "Digital Marketing Manager", "Growth Marketing Lead", "Brand Manager",
                  "Content Marketing Manager", "Product Marketing Manager"]
        locs = ["Remote (US)", "Austin, TX", "Denver, CO", "Chicago, IL", "Seattle, WA", "New York, NY", "Remote"]
        start = dt.date(2026, 8, 3); n = len(comps)
        fate = (["Rejected"] * 11 + ["Ghosted"] * 7 + ["Applied"] * 9 + ["Screening"] * 4 + ["Interviewing"] * 4 + ["Offer"] * 1 +
                ["Withdrawn"] * 1 + ["Wishlist"] * 3 + ["Rejected"] * 2)
        rnd.shuffle(fate)
        ivr = I0
        for k in range(n):
            r = A0 + k; day = start + dt.timedelta(days=int(k * 94 / n) + rnd.randint(0, 2))
            stt = fate[k]
            if k >= n - 6 and stt in ("Ghosted", "Rejected"): stt = "Applied"
            mn = rnd.choice([70, 75, 80, 85, 90, 95]) * 1000
            vals = {"B": None if stt == "Wishlist" else day, "C": comps[k], "D": rnd.choice(titles), "E": rnd.choice(locs),
                    "F": rnd.choice(["Remote", "Hybrid", "Hybrid", "On-site"]),
                    "G": rnd.choice(SOURCES[:6] + ["LinkedIn", "LinkedIn", "Referral"]), "H": f"careers.{comps[k].split()[0].lower()}.example/jobs",
                    "I": mn, "J": mn + rnd.choice([10, 15, 20, 25]) * 1000, "K": stt,
                    "L": None if stt in ("Wishlist", "Applied") else day + dt.timedelta(days=rnd.randint(5, 20)),
                    "O": rnd.randint(2, 5)}
            if stt in ("Applied",) and k % 3 == 0: vals["M"] = day + dt.timedelta(days=10)
            if stt in ("Screening", "Interviewing"): vals["M"] = dt.date(2026, 11, 12) + dt.timedelta(days=k % 6)
            if stt == "Offer": vals["P"] = "Offer $96k + 10% bonus – reply by Nov 13"
            if vals["G"] == "Referral": vals["N"] = "Sam Patel"
            for c, v in vals.items():
                if v is not None: ap[f"{c}{r}"] = v
            rounds = {"Screening": 1, "Interviewing": 2, "Offer": 3}.get(stt, 1 if (stt == "Rejected" and k % 2 == 0) else 0)
            for j in range(rounds):
                d = day + dt.timedelta(days=8 + 7 * j)
                rnd_name = ["Phone screen", "1st round", "Final round"][j]
                outcome = "Pending" if (j == rounds - 1 and stt in ("Screening", "Interviewing")) else ("Rejected" if stt == "Rejected" else ("Offer" if (stt == "Offer" and j == 2) else "Passed"))
                for c, v in zip("ADEFGHIJ", [k + 1, d, dt.time(10 + j, 0), rnd_name, ["Phone", "Video", "Video"][j], ["HR recruiter", "Hiring manager", "VP Marketing"][j],
                                              outcome, "Yes" if (k + j) % 4 else "No"]):
                    iv[f"{c}{ivr}"] = v
                ivr += 1
        cts = [("Sam Patel", "Brightline Media", "Recruiter", "linkedin.com/in/sampatel-example", "LinkedIn message", dt.date(2026, 10, 20), dt.date(2026, 11, 3)),
               ("Rita Gomez", "Kestrel Labs", "Head of Growth", "rita@kestrel.example", "Ex-colleague", dt.date(2026, 10, 2), dt.date(2026, 11, 20)),
               ("Owen Clarke", "Nimbus Cloud", "Marketing Director", "linkedin.com/in/owenclarke-example", "Webinar", dt.date(2026, 9, 15), dt.date(2026, 10, 15)),
               ("Mei Tanaka", "Agency – TalentBridge", "Recruiter", "mei@talentbridge.example", "She contacted me", dt.date(2026, 11, 2), dt.date(2026, 11, 16)),
               ("Luis Ortega", "Summit Outdoor", "Brand Lead", "linkedin.com/in/luisortega-example", "Alumni network", dt.date(2026, 8, 28), None)]
        for k, row in enumerate(cts):
            for c, v in zip("ABCDEFG", row):
                if v is not None: ct[f"{c}{C0 + k}"] = v
        ct[f"H{C0}"] = "Asked for referral to Senior role"

    # ---------- Dashboard ----------
    AR = lambda c: f"Applications!${c}${A0}:${c}${A1}"
    K.title(dash, "Job Search Dashboard", "Updates by itself from the Applications, Interviews and Contacts tabs.")
    dash["A3"] = '=Settings!E4&"  ·  target: "&Settings!E5&"  ·  search started "&TEXT(Settings!E6,"mmm d, yyyy")'; dash["A3"].font = font(bold=True)
    kp = [("Applications sent", f"=SUM({AR('W')})", "0"),
          ("Still active", f'=COUNTIF({AR("K")},"Applied")+COUNTIF({AR("K")},"Screening")+COUNTIF({AR("K")},"Interviewing")+COUNTIF({AR("K")},"Offer")', "0"),
          ("Offers", f"=SUM({AR('V')})", "0"),
          ("Reply rate", f"=IFERROR(SUM({AR('U')})/SUM({AR('W')}),0)", PCT),
          ("Interview rate", f"=IFERROR(SUM({AR('T')})/SUM({AR('W')}),0)", PCT),
          ("Follow-ups due now", f'=COUNTIF({AR("S")},"Follow up now")+COUNTIF({AR("S")},"No reply*")+COUNTIF(Contacts!$I${C0}:$I${C1},"Reach out")', "0")]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 3) * 3, 1 + (i % 3) * 2, lab, f, fmt)
        dash.cell(row=6 + (i // 3) * 3, column=1 + (i % 3) * 2).alignment = Alignment(horizontal="left")
    for c, w in zip("ABCDEFG", [24, 13, 20, 13, 20, 18, 24]): dash.column_dimensions[c].width = w
    # status pipeline
    K.section(dash, "A11", "Where your applications are")
    K.hdr(dash, 12, ["Status", "Jobs", "Share"])
    for i, stt in enumerate(STATUSES):
        r = 13 + i; dash[f"A{r}"] = stt
        dash[f"B{r}"] = f'=COUNTIF({AR("K")},A{r})'
        dash[f"C{r}"] = f"=IFERROR(B{r}/SUM($B$13:$B${12 + len(STATUSES)}),0)"
    style_range(dash, 13, 12 + len(STATUSES), 1, 3, fmt={2: "0", 3: PCT})
    # sources
    K.section(dash, "E11", "Which job sources work")
    K.hdr(dash, 12, ["Source", "Applied", "Interviews", "Interview rate"], col0=5)
    dash.column_dimensions["H"].width = 13
    for i in range(NSRC):
        r = 13 + i
        dash[f"E{r}"] = f'=IF(Settings!A{SR0 + i}="","",Settings!A{SR0 + i})'
        dash[f"F{r}"] = f'=IF(E{r}="","",SUMIFS({AR("W")},{AR("G")},E{r}))'
        dash[f"G{r}"] = f'=IF(E{r}="","",SUMIFS({AR("T")},{AR("G")},E{r}))'
        dash[f"H{r}"] = f'=IF(OR(E{r}="",N(F{r})=0),"",G{r}/F{r})'
    style_range(dash, 13, 12 + NSRC, 5, 8, fmt={6: "0", 7: "0", 8: PCT})
    # weekly
    WR = 27
    K.section(dash, f"A{WR}", "Applications per week vs. your goal")
    K.hdr(dash, WR + 1, ["Week starting", "Applications", "Goal", "% of goal"])
    for i in range(NW):
        r = WR + 2 + i
        dash[f"A{r}"] = f"={START}+{7 * i}"
        dash[f"B{r}"] = f'=COUNTIFS({AR("B")},">="&A{r},{AR("B")},"<"&(A{r}+7),{AR("W")},1)'
        dash[f"C{r}"] = f"={GOAL}"
        dash[f"D{r}"] = f"=IFERROR(B{r}/C{r},0)"
    WE = WR + 1 + NW
    style_range(dash, WR + 2, WE, 1, 4, fmt={1: "mmm d", 2: "0", 3: "0", 4: "0%"})
    dash.conditional_formatting.add(f"D{WR + 2}:D{WE}", FormulaRule(formula=[f"AND(A{WR + 2}<=TODAY(),D{WR + 2}>=1)"], font=Font(name=F, bold=True, color="15803D")))
    dash.conditional_formatting.add(f"D{WR + 2}:D{WE}", FormulaRule(formula=[f"AND(A{WR + 2}+7<=TODAY(),D{WR + 2}<1)"], font=Font(name=F, bold=True, color="B91C1C")))
    # to-do list (oldest first)
    TR = 27
    K.section(dash, f"F{TR}", "To do now (oldest first)")
    K.hdr(dash, TR + 1, ["Company", "Job title", "Status", "Action"], col0=6)
    key = AR("X")
    for r in range(A0, A1 + 1):
        ap[f"X{r}"] = f'=IF(OR(S{r}="Follow up now",LEFT(S{r},8)="No reply",S{r}="Decide on the offer"),N(B{r})+ROW()/1000000,99999999)'
    ap.column_dimensions["X"].hidden = True
    for k in range(NW):
        r = TR + 2 + k
        pos = f"MATCH(SMALL({key},{k + 1}),{key},0)"
        dash[f"F{r}"] = f'=IFERROR(IF(SMALL({key},{k + 1})>=99999999,"",INDEX({AR("C")},{pos})),"")'
        for col, src in zip("GHI", "DKS"):
            dash[f"{col}{r}"] = f'=IF(F{r}="","",INDEX({AR(src)},MATCH(SMALL({key},{k + 1}),{key},0)))'
    style_range(dash, TR + 2, TR + 1 + NW, 6, 9)
    dash.column_dimensions["I"].width = 32
    ch = BarChart(); ch.type = "bar"; ch.title = "Pipeline by status"; ch.height = 7.5; ch.width = 14; ch.legend = None
    ch.add_data(Reference(dash, min_col=2, min_row=12, max_row=12 + len(STATUSES)), titles_from_data=True)
    ch.set_categories(Reference(dash, min_col=1, min_row=13, max_row=12 + len(STATUSES)))
    ch.x_axis.scaling.orientation = "maxMin"; ch.y_axis.majorGridlines = None; ch.series[0].graphicalProperties.solidFill = MAIN
    dash.add_chart(ch, "K5")
    ch2 = BarChart(); ch2.type = "col"; ch2.title = "Applications per week"; ch2.height = 7.5; ch2.width = 14; ch2.legend = None
    ch2.add_data(Reference(dash, min_col=2, min_row=WR + 1, max_row=WE), titles_from_data=True)
    ch2.set_categories(Reference(dash, min_col=1, min_row=WR + 2, max_row=WE)); ch2.y_axis.majorGridlines = None
    ch2.series[0].graphicalProperties.solidFill = "14B8A6"; ch2.gapWidth = 40
    dash.add_chart(ch2, "K21")

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Settings tab: your name, target role, the Monday your search starts, your weekly application goal and how many days to wait before a nudge.", False),
             ("2.  Applications tab: add one row per job – date applied, company, title, source, salary range, link. Save jobs you have not applied to yet as 'Wishlist'.", False),
             ("3.  When something happens, change the Status (Applied → Screening → Interviewing → Offer / Rejected / Ghosted) and the 'Last update' date.", False),
             ("4.  'Next action' tells you what to do: Follow up now, No reply – nudge, Decide on the offer, Waiting or Closed. Set a 'Next follow-up' date to get a reminder.", False),
             ("5.  Interviews tab: type the application # (column A on Applications), the date, round and interviewer. A red cell = thank-you note not sent yet.", False),
             ("6.  Contacts tab: recruiters and people who can refer you. Add a follow-up date and it turns red when it is time to reach out.", False),
             ("7.  Dashboard: applications sent, reply rate, interview rate, offers, pipeline by status, best job sources, weekly goal and a to-do list – all automatic.", False),
             ("", False), ("Tips", True),
             ("•  Use the filter arrows on the Applications header to see only active jobs or one source.", False),
             ("•  Low interview rate from one source? Spend your time on the sources with the highest interview rate.", False),
             ("•  Paste the job description into Notes – postings often disappear after they close.", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  300 applications, 200 interviews, 100 contacts. No macros. Works in Excel 2010+ and Google Sheets.", False),
             ("•  To use Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("The companies, jobs and people are made-up examples. Use the blank copy in your download for your own search.", False)]
    K.start(st, "Job Application Tracker", "Applications, interviews, follow-ups and networking – Excel and Google Sheets.", lines)
    finish(wb, sample, path)


META = {
    "name": "Job Application Tracker (Job Search Pipeline, Interviews, Follow-Up Reminders, Networking) – Excel & Google Sheets",
    "file_base": "Job-Application-Tracker",
    "price": 12,
    "color": "#0F766E",
    "title": "Job Application Tracker",
    "tagline": "pipeline · follow-up reminders · interviews · networking · weekly goal",
    "intro": "Stop losing track of where you applied. Log every job in one place, update the status as you hear back, and the tracker tells you what to do next – follow up, nudge a silent employer or decide on an offer. It also shows your reply and interview rates, which job sources actually work and whether you hit your weekly application goal.",
    "features": [
        "Applications log for 300 jobs: company, title, location, remote/hybrid/on-site, source, link, salary range, contact and notes",
        "Status pipeline: Wishlist, Applied, Screening, Interviewing, Offer, Accepted, Rejected, Ghosted, Withdrawn",
        "Automatic 'Next action' for each job: Follow up now, No reply – nudge, Decide on the offer, Waiting",
        "Interview log linked by application number: rounds, format, interviewer, outcome, thank-you note reminder",
        "Networking contacts with follow-up dates that turn red when it is time to reach out",
        "Dashboard: applications sent, reply rate, interview rate, offers, best job sources, weekly goal vs actual, to-do list and charts",
        "Blank copy plus a sample copy filled with an example job search",
    ],
    "tags": ["job application tracker", "job search tracker", "job tracker spreadsheet", "interview tracker", "career planner", "job hunt",
             "google sheets", "excel template", "application log", "networking tracker", "resume tracker", "job search planner", "career change"],
    "summary": "Instant download: Excel + Google Sheets template, blank copy + sample copy.",
}
