"""Crypto Portfolio & Profit/Loss Tracker – transactions, average cost basis, realized + unrealized P/L, staking rewards, wallets (sample + blank).

Settings (currency, as-of, wallets) -> Coins (symbol, current price, target price) -> Transactions (buy / sell / reward / transfer, sorted by date;
running average cost per coin, realized P/L per sale, reward income) -> Dashboard (value, invested, P/L, rewards, fees, this-year table,
allocation, by wallet, charts).
"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.utils import get_column_letter as L
from _common import Kit, F, B, MONEY, DATE, font, fill, style_range, input_cell, finish

MAIN = "C2410C"; SOFT = "FDEDE3"
K = Kit(MAIN, SOFT)
NC = 30; C0 = 5; C1 = C0 + NC - 1
NT = 1000; T0 = 5; T1 = T0 + NT - 1
NW = 10; W0 = 5; W1 = W0 + NW - 1
TYPES = ["Buy", "Sell", "Reward", "Transfer"]
CO = "Coins"; TX = "Transactions"; ASOF = "Settings!$B$5"
PCT = "0.0%"; SIGNED = '+#,##0.00;-#,##0.00;0.00'; QTY = "#,##0.########"; PRICE = "#,##0.00######"
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

WALLETS = ["Coinbase", "Binance", "Kraken", "Ledger (cold wallet)", "MetaMask"]
# symbol, name, current price, target price
COINS = [("BTC", "Bitcoin", 98000, 150000), ("ETH", "Ethereum", 3400, 6000), ("SOL", "Solana", 140, 300),
         ("ADA", "Cardano", 0.45, 1.20), ("LINK", "Chainlink", 14.00, 35.00), ("DOGE", "Dogecoin", 0.16, 0.50),
         ("USDC", "USD Coin (stablecoin)", 1.00, None)]
# date, type, coin, qty, price, fee, wallet, to wallet, note
TXS = [(dt.date(2025, 1, 6), "Buy", "BTC", 0.05, 98500, 4.90, "Coinbase", None, "First buy"),
       (dt.date(2025, 1, 6), "Buy", "ETH", 1.2, 3600, 4.30, "Coinbase", None, None),
       (dt.date(2025, 2, 3), "Buy", "SOL", 10, 205, 2.05, "Binance", None, None),
       (dt.date(2025, 3, 3), "Buy", "BTC", 0.03, 86000, 2.58, "Coinbase", None, "Monthly DCA"),
       (dt.date(2025, 3, 15), "Transfer", "BTC", 0.08, 84000, 1.50, "Coinbase", "Ledger (cold wallet)", "Moved to cold storage"),
       (dt.date(2025, 4, 7), "Buy", "ADA", 2000, 0.62, 1.24, "Kraken", None, None),
       (dt.date(2025, 4, 7), "Buy", "LINK", 40, 12.50, 0.50, "Kraken", None, None),
       (dt.date(2025, 5, 5), "Buy", "BTC", 0.02, 95000, 1.90, "Coinbase", None, "Monthly DCA"),
       (dt.date(2025, 6, 1), "Reward", "SOL", 0.35, 155, None, "Binance", None, "Staking reward"),
       (dt.date(2025, 7, 14), "Buy", "DOGE", 5000, 0.19, 0.95, "Binance", None, None),
       (dt.date(2025, 8, 20), "Sell", "ETH", 0.4, 4300, 3.44, "Coinbase", None, "Took some profit"),
       (dt.date(2025, 9, 1), "Reward", "SOL", 0.30, 210, None, "Binance", None, "Staking reward"),
       (dt.date(2025, 10, 6), "Buy", "USDC", 1500, 1.00, 0, "Coinbase", None, "Dry powder"),
       (dt.date(2025, 11, 20), "Buy", "ETH", 0.5, 2900, 2.90, "Coinbase", None, "Bought the dip"),
       (dt.date(2025, 12, 15), "Sell", "DOGE", 2000, 0.14, 0.56, "Binance", None, "Cut losses"),
       (dt.date(2026, 1, 5), "Buy", "BTC", 0.015, 91000, 1.37, "Coinbase", None, "Monthly DCA"),
       (dt.date(2026, 1, 20), "Reward", "ETH", 0.02, 3300, None, "Coinbase", None, "ETH staking"),
       (dt.date(2026, 2, 9), "Buy", "SOL", 5, 120, 0.60, "Binance", None, None),
       (dt.date(2026, 3, 2), "Sell", "USDC", 500, 1.00, 0, "Coinbase", None, "Cash out"),
       (dt.date(2026, 3, 9), "Buy", "BTC", 0.01, 84000, 0.84, "Coinbase", None, "Monthly DCA"),
       (dt.date(2026, 4, 10), "Transfer", "ETH", 0.5, 3000, 2.10, "Coinbase", "MetaMask", "To self-custody"),
       (dt.date(2026, 5, 4), "Buy", "LINK", 30, 15.00, 0.45, "Kraken", None, None),
       (dt.date(2026, 6, 1), "Reward", "SOL", 0.40, 150, None, "Binance", None, "Staking reward"),
       (dt.date(2026, 7, 8), "Sell", "SOL", 4, 175, 0.70, "Binance", None, "Rebalance"),
       (dt.date(2026, 8, 3), "Buy", "BTC", 0.01, 112000, 1.12, "Coinbase", None, "Monthly DCA"),
       (dt.date(2026, 9, 1), "Reward", "SOL", 0.35, 145, None, "Binance", None, "Staking reward"),
       (dt.date(2026, 9, 22), "Sell", "ADA", 800, 0.80, 0.64, "Kraken", None, None),
       (dt.date(2026, 10, 5), "Buy", "ETH", 0.3, 3500, 1.05, "Coinbase", None, None),
       (dt.date(2026, 11, 2), "Buy", "USDC", 400, 1.00, 0, "Coinbase", None, None)]


def build(sample, path):
    wb = Workbook()
    st = wb.active; st.title = "Start Here"
    dash = wb.create_sheet("Dashboard"); co = wb.create_sheet(CO); tx = wb.create_sheet(TX); s = wb.create_sheet("Settings")

    # ---------- Settings ----------
    K.title(s, "Settings", "Yellow cells are yours to edit.")
    rows = [("Currency label", "USD", None), ("Today / as-of date", dt.date(2026, 11, 15) if sample else "=TODAY()", DATE)]
    for i, (k_, v, fmt) in enumerate(rows, 4):
        s.cell(row=i, column=1, value=k_).font = font(bold=True)
        c = s.cell(row=i, column=2); input_cell(c, v)
        if fmt: c.number_format = fmt
    s["C4"] = "The fiat currency you buy and sell in (USD, EUR, GBP, CAD, AUD...)."; s["C5"] = "Leave as =TODAY(). Its year is 'this year' on the dashboard."
    for r in (4, 5): s[f"C{r}"].font = font(italic=True, size=9)
    s["E4"] = "Wallets & exchanges"; s["E4"].font = font(bold=True); s["G4"] = "Transaction types"; s["G4"].font = font(bold=True)
    for i in range(NW):
        c = s.cell(row=W0 + i, column=5); input_cell(c, WALLETS[i] if sample and i < len(WALLETS) else None)
    for i, t in enumerate(TYPES): s.cell(row=5 + i, column=7, value=t).font = font()
    s[f"E{W1 + 2}"] = "List every exchange or wallet you use (max 10)."; s[f"E{W1 + 2}"].font = font(italic=True, size=9)
    s["G10"] = "Buy / Sell: trades for fiat.  Reward: staking, interest, airdrops (income).  Transfer: move coins between your own wallets."
    s["G10"].font = font(italic=True, size=9)
    for c, w in zip("ABCDEFG", [24, 14, 60, 3, 26, 3, 14]): s.column_dimensions[c].width = w
    WL = f"Settings!$E${W0}:$E${W1}"; TL = "Settings!$G$5:$G$8"

    # ---------- Transactions ----------
    K.title(tx, "Transactions", "One row per buy, sell, reward or transfer. Keep rows in date order (oldest first) – the average cost runs top to bottom.")
    cols = ["Date", "Type", "Coin", "Quantity", "Price per coin", "Fee (fiat)", "Wallet / exchange", "To wallet (transfers)", "Note",
            "Qty change", "Avg cost before", "Cost basis change", "Trade value", "Realized P/L", "Reward income", "Check"]
    K.hdr(tx, 4, cols, [11, 10, 8, 13, 13, 10, 18, 18, 20, 13, 13, 13, 12, 12, 12, 18])
    style_range(tx, T0, T1, 1, 16, inputs=range(1, 10),
                fmt={1: DATE, 4: QTY, 5: PRICE, 6: MONEY, 10: QTY, 11: PRICE, 12: MONEY, 13: MONEY, 14: SIGNED, 15: MONEY})
    if sample:
        for i, row in enumerate(TXS):
            for j, v in enumerate(row, 1):
                if v is not None: tx.cell(row=T0 + i, column=j, value=v)
    for r in range(T0, T1 + 1):
        e = f'C{r}=""'; p = r - 1
        held = f"SUMIFS(J${T0 - 1}:J{p},C${T0 - 1}:C{p},C{r})"
        tx[f"J{r}"] = f'=IF({e},"",IF(B{r}="Sell",-N(D{r}),IF(OR(B{r}="Buy",B{r}="Reward"),N(D{r}),0)))'
        tx[f"K{r}"] = f'=IF({e},"",IF(ABS({held})<0.000000001,0,SUMIFS(L${T0 - 1}:L{p},C${T0 - 1}:C{p},C{r})/{held}))'
        tx[f"L{r}"] = (f'=IF({e},"",IF(B{r}="Buy",N(D{r})*N(E{r})+N(F{r}),IF(B{r}="Reward",N(D{r})*N(E{r}),'
                       f'IF(B{r}="Sell",-N(D{r})*K{r},0))))')
        tx[f"M{r}"] = f'=IF({e},"",N(D{r})*N(E{r}))'
        tx[f"N{r}"] = f'=IF(OR({e},B{r}<>"Sell"),"",N(D{r})*N(E{r})-N(F{r})-N(D{r})*K{r})'
        tx[f"O{r}"] = f'=IF(OR({e},B{r}<>"Reward"),"",N(D{r})*N(E{r}))'
        tx[f"P{r}"] = (f'=IF({e},"",IF(B{r}="","Pick a type",IF(AND(OR(B{r}="Sell",B{r}="Transfer"),N(D{r})>{held}+0.000000001),"More than you hold",'
                       f'IF(AND(B{r}="Transfer",H{r}=""),"Pick the To wallet",IF(ISNA(MATCH(C{r},{CO}!$A${C0}:$A${C1},0)),"Add coin on Coins tab","OK")))))')
        tx[f"P{r}"].alignment = Alignment(horizontal="center")
    tx.conditional_formatting.add(f"P{T0}:P{T1}", FormulaRule(formula=[f'AND(P{T0}<>"",P{T0}<>"OK")'], font=Font(name=F, bold=True, color="B91C1C"), fill=fill("FEE2E2")))
    tx.conditional_formatting.add(f"P{T0}:P{T1}", FormulaRule(formula=[f'P{T0}="OK"'], font=Font(name=F, color="15803D")))
    tx.conditional_formatting.add(f"N{T0}:N{T1}", FormulaRule(formula=[f'AND(ISNUMBER(N{T0}),N{T0}<0)'], font=Font(name=F, color="B91C1C")))
    tx.conditional_formatting.add(f"N{T0}:N{T1}", FormulaRule(formula=[f'AND(ISNUMBER(N{T0}),N{T0}>0)'], font=Font(name=F, color="15803D")))
    for rng, f1 in ((f"B{T0}:B{T1}", f"={TL}"), (f"C{T0}:C{T1}", f"={CO}!$A${C0}:$A${C1}"), (f"G{T0}:G{T1}", f"={WL}"), (f"H{T0}:H{T1}", f"={WL}")):
        dv = DataValidation(type="list", formula1=f1, allow_blank=True); tx.add_data_validation(dv); dv.add(rng)
    tx[f"R4"] = "Price per coin: the price you paid or got, in your currency. For a Reward use the coin's price on that day."
    tx["R5"] = "Fee: exchange or network fee in your currency. Buy fees are added to cost basis, sell fees reduce the profit."
    tx["R6"] = "Average cost method: each sale uses the average cost of all coins you held before it."
    for r in (4, 5, 6): tx[f"R{r}"].font = font(italic=True, size=9)
    tx.freeze_panes = "D5"

    # ---------- Coins ----------
    K.title(co, "Coins", "One row per coin or token. Type the current price whenever you check (or use GOOGLEFINANCE in Google Sheets).")
    cols = ["Symbol", "Name", "Current price", "Target price", "Holdings", "Avg cost / coin", "Cost basis", "Market value",
            "Unrealized P/L", "Unrealized %", "Realized P/L", "Reward income", "Allocation", "Distance to target", "Value at target"]
    K.hdr(co, 4, cols, [9, 24, 13, 13, 15, 13, 13, 13, 13, 11, 12, 12, 10, 11, 14])
    style_range(co, C0, C1, 1, 15, inputs=range(1, 5),
                fmt={3: PRICE, 4: PRICE, 5: QTY, 6: PRICE, 7: MONEY, 8: MONEY, 9: SIGNED, 10: PCT, 11: SIGNED, 12: MONEY, 13: PCT, 14: PCT, 15: MONEY})
    if sample:
        for i, row in enumerate(COINS):
            for j, v in enumerate(row, 1):
                if v is not None: co.cell(row=C0 + i, column=j, value=v)
    TR = f"{TX}!$C${T0}:$C${T1}"; TOTV = f"$H${C1 + 1}"
    for r in range(C0, C1 + 1):
        e = f'A{r}=""'
        co[f"E{r}"] = f'=IF({e},"",ROUND(SUMIFS({TX}!$J${T0}:$J${T1},{TR},A{r}),10))'
        co[f"G{r}"] = f'=IF({e},"",IF(E{r}=0,0,ROUND(SUMIFS({TX}!$L${T0}:$L${T1},{TR},A{r}),2)))'
        co[f"F{r}"] = f'=IF(OR({e},N(E{r})=0),"",G{r}/E{r})'
        co[f"H{r}"] = f'=IF({e},"",N(E{r})*N(C{r}))'
        co[f"I{r}"] = f'=IF({e},"",H{r}-G{r})'
        co[f"J{r}"] = f'=IF(OR({e},N(G{r})=0),"",I{r}/G{r})'
        co[f"K{r}"] = f'=IF({e},"",SUMIFS({TX}!$N${T0}:$N${T1},{TR},A{r}))'
        co[f"L{r}"] = f'=IF({e},"",SUMIFS({TX}!$O${T0}:$O${T1},{TR},A{r}))'
        co[f"M{r}"] = f'=IF(OR({e},N({TOTV})=0),"",H{r}/{TOTV})'
        co[f"N{r}"] = f'=IF(OR({e},N(D{r})=0,N(C{r})=0),"",D{r}/C{r}-1)'
        co[f"O{r}"] = f'=IF(OR({e},N(D{r})=0),"",N(E{r})*D{r})'
        co[f"A{r}"].font = Font(name=F, bold=True, color="0000FF")
    T = C1 + 1
    co[f"A{T}"] = "Total"
    for c in "GHIKLO": co[f"{c}{T}"] = f"=SUM({c}{C0}:{c}{C1})"
    co[f"J{T}"] = f"=IFERROR(I{T}/G{T},0)"
    for c in "AGHIJKLO":
        x = co[f"{c}{T}"]; x.font = font(bold=True); x.border = B; x.fill = fill(SOFT)
        x.number_format = {"J": PCT, "I": SIGNED, "K": SIGNED}.get(c, MONEY)
    for rng in (f"I{C0}:K{C1}",):
        co.conditional_formatting.add(rng, FormulaRule(formula=[f'AND(ISNUMBER(I{C0}),I{C0}<0)'], font=Font(name=F, color="B91C1C")))
        co.conditional_formatting.add(rng, FormulaRule(formula=[f'AND(ISNUMBER(I{C0}),I{C0}>0)'], font=Font(name=F, color="15803D")))
    co.conditional_formatting.add(f"N{C0}:N{C1}", FormulaRule(formula=[f'AND(ISNUMBER(N{C0}),N{C0}<=0)'], fill=fill("DCFCE7"), font=Font(name=F, bold=True, color="15803D")))
    co[f"A{C1 + 3}"] = ('Google Sheets live price: in Current price type =GOOGLEFINANCE("CURRENCY:BTCUSD") (change BTC and USD). '
                        'Distance to target turns green when the price reaches your target.')
    co[f"A{C1 + 3}"].font = font(italic=True, size=9)
    co.freeze_panes = "C5"

    # ---------- Dashboard ----------
    K.title(dash, "Crypto Portfolio Dashboard", "Updates by itself from Coins and Transactions.")
    dash["A3"] = f'="As of "&TEXT({ASOF},"mmm d, yyyy")&"  ·  "&COUNTIFS({CO}!E{C0}:E{C1},">0")&" coins held  ·  Amounts in "&Settings!B4'
    dash["A3"].font = font(bold=True)
    fees = f"SUM({TX}!$F${T0}:$F${T1})"
    kp = [("Portfolio value", f"={CO}!H{T}", MONEY),
          ("Cost basis (still invested)", f"={CO}!G{T}", MONEY),
          ("Unrealized P/L", f"={CO}!I{T}", SIGNED),
          ("Realized P/L (all time)", f"={CO}!K{T}", SIGNED),
          ("Reward income (all time)", f"={CO}!L{T}", MONEY),
          ("Total P/L (unrealized + realized)", f"=E6+A9", SIGNED)]
    for i, (lab, f, fmt) in enumerate(kp):
        K.kpi(dash, 5 + (i // 3) * 3, 1 + (i % 3) * 2, lab, f, fmt)
        dash.cell(row=6 + (i // 3) * 3, column=1 + (i % 3) * 2).alignment = Alignment(horizontal="left")
    for c, w in zip("ABCDEFGH", [30, 16, 30, 16, 30, 16, 4, 16]): dash.column_dimensions[c].width = w

    yr0 = f'DATE(YEAR({ASOF}),1,1)'; yr1 = f'DATE(YEAR({ASOF})+1,1,1)'
    def yr(col, extra=""):
        return f'SUMIFS({TX}!${col}${T0}:${col}${T1},{TX}!$A${T0}:$A${T1},">="&{yr0},{TX}!$A${T0}:$A${T1},"<"&{yr1}{extra})'
    K.section(dash, "A11", "This year")
    yrows = [("Bought (incl. fees)", f'={yr("L", f",{TX}!$B${T0}:$B${T1},\"Buy\"")}', MONEY),
             ("Sold (trade value)", f'={yr("M", f",{TX}!$B${T0}:$B${T1},\"Sell\"")}', MONEY),
             ("Realized P/L", f'={yr("N")}', SIGNED),
             ("Reward income", f'={yr("O")}', MONEY),
             ("Fees paid", f'={yr("F")}', MONEY),
             ("Fees paid (all time)", f"={fees}", MONEY),
             ("Unrealized return %", f"={CO}!J{T}", PCT)]
    for i, (lab, f, fmt) in enumerate(yrows, 12):
        dash[f"A{i}"] = lab; dash[f"A{i}"].font = font(bold=True); dash[f"A{i}"].border = B; dash[f"A{i}"].fill = fill(SOFT)
        c = dash[f"B{i}"]; c.value = f; c.number_format = fmt; c.border = B; c.font = Font(name=F, bold=True, color=MAIN)
        c.alignment = Alignment(horizontal="right")
    dash["A19"] = "Realized P/L and rewards may be taxable where you live – this sheet helps you collect the numbers, it is not tax advice."
    dash["A19"].font = font(italic=True, size=9)

    # month table
    K.section(dash, "A21", "Month by month (this year)")
    K.hdr(dash, 22, ["Month", "Bought", "Sold", "Realized P/L", "Rewards"])
    for m in range(12):
        r = 23 + m
        dash[f"A{r}"] = f"=DATE(YEAR({ASOF}),{m + 1},1)"
        rng = f'{TX}!$A${T0}:$A${T1},">="&$A{r},{TX}!$A${T0}:$A${T1},"<"&DATE(YEAR($A{r}),MONTH($A{r})+1,1)'
        dash[f"B{r}"] = f'=SUMIFS({TX}!$L${T0}:$L${T1},{rng},{TX}!$B${T0}:$B${T1},"Buy")'
        dash[f"C{r}"] = f'=SUMIFS({TX}!$M${T0}:$M${T1},{rng},{TX}!$B${T0}:$B${T1},"Sell")'
        dash[f"D{r}"] = f'=SUMIFS({TX}!$N${T0}:$N${T1},{rng})'
        dash[f"E{r}"] = f'=SUMIFS({TX}!$O${T0}:$O${T1},{rng})'
    dash["A35"] = "Total"
    for c in "BCDE": dash[f"{c}35"] = f"=SUM({c}23:{c}34)"
    style_range(dash, 23, 35, 1, 5, fmt={1: "mmm yyyy", 2: MONEY, 3: MONEY, 4: SIGNED, 5: MONEY})
    for c in "ABCDE": dash[f"{c}35"].font = font(bold=True); dash[f"{c}35"].fill = fill(SOFT)

    # by wallet
    K.section(dash, "A37", "By wallet / exchange (value at current prices)")
    K.hdr(dash, 38, ["Wallet / exchange", "Value", "Share"])
    syms = f"{CO}!$A${C0}:$A${C1}"; prices = f"{CO}!$C${C0}:$C${C1}"
    for i in range(NW):
        r = 39 + i; k = W0 + i
        dash[f"A{r}"] = f'=IF(Settings!E{k}="","",Settings!E{k})'
        held = (f'SUMIFS({TX}!$J${T0}:$J${T1},{TR},{syms},{TX}!$G${T0}:$G${T1},A{r})'
                f'-SUMIFS({TX}!$D${T0}:$D${T1},{TR},{syms},{TX}!$G${T0}:$G${T1},A{r},{TX}!$B${T0}:$B${T1},"Transfer")'
                f'+SUMIFS({TX}!$D${T0}:$D${T1},{TR},{syms},{TX}!$H${T0}:$H${T1},A{r},{TX}!$B${T0}:$B${T1},"Transfer")')
        dash[f"B{r}"] = f'=IF(A{r}="",0,SUMPRODUCT({held},{prices}))'
        dash[f"C{r}"] = f"=IFERROR(B{r}/{CO}!$H${T},0)"
    style_range(dash, 39, 38 + NW, 1, 3, fmt={2: MONEY, 3: PCT})

    ch = BarChart(); ch.type = "col"; ch.grouping = "clustered"; ch.title = "Bought vs sold per month (this year)"
    ch.height = 8; ch.width = 18; ch.y_axis.majorGridlines = None; ch.roundedCorners = False
    ch.x_axis.delete = False; ch.y_axis.delete = False; ch.legend.overlay = False
    ch.add_data(Reference(dash, min_col=2, max_col=3, min_row=22, max_row=34), titles_from_data=True)
    ch.set_categories(Reference(dash, min_col=1, min_row=23, max_row=34)); ch.x_axis.number_format = "mmm"
    for ser, col in zip(ch.series, ("FDBA74", MAIN)): ser.graphicalProperties.solidFill = col
    dash.add_chart(ch, "H5")
    pie = PieChart(); pie.title = "Allocation by coin"; pie.height = 8; pie.width = 13; pie.roundedCorners = False; pie.legend.overlay = False
    pie.add_data(Reference(co, min_col=8, min_row=C0, max_row=C0 + 11)); pie.set_categories(Reference(co, min_col=1, min_row=C0, max_row=C0 + 11))
    dash.add_chart(pie, "H22")

    # ---------- Start Here ----------
    lines = [("How to start", True),
             ("1.  Settings tab: set your currency and list the exchanges and wallets you use.", False),
             ("2.  Coins tab: one row per coin (symbol + name). Type the current price whenever you check, and a target price if you have one.", False),
             ("3.  Transactions tab: log every buy, sell, staking/airdrop reward and transfer between your wallets – oldest first (if you add an old trade later, sort by Date). "
              "The Check column tells you if something is missing or if you sell more than you hold.", False),
             ("4.  Coins tab fills itself: holdings, average cost, cost basis, value, unrealized and realized profit/loss, rewards, allocation and distance to target.", False),
             ("5.  Dashboard: portfolio value, total P/L, this year's buys, sales, realized profit, rewards and fees, month by month, by wallet, charts.", False),
             ("", False), ("How profit is calculated", True),
             ("•  Average cost method: a sale's profit = sale value − fee − (coins sold × average cost of the coins you held just before).", False),
             ("•  Buy fees are added to cost basis. Rewards are counted as income at the price on the day you got them, and become part of your cost basis. Transfers between your own wallets make no profit (fee counts under Fees paid).", False),
             ("", False), ("Tips", True),
             ("•  Google Sheets: =GOOGLEFINANCE(\"CURRENCY:BTCUSD\") in a Current price cell fetches a live price (change BTC/USD to your coin and currency).", False),
             ("", False), ("Colour key", True),
             ("Yellow cells are inputs: type here.  Grey cells calculate themselves: please don't overwrite them.", False),
             ("", False), ("Good to know", True),
             ("•  30 coins, 1,000 transactions, 10 wallets, any currency. A tracking tool – not financial or tax advice.", False),
             ("•  To use Google Sheets: open drive.google.com, upload this file, then choose Open with > Google Sheets.", False)]
    if sample:
        lines += [("", False), ("This copy contains sample data", True),
                  ("The coins, prices and trades are made-up examples (as-of date Nov 15, 2026) – not current market data and not a recommendation. "
                   "Use the blank copy in your download for your own portfolio.", False)]
    K.start(st, "Crypto Portfolio Tracker", "Profit & loss, average cost basis, staking rewards and wallets. Works in Excel and Google Sheets.", lines)
    finish(wb, sample, path)


META = {
    "name": "Crypto Portfolio & Profit/Loss Tracker (Cost Basis, Staking Rewards, Wallets) – Excel & Google Sheets",
    "file_base": "Crypto-Portfolio-Profit-Loss-Tracker",
    "price": 17,
    "color": "#C2410C",
    "title": "Crypto Portfolio Tracker",
    "tagline": "Profit/loss · average cost basis · rewards · wallets — Excel & Google Sheets",
    "intro": "Know exactly where your crypto stands. Log your buys, sells, staking rewards and wallet transfers once, and see your holdings, average cost, realized and unrealized profit, rewards income and fees – per coin, per wallet and per month.",
    "features": [
        "Transaction log for up to 1,000 buys, sells, rewards (staking, interest, airdrops) and transfers between your wallets",
        "Average cost basis per coin, worked out automatically after every trade, with realized profit/loss on each sale",
        "Coins sheet for up to 30 coins: holdings, cost basis, market value, unrealized P/L and %, allocation",
        "Target prices: distance to target and value at target, turns green when a coin hits your target",
        "Dashboard: portfolio value, total P/L, this year's buys, sales, realized profit, rewards and fees, month by month",
        "Value by wallet or exchange (Coinbase, Binance, Kraken, Ledger, MetaMask... up to 10)",
        "Built-in checks: warns if you sell more than you hold or forget a field. Any currency, no macros, blank copy plus a sample copy",
    ],
    "tags": ["crypto tracker", "crypto portfolio", "bitcoin tracker", "crypto spreadsheet", "cost basis", "profit loss tracker", "staking rewards",
             "google sheets", "excel template", "investment tracker", "crypto tax", "portfolio tracker", "personal finance"],
}
