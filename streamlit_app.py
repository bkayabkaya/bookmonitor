# ============================================================================
#  BOOK MONITOR — track record dashboard (Streamlit)
#  Reads a two-sheet workbook ('Trades' + 'Daily PnL') and renders the
#  Excel-driven "Book" view. No Interactive Brokers, no local machine needed.
#
#  Deploy on Streamlit Community Cloud: push this folder to GitHub, then
#  share.streamlit.io -> New app -> pick streamlit_app.py.
# ============================================================================

import os
import io
import datetime as dt
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ---------------------------------------------------------------------------
#  CONFIG
# ---------------------------------------------------------------------------
TRADES_FILE  = os.environ.get("TRADES_FILE", "track_record.xlsx")  # bundled in the repo
TRADES_SHEET = "Trades"
DAILY_SHEET  = "Daily PnL"
EXPOSURE_SHEET = "exposure"
ANN = 252

INK, PANEL, PANEL2, LINE = "#0E1116", "#161B22", "#1C232D", "#263140"
TXT, MUTED, BRASS = "#E6EAF0", "#8A94A6", "#C8A25A"
UP, DOWN, BLUE = "#3FB77E", "#E5606E", "#5B8DEF"
# Daily-exposure bar colours, keyed by the Trades sheet's Asset Class values.
# Unmapped classes fall back to _EXPO_PALETTE (cycled).
ASSET_COLORS = {
    "Equities":   "#F0A830",  # amber
    "ETFs":       "#8E2323",  # dark red
    "Metals":     "#8B6914",  # darker yellow
    "Currency":   "#1868B7",  # Greek blue
    "Softs":      "#4C8C63",  # darker pastel green
    "Volatility": "#D64550",  # vivid red
    "Sectoral":   "#7C5CBF",  # purple
    "Index":      "#6E7681",  # dark gray
    "Cash":       "#1E5631",  # dark green
}
_EXPO_PALETTE = ["#F0A830", "#1868B7", "#8E2323", "#4C8C63", "#7C5CBF",
                 "#6E7681", "#8B6914", "#D64550", "#C8A25A", "#5FB3B3"]


def _asset_color(name, i):
    return ASSET_COLORS.get(name, _EXPO_PALETTE[i % len(_EXPO_PALETTE)])


MONO = "IBM Plex Mono, ui-monospace, monospace"
FONT = "IBM Plex Sans, system-ui, sans-serif"
PLOT = dict(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            font=dict(family=MONO, color=MUTED, size=11), margin=dict(l=48, r=18, t=28, b=32),
            xaxis=dict(gridcolor=LINE, zerolinecolor=LINE, linecolor=LINE),
            yaxis=dict(gridcolor=LINE, zerolinecolor=LINE, linecolor=LINE),
            hoverlabel=dict(font_family=MONO, bgcolor=PANEL2), showlegend=False)

FUT_MULTIPLIERS = {
    "VXM": 100, "VX": 1000, "SI": 5000, "QI": 2500, "SIL": 1000, "GC": 100, "QO": 50,
    "MGC": 10, "HG": 25000, "ZC": 5000, "ZS": 5000, "ZW": 5000, "ZL": 60000, "ZM": 100,
    "CC": 10, "KC": 37500, "SB": 112000, "EUR": 125000, "6E": 125000, "SF": 125000,
    "CHF": 125000, "6S": 125000, "6B": 62500, "6J": 12500000, "6A": 100000, "6C": 100000,
    "ES": 50, "NQ": 20, "YM": 5, "RTY": 50, "CL": 1000, "NG": 10000, "RB": 42000, "HO": 42000,
}
_MONTH_CODES = {"F": 1, "G": 2, "H": 3, "J": 4, "K": 5, "M": 6,
                "N": 7, "Q": 8, "U": 9, "V": 10, "X": 11, "Z": 12}
_REQUIRED_COLS = ["Entry Date", "Exit Date", "Symbol", "IB_TICKER", "Quantity",
                  "Entry Price", "Exit Price", "Notional", "Leg PnL", "Trade ID", "Trade Type"]


# ===========================================================================
#  ENGINE  (identical logic to the Dash app's Book tab)
# ===========================================================================
def _to_num(series):
    return pd.to_numeric(series.astype(str).str.strip().str.replace(",", "", regex=False)
                         .replace({"OPEN": np.nan, "": np.nan}), errors="coerce")


def _root_symbol(s):
    s = str(s)
    if len(s) >= 2 and s[-1].isdigit() and s[-2].upper() in _MONTH_CODES:
        return s[:-2]
    return s


def _excel(src):
    return pd.ExcelFile(io.BytesIO(src) if isinstance(src, (bytes, bytearray)) else src)


def _read_blotter(src):
    xls = _excel(src)
    sheet = next((s for s in xls.sheet_names
                  if s.strip().lower() == TRADES_SHEET.lower()), xls.sheet_names[0])
    probe = pd.read_excel(xls, sheet_name=sheet, header=None, dtype=str)
    hr = 0
    for i in range(min(len(probe), 30)):
        if "Entry Date" in [str(v).strip() for v in probe.iloc[i].tolist()]:
            hr = i
            break
    return pd.read_excel(xls, sheet_name=sheet, header=hr, dtype=str)


def load_daily_pnl(src):
    xls = _excel(src)
    name = next((s for s in xls.sheet_names if s.strip().lower() == DAILY_SHEET.lower()), None)
    if name is None:
        return None
    probe = pd.read_excel(xls, sheet_name=name, header=None, dtype=str)
    hr = 0
    for i in range(min(len(probe), 30)):
        if "Date" in [str(v).strip() for v in probe.iloc[i].tolist()]:
            hr = i
            break
    raw = pd.read_excel(xls, sheet_name=name, header=hr, dtype=str)
    raw.columns = [str(c).strip() for c in raw.columns]
    if "Date" not in raw.columns:
        return None
    bal = next((c for c in raw.columns if "balance" in c.lower()), None)
    pnl = next((c for c in raw.columns if c.lower().strip() in
                ("pnl", "p&l", "daily pnl", "profit", "return")), None)
    d = pd.DataFrame()
    d["date"] = pd.to_datetime(raw["Date"], errors="coerce")
    d["balance"] = _to_num(raw[bal]) if bal else np.nan
    d["pnl"] = _to_num(raw[pnl]) if pnl else np.nan
    d = d.dropna(subset=["date"]).sort_values("date").set_index("date")
    if d.empty:
        return None
    if d["pnl"].isna().all() and d["balance"].notna().any():
        d["pnl"] = d["balance"].diff().fillna(0.0)
    if d["balance"].isna().all() and d["pnl"].notna().any():
        d["balance"] = d["pnl"].cumsum()
    d["pnl"] = d["pnl"].fillna(0.0)
    prev = d["balance"].shift(1)
    d["ret"] = (d["pnl"] / prev).replace([np.inf, -np.inf], np.nan).fillna(0.0)
    return d


def load_exposure(src):
    xls = _excel(src)
    name = next((s for s in xls.sheet_names
                 if s.strip().lower() == EXPOSURE_SHEET.lower()), None)
    if name is None:
        return None
    probe = pd.read_excel(xls, sheet_name=name, header=None, dtype=str)
    hr = 0
    for i in range(min(len(probe), 30)):
        if "Date" in [str(v).strip() for v in probe.iloc[i].tolist()]:
            hr = i
            break
    raw = pd.read_excel(xls, sheet_name=name, header=hr, dtype=str)
    raw.columns = [str(c).strip() for c in raw.columns]
    if "Date" not in raw.columns:
        return None

    def _col(*names):
        for n in names:
            for c in raw.columns:
                if c.lower().strip() == n:
                    return c
        return None

    # display name -> possible header spellings for its *_Exposure fraction column
    wanted = {
        "Equities":   ("equity_exposure", "equities_exposure"),
        "Cash":       ("cash_exposure",),
        "Currency":   ("currency_exposure",),
        "Metals":     ("metals_exposure", "metal_exposure"),
        "Softs":      ("softs_exposure", "soft_exposure"),
        "Volatility": ("volatility_exposure", "vol_exposure"),
        "Index":      ("index_exposure",),
        "ETFs":       ("etf_exposure", "etfs_exposure"),
    }
    d = pd.DataFrame()
    d["date"] = pd.to_datetime(raw["Date"], errors="coerce")
    for disp, names in wanted.items():
        c = _col(*names)
        if c is not None:
            d[disp] = _to_num(raw[c])
    d = d.dropna(subset=["date"]).sort_values("date").set_index("date")
    d.index = d.index.normalize()
    if d.empty or d.shape[1] == 0:
        return None
    # normalise each row to 100% (guards against rounding in the sheet)
    frac = d.fillna(0.0)
    tot = frac.sum(axis=1).replace(0, np.nan)
    norm = frac.div(tot, axis=0).fillna(0.0)
    return norm.loc[:, (norm != 0).any(axis=0)]  # drop never-used categories


def load_trades(src):
    raw = _read_blotter(src)
    raw.columns = [str(c).strip() for c in raw.columns]
    missing = [c for c in _REQUIRED_COLS if c not in raw.columns]
    if missing:
        raise ValueError("Workbook is missing column(s): " + ", ".join(missing)
                         + ".\nFound: " + ", ".join(map(str, raw.columns)))
    raw = raw.dropna(how="all")
    raw = raw[raw["Entry Date"].notna() & (raw["Entry Date"].astype(str).str.strip() != "")]
    raw = raw.reset_index(drop=True)

    df = pd.DataFrame()
    df["entry_date"] = pd.to_datetime(raw["Entry Date"], errors="coerce")
    df["exit_date"] = pd.to_datetime(raw["Exit Date"], errors="coerce")
    df["symbol"] = raw["Symbol"].astype(str).str.strip()
    df["ib_ticker"] = raw["IB_TICKER"].astype(str).str.strip()
    _inm = next((c for c in raw.columns if c.strip().lower() in
                 ("instrument name", "instrument_name")), None)
    df["instrument_name"] = (raw[_inm].astype(str).str.strip().replace({"nan": "", "None": ""})
                             if _inm else "")
    df["quantity"] = _to_num(raw["Quantity"])
    df["entry_price"] = _to_num(raw["Entry Price"])
    df["exit_price"] = _to_num(raw["Exit Price"])
    df["last_price"] = _to_num(raw["Last Price"]) if "Last Price" in raw.columns else np.nan
    _uc = next((c for c in raw.columns if c.strip().lower() in
                ("unrealized", "unrealised", "unrealized pnl", "unrealised pnl",
                 "unrealized p&l", "unrealised p&l")), None)
    df["unrealised"] = _to_num(raw[_uc]) if _uc else np.nan
    df["notional"] = _to_num(raw["Notional"])
    df["comm"] = _to_num(raw["Comm"]).fillna(0.0) if "Comm" in raw.columns else 0.0
    df["leg_pnl"] = _to_num(raw["Leg PnL"])
    df["trade_pnl"] = _to_num(raw["Trade PnL"]) if "Trade PnL" in raw.columns else np.nan
    df["trade_id"] = _to_num(raw["Trade ID"]).astype("Int64")
    df["trade_type"] = raw["Trade Type"].astype(str).str.strip()
    df["is_open"] = raw["Exit Price"].astype(str).str.strip().eq("OPEN")
    df["sec_type"] = df["symbol"].apply(lambda s: "FUT" if "FUTURE" in str(s).upper() else "STK")
    if "Asset Class" in raw.columns:
        ac = raw["Asset Class"].astype(str).str.strip().replace({"": np.nan, "nan": np.nan, "None": np.nan})
        df["asset_class"] = ac.fillna(df["sec_type"].map({"FUT": "Futures", "STK": "Equity"}))
    else:
        df["asset_class"] = df["sec_type"].map({"FUT": "Futures", "STK": "Equity"})
    df["multiplier"] = df.apply(_derive_multiplier, axis=1)
    df["side"] = np.where(df["quantity"] >= 0, "Long", "Short")
    return df.reset_index(drop=True)


def _derive_multiplier(row):
    if row["sec_type"] == "STK":
        return 1.0
    if not row["is_open"] and pd.notna(row["leg_pnl"]) and pd.notna(row["exit_price"]):
        move = (row["exit_price"] - row["entry_price"]) * row["quantity"]
        if abs(move) > 1e-9:
            return round(row["leg_pnl"] / move)
    m = FUT_MULTIPLIERS.get(_root_symbol(row["ib_ticker"]))
    if m:
        return float(m)
    denom = abs(row["quantity"]) * row["entry_price"]
    return round(row["notional"] / denom) if abs(denom) > 1e-9 else 1.0


def group_trades(legs):
    rows = []
    for tid, g in legs.groupby("trade_id"):
        closed, open_legs = g[~g["is_open"]], g[g["is_open"]]
        acs = g["asset_class"].unique()
        rows.append({
            "trade_id": tid, "trade_type": g["trade_type"].iloc[0],
            "asset_class": acs[0] if len(acs) == 1 else "Mixed",
            "instruments": ", ".join(g["symbol"].tolist()),
            "instrument_names": ", ".join(dict.fromkeys(
                n for n in g["instrument_name"].tolist() if n)),
            "n_legs": len(g),
            "entry_date": g["entry_date"].min(),
            "exit_date": (g["exit_date"].max() if open_legs.empty else pd.NaT),
            "status": "Open" if not open_legs.empty else "Closed",
            "realised_pnl": closed["trade_pnl"].sum(), "comm": g["comm"].sum(),
            "holding_days": ((g["exit_date"].max() - g["entry_date"].min()).days if open_legs.empty
                             else (pd.Timestamp.today().normalize() - g["entry_date"].min()).days),
        })
    out = pd.DataFrame(rows).sort_values("trade_id").reset_index(drop=True)
    out["winning"] = np.where(out["status"] == "Closed", out["realised_pnl"] > 0, np.nan)
    return out


def marks_excel(legs):
    rows = []
    for _, leg in legs[legs["is_open"]].iterrows():
        unreal = leg.unrealised
        has = pd.notna(unreal)
        last = leg.last_price
        px = last if pd.notna(last) else leg.entry_price
        rows.append({
            "symbol": leg.symbol, "side": leg.side, "asset_class": leg.asset_class,
            "quantity": leg.quantity, "entry_price": leg.entry_price,
            "last_price": (last if pd.notna(last) else np.nan), "has_mark": has,
            "unrealised_pnl": (float(unreal) if has else np.nan),
            "signed_notional": leg.quantity * px * leg.multiplier,
            "gross_notional": abs(leg.quantity) * px * leg.multiplier,
        })
    return pd.DataFrame(rows)


def totals_excel(trades, marks):
    realised = trades.loc[trades["status"] == "Closed", "realised_pnl"].sum()
    unreal = marks.loc[marks["has_mark"], "unrealised_pnl"].sum() if not marks.empty else 0.0
    comm = trades["comm"].sum()
    unpriced = int((~marks["has_mark"]).sum()) if not marks.empty else 0
    return {"realised": float(realised), "unrealised": float(unreal), "commissions": float(comm),
            "net_total": float(realised + unreal - comm), "gross_total": float(realised + unreal),
            "unpriced": unpriced, "n_open": (0 if marks.empty else len(marks))}


def realized_curve(trades):
    today = pd.Timestamp.today().normalize()
    start = trades["entry_date"].min().normalize()
    closed = trades[trades["status"] == "Closed"].copy()
    if closed.empty:
        cal = pd.bdate_range(start, today)
        return pd.DataFrame({"cum_gross": 0.0, "cum_comm": 0.0, "cum_net": 0.0, "daily_net": 0.0}, index=cal)
    d = closed["exit_date"].dt.normalize()
    gross_by, comm_by = closed.groupby(d)["realised_pnl"].sum(), closed.groupby(d)["comm"].sum()
    cal = pd.bdate_range(start, max(today, gross_by.index.max()))
    g, c = gross_by.reindex(cal).fillna(0.0), comm_by.reindex(cal).fillna(0.0)
    out = pd.DataFrame({"cum_gross": g.cumsum(), "cum_comm": c.cumsum()})
    out["cum_net"] = out["cum_gross"] - out["cum_comm"]
    out["daily_net"] = g - c
    return out


def daily_perf(daily):
    r, pnl = daily["ret"].dropna(), daily["pnl"].dropna()
    out = {"basis": "returns",
           "best_day": float(pnl.max()) if len(pnl) else np.nan,
           "worst_day": float(pnl.min()) if len(pnl) else np.nan,
           "total_ret": (float(daily["balance"].iloc[-1] / daily["balance"].iloc[0] - 1)
                         if daily["balance"].notna().sum() >= 2 else np.nan),
           "n_days": int(len(r))}
    if len(r) < 2 or r.std(ddof=1) == 0:
        out.update({"sharpe": np.nan, "sortino": np.nan, "ann_vol": np.nan,
                    "daily_vol": float(r.std(ddof=1)) if len(r) else np.nan})
        return out
    mean, std = r.mean(), r.std(ddof=1)
    dn = r[r < 0]
    dstd = np.sqrt((dn**2).mean()) if len(dn) else np.nan
    out.update({"sharpe": float(mean / std * np.sqrt(ANN)),
                "sortino": float(mean / dstd * np.sqrt(ANN)) if dstd and dstd > 0 else np.nan,
                "ann_vol": float(std * np.sqrt(ANN)), "daily_vol": float(std)})
    return out


def performance_ratios(daily_pnl):
    d = daily_pnl.dropna()
    if (d != 0).any():
        d = d[d.index >= d[d != 0].index.min()]
    if len(d) < 2 or d.std(ddof=1) == 0:
        return {"sharpe": np.nan, "sortino": np.nan, "ann_vol": np.nan,
                "daily_vol": float(d.std(ddof=1)) if len(d) else np.nan,
                "best_day": float(d.max()) if len(d) else np.nan,
                "worst_day": float(d.min()) if len(d) else np.nan}
    mean, std = d.mean(), d.std(ddof=1)
    dn = d[d < 0]
    dstd = np.sqrt((dn**2).mean()) if len(dn) else np.nan
    return {"sharpe": float(mean / std * np.sqrt(ANN)),
            "sortino": float(mean / dstd * np.sqrt(ANN)) if dstd and dstd > 0 else np.nan,
            "ann_vol": float(std * np.sqrt(ANN)), "daily_vol": float(std),
            "best_day": float(d.max()), "worst_day": float(d.min())}


def drawdown(equity):
    eq = equity.dropna()
    if eq.empty:
        return {"series": pd.Series(dtype=float), "max_dd": np.nan, "max_dd_pct": np.nan,
                "current_dd": np.nan, "max_dd_date": None}
    peak = eq.cummax()
    dd = eq - peak
    trough = dd.idxmin()
    pk = peak.loc[trough]
    return {"series": dd, "max_dd": float(dd.min()),
            "max_dd_pct": float(dd.min() / pk * 100) if pk > 0 else np.nan,
            "current_dd": float(dd.iloc[-1]), "max_dd_date": trough}


def value_at_risk(daily_pnl, levels=(0.95, 0.99)):
    d = daily_pnl.dropna(); d = d[d != 0]
    out = {}
    for lvl in levels:
        tag = int(lvl * 100)
        if len(d) < 5:
            out[f"var_{tag}"] = out[f"cvar_{tag}"] = np.nan; continue
        q = np.quantile(d, 1 - lvl); tail = d[d <= q]
        out[f"var_{tag}"] = float(-q)
        out[f"cvar_{tag}"] = float(-tail.mean()) if len(tail) else np.nan
    return out


def exposure(marks):
    if marks is None or marks.empty:
        return {"gross": 0.0, "net": 0.0, "long": 0.0, "short": 0.0, "hhi": np.nan,
                "top_name": None, "top_share": np.nan, "by_instrument": pd.DataFrame()}
    gross = marks["gross_notional"].sum()
    w = (marks["gross_notional"] / gross) if gross else marks["gross_notional"] * 0
    top = marks["gross_notional"].idxmax()
    by = (marks.assign(weight=w)[["symbol", "side", "signed_notional", "gross_notional", "weight"]]
          .sort_values("gross_notional", ascending=False).reset_index(drop=True))
    return {"gross": float(gross), "net": float(marks["signed_notional"].sum()),
            "long": float(marks.loc[marks["signed_notional"] > 0, "signed_notional"].sum()),
            "short": float(marks.loc[marks["signed_notional"] < 0, "signed_notional"].sum()),
            "hhi": float((w**2).sum()), "top_name": marks.loc[top, "symbol"],
            "top_share": float(w.loc[top]), "by_instrument": by}


def trade_stats(trades):
    closed = trades[trades["status"] == "Closed"].copy()
    if closed.empty:
        return {"n_closed": 0}
    closed["net_pnl"] = closed["realised_pnl"] - closed["comm"]
    wins, losses = closed[closed["net_pnl"] > 0]["net_pnl"], closed[closed["net_pnl"] <= 0]["net_pnl"]
    gl = losses.sum()
    return {"n_closed": int(len(closed)), "win_rate": float(len(wins) / len(closed)),
            "profit_factor": float(wins.sum() / abs(gl)) if gl != 0 else np.inf,
            "avg_win": float(wins.mean()) if len(wins) else 0.0,
            "avg_loss": float(losses.mean()) if len(losses) else 0.0,
            "payoff": float(wins.mean() / abs(losses.mean())) if len(losses) and losses.mean() != 0 else np.inf,
            "expectancy": float(closed["net_pnl"].mean()), "avg_hold": float(closed["holding_days"].mean())}


def by_strategy(trades):
    closed = trades[trades["status"] == "Closed"].copy()
    if closed.empty:
        return pd.DataFrame()
    closed["net_pnl"] = closed["realised_pnl"] - closed["comm"]
    g = closed.groupby("trade_type")
    return pd.DataFrame({"trade_type": list(g.groups.keys()),
                         "net_pnl": g["net_pnl"].sum().values,
                         "win_rate": g.apply(lambda x: (x["net_pnl"] > 0).mean()).values})


def by_asset_class(trades, marks):
    rec = {}
    closed = trades[trades["status"] == "Closed"].copy()
    if not closed.empty:
        closed["net"] = closed["realised_pnl"] - closed["comm"]
        for ac, g in closed.groupby("asset_class"):
            rec.setdefault(ac, {}).update({"trades": int(len(g)), "realised": float(g["net"].sum()),
                                           "wins": int((g["net"] > 0).sum())})
    if marks is not None and not marks.empty:
        for ac, g in marks.groupby("asset_class"):
            d = rec.setdefault(ac, {})
            d["open_legs"] = int(len(g))
            d["unrealised"] = float(g.loc[g["has_mark"], "unrealised_pnl"].sum())
    out = []
    for ac, d in rec.items():
        realised, unreal = d.get("realised", 0.0), d.get("unrealised", 0.0)
        n = d.get("trades", 0)
        out.append({"asset_class": ac, "trades": n, "open_legs": d.get("open_legs", 0),
                    "realised": realised, "unrealised": unreal, "net": realised + unreal,
                    "win_rate": (d.get("wins", 0) / n) if n else np.nan})
    return pd.DataFrame(out).sort_values("net", ascending=False).reset_index(drop=True) if out else pd.DataFrame()


# ===========================================================================
#  FORMAT + FIGURES
# ===========================================================================
def money(x, dp=0):
    return "—" if x is None or (isinstance(x, float) and np.isnan(x)) else f"${x:,.{dp}f}"


def money_signed(x, dp=0):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "—"
    return f"{'-' if x < 0 else '+'}${abs(x):,.{dp}f}"


def num(x, dp=2):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "—"
    return "∞" if np.isinf(x) else f"{x:,.{dp}f}"


def pct(x, dp=1):
    return "—" if x is None or (isinstance(x, float) and np.isnan(x)) else f"{x*100:,.{dp}f}%"


def pnl_color(x):
    return MUTED if x is None or (isinstance(x, float) and np.isnan(x)) else (UP if x >= 0 else DOWN)


def fig_balance(daily, dd):
    bal = daily["balance"]
    fig = go.Figure(go.Scatter(x=bal.index, y=bal.values, line=dict(color=BRASS, width=2.2)))
    if dd["max_dd_date"] is not None and dd["max_dd_date"] in bal.index:
        m = dd["max_dd_date"]
        fig.add_trace(go.Scatter(x=[m], y=[bal.loc[m]], mode="markers", showlegend=False,
                      marker=dict(color=DOWN, size=8, symbol="triangle-down")))
    fig.update_layout(**PLOT, height=320); fig.update_yaxes(tickprefix="$", tickformat=",.0f")
    return fig


def fig_equity_excel(curves, dd, total_today):
    net = curves["cum_net"]
    fig = go.Figure(go.Scatter(x=curves.index, y=net, line=dict(color=BRASS, width=2.2),
                    fill="tozeroy", fillcolor="rgba(200,162,90,0.08)"))
    if total_today is not None and len(net):
        d = curves.index[-1]
        fig.add_trace(go.Scatter(x=[d, d], y=[net.iloc[-1], total_today], mode="lines",
                      line=dict(color=BLUE, width=1.4, dash="dot"), showlegend=False))
        fig.add_trace(go.Scatter(x=[d], y=[total_today], mode="markers",
                      marker=dict(color=BLUE, size=9), showlegend=False))
    fig.update_layout(**PLOT, height=320); fig.update_yaxes(tickprefix="$", tickformat=",.0f")
    return fig


def fig_drawdown(dd):
    s = dd["series"]
    fig = go.Figure(go.Scatter(x=s.index, y=s.values, line=dict(color=DOWN, width=1),
                    fill="tozeroy", fillcolor="rgba(229,96,110,0.18)"))
    fig.update_layout(**PLOT, height=160); fig.update_yaxes(tickprefix="$", tickformat=",.0f")
    return fig


def fig_exposure_bars(comp, index=None):
    if comp is None or comp.empty:
        return go.Figure(layout=dict(**PLOT, height=210))
    e = comp if index is None else comp.reindex(pd.DatetimeIndex(index).normalize())
    e = e.dropna(how="all")
    if e.empty:
        return go.Figure(layout=dict(**PLOT, height=210))
    order = e.mean().sort_values(ascending=False).index.tolist()  # biggest share stacks first
    fig = go.Figure()
    for i, col in enumerate(order):
        fig.add_trace(go.Bar(x=e.index, y=e[col], name=col,
                      marker_color=_asset_color(col, i), marker_line_width=0,
                      hovertemplate="%{x|%Y-%m-%d}<br>" + str(col) + " %{y:.0%}<extra></extra>"))
    fig.update_layout(**PLOT, height=210, barmode="stack")
    fig.update_layout(showlegend=True,
                      legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0,
                                  font=dict(family=MONO, size=10, color=MUTED)))
    fig.update_yaxes(tickformat=".0%", range=[0, 1])
    return fig


def fig_daily(daily_net):
    fig = go.Figure(go.Bar(x=daily_net.index, y=daily_net.values,
                    marker_color=[UP if v >= 0 else DOWN for v in daily_net.values], marker_line_width=0))
    fig.update_layout(**PLOT, height=230); fig.update_yaxes(tickprefix="$", tickformat=",.0f")
    return fig


def fig_strategy(bystrat):
    if bystrat.empty:
        return go.Figure(layout=dict(**PLOT, height=230))
    fig = go.Figure(go.Bar(x=bystrat["trade_type"], y=bystrat["net_pnl"], width=0.5,
                    marker_color=[UP if v >= 0 else DOWN for v in bystrat["net_pnl"]], marker_line_width=0,
                    text=[money(v) for v in bystrat["net_pnl"]], textposition="outside",
                    textfont=dict(family=MONO, size=11, color=TXT)))
    fig.update_layout(**PLOT, height=260); fig.update_yaxes(tickprefix="$", tickformat=",.0f")
    return fig


def fig_asset(byac):
    if byac.empty:
        return go.Figure(layout=dict(**PLOT, height=260))
    df = byac.iloc[::-1]
    fig = go.Figure(go.Bar(y=df["asset_class"], x=df["net"], orientation="h",
                    marker_color=[UP if v >= 0 else DOWN for v in df["net"]], marker_line_width=0,
                    text=[money(v) for v in df["net"]], textposition="auto",
                    textfont=dict(family=MONO, size=11, color=INK)))
    fig.update_layout(**PLOT, height=max(260, 34 * len(df) + 60)); fig.update_xaxes(tickprefix="$", tickformat="~s")
    return fig


def fig_exposure(exp):
    df = exp["by_instrument"]
    if df.empty:
        return go.Figure(layout=dict(**PLOT, height=260))
    df = df.iloc[::-1]
    fig = go.Figure(go.Bar(y=df["symbol"], x=df["gross_notional"], orientation="h",
                    marker_color=[UP if s == "Long" else DOWN for s in df["side"]], marker_line_width=0,
                    text=[money(v) for v in df["gross_notional"]], textposition="auto",
                    textfont=dict(family=MONO, size=10, color=INK)))
    fig.update_layout(**PLOT, height=max(260, 26 * len(df) + 40)); fig.update_xaxes(tickprefix="$", tickformat="~s")
    return fig


# ===========================================================================
#  STREAMLIT UI
# ===========================================================================
st.set_page_config(page_title="Book Monitor", page_icon="📈", layout="wide")

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500;600&display=swap');
.stApp {{ background:{INK}; }}
.block-container {{ padding-top: 3.2rem; max-width: 1400px; }}
.kpi {{ background:{PANEL}; border:1px solid {LINE}; border-radius:10px; padding:14px 16px; height:100%; }}
.kpi .l {{ color:{MUTED}; font:600 10px {MONO}; letter-spacing:.08em; }}
.kpi .v {{ font:600 22px {MONO}; margin-top:6px; line-height:1.1; }}
.kpi .s {{ color:{MUTED}; font:400 11px {MONO}; margin-top:4px; min-height:14px; }}
.sect {{ color:{TXT}; font:600 13px {FONT}; margin:6px 0 2px 2px; }}
.tile {{ background:{PANEL2}; border:1px solid {LINE}; border-radius:8px; padding:10px 12px; }}
.tile .l {{ color:{MUTED}; font:600 9.5px {MONO}; letter-spacing:.07em; }}
.tile .v {{ font:600 17px {MONO}; margin-top:3px; }}
</style>
""", unsafe_allow_html=True)


def kpi(col, label, value, sub="", color=TXT):
    col.markdown(f'<div class="kpi"><div class="l">{label.upper()}</div>'
                 f'<div class="v" style="color:{color}">{value}</div>'
                 f'<div class="s">{sub}</div></div>', unsafe_allow_html=True)


def tile(col, label, value, color=TXT):
    col.markdown(f'<div class="tile"><div class="l">{label.upper()}</div>'
                 f'<div class="v" style="color:{color}">{value}</div></div>', unsafe_allow_html=True)


def section(text):
    st.markdown(f'<div class="sect">{text}</div>', unsafe_allow_html=True)


def _pnl_style(col):
    def f(v):
        v = str(v)
        if v.startswith("-"):
            return f"color:{DOWN}"
        if v.startswith("+"):
            return f"color:{UP}"
        return ""
    return f


# ---- header ----
st.markdown(f'<div style="font:600 20px {MONO};letter-spacing:.12em">'
            f'<span style="color:{BRASS}">BOOK</span>'
            f'<span style="color:{TXT}">MONITOR</span></div>'
            f'<div style="color:{MUTED};font:400 12px {FONT}">track record · trade stats · exposure</div>',
            unsafe_allow_html=True)

# ---- data source (bundled workbook only; no uploads) ----
src = TRADES_FILE if os.path.exists(TRADES_FILE) else None

if src is None:
    st.error(f"Workbook **{TRADES_FILE}** not found in the app folder.")
    st.stop()

try:
    legs = load_trades(src)
except Exception as e:
    st.error(f"Could not load workbook:\n\n{e}")
    st.stop()

trades = group_trades(legs)
marks = marks_excel(legs)
totals = totals_excel(trades, marks)
daily = load_daily_pnl(src)
expo = load_exposure(src)
use_daily = daily is not None and len(daily) >= 2

if use_daily:
    perf = daily_perf(daily)
    dd = drawdown(daily["balance"])
    var = value_at_risk(daily["pnl"])
    daily_net = daily["pnl"]
else:
    curves = realized_curve(trades)
    perf = performance_ratios(curves["daily_net"]); perf["basis"] = "realised"
    dd = drawdown(curves["cum_net"])
    var = value_at_risk(curves["daily_net"])
    daily_net = curves["daily_net"]
    total_today = float(curves["cum_net"].iloc[-1]) + totals["unrealised"]

exp = exposure(marks)
ts = trade_stats(trades)
bystrat = by_strategy(trades)
byac = by_asset_class(trades, marks)

st.caption(f"● {'DAILY-RETURN basis' if use_daily else 'REALISED basis'} · updated {dt.datetime.now():%Y-%m-%d %H:%M}")

if totals["unpriced"]:
    st.warning(f"{totals['unpriced']} of {totals['n_open']} open leg(s) have no value in the "
               "Unrealized column — excluded from unrealised P&L.")

# ---- KPIs ----
c = st.columns(4)
kpi(c[0], "Net P&L", money(totals["net_total"]), f"gross {money(totals['gross_total'])}", pnl_color(totals["net_total"]))
kpi(c[1], "Realised", money(totals["realised"]), "closed trades", pnl_color(totals["realised"]))
kpi(c[2], "Unrealised", money(totals["unrealised"]), "open legs (Unrealized col)", pnl_color(totals["unrealised"]))
kpi(c[3], "Sharpe (ann)", num(perf.get("sharpe")),
    f"Sortino {num(perf.get('sortino'))}" + (" · daily rets" if use_daily else " · realised"),
    BLUE if use_daily else TXT)
c = st.columns(4)
kpi(c[0], "Max Drawdown", money(dd["max_dd"]),
    pct(dd["max_dd_pct"] / 100) if not np.isnan(dd["max_dd_pct"]) else "", DOWN)
kpi(c[1], "Win Rate", pct(ts.get("win_rate")) if ts.get("n_closed") else "—",
    f"PF {num(ts.get('profit_factor'))}" if ts.get("n_closed") else "")
kpi(c[2], "Gross Exposure", money(exp["gross"]), f"net {money(exp['net'])}")
if use_daily:
    kpi(c[3], "Account Balance", money(daily["balance"].iloc[-1]),
        f"return {pct(perf['total_ret'])} · {perf['n_days']}d", BLUE)
else:
    kpi(c[3], "Open Trades", str(totals["n_open"]), f"{ts.get('n_closed', 0)} closed", BLUE)

st.write("")

# ---- equity + daily ----
left, right = st.columns([2.1, 1])
with left:
    section("Account balance · daily MTM" if use_daily else "Realised P&L · by exit date")
    st.plotly_chart(fig_balance(daily, dd) if use_daily else fig_equity_excel(curves, dd, total_today),
                    use_container_width=True, config={"displayModeBar": False})
    if expo is not None and not expo.empty:
        section("Daily exposure · % of book by asset class")
        _idx = daily["balance"].index if use_daily else curves.index
        st.plotly_chart(fig_exposure_bars(expo, _idx),
                        use_container_width=True, config={"displayModeBar": False})
    section("Drawdown")
    st.plotly_chart(fig_drawdown(dd), use_container_width=True, config={"displayModeBar": False})
with right:
    section("Daily P&L")
    st.plotly_chart(fig_daily(daily_net), use_container_width=True, config={"displayModeBar": False})
    t = st.columns(3)
    tile(t[0], "Best day", money(perf["best_day"]), UP)
    tile(t[1], "Worst day", money(perf["worst_day"]), DOWN)
    tile(t[2], "Daily σ", pct(perf["daily_vol"]) if use_daily else money(perf["daily_vol"]))

st.write("")

# ---- risk grid ----
section(f"Risk & performance · {'daily returns' if use_daily else 'realised'} basis")
g = st.columns(9)
tile(g[0], "Sharpe", num(perf["sharpe"]))
tile(g[1], "Sortino", num(perf["sortino"]))
tile(g[2], "Ann. vol", pct(perf["ann_vol"]) if use_daily else money(perf["ann_vol"]))
tile(g[3], "VaR 95%", money(var["var_95"]), DOWN)
tile(g[4], "CVaR 95%", money(var["cvar_95"]), DOWN)
tile(g[5], "VaR 99%", money(var["var_99"]), DOWN)
tile(g[6], "CVaR 99%", money(var["cvar_99"]), DOWN)
tile(g[7], "Max DD", money(dd["max_dd"]), DOWN)
tile(g[8], "Current DD", money(dd["current_dd"]), DOWN if dd["current_dd"] < 0 else TXT)

st.write("")

# ---- trade stats + strategy ----
left, right = st.columns([1.3, 1])
with left:
    section("Trade statistics · closed, net of commissions")
    if ts.get("n_closed"):
        r1 = st.columns(4)
        tile(r1[0], "Closed trades", str(ts["n_closed"]))
        tile(r1[1], "Win rate", pct(ts["win_rate"]))
        tile(r1[2], "Profit factor", num(ts["profit_factor"]), UP)
        tile(r1[3], "Payoff", num(ts["payoff"]))
        r2 = st.columns(4)
        tile(r2[0], "Expectancy", money(ts["expectancy"]), pnl_color(ts["expectancy"]))
        tile(r2[1], "Avg hold", f"{ts['avg_hold']:.1f}d")
    else:
        st.caption("No closed trades yet.")
with right:
    section("By strategy")
    st.plotly_chart(fig_strategy(bystrat), use_container_width=True, config={"displayModeBar": False})

st.write("")

# ---- asset class ----
left, right = st.columns([1.1, 1.4])
with left:
    section("Net P&L by asset class · realised + open unrealised")
    st.plotly_chart(fig_asset(byac), use_container_width=True, config={"displayModeBar": False})
with right:
    section("Asset-class breakdown")
    if byac.empty:
        st.caption("No asset-class data.")
    else:
        show = pd.DataFrame({
            "Asset Class": byac["asset_class"], "Trades": byac["trades"].astype(int),
            "Open": byac["open_legs"].astype(int),
            "Realised": byac["realised"].map(money_signed), "Unrealised": byac["unrealised"].map(money_signed),
            "Net": byac["net"].map(money_signed),
            "Win%": byac["win_rate"].map(lambda v: pct(v) if pd.notna(v) else "—")})
        sty = show.style.map(_pnl_style("Net"), subset=["Realised", "Unrealised", "Net"])
        st.dataframe(sty, use_container_width=True, hide_index=True)

st.write("")

# ---- exposure ----
left, right = st.columns([1.4, 1])
with left:
    section("Open exposure by instrument")
    st.plotly_chart(fig_exposure(exp), use_container_width=True, config={"displayModeBar": False})
with right:
    section("Concentration")
    e = st.columns(2)
    tile(e[0], "Gross", money(exp["gross"]))
    tile(e[1], "Net", money(exp["net"]), pnl_color(exp["net"]))
    e = st.columns(2)
    tile(e[0], "Long", money(exp["long"]), UP)
    tile(e[1], "Short", money(exp["short"]), DOWN)
    e = st.columns(2)
    tile(e[0], "Top position", f"{exp['top_name']}" if exp["top_name"] else "—")
    tile(e[1], "Top share", pct(exp["top_share"]))
    e = st.columns(2)
    tile(e[0], "HHI", num(exp["hhi"], 3))
    tile(e[1], "Eff. names", num(1 / exp["hhi"], 1) if exp["hhi"] else "—")

st.write("")

# ---- tables ----
left, right = st.columns([1.4, 1])
with left:
    section("Trades")
    tdf = trades.copy()
    tdf["net"] = tdf["realised_pnl"] - tdf["comm"]
    tshow = pd.DataFrame({
        "ID": tdf["trade_id"].astype("Int64").astype(str),
        "Type": tdf["trade_type"], "Instruments": tdf["instruments"],
        "Instrument Name": tdf["instrument_names"], "Status": tdf["status"],
        "Hold": tdf["holding_days"].map(lambda v: f"{int(v)}d"),
        "Net P&L": tdf.apply(lambda r: "OPEN" if r["status"] == "Open" else money_signed(r["net"]), axis=1)})
    st.dataframe(tshow.style.map(_pnl_style("Net P&L"), subset=["Net P&L"]),
                 use_container_width=True, hide_index=True)
with right:
    section("Open positions · Unrealized from sheet")
    if marks.empty:
        st.caption("No open positions.")
    else:
        oshow = pd.DataFrame({
            "Instrument": marks["symbol"], "Side": marks["side"],
            "Qty": marks["quantity"].map(lambda v: f"{int(v):,}"),
            "Entry": marks["entry_price"].map(lambda v: num(v, 4)),
            "Last": marks["last_price"].map(lambda v: num(v, 4)),
            "Unreal P&L": marks["unrealised_pnl"].map(money_signed),
            "Gross Notl": marks["gross_notional"].map(lambda v: money(v))})
        st.dataframe(oshow.style.map(_pnl_style("Unreal P&L"), subset=["Unreal P&L"]),
                     use_container_width=True, hide_index=True)