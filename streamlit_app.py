# ============================================================================
#  LONG HORIZON MACRO STRATEGY · DAILY REVIEW  (Streamlit)
#  Excel/CSV-driven book monitor. Reads a workbook ('Trades' + 'Daily PnL' +
#  'exposure' sheets) and a daily-returns CSV, and reports the track record,
#  risk, exposure and monthly returns — all on the FT cream theme, with numbers
#  expressed as % of starting capital. No Interactive Brokers, no local machine.
#
#  Deploy on Streamlit Community Cloud: push this folder to GitHub (with the
#  workbook and the CSV committed next to this file), then share.streamlit.io.
# ============================================================================

import os
import datetime as dt
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ---------------------------------------------------------------------------
#  CONFIG  — relative paths so they resolve both locally and on the cloud.
# ---------------------------------------------------------------------------
TRADES_FILE    = os.environ.get("TRADES_FILE", "track_record.xlsx")
TRADES_SHEET   = "Trades"
DAILY_SHEET    = "Daily PnL"
EXPOSURE_SHEET = "exposure"
# Daily-returns CSV (Date + Daily_Return) feeding the Monthly Returns table.
MONTHLY_RETURNS_FILE = os.environ.get("MONTHLY_RETURNS_FILE",
                                      "str_II_dailyreturns_fullhistory.csv")

_MONTH_CODES = {"F": 1, "G": 2, "H": 3, "J": 4, "K": 5, "M": 6,
                "N": 7, "Q": 8, "U": 9, "V": 10, "X": 11, "Z": 12}

# Authoritative contract multipliers by root symbol (OPEN futures legs).
FUT_MULTIPLIERS = {
    "VXM": 100, "VX": 1000, "SI": 5000, "QI": 2500, "SIL": 1000, "GC": 100, "QO": 50,
    "MGC": 10, "HG": 25000, "ZC": 5000, "ZS": 5000, "ZW": 5000, "ZL": 60000, "ZM": 100,
    "CC": 10, "KC": 37500, "SB": 112000, "EUR": 125000, "6E": 125000, "SF": 125000,
    "CHF": 125000, "6S": 125000, "6B": 62500, "6J": 12500000, "6A": 100000, "6C": 100000,
    "ES": 50, "NQ": 20, "YM": 5, "RTY": 50, "CL": 1000, "NG": 10000, "RB": 42000, "HO": 42000,
}

_REQUIRED_COLS = ["Entry Date", "Exit Date", "Symbol", "IB_TICKER", "Quantity",
                  "Entry Price", "Exit Price", "Notional", "Leg PnL", "Trade ID", "Trade Type"]


def _root_symbol(local_symbol):
    s = str(local_symbol)
    if len(s) >= 2 and s[-1].isdigit() and s[-2].upper() in _MONTH_CODES:
        return s[:-2]
    return s


# ===========================================================================
#  LOADERS
# ===========================================================================
def _to_num(series):
    return pd.to_numeric(series.astype(str).str.strip().str.replace(",", "", regex=False)
                         .replace({"OPEN": np.nan, "": np.nan}), errors="coerce")


def _read_blotter():
    if not TRADES_FILE:
        raise FileNotFoundError("TRADES_FILE is not set.")
    if not os.path.exists(TRADES_FILE):
        raise FileNotFoundError(
            f"Blotter file not found:\n    {TRADES_FILE}\n"
            "Check the path (and that it's committed next to streamlit_app.py for the cloud).")
    if str(TRADES_FILE).lower().endswith((".xlsx", ".xlsm", ".xls")):
        xls = pd.ExcelFile(TRADES_FILE)
        sheet = next((s for s in xls.sheet_names
                      if s.strip().lower() == str(TRADES_SHEET).strip().lower()), xls.sheet_names[0])
        probe = pd.read_excel(xls, sheet_name=sheet, header=None, dtype=str)
        header_row = 0
        for i in range(min(len(probe), 30)):
            if "Entry Date" in [str(v).strip() for v in probe.iloc[i].tolist()]:
                header_row = i
                break
        return pd.read_excel(xls, sheet_name=sheet, header=header_row, dtype=str)
    return pd.read_csv(TRADES_FILE, dtype=str)


def load_daily_pnl():
    """'Daily PnL' sheet -> DataFrame indexed by date with balance, pnl, ret."""
    if not TRADES_FILE or not os.path.exists(TRADES_FILE):
        return None
    if not str(TRADES_FILE).lower().endswith((".xlsx", ".xlsm", ".xls")):
        return None
    try:
        xls = pd.ExcelFile(TRADES_FILE)
    except Exception:
        return None
    name = next((s for s in xls.sheet_names
                 if s.strip().lower() == str(DAILY_SHEET).strip().lower()), None)
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
    bal_col = next((c for c in raw.columns if "balance" in c.lower()), None)
    pnl_col = next((c for c in raw.columns
                    if c.lower().strip() in ("pnl", "p&l", "daily pnl", "profit", "return")), None)
    d = pd.DataFrame()
    d["date"] = pd.to_datetime(raw["Date"], errors="coerce")
    d["balance"] = _to_num(raw[bal_col]) if bal_col else np.nan
    d["pnl"] = _to_num(raw[pnl_col]) if pnl_col else np.nan
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


def load_exposure():
    """'exposure' sheet -> daily fractions per asset class (rows normalised to 100%)."""
    if not TRADES_FILE or not os.path.exists(TRADES_FILE):
        return None
    if not str(TRADES_FILE).lower().endswith((".xlsx", ".xlsm", ".xls")):
        return None
    try:
        xls = pd.ExcelFile(TRADES_FILE)
    except Exception:
        return None
    name = next((s for s in xls.sheet_names
                 if s.strip().lower() == str(EXPOSURE_SHEET).strip().lower()), None)
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

    wanted = {
        "Equities":   ("equity_exposure", "equities_exposure"),
        "Cash":       ("cash_exposure",),
        "Currency":   ("currency_exposure",),
        "Metals":     ("metals_exposure", "metal_exposure"),
        "Softs":      ("softs_exposure", "soft_exposure"),
        "Volatility": ("volatility_exposure", "vol_exposure"),
        "Index":      ("index_exposure",),
        "Rates":        ("rates_exposure", "rate_exposure"),
        "Energy":       ("energy_exposure",),
        "Agricultural": ("agricultural_exposure", "agriculture_exposure", "ag_exposure"),
        "ETFs":         ("etf_exposure", "etfs_exposure"),
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
    frac = d.fillna(0.0)
    tot = frac.sum(axis=1).replace(0, np.nan)
    norm = frac.div(tot, axis=0).fillna(0.0)
    return norm.loc[:, (norm != 0).any(axis=0)]


def load_trades():
    raw = _read_blotter()
    raw.columns = [str(c).strip() for c in raw.columns]
    missing = [c for c in _REQUIRED_COLS if c not in raw.columns]
    if missing:
        raise ValueError("Blotter is missing expected column(s): " + ", ".join(missing)
                         + ".\nFound: " + ", ".join(map(str, raw.columns)))
    raw = raw.dropna(how="all")
    raw = raw[raw["Entry Date"].notna() & (raw["Entry Date"].astype(str).str.strip() != "")]
    raw = raw.reset_index(drop=True)

    df = pd.DataFrame()
    df["entry_date"] = pd.to_datetime(raw["Entry Date"], errors="coerce")
    df["exit_date"] = pd.to_datetime(raw["Exit Date"], errors="coerce")
    df["symbol"] = raw["Symbol"].astype(str).str.strip()
    df["ib_ticker"] = raw["IB_TICKER"].astype(str).str.strip()
    _thm = next((c for c in raw.columns if c.strip().lower() == "theme"), None)
    df["theme"] = (raw[_thm].astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})
                   if _thm else "")
    df["quantity"] = _to_num(raw["Quantity"])
    df["entry_price"] = _to_num(raw["Entry Price"])
    df["exit_price"] = _to_num(raw["Exit Price"])
    df["last_price"] = _to_num(raw["Last Price"]) if "Last Price" in raw.columns else np.nan
    _ucol = next((c for c in raw.columns if c.strip().lower() in
                  ("unrealized", "unrealised", "unrealized pnl", "unrealised pnl",
                   "unrealized p&l", "unrealised p&l")), None)
    df["unrealised"] = _to_num(raw[_ucol]) if _ucol else np.nan
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
            "themes": ", ".join(dict.fromkeys(t for t in g["theme"].tolist() if t)),
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


# ===========================================================================
#  EXCEL-DRIVEN MARKS / TOTALS
# ===========================================================================
def marks_excel(legs):
    rows = []
    for _, leg in legs[legs["is_open"]].iterrows():
        unreal = leg.unrealised
        has = pd.notna(unreal)
        last = leg.last_price
        px = last if pd.notna(last) else leg.entry_price
        rows.append({
            "symbol": leg.symbol, "ib_ticker": leg.ib_ticker, "side": leg.side,
            "trade_id": leg.trade_id, "trade_type": leg.trade_type, "asset_class": leg.asset_class,
            "quantity": leg.quantity, "entry_price": leg.entry_price,
            "last_price": (last if pd.notna(last) else np.nan),
            "has_mark": has, "multiplier": leg.multiplier,
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


# ===========================================================================
#  RISK
# ===========================================================================
ANN = 252


def daily_perf(daily):
    r, pnl = daily["ret"].dropna(), daily["pnl"].dropna()
    out = {"basis": "returns",
           "best_day": float(pnl.max()) if len(pnl) else np.nan,
           "worst_day": float(pnl.min()) if len(pnl) else np.nan,
           "cum_pnl": float(pnl.sum()) if len(pnl) else np.nan,
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
    peak_at = peak.loc[trough]
    return {"series": dd, "max_dd": float(dd.min()),
            "max_dd_pct": float(dd.min() / peak_at * 100) if peak_at > 0 else np.nan,
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
    weights = (marks["gross_notional"] / gross) if gross else marks["gross_notional"] * 0
    top = marks["gross_notional"].idxmax()
    by_instrument = (marks.assign(weight=weights)
                     [["symbol", "side", "signed_notional", "gross_notional", "weight"]]
                     .sort_values("gross_notional", ascending=False).reset_index(drop=True))
    return {"gross": float(gross), "net": float(marks["signed_notional"].sum()),
            "long": float(marks.loc[marks["signed_notional"] > 0, "signed_notional"].sum()),
            "short": float(marks.loc[marks["signed_notional"] < 0, "signed_notional"].sum()),
            "hhi": float((weights**2).sum()), "top_name": marks.loc[top, "symbol"],
            "top_share": float(weights.loc[top]), "by_instrument": by_instrument}


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


def by_asset_class(trades, marks, legs=None):
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
    ls = {}
    if legs is not None and not legs.empty:
        L = legs.copy()
        closed_pl = (L["leg_pnl"].where(~L["is_open"], 0.0).fillna(0.0)
                     - L["comm"].where(~L["is_open"], 0.0).fillna(0.0))
        open_pl = L["unrealised"].where(L["is_open"], 0.0).fillna(0.0)
        L["contrib"] = closed_pl + open_pl
        for (ac, side), g in L.groupby(["asset_class", "side"]):
            ls.setdefault(ac, {})[side] = ls.get(ac, {}).get(side, 0.0) + float(g["contrib"].sum())
    out = []
    for ac, d in rec.items():
        realised, unreal = d.get("realised", 0.0), d.get("unrealised", 0.0)
        n = d.get("trades", 0)
        out.append({"asset_class": ac, "trades": n, "open_legs": d.get("open_legs", 0),
                    "realised": realised, "unrealised": unreal, "net": realised + unreal,
                    "long": ls.get(ac, {}).get("Long", 0.0), "short": ls.get(ac, {}).get("Short", 0.0),
                    "win_rate": (d.get("wins", 0) / n) if n else np.nan})
    return pd.DataFrame(out).sort_values("net", ascending=False).reset_index(drop=True) if out else pd.DataFrame()


# ===========================================================================
#  THEME + FORMAT  (FT cream palette)
# ===========================================================================
FT_COLORS = {"bg": "#e5d5c5", "text": "#2C2C2C", "grid": "#E0E0E0",
             "win": "#2E7D32", "loss": "#C41E3A", "neutral": "#FF8C42", "accent": "#1565C0"}
INK   = FT_COLORS["bg"]
PANEL = "#EFE4D7"
PANEL2 = "#DBC9B6"
LINE  = FT_COLORS["grid"]
TXT   = FT_COLORS["text"]
MUTED = "#6E6056"
BRASS = FT_COLORS["accent"]
UP, DOWN, BLUE = FT_COLORS["win"], FT_COLORS["loss"], FT_COLORS["accent"]
BAR_UP, BAR_DOWN = "#13420F", "#5E0A14"
ON_FILL = "#FFFFFF"
GREEK_BLUE = "#1868B7"
MONTHLY_POS, MONTHLY_NEG = "#0D3B66", "#8E2323"
ASSET_COLORS = {
    "Equities": GREEK_BLUE, "Cash": "#B5A694", "Currency": "#5E9C9A", "Metals": "#C9A227",
    "Softs": "#7A9E5E", "Volatility": "#8E5A9E", "Index": "#2C3E66", "Rates": "#4F7C8A",
    "Energy": "#C4622D", "Agricultural": "#6B8E3D", "ETFs": "#9C6B4E",
}
_EXPO_PALETTE = ["#1868B7", "#C9A227", "#9C6B4E", "#5E9C9A", "#8E5A9E",
                 "#2C3E66", "#7A9E5E", "#C4622D", "#4F7C8A", "#B5A694"]
FONT = "IBM Plex Sans, system-ui, sans-serif"
MONO = "IBM Plex Mono, ui-monospace, monospace"
PLOT = dict(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            font=dict(family=MONO, color=MUTED, size=11), margin=dict(l=48, r=18, t=28, b=32),
            xaxis=dict(gridcolor=LINE, zerolinecolor=LINE, linecolor=LINE),
            yaxis=dict(gridcolor=LINE, zerolinecolor=LINE, linecolor=LINE),
            hoverlabel=dict(font=dict(family=MONO, color=TXT, size=11),
                            bgcolor="#FFFFFF", bordercolor=LINE), showlegend=False)


def _asset_color(name, i):
    return ASSET_COLORS.get(name, _EXPO_PALETTE[i % len(_EXPO_PALETTE)])


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


def pct_signed(x, dp=2):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "—"
    return f"{'-' if x < 0 else '+'}{abs(x)*100:,.{dp}f}%"


def pnl_color(x):
    return MUTED if x is None or (isinstance(x, float) and np.isnan(x)) else (UP if x >= 0 else DOWN)


# ===========================================================================
#  FIGURES
# ===========================================================================
def fig_equity_excel(curves, dd, total_today=None):
    net = curves["cum_net"]
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=curves.index, y=net, name="Realised P&L",
                  line=dict(color=BRASS, width=2.2), fill="tozeroy",
                  fillcolor="rgba(21,101,192,0.10)"))
    if total_today is not None and len(net):
        d = curves.index[-1]
        fig.add_trace(go.Scatter(x=[d, d], y=[net.iloc[-1], total_today], mode="lines",
                      line=dict(color=BLUE, width=1.4, dash="dot"), showlegend=False))
        fig.add_trace(go.Scatter(x=[d], y=[total_today], mode="markers+text",
                      marker=dict(color=BLUE, size=9), text=[" incl. unreal " + money(total_today)],
                      textposition="middle left", textfont=dict(color=BLUE, size=10, family=MONO),
                      showlegend=False))
    if dd["max_dd_date"] is not None and dd["max_dd_date"] in net.index:
        m = dd["max_dd_date"]
        fig.add_trace(go.Scatter(x=[m], y=[net.loc[m]], mode="markers", showlegend=False,
                      marker=dict(color=DOWN, size=8, symbol="triangle-down")))
    fig.update_layout(**PLOT, height=300)
    fig.update_yaxes(tickprefix="$", tickformat=",.0f")
    return fig


def load_monthly_returns():
    """Compound the daily-returns CSV into a years × (Jan..Dec, YTD) pivot (decimals)."""
    if not MONTHLY_RETURNS_FILE or not os.path.exists(MONTHLY_RETURNS_FILE):
        return None
    try:
        raw = pd.read_csv(MONTHLY_RETURNS_FILE)
    except Exception:
        return None
    cols = {str(c).strip().lower(): c for c in raw.columns}
    dcol = next((cols[k] for k in ("date", "day", "dt") if k in cols), None)
    rcol = next((cols[k] for k in ("daily_return", "daily return", "return", "ret", "daily_ret")
                 if k in cols), None)
    if dcol is None or rcol is None:
        return None
    df = pd.DataFrame()
    s = raw[dcol].astype(str).str.strip()
    parts = s.str.split(r"[/.\-]", regex=True, expand=True)
    dayfirst = False
    if parts.shape[1] >= 2 and not s.str.match(r"^\d{4}").any():
        a = pd.to_numeric(parts[0], errors="coerce")
        b = pd.to_numeric(parts[1], errors="coerce")
        if (a > 12).any():
            dayfirst = True
        elif (b > 12).any():
            dayfirst = False
    df["date"] = pd.to_datetime(s, errors="coerce", format="mixed", dayfirst=dayfirst)
    df["ret"] = _to_num(raw[rcol])
    df = df.dropna(subset=["date", "ret"]).sort_values("date")
    if df.empty:
        return None
    df["Year"], df["Month"] = df["date"].dt.year, df["date"].dt.month
    monthly = df.groupby(["Year", "Month"])["ret"].apply(lambda x: (1 + x).prod() - 1)
    piv = monthly.unstack("Month").reindex(columns=range(1, 13))
    piv["YTD"] = [((np.prod(1 + row.dropna().values) - 1) if row.notna().any() else np.nan)
                  for _, row in piv.iterrows()]
    return piv


def fig_monthly_returns(piv):
    if piv is None or piv.empty:
        return None
    month_names = {1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
                   7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"}
    disp = piv.copy()
    disp.columns = [month_names.get(c, c) for c in disp.columns]
    header = ["<b>Year</b>"] + [f"<b>{c}</b>" for c in disp.columns]
    year_labels = [str(int(y)) for y in disp.index.tolist()]
    table_values, cell_colors = [year_labels], [[TXT] * len(disp)]
    for col in disp.columns:
        vals, colors = [], []
        for v in disp[col].values:
            if pd.isna(v):
                vals.append(""); colors.append(TXT)
            else:
                vals.append(f"{v:.1%}")
                colors.append(MONTHLY_POS if v > 0 else MONTHLY_NEG if v < 0 else TXT)
        table_values.append(vals); cell_colors.append(colors)
    fig = go.Figure(data=[go.Table(
        columnwidth=[46] + [38] * (len(disp.columns) - 1) + [52],
        header=dict(values=header, line_color=LINE, fill_color=PANEL2, align="center",
                    font=dict(color=TXT, family=MONO, size=12), height=30),
        cells=dict(values=table_values, line_color=LINE, fill_color=PANEL, align="center",
                   font=dict(color=cell_colors, family=MONO, size=12), height=28))])
    fig.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                      margin=dict(l=4, r=4, t=4, b=4), height=max(90, 34 + 29 * len(disp)))
    return fig


def fig_cum_return(daily, dd):
    bal = daily["balance"].dropna()
    cr = bal / bal.iloc[0] - 1.0
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=cr.index, y=cr.values, name="Cumulative return",
                  line=dict(color=BRASS, width=2.2), fill="tozeroy",
                  fillcolor="rgba(21,101,192,0.10)",
                  hovertemplate="%{x|%Y-%m-%d}<br>%{y:.1%}<extra></extra>"))
    if dd["max_dd_date"] is not None and dd["max_dd_date"] in cr.index:
        m = dd["max_dd_date"]
        fig.add_trace(go.Scatter(x=[m], y=[cr.loc[m]], mode="markers", showlegend=False,
                      marker=dict(color=DOWN, size=8, symbol="triangle-down")))
    fig.update_layout(**PLOT, height=300)
    fig.update_yaxes(tickformat=".0%")
    return fig


def _latest_positions_by_class(legs):
    """For the most recent bar: open positions per asset class as
    '<br>Long | NAME<br>Short | NAME' (direction + instrument)."""
    out = {}
    if legs is None or legs.empty:
        return out
    for ac, g in legs[legs["is_open"]].groupby("asset_class"):
        labels = []
        for _, r in g.iterrows():
            name = str(r["symbol"]).strip() or str(r["ib_ticker"]).strip()
            lab = f"{r['side']} | {name}"
            if lab not in labels:
                labels.append(lab)
        if labels:
            out[ac] = "<br>" + "<br>".join(labels)
    return out


def fig_exposure_bars(comp, index=None, legs=None, height=360):
    if comp is None or comp.empty:
        return go.Figure(layout=dict(**PLOT, height=height))
    e = comp if index is None else comp.reindex(pd.DatetimeIndex(index).normalize())
    e = e.dropna(how="all")
    e = e.loc[:, (e != 0).any(axis=0)]
    if e.empty:
        return go.Figure(layout=dict(**PLOT, height=height))
    # Re-normalise each day's row to sum to 100% so the stacked bars always fill
    # the full vertical space (guards against rounding and reindex gaps).
    e = e.fillna(0.0)
    rowsum = e.sum(axis=1).replace(0, np.nan)
    e = e.div(rowsum, axis=0).dropna(how="all")
    order = e.mean().sort_values(ascending=False).index.tolist()
    posmap = _latest_positions_by_class(legs)
    last = len(e.index) - 1
    fig = go.Figure()
    for i, c in enumerate(order):
        # Real values (zeros included) keep the stack at a true 100% — NaN would
        # leave gaps. 0%-height segments simply have no area to hover, so only
        # classes > 0% ever appear in the per-segment tooltip.
        cd = [""] * len(e.index)
        if last >= 0 and c in posmap:
            cd[last] = posmap[c]
        fig.add_trace(go.Bar(
            x=e.index, y=e[c], name=c, customdata=cd,
            marker=dict(color=_asset_color(c, i), line=dict(color="rgba(44,44,44,0.55)", width=0.6)),
            hovertemplate=f"{c}: %{{y:.0%}}%{{customdata}}<extra></extra>"))
    fig.update_layout(
        height=height, barmode="stack", bargap=0, paper_bgcolor=PANEL, plot_bgcolor=PANEL,
        font=dict(family=FONT, color=FT_COLORS["text"], size=12),
        margin=dict(l=56, r=24, t=56, b=40), hovermode="closest",
        hoverlabel=dict(bgcolor="#FFFFFF", bordercolor=LINE,
                        font=dict(family=MONO, size=11, color=TXT)),
        legend=dict(orientation="h", yanchor="bottom", y=1.08, xanchor="left", x=0,
                    font=dict(size=12), bgcolor="rgba(0,0,0,0)", traceorder="normal"))
    fig.update_yaxes(tickformat=".0%", range=[0, 1], gridcolor=FT_COLORS["grid"], zeroline=False)
    fig.update_xaxes(type="date", tickformat="%b %d", rangebreaks=[dict(bounds=["sat", "mon"])],
                     showgrid=False, showline=True, linecolor=FT_COLORS["text"],
                     ticks="outside", tickcolor=FT_COLORS["text"])
    return fig


def fig_drawdown(dd, equity=None):
    s = dd["series"]
    if equity is not None:
        eq = equity.dropna()
        peak = eq.cummax()
        dd_pct = ((eq - peak) / peak).reindex(s.index)
    else:
        mx = s.min()
        dd_pct = s / abs(mx) * (dd["max_dd_pct"] / 100.0) if mx != 0 else s * 0.0
    fig = go.Figure(go.Scatter(x=dd_pct.index, y=dd_pct.values, line=dict(color=DOWN, width=1),
                    fill="tozeroy", fillcolor="rgba(196,30,58,0.15)",
                    hovertemplate="%{x|%Y-%m-%d}<br>%{y:.1%}<extra></extra>"))
    fig.update_layout(**PLOT, height=150); fig.update_yaxes(tickformat=".0%")
    return fig


def fig_daily_returns(daily):
    r = daily["ret"].dropna()
    fig = go.Figure(go.Bar(x=r.index, y=r.values,
                    marker_color=[BAR_UP if v >= 0 else BAR_DOWN for v in r.values], marker_line_width=0,
                    hovertemplate="%{x|%Y-%m-%d}<br>%{y:.2%}<extra></extra>"))
    fig.update_layout(**PLOT, height=220); fig.update_yaxes(tickformat=".1%")
    return fig


def fig_exposure(exp, base=None):
    df = exp["by_instrument"]
    if df.empty:
        return go.Figure(layout=dict(**PLOT, height=260))
    df = df.iloc[::-1]
    use_pct = base and base != 0
    xvals = (df["gross_notional"] / base) if use_pct else df["gross_notional"]
    labels = [pct(v) for v in xvals] if use_pct else [money(v) for v in df["gross_notional"]]
    fig = go.Figure(go.Bar(y=df["symbol"], x=xvals, orientation="h",
                    marker_color=[BAR_UP if s == "Long" else BAR_DOWN for s in df["side"]], marker_line_width=0,
                    text=labels, textposition="auto", textfont=dict(family=MONO, size=10, color=ON_FILL)))
    fig.update_layout(**PLOT, height=max(260, 26 * len(df) + 40))
    fig.update_xaxes(tickformat=".0%" if use_pct else "~s", tickprefix="" if use_pct else "$")
    return fig


def fig_strategy(bystrat, base=None):
    if bystrat.empty:
        return go.Figure(layout=dict(**PLOT, height=220))
    use_pct = base and base != 0
    yvals = (bystrat["net_pnl"] / base) if use_pct else bystrat["net_pnl"]
    labels = [pct(v) for v in yvals] if use_pct else [money(v) for v in bystrat["net_pnl"]]
    fig = go.Figure(go.Bar(x=bystrat["trade_type"], y=yvals, width=0.5,
                    marker_color=[BAR_UP if v >= 0 else BAR_DOWN for v in yvals], marker_line_width=0,
                    text=labels, textposition="outside", cliponaxis=False,
                    textfont=dict(family=MONO, size=11, color=TXT)))
    fig.update_layout(**PLOT, height=220)
    fig.update_layout(margin=dict(l=48, r=18, t=40, b=32))
    vals = list(yvals)
    lo, hi = min(vals + [0]), max(vals + [0])
    pad = (hi - lo) * 0.18 or (abs(hi) or 1) * 0.18
    fig.update_yaxes(tickformat=".1%" if use_pct else ",.0f",
                     tickprefix="" if use_pct else "$", range=[lo - pad, hi + pad])
    return fig


def fig_asset(byac, base=None):
    if byac.empty:
        return go.Figure(layout=dict(**PLOT, height=240))
    df = byac.iloc[::-1]
    use_pct = base and base != 0
    xvals = [(v / base) if use_pct else v for v in df["net"].tolist()]
    cls = df["asset_class"].tolist()
    nets = [pct(v) if use_pct else money(v) for v in xvals]

    def ls_split(row):
        lo, sh = abs(float(row.get("long", 0.0))), abs(float(row.get("short", 0.0)))
        g = lo + sh
        return "—" if g <= 1e-9 else f"{lo/g*100:.0f}% | {sh/g*100:.0f}%"
    splits = [ls_split(r) for _, r in df.iterrows()]

    fig = go.Figure(go.Bar(y=cls, x=xvals, orientation="h",
                    marker_color=[BAR_UP if v >= 0 else BAR_DOWN for v in xvals], marker_line_width=0,
                    text=nets, textposition="auto", textfont=dict(family=MONO, size=11, color=ON_FILL),
                    cliponaxis=False))
    xmax, xmin = max(xvals + [0]), min(xvals + [0])
    span = (xmax - xmin) or (abs(xmax) or 1)
    pad = span * 0.015
    for c, v, sp in zip(cls, xvals, splits):
        fig.add_annotation(x=(v + pad) if v >= 0 else (v - pad), y=c, xref="x", yref="y",
                           xanchor="left" if v >= 0 else "right", showarrow=False,
                           text=sp, font=dict(family=MONO, size=11, color="#000000"))
    fig.update_layout(**PLOT, height=max(240, 34 * len(df) + 60))
    # Left margin sized to the longest asset-class label so none are clipped.
    lmargin = max(70, 9 * max((len(str(c)) for c in cls), default=8) + 16)
    fig.update_layout(margin=dict(l=lmargin, r=18, t=28, b=32))
    fig.update_xaxes(tickformat=".1%" if use_pct else "~s", tickprefix="" if use_pct else "$",
                     range=[min(xmin, 0) - span * 0.18, xmax + span * 0.26])
    return fig


def month_trade_drivers(trades, marks=None, month=None):
    if trades is None or trades.empty:
        return pd.DataFrame()
    mper = pd.Timestamp(month or pd.Timestamp.today()).to_period("M")
    df = trades.copy()
    df = df[(df["status"] == "Closed")
            & (df["entry_date"].dt.to_period("M") == mper)
            & (df["exit_date"].dt.to_period("M") == mper)]
    if df.empty:
        return pd.DataFrame()
    df["net"] = df["realised_pnl"] - df["comm"]
    rows = []
    for sym, g in df.groupby("instruments"):
        theme = ", ".join(dict.fromkeys(t for t in g["themes"].tolist() if str(t).strip())) or str(sym)
        acs = g["asset_class"].unique()
        rows.append({"symbol": sym, "asset_class": acs[0] if len(acs) == 1 else "Mixed",
                     "theme": theme, "net": float(g["net"].sum()), "n": int(len(g))})
    return pd.DataFrame(rows).sort_values("net", ascending=True).reset_index(drop=True)


def fig_trade_drivers(dr, base=None):
    if dr is None or dr.empty:
        return go.Figure(layout=dict(**PLOT, height=160))
    use_pct = base and base != 0
    xvals = [(v / base) if use_pct else v for v in dr["net"].tolist()]
    cls = dr["asset_class"].tolist()
    themes = dr["theme"].tolist()
    ypos = list(range(len(dr)))
    fig = go.Figure(go.Bar(y=ypos, x=xvals, orientation="h",
                    marker_color=[BAR_UP if v >= 0 else BAR_DOWN for v in xvals], marker_line_width=0,
                    text=themes, textposition="inside", insidetextanchor="start",
                    textfont=dict(family=FONT, size=12, color=ON_FILL), customdata=cls,
                    hovertemplate="%{customdata} · %{text}: "
                                  + ("%{x:+.1%}" if use_pct else "%{x:$,.0f}") + "<extra></extra>",
                    cliponaxis=False))
    xmax, xmin = max(xvals + [0]), min(xvals + [0])
    span = (xmax - xmin) or (abs(xmax) or 1)
    for yp, v in zip(ypos, xvals):
        fig.add_annotation(x=xmax + span * 0.15, y=yp, xref="x", yref="y", showarrow=False,
                           text=f"<b>{(pct(v) if use_pct else money(v))}</b>",
                           font=dict(family=MONO, size=13, color=UP if v >= 0 else DOWN), xanchor="left")
    fig.add_annotation(x=xmax + span * 0.15, y=1.0, xref="x", yref="paper", yanchor="bottom",
                       showarrow=False, text="NET", xanchor="left",
                       font=dict(family=MONO, size=10, color=MUTED))
    fig.update_layout(**PLOT, height=max(180, 50 * len(dr) + 60))
    fig.update_layout(margin=dict(l=112, r=92, t=28, b=36))
    fig.update_xaxes(tickformat="+.0%" if use_pct else "$,.0f",
                     range=[min(xmin, 0) - span * 0.02, xmax + span * 0.32],
                     zeroline=True, zerolinecolor=LINE)
    fig.update_yaxes(tickmode="array", tickvals=ypos, ticktext=cls, showgrid=False, autorange="reversed")
    return fig


# ===========================================================================
#  TABLE DATA (as display DataFrames; styled in the UI)
# ===========================================================================
def trades_table_df(trades, base=None):
    df = trades.copy()
    df["net"] = df["realised_pnl"] - df["comm"]
    use_pct = base and base != 0

    def _net(r):
        if r["status"] == "Open":
            return "OPEN"
        return pct_signed(r["net"] / base) if use_pct else money_signed(r["net"])
    out = pd.DataFrame({
        "ID": df["trade_id"].astype("Int64").astype(str),
        "Type": df["trade_type"],
        "Instruments": df.apply(lambda r: r["themes"] if str(r.get("themes", "")).strip()
                                else r["instruments"], axis=1),
        "Status": df["status"],
        "Hold": df["holding_days"].map(lambda v: f"{int(v)}d"),
        "Net P&L": df.apply(_net, axis=1)})
    return out


def open_table_df(marks, base=None):
    if marks.empty:
        return pd.DataFrame()
    use_pct = base and base != 0
    return pd.DataFrame({
        "Instrument": marks["symbol"], "Side": marks["side"],
        "Qty": marks["quantity"].map(lambda v: f"{int(v):,}"),
        "Entry": marks["entry_price"].map(lambda v: num(v, 4)),
        "Last": marks["last_price"].map(lambda v: num(v, 4)),
        "Unreal P&L": marks["unrealised_pnl"].map(
            (lambda v: pct_signed(v / base)) if use_pct else money_signed),
        "Gross Exp": marks["gross_notional"].map(
            (lambda v: pct(v / base)) if use_pct else (lambda v: money(v)))})


def asset_table_df(byac, base=None):
    if byac.empty:
        return pd.DataFrame()
    use_pct = base and base != 0
    _p = (lambda v: pct_signed(v / base)) if use_pct else money_signed
    return pd.DataFrame({
        "Asset Class": byac["asset_class"], "Trades": byac["trades"].map(lambda v: str(int(v))),
        "Open": byac["open_legs"].map(lambda v: str(int(v))),
        "Realized": byac["realised"].map(_p), "Unrealized": byac["unrealised"].map(_p),
        "Net": byac["net"].map(_p),
        "Win%": byac["win_rate"].map(lambda v: pct(v) if pd.notna(v) else "—")})


def fig_table(df, left_cols=(), color_cols=(), side_col=None, status_col=None,
              row_h=27, header_h=28, min_h=80):
    """Render a display DataFrame as an FT-cream Plotly table: PANEL2 header,
    PANEL cells, +/- coloured P&L columns, Long/Short + Open accents. Theme-proof
    (its colours don't depend on Streamlit's dataframe theming)."""
    cols = list(df.columns)
    n = len(df)
    header_vals = [f"<b>{str(c).upper()}</b>" for c in cols]
    cell_vals, cell_colors, aligns = [], [], []
    for c in cols:
        vals = ["" if pd.isna(v) else str(v) for v in df[c].tolist()]
        colors = []
        for v in vals:
            col = TXT
            if c in color_cols:
                if v.startswith("-"):
                    col = DOWN
                elif v.startswith("+"):
                    col = UP
            if side_col and c == side_col:
                col = UP if v == "Long" else (DOWN if v == "Short" else TXT)
            if status_col and c == status_col:
                col = BRASS if v == "Open" else col
            colors.append(col)
        cell_vals.append(vals)
        cell_colors.append(colors)
        aligns.append("left" if c in left_cols else "right")
    # Proportional column widths from the widest content (header included) so a
    # long header like "UNREALIZED" or "INSTRUMENTS" is never clipped.
    widths = [max([len(str(c)) + 2] + [len(v) for v in vals])
              for c, vals in zip(cols, cell_vals)]
    fig = go.Figure(data=[go.Table(
        columnwidth=widths,
        header=dict(values=header_vals, fill_color=PANEL2, line_color=LINE, align=aligns,
                    font=dict(color=MUTED, family=MONO, size=11), height=header_h),
        cells=dict(values=cell_vals, fill_color=PANEL, line_color=LINE, align=aligns,
                   font=dict(color=cell_colors, family=MONO, size=12), height=row_h))])
    fig.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                      margin=dict(l=2, r=2, t=2, b=2), height=max(min_h, header_h + 6 + row_h * n))
    return fig


# ===========================================================================
#  STREAMLIT UI
# ===========================================================================
st.set_page_config(page_title="Long Horizon Macro Strategy · Daily Review",
                   page_icon="📈", layout="wide")

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500;600&display=swap');
.stApp {{ background:{INK}; }}
.block-container {{ padding-top: 4.5rem; padding-left: 2rem; padding-right: 2rem; max-width: 1760px; }}
[data-testid="stVerticalBlockBorderWrapper"] {{ background:{PANEL}; border-radius:10px; }}
.hdr {{ color:{GREEK_BLUE}; font:700 21px {MONO}; letter-spacing:0.22em; margin:0.4rem 0 0.6rem 0; }}
.sect {{ color:{TXT}; font:600 13px {FONT}; letter-spacing:0.02em; margin:2px 0 8px 2px; }}
.ddlabel {{ color:{MUTED}; font:400 10px {MONO}; letter-spacing:0.06em; margin:6px 0 0 4px; }}
.note {{ color:{BRASS}; font:400 11px {MONO}; }}
.strow {{ display:flex; gap:10px; flex-wrap:wrap; margin-bottom:10px; }}
.sttile {{ flex:1; min-width:118px; background:{PANEL2}; border:1px solid {LINE};
           border-radius:8px; padding:10px 12px; font-family:{MONO}; }}
.sttile .l {{ color:{MUTED}; font-size:9.5px; letter-spacing:0.07em; text-transform:uppercase; }}
.sttile .v {{ font-size:17px; font-weight:600; margin-top:3px; }}
.stDataFrame {{ font-family:{MONO}; }}
</style>
""", unsafe_allow_html=True)


def section(text):
    st.markdown(f'<div class="sect">{text}</div>', unsafe_allow_html=True)


def stat_row(tiles):
    """tiles: list of (label, value, color)."""
    cells = "".join(
        f'<div class="sttile"><div class="l">{lab.upper()}</div>'
        f'<div class="v" style="color:{col}">{val}</div></div>'
        for lab, val, col in tiles)
    st.markdown(f'<div class="strow">{cells}</div>', unsafe_allow_html=True)


_CHART = {"displayModeBar": False}

# ---- header ----
st.markdown('<div class="hdr">LONG HORIZON MACRO STRATEGY · DAILY REVIEW</div>',
            unsafe_allow_html=True)

# ---- load ----
try:
    legs = load_trades()
except Exception as e:
    st.error(f"Could not load blotter:\n\n{e}")
    st.stop()

trades = group_trades(legs)
marks = marks_excel(legs)
totals = totals_excel(trades, marks)
daily = load_daily_pnl()
expo = load_exposure()
use_daily = daily is not None and len(daily) >= 2

base_cap = float(daily["balance"].iloc[0]) if (use_daily and daily["balance"].notna().any()) else None


def as_pct(dollars):
    return (dollars / base_cap) if (base_cap and base_cap != 0) else None


if use_daily:
    perf = daily_perf(daily)
    dd = drawdown(daily["balance"])
    var = value_at_risk(daily["ret"])
    equity_fig = fig_cum_return(daily, dd)
    dd_fig = fig_drawdown(dd, daily["balance"])
    daily_ret_fig = fig_daily_returns(daily)
    equity_index = daily["balance"].index
    net_pct = perf.get("total_ret")
    best_pct, worst_pct = daily["ret"].max(), daily["ret"].min()
else:
    curves = realized_curve(trades)
    perf = performance_ratios(curves["daily_net"]); perf["basis"] = "realized"
    dd = drawdown(curves["cum_net"])
    var = value_at_risk(curves["daily_net"])
    total_today = float(curves["cum_net"].iloc[-1]) + totals["unrealised"]
    equity_fig = fig_equity_excel(curves, dd, total_today)
    dd_fig = fig_drawdown(dd)
    daily_ret_fig = None
    equity_index = curves.index
    net_pct = as_pct(totals["net_total"])
    best_pct = as_pct(perf.get("best_day")); worst_pct = as_pct(perf.get("worst_day"))

ts = trade_stats(trades)
bystrat = by_strategy(trades)
byac = by_asset_class(trades, marks, legs)

realized_pct = as_pct(totals["realised"])
unreal_pct = as_pct(totals["unrealised"])
maxdd_pct = (dd["max_dd_pct"] / 100.0) if not np.isnan(dd["max_dd_pct"]) else None
n_days = int(perf.get("n_days") or 0)
ann_ret = (((1.0 + net_pct) ** (252.0 / n_days) - 1.0)
           if (use_daily and net_pct is not None and n_days > 0) else None)

# ---- notes ----
notes = []
if totals["unpriced"]:
    notes.append(f"{totals['unpriced']} open leg(s) have no value in the Unrealized column — "
                 "excluded from unrealized P&L (still marked at entry/last for exposure).")
if not use_daily:
    notes.append(f"'{DAILY_SHEET}' worksheet not found — %-metrics use realized trade P&L over "
                 "starting capital where available. Add the sheet for daily-return-based metrics.")
elif base_cap is None:
    notes.append("No starting Account Balance in the Daily PnL sheet — some %-metrics show '—'.")
for n in notes:
    st.markdown(f'<div class="note">⚠ {n}</div>', unsafe_allow_html=True)

# ---- metrics (two rows, % of starting capital) ----
net_label = f"Net PnL · {n_days}d" if (use_daily and n_days) else "Net PnL"
with st.container(border=True):
    stat_row([
        (net_label, pct(net_pct) if net_pct is not None else "—", pnl_color(net_pct)),
        ("Realized PnL", pct(realized_pct) if realized_pct is not None else "—", pnl_color(realized_pct)),
        ("Unrealized PnL", pct(unreal_pct) if unreal_pct is not None else "—", pnl_color(unreal_pct)),
        ("Ann Return", pct(ann_ret) if ann_ret is not None else "—", pnl_color(ann_ret)),
        ("Sharpe", num(perf.get("sharpe")), TXT),
        ("Sortino", num(perf.get("sortino")), TXT),
        ("Max Drawdown", pct(maxdd_pct) if maxdd_pct is not None else "—", DOWN),
        ("Ann Vol", pct(perf.get("ann_vol")) if use_daily else "—", TXT),
        ("Avg Hold", f"{ts['avg_hold']:.1f}d" if ts.get("n_closed") else "—", TXT),
    ])
    stat_row([
        ("Win Rate", pct(ts.get("win_rate")) if ts.get("n_closed") else "—", TXT),
        ("Closed Trades", str(ts["n_closed"]) if ts.get("n_closed") else "—", TXT),
        ("Best Day", pct(best_pct) if best_pct is not None else "—", UP),
        ("Worst Day", pct(worst_pct) if worst_pct is not None else "—", DOWN),
        ("Profit Factor", num(ts["profit_factor"]) if ts.get("n_closed") else "—", UP),
        ("VaR 95%", pct(var["var_95"]) if not np.isnan(var["var_95"]) else "—", DOWN),
        ("VaR 99%", pct(var["var_99"]) if not np.isnan(var["var_99"]) else "—", DOWN),
        ("CVaR 95%", pct(var["cvar_95"]) if not np.isnan(var["cvar_95"]) else "—", DOWN),
        ("CVaR 99%", pct(var["cvar_99"]) if not np.isnan(var["cvar_99"]) else "—", DOWN),
    ])

# ---- monthly returns ----
_monthly_fig = fig_monthly_returns(load_monthly_returns())
if _monthly_fig is not None:
    with st.container(border=True):
        section("Monthly returns  ·  %")
        st.plotly_chart(_monthly_fig, use_container_width=True, theme=None, config=_CHART)

# ---- row: curve/exposure/drawdown  +  daily returns/strategy/drivers ----
curve_title = ("Cumulative return  ·  %" if use_daily
               else "Realized P&L  ·  cumulative by exit date (blue marker = incl. open unrealized)")
dd_label = "Drawdown  ·  %" if use_daily else "Drawdown (realized)"

left, right = st.columns([2.2, 1])
with left:
    with st.container(border=True):
        section(curve_title)
        st.plotly_chart(equity_fig, use_container_width=True, theme=None, config=_CHART)
        if expo is not None and not expo.empty:
            section("Exposure by Asset Class")
            st.plotly_chart(fig_exposure_bars(expo, equity_index, legs=legs),
                            use_container_width=True, theme=None, config=_CHART)
        st.markdown(f'<div class="ddlabel">{dd_label}</div>', unsafe_allow_html=True)
        st.plotly_chart(dd_fig, use_container_width=True, theme=None, config=_CHART)
with right:
    with st.container(border=True):
        if daily_ret_fig is not None:
            section("Daily returns  ·  %")
            st.plotly_chart(daily_ret_fig, use_container_width=True, theme=None, config=_CHART)
        section("By strategy  ·  %")
        st.plotly_chart(fig_strategy(bystrat, base_cap), use_container_width=True, theme=None, config=_CHART)
        _mdr = month_trade_drivers(trades, marks)
        if _mdr is not None and not _mdr.empty:
            _mlabel = pd.Timestamp.today().strftime("%B")
            section(f"What drove {_mlabel}'s P&L  ·  trades entered this month")
            st.plotly_chart(fig_trade_drivers(_mdr, base_cap), use_container_width=True, theme=None, config=_CHART)

# ---- row: asset class / exposure / breakdown ----
c1, c2, c3 = st.columns([1, 1, 1.2])
with c1:
    with st.container(border=True):
        section("Net P&L by asset class  ·  %  |  Long % vs Short %")
        st.plotly_chart(fig_asset(byac, base_cap), use_container_width=True, theme=None, config=_CHART)
with c2:
    with st.container(border=True):
        section("Open exposure by instrument  ·  % of balance")
        st.plotly_chart(fig_exposure(exposure(marks), base_cap), use_container_width=True, theme=None, config=_CHART)
with c3:
    with st.container(border=True):
        section("Asset-class breakdown  ·  %")
        _adf = asset_table_df(byac, base_cap)
        if _adf.empty:
            st.caption("No asset-class data.")
        else:
            st.plotly_chart(fig_table(_adf, left_cols=["Asset Class"],
                                      color_cols=["Realized", "Unrealized", "Net"],
                                      row_h=42, min_h=260),
                            use_container_width=True, theme=None, config=_CHART)

# ---- row: tables ----
t1, t2 = st.columns([1.4, 1])
with t1:
    with st.container(border=True):
        section("Trades")
        _tdf = trades_table_df(trades, base_cap)
        st.plotly_chart(fig_table(_tdf, left_cols=["Type", "Instruments", "Status"],
                                  color_cols=["Net P&L"], status_col="Status"),
                        use_container_width=True, theme=None, config=_CHART)
with t2:
    with st.container(border=True):
        section("Open positions")
        _odf = open_table_df(marks, base_cap)
        if _odf.empty:
            st.caption("No open positions.")
        else:
            st.plotly_chart(fig_table(_odf, left_cols=["Instrument", "Side"],
                                      color_cols=["Unreal P&L"], side_col="Side",
                                      row_h=42, min_h=260),
                            use_container_width=True, theme=None, config=_CHART)