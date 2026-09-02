# Book Monitor — Streamlit track record

The Excel "Book" dashboard as a Streamlit app: KPIs, equity curve + drawdown,
daily P&L, risk stats (Sharpe/Sortino/vol/VaR), trade stats, by-strategy and
by-asset-class breakdowns, exposure, and the trades / open-positions tables.
No Interactive Brokers, no local machine at run time.

## Files
```
streamlit_app.py          the app
requirements.txt
.streamlit/config.toml     dark theme
track_record.xlsx          <-- you add this: your workbook (Trades + Daily PnL sheets)
```

## Run locally
```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
# opens http://localhost:8501
```
If `track_record.xlsx` isn't in the folder, the app shows an upload box so you
can drop a workbook in from the browser.

## Deploy on Streamlit Community Cloud (free)
1. Put this folder — including your `track_record.xlsx` — in a GitHub repo
   (public or private both work).
2. Go to **share.streamlit.io**, sign in with GitHub, **New app**.
3. Pick the repo + branch, set **Main file path** to `streamlit_app.py`, Deploy.
4. You get a public URL like `https://<you>-<app>.streamlit.app`.

To refresh the numbers later, replace `track_record.xlsx` in the repo and push —
it redeploys automatically.

## Workbook format
- **Trades** sheet: the blotter (Entry/Exit Date, Symbol, IB_TICKER, Quantity,
  Entry/Exit Price, Notional, Leg PnL, Trade PnL, Trade ID, Trade Type, and
  optionally Last Price, Unrealized, Comm, Asset Class).
- **Daily PnL** sheet (optional but recommended): Date, Account Balance, PnL.
  When present, Sharpe/vol/drawdown and the equity curve use these daily
  returns; without it they fall back to a realised-trade basis.

## Notes
- Open-position unrealised P&L is read from the **Unrealized** column; `Last
  Price` is used only for exposure.
- Anyone with the URL can view it. To gate it, add a password check with
  `st.secrets` (ask if you want that dropped in), or keep the repo private and
  share only the app link.
