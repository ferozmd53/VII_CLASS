# VII_CLASS — Shoonya Excel Trade Terminal

Live Shoonya trading terminal. This version is REAL mode only: trade triggers submit actual broker orders. There is no paper-trading mode. Trading involves financial risk; verify every setting, quantity, product type and broker response before enabling live triggers.

## Installation

Install Python 3.11–3.13 (64-bit) and Microsoft Excel.

Open Command Prompt and run these commands one by one:

```bat
pip uninstall VII_CLASS -y
pip install --upgrade --no-cache-dir git+https://github.com/ferozmd53/VII_CLASS.git
```

If you use Python 3.13 installed at `C:\Python313`, use:

```bat
C:\Python313\python.exe -m pip uninstall VII_CLASS -y
C:\Python313\python.exe -m pip install --upgrade --no-cache-dir git+https://github.com/ferozmd53/VII_CLASS.git
```

**Important:** Do not use `pip uninstall git+https://...`. Uninstall requires the installed package name, which is `VII_CLASS`. The install command pulls the latest repository version and its declared dependencies, including SciPy.

## Login setup

1. Open `VII_CLASS.xlsm` in Excel and enable macros.
2. Run `python run.py` once; this creates the `LOGIN` sheet and stops if the credentials are blank.
3. Open the workbook again and fill these cells in `LOGIN`, then run `python run.py` again:

| Cell | Value |
|---|---|
| B2 | Shoonya API Client ID / API key |
| B3 | Shoonya User ID |
| B4 | Shoonya password |
| B5 | TOTP secret |
| B6 | Shoonya API secret |
| B15 | Automatically generated Algo ID (`LOGIN!B3` User ID + `_U`; display only) |
| B24 | SL-LMT Market Safety percentage (default 1; allowed 0–10) |

The first login opens a temporary Chrome/Edge window and saves the OAuth session in the workbook. A valid saved session is reused; the browser is opened again only when the saved token is invalid or expired.

The access token is stored in `LOGIN!B8:B14` and a hidden `SESSION` sheet. The core engine is named `VII_CLASS.py` and is started through `run.py`. Keep the workbook private. The terminal uses the `LOGIN` sheet as the main settings sheet; legacy `Config` settings are copied into `LOGIN` on first startup.

### Indicator and historical-data settings

The indicator engine reads its settings from the same `LOGIN` sheet. Enter values in column B; if a cell is blank or invalid, the default is inserted automatically.

| LOGIN cell | Setting | Default |
|---|---|---:|
| B16 | Historical Days | 500 |
| B17 | RSI Length | 14 |
| B18 | StochRSI Length | 14 |
| B19 | SMA Stoch Length | 3 |
| B20 | Bollinger Length | 20 |
| B21 | Bollinger StdDev | 2.0 |
| B22 | Stoch Upper | 70 |
| B23 | Stoch Lower | 30 |

The terminal automatically builds the Shoonya Algo ID from `LOGIN!B3` (User ID) by appending `_U` for every order and square-off request. The generated value is displayed in `LOGIN!B15`; do not enter it manually. If the User ID already ends in `_U`, the suffix is not duplicated. The SL-LMT safety buffer is read from `LOGIN!B24` (default 1%).

These settings are loaded when the terminal starts and are also monitored while the terminal runs. If you edit any value in `LOGIN!B16:B23`, the indicator engine automatically reloads with the new settings, clears the old indicator output, and downloads fresh history in background workers; restarting the terminal is not required. The existing indicator formulas and Excel output columns are unchanged.


## Live-only trading and alerts

- The live trade sheet is named `LIVE_TRADE` (the previous `Trade_Terminal` tab is renamed automatically on first startup). The terminal forces `LIVE_TRADE!S2` to REAL; paper/simulated fills are removed.
- A live order is blocked when Shoonya login is not active, internet is unavailable, or the Excel workbook is not accessible.
- Voice alerts announce login failure, internet loss/recovery, Excel workbook disconnect/recovery, and WebSocket disconnect/connection failure.
- The Standard and PRO option-chain outputs retain IV and market data. Delta, Gamma, Theta, Vega and Rho columns are removed.
- Excel disconnect alerts require the Python terminal process to remain running. If Windows terminates Python itself, that process cannot speak an alert.

## LIVE_TRADE columns

| Columns | Purpose |
|---|---|
| M | Quantity |
| N | BUY Trigger — enter `BUY` |
| O | SELL Trigger — enter `SELL` |
| P | Entry Signal / order type |
| Q | Entry Limit Price |
| R:S:T | Entry fill price, order ID and remarks |
| U:X | Exit signal, fill price, order ID and remarks |
| Y:AD | Target, SL, trailing settings, trade status and PnL |
| AE:AL | RSI, StochRSI, SMA Stoch, Bollinger Bands and signal columns |

The old BUY/SELL dropdown is removed. Enter `BUY` in column N or `SELL` in column O; the selected trigger is cleared after the order submission attempt. Entry Signal and Entry Limit Price remain separate inputs.

## Indicators

The  main engine. RSI, StochRSI, SMA Stoch and Bollinger Bands are added to `LIVE_TRADE` columns `AG:AN`, using daily historical candles and the same indicator formulas. Daily history and indicator calculations are enabled for symbols from NSE, NFO, BSE, BFO, CDS and MCX, provided the symbol is present in `LIVE_TRADE`, resolves to a valid Shoonya instrument token, and Shoonya returns historical candles for it. History downloads run in background workers so the live quote loop is not held up. Derivative contracts may have shorter available history or no candles after expiry; the engine reports unavailable/short history in the console instead of fabricating values.

## Separate live 1-minute / 5-minute indicator sheet

A separate `Intraday_Indicators.py` module creates the `INDICATOR` worksheet without changing the existing daily indicators or trading/order columns. Enter symbols in `INDICATOR!A4:A1000` using `EXCHANGE:TRADINGSYMBOL` format; the initial list is copied from `LIVE_TRADE`.

The sheet displays the converted Pine indicators for both 1-minute and 5-minute candles:
- Mirror Upper Band: length 5, multiplier 1.7, source High, reference Open.
- ATR Step Channel: width 0.7, ATR length 200, highest ATR lookback 100, source High, with fakeout-up/down flags.

Shoonya intraday history is requested asynchronously. The current forming candle is updated from the latest WebSocket LTP, and the Excel sheet refreshes about once per second. The ATR channel needs a warm-up of approximately 300 candles; values remain blank until enough valid history is available. Fakeout flags are provisional while the candle is forming and can change before candle close. This indicator module does not place or modify orders.

## Start

```bat
python run.py
```

You can also pass a workbook path:

```bat
python run.py "C:\\full\\path\\VII_CLASS.xlsm"
```

## Source acknowledgements

- Shoonya OAuth reuse and indicator logic: [Shoonya-Multi-Symbol-Live-Historical-Data-with-Indicators](https://github.com/ferozmd53/Shoonya-Multi-Symbol-Live-Historical-Data-with-Indicators).

The original Finvasia source attribution and license are retained in this repository. This project is an integration; it is not an official Shoonya or Finvasia product.


## Live weekly OHLCV

The separate `Weekly.py` module creates a `WEEKLY_DATA` worksheet automatically. It loads recent daily candles in background workers, combines them into the current Monday-to-Friday trading week, and refreshes the current day's values from live Shoonya ticks.

Columns: Symbol, Exchange, Week Start, Week End, Open, High, Low, Close, Volume, Last Update. The Volume value is the sum of daily volumes for the current week, with today's cumulative live volume replacing the historical value for today. This sheet is a current-week summary for the symbols selected in `WEEKLY_DATA!A4:A1000`; it does not place or modify orders.


### Weekly symbol selection and row layout

Open `WEEKLY_DATA` and enter or edit symbols in `A4:A1000` using the same `EXCHANGE:TRADINGSYMBOL` format as `LIVE_TRADE`. When the sheet is first created, the current `LIVE_TRADE!A4:A1000` list is copied into it as a starting list. Weekly output is written in columns B:J on the same row as each symbol in column A. The weekly module requests subscriptions for selected weekly symbols independently; it does not create or modify orders.
