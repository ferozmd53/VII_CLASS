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

The first login opens a temporary Chrome/Edge window and saves the OAuth session in the workbook. A valid saved session is reused; the browser is opened again only when the saved token is invalid or expired.

The access token is stored in `LOGIN!B8:B14` and a hidden `SESSION` sheet. Keep the workbook private. The terminal uses the `LOGIN` sheet as the main settings sheet; legacy `Config` settings are copied into `LOGIN` on first startup.


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

The Finvasia terminal remains the main engine. RSI, StochRSI, SMA Stoch and Bollinger Bands are added to `LIVE_TRADE` columns `AE:AL`, using the same daily historical-data and indicator formulas as the Shoonya indicator project. Daily indicator history is currently enabled for NSE symbols, matching the original indicator implementation; other exchanges continue to receive the Finvasia live terminal data.

## Start

```bat
python run.py
```

You can also pass a workbook path:

```bat
python run.py "C:\\full\\path\\VII_CLASS.xlsm"
```

## Source acknowledgements

- Finvasia Excel terminal base: [PythonTraderExcelTradeTerminal](https://github.com/Indian-Algorithmic-Trading-Community/PythonTraderExcelTradeTerminal), `Finvasia` folder.
- Shoonya OAuth reuse and indicator logic: [Shoonya-Multi-Symbol-Live-Historical-Data-with-Indicators](https://github.com/ferozmd53/Shoonya-Multi-Symbol-Live-Historical-Data-with-Indicators).

The original Finvasia source attribution and license are retained in this repository. This project is an integration; it is not an official Shoonya or Finvasia product.
