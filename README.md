# VII_CLASS — Shoonya Excel Trade Terminal

Educational and research software. Trading involves financial risk; verify every setting and broker response before using real funds.

## Installation

Install Python 3.11–3.13 (64-bit) and Microsoft Excel.

Open Command Prompt and run:
##
pip install --upgrade pip
pip install git+https://github.com/ferozmd53/VII_CLASS.git
##

## Login setup

1. Open `Finvasia_Trade_Terminal_v3.xlsm` in Excel and enable macros.
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

The access token is stored in `LOGIN!B8:B14` and a hidden `SESSION` sheet. Keep the workbook private.

## Indicators

The Finvasia terminal remains the main engine. RSI, StochRSI, SMA Stoch and Bollinger Bands are added to `Trade_Terminal` columns `AD:AK`, using the same daily historical-data and indicator formulas as the Shoonya indicator project. Daily indicator history is currently enabled for NSE symbols, matching the original indicator implementation; other exchanges continue to receive the Finvasia live terminal data.

## Start

```bat
python run.py
```

You can also pass a workbook path:

```bat
python run.py "C:\\full\\path\\Finvasia_Trade_Terminal_v3.xlsm"
```

## Source acknowledgements

- Finvasia Excel terminal base: [PythonTraderExcelTradeTerminal](https://github.com/Indian-Algorithmic-Trading-Community/PythonTraderExcelTradeTerminal), `Finvasia` folder.
- Shoonya OAuth reuse and indicator logic: [Shoonya-Multi-Symbol-Live-Historical-Data-with-Indicators](https://github.com/ferozmd53/Shoonya-Multi-Symbol-Live-Historical-Data-with-Indicators).

The original Finvasia source attribution and license are retained in this repository. This project is an integration; it is not an official Shoonya or Finvasia product.
