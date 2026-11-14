"""Live weekly OHLCV sheet for the Shoonya Excel terminal.

Historical daily candles are loaded asynchronously and grouped into trading
weeks. The current day's OHLC and cumulative volume are refreshed from each
Shoonya WebSocket tick, so the current weekly candle changes live.
"""
import time
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor

import pandas as pd


class WeeklyDataEngine:
    HEADERS = [
        "Symbol", "Exchange", "Week Start", "Week End",
        "Open", "High", "Low", "Close", "Volume", "Last Update"
    ]

    def __init__(self, history_days=60, retry_seconds=300):
        self.history_days = int(history_days)
        self.retry_seconds = int(retry_seconds)
        self.history = {}
        self.pending = {}
        self.retry_at = {}
        self.latest_rows = {}
        self.executor = ThreadPoolExecutor(
            max_workers=3, thread_name_prefix="ShoonyaWeeklyHistory"
        )
        self.sheet_ready = False
        self.last_refresh_monotonic = 0.0
        self.last_history_error = {}

    @staticmethod
    def _number(value):
        try:
            if value in (None, ""):
                return None
            number = float(value)
            if pd.isna(number):
                return None
            return number
        except (TypeError, ValueError):
            return None

    @staticmethod
    def _parse_history(raw):
        rows = []
        if isinstance(raw, dict):
            return None
        for item in raw or []:
            if isinstance(item, dict):
                rows.append(item)
            elif isinstance(item, str):
                try:
                    import json
                    parsed = json.loads(item)
                    if isinstance(parsed, dict):
                        rows.append(parsed)
                except Exception:
                    continue
        if not rows:
            return None

        frame = pd.DataFrame(rows)
        frame.rename(columns={
            "time": "datetime", "into": "open", "inth": "high",
            "intl": "low", "intc": "close", "intv": "volume",
            "ssboe": "epoch"
        }, inplace=True)

        if "datetime" not in frame and "epoch" in frame:
            epoch = pd.to_numeric(frame["epoch"], errors="coerce")
            frame["datetime"] = pd.to_datetime(epoch, unit="s", errors="coerce")
        elif "datetime" in frame:
            parsed = frame["datetime"].map(
                lambda value: pd.to_datetime(value, dayfirst=True, errors="coerce")
            )
            if "epoch" in frame:
                epoch = pd.to_numeric(frame["epoch"], errors="coerce")
                parsed = parsed.fillna(
                    pd.to_datetime(epoch, unit="s", errors="coerce")
                )
            frame["datetime"] = parsed

        required = ("datetime", "open", "high", "low", "close")
        if any(column not in frame.columns for column in required):
            return None
        if "volume" not in frame:
            frame["volume"] = 0

        for column in ("open", "high", "low", "close", "volume"):
            frame[column] = pd.to_numeric(frame[column], errors="coerce")
        frame = frame.dropna(subset=["datetime", "open", "high", "low", "close"])
        frame = frame.sort_values("datetime").drop_duplicates("datetime")
        return frame.reset_index(drop=True) if not frame.empty else None

    def _fetch_history(self, api, exchange, symbol, key):
        end_date = datetime.now().replace(hour=23, minute=59, second=59, microsecond=0)
        start_date = (end_date - timedelta(days=self.history_days)).replace(
            hour=0, minute=0, second=0, microsecond=0
        )
        try:
            if exchange.upper() == "BSE" and symbol.strip().upper() == "SENSEX":
                print(
                    f"BSE:SENSEX weekly history request: symbol={symbol}, "
                    f"exchange={exchange}, days={self.history_days}, "
                    f"from={start_date.strftime('%Y-%m-%d %H:%M:%S')}, "
                    f"to={end_date.strftime('%Y-%m-%d %H:%M:%S')}"
                )
            # Keep the workbook alias, but request history using Shoonya's
            # actual NSE spot-index trading symbol.
            history_symbol = (
                "Nifty 50"
                if exchange.upper() == "NSE" and symbol.strip().upper() == "NIFTY INDEX"
                else symbol
            )
            raw = api.get_daily_price_series(
                exchange=exchange,
                tradingsymbol=history_symbol,
                startdate=str(int(start_date.timestamp())),
                enddate=str(int(end_date.timestamp()))
            )
            if exchange.upper() == "BSE" and symbol.strip().upper() == "SENSEX":
                raw_count = len(raw) if isinstance(raw, (list, tuple)) else "not-a-list"
                print(
                    f"BSE:SENSEX weekly raw history response: "
                    f"type={type(raw).__name__}, count={raw_count}, "
                    f"sample={str(raw[:2] if isinstance(raw, (list, tuple)) else raw)[:1200]}"
                )
            frame = self._parse_history(raw)
            if frame is None:
                message = raw.get("emsg") or raw.get("stat") if isinstance(raw, dict) else "no daily candles"
                print(f"Weekly history unavailable for {exchange}:{symbol}: {message}")
            else:
                print(
                    f"Weekly history loaded for {exchange}:{symbol}: "
                    f"{len(frame)} daily candles"
                )
            self.retry_at[key] = time.time() + self.retry_seconds
            return frame
        except Exception as exc:
            self.retry_at[key] = time.time() + self.retry_seconds
            print(f"Weekly history request failed for {exchange}:{symbol}: {exc}")
            return None

    def _get_history(self, api, exchange, symbol, key):
        if key in self.history:
            return self.history[key]
        future = self.pending.get(key)
        if future is None:
            if time.time() >= self.retry_at.get(key, 0):
                self.pending[key] = self.executor.submit(
                    self._fetch_history, api, exchange, symbol, key
                )
            return None
        if not future.done():
            return None
        self.pending.pop(key, None)
        try:
            frame = future.result()
        except Exception as exc:
            print(f"Weekly history worker error for {exchange}:{symbol}: {exc}")
            frame = None
        if frame is not None:
            self.history[key] = frame
        return frame

    def update(self, api, exchange, symbol, quote):
        exchange = str(exchange or "").strip().upper()
        symbol = str(symbol or "").strip().upper()
        if not exchange or not symbol or not isinstance(quote, dict):
            return None

        key = f"{exchange}|{symbol}"
        frame = self._get_history(api, exchange, symbol, key)
        if frame is None or frame.empty:
            return None

        now = datetime.now()
        today = now.date()
        ltp = self._number(quote.get("lp"))
        if ltp is None or ltp <= 0:
            return None

        frame = frame.copy()
        day_mask = frame["datetime"].dt.date == today
        day_rows = frame.index[day_mask]
        day_open = self._number(quote.get("o"))
        day_high = self._number(quote.get("h"))
        day_low = self._number(quote.get("l"))
        day_volume = self._number(quote.get("v"))

        if len(day_rows):
            row_index = day_rows[-1]
            old = frame.loc[row_index]
            frame.loc[row_index, "open"] = day_open if day_open is not None else old["open"]
            frame.loc[row_index, "high"] = max(
                value for value in (self._number(old["high"]), day_high, ltp)
                if value is not None
            )
            frame.loc[row_index, "low"] = min(
                value for value in (self._number(old["low"]), day_low, ltp)
                if value is not None
            )
            frame.loc[row_index, "close"] = ltp
            if day_volume is not None:
                frame.loc[row_index, "volume"] = day_volume
        else:
            frame = pd.concat([frame, pd.DataFrame([{
                "datetime": pd.Timestamp(today),
                "open": day_open if day_open is not None else ltp,
                "high": max(day_high, ltp) if day_high is not None else ltp,
                "low": min(day_low, ltp) if day_low is not None else ltp,
                "close": ltp,
                "volume": day_volume if day_volume is not None else 0
            }])], ignore_index=True)

        frame = frame.sort_values("datetime").reset_index(drop=True)
        self.history[key] = frame

        frame["week"] = frame["datetime"].dt.to_period("W-SUN")
        weekly = frame.groupby("week", sort=True).agg(
            open=("open", "first"),
            high=("high", "max"),
            low=("low", "min"),
            close=("close", "last"),
            volume=("volume", "sum")
        )
        current_week = pd.Timestamp(today).to_period("W-SUN")
        if current_week not in weekly.index:
            return None
        candle = weekly.loc[current_week]
        week_start = current_week.start_time.date()
        week_end = week_start + timedelta(days=4)
        self.latest_rows[key] = [
            exchange, week_start.strftime("%d/%m/%Y"),
            week_end.strftime("%d/%m/%Y"),
            round(float(candle["open"]), 4),
            round(float(candle["high"]), 4),
            round(float(candle["low"]), 4),
            round(float(candle["close"]), 4),
            int(candle["volume"]) if pd.notna(candle["volume"]) else 0,
            now.strftime("%d/%m/%Y %H:%M")
        ]
        return self.latest_rows[key]

    def _get_or_create_sheet(self, workbook):
        sheet_names = [sheet.name for sheet in workbook.sheets]
        if "WEEKLY_DATA" in sheet_names:
            return workbook.sheets["WEEKLY_DATA"], False
        sheet = workbook.sheets.add("WEEKLY_DATA")
        print("Created WEEKLY_DATA sheet for live weekly OHLCV.")
        try:
            live_sheet = next(
                (item for item in workbook.sheets
                 if item.name in ("LIVE_TRADE", "Trade_Terminal")),
                None
            )
            if live_sheet is not None:
                # Seed the new weekly input list from the current live-trade rows.
                sheet.range("A4:A1000").value = live_sheet.range("A4:A1000").value
        except Exception as exc:
            print(f"Could not copy initial LIVE_TRADE symbols to WEEKLY_DATA: {exc}")
        return sheet, True

    def refresh_from_excel(self, workbook, api, subscribe_callback, quote_callback):
        try:
            now_mono = time.monotonic()
            if now_mono - self.last_refresh_monotonic < 1.0:
                return
            self.last_refresh_monotonic = now_mono

            sheet, is_new = self._get_or_create_sheet(workbook)
            if is_new or not self.sheet_ready:
                sheet.range("A3:J3").value = [[
                    "Symbol", "Exchange", "Week Start", "Week End",
                    "Open", "High", "Low", "Close", "Volume", "Last Update"
                ]]
                sheet.range("A3:J3").api.Font.Bold = True
                self.sheet_ready = True

            selected_values = sheet.range("A4:A1000").value or []
            output_rows = []
            for value in selected_values:
                symbol = str(value or "").strip().upper()
                if len(symbol) < 6 or symbol[3] != ":":
                    output_rows.append([None] * 9)
                    continue

                exchange, tradingsymbol = symbol[:3], symbol[4:]
                subscribe_callback(symbol)
                quote = quote_callback(symbol)
                row = self.update(api, exchange, tradingsymbol, quote)
                output_rows.append(row if row else [exchange] + [None] * 8)

            # Keep each weekly output on the exact row where its symbol is entered.
            sheet.range("B4:J1000").value = output_rows
        except Exception as exc:
            print(f"Weekly Excel sheet update failed: {exc}")
