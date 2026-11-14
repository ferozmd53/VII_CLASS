"""Live weekly OHLCV sheet for the Shoonya Excel terminal.

Historical daily candles are loaded asynchronously and grouped into trading
weeks. The current day's OHLC and cumulative volume are refreshed from each
Shoonya WebSocket tick, so the current weekly candle changes live.
"""
import time
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor

import pandas as pd
from Shoonya_Indicators import fetch_bse_sensex_history


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
        self.refresh_at = {}

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
            frame["datetime"] = pd.to_datetime(
                epoch, unit="s", utc=True, errors="coerce"
            ).dt.tz_convert("Asia/Kolkata").dt.tz_localize(None)
        elif "datetime" in frame:
            parsed = frame["datetime"].map(
                lambda value: pd.to_datetime(value, dayfirst=True, errors="coerce")
            )
            if "epoch" in frame:
                epoch = pd.to_numeric(frame["epoch"], errors="coerce")
                epoch_ist = pd.to_datetime(
                    epoch, unit="s", utc=True, errors="coerce"
                ).dt.tz_convert("Asia/Kolkata").dt.tz_localize(None)
                parsed = parsed.fillna(epoch_ist)
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
            if exchange.upper() == "BSE" and symbol.strip().upper() == "SENSEX":
                parsed_count = len(frame) if frame is not None else 0
                if parsed_count < 7:
                    try:
                        print(
                            f"BSE:SENSEX weekly Shoonya history incomplete ({parsed_count} rows); "
                            "trying BSE index archive fallback"
                        )
                        frame = fetch_bse_sensex_history(start_date, end_date)
                        print(
                            f"BSE:SENSEX weekly BSE archive history loaded: {len(frame)} candles "
                            f"({frame['datetime'].iloc[0].date()} to {frame['datetime'].iloc[-1].date()})"
                        )
                    except Exception as fallback_exc:
                        print(f"BSE:SENSEX weekly BSE archive fallback failed: {fallback_exc}")
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
        frame = self.history.get(key)
        future = self.pending.get(key)

        if future is not None and future.done():
            self.pending.pop(key, None)
            try:
                refreshed = future.result()
                if refreshed is not None and not refreshed.empty:
                    frame = refreshed
                    self.history[key] = refreshed
                    self.refresh_at[key] = time.time() + 60
            except Exception as exc:
                print(f"Weekly history worker error for {exchange}:{symbol}: {exc}")
                self.refresh_at[key] = time.time() + self.retry_seconds

        if frame is None:
            if future is None and time.time() >= self.retry_at.get(key, 0):
                self.pending[key] = self.executor.submit(
                    self._fetch_history, api, exchange, symbol, key
                )
            return None

        # Keep the current weekly summary synchronized with newly published
        # broker daily candles while keeping the Excel/live loop non-blocking.
        if key not in self.pending and time.time() >= self.refresh_at.get(key, 0):
            self.refresh_at[key] = time.time() + 60
            self.pending[key] = self.executor.submit(
                self._fetch_history, api, exchange, symbol, key
            )
        return frame

    def update(self, api, exchange, symbol, quote=None):
        """Build weekly OHLCV only from timestamped broker daily candles.

        Do not overwrite broker history with WebSocket LTP or create a candle
        from the PC date. If today's broker candle is not available yet, the
        newest available broker candle remains the source of truth.
        """
        exchange = str(exchange or "").strip().upper()
        symbol = str(symbol or "").strip().upper()
        if not exchange or not symbol:
            return None

        key = f"{exchange}|{symbol}"
        frame = self._get_history(api, exchange, symbol, key)
        if frame is None or frame.empty:
            return None

        frame = frame.copy()
        frame["datetime"] = pd.to_datetime(frame["datetime"], errors="coerce")
        frame = frame.dropna(subset=["datetime"]).sort_values("datetime")
        frame = frame.drop_duplicates("datetime", keep="last").reset_index(drop=True)
        if frame.empty:
            return None

        # Keep only valid broker OHLC rows. No live quote or local-time candle
        # is inserted here; the latest broker candle controls the active week.
        for column in ("open", "high", "low", "close", "volume"):
            if column not in frame:
                frame[column] = 0
            frame[column] = pd.to_numeric(frame[column], errors="coerce")
        frame = frame.dropna(subset=["open", "high", "low", "close"])
        if frame.empty:
            return None

        frame["week"] = frame["datetime"].dt.to_period("W-SUN")
        weekly = frame.groupby("week", sort=True).agg(
            open=("open", "first"),
            high=("high", "max"),
            low=("low", "min"),
            close=("close", "last"),
            volume=("volume", "sum")
        )
        latest_week = frame["datetime"].iloc[-1].to_period("W-SUN")
        if latest_week not in weekly.index:
            return None

        candle = weekly.loc[latest_week]
        week_start = latest_week.start_time.date()
        week_end = week_start + timedelta(days=4)
        latest_candle_time = frame["datetime"].iloc[-1]
        self.latest_rows[key] = [
            exchange, week_start.strftime("%d/%m/%Y"),
            week_end.strftime("%d/%m/%Y"),
            round(float(candle["open"]), 4),
            round(float(candle["high"]), 4),
            round(float(candle["low"]), 4),
            round(float(candle["close"]), 4),
            int(candle["volume"]) if pd.notna(candle["volume"]) else 0,
            latest_candle_time.strftime("%d/%m/%Y %H:%M")
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
