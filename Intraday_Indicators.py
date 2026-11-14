"""Separate 1-minute and 5-minute live TradingView-style indicators.

The engine consumes Shoonya intraday candles in background workers and folds the
latest WebSocket LTP into the currently forming candle. Excel is only touched by
refresh_from_excel(), which is called by the terminal's main Excel thread.

Pine conversions from the user's verified TradingView reference:
1) Mirror Upper Band: SMA(low, 5) - (-0.3) * population_stdev(low, 5),
   mirrored around the current candle open.
2) ATR Step Channel: highest(ATR(200), 100) * 0.7 channel width, seeded
   at bar_index 200; only confirmed candles can move the step.
"""
import time

import sqlite3
from pathlib import Path
from contextlib import contextmanager


class CandleCache:
    """Shared persistent SQLite OHLCV cache stored on the Windows Desktop."""

    def __init__(self, database_path=None):
        desktop = Path.home() / "Desktop"
        desktop.mkdir(parents=True, exist_ok=True)
        self.database_path = Path(database_path) if database_path else desktop / "market_data_cache.db"
        self._initialize()

    def _connect(self):
        connection = sqlite3.connect(str(self.database_path), timeout=30)
        connection.execute("PRAGMA journal_mode=WAL")
        connection.execute("PRAGMA busy_timeout=30000")
        return connection

    @contextmanager
    def _connection(self):
        connection = self._connect()
        try:
            yield connection
            connection.commit()
        finally:
            connection.close()

    def _initialize(self):
        with self._connection() as connection:
            connection.execute("""
                CREATE TABLE IF NOT EXISTS candles (
                    exchange TEXT NOT NULL,
                    symbol TEXT NOT NULL,
                    timeframe TEXT NOT NULL,
                    candle_time TEXT NOT NULL,
                    open REAL, high REAL, low REAL, close REAL, volume REAL,
                    source TEXT NOT NULL DEFAULT 'Shoonya',
                    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    PRIMARY KEY (exchange, symbol, timeframe, candle_time)
                )
            """)
            connection.execute("""
                CREATE INDEX IF NOT EXISTS idx_candles_lookup
                ON candles(exchange, symbol, timeframe, candle_time)
            """)

    def save(self, exchange, symbol, timeframe, frame, source="Shoonya"):
        if frame is None or frame.empty:
            return 0
        rows = []
        for row in frame.itertuples(index=False):
            values = row._asdict()
            stamp = pd.Timestamp(values.get("datetime"))
            if pd.isna(stamp):
                continue
            def number(name):
                value = values.get(name, 0)
                return None if pd.isna(value) else float(value)
            rows.append((
                str(exchange).upper(), str(symbol).upper(), str(timeframe),
                stamp.strftime("%Y-%m-%d %H:%M:%S"),
                number("open"), number("high"), number("low"),
                number("close"), number("volume") or 0.0, source
            ))
        if not rows:
            return 0
        with self._connection() as connection:
            connection.executemany("""
                INSERT INTO candles
                (exchange,symbol,timeframe,candle_time,open,high,low,close,volume,source,updated_at)
                VALUES (?,?,?,?,?,?,?,?,?,?,CURRENT_TIMESTAMP)
                ON CONFLICT(exchange,symbol,timeframe,candle_time) DO UPDATE SET
                    open=excluded.open, high=excluded.high, low=excluded.low,
                    close=excluded.close, volume=excluded.volume,
                    source=excluded.source, updated_at=CURRENT_TIMESTAMP
            """, rows)
        return len(rows)

    def load(self, exchange, symbol, timeframe, limit=None):
        query = """
            SELECT candle_time AS datetime, open, high, low, close, volume
            FROM candles WHERE exchange=? AND symbol=? AND timeframe=?
            ORDER BY candle_time
        """
        params = [str(exchange).upper(), str(symbol).upper(), str(timeframe)]
        if limit:
            query = """
                SELECT * FROM (
                    SELECT candle_time AS datetime, open, high, low, close, volume
                    FROM candles WHERE exchange=? AND symbol=? AND timeframe=?
                    ORDER BY candle_time DESC LIMIT ?
                ) ORDER BY datetime
            """
            params.append(int(limit))
        with self._connection() as connection:
            rows = connection.execute(query, params).fetchall()
        return pd.DataFrame(rows, columns=["datetime", "open", "high", "low", "close", "volume"]) if rows else None

    def latest_time(self, exchange, symbol, timeframe):
        with self._connection() as connection:
            row = connection.execute("""
                SELECT MAX(candle_time) FROM candles
                WHERE exchange=? AND symbol=? AND timeframe=?
            """, (str(exchange).upper(), str(symbol).upper(), str(timeframe))).fetchone()
        return row[0] if row else None


from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta

import pandas as pd


class IntradayIndicatorEngine:
    SUPPORTED_INTERVALS = (1, 5)
    HISTORY_DAYS = 60
    RETRY_SECONDS = 60
    REFRESH_SECONDS = 1.0
    HISTORY_REFRESH_SECONDS = 60

    # Pine inputs supplied by the user.
    MIRROR_LENGTH = 5
    MIRROR_MULTIPLIER = -0.3         # matches the TradingView inputs you use
    LIVE_LTP_MERGE = True            # fold WebSocket LTP into the forming candle (TV updates high/low tick by tick)
    FIRST_BAR_TIME = None            # e.g. "2026-11-02 09:15": time of the FIRST bar TV loaded; trims history so bar_index==200 matches TV
    CHANNEL_WIDTH = 0.7
    CHANNEL_ATR_LENGTH = 200
    CHANNEL_HIGHEST_LENGTH = 100
    CHANNEL_INITIAL_BAR = 200

    HEADERS = [
        "Symbol",
        "1m Candle Time", "1m Open", "1m High", "1m Low", "1m Close", "1m Volume",
        "1m upperpink", "1m mirrorUpperBand",
        "5m Candle Time", "5m Open", "5m High", "5m Low", "5m Close", "5m Volume",
        "5m upperpink", "5m mirrorUpperBand",
        "Data Status",
        "1m upperpink (2nd-last)", "1m mirrorUpperBand (2nd-last)",
        "5m upperpink (2nd-last)", "5m mirrorUpperBand (2nd-last)",
        "LTP",
    ]
    INPUTS = [
        ["TIMEFRAME", "1m,5m", "MIRROR SOURCE", "open", "MIRROR SMA SOURCE", "low",
         "MIRROR LENGTH", 5, "MIRROR MULT", -0.3, "HISTORY DAYS", 60],
        ["CHANNEL SOURCE", "low", "ATR LENGTH", 200, "ATR HIGHEST", 100,
         "CHANNEL MULT", 0.7, "OUTPUT", "OHLCV + upperpink + mirrorUpperBand"],
    ]

    def __init__(self):
        self.candle_cache = CandleCache()
        print(f"SQLite candle cache: {self.candle_cache.database_path}")
        self.refresh_at = {}
        self.executor = ThreadPoolExecutor(
            max_workers=4, thread_name_prefix="ShoonyaIntradayHistory"
        )
        self.history = {}
        self.pending = {}
        self.retry_at = {}
        self.latest = {}
        self.last_refresh = 0.0
        self.sheet_ready = False
        self.last_written_row = 3

    @staticmethod
    def _number(value):
        try:
            number = float(str(value).replace(",", "").strip())
            return number if pd.notna(number) and number > 0 else None
        except (TypeError, ValueError, OverflowError):
            return None

    @staticmethod
    def _parse_history(raw):
        if isinstance(raw, dict):
            return None
        if not isinstance(raw, (list, tuple)) or not raw:
            return None

        rows = []
        for item in raw:
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
            "ssboe": "epoch",
        }, inplace=True)

        if "datetime" in frame:
            parsed = frame["datetime"].map(
                lambda value: pd.to_datetime(value, dayfirst=True, errors="coerce")
            )
        else:
            parsed = pd.Series(pd.NaT, index=frame.index, dtype="datetime64[ns]")
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

    def _fetch_history(self, api, exchange, token, interval, key, history_days, start_timestamp=None):
        now = datetime.now()
        start = datetime.fromtimestamp(start_timestamp) if start_timestamp else now - timedelta(days=history_days)
        try:
            # Shoonya API signature: get_time_price_series(exchange, token,
            # starttime, endtime, interval). Keep this REST call off Excel/tick loops.
            raw = api.get_time_price_series(
                exchange, str(token), int(start.timestamp()),
                int(now.timestamp()), str(interval)
            )
            frame = self._parse_history(raw)
            if frame is None:
                message = raw.get("emsg") or raw.get("stat") if isinstance(raw, dict) else "no candles"
                print(
                    f"Intraday {interval}m history unavailable for {key}: {message}"
                )
                return None
            if len(frame) < self.CHANNEL_INITIAL_BAR + 2:
                print(
                    f"Intraday {interval}m history short for {key}: "
                    f"{len(frame)} candles; ATR channel needs about "
                    f"{self.CHANNEL_ATR_LENGTH + self.CHANNEL_HIGHEST_LENGTH - 1} bars for full warm-up."
                )
            print(
                f"Intraday {interval}m history loaded for {key}: "
                f"{len(frame)} candles (requested {history_days} days)"
            )
            cache_symbol = key.split('|', 1)[0]
            self.candle_cache.save(exchange, cache_symbol, f'{interval}m', frame)
            return frame
        except Exception as exc:
            print(f"Intraday {interval}m history request failed for {key}: {exc}")
            return None

    def _get_history(self, api, exchange, token, interval, key, history_days):
        cache_key = (key, interval, int(history_days))
        cache_symbol = key.split("|", 1)[0]
        frame = self.history.get(cache_key)
        future = self.pending.get(cache_key)

        if frame is None:
            frame = self.candle_cache.load(exchange, cache_symbol, f"{interval}m")
            if frame is not None and not frame.empty:
                frame["datetime"] = pd.to_datetime(frame["datetime"], errors="coerce")
                self.history[cache_key] = frame
                print(f"Intraday {interval}m cache loaded for {cache_symbol}: {len(frame)} candles")

        if future is not None and future.done():
            self.pending.pop(cache_key, None)
            try:
                refreshed = future.result()
                if refreshed is not None and not refreshed.empty:
                    frame = pd.concat([frame, refreshed], ignore_index=True) if frame is not None else refreshed
                    frame = frame.sort_values("datetime").drop_duplicates("datetime", keep="last").reset_index(drop=True)
                    self.candle_cache.save(exchange, cache_symbol, f"{interval}m", frame)
                    self.history[cache_key] = frame
                    self.refresh_at[cache_key] = time.time() + self.HISTORY_REFRESH_SECONDS
            except Exception as exc:
                print(f"Intraday history worker error for {key} ({interval}m): {exc}")
                self.refresh_at[cache_key] = time.time() + self.RETRY_SECONDS

        if frame is None:
            if future is None and time.time() >= self.retry_at.get(cache_key, 0):
                self.pending[cache_key] = self.executor.submit(
                    self._fetch_history, api, exchange, token, interval, key,
                    int(history_days), None
                )
            return None

        self.history[cache_key] = frame
        if cache_key not in self.pending and time.time() >= self.refresh_at.get(cache_key, 0):
            latest = pd.to_datetime(frame["datetime"], errors="coerce").max()
            start_timestamp = max(0, int((latest - pd.Timedelta(minutes=2 * interval)).timestamp()))
            self.pending[cache_key] = self.executor.submit(
                self._fetch_history, api, exchange, token, interval, key,
                int(history_days), start_timestamp
            )
            self.refresh_at[cache_key] = time.time() + self.HISTORY_REFRESH_SECONDS
        return frame

    @staticmethod
    def _merge_live_candle(frame, ltp, now, interval):
        bucket = pd.Timestamp(now).floor(f"{interval}min")
        frame = frame.copy()
        if not frame.empty:
            frame["datetime"] = pd.to_datetime(frame["datetime"], errors="coerce")
            frame = frame.dropna(subset=["datetime"])
            # If REST already returned the current partial candle, update it.
            matches = frame.index[frame["datetime"] == bucket]
        else:
            matches = []

        if len(matches):
            index = matches[-1]
            frame.loc[index, "high"] = max(float(frame.loc[index, "high"]), ltp)
            frame.loc[index, "low"] = min(float(frame.loc[index, "low"]), ltp)
            frame.loc[index, "close"] = ltp
        else:
            previous = frame.iloc[-1] if not frame.empty else None
            frame = pd.concat([frame, pd.DataFrame([{
                "datetime": bucket,
                "open": ltp,
                "high": max(ltp, float(previous["high"])) if previous is not None and previous["datetime"] == bucket else ltp,
                "low": min(ltp, float(previous["low"])) if previous is not None and previous["datetime"] == bucket else ltp,
                "close": ltp,
                "volume": 0,
            }])], ignore_index=True)

        frame = frame.sort_values("datetime").drop_duplicates("datetime", keep="last")
        return frame.reset_index(drop=True)

    @staticmethod
    def _wilder_atr(frame, length):
        high = frame["high"].astype(float)
        low = frame["low"].astype(float)
        close = frame["close"].astype(float)
        previous_close = close.shift(1)
        true_range = pd.concat([
            high - low,
            (high - previous_close).abs(),
            (low - previous_close).abs(),
        ], axis=1).max(axis=1)
        # Pine ta.atr() uses Wilder's RMA, seeded by the first SMA(length).
        values = true_range.astype(float).tolist()
        result = [float("nan")] * len(values)
        if len(values) < length:
            return pd.Series(result, index=frame.index, dtype="float64")
        seed = sum(values[:length]) / float(length)
        result[length - 1] = seed
        previous_atr = seed
        for index in range(length, len(values)):
            previous_atr = ((previous_atr * (length - 1)) + values[index]) / length
            result[index] = previous_atr
        return pd.Series(result, index=frame.index, dtype="float64")

    @staticmethod
    def _source(frame, source_name):
        name = str(source_name or "close").strip().lower()
        if name not in ("open", "high", "low", "close"):
            name = "close"
        return frame[name].astype(float)

    def _calculate(self, frame, interval, live=True, settings=None):
        if frame is None or len(frame) < 2:
            return None, None, None, None

        settings = settings or {}
        first_bar = settings.get("first_bar_time", self.FIRST_BAR_TIME)
        if first_bar and "datetime" in frame:
            keep = pd.to_datetime(frame["datetime"], errors="coerce") >= pd.Timestamp(first_bar)
            if keep.any():
                frame = frame[keep].reset_index(drop=True)
        mirror_source = self._source(frame, settings.get("mirror_source", "open"))
        mirror_sma_source = self._source(frame, settings.get("mirror_sma_source", "low"))
        mirror_length = max(1, int(settings.get("mirror_length", self.MIRROR_LENGTH)))
        mirror_multiplier = float(settings.get("mirror_multiplier", self.MIRROR_MULTIPLIER))
        channel_source = self._source(frame, settings.get("channel_source", "low"))
        atr_length = max(1, int(settings.get("atr_length", self.CHANNEL_ATR_LENGTH)))
        highest_length = max(1, int(settings.get("atr_highest", self.CHANNEL_HIGHEST_LENGTH)))
        channel_width = float(settings.get("channel_multiplier", self.CHANNEL_WIDTH))

        sma = mirror_sma_source.rolling(mirror_length).mean()
        deviation = mirror_sma_source.rolling(mirror_length).std(ddof=0)
        real_upper = sma - (mirror_multiplier * deviation)
        mirror_upper = (2.0 * mirror_source) - real_upper

        atr = self._wilder_atr(frame, atr_length)
        channel_atr = atr.rolling(highest_length).max() * channel_width

        step_values = []
        upper_values = []
        lower_values = []
        initial_bar = int(settings.get("initial_bar", self.CHANNEL_INITIAL_BAR))  # Pine: bar_index == 200 (fixed)
        for index in range(len(frame)):
            previous_step = step_values[index - 1] if index else float("nan")
            width = channel_atr.iloc[index]
            if pd.isna(width):
                upper = lower = float("nan")
            else:
                upper = previous_step + float(width)
                lower = previous_step - float(width)

            source_value = float(channel_source.iloc[index])
            confirmed = not (live and index == len(frame) - 1)
            if index == initial_bar:
                step = source_value
            elif confirmed and not pd.isna(upper) and not pd.isna(lower) and (
                source_value > upper or source_value < lower
            ):
                step = source_value
            else:
                step = previous_step

            step_values.append(step)
            upper_values.append(upper)
            lower_values.append(lower)

        current = len(frame) - 1
        previous = current - 1
        mirror_value = mirror_upper.iloc[current]
        mirror_value = round(float(mirror_value), 4) if pd.notna(mirror_value) else None
        previous_mirror = mirror_upper.iloc[previous]
        previous_mirror = round(float(previous_mirror), 4) if pd.notna(previous_mirror) else None
        # Pine: plot(st, "upperpink", fuchsia) -> the plotted value is the STEP line `st`,
        # not the upper band (st[1] + ATR). Old code returned the band, so it never matched.
        upperpink = step_values[current]
        upperpink = round(float(upperpink), 4) if pd.notna(upperpink) and upperpink > 0 else None
        previous_upperpink = step_values[previous]
        previous_upperpink = round(float(previous_upperpink), 4) if pd.notna(previous_upperpink) and previous_upperpink > 0 else None
        return upperpink, mirror_value, previous_upperpink, previous_mirror

    def _result_for_interval(self, history, interval, settings, ltp=None):
        """Use the newest timestamped candle returned by Shoonya; never invent one."""
        if history is None or history.empty:
            return None, None, None, None, None, None

        frame = history.copy()
        frame["datetime"] = pd.to_datetime(frame["datetime"], errors="coerce")
        frame = frame.dropna(subset=["datetime"]).sort_values("datetime")
        if frame.empty:
            return None, None, None, None, None, None

        # Shoonya history is the source of truth for both candle time and OHLCV.
        # Do not synthesize a candle from local PC time or the latest WebSocket LTP.
        frame = frame.drop_duplicates("datetime", keep="last").reset_index(drop=True)
        # Broker REST candles are refreshed only every HISTORY_REFRESH_SECONDS, so the
        # forming candle's high/low/close go stale. TradingView updates them every tick,
        # and the mirror band uses the CURRENT bar's high -> fold LTP in. Only the candle
        # matching the PC clock bucket (or the one right after the last broker candle) is
        # touched, so nothing is invented after market hours.
        if self.LIVE_LTP_MERGE and ltp:
            bucket = pd.Timestamp(datetime.now()).floor(f"{interval}min")
            gap = bucket - pd.Timestamp(frame["datetime"].iloc[-1])
            if timedelta(0) <= gap <= timedelta(minutes=interval):
                frame = self._merge_live_candle(frame, ltp, datetime.now(), interval)
        candle = frame.iloc[-1]
        upperpink, mirror_upper, previous_upperpink, previous_mirror = self._calculate(
            frame, interval, live=True, settings=settings
        )

        candle_time = pd.Timestamp(candle["datetime"]).to_pydatetime()
        volume = self._number(candle.get("volume")) or 0
        ohlcv = [
            self._number(candle.get("open")),
            self._number(candle.get("high")),
            self._number(candle.get("low")),
            self._number(candle.get("close")),
            volume,
        ]
        return candle_time, ohlcv, upperpink, mirror_upper, previous_upperpink, previous_mirror

    def update_symbol(self, api, symbol, token_key, quote, settings=None):
        exchange, token = token_key.split("|", 1)
        ltp = self._number(quote.get("lp") if isinstance(quote, dict) else None)
        results = {}
        for interval in self.SUPPORTED_INTERVALS:
            key = f"{symbol}|{token}"
            history = self._get_history(
                api, exchange, token, interval, key,
                (settings or {}).get("history_days", self.HISTORY_DAYS)
            )
            if history is None:
                results[interval] = None
                continue
            result = self._result_for_interval(history, interval, settings, ltp)
            results[interval] = result

        has_data = any(
            result is not None and len(result) == 6 and result[1] is not None
            for result in results.values()
        )
        status = "BROKER CANDLE" if has_data else "LOADING BROKER HISTORY"
        self.latest[symbol] = (ltp, results, status)
        return results.get(1), results.get(5), status

    def _get_or_create_sheet(self, workbook):
        """Return the INTRADAY sheet, creating it and seeding symbols if needed."""
        names = [sheet.name for sheet in workbook.sheets]
        is_new = "INTRADAY" not in names

        if is_new:
            sheet = workbook.sheets.add("INTRADAY")
            print("Created INTRADAY sheet for live 1-minute and 5-minute indicators.")
        else:
            sheet = workbook.sheets["INTRADAY"]

        # Seed symbols both on first creation and when an existing sheet is blank.
        try:
            current_symbols = sheet.range("A4:A1000").value or []
            has_symbols = any(str(value or "").strip() for value in current_symbols)
            if not has_symbols:
                source = next(
                    (item for item in workbook.sheets
                     if item.name in ("LIVE_TRADE", "Trade_Terminal")),
                    None,
                )
                if source is not None:
                    source_symbols = source.range("A4:A1000").value or []
                    if any(str(value or "").strip() for value in source_symbols):
                        sheet.range("A4:A1000").value = source_symbols
                        print("Copied symbols from the live trade sheet into INTRADAY!A4:A1000.")
                    else:
                        print("INTRADAY sheet is ready; no symbols found in the live trade sheet.")
                else:
                    print("INTRADAY sheet is ready; add EXCHANGE:TRADINGSYMBOL values in column A.")
        except Exception as exc:
            print(f"Could not seed INTRADAY symbols from the live trade sheet: {exc}")

        return sheet, is_new

    def refresh_from_excel(
        self, workbook, api, subscribe_callback, quote_callback, token_callback
    ):
        now = time.monotonic()
        if now - self.last_refresh < self.REFRESH_SECONDS:
            return
        self.last_refresh = now

        try:
            sheet, is_new = self._get_or_create_sheet(workbook)
            if is_new or not self.sheet_ready:
                # Populate editable input controls only when blank, preserving user settings.
                for row_number, defaults in enumerate(self.INPUTS, start=1):
                    existing = sheet.range(f"A{row_number}:L{row_number}").value
                    existing = existing if isinstance(existing, list) else [existing]
                    existing = (existing + [None] * 12)[:12]
                    for col_number, default_value in enumerate(defaults, start=1):
                        if existing[col_number - 1] in (None, ""):
                            sheet.range((row_number, col_number)).value = default_value
                # NOTE: the old migration that rewrote J1 (0.5/1.7 -> -0.3) and F1
                # (high -> low) on every start-up was removed so the INTRADAY input
                # cells are never overwritten; set them to your TradingView inputs.
                try:
                    old_history_days = int(float(sheet.range("L1").value))
                    if old_history_days == 7:
                        sheet.range("L1").value = 60
                except (TypeError, ValueError):
                    pass

                # No manual candle-time selector: timestamps always come from Shoonya.
                sheet.range("M1:N1").clear_contents()
                sheet.range("A3:W3").value = [self.HEADERS]
                # Clear only unused columns beyond the 23-column INTRADAY output.
                sheet.range("X3:Z1000").clear_contents()
                sheet.range("A1:N2").api.Font.Bold = True
                sheet.range("A3:W3").api.Font.Bold = True
                sheet.range("A1:W3").api.WrapText = True
                sheet.range("B4:B1000").api.NumberFormat = "dd-mm-yyyy hh:mm"
                sheet.range("J4:J1000").api.NumberFormat = "dd-mm-yyyy hh:mm"
                self.sheet_ready = True
                try:
                    workbook.save()
                    print("INTRADAY inputs, headers and symbol list saved to the active workbook.")
                except Exception as save_error:
                    print(f"INTRADAY sheet save failed: {save_error}")

            def input_value(cell, default):
                value = sheet.range(cell).value
                return default if value in (None, "") else value

            settings = {
                "mirror_source": input_value("D1", "open"),
                "mirror_sma_source": input_value("F1", "low"),
                "mirror_length": input_value("H1", 5),
                "mirror_multiplier": input_value("J1", -0.3),
                "channel_source": input_value("B2", "low"),
                "atr_length": input_value("D2", 200),
                "atr_highest": input_value("F2", 100),
                "channel_multiplier": input_value("H2", 0.7),
                "history_days": max(1, int(float(input_value("L1", self.HISTORY_DAYS)))),
            }

            symbols = sheet.range("A4:A1000").value or []
            last_index = max(
                (index for index, value in enumerate(symbols)
                 if str(value or "").strip()),
                default=-1,
            )
            last_row = 4 + last_index
            if last_index < 0:
                if self.last_written_row > 3:
                    sheet.range(f"A4:R{self.last_written_row}").clear_contents()
                    self.last_written_row = 3
                return

            output = []
            for value in symbols[:last_index + 1]:
                symbol = str(value or "").strip().upper()
                if not symbol or ":" not in symbol:
                    output.append([None] * len(self.HEADERS))
                    continue

                subscribe_callback(symbol)
                quote = quote_callback(symbol)
                token_key = token_callback(symbol)
                if not token_key or not isinstance(quote, dict):
                    output.append([symbol] + [None] * 16 + ["WAITING FOR SUBSCRIPTION"] + [None] * 5)
                    continue

                one_minute, five_minute, status = self.update_symbol(
                    api, symbol, token_key, quote, settings=settings
                )
                def timeframe_values(result):
                    if not result or len(result) != 6 or result[1] is None:
                        return [None] * 8
                    candle_time, ohlcv, upperpink, mirror_upper, _, _ = result
                    return [candle_time] + list(ohlcv) + [upperpink, mirror_upper]

                one_previous = one_minute[4:6] if one_minute and len(one_minute) == 6 else [None, None]
                five_previous = five_minute[4:6] if five_minute and len(five_minute) == 6 else [None, None]
                output.append(
                    [symbol]
                    + timeframe_values(one_minute)
                    + timeframe_values(five_minute)
                    + [status]
                    + list(one_previous)
                    + list(five_previous)
                    + [ltp]
                )

            sheet.range(f"A4:W{last_row}").value = output
            if self.last_written_row > last_row:
                sheet.range(f"A{last_row + 1}:W{self.last_written_row}").clear_contents()
            self.last_written_row = last_row
        except Exception as exc:
            print(f"INTRADAY sheet refresh failed: {exc}")
