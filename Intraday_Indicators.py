"""Separate 1-minute and 5-minute live TradingView-style indicators.

The engine consumes Shoonya intraday candles in background workers and folds the
latest WebSocket LTP into the currently forming candle. Excel is only touched by
refresh_from_excel(), which is called by the terminal's main Excel thread.

Pine conversions:
1) Mirror Upper Band: SMA(high, 5) - 1.7 * population_stdev(high, 5),
   mirrored around the current candle open.
2) ATR Step Channel: highest(ATR(200), 100) * 0.7 channel width, with the
   confirmed-bar step state and live provisional fakeout checks.
"""
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta

import pandas as pd


class IntradayIndicatorEngine:
    SUPPORTED_INTERVALS = (1, 5)
    HISTORY_DAYS = 7
    RETRY_SECONDS = 60
    REFRESH_SECONDS = 1.0

    # Pine inputs supplied by the user.
    MIRROR_LENGTH = 5
    MIRROR_MULTIPLIER = 1.7
    CHANNEL_WIDTH = 0.7
    CHANNEL_ATR_LENGTH = 200
    CHANNEL_HIGHEST_LENGTH = 100
    CHANNEL_INITIAL_BAR = 200

    HEADERS = [
        "Symbol",
        "1m Open", "1m High", "1m Low", "1m Close", "1m Volume",
        "1m upperpink", "1m mirrorUpperBand",
        "5m Open", "5m High", "5m Low", "5m Close", "5m Volume",
        "5m upperpink", "5m mirrorUpperBand",
        "Data Status",
    ]
    INPUTS = [
        ["TIMEFRAME", "1m,5m", "MIRROR SOURCE", "open", "MIRROR SMA SOURCE", "high",
         "MIRROR LENGTH", 5, "MIRROR MULT", 1.7],
        ["CHANNEL SOURCE", "high", "ATR LENGTH", 200, "ATR HIGHEST", 100,
         "CHANNEL MULT", 0.7, "OUTPUT", "OHLCV + upperpink + mirrorUpperBand"],
    ]

    def __init__(self):
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

        if "datetime" not in frame and "epoch" in frame:
            frame["datetime"] = pd.to_numeric(frame["epoch"], errors="coerce")
            frame["datetime"] = pd.to_datetime(
                frame["datetime"], unit="s", errors="coerce"
            )
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

    def _fetch_history(self, api, exchange, token, interval, key):
        now = datetime.now()
        start = now - timedelta(days=self.HISTORY_DAYS)
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
                    f"{self.CHANNEL_INITIAL_BAR + 100} bars for full warm-up."
                )
            print(
                f"Intraday {interval}m history loaded for {key}: "
                f"{len(frame)} candles"
            )
            return frame
        except Exception as exc:
            print(f"Intraday {interval}m history request failed for {key}: {exc}")
            return None

    def _get_history(self, api, exchange, token, interval, key):
        cache_key = (key, interval)
        if cache_key in self.history:
            return self.history[cache_key]

        future = self.pending.get(cache_key)
        if future is None:
            if time.time() >= self.retry_at.get(cache_key, 0):
                self.pending[cache_key] = self.executor.submit(
                    self._fetch_history, api, exchange, token, interval, key
                )
            return None
        if not future.done():
            return None

        self.pending.pop(cache_key, None)
        try:
            frame = future.result()
        except Exception as exc:
            print(f"Intraday history worker error for {key} ({interval}m): {exc}")
            frame = None
        if frame is None:
            self.retry_at[cache_key] = time.time() + self.RETRY_SECONDS
            return None
        self.history[cache_key] = frame
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
            return None, None

        settings = settings or {}
        mirror_source = self._source(frame, settings.get("mirror_source", "open"))
        mirror_sma_source = self._source(frame, settings.get("mirror_sma_source", "high"))
        mirror_length = max(1, int(settings.get("mirror_length", self.MIRROR_LENGTH)))
        mirror_multiplier = float(settings.get("mirror_multiplier", self.MIRROR_MULTIPLIER))
        channel_source = self._source(frame, settings.get("channel_source", "high"))
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
        initial_bar = atr_length
        for index in range(len(frame)):
            previous_step = step_values[index - 1] if index else 0.0
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
            elif confirmed and not pd.isna(upper) and (
                source_value > upper or source_value < lower
            ):
                step = source_value
            else:
                step = previous_step

            step_values.append(step)
            upper_values.append(upper)
            lower_values.append(lower)

        current = len(frame) - 1
        mirror_value = mirror_upper.iloc[current]
        mirror_value = round(float(mirror_value), 4) if pd.notna(mirror_value) else None
        upperpink = step_values[current]
        upperpink = round(float(upperpink), 4) if upperpink > 0 else None
        return upperpink, mirror_value

    def update_symbol(self, api, symbol, token_key, quote, settings=None):
        exchange, token = token_key.split("|", 1)
        ltp = self._number(quote.get("lp") if isinstance(quote, dict) else None)
        if ltp is None:
            return None, None, "WAITING FOR LIVE TICK"

        results = {}
        for interval in self.SUPPORTED_INTERVALS:
            key = f"{symbol}|{token}"
            history = self._get_history(api, exchange, token, interval, key)
            if history is None:
                results[interval] = (None, None, None)
                continue
            live_frame = self._merge_live_candle(
                history, ltp, datetime.now(), interval
            )
            self.history[(key, interval)] = live_frame
            upperpink, mirror_upper = self._calculate(
                live_frame, interval, live=True, settings=settings
            )
            candle = live_frame.iloc[-1]
            volume = self._number(candle.get("volume")) or 0
            ohlcv = [
                self._number(candle.get("open")),
                self._number(candle.get("high")),
                self._number(candle.get("low")),
                self._number(candle.get("close")),
                volume,
            ]
            results[interval] = (ohlcv, upperpink, mirror_upper)

        status = "LIVE" if any(
            value[0] is not None for value in results.values()
        ) else "LOADING HISTORY"
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
                    existing = sheet.range(f"A{row_number}:J{row_number}").value
                    existing = existing if isinstance(existing, list) else [existing]
                    existing = (existing + [None] * 10)[:10]
                    for col_number, default_value in enumerate(defaults, start=1):
                        if existing[col_number - 1] in (None, ""):
                            sheet.range((row_number, col_number)).value = default_value
                sheet.range("A3:P3").value = [self.HEADERS]
                # Remove obsolete indicator columns from the previous INTRADAY layout.
                sheet.range("Q3:Z1000").clear_contents()
                sheet.range("A1:J2").api.Font.Bold = True
                sheet.range("A3:P3").api.Font.Bold = True
                sheet.range("A1:J3").api.WrapText = True
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
                "mirror_sma_source": input_value("F1", "high"),
                "mirror_length": input_value("H1", 5),
                "mirror_multiplier": input_value("J1", 1.7),
                "channel_source": input_value("B2", "high"),
                "atr_length": input_value("D2", 200),
                "atr_highest": input_value("F2", 100),
                "channel_multiplier": input_value("H2", 0.7),
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
                    sheet.range(f"A4:P{self.last_written_row}").clear_contents()
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
                    output.append([symbol] + [None] * 14 + ["WAITING FOR SUBSCRIPTION"])
                    continue

                one_minute, five_minute, status = self.update_symbol(
                    api, symbol, token_key, quote, settings=settings
                )
                def timeframe_values(result):
                    if not result or result[0] is None:
                        return [None] * 7
                    ohlcv, upperpink, mirror_upper = result
                    return list(ohlcv) + [upperpink, mirror_upper]

                output.append(
                    [symbol]
                    + timeframe_values(one_minute)
                    + timeframe_values(five_minute)
                    + [status]
                )

            sheet.range(f"A4:P{last_row}").value = output
            if self.last_written_row > last_row:
                sheet.range(f"A{last_row + 1}:P{self.last_written_row}").clear_contents()
            self.last_written_row = last_row
        except Exception as exc:
            print(f"INTRADAY sheet refresh failed: {exc}")
