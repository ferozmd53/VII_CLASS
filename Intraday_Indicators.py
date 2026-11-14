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
        "Symbol", "LTP", "Last Update",
        "1m Mirror Upper", "1m Real Upper", "1m SMA High 5", "1m STD High 5",
        "1m Channel ST", "1m Channel Upper", "1m Channel Lower",
        "1m FakeOut Up", "1m FakeOut Down", "1m Trend", "1m Candle Time",
        "5m Mirror Upper", "5m Real Upper", "5m SMA High 5", "5m STD High 5",
        "5m Channel ST", "5m Channel Upper", "5m Channel Lower",
        "5m FakeOut Up", "5m FakeOut Down", "5m Trend", "5m Candle Time",
        "Data Status",
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

    def _calculate(self, frame, interval, live=True):
        if frame is None or len(frame) < self.MIRROR_LENGTH:
            return [None] * 11

        high = frame["high"].astype(float)
        open_price = frame["open"].astype(float)
        sma = high.rolling(self.MIRROR_LENGTH).mean()
        # TradingView ta.stdev() defaults to biased/population deviation.
        deviation = high.rolling(self.MIRROR_LENGTH).std(ddof=0)
        real_upper = sma - (self.MIRROR_MULTIPLIER * deviation)
        mirror_upper = (2.0 * open_price) - real_upper

        atr = self._wilder_atr(frame, self.CHANNEL_ATR_LENGTH)
        channel_atr = atr.rolling(self.CHANNEL_HIGHEST_LENGTH).max() * self.CHANNEL_WIDTH

        step_values = []
        upper_values = []
        lower_values = []
        for index in range(len(frame)):
            previous_step = step_values[index - 1] if index else 0.0
            previous_previous_step = step_values[index - 2] if index > 1 else 0.0
            width = channel_atr.iloc[index]
            if pd.isna(width):
                upper = lower = float("nan")
            else:
                upper = previous_step + float(width)
                lower = previous_step - float(width)

            source = float(high.iloc[index])
            # The source Pine script changes its step only on confirmed bars.
            # Historical candles are confirmed; the last candle is provisional.
            confirmed = not (live and index == len(frame) - 1)
            if index == self.CHANNEL_INITIAL_BAR:
                step = source
            elif confirmed and not pd.isna(upper) and (
                source > upper or source < lower
            ):
                step = source
            else:
                step = previous_step

            step_values.append(step)
            upper_values.append(upper)
            lower_values.append(lower)

        current = len(frame) - 1
        previous = current - 1
        if previous < 0:
            return [None] * 10

        current_high = float(high.iloc[current])
        previous_high = float(high.iloc[previous])
        current_low = float(frame["low"].iloc[current])
        previous_low = float(frame["low"].iloc[previous])
        up_now = upper_values[current]
        up_prev = upper_values[previous]
        low_now = lower_values[current]
        low_prev = lower_values[previous]

        fakeout_up = (
            not pd.isna(up_now) and not pd.isna(up_prev)
            and previous_high >= up_prev and current_high < up_now
            and step_values[current - 1] == step_values[current - 2]
            if current > 1 else False
        )
        fakeout_down = (
            not pd.isna(low_now) and not pd.isna(low_prev)
            and previous_low <= low_prev and current_low > low_now
            and step_values[current - 1] == step_values[current - 2]
            if current > 1 else False
        )

        previous_step = step_values[previous]
        current_step = step_values[current]
        trend = "UP" if current_step > previous_step else (
            "DOWN" if current_step < previous_step else "FLAT"
        )
        candle_time = pd.Timestamp(frame["datetime"].iloc[current]).strftime(
            "%Y-%m-%d %H:%M"
        )

        def rounded(series):
            value = series.iloc[current]
            return round(float(value), 4) if pd.notna(value) else None

        return [
            rounded(mirror_upper), rounded(real_upper),
            rounded(sma), rounded(deviation),
            round(current_step, 4),
            round(float(up_now), 4) if not pd.isna(up_now) else None,
            round(float(low_now), 4) if not pd.isna(low_now) else None,
            "YES" if fakeout_up else "",
            "YES" if fakeout_down else "",
            trend,
            candle_time,
        ]

    def update_symbol(self, api, symbol, token_key, quote):
        exchange, token = token_key.split("|", 1)
        ltp = self._number(quote.get("lp") if isinstance(quote, dict) else None)
        if ltp is None:
            return None, None, "WAITING FOR LIVE TICK"

        results = {}
        for interval in self.SUPPORTED_INTERVALS:
            key = f"{symbol}|{token}"
            history = self._get_history(api, exchange, token, interval, key)
            if history is None:
                results[interval] = [None] * 11
                continue
            live_frame = self._merge_live_candle(
                history, ltp, datetime.now(), interval
            )
            self.history[(key, interval)] = live_frame
            results[interval] = self._calculate(live_frame, interval, live=True)

        status = "LIVE" if any(value[0] is not None for value in results.values()) else "LOADING HISTORY"
        self.latest[symbol] = (ltp, results, status)
        return results.get(1), results.get(5), status

    def _get_or_create_sheet(self, workbook):
        names = [sheet.name for sheet in workbook.sheets]
        if "INDICATOR" in names:
            return workbook.sheets["INDICATOR"], False
        sheet = workbook.sheets.add("INDICATOR")
        print("Created INDICATOR sheet for live 1-minute and 5-minute indicators.")
        try:
            source = next(
                (item for item in workbook.sheets
                 if item.name in ("LIVE_TRADE", "Trade_Terminal")),
                None,
            )
            if source is not None:
                sheet.range("A4:A1000").value = source.range("A4:A1000").value
        except Exception as exc:
            print(f"Could not seed INDICATOR symbols from LIVE_TRADE: {exc}")
        return sheet, True

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
                sheet.range("A3:Z3").value = [self.HEADERS]
                sheet.range("A3:Z3").api.Font.Bold = True
                sheet.range("A3:Z3").api.WrapText = True
                self.sheet_ready = True

            symbols = sheet.range("A4:A1000").value or []
            last_index = max(
                (index for index, value in enumerate(symbols)
                 if str(value or "").strip()),
                default=-1,
            )
            last_row = 4 + last_index
            if last_index < 0:
                if self.last_written_row > 3:
                    sheet.range(f"A4:Z{self.last_written_row}").clear_contents()
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
                    output.append([symbol, None, None] + [None] * 22 + ["WAITING FOR SUBSCRIPTION"])
                    continue

                one_minute, five_minute, status = self.update_symbol(
                    api, symbol, token_key, quote
                )
                ltp = self._number(quote.get("lp"))
                output.append(
                    [symbol, ltp, datetime.now().strftime("%H:%M:%S")]
                    + (one_minute or [None] * 11)
                    + (five_minute or [None] * 11)
                    + [status]
                )

            sheet.range(f"A4:Z{last_row}").value = output
            if self.last_written_row > last_row:
                sheet.range(f"A{last_row + 1}:Z{self.last_written_row}").clear_contents()
            self.last_written_row = last_row
        except Exception as exc:
            print(f"INDICATOR sheet refresh failed: {exc}")
