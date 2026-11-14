"""Daily TradingView-style indicators integrated into the Finvasia terminal.

Historical candle downloads run in background workers so the Excel live-data loop
and Shoonya WebSocket callback are not held up by history requests.

Outputs are written to Trade_Terminal!AD:AK:
RSI, StochRSI, SMA Stoch, BB Upper, BB Middle, BB Lower, BUY Signal, SELL Signal.
"""
import json
import time
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor

import pandas as pd


class IndicatorEngine:
    def __init__(
        self, load_days=500, retry_seconds=300, length_rsi=14,
        length_stoch=14, length_sma_stoch=3, length_bb=20,
        bb_std=2.0, stoch_upper=70, stoch_lower=30
    ):
        self.load_days = load_days
        self.retry_seconds = retry_seconds
        self.history = {}
        self.retry_at = {}
        self.pending = {}
        self.executor = ThreadPoolExecutor(max_workers=2, thread_name_prefix="ShoonyaHistory")
        self.last_ltp = {}
        self.last_result = {}
        self.length_rsi = length_rsi
        self.length_stoch = length_stoch
        self.length_sma_stoch = length_sma_stoch
        self.length_bb = length_bb
        self.bb_std = bb_std
        self.stoch_upper = stoch_upper
        self.stoch_lower = stoch_lower

    @staticmethod
    def _empty():
        return [None, None, None, None, None, None, "", ""]

    def _fetch_history(self, api, exchange, symbol, key):
        now = time.time()
        try:
            end_date = datetime.now()
            start_date = end_date - timedelta(days=self.load_days)
            raw = api.get_daily_price_series(
                exchange=exchange,
                tradingsymbol=symbol,
                startdate=str(int(start_date.timestamp())),
                enddate=str(int(end_date.timestamp()))
            )
            rows = []
            for item in raw or []:
                if isinstance(item, str):
                    try:
                        item = json.loads(item)
                    except Exception:
                        continue
                if isinstance(item, dict):
                    rows.append(item)
            if isinstance(raw, dict):
                message = raw.get("emsg") or raw.get("stat") or str(raw)
                print(f"Daily indicator history rejected for {exchange}:{symbol}: {message}")
                self.retry_at[key] = now + self.retry_seconds
                return None
            if not rows:
                print(f"Daily indicator history returned no candles for {exchange}:{symbol}")
                self.retry_at[key] = now + self.retry_seconds
                return None

            frame = pd.DataFrame(rows)
            frame.rename(columns={
                "time": "datetime", "into": "open", "inth": "high",
                "intl": "low", "intc": "close", "intv": "volume", "ssboe": "epoch"
            }, inplace=True)

            # Shoonya daily history may return DD-MMM-YYYY, DD/MM/YYYY
            # with a time, or an epoch-seconds field depending on SDK/API version.
            if "datetime" not in frame and "epoch" in frame:
                frame["datetime"] = pd.to_numeric(frame["epoch"], errors="coerce")
                frame["datetime"] = pd.to_datetime(frame["datetime"], unit="s", errors="coerce")
            elif "datetime" in frame:
                parsed = frame["datetime"].map(
                    lambda value: pd.to_datetime(value, dayfirst=True, errors="coerce")
                )
                if "epoch" in frame:
                    epoch_values = pd.to_numeric(frame["epoch"], errors="coerce")
                    epoch_dates = pd.to_datetime(epoch_values, unit="s", errors="coerce")
                    parsed = parsed.fillna(epoch_dates)
                frame["datetime"] = parsed

            if "datetime" not in frame or "close" not in frame:
                print(f"Daily indicator history has unexpected fields for {exchange}:{symbol}: {list(frame.columns)}")
                self.retry_at[key] = now + self.retry_seconds
                return None

            frame["close"] = pd.to_numeric(frame["close"], errors="coerce")
            frame = frame.dropna(subset=["datetime", "close"])
            frame = frame.sort_values("datetime").drop_duplicates("datetime").reset_index(drop=True)
            if len(frame) < self.length_bb + 2:
                print(
                    f"Daily indicator history too short for {exchange}:{symbol}: "
                    f"{len(frame)} candles; need at least {self.length_bb + 2}"
                )
                self.retry_at[key] = now + self.retry_seconds
                return None
            self.retry_at[key] = now + self.retry_seconds
            print(
                f"Daily indicator history loaded for {exchange}:{symbol}: "
                f"{len(frame)} candles ({frame['datetime'].iloc[0].date()} to "
                f"{frame['datetime'].iloc[-1].date()})"
            )
            return frame
        except Exception as exc:
            print(f"Daily indicator history unavailable for {exchange}:{symbol}: {exc}")
            self.retry_at[key] = now + self.retry_seconds
            return None

    def _request_history(self, api, exchange, symbol, key):
        if key in self.pending or time.time() < self.retry_at.get(key, 0):
            return
        self.pending[key] = self.executor.submit(
            self._fetch_history, api, exchange, symbol, key
        )

    def _get_history_without_waiting(self, api, exchange, symbol, key):
        frame = self.history.get(key)
        if frame is not None:
            return frame

        future = self.pending.get(key)
        if future is None:
            self._request_history(api, exchange, symbol, key)
            return None
        if not future.done():
            return None

        self.pending.pop(key, None)
        try:
            frame = future.result()
        except Exception as exc:
            print(f"Indicator history worker failed for {exchange}:{symbol}: {exc}")
            frame = None
        if frame is not None:
            self.history[key] = frame
        return frame

    def _calculate(self, closes):
        close = pd.Series(closes, dtype="float64")
        delta = close.diff()
        gain = delta.where(delta > 0, 0.0)
        loss = -delta.where(delta < 0, 0.0)
        alpha = 1.0 / self.length_rsi
        avg_gain = gain.ewm(alpha=alpha, adjust=False, ignore_na=True).mean()
        avg_loss = loss.ewm(alpha=alpha, adjust=False, ignore_na=True).mean()
        rs = avg_gain / avg_loss
        rsi = (100 - (100 / (1 + rs))).fillna(50.0)

        lowest = rsi.rolling(self.length_stoch).min()
        highest = rsi.rolling(self.length_stoch).max()
        stoch = (100 * (rsi - lowest) / (highest - lowest)).fillna(50.0)
        sma_stoch = stoch.rolling(self.length_sma_stoch).mean().fillna(50.0)

        middle = close.rolling(self.length_bb).mean()
        deviation = close.rolling(self.length_bb).std(ddof=0)
        upper = middle + self.bb_std * deviation
        lower = middle - self.bb_std * deviation
        return rsi, stoch, sma_stoch, upper, middle, lower

    def update(self, api, exchange, symbol, ltp):
        exchange = str(exchange or "").upper().strip()
        symbol = str(symbol or "").upper().strip()
        if exchange != "NSE" or not symbol:
            return self._empty()
        try:
            ltp = float(ltp)
        except (TypeError, ValueError):
            return self._empty()
        if ltp <= 0:
            return self._empty()

        key = f"{exchange}|{symbol}"
        if self.last_ltp.get(key) == ltp and key in self.last_result:
            return self.last_result[key]

        frame = self._get_history_without_waiting(api, exchange, symbol, key)
        if frame is None or frame.empty:
            return self._empty()

        today = datetime.now().date()
        latest_day = frame["datetime"].iloc[-1].date()
        if latest_day == today:
            frame.loc[frame.index[-1], "close"] = ltp
        else:
            # Add today's working candle once, then update its close on each tick.
            new_bar = frame.iloc[-1].copy()
            new_bar["datetime"] = pd.Timestamp(today)
            new_bar["close"] = ltp
            frame = pd.concat([frame, pd.DataFrame([new_bar])], ignore_index=True)
        self.history[key] = frame
        closes = frame["close"].astype(float).tolist()
        if len(closes) < self.length_bb + 2:
            return self._empty()

        rsi, stoch, sma_stoch, upper, middle, lower = self._calculate(closes)
        current = len(closes) - 1
        previous = current - 1
        if pd.isna(upper.iloc[current]) or pd.isna(lower.iloc[current]):
            return self._empty()

        buy_signal = (
            closes[previous] < float(lower.iloc[previous])
            and closes[current] > float(lower.iloc[current])
            and float(sma_stoch.iloc[previous]) < self.stoch_lower
        )
        sell_signal = (
            closes[previous] > float(upper.iloc[previous])
            and closes[current] < float(upper.iloc[current])
            and float(sma_stoch.iloc[previous]) > self.stoch_upper
        )

        result = [
            round(float(rsi.iloc[current]), 2),
            round(float(stoch.iloc[current]), 2),
            round(float(sma_stoch.iloc[current]), 2),
            round(float(upper.iloc[current]), 2),
            round(float(middle.iloc[current]), 2),
            round(float(lower.iloc[current]), 2),
            "BUY" if buy_signal else "",
            "SELL" if sell_signal else "",
        ]
        self.last_ltp[key] = ltp
        self.last_result[key] = result
        return result
