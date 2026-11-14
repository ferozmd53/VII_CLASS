"""Daily TradingView-style indicators integrated into the Finvasia terminal.

Historical candle downloads run in background workers so the Excel live-data loop
and Shoonya WebSocket callback are not held up by history requests.

Outputs are written to LIVE_TRADE!AG:AN for supported Shoonya exchanges:
NSE, NFO, BSE, BFO, CDS and MCX.
"""
import json
import time
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor

import pandas as pd
from Intraday_Indicators import CandleCache


def fetch_bse_sensex_history(start_date, end_date):
    """Fetch SENSEX daily OHLCV from Yahoo Finance when Shoonya history is incomplete."""
    import requests
    from datetime import timezone, timedelta

    # Yahoo Finance's chart endpoint publishes the BSE Sensex index as ^BSESN.
    # Use India-local date boundaries and try both Yahoo chart hosts.
    india_tz = timezone(timedelta(hours=5, minutes=30))
    start_local = datetime(
        start_date.year, start_date.month, start_date.day, tzinfo=india_tz
    )
    end_local = datetime(
        end_date.year, end_date.month, end_date.day, tzinfo=india_tz
    ) + timedelta(days=1)

    params = {
        "period1": int(start_local.timestamp()),
        "period2": int(end_local.timestamp()),
        "interval": "1d",
        "events": "history",
    }
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36",
        "Accept": "application/json",
    }
    errors = []
    payload = None
    for host in ("query1.finance.yahoo.com", "query2.finance.yahoo.com"):
        url = f"https://{host}/v8/finance/chart/%5EBSESN"
        try:
            response = requests.get(url, params=params, headers=headers, timeout=25)
            response.raise_for_status()
            data = response.json()
            chart = data.get("chart", {})
            if chart.get("error"):
                raise ValueError(str(chart["error"]))
            results = chart.get("result") or []
            if not results:
                raise ValueError("Yahoo Finance returned no chart result")
            payload = results[0]
            break
        except Exception as exc:
            errors.append(f"{host}: {exc}")

    if payload is None:
        raise RuntimeError("Yahoo Finance SENSEX history failed: " + " | ".join(errors))

    timestamps = payload.get("timestamp") or []
    quote_rows = (payload.get("indicators", {}).get("quote") or [{}])[0]
    if not timestamps:
        raise ValueError("Yahoo Finance returned no SENSEX daily timestamps")

    def values(name):
        return quote_rows.get(name) or [None] * len(timestamps)

    frame = pd.DataFrame({
        "datetime": pd.to_datetime(timestamps, unit="s", utc=True)
                      .tz_convert("Asia/Kolkata").tz_localize(None).normalize(),
        "open": values("open"),
        "high": values("high"),
        "low": values("low"),
        "close": values("close"),
        "volume": values("volume"),
    })
    for column in ("open", "high", "low", "close", "volume"):
        frame[column] = pd.to_numeric(frame[column], errors="coerce")
    frame = frame.dropna(subset=["datetime", "open", "high", "low", "close"])
    frame = frame.sort_values("datetime").drop_duplicates("datetime").reset_index(drop=True)
    if frame.empty:
        raise ValueError("Yahoo Finance response contained no valid SENSEX OHLC candles")
    return frame


class IndicatorEngine:
    SUPPORTED_EXCHANGES = {"NSE", "NFO", "BSE", "BFO", "CDS", "MCX"}

    def __init__(
        self, load_days=500, retry_seconds=300, length_rsi=14,
        length_stoch=14, length_sma_stoch=3, length_bb=20,
        bb_std=2.0, stoch_upper=70, stoch_lower=30
    ):
        self.candle_cache = CandleCache()
        print(f"Daily candle cache: {self.candle_cache.database_path}")
        self.load_days = load_days
        self.retry_seconds = retry_seconds
        self.history = {}
        self.retry_at = {}
        self.pending = {}
        self.executor = ThreadPoolExecutor(max_workers=2, thread_name_prefix="ShoonyaHistory")
        self.last_ltp = {}
        self.refresh_at = {}
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
            # Keep the configured historical range for SENSEX too; if Shoonya
            # returns an incomplete index series, use BSE's official index archive. 
            history_days = self.load_days
            # Use complete local calendar-day boundaries for EOD history.
            end_date = end_date.replace(hour=23, minute=59, second=59, microsecond=0)
            start_date = (end_date - timedelta(days=history_days)).replace(
                hour=0, minute=0, second=0, microsecond=0
            )
            cached = self.candle_cache.load(exchange, symbol, "1d")
            if cached is not None and not cached.empty:
                cached_latest = pd.to_datetime(cached["datetime"], errors="coerce").max()
                start_date = max(start_date, (cached_latest - timedelta(days=2)).to_pydatetime())
            if exchange.upper() == "BSE" and symbol.strip().upper() == "SENSEX":
                print(
                    f"BSE:SENSEX indicator history request: symbol={symbol}, "
                    f"exchange={exchange}, days={history_days}, "
                    f"from={start_date.strftime('%Y-%m-%d %H:%M:%S')}, "
                    f"to={end_date.strftime('%Y-%m-%d %H:%M:%S')}"
                )
            # Shoonya identifies the NSE spot index by its broker symbol
            # "Nifty 50", while the workbook alias remains NSE:NIFTY INDEX.
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
                    f"BSE:SENSEX indicator raw history response: "
                    f"type={type(raw).__name__}, count={raw_count}, "
                    f"sample={str(raw[:2] if isinstance(raw, (list, tuple)) else raw)[:1200]}"
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
                frame = None
            elif not rows:
                print(f"Daily indicator history returned no candles for {exchange}:{symbol}")
                frame = None
            else:
                frame = pd.DataFrame(rows)

            if exchange.upper() == "BSE" and symbol.strip().upper() == "SENSEX":
                parsed_count = len(frame) if frame is not None else 0
                if parsed_count < self.length_bb + 2:
                    try:
                        print(
                            f"BSE:SENSEX Shoonya history incomplete ({parsed_count} rows); "
                            "trying BSE index archive fallback"
                        )
                        frame = fetch_bse_sensex_history(start_date, end_date)
                        print(
                            f"BSE:SENSEX BSE archive history loaded: {len(frame)} candles "
                            f"({frame['datetime'].iloc[0].date()} to {frame['datetime'].iloc[-1].date()})"
                        )
                    except Exception as fallback_exc:
                        print(f"BSE:SENSEX BSE archive fallback failed: {fallback_exc}")
                        if frame is None:
                            self.retry_at[key] = now + self.retry_seconds
                            return None

            if frame is None:
                self.retry_at[key] = now + self.retry_seconds
                return None
            frame.rename(columns={
                "time": "datetime", "into": "open", "inth": "high",
                "intl": "low", "intc": "close", "intv": "volume", "ssboe": "epoch"
            }, inplace=True)

            # Shoonya daily history may return DD-MMM-YYYY, DD/MM/YYYY
            # with a time, or an epoch-seconds field depending on SDK/API version.
            if "datetime" not in frame and "epoch" in frame:
                epoch_values = pd.to_numeric(frame["epoch"], errors="coerce")
                frame["datetime"] = pd.to_datetime(
                    epoch_values, unit="s", utc=True, errors="coerce"
                ).dt.tz_convert("Asia/Kolkata").dt.tz_localize(None)
            elif "datetime" in frame:
                parsed = frame["datetime"].map(
                    lambda value: pd.to_datetime(value, dayfirst=True, errors="coerce")
                )
                if "epoch" in frame:
                    epoch_values = pd.to_numeric(frame["epoch"], errors="coerce")
                    epoch_dates = pd.to_datetime(
                        epoch_values, unit="s", utc=True, errors="coerce"
                    ).dt.tz_convert("Asia/Kolkata").dt.tz_localize(None)
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
        if frame is None:
            try:
                frame = self.candle_cache.load(exchange, symbol, "1d")
                if frame is not None and not frame.empty:
                    frame["datetime"] = pd.to_datetime(frame["datetime"], errors="coerce")
                    self.history[key] = frame
                    print(f"Daily SQLite cache loaded for {exchange}:{symbol}: {len(frame)} candles")
            except Exception as cache_error:
                print(f"Daily SQLite cache read failed for {exchange}:{symbol}: {cache_error}")
        future = self.pending.get(key)

        if future is not None and future.done():
            self.pending.pop(key, None)
            try:
                refreshed = future.result()
                if refreshed is not None and not refreshed.empty:
                    frame = pd.concat([frame, refreshed], ignore_index=True) if frame is not None else refreshed
                    frame = frame.sort_values("datetime").drop_duplicates("datetime", keep="last").reset_index(drop=True)
                    self.candle_cache.save(exchange, symbol, "1d", frame)
                    self.history[key] = frame
                    self.refresh_at[key] = time.time() + 60
            except Exception as exc:
                print(f"Indicator history worker failed for {exchange}:{symbol}: {exc}")
                self.refresh_at[key] = time.time() + self.retry_seconds

        if frame is None:
            if future is None:
                self._request_history(api, exchange, symbol, key)
                return None
            return None

        # Refresh broker daily history periodically so a forming daily candle
        # is updated from the broker, without substituting local time or LTP.
        if key not in self.pending and time.time() >= self.refresh_at.get(key, 0):
            self.refresh_at[key] = time.time() + 60
            self.pending[key] = self.executor.submit(
                self._fetch_history, api, exchange, symbol, key
            )
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

    def update(self, api, exchange, symbol, ltp=None):
        """Calculate daily indicators from broker daily candles only.

        The WebSocket LTP and the PC clock are deliberately not used to
        overwrite or fabricate daily OHLC rows. This keeps daily indicators
        aligned with the broker's actual candle timestamps.
        """
        exchange = str(exchange or "").upper().strip()
        symbol = str(symbol or "").upper().strip()
        if exchange not in self.SUPPORTED_EXCHANGES or not symbol:
            return self._empty()

        key = f"{exchange}|{symbol}"
        frame = self._get_history_without_waiting(api, exchange, symbol, key)
        if frame is None or frame.empty:
            return self._empty()

        frame = frame.copy()
        frame["datetime"] = pd.to_datetime(frame["datetime"], errors="coerce")
        frame["close"] = pd.to_numeric(frame["close"], errors="coerce")
        frame = frame.dropna(subset=["datetime", "close"])
        frame = frame.sort_values("datetime").drop_duplicates("datetime", keep="last")
        if len(frame) < self.length_bb + 2:
            return self._empty()

        # Cache by the actual newest broker candle, not by LTP. The same
        # candle can be recalculated when the broker updates its OHLC values.
        latest = frame.iloc[-1]
        candle_key = (
            str(latest["datetime"]), float(latest["close"]),
            float(latest["open"]) if "open" in frame and pd.notna(latest.get("open")) else None,
            float(latest["high"]) if "high" in frame and pd.notna(latest.get("high")) else None,
            float(latest["low"]) if "low" in frame and pd.notna(latest.get("low")) else None,
        )
        if self.last_ltp.get(key) == candle_key and key in self.last_result:
            return self.last_result[key]

        closes = frame["close"].astype(float).tolist()
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
        self.last_ltp[key] = candle_key
        self.last_result[key] = result
        return result
