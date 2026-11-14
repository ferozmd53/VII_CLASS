# -*- coding: utf-8 -*-
"""
Created on Fri Oct 14 09:13:23 2022

Disclaimer: The information provided by the Python Traders channel is for educational purposes only, so please contact your financial adviser before placing your trade. Developer is not responsible for any profit/loss happened due to coding, logical or any type of error.

Toots Feature :
1. Trade terminal places live broker orders only (REAL mode).
2. You can set target/sl together using tool.
3. You can also trail your sl.
4. Tool has voice and telegram notification/alert feature.
5. Tools have orderbook sheet where you can cancel any open orders.
6. Tools have openposition sheet where you can exit any/all running trade.
7. Both option chains show live market data and implied volatility (IV); Greek columns are removed.
"""


import warnings
warnings.filterwarnings("ignore")

from Shoonya_OAuth import login_with_excel
from get_auth import get_auth_code
from Shoonya_Indicators import IndicatorEngine
from Weekly import WeeklyDataEngine

weekly_data_engine = WeeklyDataEngine()
#import Noren
import os
import time
import json
import queue
import sys
import platform
from datetime import datetime as dt, timedelta, time, date
from time import sleep, monotonic as _monotonic
import logging
import socket
from threading import Thread, Lock, local, Event
from concurrent.futures import ThreadPoolExecutor

voice_lock = Lock()
LIVE_DATA_LOCK = Lock()
LIVE_TRADE_SHEET_LOCK = Lock()
SUBSCRIPTION_DIAGNOSTIC_LOCK = Lock()
SUBSCRIPTION_DIAGNOSTIC_TOKENS = set()
SUBSCRIPTION_CALLBACK_LOGGED = set()  # valid positive-LTP callbacks
SUBSCRIPTION_CALLBACK_SEEN = set()    # any quote/ack callback for a tracked token
EXCEL_RECONNECT_LOCK = Lock()
EXCEL_RECONNECT_EVENT_LOCK = Lock()
EXCEL_RECONNECT_GENERATION = 0
_LAST_EXCEL_RECONNECT_SIGNAL = 0.0
_LAST_EXCEL_RELAUNCH = 0.0
EXCEL_RELAUNCH_COOLDOWN_SEC = 30
SYSTEM_INTERNET_OK = None
SYSTEM_EXCEL_OK = None
BROKER_LOGIN_OK = False
import numpy as np

# NumPy 2.0 removed the legacy positive/negative infinity aliases.
# The upstream GetIVGreeks option-chain module still references both.
# Restore them before importing/loading that module for compatibility.
if not hasattr(np, "PINF"):
    np.PINF = np.inf
if not hasattr(np, "NINF"):
    np.NINF = -np.inf

try:
    import pandas as pd
except (ModuleNotFoundError, ImportError):
    print("pandas module not found")
    os.system(f"{sys.executable} -m pip install -U pandas")
finally:
    import pandas as pd
        
try:
    import pyotp
except (ModuleNotFoundError, ImportError):
    print("pyotp module not found")
    os.system(f"{sys.executable} -m pip install -U pyotp")
finally:
    import pyotp
    
try:
    import xlwings as xw
except (ModuleNotFoundError, ImportError):
    print("xlwings module not found")
    os.system(f"{sys.executable} -m pip install -U xlwings")
finally:
    import xlwings as xw

try:
    import pyttsx3
except (ModuleNotFoundError, ImportError):
    os.system(f"{sys.executable} -m pip install -U pyttsx3")
finally:
    import pyttsx3

try:
    import requests
except (ModuleNotFoundError, ImportError):
    print("requests module not found")
    os.system(f"{sys.executable} -m pip install -U requests")
finally:
    import requests
    
try:
    from GetIVGreeks import DayCountType, ExpType, TryMatchWith, CalcIvGreeks
except ImportError:
    exec(
        __import__("requests")
        .get(
            "https://gist.githubusercontent.com/ShabbirHasan1"
            + "/7695687d87053c7e3df46810ab2e3046"
            + "/raw/30551841a3fd3641c8416266fd27c5254d2150b3"
            + "/GetIVGreeks.py"
        )
        .text
    )

print("HI welcome")
    
df_ins_NFO = pd.DataFrame()

OptionChain_template = []
subs_lst = []
subs_pending_lst = []

Indices_To_check_instrument_Sheet = ["NIFTY", "BANKNIFTY", "SENSEX", "SENSEX50", "BANKEX"]
IndexList = ["NIFTY", "BANKNIFTY", "FINNIFTY","INDIAVIX","MIDCPNIFTY", "SENSEX", "SENSEX50", "BANKEX"]
Token_list = {'NIFTY':26000,'MIDCPNIFTY':26074,'BANKNIFTY':26009,'FINNIFTY':26037,'INDIAVIX':26017, 'SENSEX':1, 'SENSEX50':47, 'BANKEX':12}

try:
    TerminalSheetName = sys.argv[1]
    if not os.path.exists(TerminalSheetName):
        TerminalSheetName = os.path.join(os.path.dirname(os.path.abspath(__file__)), "VII_CLASS.xlsm")
except Exception:
    TerminalSheetName = os.path.join(os.path.dirname(os.path.abspath(__file__)), "VII_CLASS.xlsm")

# Excel COM proxies are apartment/thread-affine. Never share an xlwings
# Book/Sheet proxy between Python threads: each worker must bind its own proxy.
_excel_com_state = local()

def ensure_excel_com_initialized():
    """Initialize COM once in each Python worker thread on Windows."""
    if os.name != "nt" or getattr(_excel_com_state, "initialized", False):
        return
    try:
        import pythoncom
        pythoncom.CoInitialize()
        _excel_com_state.initialized = True
    except Exception:
        pass

def _extract_symbol(value):
    """Safely extract a symbol string from xlwings range values, which may be 1D or 2D lists."""
    if isinstance(value, list):
        val = value[0] if value else None
    else:
        val = value
    if val is None:
        return ""
    return str(val).strip().upper()

def normalize_expiry_input(value):
    """Convert Excel date, datetime, text or serial values to a Python date."""
    if value is None or (isinstance(value, str) and not value.strip()):
        return None
    if isinstance(value, dt):
        parsed_date = value.date()
        if parsed_date <= date(1899, 12, 30):
            return None
        return parsed_date
    if isinstance(value, date):
        if value <= date(1899, 12, 30):
            return None
        return value
    if isinstance(value, (int, float, np.integer, np.floating)):
        try:
            serial = float(value)
            if not np.isfinite(serial) or serial <= 0:
                return None
            return pd.to_datetime(serial, unit="D", origin="1899-12-30").date()
        except Exception:
            return None
    if isinstance(value, str):
        value = value.strip()
        for date_format in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%b-%Y", "%d %b %Y"):
            try:
                return dt.strptime(value, date_format).date()
            except ValueError:
                continue
        try:
            parsed = pd.to_datetime(value, dayfirst=True, errors="coerce")
            return None if pd.isna(parsed) else parsed.date()
        except Exception:
            return None
    return None

def resolve_option_expiry(df_instrument, symbol, requested_expiry):
    """Keep a valid requested expiry, otherwise choose the nearest upcoming expiry."""
    rows = df_instrument[
        (df_instrument.Symbol == symbol)
        & (df_instrument["OptionType"].isin(["CE", "PE"]))
    ]
    available = set()
    for value in rows["Expiry"].dropna().unique():
        value = value.date() if hasattr(value, "date") else value
        if isinstance(value, date) and value >= date.today():
            available.add(value)

    if requested_expiry in available:
        return requested_expiry
    return min(available) if available else None

def expiry_not_found_message(df_instrument, symbol, expiry_input):
    """Explain a rejected expiry: show the date entered and the valid upcoming expiries."""
    base = "Please enter correct expiry in dd-mm-YYYY (date format)"
    try:
        today = date.today()
        valid = set()
        rows = df_instrument[(df_instrument.Symbol == symbol) & (df_instrument['OptionType'].isin(['CE', 'PE']))]
        for value in rows['Expiry'].dropna().unique():
            value = value.date() if hasattr(value, "date") else value
            if value >= today:
                valid.add(value)
        upcoming = ", ".join(v.strftime("%d-%m-%Y") for v in sorted(valid)[:4])
        return f"{base}. {expiry_input.strftime('%d-%m-%Y')} has no contracts for {symbol}. Upcoming: {upcoming or 'none found'}"
    except Exception:
        return base

def is_excel_busy_error(error):
    """Return True for transient Excel/OLE automation-busy errors."""
    message = str(error).lower()
    return any(code in message for code in (
        "0x800ac472", "-2146777998",
        "0x80010001", "-2147418111",
        "0x8001010a", "-2147417846",
    ))

def is_excel_connection_lost():
    """Connection state belongs to the calling worker thread."""
    return bool(getattr(_excel_com_state, "connection_lost", False))

def _notify_excel_workbook_reconnected():
    """Signal one workbook-recovery event, coalescing worker-thread rebinds."""
    global EXCEL_RECONNECT_GENERATION, _LAST_EXCEL_RECONNECT_SIGNAL
    now = _monotonic()
    with EXCEL_RECONNECT_EVENT_LOCK:
        if now - _LAST_EXCEL_RECONNECT_SIGNAL >= 10.0:
            EXCEL_RECONNECT_GENERATION += 1
            _LAST_EXCEL_RECONNECT_SIGNAL = now

def get_terminal_workbook():
    """Get a thread-local workbook proxy attached to the already-open workbook."""
    ensure_excel_com_initialized()
    target_path = os.path.normcase(os.path.abspath(TerminalSheetName))
    workbook = getattr(_excel_com_state, "workbook", None)

    if workbook is not None:
        try:
            current_path = os.path.normcase(os.path.abspath(workbook.fullname))
            if current_path == target_path:
                _ = workbook.sheets[0].name
                return workbook
        except Exception as cached_error:
            if is_excel_busy_error(cached_error):
                raise
            _excel_com_state.workbook = None
            _excel_com_state.connection_lost = True
            if not getattr(_excel_com_state, "disconnect_reported", False):
                print("Excel COM connection lost in this worker. Waiting for a valid workbook connection...")
                _excel_com_state.disconnect_reported = True

    matches = []
    try:
        for app in xw.apps:
            try:
                for book in app.books:
                    try:
                        book_path = os.path.normcase(os.path.abspath(book.fullname))
                        if book_path != target_path:
                            continue
                        _ = book.sheets[0].name
                        matches.append(book)
                    except Exception as book_error:
                        if is_excel_busy_error(book_error):
                            raise
                        continue
            except Exception as app_error:
                if is_excel_busy_error(app_error):
                    raise
                continue
    except Exception as apps_error:
        if is_excel_busy_error(apps_error):
            raise
        pass

    if matches:
        workbook = matches[0]
        was_disconnected = bool(getattr(_excel_com_state, "connection_lost", False))
        _excel_com_state.workbook = workbook
        _excel_com_state.connection_lost = False
        _excel_com_state.disconnect_reported = False
        if was_disconnected:
            print("Excel workbook connection restored for this worker.")
            _notify_excel_workbook_reconnected()
        return workbook

    if getattr(_excel_com_state, "connection_lost", False):
        with EXCEL_RECONNECT_LOCK:
            try:
                running_apps = list(xw.apps)
            except Exception as apps_error:
                if is_excel_busy_error(apps_error):
                    raise
                running_apps = []

            if running_apps:
                for app in running_apps:
                    try:
                        for book in app.books:
                            try:
                                book_path = os.path.normcase(os.path.abspath(book.fullname))
                                if book_path != target_path:
                                    continue
                                _ = book.sheets[0].name
                                _excel_com_state.workbook = book
                                _excel_com_state.connection_lost = False
                                _excel_com_state.disconnect_reported = False
                                _excel_com_state.reconnect_error_reported = False
                                print("Excel workbook connection restored for this worker.")
                                _notify_excel_workbook_reconnected()
                                return book
                            except Exception as book_error:
                                if is_excel_busy_error(book_error):
                                    raise
                                continue
                    except Exception as app_error:
                        if is_excel_busy_error(app_error):
                            raise
                        continue

                for app in running_apps:
                    try:
                        workbook = app.books.open(TerminalSheetName)
                        _ = workbook.sheets[0].name
                        _excel_com_state.workbook = workbook
                        _excel_com_state.connection_lost = False
                        _excel_com_state.disconnect_reported = False
                        _excel_com_state.reconnect_error_reported = False
                        print("Excel workbook connection restored for this worker.")
                        _notify_excel_workbook_reconnected()
                        return workbook
                    except Exception as reconnect_error:
                        if is_excel_busy_error(reconnect_error):
                            raise
                        _excel_com_state.workbook = None
                        last_reconnect_error = reconnect_error
                if not getattr(_excel_com_state, "reconnect_error_reported", False):
                    print(f"Excel reconnect attempt failed; will retry: {locals().get('last_reconnect_error', 'No active Excel instance')}")
                    _excel_com_state.reconnect_error_reported = True

            if _excel_com_state.workbook is None:
                global _LAST_EXCEL_RELAUNCH
                now = _monotonic()
                if _LAST_EXCEL_RELAUNCH == 0.0 or now - _LAST_EXCEL_RELAUNCH >= EXCEL_RELAUNCH_COOLDOWN_SEC:
                    _LAST_EXCEL_RELAUNCH = now
                    try:
                        print("No running Excel found. Relaunching Excel and opening the workbook...")
                        new_app = xw.App(visible=True, add_book=False)
                        workbook = new_app.books.open(TerminalSheetName)
                        _ = workbook.sheets[0].name
                        _excel_com_state.workbook = workbook
                        _excel_com_state.connection_lost = False
                        _excel_com_state.disconnect_reported = False
                        _excel_com_state.reconnect_error_reported = False
                        print("Excel workbook connection restored for this worker.")
                        _notify_excel_workbook_reconnected()
                        return workbook
                    except Exception as relaunch_error:
                        if is_excel_busy_error(relaunch_error):
                            raise
                        _excel_com_state.workbook = None
                        print(f"Excel relaunch failed; will retry in {EXCEL_RELAUNCH_COOLDOWN_SEC}s: {relaunch_error}")

    if not getattr(_excel_com_state, "connection_lost", False):
        try:
            workbook = xw.Book(TerminalSheetName)
            _ = workbook.sheets[0].name
            _excel_com_state.workbook = workbook
            _excel_com_state.connection_lost = False
            _excel_com_state.disconnect_reported = False
            return workbook
        except Exception as open_error:
            if is_excel_busy_error(open_error):
                raise
            _excel_com_state.workbook = None
            _excel_com_state.connection_lost = True
            if not getattr(_excel_com_state, "disconnect_reported", False):
                print(f"Could not attach/open {os.path.basename(TerminalSheetName)}: {open_error}")
                _excel_com_state.disconnect_reported = True
            raise RuntimeError(
                f"Excel workbook could not be opened: {os.path.basename(TerminalSheetName)}"
            ) from open_error

    _excel_com_state.workbook = None
    _excel_com_state.connection_lost = True
    if not getattr(_excel_com_state, "disconnect_reported", False):
        print("VII_CLASS.xlsm is not currently accessible. Waiting for it to be opened...")
        _excel_com_state.disconnect_reported = True
    raise RuntimeError("VII_CLASS.xlsm is not available in this worker's Excel session.")

def Text2Speech(Text):
    """Speak a message using one serialized Windows TTS engine."""
    global engine
    try:
        with voice_lock:
            engine = pyttsx3.init()
            voices = engine.getProperty('voices')
            try:
                if len(voices) > 1:
                    engine.setProperty('voice', voices[1].id)
            except Exception:
                pass
            engine.setProperty('rate', 160)
            engine.say(str(Text))
            engine.runAndWait()
            del engine
    except Exception as e:
        print(f"Voice alert failed: {e}")
    
Text2Speech("Welcome to python trader trade terminal tool, have a nice day. Hope u r enjoying our work")

def internet_is_available():
    """Check internet reachability independently of the broker WebSocket."""
    for host in ("1.1.1.1", "8.8.8.8"):
        try:
            connection = socket.create_connection((host, 53), timeout=2)
            connection.close()
            return True
        except OSError:
            continue
    return False

def system_health_monitor():
    """Speak on internet or Excel workbook disconnect/recovery transitions."""
    global SYSTEM_INTERNET_OK, SYSTEM_EXCEL_OK, BROKER_LOGIN_OK
    previous_internet = None
    previous_excel = None
    broker_check_counter = 0
    while True:
        internet_ok = internet_is_available()
        SYSTEM_INTERNET_OK = internet_ok
        if internet_ok != previous_internet:
            if not internet_ok:
                Text2Speech("Warning. Internet connection lost. Live orders are blocked.")
            elif previous_internet is False:
                Text2Speech("Internet connection restored.")
            previous_internet = internet_ok

        excel_ok = False
        try:
            workbook = get_terminal_workbook()
            _ = workbook.sheets["LOGIN"].range("B3").value
            excel_ok = True
        except Exception:
            excel_ok = False
        SYSTEM_EXCEL_OK = excel_ok
        if excel_ok != previous_excel:
            if not excel_ok:
                Text2Speech("Warning. Excel workbook is closed or not responding. Live orders are blocked.")
            elif previous_excel is False:
                Text2Speech("Excel workbook connection restored.")
            previous_excel = excel_ok

        broker_check_counter += 1
        if BROKER_LOGIN_OK and internet_ok and broker_check_counter >= 4:
            broker_check_counter = 0
            try:
                broker_status = api.get_limits()
                broker_ok = isinstance(broker_status, dict) and str(broker_status.get("stat", "")).lower() == "ok"
            except Exception:
                broker_ok = False
            if not broker_ok:
                BROKER_LOGIN_OK = False
                Text2Speech("Warning. Shoonya broker session is no longer valid. Live orders are blocked.")

        sleep(8)

def get_login_settings_sheet(workbook):
    """Use LOGIN as the single settings sheet and migrate legacy sheets once."""
    sheet_names = [sheet.name for sheet in workbook.sheets]
    login_sheet = workbook.sheets["LOGIN"] if "LOGIN" in sheet_names else workbook.sheets.add("LOGIN")

    if str(login_sheet.range("C1").value or "") != "LOGIN_SETTINGS_MIGRATED_V1":
        if "Config" in sheet_names:
            try:
                workbook.sheets["Config"].used_range.copy(destination=login_sheet.range("D1"))
                print("Migrated Config sheet settings into LOGIN (starting at column D).")
            except Exception as exc:
                raise RuntimeError(f"Could not migrate Config sheet into LOGIN: {exc}")
        if "User_Credential" in sheet_names:
            legacy = workbook.sheets["User_Credential"]
            if not login_sheet.range("AZ2").value:
                login_sheet.range("AZ2").value = legacy.range("B10").value
            if not login_sheet.range("AZ3").value:
                login_sheet.range("AZ3").value = legacy.range("B11").value
        login_sheet.range("C1").value = "LOGIN_SETTINGS_MIGRATED_V1"
        workbook.save()

    # Account-specific and order-pricing settings live in the workbook, not source code.
    login_sheet.range("A15").value = "Auto Algo ID (User ID + _U)"
    login_sheet.range("A24").value = "Market Safety (%)"
    if login_sheet.range("B24").value in (None, ""):
        login_sheet.range("B24").value = 1.0

    return login_sheet


def get_shoonya_algo_id():
    """Build Shoonya Algo ID automatically from LOGIN!B3 (User ID) plus _U."""
    login_sheet = get_login_settings_sheet(get_terminal_workbook())
    user_id = str(login_sheet.range("B3").value or "").strip()
    if not user_id:
        raise RuntimeError(
            "Shoonya User ID is missing. Enter your User ID in LOGIN!B3."
        )
    # Avoid producing USERID_U_U if the suffix is already present.
    algo_id = user_id if user_id.upper().endswith("_U") else user_id + "_U"
    login_sheet.range("B15").value = algo_id
    return algo_id


def get_market_safety_percent():
    """Read and validate the SL-LMT safety buffer from LOGIN!B24."""
    login_sheet = get_login_settings_sheet(get_terminal_workbook())
    raw_value = login_sheet.range("B24").value
    try:
        value = float(raw_value) if raw_value not in (None, "") else 1.0
    except (TypeError, ValueError):
        raise RuntimeError("LOGIN!B24 Market Safety must be a number from 0 to 10.")
    if not np.isfinite(value) or value < 0 or value > 10:
        raise RuntimeError("LOGIN!B24 Market Safety must be between 0 and 10.")
    login_sheet.range("B24").value = value
    return value


def _shoonya_login_once():
    global api, TelegramBotCredential, ReceiverTelegramID, BROKER_LOGIN_OK
    BROKER_LOGIN_OK = False
    global logger, userid, client_name
    client_name = ""
    isConnected = 0
    try:
        excel_name = get_terminal_workbook()
    except Exception as exc:
        if is_excel_busy_error(exc):
            raise
        print(f"Excel workbook could not be opened: {exc}")
        Text2Speech("Warning. Excel workbook is unavailable. Shoonya login cannot start.")
        return 0

    try:
        os.environ["SHOONYA_WORKBOOK_PATH"] = os.path.abspath(TerminalSheetName)
        login_sheet = get_login_settings_sheet(excel_name)

        userid = str(login_sheet.range("B3").value or "").strip()
        client_name = userid or "Shoonya User"

        timestamp = dt.now().strftime("%d%m%Y_%H%M%S")
        os.makedirs("Logs", exist_ok=True)
        log_file = os.path.join("Logs", f"Shoonya_TT_{userid or 'user'}_{timestamp}.log")
        logging.basicConfig(filename=log_file, format="%(asctime)s %(message)s", filemode="w")
        logger = logging.getLogger()
        logger.setLevel(logging.INFO)

        TelegramBotCredential = str(login_sheet.range("AZ2").value or "")
        ReceiverTelegramID = str(login_sheet.range("AZ3").value or "")
        if "." in ReceiverTelegramID:
            ReceiverTelegramID = ReceiverTelegramID.split(".")[0]

        print("Connecting to Shoonya using reusable OAuth session...")
        api = login_with_excel(excel_name, auth_code_provider=get_auth_code)
        if api is None:
            login_sheet.range("C2").value = "Shoonya OAuth login failed. Check LOGIN sheet."
            login_sheet.range("C2").color = (255, 0, 0)
            Text2Speech("Warning. Shoonya login failed. Please check the LOGIN sheet and internet connection.")
            return 0

        login_status = api.get_limits()
        if not isinstance(login_status, dict) or str(login_status.get("stat", "")).lower() != "ok":
            print(f"Shoonya session validation failed: {login_status}")
            login_sheet.range("C2").value = "Shoonya session validation failed."
            login_sheet.range("C2").color = (255, 0, 0)
            Text2Speech("Warning. Shoonya session validation failed. No live trading will start.")
            return 0

        userid = str(login_sheet.range("B3").value or "").strip()
        client_name = userid
        login_sheet.range("C2").value = "Login Successful, Welcome " + client_name + "\nOAuth session reused/saved in LOGIN and SESSION sheets."
        login_sheet.range("C2").color = (118, 224, 180)
        Text2Speech("Shoonya login successful. Live trading mode is enabled.")
        excel_name.save()
        BROKER_LOGIN_OK = True
        isConnected = 1
        return isConnected

    except Exception as exc:
        if is_excel_busy_error(exc):
            raise
        print(f"Shoonya OAuth login error: {exc}")
        Text2Speech("Warning. Shoonya login failed. No live trading will start.")
        try:
            login_sheet.range("C2").value = "Login failed: " + str(exc)
            login_sheet.range("C2").color = (255, 0, 0)
        except Exception:
            pass
        return 0

def Shoonya_login():
    """Retry transient Excel/OLE busy errors without requesting new credentials."""
    max_attempts = 5
    for attempt in range(1, max_attempts + 1):
        try:
            result = _shoonya_login_once()
            return result
        except Exception as exc:
            if not is_excel_busy_error(exc):
                print(f"Shoonya login could not recover from Excel error: {exc}")
                return 0
            if attempt >= max_attempts:
                print(
                    "Excel remained busy after "
                    f"{max_attempts} login attempts ({exc}). "
                    "Close Excel dialogs, wait for Excel to respond, then restart run.py."
                )
                return 0
            wait_seconds = min(2 * attempt, 8)
            print(
                f"Excel is temporarily busy during Shoonya login "
                f"(attempt {attempt}/{max_attempts}). Retrying in {wait_seconds} seconds..."
            )
            sleep(wait_seconds)

feed_opened = False
SYMBOLDICT = {}
live_data = {}
SOCKET_SUBSCRIPTIONS = set()
SOCKET_SESSION_SUBSCRIPTIONS = set()
SOCKET_SUBSCRIPTIONS_LOCK = Lock()
SOCKET_SUBSCRIBE_SEND_LOCK = Lock()
indicator_engine = IndicatorEngine()
indicator_settings_signature = None
indicator_settings_last_check = 0.0

# Indicator/history controls are stored on LOGIN!A16:B23.
INDICATOR_LOGIN_DEFAULTS = [
    ("Historical Days", 500),
    ("RSI Length", 14),
    ("StochRSI Length", 14),
    ("SMA Stoch Length", 3),
    ("Bollinger Length", 20),
    ("Bollinger StdDev", 2.0),
    ("Stoch Upper", 70),
    ("Stoch Lower", 30),
]

def configure_indicator_engine_from_login(workbook=None):
    """Read indicator controls from LOGIN; write defaults for blank cells."""
    global indicator_engine, indicator_settings_signature
    workbook = workbook or get_terminal_workbook()
    login_sheet = get_login_settings_sheet(workbook)
    values = {}
    for row, (label, default) in enumerate(INDICATOR_LOGIN_DEFAULTS, start=16):
        login_sheet.range(f"A{row}").value = label
        raw_value = login_sheet.range(f"B{row}").value
        try:
            value = float(raw_value) if raw_value not in (None, "") else float(default)
        except (TypeError, ValueError):
            value = float(default)
        if value <= 0:
            value = float(default)
        if label != "Bollinger StdDev":
            value = int(value)
        login_sheet.range(f"B{row}").value = value
        values[label] = value

    # Avoid invalid rolling windows and impossible StochRSI thresholds.
    values["Historical Days"] = max(30, min(5000, values["Historical Days"]))
    values["RSI Length"] = max(2, min(200, values["RSI Length"]))
    values["StochRSI Length"] = max(2, min(200, values["StochRSI Length"]))
    values["SMA Stoch Length"] = max(1, min(50, values["SMA Stoch Length"]))
    values["Bollinger Length"] = max(2, min(200, values["Bollinger Length"]))
    values["Bollinger StdDev"] = max(0.1, min(10.0, values["Bollinger StdDev"]))
    values["Stoch Upper"] = max(1, min(100, values["Stoch Upper"]))
    values["Stoch Lower"] = max(0, min(99, values["Stoch Lower"]))
    if values["Stoch Lower"] >= values["Stoch Upper"]:
        values["Stoch Lower"], values["Stoch Upper"] = 30, 70

    # Show the effective (validated) values back in Excel.
    for row, (label, _default) in enumerate(INDICATOR_LOGIN_DEFAULTS, start=16):
        login_sheet.range(f"B{row}").value = values[label]

    indicator_engine = IndicatorEngine(
        load_days=values["Historical Days"],
        length_rsi=values["RSI Length"],
        length_stoch=values["StochRSI Length"],
        length_sma_stoch=values["SMA Stoch Length"],
        length_bb=values["Bollinger Length"],
        bb_std=values["Bollinger StdDev"],
        stoch_upper=values["Stoch Upper"],
        stoch_lower=values["Stoch Lower"],
    )
    indicator_settings_signature = tuple(
        login_sheet.range(f"B{row}").value
        for row in range(16, 24)
    )
    workbook.save()
    print("Indicator settings loaded from LOGIN: " + ", ".join(
        f"{key}={value}" for key, value in values.items()
    ))


def reload_indicator_engine_if_settings_changed(workbook):
    """Hot-reload indicator settings edited in LOGIN without restarting."""
    global indicator_settings_last_check
    now = _monotonic()
    if now - indicator_settings_last_check < 1.0:
        return False
    indicator_settings_last_check = now

    try:
        login_sheet = get_login_settings_sheet(workbook)
        current_signature = tuple(
            login_sheet.range(f"B{row}").value
            for row in range(16, 24)
        )
        if current_signature == indicator_settings_signature:
            return False

        print("LOGIN indicator settings changed; reloading indicator engine and history...")
        configure_indicator_engine_from_login(workbook)

        # Remove old calculated values immediately. The new engine will refill
        # them as fresh daily history is loaded using the updated parameters.
        live_sheet = get_live_trade_sheet(workbook)
        live_sheet.range("AG4:AN1000").clear_contents()
        workbook.save()
        print("Indicator settings applied. Fresh calculations are loading.")
        return True
    except Exception as exc:
        print(f"Could not reload changed indicator settings: {exc}")
        return False


Telegram_Message = ["Welcome to Python Trader excel based trade terminal","Have a Good Day"]
print(Telegram_Message)
Voice_Message = []
TelegramBotCredential = None
ReceiverTelegramID = None

def convert_to_float(item):
    try:
        return float(item)
    except (ValueError, TypeError):
        return 0

def SendMessageToTelegram(Message):
    global TelegramBotCredential, ReceiverTelegramID
    try:
        Url = "https://api.telegram.org/bot" + str(TelegramBotCredential) +  "/sendMessage?chat_id=" + str(ReceiverTelegramID)
        textdata ={ "text":Message}
        response = requests.request("POST",Url,params=textdata)
    except Exception as e:
        pass

def _has_valid_live_ltp(quote):
    """A subscription is live only after a real, positive LTP is present."""
    if not isinstance(quote, dict):
        return False
    try:
        ltp = float(str(quote.get("lp", "")).replace(",", "").strip())
        return bool(np.isfinite(ltp) and ltp > 0)
    except (TypeError, ValueError, OverflowError):
        return False

def event_handler_quote_update(inmessage):
    global live_data, SYMBOLDICT

    if isinstance(inmessage, str):
        try:
            inmessage = json.loads(inmessage)
        except (TypeError, ValueError, json.JSONDecodeError):
            return
    if not isinstance(inmessage, dict):
        return

    fields = [
        "ts", "lp", "pc", "c", "o", "h", "l", "v", "ltq", "ltp",
        "bp1", "sp1", "bq1", "sq1", "ap", "oi", "poi", "toi",
    ]
    message = {field: inmessage[field] for field in fields if field in inmessage}

    exchange = str(inmessage.get("e") or "").strip().upper()
    token = str(inmessage.get("tk") or "").strip()
    if not exchange or not token:
        return
    key = exchange + "|" + token

    valid_quote_tick = _has_valid_live_ltp(message)
    with SUBSCRIPTION_DIAGNOSTIC_LOCK:
        if key in SUBSCRIPTION_DIAGNOSTIC_TOKENS:
            if key not in SUBSCRIPTION_CALLBACK_SEEN:
                SUBSCRIPTION_CALLBACK_SEEN.add(key)
                print(
                    f"SHOONYA FIRST CALLBACK for {key}: "
                    f"type={inmessage.get('t')!r}, lp={message.get('lp')!r}, "
                    f"fields={','.join(message.keys()) or 'none'}"
                )
            if valid_quote_tick and key not in SUBSCRIPTION_CALLBACK_LOGGED:
                SUBSCRIPTION_CALLBACK_LOGGED.add(key)
                print(f"SHOONYA FIRST VALID LTP for {key}: lp={message.get('lp')}")

    with LIVE_DATA_LOCK:
        symbol_info = dict(SYMBOLDICT.get(key, {}))
        symbol_info.update(message)
        SYMBOLDICT[key] = symbol_info
        live_data[key] = symbol_info

def event_handler_order_update(tick_data):
    """Queue broker push updates; Excel writes are handled by the main worker."""
    try:
        if isinstance(tick_data, str):
            tick_data = json.loads(tick_data)
        if not isinstance(tick_data, dict):
            return
        order_id = str(tick_data.get("norenordno") or "").strip()
        status = str(tick_data.get("status") or "").strip().upper()
        reason = str(tick_data.get("rejreason") or "").strip()
        if not order_id or not status:
            return
        executed_price = tick_data.get("avgprc") or tick_data.get("flprc")
        broker_reason = reason if reason and reason != " " else status
        update = {
            "order_id": order_id, "status": status,
            "executed_price": executed_price, "remarks": broker_reason,
        }
        with SIDE_ORDER_TRACKER_LOCK:
            SIDE_ORDER_UPDATES[order_id] = {
                "status": status, "executed_price": executed_price,
                "remarks": broker_reason,
            }
        SIDE_ORDER_UPDATE_QUEUE.put(update)
        print(f"Queued Shoonya order update: {order_id} {status}")
    except Exception as exc:
        print(f"Could not queue broker order update: {exc}")

def process_side_order_updates(max_batch=500):
    """Apply queued broker updates from the Excel-owning thread, not the WebSocket thread."""
    processed = 0
    while processed < max_batch:
        try:
            update = SIDE_ORDER_UPDATE_QUEUE.get_nowait()
        except queue.Empty:
            break
        order_id = update["order_id"]
        status = update.get("status")
        executed_price = update.get("executed_price")
        remarks = update.get("remarks")
        with SIDE_ORDER_TRACKER_LOCK:
            tracked_order = LimitOrderBook.setdefault(
                order_id,
                {"status": "PENDING", "Remarks": None, "Executed_price": None},
            )
            tracked_order["status"] = status or tracked_order.get("status", "PENDING")
            if executed_price not in (None, "", "0", 0):
                tracked_order["Executed_price"] = executed_price
            if remarks:
                tracked_order["Remarks"] = remarks
        update_side_order_display(order_id, status, executed_price, remarks)
        processed += 1
    return processed

def _resubscribe_registered_tokens():
    """Restore every queued/lifetime token through the single subscription path."""
    if not feed_opened:
        return
    with SOCKET_SUBSCRIPTIONS_LOCK:
        tokens = sorted(SOCKET_SUBSCRIPTIONS)
    if not tokens:
        print("WebSocket reconnected; no previous quote subscriptions to restore.")
        return

    sent = 0
    already_sent = 0
    queued = 0
    failed = 0
    for offset in range(0, len(tokens), 10):
        if not feed_opened:
            break
        batch = tokens[offset:offset + 10]
        try:
            result = subscribe_symbols_batch(batch)
            sent += result["sent"]
            already_sent += result["already_sent"]
            queued += result["queued"]
        except Exception as exc:
            failed += len(batch)
            print(f"WebSocket resubscribe batch failed ({batch[0]} .. {batch[-1]}): {exc}")
            sleep(0.10)
        else:
            sleep(0.10)
    print(
        f"WebSocket resubscription finished: {sent} sent, "
        f"{already_sent} already sent, {queued} queued, {failed} failed."
    )

def open_callback():
    global feed_opened
    was_open = feed_opened
    feed_opened = True
    if was_open:
        return

    with SOCKET_SUBSCRIPTIONS_LOCK:
        SOCKET_SESSION_SUBSCRIPTIONS.clear()

    with LIVE_DATA_LOCK:
        live_data.clear()
        SYMBOLDICT.clear()

    print("Shoonya WebSocket connected/reconnected; restoring saved subscriptions...")
    Thread(
        target=_resubscribe_registered_tokens,
        daemon=True,
        name="ShoonyaResubscribe"
    ).start()

def event_handler_socket_closed():
    global feed_opened
    feed_opened = False
    print("WebSocket disconnected; waiting for reconnection.")
    Text2Speech("Warning. Shoonya live data connection disconnected.")
    sleep(2)
    
def order_status(orderid, order_book_snapshot=None):
    """Return broker status from a shared snapshot, including partial fills."""
    average_executed_price = 0
    status = ""
    try:
        order_book = order_book_snapshot if order_book_snapshot is not None else get_order_book()
        if order_book is None or order_book.empty:
            return status, average_executed_price
        order_book = order_book[order_book["Order No"].astype(str) == str(orderid)]
        if not order_book.empty:
            order = order_book.iloc[0]
            status = str(order.get("Status", "") or "").strip().upper()
            average_executed_price = order.get("Executed Price", 0) or 0
            filled = float(order.get("Filled Shares", 0) or 0)
            quantity = float(order.get("Quantity", 0) or 0)
            if filled > 0 and quantity > 0 and filled < quantity and status not in ("COMPLETE", "REJECTED", "CANCELLED", "CANCELED"):
                status = "PARTIAL (" + status + ")"
    except Exception as e:
        print(str(e) + " : Exception occur in order_status")
    return status, average_executed_price

def place_trade(symbol, quantity, buy_or_sell, order_type = None, price = None, row_product_type = "__GLOBAL__", display_row=None):
    global api, BROKER_LOGIN_OK, SYSTEM_INTERNET_OK, SYSTEM_EXCEL_OK, feed_opened
    if not BROKER_LOGIN_OK:
        print("Live order blocked: Shoonya login is not active.")
        return None
    if SYSTEM_INTERNET_OK is False or not internet_is_available():
        print("Live order blocked: internet is unavailable.")
        return None
    if not feed_opened:
        print("Live order blocked: Shoonya WebSocket is disconnected.")
        return None
    if SYSTEM_EXCEL_OK is False:
        print("Live order blocked: Excel workbook is unavailable.")
        return None
    try:
        _ = get_terminal_workbook().sheets["LOGIN"].range("B3").value
    except Exception:
        print("Live order blocked: Excel workbook is not available.")
        return None
    global Product_type
    global logger
    tradingsymbol = symbol[4:]
    exchange = symbol[:3]

    # Only MARKET, LIMIT and SL-M are supported. Anything else is blocked here
    # so price_type/trigger_price can never be left undefined.
    order_type = str(order_type or "").strip().upper()
    if order_type not in ("MARKET", "LIMIT", "SL-M"):
        print(
            f"Live order blocked for {tradingsymbol}: unsupported order type "
            f"{order_type!r}; use MARKET, LIMIT or SL-M."
        )
        return None

    if order_type == 'MARKET':
        # A MARKET selection must reach Shoonya as a true MKT order.
        # Do not convert it to LMT at the last traded price; that can leave
        # the order pending whenever the best bid/ask moves away from the LTP.
        price = 0.0
        trigger_price = None
        price_type = "MKT"
    elif order_type == 'LIMIT':
        if price is None:
            print(f"LIMIT order blocked for {tradingsymbol}: price is missing.")
            return None
        try:
            price = float(price)
        except (TypeError, ValueError):
            print(f"LIMIT order blocked for {tradingsymbol}: invalid price {price!r}.")
            return None
        if not np.isfinite(price) or price <= 0:
            print(f"LIMIT order blocked for {tradingsymbol}: invalid price {price!r}.")
            return None
        trigger_price = None
        price_type = "LMT"
    elif order_type == 'SL-M':
        price_type = "SL-LMT"
        trigger_price = price
        
        if buy_or_sell == 'BUY':
            market_safety = get_market_safety_percent()
            price = price * (100 + market_safety) / 100
            if exchange == 'CDS':
                round_to = .0025
                result = ( int(price / round_to) + 1) * round_to
                price = float( "{:.4f}".format(result) )
            else:
                round_to = .1
                result = ( int(price / round_to) + 1) * round_to
                price = float( "{:.1f}".format(result) )
        else:
            market_safety = get_market_safety_percent()
            price = price * (100 - market_safety) / 100
            if exchange == 'CDS':
                round_to = .0025
                result = ( int(price / round_to) ) * round_to
                price = float( "{:.4f}".format(result) )
            else:
                round_to = .1
                result = ( int(price / round_to) ) * round_to
                price = float( "{:.1f}".format(result) )

    product_map = {"CNC": "C", "MIS": "I", "NRML": "M", "NORMAL": "M"}
    row_product = str(row_product_type or "").strip().upper()
    global_product = str(Product_type or "").strip().upper()
    selected_product_type = row_product if row_product in product_map else global_product
    if selected_product_type not in product_map:
        print(
            f"Live order blocked for {tradingsymbol}: choose CNC, MIS or NRML "
            "in this row's Order Type cell or the global Order Type selector."
        )
        return None
    if not row_product and display_row is not None:
        try:
            product_cell = get_live_trade_sheet(get_terminal_workbook()).range(f"N{int(display_row)}")
            if not str(product_cell.value or "").strip():
                product_cell.value = selected_product_type
        except Exception as product_display_error:
            print(f"Could not display inherited Order Type for {tradingsymbol}: {product_display_error}")
    product = product_map[selected_product_type]
    Message = "Order Details : buy_or_sell = " + str(buy_or_sell[0]) + ", product_type = " + str(product)+ ", exchange = " + str(exchange)+ ", tradingsymbol = " + str(tradingsymbol)+ ", quantity = " + str(quantity)+ ", price_type = " + str(price_type)+ ", price = " + str(price)+ ", trigger_price = " + str(trigger_price)
    logger.info(Message)
    def _show_order_failure(status_text):
        # Keep the broker's exact response visible in this side's Excel status.
        if display_row is not None:
            try:
                side_status_col = "S" if str(buy_or_sell).strip().upper() == "BUY" else "Y"
                status_sheet = get_live_trade_sheet(get_terminal_workbook())
                status_sheet.range(f"{side_status_col}{int(display_row)}").value = str(status_text)[:300]
            except Exception as status_error:
                logger.error(f"Could not display broker order failure: {status_error}")

    try:
        order_response = api.place_order(
            buy_or_sell=buy_or_sell[0],
            product_type=product,
            exchange=exchange,
            tradingsymbol=tradingsymbol,
            quantity=quantity,
            discloseqty=0,
            price_type=price_type,
            price=price,
            trigger_price=trigger_price,
            retention="DAY",
            algo_id=get_shoonya_algo_id(),
            remarks="Python_Trader",
        )
    except Exception as exc:
        Message = f"BROKER SUBMISSION ERROR for {tradingsymbol}: {exc}"
        print(Message)
        logger.exception(Message)
        _show_order_failure(Message)
        return None
    logger.info(f"Raw Shoonya order response for {tradingsymbol}: {order_response!r}")
    if not isinstance(order_response, dict):
        Message = (
            f"BROKER UNCONFIRMED for {tradingsymbol}: Shoonya returned "
            f"{order_response!r}; verify broker OrderBook before retrying."
        )
        print(Message)
        logger.error(Message)
        _show_order_failure(Message)
        return None
    order_id = order_response.get("norenordno") or order_response.get("order_id")
    broker_stat = str(order_response.get("stat", "")).strip().lower()
    if not order_id or broker_stat not in ("ok", "success"):
        broker_reason = order_response.get("emsg") or order_response.get("remarks") or repr(order_response)
        Message = f"BROKER REJECTED/UNCONFIRMED for {tradingsymbol}: {broker_reason}"
        print(Message)
        logger.error(Message)
        _show_order_failure(Message)
        return None
    Message = (
        f"Broker accepted order for {tradingsymbol} {quantity} quantity "
        f"{buy_or_sell}; order id: {order_id}"
    )
    print(Message)
    logger.info(f"Raw Shoonya order response: {order_response}")
    logger.info(Message)
    Telegram_Message.append(Message)
    print(f"Order id = {order_id}")
    register_side_order_display(symbol, buy_or_sell, str(order_id), display_row=display_row)
    return str(order_id)

def get_order_book():
    standard_columns = [
        "Order No", "Status", "Exchange", "Trading Symbol", "Product Type",
        "Transaction Type", "Quantity", "Filled Shares", "Price",
        "Executed Price", "Order Type", "Token", "Lot Size", "Remarks",
        "Order rejection reason", "Trigger Price",
    ]
    empty_order_book = pd.DataFrame(columns=standard_columns)
    try:
        raw_order_book = api.get_order_book()
        if raw_order_book is None:
            print("Shoonya get_order_book returned no data; will retry on next poll.")
            return empty_order_book
        if isinstance(raw_order_book, dict):
            broker_stat = str(raw_order_book.get("stat", "")).strip().lower()
            if broker_stat and broker_stat not in ("ok", "success"):
                print(
                    "Shoonya get_order_book was rejected: "
                    f"{raw_order_book.get('emsg') or raw_order_book}"
                )
                return empty_order_book
            if isinstance(raw_order_book.get("values"), list):
                raw_order_book = raw_order_book["values"]
            elif raw_order_book.get("norenordno"):
                raw_order_book = [raw_order_book]
            else:
                print(f"Unexpected Shoonya order-book response: {raw_order_book}")
                return empty_order_book
        df_order_book = pd.DataFrame(raw_order_book)
        if df_order_book.empty:
            return empty_order_book
        column_map = {
            "norenordno": "Order No",
            "status": "Status",
            "exch": "Exchange",
            "tsym": "Trading Symbol",
            "prctyp": "Product Type",
            "trantype": "Transaction Type",
            "qty": "Quantity",
            "fillshares": "Filled Shares",
            "prc": "Price",
            "avgprc": "Executed Price",
            "prd": "Order Type",
            "token": "Token",
            "ls": "Lot Size",
            "remarks": "Remarks",
            "rejreason": "Order rejection reason",
            "trgprc": "Trigger Price",
        }
        df_order_book = df_order_book.rename(columns=column_map)
        if "Order No" in df_order_book.columns:
            df_order_book = df_order_book.sort_values(
                by="Order No", kind="stable"
            ).reset_index(drop=True)
        ordered_columns = [
            column for column in standard_columns
            if column in df_order_book.columns
        ]
        extra_columns = [
            column for column in df_order_book.columns
            if column not in ordered_columns
        ]
        return df_order_book[ordered_columns + extra_columns]
    except Exception as exc:
        print(f"Exception occur in get_order_book : {exc}")
        return empty_order_book

def GetToken_UsingSymbol(exchange, tradingsymbol):
    global df_ins_NSE, df_ins_BSE, df_ins_NFO, df_ins_CDS, df_ins_MCX, df_ins_BFO
    global api
    exchange = str(exchange or "").strip().upper()
    tradingsymbol = str(tradingsymbol or "").strip().upper()
    if not exchange or not tradingsymbol:
        raise ValueError("Exchange and trading symbol are required to resolve an instrument token.")
    # Excel-friendly alias for the Shoonya Nifty 50 spot index.
    # The live-feed token is NSE|26000; Shoonya's instrument name is "Nifty 50".
    if exchange == "NSE" and tradingsymbol == "NIFTY INDEX":
        return Token_list["NIFTY"]
    if exchange == "BSE" and tradingsymbol in IndexList:
        return Token_list[tradingsymbol]
    master_by_exchange = {
        "NSE": df_ins_NSE, "BSE": df_ins_BSE, "NFO": df_ins_NFO,
        "BFO": df_ins_BFO, "CDS": df_ins_CDS, "MCX": df_ins_MCX,
    }
    master = master_by_exchange.get(exchange)
    if master is not None and not master.empty:
        matches = master[
            master["TradingSymbol"].astype(str).str.strip().str.upper() == tradingsymbol
        ]
        if not matches.empty:
            token = str(matches.iloc[0]["Token"]).strip()
            if token and token.lower() != "nan":
                return token
        response = api.searchscrip(exchange=exchange, searchtext=tradingsymbol)
        values = response.get("values", []) if isinstance(response, dict) else []
        exact_matches = [
            item for item in values
            if str(item.get("tsym") or "").strip().upper() == tradingsymbol
            and str(item.get("token") or "").strip()
        ]
        if not exact_matches:
            candidates = [
                f"{item.get('tsym')} (token {item.get('token')})"
                for item in values[:5]
            ]
            raise ValueError(
                f"No exact instrument-master match for {exchange}:{tradingsymbol}. "
                f"Search candidates: {', '.join(candidates) or 'none'}"
            )
        token = str(exact_matches[0]["token"]).strip()
        print(f"Instrument-master fallback verified: {exchange}:{tradingsymbol} -> {token}")
        return token

def subscribe_symbol(token_key, force=False):
    global api
    token_key = str(token_key or "").strip().upper()
    if "|" not in token_key:
        raise ValueError(f"Invalid Shoonya token key {token_key!r}; expected EXCHANGE|TOKEN")
    exchange, token = token_key.split("|", 1)
    if not exchange or not token:
        raise ValueError(f"Invalid Shoonya token key {token_key!r}")
    with SOCKET_SUBSCRIPTIONS_LOCK:
        SOCKET_SUBSCRIPTIONS.add(token_key)
    if not feed_opened:
        print(f"Queued {token_key}; Shoonya WebSocket is not open yet.")
        return "queued"
    with SOCKET_SUBSCRIBE_SEND_LOCK:
        if not feed_opened:
            print(f"Queued {token_key}; WebSocket closed before the send.")
            return "queued"
        with SOCKET_SUBSCRIPTIONS_LOCK:
            sent_this_session = token_key in SOCKET_SESSION_SUBSCRIPTIONS
            if sent_this_session and not force:
                return "already_sent"
        with SUBSCRIPTION_DIAGNOSTIC_LOCK:
            SUBSCRIPTION_DIAGNOSTIC_TOKENS.add(token_key)
            SUBSCRIPTION_CALLBACK_LOGGED.discard(token_key)
            SUBSCRIPTION_CALLBACK_SEEN.discard(token_key)
        subscribe_result = api.subscribe([token_key])
        if isinstance(subscribe_result, dict):
            response_status = str(subscribe_result.get("stat", "")).strip().lower()
            if response_status not in ("", "ok"):
                raise RuntimeError(
                    f"Shoonya rejected subscribe {token_key}: {subscribe_result}"
                )
        if subscribe_result:
            print(f"Shoonya subscribe response for {token_key}: {subscribe_result}")
        with SOCKET_SUBSCRIPTIONS_LOCK:
            SOCKET_SESSION_SUBSCRIPTIONS.add(token_key)
        return "sent"

def subscribe_symbols_batch(token_keys, force=False):
    global api
    normalized = []
    for item in token_keys:
        key = str(item or "").strip().upper()
        if "|" not in key:
            raise ValueError(f"Invalid Shoonya token key {key!r}")
        exchange, token = key.split("|", 1)
        if not exchange or not token:
            raise ValueError(f"Invalid Shoonya token key {key!r}")
        normalized.append(key)
    normalized = list(dict.fromkeys(normalized))
    with SOCKET_SUBSCRIPTIONS_LOCK:
        SOCKET_SUBSCRIPTIONS.update(normalized)
    if not feed_opened:
        return {"sent": 0, "already_sent": 0, "queued": len(normalized)}
    with SOCKET_SUBSCRIBE_SEND_LOCK:
        if not feed_opened:
            return {"sent": 0, "already_sent": 0, "queued": len(normalized)}
        with SOCKET_SUBSCRIPTIONS_LOCK:
            batch = [key for key in normalized
                     if force or key not in SOCKET_SESSION_SUBSCRIPTIONS]
            already = len(normalized) - len(batch)
            if not batch:
                return {"sent": 0, "already_sent": already, "queued": 0}
        with SUBSCRIPTION_DIAGNOSTIC_LOCK:
            for key in batch:
                SUBSCRIPTION_DIAGNOSTIC_TOKENS.add(key)
                SUBSCRIPTION_CALLBACK_LOGGED.discard(key)
                SUBSCRIPTION_CALLBACK_SEEN.discard(key)
        result = api.subscribe(batch)
        if isinstance(result, dict):
            status = str(result.get("stat", "")).strip().lower()
            if status not in ("", "ok"):
                raise RuntimeError(f"Shoonya rejected subscription batch: {result}")
        with SOCKET_SUBSCRIPTIONS_LOCK:
            SOCKET_SESSION_SUBSCRIPTIONS.update(batch)
        return {"sent": len(batch), "already_sent": already, "queued": 0}

def subscribe_new_token(exchange, Token, force=False):
    token_key = f"{str(exchange).strip().upper()}|{str(Token).strip()}"
    return subscribe_symbol(token_key, force=force)

def _watch_new_symbol_subscription(symbol, exchange, token):
    token_key = f"{exchange}|{token}"
    try:
        for attempt in range(5):
            sleep(3.0)
            with LIVE_DATA_LOCK:
                quote_snapshot = dict(live_data.get(token_key, {}))
            if _has_valid_live_ltp(quote_snapshot):
                print(
                    f"First valid WebSocket LTP received for {symbol} "
                    f"({token_key}): {quote_snapshot.get('lp')}"
                )
                return
            if not feed_opened:
                print(f"Waiting for WebSocket reconnection to receive {symbol} ({token_key}).")
                continue
            if attempt < 4:
                print(
                    f"No WebSocket tick for {symbol} ({token_key}); "
                    f"resending subscription ({attempt + 1}/4)."
                )
                try:
                    subscribe_new_token(exchange, token, force=True)
                except Exception as exc:
                    print(f"Subscription retry failed for {symbol} ({token_key}): {exc}")
        with LIVE_DATA_LOCK:
            received = _has_valid_live_ltp(dict(live_data.get(token_key, {})))
        if not received:
            try:
                quote_probe = api.get_quotes(exchange=exchange, token=str(token))
                if isinstance(quote_probe, dict):
                    print(
                        f"SHOONYA REST TOKEN CHECK for {symbol} ({token_key}): "
                        f"stat={quote_probe.get('stat')!r}, "
                        f"tsym={quote_probe.get('tsym')!r}, "
                        f"token={quote_probe.get('token')!r}, "
                        f"lp={quote_probe.get('lp')!r}, "
                        f"message={quote_probe.get('emsg')!r}"
                    )
                    if str(quote_probe.get("stat", "")).strip().lower() == "ok":
                        expected_tsym = str(symbol).split(":", 1)[-1].strip().upper()
                        returned_tsym = str(quote_probe.get("tsym") or "").strip().upper()
                        returned_token = str(quote_probe.get("token") or "").strip()
                        if not returned_tsym:
                            print(
                                f"REST TOKEN IDENTITY UNVERIFIED for {symbol}: "
                                "Shoonya returned stat=Ok without tsym."
                            )
                        elif returned_tsym != expected_tsym or (
                            returned_token and returned_token != str(token).strip()
                        ):
                            print(
                                f"REST TOKEN MISMATCH for {symbol} ({token_key}): "
                                f"broker returned tsym={returned_tsym!r}, "
                                f"token={returned_token!r}."
                            )
                        else:
                            print(
                                f"REST TOKEN VERIFIED for {symbol} ({token_key}): "
                                f"tsym={returned_tsym}; investigate WebSocket delivery."
                            )
                    else:
                        print(
                            f"REST TOKEN NOT VERIFIED for {symbol} ({token_key}); "
                            "the REST response was not successful."
                        )
                else:
                    print(f"SHOONYA REST TOKEN CHECK for {symbol} ({token_key}): {quote_probe!r}")
            except Exception as probe_error:
                print(f"SHOONYA REST TOKEN CHECK failed for {symbol} ({token_key}): {probe_error}")
        with SUBSCRIPTION_DIAGNOSTIC_LOCK:
            callback_seen = token_key in SUBSCRIPTION_CALLBACK_SEEN
            valid_ltp_seen = token_key in SUBSCRIPTION_CALLBACK_LOGGED
            print(
                f"WARNING: {symbol} ({token_key}) has no valid WebSocket LTP after "
                f"four subscription retries; callback_seen={callback_seen}, "
                f"valid_ltp_seen={valid_ltp_seen}."
            )
    except Exception as exc:
        print(f"Subscription watchdog error for {symbol} ({token_key}): {exc}")

LimitOrderBook = {}
SIDE_ORDER_UPDATE_QUEUE = queue.Queue()
SIDE_ORDER_TRACKER = {}
SIDE_ORDER_TRACKER_LOCK = Lock()
SIDE_ORDER_UPDATES = {}
# Required latch state for the independent Excel BUY/SELL command dispatcher.
# Keep one state per worksheet row and side to prevent duplicate submissions.
SIDE_TRIGGER_STATE_LOCK = Lock()
SIDE_TRIGGER_STATE = {}
SIDE_ORDER_EXECUTOR = ThreadPoolExecutor(
    max_workers=64, thread_name_prefix="ShoonyaSideOrder"
)

def register_side_order_display(symbol, side, order_id, display_row=None):
    try:
        side = str(side).strip().upper()
        if side not in ("BUY", "SELL"):
            return
        workbook = get_terminal_workbook()
        sheet = get_live_trade_sheet(workbook)
        row_number = int(display_row) if display_row is not None else None
        if row_number is None or row_number < 4:
            symbols = sheet.range("A4:A1000").value
            row_number = next(
                (row_index + 4 for row_index, value in enumerate(symbols)
                if _extract_symbol(value) == str(symbol).strip().upper()),
                None,
            )
        else:
            actual_symbol = str(sheet.range(f"A{row_number}").value or "").strip()
            if actual_symbol != str(symbol).strip():
                raise ValueError(
                    f"Order row mismatch: row {row_number} contains {actual_symbol!r}, "
                    f"but submitted symbol is {symbol!r}"
                )
        if row_number is None:
            print(f"Could not locate {symbol} in LIVE_TRADE for side order display.")
            return
        id_col, status_col = ("R", "S") if side == "BUY" else ("X", "Y")
        sheet.range(f"{id_col}{row_number}").number_format = "@"
        sheet.range(f"{id_col}{row_number}").value = str(order_id)
        sheet.range(f"{status_col}{row_number}").value = None
        with SIDE_ORDER_TRACKER_LOCK:
            SIDE_ORDER_TRACKER[str(order_id)] = {
                "row": row_number, "side": side, "symbol": str(symbol),
            }
        with SIDE_ORDER_TRACKER_LOCK:
            early_update = dict(SIDE_ORDER_UPDATES.get(str(order_id)) or {})
        if early_update:
            update_side_order_display(
                str(order_id), early_update.get("status"),
                early_update.get("executed_price"), early_update.get("remarks")
            )
        if str(order_id) in LimitOrderBook:
            LimitOrderBook[str(order_id)].update({
                "status": early_update.get("status", "PENDING"),
                "Executed_price": early_update.get("executed_price"),
                "Remarks": early_update.get("remarks"),
            })
    except Exception as exc:
        print(f"Could not update {side} order display for {symbol}: {exc}")

def update_side_order_display(order_id, status, executed_price=None, remarks=None):
    try:
        with SIDE_ORDER_TRACKER_LOCK:
            tracked = SIDE_ORDER_TRACKER.get(str(order_id))
            if not tracked:
                return
            row_number = tracked["row"]
            if tracked["side"] == "BUY":
                status_col, price_col = "S", "T"
            else:
                status_col, price_col = "Y", "Z"
            sheet = get_live_trade_sheet(get_terminal_workbook())
            display_status = str(status or "").strip().upper() or "PENDING"
            old_status = str(sheet.range(f"{status_col}{row_number}").value or "").strip().upper()
            old_base = old_status.split("|", 1)[0].split(":", 1)[0].strip()
            new_base = display_status.split("|", 1)[0].split(":", 1)[0].strip()
            final_states = ("COMPLETE", "COMPLETED", "REJECTED", "CANCELLED", "CANCELED")
            if old_base in final_states and new_base not in final_states:
                display_status = old_status
            sheet.range(f"{status_col}{row_number}").value = display_status
            if executed_price not in (None, "", 0, "0"):
                try:
                    if float(executed_price) > 0:
                        sheet.range(f"{price_col}{row_number}").value = float(executed_price)
                except (TypeError, ValueError):
                    pass
            if (display_status.startswith("REJECTED") or display_status.startswith("CANCELLED")
                or display_status.startswith("CANCELED")) and remarks:
                reason_text = str(remarks).strip()
                if reason_text and reason_text.upper() not in display_status:
                    sheet.range(f"{status_col}{row_number}").value = f"{display_status}: {reason_text}"
    except Exception as exc:
        print(f"Could not refresh side order {order_id}: {exc}")

def register_pending_order(order_id):
    key = str(order_id)
    with SIDE_ORDER_TRACKER_LOCK:
        early = SIDE_ORDER_UPDATES.get(key)
        if early:
            LimitOrderBook[key] = {
                "status": early.get("status", "PENDING"),
                "Remarks": early.get("remarks"),
                "Executed_price": early.get("executed_price"),
            }
        else:
            LimitOrderBook[key] = {
                "status": "PENDING", "Remarks": None, "Executed_price": None
            }

def _is_terminal_trigger(value, side):
    if value is None:
        return False
    return str(value).strip().upper() in (side, "TRUE", "1", "YES", "ON")

def _required_positive_number(value, field_name):
    if value is None or (isinstance(value, str) and not value.strip()):
        raise ValueError(f"{field_name} is blank")
    try:
        number = float(value)
    except (TypeError, ValueError):
        raise ValueError(f"{field_name} must be a number")
    if not np.isfinite(number) or number <= 0:
        raise ValueError(f"{field_name} must be greater than zero")
    return number

def _execute_terminal_side_order(row_number, symbol, side, qty, product_type,
                                 market_type, limit_price, ltp):
    pythoncom = None
    try:
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pythoncom = None
        sheet = get_live_trade_sheet(get_terminal_workbook())
        id_col, status_col = ("R", "S") if side == "BUY" else ("X", "Y")
        quantity_value = _required_positive_number(qty, "Quantity")
        if not quantity_value.is_integer():
            raise ValueError("Quantity must be a whole number")
        quantity = int(quantity_value)
        product_type = str(product_type or "").strip().upper()
        if product_type not in ("MIS", "CNC", "NRML"):
            raise ValueError("Select a valid Product Type: MIS, CNC or NRML")
        market_type = str(market_type or "").strip().upper()
        if market_type == "MARKET":
            # The dispatcher validates that a fresh LTP exists before queueing.
            # Submit the broker's actual MARKET order, not a LIMIT at that LTP.
            broker_order_type = "MARKET"
            order_price = 0.0
        elif market_type == "LIMIT":
            broker_order_type = "LIMIT"
            order_price = _required_positive_number(limit_price, "Limit Price")
        elif market_type == "SL-LMT":
            broker_order_type = "SL-M"
            order_price = _required_positive_number(limit_price, "Trigger Price")
        else:
            raise ValueError("Select MARKET, LIMIT or SL-LMT in the BUY/SELL type cell")
        order_id = place_trade(
            symbol, quantity, side, broker_order_type, order_price,
            product_type, display_row=row_number
        )
        if order_id is None:
            current_status = str(sheet.range(f"{status_col}{row_number}").value or "").strip()
            if not current_status.startswith(("BROKER REJECTED", "BROKER UNCONFIRMED", "BROKER SUBMISSION ERROR")):
                sheet.range(f"{status_col}{row_number}").value = "SUBMISSION FAILED/UNCONFIRMED"
            print(f"{side} order for {symbol} was not confirmed; check Shoonya OrderBook before retrying.")
            return
        register_pending_order(order_id)
        sheet.range(f"{id_col}{row_number}").number_format = "@"
        sheet.range(f"{id_col}{row_number}").value = str(order_id)
        print(f"{side} order worker completed for {symbol}: {order_id}")
    except Exception as exc:
        try:
            sheet = get_live_trade_sheet(get_terminal_workbook())
            status_col = "S" if side == "BUY" else "Y"
            sheet.range(f"{status_col}{row_number}").value = "INPUT/SUBMISSION ERROR: " + str(exc)
        except Exception:
            pass
        print(f"{side} command for {symbol}, row {row_number}, failed: {exc}")
    finally:
        try:
            if pythoncom is not None:
                pythoncom.CoUninitialize()
        except Exception:
            pass

def dispatch_terminal_side_order(sheet, row_number, symbol, side, qty, product_type,
                                 market_type, limit_price, ltp):
    side = str(side).strip().upper()
    trigger_col = "O" if side == "BUY" else "U"
    id_col, status_col, price_col = (("R", "S", "T") if side == "BUY"
                                     else ("X", "Y", "Z"))
    key = (int(row_number), side)
    trigger = sheet.range(f"{trigger_col}{row_number}").value
    if not _is_terminal_trigger(trigger, side):
        with SIDE_TRIGGER_STATE_LOCK:
            SIDE_TRIGGER_STATE[key] = {"armed": False}
        return
    try:
        _required_positive_number(qty, "Quantity")
        if str(product_type or "").strip().upper() not in ("MIS", "CNC", "NRML"):
            raise ValueError("Select Product Type: MIS, CNC or NRML")
        selected_market_type = str(market_type or "").strip().upper()
        if selected_market_type == "MARKET":
            _required_positive_number(ltp, "Live LTP (waiting for fresh quote)")
        elif selected_market_type == "LIMIT":
            _required_positive_number(limit_price, "Limit Price")
        elif selected_market_type == "SL-LMT":
            _required_positive_number(limit_price, "Trigger Price")
        else:
            raise ValueError("Select MARKET, LIMIT or SL-LMT")
    except ValueError as input_error:
        waiting_message = "WAITING: " + str(input_error)
        if str(sheet.range(f"{status_col}{row_number}").value or "") != waiting_message:
            sheet.range(f"{status_col}{row_number}").value = waiting_message
        return
    with SIDE_TRIGGER_STATE_LOCK:
        state = SIDE_TRIGGER_STATE.get(key)
        existing_id = sheet.range(f"{id_col}{row_number}").value
        existing_status = str(sheet.range(f"{status_col}{row_number}").value or "").strip()
        id_is_blank = existing_id is None or not str(existing_id).strip()
        status_is_blank = not existing_status

        # If both broker tracking cells are blank, treat the active BUY/SELL
        # trigger as a fresh command even if an old in-memory latch remains.
        # QUEUED/SUBMITTING/PENDING or any broker result prevents resubmission.
        if id_is_blank and status_is_blank:
            state = {"armed": False}
            SIDE_TRIGGER_STATE[key] = state
        elif state is None:
            if existing_id not in (None, "") or (
                existing_status and not existing_status.upper().startswith("WAITING:")
            ):
                SIDE_TRIGGER_STATE[key] = {"armed": True, "recovered": True}
                return
            state = {"armed": False}
            SIDE_TRIGGER_STATE[key] = state
        if state.get("armed"):
            return
        sheet.range(f"{id_col}{row_number}").value = None
        sheet.range(f"{status_col}{row_number}").value = None
        sheet.range(f"{price_col}{row_number}").value = None
        SIDE_TRIGGER_STATE[key] = {"armed": True, "queued": True}
    try:
        SIDE_ORDER_EXECUTOR.submit(
            _execute_terminal_side_order,
            int(row_number), str(symbol), side, qty, product_type,
            market_type, limit_price, ltp
        )
    except Exception as exc:
        sheet.range(f"{status_col}{row_number}").value = "QUEUE ERROR: " + str(exc)
        print(f"Could not queue {side} order for {symbol}: {exc}")

def get_live_trade_sheet(workbook):
    with LIVE_TRADE_SHEET_LOCK:
        sheet_names = [sheet.name for sheet in workbook.sheets]
        if "LIVE_TRADE" in sheet_names:
            return workbook.sheets["LIVE_TRADE"]
        if "Trade_Terminal" in sheet_names:
            old_sheet = workbook.sheets["Trade_Terminal"]
            old_sheet.name = "LIVE_TRADE"
            try:
                workbook.save()
            except Exception as save_error:
                print(f"LIVE_TRADE renamed in Excel but workbook save failed: {save_error}")
            return old_sheet
        raise KeyError(
            "Neither LIVE_TRADE nor Trade_Terminal worksheet exists in VII_CLASS.xlsm."
        )

def start_Trade_Terminal():
    global api, live_data, BROKER_LOGIN_OK
    global SYMBOLDICT
    global Product_type
    global Trade_Mode
    global Telegram_Message, Voice_Message
    global subs_lst
    excel_TT = get_terminal_workbook()
    tt = get_live_trade_sheet(excel_TT)

    internal_marker = str(tt.range("BA3").value or "").strip()
    if internal_marker != "__PT_INTERNAL_QTY":
        legacy_headers = tt.range("M3:AE3").value
        legacy_rows = tt.range("M4:AE1000").value
        old_side_headers = [str(v or "").strip().upper() for v in (tt.range("AF3:AK3").value or [])]
        old_indicator_headers_af = [str(v or "").strip().upper() for v in (tt.range("AF3:AM3").value or [])]
        old_indicator_headers_al = [str(v or "").strip().upper() for v in (tt.range("AL3:AS3").value or [])]
        old_side_rows = tt.range("AF4:AK1000").value if old_side_headers and old_side_headers[0] == "BUY ORDER ID" else None

        tt.range("M3:AE1000").copy(destination=tt.range("BA3:BS1000"))
        tt.range("BA3:BS3").value = [
            "__PT_INTERNAL_QTY", "Internal BUY Trigger", "Internal SELL Trigger",
            "Internal Entry Signal", "Internal Entry Limit", "Internal Entry Fill",
            "Internal Entry ID", "Internal Entry Remarks", "Internal Exit Signal",
            "Internal Exit Fill", "Internal Exit ID", "Internal Exit Remarks",
            "Internal Target", "Internal SL", "Internal Trail Enable",
            "Internal Latest SL", "Internal Trade Status", "Internal Realized PnL",
            "Internal Product Type"
        ]

        if old_indicator_headers_al[0:1] == ["RSI"]:
            indicator_values = tt.range("AL3:AS1000").value
        elif old_indicator_headers_af[0:1] == ["RSI"]:
            indicator_values = tt.range("AF3:AM1000").value
        else:
            indicator_values = None

        tt.range("N4:N1000").value = [
            (str(row[18]).strip().upper() if len(row) > 18 and str(row[18] or "").strip().upper() in ("CNC", "MIS", "NRML") else None)
            for row in legacy_rows
        ]
        tt.range("O4:O1000").value = [row[1] if len(row) > 1 else None for row in legacy_rows]
        tt.range("U4:U1000").value = [row[2] if len(row) > 2 else None for row in legacy_rows]
        tt.range("P4:P1000").value = [row[3] if len(row) > 3 else None for row in legacy_rows]
        tt.range("Q4:Q1000").value = [row[4] if len(row) > 4 else None for row in legacy_rows]
        tt.range("V4:V1000").value = [row[3] if len(row) > 3 else None for row in legacy_rows]
        tt.range("W4:W1000").value = [row[4] if len(row) > 4 else None for row in legacy_rows]

        if old_side_rows is not None:
            tt.range("R4:T1000").value = [[r[0], r[1], r[2]] for r in old_side_rows]
            tt.range("X4:Z1000").value = [[r[3], r[4], r[5]] for r in old_side_rows]

        tt.range("AA4:AF1000").value = [
            [row[12] if len(row) > 12 else None,
             row[13] if len(row) > 13 else None,
             row[14] if len(row) > 14 else None,
             row[15] if len(row) > 15 else None,
             row[16] if len(row) > 16 else None,
             row[17] if len(row) > 17 else None]
            for row in legacy_rows
        ]
        if indicator_values is not None:
            tt.range("AG3:AN1000").value = indicator_values
        print("Migrated LIVE_TRADE to the grouped BUY/SELL layout; legacy trade state preserved in hidden BA:BS.")

    tt.range("M3:AF3").value = [
        "Qty", "Order Type", "BUY", "BUY Market Type", "BUY Limit Price",
        "BUY ID", "BUY Status", "BUY Executed Price",
        "SELL", "SELL Market Type", "SELL Limit Price",
        "SELL ID", "SELL Status", "SELL Executed Price",
        "Target", "SL", "Trail Enable", "Latest SL",
        "Overall Trade Status", "Realized PnL"
    ]
    tt.range("AG3:AN3").value = [
        "RSI", "StochRSI", "SMA Stoch", "BB Upper",
        "BB Middle", "BB Lower", "BUY Signal", "SELL Signal"
    ]
    try:
        tt.range("N4:N1000").api.Validation.Delete()
        tt.range("N4:N1000").api.Validation.Add(Type=3, AlertStyle=1, Operator=1, Formula1="CNC,MIS,NRML")
        tt.range("N4:N1000").api.Validation.IgnoreBlank = True
        tt.range("N4:N1000").api.Validation.InCellDropdown = True
    except Exception as validation_error:
        print(f"Order Type dropdown could not be applied: {validation_error}")
    for market_range in ("P4:P1000", "V4:V1000"):
        try:
            tt.range(market_range).api.Validation.Delete()
            tt.range(market_range).api.Validation.Add(Type=3, AlertStyle=1, Operator=1, Formula1="MARKET,LIMIT,SL-LMT")
            tt.range(market_range).api.Validation.IgnoreBlank = True
            tt.range(market_range).api.Validation.InCellDropdown = True
        except Exception as validation_error:
            print(f"Market Type dropdown could not be applied to {market_range}: {validation_error}")
    for order_id_col in ("R", "X"):
        tt.range(f"{order_id_col}4:{order_id_col}1000").number_format = "@"
    for order_status_col in ("S", "Y"):
        tt.range(f"{order_status_col}4:{order_status_col}1000").number_format = "@"
    # Preserve the column widths already configured by the user in Excel.
    # Do not assign column_width during startup; this used to resize R:Z and AG:AN
    # every time run.py initialized the LIVE_TRADE sheet.
    try:
        tt.range("BA:BS").api.EntireColumn.Hidden = True
    except Exception as hide_error:
        print(f"Internal trade-state columns could not be hidden: {hide_error}")
    tt.range('k1').value =  'PYTHON TRADER'
    tt.range('k1').color = (46,132,198)  
    Trade_Mode = "REAL"
    tt.range("S2").value = "REAL"
    try:
        excel_TT.save()
    except Exception as save_error:
        print(f"LIVE_TRADE layout is active but workbook save failed: {save_error}")
    
    AlertMessage = Trade_Mode + " trade mode enabled"
    Telegram_Message.append(AlertMessage)
    
    Symbol_Token = {}
    terminal_subscribed = set()
    subscription_pending = set()
    subscription_retry_after = {}
    force_resubscribe_symbols = set()
    subscription_lock = Lock()
    last_excel_reconnect_generation = EXCEL_RECONNECT_GENERATION
    from concurrent.futures import ThreadPoolExecutor
    subscription_executor = ThreadPoolExecutor(max_workers=2, thread_name_prefix="SymbolSubscribe")
    feed_watchdog_executor = ThreadPoolExecutor(max_workers=8, thread_name_prefix="FeedWatchdog")

    def subscribe_symbol_in_background(symbol):
        try:
            symbol = str(symbol).strip().upper()
            if len(symbol) < 6 or symbol[3] != ":":
                raise ValueError(f"Expected symbol format EXCHANGE:TRADINGSYMBOL, received {symbol!r}")
            exchange = symbol[:3]
            tradingsymbol = symbol[4:]
            token = GetToken_UsingSymbol(exchange, tradingsymbol)
            if token is None or str(token).strip() == "":
                raise ValueError(f"No instrument token found for {symbol}")
            token_key = exchange + "|" + str(token)
            if not BROKER_LOGIN_OK:
                raise RuntimeError(
                    "Shoonya REST session is not validated; refusing market-feed subscription."
                )
            with subscription_lock:
                force_resubscribe = symbol in force_resubscribe_symbols
            Symbol_Token[symbol] = token_key
            try:
                subscribe_new_token(exchange, token, force=force_resubscribe)
            except Exception:
                Symbol_Token.pop(symbol, None)
                raise
            with subscription_lock:
                if force_resubscribe:
                    force_resubscribe_symbols.discard(symbol)
                terminal_subscribed.add(symbol)
            print(f"Symbol = {symbol}, Token={token} subscription sent; waiting for first WebSocket tick")
            feed_watchdog_executor.submit(
                _watch_new_symbol_subscription, symbol, exchange, token
            )
        except Exception as exc:
            print(f"Subscribe error {symbol}: {exc}")
            with subscription_lock:
                subscription_retry_after[symbol] = _monotonic() + 5.0
        finally:
            with subscription_lock:
                subscription_pending.discard(symbol)

    def request_weekly_symbol_subscription(symbol):
        symbol = _extract_symbol(symbol)
        if not symbol:
            return
        with subscription_lock:
            retry_ready = _monotonic() >= subscription_retry_after.get(symbol, 0.0)
            needs_subscription = (
                symbol not in terminal_subscribed
                and symbol not in subscription_pending
                and retry_ready
            )
            if needs_subscription:
                subscription_pending.add(symbol)
        if needs_subscription:
            subscription_executor.submit(subscribe_symbol_in_background, symbol)

    def get_weekly_symbol_quote(symbol):
        token_key = Symbol_Token.get(symbol)
        if not token_key:
            return {}
        with LIVE_DATA_LOCK:
            return dict(live_data.get(token_key, {}))

    def write_live_market_data():
        pythoncom = None
        try:
            try:
                import pythoncom
                pythoncom.CoInitialize()
            except Exception:
                pythoncom = None

            market_sheet = get_live_trade_sheet(get_terminal_workbook())
            print("Independent live market-data writer started (LTP/OHLC/Bid/Ask/Volume/OI).")
            while True:
                try:
                    market_sheet = get_live_trade_sheet(get_terminal_workbook())
                    market_symbols = market_sheet.range("A4:A1000").value
                    market_rows = []
                    for market_symbol in market_symbols:
                        values = [None] * 11
                        symbol_str = _extract_symbol(market_symbol)
                        if symbol_str:
                            token_key = Symbol_Token.get(symbol_str)
                            if token_key:
                                with LIVE_DATA_LOCK:
                                    quote_snapshot = dict(live_data.get(token_key, {}))
                                values = [
                                    quote_snapshot.get("o", "-"),
                                    quote_snapshot.get("h", "-"),
                                    quote_snapshot.get("l", "-"),
                                    quote_snapshot.get("c", "-"),
                                    quote_snapshot.get("ap", "-"),
                                    quote_snapshot.get("bp1", "-"),
                                    quote_snapshot.get("sp1", "-"),
                                    quote_snapshot.get("v", "-"),
                                    quote_snapshot.get("oi", "-"),
                                    quote_snapshot.get("lp", "-"),
                                    quote_snapshot.get("pc", "-"),
                                ]
                        market_rows.append(values)
                    market_sheet.range("B4:L1000").value = market_rows
                except Exception as market_error:
                    if not is_excel_connection_lost():
                        print(f"Live market-data refresh error (order workers continue): {market_error}")
                    sleep(1.0 if is_excel_connection_lost() else 0.15)
        except Exception as writer_error:
            print(f"Live market-data writer stopped: {writer_error}")
        finally:
            try:
                if pythoncom is not None:
                    pythoncom.CoUninitialize()
            except Exception:
                pass

    Thread(target=write_live_market_data, daemon=True, name="LiveMarketDataWriter").start()

    global LimitOrderBook
    while True:
        try:
            excel_TT = get_terminal_workbook()
            tt = get_live_trade_sheet(excel_TT)
            reload_indicator_engine_if_settings_changed(excel_TT)
            weekly_data_engine.refresh_from_excel(
                excel_TT, api,
                request_weekly_symbol_subscription,
                get_weekly_symbol_quote
            )
            current_reconnect_generation = EXCEL_RECONNECT_GENERATION
            if current_reconnect_generation != last_excel_reconnect_generation:
                try:
                    session_check = api.get_limits()
                    session_valid = (
                        isinstance(session_check, dict)
                        and str(session_check.get("stat", "")).strip().lower() == "ok"
                    )
                except Exception as session_error:
                    session_valid = False
                    print(f"Shoonya session check failed after Excel recovery: {session_error}")

                if not session_valid:
                    BROKER_LOGIN_OK = False
                    last_excel_reconnect_generation = current_reconnect_generation
                    print(
                        "Excel reopened, but Shoonya REST session validation failed. "
                        "Fresh WebSocket subscriptions were NOT sent. Restart run.py "
                        "to perform the normal reusable-session/OAuth login again."
                    )
                elif not feed_opened:
                    print(
                        "Excel reopened and Shoonya REST session is valid, but the "
                        "WebSocket is not open yet. Subscription workers will retry "
                        "when the socket reconnects."
                    )
                else:
                    active_symbols = [
                        _extract_symbol(value)
                        for value in tt.range("A4:A1000").value
                        if _extract_symbol(value)
                    ]
                    with subscription_lock:
                        terminal_subscribed.clear()
                        force_resubscribe_symbols.update(active_symbols)
                    last_excel_reconnect_generation = current_reconnect_generation
                    print(
                        "Excel recovery confirmed; Shoonya REST session is valid and "
                        "WebSocket is open. Scheduling fresh subscriptions for "
                        f"{len(active_symbols)} active Trade_Terminal symbols."
                    )
            Product_type = tt.range(f"P2").value
            symbols = tt.range(f"a{4}:a{1000}").value
            trading_info = tt.range("BA4:BS1000").value
            terminal_inputs = tt.range("M4:W1000").value
            management_inputs = tt.range("AA4:AD1000").value
            for row_offset, internal_row in enumerate(trading_info):
                ui = terminal_inputs[row_offset]
                management = management_inputs[row_offset]
                internal_row[0] = ui[0]
                internal_row[1] = None
                internal_row[2] = None
                internal_row[18] = ui[1]
                internal_row[11] = management[0]
                internal_row[12] = management[1]
                internal_row[13] = management[2]
                internal_row[14] = management[3]
            tt.range("BA4:BS1000").value = trading_info
            main_list = []
            indicator_list = []

            idx = 0
            for i in symbols:
                lst = [None, None, None, None,None, None, None, None, None,None,None]
                indicator_row = [None, None, None, None, None, None, "", ""]
                # Excel cells may contain trailing spaces or inconsistent case, and xlwings
                # returns vertical ranges as a list of lists (e.g., [['NFO:NIFTY']]).
                # Use the helper to safely extract the canonical string key.
                i = _extract_symbol(i)
                if i:
                    with subscription_lock:
                        retry_ready = _monotonic() >= subscription_retry_after.get(i, 0.0)
                        needs_subscription = (
                            (i not in terminal_subscribed or i in force_resubscribe_symbols)
                            and i not in subscription_pending
                            and retry_ready
                        )
                        if needs_subscription:
                            subscription_pending.add(i)
                    if needs_subscription:
                        subscription_executor.submit(subscribe_symbol_in_background, i)
                    if i in terminal_subscribed and i in Symbol_Token:
                        try:
                            exchange = i[:3]
                            tradingsymbol = i[4:]
                            TokenKey = Symbol_Token[i]
                            with LIVE_DATA_LOCK:
                                quote = dict(live_data.get(TokenKey, {}))

                            lst = [
                                quote.get("o", "-"),
                                quote.get("h", "-"),
                                quote.get("l", "-"),
                                quote.get("c", "-"),
                                quote.get("ap", "-"),
                                quote.get("bp1", "-"),
                                quote.get("sp1", "-"),
                                quote.get("v", "-"),
                                quote.get("oi", "-"),
                                quote.get("lp", "-"),
                                quote.get("pc", "-"),
                            ]
                            indicator_row = indicator_engine.update(
                                api, exchange, tradingsymbol,
                                convert_to_float(quote.get("lp", 0))
                            )

                            # Weekly OHLCV is maintained by the separate Weekly.py module.
                            weekly_data_engine.update(api, exchange, tradingsymbol, quote)

                            try:
                                trade_info = trading_info[idx]
                                excel_row = idx + 4
                                ui_row = terminal_inputs[idx]
                                current_ltp = convert_to_float(quote.get("lp", 0))
                                dispatch_terminal_side_order(
                                    tt, excel_row, i, "BUY", ui_row[0], ui_row[1],
                                    ui_row[3], ui_row[4], current_ltp
                                )
                                dispatch_terminal_side_order(
                                    tt, excel_row, i, "SELL", ui_row[0], ui_row[1],
                                    ui_row[9], ui_row[10], current_ltp
                                )
                                has_new_trigger = (
                                    _is_terminal_trigger(trade_info[1], "BUY")
                                    or _is_terminal_trigger(trade_info[2], "SELL")
                                )
                                has_live_trade_state = trade_info[16] in ("Entry_Pending", "Active", "Exit_Pending")
                                if trade_info[0] is not None and (has_new_trigger or has_live_trade_state):
                                    if isinstance(trade_info[0], (int, float)):
                                        LTP = float(quote.get("lp", 0))
                                        if Trade_Mode == 'REAL':
                                            if _is_terminal_trigger(trade_info[1], "BUY") and LTP != 0:
                                                if trade_info[3] in ['True_Market' ,'True_Limit_LTP', 'Limit_Below', 'Limit_Above']:
                                                    if trade_info[16] != 'Active' and trade_info[16] != 'Entry_Pending' and trade_info[16] != 'Exit_Pending' and trade_info[16] != 'Closed' and (trade_info[16] is None or trade_info[16] == ''):
                                                        if trade_info[3] == 'True_Market':
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", " + "BUY" + ", " + str(trade_info[3])+ ", " + str(trade_info[4])+ " triggered"
                                                            logger.info(Message)
                                                            order_id = place_trade(i, int(trade_info[0]), "BUY","LIMIT",LTP, trade_info[18], display_row=idx + 4)
                                                            if order_id is None:
                                                                tt.range(f"BQ{excel_row}").value = "Order_Rejected/Unconfirmed"
                                                            else:
                                                                register_pending_order(order_id)
                                                                tt.range(f"BF{excel_row}").value = None
                                                                tt.range(f"BG{excel_row}").value = order_id
                                                                tt.range(f"BQ{excel_row}").value = 'Entry_Pending'
                                                                tt.range(f"BJ{excel_row}").value = None
                                                                tt.range(f"BK{excel_row}").value = None
                                                                tt.range(f"BP{excel_row}").value = None
                                                                tt.range(f"BR{excel_row}").value = None
                                                        elif trade_info[3] == 'True_Limit_LTP':    
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", " + "BUY" + ", " + str(trade_info[3])+ ", " + str(trade_info[4])+ " triggered"
                                                            logger.info(Message)
                                                            order_id = place_trade(i, int(trade_info[0]), "BUY","LIMIT",LTP, trade_info[18], display_row=idx + 4)
                                                            if order_id is None:
                                                                tt.range(f"BQ{excel_row}").value = "Order_Rejected/Unconfirmed"
                                                            else:
                                                                register_pending_order(order_id)
                                                                tt.range(f"BF{excel_row}").value = None
                                                                tt.range(f"BG{excel_row}").value = order_id
                                                                tt.range(f"BQ{excel_row}").value = 'Entry_Pending'
                                                                tt.range(f"BJ{excel_row}").value = None
                                                                tt.range(f"BK{excel_row}").value = None
                                                                tt.range(f"BP{excel_row}").value = None
                                                                tt.range(f"BR{excel_row}").value = None
                                                        elif trade_info[3] == 'Limit_Below' and trade_info[4] is not None:
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", " + "BUY" + ", " + str(trade_info[3])+ ", " + str(trade_info[4])+ " triggered"
                                                            logger.info(Message)
                                                            order_id = place_trade(i, int(trade_info[0]), "BUY","LIMIT",trade_info[4], trade_info[18], display_row=idx + 4)
                                                            if order_id is None:
                                                                tt.range(f"BQ{excel_row}").value = "Order_Rejected/Unconfirmed"
                                                            else:
                                                                register_pending_order(order_id)
                                                                tt.range(f"BF{excel_row}").value = None
                                                                tt.range(f"BG{excel_row}").value = order_id
                                                                tt.range(f"BQ{excel_row}").value = 'Entry_Pending'
                                                                tt.range(f"BJ{excel_row}").value = None
                                                                tt.range(f"BK{excel_row}").value = None
                                                                tt.range(f"BP{excel_row}").value = None
                                                                tt.range(f"BR{excel_row}").value = None
                                                        elif trade_info[3] == 'Limit_Above' and trade_info[4] is not None:
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", " + "BUY" + ", " + str(trade_info[3])+ ", " + str(trade_info[4])+ " triggered"
                                                            logger.info(Message)
                                                            order_id = place_trade(i, int(trade_info[0]), "BUY","SL-M",trade_info[4], trade_info[18], display_row=idx + 4)
                                                            if order_id is None:
                                                                tt.range(f"BQ{excel_row}").value = "Order_Rejected/Unconfirmed"
                                                            else:
                                                                register_pending_order(order_id)
                                                                tt.range(f"BF{excel_row}").value = None
                                                                tt.range(f"BG{excel_row}").value = order_id
                                                                tt.range(f"BQ{excel_row}").value = 'Entry_Pending'
                                                                tt.range(f"BJ{excel_row}").value = None
                                                                tt.range(f"BK{excel_row}").value = None
                                                                tt.range(f"BP{excel_row}").value = None
                                                                tt.range(f"BR{excel_row}").value = None
                                                    elif trade_info[16] == 'Entry_Pending' :
                                                        Order_id = str(int(trade_info[6]))
                                                        status = LimitOrderBook[Order_id]['status']
                                                        if status == 'COMPLETE':
                                                            Executed_price = LimitOrderBook[Order_id]['Executed_price']
                                                            tt.range(f"BF{excel_row}").value = Executed_price
                                                            tt.range(f"BQ{excel_row}").value = 'Active'
                                                            Message = str(int(trade_info[0])) + " quantity " + str(i[4:]) + " Bought at " + str(Executed_price)
                                                            Telegram_Message.append(Message)
                                                        if status in ("REJECTED", "CANCELLED", "CANCELED"):
                                                            tt.range(f"BQ{excel_row}").value = status
                                                        tt.range(f"BH{excel_row}").value = LimitOrderBook[Order_id]['Remarks']
                                                    elif trade_info[16] == 'Active':
                                                        PnL = (float(LTP) - float(trade_info[5])) * int(trade_info[0])
                                                        tt.range(f"BR{excel_row}").value = PnL
                                                        TSL = 0
                                                        if type(trade_info[13]) is float:
                                                            if trade_info[14] == True:
                                                                if(LTP >= trade_info[5]):
                                                                    CUR_TSL = LTP - trade_info[5] + trade_info[13]
                                                                else:
                                                                    CUR_TSL = trade_info[13]
                                                                OLD_TSL = trade_info[15]
                                                                if OLD_TSL is not None:
                                                                    TSL = max(CUR_TSL,OLD_TSL)
                                                                else:
                                                                    TSL = CUR_TSL 
                                                            else:
                                                                TSL = trade_info[13]
                                                            tt.range(f"BP{excel_row}").value = TSL
                                                        if trade_info[8] == 'True_Market':
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", " + str(trade_info[8]) +  " exit triggered"
                                                            logger.info(Message)
                                                            order_id = place_trade(i, int(trade_info[0]), "SELL","LIMIT",LTP, trade_info[18], display_row=idx + 4)
                                                            if order_id is None:
                                                                tt.range(f"BI{excel_row}").value = None
                                                            else:  
                                                                register_pending_order(order_id)
                                                                tt.range(f"BK{excel_row}").value = order_id
                                                                tt.range(f"BQ{excel_row}").value = 'Exit_Pending'
                                                        elif trade_info[8] == 'True_Limit_LTP':
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", " + str(trade_info[8]) +  " exit triggered"
                                                            logger.info(Message)
                                                            order_id = place_trade(i, int(trade_info[0]), "SELL","LIMIT",LTP, trade_info[18], display_row=idx + 4)
                                                            if order_id is None:
                                                                tt.range(f"BI{excel_row}").value = None
                                                            else:  
                                                                register_pending_order(order_id)
                                                                tt.range(f"BK{excel_row}").value = order_id
                                                                tt.range(f"BQ{excel_row}").value = 'Exit_Pending'
                                                        elif type(trade_info[12]) is float and trade_info[12] <= LTP:
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", target meets so exit triggered"
                                                            logger.info(Message)
                                                            order_id = place_trade(i, int(trade_info[0]), "SELL","LIMIT",LTP, trade_info[18], display_row=idx + 4)
                                                            register_pending_order(order_id)
                                                            tt.range(f"BK{excel_row}").value = order_id
                                                            tt.range(f"BQ{excel_row}").value = 'Exit_Pending'
                                                        elif TSL >= LTP and type(trade_info[13]) is float:
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", TSL meets so exit triggered"
                                                            logger.info(Message)
                                                            order_id = place_trade(i, int(trade_info[0]), "SELL","LIMIT",LTP, trade_info[18], display_row=idx + 4)
                                                            register_pending_order(order_id)
                                                            tt.range(f"BK{excel_row}").value = order_id
                                                            tt.range(f"BQ{excel_row}").value = 'Exit_Pending'
                                                    elif trade_info[16] ==  'Exit_Pending':
                                                        Order_id = str(int(trade_info[10]))
                                                        status = LimitOrderBook[Order_id]['status']
                                                        if status == 'COMPLETE':
                                                            Executed_price = LimitOrderBook[Order_id]['Executed_price']
                                                            tt.range(f"BJ{excel_row}").value = Executed_price
                                                            tt.range(f"BQ{excel_row}").value = 'Closed'
                                                            PnL = (float(Executed_price) - float(trade_info[5])) * int(trade_info[0])
                                                            tt.range(f"BR{excel_row}").value = PnL
                                                            Message = str(int(trade_info[0])) + " quantity " + str(i[4:]) + " Sold at " + str(Executed_price) + ", Trade profit or Loss = " + str(round(PnL,2))
                                                            Telegram_Message.append(Message)
                                                        if status in ("REJECTED", "CANCELLED", "CANCELED"):
                                                            tt.range(f"BQ{excel_row}").value = status
                                                        tt.range(f"BL{excel_row}").value = LimitOrderBook[Order_id]['Remarks']
                                            elif _is_terminal_trigger(trade_info[2], "SELL") and LTP != 0:
                                                if trade_info[3] in ['True_Market' , 'True_Limit_LTP' , 'Limit_Below' , 'Limit_Above']:
                                                    if trade_info[16] != 'Active' and trade_info[16] != 'Entry_Pending' and trade_info[16] != 'Exit_Pending' and trade_info[16] != 'Closed' and (trade_info[16] is None or trade_info[16] == '' ):
                                                        if trade_info[3] == 'True_Market':
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", " + "SELL" + ", " + str(trade_info[3])+ ", " + str(trade_info[4])+ " triggered"
                                                            logger.info(Message)
                                                            order_id = place_trade(i, int(trade_info[0]), "SELL","LIMIT",LTP, trade_info[18], display_row=idx + 4)
                                                            if order_id is None:
                                                                tt.range(f"BQ{excel_row}").value = "Order_Rejected/Unconfirmed"
                                                            else:
                                                                register_pending_order(order_id)
                                                                tt.range(f"BF{excel_row}").value = None
                                                                tt.range(f"BG{excel_row}").value = order_id
                                                                tt.range(f"BJ{excel_row}").value = None
                                                                tt.range(f"BK{excel_row}").value = None
                                                                tt.range(f"BP{excel_row}").value = None
                                                                tt.range(f"BQ{excel_row}").value = 'Entry_Pending'
                                                                tt.range(f"BR{excel_row}").value = None
                                                        elif trade_info[3] == 'True_Limit_LTP':    
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", " + "SELL" + ", " + str(trade_info[3])+ ", " + str(trade_info[4])+ " triggered"
                                                            logger.info(Message)
                                                            order_id = place_trade(i, int(trade_info[0]), "SELL","LIMIT",LTP, trade_info[18], display_row=idx + 4)
                                                            if order_id is None:
                                                                tt.range(f"BQ{excel_row}").value = "Order_Rejected/Unconfirmed"
                                                            else:
                                                                register_pending_order(order_id)
                                                                tt.range(f"BF{excel_row}").value = None
                                                                tt.range(f"BG{excel_row}").value = order_id
                                                                tt.range(f"BQ{excel_row}").value = 'Entry_Pending'
                                                                tt.range(f"BJ{excel_row}").value = None
                                                                tt.range(f"BK{excel_row}").value = None
                                                                tt.range(f"BP{excel_row}").value = None
                                                                tt.range(f"BR{excel_row}").value = None
                                                        elif trade_info[3] == 'Limit_Above' and trade_info[4] is not None:
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", " + "SELL" + ", " + str(trade_info[3])+ ", " + str(trade_info[4])+ " triggered"
                                                            logger.info(Message)
                                                            order_id = place_trade(i, int(trade_info[0]), "SELL","LIMIT",trade_info[4], trade_info[18], display_row=idx + 4)
                                                            if order_id is None:
                                                                tt.range(f"BQ{excel_row}").value = "Order_Rejected/Unconfirmed"
                                                            else:
                                                                register_pending_order(order_id)
                                                                tt.range(f"BF{excel_row}").value = None
                                                                tt.range(f"BG{excel_row}").value = order_id
                                                                tt.range(f"BQ{excel_row}").value = 'Entry_Pending'
                                                                tt.range(f"BJ{excel_row}").value = None
                                                                tt.range(f"BK{excel_row}").value = None
                                                                tt.range(f"BP{excel_row}").value = None
                                                                tt.range(f"BR{excel_row}").value = None
                                                        elif trade_info[3] == 'Limit_Below' and trade_info[4] is not None:
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", " + "SELL" + ", " + str(trade_info[3])+ ", " + str(trade_info[4])+ " triggered"
                                                            logger.info(Message)
                                                            order_id = place_trade(i, int(trade_info[0]), "SELL","SL-M",trade_info[4], trade_info[18], display_row=idx + 4)
                                                            if order_id is None:
                                                                tt.range(f"BQ{excel_row}").value = "Order_Rejected/Unconfirmed"
                                                            else:
                                                                register_pending_order(order_id)
                                                                tt.range(f"BF{excel_row}").value = None
                                                                tt.range(f"BG{excel_row}").value = order_id
                                                                tt.range(f"BQ{excel_row}").value = 'Entry_Pending'
                                                                tt.range(f"BJ{excel_row}").value = None
                                                                tt.range(f"BK{excel_row}").value = None
                                                                tt.range(f"BP{excel_row}").value = None
                                                                tt.range(f"BR{excel_row}").value = None
                                                    elif trade_info[16] == 'Entry_Pending' :
                                                        Order_id = str(int(trade_info[6]))
                                                        status = LimitOrderBook[Order_id]['status']
                                                        if status == 'COMPLETE':
                                                            Executed_price = LimitOrderBook[Order_id]['Executed_price']
                                                            tt.range(f"BF{excel_row}").value = Executed_price
                                                            tt.range(f"BQ{excel_row}").value = 'Active'
                                                            Message = str(int(trade_info[0])) + " quantity " + str(i[4:]) + " Sold at " + str(Executed_price) 
                                                            Telegram_Message.append(Message)
                                                        tt.range(f"BH{excel_row}").value = LimitOrderBook[Order_id]['Remarks']
                                                    elif trade_info[16] == 'Active':
                                                        PnL = (float(trade_info[5]) - float(LTP) ) * int(trade_info[0])
                                                        tt.range(f"BR{excel_row}").value = PnL
                                                        TSL = 9999999
                                                        if type(trade_info[13]) is float:
                                                            if trade_info[14] == True :
                                                                if(LTP <= trade_info[5]):
                                                                    CUR_TSL = trade_info[13] - (trade_info[5] - LTP )
                                                                else:
                                                                    CUR_TSL = trade_info[13]
                                                                OLD_TSL = trade_info[15]
                                                                if OLD_TSL is not None:
                                                                    TSL = min(CUR_TSL,OLD_TSL)
                                                                else:
                                                                    TSL = CUR_TSL 
                                                            else:
                                                                TSL = trade_info[13]
                                                            tt.range(f"BP{excel_row}").value = TSL
                                                        if trade_info[8] == 'True_Market':
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", " + str(trade_info[8]) +  " exit triggered"
                                                            logger.info(Message)
                                                            order_id = place_trade(i, int(trade_info[0]), "BUY","LIMIT",LTP, trade_info[18], display_row=idx + 4)
                                                            if order_id is None:
                                                                tt.range(f"BI{excel_row}").value = None
                                                            else:  
                                                                register_pending_order(order_id)
                                                                tt.range(f"BK{excel_row}").value = order_id
                                                                tt.range(f"BQ{excel_row}").value = "Exit_Pending"
                                                        elif trade_info[8] == 'True_Limit_LTP':
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", " + str(trade_info[8]) +  " exit triggered"
                                                            logger.info(Message)
                                                            order_id = place_trade(i, int(trade_info[0]), "BUY","LIMIT",LTP, trade_info[18], display_row=idx + 4)
                                                            if order_id is None:
                                                                tt.range(f"BI{excel_row}").value = None
                                                            else:  
                                                                register_pending_order(order_id)
                                                                tt.range(f"BK{excel_row}").value = order_id
                                                                tt.range(f"BQ{excel_row}").value = 'Exit_Pending'    
                                                        elif type(trade_info[12]) is float and trade_info[12] >= LTP:
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", target meets so exit triggered"
                                                            logger.info(Message)
                                                            order_id = place_trade(i, int(trade_info[0]), "BUY","LIMIT",LTP, trade_info[18], display_row=idx + 4)
                                                            register_pending_order(order_id)
                                                            tt.range(f"BK{excel_row}").value = order_id
                                                            tt.range(f"BQ{excel_row}").value = 'Exit_Pending'
                                                        elif TSL <= LTP and type(trade_info[13]) is float:
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", TSL meets so exit triggered"
                                                            logger.info(Message)
                                                            order_id = place_trade(i, int(trade_info[0]), "BUY","LIMIT",LTP, trade_info[18], display_row=idx + 4)
                                                            if order_id is None:
                                                                tt.range(f"BI{excel_row}").value = None
                                                            else: 
                                                                register_pending_order(order_id)
                                                                tt.range(f"BK{excel_row}").value = order_id
                                                                tt.range(f"BQ{excel_row}").value = 'Exit_Pending'
                                                    elif trade_info[16] ==  'Exit_Pending':
                                                        Order_id = str(int(trade_info[10]))
                                                        status = LimitOrderBook[Order_id]['status']
                                                        if status == 'COMPLETE':
                                                            Executed_price = LimitOrderBook[Order_id]['Executed_price']
                                                            tt.range(f"BJ{excel_row}").value = Executed_price
                                                            tt.range(f"BQ{excel_row}").value = 'Closed'
                                                            PnL = (float(trade_info[5]) - float(Executed_price)) * int(trade_info[0])
                                                            tt.range(f"BR{excel_row}").value = PnL
                                                            Message = str(int(trade_info[0])) + " quantity " + str(i[4:]) + " buyback at " + str(Executed_price) + ", Trade profit or Loss = " + str(round(PnL,2))
                                                            Telegram_Message.append(Message)
                                                        tt.range(f"BL{excel_row}").value = LimitOrderBook[Order_id]['Remarks']
                            except Exception as e:
                                Message = "Exception occur in core order book management:" + str(e)
                                logger.info(Message)
                                pass
                        except Exception as e:
                            Message = "Exception occur while processing symbol row:" + str(e)
                            logger.info(Message)
                            pass
                main_list.append(lst)
                indicator_list.append(indicator_row)
                idx += 1

            tt.range("AG4:AN1000").value = indicator_list
            weekly_data_engine.write_to_excel(excel_TT)
            tt.range("AA4:AF1000").value = tt.range("BM4:BR1000").value
            sleep(0.15)
                
        except Exception as e:
            pass
        sleep(0.20)

def get_position():
    global api
    try:
        positions = api.get_positions()
        df_positions = pd.DataFrame(positions)
        mtm = 0
        pnl = 0
        day_m2m = 0
        if len(df_positions) > 0:
            df_positions = df_positions.sort_values(by=["tsym"]).reset_index(drop=True)
            for i in positions:
                mtm += float(i['urmtom'])
                pnl += float(i['rpnl'])
            day_m2m = mtm + pnl
            df_positions = df_positions[['exch','tsym','prd','netqty','netavgprc','lp','daybuyqty','daysellqty','daybuyavgprc','daysellavgprc','openbuyqty','opensellqty']]
            df_positions = df_positions.rename(columns={'exch':'Exchange' ,'tsym':'Symbol' ,'prd':'Product' ,'netqty':'Net Quantity' ,'netavgprc':'Avg Price' ,'lp':'Last Price' ,'daybuyqty':'Buy Quantity' ,'daysellqty':'Sell Quantity' ,'daybuyavgprc':'Avg Buy Price' ,'daysellavgprc':'Avg. Sell Price' ,'openbuyqty':'Open Buy Quantity' ,'opensellqty':'Open Sell Quantity'})
            df_positions['Net Quantity'] = df_positions['Net Quantity'].astype('int')
    except Exception as e:
        print(f"Error in get_position: {e}")
    return df_positions,day_m2m

def LoadInstrument_token(Token_4_Exchange = ['NSE','BSE','NFO', 'BFO','CDS','MCX']):
    global df_ins_NSE, df_ins_BSE, df_ins_NFO, df_ins_BFO,df_ins_CDS, df_ins_MCX
    global api
    try:
        subdir = "Instrument"
        if not os.path.exists(subdir):
            os.makedirs(subdir)
        print("Finvasia intrument token download started, may take upto 2-3 minutes ..")
        if 'NSE' in Token_4_Exchange:
            zip_file = "NSE_symbols.txt.zip"
            url = f"https://api.shoonya.com/{zip_file}"
            r = requests.get(f"{url}", allow_redirects=True)
            open(zip_file, "wb").write(r.content)
            df_ins_NSE = pd.read_csv(zip_file)
            os.remove(zip_file)
            try:
                df_ins_NSE.to_csv(os.path.join(subdir,"NSE_symbols.csv"),index = False)
            except:
                pass
        if 'BSE' in Token_4_Exchange:
            zip_file = "BSE_symbols.txt.zip"
            url = f"https://api.shoonya.com/{zip_file}"
            r = requests.get(f"{url}", allow_redirects=True)
            open(zip_file, "wb").write(r.content)
            df_ins_BSE = pd.read_csv(zip_file)
            os.remove(zip_file)
            try:
                df_ins_BSE.to_csv(os.path.join(subdir,"BSE_symbols.csv"),index = False)
            except:
                pass
        if 'NFO' in Token_4_Exchange:
            zip_file = "NFO_symbols.txt.zip"
            url = f"https://api.shoonya.com/{zip_file}"
            r = requests.get(f"{url}", allow_redirects=True)
            open(zip_file, "wb").write(r.content)
            df_ins_NFO = pd.read_csv(zip_file)
            df_ins_NFO['Expiry'] = pd.to_datetime(df_ins_NFO['Expiry']).apply(lambda x: x.date())
            df_ins_NFO = df_ins_NFO.sort_values(by=['Expiry',"Symbol",'StrikePrice'], ascending=[True,True,True])
            df_ins_NFO = df_ins_NFO.astype({"StrikePrice": str})
            os.remove(zip_file)
            try:
                df_ins_NFO.to_csv(os.path.join(subdir,"NFO_symbols.csv"),index = False)
            except:
                pass
        if 'BFO' in Token_4_Exchange:
            zip_file = "BFO_symbols.txt.zip"
            url = f"https://api.shoonya.com/{zip_file}"
            r = requests.get(f"{url}", allow_redirects=True)
            open(zip_file, "wb").write(r.content)
            df_ins_BFO = pd.read_csv(zip_file, usecols=["Exchange", "Token", "LotSize", "TradingSymbol", "Expiry", "Instrument", "OptionType", "StrikePrice", "TickSize"])
            bfo_symbols = np.where(
                df_ins_BFO['TradingSymbol'].str.contains('SENSEX50', regex=False), 'SENSEX50',
                df_ins_BFO['TradingSymbol'].str.extract(r'(.*?)(?:\d)', expand=False)
            )
            df_ins_BFO.insert(3, "Symbol", bfo_symbols)
            df_ins_BFO['Expiry'] = pd.to_datetime(df_ins_BFO['Expiry']).apply(lambda x: x.date())
            df_ins_BFO = df_ins_BFO.sort_values(by=['Expiry',"Symbol",'StrikePrice'], ascending=[True,True,True])
            df_ins_BFO = df_ins_BFO.astype({"StrikePrice": str})
            os.remove(zip_file)
            try:
                df_ins_BFO.to_csv(os.path.join(subdir,"BFO_symbols.csv"),index = False)
            except:
                pass
        if 'CDS' in Token_4_Exchange:
            zip_file = "CDS_symbols.txt.zip"
            url = f"https://api.shoonya.com/{zip_file}"
            r = requests.get(f"{url}", allow_redirects=True)
            open(zip_file, "wb").write(r.content)
            df_ins_CDS = pd.read_csv(zip_file)
            df_ins_CDS['Expiry'] = pd.to_datetime(df_ins_CDS['Expiry']).apply(lambda x: x.date())
            df_ins_CDS = df_ins_CDS.sort_values(by=['Instrument','Expiry',"Symbol",'StrikePrice'], ascending=[False,True,True,True])
            df_ins_CDS = df_ins_CDS.astype({"StrikePrice": str})
            os.remove(zip_file)
            try:
                df_ins_CDS.to_csv(os.path.join(subdir,"CDS_symbols.csv"),index = False)
            except:
                pass
        if 'MCX' in Token_4_Exchange:
            zip_file = "MCX_symbols.txt.zip"
            url = f"https://api.shoonya.com/{zip_file}"
            r = requests.get(f"{url}", allow_redirects=True)
            open(zip_file, "wb").write(r.content)
            df_ins_MCX = pd.read_csv(zip_file)
            df_ins_MCX['Expiry'] = pd.to_datetime(df_ins_MCX['Expiry']).apply(lambda x: x.date())
            df_ins_MCX = df_ins_MCX.sort_values(by=['Instrument','Expiry',"Symbol",'StrikePrice'], ascending=[False,True,True,True])
            df_ins_MCX = df_ins_MCX.astype({"StrikePrice": str})
            os.remove(zip_file)
            try:
                df_ins_MCX.to_csv(os.path.join(subdir,"MCX_symbols.csv"),index = False)
            except:
                pass
        print("Finvasia intrument token download completed")
    except Exception as e:
        print(f"Exception in LoadInstrument_token : {e}")
        pass

OPTION_CHAIN_HEADERS = [
    "CE OI", "CE Change OI", "CE Volume", "CE IV", "CE LTP",
    "CE Change %", "CE Bid Qty", "CE Bid", "CE Ask", "CE Ask Qty",
    "STRIKE",
    "PE Bid Qty", "PE Bid", "PE Ask", "PE Ask Qty", "PE Change %",
    "PE LTP", "PE IV", "PE Volume", "PE Change OI", "PE OI",
]

def prepare_option_chain_output(sheet):
    sheet.range("A2:AE2").value = [OPTION_CHAIN_HEADERS + [None] * 10]
    sheet.range("V2:AE500").clear_contents()
    sheet.range("V3:AE500").clear_contents()

def start_optionchain():
    global subscribe_symbol
    global live_data
    global df_ins_NFO
    global df_ins_NSE, df_ins_BSE, df_ins_NFO, df_ins_CDS, df_ins_MCX, df_ins_BFO
    global subs_lst
    excel_name = get_terminal_workbook()
    oci = excel_name.sheets("Option_Chain_Input")
    Option_Chain_Output = excel_name.sheets("Option_Chain_Output")
    prepare_option_chain_output(Option_Chain_Output)
    oci.range("F3").value = None
    oci.range("F4").value = None
    oci.range("a2:c500").value = None
    Option_Chain_Output.range('a3:ae500').value = None
    IterationSleep = 1
    oci.range("d2").value = "Segment==>>"
    oci.range("d3").value, oci.range("d4").value = "Symbol==>>", "Expiry==>>",
    oci.range("d5").value, oci.range("d6").value = "RefreshRate==>>", "NoOfStrike==>>",
    oci.range("d7").value, oci.range("d8").value = "ExpiryType==>>" , "GreekMatch==>>"
    pre_selected_segment = pre_selected_symbol = pre_selected_expiry = ""
    pre_selected_NoOfStrike = 0
    while True:
        try:
            excel_name = get_terminal_workbook()
            oci = excel_name.sheets("Option_Chain_Input")
            Option_Chain_Output = excel_name.sheets("Option_Chain_Output")
            selected_segment = str(oci.range("E2").value or "").strip()
            Exchange = selected_segment[:3]
            if Exchange not in ['NFO','CDS','MCX', 'BFO']:
                Exchange = 'NFO'
            oci.range("F2").value = None
            if Exchange == 'NFO':
                df_instrument = df_ins_NFO
            elif Exchange == 'CDS':
                df_instrument = df_ins_CDS
            elif Exchange == 'MCX':
                df_instrument = df_ins_MCX
            elif Exchange == 'BFO':
                df_instrument = df_ins_BFO
            else:
                oci.range("F2").value = "Please select correct segment"
            if pre_selected_segment != selected_segment:
                if Exchange == 'NFO':
                    df_symbol = df_ins_NFO
                if Exchange == 'BFO':
                    df_symbol = df_ins_BFO
                if Exchange == 'CDS':
                    df_symbol = df_ins_CDS[df_ins_CDS['Instrument'] == 'UNDCUR']
                if Exchange == 'MCX':
                    df_symbol = df_ins_MCX[df_ins_MCX['Instrument'] == 'OPTFUT']
                df_symbol = df_symbol.drop_duplicates( "Symbol" , keep='first')
                df_symbol =df_symbol[['Symbol']]
                oci.range("a2:c500").value = None
                oci.range("a1").options(index=False, header=True).value = df_symbol
                pre_selected_segment = selected_segment
            input_symbol = str(oci.range("E3").value or "").strip()
            if input_symbol:
                df_instrument_temp = df_instrument[(df_instrument.Symbol == input_symbol) & (df_instrument['OptionType'].isin(['CE','PE']))]
                if len(df_instrument_temp) > 0:
                    if pre_selected_symbol != input_symbol:
                        df_exp = df_instrument_temp.drop_duplicates( "Expiry" , keep='first')
                        df_exp =df_exp[['Expiry','LotSize']]
                        df_exp.sort_values(by = 'Expiry',inplace = True)
                        oci.range("b2:C100").value = None
                        oci.range("b2").options(index=False, header=False).value = df_exp
                        pre_selected_symbol = input_symbol
                    oci.range("F3").value = None
                    expiry_input = normalize_expiry_input(oci.range("E4").value)
                    valid_expiry = resolve_option_expiry(df_instrument, input_symbol, expiry_input)
                    if valid_expiry is not None and valid_expiry != expiry_input:
                        expiry_input = valid_expiry
                        oci.range("E4").value = expiry_input
                        print(
                            f"Option Chain: invalid/blank expiry corrected for "
                            f"{input_symbol}; using {expiry_input.strftime('%d-%m-%Y')}"
                        )
                    if expiry_input != None:
                        df_instrument_temp = df_instrument_temp[df_instrument_temp.Expiry == expiry_input]
                        if len(df_instrument_temp) > 0:
                            oci.range("F4").value = None
                            try:
                                IterationSleep = oci.range("E5").value
                                if IterationSleep != None:
                                    IterationSleep = int(IterationSleep)
                                oci.range("F5").value = None
                            except:
                                IterationSleep = 5
                                Message = "Iteration should be number"
                                print(Message)
                                oci.range("F5").value = Message
                            symbol_in_template = [itr for itr in OptionChain_template if itr["symbol"] ==input_symbol ]
                            if len(symbol_in_template) == 0:
                                ind_df_specific_symbol = df_instrument[(df_instrument.Symbol == input_symbol) & (df_instrument['OptionType'].isin(['CE','PE']))]
                                List_of_Expiry = ind_df_specific_symbol["Expiry"].unique()
                                Expiry_strikelist_list = []
                                for expiry in List_of_Expiry:
                                    if str( expiry ) != 'NaT' :
                                        ind_df_specific_symbol_expiry = ind_df_specific_symbol[(ind_df_specific_symbol.Expiry == expiry) ]
                                        List_of_strikes = ind_df_specific_symbol_expiry["StrikePrice"].unique()
                                        LotSize = ind_df_specific_symbol_expiry.iloc[0]["LotSize"]
                                        strike_pe_ce_list = []
                                        for strike in List_of_strikes:
                                            ind_df_specific_symbol_expiry_pe = ind_df_specific_symbol_expiry[
                                                (ind_df_specific_symbol_expiry.StrikePrice == str(strike)) & (ind_df_specific_symbol_expiry.OptionType == "PE") ]
                                            if len(ind_df_specific_symbol_expiry_pe) > 0:
                                                pe_token = ind_df_specific_symbol_expiry_pe.iloc[0]["Token"]
                                            else:
                                                pe_token = "NA"
                                            ind_df_specific_symbol_expiry_ce = ind_df_specific_symbol_expiry[(            ind_df_specific_symbol_expiry.StrikePrice== str(strike))& (            ind_df_specific_symbol_expiry.OptionType == "CE")]
                                            if len(ind_df_specific_symbol_expiry_ce) > 0:
                                                ce_token = ind_df_specific_symbol_expiry_ce.iloc[0]["Token"]
                                            else:
                                                ce_token = "NA"
                                            strike_pe_ce_dictionary = dict(
                                                {
                                                    "strike": strike,
                                                    "PE_Token": pe_token,
                                                    "CE_Token": ce_token,
                                                }
                                            )
                                            strike_pe_ce_list.append(strike_pe_ce_dictionary)
                                        expiry_stikelist_dict = dict(
                                            {
                                                "Expiry": expiry,
                                                "LotSize": LotSize,
                                                "Strike_list": strike_pe_ce_list,
                                            }
                                        )
                                        Expiry_strikelist_list.append(expiry_stikelist_dict)
                                    else:
                                        print(f"Ignoring the wrong date : {expiry}")
                                Final_dic = dict(
                                    {
                                        "symbol": input_symbol,
                                        "Expiry_Strike_token": Expiry_strikelist_list,
                                    }
                                )
                                OptionChain_template.append(Final_dic)
                            else:
                                Message = (
                                    "symbol already available in option chain template"
                                )
                            List_of_Expiry_Strike_token = [
                                itr["Expiry_Strike_token"]
                                for itr in OptionChain_template
                                if itr["symbol"] == input_symbol
                            ][0]
                            for expiry_strike in List_of_Expiry_Strike_token:
                                if input_symbol not in subs_lst:
                                    List_of_particular_expiry_strike = [
                                        itr["Strike_list"]
                                        for itr in List_of_Expiry_Strike_token
                                        if itr["Expiry"] == expiry_strike["Expiry"]
                                    ][0]
                                    for strike_dict in List_of_particular_expiry_strike:
                                        print(f"Going to subscribe {input_symbol} strike {strike_dict.get('strike')}")
                                        PE_Token = strike_dict.get("PE_Token")
                                        CE_Token = strike_dict.get("CE_Token")
                                        if PE_Token != "NA":
                                            subscribe_new_token(Exchange, PE_Token)
                                        if CE_Token != "NA":
                                            subscribe_new_token(Exchange, CE_Token)
                                    if input_symbol not in subs_lst:
                                        subs_lst.append(input_symbol)
                                        print(f"{input_symbol} subscription completed")
                            pd_oc = pd.DataFrame(columns=[ "CE_token", "CE_oi",  "CE_poi", "CE_toi", "CE_lp", "CE_pc", "CE_bq1","CE_bp1", "CE_sq1", "CE_sp1",  "strike", "PE_bq1", "PE_bp1", "PE_sq1", "PE_sp1", "PE_pc",  "PE_lp",  "PE_toi", "PE_poi", "PE_oi", "PE_token","CE_coi","PE_coi","CE_v","PE_v"] )
                            List_of_Expiry_Strike_token = [
                                itr["Expiry_Strike_token"]
                                for itr in OptionChain_template
                                if itr["symbol"] == input_symbol
                            ][0]
                            try:
                                Lot_Size = [
                                    itr["LotSize"]
                                    for itr in List_of_Expiry_Strike_token
                                    if itr["Expiry"] == expiry_input
                                ][0]
                                List_of_particular_expiry_strike = [
                                    itr["Strike_list"]
                                    for itr in List_of_Expiry_Strike_token
                                    if itr["Expiry"] == expiry_input
                                ][0]
                                isFound, Fut_Token = GetToken(Exchange,input_symbol)
                                if Exchange == 'NFO':
                                    isFound, Spot_Token = GetToken('NSE',input_symbol)
                                    spot_ltp = convert_to_float(api.get_quotes("NSE", str(Spot_Token)).get("lp"))
                                elif Exchange == 'BFO':
                                    isFound, Spot_Token = GetToken('BSE',input_symbol)
                                    spot_ltp = convert_to_float(api.get_quotes("BSE", str(Spot_Token)).get("lp"))
                                else:    
                                    Spot_Token = Fut_Token
                                    spot_ltp = convert_to_float(api.get_quotes(Exchange, str(Spot_Token)).get("lp"))
                                future_ltp = convert_to_float(api.get_quotes(Exchange, str(Fut_Token)).get("lp"))
                                for strike_dict in List_of_particular_expiry_strike:
                                    Strike = convert_to_float(strike_dict.get("strike"))
                                    PE_Token = strike_dict.get("PE_Token")
                                    PE_Token = str(Exchange)+ "|"  + str(PE_Token)
                                    CE_Token = strike_dict.get("CE_Token")
                                    CE_Token =  str(Exchange)+ "|" + str(CE_Token)
                                    try:
                                        CE_oi = live_data[str(CE_Token)].get("oi", 0)
                                    except:
                                        CE_oi = 0
                                    try:
                                        CE_poi = live_data[str(CE_Token)].get("poi", 0)
                                    except:
                                        CE_poi = 0
                                    CE_coi = int(CE_oi) - int(CE_poi)
                                    try:
                                        CE_toi = live_data[str(CE_Token)].get("toi", "-")
                                    except:
                                        CE_toi = "-"
                                    try:
                                        CE_lp = live_data[str(CE_Token)].get("lp", 0)
                                    except:
                                        CE_lp = 0
                                    try:
                                        CE_pc = live_data[str(CE_Token)].get("pc", "-")
                                    except:
                                        CE_pc = "-"
                                    try:
                                        CE_bq1 = live_data[str(CE_Token)].get("bq1", "-")
                                    except:
                                        CE_bq1 = "-"
                                    try:
                                        CE_bp1 = live_data[str(CE_Token)].get("bp1", "-")
                                    except:
                                        CE_bp1 = "-"
                                    try:
                                        CE_sq1 = live_data[str(CE_Token)].get("sq1", "-")
                                    except:
                                        CE_sq1 = "-"
                                    try:
                                        CE_sp1 = live_data[str(CE_Token)].get("sp1", "-")
                                    except:
                                        CE_sp1 = "-"
                                    try:
                                        PE_oi = live_data[str(PE_Token)].get("oi", 0)
                                    except:
                                        PE_oi = 0
                                    try:
                                        PE_poi = live_data[str(PE_Token)].get("poi", 0)
                                    except:
                                        PE_poi = 0
                                    PE_coi = int(PE_oi) - int(PE_poi)
                                    try:
                                        PE_toi = live_data[str(PE_Token)].get("toi", "-")
                                    except:
                                        PE_toi = "-"
                                    try:
                                        PE_lp = live_data[str(PE_Token)].get("lp", 0)
                                    except:
                                        PE_lp = 0
                                    try:
                                        PE_pc = live_data[str(PE_Token)].get("pc", "-")
                                    except:
                                        PE_pc = "-"
                                    try:
                                        PE_bq1 = live_data[str(PE_Token)].get("bq1", "-")
                                    except:
                                        PE_bq1 = "-"
                                    try:
                                        PE_bp1 = live_data[str(PE_Token)].get("bp1", "-")
                                    except:
                                        PE_bp1 = "-"
                                    try:
                                        PE_sq1 = live_data[str(PE_Token)].get("sq1", "-")
                                    except:
                                        PE_sq1 = "-"
                                    try:
                                        PE_sp1 = live_data[str(PE_Token)].get("sp1", "-")
                                    except:
                                        PE_sp1 = "-"
                                    try:
                                        CE_v = live_data[str(CE_Token)].get("v", 0)
                                    except:
                                        CE_v = 0
                                    try:
                                        PE_v = live_data[str(PE_Token)].get("v", 0)
                                    except:
                                        PE_v = 0
                                    dic_data = {
                                            "CE_token": CE_Token,
                                            "CE_oi": CE_oi,
                                            "CE_poi": CE_poi,
                                            "CE_toi": CE_toi,
                                            "CE_lp": CE_lp,
                                            "CE_pc": CE_pc,
                                            "CE_bq1": CE_bq1,
                                            "CE_bp1": CE_bp1,
                                            "CE_sq1": CE_sq1,
                                            "CE_sp1": CE_sp1,
                                            "strike": Strike,
                                            "PE_bq1": PE_bq1,
                                            "PE_bp1": PE_bp1,
                                            "PE_sq1": PE_sq1,
                                            "PE_sp1": PE_sp1,
                                            "PE_pc": PE_pc,
                                            "PE_lp": PE_lp,
                                            "PE_toi": PE_toi,
                                            "PE_poi": PE_poi,
                                            "PE_oi": PE_oi,
                                            "PE_token": PE_Token,
                                            "CE_coi":CE_coi,
                                            "PE_coi":PE_coi,
                                            "CE_v":CE_v,
                                            "PE_v":PE_v,
                                        }
                                    pd_oc = pd.concat([pd_oc, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                pd_oc = pd_oc.astype({"strike": float})
                                pd_oc = pd_oc.sort_values(by="strike", ascending=True)
                                pd_oc = pd_oc.fillna(0)
                                pd_oc = pd_oc.astype({"CE_oi": int})
                                pd_oc = pd_oc.astype({"PE_oi": int})
                                pd_oc = pd_oc.astype({"CE_poi": int})
                                pd_oc = pd_oc.astype({"PE_poi": int})
                                pd_oc = pd_oc.astype({"CE_coi": int})
                                pd_oc = pd_oc.astype({"PE_coi": int})
                                pd_oc = pd_oc.astype({"CE_v": int})
                                pd_oc = pd_oc.astype({"PE_v": int})
                                pd_oc["CE_v"] = pd_oc["CE_v"] / int(Lot_Size)
                                pd_oc["PE_v"] = pd_oc["PE_v"] / int(Lot_Size)
                                pd_oc["CE_oi"] = pd_oc["CE_oi"] / int(Lot_Size)
                                pd_oc["PE_oi"] = pd_oc["PE_oi"] / int(Lot_Size)
                                pd_oc["CE_poi"] = pd_oc["CE_poi"] / int(Lot_Size)
                                pd_oc["PE_poi"] = pd_oc["PE_poi"] / int(Lot_Size)
                                pd_oc["CE_coi"] = pd_oc["CE_coi"] / int(Lot_Size)
                                pd_oc["PE_coi"] = pd_oc["PE_coi"] / int(Lot_Size)
                                pd_oc["OI_SUM"] = pd_oc["CE_oi"] + pd_oc["PE_oi"]
                                df_oc_pro = pd_oc
                                try:
                                    NoOfStrike = int(oci.range("E6").value)
                                except Exception as e:
                                    NoOfStrike = 100
                                df_oc_pro['strike_diff'] = abs(df_oc_pro['strike'] - spot_ltp)
                                df_oc_pro.sort_values(by = 'strike_diff',inplace = True)
                                AtmStrike = convert_to_float(df_oc_pro.iloc[0]['strike'])
                                AtmStrikeCallPrice = convert_to_float(df_oc_pro.iloc[0]['CE_lp'])
                                AtmStrikePutPrice = convert_to_float(df_oc_pro.iloc[0]['PE_lp'])
                                Future_LTP = future_ltp
                                Max_Pain_at_Strike = df_oc_pro[df_oc_pro.OI_SUM == df_oc_pro["OI_SUM"].max()].iloc[0]['strike']
                                Ltp_at_Max_Pain_CE = df_oc_pro[df_oc_pro.OI_SUM == df_oc_pro["OI_SUM"].max()].iloc[0]['CE_lp']
                                Ltp_at_Max_Pain_PE = df_oc_pro[df_oc_pro.OI_SUM == df_oc_pro["OI_SUM"].max()].iloc[0]['PE_lp']
                                LTP_at_ATM_CE = df_oc_pro[df_oc_pro.strike == AtmStrike].iloc[0]['CE_lp']
                                LTP_at_ATM_PE = df_oc_pro[df_oc_pro.strike == AtmStrike].iloc[0]['PE_lp']
                                Total_OI_CE = df_oc_pro["CE_oi"].sum()
                                Total_OI_PE = df_oc_pro["PE_oi"].sum()
                                Max_OI_CE = df_oc_pro["CE_oi"].max()
                                Max_OI_PE = df_oc_pro["PE_oi"].max()
                                Max_OI_at_Strike_CE = df_oc_pro[df_oc_pro.CE_oi == df_oc_pro["CE_oi"].max()].iloc[0]['strike']
                                Max_OI_at_Strike_PE = df_oc_pro[df_oc_pro.PE_oi == df_oc_pro["PE_oi"].max()].iloc[0]['strike']
                                LTP_of_Max_OI_Strike_CE = df_oc_pro[df_oc_pro.CE_oi == df_oc_pro["CE_oi"].max()].iloc[0]['CE_lp']
                                LTP_of_Max_OI_Strike_PE = df_oc_pro[df_oc_pro.PE_oi == df_oc_pro["PE_oi"].max()].iloc[0]['PE_lp']
                                Total_OI_Change_CE = df_oc_pro["CE_coi"].sum()
                                Total_OI_Change_PE = df_oc_pro["PE_coi"].sum()
                                Max_Change_in_OI_addition_CE = df_oc_pro["CE_coi"].max()
                                Max_Change_in_OI_addition_PE = df_oc_pro["PE_coi"].max()
                                Max_OI_addition_at_Srike_CE = df_oc_pro[df_oc_pro["CE_coi"] == df_oc_pro["CE_coi"].max()].iloc[0]['strike']
                                Max_OI_addition_at_Srike_PE = df_oc_pro[df_oc_pro["PE_coi"] == df_oc_pro["PE_coi"].max()].iloc[0]['strike']
                                LTP_of_Max_OI_addition_Strike_CE = df_oc_pro[df_oc_pro["CE_coi"] == df_oc_pro["CE_coi"].max()].iloc[0]['CE_lp']
                                LTP_of_Max_OI_addition_Strike_PE = df_oc_pro[df_oc_pro["PE_coi"] == df_oc_pro["PE_coi"].max()].iloc[0]['PE_lp']
                                Max_Change_in_OI_unwinding_CE = -1 * int(df_oc_pro["CE_coi"].min())
                                Max_Change_in_OI_unwinding_PE = -1 * int(df_oc_pro["PE_coi"].min())
                                Max_OI_unwinding_at_Srike_CE = df_oc_pro[df_oc_pro["CE_coi"] == df_oc_pro["CE_coi"].min()].iloc[0]['strike']
                                Max_OI_unwinding_at_Srike_PE = df_oc_pro[df_oc_pro["PE_coi"] == df_oc_pro["PE_coi"].min()].iloc[0]['strike']
                                LTP_of_Max_OI_unwinding_Strike_CE = df_oc_pro[df_oc_pro["CE_coi"] == df_oc_pro["CE_coi"].min()].iloc[0]['CE_lp']
                                LTP_of_Max_OI_unwinding_Strike_PE = df_oc_pro[df_oc_pro["PE_coi"] == df_oc_pro["PE_coi"].min()].iloc[0]['PE_lp']
                                df_additional_detail = pd.DataFrame(columns = ['CE','PE'])
                                dic_data = {'CE':Future_LTP}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                dic_data = {'CE':Max_Pain_at_Strike}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                dic_data = {'CE':Ltp_at_Max_Pain_CE,'PE': Ltp_at_Max_Pain_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                dic_data = {'CE':AtmStrike}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                dic_data = {'CE':LTP_at_ATM_CE,'PE':LTP_at_ATM_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                dic_data = {'CE':Total_OI_CE, 'PE':Total_OI_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                dic_data = {'CE':Total_OI_Change_CE, 'PE':Total_OI_Change_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                dic_data = {'CE':Max_OI_CE,  'PE':Max_OI_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                dic_data = {'CE':Max_OI_at_Strike_CE,'PE':Max_OI_at_Strike_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                dic_data = {'CE':LTP_of_Max_OI_Strike_CE,'PE':LTP_of_Max_OI_Strike_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                dic_data = {'CE':Max_Change_in_OI_addition_CE,'PE':Max_Change_in_OI_addition_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                dic_data = {'CE':Max_OI_addition_at_Srike_CE,'PE':Max_OI_addition_at_Srike_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                dic_data = {'CE':LTP_of_Max_OI_addition_Strike_CE, 'PE':LTP_of_Max_OI_addition_Strike_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                dic_data = {'CE':Max_Change_in_OI_unwinding_CE, 'PE':Max_Change_in_OI_unwinding_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                dic_data = {'CE':Max_OI_unwinding_at_Srike_CE, 'PE':Max_OI_unwinding_at_Srike_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                dic_data = {'CE':LTP_of_Max_OI_unwinding_Strike_CE,'PE':LTP_of_Max_OI_unwinding_Strike_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                oci.range("i3").options(index=False, header=False).value = df_additional_detail
                                df_oc_pro = df_oc_pro[:2*int(NoOfStrike)+1]
                                df_oc_pro.sort_values(by ='strike',inplace = True)
                                df_oc_pro['CE_IV'] = None
                                df_oc_pro['PE_IV'] = None
                                df_oc_pro = df_oc_pro.reindex(['CE_oi','CE_coi','CE_v','CE_IV','CE_lp','CE_pc','CE_bq1','CE_bp1','CE_sp1','CE_sq1','strike','PE_bq1','PE_bp1','PE_sp1','PE_sq1','PE_pc','PE_lp','PE_IV','PE_v','PE_coi','PE_oi'], axis=1)
                                SpotPrice = convert_to_float(spot_ltp)
                                FuturePrice = convert_to_float(future_ltp)
                                ExpiryDateTime = dt(expiry_input.year, expiry_input.month, expiry_input.day, 0, 0, 0)
                                ExpiryType = oci.range("F7").value
                                GreekMatch = oci.range("F8").value
                                if ExpiryType == 'WEEKLY':
                                    ExpiryDateType = ExpType.WEEKLY
                                else:
                                    ExpiryDateType = ExpType.MONTHLY
                                FromDateTime = dt.now() 
                                if Exchange == 'NFO':
                                    if dt.now().time() > time(15, 30, 0):
                                        FromDateTime = dt(dt.now().year, dt.now().month,dt.now().day, 15, 30, 0)
                                if GreekMatch == "SENSIBULL":
                                    tryMatchWith=TryMatchWith.SENSIBULL
                                else:
                                    tryMatchWith=TryMatchWith.NSE
                                dayCountType = DayCountType.CALENDARDAYS
                                IvGreeks = CalcIvGreeks( SpotPrice = SpotPrice,  FuturePrice = FuturePrice, AtmStrike = AtmStrike, AtmStrikeCallPrice = AtmStrikeCallPrice, AtmStrikePutPrice = AtmStrikePutPrice, ExpiryDateTime = ExpiryDateTime, ExpiryDateType = ExpiryDateType, FromDateTime = FromDateTime, tryMatchWith = tryMatchWith, dayCountType = dayCountType)
                                for ind in df_oc_pro.index:
                                    StrikePrice= convert_to_float(df_oc_pro['strike'][ind])
                                    StrikeCallPrice= convert_to_float(df_oc_pro['CE_lp'][ind])
                                    StrikePutPrice= convert_to_float(df_oc_pro['PE_lp'][ind])
                                    Greeks = IvGreeks.GetImpVolAndGreeks( StrikePrice = StrikePrice, StrikeCallPrice = StrikeCallPrice, StrikePutPrice = StrikePutPrice)
                                    if GreekMatch == "NSE":
                                        df_oc_pro.at[ind, "CE_IV"] = round(Greeks["CallIV"],2)
                                        df_oc_pro.at[ind, "PE_IV"] = round(Greeks["PutIV"],2)
                                    else:
                                        df_oc_pro.at[ind, "CE_IV"] = round(Greeks["ImplVol"],2)
                                        df_oc_pro.at[ind, "PE_IV"] = round(Greeks["ImplVol"],2)
                                del IvGreeks
                                df_oc_pro.round({"CE_lp":2, 'PE_lp':2})
                                if pre_selected_NoOfStrike != NoOfStrike:
                                    pre_selected_NoOfStrike = NoOfStrike
                                    Option_Chain_Output.range('a3:ae500').value = None
                                    Option_Chain_Output.range("a3:U500").color = (255,255,255)
                                df_oc_pro = df_oc_pro.reset_index(drop = True)
                                ATM_pos = df_oc_pro[df_oc_pro.strike == AtmStrike].index.values[0]
                                if NoOfStrike * 2 < len(df_oc_pro):
                                    df_oc_pro = df_oc_pro.iloc[ATM_pos - int(NoOfStrike) : ATM_pos + int(NoOfStrike) + 1]
                                    df_oc_pro = df_oc_pro.reset_index(drop=True)
                                    ATM_Row = int(NoOfStrike) + 3
                                    Option_Chain_Output.range(f"a{ATM_Row}:U{ATM_Row}").color = (46,132,198)
                                Option_Chain_Output.range('a3').options(index=False,header=False).value = df_oc_pro
                                Option_Chain_Output.range(f"k3:k{len(df_oc_pro) + 2}").color = (242,242,242)
                                for color_index, color_row in df_oc_pro.iterrows():
                                    excel_row = int(color_index) + 3
                                    strike_value = convert_to_float(color_row["strike"])
                                    if strike_value < SpotPrice:
                                        Option_Chain_Output.range(f"a{excel_row}:j{excel_row}").color = (221,235,247)
                                    elif strike_value > SpotPrice:
                                        Option_Chain_Output.range(f"l{excel_row}:u{excel_row}").color = (221,235,247)
                            except Exception as e:
                                Option_Chain_Output.range('a3:ae500').value = None
                                oci.range('I3:J18').value = None
                                Message = "Please check the all provided detail to load option chain:" + str(e)
                                print(Message)
                                oci.range("F4").value = Message
                        else:
                            Option_Chain_Output.range('a3:ae500').value = None
                            oci.range('I3:J18').value = None
                            Message = expiry_not_found_message(df_instrument, input_symbol, expiry_input)
                            print(Message)
                            oci.range("F4").value = Message
                    else:
                        Option_Chain_Output.range('a3:ae500').value = None
                        oci.range('I3:J18').value = None
                        Message = "Please enter the expiry"
                        print(Message)
                        oci.range("F4").value = Message
                else:
                    Option_Chain_Output.range('a3:ae500').value = None
                    oci.range('I3:J18').value = None
                    Message = "Please enter correct symbol"
                    print(Message)
                    oci.range("F3").value = Message
                    oci.range("F4").value = None
                    oci.range("b2:c100").value = None
            else:
                Option_Chain_Output.range('a3:ae500').value = None
                oci.range('I3:J18').value = None
                Message = "Please enter the symbol"
                print(Message)
                oci.range("F3").value = Message
                oci.range("F4").value = None
                oci.range("b2:c100").value = None
        except Exception as e:
            print(f"Excption : {e}")
            pass
        sleep(int(IterationSleep))

def GetToken(Exchange,Symbol, Type = 'FUT' ,Expiry=None,Strike = None):
    global df_ins_NSE, df_ins_BSE, df_ins_NFO, df_ins_CDS, df_ins_MCX, df_ins_BFO
    try:
        Token = None
        isTokenFound = False
        Symbol = Symbol.upper()
        if Exchange == 'NSE':
            if Symbol in IndexList:
                Token = Token_list[Symbol]
                isTokenFound = True
            else:
                df_ins_temp = df_ins_NSE[df_ins_NSE.Symbol == Symbol]
                if len(df_ins_temp) > 0:
                    Token = df_ins_temp.iloc[0]['Token']
                    isTokenFound = True
        elif Exchange == 'BSE':
            if Symbol in IndexList:
                Token = Token_list[Symbol]
                isTokenFound = True
            else:
                df_ins_temp = df_ins_BSE[df_ins_BSE.Symbol == Symbol]
                if len(df_ins_temp) > 0:
                    Token = df_ins_temp.iloc[0]['Token']
                    isTokenFound = True
        elif Exchange == 'NFO':
            df_ins_temp = df_ins_NFO[(df_ins_NFO.Symbol == Symbol) & (df_ins_NFO['Instrument'].isin(['FUTIDX','FUTSTK']))]
            if len(df_ins_temp) > 0:
                df_ins_temp.sort_values(by = 'Expiry',inplace = True)
                Token = df_ins_temp.iloc[0]['Token']
                isTokenFound = True
        elif Exchange == 'BFO':
            df_ins_temp = df_ins_BFO[(df_ins_BFO.Symbol == Symbol) & (df_ins_BFO['Instrument'].isin(['FUTIDX','FUTSTK']))]
            if len(df_ins_temp) > 0:
                df_ins_temp.sort_values(by = 'Expiry',inplace = True)
                Token = df_ins_temp.iloc[0]['Token']
                isTokenFound = True
        elif Exchange == 'CDS':
            df_ins_temp = df_ins_CDS[(df_ins_CDS.Symbol == Symbol) & (df_ins_CDS.Instrument == 'FUTCUR')]
            if len(df_ins_temp) > 0:
                df_ins_temp.sort_values(by = 'Expiry',inplace = True)
                Token = df_ins_temp.iloc[0]['Token']
                isTokenFound = True
        elif Exchange == 'MCX':
            df_ins_temp = df_ins_MCX[(df_ins_MCX.Symbol == Symbol) & (df_ins_MCX.Instrument == 'FUTCOM')]
            if len(df_ins_temp) > 0:
                df_ins_temp.sort_values(by = 'Expiry',inplace = True)
                Token = df_ins_temp.iloc[0]['Token']
                isTokenFound = True
    except Exception as e:
        print(f"Exception occur in GetToken : {e}")
    return isTokenFound, Token

def start_optionchain_Pro():
    global subscribe_symbol
    global live_data
    global df_ins_NFO
    global df_ins_NSE, df_ins_BSE, df_ins_NFO, df_ins_CDS, df_ins_MCX, df_ins_BFO
    global subs_lst
    excel_name = get_terminal_workbook()
    oci_pro = excel_name.sheets("Option_Chain_Pro_Input")
    Option_Chain_Pro_Output = excel_name.sheets("Option_Chain_Pro_Output")
    prepare_option_chain_output(Option_Chain_Pro_Output)
    oci_pro.range("F3").value = None
    oci_pro.range("F4").value = None
    oci_pro.range("a2:c500").value = None
    Option_Chain_Pro_Output.range('a3:ae500').value = None
    IterationSleep = 1
    oci_pro.range("d2").value = "Segment==>>"
    oci_pro.range("d3").value, oci_pro.range("d4").value = "Symbol==>>", "Expiry==>>",
    oci_pro.range("d5").value, oci_pro.range("d6").value = "RefreshRate==>>", "NoOfStrike==>>",
    oci_pro.range("d7").value, oci_pro.range("d8").value = "ExpiryType==>>" , "GreekMatch==>>"
    pre_selected_segment = pre_selected_symbol = pre_selected_expiry = ""
    pre_selected_NoOfStrike = 0
    while True:
        try:
            excel_name = get_terminal_workbook()
            oci_pro = excel_name.sheets("Option_Chain_Pro_Input")
            Option_Chain_Pro_Output = excel_name.sheets("Option_Chain_Pro_Output")
            selected_segment = str(oci_pro.range("E2").value or "").strip()
            Exchange = selected_segment[:3]
            if Exchange not in ['NFO','CDS','MCX','BFO']:
                Exchange = 'NFO'
            oci_pro.range("F2").value = None
            if Exchange == 'NFO':
                df_instrument = df_ins_NFO
            elif Exchange == 'CDS':
                df_instrument = df_ins_CDS
            elif Exchange == 'MCX':
                df_instrument = df_ins_MCX
            elif Exchange == 'BFO':
                df_instrument = df_ins_BFO
            else:
                oci_pro.range("F2").value = "Please select correct segment"
            if pre_selected_segment != selected_segment:
                if Exchange == 'NFO':
                    df_symbol = df_ins_NFO
                if Exchange == 'BFO':
                    df_symbol = df_ins_BFO
                if Exchange == 'CDS':
                    df_symbol = df_ins_CDS[df_ins_CDS['Instrument'] == 'UNDCUR']
                if Exchange == 'MCX':
                    df_symbol = df_ins_MCX[df_ins_MCX['Instrument'] == 'OPTFUT']
                df_symbol = df_symbol.drop_duplicates( "Symbol" , keep='first')
                df_symbol =df_symbol[['Symbol']]
                oci_pro.range("a2:c500").value = None
                oci_pro.range("a1").options(index=False, header=True).value = df_symbol
                pre_selected_segment = selected_segment
            input_symbol = str(oci_pro.range("E3").value or "").strip()
            if input_symbol:
                df_instrument_temp = df_instrument[(df_instrument.Symbol == input_symbol) & (df_instrument['OptionType'].isin(['CE','PE']))]
                if len(df_instrument_temp) > 0:
                    if pre_selected_symbol != input_symbol:
                        df_exp = df_instrument_temp.drop_duplicates( "Expiry" , keep='first')
                        df_exp =df_exp[['Expiry','LotSize']]
                        df_exp.sort_values(by = 'Expiry',inplace = True)
                        oci_pro.range("b2:C100").value = None
                        oci_pro.range("b2").options(index=False, header=False).value = df_exp
                        pre_selected_symbol = input_symbol
                    oci_pro.range("F3").value = None
                    expiry_input = normalize_expiry_input(oci_pro.range("E4").value)
                    valid_expiry = resolve_option_expiry(df_instrument, input_symbol, expiry_input)
                    if valid_expiry is not None and valid_expiry != expiry_input:
                        expiry_input = valid_expiry
                        oci_pro.range("E4").value = expiry_input
                        print(
                            f"PRO Option Chain: invalid/blank expiry corrected for "
                            f"{input_symbol}; using {expiry_input.strftime('%d-%m-%Y')}"
                        )
                    if expiry_input != None:
                        df_instrument_temp = df_instrument_temp[df_instrument_temp.Expiry == expiry_input]
                        if len(df_instrument_temp) > 0:
                            oci_pro.range("F4").value = None
                            try:
                                IterationSleep = oci_pro.range("E5").value
                                if IterationSleep != None:
                                    IterationSleep = int(IterationSleep)
                                oci_pro.range("F5").value = None
                            except:
                                IterationSleep = 5
                                Message = "Iteration should be number"
                                print(Message)
                                oci_pro.range("F5").value = Message
                            symbol_in_template = [itr for itr in OptionChain_template if itr["symbol"] ==input_symbol ]
                            if len(symbol_in_template) == 0:
                                ind_df_specific_symbol = df_instrument[(df_instrument.Symbol == input_symbol) & (df_instrument['OptionType'].isin(['CE','PE']))]
                                List_of_Expiry = ind_df_specific_symbol["Expiry"].unique()
                                Expiry_strikelist_list = []
                                for expiry in List_of_Expiry:
                                    if str( expiry ) != 'NaT' :
                                        ind_df_specific_symbol_expiry = ind_df_specific_symbol[(ind_df_specific_symbol.Expiry == expiry) ]
                                        List_of_strikes = ind_df_specific_symbol_expiry["StrikePrice"].unique()
                                        LotSize = ind_df_specific_symbol_expiry.iloc[0]["LotSize"]
                                        strike_pe_ce_list = []
                                        for strike in List_of_strikes:
                                            ind_df_specific_symbol_expiry_pe = ind_df_specific_symbol_expiry[
                                                (ind_df_specific_symbol_expiry.StrikePrice == str(strike)) & (ind_df_specific_symbol_expiry.OptionType == "PE") ]
                                            if len(ind_df_specific_symbol_expiry_pe) > 0:
                                                pe_token = ind_df_specific_symbol_expiry_pe.iloc[0]["Token"]
                                            else:
                                                pe_token = "NA"
                                            ind_df_specific_symbol_expiry_ce = ind_df_specific_symbol_expiry[(            ind_df_specific_symbol_expiry.StrikePrice== str(strike))& (            ind_df_specific_symbol_expiry.OptionType == "CE")]
                                            if len(ind_df_specific_symbol_expiry_ce) > 0:
                                                ce_token = ind_df_specific_symbol_expiry_ce.iloc[0]["Token"]
                                            else:
                                                ce_token = "NA"
                                            strike_pe_ce_dictionary = dict(
                                                {
                                                    "strike": strike,
                                                    "PE_Token": pe_token,
                                                    "CE_Token": ce_token,
                                                }
                                            )
                                            strike_pe_ce_list.append(strike_pe_ce_dictionary)
                                        expiry_stikelist_dict = dict(
                                            {
                                                "Expiry": expiry,
                                                "LotSize": LotSize,
                                                "Strike_list": strike_pe_ce_list,
                                            }
                                        )
                                        Expiry_strikelist_list.append(expiry_stikelist_dict)
                                    else:
                                        print(f"Ignoring the wrong date : {expiry}")
                                Final_dic = dict(
                                    {
                                        "symbol": input_symbol,
                                        "Expiry_Strike_token": Expiry_strikelist_list,
                                    }
                                )
                                OptionChain_template.append(Final_dic)
                            else:
                                Message = (
                                    "symbol already available in option chain template"
                                )
                            List_of_Expiry_Strike_token = [
                                itr["Expiry_Strike_token"]
                                for itr in OptionChain_template
                                if itr["symbol"] == input_symbol
                            ][0]
                            for expiry_strike in List_of_Expiry_Strike_token:
                                if input_symbol not in subs_lst:
                                    List_of_particular_expiry_strike = [
                                        itr["Strike_list"]
                                        for itr in List_of_Expiry_Strike_token
                                        if itr["Expiry"] == expiry_strike["Expiry"]
                                    ][0]
                                    for strike_dict in List_of_particular_expiry_strike:
                                        print(f"Going to subscribe {input_symbol} strike {strike_dict.get('strike')}")
                                        PE_Token = strike_dict.get("PE_Token")
                                        CE_Token = strike_dict.get("CE_Token")
                                        if PE_Token != "NA":
                                            subscribe_new_token(Exchange, PE_Token)
                                        if CE_Token != "NA":
                                            subscribe_new_token(Exchange, CE_Token)
                                    if input_symbol not in subs_lst:
                                        subs_lst.append(input_symbol)
                                        print(f"{input_symbol} subscription completed")
                            pd_oc = pd.DataFrame(columns=[ "CE_token", "CE_oi",  "CE_poi", "CE_toi", "CE_lp", "CE_pc", "CE_bq1","CE_bp1", "CE_sq1", "CE_sp1",  "strike", "PE_bq1", "PE_bp1", "PE_sq1", "PE_sp1", "PE_pc",  "PE_lp",  "PE_toi", "PE_poi", "PE_oi", "PE_token","CE_coi","PE_coi","CE_v","PE_v"] )
                            List_of_Expiry_Strike_token = [
                                itr["Expiry_Strike_token"]
                                for itr in OptionChain_template
                                if itr["symbol"] == input_symbol
                            ][0]
                            try:
                                Lot_Size = [
                                    itr["LotSize"]
                                    for itr in List_of_Expiry_Strike_token
                                    if itr["Expiry"] == expiry_input
                                ][0]
                                List_of_particular_expiry_strike = [
                                    itr["Strike_list"]
                                    for itr in List_of_Expiry_Strike_token
                                    if itr["Expiry"] == expiry_input
                                ][0]
                                isFound, Fut_Token = GetToken(Exchange,input_symbol)
                                if Exchange == 'NFO':
                                    isFound, Spot_Token = GetToken('NSE',input_symbol)
                                    spot_ltp = float(api.get_quotes("NSE", str(Spot_Token)).get("lp"))
                                if Exchange == 'BFO':
                                    isFound, Spot_Token = GetToken('BSE',input_symbol)
                                    spot_ltp = convert_to_float(api.get_quotes("BSE", str(Spot_Token)).get("lp"))
                                else:    
                                    Spot_Token = Fut_Token
                                    spot_ltp = float(api.get_quotes(Exchange, str(Spot_Token)).get("lp"))
                                future_ltp = float(api.get_quotes(Exchange, str(Fut_Token)).get("lp"))
                                for strike_dict in List_of_particular_expiry_strike:
                                    Strike = convert_to_float(strike_dict.get("strike"))
                                    PE_Token = strike_dict.get("PE_Token")
                                    PE_Token = str(Exchange)+ "|"  + str(PE_Token)
                                    CE_Token = strike_dict.get("CE_Token")
                                    CE_Token =  str(Exchange)+ "|" + str(CE_Token)
                                    try:
                                        CE_oi = live_data[str(CE_Token)].get("oi", 0)
                                    except:
                                        CE_oi = 0
                                    try:
                                        CE_poi = live_data[str(CE_Token)].get("poi", 0)
                                    except:
                                        CE_poi = 0
                                    CE_coi = int(CE_oi) - int(CE_poi)
                                    try:
                                        CE_toi = live_data[str(CE_Token)].get("toi", "-")
                                    except:
                                        CE_toi = "-"
                                    try:
                                        CE_lp = live_data[str(CE_Token)].get("lp", 0)
                                    except:
                                        CE_lp = 0
                                    try:
                                        CE_pc = live_data[str(CE_Token)].get("pc", "-")
                                    except:
                                        CE_pc = "-"
                                    try:
                                        CE_bq1 = live_data[str(CE_Token)].get("bq1", "-")
                                    except:
                                        CE_bq1 = "-"
                                    try:
                                        CE_bp1 = live_data[str(CE_Token)].get("bp1", "-")
                                    except:
                                        CE_bp1 = "-"
                                    try:
                                        CE_sq1 = live_data[str(CE_Token)].get("sq1", "-")
                                    except:
                                        CE_sq1 = "-"
                                    try:
                                        CE_sp1 = live_data[str(CE_Token)].get("sp1", "-")
                                    except:
                                        CE_sp1 = "-"
                                    try:
                                        PE_oi = live_data[str(PE_Token)].get("oi", 0)
                                    except:
                                        PE_oi = 0
                                    try:
                                        PE_poi = live_data[str(PE_Token)].get("poi", 0)
                                    except:
                                        PE_poi = 0
                                    PE_coi = int(PE_oi) - int(PE_poi)
                                    try:
                                        PE_toi = live_data[str(PE_Token)].get("toi", "-")
                                    except:
                                        PE_toi = "-"
                                    try:
                                        PE_lp = live_data[str(PE_Token)].get("lp", 0)
                                    except:
                                        PE_lp = 0
                                    try:
                                        PE_pc = live_data[str(PE_Token)].get("pc", "-")
                                    except:
                                        PE_pc = "-"
                                    try:
                                        PE_bq1 = live_data[str(PE_Token)].get("bq1", "-")
                                    except:
                                        PE_bq1 = "-"
                                    try:
                                        PE_bp1 = live_data[str(PE_Token)].get("bp1", "-")
                                    except:
                                        PE_bp1 = "-"
                                    try:
                                        PE_sq1 = live_data[str(PE_Token)].get("sq1", "-")
                                    except:
                                        PE_sq1 = "-"
                                    try:
                                        PE_sp1 = live_data[str(PE_Token)].get("sp1", "-")
                                    except:
                                        PE_sp1 = "-"
                                    try:
                                        CE_v = live_data[str(CE_Token)].get("v", 0)
                                    except:
                                        CE_v = 0
                                    try:
                                        PE_v = live_data[str(PE_Token)].get("v", 0)
                                    except:
                                        PE_v = 0
                                    dic_data = {
                                            "CE_token": CE_Token,
                                            "CE_oi": CE_oi,
                                            "CE_poi": CE_poi,
                                            "CE_toi": CE_toi,
                                            "CE_lp": CE_lp,
                                            "CE_pc": CE_pc,
                                            "CE_bq1": CE_bq1,
                                            "CE_bp1": CE_bp1,
                                            "CE_sq1": CE_sq1,
                                            "CE_sp1": CE_sp1,
                                            "strike": Strike,
                                            "PE_bq1": PE_bq1,
                                            "PE_bp1": PE_bp1,
                                            "PE_sq1": PE_sq1,
                                            "PE_sp1": PE_sp1,
                                            "PE_pc": PE_pc,
                                            "PE_lp": PE_lp,
                                            "PE_toi": PE_toi,
                                            "PE_poi": PE_poi,
                                            "PE_oi": PE_oi,
                                            "PE_token": PE_Token,
                                            "CE_coi":CE_coi,
                                            "PE_coi":PE_coi,
                                            "CE_v":CE_v,
                                            "PE_v":PE_v,
                                        }
                                    pd_oc = pd.concat([pd_oc, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                pd_oc = pd_oc.astype({"strike": float})
                                pd_oc = pd_oc.sort_values(by="strike", ascending=True)
                                pd_oc = pd_oc.fillna(0)
                                pd_oc = pd_oc.astype({"CE_oi": int})
                                pd_oc = pd_oc.astype({"PE_oi": int})
                                pd_oc = pd_oc.astype({"CE_poi": int})
                                pd_oc = pd_oc.astype({"PE_poi": int})
                                pd_oc = pd_oc.astype({"CE_coi": int})
                                pd_oc = pd_oc.astype({"PE_coi": int})
                                pd_oc = pd_oc.astype({"CE_v": int})
                                pd_oc = pd_oc.astype({"PE_v": int})
                                pd_oc["CE_v"] = pd_oc["CE_v"] / int(Lot_Size)
                                pd_oc["PE_v"] = pd_oc["PE_v"] / int(Lot_Size)
                                pd_oc["CE_oi"] = pd_oc["CE_oi"] / int(Lot_Size)
                                pd_oc["PE_oi"] = pd_oc["PE_oi"] / int(Lot_Size)
                                pd_oc["CE_poi"] = pd_oc["CE_poi"] / int(Lot_Size)
                                pd_oc["PE_poi"] = pd_oc["PE_poi"] / int(Lot_Size)
                                pd_oc["CE_coi"] = pd_oc["CE_coi"] / int(Lot_Size)
                                pd_oc["PE_coi"] = pd_oc["PE_coi"] / int(Lot_Size)
                                pd_oc["OI_SUM"] = pd_oc["CE_oi"] + pd_oc["PE_oi"]
                                df_oc_pro = pd_oc
                                try:
                                    NoOfStrike = int(oci_pro.range("E6").value)
                                except Exception as e:
                                    NoOfStrike = 100
                                df_oc_pro['strike_diff'] = abs(df_oc_pro['strike'] - spot_ltp)
                                df_oc_pro.sort_values(by = 'strike_diff',inplace = True)
                                AtmStrike = convert_to_float(df_oc_pro.iloc[0]['strike'])
                                AtmStrikeCallPrice = convert_to_float(df_oc_pro.iloc[0]['CE_lp'])
                                AtmStrikePutPrice = convert_to_float(df_oc_pro.iloc[0]['PE_lp'])
                                Future_LTP = future_ltp
                                Max_Pain_at_Strike = df_oc_pro[df_oc_pro.OI_SUM == df_oc_pro["OI_SUM"].max()].iloc[0]['strike']
                                Ltp_at_Max_Pain_CE = df_oc_pro[df_oc_pro.OI_SUM == df_oc_pro["OI_SUM"].max()].iloc[0]['CE_lp']
                                Ltp_at_Max_Pain_PE = df_oc_pro[df_oc_pro.OI_SUM == df_oc_pro["OI_SUM"].max()].iloc[0]['PE_lp']
                                LTP_at_ATM_CE = df_oc_pro[df_oc_pro.strike == AtmStrike].iloc[0]['CE_lp']
                                LTP_at_ATM_PE = df_oc_pro[df_oc_pro.strike == AtmStrike].iloc[0]['PE_lp']
                                Total_OI_CE = df_oc_pro["CE_oi"].sum()
                                Total_OI_PE = df_oc_pro["PE_oi"].sum()
                                Max_OI_CE = df_oc_pro["CE_oi"].max()
                                Max_OI_PE = df_oc_pro["PE_oi"].max()
                                Max_OI_at_Strike_CE = df_oc_pro[df_oc_pro.CE_oi == df_oc_pro["CE_oi"].max()].iloc[0]['strike']
                                Max_OI_at_Strike_PE = df_oc_pro[df_oc_pro.PE_oi == df_oc_pro["PE_oi"].max()].iloc[0]['strike']
                                LTP_of_Max_OI_Strike_CE = df_oc_pro[df_oc_pro.CE_oi == df_oc_pro["CE_oi"].max()].iloc[0]['CE_lp']
                                LTP_of_Max_OI_Strike_PE = df_oc_pro[df_oc_pro.PE_oi == df_oc_pro["PE_oi"].max()].iloc[0]['PE_lp']
                                Total_OI_Change_CE = df_oc_pro["CE_coi"].sum()
                                Total_OI_Change_PE = df_oc_pro["PE_coi"].sum()
                                Max_Change_in_OI_addition_CE = df_oc_pro["CE_coi"].max()
                                Max_Change_in_OI_addition_PE = df_oc_pro["PE_coi"].max()
                                Max_OI_addition_at_Srike_CE = df_oc_pro[df_oc_pro["CE_coi"] == df_oc_pro["CE_coi"].max()].iloc[0]['strike']
                                Max_OI_addition_at_Srike_PE = df_oc_pro[df_oc_pro["PE_coi"] == df_oc_pro["PE_coi"].max()].iloc[0]['strike']
                                LTP_of_Max_OI_addition_Strike_CE = df_oc_pro[df_oc_pro["CE_coi"] == df_oc_pro["CE_coi"].max()].iloc[0]['CE_lp']
                                LTP_of_Max_OI_addition_Strike_PE = df_oc_pro[df_oc_pro["PE_coi"] == df_oc_pro["PE_coi"].max()].iloc[0]['PE_lp']
                                Max_Change_in_OI_unwinding_CE = -1 * int(df_oc_pro["CE_coi"].min())
                                Max_Change_in_OI_unwinding_PE = -1 * int(df_oc_pro["PE_coi"].min())
                                Max_OI_unwinding_at_Srike_CE = df_oc_pro[df_oc_pro["CE_coi"] == df_oc_pro["CE_coi"].min()].iloc[0]['strike']
                                Max_OI_unwinding_at_Srike_PE = df_oc_pro[df_oc_pro["PE_coi"] == df_oc_pro["PE_coi"].min()].iloc[0]['strike']
                                LTP_of_Max_OI_unwinding_Strike_CE = df_oc_pro[df_oc_pro["CE_coi"] == df_oc_pro["CE_coi"].min()].iloc[0]['CE_lp']
                                LTP_of_Max_OI_unwinding_Strike_PE = df_oc_pro[df_oc_pro["PE_coi"] == df_oc_pro["PE_coi"].min()].iloc[0]['PE_lp']
                                df_additional_detail = pd.DataFrame(columns = ['CE','PE'])
                                dic_data = {'CE':Future_LTP}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                dic_data = {'CE':Max_Pain_at_Strike}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                dic_data = {'CE':Ltp_at_Max_Pain_CE,'PE': Ltp_at_Max_Pain_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                dic_data = {'CE':AtmStrike}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                dic_data = {'CE':LTP_at_ATM_CE,'PE':LTP_at_ATM_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                dic_data = {'CE':Total_OI_CE, 'PE':Total_OI_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                dic_data = {'CE':Total_OI_Change_CE, 'PE':Total_OI_Change_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                dic_data = {'CE':Max_OI_CE,  'PE':Max_OI_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                dic_data = {'CE':Max_OI_at_Strike_CE,'PE':Max_OI_at_Strike_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                dic_data = {'CE':LTP_of_Max_OI_Strike_CE,'PE':LTP_of_Max_OI_Strike_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                dic_data = {'CE':Max_Change_in_OI_addition_CE,'PE':Max_Change_in_OI_addition_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                dic_data = {'CE':Max_OI_addition_at_Srike_CE,'PE':Max_OI_addition_at_Srike_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                dic_data = {'CE':LTP_of_Max_OI_addition_Strike_CE, 'PE':LTP_of_Max_OI_addition_Strike_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                dic_data = {'CE':Max_Change_in_OI_unwinding_CE, 'PE':Max_Change_in_OI_unwinding_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                dic_data = {'CE':Max_OI_unwinding_at_Srike_CE, 'PE':Max_OI_unwinding_at_Srike_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                dic_data = {'CE':LTP_of_Max_OI_unwinding_Strike_CE,'PE':LTP_of_Max_OI_unwinding_Strike_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                oci_pro.range("i3").options(index=False, header=False).value = df_additional_detail
                                df_oc_pro = df_oc_pro[:2*int(NoOfStrike)+1]
                                df_oc_pro.sort_values(by ='strike',inplace = True)
                                df_oc_pro['CE_IV'] = None
                                df_oc_pro['PE_IV'] = None
                                df_oc_pro = df_oc_pro.reindex(['CE_oi','CE_coi','CE_v','CE_IV','CE_lp','CE_pc','CE_bq1','CE_bp1','CE_sp1','CE_sq1','strike','PE_bq1','PE_bp1','PE_sp1','PE_sq1','PE_pc','PE_lp','PE_IV','PE_v','PE_coi','PE_oi'], axis=1)
                                SpotPrice = convert_to_float(spot_ltp)
                                FuturePrice = convert_to_float(future_ltp)
                                ExpiryDateTime = dt(expiry_input.year, expiry_input.month, expiry_input.day, 0, 0, 0)
                                ExpiryType = oci_pro.range("F7").value
                                GreekMatch = oci_pro.range("F8").value
                                if ExpiryType == 'WEEKLY':
                                    ExpiryDateType = ExpType.WEEKLY
                                else:
                                    ExpiryDateType = ExpType.MONTHLY
                                FromDateTime = dt.now() 
                                if Exchange == 'NFO':
                                    if dt.now().time() > time(15, 30, 0):
                                        FromDateTime = dt(dt.now().year, dt.now().month,dt.now().day, 15, 30, 0)
                                if GreekMatch == "SENSIBULL":
                                    tryMatchWith=TryMatchWith.SENSIBULL
                                else:
                                    tryMatchWith=TryMatchWith.NSE
                                dayCountType = DayCountType.CALENDARDAYS
                                IvGreeks = CalcIvGreeks( SpotPrice = SpotPrice,  FuturePrice = FuturePrice, AtmStrike = AtmStrike, AtmStrikeCallPrice = AtmStrikeCallPrice, AtmStrikePutPrice = AtmStrikePutPrice, ExpiryDateTime = ExpiryDateTime, ExpiryDateType = ExpiryDateType, FromDateTime = FromDateTime, tryMatchWith = tryMatchWith, dayCountType = dayCountType)
                                for ind in df_oc_pro.index:
                                    StrikePrice= convert_to_float(df_oc_pro['strike'][ind])
                                    StrikeCallPrice= convert_to_float(df_oc_pro['CE_lp'][ind])
                                    StrikePutPrice= convert_to_float(df_oc_pro['PE_lp'][ind])
                                    Greeks = IvGreeks.GetImpVolAndGreeks( StrikePrice = StrikePrice, StrikeCallPrice = StrikeCallPrice, StrikePutPrice = StrikePutPrice)
                                    if GreekMatch == "NSE":
                                        df_oc_pro.at[ind, "CE_IV"] = round(Greeks["CallIV"],2)
                                        df_oc_pro.at[ind, "PE_IV"] = round(Greeks["PutIV"],2)
                                    else:
                                        df_oc_pro.at[ind, "CE_IV"] = round(Greeks["ImplVol"],2)
                                        df_oc_pro.at[ind, "PE_IV"] = round(Greeks["ImplVol"],2)
                                del IvGreeks
                                df_oc_pro.round({"CE_lp":2, 'PE_lp':2})
                                if pre_selected_NoOfStrike != NoOfStrike:
                                    pre_selected_NoOfStrike = NoOfStrike
                                    Option_Chain_Pro_Output.range('a3:ae500').value = None
                                    Option_Chain_Pro_Output.range("a3:U500").color = (255,255,255)
                                df_oc_pro = df_oc_pro.reset_index(drop = True)
                                ATM_pos = df_oc_pro[df_oc_pro.strike == AtmStrike].index.values[0]
                                if NoOfStrike * 2 < len(df_oc_pro):
                                    df_oc_pro = df_oc_pro.iloc[ATM_pos - int(NoOfStrike) : ATM_pos + int(NoOfStrike) + 1]
                                    df_oc_pro = df_oc_pro.reset_index(drop=True)
                                    ATM_Row = int(NoOfStrike) + 3
                                    Option_Chain_Pro_Output.range(f"a{ATM_Row}:U{ATM_Row}").color = (46,132,198)
                                Option_Chain_Pro_Output.range('a3').options(index=False,header=False).value = df_oc_pro
                                Option_Chain_Pro_Output.range(f"k3:k{len(df_oc_pro) + 2}").color = (242,242,242)
                                for color_index, color_row in df_oc_pro.iterrows():
                                    excel_row = int(color_index) + 3
                                    strike_value = convert_to_float(color_row["strike"])
                                    if strike_value < SpotPrice:
                                        Option_Chain_Pro_Output.range(f"a{excel_row}:j{excel_row}").color = (221,235,247)
                                    elif strike_value > SpotPrice:
                                        Option_Chain_Pro_Output.range(f"l{excel_row}:u{excel_row}").color = (221,235,247)
                            except Exception as e:
                                Option_Chain_Pro_Output.range('a3:ae500').value = None
                                oci_pro.range('I3:J18').value = None
                                Message = "Please check the all provided detail to load option chain:" + str(e)
                                print(Message)
                                oci_pro.range("F4").value = Message
                        else:
                            Option_Chain_Pro_Output.range('a3:ae500').value = None
                            oci_pro.range('I3:J18').value = None
                            Message = expiry_not_found_message(df_instrument, input_symbol, expiry_input)
                            print(Message)
                            oci_pro.range("F4").value = Message
                    else:
                        Option_Chain_Pro_Output.range('a3:ae500').value = None
                        oci_pro.range('I3:J18').value = None
                        Message = "Please enter the expiry"
                        print(Message)
                        oci_pro.range("F4").value = Message
                else:
                    Option_Chain_Pro_Output.range('a3:ae500').value = None
                    oci_pro.range('I3:J18').value = None
                    Message = "Please enter correct symbol"
                    print(Message)
                    oci_pro.range("F3").value = Message
                    oci_pro.range("F4").value = None
                    oci_pro.range("b2:c100").value = None
            else:
                Option_Chain_Pro_Output.range('a3:ae500').value = None
                oci_pro.range('I3:J18').value = None
                Message = "Please enter the symbol"
                print(Message)
                oci_pro.range("F3").value = Message
                oci_pro.range("F4").value = None
                oci_pro.range("b2:c100").value = None
        except Exception as e:
            print(f"Excption : {e}")
            pass
        sleep(int(IterationSleep))

def CloseTrade():
    global api, BROKER_LOGIN_OK
    if not BROKER_LOGIN_OK or not internet_is_available() or SYSTEM_EXCEL_OK is False:
        Text2Speech("Square off blocked. Broker login, internet or Excel is unavailable.")
        return
    print("I am inside CloseTrade")
    try:
        df_open_position_net,OverAllPnL  = get_position()
        price_type = "MKT"
        price = 0.0
        df_open_position_net['Net Quantity'] = df_open_position_net['Net Quantity'].astype('int')
        df_open_position_net = df_open_position_net[df_open_position_net['Net Quantity']!= 0 ]
        if(len(df_open_position_net) > 0):
            df_position_net_short = df_open_position_net[df_open_position_net['Net Quantity'] < 0]
            if(len(df_position_net_short) > 0 ):
                for ind in df_position_net_short.index:
                    try:
                        api.place_order(
                            buy_or_sell='B',
                            product_type=df_position_net_short['Product'][ind],
                            exchange=df_position_net_short['Exchange'][ind],
                            tradingsymbol=df_position_net_short['Symbol'][ind],
                            quantity=abs(df_position_net_short['Net Quantity'][ind]),
                            discloseqty=0,
                            price_type=price_type,
                            price=price,
                            trigger_price=None,
                            retention="DAY",
                            algo_id=get_shoonya_algo_id(),
                            remarks="Python_Trader_UserSelected_squareoff",
                        )
                    except Exception as e:
                        Message =  str(e) + " : Exception occur in Finvasia order placement"
                        print(Message) 
            df_position_net_long = df_open_position_net[df_open_position_net['Net Quantity'] > 0]
            if(len(df_position_net_long) > 0 ):
                for ind in df_position_net_long.index:
                    try:
                        api.place_order(
                            buy_or_sell='S',
                            product_type=df_position_net_long['Product'][ind],
                            exchange=df_position_net_long['Exchange'][ind],
                            tradingsymbol=df_position_net_long['Symbol'][ind],
                            quantity=abs(df_position_net_long['Net Quantity'][ind]),
                            discloseqty=0,
                            price_type=price_type,
                            price=price,
                            trigger_price=None,
                            retention="DAY",
                            algo_id=get_shoonya_algo_id(),
                            remarks="Python_Trader_UserSelected_squareoff",)
                    except Exception as e:
                        Message =  str(e) + " : Exception occur in finvasia order placement"
                        print(Message)
        else:
            Message = "No Open position found"
            print(Message)
    except Exception as e:
        print(f"Exception occur in CloseTrade : {e}")

df_orderbook = pd.DataFrame()
df_openPosition = pd.DataFrame()

def start_Open_Position():
    excel_op = get_terminal_workbook()
    op_op = excel_op.sheets("OpenPosition")
    op_tt = get_live_trade_sheet(excel_op)
    op_hold = excel_op.sheets("Holdings")
    op_config = get_login_settings_sheet(excel_op)
    op_ob = excel_op.sheets("OrderBook")
    isTelegramEnable = False
    if op_config.range("E3").value == True:
        isTelegramEnable = True
    isVoiceEnable = False
    if op_config.range("E6").value == True:
        isVoiceEnable = True
    op_op.range(f"d2").value = False
    op_op.range(f"e2").value = 0
    op_op.range("b3:az1000").value = None
    op_hold.range("a1:w500").value  = None
    op_ob.range("b1:az100").value = None
    global LimitOrderBook
    global df_orderbook, df_openPosition
    global api
    global Telegram_Message, Voice_Message
    while True:
        try:
            excel_op = get_terminal_workbook()
            op_op = excel_op.sheets("OpenPosition")
            op_tt = get_live_trade_sheet(excel_op)
            op_hold = excel_op.sheets("Holdings")
            op_ob = excel_op.sheets("OrderBook")
            get_limits = api.get_limits()
            op_tt.range("a2").value = get_limits['cash']
            try:
                op_tt.range("b2").value = get_limits['marginused']
                op_tt.range("c2").value = get_limits['expo']
                op_tt.range("d2").value = get_limits['span']
            except Exception as e:
                op_tt.range("b2").value = 0
                op_tt.range("c2").value = 0
                op_tt.range("d2").value = 0
            df_openPosition, OverAllPnL  = get_position()
            if(len(df_openPosition) > 0):
                op_op.range(f"a2").value = OverAllPnL
                op_tt.range(f"f2").value = OverAllPnL
                if excel_op.sheets.active.name == "OpenPosition":
                    op_op.range("b3").options(index=False,header=True).value = df_openPosition
                    KillSwitch = op_op.range(f"d2").value
                    Reconfirm = int(op_op.range(f"e2").value)
                    if(KillSwitch == 'Execute' and Reconfirm == 1):
                        CloseTrade()        
                        op_op.range(f"d2").value = False
                        op_op.range(f"e2").value = 0
                    else:
                        UserAction = op_op.range(f"a{4}:a{3 + len(df_openPosition)}").value
                        if UserAction != None:
                            for i in range (0,len(df_openPosition)):
                                if UserAction[i] == 'Square_Off' or UserAction[i] == 'S':
                                    exchange = df_openPosition['Exchange'][i]
                                    tradingsymbol = df_openPosition['Symbol'][i]
                                    if df_openPosition['Net Quantity'][i] < 0:
                                        buy_or_sell = 'B'
                                    else:
                                        buy_or_sell = 'S'
                                    quantity = abs(df_openPosition['Net Quantity'][i])
                                    product_type = df_openPosition['Product'][i]
                                    tag = "PT individual Exit"
                                    price_type = "MKT"
                                    price = 0.0
                                    try:
                                        if not BROKER_LOGIN_OK or not internet_is_available() or SYSTEM_EXCEL_OK is False:
                                            raise RuntimeError("Live square-off blocked: broker login, internet or Excel unavailable")
                                        api.place_order(buy_or_sell=buy_or_sell,product_type=product_type,exchange=exchange,tradingsymbol=tradingsymbol,quantity=quantity,discloseqty=0,price_type=price_type,price=price,trigger_price=None,retention="DAY",remarks=tag,algo_id=get_shoonya_algo_id())
                                    except Exception as e:
                                        print(f"squareoff can't be done for {df_openPosition['Symbol'][i]} with reason : {e}")
                                        pass
                                    op_op.range(f"a{4+i}").value = None
        except Exception as e:
            if not is_excel_connection_lost():
                print(f"Exception occur in OpenPosition : {e}")
            pass
        try:
            if excel_op.sheets.active.name == "Holdings":
                df_holding = getholdings()
                if(len(df_holding)> 0 ):
                    op_hold.range("a1").options(index=False,header=True).value = df_holding
        except Exception as e:
            if not is_excel_connection_lost():
                print(f"Exception occur in Holdings : {e}")
            pass
        try:
            # Refresh OrderBook continuously, even when another Excel sheet is active.
            # Excel can update an inactive worksheet directly; do not gate this on the active tab.
            df_orderbook = get_order_book()
            if (len(df_orderbook) > 0):
                op_ob.range("b1").options(index=False,header=True).value = df_orderbook
                UserAction = op_ob.range(f"a{2}:a{1 + len(df_orderbook)}").value
                if UserAction is not None:
                    for i in range (0,len(df_orderbook)):
                        if UserAction[i] == 'CANCEL' or UserAction[i] == 'C':
                            try:
                                api.cancel_order(orderno=df_orderbook['Order No'][i])
                            except Exception as e:
                                print(f"Order can't be cancelled with reason : {e}")
                                pass
                            op_ob.range(f"a{2+i}").value = None
        except Exception as e:
            if not is_excel_connection_lost():
                print(f"Exception occur in OrderBook : {e}")
            pass
        try:
            process_side_order_updates()
            order_book_snapshot = get_order_book()
            for key, value in list(LimitOrderBook.items()):
                if value.get('status') not in ('COMPLETE', 'REJECTED', 'CANCELLED', 'CANCELED'):
                    status, Executed_price = order_status(key, order_book_snapshot)
                    if status:
                        value['status'] = status
                    if Executed_price not in (None, "", 0, "0"):
                        value['Executed_price'] = Executed_price
                    if status and status not in ("", "PENDING", "OPEN"):
                        value['Remarks'] = value.get('Remarks') or status
                    update_side_order_display(
                        key, value.get('status'), value.get('Executed_price'), value.get('Remarks')
                    )
        except Exception as e:
            print(f"Order status refresh failed without stopping live quotes/orders: {e}")
        try:
            if isTelegramEnable:
                if len(Telegram_Message) > 0:
                    SendMessageToTelegram(Telegram_Message[0])
                    del Telegram_Message[0]
        except Exception as e:
            pass

def getholdings():  
    global api
    df_holding = pd.DataFrame(columns=['Exchange','Name','Holding quantity','Non Poa display quantity','Average price'])
    holdings = api.get_holdings()
    for holding in holdings:
        exch = holding['exch_tsym'][0]['exch']
        tsym = holding['exch_tsym'][0]['tsym']
        holdqty = holding['holdqty']
        try:
            npoadqty = holding['npoadqty']
        except Exception as e:
            npoadqty = 0
        upldprc = holding['upldprc']
        dic_data = {'Exchange':exch,'Name':tsym,'Holding quantity':holdqty,'Non Poa display quantity':npoadqty,'Average price':upldprc}
        df_holding = pd.concat([df_holding, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
    return df_holding
    
def StartThread():
    excel_name = get_terminal_workbook()
    Config_sheet = get_login_settings_sheet(excel_name)
    configure_indicator_engine_from_login()
    try:
        threads = []
        if Config_sheet.range("E2").value == True:
            threads.append(Thread(target=start_Trade_Terminal))
            threads.append(Thread(target=start_Open_Position))      
        if Config_sheet.range("E4").value == True:
            threads.append(Thread(target=start_optionchain))
        if Config_sheet.range("E5").value == True:
            threads.append(Thread(target=start_optionchain_Pro))
        if len(threads) != 0 :    
            for thread in threads:
                thread.start()
            for thread in threads:
                thread.join()
        else:
            print("Please select atlease one feature in config and restart the algo")
    except Exception as e:
        Message = str(e) + " : Exception occur"
        print(Message)

print("Python Trader Excel Based Terminal program initialised")
if Shoonya_login() == 1:
    Thread(target=system_health_monitor, daemon=True).start()
    LoadInstrument_token()
    api.start_websocket(
        order_update_callback=event_handler_order_update,
        subscribe_callback=event_handler_quote_update,
        socket_open_callback=open_callback,
        socket_close_callback=event_handler_socket_closed,
    )
    for websocket_attempt in range(60):
        if feed_opened:
            break
        print("Trying to connect WebSocket...")
        sleep(1)
    if not feed_opened:
        Text2Speech("Warning. Shoonya live data connection failed. Live trading will not start.")
        raise SystemExit("Shoonya WebSocket did not connect within 60 seconds.")
    print("Connected to WebSocket...")
    StartThread()
    print("Enjoy the automation...")
else:
    BROKER_LOGIN_OK = False
    Text2Speech("Warning. Shoonya login failed. No live trading will start.")
    print("Shoonya login failed. Please check LOGIN credentials, API activation and internet connection.")