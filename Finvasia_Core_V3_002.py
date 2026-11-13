# -*- coding: utf-8 -*-
"""
Created on Fri Oct 14 09:13:23 2022

This tool is to fetch two option chain (of any segment/symbol) data and trade terminal sheet, where all input comes using web socket data which leads to low lag and no api restriction. 

Contact details :
Telegram Channel:  https://t.me/pythontrader
Developer Telegram ID : https://t.me/pythontrader_admin
Gmail ID:   mnkumar2020@gmail.com 
Whatsapp : 9470669031 
Youtube : youtube.com/@pythontraders

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

Market_Safety = 1

# Shoonya algo identifier registered for this trading account.
SHOONYA_ALGO_ID = "FA78603_U"

import warnings
warnings.filterwarnings("ignore")

from Shoonya_OAuth import login_with_excel
from get_auth import get_auth_code
from Shoonya_Indicators import IndicatorEngine
#from Noren import NorenApi
import os
import time
import json
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
    #print("pyttsx3 module not found")
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
    
#logging.basicConfig(level=logging.DEBUG)

df_ins_NFO = pd.DataFrame()

OptionChain_template = []
subs_lst = []
subs_pending_lst = []

Indices_To_check_instrument_Sheet = ["NIFTY","BANKNIFTY","SENSEX" "SENSEX50", "BANKEX"]
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



def normalize_expiry_input(value):
    """Convert Excel date, datetime, text or serial values to a Python date."""
    if value is None or (isinstance(value, str) and not value.strip()):
        return None
    if isinstance(value, dt):
        parsed_date = value.date()
        # xlwings can return a blank date-formatted cell as the Excel epoch.
        if parsed_date <= date(1899, 12, 30):
            return None
        return parsed_date
    if isinstance(value, date):
        if value <= date(1899, 12, 30):
            return None
        return value
    if isinstance(value, (int, float, np.integer, np.floating)):
        try:
            # Excel may return 0 for a blank date-formatted cell. Serial
            # zero is 30-12-1899, not a usable option expiry.
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
        # Several workers can independently rebind after one Excel restart.
        # Coalesce those notices so the terminal performs one replay, not many.
        if now - _LAST_EXCEL_RECONNECT_SIGNAL >= 10.0:
            EXCEL_RECONNECT_GENERATION += 1
            _LAST_EXCEL_RECONNECT_SIGNAL = now


def get_terminal_workbook():
    """Get a thread-local workbook proxy attached to the already-open workbook.

    xlwings COM objects must not be passed between worker threads. Cache one
    workbook proxy per thread, validate it in that same COM apartment, and
    only report reconnection after this thread has actually recovered.
    """
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
                        # Validate the candidate in this worker before caching it.
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

    # Recovery: Excel may have been restarted as a new COM server after a
    # crash. A worker's old COM apartment cannot reuse its dead workbook proxy.
    # Serialize rebinding attempts so the health monitor and data workers do
    # not race to open the same file. Only try when an Excel instance is
    # already running; this avoids starting Excel from a background worker.
    if getattr(_excel_com_state, "connection_lost", False):
        with EXCEL_RECONNECT_LOCK:
            try:
                running_apps = list(xw.apps)
            except Exception as apps_error:
                if is_excel_busy_error(apps_error):
                    raise
                running_apps = []

            if running_apps:
                # Recheck inside the lock: another worker may already have
                # reopened the workbook while this worker was waiting.
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

                # Only open when the workbook is genuinely absent from all
                # running Excel instances. This prevents duplicate workbook tabs.
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
                # Reached when no Excel process exists (typical after a crash)
                # OR a dead/hung Excel process is still listed but could not
                # open the workbook. Relaunch Excel and open the workbook. The
                # cooldown stops a failing launch from spawning Excel windows
                # every 8 seconds.
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

    # First-time startup: use xlwings' path lookup. It attaches to the workbook
    # if already open, or opens it on first startup if it is not open yet.
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

    # Migrate the complete legacy Config used range to LOGIN starting at D1.
    # Existing LOGIN credentials in columns A:B are left untouched.
    if str(login_sheet.range("C1").value or "") != "LOGIN_SETTINGS_MIGRATED_V1":
        if "Config" in sheet_names:
            try:
                workbook.sheets["Config"].used_range.copy(destination=login_sheet.range("D1"))
                print("Migrated Config sheet settings into LOGIN (starting at column D).")
            except Exception as exc:
                raise RuntimeError(f"Could not migrate Config sheet into LOGIN: {exc}")
        # Preserve legacy Telegram settings without relying on User_Credential.
        if "User_Credential" in sheet_names:
            legacy = workbook.sheets["User_Credential"]
            if not login_sheet.range("AZ2").value:
                login_sheet.range("AZ2").value = legacy.range("B10").value
            if not login_sheet.range("AZ3").value:
                login_sheet.range("AZ3").value = legacy.range("B11").value
        login_sheet.range("C1").value = "LOGIN_SETTINGS_MIGRATED_V1"
        workbook.save()

    return login_sheet


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
        # Excel may briefly reject COM calls while recalculating, saving, or
        # recovering after a forced close. Let the outer login retry handle it.
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
# Keep every token subscribed during this API session so a re-opened
# Shoonya WebSocket can restore the complete feed, including option strikes.
SOCKET_SUBSCRIPTIONS = set()
SOCKET_SUBSCRIPTIONS_LOCK = Lock()
# Serialize subscribe commands: reconnect restoration and newly added symbols
# can otherwise write to the Shoonya WebSocket from different threads at once.
SOCKET_SUBSCRIBE_SEND_LOCK = Lock()
indicator_engine = IndicatorEngine()

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
    #print(f"SendMessageToTelegram = {Message}")
    global TelegramBotCredential, ReceiverTelegramID
    try:
        Url = "https://api.telegram.org/bot" + str(TelegramBotCredential) +  "/sendMessage?chat_id=" + str(ReceiverTelegramID)
        
        textdata ={ "text":Message}
        response = requests.request("POST",Url,params=textdata)
    except Exception as e:
        #print(f"SendMessageToTelegram exception occur : {e}")
        pass

def event_handler_quote_update(inmessage):
    global live_data

    global SYMBOLDICT
    # e   Exchange
    # tk  Token
    # lp  LTP
    # pc  Percentage change
    # v   volume
    # o   Open price
    # h   High price
    # l   Low price
    # c   Close price
    # ap  Average trade price
    # oi  Open interest

    fields = [
        "ts",
        "lp",
        "pc",
        "c",
        "o",
        "h",
        "l",
        "v",
        "ltq",
        "ltp",
        "bp1",
        "sp1",
        "bq1",
        "sq1",
        "ap",
        "oi",
        "ap",
        "poi",
        "toi",
    ]

    message = {field: inmessage[field] for field in set(fields) & set(inmessage.keys())}
    # print(message)
    exchange = str(inmessage.get("e") or "").strip().upper()
    token = str(inmessage.get("tk") or "").strip()
    if not exchange or not token:
        return
    key = exchange + "|" + token

    # Keep quote ingestion short and independent from Excel/order workers.
    # Replace the snapshot atomically so readers never see a half-updated OHLC tick.
    with LIVE_DATA_LOCK:
        symbol_info = dict(SYMBOLDICT.get(key, {}))
        symbol_info.update(message)
        SYMBOLDICT[key] = symbol_info
        live_data[key] = symbol_info

def event_handler_order_update(tick_data):
    """Apply broker push updates to the exact order row when the ID is known."""
    print(f"Order update {tick_data}")
    try:
        order_id = str(tick_data.get("norenordno") or "").strip()
        status = str(tick_data.get("status") or "").strip().upper()
        reason = str(tick_data.get("rejreason") or "").strip()
        if not order_id or not status:
            return
        executed_price = tick_data.get("avgprc") or tick_data.get("flprc")
        broker_reason = reason if reason and reason != " " else status
        with SIDE_ORDER_TRACKER_LOCK:
            SIDE_ORDER_UPDATES[order_id] = {
                "status": status, "executed_price": executed_price,
                "remarks": broker_reason,
            }
        if order_id in LimitOrderBook:
            LimitOrderBook[order_id]["status"] = status
            if executed_price not in (None, "", "0", 0):
                LimitOrderBook[order_id]["Executed_price"] = executed_price
            LimitOrderBook[order_id]["Remarks"] = broker_reason
        update_side_order_display(order_id, status, executed_price, broker_reason)
    except Exception as exc:
        print(f"Could not apply broker order update: {exc}")

def _resubscribe_registered_tokens():
    """Restore all previously requested quote tokens after WebSocket reconnect."""
    if not feed_opened:
        return
    with SOCKET_SUBSCRIPTIONS_LOCK:
        tokens = sorted(SOCKET_SUBSCRIPTIONS)
    if not tokens:
        print("WebSocket reconnected; no previous quote subscriptions to restore.")
        return

    restored = 0
    failed = 0
    for token_key in tokens:
        if not feed_opened:
            break
        try:
            with SOCKET_SUBSCRIBE_SEND_LOCK:
                if not feed_opened:
                    break
                api.subscribe([token_key])
            restored += 1
            sleep(0.02)
        except Exception as exc:
            failed += 1
            print(f"WebSocket resubscribe failed for {token_key}: {exc}")
            sleep(0.10)
    print(
        f"WebSocket resubscription finished: {restored}/{len(tokens)} "
        f"tokens restored; {failed} failed."
    )


def open_callback():
    global feed_opened
    was_open = feed_opened
    feed_opened = True
    if was_open:
        return

    # Old quotes must not be used to place orders while the new socket is
    # still restoring subscriptions. Fresh quote callbacks repopulate them.
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
    
    if order_type == 'MARKET':
        # Shoonya API does not permit MARKET orders. Convert to a LIMIT
        # order at the supplied live LTP; never submit a zero-price order.
        if price is None:
            print(f"Live LIMIT order blocked for {tradingsymbol}: valid live LTP is unavailable.")
            return None
        try:
            price = float(price)
        except (TypeError, ValueError):
            print(f"Live LIMIT order blocked for {tradingsymbol}: invalid LTP {price!r}.")
            return None
        if not np.isfinite(price) or price <= 0:
            print(f"Live LIMIT order blocked for {tradingsymbol}: invalid LTP {price!r}.")
            return None
        trigger_price = None
        price_type = "LMT"
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
            price = price * (100 + Market_Safety) / 100
            if exchange == 'CDS':
                round_to = .0025
                result = ( int(price / round_to) + 1) * round_to
                price = float( "{:.4f}".format(result) )
            else:
                round_to = .1
                result = ( int(price / round_to) + 1) * round_to
                price = float( "{:.1f}".format(result) )
        else:
            price = price * (100 - Market_Safety) / 100
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
    # When a row has no override, show the inherited global product in its cell.
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
            algo_id=SHOONYA_ALGO_ID,
            remarks="Python_Trader",
        )
    except Exception as exc:
        Message = f"ORDER SUBMISSION ERROR for {tradingsymbol}: {exc}"
        print(Message)
        logger.error(Message)
        return None

    if not isinstance(order_response, dict):
        Message = (
            f"ORDER NOT CONFIRMED for {tradingsymbol}: Shoonya returned "
            f"{order_response!r}; check broker OrderBook before retrying."
        )
        print(Message)
        logger.error(Message)
        return None

    order_id = order_response.get("norenordno") or order_response.get("order_id")
    broker_stat = str(order_response.get("stat", "")).strip().lower()
    if not order_id or broker_stat in ("not_ok", "not ok", "error", "failed"):
        Message = (
            f"ORDER REJECTED/UNCONFIRMED for {tradingsymbol}: "
            f"{order_response.get('emsg') or order_response}"
        )
        print(Message)
        logger.error(Message)
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
    try:
        global api
        order_book = api.get_order_book()
        df_order_book = pd.DataFrame(order_book)
        if(len(df_order_book) > 0):
            df_order_book = df_order_book.sort_values(by=["norenordno"]).reset_index(drop=True)
            try:
                df_order_book = df_order_book[['norenordno' ,'status','exch','tsym','prctyp','trantype','qty','fillshares','prc','avgprc','prd','token','ls','remarks','rejreason','trgprc']]
                df_order_book = df_order_book.rename(columns={'norenordno' : 'Order No','exch' : 'Exchange','tsym' : 'Trading Symbol','qty' : 'Quantity','trantype' : 'Transaction Type','prctyp' : 'Product Type','trgprc':'Trigger Price','token' : 'Token','ls' : 'Lot Size','status' : 'Status','remarks' : 'Remarks','rejreason' : 'Order rejection reason','prd':'Order Type','avgprc':'Executed Price','fillshares':'Filled Shares','prc':'Price'})
            except Exception as e:
                try:
                    df_order_book = df_order_book[['norenordno' ,'status','exch','tsym','prctyp','trantype','qty','fillshares','prc','avgprc','prd','token','ls','remarks','trgprc']]
                    df_order_book = df_order_book.rename(columns={'norenordno' : 'Order No','exch' : 'Exchange','tsym' : 'Trading Symbol','qty' : 'Quantity','trantype' : 'Transaction Type','prctyp' : 'Product Type','trgprc':'Trigger Price','token' : 'Token','ls' : 'Lot Size','status' : 'Status','remarks' : 'Remarks','prd':'Order Type','avgprc':'Executed Price','fillshares':'Filled Shares','prc':'Price'})
                except Exception as e:
                    try:
                        df_order_book = df_order_book[['norenordno' ,'status','exch','tsym','prctyp','trantype','qty','prc','prd','token','ls','remarks','avgprc']]
                        df_order_book = df_order_book.rename(columns={'norenordno' : 'Order No','exch' : 'Exchange','tsym' : 'Trading Symbol','qty' : 'Quantity','trantype' : 'Transaction Type','prctyp' : 'Product Type','token' : 'Token','ls' : 'Lot Size','status' : 'Status','remarks' : 'Remarks','prd':'Order Type','prc':'Price','avgprc':'Executed Price'})
                    except Exception as e:
                        df_order_book = df_order_book[['norenordno' ,'status','exch','tsym','prctyp','trantype','qty','prc','prd','token','ls','remarks']]
                        df_order_book = df_order_book.rename(columns={'norenordno' : 'Order No','exch' : 'Exchange','tsym' : 'Trading Symbol','qty' : 'Quantity','trantype' : 'Transaction Type','prctyp' : 'Product Type','token' : 'Token','ls' : 'Lot Size','status' : 'Status','remarks' : 'Remarks','prd':'Order Type','prc':'Price'})
        #print(df_order_book)
    except Exception as e:
        print(f"Exception occur in get_order_book : {e}")

    return df_order_book

def GetToken_UsingSymbol(exchange, tradingsymbol):
    global df_ins_NSE, df_ins_BSE, df_ins_NFO, df_ins_CDS, df_ins_MCX, df_ins_BFO
    global api
    try:
        if exchange == 'NSE':
            df_ins_temp = df_ins_NSE[df_ins_NSE.TradingSymbol == tradingsymbol]
            Token = df_ins_temp.iloc[0]['Token']
        elif exchange == 'BSE':
            if tradingsymbol in IndexList:
                Token = Token_list[tradingsymbol]
            else:
                df_ins_temp = df_ins_BSE[df_ins_BSE.TradingSymbol == tradingsymbol]
                Token = df_ins_temp.iloc[0]['Token']
        elif exchange == 'NFO':
            df_ins_temp = df_ins_NFO[df_ins_NFO.TradingSymbol == tradingsymbol]
            Token = df_ins_temp.iloc[0]['Token']
        elif exchange == 'BFO':
            df_ins_temp = df_ins_BFO[df_ins_BFO.TradingSymbol == tradingsymbol]
            Token = df_ins_temp.iloc[0]['Token']
        elif exchange == 'CDS':
            df_ins_temp = df_ins_CDS[df_ins_CDS.TradingSymbol == tradingsymbol]
            Token = df_ins_temp.iloc[0]['Token']
        elif exchange == 'MCX':
            df_ins_temp = df_ins_MCX[df_ins_MCX.TradingSymbol == tradingsymbol]
            Token = df_ins_temp.iloc[0]['Token']
        #print(f"Token = {Token}")
    except Exception as e:
        #print(f"Exception : {e}")
        Token = api.searchscrip(exchange=exchange, searchtext=tradingsymbol).get("values")[0].get("token")
        #print(f"latest Token = {Token}")
        pass
    return Token

def subscribe_new_token(exchange, Token, force=False):
    """Send a market-feed subscription immediately; optionally resend it."""
    global api
    token_key = f"{str(exchange).strip().upper()}|{str(Token).strip()}"
    if not feed_opened:
        raise RuntimeError("Shoonya WebSocket is not open; subscription will be retried.")

    # Register before sending so a reconnect can restore this token even if
    # the socket drops immediately after the subscribe command.
    with SOCKET_SUBSCRIPTIONS_LOCK:
        already_registered = token_key in SOCKET_SUBSCRIPTIONS
        SOCKET_SUBSCRIPTIONS.add(token_key)

    try:
        with SOCKET_SUBSCRIBE_SEND_LOCK:
            if not feed_opened:
                raise RuntimeError("Shoonya WebSocket closed before subscription was sent.")
            if force or not already_registered:
                api.subscribe([token_key])
    except Exception:
        # Do not retain a failed first-time subscription; the terminal worker
        # will retry it. Existing registrations remain available for reconnect.
        if not already_registered:
            with SOCKET_SUBSCRIPTIONS_LOCK:
                SOCKET_SUBSCRIPTIONS.discard(token_key)
        raise


def _watch_new_symbol_subscription(symbol, exchange, token):
    """Verify that a newly subscribed symbol produces a real WebSocket tick.

    If no tick arrives, resend only the market-feed subscription. This check
    deliberately does not call REST quote APIs or write synthetic quote data.
    """
    token_key = f"{exchange}|{token}"
    try:
        for attempt in range(5):
            sleep(3.0)
            with LIVE_DATA_LOCK:
                if live_data.get(token_key):
                    print(f"First WebSocket tick received for {symbol} ({token_key}).")
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
            received = bool(live_data.get(token_key))
        if not received:
            print(
                f"WARNING: {symbol} ({token_key}) has no WebSocket tick after "
                "four subscription retries. Check the symbol/token or market status."
            )
    except Exception as exc:
        print(f"Subscription watchdog error for {symbol} ({token_key}): {exc}")

LimitOrderBook = {}

# Side-specific order display is independent of the hidden trade-state engine.
SIDE_ORDER_TRACKER = {}
SIDE_ORDER_TRACKER_LOCK = Lock()
SIDE_ORDER_UPDATES = {}
# Per-row/per-side command latch. A visible BUY/SELL command is submitted once;
# clearing the cell rearms that side. Accepted IDs also protect across restarts.
SIDE_TRIGGER_STATE = {}
SIDE_TRIGGER_STATE_LOCK = Lock()

# Manual BUY/SELL requests run outside the market-data/Excel refresh loop.
# Multiple workers ensure one slow broker request cannot hold up other symbols.
SIDE_ORDER_EXECUTOR = ThreadPoolExecutor(
    max_workers=64, thread_name_prefix="ShoonyaSideOrder"
)



def register_side_order_display(symbol, side, order_id, display_row=None):
    """Write the order ID into the exact triggering row; symbol lookup is only a fallback."""
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
                 if str(value or "").strip() == str(symbol).strip()),
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
        sheet.range(f"{status_col}{row_number}").value = "SUBMITTED"
        with SIDE_ORDER_TRACKER_LOCK:
            SIDE_ORDER_TRACKER[str(order_id)] = {
                "row": row_number, "side": side, "symbol": str(symbol),
            }
            early_update = SIDE_ORDER_UPDATES.get(str(order_id))
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
    """Reflect broker-confirmed status/fill in the matching side-specific columns."""
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
        display_status = str(status or "").strip().upper()
        if not display_status or display_status == "PENDING":
            display_status = "PENDING"
        sheet.range(f"{status_col}{row_number}").value = display_status
        # Shoonya avgprc is meaningful for both partial and complete fills.
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
    """Add a local pending placeholder without overwriting an earlier broker push update."""
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
    """Accept an explicit BUY/SELL command or a truthy Excel trigger value."""
    if value is None:
        return False
    return str(value).strip().upper() in (side, "TRUE", "1", "YES", "ON")



def _required_positive_number(value, field_name):
    """Parse a required positive order input without ever calling float(None)."""
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
    """Submit one side order on a worker thread; never block the quote refresh loop."""
    pythoncom = None
    try:
        # xlwings uses COM on Windows; initialize COM for this worker when available.
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pythoncom = None

        sheet = get_live_trade_sheet(get_terminal_workbook())
        id_col, status_col = ("R", "S") if side == "BUY" else ("X", "Y")
        sheet.range(f"{status_col}{row_number}").value = "SUBMITTING"

        quantity_value = _required_positive_number(qty, "Quantity")
        if not quantity_value.is_integer():
            raise ValueError("Quantity must be a whole number")
        quantity = int(quantity_value)

        product_type = str(product_type or "").strip().upper()
        if product_type not in ("MIS", "CNC", "NRML"):
            raise ValueError("Select a valid Product Type: MIS, CNC or NRML")

        market_type = str(market_type or "").strip().upper()
        if market_type == "MARKET":
            broker_order_type = "LIMIT"
            order_price = _required_positive_number(ltp, "Live LTP")
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
            sheet.range(f"{status_col}{row_number}").value = "SUBMISSION FAILED/UNCONFIRMED"
            print(f"{side} order for {symbol} was not confirmed; check Shoonya OrderBook before retrying.")
            return

        register_pending_order(order_id)
        sheet.range(f"{id_col}{row_number}").number_format = "@"
        sheet.range(f"{id_col}{row_number}").value = str(order_id)
        print(f"{side} order worker completed for {symbol}: {order_id}")
        # The BUY/SELL command cell is deliberately not cleared.

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
    """Latch a row/side command and queue it without waiting for the broker."""
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

    # Users may enter BUY/SELL before completing the other cells. Do not latch
    # or submit an incomplete command; keep checking until every required
    # input (and a fresh LTP for MARKET) is available.
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
        if state is None:
            existing_id = sheet.range(f"{id_col}{row_number}").value
            existing_status = str(sheet.range(f"{status_col}{row_number}").value or "").strip()
            # A WAITING message is only an incomplete user input, not a
            # previously submitted order. Allow the command to queue once the
            # user finishes the missing fields.
            if existing_id not in (None, "") or (
                existing_status and not existing_status.upper().startswith("WAITING:")
            ):
                SIDE_TRIGGER_STATE[key] = {"armed": True, "recovered": True}
                return
            state = {"armed": False}
            SIDE_TRIGGER_STATE[key] = state

        if state.get("armed"):
            return

        # Clear old result fields only after the user clears and re-enters the command.
        sheet.range(f"{id_col}{row_number}").value = None
        sheet.range(f"{status_col}{row_number}").value = "QUEUED"
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
    """Return LIVE_TRADE and permanently rename the old worksheet on first run."""
    try:
        return workbook.sheets("LIVE_TRADE")
    except Exception:
        try:
            old_sheet = workbook.sheets("Trade_Terminal")
            old_sheet.name = "LIVE_TRADE"
            try:
                workbook.save()
            except Exception as save_error:
                print(f"LIVE_TRADE renamed in Excel but workbook save failed: {save_error}")
            return old_sheet
        except Exception:
            # Another worker may have completed the rename at the same time.
            return workbook.sheets("LIVE_TRADE")


def start_Trade_Terminal():
    #print("I am inside start_Trade_Terminal")
    global api, live_data
    global SYMBOLDICT
    global Product_type
    global Trade_Mode
    global Telegram_Message, Voice_Message    
    global subs_lst
    excel_TT = get_terminal_workbook()
    tt = get_live_trade_sheet(excel_TT)

    # Migrate once to the requested BUY/SELL grouped layout.
    # The legacy Entry/Exit state machine is retained in hidden BA:BS columns.
    internal_marker = str(tt.range("BA3").value or "").strip()
    if internal_marker != "__PT_INTERNAL_QTY":
        legacy_headers = tt.range("M3:AE3").value
        legacy_rows = tt.range("M4:AE1000").value
        old_side_headers = [str(v or "").strip().upper() for v in (tt.range("AF3:AK3").value or [])]
        old_indicator_headers_af = [str(v or "").strip().upper() for v in (tt.range("AF3:AM3").value or [])]
        old_indicator_headers_al = [str(v or "").strip().upper() for v in (tt.range("AL3:AS3").value or [])]
        old_side_rows = tt.range("AF4:AK1000").value if old_side_headers and old_side_headers[0] == "BUY ORDER ID" else None

        # Preserve all 19 legacy engine fields without exposing them in the new UI.
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

        # Read the original indicator block before writing the new visible headers.
        if old_indicator_headers_al[0:1] == ["RSI"]:
            indicator_values = tt.range("AL3:AS1000").value
        elif old_indicator_headers_af[0:1] == ["RSI"]:
            indicator_values = tt.range("AF3:AM1000").value
        else:
            indicator_values = None

        # Visible controls: Qty, product, BUY trigger/type/price, SELL trigger/type/price.
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

        # Carry forward side-specific IDs/statuses from the immediately previous layout, if present.
        if old_side_rows is not None:
            tt.range("R4:T1000").value = [[r[0], r[1], r[2]] for r in old_side_rows]
            tt.range("X4:Z1000").value = [[r[3], r[4], r[5]] for r in old_side_rows]

        # Carry forward trade management and indicator values.
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

    # Explicit headers keep every displayed value under its matching action.
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
    tt.range("R:Z").column_width = 18
    tt.range("S:S").column_width = 26
    tt.range("Y:Y").column_width = 26
    tt.range("AG:AN").column_width = 13
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
    # Trade-mode announcement is not spoken; system alerts remain enabled.
    
    Symbol_Token = {}
    terminal_subscribed = set()
    subscription_pending = set()
    subscription_retry_after = {}
    force_resubscribe_symbols = set()
    subscription_lock = Lock()
    last_excel_reconnect_generation = EXCEL_RECONNECT_GENERATION
    from concurrent.futures import ThreadPoolExecutor
    subscription_executor = ThreadPoolExecutor(max_workers=2, thread_name_prefix="SymbolSubscribe")

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
            # A workbook-recovery replay forces a fresh Shoonya subscribe even
            # when this token is already present in the socket registry.
            with subscription_lock:
                force_resubscribe = symbol in force_resubscribe_symbols
            # Make the token available to the independent Excel writer as soon
            # as it is resolved; remove it again if the subscribe call fails.
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
            Thread(
                target=_watch_new_symbol_subscription,
                args=(symbol, exchange, token),
                daemon=True,
                name=f"SymbolFeedCheck-{tradingsymbol}"
            ).start()
        except Exception as exc:
            print(f"Subscribe error {symbol}: {exc}")
            with subscription_lock:
                subscription_retry_after[symbol] = _monotonic() + 5.0
        finally:
            with subscription_lock:
                subscription_pending.discard(symbol)

    def write_live_market_data():
        """Write LTP/OHLC/quotes on its own loop, independent of order processing."""
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
                    # Reacquire the workbook/sheet each cycle. A worksheet COM
                    # proxy from before an Excel crash can never be reused.
                    market_sheet = get_live_trade_sheet(get_terminal_workbook())
                    market_symbols = market_sheet.range("A4:A1000").value
                    market_rows = []
                    for market_symbol in market_symbols:
                        values = [None] * 11
                        if market_symbol:
                            token_key = Symbol_Token.get(str(market_symbol).strip().upper())
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
                    # Avoid flooding the console while Excel is closed; retry
                    # slowly until the user reopens the workbook.
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
    #run a parallel thread to update the status
    while True:
        try:
            # Rebind after Excel is closed/reopened; never reuse stale COM sheets.
            excel_TT = get_terminal_workbook()
            tt = get_live_trade_sheet(excel_TT)
            current_reconnect_generation = EXCEL_RECONNECT_GENERATION
            if current_reconnect_generation != last_excel_reconnect_generation:
                active_symbols = [
                    str(value).strip().upper() for value in tt.range("A4:A1000").value
                    if value is not None and str(value).strip()
                ]
                with subscription_lock:
                    terminal_subscribed.clear()
                    force_resubscribe_symbols.update(active_symbols)
                last_excel_reconnect_generation = current_reconnect_generation
                print(
                    "Excel recovery detected: scheduling fresh WebSocket subscriptions "
                    f"for {len(active_symbols)} active Trade_Terminal symbols."
                )
            Product_type = tt.range(f"P2").value
            symbols = tt.range(f"a{4}:a{1000}").value
            trading_info = tt.range("BA4:BS1000").value
            terminal_inputs = tt.range("M4:W1000").value
            management_inputs = tt.range("AA4:AD1000").value
            for row_offset, internal_row in enumerate(trading_info):
                ui = terminal_inputs[row_offset]
                management = management_inputs[row_offset]
                internal_row[0] = ui[0]                         # Qty (M)
                # Manual BUY/SELL cells are dispatched independently below.
                # Keep them out of the legacy single-position entry state machine.
                internal_row[1] = None
                internal_row[2] = None
                internal_row[18] = ui[1]                        # Order Type (N)
                internal_row[11] = management[0]                # Target (AA)
                internal_row[12] = management[1]                # SL (AB)
                internal_row[13] = management[2]                # Trail Enable (AC)
                internal_row[14] = management[3]                # Latest SL (AD)
            tt.range("BA4:BS1000").value = trading_info
            main_list = []
            indicator_list = []

            idx = 0
            for i in symbols:
                lst = [None, None, None, None,None, None, None, None, None,None,None]
                indicator_row = [None, None, None, None, None, None, "", ""]
                # Excel cells may contain trailing spaces or inconsistent case.
                # Use one canonical key for lookup, subscription and quote writing.
                i = str(i).strip().upper() if i is not None else ""
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

                            try:
                                trade_info = trading_info[idx]
                                excel_row = idx + 4
                                ui_row = terminal_inputs[idx]
                                current_ltp = convert_to_float(quote.get("lp", 0))
                                # Both sides are evaluated independently; neither command cell is cleared.
                                dispatch_terminal_side_order(
                                    tt, excel_row, i, "BUY", ui_row[0], ui_row[1],
                                    ui_row[3], ui_row[4], current_ltp
                                )
                                dispatch_terminal_side_order(
                                    tt, excel_row, i, "SELL", ui_row[0], ui_row[1],
                                    ui_row[9], ui_row[10], current_ltp
                                )
                                #print(f" {i} : {trade_info}")
                                has_new_trigger = (
                                    _is_terminal_trigger(trade_info[1], "BUY")
                                    or _is_terminal_trigger(trade_info[2], "SELL")
                                )
                                has_live_trade_state = trade_info[16] in ("Entry_Pending", "Active", "Exit_Pending")
                                if trade_info[0] is not None and (has_new_trigger or has_live_trade_state):
                                    if isinstance(trade_info[0], (int, float)):
                                        
                                        LTP = float(quote.get("lp", 0))
                                        
                                        if Trade_Mode == 'REAL':
                                            #Real trade mode handling will handle here
                                            if _is_terminal_trigger(trade_info[1], "BUY") and LTP != 0:
                                                if trade_info[3] in ['True_Market' ,'True_Limit_LTP', 'Limit_Below', 'Limit_Above']:
                                                    
                                                    
                                                    #0:Qty    1:BUY/SELL    2:Entry_Signal  3:Entry_Limit_Price  4:Entry_Done@  5:Entry_Order_ID 6: Entry_Remarks  7:Exit_Signal 8:Exit_Done @ 9:Exit_Order_ID  10: Exit_Remarks 11:Target    12:SL  13:Trail Enable 14: Latest_SL 15:Trade_status 16: PnL
                                                    
            
                                                    if trade_info[16] != 'Active' and trade_info[16] != 'Entry_Pending' and trade_info[16] != 'Exit_Pending' and trade_info[16] != 'Closed' and (trade_info[16] is None or trade_info[16] == ''):
                                                        
                                                        
                                                        if trade_info[3] == 'True_Market':
                                                            #Entry buy trade immediately
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
                                                            #Entry buy trade immediately at ltp price
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
                                                            #print(f"LTP={LTP}, trade_info[4] ={trade_info[4]}")
                                                            
                                                            #handling real Limit_Below based trade
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
                                                            #print(f"LTP={LTP}, trade_info[4] ={trade_info[4]}")
                                                            
                                                            #handling real Limit_Above based trade
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
                                                        #print(f"Order_id={Order_id}")
                                                        status = LimitOrderBook[Order_id]['status']
                                                        if status == 'COMPLETE':
                                                            Executed_price = LimitOrderBook[Order_id]['Executed_price']
                                                            tt.range(f"BF{excel_row}").value = Executed_price
                                                            tt.range(f"BQ{excel_row}").value = 'Active'
                                                        
                                                            Message = str(int(trade_info[0])) + " quantity " + str(i[4:]) + " Bought at " + str(Executed_price)
                                                            
                                                            Telegram_Message.append(Message)
                                                            # Order and execution messages are intentionally not spoken.
                        
                        
                                                        if status in ("REJECTED", "CANCELLED", "CANCELED"):
                                                            tt.range(f"BQ{excel_row}").value = status
                                                        tt.range(f"BH{excel_row}").value = LimitOrderBook[Order_id]['Remarks']
                                                        
                                                    elif trade_info[16] == 'Active':
                                                        
                                                        
                                                        PnL = (float(LTP) - float(trade_info[5])) * int(trade_info[0])
                                                        tt.range(f"BR{excel_row}").value = PnL
                                                        #print(f"PnL = {PnL}")
                                                        
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
                                                            #exit buy order immediately
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
                                                            #exit buy at ltp limit 
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
                                                            #target meets, so exit the buy order
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", target meets so exit triggered"
                                                            logger.info(Message)
                                                            
                                                            order_id = place_trade(i, int(trade_info[0]), "SELL","LIMIT",LTP, trade_info[18], display_row=idx + 4)
                                                            
                                                            register_pending_order(order_id)
                                                                
                                                            tt.range(f"BK{excel_row}").value = order_id
                                                            tt.range(f"BQ{excel_row}").value = 'Exit_Pending'
                                                        
                                                        elif TSL >= LTP and type(trade_info[13]) is float:
                                                            #sl meets, so exiting the buy order
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", TSL meets so exit triggered"
                                                            logger.info(Message)
                                                            
                                                            order_id = place_trade(i, int(trade_info[0]), "SELL","LIMIT",LTP, trade_info[18], display_row=idx + 4)
                                                            
                                                            register_pending_order(order_id)
                                                            
                                                            tt.range(f"BK{excel_row}").value = order_id
                                                            tt.range(f"BQ{excel_row}").value = 'Exit_Pending'
                                                            
                                                    
                                                    elif trade_info[16] ==  'Exit_Pending':
                                                        Order_id = str(int(trade_info[10]))
                                                        #print(f"Order_id={Order_id}")
                                                        status = LimitOrderBook[Order_id]['status']
                                                        if status == 'COMPLETE':
                                                            Executed_price = LimitOrderBook[Order_id]['Executed_price']
                                                            tt.range(f"BJ{excel_row}").value = Executed_price
                                                            tt.range(f"BQ{excel_row}").value = 'Closed'
                                                            
                                                            PnL = (float(Executed_price) - float(trade_info[5])) * int(trade_info[0])
                                                            tt.range(f"BR{excel_row}").value = PnL
                                                            
                                                            Message = str(int(trade_info[0])) + " quantity " + str(i[4:]) + " Sold at " + str(Executed_price) + ", Trade profit or Loss = " + str(round(PnL,2))
                                                            
                                                            Telegram_Message.append(Message)
                                                            # Order and execution messages are intentionally not spoken.
                                                            
                                                        if status in ("REJECTED", "CANCELLED", "CANCELED"):
                                                            tt.range(f"BQ{excel_row}").value = status
                                                        tt.range(f"BL{excel_row}").value = LimitOrderBook[Order_id]['Remarks']
                                                        
                                            elif _is_terminal_trigger(trade_info[2], "SELL") and LTP != 0:
                                                
                                                if trade_info[3] in ['True_Market' , 'True_Limit_LTP' , 'Limit_Below' , 'Limit_Above']:
                                                    #0:Qty    1:BUY/SELL    2:Entry_Signal  3:Entry_Limit_Price  4:Entry_Done@  5:Entry_Order_ID 6: Entry_Remarks  7:Exit_Signal 8:Exit_Done @ 9:Exit_Order_ID   10:Target    11:SL  12:Trail Enable 13: Latest_SL 14:Trade_status 15: PnL
                                                    
                                                    
                                                    if trade_info[16] != 'Active' and trade_info[16] != 'Entry_Pending' and trade_info[16] != 'Exit_Pending' and trade_info[16] != 'Closed' and (trade_info[16] is None or trade_info[16] == '' ):
                                                
                                                        if trade_info[3] == 'True_Market':
                                                            #entry sell order immediately
                                                            
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
                                                                #print("b")
                                                            
                                                        elif trade_info[3] == 'True_Limit_LTP':    
                                                            #Entry buy trade immediately at ltp price
                                                            
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
                                                            
                                                            #handling real Limit_Above based trade
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
                                                                                                                       #handling real Limit_Above based trade
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
                                                        #print(f"Order_id={Order_id}")
                                                        status = LimitOrderBook[Order_id]['status']
                                                        if status == 'COMPLETE':
                                                            Executed_price = LimitOrderBook[Order_id]['Executed_price']
                                                            tt.range(f"BF{excel_row}").value = Executed_price
                                                            tt.range(f"BQ{excel_row}").value = 'Active'
                                                            
                                                            Message = str(int(trade_info[0])) + " quantity " + str(i[4:]) + " Sold at " + str(Executed_price) 
                                                            
                                                            Telegram_Message.append(Message)
                                                            # Order and execution messages are intentionally not spoken.
                                                            
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
                                                            #exit sell order immediately
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
                                                            #exit SELL at ltp limit 
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
                                                            #target meets, so exiting the sell order
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", target meets so exit triggered"
                                                            logger.info(Message)
                                                            
                                                            order_id = place_trade(i, int(trade_info[0]), "BUY","LIMIT",LTP, trade_info[18], display_row=idx + 4)
                                                            
                                                            register_pending_order(order_id)
                                                            
                                                            tt.range(f"BK{excel_row}").value = order_id
                                                            tt.range(f"BQ{excel_row}").value = 'Exit_Pending'
                                                            
                                                        
                                                        elif TSL <= LTP and type(trade_info[13]) is float:
                                                            #sl hit, so exiting the sell order
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
                                                        #print(f"Order_id={Order_id}")
                                                        status = LimitOrderBook[Order_id]['status']
                                                        if status == 'COMPLETE':
                                                            Executed_price = LimitOrderBook[Order_id]['Executed_price']
                                                            tt.range(f"BJ{excel_row}").value = Executed_price
                                                            tt.range(f"BQ{excel_row}").value = 'Closed'
                                                            
                                                            PnL = (float(trade_info[5]) - float(Executed_price)) * int(trade_info[0])
                                                            tt.range(f"BR{excel_row}").value = PnL
                                                            
                                                            Message = str(int(trade_info[0])) + " quantity " + str(i[4:]) + " buyback at " + str(Executed_price) + ", Trade profit or Loss = " + str(round(PnL,2))
                                                            
                                                            Telegram_Message.append(Message)
                                                            # Order and execution messages are intentionally not spoken.
                                                            
                                                        tt.range(f"BL{excel_row}").value = LimitOrderBook[Order_id]['Remarks']
                            except Exception as e:
                                Message = "Exception occur in core order book management:" + str(e)
                                #print(Message)
                                logger.info(Message)
                                pass
                        except Exception as e:
                            #print(f"exception : {str(e)}")
                            pass
                main_list.append(lst)
                indicator_list.append(indicator_row)

                idx += 1

            # B:L market quotes are written by LiveMarketDataWriter so a
            # slow order/strategy task cannot pause LTP and OHLC updates.
            tt.range("AG4:AN1000").value = indicator_list
            tt.range("AA4:AF1000").value = tt.range("BM4:BR1000").value
            # Yield to Excel's UI/event queue; the previous tight loop starved typing.
            sleep(0.15)
                
        except Exception as e:
            #print(f"Exception : {str(e)}")
            pass
        # Keep Excel responsive even when a COM read/write cycle raises an exception.
        sleep(0.20)

def get_position():
    #print(f"I am inside get_position")
    global api
    try:
        positions = api.get_positions()
        df_positions = pd.DataFrame(positions)
        #df_positions.to_csv("positions.csv")
        #print(df_positions)
        mtm = 0
        pnl = 0
        day_m2m = 0
        if len(df_positions) > 0:
            df_positions = df_positions.sort_values(by=["tsym"]).reset_index(drop=True)
            for i in positions:
                mtm += float(i['urmtom'])
                pnl += float(i['rpnl'])
                day_m2m = mtm + pnl
            #print(f'{day_m2m} is your Daily MTM')
        
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
            #reading nse instrument symbol
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
            #reading BSE instrument symbol
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
            #reading nfo instrument symbol
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
            #reading nfo instrument symbol
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
            #reading CDS instrument symbol
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
            #reading MCX instrument symbol
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
    """Replace legacy Greek headings and clear the old Greek output columns."""
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
            # Refresh sheet proxies after an Excel crash/reopen.
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
            #print(f"pre : pre_selected_segment = {pre_selected_segment}, selected_segment={selected_segment}")
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
            #print(f"post : pre_selected_segment = {pre_selected_segment}, selected_segment={selected_segment}")
            
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

                    #print(f"Option Chain Details: input_symbol={input_symbol},expiry_input={expiry_input}")
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
                                #print("Symbol Not found, so append the details")

                                ind_df_specific_symbol = df_instrument[(df_instrument.Symbol == input_symbol) & (df_instrument['OptionType'].isin(['CE','PE']))]
                                #print(ind_df_specific_symbol)

                                List_of_Expiry = ind_df_specific_symbol["Expiry"].unique()
                                #print(List_of_Expiry)

                                Expiry_strikelist_list = []

                                for expiry in List_of_Expiry:
                                    if str( expiry ) != 'NaT' :
                                        #print(expiry)
                                        ind_df_specific_symbol_expiry = ind_df_specific_symbol[(ind_df_specific_symbol.Expiry == expiry) ]
                                        #print(ind_df_specific_symbol_expiry)

                                        List_of_strikes = ind_df_specific_symbol_expiry["StrikePrice"].unique()

                                        LotSize = ind_df_specific_symbol_expiry.iloc[0]["LotSize"]
                                        #print(LotSize)
                                        #print(List_of_strikes)
                                        strike_pe_ce_list = []

                                        for strike in List_of_strikes:
                                            #print(f"\nsymbol:{input_symbol},Expiry:{expiry},strike:{strike}")
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

                                            # print(f"symbol:{input_symbol},Expiry:{expiry},strike:{strike},pe_token={pe_token},ce_token={ce_token}")

                                            strike_pe_ce_dictionary = dict(
                                                {
                                                    "strike": strike,
                                                    "PE_Token": pe_token,
                                                    "CE_Token": ce_token,
                                                }
                                            )

                                            strike_pe_ce_list.append(strike_pe_ce_dictionary)

                                        # print(strike_pe_ce_list)

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
                                # print("\n\n")
                                # print(Expiry_strikelist_list)

                                Final_dic = dict(
                                    {
                                        "symbol": input_symbol,
                                        "Expiry_Strike_token": Expiry_strikelist_list,
                                    }
                                )

                                OptionChain_template.append(Final_dic)
                                # print("\n\n")                                #print(OptionChain_template)
                            else:
                                Message = (
                                    "symbol already available in option chain template"
                                )

                            List_of_Expiry_Strike_token = [
                                itr["Expiry_Strike_token"]
                                for itr in OptionChain_template
                                if itr["symbol"] == input_symbol
                            ][0]
                            #print(f"\n\n******\n\n{List_of_Expiry_Strike_token}")

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
                                        # print(PE_Token)
                                        CE_Token = strike_dict.get("CE_Token")
                                        # print(CE_Token)
                                        if PE_Token != "NA":
                                            subscribe_new_token(Exchange, PE_Token)
                                        if CE_Token != "NA":
                                            subscribe_new_token(Exchange, CE_Token)

                            if input_symbol not in subs_lst:
                                subs_lst.append(input_symbol)
                                print(f"{input_symbol} subscription completed")

                            pd_oc = pd.DataFrame(columns=[ "CE_token", "CE_oi",  "CE_poi", "CE_toi", "CE_lp", "CE_pc", "CE_bq1","CE_bp1", "CE_sq1", "CE_sp1",  "strike", "PE_bq1", "PE_bp1", "PE_sq1", "PE_sp1", "PE_pc",  "PE_lp",  "PE_toi", "PE_poi", "PE_oi", "PE_token","CE_coi","PE_coi","CE_v","PE_v"] )

                            # prepare option chain
                            List_of_Expiry_Strike_token = [
                                itr["Expiry_Strike_token"]
                                for itr in OptionChain_template
                                if itr["symbol"] == input_symbol
                            ][0]
                            #print(f"List_of_Expiry_Strike_token = {List_of_Expiry_Strike_token}")
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
                                #print(f"****{live_data}")
                                #print(f"List_of_particular_expiry_strike= {List_of_particular_expiry_strike}")
                                
                                isFound, Fut_Token = GetToken(Exchange,input_symbol)
                                #print(f"isFound {isFound} Fut_Token {Fut_Token}")
                                if Exchange == 'NFO':
                                    isFound, Spot_Token = GetToken('NSE',input_symbol)
                                    spot_ltp = convert_to_float(api.get_quotes("NSE", str(Spot_Token)).get("lp"))
                                elif Exchange == 'BFO':
                                    isFound, Spot_Token = GetToken('BSE',input_symbol)
                                    spot_ltp = convert_to_float(api.get_quotes("BSE", str(Spot_Token)).get("lp"))
                                else:    
                                    Spot_Token = Fut_Token
                                    spot_ltp = convert_to_float(api.get_quotes(Exchange, str(Spot_Token)).get("lp"))
                                #print(f"Spot_Token {Spot_Token} spot_ltp {spot_ltp}")
                                future_ltp = convert_to_float(api.get_quotes(Exchange, str(Fut_Token)).get("lp"))
                                
                                #print(f"{input_symbol} spot ltp = {spot_ltp} future ltp = {future_ltp}")

                                for strike_dict in List_of_particular_expiry_strike:
                                    
                                    Strike = convert_to_float(strike_dict.get("strike"))
                                    #print(Strike)
                                    PE_Token = strike_dict.get("PE_Token")
                                    PE_Token = str(Exchange)+ "|"  + str(PE_Token)
                                    #print(PE_Token)
                                    CE_Token = strike_dict.get("CE_Token")
                                    CE_Token =  str(Exchange)+ "|" + str(CE_Token)
                                    #print(CE_Token)
                                    
                                    
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
                                    
                                    #print(f"CE COI : {CE_oi} {CE_poi} {CE_coi}")
                                    #print(f"PE COI : {PE_oi} {PE_poi} {PE_coi}")
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
                                    
                                #pd_oc.to_csv("pd_oc.csv")                            
                                #print(pd_oc)
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

                                #print(pd_oc)
                                
                                df_oc_pro = pd_oc
                                
                                try:
                                    NoOfStrike = int(oci.range("E6").value)
                                except Exception as e:
                                    NoOfStrike = 100
                                
                                df_oc_pro['strike_diff'] = abs(df_oc_pro['strike'] - spot_ltp)
            
                                df_oc_pro.sort_values(by = 'strike_diff',inplace = True)
                                
                                #print(f"\n\n***\n\n{df_oc_pro}")
                                AtmStrike = convert_to_float(df_oc_pro.iloc[0]['strike'])
                                AtmStrikeCallPrice = convert_to_float(df_oc_pro.iloc[0]['CE_lp'])
                                AtmStrikePutPrice = convert_to_float(df_oc_pro.iloc[0]['PE_lp'])
                                
                                #additional detail related to dump on input page
                                Future_LTP = future_ltp
                                Max_Pain_at_Strike = df_oc_pro[df_oc_pro.OI_SUM == df_oc_pro["OI_SUM"].max()].iloc[0]['strike']
                                Ltp_at_Max_Pain_CE = df_oc_pro[df_oc_pro.OI_SUM == df_oc_pro["OI_SUM"].max()].iloc[0]['CE_lp']
                                Ltp_at_Max_Pain_PE = df_oc_pro[df_oc_pro.OI_SUM == df_oc_pro["OI_SUM"].max()].iloc[0]['PE_lp']
                                #ATM_Strike = 
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
                                
                                #print(f"SpotPrice = {SpotPrice}, FuturePrice={FuturePrice}, AtmStrike={AtmStrike}, AtmStrikeCallPrice={AtmStrikeCallPrice}, AtmStrikePutPrice={AtmStrikePutPrice}, ExpiryDateTime={ExpiryDateTime}")
                                
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
                
                                #print(f"SpotPrice={SpotPrice}, FuturePrice={FuturePrice},  AtmStrike={AtmStrike}, AtmStrikeCallPrice={AtmStrikeCallPrice}, AtmStrikePutPrice={AtmStrikePutPrice}, ExpiryDateTime={ExpiryDateTime},  ExpiryDateType={ExpiryDateType}, FromDateTime={FromDateTime}, tryMatchWith={tryMatchWith}")
                                
                                for ind in df_oc_pro.index:
                                    
                                    StrikePrice= convert_to_float(df_oc_pro['strike'][ind])
                                    StrikeCallPrice= convert_to_float(df_oc_pro['CE_lp'][ind])
                                    StrikePutPrice= convert_to_float(df_oc_pro['PE_lp'][ind])
                                    #print(f"StrikePrice={StrikePrice}, StrikeCallPrice={StrikeCallPrice}, StrikePutPrice={StrikePutPrice}")
                                    Greeks = IvGreeks.GetImpVolAndGreeks( StrikePrice = StrikePrice, StrikeCallPrice = StrikeCallPrice, StrikePutPrice = StrikePutPrice)
                                    #print(Greeks)
                                    
                                    
                                    
                                    
                                    if GreekMatch == "NSE":
                                        df_oc_pro.at[ind, "CE_IV"] = round(Greeks["CallIV"],2)
                                        df_oc_pro.at[ind, "PE_IV"] = round(Greeks["PutIV"],2)
                                    else:
                                        df_oc_pro.at[ind, "CE_IV"] = round(Greeks["ImplVol"],2)
                                        df_oc_pro.at[ind, "PE_IV"] = round(Greeks["ImplVol"],2)
                                
                                del IvGreeks
                                df_oc_pro.round({"CE_lp":2, 'PE_lp':2})
                                #print(df_oc_pro)
                                
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

                                # Shade ITM option sides and keep the strike column distinct.
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
    #print(f"GetToken called with parametrer \nExchange = {Exchange}, Symbol = {Symbol}")
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
    
    #print(f"Returning value isTokenFound = {isTokenFound}, Token = {Token} ")
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
            # Refresh sheet proxies after an Excel crash/reopen.
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
            #print(f"pre : pre_selected_segment = {pre_selected_segment}, selected_segment={selected_segment}")
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
            #print(f"post : pre_selected_segment = {pre_selected_segment}, selected_segment={selected_segment}")
            
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

                    #print(f"Option Chain Details: input_symbol={input_symbol},expiry_input={expiry_input}")
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
                                #print("Symbol Not found, so append the details")

                                ind_df_specific_symbol = df_instrument[(df_instrument.Symbol == input_symbol) & (df_instrument['OptionType'].isin(['CE','PE']))]
                                #print(ind_df_specific_symbol)

                                List_of_Expiry = ind_df_specific_symbol["Expiry"].unique()
                                #print(List_of_Expiry)

                                Expiry_strikelist_list = []

                                for expiry in List_of_Expiry:
                                    if str( expiry ) != 'NaT' :
                                        #print(expiry)
                                        ind_df_specific_symbol_expiry = ind_df_specific_symbol[(ind_df_specific_symbol.Expiry == expiry) ]
                                        #print(ind_df_specific_symbol_expiry)

                                        List_of_strikes = ind_df_specific_symbol_expiry["StrikePrice"].unique()

                                        LotSize = ind_df_specific_symbol_expiry.iloc[0]["LotSize"]
                                        #print(LotSize)
                                        #print(List_of_strikes)
                                        strike_pe_ce_list = []

                                        for strike in List_of_strikes:
                                            #print(f"\nsymbol:{input_symbol},Expiry:{expiry},strike:{strike}")
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

                                            # print(f"symbol:{input_symbol},Expiry:{expiry},strike:{strike},pe_token={pe_token},ce_token={ce_token}")

                                            strike_pe_ce_dictionary = dict(
                                                {
                                                    "strike": strike,
                                                    "PE_Token": pe_token,
                                                    "CE_Token": ce_token,
                                                }
                                            )

                                            strike_pe_ce_list.append(strike_pe_ce_dictionary)

                                        # print(strike_pe_ce_list)

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
                                # print("\n\n")
                                # print(Expiry_strikelist_list)

                                Final_dic = dict(
                                    {
                                        "symbol": input_symbol,
                                        "Expiry_Strike_token": Expiry_strikelist_list,
                                    }
                                )

                                OptionChain_template.append(Final_dic)
                                # print("\n\n")
                                #print(OptionChain_template)
                            else:
                                Message = (
                                    "symbol already available in option chain template"
                                )

                            List_of_Expiry_Strike_token = [
                                itr["Expiry_Strike_token"]
                                for itr in OptionChain_template
                                if itr["symbol"] == input_symbol
                            ][0]
                            #print(f"\n\n******\n\n{List_of_Expiry_Strike_token}")

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
                                        # print(PE_Token)
                                        CE_Token = strike_dict.get("CE_Token")
                                        # print(CE_Token)
                                        if PE_Token != "NA":
                                            subscribe_new_token(Exchange, PE_Token)
                                        if CE_Token != "NA":
                                            subscribe_new_token(Exchange, CE_Token)

                            if input_symbol not in subs_lst:
                                subs_lst.append(input_symbol)
                                print(f"{input_symbol} subscription completed")

                            pd_oc = pd.DataFrame(columns=[ "CE_token", "CE_oi",  "CE_poi", "CE_toi", "CE_lp", "CE_pc", "CE_bq1","CE_bp1", "CE_sq1", "CE_sp1",  "strike", "PE_bq1", "PE_bp1", "PE_sq1", "PE_sp1", "PE_pc",  "PE_lp",  "PE_toi", "PE_poi", "PE_oi", "PE_token","CE_coi","PE_coi","CE_v","PE_v"] )

                            # prepare option chain
                            List_of_Expiry_Strike_token = [
                                itr["Expiry_Strike_token"]
                                for itr in OptionChain_template
                                if itr["symbol"] == input_symbol
                            ][0]
                            #print(f"List_of_Expiry_Strike_token = {List_of_Expiry_Strike_token}")
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
                                #print(f"****{live_data}")
                                #print(f"List_of_particular_expiry_strike= {List_of_particular_expiry_strike}")
                                
                                isFound, Fut_Token = GetToken(Exchange,input_symbol)
                                #print(f"isFound {isFound} Fut_Token {Fut_Token}")
                                if Exchange == 'NFO':
                                    isFound, Spot_Token = GetToken('NSE',input_symbol)
                                    spot_ltp = float(api.get_quotes("NSE", str(Spot_Token)).get("lp"))
                                if Exchange == 'BFO':
                                    isFound, Spot_Token = GetToken('BSE',input_symbol)
                                    spot_ltp = convert_to_float(api.get_quotes("BSE", str(Spot_Token)).get("lp"))
                                else:    
                                    Spot_Token = Fut_Token
                                    spot_ltp = float(api.get_quotes(Exchange, str(Spot_Token)).get("lp"))
                                #print(f"Spot_Token {Spot_Token} spot_ltp {spot_ltp}")
                                future_ltp = float(api.get_quotes(Exchange, str(Fut_Token)).get("lp"))
                                
                                #print(f"{input_symbol} spot ltp = {spot_ltp} future ltp = {future_ltp}")

                                for strike_dict in List_of_particular_expiry_strike:
                                    
                                    Strike = convert_to_float(strike_dict.get("strike"))
                                    #print(Strike)
                                    PE_Token = strike_dict.get("PE_Token")
                                    PE_Token = str(Exchange)+ "|"  + str(PE_Token)
                                    #print(PE_Token)
                                    CE_Token = strike_dict.get("CE_Token")
                                    CE_Token =  str(Exchange)+ "|" + str(CE_Token)
                                    #print(CE_Token)
                                    
                                    
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
                                    
                                    #print(f"CE COI : {CE_oi} {CE_poi} {CE_coi}")
                                    #print(f"PE COI : {PE_oi} {PE_poi} {PE_coi}")
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
                                    
                                #pd_oc.to_csv("pd_oc.csv")                            
                                #print(pd_oc)
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

                                #print(pd_oc)
                                
                                df_oc_pro = pd_oc
                                
                                try:
                                    NoOfStrike = int(oci_pro.range("E6").value)
                                except Exception as e:
                                    NoOfStrike = 100
                                
                                df_oc_pro['strike_diff'] = abs(df_oc_pro['strike'] - spot_ltp)
            
                                df_oc_pro.sort_values(by = 'strike_diff',inplace = True)
                                
                                #print(f"\n\n***\n\n{df_oc_pro}")
                                AtmStrike = convert_to_float(df_oc_pro.iloc[0]['strike'])
                                AtmStrikeCallPrice = convert_to_float(df_oc_pro.iloc[0]['CE_lp'])
                                AtmStrikePutPrice = convert_to_float(df_oc_pro.iloc[0]['PE_lp'])
                                
                                #additional detail related to dump on input page
                                Future_LTP = future_ltp
                                Max_Pain_at_Strike = df_oc_pro[df_oc_pro.OI_SUM == df_oc_pro["OI_SUM"].max()].iloc[0]['strike']
                                Ltp_at_Max_Pain_CE = df_oc_pro[df_oc_pro.OI_SUM == df_oc_pro["OI_SUM"].max()].iloc[0]['CE_lp']
                                Ltp_at_Max_Pain_PE = df_oc_pro[df_oc_pro.OI_SUM == df_oc_pro["OI_SUM"].max()].iloc[0]['PE_lp']
                                #ATM_Strike = 
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
                                
                                #print(f"SpotPrice = {SpotPrice}, FuturePrice={FuturePrice}, AtmStrike={AtmStrike}, AtmStrikeCallPrice={AtmStrikeCallPrice}, AtmStrikePutPrice={AtmStrikePutPrice}, ExpiryDateTime={ExpiryDateTime}")
                                
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
                
                                #print(f"SpotPrice={SpotPrice}, FuturePrice={FuturePrice},  AtmStrike={AtmStrike}, AtmStrikeCallPrice={AtmStrikeCallPrice}, AtmStrikePutPrice={AtmStrikePutPrice}, ExpiryDateTime={ExpiryDateTime},  ExpiryDateType={ExpiryDateType}, FromDateTime={FromDateTime}, tryMatchWith={tryMatchWith}")
                                
                                for ind in df_oc_pro.index:
                                    
                                    StrikePrice= convert_to_float(df_oc_pro['strike'][ind])
                                    StrikeCallPrice= convert_to_float(df_oc_pro['CE_lp'][ind])
                                    StrikePutPrice= convert_to_float(df_oc_pro['PE_lp'][ind])
                                    #print(f"StrikePrice={StrikePrice}, StrikeCallPrice={StrikeCallPrice}, StrikePutPrice={StrikePutPrice}")
                                    Greeks = IvGreeks.GetImpVolAndGreeks( StrikePrice = StrikePrice, StrikeCallPrice = StrikeCallPrice, StrikePutPrice = StrikePutPrice)
                                    #print(Greeks)
                                    
                                    
                                    
                                    
                                    if GreekMatch == "NSE":
                                        df_oc_pro.at[ind, "CE_IV"] = round(Greeks["CallIV"],2)
                                        df_oc_pro.at[ind, "PE_IV"] = round(Greeks["PutIV"],2)
                                    else:
                                        df_oc_pro.at[ind, "CE_IV"] = round(Greeks["ImplVol"],2)
                                        df_oc_pro.at[ind, "PE_IV"] = round(Greeks["ImplVol"],2)
                                
                                del IvGreeks
                                df_oc_pro.round({"CE_lp":2, 'PE_lp':2})
                                #print(df_oc_pro)
                                
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

                                # Shade ITM option sides and keep the strike column distinct.
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
        df_open_position_net = df_open_position_net[df_open_position_net['Net Quantity']!= 0 ] #make it 0 
        
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
                            algo_id=SHOONYA_ALGO_ID,
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
                            algo_id=SHOONYA_ALGO_ID,
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
            # Rebind all worksheet proxies to the current Excel process.
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
                    #print(f"KillSwitch={KillSwitch} Reconfirm={Reconfirm}")
                    if(KillSwitch == 'Execute' and Reconfirm == 1):
                        CloseTrade()        
                        op_op.range(f"d2").value = False
                        op_op.range(f"e2").value = 0
                    else:
                        
                        #print("check if any open order needs to be cancel")
                        UserAction = op_op.range(f"a{4}:a{3 + len(df_openPosition)}").value
                        #print(f"UserAction = {UserAction}")
                        if UserAction != None:
                            for i in range (0,len(df_openPosition)):
                                #print(f"{i} : {UserAction[i]}")
                                if UserAction[i] == 'Square_Off' or UserAction[i] == 'S':
                                    #print(f"*{df_openPosition}**")
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
                                    #print(f"squareoff {tradingsymbol}  {exchange} {product} {quantity}{transaction_type}")
                                    try:
                                        if not BROKER_LOGIN_OK or not internet_is_available() or SYSTEM_EXCEL_OK is False:
                                            raise RuntimeError("Live square-off blocked: broker login, internet or Excel unavailable")
                                        api.place_order(buy_or_sell=buy_or_sell,product_type=product_type,exchange=exchange,tradingsymbol=tradingsymbol,quantity=quantity,discloseqty=0,price_type=price_type,price=price,trigger_price=None,retention="DAY",remarks=tag,algo_id=SHOONYA_ALGO_ID)
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
            if excel_op.sheets.active.name == "OrderBook":
                df_orderbook = get_order_book()
                if (len(df_orderbook) > 0):
                    op_ob.range("b1").options(index=False,header=True).value = df_orderbook
                    
                    
                    #check if any open order needs to be cancel
                    UserAction = op_ob.range(f"a{2}:a{1 + len(df_orderbook)}").value
                    #print(f"UserAction = {UserAction}")
                    if UserAction is not None:
                        for i in range (0,len(df_orderbook)):
                            #print(f"{i} : {UserAction[i]}")
                            if UserAction[i] == 'CANCEL' or UserAction[i] == 'C':
                                #print(f"Cancel order id {df_orderbook['order_id'][i]} variety {df_orderbook['variety'][i]}")
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
            # Fetch the broker order book once per cycle, not once per order.
            # This prevents dozens of pending symbols from flooding the API and
            # competing with new order submissions.
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

        #print(f"Telegram_Message queue = {Telegram_Message}")
        try:
            if isTelegramEnable:
                if len(Telegram_Message) > 0:
                    SendMessageToTelegram(Telegram_Message[0])
                    del Telegram_Message[0]
        
        
        except Exception as e:
            #print(f"Exception in telegram Send Message : {e}")
            pass
            
        # Trade/order voice queue intentionally disabled; direct system alerts remain active.
def getholdings():  
    #print("I am inside getholdings")
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
        
        #print(f"{exch} {tsym} {holdqty} {npoadqty} {upldprc}")
        dic_data = {'Exchange':exch,'Name':tsym,'Holding quantity':holdqty,'Non Poa display quantity':npoadqty,'Average price':upldprc}
        df_holding = pd.concat([df_holding, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
    return df_holding
    
def StartThread():
    excel_name = get_terminal_workbook()
    Config_sheet = get_login_settings_sheet(excel_name)
    try:
        
        # Define the threads and put them in an array
        threads = []
        
        if Config_sheet.range("E2").value == True:
            threads.append(Thread(target=start_Trade_Terminal))
            threads.append(Thread(target=start_Open_Position))      
           
        if Config_sheet.range("E4").value == True:
            threads.append(Thread(target=start_optionchain))
        
        if Config_sheet.range("E5").value == True:
            threads.append(Thread(target=start_optionchain_Pro))
            
        if len(threads) != 0 :    
            # Func1 and Func2 run in separate threads
            for thread in threads:
                thread.start()

            # Wait until both Func1 and Func2 have finished
            for thread in threads:
                thread.join()
        else:
            print("Please select atlease one feature in config and restart the algo")
        
    except Exception as e:
        Message = str(e) + " : Exception occur"
        print(Message)

print("Python Trader Excel Based Terminal program initialised")
if Shoonya_login() == 1:
    # Start the monitor only after the main thread has attached to the workbook.
    # This prevents the monitor racing startup and opening a second Excel copy.
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
