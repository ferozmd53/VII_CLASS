# get_auth.py
# CDP-based Shoonya OAuth login (ChromeDriver not required).
#   • Robust workbook finder (Desktop / OneDrive / Documents / CWD / script dir)
#   • Fills fields by known IDs: lgnusrid / lgnpwd / lgnotp
#   • 3-second pause after each field
#   • Session cache — rapid re-calls don't reopen Chrome
#   • Chrome closes automatically after auth

import json
import os
import re
import shutil
import socket
import subprocess
import sys
import tempfile
import threading
import time
import urllib.request
from pathlib import Path
from urllib.parse import unquote

import pyotp
import requests
import websocket
import xlwings as xw


INPUT_FILL_DELAY_SEC = 3.0
LAUNCH_TIMEOUT_SEC   = 20
AUTH_CACHE_TTL_SEC   = 120.0
WORKBOOK_NAME        = "Finvasia_Trade_Terminal_v3.xlsm"

_auth_cache = {"code": None, "ts": 0.0}
_auth_lock  = threading.Lock()


def _find_workbook():
    home = Path(os.path.expanduser("~"))
    env_workbook = os.environ.get("SHOONYA_WORKBOOK_PATH")
    candidates = []
    if env_workbook:
        candidates.append(Path(env_workbook))
    candidates += [
        home / "Desktop" / WORKBOOK_NAME,
        home / "OneDrive" / "Desktop" / WORKBOOK_NAME,
        home / "OneDrive" / WORKBOOK_NAME,
        home / "Documents" / WORKBOOK_NAME,
        home / "Downloads" / WORKBOOK_NAME,
        home / WORKBOOK_NAME,
        Path.cwd() / WORKBOOK_NAME,
    ]
    if sys.argv and sys.argv[0]:
        candidates.append(Path(sys.argv[0]).resolve().parent / WORKBOOK_NAME)
    candidates.append(Path(__file__).resolve().parent / WORKBOOK_NAME)

    for c in candidates:
        try:
            if c.is_file():
                return c
        except (OSError, ValueError):
            continue

    for root in [home / "Desktop", home / "OneDrive" / "Desktop",
                 home / "OneDrive", home / "Documents"]:
        if not root.is_dir():
            continue
        try:
            for p in root.rglob(WORKBOOK_NAME):
                if len(p.relative_to(root).parts) - 1 <= 3 and p.is_file():
                    print(f"  ✅ Found workbook: {p}")
                    return p
        except (OSError, PermissionError):
            continue

    raise FileNotFoundError(
        f"Cannot find {WORKBOOK_NAME}. Searched:\n  "
        + "\n  ".join(str(c) for c in candidates)
    )


def _free_port():
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def _http_json(url, timeout=3):
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
    with opener.open(url, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8"))


def _kill_temp_browsers(marker="shoonya_browser_"):
    if os.name != "nt":
        return 0
    try:
        ps = ("Get-CimInstance Win32_Process "
              "-Filter \"Name='chrome.exe' OR Name='msedge.exe'\" | "
              f"Where-Object {{ $_.CommandLine -like '*{marker}*' }} | "
              "ForEach-Object { Stop-Process -Id $_.ProcessId -Force "
              "-ErrorAction SilentlyContinue; $_.ProcessId }")
        r = subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                           capture_output=True, text=True, timeout=15)
        return len([p for p in r.stdout.split() if p.strip().isdigit()])
    except Exception:
        return 0


def _find_chrome():
    for p in [
        os.path.join(os.environ.get("PROGRAMFILES", r"C:\Program Files"),
                     "Google", "Chrome", "Application", "chrome.exe"),
        os.path.join(os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)"),
                     "Google", "Chrome", "Application", "chrome.exe"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""),
                     "Google", "Chrome", "Application", "chrome.exe"),
    ]:
        if p and os.path.isfile(p):
            return p
    return shutil.which("chrome.exe") or shutil.which("chrome")


def _find_edge():
    for p in [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]:
        if os.path.isfile(p):
            return p
    return shutil.which("msedge.exe") or shutil.which("msedge")


class CDP:
    def __init__(self, ws):
        self.ws = ws
        self.command_id = 0

    def call(self, method, params=None):
        self.command_id += 1
        cid = self.command_id
        self.ws.send(json.dumps({"id": cid, "method": method, "params": params or {}}))
        while True:
            msg = json.loads(self.ws.recv())
            if msg.get("id") == cid:
                if "error" in msg:
                    raise RuntimeError(f"CDP error: {msg['error']}")
                return msg.get("result", {})

    def evaluate(self, expr):
        r = self.call("Runtime.evaluate",
                      {"expression": expr, "returnByValue": True, "awaitPromise": True})
        rem = r.get("result", {})
        if rem.get("subtype") == "error":
            raise RuntimeError(rem.get("description", "JS failed"))
        return rem.get("value")


def _launch(exe, url, headless=False):
    port = _free_port()
    temp = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "Temp"
    temp.mkdir(parents=True, exist_ok=True)
    profile = tempfile.mkdtemp(prefix="shoonya_browser_", dir=str(temp))

    cmd = [
        exe,
        f"--remote-debugging-port={port}",
        "--remote-debugging-address=127.0.0.1",
        "--remote-allow-origins=*",
        f"--user-data-dir={profile}",
        "--enable-automation", "--no-first-run",
        "--no-default-browser-check", "--disable-sync",
        "--disable-extensions", "--disable-gpu",
        "--disable-blink-features=AutomationControlled",
        "--disable-features=IsolateOrigins,site-per-process",
    ]
    if headless:
        cmd += ["--headless=new", "--window-size=1920,1080"]
    else:
        cmd.append("--new-window")
    cmd.append(url)

    cf, si = 0, None
    if os.name == "nt":
        cf = subprocess.CREATE_NO_WINDOW
        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW

    log = tempfile.NamedTemporaryFile(delete=False, suffix=".log",
                                       dir=str(temp), mode="w").name
    log_f = open(log, "w", encoding="utf-8", errors="replace")
    proc = subprocess.Popen(cmd, stdout=log_f, stderr=subprocess.STDOUT,
                            stdin=subprocess.DEVNULL, creationflags=cf,
                            startupinfo=si)

    print(f"  port {port}...", end="", flush=True)
    deadline = time.time() + LAUNCH_TIMEOUT_SEC
    while time.time() < deadline:
        try:
            _http_json(f"http://127.0.0.1:{port}/json/version", timeout=1)
            print(" OK")
            return port, profile, log, proc, log_f
        except Exception:
            print(".", end="", flush=True)
            time.sleep(0.5)
    print(" TIMEOUT")
    try:
        log_f.flush(); log_f.close()
        proc.terminate()
    except Exception:
        pass
    raise RuntimeError(f"Port {port} not open. Log: {log}")


_FILL_JS = """
(args => {
  const {which, value} = args;
  const vis = e => !!(e.offsetWidth||e.offsetHeight||e.getClientRects().length);
  const KNOWN = {user:'lgnusrid', password:'lgnpwd', totp:'lgnotp'};
  let t = document.getElementById(KNOWN[which]);
  let by = 'id';
  if (!t || !vis(t) || t.disabled) {
    const inputs = [...document.querySelectorAll('input')].filter(e =>
      vis(e) && !e.disabled &&
      !['hidden','submit','button','checkbox','radio','image','file','reset']
        .includes((e.type||'').toLowerCase()));
    const otpPat = /otp|totp|2fa|two.?factor|authenticat|verification.?code/i;
    const isOtp = e => otpPat.test((e.id||'')+'|'+(e.name||'')+'|'+(e.placeholder||''));
    const pw = inputs.find(e => (e.type||'').toLowerCase() === 'password' && !isOtp(e));
    const otp = inputs.find(e => e !== pw && isOtp(e));
    const userPat = /user.?id|user.?name|client.?id|login.?id|login.?name|account|lgn.?usr/i;
    const user = inputs.find(e => e !== pw && e !== otp && !isOtp(e) &&
      userPat.test((e.name||'')+'|'+(e.id||'')+'|'+(e.placeholder||'')));
    t = which === 'user' ? user : which === 'password' ? pw : otp;
    by = 'pattern';
  }
  if (!t) return {ok:false, which};
  const s = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
  t.focus(); s.call(t, value);
  t.dispatchEvent(new Event('input', {bubbles:true}));
  t.dispatchEvent(new Event('change', {bubbles:true}));
  t.dispatchEvent(new KeyboardEvent('keyup', {bubbles:true}));
  t.blur();
  return {ok:true, which, got:t.value, expected:value,
          id:t.id||'', matchedBy:by};
})(__ARGS__)
"""


def get_auth_code():
    with _auth_lock:
        if _auth_cache["code"]:
            age = time.time() - _auth_cache["ts"]
            if age < AUTH_CACHE_TTL_SEC:
                print(f"↺ Reusing cached auth code ({age:.0f}s old) — Chrome not reopening")
                return _auth_cache["code"]

        print("=" * 60)
        print("SHOONYA AUTHENTICATION")
        print("=" * 60)
        print(f"get_auth.py loaded from: {os.path.abspath(__file__)}")

        workbook_path = _find_workbook()
        print(f"📘 Excel workbook: {workbook_path}")
        wb = xw.Book(str(workbook_path))

        try:
            ws = wb.sheets["LOGIN"]
        except Exception:
            ws = wb.sheets.add("LOGIN")
            ws.range("A1:B11").value = [
                ["Field","Value"],["CLIENT_ID",""],["USER_ID",""],["PASSWORD",""],
                ["TOTP_SECRET",""],["API_SECRET",""],["AUTH_CODE",""],
                ["",""],["",""],["",""],["IP_ADDRESS",""],
            ]
        if not ws.range("A1").value:
            ws.range("A1:B11").value = [
                ["Field","Value"],["CLIENT_ID",""],["USER_ID",""],["PASSWORD",""],
                ["TOTP_SECRET",""],["API_SECRET",""],["AUTH_CODE",""],
                ["",""],["",""],["",""],["IP_ADDRESS",""],
            ]

        cid      = str(ws.range("B2").value or "").strip()
        uid      = str(ws.range("B3").value or "").strip()
        pwd      = str(ws.range("B4").value or "")
        totp_sec = str(ws.range("B5").value or "").replace(" ", "").strip()

        try:
            ip = requests.get("https://api.ipify.org", timeout=5).text.strip()
            print(f"Your IP: {ip}")
            ws.range("B11").value = ip
            wb.save()
        except Exception:
            pass

        print(f"Client ID: {cid}")
        print(f"User ID: {uid}")

        login_url = (
            "https://trade.shoonya.com/OAuthlogin/investor-entry-level/login"
            f"?api_key={cid}&route_to={uid}"
        )

        n = _kill_temp_browsers()
        if n:
            print(f"🧹 Cleaned {n} stale temp browser(s).")

        candidates = []
        c = _find_chrome()
        if c:
            candidates.append(("Chrome", c))
        e = _find_edge()
        if e:
            candidates.append(("Edge", e))
        if not candidates:
            raise FileNotFoundError("No Chrome or Edge found on this PC.")

        launched = None
        for name, path in candidates:
            for headless in (False, True):
                mode = "HEADLESS" if headless else "headed"
                print(f"\n═══ {name} ({mode}) ═══")
                try:
                    launched = (name,) + _launch(path, login_url, headless)
                    print(f"  ✅ {name} ({mode}) opened debug port.")
                    break
                except Exception as ex:
                    print(f"  ❌ {ex}")
            if launched:
                break
        if not launched:
            raise RuntimeError("Chrome and Edge both failed to open debug port.")

        _, port, profile, log_path, proc, log_f = launched
        cdp = None
        obtained = False
        try:
            page = None
            for _ in range(30):
                try:
                    tgts = _http_json(f"http://127.0.0.1:{port}/json")
                except Exception:
                    time.sleep(0.5); continue
                page = next((t for t in tgts if t.get("type") == "page"), None)
                if page:
                    break
                time.sleep(0.5)
            if not page:
                raise RuntimeError("No login tab.")

            cdp = CDP(websocket.create_connection(
                page["webSocketDebuggerUrl"], timeout=15, suppress_origin=True))
            cdp.call("Page.enable")
            cdp.call("Runtime.enable")
            time.sleep(3)

            totp = pyotp.TOTP(totp_sec).now()
            print(f"\nComputed TOTP: {totp}")

            for which, val, label in [
                ("user",     uid,  "User ID"),
                ("password", pwd,  "Password"),
                ("totp",     totp, "TOTP"),
            ]:
                print(f"  Filling {label}…", end=" ", flush=True)
                r = cdp.evaluate(_FILL_JS.replace(
                    "__ARGS__", json.dumps({"which": which, "value": val})))
                if not r or not r.get("ok") or r.get("got", "").strip() != val.strip():
                    raise RuntimeError(f"{label} fill failed: {r}")
                shown = "*" * len(val) if which == "password" else val
                print(f"✅ {shown}  ({r.get('matchedBy')}, id={r.get('id')!r})")
                if INPUT_FILL_DELAY_SEC > 0:
                    print(f"     ⏳ {INPUT_FILL_DELAY_SEC:.0f}s…")
                    time.sleep(INPUT_FILL_DELAY_SEC)

            cdp.evaluate("""
            (() => {
              const b = [...document.querySelectorAll(
                'button,[role="button"],input[type="submit"]')].find(e => {
                const t = (e.innerText||e.value||'').trim().toUpperCase();
                return t === 'LOGIN' || t.includes('LOGIN');
              });
              if (!b) return 'LOGIN button not found';
              b.click();
              return 'LOGIN clicked';
            })()
            """)

            print("Waiting up to 3 min for OAuth redirect…")
            deadline = time.time() + 180
            while time.time() < deadline:
                try:
                    url = cdp.evaluate("location.href") or ""
                except Exception:
                    url = ""
                m = re.search(r"[?&#]code=([^&#]+)", url)
                if m:
                    code = unquote(m.group(1))
                    ws.range("B7").value = code
                    wb.save()
                    _auth_cache["code"] = code
                    _auth_cache["ts"]   = time.time()
                    print("✅ AUTH CODE saved to LOGIN!B7")
                    obtained = True
                    return code
                time.sleep(0.5)
            raise TimeoutError("No OAuth code in 3 min.")
        finally:
            if cdp:
                try: cdp.ws.close()
                except Exception: pass
            try:
                log_f.flush(); log_f.close()
            except Exception: pass
            if obtained:
                print("Closing temp browser.")
                _kill_temp_browsers()
            else:
                print(f"Temp browser left open for inspection. Log: {log_path}")


if __name__ == "__main__":
    get_auth_code()
