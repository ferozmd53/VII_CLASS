# Extreme_Reversal_Signal.py - FIXED DATA POSITIONS

# ─── NorenApi import with automatic fallback ───
# Prefers NorenRestApiOAuth (real package — correctly formats jData).
# Falls back to NorenRestApiPy + a shim that mimics NorenRestApiOAuth's
# raw-string jData payload (NOT a URL-encoded dict).
try:
    from NorenRestApiOAuth.NorenApi import NorenApi
    if not hasattr(NorenApi, "getAccessToken"):
        raise ImportError("getAccessToken missing")
    print("✅ Using NorenRestApiOAuth (real package)")
except Exception:
    try:
        from NorenRestApiPy.NorenApi import NorenApi as _BaseNorenApi
        import requests as _req
        import json as _json
        import hashlib as _hashlib

        class NorenApi(_BaseNorenApi):
            """NorenRestApiPy with getAccessToken that mimics NorenRestApiOAuth exactly."""

            def getAccessToken(self, auth_code, secret_key, client_id, uid):
                """Exchange a fresh Shoonya OAuth code with its required SHA256 checksum."""
                if not all(str(v or "").strip() for v in (auth_code, secret_key, client_id, uid)):
                    print("❌ OAuth exchange stopped: auth code or API credentials are blank")
                    return None

                url = "https://api.shoonya.com/NorenWClientAPI/GenAcsTok"
                checksum = _hashlib.sha256(
                    (str(client_id).strip() + str(secret_key).strip() + str(auth_code).strip()).encode("utf-8")
                ).hexdigest()
                payload = "jData=" + _json.dumps({
                    "code": str(auth_code).strip(),
                    "checksum": checksum,
                })

                try:
                    response = _req.post(
                        url,
                        data=payload,
                        headers={"Content-Type": "application/x-www-form-urlencoded"},
                        timeout=30,
                    )
                    ret = response.json()
                except Exception as exc:
                    print(f"❌ getAccessToken HTTP error: {exc}")
                    return None

                if ret.get("stat") == "Ok" and (ret.get("access_token") or ret.get("susertoken")):
                    access_token = ret.get("access_token") or ret.get("susertoken")
                    returned_uid = ret.get("uid") or ret.get("USERID") or uid
                    actid = ret.get("actid")
                    refresh_token = ret.get("refresh_token")
                    try:
                        self._NorenApi__access_token = access_token
                        self._NorenApi__username = returned_uid
                        self._NorenApi__accountid = actid
                        self._NorenApi__susertoken = ret.get("susertoken", access_token)
                    except Exception:
                        pass
                    self.access_token = access_token
                    self.username = returned_uid
                    self.accountid = actid
                    print(f"✅ getAccessToken OK (uid={returned_uid})")
                    return (access_token, returned_uid, refresh_token, actid)

                print(f"❌ getAccessToken failed: {ret}")
                return None

            def injectOAuthHeader(self, token, uid, actid=None):
                try:
                    self._NorenApi__susertoken = token
                    self._NorenApi__access_token = token
                    self._NorenApi__username = uid
                    self._NorenApi__accountid = actid or uid
                    self._NorenApi__OAuthHeaders = {
                        "Authorization": f"Bearer {token}",
                        "Content-Type": "application/json; charset=utf-8",
                    }
                except Exception:
                    pass
                self.access_token = token
                self.username     = uid
                self.accountid    = actid or uid
                print(f"✅ injectOAuthHeader OK (uid={uid})")

            def place_order(
                self, buy_or_sell, product_type, exchange, tradingsymbol,
                quantity, discloseqty, price_type, price=0.0,
                trigger_price=None, retention="DAY", amo=None, remarks=None,
                bookloss_price=0.0, bookprofit_price=0.0, trail_price=0.0,
                algo_id=None
            ):
                """OAuth-compatible PlaceOrder for the legacy NorenRestApiPy fallback."""
                config = self._NorenApi__service_config
                url = f"{config['host']}{config['routes']['placeorder']}"
                values = {
                    "ordersource": "API",
                    "uid": self._NorenApi__username,
                    "actid": self._NorenApi__accountid or self._NorenApi__username,
                    "trantype": getattr(buy_or_sell, "value", buy_or_sell),
                    "prd": getattr(product_type, "value", product_type),
                    "exch": exchange,
                    "tsym": tradingsymbol,
                    "qty": str(quantity),
                    "dscqty": str(discloseqty),
                    "prctyp": getattr(price_type, "value", price_type),
                    "prc": str(price),
                    "trgprc": str(trigger_price if trigger_price is not None else 0),
                    "ret": retention,
                    "remarks": remarks,
                    "algo_id": algo_id,
                }
                if amo is not None:
                    values["amo"] = amo
                if product_type in ("H", "B"):
                    values["blprc"] = str(bookloss_price)
                if product_type == "B":
                    values["bpprc"] = str(bookprofit_price)
                if trail_price:
                    values["trailprc"] = str(trail_price)

                payload = "jData=" + _json.dumps(values)
                headers = getattr(self, "_NorenApi__OAuthHeaders", None)
                if headers:
                    response = _req.post(url, data=payload, headers=headers, timeout=30)
                else:
                    token = getattr(self, "_NorenApi__susertoken", "")
                    response = _req.post(
                        url, data=payload + "&jKey=" + token, timeout=30
                    )
                response.raise_for_status()
                try:
                    return response.json()
                except Exception:
                    return {"stat": "Not_Ok", "emsg": response.text}

            # Override the legacy NorenRestApiPy WebSocket callbacks.
            # The legacy package sends t="c" with pwd/susertoken; OAuth requires
            # t="a" with accesstoken and an ak/ck connection acknowledgement.
            def _NorenApi__on_open_callback(self, ws=None):
                self._NorenApi__websocket_connected = True
                payload = {
                    "t": "a",
                    "uid": self._NorenApi__username,
                    "actid": self._NorenApi__accountid or self._NorenApi__username,
                    "source": "API",
                    "accesstoken": self._NorenApi__access_token,
                }
                self._NorenApi__ws_send(_json.dumps(payload))

            def _NorenApi__on_data_callback(
                self, ws=None, message=None, data_type=None, continue_flag=None
            ):
                try:
                    data = _json.loads(message) if isinstance(message, str) else message
                    if not isinstance(data, dict):
                        return
                    msg_type = data.get("t")
                    if msg_type in ("ak", "ck"):
                        if str(data.get("s", "")).strip().lower() in ("ok", "success"):
                            callback = self._NorenApi__on_open
                            if callback:
                                callback()
                        else:
                            error_callback = self._NorenApi__on_error
                            if error_callback:
                                error_callback(data)
                            else:
                                print(f"❌ Shoonya WebSocket authentication rejected: {data}")
                        return
                    if msg_type in ("tk", "tf", "dk", "df"):
                        callback = self._NorenApi__subscribe_callback
                        if callback:
                            callback(data)
                        return
                    if msg_type == "om":
                        callback = self._NorenApi__order_update_callback
                        if callback:
                            callback(data)
                        return
                except Exception as e:
                    error_callback = self._NorenApi__on_error
                    if error_callback:
                        error_callback(e)
                    else:
                        print(f"❌ WebSocket message error: {e}")

        print("⚠️  Using NorenRestApiPy with jData-raw-string shim")
        print("   For the real package: pip install --no-cache-dir --no-deps NorenRestApiOAuth")
    except ImportError:
        raise ImportError(
            "Install one of:\n"
            "  pip install --no-cache-dir --no-deps NorenRestApiOAuth\n"
            "  pip install --no-cache-dir --no-deps NorenRestApiPy"
        )


import time
from datetime import datetime, timedelta
import xlwings as xw


class ShoonyaApiPy(NorenApi):
    """Shoonya OAuth API client using the OAuth-compatible SDK/shim above."""

    def __init__(self):
        super().__init__(
            host="https://api.shoonya.com/NorenWClientAPI/",
            websocket="wss://api.shoonya.com/NorenWSAPI/"
        )

    def injectOAuthHeader(self, token, uid, actid=None):
        """Restore an already-issued OAuth access token into the SDK client."""
        account_id = str(actid or uid)
        try:
            self._NorenApi__susertoken = token
            self._NorenApi__access_token = token
            self._NorenApi__username = uid
            self._NorenApi__accountid = account_id
            self._NorenApi__OAuthHeaders = {
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json; charset=utf-8",
            }
        except Exception:
            pass
        self.access_token = token
        self.username = uid
        self.accountid = account_id


def _ensure_sheet(workbook, name):
    try:
        return workbook.sheets[name]
    except Exception:
        return workbook.sheets.add(name)


def _session_age_seconds(login_sheet, session_sheet):
    value = login_sheet.range("B12").value
    if value not in (None, ""):
        try:
            if isinstance(value, datetime):
                stamp = value.replace(tzinfo=None).timestamp()
            else:
                stamp = float(value)
            return time.time() - stamp
        except Exception:
            try:
                return (datetime.now() - datetime.fromisoformat(str(value)).replace(tzinfo=None)).total_seconds()
            except Exception:
                pass

    value = session_sheet.range("B3").value
    if value not in (None, ""):
        try:
            parsed = value if isinstance(value, datetime) else datetime.fromisoformat(str(value).strip())
            return (datetime.now() - parsed.replace(tzinfo=None)).total_seconds()
        except Exception:
            pass
    return -1


def _valid_session(client):
    try:
        result = client.get_limits()
        return isinstance(result, dict) and str(result.get("stat", "")).lower() == "ok"
    except Exception as exc:
        print(f"Saved Shoonya session validation failed: {exc}")
        return False


def login_with_excel(workbook, auth_code_provider=None):
    """
    Use the workbook LOGIN sheet as the only credential/session store.
    A valid OAuth token is reused for up to 23h 50m and is broker-validated.
    """
    login_sheet = _ensure_sheet(workbook, "LOGIN")
    session_sheet = _ensure_sheet(workbook, "SESSION")

    if not login_sheet.range("A1").value:
        login_sheet.range("A1:B14").value = [
            ["FIELD", "VALUE"],
            ["CLIENT_ID", ""],
            ["USER_ID", ""],
            ["PASSWORD", ""],
            ["TOTP_SECRET", ""],
            ["API_SECRET", ""],
            ["AUTH_CODE", ""],
            ["WS_TOKEN", ""],
            ["ACCESS_TOKEN", ""],
            ["", ""],
            ["IP_ADDRESS", ""],
            ["TOKEN_CREATED_EPOCH", ""],
            ["ACCOUNT_ID", ""],
            ["REFRESH_TOKEN", ""],
        ]

    client_id = str(login_sheet.range("B2").value or "").strip()
    uid = str(login_sheet.range("B3").value or "").strip()
    secret = str(login_sheet.range("B6").value or "").strip()
    if not uid or not secret:
        print("LOGIN sheet requires USER_ID in B3 and API_SECRET in B6.")
        return None
    if not client_id:
        client_id = f"{uid}_U"
        login_sheet.range("B2").value = client_id

    try:
        session_sheet.range("A1:A6").value = [
            ["ACCESS_TOKEN"], ["WS_TOKEN"], ["CREATED_AT"],
            ["ACCOUNT_ID"], ["REFRESH_TOKEN"], ["USER_ID"]
        ]
        session_sheet.range("A:A").column_width = 24
        session_sheet.range("B:B").column_width = 55
    except Exception:
        pass

    cached_token = str(login_sheet.range("B9").value or session_sheet.range("B1").value or "").strip()
    cached_uid = str(session_sheet.range("B6").value or uid).strip()
    age = _session_age_seconds(login_sheet, session_sheet)

    if cached_token and cached_uid == uid and 0 <= age < (23 * 60 * 60 + 50 * 60):
        candidate = ShoonyaApiPy()
        candidate.injectOAuthHeader(
            cached_token, uid,
            str(login_sheet.range("B13").value or session_sheet.range("B4").value or uid).strip()
        )
        if _valid_session(candidate):
            # Restore the visible LOGIN cells from the valid saved session.
            # OAuth authorization codes are one-time values and must stay blank
            # after a successful token exchange.
            created_value = login_sheet.range("B12").value
            if created_value in (None, ""):
                saved_created = session_sheet.range("B3").value
                try:
                    if isinstance(saved_created, datetime):
                        created_epoch = int(saved_created.replace(tzinfo=None).timestamp())
                    else:
                        created_epoch = int(
                            datetime.fromisoformat(str(saved_created).strip())
                            .replace(tzinfo=None).timestamp()
                        )
                except Exception:
                    created_epoch = int(time.time() - age)
            else:
                try:
                    created_epoch = int(float(created_value))
                except Exception:
                    try:
                        created_epoch = int(
                            datetime.fromisoformat(str(created_value).strip())
                            .replace(tzinfo=None).timestamp()
                        )
                    except Exception:
                        created_epoch = int(time.time() - age)

            account_id = str(
                login_sheet.range("B13").value
                or session_sheet.range("B4").value
                or uid
            ).strip()
            refresh_token = str(
                login_sheet.range("B14").value
                or session_sheet.range("B5").value
                or ""
            ).strip()
            created_iso = session_sheet.range("B3").value
            if created_iso in (None, ""):
                created_iso = datetime.fromtimestamp(created_epoch).isoformat(timespec="seconds")

            login_sheet.range("B8").value = cached_token
            login_sheet.range("B9").value = cached_token
            login_sheet.range("B12").value = created_epoch
            login_sheet.range("B13").value = account_id
            login_sheet.range("B14").value = refresh_token

            session_sheet.range("B1").value = cached_token
            session_sheet.range("B2").value = cached_token
            session_sheet.range("B3").value = created_iso
            session_sheet.range("B4").value = account_id
            session_sheet.range("B5").value = refresh_token
            session_sheet.range("B6").value = uid
            try:
                session_sheet.visible = False
            except Exception:
                pass
            workbook.save()
            print(
                f"Reusing valid Shoonya OAuth session ({age / 3600:.1f} hours old); "
                "blank LOGIN session cells have been restored."
            )
            return candidate
        print("Saved OAuth token failed broker validation; requesting a fresh session.")
    elif cached_token:
        print("Saved OAuth session is expired, invalid or belongs to another USER_ID.")

    auth_code = str(login_sheet.range("B7").value or "").strip()

    def fresh_code():
        if not callable(auth_code_provider):
            return ""
        print("Starting browser OAuth because no valid saved session is available.")
        value = str(auth_code_provider() or "").strip()
        if value:
            login_sheet.range("B7").value = value
            workbook.save()
        return value

    if not auth_code:
        auth_code = fresh_code()
    if not auth_code:
        print("No OAuth authorization code is available.")
        return None

    client = ShoonyaApiPy()
    try:
        token_result = client.getAccessToken(auth_code, secret, client_id, uid)
    except Exception as exc:
        print(f"Shoonya OAuth token exchange failed: {exc}")
        token_result = None

    if not token_result and callable(auth_code_provider):
        print("Saved OAuth code was rejected; obtaining one fresh code and retrying.")
        login_sheet.range("B7").value = ""
        auth_code = fresh_code()
        if auth_code:
            client = ShoonyaApiPy()
            try:
                token_result = client.getAccessToken(auth_code, secret, client_id, uid)
            except Exception as exc:
                print(f"Fresh OAuth token exchange failed: {exc}")
                token_result = None

    if not token_result:
        print("Shoonya did not return an access token.")
        return None

    if isinstance(token_result, dict):
        access_token = token_result.get("access_token") or token_result.get("susertoken")
        returned_uid = token_result.get("uid") or token_result.get("USERID") or uid
        refresh_token = token_result.get("refresh_token")
        account_id = token_result.get("actid") or uid
    else:
        access_token, returned_uid, refresh_token, account_id = (list(token_result) + [None] * 4)[:4]
        account_id = account_id or uid

    if not access_token:
        print("OAuth exchange response did not contain an access token.")
        return None

    client.injectOAuthHeader(str(access_token), uid, str(account_id))
    if not _valid_session(client):
        print("New OAuth token was issued but broker validation failed.")
        return None

    created_epoch = int(time.time())
    created_iso = datetime.now().isoformat(timespec="seconds")
    login_sheet.range("B8").value = str(access_token)
    login_sheet.range("B9").value = str(access_token)
    login_sheet.range("B12").value = created_epoch
    login_sheet.range("B13").value = str(account_id)
    login_sheet.range("B14").value = str(refresh_token or "")
    login_sheet.range("B7").value = ""
    session_sheet.range("B1").value = str(access_token)
    session_sheet.range("B2").value = str(access_token)
    session_sheet.range("B3").value = created_iso
    session_sheet.range("B4").value = str(account_id)
    session_sheet.range("B5").value = str(refresh_token or "")
    session_sheet.range("B6").value = uid
    try:
        session_sheet.visible = False
    except Exception:
        pass
    workbook.save()
    print("New Shoonya OAuth session saved in the Excel workbook.")
    return client
