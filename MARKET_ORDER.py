# -*- coding: utf-8 -*-
"""
Shoonya Web Client MARKET-order adapter.

The existing terminal keeps its normal Shoonya OAuth API client for quotes,
order updates and LMT/SL-LMT orders. This module creates a separate client
pointing at the Shoonya Web Client host only when a MARKET order is selected.
The session token is read from LOGIN!B9 (ACCESS_TOKEN), with LOGIN!B8 as the
legacy fallback. LOGIN!B14 is a refresh token and is intentionally not used
as jKey.
"""

from Shoonya_OAuth import ShoonyaApiPy

WEB_CLIENT_HOST = "https://trade.shoonya.com/NorenWClientWeb/"
WEB_CLIENT_WEBSOCKET = "wss://trade.shoonya.com/NorenWSWeb/"


class ShoonyaMarketWebClient(ShoonyaApiPy):
    """OAuth client using the Shoonya Web Client host for order submission."""

    def __init__(self):
        super().__init__()
        # Keep the SDK's standard endpoint routes; change only the HTTP host.
        config = getattr(self, "_NorenApi__service_config", None)
        if not isinstance(config, dict):
            raise RuntimeError("Shoonya SDK service configuration is unavailable")
        config["host"] = WEB_CLIENT_HOST


def _read_login_value(login_sheet, cell):
    try:
        value = login_sheet.range(cell).value
        return str(value or "").strip()
    except Exception:
        return ""


def place_market_order(api, login_sheet, **order):
    """
    Submit one real MARKET order through the Shoonya Web Client host.

    The same already-authenticated OAuth session is reused. No credentials,
    access tokens or refresh tokens are hard-coded in this module.
    """
    uid = _read_login_value(login_sheet, "B3")
    access_token = (
        _read_login_value(login_sheet, "B9")
        or _read_login_value(login_sheet, "B8")
    )
    account_id = _read_login_value(login_sheet, "B13") or uid

    # Fall back to the active API object's in-memory session only if the
    # workbook token cells are blank.
    if not uid:
        uid = str(getattr(api, "username", "") or "").strip()
    if not access_token:
        access_token = str(getattr(api, "access_token", "") or "").strip()
    if not account_id:
        account_id = str(getattr(api, "accountid", "") or uid).strip()

    if not uid or not access_token:
        raise RuntimeError(
            "MARKET order blocked: valid Shoonya session token is missing "
            "from LOGIN!B9/B8."
        )

    market_api = ShoonyaMarketWebClient()
    market_api.injectOAuthHeader(access_token, uid, account_id)

    # The screenshot's verified Shoonya order example uses MKT with price=0.
    order["price_type"] = "MKT"
    order["price"] = 0
    order["trigger_price"] = 0
    return market_api.place_order(**order)
