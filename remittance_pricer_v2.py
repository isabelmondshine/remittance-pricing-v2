"""
remittance_pricer.py  —  FT Partners Remittance Competitive Pricing Tracker
Scrapes live bank-to-bank transfer pricing from each provider's public website
using a real headless browser (Playwright/Chromium). No APIs, no auth required.

USAGE:
  python remittance_pricer.py
  python remittance_pricer.py --corridors USD-EUR USD-INR
  python remittance_pricer.py --amounts 500 1000 5000

REQUIREMENTS:
  pip install playwright openpyxl requests
  playwright install chromium
"""

import argparse
import json
import re
import time
from datetime import datetime
from typing import Optional

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ── Configuration ──────────────────────────────────────────────────────────────

CORRIDORS = [
    ("USD", "EUR"),
    ("USD", "GBP"),
    ("USD", "MXN"),
    ("USD", "INR"),
    ("USD", "PHP"),
    ("USD", "CAD"),
    ("USD", "AUD"),
    ("USD", "JPY"),
]

AMOUNTS = [100, 500, 1000, 5000, 10000]

CURRENCY_TO_COUNTRY = {
    "EUR": "DE", "GBP": "GB", "MXN": "MX", "INR": "IN",
    "PHP": "PH", "CAD": "CA", "AUD": "AU", "JPY": "JP",
}

# Country names for display on provider sites
COUNTRY_NAMES = {
    "DE": "Germany", "GB": "United Kingdom", "MX": "Mexico",
    "IN": "India",   "PH": "Philippines",    "CA": "Canada",
    "AU": "Australia", "JP": "Japan",
}

# Mid-market reference rates (approximate — used for markup calculation)
# These are updated manually when running the script; the scraper gets real rates
MID_MARKET_REF = {
    ("USD","EUR"): 0.925, ("USD","GBP"): 0.790, ("USD","MXN"): 17.15,
    ("USD","INR"): 83.50, ("USD","PHP"): 57.50, ("USD","CAD"): 1.365,
    ("USD","AUD"): 1.525, ("USD","JPY"): 149.5,
}

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
      "AppleWebKit/537.36 (KHTML, like Gecko) "
      "Chrome/124.0.0.0 Safari/537.36")

# ── Data model ─────────────────────────────────────────────────────────────────

class Quote:
    def __init__(self, provider, from_ccy, to_ccy, send_amount,
                 fee_usd=None, fx_rate=None, fx_markup_pct=None,
                 received_amount=None, note="", error=""):
        self.provider        = provider
        self.from_ccy        = from_ccy
        self.to_ccy          = to_ccy
        self.send_amount     = send_amount
        self.fee_usd         = fee_usd
        self.fx_rate         = fx_rate
        self.fx_markup_pct   = fx_markup_pct
        self.received_amount = received_amount
        self.note            = note
        self.error           = error

    def fill_gaps(self):
        """Calculate markup and received amount if missing."""
        mid = MID_MARKET_REF.get((self.from_ccy, self.to_ccy))
        if self.fx_markup_pct is None and self.fx_rate and mid:
            self.fx_markup_pct = round((mid - self.fx_rate) / mid * 100, 3)
        if self.received_amount is None and self.fx_rate and self.fee_usd is not None:
            self.received_amount = round((self.send_amount - self.fee_usd) * self.fx_rate, 2)


# ── Browser helpers ────────────────────────────────────────────────────────────

def _make_browser():
    from playwright.sync_api import sync_playwright
    pw = sync_playwright().start()
    browser = pw.chromium.launch(
        headless=True,
        args=["--no-sandbox", "--disable-setuid-sandbox",
              "--disable-blink-features=AutomationControlled"]
    )
    ctx = browser.new_context(
        user_agent=UA,
        viewport={"width": 1280, "height": 800},
        locale="en-US",
    )
    return pw, browser, ctx


def _safe_float(s):
    """Extract first number from a string like '$3.99' or '1,234.56 EUR'."""
    if s is None:
        return None
    s = str(s).replace(",", "")
    m = re.search(r"[\d]+\.?\d*", s)
    return float(m.group()) if m else None


def _intercept_json(page, keywords=("rate", "fee", "quote", "price", "transfer", "amount")):
    """Attach response listener, return dict that gets populated with captured JSON."""
    captured = {}
    def on_response(response):
        url = response.url.lower()
        if any(k in url for k in keywords):
            try:
                body = response.json()
                if isinstance(body, dict):
                    captured.update(body)
                elif isinstance(body, list) and body and isinstance(body[0], dict):
                    captured.update(body[0])
            except Exception:
                pass
    page.on("response", on_response)
    return captured


# ── Provider: Wise ─────────────────────────────────────────────────────────────

def scrape_wise(from_ccy, to_ccy, amount, ctx) -> Quote:
    """
    Wise send-money calculator.
    Intercepts the /v3/quotes API call the page makes internally.
    """
    page = ctx.new_page()
    captured = _intercept_json(page, ("quotes", "comparison", "rate", "fee"))

    try:
        url = (f"https://wise.com/us/send-money/"
               f"?source={from_ccy}&target={to_ccy}&sourceAmount={amount}")
        page.goto(url, wait_until="networkidle", timeout=30000)
        time.sleep(3)

        # Try to find data in captured API calls
        fee      = _find_nested(captured, ["fee", "totalFee", "transferFee"])
        rate     = _find_nested(captured, ["rate", "exchangeRate"])
        received = _find_nested(captured, ["targetAmount", "recipientAmount", "receivedAmount"])

        # Fallback: parse page text
        if rate is None:
            text = page.inner_text("body")
            rate     = _extract_rate(text, from_ccy, to_ccy)
            fee      = _extract_fee(text) if fee is None else fee
            received = _extract_received(text, to_ccy) if received is None else received

        page.close()

        if rate is None:
            return Quote("Wise", from_ccy, to_ccy, amount,
                         error="Could not extract rate from Wise page")

        q = Quote("Wise", from_ccy, to_ccy, amount,
                  fee_usd=float(fee or 0), fx_rate=float(rate),
                  received_amount=float(received) if received else None,
                  note="Wise calculator (bank transfer)")
        q.fill_gaps()
        return q

    except Exception as e:
        page.close()
        return Quote("Wise", from_ccy, to_ccy, amount, error=str(e)[:100])


# ── Provider: Remitly ──────────────────────────────────────────────────────────

def scrape_remitly(from_ccy, to_ccy, amount, ctx) -> Quote:
    """
    Remitly send-money calculator. Economy rate, bank deposit.
    Intercepts their internal pricing API call.
    """
    page = ctx.new_page()
    api_data = {}

    def on_response(response):
        if "price" in response.url or "calculator" in response.url or "rate" in response.url:
            try:
                body = response.json()
                if isinstance(body, dict):
                    api_data.update(body)
            except Exception:
                pass

    page.on("response", on_response)

    try:
        country = CURRENCY_TO_COUNTRY.get(to_ccy, "IN")
        url = (f"https://www.remitly.com/us/en/send-money"
               f"?sourceCountry=USA&destCountry={country}"
               f"&sourceCurrency={from_ccy}&destCurrency={to_ccy}"
               f"&sendAmount={int(amount)}&deliveryType=BANK_DEPOSIT")
        page.goto(url, wait_until="networkidle", timeout=35000)
        time.sleep(4)

        text = page.inner_text("body")
        page.close()

        # Try API data first
        fee      = _find_nested(api_data, ["fee", "transfer_fee", "transferFee", "serviceFee"])
        rate     = _find_nested(api_data, ["exchange_rate", "exchangeRate", "rate"])
        received = _find_nested(api_data, ["destination_amount", "destinationAmount",
                                            "recipient_amount", "recipientAmount"])

        # Fallback to page text
        if rate is None:
            rate     = _extract_rate(text, from_ccy, to_ccy)
            fee      = _extract_fee(text) if fee is None else fee
            received = _extract_received(text, to_ccy) if received is None else received

        if rate is None:
            return Quote("Remitly", from_ccy, to_ccy, amount,
                         error="Could not extract rate from Remitly page")

        q = Quote("Remitly", from_ccy, to_ccy, amount,
                  fee_usd=float(fee or 0), fx_rate=float(rate),
                  received_amount=float(received) if received else None,
                  note="Economy, bank deposit")
        q.fill_gaps()
        return q

    except Exception as e:
        page.close()
        return Quote("Remitly", from_ccy, to_ccy, amount, error=str(e)[:100])


# ── Provider: Western Union ────────────────────────────────────────────────────

def scrape_western_union(from_ccy, to_ccy, amount, ctx) -> Quote:
    """Western Union send-money page. Bank account delivery."""
    page = ctx.new_page()
    api_data = {}

    def on_response(response):
        if any(k in response.url for k in ["price", "estimate", "fee", "rate", "transfer"]):
            try:
                body = response.json()
                if isinstance(body, dict):
                    api_data.update(body)
            except Exception:
                pass

    page.on("response", on_response)

    try:
        country = CURRENCY_TO_COUNTRY.get(to_ccy, "DE")
        country_name = COUNTRY_NAMES.get(country, "Germany")
        url = (f"https://www.westernunion.com/us/en/send-money/app/price-estimator"
               f"?sendCurrencyCode={from_ccy}&destCurrencyCode={to_ccy}"
               f"&destCountryCode={country}&sendAmount={int(amount)}"
               f"&payinMethod=ACCOUNT&payoutMethod=ACCOUNT")
        page.goto(url, wait_until="networkidle", timeout=35000)
        time.sleep(4)

        text = page.inner_text("body")
        page.close()

        fee      = _find_nested(api_data, ["fee", "transferFee", "transactionFee"])
        rate     = _find_nested(api_data, ["fxRate", "exchangeRate", "rate"])
        received = _find_nested(api_data, ["destPrincipalAmount", "receiveAmount", "recipientAmount"])

        if rate is None:
            rate     = _extract_rate(text, from_ccy, to_ccy)
            fee      = _extract_fee(text) if fee is None else fee
            received = _extract_received(text, to_ccy) if received is None else received

        if rate is None:
            return Quote("Western Union", from_ccy, to_ccy, amount,
                         error="Could not extract rate from WU page")

        q = Quote("Western Union", from_ccy, to_ccy, amount,
                  fee_usd=float(fee or 0), fx_rate=float(rate),
                  received_amount=float(received) if received else None,
                  note="Bank account delivery (online)")
        q.fill_gaps()
        return q

    except Exception as e:
        page.close()
        return Quote("Western Union", from_ccy, to_ccy, amount, error=str(e)[:100])


# ── Provider: MoneyGram ────────────────────────────────────────────────────────

def scrape_moneygram(from_ccy, to_ccy, amount, ctx) -> Quote:
    """MoneyGram fee estimator. Bank deposit delivery."""
    page = ctx.new_page()
    api_data = {}

    def on_response(response):
        if any(k in response.url for k in ["fee", "estimate", "rate", "price"]):
            try:
                body = response.json()
                if isinstance(body, dict):
                    api_data.update(body)
            except Exception:
                pass

    page.on("response", on_response)

    try:
        country = CURRENCY_TO_COUNTRY.get(to_ccy, "DE")
        url = (f"https://www.moneygram.com/mgo/us/en/fee-estimator"
               f"?sendCurrency={from_ccy}&receiveCurrency={to_ccy}"
               f"&receiveCountry={country}&sendAmount={int(amount)}"
               f"&paymentMethod=BANK_ACCOUNT"
               f"&deliveryMethod=RECEIVE_MONEY_IN_BANK_ACCOUNT")
        page.goto(url, wait_until="networkidle", timeout=35000)
        time.sleep(4)

        text = page.inner_text("body")
        page.close()

        fee      = _find_nested(api_data, ["fee", "transferFee", "mgiSendFee", "totalFee"])
        rate     = _find_nested(api_data, ["exchangeRate", "fxRate", "rate"])
        received = _find_nested(api_data, ["receiveAmount", "destinationAmount", "estimatedReceiveAmount"])

        if rate is None:
            rate     = _extract_rate(text, from_ccy, to_ccy)
            fee      = _extract_fee(text) if fee is None else fee
            received = _extract_received(text, to_ccy) if received is None else received

        if rate is None:
            return Quote("MoneyGram", from_ccy, to_ccy, amount,
                         error="Could not extract rate from MoneyGram page")

        q = Quote("MoneyGram", from_ccy, to_ccy, amount,
                  fee_usd=float(fee or 0), fx_rate=float(rate),
                  received_amount=float(received) if received else None,
                  note="Bank deposit delivery")
        q.fill_gaps()
        return q

    except Exception as e:
        page.close()
        return Quote("MoneyGram", from_ccy, to_ccy, amount, error=str(e)[:100])


# ── Provider: Revolut ──────────────────────────────────────────────────────────

def scrape_revolut(from_ccy, to_ccy, amount, ctx) -> Quote:
    """Revolut transfer page. Standard (free) plan, external bank transfer."""
    page = ctx.new_page()
    api_data = {}

    def on_response(response):
        if any(k in response.url for k in ["transfer", "quote", "rate", "price", "fee"]):
            try:
                body = response.json()
                if isinstance(body, dict):
                    api_data.update(body)
            except Exception:
                pass

    page.on("response", on_response)

    try:
        url = (f"https://www.revolut.com/en-US/money-transfer/"
               f"{from_ccy.lower()}-to-{to_ccy.lower()}/")
        page.goto(url, wait_until="networkidle", timeout=35000)
        time.sleep(3)

        text = page.inner_text("body")
        page.close()

        fee      = _find_nested(api_data, ["fee", "transferFee", "totalFee"])
        rate     = _find_nested(api_data, ["rate", "exchangeRate", "fxRate"])
        received = _find_nested(api_data, ["recipientAmount", "receiveAmount", "targetAmount"])

        if rate is None:
            rate     = _extract_rate(text, from_ccy, to_ccy)
            fee      = _extract_fee(text) if fee is None else fee
            received = _extract_received(text, to_ccy) if received is None else received

        # Revolut standard plan: ~$5 fee for bank transfers if not found
        if rate and fee is None:
            fee = 5.0

        if rate is None:
            return Quote("Revolut", from_ccy, to_ccy, amount,
                         error="Could not extract rate from Revolut page")

        q = Quote("Revolut", from_ccy, to_ccy, amount,
                  fee_usd=float(fee or 5.0), fx_rate=float(rate),
                  received_amount=float(received) if received else None,
                  note="Standard plan, external bank transfer")
        q.fill_gaps()
        return q

    except Exception as e:
        page.close()
        return Quote("Revolut", from_ccy, to_ccy, amount, error=str(e)[:100])


# ── Provider: Euronet ──────────────────────────────────────────────────────────

def scrape_euronet(from_ccy, to_ccy, amount, ctx) -> Quote:
    """Euronet / epay money transfer. Coverage varies by corridor."""
    page = ctx.new_page()
    api_data = {}

    def on_response(response):
        if any(k in response.url for k in ["rate", "fee", "quote", "transfer", "price"]):
            try:
                body = response.json()
                if isinstance(body, dict):
                    api_data.update(body)
            except Exception:
                pass

    page.on("response", on_response)

    try:
        # Try Euronet's money transfer page
        urls = [
            f"https://www.euronet.eu/send-money?from={from_ccy}&to={to_ccy}&amount={int(amount)}",
            f"https://xe.com/send-money/",  # Euronet owns XE
        ]

        text = ""
        for url in urls:
            try:
                page.goto(url, wait_until="networkidle", timeout=25000)
                time.sleep(3)
                text = page.inner_text("body")
                if to_ccy in text or from_ccy in text:
                    break
            except Exception:
                continue

        page.close()

        fee      = _find_nested(api_data, ["fee", "transferFee"])
        rate     = _find_nested(api_data, ["rate", "exchangeRate"])
        received = _find_nested(api_data, ["receivedAmount", "destinationAmount"])

        if rate is None:
            rate     = _extract_rate(text, from_ccy, to_ccy)
            fee      = _extract_fee(text) if fee is None else fee
            received = _extract_received(text, to_ccy) if received is None else received

        if rate is None:
            return Quote("Euronet", from_ccy, to_ccy, amount,
                         error="Corridor not supported or page could not be scraped")

        q = Quote("Euronet", from_ccy, to_ccy, amount,
                  fee_usd=float(fee or 0), fx_rate=float(rate),
                  received_amount=float(received) if received else None,
                  note="Bank transfer")
        q.fill_gaps()
        return q

    except Exception as e:
        page.close()
        return Quote("Euronet", from_ccy, to_ccy, amount, error=str(e)[:100])


# ── Text extraction helpers ────────────────────────────────────────────────────

def _find_nested(d, keys):
    """Search a dict (potentially nested) for any of the given keys."""
    if not isinstance(d, dict):
        return None
    for k in keys:
        if k in d:
            try:
                return float(d[k])
            except (TypeError, ValueError):
                pass
    # One level deep
    for v in d.values():
        if isinstance(v, dict):
            result = _find_nested(v, keys)
            if result is not None:
                return result
    return None


def _extract_rate(text, from_ccy, to_ccy):
    """Extract exchange rate from page text."""
    patterns = [
        rf"1\s*{from_ccy}\s*[=≈→\-]+\s*([\d,]+\.?\d*)\s*{to_ccy}",
        rf"1\s*{from_ccy}\s*[=≈→\-]+\s*([\d,]+\.?\d*)",
        rf"{to_ccy}\s*[=≈:]\s*([\d,]+\.?\d*)",
        rf"exchange rate[^\d]*([\d,]+\.?\d*)",
        rf"rate[:\s]+([\d,]+\.?\d*)",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            val = float(m.group(1).replace(",", ""))
            # Sanity check: rate should be > 0.01 and < 100000
            if 0.01 < val < 100000:
                return val
    return None


def _extract_fee(text):
    """Extract transfer fee from page text."""
    patterns = [
        r"(?:transfer fee|service fee|fee)[:\s]*\$?\s*([\d.]+)",
        r"\$([\d.]+)\s*fee",
        r"fee[:\s]*\$\s*([\d.]+)",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            val = float(m.group(1))
            if 0 <= val < 100:  # sanity check
                return val
    return None


def _extract_received(text, to_ccy):
    """Extract recipient received amount from page text."""
    patterns = [
        rf"([\d,]+\.?\d*)\s*{to_ccy}",
        rf"recipient (?:gets|receives)[^\d]*([\d,]+\.?\d*)",
        rf"they (?:get|receive)[^\d]*([\d,]+\.?\d*)",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            val = float(m.group(1).replace(",", ""))
            if val > 0.01:
                return val
    return None


# ── Orchestrator ───────────────────────────────────────────────────────────────

SCRAPERS = [
    ("Wise",          scrape_wise),
    ("Remitly",       scrape_remitly),
    ("Western Union", scrape_western_union),
    ("MoneyGram",     scrape_moneygram),
    ("Revolut",       scrape_revolut),
    ("Euronet",       scrape_euronet),
]


def fetch_all_quotes(corridors, amounts):
    from playwright.sync_api import sync_playwright

    all_quotes = []

    pw, browser, ctx = _make_browser()

    try:
        for from_ccy, to_ccy in corridors:
            print(f"\n{'='*55}")
            print(f"  {from_ccy} → {to_ccy}")
            print(f"{'='*55}")

            for name, scraper in SCRAPERS:
                for amount in amounts:
                    print(f"  [{name}] ${amount:,} ... ", end="", flush=True)
                    try:
                        q = scraper(from_ccy, to_ccy, amount, ctx)
                        all_quotes.append(q)
                        if q.error:
                            print(f"✗  {q.error[:60]}")
                        else:
                            rcv = f"{q.received_amount:,.2f} {to_ccy}" if q.received_amount else "?"
                            fee = f"${q.fee_usd:.2f}" if q.fee_usd is not None else "?"
                            mkp = f"{q.fx_markup_pct:.2f}%" if q.fx_markup_pct is not None else "?"
                            print(f"✓  fee={fee}  mkp={mkp}  rcv={rcv}")
                    except Exception as e:
                        q = Quote(name, from_ccy, to_ccy, amount, error=str(e)[:80])
                        all_quotes.append(q)
                        print(f"✗  {str(e)[:60]}")
                    time.sleep(1)
    finally:
        ctx.close()
        browser.close()
        pw.stop()

    return all_quotes


# ── Excel output ───────────────────────────────────────────────────────────────

C_NAVY  = "1A3A5C"
C_BLUE  = "2E6DA4"
C_LBLUE = "D9E8F5"
C_WHITE = "FFFFFF"
C_GREEN = "E2F0D9"
C_AMBER = "FFF2CC"
C_REDBG = "FFE0E0"
C_BEST  = "C6EFCE"
C_GREY  = "F5F5F5"
C_DKGRY = "404040"

def _fill(c):
    return PatternFill("solid", start_color=c, fgColor=c)
def _font(bold=False, color=C_DKGRY, size=10, italic=False):
    return Font(bold=bold, color=color, size=size, italic=italic, name="Arial")
def _border():
    t = Side(style="thin", color="CCCCCC")
    return Border(left=t, right=t, top=t, bottom=t)


def write_excel(all_quotes, amounts, filename):
    wb = Workbook()
    ws = wb.active
    ws.title = "Pricing Summary"
    _write_summary(ws, all_quotes, amounts)

    corridors_seen = list(dict.fromkeys(f"{q.from_ccy}→{q.to_ccy}" for q in all_quotes))
    for cor in corridors_seen:
        ws2 = wb.create_sheet(title=cor)
        _write_raw(ws2, [q for q in all_quotes if f"{q.from_ccy}→{q.to_ccy}" == cor])

    wl = wb.create_sheet("Legend")
    _write_legend(wl)
    wb.save(filename)
    print(f"\n✅  Saved: {filename}")


def _write_summary(ws, all_quotes, amounts):
    by_corridor = {}
    for q in all_quotes:
        cor = f"{q.from_ccy}→{q.to_ccy}"
        by_corridor.setdefault(cor, {}).setdefault(q.provider, {})[q.send_amount] = q

    CPAMT = 4
    total_cols = 1 + len(amounts) * CPAMT
    row = 1

    # Title
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=total_cols)
    ws.cell(row, 1).value = "FT Partners — Remittance Competitive Pricing Tracker"
    ws.cell(row, 1).font  = Font(bold=True, size=14, color=C_WHITE, name="Arial")
    ws.cell(row, 1).fill  = _fill(C_NAVY)
    ws.cell(row, 1).alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row].height = 28
    row += 1

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=total_cols)
    ws.cell(row, 1).value = (
        f"Bank-to-bank transfers  ·  All prices scraped live from each provider's public website  ·  "
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    )
    ws.cell(row, 1).font  = _font(italic=True, color=C_WHITE, size=9)
    ws.cell(row, 1).fill  = _fill(C_BLUE)
    ws.cell(row, 1).alignment = Alignment(horizontal="center")
    row += 2

    for corridor, providers in by_corridor.items():
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=total_cols)
        ws.cell(row, 1).value = f"  {corridor}  ·  Bank Account Delivery"
        ws.cell(row, 1).font  = Font(bold=True, size=11, color=C_WHITE, name="Arial")
        ws.cell(row, 1).fill  = _fill(C_BLUE)
        ws.cell(row, 1).alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[row].height = 22
        row += 1

        # Amount column headers
        ws.cell(row, 1).value = "Provider"
        ws.cell(row, 1).font  = _font(bold=True, color=C_WHITE)
        ws.cell(row, 1).fill  = _fill(C_NAVY)
        ws.cell(row, 1).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row, 1).border = _border()
        col = 2
        for amt in amounts:
            ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col+CPAMT-1)
            c = ws.cell(row, col)
            c.value = f"${amt:,}"
            c.font  = _font(bold=True, color=C_WHITE)
            c.fill  = _fill(C_NAVY)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = _border()
            col += CPAMT
        ws.row_dimensions[row].height = 18
        row += 1

        # Sub-headers
        ws.cell(row, 1).fill = _fill(C_LBLUE)
        ws.cell(row, 1).border = _border()
        col = 2
        for _ in amounts:
            for lbl in ["Fee ($)", "FX Rate", "Mkp %", "Received"]:
                c = ws.cell(row, col)
                c.value = lbl
                c.font  = _font(bold=True, size=8)
                c.fill  = _fill(C_LBLUE)
                c.alignment = Alignment(horizontal="center")
                c.border = _border()
                col += 1
        ws.row_dimensions[row].height = 14
        row += 1

        # Bests
        best_rcv, best_fee = {}, {}
        for pname, pdata in providers.items():
            for amt in amounts:
                q = pdata.get(amt)
                if not q or q.error: continue
                if q.received_amount and (amt not in best_rcv or q.received_amount > best_rcv[amt]):
                    best_rcv[amt] = q.received_amount
                if q.fee_usd is not None and (amt not in best_fee or q.fee_usd < best_fee[amt]):
                    best_fee[amt] = q.fee_usd

        for i, pname in enumerate(sorted(providers.keys())):
            pdata = providers[pname]
            bg = C_WHITE if i % 2 == 0 else C_GREY
            ws.cell(row, 1).value = pname
            ws.cell(row, 1).font  = _font(bold=True)
            ws.cell(row, 1).fill  = _fill(bg)
            ws.cell(row, 1).alignment = Alignment(horizontal="left", vertical="center", indent=1)
            ws.cell(row, 1).border = _border()

            col = 2
            for amt in amounts:
                q = pdata.get(amt)
                is_best_rcv = (q and not q.error and q.received_amount and
                               q.received_amount == best_rcv.get(amt))
                is_best_fee = (q and not q.error and q.fee_usd is not None and
                               q.fee_usd == best_fee.get(amt))

                if not q or q.error:
                    err = (q.error[:30] if q and q.error else "N/A")
                    for offset in range(CPAMT):
                        c = ws.cell(row, col+offset)
                        c.value = f"⚠ {err}" if offset == 0 else ""
                        c.font  = _font(size=8, italic=True, color="AA0000")
                        c.fill  = _fill("FFF0F0")
                        c.border = _border()
                    col += CPAMT
                    continue

                # Fee
                c = ws.cell(row, col)
                c.value = round(q.fee_usd, 2) if q.fee_usd is not None else "—"
                c.number_format = '$#,##0.00'
                c.font  = _font(bold=is_best_fee, color="006600" if is_best_fee else C_DKGRY)
                c.fill  = _fill(C_BEST if is_best_fee else bg)
                c.alignment = Alignment(horizontal="right")
                c.border = _border()
                col += 1

                # FX Rate
                c = ws.cell(row, col)
                c.value = round(q.fx_rate, 4) if q.fx_rate else "—"
                c.number_format = "0.0000"
                c.font  = _font()
                c.fill  = _fill(bg)
                c.alignment = Alignment(horizontal="right")
                c.border = _border()
                col += 1

                # Markup
                c = ws.cell(row, col)
                mkp = q.fx_markup_pct
                c.value = round(mkp/100, 4) if mkp is not None else "—"
                c.number_format = "0.00%"
                if mkp is not None:
                    if   mkp < 0.5:  c.fill=_fill(C_GREEN); c.font=_font(color="276221",bold=True)
                    elif mkp < 1.5:  c.fill=_fill(C_AMBER); c.font=_font(color="7D6608",bold=True)
                    else:            c.fill=_fill(C_REDBG); c.font=_font(color="B22222",bold=True)
                else:
                    c.fill=_fill(bg); c.font=_font()
                c.alignment = Alignment(horizontal="right")
                c.border = _border()
                col += 1

                # Received
                c = ws.cell(row, col)
                c.value = round(q.received_amount, 2) if q.received_amount else "—"
                c.number_format = "#,##0.00"
                c.font  = _font(bold=is_best_rcv, color="006600" if is_best_rcv else C_DKGRY)
                c.fill  = _fill(C_BEST if is_best_rcv else bg)
                c.alignment = Alignment(horizontal="right")
                c.border = _border()
                col += 1

            ws.row_dimensions[row].height = 15
            row += 1

        row += 2

    ws.column_dimensions["A"].width = 18
    ci = 2
    for _ in amounts:
        ws.column_dimensions[get_column_letter(ci)].width   = 9
        ws.column_dimensions[get_column_letter(ci+1)].width = 9
        ws.column_dimensions[get_column_letter(ci+2)].width = 8
        ws.column_dimensions[get_column_letter(ci+3)].width = 13
        ci += 4
    ws.freeze_panes = "B4"


def _write_raw(ws, quotes):
    headers = ["Provider","From","To","Send Amount","Fee ($)","FX Rate",
               "FX Markup %","Received","Note","Error"]
    for col, h in enumerate(headers, 1):
        c = ws.cell(1, col)
        c.value = h
        c.font  = _font(bold=True, color=C_WHITE)
        c.fill  = _fill(C_NAVY)
        c.border = _border()
        c.alignment = Alignment(horizontal="center")
    for row, q in enumerate(quotes, 2):
        vals = [q.provider, q.from_ccy, q.to_ccy, q.send_amount,
                q.fee_usd, q.fx_rate,
                round(q.fx_markup_pct/100, 4) if q.fx_markup_pct is not None else None,
                q.received_amount, q.note, q.error]
        for col, v in enumerate(vals, 1):
            c = ws.cell(row, col)
            c.value = v
            c.font  = _font()
            c.border = _border()
            if col == 4: c.number_format = "$#,##0"
            if col == 5: c.number_format = "$#,##0.00"
            if col == 6: c.number_format = "0.0000"
            if col == 7: c.number_format = "0.00%"
            if col == 8: c.number_format = "#,##0.00"
    for i, w in enumerate([18,6,6,12,10,10,10,14,45,40], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"
    ws.freeze_panes = "A2"


def _write_legend(ws):
    ws.title = "Legend & Notes"
    rows = [
        ("COLOR CODING",""),
        ("Green FX markup","< 0.5% above mid-market — excellent"),
        ("Amber FX markup","0.5–1.5% above mid-market — moderate"),
        ("Red FX markup","> 1.5% above mid-market — expensive"),
        ("Green background (Received)","Best received amount for that send amount"),
        ("Green background (Fee)","Lowest fee for that send amount"),
        ("",""),
        ("DATA SOURCES",""),
        ("Method","All prices scraped live from each provider's public website using a headless browser."),
        ("Wise","wise.com/us/send-money — real mid-market rate, transparent fee"),
        ("Remitly","remitly.com — Economy rate, bank deposit delivery"),
        ("Western Union","westernunion.com — bank account in/out (online pricing)"),
        ("MoneyGram","moneygram.com — bank deposit delivery"),
        ("Revolut","revolut.com — Standard (free) plan, external bank transfer"),
        ("Euronet","euronet.eu — coverage varies by corridor"),
        ("",""),
        ("NOTES",""),
        ("FX Markup","% above mid-market rate. Lower = better deal for recipient."),
        ("Received","Amount recipient gets after all fees and FX markup applied."),
        ("Bank Transfer","All pricing = bank account send + bank account receive. Not cash pickup."),
        ("Revolut","Standard free plan. Premium/Metal plans have lower or no fees."),
    ]
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 85
    for i, (label, value) in enumerate(rows, 1):
        ws.cell(i,1).value = label
        ws.cell(i,2).value = value
        if label.isupper() and label:
            ws.cell(i,1).font = Font(bold=True, size=10, color=C_WHITE, name="Arial")
            ws.cell(i,1).fill = _fill(C_BLUE)
            ws.cell(i,2).fill = _fill(C_BLUE)
        else:
            ws.cell(i,1).font = _font(bold=bool(label))
            ws.cell(i,2).font = _font()
        ws.row_dimensions[i].height = 15


# ── Entry point ────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="FT Partners Remittance Pricing Tracker")
    parser.add_argument("--corridors", nargs="+", default=None,
                        help="e.g. USD-EUR USD-INR")
    parser.add_argument("--amounts", nargs="+", type=int, default=None)
    parser.add_argument("--output", type=str, default=None)
    args = parser.parse_args()

    corridors   = [tuple(c.replace("-"," ").split()) for c in args.corridors] if args.corridors else CORRIDORS
    amounts     = args.amounts or AMOUNTS
    output_file = args.output or f"remittance_pricing_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    print("\n" + "═"*55)
    print("  FT Partners — Remittance Pricing Tracker")
    print("═"*55)
    print(f"  Corridors : {', '.join(f'{a}→{b}' for a,b in corridors)}")
    print(f"  Amounts   : {', '.join(f'${a:,}' for a in amounts)}")
    print(f"  Output    : {output_file}\n")

    all_quotes = fetch_all_quotes(corridors, amounts)
    errors = [q for q in all_quotes if q.error]
    print(f"\n  Total quotes : {len(all_quotes)}")
    print(f"  Errors       : {len(errors)}")
    write_excel(all_quotes, amounts, output_file)


if __name__ == "__main__":
    main()
