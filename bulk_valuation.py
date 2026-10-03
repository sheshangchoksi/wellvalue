"""
bulk_valuation.py
==================
Bulk DCF Valuation mode: run a fast-and-forgiving DCF (FCFF) across a batch
of companies sourced from Screener.in, using ONE user-supplied beta for the
whole batch (no per-stock peer-beta calculation, no comparative valuation).

Design goals (per product spec):
  1. Tickers from the existing NSE/BSE dropdown lists, OR pasted manually,
     OR both mixed together in one run.
  2. For each ticker: try the Screener "consolidated" page first. Only fall
     back to "standalone" if consolidated genuinely has no usable P&L/BS
     data — NOT just because the HTTP status was 200 (Screener returns 200
     for the consolidated URL of companies that don't have any consolidated
     financials, e.g. screener.in/company/DAVANGERE/consolidated/ — the
     page renders fine, it just has nothing in the consolidated tables).
     This module checks for actual extracted numbers, not status codes.
  3. Stock price is only needed for the current-price-vs-fair-value display,
     never for the DCF math itself (WACC here uses book equity + a manual
     beta, not market cap) — so a price-fetch failure must never block a
     company's DCF result. Fallback chain: Screener's own "Current Price"
     ratio (already on the page we scraped — zero extra requests) →
     yfinance via yf_ratelimit (rate-limit-shielded) → BSE official quote
     API (for companies with a BSE code) → mark "N/A" and move on.
  4. One company's failure at any stage must never abort the batch. Every
     stage is isolated in try/except, and each company's full processing
     (including any nested st.warning/st.info emitted by the shared DCF
     functions) is captured inside its own collapsed st.status(...) block,
     so the page stays readable across dozens of companies while every
     failure is still one click away from full detail.

This module is intentionally decoupled from PHASE5_DCF_valuation.py: it
takes the DCF engine functions (classify_business_model,
calculate_working_capital_metrics, project_financials, calculate_wacc,
calculate_dcf_valuation, ensure_valid_number) as arguments to
render_bulk_valuation_ui(...) instead of importing them, because
PHASE5_DCF_valuation.py is the one that imports and calls this module —
a direct two-way import would be circular.
"""

from __future__ import annotations

import random
import re
import time
import traceback
from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional

import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st
from bs4 import BeautifulSoup

import ticker_dropdown_utils as tdu

try:
    from yf_ratelimit import safe_ticker
    _HAS_YF_SHIELD = True
except Exception:
    _HAS_YF_SHIELD = False

try:
    from utils_indian_apis import get_bse_quote
    _HAS_BSE_QUOTE = True
except Exception:
    _HAS_BSE_QUOTE = False


CR_TO_LAC = 100.0  # Screener publishes ₹ Crores; the shared DCF engine works in ₹ Lacs.
                    # 1 Crore = 100 Lacs (confirmed against screener_excel_mode.py's
                    # `revenue * 100` conversion, the path your working Excel-upload
                    # mode uses). Do NOT change this back to 10 — that was a bug
                    # inherited from a stale comment in utils_indian_apis.py and is
                    # what caused every monetary figure here to be 10x too small.
SCREENER_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}


# ══════════════════════════════════════════════════════════════════════════
# RESULT ROW
# ══════════════════════════════════════════════════════════════════════════

@dataclass
class BulkRow:
    display_name: str
    symbol: str
    exchange: str                       # "NSE" | "BSE" | "MANUAL"
    status: str = "pending"             # pending | success | skipped | error
    stage_failed: Optional[str] = None  # which pipeline stage failed, for debugging
    error_detail: str = ""
    data_source: str = ""               # "Consolidated" | "Standalone"
    company_name: str = ""
    current_price: Optional[float] = None
    price_source: str = ""
    fair_value_per_share: Optional[float] = None
    upside_pct: Optional[float] = None
    enterprise_value_cr: Optional[float] = None
    equity_value_cr: Optional[float] = None
    wacc_pct: Optional[float] = None
    shares_outstanding: Optional[int] = None
    debug_log: list = field(default_factory=list)

    def log(self, msg: str):
        self.debug_log.append(msg)


# ══════════════════════════════════════════════════════════════════════════
# TICKER RESOLUTION
# ══════════════════════════════════════════════════════════════════════════

def _resolve_picked_options(display_options: list[str]) -> list[BulkRow]:
    """Turn 'Name (NSE: TICKER)' / 'Name (BSE: CODE)' display strings (from
    the shared dropdown list) into BulkRow stubs with a Screener-ready symbol."""
    rows = []
    for disp in display_options:
        full = tdu.resolve_peer_display_to_full_ticker(disp)  # "TCS.NS" / "500325.BO"
        if not full or "." not in full:
            continue
        raw, suffix = full.rsplit(".", 1)
        exchange = "NSE" if suffix.upper() == "NS" else "BSE"
        name = disp[: disp.rfind("(")].strip() if "(" in disp else disp
        rows.append(BulkRow(display_name=name, symbol=raw, exchange=exchange))
    return rows


def _resolve_manual_tickers(raw_text: str) -> list[BulkRow]:
    """Free-text tickers/codes, one per line or comma-separated. Numeric-only
    tokens are treated as BSE codes (matching the user's own example:
    'https://www.screener.in/company/500325/'), everything else as a plain
    Screener symbol (NSE ticker or any other slug Screener accepts)."""
    if not raw_text or not raw_text.strip():
        return []
    tokens = re.split(r"[,\n\r\t]+", raw_text)
    rows = []
    for tok in tokens:
        sym = tok.strip().upper()
        if not sym:
            continue
        exchange = "BSE" if sym.isdigit() else "MANUAL"
        rows.append(BulkRow(display_name=sym, symbol=sym, exchange=exchange))
    return rows


def _dedupe(rows: list[BulkRow]) -> list[BulkRow]:
    seen, out = set(), []
    for r in rows:
        key = r.symbol.upper()
        if key in seen:
            continue
        seen.add(key)
        out.append(r)
    return out


# ══════════════════════════════════════════════════════════════════════════
# SCREENER FETCH — WITH REAL consolidated → standalone DETECTION
# ══════════════════════════════════════════════════════════════════════════

def _get_soup(url: str, session: requests.Session, timeout: int = 20):
    resp = session.get(url, headers=SCREENER_HEADERS, timeout=timeout)
    if resp.status_code != 200:
        return None, f"HTTP {resp.status_code} for {url}"
    return BeautifulSoup(resp.content, "lxml"), None


def _find_table_by_heading(soup, heading_keywords):
    for heading in soup.find_all(["h2", "h3"]):
        text_lower = heading.get_text(strip=True).lower()
        if all(kw.lower() in text_lower for kw in heading_keywords):
            sibling = heading.find_next_sibling()
            while sibling:
                if sibling.name == "table":
                    return sibling
                tbl = sibling.find("table") if sibling.name else None
                if tbl:
                    return tbl
                sibling = sibling.find_next_sibling()
    return None


def _table_header_years(table):
    """Column headers of a table's first row, e.g. ['Mar 2015', ..., 'TTM']."""
    if table is None:
        return []
    first_row = table.find("tr")
    if first_row is None:
        return []
    return [c.get_text(strip=True) for c in first_row.find_all(["td", "th"])]


def _is_annual_table(table) -> bool:
    """Screener's annual P&L/BS tables span many 'Mon YYYY' columns (often
    10-12 years) and the P&L one ends in 'TTM'. The Quarterly Results table
    only ever has ~8-13 recent quarters and never 'TTM'. A wide span of
    distinct years is what actually distinguishes them — both use 'Mar' in
    column labels, so checking for the word alone isn't enough."""
    headers = _table_header_years(table)
    if not headers:
        return False
    years_seen = set(re.findall(r"(20\d{2})", " ".join(headers)))
    return len(years_seen) >= 5  # quarterly tables cover at most ~3-4 distinct years


def _find_screener_section_table(soup, section_id: str, heading_keywords, must_be_annual: bool):
    """Primary strategy: Screener's real page anchors each section with a
    stable id (id="profit-loss", id="balance-sheet" — visible in the page's
    own nav links). That's far more reliable than matching on heading text
    or table position, and is what fixes a real bug found on RELIANCE: the
    Quarterly Results table sits *before* the annual Profit & Loss table in
    the DOM, so a naive 'first table found' or fragile heading-sibling
    search can silently grab quarterly figures instead of annual ones —
    producing wildly inconsistent errors from company to company (exactly
    the 'random-looking' pattern reported), not a uniform scale error."""
    table = None
    section = soup.find(id=section_id)
    if section is not None:
        table = section.find("table") if section.name != "table" else section
        if table is None:
            table = section.find_next("table")

    if table is None or (must_be_annual and not _is_annual_table(table)):
        # Fall back to heading-text search, but only accept it if it also
        # passes the annual-vs-quarterly content check (when applicable).
        candidate = _find_table_by_heading(soup, heading_keywords)
        if candidate is not None and (not must_be_annual or _is_annual_table(candidate)):
            table = candidate

    if must_be_annual and table is not None and not _is_annual_table(table):
        # Last resort: scan every table on the page for one that both
        # matches the keywords nearby AND looks annual. Never silently fall
        # back to "first table on the page" — that's exactly how quarterly
        # data slipped in undetected before.
        table = None

    return table


def _parse_row(table, keywords):
    if table is None:
        return []
    for tr in table.find_all("tr"):
        cells = tr.find_all(["td", "th"])
        if not cells:
            continue
        label = cells[0].get_text(strip=True).lower()
        label = label.replace("\xa0", " ").replace("–", "").replace("-", "").strip()
        for kw in keywords:
            if kw.lower() in label:
                values = []
                for cell in cells[1:]:
                    raw = cell.get_text(strip=True).replace(",", "").replace("\xa0", "")
                    try:
                        values.append(float(raw))
                    except (ValueError, TypeError):
                        values.append(0.0)
                return values
    return []


def _safe_parse(table, keyword_variants):
    for kws in keyword_variants:
        result = _parse_row(table, kws)
        if result:
            return result
    return []


def _pad(lst, n):
    lst = [v for v in lst if v is not None]
    if len(lst) < n:
        lst = [0.0] * (n - len(lst)) + lst
    return lst[-n:]


def _extract_ratio_value(soup, label: str) -> Optional[float]:
    """Best-effort parse of a labeled value in Screener's top ratio list
    (e.g. 'Current Price', 'Market Cap'). Multiple strategies since Screener's
    exact markup can shift; every strategy is defensive and returns None
    (never raises) on no match."""
    label_lower = label.lower()
    try:
        for li in soup.find_all("li"):
            name_el = li.find(class_=re.compile("name"))
            if not name_el:
                continue
            if label_lower not in name_el.get_text(strip=True).lower():
                continue
            num_el = li.find(class_=re.compile("number"))
            if num_el:
                raw = num_el.get_text(strip=True).replace(",", "")
                try:
                    return float(raw)
                except ValueError:
                    pass
            m = re.search(r"[\d,]+\.?\d*", li.get_text(" ", strip=True).replace(label, ""))
            if m:
                try:
                    return float(m.group(0).replace(",", ""))
                except ValueError:
                    pass
    except Exception:
        pass
    try:
        text = soup.get_text(" ", strip=True)
        m = re.search(re.escape(label) + r"\D{0,15}?([\d,]+\.?\d{0,2})", text)
        if m:
            return float(m.group(1).replace(",", ""))
    except Exception:
        pass
    return None


def _extract_current_price_from_soup(soup) -> Optional[float]:
    """Screener's 'Current Price' ratio — only affects the price-comparison
    display, never blocks a DCF result if it comes back None."""
    return _extract_ratio_value(soup, "Current Price")


def _extract_shares_from_market_cap(soup, price: Optional[float]) -> Optional[int]:
    """Shares outstanding = Market Cap ÷ Current Price, both already printed
    on the page. Preferred over back-solving from Net Profit ÷ EPS, which
    overstates share count for companies with material minority interests
    (consolidated Net Profit includes minority profit; EPS doesn't) — e.g.
    for RELIANCE this formula gives ~1,363 Cr shares (matches Screener's own
    Market Cap ÷ Price), while Net Profit/EPS overstates it to ~1,597 Cr."""
    mcap_cr = _extract_ratio_value(soup, "Market Cap")
    if mcap_cr is None or not price or price <= 0:
        return None
    try:
        return int((mcap_cr * 1_00_00_000) / price)  # ₹ Cr -> ₹, then / price
    except Exception:
        return None


def _parse_screener_page(soup, num_years: int):
    """Returns (financials_dict_or_None, shares, company_name, reason_if_failed).
    financials_dict is None (not an empty-but-present dict) when the page
    had no usable P&L/BS data — this is the signal the caller uses to decide
    whether to retry with the standalone URL."""
    title_tag = soup.find("h1")
    company_name = title_tag.get_text(strip=True) if title_tag else ""

    pl_table = _find_screener_section_table(soup, "profit-loss", ["profit", "loss"], must_be_annual=True)
    bs_table = _find_screener_section_table(soup, "balance-sheet", ["balance", "sheet"], must_be_annual=False)

    if pl_table is None or bs_table is None:
        return None, 0, company_name, (
            "Could not confidently locate the annual Profit & Loss / Balance Sheet tables "
            "(section anchors missing and heading-based search either found nothing or "
            "matched what looks like the Quarterly Results table instead of annual data)."
        )

    raw_revenue = _parse_row(pl_table, ["revenue", "sales"])
    raw_interest = _parse_row(pl_table, ["interest"])
    raw_expenses = _parse_row(pl_table, ["expenses"])
    raw_depreciation = _parse_row(pl_table, ["depreciation"])
    raw_pbt = _parse_row(pl_table, ["profit before tax"])
    raw_tax_pct = _parse_row(pl_table, ["tax %"])
    raw_net_profit = _parse_row(pl_table, ["net profit"])
    raw_eps = _parse_row(pl_table, ["eps in rs", "eps"])

    # THE KEY CHECK: consolidated URL returned HTTP 200 but has no real
    # revenue data (e.g. DAVANGERE-style single-entity companies) → treat as
    # a failed parse so the caller retries standalone.
    if not raw_revenue or all(v == 0 for v in raw_revenue):
        return None, 0, company_name, "P&L table found but Revenue/Sales row is empty or all-zero"

    raw_equity_capital = _safe_parse(bs_table, [["equity capital"], ["equity share capital"], ["share capital"]])
    raw_reserves = _safe_parse(bs_table, [["reserves"], ["reserves and surplus"], ["other equity"]])
    raw_borrowing = _safe_parse(bs_table, [["borrowing"], ["borrowings"], ["total borrowings"], ["debt"]])
    raw_payables = _safe_parse(bs_table, [["trade payables"], ["payables"], ["accounts payable"], ["creditors"]])
    raw_receivables = _safe_parse(bs_table, [["trade receivables"], ["receivables"], ["accounts receivable"], ["debtors"]])
    raw_gross_block = _safe_parse(bs_table, [["gross block"], ["fixed assets"], ["property plant equipment"], ["ppe"]])
    raw_accum_dep = _safe_parse(bs_table, [["accumulated depreciation"], ["depreciation"]])
    raw_cash = _safe_parse(bs_table, [["cash equivalents"], ["cash and bank"], ["cash"], ["bank balances"]])
    raw_inventory = _safe_parse(bs_table, [["inventories"], ["inventory"], ["stock"]])

    n = num_years
    revenue = _pad(raw_revenue, n)
    interest = _pad(raw_interest, n)
    expenses = _pad(raw_expenses, n)
    depreciation = _pad(raw_depreciation, n)
    pbt = _pad(raw_pbt, n)
    tax_pct = _pad(raw_tax_pct, n)
    net_profit = _pad(raw_net_profit, n)
    eps = _pad(raw_eps, n)
    equity_capital = _pad(raw_equity_capital, n)
    reserves = _pad(raw_reserves, n)
    borrowing = _pad(raw_borrowing, n)
    payables = _pad(raw_payables, n)
    receivables = _pad(raw_receivables, n)
    gross_block = _pad(raw_gross_block, n)
    accum_dep = _pad(raw_accum_dep, n)
    cash_vals = _pad(raw_cash, n)
    inventory_vals = _pad(raw_inventory, n)

    price_for_shares = _extract_current_price_from_soup(soup)
    shares = _extract_shares_from_market_cap(soup, price_for_shares) or 0
    shares_source = "Market Cap ÷ Price" if shares else ""
    if not shares:
        # Fallback: back-solve from Net Profit ÷ EPS. Less accurate for
        # companies with material minority interests (consolidated Net
        # Profit includes minority profit; EPS doesn't), but works when
        # Market Cap isn't parseable from the page.
        for i in range(n - 1, -1, -1):
            if eps[i] != 0 and net_profit[i] != 0:
                shares = int((net_profit[i] * 10_000_000) / eps[i])
                shares_source = "Net Profit ÷ EPS (fallback)"
                break

    financials_out = {
        "years": [str(datetime.now().year - i) for i in range(n)],
        "revenue": [], "cogs": [], "opex": [], "ebitda": [], "depreciation": [],
        "ebit": [], "interest": [], "interest_income": [], "tax": [], "nopat": [],
        "fixed_assets": [], "inventory": [], "receivables": [], "payables": [],
        "cash": [], "equity": [], "st_debt": [], "lt_debt": [],
    }

    for i in range(n - 1, -1, -1):
        rev = revenue[i] * CR_TO_LAC
        dep_val = depreciation[i] * CR_TO_LAC
        int_val = interest[i] * CR_TO_LAC
        pbt_val = pbt[i] * CR_TO_LAC
        ebitda_val = pbt_val + int_val + dep_val
        cogs_val = rev * 0.55 if rev > 0 else 0.0
        opex_val = rev - cogs_val - ebitda_val
        if opex_val < 0:
            opex_val = expenses[i] * CR_TO_LAC
            cogs_val = rev - opex_val - ebitda_val
            if cogs_val < 0:
                cogs_val = 0.0
                opex_val = rev - ebitda_val
        ebit_val = ebitda_val - dep_val
        t_rate = tax_pct[i]
        if t_rate > 1:
            t_rate = t_rate / 100.0
        t_rate = max(0.0, min(t_rate, 0.40))
        tax_val = pbt_val * t_rate
        nopat_val = ebit_val * (1 - t_rate)
        eq_val = (equity_capital[i] + reserves[i]) * CR_TO_LAC
        fa_val = (gross_block[i] - accum_dep[i]) * CR_TO_LAC
        if fa_val < 0:
            fa_val = gross_block[i] * CR_TO_LAC
        pay_val = payables[i] * CR_TO_LAC
        rec_val = receivables[i] * CR_TO_LAC
        cash_val = cash_vals[i] * CR_TO_LAC
        inv_val = inventory_vals[i] * CR_TO_LAC
        borrow_val = borrowing[i] * CR_TO_LAC

        financials_out["revenue"].append(rev)
        financials_out["cogs"].append(cogs_val)
        financials_out["opex"].append(opex_val)
        financials_out["ebitda"].append(ebitda_val)
        financials_out["depreciation"].append(dep_val)
        financials_out["ebit"].append(ebit_val)
        financials_out["interest"].append(int_val)
        financials_out["interest_income"].append(int_val)
        financials_out["tax"].append(tax_val)
        financials_out["nopat"].append(nopat_val)
        financials_out["fixed_assets"].append(fa_val)
        financials_out["inventory"].append(inv_val)
        financials_out["receivables"].append(rec_val)
        financials_out["payables"].append(pay_val)
        financials_out["cash"].append(cash_val)
        financials_out["equity"].append(eq_val)
        financials_out["st_debt"].append(borrow_val * 0.40)
        financials_out["lt_debt"].append(borrow_val * 0.60)

    return financials_out, shares, company_name, shares_source, None


def fetch_screener_bulk(symbol: str, num_years: int, session: requests.Session):
    """Try consolidated, then standalone, based on actually-parsed data (not
    HTTP status). Returns dict with keys: financials, shares, company_name,
    data_source, price, price_source, log (list[str]), error (str|None)."""
    log = []
    for label, url in [
        ("Consolidated", f"https://www.screener.in/company/{symbol}/consolidated/"),
        ("Standalone", f"https://www.screener.in/company/{symbol}/"),
    ]:
        time.sleep(random.uniform(1.2, 2.4))  # be respectful to Screener
        soup, fetch_err = _get_soup(url, session)
        if soup is None:
            log.append(f"{label}: fetch failed — {fetch_err}")
            continue
        financials, shares, company_name, shares_source, parse_err = _parse_screener_page(soup, num_years)
        if financials is None:
            log.append(f"{label}: {parse_err} — trying next source" if label == "Consolidated" else f"{label}: {parse_err}")
            continue
        log.append(f"{label}: parsed OK ({len(financials['years'])} yrs); shares via {shares_source or 'unresolved'}")
        price = _extract_current_price_from_soup(soup)
        return {
            "financials": financials, "shares": shares, "company_name": company_name or symbol,
            "data_source": label, "price": price,
            "price_source": "Screener.in page" if price is not None else "",
            "log": log, "error": None,
        }
    return {"financials": None, "shares": 0, "company_name": symbol, "data_source": "",
            "price": None, "price_source": "", "log": log,
            "error": "Neither consolidated nor standalone Screener page had usable financial data"}


# ══════════════════════════════════════════════════════════════════════════
# PRICE FALLBACK CHAIN (never fatal to the row)
# ══════════════════════════════════════════════════════════════════════════

def get_price_with_fallback(screener_price, symbol: str, exchange: str, log: list) -> tuple[Optional[float], str]:
    if screener_price is not None and screener_price > 0:
        return screener_price, "Screener.in page"

    if _HAS_YF_SHIELD:
        yf_symbol = f"{symbol}.BO" if exchange == "BSE" else f"{symbol}.NS"
        try:
            t = safe_ticker(yf_symbol)
            fi = t.fast_info
            px = fi.get("last_price") if isinstance(fi, dict) else getattr(fi, "last_price", None)
            if px:
                return float(px), f"yfinance ({yf_symbol})"
            log.append(f"yfinance fast_info had no last_price for {yf_symbol}")
        except Exception as e:
            log.append(f"yfinance price fetch failed for {yf_symbol}: {e}")

    if _HAS_BSE_QUOTE and exchange == "BSE":
        try:
            data = get_bse_quote(symbol)
            if data and data.get("price"):
                return float(data["price"]), "BSE India (official API)"
            log.append("BSE India API returned no price")
        except Exception as e:
            log.append(f"BSE India API failed: {e}")

    log.append("All price sources exhausted — marking price N/A (DCF result is unaffected)")
    return None, "N/A"


# ══════════════════════════════════════════════════════════════════════════
# MAIN UI ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════

def render_bulk_valuation_ui(
    classify_business_model,
    calculate_working_capital_metrics,
    project_financials,
    calculate_wacc,
    calculate_dcf_valuation,
    ensure_valid_number,
):
    st.subheader("📦 Bulk Valuation — DCF Across Many Companies")
    st.caption(
        "Runs a Free-Cash-Flow-to-Firm DCF for each company using Screener.in data, "
        "one shared beta for the whole batch, and no peer/relative valuation. "
        "Best for scanning a large watchlist quickly, not for a boardroom-grade "
        "single-stock report."
    )

    # ── Ticker selection ─────────────────────────────────────────────────
    st.markdown("### 1️⃣ Choose companies")
    tab_pick, tab_manual = st.tabs(["📋 Pick from NSE/BSE list", "✍️ Manual entry"])

    picked_display = []
    with tab_pick:
        load_err_nse = tdu.get_load_error("NSE")
        load_err_bse = tdu.get_load_error("BSE")
        all_options = tdu.get_peer_display_options()
        if not all_options:
            st.error(
                "Couldn't load the NSE/BSE reference lists. "
                f"NSE error: {load_err_nse or 'n/a'} | BSE error: {load_err_bse or 'n/a'}"
            )
        else:
            exch_choice = st.radio(
                "List", ["Both", "NSE only", "BSE only"], horizontal=True, key="bulk_exch_choice",
            )
            if exch_choice == "NSE only":
                options = [o for o in all_options if "(NSE:" in o]
            elif exch_choice == "BSE only":
                options = [o for o in all_options if "(BSE:" in o]
            else:
                options = all_options
            st.caption(f"{len(options)} companies in this list, sorted alphabetically by name.")

            # NOTE on the pattern below: once a widget has a `key`, Streamlit
            # ignores `default=` on every rerun after the first — the keyed
            # session_state entry is the only source of truth from then on.
            # So "Add range" must write into st.session_state["bulk_multiselect"]
            # directly (before the multiselect widget is instantiated further
            # down in this same script run), not into a separate variable.
            if "bulk_multiselect" not in st.session_state:
                st.session_state.bulk_multiselect = []
            # Prior selections that fell outside the current NSE/BSE filter
            # would otherwise trip Streamlit's "default value not in options"
            # error — drop anything not in the currently visible list.
            st.session_state.bulk_multiselect = [o for o in st.session_state.bulk_multiselect if o in options]

            rc1, rc2, rc3 = st.columns([1, 1, 1.4])
            with rc1:
                range_from = st.number_input(
                    "From #", min_value=1, max_value=max(len(options), 1), value=1, key="bulk_range_from",
                )
            with rc2:
                range_to = st.number_input(
                    "To #", min_value=1, max_value=max(len(options), 1),
                    value=min(20, len(options)) or 1, key="bulk_range_to",
                )
            with rc3:
                st.markdown("&nbsp;")
                if st.button("➕ Add range to selection", key="bulk_add_range"):
                    lo, hi = sorted([int(range_from), int(range_to)])
                    picked_slice = options[lo - 1: hi]
                    st.session_state.bulk_multiselect = list(
                        dict.fromkeys(st.session_state.bulk_multiselect + picked_slice)
                    )
                    st.toast(f"Added {len(picked_slice)} companies (#{lo}–{hi} of the {exch_choice.lower()} list).")

            picked_display = st.multiselect(
                "Search and select companies (type to filter, or use the range picker above)",
                options=options,
                key="bulk_multiselect",
                help="Sourced from NSE_Tickers_List.csv / BSE_Codes_List.csv in the repo.",
            )

    manual_text = ""
    with tab_manual:
        manual_text = st.text_area(
            "One ticker/code per line (or comma-separated). "
            "Numeric entries are treated as BSE codes; anything else as a Screener symbol "
            "(usually the NSE ticker).",
            placeholder="RELIANCE\n500325\nTATASTEEL\nDAVANGERE",
            height=120,
        )

    rows = _dedupe(_resolve_picked_options(picked_display) + _resolve_manual_tickers(manual_text))

    if rows:
        st.info(f"**{len(rows)}** unique companies queued: " + ", ".join(r.symbol for r in rows[:15]) +
                (f" … +{len(rows) - 15} more" if len(rows) > 15 else ""))

    # ── Assumptions ──────────────────────────────────────────────────────
    st.markdown("### 2️⃣ Batch assumptions")
    c1, c2, c3 = st.columns(3)
    with c1:
        beta = st.number_input(
            "Beta (applied to every company)", min_value=0.0, max_value=5.0,
            value=0.96, step=0.01,
            help="No per-company beta calculation in bulk mode — this single value is used for all.",
        )
        num_years = st.slider("Historical years to fetch", 3, 10, 5)
    with c2:
        tax_rate_pct = st.number_input("Tax rate (%)", min_value=0.0, max_value=45.0, value=25.17, step=0.01)
        projection_years = st.slider("Projection years", 3, 10, 5)
    with c3:
        terminal_growth = st.number_input("Terminal growth (%)", min_value=0.0, max_value=10.0, value=4.0, step=0.1)
        default_rf = st.session_state.get("cached_rf_rate_listed", 6.83)
        default_rm = st.session_state.get("cached_rm_rate_listed", 12.0)
        rf_rate = st.number_input("Risk-free rate (%)", value=float(default_rf), step=0.01)
        rm_rate = st.number_input("Market return (%)", value=float(default_rm), step=0.01)

    run = st.button("🚀 Run Bulk Valuation", type="primary", disabled=(len(rows) == 0))

    if not run:
        return

    if "bulk_results" not in st.session_state:
        st.session_state.bulk_results = []

    session = requests.Session()
    results: list[BulkRow] = []
    progress = st.progress(0.0, text="Starting…")

    for idx, row in enumerate(rows):
        with st.status(f"{row.display_name or row.symbol} ({row.symbol})", expanded=False) as status_box:
            try:
                fetch = fetch_screener_bulk(row.symbol, num_years, session)
                for line in fetch["log"]:
                    row.log(line)
                    st.caption(line)

                if fetch["financials"] is None:
                    row.status = "error"
                    row.stage_failed = "screener_fetch"
                    row.error_detail = fetch["error"]
                    status_box.update(label=f"❌ {row.symbol} — Screener fetch failed", state="error")
                    results.append(row)
                    continue

                financials = fetch["financials"]
                row.company_name = fetch["company_name"]
                row.data_source = fetch["data_source"]
                row.shares_outstanding = fetch["shares"]

                price, price_source = get_price_with_fallback(fetch["price"], row.symbol, row.exchange, row.debug_log)
                row.current_price = price
                row.price_source = price_source
                for line in row.debug_log[len(fetch["log"]):]:
                    st.caption(line)

                if not row.shares_outstanding or row.shares_outstanding <= 0:
                    row.status = "error"
                    row.stage_failed = "shares_derivation"
                    row.error_detail = "Could not derive shares outstanding from EPS/Net Profit on the Screener page."
                    status_box.update(label=f"❌ {row.symbol} — shares outstanding unavailable", state="error")
                    results.append(row)
                    continue

                classification = classify_business_model(financials, income_stmt=None, balance_sheet=None)
                if classification.get("type") == "INTEREST_DOMINANT":
                    row.status = "skipped"
                    row.stage_failed = "classification"
                    row.error_detail = (
                        "Classified as a bank/NBFC/interest-dominant business — standard FCFF DCF "
                        "isn't applicable. (" + "; ".join(classification.get("criteria_met", [])) + ")"
                    )
                    status_box.update(label=f"⚠️ {row.symbol} — skipped (bank/NBFC)", state="complete")
                    results.append(row)
                    continue

                wc_metrics = calculate_working_capital_metrics(financials)
                projections, _drivers = project_financials(
                    financials, wc_metrics, projection_years, tax_rate_pct,
                    None, None, None,
                )
                wacc_details = calculate_wacc(
                    financials, tax_rate_pct / 100, peer_tickers=None,
                    manual_rf_rate=rf_rate, manual_rm_rate=rm_rate,
                    manual_beta=beta,
                )
                cash_balance = financials["cash"][0] if financials["cash"][0] > 0 else 0
                valuation, dcf_err = calculate_dcf_valuation(
                    projections, wacc_details, terminal_growth, row.shares_outstanding, cash_balance,
                )

                if dcf_err:
                    row.status = "error"
                    row.stage_failed = "dcf_calculation"
                    row.error_detail = dcf_err
                    status_box.update(label=f"❌ {row.symbol} — DCF calculation error", state="error")
                    results.append(row)
                    continue

                row.status = "success"
                row.fair_value_per_share = valuation.get("fair_value_per_share")
                row.enterprise_value_cr = ensure_valid_number(valuation.get("enterprise_value", 0)) / 10.0
                row.equity_value_cr = ensure_valid_number(valuation.get("equity_value", 0)) / 10.0
                row.wacc_pct = wacc_details.get("wacc")
                if row.current_price and row.fair_value_per_share:
                    row.upside_pct = (row.fair_value_per_share / row.current_price - 1) * 100

                status_box.update(label=f"✅ {row.symbol} — Fair Value ₹{row.fair_value_per_share:.2f}", state="complete")
                results.append(row)

            except Exception:
                row.status = "error"
                row.stage_failed = row.stage_failed or "unexpected"
                row.error_detail = traceback.format_exc(limit=3)
                st.code(row.error_detail)
                status_box.update(label=f"❌ {row.symbol} — unexpected error", state="error")
                results.append(row)

        progress.progress((idx + 1) / len(rows), text=f"{idx + 1}/{len(rows)} processed")

    progress.empty()
    st.session_state.bulk_results = results
    _render_results(results)


# ══════════════════════════════════════════════════════════════════════════
# RESULTS DISPLAY
# ══════════════════════════════════════════════════════════════════════════

def _render_results(results: list[BulkRow]):
    if not results:
        return

    n_ok = sum(1 for r in results if r.status == "success")
    n_skip = sum(1 for r in results if r.status == "skipped")
    n_err = sum(1 for r in results if r.status == "error")

    st.markdown("### 3️⃣ Results")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total", len(results))
    m2.metric("✅ Valued", n_ok)
    m3.metric("⚠️ Skipped", n_skip)
    m4.metric("❌ Errors", n_err)

    df = pd.DataFrame([{
        "Company": r.company_name or r.display_name,
        "Symbol": r.symbol,
        "Status": {"success": "✅", "skipped": "⚠️", "error": "❌"}.get(r.status, r.status),
        "Data Source": r.data_source,
        "Current Price (₹)": r.current_price,
        "Price Source": r.price_source,
        "Fair Value/Share (₹)": r.fair_value_per_share,
        "Upside/Downside %": r.upside_pct,
        "Enterprise Value (₹ Cr)": r.enterprise_value_cr,
        "Equity Value (₹ Cr)": r.equity_value_cr,
        "WACC %": r.wacc_pct,
        "Notes": r.error_detail,
    } for r in results])

    st.dataframe(df, use_container_width=True, hide_index=True)

    csv_bytes = df.to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇️ Download results (CSV)", data=csv_bytes,
        file_name=f"bulk_valuation_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
        mime="text/csv",
    )

    ok_rows = [r for r in results if r.status == "success"]
    if not ok_rows:
        return

    st.markdown("#### 📊 Fair Value vs Current Price")
    ok_sorted = sorted(ok_rows, key=lambda r: (r.upside_pct if r.upside_pct is not None else -1e9), reverse=True)
    names = [r.symbol for r in ok_sorted]

    fig1 = go.Figure()
    fig1.add_bar(name="Current Price", x=names, y=[r.current_price for r in ok_sorted],
                 marker_color="#94a3b8",
                 hovertemplate="%{x}<br>Current: ₹%{y:.2f}<extra></extra>")
    fig1.add_bar(name="Fair Value", x=names, y=[r.fair_value_per_share for r in ok_sorted],
                 marker_color="#2563eb",
                 hovertemplate="%{x}<br>Fair Value: ₹%{y:.2f}<extra></extra>")
    fig1.update_layout(barmode="group", template="plotly_white", height=460,
                        legend=dict(orientation="h", yanchor="bottom", y=1.02),
                        margin=dict(l=10, r=10, t=40, b=10))
    st.plotly_chart(fig1, use_container_width=True)

    st.markdown("#### 📈 Upside / Downside %")
    colors = ["#16a34a" if (r.upside_pct or 0) >= 0 else "#dc2626" for r in ok_sorted]
    fig2 = go.Figure(go.Bar(
        x=names, y=[r.upside_pct for r in ok_sorted], marker_color=colors,
        hovertemplate="%{x}<br>%{y:.1f}%<extra></extra>",
    ))
    fig2.add_hline(y=0, line_color="#334155", line_width=1)
    fig2.update_layout(template="plotly_white", height=420, showlegend=False,
                        margin=dict(l=10, r=10, t=20, b=10),
                        yaxis_title="Upside / Downside (%)")
    st.plotly_chart(fig2, use_container_width=True)

    priced = [r for r in ok_rows if r.current_price]
    if priced:
        st.markdown("#### 🎯 Current Price vs Fair Value (each dot = one company)")
        max_val = max(max(r.current_price for r in priced), max(r.fair_value_per_share for r in priced)) * 1.1
        fig3 = go.Figure()
        fig3.add_trace(go.Scatter(
            x=[r.current_price for r in priced], y=[r.fair_value_per_share for r in priced],
            mode="markers+text", text=[r.symbol for r in priced], textposition="top center",
            marker=dict(size=10, color=[("#16a34a" if (r.upside_pct or 0) >= 0 else "#dc2626") for r in priced]),
            hovertemplate="%{text}<br>Current: ₹%{x:.2f}<br>Fair Value: ₹%{y:.2f}<extra></extra>",
            name="Companies",
        ))
        fig3.add_trace(go.Scatter(x=[0, max_val], y=[0, max_val], mode="lines",
                                   line=dict(dash="dash", color="#94a3b8"),
                                   name="Fairly Valued (y = x)"))
        fig3.update_layout(template="plotly_white", height=520,
                            xaxis_title="Current Price (₹)", yaxis_title="Fair Value / Share (₹)",
                            margin=dict(l=10, r=10, t=20, b=10))
        st.plotly_chart(fig3, use_container_width=True)

    err_rows = [r for r in results if r.status in ("error", "skipped")]
    if err_rows:
        with st.expander(f"🔍 Debug detail for {len(err_rows)} non-success rows"):
            for r in err_rows:
                st.markdown(f"**{r.symbol}** — stage: `{r.stage_failed}`")
                st.write(r.error_detail)
                if r.debug_log:
                    st.caption(" → ".join(r.debug_log))
                st.markdown("---")
