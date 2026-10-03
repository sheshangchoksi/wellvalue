"""
ticker_dropdown_utils.py
=========================
Loads the NSE ticker list and BSE code list (expected to sit in the same
directory as this file / repo root) and exposes cached name<->ticker
lookups so the app can offer "select company by name" dropdowns instead of
free-text ticker/code entry, both for the company being valued and for peer
companies used in beta / relative-valuation calculations.

File discovery is deliberately flexible: it will match
"NSE_Tickers_List.csv", "NSE Tickers List.csv", different casing, etc. —
anything in the same folder whose name contains "nse" + "ticker" (or
"bse" + "code") and ends in .csv. If nothing matches, get_load_error()
returns a precise diagnostic (folder scanned + files found) instead of a
silently-empty dropdown.

Public API
----------
get_company_display_options(exchange)      -> list[str]  "Name (TICKER)"
resolve_company_display(display, exchange) -> str  bare ticker/code
get_peer_display_options()                 -> list[str]  "Name (NSE: TICKER)" / "Name (BSE: CODE)"
resolve_peer_display_to_full_ticker(disp)  -> str  e.g. "TCS.NS" / "500325.BO"
bare_ticker_to_peer_display(symbol, exch)  -> str  reverse-map for auto-fetched peers
get_load_error(exchange)                   -> str | None  diagnostic for the UI to display
"""
import os
import glob
import pandas as pd
import streamlit as st

_THIS_DIR = os.path.dirname(os.path.abspath(__file__))

# Last diagnostic message per exchange, so the calling UI can show *why*
# a dropdown came back empty instead of a generic "not found".
_LOAD_ERRORS = {"NSE": None, "BSE": None}


def _find_csv(must_contain_all):
    """Look in _THIS_DIR for a .csv file whose lowercased name contains every
    keyword in must_contain_all (e.g. ['nse', 'ticker']). Exact/underscore/space
    variants are all matched this way. Returns the path, or None if nothing
    matches."""
    try:
        candidates = glob.glob(os.path.join(_THIS_DIR, "*.csv"))
    except Exception:
        candidates = []
    for path in candidates:
        fname = os.path.basename(path).lower()
        if all(kw in fname for kw in must_contain_all):
            return path
    return None


def _list_csvs_for_diagnostic():
    try:
        return [os.path.basename(p) for p in glob.glob(os.path.join(_THIS_DIR, "*.csv"))]
    except Exception:
        return []


@st.cache_data(show_spinner=False)
def _load_nse():
    path = _find_csv(["nse", "ticker"])
    if not path:
        found = _list_csvs_for_diagnostic()
        raise FileNotFoundError(
            f"No NSE ticker CSV found in {_THIS_DIR}. "
            f"Expected a .csv file with 'nse' and 'ticker' in its name. "
            f"CSV files present: {found or 'none'}."
        )
    df = pd.read_csv(path, encoding="utf-8-sig")
    df.columns = [c.strip() for c in df.columns]

    ticker_col = next((c for c in df.columns if "ticker" in c.lower()), None)
    name_col = next((c for c in df.columns if c.lower() == "name" or "name" in c.lower()), None)
    if ticker_col is None or name_col is None:
        raise ValueError(
            f"'{os.path.basename(path)}' doesn't have the expected columns. "
            f"Found columns: {list(df.columns)}. Expected something like 'NSE Ticker' and 'Name'."
        )

    df[ticker_col] = df[ticker_col].astype(str).str.strip()
    df[name_col] = df[name_col].astype(str).str.strip()

    name_to_ticker, ticker_to_name = {}, {}
    for _, row in df.iterrows():
        name, tkr = row[name_col], row[ticker_col]
        if not name or not tkr or name.lower() == "nan" or tkr.lower() == "nan":
            continue
        name_to_ticker[name] = tkr
        ticker_to_name[tkr.upper()] = name
    return name_to_ticker, ticker_to_name


@st.cache_data(show_spinner=False)
def _load_bse():
    path = _find_csv(["bse", "code"])
    if not path:
        found = _list_csvs_for_diagnostic()
        raise FileNotFoundError(
            f"No BSE code CSV found in {_THIS_DIR}. "
            f"Expected a .csv file with 'bse' and 'code' in its name. "
            f"CSV files present: {found or 'none'}."
        )
    df = pd.read_csv(path, encoding="utf-8-sig")
    df.columns = [c.strip() for c in df.columns]

    code_col = next((c for c in df.columns if "code" in c.lower()), None)
    name_col = next((c for c in df.columns if c.lower() == "name" or "name" in c.lower()), None)
    if code_col is None or name_col is None:
        raise ValueError(
            f"'{os.path.basename(path)}' doesn't have the expected columns. "
            f"Found columns: {list(df.columns)}. Expected something like 'BSE Code' and 'Name'."
        )

    df[code_col] = df[code_col].astype(str).str.strip()
    df[name_col] = df[name_col].astype(str).str.strip()

    name_to_code, code_to_name = {}, {}
    for _, row in df.iterrows():
        name, code = row[name_col], row[code_col]
        if not name or not code or name.lower() == "nan" or code.lower() == "nan":
            continue
        name_to_code[name] = code
        code_to_name[code] = name
    return name_to_code, code_to_name


def get_load_error(exchange: str):
    """Returns the diagnostic message from the last failed load for this
    exchange ('NSE' or 'BSE'), or None if the last load succeeded / hasn't
    been attempted yet."""
    return _LOAD_ERRORS.get(exchange)


@st.cache_data(show_spinner=False)
def get_company_display_options(exchange: str):
    """Sorted 'Name (TICKER)' options for the company-to-be-valued dropdown."""
    try:
        if exchange == "NSE":
            name_to_ticker, _ = _load_nse()
            opts = [f"{name} ({tkr})" for name, tkr in name_to_ticker.items()]
        elif exchange == "BSE":
            name_to_code, _ = _load_bse()
            opts = [f"{name} ({code})" for name, code in name_to_code.items()]
        else:
            opts = []
        _LOAD_ERRORS[exchange] = None
        return sorted(opts, key=lambda s: s.lower())
    except Exception as e:
        _LOAD_ERRORS[exchange] = str(e)
        return []


def resolve_company_display(display: str, exchange: str) -> str:
    """'RELIANCE INDUSTRIES LTD (RELIANCE)' -> 'RELIANCE'. Empty in -> empty out."""
    if not display:
        return ""
    if "(" in display and display.rstrip().endswith(")"):
        return display[display.rfind("(") + 1: display.rfind(")")].strip()
    return display.strip()


@st.cache_data(show_spinner=False)
def get_peer_display_options():
    """Sorted combined NSE+BSE options for a single peer-selection dropdown,
    e.g. 'TATA CONSULTANCY SERVICES LTD (NSE: TCS)', 'ABB INDIA LIMITED (BSE: 500002)'."""
    opts = []
    try:
        name_to_ticker, _ = _load_nse()
        opts += [f"{name} (NSE: {tkr})" for name, tkr in name_to_ticker.items()]
        _LOAD_ERRORS["NSE"] = None
    except Exception as e:
        _LOAD_ERRORS["NSE"] = str(e)
    try:
        name_to_code, _ = _load_bse()
        opts += [f"{name} (BSE: {code})" for name, code in name_to_code.items()]
        _LOAD_ERRORS["BSE"] = None
    except Exception as e:
        _LOAD_ERRORS["BSE"] = str(e)
    return sorted(opts, key=lambda s: s.lower())


def resolve_peer_display_to_full_ticker(display: str) -> str:
    """'TCS LTD (NSE: TCS)' -> 'TCS.NS'; 'ABB INDIA LIMITED (BSE: 500002)' -> '500002.BO'."""
    if not display or "(" not in display or ":" not in display:
        return ""
    inner = display[display.rfind("(") + 1: display.rfind(")")].strip()  # "NSE: TCS"
    if ":" not in inner:
        return ""
    exch, raw = inner.split(":", 1)
    exch = exch.strip().upper()
    raw = raw.strip()
    suffix = "NS" if exch == "NSE" else "BO"
    return f"{raw}.{suffix}" if raw else ""


def bare_ticker_to_peer_display(bare_symbol: str, exchange: str) -> str:
    """Reverse-map an auto-fetched bare NSE ticker / BSE code back to a
    'Name (NSE: TICKER)' style display string. Falls back to the raw symbol
    itself (still tagged with exchange) if it isn't found in the reference list,
    so auto-fetch never silently drops a peer."""
    bare_symbol = str(bare_symbol).strip()
    if not bare_symbol:
        return ""
    try:
        if exchange == "NSE":
            _, ticker_to_name = _load_nse()
            name = ticker_to_name.get(bare_symbol.upper())
            return f"{name} (NSE: {bare_symbol.upper()})" if name else f"{bare_symbol.upper()} (NSE: {bare_symbol.upper()})"
        else:
            _, code_to_name = _load_bse()
            name = code_to_name.get(bare_symbol)
            return f"{name} (BSE: {bare_symbol})" if name else f"{bare_symbol} (BSE: {bare_symbol})"
    except Exception:
        # Reference list unavailable — still tag with exchange so the peer isn't lost
        tag = "NSE" if exchange == "NSE" else "BSE"
        return f"{bare_symbol} ({tag}: {bare_symbol})"
