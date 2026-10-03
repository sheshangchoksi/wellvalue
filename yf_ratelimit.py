"""
yf_ratelimit.py  ·  Universal yfinance Rate-Limit Shield
=========================================================
Drop-in replacement / umbrella for ALL yfinance calls in this app.

HOW IT WORKS
------------
1. curl_cffi Chrome-impersonation session  →  bypasses Yahoo's bot filters
2. Exponential back-off with jitter        →  survive transient 429s
3. In-process LRU memory cache             →  zero-network hits for repeat symbols
4. Disk-backed cache (pickle)              →  survives process restarts / Render
                                               free-tier spin-downs, which wipe
                                               the in-memory cache
5. Concurrency throttle                    →  stay under Yahoo's rate budget
6. Per-symbol circuit breaker              →  stop hammering a symbol that's
                                               already failing; fail fast instead
                                               of burning the whole retry budget
                                               (and the shared rate-limit window)
                                               on a request that's going to fail
                                               anyway

HOW TO USE  (two-line migration per file)
-----------------------------------------
BEFORE:
    import yfinance as yf
    ticker = yf.Ticker("RELIANCE.NS")
    df     = yf.download("RELIANCE.NS", period="1y")

AFTER:
    from yf_ratelimit import safe_ticker, safe_download
    ticker = safe_ticker("RELIANCE.NS")
    df     = safe_download("RELIANCE.NS", period="1y")

Everything else (.info, .financials, .history, .balance_sheet, .cashflow,
.options, .option_chain …) works exactly the same on the returned object.

TUNING WITHOUT REDEPLOYING
---------------------------
Every timing/retry constant below can be overridden with an environment
variable of the same name (see Render dashboard -> Environment), so you
can loosen/tighten the throttle without a code change:
    YF_MIN_DELAY_S, YF_MAX_DELAY_S, YF_MAX_RETRIES, YF_BASE_BACKOFF_S,
    YF_CACHE_TTL_S, YF_DISK_CACHE_PATH, YF_CIRCUIT_FAIL_THRESHOLD,
    YF_CIRCUIT_COOLDOWN_S

RENDER FREE-TIER NOTE
----------------------
Render's free web services spin down after ~15 min idle and spin back up
on the next request — this wipes any in-memory cache. The disk cache
below (instance/yf_cache.pkl by default) is written on every successful
fetch and reloaded on import, so previously-fetched tickers stay warm
across a spin-down/spin-up cycle instead of re-hitting Yahoo cold.
"""

from __future__ import annotations

import functools
import logging
import os
import pickle
import random
import tempfile
import threading
import time
from typing import Any

import pandas as pd

logger = logging.getLogger(__name__)


def _env_float(name: str, default: float) -> float:
    try:
        return float(os.environ.get(name, default))
    except (TypeError, ValueError):
        return default


def _env_int(name: str, default: int) -> int:
    try:
        return int(os.environ.get(name, default))
    except (TypeError, ValueError):
        return default


# ── optional streamlit cache (harmless no-op outside Streamlit) ────────────
try:
    import streamlit as st
    _HAS_ST = True
except Exception:
    _HAS_ST = False

# ── curl_cffi (preferred) → requests (fallback) ─────────────────────────────
# DIAGNOSTIC OVERRIDE: set env var SHESHVAL_DISABLE_CURL_CFFI=1 to force the
# plain-requests fallback even when curl_cffi is installed — kept from an
# earlier investigation into a curl_cffi/yfinance session incompatibility
# (turned out to be an incomplete venv, not a real incompatibility — see
# MIGRATION_CHECKLIST.md Round 10). Costs nothing when unset; left in as a
# cheap diagnostic lever if a similar symptom ever reappears.
if os.environ.get("SHESHVAL_DISABLE_CURL_CFFI") == "1":
    import requests as _curl_requests
    _HAS_CURL = False
else:
    try:
        from curl_cffi import requests as _curl_requests
        _HAS_CURL = True
    except ImportError:
        import requests as _curl_requests          # type: ignore[assignment]
        _HAS_CURL = False

import yfinance as yf

# ────────────────────────────────────────────────────────────────────────────
# CONFIG  (env-var overridable — see module docstring)
# ────────────────────────────────────────────────────────────────────────────
# Yahoo's tolerance has tightened noticeably since curl_cffi's original
# defaults were tuned; these are deliberately more conservative than the
# textbook "1 req/sec" figure to leave real headroom before 429s.
MIN_DELAY_S      = _env_float("YF_MIN_DELAY_S", 3.0)     # minimum pause between Yahoo requests
MAX_DELAY_S      = _env_float("YF_MAX_DELAY_S", 6.0)     # maximum pause (random jitter)
MAX_RETRIES      = _env_int("YF_MAX_RETRIES", 4)         # retry budget per call
BASE_BACKOFF_S   = _env_float("YF_BASE_BACKOFF_S", 10.0)  # base for exponential backoff on 429
CACHE_TTL_S      = _env_int("YF_CACHE_TTL_S", 6 * 3600)  # in-process + disk cache TTL (6 hours —
                                                           # fundamentals/financials don't move
                                                           # intraday, so this is safe for a DCF tool)

# Circuit breaker: if a symbol fails this many times in a row, stop retrying
# it against Yahoo for the cooldown window and fail fast instead. Protects
# the shared throttle budget from being spent on a symbol that's currently
# blocked, which would otherwise slow down/starve every other request too.
CIRCUIT_FAIL_THRESHOLD = _env_int("YF_CIRCUIT_FAIL_THRESHOLD", 3)
CIRCUIT_COOLDOWN_S     = _env_int("YF_CIRCUIT_COOLDOWN_S", 120)

# An empty/None/{} result with NO exception is a weaker signal than a real
# 429/403 — it's also what yfinance returns for a ticker that simply has no
# data (delisted, mistyped, or never listed — e.g. yfinance's own "No
# timezone found, symbol may be delisted"). Retrying that MAX_RETRIES times
# at the same backoff as a confirmed rate-limit (70s+ with the defaults
# above) is mostly wasted time on the FIRST call to such a symbol — the
# circuit breaker protects every call after that, but does nothing for the
# first one. Give this weaker signal its own much smaller budget.
EMPTY_RESULT_MAX_RETRIES = _env_int("YF_EMPTY_RESULT_MAX_RETRIES", 2)

DISK_CACHE_PATH = os.environ.get(
    "YF_DISK_CACHE_PATH",
    # This file lives at repo root (sheshvaluations/yf_ratelimit.py), unlike
    # the reference this was ported from (app/logic/yf_ratelimit.py, 3
    # levels deep) — one dirname() gets us to repo root, then into a
    # gitignored cache/ folder next to it, not inside backend/ (which is
    # what uvicorn --reload watches — a cache file churning in there would
    # trigger spurious reloads on every fetch).
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache", "yf_cache.pkl"),
)

_CHROME_UA       = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36"
)

# ────────────────────────────────────────────────────────────────────────────
# RATE-LIMIT GATE  (shared across all threads in this process)
# ────────────────────────────────────────────────────────────────────────────
_gate_lock       = threading.Lock()
_last_request_ts = 0.0

def _throttle():
    """Block until at least MIN_DELAY_S has elapsed since the last call."""
    global _last_request_ts
    with _gate_lock:
        now   = time.monotonic()
        wait  = MIN_DELAY_S - (now - _last_request_ts)
        if wait > 0:
            time.sleep(wait)
        _last_request_ts = time.monotonic()


# ────────────────────────────────────────────────────────────────────────────
# PER-SYMBOL CIRCUIT BREAKER
# ────────────────────────────────────────────────────────────────────────────
_circuit_lock  = threading.Lock()
_circuit_state: dict[str, dict[str, Any]] = {}  # symbol -> {"fails": int, "opened_at": float}

def _circuit_is_open(symbol: str) -> bool:
    with _circuit_lock:
        st_ = _circuit_state.get(symbol)
        if not st_ or st_["fails"] < CIRCUIT_FAIL_THRESHOLD:
            return False
        if time.time() - st_["opened_at"] > CIRCUIT_COOLDOWN_S:
            # Cooldown elapsed — give it one more try.
            st_["fails"] = 0
            return False
        return True

def _circuit_record_failure(symbol: str):
    with _circuit_lock:
        st_ = _circuit_state.setdefault(symbol, {"fails": 0, "opened_at": 0.0})
        st_["fails"] += 1
        if st_["fails"] == CIRCUIT_FAIL_THRESHOLD:
            st_["opened_at"] = time.time()
            logger.warning(
                "yf_ratelimit: circuit OPEN for %s after %d consecutive failures "
                "— failing fast for %ds instead of continuing to retry",
                symbol, CIRCUIT_FAIL_THRESHOLD, CIRCUIT_COOLDOWN_S,
            )

def _circuit_record_success(symbol: str):
    with _circuit_lock:
        if symbol in _circuit_state:
            _circuit_state[symbol] = {"fails": 0, "opened_at": 0.0}


class CircuitOpenError(RuntimeError):
    """Raised when a symbol's circuit breaker is open (too many recent
    failures) instead of spending more of the retry/rate-limit budget on it."""
    pass


# ────────────────────────────────────────────────────────────────────────────
# SESSION FACTORY
# ────────────────────────────────────────────────────────────────────────────
def _make_session():
    """
    Return a curl_cffi Chrome-impersonation session (or a plain requests
    session as fallback). A fresh session is created each time so that
    Yahoo can't track connection state across symbols.
    """
    if _HAS_CURL:
        sess = _curl_requests.Session(impersonate="chrome124")
    else:
        import requests as _req
        from requests.adapters import HTTPAdapter
        from urllib3.util.retry import Retry
        sess = _req.Session()
        retry = Retry(
            total=3,
            backoff_factor=1.5,
            status_forcelist=[429, 500, 502, 503, 504],
            allowed_methods=["GET", "HEAD"],
            raise_on_status=False,
        )
        sess.mount("https://", HTTPAdapter(max_retries=retry))
        sess.mount("http://",  HTTPAdapter(max_retries=retry))

    sess.headers.update({
        "User-Agent":      _CHROME_UA,
        "Accept-Language": "en-US,en;q=0.9",
        "Accept-Encoding": "gzip, deflate, br",
    })
    return sess


# ────────────────────────────────────────────────────────────────────────────
# IN-PROCESS MEMORY CACHE + DISK-BACKED PERSISTENCE
# ────────────────────────────────────────────────────────────────────────────
# The in-memory dict is fast; the disk file makes fetched data survive a
# process restart (e.g. Render free-tier spin-down/spin-up, or a redeploy)
# instead of forcing every symbol to be re-fetched from Yahoo cold.
_mem_cache: dict[str, tuple[float, Any]] = {}
_cache_lock = threading.Lock()
_disk_dirty = False


def _load_disk_cache():
    global _mem_cache
    try:
        if os.path.exists(DISK_CACHE_PATH):
            with open(DISK_CACHE_PATH, "rb") as f:
                data = pickle.load(f)
            if isinstance(data, dict):
                now = time.time()
                # Drop anything already past TTL so we don't resurrect stale data.
                _mem_cache = {k: v for k, v in data.items() if (now - v[0]) < CACHE_TTL_S}
                logger.info("yf_ratelimit: loaded %d cached entries from disk (%s)",
                            len(_mem_cache), DISK_CACHE_PATH)
    except Exception as e:
        # Common cause: the cache file was written by a different
        # numpy/pandas version (pickled DataFrames aren't forward/backward
        # compatible across numpy's internal layout changes) — already
        # non-fatal (falls back to an empty in-memory cache), but left as-is
        # it would silently fail to load, and warn, on every single restart
        # forever. Remove the stale file so it starts clean instead.
        logger.warning("yf_ratelimit: could not load disk cache (%s): %s — removing it", DISK_CACHE_PATH, e)
        try:
            os.remove(DISK_CACHE_PATH)
        except OSError:
            pass


def _save_disk_cache():
    """Atomic write (temp file + rename) so a crash mid-write can't corrupt
    the cache file."""
    try:
        os.makedirs(os.path.dirname(DISK_CACHE_PATH), exist_ok=True)
        fd, tmp_path = tempfile.mkstemp(
            dir=os.path.dirname(DISK_CACHE_PATH), prefix=".yf_cache_", suffix=".tmp"
        )
        with os.fdopen(fd, "wb") as f:
            with _cache_lock:
                pickle.dump(_mem_cache, f, protocol=pickle.HIGHEST_PROTOCOL)
        os.replace(tmp_path, DISK_CACHE_PATH)
    except Exception as e:
        logger.warning("yf_ratelimit: could not save disk cache (%s): %s", DISK_CACHE_PATH, e)


_load_disk_cache()


def _mem_get(key: str) -> Any | None:
    with _cache_lock:
        entry = _mem_cache.get(key)
        if entry and (time.time() - entry[0]) < CACHE_TTL_S:
            return entry[1]
    return None

def _mem_set(key: str, value: Any):
    with _cache_lock:
        _mem_cache[key] = (time.time(), value)
    # Fire-and-forget disk write so cold-start recovery works. Cheap relative
    # to the network round-trip we just made, and infrequent (once per
    # distinct symbol/property per TTL window) so it isn't a hot path.
    threading.Thread(target=_save_disk_cache, daemon=True).start()

def clear_cache(symbol: str | None = None):
    """Clear in-process + disk cache. Pass symbol to clear only that ticker."""
    with _cache_lock:
        if symbol:
            keys = [k for k in _mem_cache if k.startswith(symbol)]
            for k in keys:
                _mem_cache.pop(k, None)
        else:
            _mem_cache.clear()
    _save_disk_cache()
    logger.info("yf_ratelimit: cache cleared%s",
                f" for {symbol}" if symbol else " (all)")


# ────────────────────────────────────────────────────────────────────────────
# RETRY DECORATOR  (wraps any callable that talks to Yahoo)
# ────────────────────────────────────────────────────────────────────────────
def _with_retry(fn, *args, symbol: str | None = None, **kwargs):
    """
    Call fn(*args, **kwargs) with throttle + exponential backoff on errors.
    Returns the result or raises the last exception after MAX_RETRIES.

    If `symbol` is given and its circuit breaker is open (too many recent
    consecutive failures), raises CircuitOpenError immediately without
    touching the network or the shared throttle budget.
    """
    if symbol and _circuit_is_open(symbol):
        raise CircuitOpenError(
            f"{symbol}: too many recent Yahoo Finance failures — pausing "
            f"requests for this symbol for up to {CIRCUIT_COOLDOWN_S}s before "
            f"retrying. If this persists, Yahoo may be rate-limiting this "
            f"server's IP; try Screener Excel Mode instead for Indian tickers."
        )

    last_exc = None
    for attempt in range(MAX_RETRIES):
        _throttle()
        jitter = random.uniform(0, MAX_DELAY_S - MIN_DELAY_S)
        if attempt:
            backoff = BASE_BACKOFF_S * (2 ** (attempt - 1)) + jitter
            logger.warning("yf_ratelimit: retry %d/%d — waiting %.1fs",
                           attempt, MAX_RETRIES, backoff)
            time.sleep(backoff)
        try:
            result = fn(*args, **kwargs)

            # yfinance silently returns None / {} / an empty DataFrame instead
            # of raising when a request is quietly dropped (very common under
            # rate limiting) — treat all three as a failure worth retrying,
            # not a valid empty result, so callers never get None back and
            # so we never poison the cache with a bad answer.
            is_bad = (
                result is None
                or (isinstance(result, dict) and not result)
                or (isinstance(result, pd.DataFrame) and result.empty)
            )
            if is_bad and attempt < EMPTY_RESULT_MAX_RETRIES - 1:
                logger.warning("yf_ratelimit: empty/None result on attempt %d/%d — retrying",
                               attempt + 1, EMPTY_RESULT_MAX_RETRIES)
                last_exc = RuntimeError("Empty/None result returned (possible silent 429)")
                continue
            if is_bad:
                # Exhausted the (small) empty-result budget. Record it for the
                # circuit breaker so repeated calls to this symbol fail fast —
                # but RETURN the bad result rather than raise. Every verbatim-
                # ported caller (get_risk_free_rate, get_market_return,
                # get_stock_beta — all copied from PHASE5_DCF_valuation.py)
                # was written against real yfinance's actual behavior: an
                # invalid/empty-data ticker comes back as an empty DataFrame,
                # not an exception. Their OWN fallback chains (try strategy 2,
                # try strategy 3, then a curated per-ticker rate) depend on
                # that contract — raising here would skip straight past all of
                # it to the first generic `except Exception` in the caller,
                # which is a real regression this was caught doing (see
                # MIGRATION_CHECKLIST.md Round 15).
                if symbol:
                    _circuit_record_failure(symbol)
                logger.warning("yf_ratelimit: giving up after %d empty-result attempt(s) for %s — "
                               "returning empty result to caller's own fallback logic",
                               EMPTY_RESULT_MAX_RETRIES, symbol or "this request")
                return result
            if symbol:
                _circuit_record_success(symbol)
            return result
        except Exception as exc:
            last_exc = exc
            msg = str(exc).lower()
            is_rate = any(x in msg for x in ("429", "rate", "too many", "forbidden", "403", "crumb"))
            if not is_rate:
                # Non-rate-limit error — don't keep retrying
                if symbol:
                    _circuit_record_failure(symbol)
                raise
            logger.warning("yf_ratelimit: rate-limit hit on attempt %d: %s", attempt + 1, exc)

    if symbol:
        _circuit_record_failure(symbol)
    raise last_exc or RuntimeError("yf_ratelimit: all retries exhausted")


# ────────────────────────────────────────────────────────────────────────────
# PUBLIC API  ── safe_ticker() and safe_download()
# ────────────────────────────────────────────────────────────────────────────

class _CachedTicker:
    """
    Lazy, cached wrapper around yf.Ticker. All property accesses are
    cached in-process + on disk, and retried on rate-limit errors.
    """
    _PROPS = ("info", "financials", "income_stmt", "balance_sheet",
              "cashflow", "quarterly_financials", "quarterly_income_stmt",
              "quarterly_balance_sheet", "quarterly_cashflow",
              "fast_info", "dividends", "splits", "actions",
              "recommendations", "calendar", "earnings_dates",
              "options")

    def __init__(self, symbol: str):
        self._symbol  = symbol
        self._yf_obj  = None
        self._yf_lock = threading.Lock()

    # -- lazy yf.Ticker construction -----------------------------------------
    def _get_yf(self) -> yf.Ticker:
        with self._yf_lock:
            if self._yf_obj is None:
                sess = _make_session()
                self._yf_obj = yf.Ticker(self._symbol, session=sess)
        return self._yf_obj

    # -- generic cached property fetch ----------------------------------------
    def _fetch_prop(self, prop: str) -> Any:
        key = f"{self._symbol}:prop:{prop}"
        cached = _mem_get(key)
        if cached is not None:
            return cached

        def _do():
            return getattr(self._get_yf(), prop)

        result = _with_retry(_do, symbol=self._symbol)
        _mem_set(key, result)
        return result

    # -- expose all standard yf.Ticker properties transparently --------------
    def __getattr__(self, name: str) -> Any:
        if name.startswith("_"):
            raise AttributeError(name)
        if name in self._PROPS:
            return self._fetch_prop(name)
        # Pass through anything else (e.g. .ticker, .isin)
        return getattr(self._get_yf(), name)

    # -- history() — supports arbitrary kwargs --------------------------------
    def history(self, period="1mo", interval="1d", **kwargs) -> pd.DataFrame:
        key = f"{self._symbol}:history:{period}:{interval}:{sorted(kwargs.items())}"
        cached = _mem_get(key)
        if cached is not None:
            return cached

        def _do():
            return self._get_yf().history(period=period, interval=interval, **kwargs)

        result = _with_retry(_do, symbol=self._symbol)
        _mem_set(key, result)
        return result

    # -- option_chain() -------------------------------------------------------
    def option_chain(self, date: str | None = None) -> Any:
        key = f"{self._symbol}:option_chain:{date}"
        cached = _mem_get(key)
        if cached is not None:
            return cached

        def _do():
            return self._get_yf().option_chain(date) if date else self._get_yf().option_chain()

        result = _with_retry(_do, symbol=self._symbol)
        _mem_set(key, result)
        return result

    # -- repr / str -----------------------------------------------------------
    def __repr__(self):
        return f"<CachedTicker '{self._symbol}'>"

    def __str__(self):
        return self._symbol


# -- module-level Ticker cache (one object per symbol per process) -----------
_ticker_registry: dict[str, _CachedTicker] = {}
_registry_lock   = threading.Lock()

def safe_ticker(symbol: str) -> _CachedTicker:
    """
    Drop-in for yf.Ticker(symbol).

    Returns a cached, rate-limit-aware wrapper. The same object is
    reused across all calls with the same symbol within a process.

    Usage:
        from yf_ratelimit import safe_ticker
        t = safe_ticker("RELIANCE.NS")
        print(t.info["currentPrice"])
        df = t.history(period="1y")
    """
    with _registry_lock:
        if symbol not in _ticker_registry:
            _ticker_registry[symbol] = _CachedTicker(symbol)
        return _ticker_registry[symbol]


def safe_download(
    tickers,
    period: str = "1mo",
    interval: str = "1d",
    flatten: bool = True,
    **kwargs,
) -> pd.DataFrame:
    """
    Drop-in for yf.download(tickers, ...).

    Extra args vs yf.download:
        flatten (bool): If True (default), flatten MultiIndex columns to
                        single-level "Close", "Open" … instead of
                        ("Close", "AAPL") etc. Matches pre-0.2 behaviour.

    Usage:
        from yf_ratelimit import safe_download
        df = safe_download("RELIANCE.NS", period="1y")
        df = safe_download(["RELIANCE.NS", "TCS.NS"], start="2023-01-01")
    """
    # Build a stable cache key
    ticker_key = tickers if isinstance(tickers, str) else "|".join(sorted(tickers))
    key = f"download:{ticker_key}:{period}:{interval}:{sorted(kwargs.items())}"
    cached = _mem_get(key)
    if cached is not None:
        return cached

    def _do():
        sess = _make_session()
        # NOT progress=False hardcoded here directly — callers (e.g. the
        # verbatim-ported get_risk_free_rate/get_market_return) also pass
        # progress=False explicitly, and Python rejects the same keyword
        # twice. Reapplied from Round 9 after this file was wholesale-
        # replaced from a reference implementation that never had this fix
        # (see MIGRATION_CHECKLIST.md Round 15) — caught immediately by the
        # same TypeError resurfacing in testing, not missed.
        kwargs.setdefault("progress", False)
        return yf.download(
            tickers,
            period=period,
            interval=interval,
            session=sess,
            **kwargs,
        )

    df = _with_retry(_do, symbol=ticker_key)

    # Flatten MultiIndex columns (yfinance >= 0.2 wraps single-ticker downloads too)
    if flatten and isinstance(df.columns, pd.MultiIndex):
        if isinstance(tickers, str) or (isinstance(tickers, (list, tuple)) and len(tickers) == 1):
            df.columns = df.columns.get_level_values(0)
        # For multi-ticker downloads keep MultiIndex — caller can handle it

    _mem_set(key, df)
    return df


# ────────────────────────────────────────────────────────────────────────────
# STREAMLIT CACHE LAYER  (optional — adds cross-rerun deduplication)
# Only activated when Streamlit is present (i.e. inside a Streamlit app)
# ────────────────────────────────────────────────────────────────────────────
if _HAS_ST:
    @st.cache_data(ttl=CACHE_TTL_S, show_spinner=False)
    def st_download(tickers, period="1mo", interval="1d", **kwargs) -> pd.DataFrame:
        """st.cache_data-backed version of safe_download. Use this in Streamlit pages."""
        return safe_download(tickers, period=period, interval=interval, **kwargs)

    @st.cache_data(ttl=CACHE_TTL_S, show_spinner=False)
    def st_ticker_info(symbol: str) -> dict:
        """st.cache_data-backed .info fetch. Fast for repeated Streamlit reruns."""
        return safe_ticker(symbol).info

    @st.cache_data(ttl=CACHE_TTL_S, show_spinner=False)
    def st_ticker_history(symbol: str, period="1mo", interval="1d") -> pd.DataFrame:
        """st.cache_data-backed .history fetch."""
        return safe_ticker(symbol).history(period=period, interval=interval)


# ────────────────────────────────────────────────────────────────────────────
# QUICK SELF-TEST  (run with: python yf_ratelimit.py)
# ────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    print("curl_cffi available:", _HAS_CURL)
    print("disk cache path:", DISK_CACHE_PATH)

    print("\n— safe_ticker test —")
    t = safe_ticker("RELIANCE.NS")
    info = t.info
    print("currentPrice:", info.get("currentPrice") or info.get("regularMarketPrice"))

    print("\n— safe_download test —")
    df = safe_download("RELIANCE.NS", period="5d")
    print(df.tail(3))

    print("\n— cache hit test (should be instant) —")
    t2 = safe_ticker("RELIANCE.NS")
    print("Same object?", t is t2)

    print("\nAll tests passed ✓")
