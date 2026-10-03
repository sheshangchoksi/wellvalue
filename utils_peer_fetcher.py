"""
Industry Peer Fetcher
=====================
Uses Yahoo Finance v6 recommendationsbysymbol API (JSON, no JS rendering needed).
Supports all exchanges: NSE (.NS), BSE (.BO), LSE (.L), SSE (.SS), HKEX (.HK),
NASDAQ/NYSE (bare tickers), and any other Yahoo Finance suffix.
NO HARDCODED PEERS
"""
import logging
import requests
import re

logger = logging.getLogger(__name__)

_INDIAN_SUFFIXES = {'.NS', '.BO'}


def _detect_suffix(ticker: str) -> str:
    """
    Detect exchange suffix from ticker string.
    Returns suffix with dot (e.g. '.NS', '.L', '.SS') or '' for bare tickers.
    """
    ticker = ticker.strip().upper()
    dot = ticker.rfind('.')
    if dot < 0:
        return ''
    candidate = ticker[dot:]
    if re.match(r'\.[A-Z]{1,5}$', candidate):
        return candidate
    return ''


def get_peers_from_yahoo_comparison(ticker: str, max_peers: int = 20, exclude_self: bool = True):
    """
    Fetches peers via Yahoo Finance v6 recommendationsbysymbol API, and tops
    up with the v1 similarsecurities API if the first endpoint didn't reach
    max_peers distinct peers. Works for all exchanges.

    Returns: List of bare base peer tickers (suffix stripped).
             Caller is responsible for re-appending the correct suffix.
    """
    ticker_upper = ticker.strip().upper()
    suffix = _detect_suffix(ticker_upper)
    ticker_base = ticker_upper[:-len(suffix)] if suffix else ticker_upper
    full_ticker = ticker_upper  # pass exactly as given — GEV, SHEL.L, RELIANCE.NS, 600519.SS

    print(f"[PeerFetcher] Input: '{ticker}' → base='{ticker_base}' suffix='{suffix}' → API ticker='{full_ticker}'")

    peers = []

    # ── Primary: v6 recommendationsbysymbol ──────────────────────────────
    for host in ["query2.finance.yahoo.com", "query1.finance.yahoo.com"]:
        try:
            url = f"https://{host}/v6/finance/recommendationsbysymbol/{full_ticker}"
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                              '(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Accept': 'application/json',
            }
            print(f"[PeerFetcher] GET {url}")
            response = requests.get(url, headers=headers, timeout=15)
            print(f"[PeerFetcher] HTTP {response.status_code}")

            if response.status_code != 200:
                print(f"[PeerFetcher] Non-200 from {host}, trying next...")
                continue

            data = response.json()
            print(f"[PeerFetcher] Raw JSON keys: {list(data.keys())}")
            results = data.get('finance', {}).get('result', [])
            print(f"[PeerFetcher] Result count: {len(results)}")

            if not results:
                print(f"[PeerFetcher] Empty result — full response: {str(data)[:500]}")
                break  # got a 200 but empty — no point trying other host

            for rec in results:
                print(f"[PeerFetcher] rec keys: {list(rec.keys())}")
                for item in rec.get('recommendedSymbols', []):
                    symbol = item.get('symbol', '').strip()
                    if not symbol:
                        continue
                    peer_suffix = _detect_suffix(symbol)
                    peer_base = symbol[:-len(peer_suffix)] if peer_suffix else symbol
                    if not peer_base:
                        continue
                    if exclude_self and peer_base == ticker_base:
                        continue
                    if peer_base not in peers:
                        peers.append(peer_base)
                        print(f"   Found: {peer_base} (raw: {symbol})")
            break  # success — don't try other host

        except Exception as e:
            print(f"[PeerFetcher] Exception on {host}: {e}")
            continue

    # ── Fallback / top-up: v1 similar-securities ──────────────────────────
    # Yahoo's v6 endpoint commonly caps out around 5 recommendations
    # regardless of what we ask for — that's a Yahoo-side limit, not one we
    # impose. We used to only try this second endpoint when v6 returned
    # nothing at all, which meant a v6 result of exactly 5 peers silently
    # became the ceiling even when the caller asked for more (max_peers=20,
    # say). Now we also top up whenever we're still short of max_peers, and
    # merge/dedupe with whatever v6 already found.
    if len(peers) < max_peers:
        print(f"[PeerFetcher] Have {len(peers)}/{max_peers} peers — topping up via v1 similar-securities...")
        try:
            url = f"https://query1.finance.yahoo.com/v1/finance/similarsecurities?symbol={full_ticker}"
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                              '(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Accept': 'application/json',
            }
            print(f"[PeerFetcher] Top-up GET {url}")
            r = requests.get(url, headers=headers, timeout=15)
            print(f"[PeerFetcher] Top-up HTTP {r.status_code}")
            if r.status_code == 200:
                data = r.json()
                print(f"[PeerFetcher] Top-up raw: {str(data)[:500]}")
                for item in data.get('finance', {}).get('result', [{}])[0].get('quotes', []):
                    symbol = item.get('symbol', '').strip()
                    if not symbol:
                        continue
                    peer_suffix = _detect_suffix(symbol)
                    peer_base = symbol[:-len(peer_suffix)] if peer_suffix else symbol
                    if peer_base and peer_base != ticker_base and peer_base not in peers:
                        peers.append(peer_base)
                        print(f"   Top-up found: {peer_base} (raw: {symbol})")
        except Exception as e:
            print(f"[PeerFetcher] Top-up exception: {e}")

    result = peers[:max_peers]

    if result:
        print(f"\n[PeerFetcher] SUCCESS: {len(result)} peers for {full_ticker}: {', '.join(result)}")
    else:
        print(f"\n[PeerFetcher] No peers found for {full_ticker}")

    return result


def get_industry_peers(ticker: str, max_peers: int = 20, exclude_self: bool = True):
    """
    Main entry point. Returns bare peer base tickers (no exchange suffix).
    Caller re-appends the correct exchange suffix.
    """
    ticker_upper = ticker.strip().upper()

    print(f"\n{'='*70}")
    print(f"[PeerFetcher] Finding peers for {ticker_upper}")
    print('='*70)

    peers = get_peers_from_yahoo_comparison(ticker_upper, max_peers, exclude_self)

    if peers:
        print(f"\n{'='*70}")
        print(f"[PeerFetcher] FINAL RESULT: {len(peers)} peers")
        for i, peer in enumerate(peers, 1):
            print(f"   {i:2d}. {peer}")
        print('='*70)
    else:
        print(f"\n{'='*70}")
        print(f"[PeerFetcher] FAILED: No peers found for {ticker_upper}")
        print("="*70)

    return peers


# Alias for compatibility
get_industry_peers_fast = get_industry_peers


if __name__ == "__main__":
    test_tickers = [
        "TATASTEEL.NS",
        "ADVAIT.BO",
        "GEV",
        "AAPL",
        "SHEL.L",
        "600519.SS",
    ]
    for ticker in test_tickers:
        peers = get_industry_peers(ticker, max_peers=10)
        print(f"\n{ticker}: {len(peers)} peers\n")
