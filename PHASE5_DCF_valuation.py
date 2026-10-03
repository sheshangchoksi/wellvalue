import streamlit as st

st.set_page_config(layout="wide", page_title="DCF Valuation Engine")

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import time
import random
import requests
import logging
import traceback
from bs4 import BeautifulSoup
import re
from io import StringIO
# ── yf_ratelimit shim ──────────────────────────────────────────
# Replaces direct yfinance calls with rate-limit-safe wrappers.
# DO NOT remove this block.
from yf_ratelimit import safe_ticker as _rl_ticker, safe_download as _rl_download

class _YFShim:
    """Makes existing yf.Ticker() / yf.download() calls use safe wrappers."""
    @staticmethod
    def Ticker(symbol, **_):
        return _rl_ticker(symbol)
    @staticmethod
    def download(tickers, **kwargs):
        return _rl_download(tickers, **kwargs)

yf = _YFShim()
# ── end shim ───────────────────────────────────────────────────

import plotly.graph_objects as go
import plotly.express as px
import matplotlib
matplotlib.use('Agg')  # headless backend — no display/browser needed
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap
from plotly.subplots import make_subplots

# Stock Price Comparison Module
try:
    from stock_price_comparison import (
        get_stock_comparison_data_listed,
        get_stock_comparison_data_screener
    )
    STOCK_COMPARISON_AVAILABLE = True
except ImportError as e:
    STOCK_COMPARISON_AVAILABLE = False
    STOCK_COMPARISON_ERROR = str(e)

# Screener Excel Mode Module
try:
    from screener_excel_mode import (
        parse_screener_excel_to_dataframes,
        prepare_uploaded_screener_excel,
        get_value_from_screener_df,
        detect_screener_year_columns,
        extract_screener_financials,
        get_screener_shares_outstanding,
        calculate_screener_ddm_valuation,
        calculate_screener_rim_valuation,
        display_screener_financial_summary,
        display_screener_ddm_results,
        display_screener_rim_results,
        fetch_ticker_data_for_screener
    )
    SCREENER_MODE_AVAILABLE = True
except ImportError as e:
    SCREENER_MODE_AVAILABLE = False
    SCREENER_MODE_ERROR = str(e)

# Screener Auto Download Module
try:
    from screener_auto_download_streamlit import integrate_with_existing_upload_section
    AUTO_DOWNLOAD_AVAILABLE = True
except ImportError as e:
    AUTO_DOWNLOAD_AVAILABLE = False
    AUTO_DOWNLOAD_ERROR = str(e)

# Indian Stock Market APIs (fallback for Yahoo Finance)
SCREENER_IMPORT_ERROR = None
try:
    from utils_indian_apis import get_indian_stock_data, get_nse_quote, get_screener_data, fetch_screener_financials
    INDIAN_APIS_AVAILABLE = True
except ImportError as e:
    INDIAN_APIS_AVAILABLE = False
    SCREENER_IMPORT_ERROR = f"ImportError: {e}"
    # Define robust embedded screener with better parsing
    def fetch_screener_financials(symbol, num_years=5):
        """Robust Screener.in scraper with detailed logging and Streamlit Cloud compatibility"""
        import time as _time
        import random as _random
        try:
            from bs4 import BeautifulSoup
            import requests
            import streamlit as st
            from requests.adapters import HTTPAdapter
            from urllib3.util.retry import Retry
            import os
            
            # Disable proxy that may be blocking screener.in
            os.environ.pop('HTTP_PROXY', None)
            os.environ.pop('HTTPS_PROXY', None)
            os.environ.pop('http_proxy', None)
            os.environ.pop('https_proxy', None)
            
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
                'Accept-Language': 'en-US,en;q=0.5',
                'Connection': 'keep-alive'
            }
            
            # Set up session with retries
            session = requests.Session()
            session.trust_env = False  # Bypass proxy blocking screener.in
            retry_strategy = Retry(
                total=3,
                backoff_factor=1,
                status_forcelist=[429, 500, 502, 503, 504],
                allowed_methods=["HEAD", "GET", "OPTIONS"]
            )
            adapter = HTTPAdapter(max_retries=retry_strategy)
            session.mount("https://", adapter)
            session.mount("http://", adapter)
            
            urls_to_try = [f"https://www.screener.in/company/{symbol}/consolidated/", f"https://www.screener.in/company/{symbol}/"]
            
            soup = None
            connection_error = False
            last_error = None
            
            for url in urls_to_try:
                try:
                    _time.sleep(_random.uniform(1.5, 3.0))
                    st.info(f"🔍 Attempting to fetch from: {url}")
                    
                    # Try with SSL verification first
                    try:
                        resp = session.get(url, headers=headers, timeout=30, verify=True)
                    except requests.exceptions.SSLError:
                        st.warning("⚠️ SSL verification failed, retrying without verification...")
                        resp = session.get(url, headers=headers, timeout=30, verify=False)
                    
                    if resp.status_code == 200:
                        soup = BeautifulSoup(resp.content, 'lxml')
                        st.success(f"✅ Successfully connected to Screener.in")
                        break
                    else:
                        st.warning(f"⚠️ Received status code {resp.status_code}, trying next URL...")
                        
                except requests.exceptions.ConnectionError as e:
                    connection_error = True
                    last_error = str(e)
                    st.error(f"❌ CONNECTION ERROR: Cannot reach www.screener.in")
                    
                    # Check if it's Streamlit Cloud specific issue
                    if "Connection refused" in str(e) or "Errno 111" in str(e):
                        st.error("🔴 **STREAMLIT CLOUD NETWORK RESTRICTION DETECTED**")
                        st.markdown("---")
                        st.markdown("### ✅ Recommended Solutions:")
                        st.markdown("""
                        **Option 1: Use Screener Excel Mode (Recommended)**
                        1. Visit [screener.in/company/{}/consolidated/](https://www.screener.in/company/{}/consolidated/)
                        2. Click the **Export** button to download Excel file
                        3. Return to this app and select **"Screener Excel Mode"**
                        4. Upload the downloaded Excel file
                        
                        **Option 2: Use Yahoo Finance Mode**
                        - For listed companies with NSE/BSE tickers
                        - Select "Listed Company (Yahoo Finance)" mode
                        
                        **Option 3: Deploy Elsewhere**
                        - Deploy on Heroku, Railway, or your own server
                        - These platforms typically have fewer network restrictions
                        
                        **Option 4: Upgrade Streamlit Cloud**
                        - Streamlit Cloud Teams/Enterprise may have better network access
                        """.format(symbol, symbol))
                        st.markdown("---")
                    return None
                    
                except requests.exceptions.Timeout:
                    st.warning(f"⚠️ Timeout accessing {url} (30s), trying next URL...")
                    continue
                    
                except Exception as e:
                    st.warning(f"⚠️ Error accessing {url}: {type(e).__name__}: {str(e)}")
                    last_error = str(e)
                    continue
            
            if soup is None:
                if not connection_error:
                    st.error(f"❌ Could not access Screener.in page for {symbol}")
                    if last_error:
                        st.error(f"Last error: {last_error}")
                    st.info("💡 Try using **Screener Excel Mode** instead - upload a manually downloaded file from screener.in")
                return None
            
            company_name = soup.find('h1').get_text(strip=True) if soup.find('h1') else symbol
            
            # Try to extract current price from page
            current_price = 0
            try:
                # Read the LABELLED 'Current Price' ratio. The first span.number on the page is the
                # Market Cap (in Crores), which this used to return as the "price".
                for _li in soup.select('#top-ratios li'):
                    _nm, _vl = _li.find(class_='name'), _li.find(class_='number')
                    if _nm and _vl and 'current price' in _nm.get_text(strip=True).lower():
                        current_price = float(_vl.get_text(strip=True).replace(',', ''))
                        st.write(f"✓ Found Current Price: ₹{current_price:.2f}")
                        break
            except:
                pass
            
            # More flexible table parsing
            def parse_row(table, keywords, debug_name=""):
                if table is None:
                    return []
                for tr in table.find_all('tr'):
                    cells = tr.find_all(['td', 'th'])
                    if not cells:
                        continue
                    label = cells[0].get_text(strip=True).lower().replace('\xa0', ' ').replace('–', '-').replace('+', '').strip()
                    
                    # Try each keyword
                    for kw in keywords:
                        if kw.lower() in label:
                            values = []
                            for cell in cells[1:]:
                                raw = cell.get_text(strip=True).replace(',', '').replace('\xa0', '')
                                try:
                                    values.append(float(raw))
                                except:
                                    values.append(0.0)
                            if debug_name and values:
                                st.write(f"✓ Found {debug_name}: {len(values)} years")
                            return values
                return []
            
            # Get ALL tables on page
            all_tables = soup.find_all('table')
            st.info(f"📊 Found {len(all_tables)} tables on Screener.in page")
            
            # Find tables by their parent section IDs (more reliable)
            pl_table = None
            bs_table = None
            
            # Method 1: Find by section ID
            pl_section = soup.find('section', {'id': 'profit-loss'})
            if pl_section:
                # Find the main results table (not segment table)
                main_table = pl_section.find('div', {'data-result-table': ''})
                if main_table:
                    pl_table = main_table.find('table')
                    if pl_table:
                        st.write("✓ Found P&L table by section ID and data-result-table")
            
            bs_section = soup.find('section', {'id': 'balance-sheet'})
            if bs_section:
                main_table = bs_section.find('div', {'data-result-table': ''})
                if main_table:
                    bs_table = main_table.find('table')
                    if bs_table:
                        st.write("✓ Found Balance Sheet table by section ID and data-result-table")
            
            # Method 2: Fallback - look for class="data-table"
            if pl_table is None or bs_table is None:
                st.warning("⚠️ Using fallback method - searching by class")
                data_tables = soup.find_all('table', class_='data-table')
                
                for idx, table in enumerate(data_tables):
                    # Check what section this table is in
                    parent_section = table.find_parent('section')
                    if parent_section and parent_section.get('id'):
                        section_id = parent_section.get('id')
                        st.write(f"  Table #{idx+1} is in section: {section_id}")
                        
                        if 'profit' in section_id or 'loss' in section_id:
                            if pl_table is None:
                                pl_table = table
                                st.write(f"✓ Using Table #{idx+1} as P&L")
                        elif 'balance' in section_id or 'sheet' in section_id:
                            if bs_table is None:
                                bs_table = table
                                st.write(f"✓ Using Table #{idx+1} as Balance Sheet")
            
            if pl_table is None or bs_table is None:
                st.error(f"❌ Could not locate financial tables. Found {len(all_tables)} total tables.")
                if pl_table is None:
                    st.error("  Missing: P&L table")
                if bs_table is None:
                    st.error("  Missing: Balance Sheet table")
                return None
            
            st.write("### 📋 Parsing P&L Statement")
            # More flexible parsing - return first match found
            def parse_row_flexible(table, keywords, debug_name=""):
                if table is None:
                    return []
                
                for tr in table.find_all('tr'):
                    cells = tr.find_all(['td', 'th'])
                    if not cells:
                        continue
                    
                    # Get label from first cell
                    first_cell = cells[0]
                    label = first_cell.get_text(strip=True).lower()
                    
                    # Remove special characters
                    label = label.replace('\xa0', ' ').replace('–', '-').replace('+', '').replace('&amp;', '').replace('  ', ' ').strip()
                    
                    # Try each keyword
                    for kw in keywords:
                        if kw.lower() in label:
                            values = []
                            for cell in cells[1:]:
                                raw = cell.get_text(strip=True).replace(',', '').replace('\xa0', '').replace('%', '')
                                try:
                                    val = float(raw)
                                    values.append(val)
                                except:
                                    values.append(0.0)
                            
                            # The annual P&L ends with a 'TTM' column that the Balance Sheet lacks; pad()
                            # aligns from the right, so keeping it shifted every P&L year one year off.
                            _first_tr = table.find('tr')
                            _hdr = [c.get_text(strip=True) for c in _first_tr.find_all(['td', 'th'])] if _first_tr else []
                            if _hdr and _hdr[-1].strip().upper() == 'TTM' and values:
                                values = values[:-1]
                            
                            # Only return if we found actual non-zero values
                            if values and any(v != 0 for v in values):
                                if debug_name:
                                    st.write(f"✓ Found {debug_name}: {len(values)} years - Label: '{label[:60]}'")
                                return values
                
                return []
            
            raw_revenue = parse_row_flexible(pl_table, ['revenue', 'sales'], "Revenue")  # Screener labels it 'Sales' for non-financial cos
            raw_expenses = parse_row_flexible(pl_table, ['expenses'], "Expenses")
            raw_operating_profit = parse_row_flexible(pl_table, ['financing profit', 'operating profit'], "Financing/Operating Profit")
            raw_other_income = parse_row_flexible(pl_table, ['other income'], "Other Income")
            raw_interest = parse_row_flexible(pl_table, ['interest'], "Interest")
            raw_depreciation = parse_row_flexible(pl_table, ['depreciation'], "Depreciation")
            raw_pbt = parse_row_flexible(pl_table, ['profit before tax'], "PBT")
            raw_tax = []
            raw_tax_pct = parse_row_flexible(pl_table, ['tax %'], "Tax %")
            raw_net_profit = parse_row_flexible(pl_table, ['net profit'], "Net Profit")
            raw_eps = parse_row_flexible(pl_table, ['eps in rs'], "EPS")
            
            st.write("### 🏦 Parsing Balance Sheet")
            # Equity & Liabilities
            raw_equity_capital = parse_row_flexible(bs_table, ['equity capital'], "Equity Capital")
            raw_reserves = parse_row_flexible(bs_table, ['reserves'], "Reserves")
            raw_borrowing = parse_row_flexible(bs_table, ['borrowing'], "Borrowing")
            raw_other_liabilities = parse_row_flexible(bs_table, ['other liabilities'], "Other Liabilities")
            raw_trade_payables = parse_row_flexible(bs_table, ['trade payables'], "Trade Payables")
            raw_advance_customers = parse_row_flexible(bs_table, ['advance from customers'], "Advance from Customers")
            
            # Assets - main items
            raw_fixed_assets = parse_row_flexible(bs_table, ['fixed assets'], "Fixed Assets")
            raw_gross_block = parse_row_flexible(bs_table, ['gross block'], "Gross Block")
            raw_accumulated_dep = parse_row_flexible(bs_table, ['accumulated depreciation'], "Accumulated Depreciation")
            raw_cwip = parse_row_flexible(bs_table, ['cwip'], "CWIP")
            raw_investments = parse_row_flexible(bs_table, ['investments'], "Investments")
            
            # Current Assets
            raw_trade_receivables = parse_row_flexible(bs_table, ['trade receivables'], "Trade Receivables")
            raw_cash = parse_row_flexible(bs_table, ['cash equivalents'], "Cash")
            raw_inventory = parse_row_flexible(bs_table, ['inventories'], "Inventory")
            raw_loans_advances = parse_row_flexible(bs_table, ['loans n advances'], "Loans & Advances")
            raw_other_assets = parse_row_flexible(bs_table, ['other assets'], "Other Assets")
            
            # SECOND PASS: If main items are missing, look for them as nested items
            # This catches items that appear AFTER expandable section headers
            if not raw_trade_receivables:
                st.info("  🔍 Trade receivables not found as main item, searching nested items...")
                # Look for rows that come after "Other Assets" or similar
                found_nested = False
                in_other_assets_section = False
                for tr in bs_table.find_all('tr'):
                    cells = tr.find_all(['td', 'th'])
                    if not cells:
                        continue
                    label = cells[0].get_text(strip=True).lower()
                    
                    # Check if we're in "Other Assets" section
                    if 'other asset' in label:
                        in_other_assets_section = True
                        continue
                    
                    # If we're in the section and find trade receivables
                    if in_other_assets_section and ('trade receivable' in label or 'receivable' in label):
                        values = []
                        for cell in cells[1:]:
                            raw = cell.get_text(strip=True).replace(',', '').replace('\xa0', '')
                            try:
                                values.append(float(raw))
                            except:
                                values.append(0.0)
                        if values and any(v != 0 for v in values):
                            raw_trade_receivables = values
                            st.write(f"  ✓ Found nested Trade Receivables: {len(values)} years")
                            found_nested = True
                            break
                    
                    # Stop if we hit next major section
                    if in_other_assets_section and ('total' in label or 'fixed asset' in label or 'investment' in label):
                        break
            
            if not raw_inventory:
                st.info("  🔍 Inventory not found as main item, searching nested items...")
                in_other_assets_section = False
                for tr in bs_table.find_all('tr'):
                    cells = tr.find_all(['td', 'th'])
                    if not cells:
                        continue
                    label = cells[0].get_text(strip=True).lower()
                    
                    if 'other asset' in label:
                        in_other_assets_section = True
                        continue
                    
                    if in_other_assets_section and ('inventor' in label or 'stock' in label):
                        values = []
                        for cell in cells[1:]:
                            raw = cell.get_text(strip=True).replace(',', '').replace('\xa0', '')
                            try:
                                values.append(float(raw))
                            except:
                                values.append(0.0)
                        if values and any(v != 0 for v in values):
                            raw_inventory = values
                            st.write(f"  ✓ Found nested Inventory: {len(values)} years")
                            break
                    
                    if in_other_assets_section and ('total' in label or 'fixed asset' in label):
                        break
            
            raw_total_assets = parse_row_flexible(bs_table, ['total assets', 'total asset'], "Total Assets")
            raw_total_liabilities = parse_row_flexible(bs_table, ['total liabilities', 'total liability'], "Total Liabilities")
            
            # Determine main receivables based on what's available
            st.write("### 🔍 Determining Working Capital Items")
            
            # Use trade receivables if available
            raw_receivables = raw_trade_receivables if raw_trade_receivables and sum(raw_trade_receivables) > 0 else []
            
            # If no trade receivables but has loans & advances (NBFC), use that
            if not raw_receivables and raw_loans_advances and sum(raw_loans_advances) > 0:
                raw_receivables = raw_loans_advances
                st.info("  🏦 Using Loans & Advances as Receivables (NBFC detected)")
            
            # Use trade payables if available
            raw_payables = raw_trade_payables if raw_trade_payables and sum(raw_trade_payables) > 0 else []
            
            # If no trade payables but has advances from customers, use that
            if not raw_payables and raw_advance_customers and sum(raw_advance_customers) > 0:
                raw_payables = raw_advance_customers
                st.info("  📦 Using Advances from Customers as Payables")
            
            # If no payables at all, use other liabilities
            if not raw_payables and raw_other_liabilities and sum(raw_other_liabilities) > 0:
                raw_payables = raw_other_liabilities
                st.info("  📊 Using Other Liabilities as Payables")
            
            # Show what we found
            if raw_receivables:
                st.write(f"  ✓ Receivables: {len(raw_receivables)} years")
            else:
                st.warning("  ⚠️ No Receivables found")
            
            if raw_payables:
                st.write(f"  ✓ Payables: {len(raw_payables)} years")
            else:
                st.warning("  ⚠️ No Payables found")
            
            if raw_inventory:
                st.write(f"  ✓ Inventory: {len(raw_inventory)} years")
            else:
                st.info("  ℹ️ No Inventory (may be service/NBFC company)")
            
            # Check what we got
            items_found = sum([
                1 if raw_revenue else 0,
                1 if raw_net_profit else 0,
                1 if raw_equity_capital else 0,
                1 if raw_reserves else 0,
                1 if raw_borrowing else 0,
                1 if raw_fixed_assets else 0,
                1 if raw_operating_profit or raw_pbt else 0
            ])
            
            st.info(f"📊 Extracted {items_found}/7 key line items successfully")
            
            # Show what's missing
            missing_items = []
            if not raw_revenue:
                missing_items.append("Revenue")
            if not raw_net_profit:
                missing_items.append("Net Profit")
            if not raw_equity_capital:
                missing_items.append("Equity Capital")
            if not raw_reserves:
                missing_items.append("Reserves")
            if not raw_borrowing:
                missing_items.append("Borrowings")
            if not raw_fixed_assets:
                missing_items.append("Fixed Assets")
            if not (raw_operating_profit or raw_pbt):
                missing_items.append("Operating Profit/PBT")
            
            if missing_items:
                st.warning(f"⚠️ Missing items: {', '.join(missing_items)}")
            
            if items_found < 3:
                st.error("❌ Insufficient data extracted from Screener.in. Too few key line items found.")
                st.warning("💡 This can happen if:\n- Company has limited financial history\n- Screener.in changed their HTML structure\n- Company uses non-standard accounting labels")
                return None
            
            def pad(lst, n):
                lst = [v for v in lst if v is not None]
                if len(lst) < n:
                    lst = [0.0] * (n - len(lst)) + lst
                return lst[-n:]
            
            n = num_years
            revenue = pad(raw_revenue, n)
            expenses = pad(raw_expenses, n)
            operating_profit = pad(raw_operating_profit, n)
            other_income = pad(raw_other_income, n)
            interest = pad(raw_interest, n)
            depreciation = pad(raw_depreciation, n)
            pbt = pad(raw_pbt, n)
            tax = pad(raw_tax, n)
            tax_pct = pad(raw_tax_pct, n)
            net_profit = pad(raw_net_profit, n)
            eps = pad(raw_eps, n)
            
            equity_capital = pad(raw_equity_capital, n)
            reserves = pad(raw_reserves, n)
            borrowing = pad(raw_borrowing, n)
            other_liabilities = pad(raw_other_liabilities, n)
            payables = pad(raw_payables, n)
            receivables = pad(raw_receivables, n)
            fixed_assets = pad(raw_fixed_assets, n)
            gross_block = pad(raw_gross_block, n)
            accumulated_dep = pad(raw_accumulated_dep, n)
            cwip = pad(raw_cwip, n)
            cash_vals = pad(raw_cash, n)
            inventory_vals = pad(raw_inventory, n)
            investments = pad(raw_investments, n)
            other_assets = pad(raw_other_assets, n)
            total_assets = pad(raw_total_assets, n)
            
            # Shares: prefer Market Cap / Current Price (printed on the page). Net Profit / EPS
            # overstates shares when there is minority interest (consolidated profit includes it).
            shares = 0
            try:
                _tr = {}
                for _li in soup.select('#top-ratios li'):
                    _n, _v = _li.find(class_='name'), _li.find(class_='number')
                    if _n and _v:
                        _tr[_n.get_text(strip=True).lower()] = float(_v.get_text(strip=True).replace(',', ''))
                _mc, _px = _tr.get('market cap'), _tr.get('current price')
                if _mc and _px and _px > 0:
                    shares = int(_mc * 10_000_000 / _px)
                    st.success(f"✅ Shares: {shares:,} (Market Cap ÷ Price)")
            except Exception:
                shares = 0
            if not shares:
                for i in range(n - 1, -1, -1):
                    if eps[i] != 0 and net_profit[i] != 0:
                        shares = int((net_profit[i] * 10_000_000) / eps[i])
                        st.success(f"✅ Calculated shares: {shares:,} (from Year {i+1} EPS: ₹{eps[i]:.2f})")
                        break
            
            # If EPS method failed, try NSEPy
            if shares == 0:
                st.warning("⚠️ Could not calculate from EPS. Trying NSEPy...")
                try:
                    # Try to get from NSE
                    import requests
                    nse_url = f"https://www.nseindia.com/api/quote-equity?symbol={symbol}"
                    headers = {
                        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
                        'Accept': '*/*',
                        'Accept-Language': 'en-US,en;q=0.9',
                    }
                    # Get session cookies first
                    session = requests.Session()
                    session.get("https://www.nseindia.com", headers=headers, timeout=5)
                    resp = session.get(nse_url, headers=headers, timeout=5)
                    if resp.status_code == 200:
                        data = resp.json()
                        # NSE gives "issuedSize" in number of shares
                        shares = data.get('securityInfo', {}).get('issuedSize', 0)
                        if shares > 0:
                            st.success(f"✅ Fetched shares from NSE: {shares:,}")
                        else:
                            st.info("💡 NSE data available but no issuedSize found")
                    else:
                        st.info(f"💡 NSE API returned status {resp.status_code}")
                except Exception as e:
                    st.info(f"💡 NSE fetch failed: {str(e)[:100]}")
            
            if shares == 0:
                st.warning("⚠️ Could not calculate shares outstanding. Will need manual input.")
            
            CR_TO_LAC = 100.0  # 1 Crore = 100 Lacs (was 10.0 -> every monetary figure 10x too small)
            from datetime import datetime as _dt
            current_year = _dt.now().year
            years_labels = [str(current_year - i) for i in range(n)]
            
            financials_out = {
                'years': years_labels,
                'revenue': [], 'cogs': [], 'opex': [], 'ebitda': [], 'depreciation': [],
                'ebit': [], 'interest': [], 'interest_income': [], 'tax': [], 'nopat': [],
                'fixed_assets': [], 'inventory': [], 'receivables': [], 'payables': [],
                'cash': [], 'equity': [], 'st_debt': [], 'lt_debt': []
            }
            
            for i in range(n - 1, -1, -1):
                rev = revenue[i] * CR_TO_LAC
                dep_val = depreciation[i] * CR_TO_LAC
                int_val = interest[i] * CR_TO_LAC
                other_inc = other_income[i] * CR_TO_LAC
                
                # Use operating profit if available, else derive from PBT
                if operating_profit[i] != 0:
                    ebit_val = operating_profit[i] * CR_TO_LAC
                    ebitda_val = ebit_val + dep_val
                else:
                    pbt_val = pbt[i] * CR_TO_LAC
                    ebitda_val = pbt_val + int_val + dep_val
                    ebit_val = ebitda_val - dep_val
                
                # COGS and OpEx estimation
                cogs_val = rev * 0.55 if rev > 0 else 0.0
                opex_val = max(0, rev - cogs_val - ebitda_val)
                if opex_val < 0 and expenses[i] != 0:
                    total_exp = expenses[i] * CR_TO_LAC
                    cogs_val = total_exp * 0.65
                    opex_val = total_exp * 0.35
                
                # Tax calculation
                tax_val = tax[i] * CR_TO_LAC if tax[i] != 0 else ebit_val * 0.25
                t_rate = min(0.35, tax_val / ebit_val) if ebit_val != 0 else 0.25
                nopat_val = ebit_val * (1 - t_rate)
                
                # Balance sheet
                eq_val = (equity_capital[i] + reserves[i]) * CR_TO_LAC
                fa_val = fixed_assets[i] * CR_TO_LAC
                pay_val = payables[i] * CR_TO_LAC
                rec_val = receivables[i] * CR_TO_LAC
                cash_val = cash_vals[i] * CR_TO_LAC
                inv_val = inventory_vals[i] * CR_TO_LAC
                borrow_val = borrowing[i] * CR_TO_LAC
                st_debt_val = borrow_val * 0.30
                lt_debt_val = borrow_val * 0.70
                
                for key, val in zip(
                    ['revenue', 'cogs', 'opex', 'ebitda', 'depreciation', 'ebit', 'interest', 'interest_income', 'tax', 'nopat', 'fixed_assets', 'inventory', 'receivables', 'payables', 'cash', 'equity', 'st_debt', 'lt_debt'],
                    [rev, cogs_val, opex_val, ebitda_val, dep_val, ebit_val, int_val, other_inc, tax_val, nopat_val, fa_val, inv_val, rec_val, pay_val, cash_val, eq_val, st_debt_val, lt_debt_val]
                ):
                    financials_out[key].append(val)
            
            st.success(f"✅ Successfully parsed financials for {company_name}")
            return {'financials': financials_out, 'shares': shares, 'company_name': company_name, 'current_price': current_price}
            
        except Exception as e:
            st.error(f"❌ Scraper error: {str(e)}")
            import traceback
            st.code(traceback.format_exc())
            return None
except Exception as e:
    INDIAN_APIS_AVAILABLE = False
    SCREENER_IMPORT_ERROR = f"Exception: {e}"
    fetch_screener_financials = None

import os


# ================================
# AGGRESSIVE RATE LIMIT PREVENTION WITH DATA CACHING
# ================================

# Global cache for yfinance DATA (not just ticker objects)
# This caches .info, .financials, .balance_sheet, .cashflow to prevent duplicate API calls
_TICKER_DATA_CACHE = {}
_CACHE_TIMESTAMP = {}
CACHE_DURATION = 300  # 5 minutes cache

class CachedTickerData:
    """
    Wrapper class that caches yfinance ticker data properties
    This prevents rate limits by caching .info, .financials, etc.
    """
    def __init__(self, ticker_symbol, force_refresh=False):
        self.symbol = ticker_symbol
        self._ticker_obj = None
        self._info = None
        self._financials = None
        self._balance_sheet = None
        self._cashflow = None
        self._history = None
        self._loaded = False
        
        # Load from cache or fetch new
        if not force_refresh and ticker_symbol in _TICKER_DATA_CACHE:
            cache_age = time.time() - _CACHE_TIMESTAMP.get(ticker_symbol, 0)
            if cache_age < CACHE_DURATION:
                cached_data = _TICKER_DATA_CACHE[ticker_symbol]
                self._info = cached_data.get('info')
                self._financials = cached_data.get('financials')
                self._balance_sheet = cached_data.get('balance_sheet')
                self._cashflow = cached_data.get('cashflow')
                self._history = cached_data.get('history')
                self._loaded = True
    
    def _ensure_loaded(self):
        """Lazy load ticker data on first access"""
        if not self._loaded:
            try:
                self._ticker_obj = yf.Ticker(self.symbol)
                # Fetch all data at once to minimize API calls
                self._info = self._ticker_obj.info
                self._financials = self._ticker_obj.financials
                self._balance_sheet = self._ticker_obj.balance_sheet
                self._cashflow = self._ticker_obj.cashflow
                self._loaded = True
                
                # Cache the data
                _TICKER_DATA_CACHE[self.symbol] = {
                    'info': self._info,
                    'financials': self._financials,
                    'balance_sheet': self._balance_sheet,
                    'cashflow': self._cashflow,
                    'history': self._history
                }
                _CACHE_TIMESTAMP[self.symbol] = time.time()
            except Exception as e:
                # Return empty data on error
                self._info = {}
                self._financials = pd.DataFrame()
                self._balance_sheet = pd.DataFrame()
                self._cashflow = pd.DataFrame()
                raise e
    
    @property
    def info(self):
        """Get cached info or fetch if needed"""
        if self._info is None:
            self._ensure_loaded()
        return self._info if self._info is not None else {}
    
    @property
    def financials(self):
        """Get cached financials or fetch if needed"""
        if self._financials is None:
            self._ensure_loaded()
        return self._financials if self._financials is not None else pd.DataFrame()
    
    @property
    def balance_sheet(self):
        """Get cached balance sheet or fetch if needed"""
        if self._balance_sheet is None:
            self._ensure_loaded()
        return self._balance_sheet if self._balance_sheet is not None else pd.DataFrame()
    
    @property
    def cashflow(self):
        """Get cached cashflow or fetch if needed"""
        if self._cashflow is None:
            self._ensure_loaded()
        return self._cashflow if self._cashflow is not None else pd.DataFrame()
    
    def history(self, *args, **kwargs):
        """Get historical data - cache by parameters"""
        cache_key = f"{self.symbol}_history_{str(args)}_{str(kwargs)}"
        
        if cache_key in _TICKER_DATA_CACHE:
            cache_age = time.time() - _CACHE_TIMESTAMP.get(cache_key, 0)
            if cache_age < CACHE_DURATION:
                return _TICKER_DATA_CACHE[cache_key]
        
        # Fetch new history
        if self._ticker_obj is None:
            self._ticker_obj = yf.Ticker(self.symbol)
        
        hist_data = self._ticker_obj.history(*args, **kwargs)
        
        # Cache it
        _TICKER_DATA_CACHE[cache_key] = hist_data
        _CACHE_TIMESTAMP[cache_key] = time.time()
        
        return hist_data

def get_cached_ticker(ticker_symbol, force_refresh=False):
    """
    Get cached ticker data wrapper
    This returns a CachedTickerData object that caches .info, .financials, etc.
    to prevent rate limits from repeated API calls
    """
    return CachedTickerData(ticker_symbol, force_refresh)

def clear_ticker_cache():
    """Clear the ticker data cache"""
    global _TICKER_DATA_CACHE, _CACHE_TIMESTAMP
    _TICKER_DATA_CACHE = {}
    _CACHE_TIMESTAMP = {}


# ================================

# ================================

# Helper function for ticker exchange suffix
def get_currency_symbol(info_dict):
    """
    Return the correct currency symbol for a Yahoo Finance info dict.
    Defaults to ₹ for INR (Indian stocks), otherwise uses the ISO code
    wrapped neatly (e.g. '$', '€', 'CNY ', 'USD ').
    """
    _SYMBOLS = {
        'INR': '₹',
        'USD': '$',
        'EUR': '€',
        'GBP': '£',
        'JPY': '¥',
        'CNY': '¥',
        'HKD': 'HK$',
        'SGD': 'S$',
        'AUD': 'A$',
        'CAD': 'C$',
        'CHF': 'CHF ',
        'KRW': '₩',
        'TWD': 'NT$',
        'BRL': 'R$',
    }
    if not info_dict:
        return '₹'
    # yfinance exposes 'currency' (price currency) and 'financialCurrency' (statement currency)
    ccy = info_dict.get('currency') or info_dict.get('financialCurrency') or 'INR'
    return _SYMBOLS.get(ccy.upper(), f'{ccy} ')

def get_ticker_with_exchange(ticker, exchange_suffix):
    """Add exchange suffix to ticker. If suffix is empty or ticker already has one, return as-is."""
    ticker = ticker.strip().upper()
    # Already has a recognised suffix — leave it alone
    known = ('.NS', '.BO', '.L', '.SS', '.HK', '.DE', '.F', '.KS', '.TW', '.SI', '.AX', '.TO', '.V')
    if any(ticker.endswith(s) for s in known):
        return ticker
    if not exchange_suffix:          # NASDAQ/NYSE / Other — bare ticker is correct
        return ticker
    return f"{ticker}.{exchange_suffix}"


def get_price_as_of(stock_obj, as_of_date, max_lookback_days=180):
    """
    Return the closing price of a yfinance Ticker object on (or the most
    recent trading day before) as_of_date — i.e. the T-1, T-2, T-3 ... close,
    whichever is the first one actually traded at or before that date.
    Falls back to None only if no price data is available at all within
    max_lookback_days, so callers can decide how to handle it (e.g. skip
    the peer, or fall back to a live price).

    Args:
        stock_obj: a yfinance Ticker (or CachedTickerData wrapper exposing
                   .history(start=, end=))
        as_of_date: datetime.date / datetime.datetime / pd.Timestamp
        max_lookback_days: widest window (in calendar days) to search
                            backwards for a trading day. The search starts
                            narrow (15 days, covering ordinary weekends/
                            holidays) and widens automatically if the stock
                            didn't trade in that window (e.g. long halts,
                            illiquid names, newly-listed peers).
    """
    import pandas as pd
    from datetime import timedelta
    try:
        as_of_date = pd.Timestamp(as_of_date)
    except Exception:
        return None
    end = as_of_date + timedelta(days=1)

    # Start with a narrow window (handles ordinary weekends/holidays) and
    # widen progressively until we find the last traded close on or before
    # as_of_date, or give up at max_lookback_days. Each attempt is
    # independently guarded — a transient error (rate limit, timeout) on
    # one window must not abort the wider retries that follow.
    last_error = None
    for lookback in (15, 30, 60, max_lookback_days):
        try:
            start = as_of_date - timedelta(days=lookback)
            hist = stock_obj.history(start=start, end=end)
            if hist is None or hist.empty:
                continue
            # yfinance's history() index is timezone-aware (localized to the
            # exchange's own timezone, e.g. Asia/Kolkata for .NS tickers),
            # while as_of_date is a plain naive Timestamp. Comparing the two
            # directly raises TypeError ("Invalid comparison between
            # tz-aware and tz-naive") — which was being silently swallowed
            # by this same try/except and misreported as "no price found"
            # for every single ticker and every lookback window, since the
            # mismatch is identical regardless of window size. Strip the
            # index's tz before comparing so the comparison is well-defined.
            _hist_idx = hist.index
            if getattr(_hist_idx, 'tz', None) is not None:
                _hist_idx = _hist_idx.tz_localize(None)
            hist = hist.set_axis(_hist_idx)
            hist = hist[hist.index <= as_of_date]
            if hist.empty:
                continue
            # Last row on/before as_of_date = the most recent prior trading
            # day (T-0 if it traded that day, else T-1, T-2, ... as needed).
            return float(hist['Close'].iloc[-1])
        except Exception as e:
            last_error = e
            continue

    if last_error is not None:
        logging.getLogger(__name__).warning(
            "get_price_as_of: all lookback windows failed for %s as of %s: %s",
            getattr(stock_obj, 'symbol', stock_obj), as_of_date, last_error
        )
    return None


def is_valuation_date_today(valuation_date):
    """True if valuation_date is None or equals today's date (i.e. 'use live data')."""
    if not valuation_date:
        return True
    from datetime import date
    vd = valuation_date.date() if hasattr(valuation_date, 'date') else valuation_date
    return vd == date.today()

# ================================
# STREAMLIT UI
# ================================

# Initialize session state for rate limiting
if 'last_yahoo_request' not in st.session_state:
    st.session_state.last_yahoo_request = 0

if 'yahoo_request_count' not in st.session_state:
    st.session_state.yahoo_request_count = 0

if 'session_start_time' not in st.session_state:
    import time
    st.session_state.session_start_time = time.time()

# Reset counter every hour
import time
if time.time() - st.session_state.session_start_time > 3600:
    st.session_state.yahoo_request_count = 0
    st.session_state.session_start_time = time.time()

# Fix text truncation in metrics and throughout the app
st.markdown("""
    <style>
    /* Fix metric value truncation */
    [data-testid="stMetricValue"] {
        width: fit-content;
        white-space: nowrap;
        overflow: visible !important;
        text-overflow: clip !important;
    }
    
    /* Fix metric label truncation */
    [data-testid="stMetricLabel"] {
        width: fit-content;
        white-space: nowrap;
        overflow: visible !important;
        text-overflow: clip !important;
    }
    
    /* Fix metric container */
    [data-testid="metric-container"] {
        overflow: visible !important;
        width: fit-content;
        min-width: max-content;
    }
    
    /* Fix metric delta truncation */
    [data-testid="stMetricDelta"] {
        white-space: nowrap;
        overflow: visible !important;
    }
    
    /* General fix for all div elements in columns */
    [data-testid="column"] > div {
        overflow: visible !important;
    }
    
    /* Ensure proper spacing for columns */
    [data-testid="column"] {
        overflow: visible !important;
        min-width: fit-content;
    }
    </style>
    """, unsafe_allow_html=True)


# UTILITY FUNCTIONS
# ================================


# Auto peer fetching utility - MANDATORY
try:
    from utils_peer_fetcher import get_industry_peers
    PEER_FETCHER_AVAILABLE = True
except Exception as e:
    PEER_FETCHER_AVAILABLE = False
    print(f"[DCF] ERROR: Peer fetcher not available: {e}")

# NSE/BSE name-based dropdown utility (company-to-be-valued + peer selection)
try:
    from ticker_dropdown_utils import (
        get_company_display_options,
        resolve_company_display,
        get_peer_display_options,
        resolve_peer_display_to_full_ticker,
        bare_ticker_to_peer_display,
        get_load_error,
    )
    TICKER_DROPDOWN_AVAILABLE = True
except Exception as e:
    TICKER_DROPDOWN_AVAILABLE = False
    print(f"[DCF] ERROR: Ticker dropdown utils not available: {e}")

def sanitize_value(val):
    """Convert string values to float, handling various formats"""
    if pd.isna(val) or val == '' or val == '-':
        return 0.0
    try:
        return float(str(val).replace(',', ''))
    except:
        return 0.0

def safe_extract(data, key, year, default=0.0):
    """
    ROBUST EXTRACTOR: Safely extract value from DataFrame/dict, handling None/NaN/missing data
    
    This is THE solution to the VEDL/TATAMOTORS None-value problem.
    Returns 0.0 instead of None to prevent NaN cascading through calculations.
    
    Args:
        data: DataFrame or dict containing financial data
        key: Row index or dictionary key to extract
        year: Column name (for DataFrame) or nested key
        default: Default value if extraction fails (default: 0.0)
    
    Returns:
        float: Extracted value or default (never None or NaN)
    """
    try:
        if isinstance(data, pd.DataFrame):
            if key in data.index and year in data.columns:
                val = data.loc[key, year]
                # Handle None, NaN, inf
                if val is None or pd.isna(val) or np.isinf(val):
                    return default
                return abs(float(val))
            return default
        elif isinstance(data, dict):
            val = data.get(key, default)
            if val is None or (isinstance(val, float) and (np.isnan(val) or np.isinf(val))):
                return default
            return abs(float(val))
        return default
    except Exception:
        return default

def safe_divide(numerator, denominator, default=0.0):
    """
    ROBUST DIVISION: Safely divide two numbers, handling None/NaN/zero division
    
    Args:
        numerator: Number to divide
        denominator: Number to divide by
        default: Default value if division fails (default: 0.0)
    
    Returns:
        float: Result of division or default (never None or NaN)
    """
    try:
        # Handle None inputs
        if numerator is None or denominator is None:
            return default
        
        # Convert to float and check for NaN/inf
        num = float(numerator)
        den = float(denominator)
        
        if np.isnan(num) or np.isinf(num) or np.isnan(den) or np.isinf(den):
            return default
        
        # Handle zero denominator
        if den == 0:
            return default
        
        result = num / den
        
        # Check result validity
        if np.isnan(result) or np.isinf(result):
            return default
        
        return result
    except Exception:
        return default

def ensure_valid_number(val, default=0.0):
    """
    ROBUST VALIDATOR: Ensure a value is a valid number (not None, NaN, or inf)
    
    Args:
        val: Value to validate
        default: Default value if validation fails (default: 0.0)
    
    Returns:
        float: Valid number or default (never None or NaN)
    """
    try:
        if val is None:
            return default
        
        num = float(val)
        
        if np.isnan(num) or np.isinf(num):
            return default
        
        return num
    except Exception:
        return default

def parse_excel_to_dataframes(excel_file):
    """Parse Excel file and extract Balance Sheet and P&L as DataFrames"""
    try:
        # Read both sheets from Excel
        df_bs = pd.read_excel(excel_file, sheet_name='BalanceSheet', header=None)
        df_pl = pd.read_excel(excel_file, sheet_name='Profit&Loss', header=None)
        
        # Set first column as 'Item' and rest as year columns
        # Column 0 = Item names, Columns 1,2,3... = Year data
        
        # Extract header row (first row contains year numbers)
        bs_years = df_bs.iloc[0, 1:].values  # Get year values from first row
        pl_years = df_pl.iloc[0, 1:].values
        
        # Create column names: 'Item' for first column, '_XX' for year columns
        bs_columns = ['Item'] + [f'_{int(year)}' if pd.notna(year) else f'_col{i}' 
                                  for i, year in enumerate(bs_years, 1)]
        pl_columns = ['Item'] + [f'_{int(year)}' if pd.notna(year) else f'_col{i}' 
                                  for i, year in enumerate(pl_years, 1)]
        
        # Remove header row and set column names
        df_bs = df_bs.iloc[1:].copy()  # Skip header row
        df_bs.columns = bs_columns
        
        df_pl = df_pl.iloc[1:].copy()
        df_pl.columns = pl_columns
        
        # Reset index
        df_bs = df_bs.reset_index(drop=True)
        df_pl = df_pl.reset_index(drop=True)
        
        # Convert year columns to numeric (sanitize values)
        for col in df_bs.columns[1:]:  # Skip 'Item' column
            df_bs[col] = df_bs[col].apply(sanitize_value)
        
        for col in df_pl.columns[1:]:
            df_pl[col] = df_pl[col].apply(sanitize_value)
        
        # Remove rows where Item is NaN or empty
        df_bs = df_bs[df_bs['Item'].notna() & (df_bs['Item'] != '')]
        df_pl = df_pl[df_pl['Item'].notna() & (df_pl['Item'] != '')]
        
        return df_bs, df_pl
        
    except Exception as e:
        st.error(f"Error parsing Excel: {str(e)}")
        return None, None

def get_value_from_df(df, item_name, year_col):
    """Extract value from DataFrame by item name (case-insensitive partial match)"""
    if df is None or df.empty:
        return 0.0
    
    item_name_lower = item_name.lower()
    mask = df['Item'].str.lower().str.contains(item_name_lower, na=False, regex=False)
    matching = df[mask]
    
    if not matching.empty and year_col in matching.columns:
        return matching.iloc[0][year_col]
    return 0.0

def detect_year_columns(df):
    """Detect year columns dynamically (columns starting with _)"""
    if df is None or df.empty:
        return []
    
    year_cols = [col for col in df.columns if col.startswith('_') and col != 'Item']
    # Sort by numeric value after underscore
    year_cols.sort(key=lambda x: int(x[1:]) if x[1:].isdigit() else 0)
    return year_cols

# ================================
# BUSINESS MODEL CLASSIFICATION (RULEBOOK COMPLIANT)
# ================================

def classify_business_model(financials, income_stmt=None, balance_sheet=None):
    """
    Classify company as OPERATING or INTEREST-DOMINANT per Rulebook Section 2.
    
    Returns:
        dict: {
            'type': 'OPERATING' or 'INTEREST_DOMINANT',
            'criteria_met': list of criteria that triggered classification,
            'metrics': dict of calculated ratios
        }
    """
    criteria_met = []
    metrics = {}
    
    # Calculate 3-year averages for classification
    try:
        # Method 1: From extracted financials dict
        if financials and 'revenue' in financials and 'interest' in financials:
            revenues = financials['revenue']
            interest_expenses = financials['interest']
            
            # Get interest income if available (from additional fields)
            interest_income = financials.get('interest_income', [0] * len(revenues))
            
            # Calculate averages
            avg_revenue = np.mean(revenues) if revenues else 0
            avg_interest_expense = np.mean(interest_expenses) if interest_expenses else 0
            avg_interest_income = np.mean(interest_income) if interest_income else 0
            
            # Criterion 1: Interest Income / Total Revenue ≥ 50%
            # Screener-sourced data has no interest-income line (stored as 0.0 = 'not available').
            # In that case criterion 1 can never fire, so use a labelled PROXY instead:
            # Interest EXPENSE / Revenue ≥ 35% (typical of banks/NBFCs, far above normal operating cos).
            if avg_revenue > 0:
                interest_income_available = any(abs(float(x or 0)) > 0 for x in interest_income)
                if interest_income_available:
                    interest_income_ratio = (avg_interest_income / avg_revenue) * 100
                    metrics['interest_income_ratio'] = interest_income_ratio
                    if interest_income_ratio >= 50:
                        criteria_met.append(f"Interest Income / Revenue = {interest_income_ratio:.1f}% (≥50%)")
                else:
                    interest_expense_to_revenue = (avg_interest_expense / avg_revenue) * 100
                    metrics['interest_expense_to_revenue'] = interest_expense_to_revenue
                    metrics['interest_income_available'] = False
                    if interest_expense_to_revenue >= 35:
                        criteria_met.append(
                            f"Interest Expense / Revenue = {interest_expense_to_revenue:.1f}% (≥35%) "
                            f"[proxy: interest income not available from source]")
            
            # Criterion 2: Interest Expense / Total Expenses ≥ 40%
            # Total Expenses = COGS + Opex + Interest + Depreciation
            total_expenses = []
            for i in range(len(revenues)):
                exp = (financials.get('cogs', [0]*len(revenues))[i] + 
                       financials.get('opex', [0]*len(revenues))[i] + 
                       interest_expenses[i] +
                       financials.get('depreciation', [0]*len(revenues))[i])
                total_expenses.append(exp)
            
            avg_total_expenses = np.mean(total_expenses) if total_expenses else 0
            if avg_total_expenses > 0:
                interest_expense_ratio = (avg_interest_expense / avg_total_expenses) * 100
                metrics['interest_expense_ratio'] = interest_expense_ratio
                if interest_expense_ratio >= 40:
                    criteria_met.append(f"Interest Expense / Total Expenses = {interest_expense_ratio:.1f}% (≥40%)")
        
        # Method 2: Check raw statements for Net Interest Income presence
        if income_stmt is not None:
            nii_fields = ['Net Interest Income', 'Net Interest Margin', 'Interest Income Net']
            for field in nii_fields:
                if field in income_stmt.index:
                    criteria_met.append(f"Presence of '{field}' line item")
                    break
        
        # Method 3: Balance sheet structure check
        if balance_sheet is not None:
            # Check for lending business indicators
            lending_indicators = ['Loans', 'Advances', 'Loans And Advances', 'Net Loans']
            financial_assets = ['Financial Assets', 'Investment Securities']
            
            for indicator in lending_indicators + financial_assets:
                if indicator in balance_sheet.index:
                    # Check if it's a significant portion
                    if 'Total Assets' in balance_sheet.index:
                        try:
                            asset_val = abs(balance_sheet.loc[indicator, balance_sheet.columns[0]])
                            total_assets = abs(balance_sheet.loc['Total Assets', balance_sheet.columns[0]])
                            if total_assets > 0 and (asset_val / total_assets) > 0.5:
                                criteria_met.append(f"Balance Sheet dominated by {indicator} ({asset_val/total_assets*100:.1f}% of assets)")
                        except:
                            pass
    
    except Exception as e:
        st.warning(f"Classification warning: {str(e)}")
    
    # Decision: INTEREST-DOMINANT if 2+ criteria met
    is_interest_dominant = len(criteria_met) >= 2
    
    classification = {
        'type': 'INTEREST_DOMINANT' if is_interest_dominant else 'OPERATING',
        'criteria_met': criteria_met,
        'metrics': metrics
    }
    
    return classification

def validate_fcff_eligibility(classification):
    """
    Check if FCFF DCF is valid per Rulebook Section 3.
    Returns: (is_valid: bool, reason: str)
    """
    if classification['type'] == 'INTEREST_DOMINANT':
        return False, "🚫 FCFF DCF is NOT VALID for Interest-Dominant entities. Debt is operating raw material, not financing."
    
    return True, "✅ FCFF DCF is valid for Operating Companies"

def show_classification_warning(classification):
    """Display business model classification and restrictions"""
    if classification['type'] == 'INTEREST_DOMINANT':
        st.error("""
        🚫 **INTEREST-DOMINANT ENTITY DETECTED**
        
        This company derives significant income from interest operations (lending/banking).
        
        **Why FCFF DCF is Invalid:**
        - Interest expense = Operating Cost (like COGS), not financing cost
        - Interest income = Revenue
        - Debt = Operating raw material (inventory equivalent)
        - EBIT/NOPAT/WACC are economically meaningless
        
        **Criteria Met:**
        """)
        for criterion in classification['criteria_met']:
            st.write(f"  • {criterion}")
        
        st.info("""
        **Recommended Valuation Methods:**
        - ✅ Residual Income Model (preferred)
        - ✅ Dividend Discount Model
        - ✅ P/B with ROE analysis
        - ✅ Relative valuation (P/E, P/B)
        
        ❌ FCFF DCF is blocked to prevent economically invalid valuation.
        """)
        
        return True  # Should stop execution
    
    else:
        st.success(f"""
        ✅ **OPERATING COMPANY CLASSIFICATION**
        
        FCFF DCF valuation is appropriate for this company.
        """)
        
        if classification['criteria_met']:
            with st.expander("ℹ️ Classification Details"):
                st.write("The following interest-related metrics were detected but did not exceed thresholds:")
                for criterion in classification['criteria_met']:
                    st.write(f"  • {criterion}")
        
        return False  # Can continue

# ================================
# BANK VALUATION METHODS
# ================================

def calculate_residual_income_model(financials, shares, cost_of_equity, terminal_growth=3.5, projection_years=5, assumed_roe=None, dcf_projections=None):
    """
    Residual Income Model - Suitable for both banks and non-banking companies
    RI = Net Income - (Cost of Equity × Book Value of Equity)
    Value = Book Value + PV(Future Residual Income)
    
    Particularly effective for:
    - Companies with stable book value
    - Banks and financial institutions
    - Asset-heavy businesses
    
    Args:
        terminal_growth: Terminal growth rate (default 3.5%)
        projection_years: Years to project (default 5)
        assumed_roe: Override ROE if provided
        dcf_projections: DCF projections dict with 'nopat' key - if provided, uses projected NOPAT as Net Income (NO DUPLICATION!)
    """
    try:
        # Validation checks with specific error messages
        if 'equity' not in financials or len(financials['equity']) == 0:
            return {
                'error': True,
                'reason': 'No equity (book value) data available',
                'suggestion': 'RIM requires balance sheet equity data. Use DCF or DDM instead.'
            }
        
        # Get latest data - USE NEWEST (index 0, not -1)
        latest_equity = financials['equity'][0] * 100000  # Convert from Lacs to Rupees
        
        if latest_equity <= 0:
            return {
                'error': True,
                'reason': f'Company has negative or zero book value (₹{latest_equity:,.0f})',
                'suggestion': 'RIM requires positive equity. This company may be distressed. Use DCF instead.'
            }
        
        # Calculate average ROE
        net_incomes = []
        equities = []
        for i in range(len(financials['years'])):
            # Approximate net income from NOPAT (banks don't really have NOPAT, using as proxy)
            ni = financials['nopat'][i] * 100000
            eq = financials['equity'][i] * 100000
            net_incomes.append(ni)
            equities.append(eq)
        
        avg_net_income = np.mean(net_incomes)
        
        if avg_net_income <= 0:
            return {
                'error': True,
                'reason': f'Company has negative average net income (₹{avg_net_income:,.0f} Lacs)',
                'suggestion': 'RIM requires profitable companies. This company is loss-making. Use DCF or asset-based valuation.'
            }
        
        # Use provided ROE or calculate
        if assumed_roe:
            roe = assumed_roe
        else:
            roe = (avg_net_income / latest_equity * 100) if latest_equity > 0 else 15
        
        if roe < 0:
            return {
                'error': True,
                'reason': f'Company has negative ROE ({roe:.1f}%)',
                'suggestion': 'RIM requires positive ROE. Company is destroying shareholder value. Use DCF instead.'
            }
        
        # Calculate historical book value growth rate using CAGR (data newest to oldest)
        if len(equities) > 1 and equities[-1] > 0 and equities[0] > 0:
            num_years = len(equities) - 1
            # Start = oldest (last), End = newest (first)
            bv_growth = ((equities[0] / equities[-1]) ** (1 / num_years) - 1) * 100
            bv_growth = max(-50, min(bv_growth, 150))  # Allow up to 150% growth
        else:
            bv_growth = 10.0  # Default
        
        # Project N years of residual income
        # ✅ USE EXISTING DCF PROJECTIONS IF AVAILABLE - NO DUPLICATION!
        projections = []
        current_bv = latest_equity
        
        if dcf_projections and 'nopat' in dcf_projections and len(dcf_projections['nopat']) >= projection_years:
            # Use projected NOPAT from DCF as Net Income (already calculated!)
            for year in range(1, projection_years + 1):
                projected_nopat_lacs = dcf_projections['nopat'][year-1]  # 0-indexed
                current_ni = projected_nopat_lacs * 100000  # Convert to Rupees
                
                # Book value grows with retained earnings
                # Simplified: BV grows at historical rate
                current_bv = current_bv * (1 + bv_growth / 100)
                
                # Residual income = NI - (Ke × BV)
                ri = current_ni - (cost_of_equity / 100 * current_bv)
                
                # Present value
                pv_ri = ri / ((1 + cost_of_equity / 100) ** year)
                projections.append({
                    'year': year,
                    'book_value': current_bv,
                    'net_income': current_ni,
                    'residual_income': ri,
                    'pv_ri': pv_ri,
                    'source': 'DCF Projected NOPAT'
                })
        else:
            # Fallback: use ROE-based projection
            current_ni = avg_net_income
            for year in range(1, projection_years + 1):
                # Growth in book value
                current_bv = current_bv * (1 + bv_growth / 100)
                current_ni = current_bv * (roe / 100)
                
                # Residual income = NI - (Ke × BV)
                ri = current_ni - (cost_of_equity / 100 * current_bv)
                
                # Present value
                pv_ri = ri / ((1 + cost_of_equity / 100) ** year)
                projections.append({
                    'year': year,
                'book_value': current_bv,
                'net_income': current_ni,
                'residual_income': ri,
                'pv_ri': pv_ri,
                'source': 'ROE-based projection'
            })
        
        # Terminal value of residual income
        if cost_of_equity / 100 > terminal_growth / 100:
            terminal_ri = projections[-1]['residual_income'] * (1 + terminal_growth / 100) / (cost_of_equity / 100 - terminal_growth / 100)
            pv_terminal_ri = terminal_ri / ((1 + cost_of_equity / 100) ** projection_years)  # was hard-coded ** 5
        else:
            pv_terminal_ri = 0
        
        # Total value = Current BV + Sum of PV(RI) + PV(Terminal RI)
        sum_pv_ri = sum([p['pv_ri'] for p in projections])
        total_equity_value = latest_equity + sum_pv_ri + pv_terminal_ri
        
        value_per_share = total_equity_value / shares if shares > 0 else 0
        book_value_per_share = latest_equity / shares if shares > 0 else 0
        current_eps = avg_net_income / shares if shares > 0 else 0
        
        return {
            'method': 'Residual Income Model',
            'current_book_value': latest_equity,
            'book_value_per_share': book_value_per_share,
            'current_eps': current_eps,
            'roe': roe,
            'bv_growth': bv_growth,
            'cost_of_equity': cost_of_equity,
            'terminal_growth': terminal_growth,
            'projections': projections,
            'sum_pv_ri': sum_pv_ri,
            'terminal_ri_pv': pv_terminal_ri,
            'total_equity_value': total_equity_value,
            'value_per_share': value_per_share,
            'using_dcf_projections': bool(dcf_projections)
        }
    except Exception as e:
        return {
            'error': True,
            'reason': f'Calculation error: {str(e)}',
            'suggestion': 'Check if financial data is complete and valid. Use DCF as primary valuation method.',
            'technical_details': str(e)
        }

def calculate_dividend_discount_model(financials, shares, cost_of_equity, ticker=None, div_growth_override=None, payout_ratio_override=None, dcf_projections=None, exchange_suffix="NS"):
    """
    Dividend Discount Model (Gordon Growth Model)
    Value = D1 / (Ke - g)
    
    Args:
        div_growth_override: Override dividend growth rate
        payout_ratio_override: Override payout ratio
        dcf_projections: DCF projections dict with 'nopat' key - if provided, uses projected NOPAT for dividend projections (NO DUPLICATION!)
    """
    try:
        # Try to fetch actual dividend data from yfinance
        actual_dividends = []
        div_growth_calculated = None
        payout_ratio_calculated = None
        
        if ticker:
            try:
                stock = get_cached_ticker(get_ticker_with_exchange(ticker, exchange_suffix))
                dividends_hist = stock.dividends

                if not dividends_hist.empty and len(dividends_hist) > 0:
                    # Annual dividends. 'Y' was removed in pandas 3 ('YE' is the year-end alias; older
                    # pandas only knows 'Y'), so try the new alias first.
                    try:
                        dividends_by_year = dividends_hist.resample('YE').sum()
                    except ValueError:
                        dividends_by_year = dividends_hist.resample('Y').sum()
                    # The last bucket is the CURRENT, incomplete calendar year (often 0 or a partial
                    # payment). Drop it when at least 2 complete years remain.
                    _complete = dividends_by_year[dividends_by_year.index.year < pd.Timestamp.today().year]
                    if len(_complete) >= 2:
                        dividends_by_year = _complete
                    if len(dividends_by_year) >= 2:
                        # resample() returns CHRONOLOGICAL order: [oldest ... newest]
                        recent_divs = dividends_by_year[-3:].values
                        actual_dividends = recent_divs.tolist()

                        # Growth = CAGR from oldest (index 0) to newest (index -1). The old code treated
                        # index 0 as the newest, which inverted the sign (rising dividends -> negative growth).
                        if len(actual_dividends) >= 2 and actual_dividends[-1] > 0 and actual_dividends[0] > 0:
                            num_years = len(actual_dividends) - 1
                            div_growth_calculated = ((actual_dividends[-1] / actual_dividends[0]) ** (1 / num_years) - 1) * 100
                            div_growth_calculated = max(-50, min(div_growth_calculated, 150))  # Allow up to 150%

                        # Calculate payout ratio from actual data - USE NEWEST nopat (index 0)
                        latest_div = actual_dividends[-1] if actual_dividends else 0
                        latest_ni = financials['nopat'][0] * 100000
                        if latest_ni > 0 and latest_div > 0:
                            payout_ratio_calculated = (latest_div * shares) / latest_ni
                            payout_ratio_calculated = max(0.1, min(payout_ratio_calculated, 0.9))
            except Exception as e:
                actual_dividends = []
                st.caption(f"ℹ️ DDM: actual dividend history unavailable ({type(e).__name__}: {e}); using payout-ratio estimate instead.")
        
        # Calculate average earnings
        net_incomes = []
        for i in range(len(financials['years'])):
            ni = financials['nopat'][i] * 100000
            net_incomes.append(ni)
        
        avg_net_income = np.mean(net_incomes)
        
        # Use overrides or calculated or default values
        if payout_ratio_override:
            payout_ratio = payout_ratio_override / 100
        elif payout_ratio_calculated:
            payout_ratio = payout_ratio_calculated
        else:
            payout_ratio = 0.40
        
        if div_growth_override:
            div_growth = div_growth_override
        elif div_growth_calculated:
            div_growth = div_growth_calculated
        else:
            div_growth = 8.0
        
        # Calculate dividends
        total_dividends = avg_net_income * payout_ratio
        dps = total_dividends / shares if shares > 0 else 0
        
        # If we have actual dividends, use the latest as current DPS
        if actual_dividends:
            dps = actual_dividends[-1]
        
        # Next year dividend
        d1 = dps * (1 + div_growth / 100)
        
        # DDM valuation
        if cost_of_equity <= div_growth:
            return None
        
        value_per_share = d1 / ((cost_of_equity - div_growth) / 100)
        
        # 5-year dividend projection
        # ✅ USE EXISTING DCF PROJECTIONS IF AVAILABLE - NO DUPLICATION!
        projections_list = []
        
        if dcf_projections and 'nopat' in dcf_projections and len(dcf_projections['nopat']) > 0:
            # Use projected NOPAT from DCF (already calculated!)
            for year_idx, projected_nopat_lacs in enumerate(dcf_projections['nopat'], 1):
                projected_ni = projected_nopat_lacs * 100000  # Convert to Rupees
                projected_dividend = projected_ni * payout_ratio
                projected_dps = projected_dividend / shares if shares > 0 else 0
                pv_div = projected_dps / ((1 + cost_of_equity / 100) ** year_idx)
                projections_list.append({
                    'year': year_idx,
                    'dividend': projected_dps,
                    'pv_dividend': pv_div,
                    'source': 'DCF Projected NOPAT'
                })
        else:
            # Fallback: use growth-based projection
            current_div = dps
            for year in range(1, 6):
                current_div = current_div * (1 + div_growth / 100)
                pv_div = current_div / ((1 + cost_of_equity / 100) ** year)
                projections_list.append({
                    'year': year,
                    'dividend': current_div,
                    'pv_dividend': pv_div,
                    'source': 'Growth-based projection'
                })
        
        return {
            'method': 'Dividend Discount Model',
            'current_dps': dps,
            'payout_ratio': payout_ratio * 100,
            'dividend_growth': div_growth,
            'required_return': cost_of_equity,
            'next_year_dps': d1,
            'projections': projections_list,
            'value_per_share': value_per_share,
            'using_actual_data': bool(actual_dividends),
            'using_dcf_projections': bool(dcf_projections),
            'historical_dividends': actual_dividends if actual_dividends else None
        }
    except Exception as e:
        st.error(f"DDM error: {str(e)}")
        return None

def calculate_pb_roe_valuation(financials, shares, cost_of_equity, assumed_roe=None):
    """
    P/B with ROE Analysis
    Fair P/B = ROE / Cost of Equity
    
    Args:
        assumed_roe: Override ROE if provided
    """
    try:
        # Latest book value - USE NEWEST (index 0)
        latest_equity = financials['equity'][0] * 100000
        book_value_per_share = latest_equity / shares if shares > 0 else 0
        
        # Calculate ROE
        net_incomes = []
        equities = []
        for i in range(len(financials['years'])):
            ni = financials['nopat'][i] * 100000
            eq = financials['equity'][i] * 100000
            net_incomes.append(ni)
            equities.append(eq)
        
        avg_net_income = np.mean(net_incomes)
        avg_equity = np.mean(equities)
        
        # Use provided ROE or calculate
        if assumed_roe:
            roe = assumed_roe
        else:
            roe = (avg_net_income / avg_equity * 100) if avg_equity > 0 else 15
        
        # Fair P/B ratio
        fair_pb = roe / cost_of_equity
        
        # Fair value per share
        value_per_share = book_value_per_share * fair_pb
        
        return {
            'method': 'P/B with ROE Analysis',
            'book_value_per_share': book_value_per_share,
            'roe': roe,
            'cost_of_equity': cost_of_equity,
            'fair_pb_ratio': fair_pb,
            'value_per_share': value_per_share,
            'historical_roe': [(net_incomes[i] / equities[i] * 100) for i in range(len(net_incomes))]
        }
    except Exception as e:
        st.error(f"P/B ROE error: {str(e)}")
        return None


def get_auto_peers_or_default(ticker):
    """
    Try to auto-fetch real industry peers for this specific ticker.
    Returns a comma-separated peer string, or "" if none could be found —
    NEVER a hardcoded generic peer list (previously fell back to a fixed
    set of banking stocks regardless of the company being valued, which
    silently fabricated peer data for non-bank companies).
    """
    if PEER_FETCHER_AVAILABLE:
        try:
            print(f"[DCF] Auto-fetching peers for {ticker}...")
            peers = get_industry_peers(ticker, max_peers=20, exclude_self=True)
            if peers:
                # VERIFY: ticker not in peer list
                ticker_upper = ticker.upper()
                peers = [p for p in peers if p.upper() != ticker_upper]
                peer_str = ",".join(peers)
                print(f"[DCF] ✅ Auto-fetched {len(peers)} peers: {peers}")
                return peer_str
            else:
                print("[DCF] ⚠️ No peers found")
        except Exception as e:
            print(f"[DCF] ⚠️ Error fetching peers: {e}")
    return ""

def calculate_relative_valuation(ticker, financials, shares, peer_tickers=None, exchange_suffix="NS"):
    """
    Relative Valuation using peer multiples
    P/E and P/B comparisons with actual peer data
    
    Handles rate limiting gracefully with retries and delays
    """
    import time
    import random
    
    try:
        if not ticker:
            return None
        
        # Get stock info with rate limit handling
        max_retries = 3
        retry_delay = 2  # seconds
        current_price = 0  # was only assigned inside `if info:` -> UnboundLocalError when Yahoo returns empty info
        
        for attempt in range(max_retries):
            try:
                stock = get_cached_ticker(get_ticker_with_exchange(ticker, exchange_suffix))
                info = stock.info if stock else None
                # Robust price fetching - try multiple methods
                if info:
                    current_price = info.get('currentPrice', 0)
                    if not current_price or current_price == 0:
                        current_price = info.get('regularMarketPrice', 0)
                if not current_price or current_price == 0:
                    try:
                        hist = stock.history(period='1d')
                        if not hist.empty:
                            current_price = hist['Close'].iloc[-1]
                    except:
                        pass
                break
            except Exception as e:
                if "rate" in str(e).lower() or "429" in str(e):
                    if attempt < max_retries - 1:
                        wait_time = retry_delay * (attempt + 1) + random.uniform(0.5, 1.5)
                        st.warning(f"⏳ Rate limit hit. Waiting {wait_time:.1f}s... (Attempt {attempt + 1}/{max_retries})")
                        time.sleep(wait_time)
                    else:
                        st.error("❌ Rate limit exceeded. Using fallback calculations.")
                        current_price = 0
                else:
                    raise
        
        # Calculate company metrics - USE NEWEST (index 0)
        latest_ni = financials['nopat'][0] * 100000
        eps = latest_ni / shares if shares > 0 else 0
        
        # USE NEWEST equity (index 0)
        latest_equity = financials['equity'][0] * 100000
        bvps = latest_equity / shares if shares > 0 else 0
        
        current_pe = current_price / eps if eps > 0 else 0
        current_pb = current_price / bvps if bvps > 0 else 0
        
        # Fetch peer multiples with rate limiting
        peer_pe_list = []
        peer_pb_list = []
        peer_data = []
        
        # Auto-fetch peers if none were provided by the caller. No hardcoded
        # fallback peer list here — if auto-fetch can't find real peers for
        # this specific company, we say so and skip the multiples-based
        # valuation rather than substituting unrelated companies.
        if not peer_tickers:
            peer_tickers = get_auto_peers_or_default(ticker)

        if not peer_tickers:
            st.warning("⚠️ No peer companies available (auto-fetch found none and none were provided manually) — "
                       "relative (P/E / P/B) valuation skipped. Enter peers manually to enable it.")
            return None

        peers = [t.strip() for t in peer_tickers.split(',') if t.strip()]
        
        st.info(f"📊 Fetching data for {len(peers)} peer companies...")
        
        for i, peer in enumerate(peers):
            try:
                # Add delay between requests to avoid rate limiting
                if i > 0:
                    time.sleep(random.uniform(1.0, 1.5))  # Reduced from 2-3s since caching prevents duplicates
                
                peer_stock = get_cached_ticker(get_ticker_with_exchange(peer, exchange_suffix))
                peer_info = peer_stock.info if peer_stock else None
                
                if not peer_info:
                    continue
                
                peer_pe = peer_info.get('trailingPE', 0)
                peer_pb = peer_info.get('priceToBook', 0)
                # Robust price fetching for peers
                peer_price = peer_info.get('currentPrice', 0)
                if not peer_price or peer_price == 0:
                    peer_price = peer_info.get('regularMarketPrice', 0)
                if not peer_price or peer_price == 0:
                    try:
                        hist = peer_stock.history(period='1d')
                        if not hist.empty:
                            peer_price = hist['Close'].iloc[-1]
                    except:
                        pass
                
                if peer_pe and peer_pe > 0 and peer_pe < 100:  # Sanity check
                    peer_pe_list.append(peer_pe)
                
                if peer_pb and peer_pb > 0 and peer_pb < 20:  # Sanity check
                    peer_pb_list.append(peer_pb)
                
                if peer_price > 0:
                    peer_data.append({
                        'ticker': peer,
                        'price': peer_price,
                        'pe': peer_pe if peer_pe else 'N/A',
                        'pb': peer_pb if peer_pb else 'N/A'
                    })
            except Exception as e:
                if "rate" in str(e).lower() or "429" in str(e):
                    st.warning(f"⏳ Rate limit hit on peer {peer}. Skipping...")
                    time.sleep(2)  # Longer delay after rate limit
                continue
        
        # Calculate sector averages — only from real peer data actually
        # fetched. No fabricated "typical industry P/E of 20" fallback:
        # if we have no real peer multiples, we report that plainly instead
        # of presenting an invented number as a data-driven fair value.
        have_pe = bool(peer_pe_list)
        have_pb = bool(peer_pb_list)

        if have_pe:
            sector_avg_pe = np.median(peer_pe_list)  # Use median to avoid outliers
            sector_low_pe = np.percentile(peer_pe_list, 25)
            sector_high_pe = np.percentile(peer_pe_list, 75)
            st.success(f"✅ Fetched {len(peer_pe_list)} peer P/E ratios")
        else:
            st.warning("⚠️ No usable peer P/E data fetched — P/E-based fair value skipped (not fabricated).")
            sector_avg_pe = sector_low_pe = sector_high_pe = None
        
        if have_pb:
            sector_avg_pb = np.median(peer_pb_list)
            sector_low_pb = np.percentile(peer_pb_list, 25)
            sector_high_pb = np.percentile(peer_pb_list, 75)
        else:
            st.warning("⚠️ No usable peer P/B data fetched — P/B-based fair value skipped (not fabricated).")
            sector_avg_pb = sector_low_pb = sector_high_pb = None
        
        # Fair value based on sector multiples — only computed where we
        # actually have real peer-derived multiples for that metric.
        fair_value_pe = eps * sector_avg_pe if (have_pe and eps) else None
        fair_value_pb = bvps * sector_avg_pb if (have_pb and bvps) else None
        
        # Conservative and aggressive estimates (P/E-range based)
        conservative_value = eps * sector_low_pe if (have_pe and eps) else None
        aggressive_value = eps * sector_high_pe if (have_pe and eps) else None
        
        _valid_fair_values = [v for v in (fair_value_pe, fair_value_pb) if v is not None]

        return {
            'method': 'Relative Valuation',
            'current_price': current_price,
            'eps': eps,
            'bvps': bvps,
            'current_pe': current_pe,
            'current_pb': current_pb,
            'sector_avg_pe': sector_avg_pe,
            'sector_avg_pb': sector_avg_pb,
            'sector_low_pe': sector_low_pe,
            'sector_high_pe': sector_high_pe,
            'sector_low_pb': sector_low_pb,
            'sector_high_pb': sector_high_pb,
            'fair_value_pe_based': fair_value_pe,
            'fair_value_pb_based': fair_value_pb,
            'conservative_value': conservative_value,
            'aggressive_value': aggressive_value,
            'avg_fair_value': (sum(_valid_fair_values) / len(_valid_fair_values)) if _valid_fair_values else None,
            'peer_count': len(peer_pe_list),
            'peer_data': peer_data,
            'rate_limited': len(peer_pe_list) == 0  # Flag if we got rate limited
        }
    except Exception as e:
        error_msg = str(e)
        if "rate" in error_msg.lower() or "429" in error_msg:
            st.error("❌ **Rate Limit Exceeded**")
            st.info("""
            **Why this happens:**
            - Yahoo Finance limits requests to prevent abuse
            - Too many requests in short time
            
            **Solutions:**
            1. Wait 5-10 minutes and try again
            2. Reduce number of peer companies
            3. Use the app during off-peak hours
            4. Consider using cached/default multiples
            
            **Alternative:** Use P/B ROE model or DDM for valuation instead
            """)
        else:
            st.error(f"Relative valuation error: {error_msg}")
        return None

# ================================
# YAHOO FINANCE SCRAPING WITH CACHING AND RETRY
# ================================

@st.cache_data(ttl=3600)  # Cache for 1 hour
def fetch_yahoo_financials_cached(ticker, exchange_suffix="NS"):
    """Cached wrapper for Yahoo Finance fetch"""
    return fetch_yahoo_financials_internal(ticker, exchange_suffix)

def fetch_yahoo_financials(ticker, exchange_suffix="NS"):
    """
    Fetch financial statements from Yahoo Finance with comprehensive error handling
    
    Features:
    - Retry logic with exponential backoff
    - Rate limit detection
    - Session caching
    - Automatic delays
    """
    import time
    import random
    
    max_retries = 3
    base_delay = 3  # seconds
    
    for attempt in range(max_retries):
        try:
            # Add delay before request (except first attempt)
            if attempt > 0:
                delay = base_delay * (2 ** attempt) + random.uniform(1, 3)
                st.warning(f"⏳ Rate limit detected. Waiting {delay:.1f} seconds... (Attempt {attempt + 1}/{max_retries})")
                time.sleep(delay)
            elif attempt == 0 and hasattr(st.session_state, 'last_yahoo_request'):
                # Add small delay between different requests
                elapsed = time.time() - st.session_state.last_yahoo_request
                if elapsed < 1:
                    time.sleep(1 - elapsed)
            
            # Try cached version first
            if attempt == 0:
                try:
                    return fetch_yahoo_financials_cached(ticker, exchange_suffix)
                except:
                    pass  # If cache fails, continue to direct fetch
            
            # Direct fetch
            result = fetch_yahoo_financials_internal(ticker, exchange_suffix)
            
            # Track request time and count
            st.session_state.last_yahoo_request = time.time()
            st.session_state.yahoo_request_count += 1
            
            return result
            
        except Exception as e:
            error_str = str(e).lower()
            
            if any(keyword in error_str for keyword in ['rate', '429', 'too many', 'limit']):
                if attempt < max_retries - 1:
                    continue  # Try again with delay
                else:
                    # All retries exhausted
                    st.error("❌ **Yahoo Finance Rate Limit Exceeded**")
                    st.warning("""
                    **Too many requests to Yahoo Finance. Please try one of these options:**
                    
                    1. ⏰ **Wait 10-15 minutes** and try again
                    2. 🌐 **Use a different network** (mobile hotspot, VPN)
                    3. 🔄 **Restart your Streamlit app** to clear session
                    4. 📊 **Use the Excel upload feature** for unlisted companies instead
                    
                    **Why this happens:**
                    Yahoo Finance limits free API requests to prevent abuse. This is a Yahoo limitation, not an issue with our app.
                    """)
                    return None, "Rate limit exceeded. Please try again later or use alternative data sources."
            else:
                # Other error
                if attempt < max_retries - 1:
                    continue
                else:
                    return None, f"Error fetching data: {str(e)}"
    
    return None, "Failed to fetch data after multiple retries"

def fetch_yahoo_financials_internal(ticker, exchange_suffix="NS"):
    """Internal function for actual Yahoo Finance fetch"""
    try:
        stock = get_cached_ticker(get_ticker_with_exchange(ticker, exchange_suffix))
        
        # Get financial statements
        income_stmt = stock.financials
        balance_sheet = stock.balance_sheet
        cash_flow = stock.cashflow
        
        # Get company info
        info = stock.info if stock else None
        
        if income_stmt.empty or balance_sheet.empty:
            return None, "No financial data available"
        
        # Get shares outstanding - ROBUST MULTI-METHOD APPROACH
        shares = 0
        
        # Method 1: Direct from info
        if info:
            shares = info.get('sharesOutstanding', 0)
        
        # Method 2: Implied shares outstanding
        if (shares == 0 or shares is None) and info:
            shares = info.get('impliedSharesOutstanding', 0)
        
        # Method 3: From balance sheet (Total Common Stock / Par Value)
        if (shares == 0 or shares is None) and not balance_sheet.empty:
            try:
                # Try to get from Common Stock or Share Capital
                for row_name in ['Common Stock', 'Share Capital', 'Ordinary Shares Capital', 'Common Stock Equity']:
                    if row_name in balance_sheet.index:
                        common_stock = balance_sheet.loc[row_name].iloc[0]
                        # Typically par value is ₹1, ₹2, ₹5, or ₹10
                        # Try different par values to estimate
                        for par_value in [1, 2, 5, 10]:
                            estimated_shares = abs(common_stock) / par_value
                            # Sanity check: shares should be reasonable (between 1M and 100B)
                            if 1_000_000 < estimated_shares < 100_000_000_000:
                                shares = estimated_shares
                                break
                        if shares > 0:
                            break
            except:
                pass
        
        # Method 4: Calculate from market cap and price
        if (shares == 0 or shares is None) and info and 'marketCap' in info and 'currentPrice' in info:
            market_cap = info.get('marketCap', 0)
            current_price = info.get('currentPrice', 0)
            if market_cap > 0 and current_price > 0:
                shares = market_cap / current_price
        
        # Method 5: From enterprise value and price
        if (shares == 0 or shares is None) and info and 'enterpriseValue' in info and 'currentPrice' in info:
            ev = info.get('enterpriseValue', 0)
            price = info.get('currentPrice', 0)
            if ev > 0 and price > 0:
                # Rough estimate assuming EV ≈ Market Cap for many companies
                shares = ev / price
        
        # Final check
        if shares is None:
            shares = 0
        
        shares_source = "Unknown"
        if shares > 0:
            if info and info.get('sharesOutstanding', 0) > 0:
                shares_source = "Direct (sharesOutstanding)"
            elif info.get('impliedSharesOutstanding', 0) > 0:
                shares_source = "Implied shares"
            elif 'marketCap' in info and 'currentPrice' in info:
                shares_source = "Calculated from Market Cap"
            else:
                shares_source = "Estimated from Balance Sheet"
        
        return {
            'income_statement': income_stmt,
            'balance_sheet': balance_sheet,
            'cash_flow': cash_flow,
            'info': info,
            'shares': shares,
            'shares_source': shares_source  # Track how shares were obtained
        }, None
        
    except Exception as e:
        return None, f"Error fetching data: {str(e)}"

def extract_financials_listed(yahoo_data, num_years=3):
    """
    Extract financial metrics from Yahoo Finance OR Screener.in data
    
    Args:
        yahoo_data: Dictionary containing either Yahoo Finance data OR Screener.in data
        num_years: Number of historical years to extract (default 3)
    
    Returns:
        Dictionary with financial metrics in ₹ Lacs
    """
    
    # CHECK IF THIS IS SCREENER.IN DATA
    if '_screener_financials' in yahoo_data:
        # SCREENER.IN DATA PATH
        st.info("📊 Using Screener.in financial data")
        financials = yahoo_data['_screener_financials']
        
        # BUGFIX: Validate that all required fields exist and have data
        required_fields = [
            'years', 'revenue', 'cogs', 'opex', 'ebitda', 'depreciation',
            'ebit', 'interest', 'tax', 'nopat', 'fixed_assets', 'inventory',
            'receivables', 'payables', 'cash', 'equity', 'st_debt', 'lt_debt'
        ]
        
        missing_fields = []
        empty_fields = []
        for field in required_fields:
            if field not in financials:
                missing_fields.append(field)
            elif isinstance(financials[field], list) and len(financials[field]) == 0:
                empty_fields.append(field)
            elif isinstance(financials[field], list) and all(v == 0 for v in financials[field]):
                empty_fields.append(field + " (all zeros)")
        
        if missing_fields:
            st.error(f"❌ Missing financial fields from Screener.in: {', '.join(missing_fields)}")
            st.info("💡 Try using Yahoo Finance data source instead")
        
        if empty_fields:
            st.warning(f"⚠️ Empty or zero financial fields from Screener.in: {', '.join(empty_fields)}")
            st.info("💡 This may be due to incomplete data on Screener.in for this stock")
            
            # DEBUG: Show actual values for troubleshooting
            with st.expander("🔍 Debug: View Raw Screener Data"):
                debug_data = {
                    'Fixed Assets': financials.get('fixed_assets', []),
                    'Inventory': financials.get('inventory', []),
                    'Receivables': financials.get('receivables', []),
                    'Payables': financials.get('payables', []),
                    'Equity': financials.get('equity', []),
                    'ST Debt': financials.get('st_debt', []),
                    'LT Debt': financials.get('lt_debt', [])
                }
                st.json(debug_data)
        
        # Screener data is already in the correct format (₹ Lacs)
        # Just ensure we only return the requested number of years
        if num_years < len(financials['years']):
            # Truncate to requested years (newest first)
            for key in financials.keys():
                if isinstance(financials[key], list):
                    financials[key] = financials[key][:num_years]
        
        return financials
    
    # YAHOO FINANCE DATA PATH (ORIGINAL CODE)
    try:
        income_stmt = yahoo_data['income_statement']
        balance_sheet = yahoo_data['balance_sheet']
        cash_flow = yahoo_data['cash_flow']
        
        # Get last N years (columns are sorted newest to oldest)
        years = income_stmt.columns[:min(num_years, len(income_stmt.columns))]
        
        opex_methods_used = []  # Track which method was used for each year
        
        financials = {
            'years': [str(y.year) for y in years],
            'revenue': [],
            'cogs': [],
            'opex': [],
            'ebitda': [],
            'depreciation': [],
            'ebit': [],
            'interest': [],
            'interest_income': [],  # Added for business classification
            'tax': [],
            'nopat': [],
            'fixed_assets': [],
            'inventory': [],
            'receivables': [],
            'payables': [],
            'cash': [],
            'equity': [],
            'st_debt': [],
            'lt_debt': [],
        }
        
        for year in years:
            # Income Statement - Values are already in the correct currency
            # ROBUST: Use safe_extract to handle None values
            revenue = safe_extract(income_stmt, 'Total Revenue', year, default=0)
            cogs = safe_extract(income_stmt, 'Cost Of Revenue', year, default=0)
            
            # Try to get Operating Expenses directly from various fields
            opex = 0
            opex_method = "None"
            
            # Method 1: Try direct operating expense fields
            opex_fields = [
                'Operating Expense',
                'Total Operating Expenses',
                'Operating Expenses',
                'Selling General And Administration',
                'Selling General Administrative',
                'SG&A Expense'
            ]
            
            for field in opex_fields:
                if field in income_stmt.index:
                    opex = abs(income_stmt.loc[field, year])
                    opex_method = f"Method 1: Direct field '{field}'"
                    break
            
            # Method 2: If not found, try to calculate from Gross Profit - Operating Income
            if opex == 0:
                gross_profit = 0
                if 'Gross Profit' in income_stmt.index:
                    gross_profit = abs(income_stmt.loc['Gross Profit', year])
                elif revenue > 0 and cogs > 0:
                    gross_profit = revenue - cogs
                
                operating_income = abs(income_stmt.loc['Operating Income', year]) if 'Operating Income' in income_stmt.index else 0
                
                if gross_profit > 0 and operating_income > 0:
                    opex = gross_profit - operating_income
                    opex_method = "Method 2: Gross Profit - Operating Income"
            
            # Method 3: If still not found, try SG&A + R&D + Other
            if opex == 0:
                sga = abs(income_stmt.loc['Selling General And Administration', year]) if 'Selling General And Administration' in income_stmt.index else 0
                rd = abs(income_stmt.loc['Research And Development', year]) if 'Research And Development' in income_stmt.index else 0
                other_opex = abs(income_stmt.loc['Other Operating Expenses', year]) if 'Other Operating Expenses' in income_stmt.index else 0
                opex = sga + rd + other_opex
                if opex > 0:
                    opex_method = f"Method 3: SG&A({sga/100000:.2f}) + R&D({rd/100000:.2f}) + Other({other_opex/100000:.2f})"
            
            # Get EBITDA - ROBUST: Handle None values
            ebitda = (
                safe_extract(income_stmt, 'EBITDA', year) or
                safe_extract(income_stmt, 'Normalized EBITDA', year) or
                max(0, revenue - cogs - opex)  # Calculate if not available
            )
            ebitda = ensure_valid_number(ebitda, default=0)
            
            # Get depreciation separately for projections - ROBUST
            depreciation = (
                safe_extract(income_stmt, 'Reconciled Depreciation', year) or
                safe_extract(cash_flow, 'Depreciation And Amortization', year) or
                safe_extract(income_stmt, 'Depreciation', year) or
                0
            )
            
            # Fallback: Calculate from Operating Income vs EBITDA or use revenue-based estimate
            if depreciation == 0:
                operating_income = safe_extract(income_stmt, 'Operating Income', year, default=0)
                if ebitda > operating_income and operating_income > 0:
                    depreciation = ebitda - operating_income
                else:
                    depreciation = revenue * 0.02 if revenue > 0 else 0  # 2% of revenue fallback
            
            depreciation = ensure_valid_number(depreciation, default=0)
            
            # Final safety check: If opex is still 0 or unreasonable, derive from EBITDA
            if opex == 0 or opex < 0:
                opex = revenue - cogs - ebitda
                opex_method = "Method 4 (Fallback): Revenue - COGS - EBITDA"
                if opex < 0:
                    opex = revenue * 0.15  # Assume 15% of revenue as default
                    opex_method = "Method 5 (Default): 15% of Revenue"
            
            # EBIT
            ebit = ebitda - depreciation
            
            # Interest Expense - ROBUST: Handle None values
            interest = (
                safe_extract(income_stmt, 'Interest Expense', year) or
                safe_extract(income_stmt, 'Interest Expense Non Operating', year) or
                0
            )
            
            # For banks: check if Net Interest Income is negative (then it's an expense)
            if interest == 0 and 'Net Interest Income' in income_stmt.index:
                net_int = safe_extract(income_stmt, 'Net Interest Income', year, default=0)
                if net_int < 0:
                    interest = abs(net_int)
            
            interest = ensure_valid_number(interest, default=0)
            
            # Interest Income (for business classification) - ROBUST
            interest_income = (
                safe_extract(income_stmt, 'Interest Income', year) or
                safe_extract(income_stmt, 'Interest And Dividend Income', year) or
                0
            )
            
            # For banks: Net Interest Income is primary revenue if positive
            if interest_income == 0 and 'Net Interest Income' in income_stmt.index:
                net_int = safe_extract(income_stmt, 'Net Interest Income', year, default=0)
                if net_int > 0:
                    interest_income = abs(net_int)
            
            interest_income = ensure_valid_number(interest_income, default=0)
            
            # Tax - ROBUST
            tax = safe_extract(income_stmt, 'Tax Provision', year, default=0)
            tax = ensure_valid_number(tax, default=0)
            
            # NOPAT (using 25% tax as default) - ROBUST calculation
            ebt = ebit - interest
            if ebt > 0 and tax > 0:
                tax_rate_effective = safe_divide(tax, ebt, default=0.25)
                tax_rate_effective = min(max(tax_rate_effective, 0), 0.35)  # Clamp between 0 and 35%
            else:
                tax_rate_effective = 0.25
            
            nopat = ebit * (1 - tax_rate_effective)
            nopat = ensure_valid_number(nopat, default=0)
            
            # Balance Sheet - Values are already in the correct currency
            # ROBUST: Use safe_extract for all balance sheet items to handle None values
            total_assets = safe_extract(balance_sheet, 'Total Assets', year, default=0)
            
            # Fixed Assets
            fixed_assets = (
                safe_extract(balance_sheet, 'Net PPE', year) or
                safe_extract(balance_sheet, 'Gross PPE', year) or
                safe_extract(balance_sheet, 'Properties', year) or
                (total_assets * 0.3 if total_assets > 0 else 0)  # Fallback: 30% of total assets
            )
            
            # Current Assets - ROBUST: These can be None for many companies
            inventory = safe_extract(balance_sheet, 'Inventory', year, default=0)
            
            receivables = (
                safe_extract(balance_sheet, 'Receivables', year) or
                safe_extract(balance_sheet, 'Accounts Receivable', year) or
                safe_extract(balance_sheet, 'Gross Accounts Receivable', year) or
                0
            )
            
            cash = (
                safe_extract(balance_sheet, 'Cash And Cash Equivalents', year) or
                safe_extract(balance_sheet, 'Cash Cash Equivalents And Short Term Investments', year) or
                0
            )
            
            # Liabilities - ROBUST: These can also be None
            payables = (
                safe_extract(balance_sheet, 'Payables', year) or
                safe_extract(balance_sheet, 'Accounts Payable', year) or
                safe_extract(balance_sheet, 'Payables And Accrued Expenses', year) or
                0
            )
            
            # Debt - ROBUST: Handle None debt values
            st_debt = (
                safe_extract(balance_sheet, 'Current Debt', year) or
                safe_extract(balance_sheet, 'Current Debt And Capital Lease Obligation', year) or
                0
            )
            
            lt_debt = (
                safe_extract(balance_sheet, 'Long Term Debt', year) or
                safe_extract(balance_sheet, 'Long Term Debt And Capital Lease Obligation', year) or
                0
            )
            
            # Equity - ROBUST: Critical for WACC calculation
            equity = (
                safe_extract(balance_sheet, 'Stockholders Equity', year) or
                safe_extract(balance_sheet, 'Total Equity Gross Minority Interest', year) or
                safe_extract(balance_sheet, 'Common Stock Equity', year) or
                0
            )
            
            # Convert to Lacs (divide by 100,000)
            # ROBUST: Ensure all values are valid numbers before storage
            # Yahoo Finance data is in actual currency (Rupees for Indian stocks)
            financials['revenue'].append(ensure_valid_number(revenue / 100000, 0))
            financials['cogs'].append(ensure_valid_number(cogs / 100000, 0))
            financials['opex'].append(ensure_valid_number(opex / 100000, 0))
            financials['ebitda'].append(ensure_valid_number(ebitda / 100000, 0))
            financials['depreciation'].append(ensure_valid_number(depreciation / 100000, 0))
            financials['ebit'].append(ensure_valid_number(ebit / 100000, 0))
            financials['interest'].append(ensure_valid_number(interest / 100000, 0))
            financials['interest_income'].append(ensure_valid_number(interest_income / 100000, 0))
            financials['tax'].append(ensure_valid_number(tax / 100000, 0))
            financials['nopat'].append(ensure_valid_number(nopat / 100000, 0))
            
            financials['fixed_assets'].append(ensure_valid_number(fixed_assets / 100000, 0))
            financials['inventory'].append(ensure_valid_number(inventory / 100000, 0))
            financials['receivables'].append(ensure_valid_number(receivables / 100000, 0))
            financials['payables'].append(ensure_valid_number(payables / 100000, 0))
            financials['cash'].append(ensure_valid_number(cash / 100000, 0))
            financials['equity'].append(ensure_valid_number(equity / 100000, 0))
            financials['st_debt'].append(ensure_valid_number(st_debt / 100000, 0))
            financials['lt_debt'].append(ensure_valid_number(lt_debt / 100000, 0))
            
            # Track which method was used for opex
            opex_methods_used.append(f"{year.year}: {opex_method}")
        
        
        return financials
        
    except Exception as e:
        st.error(f"Error extracting financials: {str(e)}")
        import traceback
        st.error(traceback.format_exc())
        return None

def get_stock_beta(ticker, market_ticker=None, period_years=3,
                   beta_start_date=None, beta_end_date=None):
    """Calculate beta using daily return regression of stock vs market.

    Date priority:
        beta_start_date / beta_end_date  →  explicit window (user-chosen)
        period_years                     →  fallback rolling window
    
    Market index auto-selection:
        NSE (.NS)  →  NIFTY 50  (^NSEI)
        BSE (.BO)  →  SENSEX    (^BSESN)
    
    Returns: float beta (clamped 0.1 – 3.0)
    """
    try:
        if not ticker:
            st.warning("⚠️ Invalid ticker provided — using default β=1.0")
            return 1.0

        # Ensure exchange suffix
        if '.NS' not in ticker and '.BO' not in ticker:
            ticker = ticker + '.NS'

        # Auto-select market index
        if market_ticker is None:
            market_ticker = '^NSEI' if '.NS' in ticker else '^BSESN'

        # ── Date window ───────────────────────────────────────────────────
        if beta_start_date is not None and beta_end_date is not None:
            start_date = pd.Timestamp(beta_start_date)
            end_date   = pd.Timestamp(beta_end_date)
        else:
            end_date   = pd.Timestamp(datetime.now())
            start_date = end_date - pd.DateOffset(years=period_years)

        # ── Fetch daily OHLCV ─────────────────────────────────────────────
        stock_obj  = yf.Ticker(ticker)
        market_obj = yf.Ticker(market_ticker)

        stock  = stock_obj.history(start=start_date, end=end_date, period=None)
        market = market_obj.history(start=start_date, end=end_date, period=None)

        # Flatten MultiIndex if present (yfinance ≥ 0.2.x)
        if isinstance(stock.columns, pd.MultiIndex):
            stock.columns = stock.columns.get_level_values(0)
        if isinstance(market.columns, pd.MultiIndex):
            market.columns = market.columns.get_level_values(0)

        if stock.empty or market.empty:
            st.warning(f"⚠️ No price data for {ticker} in chosen period — using β=1.0")
            return 1.0

        # ── Daily returns (pct_change = day-over-day) ─────────────────────
        stock_returns  = stock['Close'].pct_change().dropna()
        market_returns = market['Close'].pct_change().dropna()

        # Align on common trading days
        aligned = pd.concat([stock_returns, market_returns], axis=1, join='inner').dropna()
        aligned.columns = ['stock', 'market']

        n_obs = len(aligned)
        if n_obs < 20:
            st.warning(f"⚠️ Only {n_obs} overlapping trading days for {ticker} — using β=1.0")
            return 1.0

        # ── OLS β = Cov(Rs, Rm) / Var(Rm) ───────────────────────────────
        cov    = aligned['stock'].cov(aligned['market'])
        var_m  = aligned['market'].var()

        if var_m == 0:
            return 1.0

        beta = cov / var_m

        # Caption for transparency
        index_name = "NIFTY 50" if market_ticker == '^NSEI' else "SENSEX"
        st.caption(
            f"   β vs {index_name} | {n_obs} daily obs "
            f"({aligned.index[0].strftime('%d-%b-%Y')} → {aligned.index[-1].strftime('%d-%b-%Y')})"
        )

        return max(0.1, min(beta, 3.0))

    except Exception as e:
        st.warning(f"Could not calculate beta for {ticker}: {str(e)}")
        return 1.0

def get_risk_free_rate(custom_ticker=None, as_of_date=None):
    """
    Get risk-free rate by calculating historical CAGR from Yahoo Finance ticker data.
    
    For bond indices/yields: Uses the closing value directly if it's already a percentage
    For stocks/indices: Calculates CAGR from price history

    Args:
        custom_ticker: override ticker (defaults to ^TNX)
        as_of_date: if provided, treat this date as "today" — i.e. only use
                    data up to and including this date, so the rate reflects
                    what was known/observable as of the valuation date rather
                    than the live market as of right now.
    
    Returns:
        tuple: (rate, debug_messages_list)
    """
    from datetime import datetime, timedelta
    pass  # yf available via module-level shim (yf_ratelimit)
    import pandas as pd
    
    ticker = custom_ticker if custom_ticker else '^TNX'
    debug = []
    
    debug.append(f"🔍 Fetching data for: `{ticker}`")
    
    try:
        # Fetch maximum available historical data
        end_date = pd.Timestamp(as_of_date) if as_of_date else datetime.now()
        start_date = end_date - timedelta(days=365*20)  # Try to get 20 years
        if as_of_date:
            debug.append(f"📌 **Valuation date mode** — using data as of {end_date.date()}")
        
        debug.append(f"📅 **Fetching data** from {start_date.date()} to {end_date.date()}")
        
        # Try to get data from Yahoo Finance
        ticker_obj = yf.Ticker(ticker)
        if as_of_date:
            gsec_data = ticker_obj.history(start=start_date, end=end_date + timedelta(days=1))
        else:
            gsec_data = ticker_obj.history(period='max')
        
        if len(gsec_data) < 2:
            gsec_data = ticker_obj.history(start=start_date, end=end_date + timedelta(days=1), period=None)
        
        if len(gsec_data) < 2:
            gsec_data = yf.download(ticker, start=start_date, end=end_date + timedelta(days=1), progress=False)
            if isinstance(gsec_data.columns, pd.MultiIndex):
                gsec_data.columns = gsec_data.columns.get_level_values(0)

        # If an as_of_date was given, make sure we never look past it
        # (period='max'/download windows can otherwise run to "now").
        if as_of_date and len(gsec_data) > 0:
            gsec_data = gsec_data[gsec_data.index <= end_date]
        
        # If we got good data, proceed with it
        if len(gsec_data) >= 2:
            debug.append(f"✅ **Successfully fetched {len(gsec_data)} rows of data**")
        else:
            # Check if we have a curated rate for this ticker
            ticker_upper = ticker.upper()
            curated_rates = {
                'NIFTYGS10YR.NS': (6.68, 'India 10Y G-Sec'),
                'NIFTYGS10YR': (6.68, 'India 10Y G-Sec'),
                'NIFTY GS 10YR': (6.68, 'India 10Y G-Sec'),
                'NIFTYGS5YR.NS': (6.43, 'India 5Y G-Sec'),
                'NIFTYGS5YR': (6.43, 'India 5Y G-Sec'),
                'NIFTY GS 5YR': (6.43, 'India 5Y G-Sec'),
                'NIFTYGS3YR.NS': (6.05, 'India 3Y G-Sec'),
                'NIFTYGS3YR': (6.05, 'India 3Y G-Sec'),
                'NIFTY GS 3YR': (6.05, 'India 3Y G-Sec'),
                '^TNX': (4.50, 'US 10Y Treasury'),
            }
            
            if ticker_upper in curated_rates:
                curated_rate, description = curated_rates[ticker_upper]
                debug.append(f"📊 **Using market rate for {description}**: {curated_rate}%")
                debug.append(f"💡 Source: Current market data (Feb 2026)")
                debug.append(f"✅ You can manually override this value in the field below")
                return curated_rate, debug
            else:
                # No curated rate available
                debug.append(f"⚠️ Unable to fetch historical data for ticker: {ticker}")
                debug.append(f"💡 **Try these alternatives:**")
                debug.append(f"   • Use ^TNX (US 10Y Treasury) - reliable data")
                debug.append(f"   • Manually enter rate in the field below")
                fallback = 6.83
                debug.append(f"⚠️ Using default fallback: {fallback}%")
                return fallback, debug
        
        debug.append(f"📊 **Downloaded {len(gsec_data)} rows of data**")
        
        # Show actual date range and data
        if len(gsec_data) > 0:
            debug.append(f"📋 Date range: {gsec_data.index[0].date()} to {gsec_data.index[-1].date()}")
        
        # Extract close prices
        close_prices = gsec_data['Close'].dropna()
        
        if len(close_prices) < 2:
            debug.append(f"⚠️ Insufficient data points: {len(close_prices)}")
            fallback = 6.83
            return fallback, debug
        
        # Get first and last prices
        first_price = float(close_prices.iloc[0])
        last_price = float(close_prices.iloc[-1])
        
        # Calculate time period
        first_date = close_prices.index[0]
        last_date = close_prices.index[-1]
        days_diff = (last_date - first_date).days
        years = days_diff / 365.25
        
        debug.append(f"📊 Period: {years:.1f} years | Price: {first_price:.2f} → {last_price:.2f}")
        
        # Determine if this is a yield/rate or a price
        avg_price = close_prices.mean()
        price_volatility = close_prices.std() / avg_price if avg_price > 0 else 0
        
        # Decision logic
        if avg_price < 50 and price_volatility < 0.5:
            # Likely already a yield/rate
            days_to_use = min(90, len(close_prices))
            recent_prices = close_prices.iloc[-days_to_use:]
            avg_rate = recent_prices.mean()
            
            debug.append(f"📈 Interpretation: Direct yield/rate (avg={avg_price:.2f})")
            debug.append(f"✅ Result: {avg_rate:.2f}% (90-day average)")
            
            return round(avg_rate, 2), debug
            
        else:
            # Calculate CAGR for stocks/indices
            if first_price <= 0 or years <= 0:
                debug.append(f"❌ Cannot calculate CAGR (invalid data)")
                fallback = 6.83
                return fallback, debug
            
            cagr = ((last_price / first_price) ** (1 / years) - 1) * 100
            
            debug.append(f"📈 Interpretation: Price data, calculating CAGR")
            debug.append(f"✅ Result: {cagr:.2f}% CAGR over {years:.1f} years")
            
            return round(cagr, 2), debug
        
    except Exception as e:
        debug.append(f"❌ Error: {type(e).__name__}: {str(e)[:100]}")
        fallback = 6.83
        return fallback, debug

def get_market_return(custom_ticker=None, as_of_date=None):
    """
    Calculate market return (CAGR) from historical index data.
    
    Args:
        custom_ticker: Yahoo Finance ticker (use % prefix, e.g., %5ENSEBANK for ^NSEBANK)
        as_of_date: if provided, treat this date as "today" — only use data up
                    to and including this date, so the return reflects what was
                    observable as of the valuation date rather than live data.
    
    Returns:
        tuple: (return_rate, debug_messages_list)
    """
    from datetime import datetime, timedelta
    pass  # yf available via module-level shim (yf_ratelimit)
    import pandas as pd
    
    ticker = custom_ticker if custom_ticker else '%5EBSESN'  # Default to BSE Sensex
    
    debug = []
    debug.append(f"🔍 Fetching market data for: `{ticker}`")
    
    try:
        end_date = pd.Timestamp(as_of_date) if as_of_date else datetime.now()
        if as_of_date:
            debug.append(f"📌 **Valuation date mode** — using data as of {end_date.date()}")
        
        debug.append(f"📅 Requesting maximum available historical data")
        
        # Try to get ALL available data using period='max'
        ticker_obj = yf.Ticker(ticker)
        if as_of_date:
            start_date = end_date - timedelta(days=365*30)
            market_data = ticker_obj.history(start=start_date, end=end_date + timedelta(days=1))
        else:
            market_data = ticker_obj.history(period='max')
        
        if len(market_data) < 2:
            # Fallback: try with date range
            start_date = end_date - timedelta(days=365*30)  # 30 years as fallback
            market_data = ticker_obj.history(start=start_date, end=end_date + timedelta(days=1), period=None)
        
        if len(market_data) < 2:
            if as_of_date:
                market_data = yf.download(ticker, start=end_date - timedelta(days=365*30), end=end_date + timedelta(days=1), progress=False)
            else:
                market_data = yf.download(ticker, period='max', progress=False)
            if isinstance(market_data.columns, pd.MultiIndex):
                market_data.columns = market_data.columns.get_level_values(0)

        # Never look past the valuation date, even if a period='max'/download
        # window pulled data all the way to "now".
        if as_of_date and len(market_data) > 0:
            market_data = market_data[market_data.index <= end_date]
        
        # If we got good data, calculate CAGR
        if len(market_data) >= 252:  # At least 1 year
            debug.append(f"✅ Successfully fetched {len(market_data)} rows")
            
            start_price = float(market_data['Close'].iloc[0])
            end_price = float(market_data['Close'].iloc[-1])
            num_years = len(market_data) / 252  # 252 trading days per year
            
            if start_price > 0 and num_years > 0:
                cagr = ((end_price / start_price) ** (1 / num_years) - 1) * 100
                
                debug.append(f"📊 Period: {num_years:.1f} years | Price: {start_price:.2f} → {end_price:.2f}")
                debug.append(f"✅ Market CAGR: {cagr:.2f}%")
                
                return round(cagr, 2), debug
        
        # Curated fallback rates
        ticker_upper = ticker.upper()
        curated_market_returns = {
            '%5EBSESN': (12.5, 'BSE Sensex'),
            '%5ENSEI': (12.8, 'NSE Nifty 50'),
            '%5ENSEBANK': (15.2, 'Nifty Bank'),
            '%5EGSPC': (10.5, 'S&P 500'),
            '%5EDJI': (9.8, 'Dow Jones'),
        }
        
        if ticker_upper in curated_market_returns:
            curated_rate, description = curated_market_returns[ticker_upper]
            debug.append(f"📊 Using historical average for {description}: {curated_rate}%")
            debug.append(f"💡 Source: Long-term market averages")
            debug.append(f"✅ You can manually override this value below")
            return curated_rate, debug
        
        # Generic fallback
        debug.append(f"⚠️ Unable to fetch data for {ticker}")
        debug.append(f"💡 Try: %5EBSESN (Sensex), %5ENSEI (Nifty), or manual entry")
        fallback = 12.0
        return fallback, debug
        
    except Exception as e:
        debug.append(f"❌ Error: {type(e).__name__}: {str(e)[:100]}")
        fallback = 12.0
        return fallback, debug

# ================================
# ADVANCED CHARTING FUNCTIONS
# ================================

def create_waterfall_chart(valuation):
    """Create waterfall chart showing DCF value buildup"""
    fig = go.Figure(go.Waterfall(
        name = "DCF Waterfall",
        orientation = "v",
        measure = ["relative", "relative", "total"],
        x = ["PV of Projected FCFF", "PV of Terminal Value", "Enterprise Value"],
        textposition = "outside",
        text = [f"₹{valuation['sum_pv_fcff']:.2f}L", 
                f"₹{valuation['pv_terminal_value']:.2f}L",
                f"₹{valuation['enterprise_value']:.2f}L"],
        y = [valuation['sum_pv_fcff'], valuation['pv_terminal_value'], 0],
        connector = {"line":{"color":"rgb(63, 63, 63)"}},
    ))
    
    fig.update_layout(
        title = "DCF Valuation Waterfall",
        showlegend = False,
        height = 500,
        yaxis_title="Value (₹ Lacs)"
    )
    
    return fig

def create_fcff_projection_chart(projections):
    """Create detailed FCFF projection chart with components"""
    years = [f"Year {y}" for y in projections['year']]
    
    fig = make_subplots(
        rows=2, cols=2,
        subplot_titles=('FCFF Projection', 'Revenue & EBITDA', 'Working Capital Changes', 'Capex & Depreciation'),
        specs=[[{"secondary_y": False}, {"secondary_y": False}],
               [{"secondary_y": False}, {"secondary_y": False}]]
    )
    
    # FCFF Projection
    fig.add_trace(
        go.Bar(name='FCFF', x=years, y=projections['fcff'], marker_color='#2E86AB'),
        row=1, col=1
    )
    
    # Revenue & EBITDA
    fig.add_trace(
        go.Scatter(name='Revenue', x=years, y=projections['revenue'], mode='lines+markers', line=dict(color='#06A77D', width=3)),
        row=1, col=2
    )
    fig.add_trace(
        go.Scatter(name='EBITDA', x=years, y=projections['ebitda'], mode='lines+markers', line=dict(color='#F77F00', width=3)),
        row=1, col=2
    )
    
    # Working Capital
    fig.add_trace(
        go.Bar(name='Δ WC', x=years, y=projections['delta_wc'], marker_color='#D62828'),
        row=2, col=1
    )
    
    # Capex & Depreciation
    fig.add_trace(
        go.Bar(name='Capex', x=years, y=projections['capex'], marker_color='#F77F00'),
        row=2, col=2
    )
    fig.add_trace(
        go.Bar(name='Depreciation', x=years, y=projections['depreciation'], marker_color='#06A77D'),
        row=2, col=2
    )
    
    fig.update_layout(height=800, showlegend=True, title_text="Comprehensive Financial Projections")
    fig.update_xaxes(title_text="Year", row=2, col=1)
    fig.update_xaxes(title_text="Year", row=2, col=2)
    fig.update_yaxes(title_text="₹ Lacs", row=1, col=1)
    fig.update_yaxes(title_text="₹ Lacs", row=1, col=2)
    fig.update_yaxes(title_text="₹ Lacs", row=2, col=1)
    fig.update_yaxes(title_text="₹ Lacs", row=2, col=2)
    
    return fig

def create_sensitivity_heatmap(projections, wacc_range, g_range, num_shares, net_debt=0, terminal_fcff=None):
    """Create sensitivity analysis heatmap (fair value per share = (EV - net debt) / shares).

    net_debt: Lacs. Without it the chart showed ENTERPRISE value per share under a 'Fair Value' label.
    terminal_fcff: the terminal FCFF the DCF actually used (differs from projections['fcff'][-1] when
    FCFF recovery fired); falls back to the last projected FCFF."""
    last_fcff = terminal_fcff if terminal_fcff is not None else projections['fcff'][-1]
    n = len(projections['fcff'])
    
    # Create matrix
    matrix = []
    for w in wacc_range:
        row = []
        for g in g_range:
            if g >= w - 0.1:
                row.append(None)
            else:
                try:
                    fcff_n1 = last_fcff * (1 + g/100)
                    tv = fcff_n1 / ((w/100) - (g/100))
                    pv_tv = tv / ((1 + w/100) ** n)
                    
                    # Calculate sum_pv_fcff (approximate from first calc)
                    sum_pv_fcff = sum([projections['fcff'][i] / ((1 + w/100) ** (i+1)) for i in range(len(projections['fcff']))])
                    
                    ev = sum_pv_fcff + pv_tv
                    eq_val = (ev - net_debt) * 100000
                    fv = eq_val / num_shares if num_shares > 0 else 0
                    row.append(fv)
                except:
                    row.append(None)
        matrix.append(row)
    
    fig = go.Figure(data=go.Heatmap(
        z=matrix,
        x=[f"{g:.1f}%" for g in g_range],
        y=[f"{w:.1f}%" for w in wacc_range],
        colorscale='RdYlGn',
        text=[[f"₹{val:.1f}" if val else "N/A" for val in row] for row in matrix],
        texttemplate="%{text}",
        textfont={"size":10},
        colorbar=dict(title="Fair Value ₹")
    ))
    
    fig.update_layout(
        title='Sensitivity Analysis: Fair Value per Share',
        xaxis_title='Terminal Growth Rate (g)',
        yaxis_title='WACC',
        height=600
    )
    
    return fig

def create_historical_financials_chart(financials, reverse_years=False):
    """
    Create comprehensive historical financials overview
    
    Args:
        financials: Financial data dict
        reverse_years: If True, reverse the years for chronological display (used in Screener mode)
    """
    years = financials['years']
    
    # Reverse years and all data arrays for chronological display if needed
    if reverse_years:
        years = list(reversed(years))
        # Create reversed version of financials dict for plotting
        financials_plot = {}
        for key, value in financials.items():
            if isinstance(value, list) and key != 'years':
                financials_plot[key] = list(reversed(value))
            else:
                financials_plot[key] = value
        financials_plot['years'] = years
    else:
        financials_plot = financials
    
    fig = make_subplots(
        rows=3, cols=2,
        subplot_titles=('Revenue & EBITDA Trend', 'Profitability Margins', 
                       'Balance Sheet Health', 'Cash Flow Quality',
                       'Working Capital Efficiency', 'Leverage Ratios'),
        specs=[[{"secondary_y": True}, {"secondary_y": False}],
               [{"secondary_y": True}, {"secondary_y": False}],
               [{"secondary_y": False}, {"secondary_y": False}]]
    )
    
    # Revenue & EBITDA
    fig.add_trace(
        go.Bar(name='Revenue', x=years, y=financials_plot['revenue'], marker_color='#06A77D'),
        row=1, col=1, secondary_y=False
    )
    fig.add_trace(
        go.Scatter(name='EBITDA Margin %', x=years, 
                  y=[(financials_plot['ebitda'][i]/financials_plot['revenue'][i]*100) if financials_plot['revenue'][i] > 0 else 0 
                     for i in range(len(years))],
                  mode='lines+markers', line=dict(color='#F77F00', width=3), marker=dict(size=10)),
        row=1, col=1, secondary_y=True
    )
    
    # Profitability Margins
    ebitda_margins = [(financials_plot['ebitda'][i]/financials_plot['revenue'][i]*100) if financials_plot['revenue'][i] > 0 else 0 
                      for i in range(len(years))]
    ebit_margins = [(financials_plot['ebit'][i]/financials_plot['revenue'][i]*100) if financials_plot['revenue'][i] > 0 else 0 
                    for i in range(len(years))]
    
    fig.add_trace(
        go.Scatter(name='EBITDA Margin', x=years, y=ebitda_margins, mode='lines+markers', line=dict(width=3)),
        row=1, col=2
    )
    fig.add_trace(
        go.Scatter(name='EBIT Margin', x=years, y=ebit_margins, mode='lines+markers', line=dict(width=3)),
        row=1, col=2
    )
    
    # Balance Sheet
    fig.add_trace(
        go.Bar(name='Equity', x=years, y=financials_plot['equity'], marker_color='#06A77D'),
        row=2, col=1, secondary_y=False
    )
    total_debt = [financials_plot['st_debt'][i] + financials_plot['lt_debt'][i] for i in range(len(years))]
    fig.add_trace(
        go.Bar(name='Debt', x=years, y=total_debt, marker_color='#D62828'),
        row=2, col=1, secondary_y=False
    )
    fig.add_trace(
        go.Scatter(name='Debt/Equity', x=years,
                  y=[(total_debt[i]/financials_plot['equity'][i]) if financials_plot['equity'][i] > 0 else 0 
                     for i in range(len(years))],
                  mode='lines+markers', line=dict(color='#2E86AB', width=3)),
        row=2, col=1, secondary_y=True
    )
    
    # Cash Flow Quality (NOPAT vs EBIT)
    fig.add_trace(
        go.Bar(name='EBIT', x=years, y=financials_plot['ebit'], marker_color='#F77F00'),
        row=2, col=2
    )
    fig.add_trace(
        go.Bar(name='NOPAT', x=years, y=financials_plot['nopat'], marker_color='#06A77D'),
        row=2, col=2
    )
    
    # Working Capital Components
    fig.add_trace(
        go.Bar(name='Inventory', x=years, y=financials_plot['inventory'], marker_color='#2E86AB'),
        row=3, col=1
    )
    fig.add_trace(
        go.Bar(name='Receivables', x=years, y=financials_plot['receivables'], marker_color='#06A77D'),
        row=3, col=1
    )
    fig.add_trace(
        go.Bar(name='Payables', x=years, y=financials_plot['payables'], marker_color='#D62828'),
        row=3, col=1
    )
    
    # Leverage Ratios
    debt_to_ebitda = [(total_debt[i]/financials_plot['ebitda'][i]) if financials_plot['ebitda'][i] > 0 else 0 
                      for i in range(len(years))]
    interest_coverage = [(financials_plot['ebit'][i]/financials_plot['interest'][i]) if financials_plot['interest'][i] > 0 else 0 
                         for i in range(len(years))]
    
    fig.add_trace(
        go.Scatter(name='Debt/EBITDA', x=years, y=debt_to_ebitda, mode='lines+markers', line=dict(width=3)),
        row=3, col=2
    )
    fig.add_trace(
        go.Scatter(name='Interest Coverage', x=years, y=interest_coverage, mode='lines+markers', line=dict(width=3)),
        row=3, col=2
    )
    
    fig.update_layout(height=1200, showlegend=True, title_text="Historical Financial Analysis Dashboard")
    fig.update_yaxes(title_text="₹ Lacs", row=1, col=1, secondary_y=False)
    fig.update_yaxes(title_text="Margin %", row=1, col=1, secondary_y=True)
    fig.update_yaxes(title_text="Margin %", row=1, col=2)
    fig.update_yaxes(title_text="₹ Lacs", row=2, col=1, secondary_y=False)
    fig.update_yaxes(title_text="Ratio", row=2, col=1, secondary_y=True)
    fig.update_yaxes(title_text="₹ Lacs", row=2, col=2)
    fig.update_yaxes(title_text="₹ Lacs", row=3, col=1)
    fig.update_yaxes(title_text="Ratio", row=3, col=2)
    
    return fig

def create_wacc_breakdown_chart(wacc_details):
    """Create visual breakdown of WACC components"""
    labels = ['Cost of Equity (Ke)', 'After-tax Cost of Debt (Kd)']
    values = [wacc_details['ke'], wacc_details['kd_after_tax']]
    weights = [wacc_details['we'], wacc_details['wd']]
    contributions = [wacc_details['ke'] * wacc_details['we'] / 100, 
                    wacc_details['kd_after_tax'] * wacc_details['wd'] / 100]
    
    fig = make_subplots(
        rows=1, cols=2,
        specs=[[{"type": "pie"}, {"type": "bar"}]],
        subplot_titles=('Capital Structure Weights', 'WACC Components Contribution')
    )
    
    # Capital structure pie
    fig.add_trace(
        go.Pie(labels=['Equity', 'Debt'], values=[wacc_details['we'], wacc_details['wd']],
               marker_colors=['#06A77D', '#D62828']),
        row=1, col=1
    )
    
    # WACC contribution bar
    fig.add_trace(
        go.Bar(name='Contribution to WACC', x=labels, y=contributions,
               marker_color=['#06A77D', '#D62828'],
               text=[f"{c:.2f}%" for c in contributions],
               textposition='auto'),
        row=1, col=2
    )
    
    fig.update_layout(height=400, showlegend=True, title_text=f"WACC Breakdown (Total: {wacc_details['wacc']:.2f}%)")
    
    return fig

def create_bank_valuation_comparison_chart(valuations_dict):
    """Create comparison chart for multiple bank valuation methods"""
    methods = []
    values = []
    colors = []
    
    color_map = {
        'Residual Income Model': '#2E86AB',
        'Dividend Discount Model': '#06A77D',
        'P/B with ROE Analysis': '#F77F00',
        'Relative Valuation (P/E)': '#D62828',
        'Relative Valuation (P/B)': '#9D4EDD'
    }
    
    for method, val_data in valuations_dict.items():
        if val_data and 'value_per_share' in val_data:
            methods.append(method)
            values.append(val_data['value_per_share'])
            colors.append(color_map.get(method, '#888888'))
    
    fig = go.Figure()
    
    fig.add_trace(go.Bar(
        x=methods,
        y=values,
        marker_color=colors,
        text=[f"₹{v:.2f}" for v in values],
        textposition='auto',
    ))
    
    if values:
        avg_value = np.mean(values)
        fig.add_hline(y=avg_value, line_dash="dash", line_color="red",
                     annotation_text=f"Average: ₹{avg_value:.2f}",
                     annotation_position="right")
    
    fig.update_layout(
        title="Bank Valuation Methods Comparison",
        xaxis_title="Valuation Method",
        yaxis_title="Fair Value per Share (₹)",
        height=500,
        showlegend=False
    )
    
    return fig

def create_price_vs_value_gauge(current_price, fair_value):
    """Create gauge chart showing current price vs fair value"""
    # Check for invalid fair value (zero, negative, or unrealistic)
    if fair_value <= 0:
        return None
    
    ratio = (current_price / fair_value) * 100
    
    fig = go.Figure(go.Indicator(
        mode = "gauge+number+delta",
        value = current_price,
        domain = {'x': [0, 1], 'y': [0, 1]},
        title = {'text': f"Current Price vs Fair Value (₹{fair_value:.2f})", 'font': {'size': 20}},
        delta = {'reference': fair_value, 'valueformat': '.2f'},
        gauge = {
            'axis': {'range': [None, max(current_price, fair_value) * 1.5], 'tickformat': '₹.2f'},
            'bar': {'color': "#2E86AB"},
            'bgcolor': "white",
            'borderwidth': 2,
            'bordercolor': "gray",
            'steps': [
                {'range': [0, fair_value * 0.8], 'color': '#06A77D'},
                {'range': [fair_value * 0.8, fair_value * 1.2], 'color': '#F4D35E'},
                {'range': [fair_value * 1.2, fair_value * 2], 'color': '#D62828'}],
            'threshold': {
                'line': {'color': "red", 'width': 4},
                'thickness': 0.75,
                'value': fair_value}}))
    
    fig.update_layout(height=400)
    
    if ratio < 80:
        recommendation = "🟢 UNDERVALUED - Potential Buy"
    elif ratio > 120:
        recommendation = "🔴 OVERVALUED - Potential Sell"
    else:
        recommendation = "🟡 FAIRLY VALUED - Hold"
    
    fig.add_annotation(
        text=recommendation,
        xref="paper", yref="paper",
        x=0.5, y=-0.1,
        showarrow=False,
        font=dict(size=16, color="black", family="Arial Black")
    )
    
    return fig

# ================================
# DCF CALCULATION FUNCTIONS
# ================================

def extract_financials_unlisted(df_bs, df_pl, year_cols):
    """Extract financial metrics from Excel DataFrames"""
    num_years = min(3, len(year_cols))
    # year_cols is sorted oldest->newest. Take the most recent N years,
    # then REVERSE to newest-first so financials['...'][0] = latest year,
    # matching the convention used everywhere else (Yahoo/Screener data,
    # WACC, DCF projections, comp valuation, etc.)
    last_years = list(reversed(year_cols[-num_years:]))
    
    financials = {
        'years': last_years,
        'revenue': [],
        'cogs': [],
        'opex': [],
        'ebitda': [],
        'depreciation': [],
        'ebit': [],
        'interest': [],
        'interest_income': [],  # Added for business classification
        'tax': [],
        'nopat': [],
        'fixed_assets': [],
        'inventory': [],
        'receivables': [],
        'payables': [],
        'cash': [],
        'equity': [],
        'st_debt': [],
        'lt_debt': [],
    }
    
    for year_col in last_years:
        # Income Statement
        revenue = get_value_from_df(df_pl, 'Net Revenue', year_col)
        # COGS: Use 'Cost of Materials' for manufacturers; fall back to
        # Purchases of Stock-in-trade +/- Changes in Inventories for trading
        # companies (e.g. Zepto, Blinkit) where Cost of Materials = 0.
        cogs_materials = get_value_from_df(df_pl, 'Cost of Materials', year_col)
        purchases_sit  = get_value_from_df(df_pl, 'Purchases of Stock-in-trade', year_col)
        changes_inv    = get_value_from_df(df_pl, 'Changes in Inventories', year_col)
        if cogs_materials > 0:
            cogs = cogs_materials
        elif purchases_sit > 0:
            # Trading COGS = Purchases ± Changes in Inventories
            # "Changes in Inventories" row is typically positive when stock decreases
            cogs = purchases_sit + changes_inv
        else:
            cogs = 0.0
        employee_exp = get_value_from_df(df_pl, 'Employee Benefit', year_col)
        other_exp = get_value_from_df(df_pl, 'Other Expenses', year_col)
        depreciation = get_value_from_df(df_pl, 'Depreciation', year_col)
        interest = get_value_from_df(df_pl, 'Finance Costs', year_col)
        interest_income = get_value_from_df(df_pl, 'Finance Income', year_col)  # For classification
        tax = get_value_from_df(df_pl, 'Income Tax', year_col)
        
        opex = employee_exp + other_exp
        ebitda = revenue - opex - cogs
        ebit = ebitda - depreciation
        pbt = ebit - interest
        pat = pbt - tax
        nopat = ebit * (1 - 0.25)  # Assuming 25% tax
        
        financials['revenue'].append(revenue)
        financials['cogs'].append(cogs)
        financials['opex'].append(opex)
        financials['ebitda'].append(ebitda)
        financials['depreciation'].append(depreciation)
        financials['ebit'].append(ebit)
        financials['interest'].append(interest)
        financials['interest_income'].append(interest_income)  # Store for classification
        financials['tax'].append(tax)
        financials['nopat'].append(nopat)
        
        # Balance Sheet
        fixed_assets = get_value_from_df(df_bs, 'Tangible Assets', year_col)
        inventory = get_value_from_df(df_bs, 'Inventories', year_col)
        receivables = get_value_from_df(df_bs, 'Trade Receivables', year_col)
        payables = get_value_from_df(df_bs, 'Trade Payables', year_col)
        cash = get_value_from_df(df_bs, 'Cash and Bank', year_col)
        equity = get_value_from_df(df_bs, 'Total Equity', year_col)
        st_debt = get_value_from_df(df_bs, 'Short Term Borrowings', year_col)
        lt_debt = get_value_from_df(df_bs, 'Long Term Borrowings', year_col)
        
        financials['fixed_assets'].append(fixed_assets)
        financials['inventory'].append(inventory)
        financials['receivables'].append(receivables)
        financials['payables'].append(payables)
        financials['cash'].append(cash)
        financials['equity'].append(equity)
        financials['st_debt'].append(st_debt)
        financials['lt_debt'].append(lt_debt)
    
    return financials

def fetch_screener_peer_data(ticker_symbol):
    """
    Fetch peer company data from Screener.in for comparative valuation
    
    Args:
        ticker_symbol: NSE/BSE ticker (without .NS/.BO suffix)
    
    Returns:
        dict with peer financial data or None
    """
    try:
        import time
        import random
        
        # Clean ticker
        ticker_clean = ticker_symbol.replace('.NS', '').replace('.BO', '')
        
        # Add delay for respectful scraping
        time.sleep(random.uniform(1.5, 3.0))
        
        screener_data = fetch_screener_financials(ticker_clean, num_years=3)
        
        if not screener_data or not screener_data.get('financials'):
            return None
        
        financials = screener_data['financials']
        shares = screener_data.get('shares', 0)
        
        # Convert to format expected by perform_comparative_valuation
        # Values are already in Lacs, need to convert to rupees for compatibility
        revenue = ensure_valid_number(financials['revenue'][0] * 100000, 0)  # Most recent year
        ebitda = ensure_valid_number(financials['ebitda'][0] * 100000, 0)
        net_income = ensure_valid_number(financials['nopat'][0] * 100000, 0)
        total_debt = ensure_valid_number((financials['st_debt'][0] + financials['lt_debt'][0]) * 100000, 0)
        cash = ensure_valid_number(financials['cash'][0] * 100000, 0)
        equity = ensure_valid_number(financials['equity'][0] * 100000, 0)
        
        eps = safe_divide(net_income, shares, default=0)
        book_value = safe_divide(equity, shares, default=0)
        
        return {
            'ticker': ticker_clean,
            'name': screener_data.get('company_name', ticker_clean),
            'price': 0,  # Screener doesn't provide live price
            'shares': shares,
            'market_cap': 0,  # Cannot calculate without price
            'revenue': revenue,
            'ebitda': ebitda,
            'net_income': net_income,
            'eps': eps,
            'book_value': book_value,
            'total_debt': total_debt,
            'cash': cash,
            'enterprise_value': 0,  # Cannot calculate without market cap
            '_source': 'screener'
        }
    
    except Exception as e:
        print(f"Error fetching Screener data for {ticker_symbol}: {e}")
        return None


def perform_comparative_valuation(target_ticker, comp_tickers_str, target_financials=None, target_shares=None, exchange_suffix="NS", projections=None, use_screener_peers=False, valuation_date=None):
    """
    Perform comparative valuation using peer multiples
    
    Args:
        target_ticker: Target company ticker
        comp_tickers_str: Comma-separated peer tickers
        target_financials: Target company financials dict
        target_shares: Target company shares outstanding
        exchange_suffix: NS or BO
        projections: DCF projections dict with 'nopat' key
        use_screener_peers: If True, fetch peer data from Screener.in instead of Yahoo Finance
        valuation_date: If provided (and not today), prices for the target and
                        every Yahoo Finance peer are pulled as of this date
                        instead of the live/current price, so multiples reflect
                        market conditions on the valuation date rather than
                        the day the app happens to be run.
    """
    _use_hist_price = not is_valuation_date_today(valuation_date)
    try:
        comp_tickers = [t.strip() for t in comp_tickers_str.split(',') if t.strip()]
        
        if not comp_tickers:
            return None
        
        results = {
            'target': {},
            'comparables': [],
            'multiples_stats': {},
            'valuations': {},
            '_peer_source': 'screener' if use_screener_peers else 'yahoo',
            '_currency_symbol': '₹'  # default; updated below from target info
        }
        
        # Get target company data
        if target_ticker:
            # Listed company
            target_stock = get_cached_ticker(get_ticker_with_exchange(target_ticker, exchange_suffix))
            target_info = target_stock.info if target_stock else None
            target_financials_yf = target_stock.financials
            target_bs = target_stock.balance_sheet
            
            if not target_info:
                st.error(f"Could not fetch data for {target_ticker}")
                return results
            
            # Detect currency from target ticker
            results['_currency_symbol'] = get_currency_symbol(target_info)
            _csym = results['_currency_symbol']  # shorthand for formula strings below

            # Robust price fetching — live by default, or historical as-of the
            # valuation date if one was supplied.
            _live_price = target_info.get('currentPrice', 0) or target_info.get('regularMarketPrice', 0) or 0
            _target_price = _live_price
            if _use_hist_price:
                _hist_price = get_price_as_of(target_stock, valuation_date)
                if _hist_price:
                    _target_price = _hist_price
                else:
                    st.warning(f"⚠️ No {target_ticker} price found on/before {pd.Timestamp(valuation_date).date()} — using live price instead.")

            _target_shares_outstanding = target_info.get('sharesOutstanding', 0)
            _target_total_debt = safe_extract(target_bs, 'Long Term Debt', target_bs.columns[0]) if 'Long Term Debt' in target_bs.index else 0
            _target_cash = safe_extract(target_bs, 'Cash And Cash Equivalents', target_bs.columns[0]) if 'Cash And Cash Equivalents' in target_bs.index else 0

            # Market cap / EV need to move with the historical price too —
            # info['marketCap'] / info['enterpriseValue'] are always "live".
            if _use_hist_price and _target_shares_outstanding:
                _target_market_cap = _target_price * _target_shares_outstanding
                _target_ev = _target_market_cap + _target_total_debt - _target_cash
            else:
                _target_market_cap = target_info.get('marketCap', 0)
                _target_ev = target_info.get('enterpriseValue', 0)

            results['target'] = {
                'name': target_info.get('longName', target_ticker),
                'current_price': _target_price,
                'shares': _target_shares_outstanding,
                'market_cap': _target_market_cap,
                'enterprise_value': _target_ev,
                'revenue': safe_extract(target_financials_yf, 'Total Revenue', target_financials_yf.columns[0]) if 'Total Revenue' in target_financials_yf.index else 0,
                'ebitda': target_info.get('ebitda', 0),
                'net_income': safe_extract(target_financials_yf, 'Net Income', target_financials_yf.columns[0]) if 'Net Income' in target_financials_yf.index else 0,
                'book_value_per_share': target_info.get('bookValue', 0),
                'total_debt': _target_total_debt,
                'cash': _target_cash,
            }
            
            # Calculate EPS and Book Value - ALWAYS set to avoid KeyError
            if results['target']['shares'] > 0 and results['target']['net_income'] != 0:
                results['target']['eps'] = results['target']['net_income'] / results['target']['shares']
            else:
                results['target']['eps'] = 0
                
        else:
            # Unlisted company - arrays have NEWEST first, so [0] = latest
            equity_lacs = target_financials['equity'][0]  # Latest equity in Lacs
            equity_rupees = equity_lacs * 100000  # Convert to Rupees
            book_value_per_share = equity_rupees / target_shares if target_shares > 0 else 0
            
            results['target'] = {
                'name': 'Target Company (Unlisted)',
                'current_price': 0,
                'shares': target_shares,
                'market_cap': 0,
                'enterprise_value': 0,
                'revenue': target_financials['revenue'][0] * 100000,  # [0] = latest year
                'ebitda': target_financials['ebitda'][0] * 100000,
                'net_income': target_financials['nopat'][0] * 100000,  # Using NOPAT as proxy
                'book_value_per_share': book_value_per_share,  # CALCULATE from equity
                'total_debt': (target_financials['st_debt'][0] + target_financials['lt_debt'][0]) * 100000,
                'cash': target_financials['cash'][0] * 100000,
                'eps': (target_financials['nopat'][0] * 100000) / target_shares if target_shares > 0 else 0,
            }
        
        # Get comparable companies data
        comp_data = []
        
        if use_screener_peers:
            # FETCH PEER DATA FROM SCREENER.IN
            st.info(f"🌐 Fetching peer data from Screener.in for {len(comp_tickers)} companies...")
            
            for idx, ticker in enumerate(comp_tickers):
                try:
                    peer_data = fetch_screener_peer_data(ticker)
                    
                    if peer_data:
                        # Calculate multiples (most will be 0 since no price/market cap)
                        pe = 0  # Cannot calculate without price
                        pb = 0  # Cannot calculate without price
                        ps = 0  # Cannot calculate without market cap
                        
                        # Can calculate EV-based multiples if we have enterprise value
                        enterprise_value = peer_data['enterprise_value']
                        ev_ebitda = safe_divide(enterprise_value, peer_data['ebitda'], default=0)
                        ev_sales = safe_divide(enterprise_value, peer_data['revenue'], default=0)
                        
                        comp_data.append({
                            'ticker': peer_data['ticker'],
                            'name': peer_data['name'],
                            'price': peer_data['price'],
                            'market_cap': peer_data['market_cap'],
                            'revenue': peer_data['revenue'],
                            'ebitda': peer_data['ebitda'],
                            'net_income': peer_data['net_income'],
                            'eps': peer_data['eps'],
                            'book_value': peer_data['book_value'],
                            'pe': pe,
                            'pb': pb,
                            'ps': ps,
                            'ev_ebitda': ev_ebitda,
                            'ev_sales': ev_sales,
                            'enterprise_value': enterprise_value,
                            'shares': peer_data['shares']
                        })
                        
                        st.success(f"✅ {peer_data['name']}")
                    else:
                        st.warning(f"⚠️ Could not fetch data for {ticker}")
                
                except Exception as e:
                    st.warning(f"⚠️ Error fetching {ticker}: {str(e)}")
                    continue
            
            if not comp_data:
                st.warning("⚠️ Could not fetch any peer data from Screener.in. Try Yahoo Finance instead.")
                return None
        
        else:
            # FETCH PEER DATA FROM YAHOO FINANCE (ORIGINAL CODE)
            for idx, ticker in enumerate(comp_tickers):
                try:
                    # Add delay between requests to avoid rate limiting
                    if idx > 0:
                        time.sleep(random.uniform(1.0, 1.5))
                    
                    # Ticker already has suffix (.NS or .BO) from UI combination
                    comp_stock = get_cached_ticker(ticker)
                    comp_info = comp_stock.info if comp_stock else None
                    
                    if not comp_info:
                        st.warning(f"Could not fetch data for {ticker}")
                        continue
                    
                    comp_financials_yf = comp_stock.financials
                    comp_bs = comp_stock.balance_sheet
                    
                    # Extract financial data
                    shares = comp_info.get('sharesOutstanding', 0)
                    # Robust price fetching - try multiple methods
                    price = comp_info.get('currentPrice', 0)
                    if not price or price == 0:
                        price = comp_info.get('regularMarketPrice', 0)
                    if not price or price == 0:
                        try:
                            hist = comp_stock.history(period='1d')
                            if not hist.empty:
                                price = hist['Close'].iloc[-1]
                        except:
                            pass
                    market_cap = comp_info.get('marketCap', 0)

                    if _use_hist_price:
                        _peer_hist_price = get_price_as_of(comp_stock, valuation_date)
                        if _peer_hist_price:
                            price = _peer_hist_price
                            if shares:
                                market_cap = price * shares  # keep market cap consistent with the historical price
                        else:
                            st.warning(f"⚠️ No {ticker} price found on/before {pd.Timestamp(valuation_date).date()} — using live price for this peer instead.")
                    
                    revenue = safe_extract(comp_financials_yf, 'Total Revenue', comp_financials_yf.columns[0]) if 'Total Revenue' in comp_financials_yf.index and not comp_financials_yf.empty else 0
                    ebitda = comp_info.get('ebitda', 0)
                    net_income = safe_extract(comp_financials_yf, 'Net Income', comp_financials_yf.columns[0]) if 'Net Income' in comp_financials_yf.index and not comp_financials_yf.empty else 0
                    
                    total_debt = safe_extract(comp_bs, 'Long Term Debt', comp_bs.columns[0]) if 'Long Term Debt' in comp_bs.index and not comp_bs.empty else 0
                    cash = safe_extract(comp_bs, 'Cash And Cash Equivalents', comp_bs.columns[0]) if 'Cash And Cash Equivalents' in comp_bs.index and not comp_bs.empty else 0
                    
                    book_value = comp_info.get('bookValue', 0)
                    eps = net_income / shares if shares > 0 else 0
                    
                    # Calculate multiples
                    pe = price / eps if eps > 0 else 0
                    pb = price / book_value if book_value > 0 else 0
                    ps = market_cap / revenue if revenue > 0 else 0
                    
                    enterprise_value = market_cap + total_debt - cash
                    ev_ebitda = enterprise_value / ebitda if ebitda > 0 else 0
                    ev_sales = enterprise_value / revenue if revenue > 0 else 0
                    
                    comp_data.append({
                        'ticker': ticker,
                        'name': comp_info.get('longName', ticker),
                        'price': price,
                        'market_cap': market_cap,
                        'revenue': revenue,
                        'ebitda': ebitda,
                        'net_income': net_income,
                        'eps': eps,
                        'book_value': book_value,
                        'pe': pe,
                        'pb': pb,
                        'ps': ps,
                        'ev_ebitda': ev_ebitda,
                        'ev_sales': ev_sales,
                        'enterprise_value': enterprise_value,
                        'shares': shares
                    })
                    
                except Exception as e:
                    st.warning(f"Could not fetch data for {ticker}: {str(e)}")
                    continue
        
        results['comparables'] = comp_data
        
        if not comp_data:
            st.error("No comparable company data could be fetched")
            return None
        
        # Calculate statistics for each multiple
        multiples = ['pe', 'pb', 'ps', 'ev_ebitda', 'ev_sales']
        
        for multiple in multiples:
            valid_values = [c[multiple] for c in comp_data if c.get(multiple, 0) > 0]
            
            if not valid_values:
                continue
            
            results['multiples_stats'][multiple] = {
                'average': np.mean(valid_values),
                'median': np.median(valid_values),
                'min': np.min(valid_values),
                'max': np.max(valid_values),
                'std': np.std(valid_values),
                'values': valid_values
            }
        
        # Calculate implied valuations
        target = results['target']
        valuations_summary = {}
        _csym = results.get('_currency_symbol', '₹')  # ensure _csym is always defined
        
        # P/E Method
        if 'pe' in results['multiples_stats'] and target['eps'] > 0:
            stats = results['multiples_stats']['pe']
            
            fair_value_avg = target['eps'] * stats['average']
            fair_value_median = target['eps'] * stats['median']
            
            valuations_summary['pe'] = {
                'method': 'Price-to-Earnings (P/E)',
                'target_metric': target['eps'],
                'metric_name': 'EPS',
                'avg_multiple': stats['average'],
                'median_multiple': stats['median'],
                'fair_value_avg': fair_value_avg,
                'fair_value_median': fair_value_median,
                'current_price': target['current_price'],
                'upside_avg': ((fair_value_avg - target['current_price']) / target['current_price'] * 100) if target['current_price'] else 0,
                'upside_median': ((fair_value_median - target['current_price']) / target['current_price'] * 100) if target['current_price'] else 0,
                'formula_avg': f"EPS × Avg P/E = {_csym}{target['eps']:.2f} × {stats['average']:.2f} = {_csym}{fair_value_avg:.2f}",
                'formula_median': f"EPS × Median P/E = {_csym}{target['eps']:.2f} × {stats['median']:.2f} = {_csym}{fair_value_median:.2f}"
            }
        
        # P/B Method
        if 'pb' in results['multiples_stats'] and target['book_value_per_share'] > 0:
            stats = results['multiples_stats']['pb']
            
            fair_value_avg = target['book_value_per_share'] * stats['average']
            fair_value_median = target['book_value_per_share'] * stats['median']
            
            valuations_summary['pb'] = {
                'method': 'Price-to-Book (P/B)',
                'target_metric': target['book_value_per_share'],
                'metric_name': 'Book Value per Share',
                'avg_multiple': stats['average'],
                'median_multiple': stats['median'],
                'fair_value_avg': fair_value_avg,
                'fair_value_median': fair_value_median,
                'current_price': target['current_price'],
                'upside_avg': ((fair_value_avg - target['current_price']) / target['current_price'] * 100) if target['current_price'] else 0,
                'upside_median': ((fair_value_median - target['current_price']) / target['current_price'] * 100) if target['current_price'] else 0,
                'formula_avg': f"BVPS × Avg P/B = {_csym}{target['book_value_per_share']:.2f} × {stats['average']:.2f} = {_csym}{fair_value_avg:.2f}",
                'formula_median': f"BVPS × Median P/B = {_csym}{target['book_value_per_share']:.2f} × {stats['median']:.2f} = {_csym}{fair_value_median:.2f}"
            }
        
        # P/S Method
        if 'ps' in results['multiples_stats'] and target['revenue'] > 0 and target['shares'] > 0:
            stats = results['multiples_stats']['ps']
            
            revenue_per_share = target['revenue'] / target['shares']
            fair_value_avg = revenue_per_share * stats['average']
            fair_value_median = revenue_per_share * stats['median']
            
            valuations_summary['ps'] = {
                'method': 'Price-to-Sales (P/S)',
                'target_metric': revenue_per_share,
                'metric_name': 'Revenue per Share',
                'avg_multiple': stats['average'],
                'median_multiple': stats['median'],
                'fair_value_avg': fair_value_avg,
                'fair_value_median': fair_value_median,
                'current_price': target['current_price'],
                'upside_avg': ((fair_value_avg - target['current_price']) / target['current_price'] * 100) if target['current_price'] else 0,
                'upside_median': ((fair_value_median - target['current_price']) / target['current_price'] * 100) if target['current_price'] else 0,
                'formula_avg': f"Revenue/Share × Avg P/S = {_csym}{revenue_per_share:.2f} × {stats['average']:.2f} = {_csym}{fair_value_avg:.2f}",
                'formula_median': f"Revenue/Share × Median P/S = {_csym}{revenue_per_share:.2f} × {stats['median']:.2f} = {_csym}{fair_value_median:.2f}"
            }
        
        # EV/EBITDA Method
        if 'ev_ebitda' in results['multiples_stats'] and target['ebitda'] > 0 and target['shares'] > 0:
            stats = results['multiples_stats']['ev_ebitda']
            
            implied_ev_avg = target['ebitda'] * stats['average']
            implied_ev_median = target['ebitda'] * stats['median']
            
            net_debt = target['total_debt'] - target['cash']
            
            equity_value_avg = implied_ev_avg - net_debt
            equity_value_median = implied_ev_median - net_debt
            
            fair_value_avg = equity_value_avg / target['shares']
            fair_value_median = equity_value_median / target['shares']
            
            valuations_summary['ev_ebitda'] = {
                'method': 'EV/EBITDA',
                'target_metric': target['ebitda'],
                'metric_name': 'EBITDA',
                'avg_multiple': stats['average'],
                'median_multiple': stats['median'],
                'implied_ev_avg': implied_ev_avg,
                'implied_ev_median': implied_ev_median,
                'total_debt': total_debt,
        'cash': cash,
        'net_debt': net_debt,
                'fair_value_avg': fair_value_avg,
                'fair_value_median': fair_value_median,
                'current_price': target['current_price'],
                'upside_avg': ((fair_value_avg - target['current_price']) / target['current_price'] * 100) if target['current_price'] else 0,
                'upside_median': ((fair_value_median - target['current_price']) / target['current_price'] * 100) if target['current_price'] else 0,
                'formula_avg': f"(EBITDA × Avg EV/EBITDA - Net Debt) / Shares = ({_csym}{target['ebitda']/1e7:.2f}{'Cr' if _csym=='₹' else 'M'} × {stats['average']:.2f} - {_csym}{net_debt/1e7:.2f}{'Cr' if _csym=='₹' else 'M'}) / {target['shares']/1e7:.2f}{'Cr' if _csym=='₹' else 'M'}",
                'formula_median': f"(EBITDA × Median EV/EBITDA - Net Debt) / Shares = ({_csym}{target['ebitda']/1e7:.2f}{'Cr' if _csym=='₹' else 'M'} × {stats['median']:.2f} - {_csym}{net_debt/1e7:.2f}{'Cr' if _csym=='₹' else 'M'}) / {target['shares']/1e7:.2f}{'Cr' if _csym=='₹' else 'M'}"
            }
        
        # EV/Sales Method
        if 'ev_sales' in results['multiples_stats'] and target['revenue'] > 0 and target['shares'] > 0:
            stats = results['multiples_stats']['ev_sales']
            
            implied_ev_avg = target['revenue'] * stats['average']
            implied_ev_median = target['revenue'] * stats['median']
            
            net_debt = target['total_debt'] - target['cash']
            
            equity_value_avg = implied_ev_avg - net_debt
            equity_value_median = implied_ev_median - net_debt
            
            fair_value_avg = equity_value_avg / target['shares']
            fair_value_median = equity_value_median / target['shares']
            
            valuations_summary['ev_sales'] = {
                'method': 'EV/Sales',
                'target_metric': target['revenue'],
                'metric_name': 'Revenue',
                'avg_multiple': stats['average'],
                'median_multiple': stats['median'],
                'implied_ev_avg': implied_ev_avg,
                'implied_ev_median': implied_ev_median,
                'net_debt': net_debt,
                'fair_value_avg': fair_value_avg,
                'fair_value_median': fair_value_median,
                'current_price': target['current_price'],
                'upside_avg': ((fair_value_avg - target['current_price']) / target['current_price'] * 100) if target['current_price'] else 0,
                'upside_median': ((fair_value_median - target['current_price']) / target['current_price'] * 100) if target['current_price'] else 0,
                'formula_avg': f"(Revenue × Avg EV/Sales - Net Debt) / Shares = ({_csym}{target['revenue']/1e7:.2f}{'Cr' if _csym=='₹' else 'M'} × {stats['average']:.2f} - {_csym}{net_debt/1e7:.2f}{'Cr' if _csym=='₹' else 'M'}) / {target['shares']/1e7:.2f}{'Cr' if _csym=='₹' else 'M'}",
                'formula_median': f"(Revenue × Median EV/Sales - Net Debt) / Shares = ({_csym}{target['revenue']/1e7:.2f}{'Cr' if _csym=='₹' else 'M'} × {stats['median']:.2f} - {_csym}{net_debt/1e7:.2f}{'Cr' if _csym=='₹' else 'M'}) / {target['shares']/1e7:.2f}{'Cr' if _csym=='₹' else 'M'}"
            }
        
        results['valuations'] = valuations_summary
        
        # Calculate 12-Month Forward P/E using EXISTING PROJECTED NOPAT (NO DUPLICATION!)
        # Forward P/E = Current Price / Future EPS
        # Future EPS = Projected Year 1 NOPAT / Current Outstanding Shares
        
        # Get current EPS for comparison
        current_eps = 0
        if target_financials and 'nopat' in target_financials:
            current_nopat = target_financials['nopat'][0] * 100000  # Most recent historical year
            current_eps = current_nopat / target['shares'] if target['shares'] > 0 else 0
        
        # Use EXISTING projections if available, otherwise calculate growth
        if projections and 'nopat' in projections and len(projections['nopat']) > 0:
            # ✅ USE EXISTING PROJECTED NOPAT - NO DUPLICATION!
            projected_nopat_year1 = projections['nopat'][0] * 100000  # Year 1 projection (already calculated in DCF!)
            future_eps = projected_nopat_year1 / target['shares'] if target['shares'] > 0 else 0
            
            # Calculate growth rate for display purposes only
            if current_eps > 0 and future_eps > 0:
                growth_rate = ((future_eps / current_eps) - 1) * 100
            else:
                growth_rate = 0.0
            
            calculation_method = "Using DCF Year 1 Projected NOPAT"
            
        elif target_financials and 'nopat' in target_financials and len(target_financials['nopat']) >= 2:
            # Fallback: Calculate growth if projections not available
            recent_nopat = target_financials['nopat'][:min(3, len(target_financials['nopat']))]
            if len(recent_nopat) >= 2 and recent_nopat[-1] > 0:
                num_years = len(recent_nopat) - 1
                growth_rate = ((recent_nopat[0] / recent_nopat[-1]) ** (1/num_years) - 1) * 100
                growth_rate = max(-5, min(growth_rate, 25))
            else:
                growth_rate = 10.0
            
            future_eps = current_eps * (1 + growth_rate / 100)
            calculation_method = f"Using historical CAGR ({growth_rate:.1f}%)"
        else:
            future_eps = current_eps * 1.10  # Default 10% growth
            growth_rate = 10.0
            calculation_method = "Using default 10% growth"
        
        if 'pe' in results['multiples_stats'] and future_eps > 0:
            stats = results['multiples_stats']['pe']
            forward_fair_value_avg = future_eps * stats['average']
            forward_fair_value_median = future_eps * stats['median']
            
            results['forward_pe'] = {
                'current_eps': current_eps,
                'forward_eps': future_eps,
                'earnings_growth_rate': growth_rate,
                'fair_value_avg': forward_fair_value_avg,
                'fair_value_median': forward_fair_value_median,
                'formula_avg': f"Forward EPS × Avg P/E = {_csym}{future_eps:.2f} × {stats['average']:.2f} = {_csym}{forward_fair_value_avg:.2f}",
                'formula_median': f"Forward EPS × Median P/E = {_csym}{future_eps:.2f} × {stats['median']:.2f} = {_csym}{forward_fair_value_median:.2f}",
                'calculation_method': calculation_method,
                'calculation_note': f"12-Month Forward EPS: {calculation_method}",
                'using_dcf_projections': bool(projections)
            }
        
        return results
        
    except Exception as e:
        st.error(f"Comparative valuation error: {str(e)}")
        import traceback
        st.error(traceback.format_exc())
        return None

def calculate_working_capital_metrics(financials):
    """
    Calculate working capital days with ROBUST None/zero handling
    
    CRITICAL FIX: If inventory, receivables, or payables are None/zero for ALL years,
    exclude them from working capital calculations to prevent NaN cascade
    """
    wc_metrics = {
        'inventory_days': [],
        'debtor_days': [],
        'creditor_days': []
    }
    
    # ROBUST: Check if we have valid data for each WC component across ALL years
    has_valid_inventory = any(ensure_valid_number(financials['inventory'][i], 0) > 0 for i in range(len(financials['years'])))
    has_valid_receivables = any(ensure_valid_number(financials['receivables'][i], 0) > 0 for i in range(len(financials['years'])))
    has_valid_payables = any(ensure_valid_number(financials['payables'][i], 0) > 0 for i in range(len(financials['years'])))
    
    for i in range(len(financials['years'])):
        # ROBUST: Ensure all values are valid numbers
        revenue = ensure_valid_number(financials['revenue'][i], 0)
        cogs = ensure_valid_number(financials['cogs'][i], 0)
        inventory = ensure_valid_number(financials['inventory'][i], 0)
        receivables = ensure_valid_number(financials['receivables'][i], 0)
        payables = ensure_valid_number(financials['payables'][i], 0)

        # For trading companies (e.g. Zepto) COGS may be 0 because Cost of Materials = 0.
        # In that case fall back to revenue as the denominator so inventory days and
        # creditor days are still meaningful (same base as debtors days).
        cogs_or_rev = cogs if cogs > 0 else revenue

        # Calculate days only if we have valid denominators
        inv_days = safe_divide(inventory * 365, cogs_or_rev, default=0) if has_valid_inventory and cogs_or_rev > 0 else 0
        deb_days = safe_divide(receivables * 365, revenue, default=0) if has_valid_receivables and revenue > 0 else 0
        cred_days = safe_divide(payables * 365, cogs_or_rev, default=0) if has_valid_payables and cogs_or_rev > 0 else 0
        
        wc_metrics['inventory_days'].append(ensure_valid_number(inv_days, 0))
        wc_metrics['debtor_days'].append(ensure_valid_number(deb_days, 0))
        wc_metrics['creditor_days'].append(ensure_valid_number(cred_days, 0))
    
    # Average days - ROBUST: Only calculate if we have valid data
    if has_valid_inventory and wc_metrics['inventory_days']:
        valid_inv_days = [d for d in wc_metrics['inventory_days'] if d > 0]
        wc_metrics['avg_inv_days'] = np.mean(valid_inv_days) if valid_inv_days else 0
    else:
        wc_metrics['avg_inv_days'] = 0  # No inventory data available
    
    if has_valid_receivables and wc_metrics['debtor_days']:
        valid_deb_days = [d for d in wc_metrics['debtor_days'] if d > 0]
        wc_metrics['avg_deb_days'] = np.mean(valid_deb_days) if valid_deb_days else 0
    else:
        wc_metrics['avg_deb_days'] = 0  # No receivables data available
    
    if has_valid_payables and wc_metrics['creditor_days']:
        valid_cred_days = [d for d in wc_metrics['creditor_days'] if d > 0]
        wc_metrics['avg_cred_days'] = np.mean(valid_cred_days) if valid_cred_days else 0
    else:
        wc_metrics['avg_cred_days'] = 0  # No payables data available
    
    # Add flags to indicate which components are available
    wc_metrics['has_inventory'] = has_valid_inventory
    wc_metrics['has_receivables'] = has_valid_receivables
    wc_metrics['has_payables'] = has_valid_payables
    
    # Ensure all values are valid numbers (not NaN or inf)
    wc_metrics['avg_inv_days'] = ensure_valid_number(wc_metrics['avg_inv_days'], 0)
    wc_metrics['avg_deb_days'] = ensure_valid_number(wc_metrics['avg_deb_days'], 0)
    wc_metrics['avg_cred_days'] = ensure_valid_number(wc_metrics['avg_cred_days'], 0)
    
    return wc_metrics

def calculate_historical_capex_ratio(financials):
    """
    Calculate historical CapEx as % of Revenue for each year and return average
    This provides a more realistic CapEx projection method
    
    Returns:
        dict: {
            'capex_ratios': list of ratios for each year,
            'avg_capex_ratio': average ratio to use for projections
        }
    """
    revenues = financials['revenue']
    fixed_assets = financials['fixed_assets']
    
    capex_ratios = []
    
    # Calculate CapEx for each historical year
    # CapEx = Change in Fixed Assets + Depreciation
    for i in range(len(revenues) - 1):
        # Yahoo data is newest to oldest, so:
        # i=0 is newest year, i=1 is previous year
        newer_fa = fixed_assets[i]
        older_fa = fixed_assets[i + 1]
        revenue = revenues[i]
        
        if revenue > 0:
            # Estimate depreciation as ~5% of older FA (conservative estimate)
            estimated_dep = older_fa * 0.05
            capex = (newer_fa - older_fa) + estimated_dep
            
            # Calculate as % of revenue
            capex_ratio = (capex / revenue) * 100
            
            # Sanity check: Flag unusually high CapEx ratios but allow them
            if capex_ratio > 150.0:
                capex_ratio = 150.0  # Maximum cap at 150%
            elif capex_ratio > 50.0:
                # Just log high CapEx, don't cap it
                pass
            
            capex_ratios.append(capex_ratio)
    
    # Calculate average
    avg_capex_ratio = np.mean(capex_ratios) if capex_ratios else 5.0  # Default 5% if no data
    
    return {
        'capex_ratios': capex_ratios,
        'avg_capex_ratio': avg_capex_ratio
    }

def calculate_bank_fcfe_valuation(projections, ke, terminal_growth, num_shares):
    """
    Bank Valuation using FCFE (Free Cash Flow to Equity)
    
    CRITICAL DIFFERENCES FROM NORMAL DCF:
    - Uses Ke (Cost of Equity), NOT WACC
    - Discounts FCFE, NOT FCFF
    - Values EQUITY directly, NOT Enterprise Value
    - No debt adjustment (debt is raw material for banks)
    
    Formula:
    Equity Value = Σ PV(FCFE) + Terminal Value
    Value per Share = Equity Value / Shares
    
    Args:
        projections: Bank projections with 'fcfe' key
        ke: Cost of Equity % (NOT WACC!)
        terminal_growth: Terminal growth rate %
        num_shares: Outstanding shares
    """
    
    if 'fcfe' not in projections:
        return None, "❌ Bank projections must have 'fcfe' (Free Cash Flow to Equity)"
    
    fcfe_list = projections['fcfe']
    g = terminal_growth
    
    # Validation: Terminal growth must be < Ke
    if g >= ke:
        return None, f"❌ ERROR: Terminal growth ({g:.1f}%) must be less than Ke ({ke:.1f}%)"
    
    # =====================================================
    # STEP 1: Discount FCFE at Ke
    # =====================================================
    pv_fcfe = []
    for year, fcfe in enumerate(fcfe_list, start=1):
        fcfe_rupees = fcfe * 100000  # Convert from Lacs to Rupees
        pv = fcfe_rupees / ((1 + ke / 100) ** year)
        pv_fcfe.append(pv)
    
    sum_pv_fcfe = sum(pv_fcfe)
    
    # =====================================================
    # STEP 2: Terminal Value using Gordon Growth
    # =====================================================
    last_fcfe = fcfe_list[-1] * 100000  # Rupees
    fcfe_terminal = last_fcfe * (1 + g / 100)
    
    terminal_value = fcfe_terminal / ((ke / 100) - (g / 100))
    
    n = len(fcfe_list)
    pv_terminal = terminal_value / ((1 + ke / 100) ** n)
    
    # =====================================================
    # STEP 3: EQUITY VALUE (NOT Enterprise Value!)
    # =====================================================
    equity_value = sum_pv_fcfe + pv_terminal  # In Rupees
    
    # Convert to per share
    value_per_share = equity_value / num_shares if num_shares > 0 else 0
    
    # For banks, sustainable growth rate
    if 'roe' in projections:
        roe = projections.get('roe', 15)
        # Retention ratio = (PAT - FCFE) / PAT
        avg_pat = np.mean([projections['pat'][i] for i in range(len(projections['pat']))])
        avg_fcfe = np.mean(fcfe_list)
        retention_ratio = (avg_pat - avg_fcfe) / avg_pat if avg_pat > 0 else 0.5
        sustainable_growth = roe * retention_ratio
    else:
        sustainable_growth = g
    
    return {
        'valuation_method': 'Bank FCFE (Equity DCF)',
        'fair_value_per_share': value_per_share,
        'equity_value': equity_value / 100000,  # In Lacs
        'sum_pv_fcfe': sum_pv_fcfe / 100000,
        'terminal_value_pv': pv_terminal / 100000,
        'cost_of_equity': ke,
        'terminal_growth': g,
        'sustainable_growth': sustainable_growth,
        'projections': projections,
        'pv_fcfe_by_year': [pv / 100000 for pv in pv_fcfe]
    }, None

def project_financials_bank(financials, years, tax_rate, car_ratio=14.0, rwa_percentage=75.0):
    """
    Bank-Specific Financial Projection using FCFE (Free Cash Flow to Equity) Methodology
    
    CRITICAL: Banks CANNOT use FCFF/WACC. Must use FCFE/Ke.
    
    Why FCFE for banks:
    - Debt is raw material (not financing)
    - Interest expense is operating cost (not financing cost)
    - Capital Adequacy Ratio (CAR) constrains growth
    - Only equity capital matters
    
    FCFE Formula for Banks:
    FCFE = PAT - (Growth in Advances × RWA% × CAR)
    
    Where:
    - PAT = Net Profit After Tax
    - Growth in Advances = Loan book growth
    - RWA = Risk Weighted Assets ≈ 70-80% of Advances
    - CAR = Capital Adequacy Ratio (13-15% in India)
    
    Args:
        financials: Historical financial data
        years: Projection period
        tax_rate: Tax rate %
        car_ratio: Capital Adequacy Ratio % (default 14%)
        rwa_percentage: Risk Weight % (default 75%)
    """
    
    st.info("🏦 **Using Bank FCFE Methodology** (Not FCFF - that's wrong for banks!)")
    
    # =====================================================
    # STEP 1: Get Historical Data
    # =====================================================
    
    # Revenue = Total Revenue (Interest Income + Non-Interest Income)
    revenue_history = financials['revenue'] if 'revenue' in financials else []
    
    # Net Profit After Tax
    nopat_history = financials['nopat'] if 'nopat' in financials else []
    
    # Equity (Book Value)
    equity_history = financials['equity'] if 'equity' in financials else []
    
    if not revenue_history or not nopat_history or not equity_history:
        st.error("❌ Insufficient bank data for FCFE projection")
        return None, None
    
    # =====================================================
    # STEP 2: Calculate Historical Growth Rates
    # =====================================================
    
    # Revenue growth (proxy for advances growth in absence of loan book data)
    if len(revenue_history) >= 2 and revenue_history[-1] > 0 and revenue_history[0] > 0:
        num_years_hist = len(revenue_history) - 1
        revenue_growth = ((revenue_history[0] / revenue_history[-1]) ** (1/num_years_hist) - 1) * 100
        revenue_growth = max(5, min(revenue_growth, 150))  # Cap between 5-150% for banks
    else:
        revenue_growth = 12.0  # Default Indian bank growth
    
    # Historical ROE
    latest_equity = equity_history[0] * 100000  # Convert from Lacs to Rupees
    latest_pat = nopat_history[0] * 100000
    current_roe = (latest_pat / latest_equity * 100) if latest_equity > 0 else 15.0
    
    st.info(f"""
    📊 **Bank Metrics Calculated:**
    - Revenue Growth (Advances proxy): {revenue_growth:.2f}%
    - Current ROE: {current_roe:.2f}%
    - CAR Target: {car_ratio:.1f}%
    - RWA Weight: {rwa_percentage:.1f}%
    """)
    
    # =====================================================
    # STEP 3: Project PAT (Net Profit)
    # =====================================================
    
    # Use ROE × Book Value approach for PAT projection
    projected_pat = []
    projected_equity = []
    projected_revenue = []
    projected_fcfe = []
    
    current_revenue = revenue_history[0]
    current_equity_val = latest_equity
    
    for year in range(1, years + 1):
        # Project revenue (advances)
        current_revenue = current_revenue * (1 + revenue_growth / 100)
        projected_revenue.append(current_revenue / 100000)  # Back to Lacs
        
        # Project equity (retained earnings accumulate)
        # Equity grows with retained PAT
        projected_equity.append(current_equity_val / 100000)
        
        # Project PAT using ROE
        pat = current_equity_val * (current_roe / 100)
        projected_pat.append(pat / 100000)
        
        # =====================================================
        # STEP 4: Calculate FCFE (KEY FORMULA)
        # =====================================================
        # FCFE = PAT - Equity Required for Growth
        # Equity Required = Growth in Advances × RWA% × CAR%
        
        advances_growth_amount = current_revenue * (revenue_growth / 100)  # Rupees
        equity_required = advances_growth_amount * (rwa_percentage / 100) * (car_ratio / 100)
        
        fcfe = pat - equity_required
        projected_fcfe.append(fcfe / 100000)  # Back to Lacs
        
        # Update equity for next year (add retained earnings)
        # Retained = PAT - Dividends, approximate as PAT - FCFE
        retained = pat - fcfe
        current_equity_val += retained
    
    projections = {
        'year': list(range(1, years + 1)),
        'revenue': projected_revenue,  # Total revenue
        'pat': projected_pat,  # Net Profit After Tax
        'equity': projected_equity,  # Book Value
        'fcfe': projected_fcfe  # Free Cash Flow to Equity (THIS IS WHAT WE DISCOUNT!)
    }
    
    drivers = {
        'revenue_growth': revenue_growth,
        'roe': current_roe,
        'car_ratio': car_ratio,
        'rwa_percentage': rwa_percentage,
        'tax_rate': tax_rate
    }
    
    st.success(f"✅ **Bank FCFE Projections Complete:** {years} years projected")
    st.info(f"💡 **FCFE** (not FCFF!) will be discounted at **Ke** (not WACC!) to get equity value directly")
    
    return projections, drivers

def project_financials(financials, wc_metrics, years, tax_rate, 
                      rev_growth_override, opex_margin_override, capex_ratio_override=None,
                      # PER-YEAR overrides (lists of length == years, index 0 = Year 1)
                      rev_growth_per_year=None,   # list[float] | None
                      ebitda_margin_per_year=None, # list[float] | None
                      # NEW PARAMETERS - Complete user control
                      ebitda_margin_override=None,
                      depreciation_rate_override=None,
                      depreciation_method="Auto",
                      inventory_days_override=None,
                      debtor_days_override=None,
                      creditor_days_override=None,
                      interest_rate_override=None,
                      working_capital_pct_override=None):
    """
    INDUSTRY-GRADE Financial Projection Engine with FULL USER CONTROL
    ==================================================================
    
    Uses multiple validation layers and industry best practices:
    1. Revenue: CAGR with economic floor (user can override)
    2. Margins: Normalized with trend analysis (user can override)
    3. CapEx: % of Revenue (user can override)
    4. Working Capital: Days-based methodology (user can override each component)
    5. Depreciation: Multiple methods (user can select)
    6. Interest: Auto or user override
    7. Sanity checks at every step
    
    ALL parameters can be overridden by the user for complete control.
    """
    
    # ============================================
    # STEP 1: CALCULATE HISTORICAL CAPEX RATIO
    # ============================================
    capex_info = calculate_historical_capex_ratio(financials)
    avg_capex_ratio = capex_info['avg_capex_ratio']
    
    # Apply override if provided
    if capex_ratio_override:
        try:
            avg_capex_ratio = float(capex_ratio_override)
        except:
            pass  # Keep calculated value if override is invalid
    
    # ============================================
    # STEP 2: REVENUE GROWTH WITH INTELLIGENT FLOORS
    # ============================================
    revenues = financials['revenue']
    num_years = len(revenues) - 1
    
    if num_years > 0 and revenues[0] > 0 and revenues[-1] > 0:
        # CAGR: Start = OLDEST (last element), End = NEWEST (first element)
        start_revenue = revenues[-1]  # Oldest year
        end_revenue = revenues[0]     # Newest year
        historical_cagr = ((end_revenue / start_revenue) ** (1 / num_years) - 1) * 100
        
        # INDUSTRY PRACTICE: Apply GDP floor for low-growth companies
        # But also cap unrealistic high growth
        if historical_cagr < 4.0:
            # Blend historical with GDP floor (7% for India)
            avg_growth = (historical_cagr * 0.6) + (7.0 * 0.4)
        elif historical_cagr > 150.0:
            # Cap extremely excessive growth - maximum 150%
            avg_growth = 150.0
            try:
                st.warning(f"⚠️ Historical Revenue CAGR ({historical_cagr:.1f}%) capped at 150%. Original: {historical_cagr:.1f}%")
            except:
                pass
        else:
            avg_growth = historical_cagr
            if historical_cagr > 50.0:
                try:
                    st.info(f"📊 High Growth Detected: Revenue CAGR = {historical_cagr:.1f}% (using actual historical rate)")
                except:
                    pass
    else:
        avg_growth = 8.0  # Reasonable default for Indian economy
    
    if rev_growth_override:
        avg_growth = float(rev_growth_override)
    
    # ============================================
    # CRITICAL CAPEX NORMALIZATION RULE
    # ============================================
    # RULE: If CapEx/Revenue ratio > Revenue Growth rate,
    # then CapEx is consuming too much cash relative to growth
    # SOLUTION: Cap CapEx at 1/4 of revenue growth rate
    
    if avg_capex_ratio > avg_growth:
        original_capex_ratio = avg_capex_ratio
        # Floor at MAINTENANCE capex (historical depreciation / revenue). Without a floor this rule
        # could cut CapEx far below depreciation (e.g. 10% -> 1.5% of revenue at 6% growth), which
        # lets the asset base shrink while revenue grows and inflates FCFF.
        _dep_ratios = [
            (financials['depreciation'][i] / financials['revenue'][i]) * 100
            for i in range(len(revenues))
            if financials['revenue'][i] > 0 and financials['depreciation'][i] > 0
        ]
        maintenance_capex_ratio = float(np.median(_dep_ratios)) if _dep_ratios else 0.0
        avg_capex_ratio = min(original_capex_ratio, max(avg_growth / 4.0, maintenance_capex_ratio))
        
        # Log the adjustment for transparency
        import streamlit as st
        st.warning(f"⚠️ **CapEx Normalization Applied**")
        st.info(f"""
        📊 **Original CapEx/Revenue:** {original_capex_ratio:.2f}%
        📈 **Revenue Growth Rate:** {avg_growth:.2f}%
        
        **Issue Detected:** CapEx ratio ({original_capex_ratio:.2f}%) exceeds revenue growth ({avg_growth:.2f}%)
        
        **Action Taken:** CapEx ratio normalized to **{avg_capex_ratio:.2f}%** (the higher of 1/4 of revenue growth and maintenance CapEx = historical depreciation/revenue)
        
        **Rationale:** Sustainable companies cannot indefinitely spend more on CapEx (as % of revenue) 
        than their revenue growth rate. This normalization ensures long-term financial viability.
        """)
    
    # ============================================
    # STEP 3: MARGIN ANALYSIS WITH NORMALIZATION
    # ============================================
    
    # COGS Margin - Use median to avoid outliers
    cogs_margins = []
    for i in range(len(revenues)):
        if financials['revenue'][i] > 0:
            margin = (financials['cogs'][i] / financials['revenue'][i]) * 100
            # Sanity check: COGS should be 20-85% of revenue
            if 20 <= margin <= 85:
                cogs_margins.append(margin)
    
    avg_cogs_margin = np.median(cogs_margins) if cogs_margins else 55.0
    
    # OpEx Margin - Use median and exclude outliers
    opex_margins = []
    for i in range(len(revenues)):
        if financials['revenue'][i] > 0:
            margin = (financials['opex'][i] / financials['revenue'][i]) * 100
            # Sanity check: OpEx should be 5-50% of revenue
            if 5 <= margin <= 50:
                opex_margins.append(margin)
    
    avg_opex_margin = np.median(opex_margins) if opex_margins else 15.0
    
    if opex_margin_override:
        avg_opex_margin = float(opex_margin_override)
    
    # ============================================
    # STEP 4: DEPRECIATION RATE - NORMALIZED
    # ============================================
    dep_rates = []
    for i in range(len(revenues)):
        if financials['fixed_assets'][i] > 0:
            rate = (financials['depreciation'][i] / financials['fixed_assets'][i]) * 100
            # Sanity: Depreciation typically 3-15% of FA
            if 3 <= rate <= 15:
                dep_rates.append(rate)
    
    avg_dep_rate = np.median(dep_rates) if dep_rates else 6.0
    
    # User override for depreciation rate
    if depreciation_rate_override:
        avg_dep_rate = float(depreciation_rate_override)
    
    # ============================================
    # STEP 5: INTEREST RATE ON DEBT
    # ============================================
    total_debts = [financials['st_debt'][i] + financials['lt_debt'][i] for i in range(len(revenues))]
    fin_cost_rates = []
    
    for i in range(len(revenues)):
        if total_debts[i] > 0 and financials['interest'][i] > 0:
            rate = (financials['interest'][i] / total_debts[i]) * 100
            # Sanity: Interest rate should be 4-18%
            if 4 <= rate <= 18:
                fin_cost_rates.append(rate)
    
    avg_fin_cost_rate = np.median(fin_cost_rates) if fin_cost_rates else 8.0
    
    # User override for interest rate
    if interest_rate_override:
        avg_fin_cost_rate = float(interest_rate_override)
    
    # ============================================
    # STEP 6: BALANCE SHEET GROWTH RATES
    # ============================================
    
    # Equity growth (for completeness)
    equity_values = financials['equity']
    if len(equity_values) > 1 and equity_values[-1] > 0 and equity_values[0] > 0:
        avg_equity_growth = ((equity_values[0] / equity_values[-1]) ** (1 / (len(equity_values) - 1)) - 1) * 100
        # Cap equity growth at revenue growth + 5%
        avg_equity_growth = min(avg_equity_growth, avg_growth + 5)
    else:
        avg_equity_growth = avg_growth
    
    # Debt growth - conservative assumption
    if len(total_debts) > 1 and total_debts[-1] > 0 and total_debts[0] > 0:
        historical_debt_growth = ((total_debts[0] / total_debts[-1]) ** (1 / (len(total_debts) - 1)) - 1) * 100
        # INDUSTRY PRACTICE: Debt shouldn't grow faster than revenue
        avg_debt_growth = min(historical_debt_growth, avg_growth * 0.8)
    else:
        avg_debt_growth = 0.0  # Conservative: assume no debt growth
    
    # ============================================
    # STEP 7: PROJECTIONS WITH ROBUST VALIDATION
    # ============================================
    projections = {
        'year': [],
        'revenue': [],
        'cogs': [],
        'opex': [],
        'ebitda': [],
        'depreciation': [],
        'ebit': [],
        'interest': [],
        'nopat': [],
        'fixed_assets': [],
        'equity': [],
        'debt': [],
        'wc': [],
        'delta_wc': [],
        'capex': [],
        'fcff': []
    }
    
    # Starting point - USE NEWEST YEAR DATA (index 0)
    last_revenue = revenues[0]
    last_fa = financials['fixed_assets'][0]
    last_equity = financials['equity'][0]
    last_debt = total_debts[0] if total_debts[0] > 0 else 0
    
    # CRITICAL FIX: Calculate initial working capital from most recent historical year
    # This ensures delta_wc calculations are based on actual historical WC, not zero
    last_inventory = ensure_valid_number(financials['inventory'][0], 0)
    last_receivables = ensure_valid_number(financials['receivables'][0], 0)
    last_payables = ensure_valid_number(financials['payables'][0], 0)
    last_wc = last_inventory + last_receivables - last_payables
    last_wc = ensure_valid_number(last_wc, 0)
    
    for year in range(1, years + 1):
        # ============================================
        # REVENUE PROJECTION
        # ============================================
        # Per-year growth takes priority → single override → auto CAGR
        year_idx = year - 1  # 0-based index
        if rev_growth_per_year and year_idx < len(rev_growth_per_year) and rev_growth_per_year[year_idx] > 0:
            year_growth = rev_growth_per_year[year_idx]
        else:
            year_growth = avg_growth
        projected_revenue = last_revenue * (1 + year_growth / 100)
        
        # ============================================
        # P&L PROJECTIONS
        # ============================================
        projected_cogs = projected_revenue * (avg_cogs_margin / 100)
        projected_opex = projected_revenue * (avg_opex_margin / 100)
        projected_ebitda = projected_revenue - projected_cogs - projected_opex
        
        # Per-year EBITDA margin takes priority → single override → derived
        if ebitda_margin_per_year and year_idx < len(ebitda_margin_per_year) and ebitda_margin_per_year[year_idx] > 0:
            projected_ebitda = projected_revenue * (ebitda_margin_per_year[year_idx] / 100)
        elif ebitda_margin_override:
            # USER OVERRIDE: EBITDA margin directly overrides the derived EBITDA
            projected_ebitda = projected_revenue * (float(ebitda_margin_override) / 100)
        
        # Sanity check: EBITDA should be positive for healthy companies
        # (skip auto-correction if user explicitly set the EBITDA margin)
        active_ebitda_override = (
            (ebitda_margin_per_year and year_idx < len(ebitda_margin_per_year) and ebitda_margin_per_year[year_idx] > 0)
            or ebitda_margin_override
        )
        if projected_ebitda < 0 and not active_ebitda_override:
            # Adjust opex to maintain 5% EBITDA margin
            projected_opex = projected_revenue * 0.85 - projected_cogs
            projected_ebitda = projected_revenue * 0.15
        
        # ============================================
        # CAPEX PROJECTION (INDUSTRY METHOD)
        # ============================================
        # CapEx as % of Revenue (most reliable method)
        capex = projected_revenue * (avg_capex_ratio / 100)
        
        # Depreciation: Apply to growing FA base
        # FA will grow based on net CapEx
        projected_fa = last_fa + capex  # Will subtract depreciation next
        
        # USER OVERRIDE: Depreciation method selector
        if depreciation_method == "% of Revenue":
            projected_dep = projected_revenue * (avg_dep_rate / 100)
        elif depreciation_method == "Absolute Value":
            # depreciation_rate_override (if provided) is treated as an absolute ₹ Lacs value
            projected_dep = float(depreciation_rate_override) if depreciation_rate_override else avg_dep_rate
        else:
            # "Auto" or "% of Fixed Assets" (default)
            projected_dep = projected_fa * (avg_dep_rate / 100)
        
        # Adjust FA after depreciation
        projected_fa = projected_fa - projected_dep
        
        # ============================================
        # EBIT & NOPAT
        # ============================================
        projected_ebit = projected_ebitda - projected_dep
        projected_nopat = projected_ebit * (1 - tax_rate / 100)
        
        # ============================================
        # DEBT & INTEREST
        # ============================================
        projected_debt = last_debt * (1 + avg_debt_growth / 100) if last_debt > 0 else 0
        projected_interest = projected_debt * (avg_fin_cost_rate / 100) if projected_debt > 0 else 0
        
        # ============================================
        # WORKING CAPITAL (INDUSTRY STANDARD METHOD)
        # ROBUST: Handle cases where inventory/receivables/payables data is None
        # CRITICAL FIX: User overrides take priority over historical data
        # ============================================
        
        # Initialize WC components
        projected_inventory = 0
        projected_receivables = 0
        projected_payables = 0
        
        # Determine inventory days: USER OVERRIDE > Historical Data > 0
        if inventory_days_override and inventory_days_override > 0:
            inv_days_to_use = inventory_days_override
        elif wc_metrics.get('has_inventory', False) and wc_metrics['avg_inv_days'] > 0:
            inv_days_to_use = wc_metrics['avg_inv_days']
        else:
            inv_days_to_use = 0
        
        # Determine debtor days: USER OVERRIDE > Historical Data > 0
        if debtor_days_override and debtor_days_override > 0:
            deb_days_to_use = debtor_days_override
        elif wc_metrics.get('has_receivables', False) and wc_metrics['avg_deb_days'] > 0:
            deb_days_to_use = wc_metrics['avg_deb_days']
        else:
            deb_days_to_use = 0
        
        # Determine creditor days: USER OVERRIDE > Historical Data > 0
        if creditor_days_override and creditor_days_override > 0:
            cred_days_to_use = creditor_days_override
        elif wc_metrics.get('has_payables', False) and wc_metrics['avg_cred_days'] > 0:
            cred_days_to_use = wc_metrics['avg_cred_days']
        else:
            cred_days_to_use = 0
        
        # Calculate WC components using determined days
        # For trading companies cogs may be 0; fall back to revenue so that
        # inventory and creditor projections use the same base as debtor days.
        projected_cogs_or_rev = projected_cogs if projected_cogs > 0 else projected_revenue

        if inv_days_to_use > 0:
            projected_inventory = safe_divide(projected_cogs_or_rev * inv_days_to_use, 365, default=0)
        
        if deb_days_to_use > 0:
            projected_receivables = safe_divide(projected_revenue * deb_days_to_use, 365, default=0)
        
        if cred_days_to_use > 0:
            projected_payables = safe_divide(projected_cogs_or_rev * cred_days_to_use, 365, default=0)
        
        # ROBUST: Ensure all WC components are valid numbers
        projected_inventory = ensure_valid_number(projected_inventory, 0)
        projected_receivables = ensure_valid_number(projected_receivables, 0)
        projected_payables = ensure_valid_number(projected_payables, 0)
        
        # Calculate projected WC
        projected_wc = projected_inventory + projected_receivables - projected_payables
        projected_wc = ensure_valid_number(projected_wc, 0)
        
        # USER OVERRIDE: Working Capital as % of Revenue takes priority over
        # the days-based build-up above
        if working_capital_pct_override:
            projected_wc = projected_revenue * (float(working_capital_pct_override) / 100)
            projected_wc = ensure_valid_number(projected_wc, 0)
        
        # Calculate change in WC
        delta_wc = projected_wc - last_wc
        delta_wc = ensure_valid_number(delta_wc, 0)
        
        # CRITICAL: Cap WC changes to prevent unrealistic swings
        # Industry practice: ΔWC shouldn't exceed 20% of revenue
        max_delta_wc = projected_revenue * 0.20
        if abs(delta_wc) > max_delta_wc:
            delta_wc = max_delta_wc if delta_wc > 0 else -max_delta_wc
            projected_wc = last_wc + delta_wc
            projected_wc = ensure_valid_number(projected_wc, 0)
            delta_wc = ensure_valid_number(delta_wc, 0)
        
        # ============================================
        # FCFF CALCULATION
        # ROBUST: Ensure all components are valid numbers before calculation
        # ============================================
        projected_nopat = ensure_valid_number(projected_nopat, 0)
        projected_dep = ensure_valid_number(projected_dep, 0)
        delta_wc = ensure_valid_number(delta_wc, 0)
        capex = ensure_valid_number(capex, 0)
        
        fcff = projected_nopat + projected_dep - delta_wc - capex
        fcff = ensure_valid_number(fcff, 0)  # Ensure FCFF is never NaN
        
        # SANITY CHECK: Only intervene if FCFF is extremely negative AND components are likely invalid
        # Removed aggressive normalization that was incorrectly zeroing out legitimate working capital changes
        # Negative FCFF is acceptable for high-growth companies with working capital needs
        
        # ============================================
        # STORE PROJECTIONS
        # ============================================
        projections['year'].append(year)
        projections['revenue'].append(projected_revenue)
        projections['cogs'].append(projected_cogs)
        projections['opex'].append(projected_opex)
        projections['ebitda'].append(projected_ebitda)
        projections['depreciation'].append(projected_dep)
        projections['ebit'].append(projected_ebit)
        projections['interest'].append(projected_interest)
        projections['nopat'].append(projected_nopat)
        projections['fixed_assets'].append(projected_fa)
        projections['equity'].append(last_equity * (1 + avg_equity_growth / 100))
        projections['debt'].append(projected_debt)
        projections['wc'].append(projected_wc)
        projections['delta_wc'].append(delta_wc)
        projections['capex'].append(capex)
        projections['fcff'].append(fcff)
        
        # Update for next iteration
        last_revenue = projected_revenue
        last_fa = projected_fa
        last_equity = projections['equity'][-1]
        last_debt = projected_debt
        last_wc = projected_wc
    
    return projections, {
        'avg_growth': avg_growth,
        'avg_cogs_margin': avg_cogs_margin,
        'avg_opex_margin': avg_opex_margin,
        'avg_dep_rate': avg_dep_rate,
        'avg_fin_cost_rate': avg_fin_cost_rate,
        'avg_capex_ratio': avg_capex_ratio
    }

def calculate_peer_unlevered_beta(peer_tickers, target_financials, tax_rate, period_years=3,
                                   beta_start_date=None, beta_end_date=None):
    """
    Hamada-equation beta pipeline:
      1. Fetch 3-year raw (levered) beta for every comp ticker.
      2. Unlever each beta using the peer's own D/E ratio.
      3. Average the unlevered betas.
      4. Relever the average using the TARGET company's D/E ratio.

    Hamada formula:
        β_L  = β_U × [1 + (1 − t) × D/E]
        β_U  = β_L  / [1 + (1 − t) × D/E]

    Args:
        peer_tickers  : comma-separated ticker strings (may include .NS/.BO).
        target_financials : financials dict of the company being valued.
        tax_rate      : effective tax rate % (e.g. 25.0).
        period_years  : lookback for beta regression (default 3).

    Returns:
        dict with keys:
            relevered_beta, avg_unlevered_beta, peer_details,
            target_de_ratio, n_peers_used
    """
    import time, random

    ticker_list = [t.strip() for t in peer_tickers.split(',') if t.strip()]
    t = tax_rate / 100.0

    # ── Target D/E from the TARGET company's own balance sheet ──────────
    tgt_eq   = ensure_valid_number(target_financials['equity'][0], 1)   # Lacs
    tgt_std  = ensure_valid_number(target_financials['st_debt'][0], 0)
    tgt_ltd  = ensure_valid_number(target_financials['lt_debt'][0], 0)
    tgt_debt = tgt_std + tgt_ltd
    target_de = tgt_debt / tgt_eq if tgt_eq > 0 else 0.0

    peer_details    = []
    unlevered_betas = []

    for ticker in ticker_list:
        try:
            # Small polite delay between API calls
            if peer_details:
                time.sleep(random.uniform(0.5, 1.0))

            # ── Step 1: raw (levered) beta via regression ─────────────────
            raw_beta = get_stock_beta(ticker, period_years=period_years,
                                       beta_start_date=beta_start_date,
                                       beta_end_date=beta_end_date)
            if raw_beta <= 0:
                st.caption(f"   ↳ {ticker}: skipped (invalid beta {raw_beta:.3f})")
                continue

            # ── Step 2: peer D/E from yfinance ───────────────────────────
            peer_de = 0.0
            try:
                peer_stock = get_cached_ticker(ticker)

                # .info can legitimately return None – guard every access
                raw_info = peer_stock.info
                info = raw_info if isinstance(raw_info, dict) else {}

                peer_equity = ensure_valid_number(info.get('bookValue', 0), 0)
                peer_price  = ensure_valid_number(
                    info.get('currentPrice') or info.get('regularMarketPrice') or 0, 0)
                peer_shares = ensure_valid_number(info.get('sharesOutstanding', 0), 0)
                peer_mktcap = peer_price * peer_shares if peer_price > 0 and peer_shares > 0 else 0

                # Try balance-sheet debt (balance_sheet may be empty / have no columns)
                peer_bs = peer_stock.balance_sheet
                if peer_bs is not None and not peer_bs.empty and len(peer_bs.columns) > 0:
                    bs_col = peer_bs.columns[0]
                    peer_ltd = safe_extract(peer_bs, 'Long Term Debt', bs_col)                         if 'Long Term Debt' in peer_bs.index else 0
                    peer_std = safe_extract(peer_bs, 'Current Debt', bs_col)                         if 'Current Debt' in peer_bs.index else 0
                else:
                    peer_ltd = peer_std = 0
                peer_total_debt = peer_ltd + peer_std

                # Equity denominator: book value × shares → market cap → 0
                peer_eq_value = (peer_equity * peer_shares
                                 if peer_equity > 0 and peer_shares > 0
                                 else peer_mktcap)
                peer_de = peer_total_debt / peer_eq_value if peer_eq_value > 0 else 0.0
                peer_de = max(0.0, peer_de)
            except Exception:
                peer_de = 0.0   # assume all-equity if we can't fetch

            # ── Step 3: unlever ───────────────────────────────────────────
            # β_U = β_L / [1 + (1-t) × D/E]
            unlevered_beta = raw_beta / (1 + (1 - t) * peer_de)
            unlevered_beta = max(0.1, min(unlevered_beta, 3.0))

            unlevered_betas.append(unlevered_beta)
            peer_details.append({
                'ticker'        : ticker,
                'raw_beta'      : raw_beta,
                'peer_de_ratio' : peer_de,
                'unlevered_beta': unlevered_beta,
            })

            st.caption(
                f"   {ticker} → βL={raw_beta:.3f}, D/E={peer_de:.2f}, "
                f"βU={unlevered_beta:.3f}"
            )

        except Exception as e:
            st.warning(f"⚠️ Beta unlever failed for {ticker}: {str(e)[:80]}")
            continue

    if not unlevered_betas:
        # Hard fallback: return levered beta of 1.0
        return {
            'relevered_beta'    : 1.0,
            'avg_unlevered_beta': 1.0,
            'peer_details'      : [],
            'target_de_ratio'   : target_de,
            'n_peers_used'      : 0,
            'fallback'          : True,
        }

    # ── Step 4: average unlevered betas ──────────────────────────────────
    avg_bu = float(np.mean(unlevered_betas))

    # ── Step 5: relever using target company's D/E ───────────────────────
    # β_L = β_U × [1 + (1-t) × D/E_target]
    relevered_beta = avg_bu * (1 + (1 - t) * target_de)
    relevered_beta = max(0.1, min(relevered_beta, 3.0))

    return {
        'relevered_beta'    : relevered_beta,
        'avg_unlevered_beta': avg_bu,
        'peer_details'      : peer_details,
        'target_de_ratio'   : target_de,
        'n_peers_used'      : len(unlevered_betas),
        'fallback'          : False,
    }


def calculate_wacc(financials, tax_rate, peer_tickers=None, manual_rf_rate=None, manual_rm_rate=None,
                   beta_start_date=None, beta_end_date=None, manual_beta=None):
    """Calculate WACC with proper beta calculation from peers.

    manual_beta: if provided (not None), this beta is used directly and the
    entire peer-based Hamada unlever/relever calculation is skipped — no
    peer_tickers required, no st.* UI emitted for the beta section. Used by
    Bulk Valuation mode, which applies one user-supplied beta across every
    company in a batch instead of computing it per company. Existing callers
    that don't pass manual_beta are completely unaffected.
    """
    # Cost of Equity (Ke)
    # ALWAYS use manual rates (passed from session state), never fetch
    rf = manual_rf_rate if manual_rf_rate is not None else 6.83
    rm = manual_rm_rate if manual_rm_rate is not None else 12.0
    
    # ── Beta: unlever peer betas → average → relever for target ──────────
    beta = 1.0
    beta_details = {}
    if manual_beta is not None:
        beta = float(manual_beta)
        beta_details = {'source': 'manual', 'manual_beta': beta}
    elif peer_tickers and peer_tickers.strip():
        st.markdown("#### 🔢 Beta Calculation (Hamada Unlevering / Relevering)")
        if beta_start_date and beta_end_date:
            _bd_label = f"{pd.Timestamp(beta_start_date).strftime('%d-%b-%Y')} → {pd.Timestamp(beta_end_date).strftime('%d-%b-%Y')}"
        else:
            _bd_label = "3-year rolling window (default)"
        st.caption(
            f"Beta window: {_bd_label} | Daily returns | "
            "Unlever each peer β (Hamada) → average → relever to target D/E."
        )
        beta_result = calculate_peer_unlevered_beta(
            peer_tickers, financials, tax_rate, period_years=3,
            beta_start_date=beta_start_date, beta_end_date=beta_end_date
        )
        beta         = beta_result['relevered_beta']
        beta_details = beta_result

        if beta_result.get('fallback'):
            st.warning("⚠️ Could not compute unlevered beta from peers — using default β=1.0")
        else:
            st.success(
                f"✅ Avg Unlevered β = {beta_result['avg_unlevered_beta']:.3f} | "
                f"Target D/E = {beta_result['target_de_ratio']:.2f} | "
                f"**Relevered β = {beta:.3f}** "
                f"(from {beta_result['n_peers_used']} peers)"
            )
            with st.expander("📋 Beta Breakdown by Peer"):
                for pd_row in beta_result['peer_details']:
                    st.write(
                        f"• **{pd_row['ticker']}** | "
                        f"Levered β = {pd_row['raw_beta']:.3f} | "
                        f"Peer D/E = {pd_row['peer_de_ratio']:.2f} | "
                        f"Unlevered β = {pd_row['unlevered_beta']:.3f}"
                    )
    else:
        st.warning("⚠️ No peer tickers provided, using default β=1.0")

    ke = rf + (beta * (rm - rf))
    
    # Cost of Debt (Kd) - USE NEWEST values (index 0)
    # Handle debt properly - could be 0 or NaN
    st_debt = financials['st_debt'][0] if financials['st_debt'][0] > 0 else 0
    lt_debt = financials['lt_debt'][0] if financials['lt_debt'][0] > 0 else 0
    total_debt = st_debt + lt_debt
    interest = financials['interest'][0]
    
    # Cost of Debt - handle zero/NaN debt
    if total_debt > 0 and interest > 0:
        kd = (interest / total_debt * 100)
    else:
        kd = 0.0  # Debt-free company has no cost of debt
    kd_after_tax = kd * (1 - tax_rate / 100)
    
    # WACC - USE NEWEST equity (index 0)
    equity = financials['equity'][0]
    total_capital = equity + total_debt
    
    # Handle weights - CRITICAL: Support negative equity scenarios
    # When equity is negative, weights can be >100% or negative
    if total_capital != 0:
        we = equity / total_capital
        wd = total_debt / total_capital
    else:
        # Edge case: total capital = 0 (equity = -debt)
        we = 1.0
        wd = 0.0
    
    wacc = (we * ke) + (wd * kd_after_tax)
    
    return {
        'wacc': wacc,
        'ke': ke,
        'kd': kd,
        'kd_after_tax': kd_after_tax,
        'rf': rf,
        'rm': rm,
        'beta': beta,
        'we': we * 100,
        'wd': wd * 100,
        'equity': equity,
        'debt': total_debt,
        'beta_details': beta_details,
    }

def calculate_wacc_bank(financials, tax_rate, peer_tickers=None, manual_rf_rate=None, manual_rm_rate=None,
                        beta_start_date=None, beta_end_date=None):
    """
    Calculate WACC for BANKS/NBFCs with proper Cost of Funds methodology
    
    For banks:
    - Kd = Cost of Funds (WACF) = Interest Expended / Average Interest-Bearing Liabilities
    - Interest-bearing liabilities include: Deposits, Borrowings, Bonds, Subordinated debt
    - NOT just simple Interest/Debt ratio (that's wrong for banks!)
    """
    
    # Cost of Equity (Ke) - Same as normal companies
    # ALWAYS use manual rates (passed from session state), never fetch
    rf = manual_rf_rate if manual_rf_rate is not None else 6.83
    rm = manual_rm_rate if manual_rm_rate is not None else 12.0
    
    # ── Beta: unlever peer betas → average → relever for target (Bank) ────
    beta = 1.0
    beta_details = {}
    if peer_tickers and peer_tickers.strip():
        st.markdown("#### 🔢 Beta Calculation (Hamada Unlevering / Relevering) — Bank Mode")
        if beta_start_date and beta_end_date:
            _bd_label = f"{pd.Timestamp(beta_start_date).strftime('%d-%b-%Y')} → {pd.Timestamp(beta_end_date).strftime('%d-%b-%Y')}"
        else:
            _bd_label = "3-year rolling window (default)"
        st.caption(
            f"Beta window: {_bd_label} | Daily returns | "
            "Unlevering peer betas then relevering to target bank's own D/E."
        )
        beta_result = calculate_peer_unlevered_beta(
            peer_tickers, financials, tax_rate, period_years=3,
            beta_start_date=beta_start_date, beta_end_date=beta_end_date
        )
        beta         = beta_result['relevered_beta']
        beta_details = beta_result

        if beta_result.get('fallback'):
            st.warning("⚠️ Could not compute unlevered beta — using default β=1.0")
        else:
            st.success(
                f"✅ Avg Unlevered β = {beta_result['avg_unlevered_beta']:.3f} | "
                f"Target D/E = {beta_result['target_de_ratio']:.2f} | "
                f"**Relevered β = {beta:.3f}** "
                f"(from {beta_result['n_peers_used']} bank peers)"
            )
            with st.expander("📋 Bank Beta Breakdown by Peer"):
                for pd_row in beta_result['peer_details']:
                    st.write(
                        f"• **{pd_row['ticker']}** | "
                        f"Levered β = {pd_row['raw_beta']:.3f} | "
                        f"Peer D/E = {pd_row['peer_de_ratio']:.2f} | "
                        f"Unlevered β = {pd_row['unlevered_beta']:.3f}"
                    )
    
    ke = rf + (beta * (rm - rf))
    
    # Cost of Debt (Kd) - BANK METHODOLOGY
    # Kd = WACF (Weighted Average Cost of Funds)
    # Formula: Interest Expended / Average Interest-Bearing Liabilities
    
    interest_expense = financials['interest'][0]  # Total interest paid
    
    # Interest-bearing liabilities (need to calculate from balance sheet)
    # For banks: Deposits + Borrowings = primary interest-bearing liabilities
    # We approximate: Total Liabilities - Equity - Other non-interest liabilities
    
    equity_current = financials['equity'][0]
    equity_previous = financials['equity'][1] if len(financials['equity']) > 1 else equity_current
    
    # For banks, we need total assets - equity = total liabilities (most are interest-bearing)
    # Rough approximation: Use debt as proxy for interest-bearing liabilities
    st_debt_current = financials['st_debt'][0] if financials['st_debt'][0] > 0 else 0
    lt_debt_current = financials['lt_debt'][0] if financials['lt_debt'][0] > 0 else 0
    total_liabilities_current = st_debt_current + lt_debt_current
    
    st_debt_previous = financials['st_debt'][1] if len(financials['st_debt']) > 1 and financials['st_debt'][1] > 0 else st_debt_current
    lt_debt_previous = financials['lt_debt'][1] if len(financials['lt_debt']) > 1 and financials['lt_debt'][1] > 0 else lt_debt_current
    total_liabilities_previous = st_debt_previous + lt_debt_previous
    
    # Average interest-bearing liabilities
    avg_interest_bearing_liabilities = (total_liabilities_current + total_liabilities_previous) / 2
    
    # Cost of Funds (Kd for banks)
    if avg_interest_bearing_liabilities > 0 and interest_expense > 0:
        kd = (interest_expense / avg_interest_bearing_liabilities) * 100
        st.info(f"💡 **Bank Cost of Funds (Kd):** {kd:.2f}% = Interest Expended ₹{interest_expense:,.0f} / Avg Liabilities ₹{avg_interest_bearing_liabilities:,.0f}")
    else:
        kd = 5.0  # Default reasonable cost of funds for banks
        st.warning("⚠️ Using default bank cost of funds: 5%")
    
    # For banks, tax shield on interest is different (interest is operating expense, not financing)
    # But for WACC calculation, we still apply tax shield
    kd_after_tax = kd * (1 - tax_rate / 100)
    
    # Capital structure weights
    total_capital = equity_current + total_liabilities_current
    
    # Handle weights - CRITICAL: Support negative equity scenarios
    if total_capital != 0:
        we = equity_current / total_capital
        wd = total_liabilities_current / total_capital
    else:
        # Edge case: total capital = 0 (equity = -debt)
        we = 0.20  # Banks typically have low equity weight
        wd = 0.80
    
    wacc = (we * ke) + (wd * kd_after_tax)
    
    st.success(f"✅ **Bank WACC Calculated:** {wacc:.2f}% | Equity Weight: {we*100:.1f}% | Debt Weight: {wd*100:.1f}%")
    
    return {
        'wacc': wacc,
        'ke': ke,
        'kd': kd,
        'kd_after_tax': kd_after_tax,
        'rf': rf,
        'rm': rm,
        'beta': beta,
        'we': we * 100,
        'wd': wd * 100,
        'equity': equity_current,
        'debt': total_liabilities_current,
        'calculation_method': 'Bank Methodology (Cost of Funds)',
        'interest_expense': interest_expense,
        'avg_liabilities': avg_interest_bearing_liabilities,
        'beta_details': beta_details,
    }

def calculate_dcf_valuation(projections, wacc_details, terminal_growth, num_shares, cash_balance=0, manual_discount_rate=None,
                            allow_fcff_recovery=True):
    """
    Calculate DCF valuation with Rulebook-compliant validations and (optional) FCFF recovery
    
    Args:
        manual_discount_rate: Optional manual override for discount rate (instead of WACC)
        allow_fcff_recovery: If True (default, interactive UI) a non-positive terminal-year FCFF is
            replaced by a normalized terminal FCFF for the TERMINAL VALUE only, and the caller's
            projections are never modified. If False (used by unattended Bulk mode) the function
            returns an error instead of valuing a company on an invented cash flow.
    """
    # Use manual discount rate if provided, otherwise use WACC
    if manual_discount_rate and manual_discount_rate > 0:
        wacc = manual_discount_rate
        discount_rate_source = f"Manual Override ({manual_discount_rate:.2f}%)"
    else:
        wacc = wacc_details['wacc']
        discount_rate_source = f"WACC ({wacc:.2f}%)"
    
    g = terminal_growth
    
    # RULEBOOK SECTION 8.2: Terminal growth must be < discount rate
    if g >= wacc:
        return None, f"❌ HARD ERROR: Terminal growth rate must be less than discount rate (Rulebook 8.2). Current: TG={g:.1f}%, Discount={wacc:.1f}%"
    
    # CRITICAL FIX: Terminal growth should generally be lower than long-term revenue growth
    # Extract revenue growth from projections
    if len(projections['revenue']) >= 2:
        first_rev = projections['revenue'][0]
        last_rev = projections['revenue'][-1]
        num_years = len(projections['revenue']) - 1
        
        # BUGFIX: Protect against zero/negative revenues to prevent ZeroDivisionError
        if first_rev > 0 and last_rev > 0:
            implied_revenue_cagr = ((last_rev / first_rev) ** (1 / num_years) - 1) * 100
        else:
            implied_revenue_cagr = 0  # Default to 0 if revenue data is invalid
            st.warning("⚠️ Revenue data contains zero or negative values. Cannot calculate revenue CAGR.")
        
        # Warning if terminal growth is too close to revenue growth
        if g > implied_revenue_cagr * 0.9 and implied_revenue_cagr > 0:
            st.warning(f"⚠️ Terminal Growth Rate ({g:.1f}%) is very close to projected revenue CAGR ({implied_revenue_cagr:.1f}%)")
            st.info("💡 **Recommendation:** Terminal growth should typically be 40-60% of revenue growth for conservative valuations")
    
    # RULEBOOK SECTION 8.2: Check terminal year FCFF
    last_fcff = projections['fcff'][-1]
    fcff_adjusted = False
    adjustment_details = {}
    
    if last_fcff <= 0 and not allow_fcff_recovery:
        return None, (f"❌ Terminal-year FCFF is {last_fcff:.2f} Lacs (≤ 0) and FCFF recovery is disabled for this run — "
                      "a DCF on this company would require inventing a terminal cash flow.")
    
    if last_fcff <= 0:
        # =====================================================
        # INTELLIGENT FCFF RECOVERY MECHANISM
        # =====================================================
        st.warning(f"⚠️ Terminal year FCFF is {last_fcff:.2f} Lacs (negative or zero)")
        st.info("🔧 **Activating Intelligent FCFF Recovery Mechanism**")
        
        # Analyze all projected FCFFs to understand the issue
        all_fcffs = projections['fcff']
        positive_fcffs = [f for f in all_fcffs if f > 0]
        
        # Even if NO positive FCFFs, we can still recover using ULTRA-AGGRESSIVE strategies
        if len(positive_fcffs) == 0:
            st.error("⚠️ **SEVERE CASE:** All projected FCFFs are negative or zero")
            st.warning("🔧 **Activating ULTRA-AGGRESSIVE Recovery Mechanisms**")
            st.caption("Using fundamental value drivers to construct viable terminal FCFF")
        
        # Calculate intelligent adjustments
        avg_positive_fcff = np.mean(positive_fcffs) if positive_fcffs else 0
        median_positive_fcff = np.median(positive_fcffs) if positive_fcffs else 0
        max_fcff = max(all_fcffs) if all_fcffs else 0
        
        # Analyze components to find best recovery path
        last_nopat = projections['nopat'][-1]
        last_dep = projections['depreciation'][-1]
        last_dwc = projections['delta_wc'][-1]
        last_capex = projections['capex'][-1]
        
        # Calculate component ratios
        revenue = projections['revenue'][-1]
        ebitda = projections['ebitda'][-1]
        
        # INTELLIGENT RECOVERY STRATEGIES
        recovery_options = []
        
        # =====================================================
        # ULTRA-AGGRESSIVE STRATEGIES (For severe cases with all negative FCFFs)
        # =====================================================
        
        # Ultra Strategy 1: Revenue-Based Proxy with Industry Margins
        if revenue > 0:
            # Assume industry-standard metrics:
            # - Operating Margin: 10-15% (conservative for mature company)
            # - Tax Rate: from WACC details or default 25%
            # - FCFF/Revenue: 5-8% for mature companies
            tax_rate = wacc_details.get('tax_rate', 25)
            
            # Conservative approach: 6% FCFF/Revenue ratio
            fcff_from_revenue = revenue * 0.06
            
            if fcff_from_revenue > 0:
                recovery_options.append({
                    'strategy': 'Revenue-Based Proxy (Ultra-Aggressive)',
                    'fcff': fcff_from_revenue,
                    'adjustments': {
                        'revenue': f'₹{revenue:.2f} Lacs',
                        'fcff_margin': '6% of Revenue (conservative industry standard)',
                        'rationale': 'Assumes normalized mature company cash conversion',
                        'calculated_fcff': f'{fcff_from_revenue:.2f} Lacs'
                    }
                })
        
        # Ultra Strategy 2: Reverse-Engineered from Growth Rate
        # If we know terminal growth rate, work backwards from sustainable metrics
        if terminal_growth > 0 and revenue > 0:
            # Sustainable FCFF = Revenue × Terminal Growth × FCFF/Growth ratio
            # Typical FCFF/Growth ratio: 2-4 (i.e., if growing at 5%, FCFF ~10-20% of revenue)
            sustainable_fcff = revenue * (terminal_growth / 100) * 2.5
            
            if sustainable_fcff > 0:
                recovery_options.append({
                    'strategy': 'Growth-Reverse-Engineered (Ultra-Aggressive)',
                    'fcff': sustainable_fcff,
                    'adjustments': {
                        'terminal_growth': f'{terminal_growth}%',
                        'revenue': f'₹{revenue:.2f} Lacs',
                        'fcff_growth_ratio': '2.5x (industry standard)',
                        'logic': f'Sustainable FCFF = Revenue × {terminal_growth}% × 2.5',
                        'calculated_fcff': f'{sustainable_fcff:.2f} Lacs'
                    }
                })
        
        # Ultra Strategy 3: NOPAT-Based with Zero CapEx/WC (Absolute Floor)
        # Assumes company transitions to asset-light model
        if last_nopat > 0:
            fcff_nopat_only = last_nopat + last_dep  # Just operating cash
            
            if fcff_nopat_only > 0:
                recovery_options.append({
                    'strategy': 'NOPAT-Based Floor (Ultra-Aggressive)',
                    'fcff': fcff_nopat_only,
                    'adjustments': {
                        'nopat': f'₹{last_nopat:.2f} Lacs',
                        'depreciation': f'₹{last_dep:.2f} Lacs',
                        'capex': 'Assumed ZERO (asset-light transition)',
                        'working_capital': 'Assumed ZERO (normalized)',
                        'calculated_fcff': f'{fcff_nopat_only:.2f} Lacs'
                    }
                })
        
        # Ultra Strategy 4: Operating Cash Flow Proxy
        # EBITDA - Tax (ignoring all capital requirements)
        if ebitda > 0:
            tax_rate = wacc_details.get('tax_rate', 25)
            operating_cash = ebitda * (1 - tax_rate / 100)
            
            if operating_cash > 0:
                recovery_options.append({
                    'strategy': 'Operating Cash Proxy (Ultra-Aggressive)',
                    'fcff': operating_cash,
                    'adjustments': {
                        'ebitda': f'₹{ebitda:.2f} Lacs',
                        'tax_rate': f'{tax_rate}%',
                        'assumption': 'Pure operating cash, no capital requirements',
                        'calculated_fcff': f'{operating_cash:.2f} Lacs'
                    }
                })
        
        # Ultra Strategy 5: Minimum Viable FCFF (Absolute Last Resort)
        # Use 3% of revenue as bare minimum sustainable cash generation
        if revenue > 0:
            min_viable_fcff = revenue * 0.03
            
            recovery_options.append({
                'strategy': 'Minimum Viable FCFF (Last Resort)',
                'fcff': min_viable_fcff,
                'adjustments': {
                    'revenue': f'₹{revenue:.2f} Lacs',
                    'fcff_margin': '3% of Revenue (bare minimum)',
                    'rationale': 'Absolute floor for sustainable operations',
                    'note': 'This is the MOST aggressive assumption',
                    'calculated_fcff': f'{min_viable_fcff:.2f} Lacs'
                }
            })
        
        # =====================================================
        # STANDARD RECOVERY STRATEGIES
        # =====================================================
        
        # Strategy 1: Reduce CapEx to sustainable level (typically 80% of depreciation)
        sustainable_capex = last_dep * 0.8
        capex_savings = last_capex - sustainable_capex if last_capex > sustainable_capex else 0
        fcff_with_capex_adj = last_nopat + last_dep - last_dwc - sustainable_capex
        if fcff_with_capex_adj > 0:
            recovery_options.append({
                'strategy': 'Reduced CapEx to Sustainable Level',
                'fcff': fcff_with_capex_adj,
                'adjustments': {
                    'capex': f'Reduced from {last_capex:.2f} to {sustainable_capex:.2f} Lacs (80% of Depreciation)',
                    'savings': f'CapEx savings: {capex_savings:.2f} Lacs'
                }
            })
        
        # Strategy 2: Normalize Working Capital (assume zero working capital change in terminal year)
        fcff_with_wc_normalization = last_nopat + last_dep - last_capex
        if fcff_with_wc_normalization > 0:
            recovery_options.append({
                'strategy': 'Normalized Working Capital',
                'fcff': fcff_with_wc_normalization,
                'adjustments': {
                    'working_capital': f'Set ΔWC to 0 (from {last_dwc:.2f} Lacs)',
                    'assumption': 'Working capital stabilizes at terminal year'
                }
            })
        
        # Strategy 3: Combined approach (sustainable capex + normalized WC)
        fcff_combined = last_nopat + last_dep - sustainable_capex
        if fcff_combined > 0:
            recovery_options.append({
                'strategy': 'Combined: Sustainable CapEx + Normalized WC',
                'fcff': fcff_combined,
                'adjustments': {
                    'capex': f'CapEx at 80% depreciation: {sustainable_capex:.2f} Lacs',
                    'working_capital': 'ΔWC = 0 (normalized)',
                    'combined_impact': f'Total improvement: {fcff_combined - last_fcff:.2f} Lacs'
                }
            })
        
        # Strategy 4: Use average of positive FCFFs (conservative)
        if avg_positive_fcff > 0:
            recovery_options.append({
                'strategy': 'Average of Positive Historical FCFFs',
                'fcff': avg_positive_fcff,
                'adjustments': {
                    'basis': f'Average of {len(positive_fcffs)} positive FCFF years',
                    'value': f'{avg_positive_fcff:.2f} Lacs',
                    'rationale': 'Uses historical positive cash generation as sustainable baseline'
                }
            })
        
        # Strategy 5: EBITDA-based proxy (if operations are profitable)
        if ebitda > 0:
            # Typical FCFF = ~40-60% of EBITDA for mature companies
            fcff_from_ebitda = ebitda * 0.5 * (1 - wacc_details.get('tax_rate', 25) / 100)
            if fcff_from_ebitda > 0:
                recovery_options.append({
                    'strategy': 'EBITDA-Based Proxy',
                    'fcff': fcff_from_ebitda,
                    'adjustments': {
                        'ebitda': f'{ebitda:.2f} Lacs',
                        'conversion': '50% FCFF/EBITDA ratio (industry standard)',
                        'tax_adjusted': 'Applied tax rate',
                        'calculated_fcff': f'{fcff_from_ebitda:.2f} Lacs'
                    }
                })
        
        # Select BEST recovery option
        # With ultra-aggressive strategies, we ALWAYS have at least one option
        if not recovery_options:
            # Absolute fallback - should never reach here
            # Create emergency FCFF based on revenue if available
            if revenue > 0:
                emergency_fcff = revenue * 0.02  # 2% of revenue
                recovery_options.append({
                    'strategy': 'Emergency Fallback (2% of Revenue)',
                    'fcff': emergency_fcff,
                    'adjustments': {
                        'note': 'Emergency fallback - uses minimal assumptions',
                        'revenue': f'₹{revenue:.2f} Lacs',
                        'calculated_fcff': f'{emergency_fcff:.2f} Lacs'
                    }
                })
            elif ebitda > 0:
                emergency_fcff = ebitda * 0.3  # 30% of EBITDA
                recovery_options.append({
                    'strategy': 'Emergency Fallback (30% of EBITDA)',
                    'fcff': emergency_fcff,
                    'adjustments': {
                        'note': 'Emergency fallback - uses minimal assumptions',
                        'ebitda': f'₹{ebitda:.2f} Lacs',
                        'calculated_fcff': f'{emergency_fcff:.2f} Lacs'
                    }
                })
            else:
                # Truly desperate - use 1 Lac as minimum
                recovery_options.append({
                    'strategy': 'Absolute Minimum (₹1 Lac)',
                    'fcff': 1.0,
                    'adjustments': {
                        'note': '⚠️ No viable metrics available - using symbolic minimum',
                        'warning': 'Valuation highly uncertain - use with extreme caution'
                    }
                })
        
        # Selection rule: prefer options built from the company's OWN components (capex / working-capital
        # normalization, its own positive FCFFs, EBITDA conversion). Revenue-only proxies and
        # emergency fallbacks are used only if no company-derived option exists. Within a tier pick the
        # LOWEST FCFF (most conservative) - previously the option with the highest hard-coded label
        # priority won, regardless of how aggressive its FCFF was.
        _company_derived = {
            'Combined: Sustainable CapEx + Normalized WC',
            'Reduced CapEx to Sustainable Level',
            'Normalized Working Capital',
            'Average of Positive Historical FCFFs',
            'EBITDA-Based Proxy',
        }
        _tier1 = [o for o in recovery_options if o['strategy'] in _company_derived and o['fcff'] > 0]
        _pool = _tier1 if _tier1 else recovery_options
        best_option = min(_pool, key=lambda x: x['fcff'])
        
        # Apply the best recovery strategy
        last_fcff = best_option['fcff']
        fcff_adjusted = True
        adjustment_details = best_option
        
        # Display the recovery strategy
        st.success(f"✅ **Selected Recovery Strategy:** {best_option['strategy']}")
        st.write(f"**Adjusted Terminal FCFF:** ₹{last_fcff:.2f} Lacs (Original: ₹{projections['fcff'][-1]:.2f} Lacs)")
        
        # Add extra warning for ultra-aggressive strategies
        if 'Ultra-Aggressive' in best_option['strategy'] or 'Last Resort' in best_option['strategy'] or 'Emergency' in best_option['strategy'] or 'Absolute' in best_option['strategy']:
            st.warning("⚠️ **CAUTION:** This strategy uses aggressive assumptions due to severe cash flow issues")
            st.caption("💡 Recommendation: Cross-validate with alternative valuation methods (P/E, P/B, EV/EBITDA)")
        
        with st.expander("📋 View All Recovery Options Considered"):
            for i, opt in enumerate(recovery_options, 1):
                st.markdown(f"### Option {i}: {opt['strategy']}")
                st.write(f"**Resulting FCFF:** ₹{opt['fcff']:.2f} Lacs")
                st.write("**Adjustments:**")
                for key, value in opt['adjustments'].items():
                    st.write(f"  - {key.replace('_', ' ').title()}: {value}")
                if opt == best_option:
                    st.success("✅ **SELECTED** (most conservative company-derived option)")
                st.markdown("---")
        
        st.info("💡 **Note:** These adjustments reflect sustainable long-term assumptions required for terminal value calculation.")
        
        # NOTE: the caller's projections are intentionally NOT modified. The explicit-period cash
        # flows stay as projected; the normalized FCFF is used for the terminal value only and is
        # reported via valuation['adjusted_terminal_fcff'].
        st.caption(f"📌 Normalized FCFF of ₹{last_fcff:.2f} Lacs is used for the terminal value only; projected FCFFs are unchanged")
    
    # Present Value of FCFFs
    pv_fcffs = []
    for i, fcff in enumerate(projections['fcff']):
        year = i + 1
        pv = fcff / ((1 + wacc / 100) ** year)
        pv_fcffs.append(pv)
    
    sum_pv_fcff = sum(pv_fcffs)
    
    # If the explicit-period PV is negative (cash burn during growth), KEEP it. It used to be set to 0,
    # which ignored real cash burn and overstated value. Terminal value may then legitimately exceed
    # 100% of EV (see the sanity check below).
    growth_phase_adjusted = False
    original_sum_pv_fcff = sum_pv_fcff
    negative_pv_fcff = sum_pv_fcff < 0
    if negative_pv_fcff:
        st.warning(f"⚠️ Sum of PV(FCFF) over the projection period is negative (₹{sum_pv_fcff:.2f} Lacs): "
                   "the company burns cash before the terminal year. This burn is deducted from value, not ignored.")
    
    # Terminal Value (Rulebook Section 8.1)
    fcff_n_plus_1 = last_fcff * (1 + g / 100)
    terminal_value = fcff_n_plus_1 / ((wacc / 100) - (g / 100))
    
    n = len(projections['fcff'])
    pv_terminal_value = terminal_value / ((1 + wacc / 100) ** n)
    
    # Enterprise Value (in Lacs)
    enterprise_value = sum_pv_fcff + pv_terminal_value
    
    # RULEBOOK SECTION 13.1: Terminal Value sanity checks
    tv_percentage = (pv_terminal_value / enterprise_value * 100) if enterprise_value > 0 else 0
    
    if tv_percentage > 100 and not negative_pv_fcff:
        return None, f"❌ ERROR: Terminal Value ({tv_percentage:.1f}%) exceeds 100% of Enterprise Value (Rulebook 13.1)"
    
    # Equity Value Calculation: EV - Net Debt
    # Handle debt properly (could be 0, None, or NaN for debt-free companies)
    total_debt = wacc_details.get('debt', 0)
    if total_debt is None or (isinstance(total_debt, float) and np.isnan(total_debt)):
        total_debt = 0
    total_debt = float(total_debt) if total_debt > 0 else 0
    
    # Handle cash properly (extract from parameter, could be 0)
    cash = float(cash_balance) if cash_balance and cash_balance > 0 else 0
    
    # Net Debt = Total Debt - Cash
    # Can be negative (net cash position) if cash > debt
    net_debt = total_debt - cash
    
    # Equity Value = Enterprise Value - Net Debt
    # For debt-free companies: EV - (-cash) = EV + cash
    equity_value = enterprise_value - net_debt
    
    # Convert Equity Value from Lacs to absolute Rupees, then divide by shares
    # Equity Value is in Lacs, so multiply by 100,000 to get Rupees
    equity_value_rupees = equity_value * 100000
    fair_value_per_share = equity_value_rupees / num_shares if num_shares > 0 else 0
    
    # CRITICAL VALIDATION: Check if fair value is negative
    negative_fair_value_warning = None
    if fair_value_per_share < 0:
        negative_fair_value_warning = {
            'enterprise_value': enterprise_value,
            'total_debt': total_debt,
            'cash': cash,
            'net_debt': net_debt,
            'equity_value': equity_value,
            'num_shares': num_shares,
            'reason': []
        }
        
        # Diagnose the problem
        if enterprise_value < net_debt:
            negative_fair_value_warning['reason'].append(
                f"Enterprise Value (₹{enterprise_value:.2f} Lacs) is less than Net Debt (₹{net_debt:.2f} Lacs)"
            )
        if enterprise_value <= 0:
            negative_fair_value_warning['reason'].append(
                f"Enterprise Value is zero or negative (₹{enterprise_value:.2f} Lacs)"
            )
        if net_debt > enterprise_value * 2:
            negative_fair_value_warning['reason'].append(
                f"Net Debt (₹{net_debt:.2f} Lacs) is more than 2x Enterprise Value"
            )
    
    return {
        'pv_fcffs': pv_fcffs,
        'sum_pv_fcff': sum_pv_fcff,
        'original_sum_pv_fcff': original_sum_pv_fcff if growth_phase_adjusted else sum_pv_fcff,
        'growth_phase_adjusted': growth_phase_adjusted,
        'negative_pv_fcff': negative_pv_fcff,
        'terminal_value': terminal_value,
        'pv_terminal_value': pv_terminal_value,
        'enterprise_value': enterprise_value,
        'total_debt': total_debt,
        'cash': cash,
        'net_debt': net_debt,
        'equity_value': equity_value,
        'equity_value_rupees': equity_value_rupees,
        'fair_value_per_share': fair_value_per_share,
        'tv_percentage': tv_percentage,
        'tv_warning': tv_percentage > 90,  # Flag for warning
        'fcff_adjusted': fcff_adjusted,
        'adjustment_details': adjustment_details if fcff_adjusted else None,
        'adjusted_terminal_fcff': last_fcff if fcff_adjusted else projections['fcff'][-1],
        'negative_fair_value_warning': negative_fair_value_warning,
        'wacc': wacc,
        'discount_rate_source': discount_rate_source
    }, None


def render_scenario_settings_ui(key_prefix, default_pct=10.0):
    """
    Scenario Analysis INPUT controls, rendered ABOVE the Run Valuation
    button so the choice is locked in before the model runs.

    Lets the user pick which case to spotlight by default, how big a swing
    (%) to apply, and which of the five levers (Revenue Growth, EBITDA
    Margin, WACC, CapEx Ratio, Working Capital) should move across
    scenarios. Returns (default_case, override_pct, flex_drivers_set).
    """
    with st.expander("🎭 Scenario Analysis Settings (Base / Best / Worst)", expanded=False):
        st.caption(
            "All three cases are always computed and shown side-by-side in the "
            "Scenario Analysis tab after you run the valuation. This just sets "
            "the swing size, which levers move, and which case opens by default."
        )
        col_a, col_b = st.columns(2)
        with col_a:
            default_case = st.selectbox(
                "Default case to open:", ["Base Case", "Best Case", "Worst Case"],
                index=0, key=f"{key_prefix}_case"
            )
        with col_b:
            override_pct = st.number_input(
                "Adjustment % (applied to every selected lever):",
                min_value=0.5, max_value=100.0, value=default_pct, step=0.5,
                key=f"{key_prefix}_pct",
                help="Best case moves each selected lever this % in the favorable "
                     "direction; Worst case moves it this % unfavorably."
            )
        driver_labels = st.multiselect(
            "Levers to flex across scenarios:",
            ["Revenue Growth", "EBITDA Margin", "WACC", "CapEx Ratio", "Working Capital"],
            default=["Revenue Growth", "EBITDA Margin", "WACC", "CapEx Ratio", "Working Capital"],
            key=f"{key_prefix}_drivers",
            help="Unchecked levers stay fixed at the base-case value in every scenario."
        )
        label_to_key = {
            'Revenue Growth': 'growth', 'EBITDA Margin': 'ebitda_margin',
            'WACC': 'wacc', 'CapEx Ratio': 'capex_ratio', 'Working Capital': 'wc_pct'
        }
        flex_drivers = {label_to_key[l] for l in driver_labels}
    return default_case, override_pct, flex_drivers


# Direction a "favorable" (Best Case) move takes each lever, relative to base.
# Revenue growth & EBITDA margin: UP is favorable. WACC, CapEx ratio, and
# Working Capital: DOWN is favorable (cheaper capital / less cash tied up).
_SCENARIO_DRIVER_DIRECTION = {
    'growth': +1,
    'ebitda_margin': +1,
    'wacc': -1,
    'capex_ratio': -1,
    'wc_pct': -1,
}
_SCENARIO_DRIVER_LABELS = {
    'growth': 'Revenue Growth',
    'ebitda_margin': 'EBITDA Margin',
    'wacc': 'WACC',
    'capex_ratio': 'CapEx Ratio',
    'wc_pct': 'Working Capital (% Rev)',
}


def _scenario_driver_value(base_value, key, case_sign, override_pct, flex_drivers):
    """case_sign: +1 = Best Case, -1 = Worst Case, 0 = Base Case (or driver not flexed)."""
    if case_sign == 0 or key not in flex_drivers:
        return base_value
    direction = _SCENARIO_DRIVER_DIRECTION[key]
    factor = 1 + (case_sign * direction * override_pct / 100.0)
    return base_value * factor


def _run_scenario_point(financials, wc_metrics, years, tax_rate,
                         growth, ebitda_margin, wacc_pct, capex_ratio, wc_pct,
                         terminal_growth, num_shares, cash_balance,
                         depreciation_rate_override=None, depreciation_method="Auto",
                         inventory_days_override=None, debtor_days_override=None,
                         creditor_days_override=None, interest_rate_override=None):
    """Runs ONE DCF point for an explicit set of driver values (not factors)."""
    try:
        proj, drv = project_financials(
            financials, wc_metrics, years, tax_rate,
            rev_growth_override=growth,
            opex_margin_override=None,
            capex_ratio_override=capex_ratio,
            ebitda_margin_override=ebitda_margin,
            depreciation_rate_override=depreciation_rate_override,
            depreciation_method=depreciation_method,
            inventory_days_override=inventory_days_override,
            debtor_days_override=debtor_days_override,
            creditor_days_override=creditor_days_override,
            interest_rate_override=interest_rate_override,
            working_capital_pct_override=wc_pct
        )
        # The DCF reads debt from wacc_details['debt'] (default 0). Passing only wacc/tax_rate made every
        # scenario / tornado / Monte-Carlo valuation ignore the company's debt, so 'Base Case' did not
        # reconcile with the headline DCF. Use the same debt definition as calculate_wacc().
        _std = financials['st_debt'][0] if financials['st_debt'][0] > 0 else 0
        _ltd = financials['lt_debt'][0] if financials['lt_debt'][0] > 0 else 0
        val, err = calculate_dcf_valuation(
            proj, {'wacc': wacc_pct, 'tax_rate': tax_rate, 'debt': _std + _ltd}, terminal_growth,
            num_shares, cash_balance, manual_discount_rate=wacc_pct
        )
        return proj, drv, val, err
    except Exception as e:
        return None, None, None, str(e)


def calculate_scenario_dcf(financials, wc_metrics, years, tax_rate,
                            base_drivers, override_pct, flex_drivers,
                            terminal_growth, num_shares, cash_balance,
                            fixed_overrides=None):
    """
    Runs the combined Base / Best / Worst DCF, moving every lever in
    flex_drivers TOGETHER (this is what makes it scenario analysis rather
    than one-variable-at-a-time sensitivity analysis).

    base_drivers: dict with keys 'growth', 'ebitda_margin', 'wacc',
    'capex_ratio', 'wc_pct' — the base-case value for each lever.
    fixed_overrides: dict of the non-flexed project_financials kwargs
    (depreciation, WC days, interest rate) held constant across all cases.
    """
    fixed_overrides = fixed_overrides or {}
    case_signs = {'Best Case': +1, 'Base Case': 0, 'Worst Case': -1}
    results = {}
    for case_name, sign in case_signs.items():
        used = {
            key: _scenario_driver_value(base_drivers[key], key, sign, override_pct, flex_drivers)
            for key in base_drivers
        }
        proj, drv, val, err = _run_scenario_point(
            financials, wc_metrics, years, tax_rate,
            used['growth'], used['ebitda_margin'], used['wacc'], used['capex_ratio'], used['wc_pct'],
            terminal_growth, num_shares, cash_balance, **fixed_overrides
        )
        results[case_name] = {
            'projections': proj, 'drivers': drv, 'valuation': val, 'error': err,
            'used': used,
        }
    return results


def calculate_driver_tornado(financials, wc_metrics, years, tax_rate,
                              base_drivers, override_pct, flex_drivers,
                              terminal_growth, num_shares, cash_balance,
                              fixed_overrides=None):
    """
    Isolates each lever ONE AT A TIME (all others held at base) to show how
    much of the Base→Best/Worst swing each individual driver is responsible
    for — a classic tornado chart. Complements the combined scenario view.
    Returns a list of dicts sorted by swing size, largest first.
    """
    fixed_overrides = fixed_overrides or {}
    tornado = []
    for key in ['growth', 'ebitda_margin', 'wacc', 'capex_ratio', 'wc_pct']:
        if key not in flex_drivers:
            continue
        favorable = dict(base_drivers)
        unfavorable = dict(base_drivers)
        favorable[key] = _scenario_driver_value(base_drivers[key], key, +1, override_pct, flex_drivers)
        unfavorable[key] = _scenario_driver_value(base_drivers[key], key, -1, override_pct, flex_drivers)

        _, _, val_up, err_up = _run_scenario_point(
            financials, wc_metrics, years, tax_rate,
            favorable['growth'], favorable['ebitda_margin'], favorable['wacc'],
            favorable['capex_ratio'], favorable['wc_pct'],
            terminal_growth, num_shares, cash_balance, **fixed_overrides
        )
        _, _, val_down, err_down = _run_scenario_point(
            financials, wc_metrics, years, tax_rate,
            unfavorable['growth'], unfavorable['ebitda_margin'], unfavorable['wacc'],
            unfavorable['capex_ratio'], unfavorable['wc_pct'],
            terminal_growth, num_shares, cash_balance, **fixed_overrides
        )
        if val_up and val_down and not err_up and not err_down:
            fv_up = val_up['fair_value_per_share']
            fv_down = val_down['fair_value_per_share']
            tornado.append({
                'driver': _SCENARIO_DRIVER_LABELS[key],
                'favorable_fv': fv_up,
                'unfavorable_fv': fv_down,
                'swing': abs(fv_up - fv_down)
            })
    tornado.sort(key=lambda x: x['swing'], reverse=True)
    return tornado


def render_scenario_analysis_tab(financials, wc_metrics, years, tax_rate,
                                  base_drivers, override_pct, flex_drivers, default_case,
                                  terminal_growth, num_shares, cash_balance,
                                  csym, current_price=0, key_prefix="scenario",
                                  fixed_overrides=None, base_fair_value=None):
    """
    Scenario Analysis results view: comparison table, fair-value chart,
    driver radar chart, and a tornado chart isolating each lever's swing.
    The case / adjustment % / which-levers-to-flex choices are made ABOVE
    the Run Valuation button (via render_scenario_settings_ui) — this tab
    just displays the outcome and lets you browse the 3 already-computed
    cases without re-running anything.
    """
    st.subheader("🎭 Scenario Analysis")
    st.caption(
        "Moves Revenue Growth, EBITDA Margin, WACC, CapEx Ratio and Working "
        "Capital **together** under Base, Best and Worst case assumptions — "
        "unlike Sensitivity Analysis, which flexes WACC and terminal growth "
        "one at a time. Adjust the swing % and which levers move under "
        "**🎭 Scenario Analysis Settings**, above the Run Valuation button."
    )

    if not flex_drivers:
        st.warning("No levers are selected to flex. Enable at least one under Scenario Analysis Settings above.")
        return

    scenario_results = calculate_scenario_dcf(
        financials, wc_metrics, years, tax_rate,
        base_drivers, override_pct, flex_drivers,
        terminal_growth, num_shares, cash_balance,
        fixed_overrides=fixed_overrides
    )

    # ---------- Comparison table ----------
    st.markdown("#### 📋 Scenario Comparison")
    driver_order = ['growth', 'ebitda_margin', 'wacc', 'capex_ratio', 'wc_pct']
    summary_rows = []
    for case_name in ["Best Case", "Base Case", "Worst Case"]:
        res = scenario_results.get(case_name, {})
        val = res.get('valuation')
        row = {'Scenario': case_name}
        for key in driver_order:
            label = _SCENARIO_DRIVER_LABELS[key]
            if key in flex_drivers:
                row[label] = f"{res['used'][key]:.2f}%"
            else:
                row[label] = "— (fixed)"
        if val and not res.get('error'):
            fv = val['fair_value_per_share']
            row['Fair Value / Share'] = f"{csym}{fv:.2f}"
            row['vs CMP'] = f"{((fv - current_price) / current_price * 100):+.1f}%" if current_price and current_price > 0 else "—"
        else:
            row['Fair Value / Share'] = "Error"
            row['vs CMP'] = "—"
        summary_rows.append(row)
    summary_df = pd.DataFrame(summary_rows)

    def _highlight_scenario(row):
        color = {'Best Case': 'background-color: #d4f7dc', 'Worst Case': 'background-color: #fbd8d8'}.get(row['Scenario'], '')
        return [color] * len(row)
    try:
        st.dataframe(summary_df.style.apply(_highlight_scenario, axis=1), use_container_width=True, hide_index=True)
    except Exception:
        st.dataframe(summary_df, use_container_width=True, hide_index=True)

    # ---------- Selected case detail ----------
    view_case = st.radio(
        "View detail for:", ["Best Case", "Base Case", "Worst Case"],
        index=["Best Case", "Base Case", "Worst Case"].index(default_case),
        horizontal=True, key=f"{key_prefix}_view_case"
    )
    selected_res = scenario_results.get(view_case)

    if selected_res and selected_res.get('valuation') and not selected_res.get('error'):
        sel_val = selected_res['valuation']
        st.markdown(f"#### 🎯 {view_case} Detail")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            delta = None
            if current_price and current_price > 0:
                delta = f"{((sel_val['fair_value_per_share'] - current_price) / current_price * 100):+.1f}%"
            st.metric("Fair Value / Share", f"{csym}{sel_val['fair_value_per_share']:.2f}", delta=delta)
        with c2:
            st.metric("Enterprise Value", f"{csym}{sel_val.get('enterprise_value', 0):.1f} Lacs")
        with c3:
            st.metric("Equity Value", f"{csym}{sel_val.get('equity_value', 0):.1f} Lacs")
        with c4:
            st.metric("Discount Rate Used", f"{selected_res['used']['wacc']:.2f}%")

        # ---------- Fair value comparison chart ----------
        case_order = ["Worst Case", "Base Case", "Best Case"]
        fvs = []
        for c in case_order:
            r = scenario_results.get(c)
            fvs.append(r['valuation']['fair_value_per_share'] if r and r.get('valuation') and not r.get('error') else 0)
        fig_sc = go.Figure()
        fig_sc.add_trace(go.Bar(
            x=case_order, y=fvs,
            marker_color=['#e74c3c', '#3498db', '#2ecc71'],
            marker_line_color='#2c3e50', marker_line_width=1.5,
            text=[f"{csym}{v:.2f}" for v in fvs], textposition='outside',
            name='Fair Value / Share'
        ))
        if current_price and current_price > 0:
            fig_sc.add_hline(y=current_price, line_dash="dash", line_color="black", line_width=2,
                              annotation_text=f"Current Price: {csym}{current_price:.2f}",
                              annotation_position="top left")
        fig_sc.update_layout(
            title="Fair Value Across Scenarios",
            yaxis_title=f"Fair Value ({csym})", showlegend=False,
            plot_bgcolor='rgba(240,240,240,0.4)', height=380
        )
        st.plotly_chart(fig_sc, use_container_width=True)

        # ---------- Driver radar chart ----------
        st.markdown("#### 🕸️ Driver Assumptions by Scenario")
        radar_keys = [k for k in driver_order if k in flex_drivers]
        if len(radar_keys) >= 3:
            fig_radar = go.Figure()
            radar_colors = {'Worst Case': '#e74c3c', 'Base Case': '#3498db', 'Best Case': '#2ecc71'}
            base_vals = base_drivers
            for c in ["Worst Case", "Base Case", "Best Case"]:
                r = scenario_results.get(c)
                if not r or r.get('error'):
                    continue
                # Normalize each lever as % of its own base value so all axes are comparable
                vals = []
                for k in radar_keys:
                    b = base_vals[k]
                    used_v = r['used'][k]
                    vals.append((used_v / b * 100) if b else 100)
                vals.append(vals[0])
                labels = [_SCENARIO_DRIVER_LABELS[k] for k in radar_keys] + [_SCENARIO_DRIVER_LABELS[radar_keys[0]]]
                fig_radar.add_trace(go.Scatterpolar(
                    r=vals, theta=labels, fill='toself', name=c,
                    line_color=radar_colors[c], opacity=0.6
                ))
            fig_radar.update_layout(
                polar=dict(radialaxis=dict(visible=True, title="% of Base Case")),
                showlegend=True, height=420,
                title="Each Lever as % of its Base-Case Value"
            )
            st.plotly_chart(fig_radar, use_container_width=True)
        else:
            st.caption("Select at least 3 levers under Scenario Analysis Settings to see the radar chart.")

        # ---------- Tornado chart: which lever moves the needle most ----------
        st.markdown("#### 🌪️ Driver Impact (Tornado Chart)")
        st.caption("Each bar isolates ONE lever at a time (others held at base) to show its individual swing on fair value.")
        tornado = calculate_driver_tornado(
            financials, wc_metrics, years, tax_rate,
            base_drivers, override_pct, flex_drivers,
            terminal_growth, num_shares, cash_balance,
            fixed_overrides=fixed_overrides
        )
        if tornado:
            drivers_t = [t['driver'] for t in tornado]
            base_fv_t = base_fair_value if base_fair_value else scenario_results['Base Case']['valuation']['fair_value_per_share']
            low = [min(t['favorable_fv'], t['unfavorable_fv']) for t in tornado]
            high = [max(t['favorable_fv'], t['unfavorable_fv']) for t in tornado]
            fig_tor = go.Figure()
            fig_tor.add_trace(go.Bar(
                y=drivers_t, x=[h - l for h, l in zip(high, low)], base=low,
                orientation='h', marker_color='#f39c12', marker_line_color='#c9770a', marker_line_width=1,
                text=[f"{csym}{l:.2f} – {csym}{h:.2f}" for l, h in zip(low, high)],
                textposition='inside', name='Fair Value Range'
            ))
            fig_tor.add_vline(x=base_fv_t, line_dash="dash", line_color="#2c3e50", line_width=2,
                               annotation_text=f"Base: {csym}{base_fv_t:.2f}")
            fig_tor.update_layout(
                title="Fair Value Swing by Lever (Base ± Adjustment %)",
                xaxis_title=f"Fair Value ({csym})", height=max(300, len(drivers_t) * 70),
                plot_bgcolor='rgba(240,240,240,0.4)'
            )
            st.plotly_chart(fig_tor, use_container_width=True)
        else:
            st.caption("No levers selected — enable at least one under Scenario Analysis Settings.")
    else:
        err_msg = selected_res.get('error') if selected_res else 'Unknown error'
        st.error(f"Could not compute {view_case}: {err_msg}")


def render_football_field_chart(methods_data, current_price, csym, key_prefix="football",
                                 chart_title="Football Field — Valuation Range by Method"):
    """
    Standard banker's "football field" chart: a horizontal bar per valuation
    method spanning its low-high range, with a tick mark at the midpoint and
    a vertical line at the current market price for reference.

    methods_data: list of dicts, each with:
        'method' (str label), 'low' (float), 'high' (float),
        and optionally 'mid' (float — defaults to the low/high midpoint).
    Methods with low == high (a single point estimate, e.g. no range could
    be computed) are still drawn as a thin marker so they aren't silently
    dropped.
    """
    if not methods_data:
        st.info("Run more than one valuation method (e.g. DDM, RIM, or Comparable Valuation alongside DCF) to populate the football field — it needs at least one method's range to draw.")
        return

    # Keep a consistent top-to-bottom order matching the order passed in,
    # since Plotly draws categorical y-axes bottom-to-top by default.
    methods_data = list(reversed(methods_data))

    fig = go.Figure()
    for m in methods_data:
        low, high = m['low'], m['high']
        if high < low:
            low, high = high, low
        mid = m.get('mid', (low + high) / 2)
        span = max(high - low, 0.01)  # avoid a zero-width bar being invisible
        fig.add_trace(go.Bar(
            y=[m['method']], x=[span], base=[low],
            orientation='h',
            marker_color='rgba(52, 152, 219, 0.55)',
            marker_line_color='#2980b9', marker_line_width=1.5,
            customdata=[[low, high]],
            hovertemplate=f"<b>{m['method']}</b><br>Low: {csym}%{{customdata[0]:.2f}}<br>High: {csym}%{{customdata[1]:.2f}}<extra></extra>",
            showlegend=False
        ))
        fig.add_trace(go.Scatter(
            y=[m['method']], x=[mid], mode='markers',
            marker=dict(symbol='line-ns', size=20, color='#1a1a1a', line=dict(width=2)),
            showlegend=False,
            hovertemplate=f"<b>{m['method']} — Midpoint</b><br>{csym}%{{x:.2f}}<extra></extra>"
        ))

    if current_price and current_price > 0:
        fig.add_vline(x=current_price, line_dash="dash", line_color="#e74c3c", line_width=2,
                       annotation_text=f"Current Price: {csym}{current_price:.2f}",
                       annotation_position="top")

    fig.update_layout(
        title=chart_title,
        xaxis_title=f"Fair Value per Share ({csym})",
        height=max(320, 90 * len(methods_data)),
        plot_bgcolor='rgba(240,240,240,0.4)',
        margin=dict(l=10, r=10, t=60, b=40)
    )
    st.plotly_chart(fig, use_container_width=True, key=f"{key_prefix}_chart")
    st.caption("The vertical black tick on each bar is that method's point/average estimate; the dashed red line is the current market price.")


def run_monte_carlo_dcf(financials, wc_metrics, years, tax_rate,
                         base_drivers, volatility_pct, flex_drivers,
                         terminal_growth, num_shares, cash_balance,
                         n_iterations=1000, distribution="Normal", seed=None,
                         fixed_overrides=None):
    """
    Runs n_iterations DCF valuations, sampling each driver in flex_drivers
    from a distribution centered on its base-case value with the given
    volatility (std dev, as % of the base value). Unlike Scenario Analysis's
    3 fixed points, this builds a full fair-value distribution. Returns
    (fair_values: np.array, error_count: int, total_run: int).
    """
    fixed_overrides = fixed_overrides or {}
    rng = np.random.default_rng(seed)
    fair_values = []
    errors = 0

    for _ in range(n_iterations):
        sampled = {}
        for key, base_val in base_drivers.items():
            if key not in flex_drivers:
                sampled[key] = base_val
                continue
            std = abs(base_val) * (volatility_pct / 100.0)
            if std <= 0:
                std = max(abs(base_val), 1.0) * 0.05  # tiny floor so sampling isn't degenerate at base_val == 0
            if distribution == "Normal":
                val = rng.normal(base_val, std)
            elif distribution == "Triangular":
                spread = std * 2
                val = rng.triangular(base_val - spread, base_val, base_val + spread)
            else:  # Uniform
                spread = std * 1.73  # ~matches std dev of a uniform distribution
                val = rng.uniform(base_val - spread, base_val + spread)
            # Sanity clips so sampled draws stay in a sensible range
            if key == 'wacc':
                val = max(val, 1.0)
            elif key == 'ebitda_margin':
                val = max(min(val, 95.0), -50.0)
            elif key == 'capex_ratio':
                val = max(val, 0.0)
            sampled[key] = val

        _, _, val_result, err = _run_scenario_point(
            financials, wc_metrics, years, tax_rate,
            sampled['growth'], sampled['ebitda_margin'], sampled['wacc'],
            sampled['capex_ratio'], sampled['wc_pct'],
            terminal_growth, num_shares, cash_balance, **fixed_overrides
        )
        fv = val_result.get('fair_value_per_share') if val_result else None
        if fv is not None and not err and np.isfinite(fv):
            fair_values.append(fv)
        else:
            errors += 1

    return np.array(fair_values), errors, n_iterations


def render_monte_carlo_tab(financials, wc_metrics, years, tax_rate,
                            base_drivers, terminal_growth, num_shares, cash_balance,
                            csym, current_price=0, key_prefix="mc",
                            fixed_overrides=None, base_fair_value=None):
    """
    Monte Carlo Simulation tab: settings (all user-adjustable — iteration
    count, volatility, distribution shape, which levers randomize, and an
    optional fixed seed for reproducibility), a Run button (simulation is
    gated behind an explicit click so it doesn't re-run on every unrelated
    Streamlit interaction), and — once run — a fair-value distribution
    histogram with percentiles and probability of upside vs current price.
    """
    st.subheader("🎲 Monte Carlo Simulation")
    st.caption(
        "Runs the DCF thousands of times, sampling Revenue Growth, EBITDA Margin, WACC, "
        "CapEx Ratio and Working Capital from a probability distribution around their base-case "
        "values — instead of the 3 fixed points in Scenario Analysis — to show the full range and "
        "likelihood of outcomes."
    )

    with st.expander("🎲 Monte Carlo Settings", expanded=False):
        col_a, col_b = st.columns(2)
        with col_a:
            n_iterations = st.number_input(
                "Number of simulations:", min_value=200, max_value=5000, value=1000, step=100,
                key=f"{key_prefix}_n_iter",
                help="More simulations = smoother distribution but slower to run. 1,000 is usually plenty; "
                     "keep it lower if you're on a slow connection or the app feels sluggish."
            )
            volatility_pct = st.number_input(
                "Volatility — Std Dev (% of base value):", min_value=1.0, max_value=100.0, value=15.0, step=1.0,
                key=f"{key_prefix}_vol",
                help="How much each selected lever is allowed to swing, as a standard deviation relative "
                     "to its base-case value. Higher = wider, more uncertain distribution."
            )
        with col_b:
            distribution = st.selectbox(
                "Distribution shape:", ["Normal", "Triangular", "Uniform"], index=0,
                key=f"{key_prefix}_dist",
                help="Normal = classic bell curve (most common default). Triangular = bounded, peaked at "
                     "the base case. Uniform = every value in range equally likely (widest/most conservative)."
            )
            use_seed = st.checkbox("Fix random seed (reproducible results)", value=False, key=f"{key_prefix}_use_seed")
            seed = None
            if use_seed:
                seed = int(st.number_input("Seed:", min_value=0, max_value=999999, value=42, step=1, key=f"{key_prefix}_seed"))
        driver_labels = st.multiselect(
            "Levers to randomize:",
            ["Revenue Growth", "EBITDA Margin", "WACC", "CapEx Ratio", "Working Capital"],
            default=["Revenue Growth", "EBITDA Margin", "WACC", "CapEx Ratio", "Working Capital"],
            key=f"{key_prefix}_drivers"
        )
        label_to_key = {
            'Revenue Growth': 'growth', 'EBITDA Margin': 'ebitda_margin',
            'WACC': 'wacc', 'CapEx Ratio': 'capex_ratio', 'Working Capital': 'wc_pct'
        }
        flex_drivers_mc = {label_to_key[l] for l in driver_labels}

    run_clicked = st.button("🎲 Run Monte Carlo Simulation", key=f"{key_prefix}_run_btn", type="primary")
    results_key = f"{key_prefix}_results"

    if run_clicked:
        if not flex_drivers_mc:
            st.warning("Select at least one lever to randomize under Monte Carlo Settings.")
        else:
            with st.spinner(f"Running {int(n_iterations):,} DCF simulations..."):
                fair_values, error_count, total = run_monte_carlo_dcf(
                    financials, wc_metrics, years, tax_rate,
                    base_drivers, volatility_pct, flex_drivers_mc,
                    terminal_growth, num_shares, cash_balance,
                    n_iterations=int(n_iterations), distribution=distribution, seed=seed,
                    fixed_overrides=fixed_overrides
                )
            st.session_state[results_key] = {
                'fair_values': fair_values, 'errors': error_count, 'total': total
            }

    res = st.session_state.get(results_key)
    if res and res['fair_values'] is not None and len(res['fair_values']) > 0:
        fvs = res['fair_values']
        if res['errors'] > 0:
            st.caption(f"⚠️ {res['errors']} of {res['total']} simulations failed and were excluded (e.g. negative-FCF edge cases from extreme sampled draws).")

        p10, p25, p50, p75, p90 = np.percentile(fvs, [10, 25, 50, 75, 90])
        mean_fv, std_fv = np.mean(fvs), np.std(fvs)

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("P10 (Bear)", f"{csym}{p10:.2f}")
        c2.metric("P25", f"{csym}{p25:.2f}")
        c3.metric("P50 (Median)", f"{csym}{p50:.2f}")
        c4.metric("P75", f"{csym}{p75:.2f}")
        c5.metric("P90 (Bull)", f"{csym}{p90:.2f}")
        st.caption(f"Mean: {csym}{mean_fv:.2f} | Std Dev: {csym}{std_fv:.2f} | Valid simulations: {len(fvs):,}")

        fig_hist = go.Figure()
        fig_hist.add_trace(go.Histogram(
            x=fvs, nbinsx=50, marker_color='rgba(52,152,219,0.7)',
            marker_line_color='#2980b9', marker_line_width=1, name='Fair Value Distribution'
        ))
        if base_fair_value:
            fig_hist.add_vline(x=base_fair_value, line_dash="dash", line_color="#2c3e50", line_width=2,
                                annotation_text=f"Base Case: {csym}{base_fair_value:.2f}")
        if current_price and current_price > 0:
            fig_hist.add_vline(x=current_price, line_dash="dot", line_color="#e74c3c", line_width=2,
                                annotation_text=f"Current Price: {csym}{current_price:.2f}")
        fig_hist.update_layout(
            title="Fair Value Distribution (Monte Carlo)",
            xaxis_title=f"Fair Value per Share ({csym})", yaxis_title="Frequency",
            height=420, plot_bgcolor='rgba(240,240,240,0.4)'
        )
        st.plotly_chart(fig_hist, use_container_width=True)

        if current_price and current_price > 0:
            prob_upside = (fvs > current_price).mean() * 100
            st.info(f"📊 Based on this simulation, there's a **{prob_upside:.1f}%** probability the fair value exceeds the current market price of {csym}{current_price:.2f}.")

        if st.button("🔄 Clear results", key=f"{key_prefix}_clear_btn"):
            del st.session_state[results_key]
            st.rerun()
    else:
        st.info("💡 Configure the settings above and click **Run Monte Carlo Simulation** to generate the fair-value distribution.")


def build_multistage_growth_array(total_years, stage1_years, stage1_growth, stage2_years, terminal_growth,
                                   stage1_margin=None, stage2_margin=None):
    """
    Builds per-year revenue growth (and optionally EBITDA margin) arrays for
    a classic 3-stage profile instead of one flat rate cliffing straight
    into the terminal value:
      Stage 1 (High Growth): stage1_years at stage1_growth (or stage1_margin)
      Stage 2 (Fade):         stage2_years linearly interpolating down to
                               terminal_growth (or stage2_margin, the mature level)
      Stage 3 (Stable):       remaining years held flat at terminal_growth
                               (or stage2_margin)
    Returns (growth_array, margin_array_or_None), both length total_years.
    """
    fade_margins = stage1_margin is not None and stage2_margin is not None
    growth_array, margin_array = [], [] if fade_margins else None

    for yr in range(1, total_years + 1):
        if yr <= stage1_years:
            g = stage1_growth
            m = stage1_margin if fade_margins else None
        elif yr <= stage1_years + stage2_years:
            progress = (yr - stage1_years) / max(stage2_years, 1)
            g = stage1_growth + (terminal_growth - stage1_growth) * progress
            m = (stage1_margin + (stage2_margin - stage1_margin) * progress) if fade_margins else None
        else:
            g = terminal_growth
            m = stage2_margin if fade_margins else None
        growth_array.append(round(g, 3))
        if fade_margins:
            margin_array.append(round(m, 3))

    return growth_array, margin_array


def render_multistage_growth_ui(key_prefix, total_years, default_stage1_growth, terminal_growth,
                                 default_stage1_margin=None):
    """
    Multi-Stage DCF growth profile settings — an alternative to a single flat
    growth rate or hand-typing every year's growth/margin. All stage lengths
    and rates are adjustable; nothing here is a fixed/hardcoded assumption.
    Returns (enabled: bool, growth_array: list|None, margin_array: list|None)
    — both arrays are None when disabled, so the caller falls back to
    whatever single-override or manual per-year values it already had.
    """
    with st.expander("🎯 Multi-Stage Growth Profile (High-Growth → Fade → Stable)", expanded=False):
        st.caption(
            "Models a high-growth phase that gradually fades to the terminal growth rate, instead of one "
            "flat growth rate cliffing straight into the terminal value — closer to how this is actually "
            "modeled for high-growth or turnaround companies. When enabled, this overrides the single "
            "Revenue Growth override and the manual Year-by-Year entries above."
        )
        use_multistage = st.checkbox("Use Multi-Stage Growth Profile", value=False, key=f"{key_prefix}_enable")
        if not use_multistage:
            return False, None, None

        total_years = int(total_years)
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown("**Stage 1: High Growth**")
            stage1_years = int(st.number_input(
                "Years", min_value=1, max_value=max(total_years - 1, 1),
                value=min(3, max(total_years - 1, 1)), step=1, key=f"{key_prefix}_s1_years"
            ))
            stage1_growth = st.number_input(
                "Growth Rate (%)", min_value=-50.0, max_value=200.0,
                value=float(default_stage1_growth), step=0.5, key=f"{key_prefix}_s1_growth"
            )
        with col2:
            st.markdown("**Stage 2: Fade**")
            max_fade_years = max(total_years - stage1_years, 0)
            stage2_years = int(st.number_input(
                "Years", min_value=0, max_value=max_fade_years,
                value=min(2, max_fade_years), step=1, key=f"{key_prefix}_s2_years"
            ))
            st.caption(f"Fades linearly from {stage1_growth:.1f}% down to the terminal growth rate ({terminal_growth:.1f}%) over these years.")
        with col3:
            st.markdown("**Stage 3: Stable**")
            stage3_years = max(total_years - stage1_years - stage2_years, 0)
            st.metric("Years", stage3_years)
            st.caption(f"Held flat at the terminal growth rate ({terminal_growth:.1f}%).")

        use_margin_fade = False
        stage1_margin = mature_margin = None
        if default_stage1_margin is not None:
            use_margin_fade = st.checkbox(
                "Also fade EBITDA Margin toward a mature/long-run level",
                value=False, key=f"{key_prefix}_margin_enable",
                help="Competition typically compresses super-normal margins over time — this lets the "
                     "EBITDA margin fade alongside growth instead of staying flat at today's level."
            )
            if use_margin_fade:
                colm1, colm2 = st.columns(2)
                with colm1:
                    stage1_margin = st.number_input(
                        "Stage 1 EBITDA Margin (%)", min_value=0.0, max_value=95.0,
                        value=float(default_stage1_margin), step=0.5, key=f"{key_prefix}_s1_margin"
                    )
                with colm2:
                    mature_margin = st.number_input(
                        "Mature/Stable EBITDA Margin (%)", min_value=0.0, max_value=95.0,
                        value=float(default_stage1_margin), step=0.5, key=f"{key_prefix}_mature_margin"
                    )

        growth_array, margin_array = build_multistage_growth_array(
            total_years, stage1_years, stage1_growth, stage2_years, terminal_growth,
            stage1_margin=stage1_margin if use_margin_fade else None,
            stage2_margin=mature_margin if use_margin_fade else None
        )

        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=list(range(1, total_years + 1)), y=growth_array, mode='lines+markers',
            name='Revenue Growth %', line=dict(color='#3498db', width=3)
        ))
        layout_kwargs = dict(
            title="Multi-Stage Growth Profile Preview", xaxis_title="Projection Year",
            yaxis_title="Revenue Growth %", height=320, plot_bgcolor='rgba(240,240,240,0.4)'
        )
        if margin_array:
            fig.add_trace(go.Scatter(
                x=list(range(1, total_years + 1)), y=margin_array, mode='lines+markers',
                name='EBITDA Margin %', line=dict(color='#e67e22', width=3), yaxis='y2'
            ))
            layout_kwargs['yaxis2'] = dict(title='EBITDA Margin %', overlaying='y', side='right')
        fig.update_layout(**layout_kwargs)
        st.plotly_chart(fig, use_container_width=True, key=f"{key_prefix}_preview_chart")

        return True, growth_array, (margin_array if margin_array else None)


def render_implied_multiples_panel(valuation, financials, shares, comp_results=None, current_price=0, key_prefix="implied"):
    """
    Backs out what P/E, P/B, P/S and EV/EBITDA multiple the DCF fair value
    implies, and puts it side by side with actual peer trading multiples
    (from Comparable Valuation, if run) and the company's own current
    multiple (if a market price is available) — the standard triangulation
    sanity check: does the DCF's implicit multiple look reasonable next to
    how the market actually prices similar companies?
    """
    st.markdown("#### 🔄 Implied Multiples — DCF vs Market")
    st.caption("What P/E / P/B / P/S / EV-EBITDA multiple does the DCF fair value imply, and how does that compare to actual peer multiples?")

    if not valuation or valuation.get('fair_value_per_share', 0) <= 0:
        st.info("Run a DCF valuation to see its implied multiples.")
        return

    fair_value = valuation['fair_value_per_share']
    ev = valuation.get('enterprise_value', 0)

    net_income = ensure_valid_number(financials['nopat'][0] * 100000, 0) if financials.get('nopat') else 0
    equity_val = ensure_valid_number(financials['equity'][0] * 100000, 0) if financials.get('equity') else 0
    revenue_val = ensure_valid_number(financials['revenue'][0] * 100000, 0) if financials.get('revenue') else 0
    ebitda_val = financials['ebitda'][0] if financials.get('ebitda') else 0  # Lacs — same units as `ev`

    eps = safe_divide(net_income, shares, default=0) if shares else 0
    bvps = safe_divide(equity_val, shares, default=0) if shares else 0
    sales_ps = safe_divide(revenue_val, shares, default=0) if shares else 0

    implied_pe = (fair_value / eps) if eps > 0 else None
    implied_pb = (fair_value / bvps) if bvps > 0 else None
    implied_ps = (fair_value / sales_ps) if sales_ps > 0 else None
    implied_ev_ebitda = (ev / ebitda_val) if ebitda_val > 0 else None

    peer_stats = (comp_results or {}).get('multiples_stats', {})

    rows = []
    flag_candidates = []

    def _fmt(x):
        return f"{x:.2f}x" if x is not None and np.isfinite(x) else "—"

    def _add_row(label, implied, stats_key, current_val=None):
        stats = peer_stats.get(stats_key, {})
        avg = stats.get('average')
        rows.append({
            'Multiple': label,
            'DCF-Implied': _fmt(implied),
            'Peer Average': _fmt(avg),
            'Peer Median': _fmt(stats.get('median')),
            'Peer Min – Max': f"{_fmt(stats.get('min'))} – {_fmt(stats.get('max'))}" if stats.get('min') is not None else "—",
            'Company Current': _fmt(current_val)
        })
        if implied is not None and avg:
            flag_candidates.append((label, implied, avg))

    _add_row("P/E", implied_pe, 'pe', current_val=(current_price / eps if current_price and eps > 0 else None))
    _add_row("P/B", implied_pb, 'pb', current_val=(current_price / bvps if current_price and bvps > 0 else None))
    _add_row("P/S", implied_ps, 'ps', current_val=(current_price / sales_ps if current_price and sales_ps > 0 else None))
    _add_row("EV/EBITDA", implied_ev_ebitda, 'ev_ebitda')

    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    if not peer_stats:
        st.caption("💡 Run Comparative Valuation with peer tickers above to see how these implied multiples stack up against actual peer trading multiples.")
        return

    flag_threshold = st.slider(
        "Flag DCF-implied multiples that deviate from the peer average by more than:",
        min_value=10, max_value=100, value=30, step=5, format="%d%%",
        key=f"{key_prefix}_flag_threshold",
        help="Purely a display threshold for the warnings below — doesn't change any calculation."
    )
    any_flagged = False
    for label, implied, avg in flag_candidates:
        pct_diff = (implied - avg) / avg * 100
        if abs(pct_diff) > flag_threshold:
            any_flagged = True
            direction = "above" if pct_diff > 0 else "below"
            st.warning(f"⚠️ DCF-implied {label} ({implied:.2f}x) is {abs(pct_diff):.0f}% {direction} the peer average ({avg:.2f}x) — worth double-checking the growth/margin/WACC assumptions driving it.")
    if not any_flagged:
        st.success(f"✅ All DCF-implied multiples are within {flag_threshold}% of their peer averages.")


def render_ibc_liquidation_tab(financials, csym, key_prefix, dcf_equity_value=None,
                                dcf_fair_value_per_share=None, num_shares=None):
    """
    Estimates Liquidation Value under Regulation 35 of the IBBI (Insolvency
    Resolution Process for Corporate Persons) Regulations, 2016 — the
    asset-by-asset distressed/forced-sale realizable value, as distinct from
    the Fair Value (going-concern DCF) computed elsewhere in this app.
    Registered Valuer reports under IBC typically present both side by side.

    Every book value and haircut % is pre-filled from whatever this app
    already has on hand, but is fully editable — none of it is fixed. The
    Section 53 waterfall is a simplified illustration, clearly labeled as
    such, since actual distribution depends on case-specific facts (which
    creditors relinquished security, workmen dues caps, etc.).
    """
    st.subheader("⚖️ IBC Liquidation Value")
    st.caption(
        "Estimates Liquidation Value under Reg. 35 of the IBBI (Insolvency Resolution Process for "
        "Corporate Persons) Regulations, 2016 — the distressed, asset-by-asset realizable value on a "
        "forced/timeline-bound sale, distinct from the Fair Value (going-concern DCF) shown elsewhere."
    )
    st.info(
        "💡 The book values and haircuts below are pre-filled as illustrative starting points based on "
        "common registered-valuer practice in India — every engagement differs by asset condition, "
        "industry, and case facts. Edit every figure (including adding/removing rows) to match the actual "
        "asset register / valuation report for this case."
    )

    _fixed_assets = ensure_valid_number(financials['fixed_assets'][0] * 100000, 0) if financials.get('fixed_assets') else 0
    _inventory = ensure_valid_number(financials['inventory'][0] * 100000, 0) if financials.get('inventory') else 0
    _receivables = ensure_valid_number(financials['receivables'][0] * 100000, 0) if financials.get('receivables') else 0
    _cash = ensure_valid_number(financials['cash'][0] * 100000, 0) if financials.get('cash') else 0

    default_rows = [
        {"Asset Category": "Land & Building", "Book Value": round(_fixed_assets * 0.4, 0), "Haircut %": 20.0},
        {"Asset Category": "Plant & Machinery", "Book Value": round(_fixed_assets * 0.6, 0), "Haircut %": 50.0},
        {"Asset Category": "Furniture, Fixtures & Other Fixed Assets", "Book Value": 0.0, "Haircut %": 70.0},
        {"Asset Category": "Inventory (Raw Material / WIP / Finished Goods)", "Book Value": round(_inventory, 0), "Haircut %": 50.0},
        {"Asset Category": "Trade Receivables", "Book Value": round(_receivables, 0), "Haircut %": 60.0},
        {"Asset Category": "Cash & Bank Balances", "Book Value": round(_cash, 0), "Haircut %": 0.0},
        {"Asset Category": "Investments (Quoted & Unquoted)", "Book Value": 0.0, "Haircut %": 30.0},
        {"Asset Category": "Other Current Assets / Loans & Advances", "Book Value": 0.0, "Haircut %": 80.0},
        {"Asset Category": "Intangible Assets (Goodwill, Brand, IP)", "Book Value": 0.0, "Haircut %": 95.0},
    ]

    table_key = f"{key_prefix}_asset_table"
    if table_key not in st.session_state:
        st.session_state[table_key] = pd.DataFrame(default_rows)

    st.markdown("#### 📋 Asset-wise Realizable Value")
    st.caption("Add or remove rows to match the actual fixed asset register / balance sheet schedules for this case.")
    edited_df = st.data_editor(
        st.session_state[table_key],
        column_config={
            "Asset Category": st.column_config.TextColumn("Asset Category"),
            "Book Value": st.column_config.NumberColumn(f"Book Value ({csym})", min_value=0.0, format="%.0f"),
            "Haircut %": st.column_config.NumberColumn("Haircut %", min_value=0.0, max_value=100.0, format="%.1f"),
        },
        num_rows="dynamic",
        use_container_width=True,
        key=f"{key_prefix}_editor",
        hide_index=True
    )
    st.session_state[table_key] = edited_df

    calc_df = edited_df.copy()
    calc_df["Book Value"] = pd.to_numeric(calc_df["Book Value"], errors='coerce').fillna(0)
    calc_df["Haircut %"] = pd.to_numeric(calc_df["Haircut %"], errors='coerce').fillna(0).clip(0, 100)
    calc_df["Realizable Value"] = calc_df["Book Value"] * (1 - calc_df["Haircut %"] / 100)

    total_book_value = calc_df["Book Value"].sum()
    gross_realizable = calc_df["Realizable Value"].sum()

    c1, c2, c3 = st.columns(3)
    c1.metric("Total Book Value", f"{csym}{total_book_value:,.0f}")
    c2.metric("Gross Realizable Value", f"{csym}{gross_realizable:,.0f}")
    c3.metric("Overall Recovery %", f"{safe_divide(gross_realizable, total_book_value, 0) * 100:.1f}%")

    if not calc_df.empty and calc_df["Realizable Value"].sum() > 0:
        fig = go.Figure(go.Bar(
            x=calc_df["Asset Category"], y=calc_df["Realizable Value"],
            marker_color='rgba(231, 76, 60, 0.7)', marker_line_color='#c0392b', marker_line_width=1,
            text=[f"{csym}{v:,.0f}" for v in calc_df["Realizable Value"]], textposition='outside'
        ))
        fig.update_layout(
            title="Realizable Value by Asset Category", yaxis_title=f"Realizable Value ({csym})",
            height=420, plot_bgcolor='rgba(240,240,240,0.4)', xaxis_tickangle=-30
        )
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### 💸 Liquidation Process Costs")
    col_a, col_b = st.columns(2)
    with col_a:
        liq_cost_pct = st.number_input(
            "Liquidator's Fee / Process Cost (% of Gross Realizable Value):",
            min_value=0.0, max_value=20.0, value=2.0, step=0.25, key=f"{key_prefix}_liq_cost_pct",
            help="IBBI (Liquidation Process) Regulations prescribe a tiered fee schedule based on realized "
                 "value — enter your applicable effective rate here."
        )
    with col_b:
        other_cirp_costs = st.number_input(
            f"Other CIRP/Liquidation Costs ({csym}):",
            min_value=0.0, value=0.0, step=10000.0, key=f"{key_prefix}_other_costs",
            help="Any additional insolvency resolution process costs not captured by the fee % above."
        )

    liq_fee_amount = gross_realizable * (liq_cost_pct / 100)
    net_liquidation_value = max(gross_realizable - liq_fee_amount - other_cirp_costs, 0)

    st.metric("🏦 Net Liquidation Value (available for distribution)", f"{csym}{net_liquidation_value:,.0f}")

    if dcf_equity_value and dcf_equity_value > 0:
        st.markdown("---")
        st.markdown("#### ⚖️ Liquidation Value vs. Fair Value (Going Concern DCF)")
        st.caption("IBC Reg. 35 requires both to be reported — the gap between them is itself informative for resolution planning.")
        comp_col1, comp_col2, comp_col3 = st.columns(3)
        comp_col1.metric("Fair Value (DCF, Going Concern)", f"{csym}{dcf_equity_value:,.0f}")
        comp_col2.metric("Liquidation Value", f"{csym}{net_liquidation_value:,.0f}")
        gap_pct = safe_divide(dcf_equity_value - net_liquidation_value, dcf_equity_value, 0) * 100
        comp_col3.metric("Value Erosion in Liquidation", f"{gap_pct:.1f}%")
        if num_shares and num_shares > 0 and dcf_fair_value_per_share:
            st.caption(f"Per-share equivalent: Fair Value {csym}{dcf_fair_value_per_share:.2f} vs. Liquidation {csym}{safe_divide(net_liquidation_value, num_shares, 0):.2f}")

    with st.expander("🧮 Section 53 Distribution Waterfall (simplified, illustrative)", expanded=False):
        st.caption(
            "A simplified illustration of the IBC Section 53 priority waterfall. Actual distribution depends "
            "on which creditors relinquished security interests, workmen dues caps, and case-specific facts — "
            "verify against the actual resolution professional's records and legal advice before relying on this."
        )
        wf_col1, wf_col2 = st.columns(2)
        with wf_col1:
            secured_fc = st.number_input(f"Secured Financial Creditors ({csym}):", min_value=0.0, value=0.0, step=10000.0, key=f"{key_prefix}_wf_secured")
            workmen_dues = st.number_input(f"Workmen's Dues – 24 months ({csym}):", min_value=0.0, value=0.0, step=10000.0, key=f"{key_prefix}_wf_workmen")
            employee_dues = st.number_input(f"Employee Dues – 12 months ({csym}):", min_value=0.0, value=0.0, step=10000.0, key=f"{key_prefix}_wf_employee")
        with wf_col2:
            unsecured_fc = st.number_input(f"Unsecured Financial Creditors ({csym}):", min_value=0.0, value=0.0, step=10000.0, key=f"{key_prefix}_wf_unsecured")
            operational_creditors = st.number_input(f"Operational Creditors ({csym}):", min_value=0.0, value=0.0, step=10000.0, key=f"{key_prefix}_wf_opcred")
            govt_dues = st.number_input(f"Government Dues ({csym}):", min_value=0.0, value=0.0, step=10000.0, key=f"{key_prefix}_wf_govt")

        remaining = net_liquidation_value
        waterfall_rows = []

        tier1_claim = workmen_dues + secured_fc
        tier1_paid = min(remaining, tier1_claim)
        remaining -= tier1_paid
        waterfall_rows.append(("Tier 1: Workmen Dues (24mo) + Secured Financial Creditors (pari passu)", tier1_claim, tier1_paid))

        tier2_paid = min(remaining, employee_dues)
        remaining -= tier2_paid
        waterfall_rows.append(("Tier 2: Employee Dues (12mo)", employee_dues, tier2_paid))

        tier3_paid = min(remaining, unsecured_fc)
        remaining -= tier3_paid
        waterfall_rows.append(("Tier 3: Unsecured Financial Creditors", unsecured_fc, tier3_paid))

        tier4_paid = min(remaining, govt_dues)
        remaining -= tier4_paid
        waterfall_rows.append(("Tier 4: Government Dues", govt_dues, tier4_paid))

        tier5_paid = min(remaining, operational_creditors)
        remaining -= tier5_paid
        waterfall_rows.append(("Tier 5: Operational Creditors", operational_creditors, tier5_paid))

        waterfall_rows.append(("Remaining (to Preference/Equity Shareholders)", None, max(remaining, 0)))

        wf_df = pd.DataFrame(waterfall_rows, columns=["Priority Class", "Claim Amount", "Amount Recovered"])
        wf_df["Recovery %"] = wf_df.apply(
            lambda r: f"{safe_divide(r['Amount Recovered'], r['Claim Amount'], 0) * 100:.1f}%" if r["Claim Amount"] else "—",
            axis=1
        )
        wf_df["Claim Amount"] = wf_df["Claim Amount"].apply(lambda v: f"{csym}{v:,.0f}" if v is not None else "—")
        wf_df["Amount Recovered"] = wf_df["Amount Recovered"].apply(lambda v: f"{csym}{v:,.0f}")
        st.dataframe(wf_df, use_container_width=True, hide_index=True)


def _ratio_interpretation(name, value, csym="₹"):
    """Plain-English, number-substituted explanation of what a ratio actually means."""
    interpretations = {
        'P/E': f"You're paying {csym}{value:.2f} for every {csym}1 of current annual earnings per share.",
        'P/B': f"You're paying {csym}{value:.2f} for every {csym}1 of net worth (book value) on the balance sheet.",
        'P/S': f"You're paying {csym}{value:.2f} for every {csym}1 of annual revenue.",
        'EV/EBITDA': f"The business is priced at {value:.2f}x its annual EBITDA — a yardstick for comparing companies regardless of debt levels or tax differences.",
        'Gross Margin': f"For every {csym}100 of revenue, {csym}{value:.2f} is left after direct costs of production (before overheads, interest and tax).",
        'EBITDA Margin': f"For every {csym}100 of revenue, {csym}{value:.2f} is left as core operating cash profit, before depreciation, interest and tax.",
        'Net Margin': f"For every {csym}100 of revenue, {csym}{value:.2f} ultimately reaches the bottom line as net profit.",
        'Current Ratio': f"The company has {csym}{value:.2f} of current assets for every {csym}1 of current liabilities — generally, above 1x signals it can cover near-term obligations.",
        'Quick Ratio': f"Excluding inventory, the company has {csym}{value:.2f} of liquid assets for every {csym}1 of current liabilities.",
        'Cash Ratio': f"The company holds {csym}{value:.2f} of pure cash for every {csym}1 of current liabilities — the most conservative liquidity measure.",
        'Debt-to-Equity': f"For every {csym}1 of shareholder equity, the company carries {csym}{value:.2f} of debt — higher means more reliance on borrowed capital.",
        'Debt-to-Assets': f"{value:.1f}% of the company's assets are funded by debt rather than equity.",
        'Interest Coverage': f"Operating profit (EBIT) covers interest expense {value:.2f} times over — below ~2x is often considered a solvency warning sign.",
        'ROE': f"For every {csym}100 of shareholder equity, the business generates {csym}{value:.2f} of net profit in a year.",
        'ROCE': f"For every {csym}100 of total capital employed (equity + debt), the business generates {csym}{value:.2f} of operating profit.",
        'ROA': f"For every {csym}100 of total assets, the business generates {csym}{value:.2f} of net profit.",
        'Asset Turnover': f"The company generates {csym}{value:.2f} of revenue for every {csym}1 of assets employed — higher generally means more efficient asset use.",
        'Equity Multiplier': f"Total assets are {value:.2f}x shareholder equity — the leverage component of ROE; higher means more of the asset base is debt-funded.",
        'Tax Burden': f"The company keeps {value*100:.1f}% of pre-tax profit after tax — i.e. its effective tax rate is roughly {(1-value)*100:.1f}%.",
        'Interest Burden': f"After interest expense, {value*100:.1f}% of operating profit (EBIT) remains as pre-tax profit.",
        'Operating Margin': f"For every {csym}100 of revenue, {csym}{value:.2f} remains as operating profit (EBIT), before interest and tax.",
        'Cash Conversion Cycle': f"It takes the company roughly {value:.0f} days, on net, to convert money spent on inventory back into cash from sales.",
    }
    return interpretations.get(name, "")


def render_ratio_dashboard(financials, wc_metrics, csym, key_prefix, current_price=0, shares=None, tax_rate=25.0):
    """
    Ratio & red-flag dashboard built entirely from the financials already
    ingested for this company — DuPont ROE decomposition, liquidity/solvency
    trends, efficiency, and (where a price/share count is available)
    valuation multiples — each with a plain-English, number-substituted
    interpretation. Also runs a handful of adjustable-threshold anomaly
    checks (margin swings, receivables/inventory outpacing sales growth,
    thin coverage, high leverage), plus a manual checklist for things that
    genuinely can't be detected from summary financials alone (related-party
    transactions, contingent liabilities, auditor qualifications).
    """
    st.subheader("📐 Ratio Analysis & Red Flags")
    st.caption("Built from the financials already extracted for this company — no new data entry required.")

    years = financials.get('years', [])
    n = len(years)
    if n == 0:
        st.info("No historical financials available to analyze yet.")
        return

    # Oldest → newest, for left-to-right trend charts (financials lists are newest-first)
    idx_chron = list(reversed(range(n)))
    yrs_chron = [years[i] for i in idx_chron]

    def _get(key, i):
        arr = financials.get(key, [])
        return ensure_valid_number(arr[i], 0) if i < len(arr) else 0

    total_assets, equity_l, debt_l, net_income_l, ebt_l = [], [], [], [], []
    for i in range(n):
        rev = _get('revenue', i)
        ebit = _get('ebit', i)
        interest = _get('interest', i)
        eq = _get('equity', i)
        debt = _get('st_debt', i) + _get('lt_debt', i)
        ebt = ebit - interest
        ni = ebt * (1 - tax_rate / 100)
        assets = _get('fixed_assets', i) + _get('inventory', i) + _get('receivables', i) + _get('cash', i)
        total_assets.append(assets)
        equity_l.append(eq)
        debt_l.append(debt)
        net_income_l.append(ni)
        ebt_l.append(ebt)

    st.caption(
        "Note: Total Assets is approximated as Fixed Assets + Inventory + Receivables + Cash (the asset "
        "categories this app tracks) — other items like investments or intangibles aren't separately "
        "captured, so asset-based ratios below are directional, not exact from an audited balance sheet."
    )

    # ---------- Settings (adjustable red-flag thresholds) ----------
    with st.expander("🚩 Red-Flag Thresholds (adjustable)", expanded=False):
        col_t1, col_t2 = st.columns(2)
        with col_t1:
            margin_jump_threshold = st.number_input("Flag EBITDA margin swings greater than (pp YoY):", min_value=1.0, max_value=50.0, value=8.0, step=1.0, key=f"{key_prefix}_margin_thresh")
            outpace_threshold = st.number_input("Flag receivables/inventory growth exceeding revenue growth by (pp):", min_value=5.0, max_value=100.0, value=20.0, step=5.0, key=f"{key_prefix}_outpace_thresh")
        with col_t2:
            current_ratio_threshold = st.number_input("Flag Current Ratio below:", min_value=0.5, max_value=3.0, value=1.0, step=0.1, key=f"{key_prefix}_cr_thresh")
            de_threshold = st.number_input("Flag Debt-to-Equity above:", min_value=0.5, max_value=10.0, value=2.0, step=0.5, key=f"{key_prefix}_de_thresh")
            interest_cov_threshold = st.number_input("Flag Interest Coverage below:", min_value=0.5, max_value=10.0, value=2.0, step=0.5, key=f"{key_prefix}_ic_thresh")

    # ============================================================
    # TREND OVERVIEW
    # ============================================================
    st.markdown("#### 📈 Trend Overview")
    revenue_chron = [_get('revenue', i) for i in idx_chron]
    ebitda_margin_chron = [safe_divide(_get('ebitda', i), _get('revenue', i), 0) * 100 for i in idx_chron]
    net_margin_chron = [safe_divide(net_income_l[i], _get('revenue', i), 0) * 100 for i in idx_chron]

    fig_trend = go.Figure()
    fig_trend.add_trace(go.Bar(x=yrs_chron, y=revenue_chron, name=f'Revenue ({csym} Lacs)', marker_color='rgba(52,152,219,0.6)', yaxis='y'))
    fig_trend.add_trace(go.Scatter(x=yrs_chron, y=ebitda_margin_chron, name='EBITDA Margin %', mode='lines+markers', line=dict(color='#e67e22', width=3), yaxis='y2'))
    fig_trend.add_trace(go.Scatter(x=yrs_chron, y=net_margin_chron, name='Net Margin %', mode='lines+markers', line=dict(color='#27ae60', width=3, dash='dash'), yaxis='y2'))
    fig_trend.update_layout(
        title="Revenue & Margin Trend", yaxis=dict(title=f'Revenue ({csym} Lacs)'),
        yaxis2=dict(title='Margin %', overlaying='y', side='right'),
        height=380, plot_bgcolor='rgba(240,240,240,0.4)'
    )
    st.plotly_chart(fig_trend, use_container_width=True)

    # ============================================================
    # DUPONT ROE DECOMPOSITION
    # ============================================================
    st.markdown("#### 🧮 DuPont ROE Decomposition (5-Step)")
    st.caption("ROE = Tax Burden × Interest Burden × Operating Margin × Asset Turnover × Equity Multiplier")
    dupont_rows = []
    for i in range(n):
        ebit = _get('ebit', i)
        rev = _get('revenue', i)
        tax_burden = safe_divide(net_income_l[i], ebt_l[i], 0)
        interest_burden = safe_divide(ebt_l[i], ebit, 0)
        op_margin = safe_divide(ebit, rev, 0)
        asset_turnover = safe_divide(rev, total_assets[i], 0)
        equity_multiplier = safe_divide(total_assets[i], equity_l[i], 0)
        roe = safe_divide(net_income_l[i], equity_l[i], 0)
        dupont_rows.append({
            'Year': years[i], 'Tax Burden': tax_burden, 'Interest Burden': interest_burden,
            'Operating Margin %': op_margin * 100, 'Asset Turnover': asset_turnover,
            'Equity Multiplier': equity_multiplier, 'ROE %': roe * 100
        })
    dupont_df = pd.DataFrame(dupont_rows)
    st.dataframe(
        dupont_df.style.format({
            'Tax Burden': '{:.2f}', 'Interest Burden': '{:.2f}', 'Operating Margin %': '{:.1f}%',
            'Asset Turnover': '{:.2f}x', 'Equity Multiplier': '{:.2f}x', 'ROE %': '{:.1f}%'
        }),
        use_container_width=True, hide_index=True
    )
    if dupont_rows:
        latest = dupont_rows[0]
        with st.expander("💬 What do these mean? (latest year)", expanded=False):
            st.write(f"**Tax Burden ({latest['Tax Burden']:.2f}):** {_ratio_interpretation('Tax Burden', latest['Tax Burden'], csym)}")
            st.write(f"**Interest Burden ({latest['Interest Burden']:.2f}):** {_ratio_interpretation('Interest Burden', latest['Interest Burden'], csym)}")
            st.write(f"**Operating Margin ({latest['Operating Margin %']:.1f}%):** {_ratio_interpretation('Operating Margin', latest['Operating Margin %']/100, csym)}")
            st.write(f"**Asset Turnover ({latest['Asset Turnover']:.2f}x):** {_ratio_interpretation('Asset Turnover', latest['Asset Turnover'], csym)}")
            st.write(f"**Equity Multiplier ({latest['Equity Multiplier']:.2f}x):** {_ratio_interpretation('Equity Multiplier', latest['Equity Multiplier'], csym)}")
            st.write(f"**ROE ({latest['ROE %']:.1f}%):** {_ratio_interpretation('ROE', latest['ROE %'], csym)}")

    # ============================================================
    # LIQUIDITY & SOLVENCY
    # ============================================================
    st.markdown("#### 💧 Liquidity & Solvency")
    liq_rows = []
    for i in range(n):
        cur_liab = _get('payables', i) + _get('st_debt', i)
        cur_assets = _get('inventory', i) + _get('receivables', i) + _get('cash', i)
        current_ratio = safe_divide(cur_assets, cur_liab, 0)
        quick_ratio = safe_divide(_get('receivables', i) + _get('cash', i), cur_liab, 0)
        cash_ratio = safe_divide(_get('cash', i), cur_liab, 0)
        de_ratio = safe_divide(debt_l[i], equity_l[i], 0)
        interest_cov = safe_divide(_get('ebit', i), _get('interest', i), 0) if _get('interest', i) > 0 else None
        debt_to_assets = safe_divide(debt_l[i], total_assets[i], 0) * 100
        liq_rows.append({
            'Year': years[i], 'Current Ratio': current_ratio, 'Quick Ratio': quick_ratio,
            'Cash Ratio': cash_ratio, 'Debt-to-Equity': de_ratio,
            'Interest Coverage': interest_cov, 'Debt-to-Assets %': debt_to_assets
        })
    liq_df = pd.DataFrame(liq_rows)
    st.dataframe(
        liq_df.style.format({
            'Current Ratio': '{:.2f}x', 'Quick Ratio': '{:.2f}x', 'Cash Ratio': '{:.2f}x',
            'Debt-to-Equity': '{:.2f}x', 'Interest Coverage': lambda v: f"{v:.2f}x" if v is not None else "N/A (no interest expense)",
            'Debt-to-Assets %': '{:.1f}%'
        }),
        use_container_width=True, hide_index=True
    )
    if liq_rows:
        latest_liq = liq_rows[0]
        with st.expander("💬 What do these mean? (latest year)", expanded=False):
            st.write(f"**Current Ratio ({latest_liq['Current Ratio']:.2f}x):** {_ratio_interpretation('Current Ratio', latest_liq['Current Ratio'], csym)}")
            st.write(f"**Quick Ratio ({latest_liq['Quick Ratio']:.2f}x):** {_ratio_interpretation('Quick Ratio', latest_liq['Quick Ratio'], csym)}")
            st.write(f"**Cash Ratio ({latest_liq['Cash Ratio']:.2f}x):** {_ratio_interpretation('Cash Ratio', latest_liq['Cash Ratio'], csym)}")
            st.write(f"**Debt-to-Equity ({latest_liq['Debt-to-Equity']:.2f}x):** {_ratio_interpretation('Debt-to-Equity', latest_liq['Debt-to-Equity'], csym)}")
            if latest_liq['Interest Coverage'] is not None:
                st.write(f"**Interest Coverage ({latest_liq['Interest Coverage']:.2f}x):** {_ratio_interpretation('Interest Coverage', latest_liq['Interest Coverage'], csym)}")
            st.write(f"**Debt-to-Assets ({latest_liq['Debt-to-Assets %']:.1f}%):** {_ratio_interpretation('Debt-to-Assets', latest_liq['Debt-to-Assets %'], csym)}")

    # ============================================================
    # PROFITABILITY & EFFICIENCY
    # ============================================================
    st.markdown("#### ⚡ Profitability & Efficiency")
    eff_rows = []
    inv_days = wc_metrics.get('inventory_days', []) if wc_metrics else []
    deb_days = wc_metrics.get('debtor_days', []) if wc_metrics else []
    cred_days = wc_metrics.get('creditor_days', []) if wc_metrics else []
    for i in range(n):
        rev = _get('revenue', i)
        gross_margin = safe_divide(rev - _get('cogs', i), rev, 0) * 100
        ebitda_margin = safe_divide(_get('ebitda', i), rev, 0) * 100
        net_margin = safe_divide(net_income_l[i], rev, 0) * 100
        roce = safe_divide(_get('ebit', i), equity_l[i] + debt_l[i], 0) * 100
        roa = safe_divide(net_income_l[i], total_assets[i], 0) * 100
        i_days = inv_days[i] if i < len(inv_days) else 0
        d_days = deb_days[i] if i < len(deb_days) else 0
        c_days = cred_days[i] if i < len(cred_days) else 0
        ccc = i_days + d_days - c_days
        eff_rows.append({
            'Year': years[i], 'Gross Margin %': gross_margin, 'EBITDA Margin %': ebitda_margin,
            'Net Margin %': net_margin, 'ROCE %': roce, 'ROA %': roa,
            'Inventory Days': i_days, 'Debtor Days': d_days, 'Creditor Days': c_days,
            'Cash Conversion Cycle (days)': ccc
        })
    eff_df = pd.DataFrame(eff_rows)
    _eff_fmt = {c: '{:.1f}%' for c in ['Gross Margin %', 'EBITDA Margin %', 'Net Margin %', 'ROCE %', 'ROA %']}
    _eff_fmt.update({c: '{:.0f}' for c in ['Inventory Days', 'Debtor Days', 'Creditor Days', 'Cash Conversion Cycle (days)']})
    st.dataframe(
        eff_df.style.format(_eff_fmt),
        use_container_width=True, hide_index=True
    )
    if eff_rows:
        latest_eff = eff_rows[0]
        with st.expander("💬 What do these mean? (latest year)", expanded=False):
            st.write(f"**Gross Margin ({latest_eff['Gross Margin %']:.1f}%):** {_ratio_interpretation('Gross Margin', latest_eff['Gross Margin %'], csym)}")
            st.write(f"**EBITDA Margin ({latest_eff['EBITDA Margin %']:.1f}%):** {_ratio_interpretation('EBITDA Margin', latest_eff['EBITDA Margin %'], csym)}")
            st.write(f"**Net Margin ({latest_eff['Net Margin %']:.1f}%):** {_ratio_interpretation('Net Margin', latest_eff['Net Margin %'], csym)}")
            st.write(f"**ROCE ({latest_eff['ROCE %']:.1f}%):** {_ratio_interpretation('ROCE', latest_eff['ROCE %'], csym)}")
            st.write(f"**ROA ({latest_eff['ROA %']:.1f}%):** {_ratio_interpretation('ROA', latest_eff['ROA %'], csym)}")
            st.write(f"**Cash Conversion Cycle ({latest_eff['Cash Conversion Cycle (days)']:.0f} days):** {_ratio_interpretation('Cash Conversion Cycle', latest_eff['Cash Conversion Cycle (days)'], csym)}")

    # ============================================================
    # VALUATION RATIOS (only if a price/share count is available)
    # ============================================================
    if current_price and current_price > 0 and shares and shares > 0:
        st.markdown("#### 💰 Valuation Ratios (Current Market Price)")
        eps = safe_divide(net_income_l[0] * 100000, shares, 0)
        bvps = safe_divide(equity_l[0] * 100000, shares, 0)
        sps = safe_divide(_get('revenue', 0) * 100000, shares, 0)
        pe = safe_divide(current_price, eps, None) if eps > 0 else None
        pb = safe_divide(current_price, bvps, None) if bvps > 0 else None
        ps = safe_divide(current_price, sps, None) if sps > 0 else None
        market_cap = current_price * shares
        ev = market_cap + (debt_l[0] * 100000) - (_get('cash', 0) * 100000)
        ev_ebitda = safe_divide(ev, _get('ebitda', 0) * 100000, None) if _get('ebitda', 0) > 0 else None

        val_cols = st.columns(4)
        for col, (label, val) in zip(val_cols, [('P/E', pe), ('P/B', pb), ('P/S', ps), ('EV/EBITDA', ev_ebitda)]):
            col.metric(label, f"{val:.2f}x" if val is not None else "N/A")
        with st.expander("💬 What do these mean?", expanded=False):
            for label, val in [('P/E', pe), ('P/B', pb), ('P/S', ps), ('EV/EBITDA', ev_ebitda)]:
                if val is not None:
                    st.write(f"**{label} ({val:.2f}x):** {_ratio_interpretation(label, val, csym)}")

    # ============================================================
    # RED FLAGS
    # ============================================================
    st.markdown("#### 🚩 Red Flags")
    flags = []
    if n >= 2:
        rev_growth = safe_divide(_get('revenue', 0) - _get('revenue', 1), _get('revenue', 1), None)
        recv_growth = safe_divide(_get('receivables', 0) - _get('receivables', 1), _get('receivables', 1), None)
        inv_growth = safe_divide(_get('inventory', 0) - _get('inventory', 1), _get('inventory', 1), None)
        margin_now = safe_divide(_get('ebitda', 0), _get('revenue', 0), 0) * 100
        margin_prev = safe_divide(_get('ebitda', 1), _get('revenue', 1), 0) * 100

        if rev_growth is not None and recv_growth is not None and (recv_growth - rev_growth) * 100 > outpace_threshold:
            flags.append(f"📈 **Receivables outpacing revenue growth**: Receivables grew {recv_growth*100:.1f}% vs revenue growth of {rev_growth*100:.1f}% — worth checking collection quality or aggressive revenue recognition.")
        if rev_growth is not None and inv_growth is not None and (inv_growth - rev_growth) * 100 > outpace_threshold:
            flags.append(f"📦 **Inventory building up faster than sales**: Inventory grew {inv_growth*100:.1f}% vs revenue growth of {rev_growth*100:.1f}% — worth checking for slow-moving/obsolete stock.")
        if abs(margin_now - margin_prev) > margin_jump_threshold:
            flags.append(f"📊 **Sudden EBITDA margin swing**: Margin moved from {margin_prev:.1f}% to {margin_now:.1f}% ({margin_now - margin_prev:+.1f}pp) year-on-year — worth checking for one-off items or accounting policy changes.")

    if liq_rows and liq_rows[0]['Current Ratio'] < current_ratio_threshold:
        flags.append(f"💧 **Thin liquidity**: Current Ratio of {liq_rows[0]['Current Ratio']:.2f}x is below your {current_ratio_threshold:.1f}x threshold.")
    if liq_rows and liq_rows[0]['Debt-to-Equity'] > de_threshold:
        flags.append(f"⚖️ **High leverage**: Debt-to-Equity of {liq_rows[0]['Debt-to-Equity']:.2f}x is above your {de_threshold:.1f}x threshold.")
    if liq_rows and liq_rows[0]['Interest Coverage'] is not None and liq_rows[0]['Interest Coverage'] < interest_cov_threshold:
        flags.append(f"🔻 **Thin interest coverage**: EBIT covers interest only {liq_rows[0]['Interest Coverage']:.2f}x, below your {interest_cov_threshold:.1f}x threshold.")

    if flags:
        for f in flags:
            st.warning(f)
    else:
        st.success("✅ No automated red flags triggered at the current thresholds.")

    with st.expander("📋 Items that can't be auto-detected from summary financials — check manually", expanded=False):
        st.markdown(
            "- **Related-party transactions** — check the notes to accounts, not visible in summary financials\n"
            "- **Contingent liabilities & pending litigation** — check notes to accounts / auditor's report\n"
            "- **Auditor qualifications, emphasis of matter, or going-concern remarks**\n"
            "- **Off-balance-sheet items** (operating leases pre-Ind AS 116, guarantees given, LCs)\n"
            "- **Changes in accounting policy or estimates** between years that could distort trend comparisons\n"
            "- **Related entity/group company transactions** and cross-holdings"
        )


def render_valuation_health_check(valuation, terminal_growth, wacc, csym="₹", key_prefix="health",
                                   ddm_result=None, rim_result=None, comp_results=None):
    """
    Consolidates every valuation-quality diagnostic this app already computes
    (terminal value dominance, negative/adjusted FCFF, growth-vs-WACC
    sanity) plus cross-method triangulation (DCF vs DDM vs RIM vs Comps
    average) into one dashboard, so nothing gets missed by being buried in
    a different tab. Every threshold is adjustable — none of this is a
    silent pass/fail baked into the underlying calculation.
    """
    st.subheader("🩺 Valuation Health Check")
    st.caption("Pulls together the sanity checks already computed elsewhere in this app, plus a cross-method comparison, into one place.")

    if not valuation or valuation.get('fair_value_per_share', 0) <= 0:
        st.info("Run a DCF valuation to see its health check.")
        return

    with st.expander("⚙️ Health-Check Thresholds (adjustable)", expanded=False):
        col_t1, col_t2 = st.columns(2)
        with col_t1:
            tv_threshold = st.number_input("Flag if Terminal Value is more than (% of EV):", min_value=50.0, max_value=99.0, value=85.0, step=1.0, key=f"{key_prefix}_tv_thresh")
            spread_threshold = st.number_input("Flag if (WACC − Terminal Growth) is less than (pp):", min_value=0.5, max_value=10.0, value=2.0, step=0.5, key=f"{key_prefix}_spread_thresh")
            triangulation_threshold = st.number_input("Flag if DCF deviates from other methods' average by more than (%):", min_value=10.0, max_value=100.0, value=30.0, step=5.0, key=f"{key_prefix}_tri_thresh")
        with col_t2:
            macro_growth_ceiling = st.number_input("Flag if Terminal Growth exceeds long-run economic growth of (%):", min_value=3.0, max_value=15.0, value=7.0, step=0.5, key=f"{key_prefix}_macro_thresh",
                                                    help="A company can't outgrow the whole economy forever — terminal growth meaningfully above long-run nominal GDP growth is a classic DCF red flag.")
            wacc_low_bound = st.number_input("Flag if WACC is below (%):", min_value=1.0, max_value=15.0, value=6.0, step=0.5, key=f"{key_prefix}_wacc_low")
            wacc_high_bound = st.number_input("Flag if WACC is above (%):", min_value=10.0, max_value=40.0, value=25.0, step=1.0, key=f"{key_prefix}_wacc_high")

    checks = []  # (passed: bool, message: str)

    tv_pct = valuation.get('tv_percentage', 0)
    if tv_pct > tv_threshold:
        checks.append((False, f"**Terminal Value dominates the valuation**: {tv_pct:.1f}% of Enterprise Value comes from the terminal value, above your {tv_threshold:.0f}% threshold — most of this valuation rests on assumptions about the distant future, not the explicit forecast."))
    else:
        checks.append((True, f"Terminal Value is {tv_pct:.1f}% of Enterprise Value — within your {tv_threshold:.0f}% threshold."))

    spread = wacc - terminal_growth
    if spread < spread_threshold:
        checks.append((False, f"**WACC and Terminal Growth are too close**: the spread is only {spread:.1f}pp (WACC {wacc:.1f}% vs Terminal Growth {terminal_growth:.1f}%), below your {spread_threshold:.1f}pp threshold — the Gordon Growth terminal value formula becomes extremely sensitive to small assumption changes as this spread shrinks."))
    else:
        checks.append((True, f"WACC − Terminal Growth spread is {spread:.1f}pp — comfortably above your {spread_threshold:.1f}pp threshold."))

    if terminal_growth > macro_growth_ceiling:
        checks.append((False, f"**Terminal Growth looks unsustainable**: {terminal_growth:.1f}% is above your {macro_growth_ceiling:.1f}% long-run economic growth ceiling — implies the company outgrows the broader economy forever."))
    else:
        checks.append((True, f"Terminal Growth of {terminal_growth:.1f}% is at or below your {macro_growth_ceiling:.1f}% long-run growth ceiling."))

    if wacc < wacc_low_bound:
        checks.append((False, f"**WACC looks unusually low**: {wacc:.1f}% is below your {wacc_low_bound:.1f}% floor — double-check the risk-free rate, beta, and cost of debt inputs."))
    elif wacc > wacc_high_bound:
        checks.append((False, f"**WACC looks unusually high**: {wacc:.1f}% is above your {wacc_high_bound:.1f}% ceiling — double-check the risk premium and cost of debt inputs."))
    else:
        checks.append((True, f"WACC of {wacc:.1f}% is within your {wacc_low_bound:.1f}%–{wacc_high_bound:.1f}% expected range."))

    if valuation.get('fcff_adjusted'):
        adj = valuation.get('adjustment_details') or {}
        checks.append((False, f"**FCFF was adjusted**: the model had to override projected free cash flow (e.g. persistently negative FCFF) — check the adjustment details in the DCF Summary/FCF Working tabs before relying on this fair value. {adj if adj else ''}"))
    else:
        checks.append((True, "Projected FCFF required no manual adjustment."))

    if valuation.get('negative_fair_value_warning'):
        checks.append((False, f"**Negative fair value flagged**: {valuation['negative_fair_value_warning']}"))

    # ---------- Cross-method triangulation ----------
    dcf_fv = valuation['fair_value_per_share']
    other_values = []
    if ddm_result and ddm_result.get('value_per_share', 0) > 0:
        other_values.append(('DDM', ddm_result['value_per_share']))
    if rim_result and rim_result.get('value_per_share', 0) > 0:
        other_values.append(('RIM', rim_result['value_per_share']))
    if comp_results and comp_results.get('valuations'):
        comp_avgs = [v['fair_value_avg'] for v in comp_results['valuations'].values() if v.get('fair_value_avg', 0) > 0]
        if comp_avgs:
            other_values.append(('Comps (avg of methods)', np.mean(comp_avgs)))

    if other_values:
        avg_other = np.mean([v for _, v in other_values])
        pct_diff = safe_divide(dcf_fv - avg_other, avg_other, 0) * 100
        methods_str = ", ".join(f"{label}: {csym}{val:.2f}" for label, val in other_values)
        if abs(pct_diff) > triangulation_threshold:
            direction = "above" if pct_diff > 0 else "below"
            checks.append((False, f"**DCF diverges from other methods**: DCF fair value ({csym}{dcf_fv:.2f}) is {abs(pct_diff):.0f}% {direction} the average of other methods run ({csym}{avg_other:.2f}, from {methods_str}) — above your {triangulation_threshold:.0f}% threshold."))
        else:
            checks.append((True, f"DCF fair value ({csym}{dcf_fv:.2f}) is within {triangulation_threshold:.0f}% of the average of other methods run ({csym}{avg_other:.2f}, from {methods_str})."))
    else:
        checks.append((True, "No other valuation methods (DDM/RIM/Comps) have been run yet to triangulate against — run one alongside DCF for a cross-check."))

    n_flagged = sum(1 for passed, _ in checks if not passed)
    c1, c2 = st.columns(2)
    c1.metric("Checks Passed", f"{len(checks) - n_flagged} / {len(checks)}")
    c2.metric("Flags Raised", n_flagged)

    for passed, msg in checks:
        if passed:
            st.success(f"✅ {msg}")
        else:
            st.warning(f"⚠️ {msg}")


@st.cache_data(ttl=86400, show_spinner=False)  # company descriptions barely change — cache for a day
def fetch_screener_company_info(ticker_symbol):
    """
    Scrapes the public "About" blurb (and "Key Points", if present) from the
    company's Screener.in page — this is ordinary public page content, not
    the authenticated Excel export, so no login/cookies are needed.
    Returns {'about': str, 'key_points': str, 'source_url': str} or None if
    the page/section couldn't be found (never raises).
    """
    if not ticker_symbol or not ticker_symbol.strip():
        return None
    ticker_clean = ticker_symbol.strip().upper().replace('.NS', '').replace('.BO', '')
    for url in (f"https://www.screener.in/company/{ticker_clean}/consolidated/",
                f"https://www.screener.in/company/{ticker_clean}/"):
        try:
            resp = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=8)
            if resp.status_code != 200:
                continue
            soup = BeautifulSoup(resp.text, "html.parser")
            profile = soup.select_one("div.company-profile")
            if not profile:
                continue
            about_p = profile.select_one("div.about p")
            about_text = about_p.get_text(" ", strip=True) if about_p else ""
            key_points_text = ""
            for title_div in profile.select("div.title"):
                if title_div.get_text(strip=True) == "Key Points":
                    kp_div = title_div.find_next_sibling("div")
                    if kp_div:
                        for br in kp_div.find_all("br"):
                            br.replace_with("\n")
                        key_points_text = kp_div.get_text(strip=True)
                    break
            if about_text or key_points_text:
                return {'about': about_text, 'key_points': key_points_text, 'source_url': url}
        except Exception:
            continue
    return None


def render_screener_company_info(ticker_symbol, key_prefix="company_info"):
    """Displays the scraped About/Key Points blurb, if found, and stays silent (no error) otherwise."""
    info = fetch_screener_company_info(ticker_symbol)
    if not info:
        return
    with st.expander(f"ℹ️ About {ticker_symbol.strip().upper()}", expanded=True):
        if info.get('about'):
            st.write(info['about'])
        if info.get('key_points'):
            st.markdown("**Key Points**")
            st.caption(info['key_points'])
        st.caption(f"Source: [{info['source_url']}]({info['source_url']})")


def create_zscore_gauge(z_score, distress_threshold, safe_threshold):
    """Gauge chart for the Altman Z-Score, styled to match the price-vs-fair-value gauge elsewhere in the app."""
    axis_max = max(z_score, safe_threshold) * 1.5 + 1
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=z_score,
        domain={'x': [0, 1], 'y': [0, 1]},
        title={'text': "Altman Z-Score", 'font': {'size': 20}},
        number={'valueformat': '.2f'},
        gauge={
            'axis': {'range': [None, axis_max]},
            'bar': {'color': "#2E86AB"},
            'bgcolor': "white",
            'borderwidth': 2,
            'bordercolor': "gray",
            'steps': [
                {'range': [0, distress_threshold], 'color': '#D62828'},
                {'range': [distress_threshold, safe_threshold], 'color': '#F4D35E'},
                {'range': [safe_threshold, axis_max], 'color': '#06A77D'}],
            'threshold': {'line': {'color': "black", 'width': 4}, 'thickness': 0.75, 'value': z_score}
        }
    ))
    fig.update_layout(height=380)
    return fig


def render_altman_zscore_tab(financials, csym, key_prefix, current_price=0, shares=None):
    """
    Altman Z-Score distress prediction, with a choice of three published
    model variants and every coefficient, weight and zone threshold fully
    adjustable — none of Altman's published numbers are silently locked in.
    Computed for the latest year as the primary distress read (the standard
    way this score is used); an optional historical-trend section lets the
    user supply Retained Earnings for prior years too, since that line item
    isn't separately tracked elsewhere in this app.
    """
    st.subheader("🧮 Altman Z-Score — Financial Distress Prediction")
    st.caption("A published bankruptcy-prediction score combining balance sheet and income statement ratios into a single distress indicator.")

    years = financials.get('years', [])
    n = len(years)
    if n == 0:
        st.info("No historical financials available yet.")
        return

    def _get(key, i):
        arr = financials.get(key, [])
        return ensure_valid_number(arr[i], 0) if i < len(arr) else 0

    model = st.selectbox(
        "Model variant:",
        ["Original Z-Score (Public Manufacturers)", "Z'-Score (Private Companies)", "Z''-Score (Non-Manufacturers / Emerging Markets)"],
        key=f"{key_prefix}_model",
        help="Original Z was built for public manufacturers using market value of equity. Z' substitutes book value of "
             "equity for private companies. Z'' drops the Sales/Assets ratio (unstable across industries) and is "
             "commonly used for non-manufacturers and emerging-market companies."
    )
    if model.startswith("Original"):
        defaults = {'c1': 1.2, 'c2': 1.4, 'c3': 3.3, 'c4': 0.6, 'c5': 1.0, 'safe': 2.99, 'distress': 1.81, 'use_x5': True, 'use_market_value': True}
    elif model.startswith("Z'-"):
        defaults = {'c1': 0.717, 'c2': 0.847, 'c3': 3.107, 'c4': 0.420, 'c5': 0.998, 'safe': 2.9, 'distress': 1.23, 'use_x5': True, 'use_market_value': False}
    else:
        defaults = {'c1': 6.56, 'c2': 3.26, 'c3': 6.72, 'c4': 1.05, 'c5': 0.0, 'safe': 2.6, 'distress': 1.1, 'use_x5': False, 'use_market_value': False}

    with st.expander("⚙️ Coefficients & Zone Thresholds (adjustable)", expanded=False):
        st.caption("Pre-filled with Altman's published values for the selected model above — edit any of them to match your own research or a sector-specific recalibration.")
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            c1 = st.number_input("X1 Weight", value=defaults['c1'], step=0.01, key=f"{key_prefix}_c1", help="Working Capital / Total Assets")
        with col2:
            c2 = st.number_input("X2 Weight", value=defaults['c2'], step=0.01, key=f"{key_prefix}_c2", help="Retained Earnings / Total Assets")
        with col3:
            c3 = st.number_input("X3 Weight", value=defaults['c3'], step=0.01, key=f"{key_prefix}_c3", help="EBIT / Total Assets")
        with col4:
            c4 = st.number_input("X4 Weight", value=defaults['c4'], step=0.01, key=f"{key_prefix}_c4", help="Equity / Total Liabilities")
        with col5:
            use_x5 = st.checkbox("Include X5?", value=defaults['use_x5'], key=f"{key_prefix}_use_x5", help="Sales / Total Assets — Altman's Z'' model drops this ratio")
            c5 = st.number_input("X5 Weight", value=defaults['c5'], step=0.01, key=f"{key_prefix}_c5", disabled=not use_x5)

        col_z1, col_z2, col_z3 = st.columns(3)
        with col_z1:
            safe_threshold = st.number_input("Safe Zone (Z above)", value=defaults['safe'], step=0.05, key=f"{key_prefix}_safe")
        with col_z2:
            distress_threshold = st.number_input("Distress Zone (Z below)", value=defaults['distress'], step=0.05, key=f"{key_prefix}_distress")
        with col_z3:
            use_market_value = st.checkbox(
                "Use Market Value of Equity for X4?", value=defaults['use_market_value'], key=f"{key_prefix}_use_mv",
                help="Unchecked uses Book Value of Equity instead — needed when no reliable market price exists (unlisted companies)."
            )

    st.markdown("#### 📋 Inputs")
    total_assets_latest = _get('fixed_assets', 0) + _get('inventory', 0) + _get('receivables', 0) + _get('cash', 0)
    retained_earnings_input = st.number_input(
        f"Retained Earnings / Reserves & Surplus, latest year ({csym} Lacs):", value=0.0, step=100.0, key=f"{key_prefix}_re",
        help="Not separately tracked elsewhere in this app (only a single lumped Equity figure is) — enter it from the "
             "balance sheet. Leaving this at 0 will understate X2 and the overall Z-Score, not eliminate that component."
    )
    st.caption(f"Approximate Total Assets used below: {csym}{total_assets_latest * 100000:,.0f} (Fixed Assets + Inventory + Receivables + Cash — the same approximation used in the Ratio Analysis and IBC Liquidation tabs).")

    total_assets = total_assets_latest
    total_liabilities = _get('st_debt', 0) + _get('lt_debt', 0) + _get('payables', 0)
    working_capital = (_get('inventory', 0) + _get('receivables', 0) + _get('cash', 0)) - (_get('payables', 0) + _get('st_debt', 0))
    ebit = _get('ebit', 0)
    sales = _get('revenue', 0)
    book_equity = _get('equity', 0)

    equity_for_x4 = book_equity
    equity_source = "Book Value"
    if use_market_value:
        if current_price and shares:
            equity_for_x4 = (current_price * shares) / 100000  # to Lacs, matching other figures
            equity_source = "Market Value"
        else:
            st.caption("💡 No market price/share count available — falling back to Book Value of Equity for X4 even though Market Value was selected.")

    x1 = safe_divide(working_capital, total_assets, 0)
    x2 = safe_divide(retained_earnings_input, total_assets, 0)
    x3 = safe_divide(ebit, total_assets, 0)
    x4 = safe_divide(equity_for_x4, total_liabilities, 0)
    x5 = safe_divide(sales, total_assets, 0) if use_x5 else 0

    z_score = c1 * x1 + c2 * x2 + c3 * x3 + c4 * x4 + (c5 * x5 if use_x5 else 0)

    if z_score > safe_threshold:
        zone, zone_msg = "🟢 Safe Zone", "Low probability of financial distress based on this score."
    elif z_score < distress_threshold:
        zone, zone_msg = "🔴 Distress Zone", "Elevated probability of financial distress — worth deeper investigation."
    else:
        zone, zone_msg = "🟡 Grey Zone", "Ambiguous — neither clearly safe nor clearly distressed on this score alone."

    st.markdown("#### 🎯 Result")
    col_r1, col_r2 = st.columns([1, 1])
    with col_r1:
        st.plotly_chart(create_zscore_gauge(z_score, distress_threshold, safe_threshold), use_container_width=True)
    with col_r2:
        st.metric("Z-Score", f"{z_score:.2f}")
        st.markdown(f"**Zone: {zone}**")
        st.caption(zone_msg)
        st.caption(f"Equity basis used for X4: **{equity_source}**")

    st.markdown("#### 🧮 Component Breakdown")
    breakdown_rows = [
        {'Ratio': 'X1 (Working Capital / Total Assets)', 'Value': x1, 'Weight': c1, 'Contribution': c1 * x1},
        {'Ratio': 'X2 (Retained Earnings / Total Assets)', 'Value': x2, 'Weight': c2, 'Contribution': c2 * x2},
        {'Ratio': 'X3 (EBIT / Total Assets)', 'Value': x3, 'Weight': c3, 'Contribution': c3 * x3},
        {'Ratio': f'X4 (Equity [{equity_source}] / Total Liabilities)', 'Value': x4, 'Weight': c4, 'Contribution': c4 * x4},
    ]
    if use_x5:
        breakdown_rows.append({'Ratio': 'X5 (Sales / Total Assets)', 'Value': x5, 'Weight': c5, 'Contribution': c5 * x5})
    bd_df = pd.DataFrame(breakdown_rows)
    st.dataframe(bd_df.style.format({'Value': '{:.3f}', 'Weight': '{:.3f}', 'Contribution': '{:.3f}'}), use_container_width=True, hide_index=True)

    with st.expander("💬 What does this mean?", expanded=False):
        st.write(f"**X1 ({x1:.3f}):** For every {csym}1 of assets, {csym}{x1:.3f} is tied up as net working capital — a negative value suggests short-term liquidity strain.")
        st.write(f"**X2 ({x2:.3f}):** Retained earnings make up {x2*100:.1f}% of total assets — a proxy for how much of the asset base has been built from accumulated profits rather than debt/fresh equity.")
        st.write(f"**X3 ({x3:.3f}):** Operating profit (EBIT) is {x3*100:.1f}% of total assets — the core profitability-per-asset-rupee driver of the score.")
        st.write(f"**X4 ({x4:.3f}):** Equity covers {x4:.2f}x total liabilities — the solvency cushion available before liabilities exceed the equity base.")
        if use_x5:
            st.write(f"**X5 ({x5:.3f}):** The business generates {csym}{x5:.2f} of sales for every {csym}1 of assets — an asset-efficiency measure.")
        st.write(f"**Overall ({z_score:.2f}, {zone}):** {zone_msg}")

    # ---------- Optional historical trend ----------
    with st.expander("📈 Historical Trend (optional — needs Retained Earnings per year)", expanded=False):
        st.caption(
            "X1, X3, X4 and X5 can be computed for every year automatically, but Retained Earnings for prior years "
            "isn't available elsewhere in this app — enter it below (Lacs) to see the Z-Score trend rather than "
            "just the latest year. Leave a year at 0 and its score will understate X2 for that year."
        )
        re_table_key = f"{key_prefix}_re_table"
        default_re_rows = pd.DataFrame({'Year': years, 'Retained Earnings': [retained_earnings_input if i == 0 else 0.0 for i in range(n)]})
        if re_table_key not in st.session_state:
            st.session_state[re_table_key] = default_re_rows
        edited_re = st.data_editor(
            st.session_state[re_table_key],
            column_config={
                "Year": st.column_config.TextColumn("Year", disabled=True),
                "Retained Earnings": st.column_config.NumberColumn(f"Retained Earnings ({csym} Lacs)", step=100.0)
            },
            use_container_width=True, hide_index=True, key=f"{key_prefix}_re_editor"
        )
        st.session_state[re_table_key] = edited_re

        trend_rows = []
        for i in range(n):
            ta_i = _get('fixed_assets', i) + _get('inventory', i) + _get('receivables', i) + _get('cash', i)
            tl_i = _get('st_debt', i) + _get('lt_debt', i) + _get('payables', i)
            wc_i = (_get('inventory', i) + _get('receivables', i) + _get('cash', i)) - (_get('payables', i) + _get('st_debt', i))
            re_i = float(edited_re.iloc[i]['Retained Earnings']) if i < len(edited_re) else 0.0
            eq_i = _get('equity', i) if not (use_market_value and i == 0 and current_price and shares) else equity_for_x4

            x1_i = safe_divide(wc_i, ta_i, 0)
            x2_i = safe_divide(re_i, ta_i, 0)
            x3_i = safe_divide(_get('ebit', i), ta_i, 0)
            x4_i = safe_divide(eq_i, tl_i, 0)
            x5_i = safe_divide(_get('revenue', i), ta_i, 0) if use_x5 else 0
            z_i = c1 * x1_i + c2 * x2_i + c3 * x3_i + c4 * x4_i + (c5 * x5_i if use_x5 else 0)
            trend_rows.append({'Year': years[i], 'Z-Score': z_i})

        trend_chron = list(reversed(trend_rows))
        fig_trend = go.Figure()
        fig_trend.add_trace(go.Scatter(
            x=[r['Year'] for r in trend_chron], y=[r['Z-Score'] for r in trend_chron],
            mode='lines+markers', line=dict(color='#2E86AB', width=3), name='Z-Score'
        ))
        fig_trend.add_hline(y=safe_threshold, line_dash="dash", line_color="#06A77D", annotation_text="Safe Zone")
        fig_trend.add_hline(y=distress_threshold, line_dash="dash", line_color="#D62828", annotation_text="Distress Zone")
        fig_trend.update_layout(title="Z-Score Trend", yaxis_title="Z-Score", height=380, plot_bgcolor='rgba(240,240,240,0.4)')
        st.plotly_chart(fig_trend, use_container_width=True)


def run_lbo_debt_schedule(tranches, hold_years, revenue0, rev_growth_pct, ebitda_margin_pct,
                           capex_pct, da_pct, wc_pct_of_growth, tax_rate, min_cash_reserve):
    """
    Shared LBO projection + debt-paydown waterfall engine. Works identically
    for one tranche (Simple mode) or several (Advanced mode) — Simple mode
    is just this same engine called with a single-tranche list, so the two
    modes can never silently drift apart in how paydown is computed.

    Each year: project Revenue/EBITDA off flat growth & margin assumptions,
    compute cash available after CapEx, working capital, cash taxes and
    interest, then pay down tranches in seniority order — mandatory
    amortization first, then the cash-sweep percentage of whatever's left
    above the minimum cash reserve. All figures in the same currency units
    as revenue0/tranche principals (Lacs, matching the rest of this app).

    Returns (schedule: list[dict] per year, final_balances: dict, final_cash: float).
    """
    tranches = sorted(tranches, key=lambda t: t['seniority'])
    balances = {t['name']: t['principal'] for t in tranches}
    original_principal = {t['name']: t['principal'] for t in tranches}
    cash_balance = 0.0

    schedule = []
    revenue = revenue0
    for yr in range(1, int(hold_years) + 1):
        prev_revenue = revenue
        revenue = revenue * (1 + rev_growth_pct / 100)
        ebitda = revenue * (ebitda_margin_pct / 100)
        da = revenue * (da_pct / 100)
        capex = revenue * (capex_pct / 100)
        delta_wc = (revenue - prev_revenue) * (wc_pct_of_growth / 100)

        total_interest = sum(balances[t['name']] * (t['rate'] / 100) for t in tranches)
        ebt = ebitda - da - total_interest
        cash_taxes = max(ebt, 0) * (tax_rate / 100)

        fcf_before_debt_service = ebitda - capex - delta_wc - cash_taxes - total_interest
        cash_available = cash_balance + fcf_before_debt_service

        year_payments = {t['name']: 0.0 for t in tranches}

        # Mandatory amortization first, by seniority
        for t in tranches:
            mand = original_principal[t['name']] * (t['mandatory_amort_pct'] / 100)
            mand = max(min(mand, balances[t['name']], max(cash_available, 0)), 0)
            balances[t['name']] -= mand
            cash_available -= mand
            year_payments[t['name']] += mand

        # Cash sweep (optional paydown), by seniority, respecting the minimum cash reserve
        sweepable = max(cash_available - min_cash_reserve, 0)
        for t in tranches:
            if sweepable <= 0 or balances[t['name']] <= 0:
                continue
            sweep_amt = min(sweepable * (t['sweep_pct'] / 100), balances[t['name']])
            balances[t['name']] -= sweep_amt
            cash_available -= sweep_amt
            sweepable -= sweep_amt
            year_payments[t['name']] += sweep_amt

        cash_balance = cash_available

        row = {
            'Year': yr, 'Revenue': revenue, 'EBITDA': ebitda, 'Interest': total_interest,
            'Cash Taxes': cash_taxes, 'FCF Before Debt Service': fcf_before_debt_service,
            'Total Debt (End)': sum(balances.values()), 'Cash Balance': cash_balance,
        }
        for t in tranches:
            row[f"{t['name']} Balance"] = balances[t['name']]
            row[f"{t['name']} Paydown"] = year_payments[t['name']]
        schedule.append(row)

    return schedule, dict(balances), cash_balance


def render_lbo_analysis_tab(financials, csym, key_prefix, tax_rate=25.0):
    """
    LBO Analysis — the return a financial/PE sponsor would earn buying the
    business today with leverage, and the implied 'floor value' a leveraged
    buyer could pay while still hitting target returns. A standard fourth
    valuation lens alongside DCF, Comps, and Precedent Transactions.

    Supports both a Simple single blended debt tranche and an Advanced
    multi-tranche structure (different rates/amortization/seniority per
    tranche) through the exact same underlying paydown engine — nothing
    about the mechanics differs between the two, only how many tranches
    are fed into it. Every assumption is adjustable; none of Altman/LBO
    convention defaults are silently locked in.
    """
    st.subheader("💼 LBO Analysis")
    st.caption(
        "Estimates the return a private equity / financial sponsor would earn buying this business with leverage "
        "today, and the 'floor value' a leveraged buyer could pay while still hitting target returns."
    )

    years_hist = financials.get('years', [])
    if not years_hist:
        st.info("No historical financials available yet.")
        return

    latest_ebitda = ensure_valid_number(financials['ebitda'][0], 0) if financials.get('ebitda') else 0
    latest_revenue = ensure_valid_number(financials['revenue'][0], 0) if financials.get('revenue') else 0
    latest_da = ensure_valid_number(financials['depreciation'][0], 0) if financials.get('depreciation') else 0

    if latest_ebitda <= 0:
        st.warning("Latest EBITDA is zero or negative — LBO analysis needs positive EBITDA to be meaningful.")
        return

    # ---------- Entry Assumptions ----------
    st.markdown("#### 🚪 Entry Assumptions")
    col1, col2, col3 = st.columns(3)
    with col1:
        entry_multiple = st.number_input("Entry EV/EBITDA Multiple (x):", min_value=1.0, max_value=30.0, value=8.0, step=0.5, key=f"{key_prefix}_entry_mult")
        hold_years = st.number_input("Hold Period (years):", min_value=1, max_value=10, value=5, step=1, key=f"{key_prefix}_hold_years")
    with col2:
        exit_multiple = st.number_input("Exit EV/EBITDA Multiple (x):", min_value=1.0, max_value=30.0, value=entry_multiple, step=0.5, key=f"{key_prefix}_exit_mult",
                                         help="Defaults to the entry multiple (no multiple expansion/contraction assumed) — edit to model a different exit environment.")
        transaction_fee_pct = st.number_input("Transaction Fees (% of EV):", min_value=0.0, max_value=10.0, value=2.0, step=0.5, key=f"{key_prefix}_fees")
    with col3:
        target_debt_ebitda = st.number_input("Total Debt / EBITDA at Entry (x):", min_value=0.0, max_value=10.0, value=4.0, step=0.25, key=f"{key_prefix}_debt_ebitda")
        min_cash_reserve = st.number_input(f"Minimum Cash Reserve ({csym} Lacs):", min_value=0.0, value=0.0, step=100.0, key=f"{key_prefix}_min_cash")

    entry_ev = entry_multiple * latest_ebitda
    total_fees = entry_ev * (transaction_fee_pct / 100)
    total_debt_entry = target_debt_ebitda * latest_ebitda
    total_uses = entry_ev + total_fees
    sponsor_equity_entry = total_uses - total_debt_entry

    if sponsor_equity_entry <= 0:
        st.error("At this Debt/EBITDA multiple, debt alone exceeds the purchase price plus fees — reduce leverage or increase the entry multiple.")
        return

    st.markdown("##### 💰 Sources & Uses")
    su_col1, su_col2 = st.columns(2)
    with su_col1:
        st.markdown("**Uses**")
        st.write(f"Purchase Enterprise Value: {csym}{entry_ev:,.1f} Lacs")
        st.write(f"Transaction Fees: {csym}{total_fees:,.1f} Lacs")
        st.write(f"**Total Uses: {csym}{total_uses:,.1f} Lacs**")
    with su_col2:
        st.markdown("**Sources**")
        st.write(f"Total Debt: {csym}{total_debt_entry:,.1f} Lacs ({target_debt_ebitda:.2f}x EBITDA)")
        st.write(f"Sponsor Equity: {csym}{sponsor_equity_entry:,.1f} Lacs")
        st.write(f"**Total Sources: {csym}{total_debt_entry + sponsor_equity_entry:,.1f} Lacs**")

    # ---------- Debt Structure ----------
    st.markdown("#### 🏦 Debt Structure")
    debt_mode = st.radio(
        "Debt structure detail:", ["Simple (single blended tranche)", "Advanced (multiple tranches)"],
        key=f"{key_prefix}_debt_mode", horizontal=True
    )

    if debt_mode.startswith("Simple"):
        col_d1, col_d2, col_d3 = st.columns(3)
        with col_d1:
            interest_rate = st.number_input("Interest Rate (%):", min_value=0.0, max_value=30.0, value=11.0, step=0.25, key=f"{key_prefix}_rate_simple")
        with col_d2:
            mandatory_amort_pct = st.number_input("Mandatory Amortization (% of principal/yr):", min_value=0.0, max_value=50.0, value=5.0, step=1.0, key=f"{key_prefix}_amort_simple")
        with col_d3:
            cash_sweep_pct = st.number_input("Cash Sweep (% of excess FCF):", min_value=0.0, max_value=100.0, value=75.0, step=5.0, key=f"{key_prefix}_sweep_simple")
        tranches = [{
            'name': 'Total Debt', 'principal': total_debt_entry, 'rate': interest_rate,
            'mandatory_amort_pct': mandatory_amort_pct, 'sweep_pct': cash_sweep_pct, 'seniority': 1
        }]
    else:
        st.caption("Add/edit tranches below. '% of Total Debt' across all rows should sum to 100% (auto-scaled if not). Lower Seniority Rank = paid first, for both mandatory amortization and cash sweep.")
        table_key = f"{key_prefix}_tranche_table"
        if table_key not in st.session_state:
            st.session_state[table_key] = pd.DataFrame([
                {'Tranche': 'Senior Term Loan', '% of Total Debt': 70.0, 'Interest Rate %': 9.0, 'Mandatory Amort % / yr': 10.0, 'Seniority Rank': 1},
                {'Tranche': 'Subordinated / Mezzanine', '% of Total Debt': 30.0, 'Interest Rate %': 14.0, 'Mandatory Amort % / yr': 0.0, 'Seniority Rank': 2},
            ])
        edited_tranches = st.data_editor(
            st.session_state[table_key],
            column_config={
                'Tranche': st.column_config.TextColumn('Tranche'),
                '% of Total Debt': st.column_config.NumberColumn('% of Total Debt', min_value=0.0, max_value=100.0),
                'Interest Rate %': st.column_config.NumberColumn('Interest Rate %', min_value=0.0, max_value=30.0),
                'Mandatory Amort % / yr': st.column_config.NumberColumn('Mandatory Amort % / yr', min_value=0.0, max_value=100.0),
                'Seniority Rank': st.column_config.NumberColumn('Seniority Rank (1 = paid first)', min_value=1, max_value=10, step=1),
            },
            num_rows="dynamic", use_container_width=True, hide_index=True, key=f"{key_prefix}_tranche_editor"
        )
        st.session_state[table_key] = edited_tranches

        pct_sum = pd.to_numeric(edited_tranches['% of Total Debt'], errors='coerce').fillna(0).sum()
        if pct_sum > 0 and abs(pct_sum - 100) > 0.5:
            st.caption(f"💡 Tranche percentages sum to {pct_sum:.1f}%, not 100% — debt allocation below is scaled proportionally so the total still matches Total Debt from Sources & Uses.")

        cash_sweep_pct_advanced = st.number_input(
            "Cash Sweep (% of excess FCF, applied by seniority order):", min_value=0.0, max_value=100.0, value=75.0, step=5.0, key=f"{key_prefix}_sweep_advanced"
        )

        tranches = []
        if pct_sum > 0:
            for _, row in edited_tranches.sort_values('Seniority Rank').iterrows():
                pct = ensure_valid_number(row['% of Total Debt'], 0) / pct_sum
                tranches.append({
                    'name': str(row['Tranche']) if row['Tranche'] else f"Tranche {len(tranches)+1}",
                    'principal': total_debt_entry * pct,
                    'rate': ensure_valid_number(row['Interest Rate %'], 0),
                    'mandatory_amort_pct': ensure_valid_number(row['Mandatory Amort % / yr'], 0),
                    'sweep_pct': cash_sweep_pct_advanced,
                    'seniority': ensure_valid_number(row['Seniority Rank'], 1)
                })
        if not tranches:
            st.warning("Add at least one tranche with a non-zero % of Total Debt to run the paydown schedule.")
            return

    # ---------- Operating Assumptions ----------
    st.markdown("#### 📈 Operating Assumptions (Hold Period)")
    default_margin = safe_divide(latest_ebitda, latest_revenue, 0) * 100 if latest_revenue > 0 else 20.0
    default_da_pct = safe_divide(latest_da, latest_revenue, 0) * 100 if latest_revenue > 0 and latest_da > 0 else 3.0
    col_o1, col_o2, col_o3, col_o4 = st.columns(4)
    with col_o1:
        rev_growth = st.number_input("Revenue Growth %/yr:", value=10.0, step=0.5, key=f"{key_prefix}_rev_growth")
    with col_o2:
        ebitda_margin = st.number_input("EBITDA Margin % (held flat):", value=round(default_margin, 1), step=0.5, key=f"{key_prefix}_ebitda_margin")
    with col_o3:
        capex_pct = st.number_input("CapEx % of Revenue:", value=5.0, step=0.5, key=f"{key_prefix}_capex_pct")
    with col_o4:
        da_pct = st.number_input("D&A % of Revenue:", value=round(default_da_pct, 1), step=0.5, key=f"{key_prefix}_da_pct")
    wc_pct_of_growth = st.number_input(
        "Incremental Working Capital (% of Revenue Growth):", value=10.0, step=1.0, key=f"{key_prefix}_wc_pct",
        help="Extra working capital needed each year, as a % of that year's rupee revenue growth — a simplified stand-in for a full working-capital-days build."
    )

    # ---------- Run schedule & exit ----------
    schedule, final_balances, final_cash = run_lbo_debt_schedule(
        tranches, hold_years, latest_revenue, rev_growth, ebitda_margin,
        capex_pct, da_pct, wc_pct_of_growth, tax_rate, min_cash_reserve
    )

    exit_ebitda = schedule[-1]['EBITDA']
    exit_ev = exit_multiple * exit_ebitda
    exit_net_debt = sum(final_balances.values()) - final_cash
    exit_equity_value = exit_ev - exit_net_debt

    moic = safe_divide(exit_equity_value, sponsor_equity_entry, None)
    irr_pct = ((moic ** (1 / hold_years)) - 1) * 100 if moic and moic > 0 else None

    st.markdown("#### 📅 Debt Paydown Schedule")
    sched_df = pd.DataFrame(schedule)
    display_cols = ['Year', 'Revenue', 'EBITDA', 'Interest', 'FCF Before Debt Service', 'Total Debt (End)', 'Cash Balance']
    fmt = {c: f'{csym}{{:,.0f}}' for c in display_cols if c != 'Year'}
    st.dataframe(sched_df[display_cols].style.format(fmt), use_container_width=True, hide_index=True)

    fig_debt = go.Figure()
    balance_cols = [c for c in sched_df.columns if c.endswith(' Balance') and c != 'Cash Balance']
    for col in balance_cols:
        fig_debt.add_trace(go.Scatter(x=sched_df['Year'], y=sched_df[col], mode='lines+markers', name=col.replace(' Balance', ''), stackgroup='debt'))
    fig_debt.update_layout(title="Debt Paydown Over the Hold Period", xaxis_title="Year", yaxis_title=f"Balance ({csym} Lacs)", height=380, plot_bgcolor='rgba(240,240,240,0.4)')
    st.plotly_chart(fig_debt, use_container_width=True)

    st.markdown("#### 🚪 Exit & Returns")
    ex_col1, ex_col2, ex_col3 = st.columns(3)
    with ex_col1:
        st.metric(f"Exit Enterprise Value (Yr {int(hold_years)})", f"{csym}{exit_ev:,.1f} Lacs")
        st.caption(f"Exit EBITDA {csym}{exit_ebitda:,.1f} Lacs × {exit_multiple:.1f}x")
    with ex_col2:
        st.metric("Exit Equity Value", f"{csym}{exit_equity_value:,.1f} Lacs")
        st.caption(f"Less Exit Net Debt of {csym}{exit_net_debt:,.1f} Lacs")
    with ex_col3:
        st.metric("MOIC", f"{moic:.2f}x" if moic is not None else "N/A")
        st.metric("IRR", f"{irr_pct:.1f}%" if irr_pct is not None else "N/A")

    if moic is not None:
        with st.expander("💬 What does this mean?", expanded=False):
            st.write(f"A sponsor investing {csym}{sponsor_equity_entry:,.1f} Lacs of equity today would end up with {csym}{exit_equity_value:,.1f} Lacs after {int(hold_years)} years — a **{moic:.2f}x** return on invested capital, equivalent to roughly **{irr_pct:.1f}%** per year (IRR), assuming no interim dividends and the entry/exit multiple, leverage and operating assumptions above.")
            debt_paydown_pct = safe_divide(total_debt_entry - sum(final_balances.values()), total_debt_entry, 0) * 100
            st.write(f"Over the hold period, {debt_paydown_pct:.0f}% of the entry debt was paid down, leaving {csym}{sum(final_balances.values()):,.1f} Lacs outstanding at exit (plus {csym}{final_cash:,.1f} Lacs of accumulated cash).")

    # ---------- IRR sensitivity ----------
    with st.expander("🔥 IRR Sensitivity — Entry vs Exit Multiple", expanded=False):
        st.caption("Reruns the same debt schedule and returns for a grid of entry/exit multiples around your inputs above — everything else held constant.")
        entry_range = [entry_multiple + d for d in (-1.5, -0.75, 0, 0.75, 1.5)]
        exit_range = [exit_multiple + d for d in (-1.5, -0.75, 0, 0.75, 1.5)]
        sens_rows = []
        for em in entry_range:
            row = {'Entry Multiple': f"{em:.2f}x"}
            for xm in exit_range:
                ev_s = em * latest_ebitda
                fees_s = ev_s * (transaction_fee_pct / 100)
                debt_s = target_debt_ebitda * latest_ebitda
                equity_s = (ev_s + fees_s) - debt_s
                if equity_s <= 0:
                    row[f"{xm:.2f}x"] = None
                    continue
                sched_s, bal_s, cash_s = run_lbo_debt_schedule(
                    [{**t, 'principal': t['principal'] * safe_divide(debt_s, total_debt_entry, 1)} for t in tranches],
                    hold_years, latest_revenue, rev_growth, ebitda_margin, capex_pct, da_pct, wc_pct_of_growth, tax_rate, min_cash_reserve
                )
                exit_ev_s = xm * sched_s[-1]['EBITDA']
                exit_eq_s = exit_ev_s - (sum(bal_s.values()) - cash_s)
                moic_s = safe_divide(exit_eq_s, equity_s, None)
                irr_s = ((moic_s ** (1 / hold_years)) - 1) * 100 if moic_s and moic_s > 0 else None
                row[f"{xm:.2f}x"] = f"{irr_s:.1f}%" if irr_s is not None else "N/A"
            sens_rows.append(row)
        sens_df = pd.DataFrame(sens_rows).set_index('Entry Multiple')
        st.caption("Columns = Exit Multiple, Rows = Entry Multiple. Values are IRR.")
        st.dataframe(sens_df, use_container_width=True)


# ================================
# MAIN UI FUNCTION
# ================================
def main():
    """Main DCF UI function - can be called from dashboard or run standalone"""
    
    # ===== INITIALIZE RISK-FREE RATE AND MARKET RETURN ONCE AT STARTUP =====
    # This runs ONLY ONCE when app starts, not on every interaction
    if 'rf_rate_initialized' not in st.session_state:
        st.session_state.rf_rate_initialized = True
        st.session_state.cached_rf_rate_listed = 6.83  # Default fallback
        st.session_state.cached_rf_rate_screener = 6.83
        st.session_state.cached_rf_rate_unlisted = 6.83
        
        # Calculate actual Sensex CAGR for market return (not hardcoded 12%)
        try:
            sensex_return, _ = get_market_return('%5EBSESN')
            st.session_state.cached_rm_rate_listed = sensex_return
            st.session_state.cached_rm_rate_screener = sensex_return
            st.session_state.cached_rm_rate_unlisted = sensex_return
        except:
            # Fallback only if fetch fails
            st.session_state.cached_rm_rate_listed = 12.0
            st.session_state.cached_rm_rate_screener = 12.0
            st.session_state.cached_rm_rate_unlisted = 12.0
    
    # Initialize session state for peer auto-fetch (Screener mode)
    if 'nse_peers_input' not in st.session_state:
        st.session_state.nse_peers_input = ''
    if 'bse_peers_input' not in st.session_state:
        st.session_state.bse_peers_input = ''
    
    st.title("🏦 DCF Valuation Engine")
    st.markdown("**Listed + Unlisted | Excel-Integrated | Traditional WACC**")

    # Show rate limit status
    if st.session_state.yahoo_request_count > 0:
        with st.expander("📊 API Usage Status", expanded=False):
            st.caption(f"Yahoo Finance requests this hour: {st.session_state.yahoo_request_count}")
            if st.session_state.yahoo_request_count > 10:
                st.warning("⚠️ High API usage detected. Consider waiting if you encounter rate limits.")

    # Display current India inflation rate - USING ACTUAL DATA
    try:
        @st.cache_data(ttl=21600)  # Cache for 6 hours — keeps it current without hammering the source
        def get_india_inflation():
            """
            Fetch the latest India CPI YoY inflation rate live. Returns
            (rate: float, source_label: str), or (None, None) if no live
            source could be reached — we never fall back to a hardcoded
            guess, since a stale/fake number is worse than no banner.
            """
            # Method 1: TradingEconomics — scrapes the page's SEO meta
            # description, which TE keeps in sync with the headline print,
            # e.g. "increased to 4.82 percent in August from 4.44 percent
            # in July of 2026." This is the same page a person would check
            # by hand, so it's live as of whenever TE last updated it.
            try:
                resp = requests.get(
                    "https://tradingeconomics.com/india/inflation-cpi",
                    headers={"User-Agent": "Mozilla/5.0"},
                    timeout=8
                )
                if resp.status_code == 200:
                    soup = BeautifulSoup(resp.text, "html.parser")
                    meta = soup.find("meta", {"name": "description"})
                    desc = meta.get("content", "") if meta else ""
                    matches = re.findall(r'(\d+\.?\d*)\s*percent in (\w+)', desc)
                    year_match = re.search(r'of (\d{4})', desc)
                    if matches:
                        value, month = matches[0]
                        year = year_match.group(1) if year_match else ""
                        label = f"TradingEconomics, {month} {year}".strip().rstrip(',')
                        return float(value), label
            except Exception:
                pass

            # Method 2: World Bank (FP.CPI.TOTL.ZG) — annual, only used if
            # TradingEconomics couldn't be reached.
            try:
                import world_bank_data as wb
                # India CPI inflation: FP.CPI.TOTL.ZG
                df = wb.get_series('FP.CPI.TOTL.ZG', id_or_value='id', simplify_index=True, mrv=1)
                india_data = df[df.index.get_level_values(0) == 'IND']
                if not india_data.empty:
                    inflation = float(india_data.iloc[0])
                    year = india_data.index.get_level_values(1)[0]
                    return inflation, f"World Bank ({year})"
            except Exception:
                pass

            # Method 3: FRED via pandas_datareader — last resort.
            try:
                import pandas_datareader as pdr
                from datetime import datetime, timedelta

                end_date = datetime.now()
                start_date = end_date - timedelta(days=365)

                # India CPI (INDCPIALLMINMEI)
                india_cpi = pdr.get_data_fred('INDCPIALLMINMEI', start_date, end_date)
                if not india_cpi.empty:
                    # Calculate YoY inflation from latest data
                    latest_cpi = india_cpi.iloc[-1].values[0]
                    year_ago_cpi = india_cpi.iloc[0].values[0]
                    inflation = ((latest_cpi - year_ago_cpi) / year_ago_cpi) * 100
                    date = india_cpi.index[-1].strftime('%b %Y')
                    return inflation, f"FRED ({date})"
            except Exception:
                pass

            # No live source reachable — don't show a guessed number.
            return None, None

        india_inflation, inflation_source = get_india_inflation()

        # Display inflation banner (only if a live figure was actually fetched)
        if india_inflation is not None:
            col_inf1, col_inf2, col_inf3 = st.columns([1, 2, 1])
            with col_inf2:
                if "TradingEconomics" in inflation_source:
                    st.success(f"📊 **India CPI Inflation ({inflation_source}):** {india_inflation:.2f}% | Reference for terminal growth")
                else:
                    st.warning(f"📊 **India CPI Inflation ({inflation_source}):** {india_inflation:.2f}% | Reference for terminal growth")
    except Exception as e:
        pass  # Silently fail if inflation display doesn't work

    # ================================
    # SESSION STATE INITIALIZATION
    # ================================
    if 'show_results_listed' not in st.session_state:
        st.session_state.show_results_listed = False
    if 'show_results_unlisted' not in st.session_state:
        st.session_state.show_results_unlisted = False
    if 'fetch_status' not in st.session_state:
        st.session_state.fetch_status = {}
    if 'current_ticker' not in st.session_state:
        st.session_state.current_ticker = None
    if 'yahoo_request_count' not in st.session_state:
        st.session_state.yahoo_request_count = 0

    # Cache status display and control
    col_cache1, col_cache2 = st.columns([3, 1])
    with col_cache1:
        cache_size = len(_TICKER_DATA_CACHE)
        if cache_size > 0:
            st.info(f"🚀 **Rate Limit Protection Active:** {cache_size} data entries cached (prevents duplicate API calls)")
        else:
            st.info("🚀 **Rate Limit Protection:** Caching enabled to prevent duplicate API calls")
    with col_cache2:
        if st.button("🗑️ Clear Cache", help="Clear ticker cache if you need fresh data"):
            clear_ticker_cache()
            st.success("Cache cleared!")
            st.rerun()

    mode = st.radio("Select Mode:", 
                    ["Listed Company (Yahoo Finance)", 
                     "Unlisted Company (Excel Upload)",
                     "Screener Excel Mode (Screener.in Template)",
                     "Bulk Valuation (Screener Web)"], 
                    horizontal=True)

    if mode == "Bulk Valuation (Screener Web)":
        from bulk_valuation import render_bulk_valuation_ui
        render_bulk_valuation_ui(
            classify_business_model=classify_business_model,
            calculate_working_capital_metrics=calculate_working_capital_metrics,
            project_financials=project_financials,
            calculate_wacc=calculate_wacc,
            calculate_dcf_valuation=calculate_dcf_valuation,
            ensure_valid_number=ensure_valid_number,
        )

    elif mode == "Listed Company (Yahoo Finance)":
        st.subheader("📈 Listed Company Valuation")

        # ===== VALUATION DATE — controls what "today" means for Rf, Rm, and =====
        # ===== the current price / multiples used in comparative valuation =====
        st.markdown("### 📅 Valuation Date")
        col_vd1, col_vd2 = st.columns([2, 3])
        with col_vd1:
            valuation_date_listed = st.date_input(
                "Valuation Date",
                value=st.session_state.get('valuation_date_listed', datetime.now().date()),
                key='valuation_date_listed',
                help="Leave as today for a live valuation. Set an earlier date to price the company, its peers, Rf and Rm as of that date instead of today's market — useful for valuing 'as of' a specific past date."
            )
        with col_vd2:
            st.write("")
            if is_valuation_date_today(valuation_date_listed):
                st.caption("💡 Using **live/current** market data (today).")
            else:
                st.info(f"📌 Valuing **as of {valuation_date_listed.strftime('%d-%b-%Y')}** — current price, peer multiples, Rf and Rm will use data as of this date (not today).")
        st.markdown("---")

        # ===== RISK-FREE RATE TICKER - AT THE TOP, ALWAYS VISIBLE =====
        st.markdown("### 🏛️ Risk-Free Rate Configuration")
        st.info("""
        💡 **Three Ways to Set Risk-Free Rate:**
        1. **Fetch from Ticker** (^TNX works best, NIFTYGS10YR.NS has limited data)
        2. **Use Curated Rate** (NIFTYGS10YR.NS → 6.68% based on current market data)
        3. **Manual Entry** (Recommended - type rate in the number field below)
        
        **Default**: 6.83% | **Current India 10Y G-Sec**: 6.68% (as of Feb 13, 2026)
        """)
        
        # Initialize if not exists
        if 'cached_rf_rate_listed' not in st.session_state:
            st.session_state.cached_rf_rate_listed = 6.83
        
        rf_col1, rf_col2, rf_col3 = st.columns([3, 2, 1])
        with rf_col1:
            custom_rf_ticker_listed = st.text_input(
                "Yahoo Finance Ticker for Risk-Free Rate",
                value="NIFTYGS10YR.NS",
                key='custom_rf_ticker_listed_top',
                help="Enter any ticker. For Indian bonds use .NS/.BO suffix. For US indices use %5E prefix (e.g., %5ETNX)"
            )
        with rf_col2:
            st.metric("Current RF Rate", f"{st.session_state.cached_rf_rate_listed:.2f}%")
        with rf_col3:
            st.write("")
            st.write("")
            
            # Add a counter to track button clicks
            if 'rf_fetch_click_count_listed' not in st.session_state:
                st.session_state.rf_fetch_click_count_listed = 0
            
            if st.button("🔄 Fetch", key='refresh_rf_listed_top'):
                # Increment click counter FIRST
                st.session_state.rf_fetch_click_count_listed += 1
                
                # Store debug output in session state so it persists across reruns
                debug_output = []
                debug_output.append(f"🔄 **FETCH BUTTON CLICKED #{st.session_state.rf_fetch_click_count_listed} - LISTED MODE**")
                debug_output.append(f"📝 Input ticker: `{custom_rf_ticker_listed}`")
                
                ticker_to_use = custom_rf_ticker_listed.strip() if custom_rf_ticker_listed.strip() else None
                debug_output.append(f"📝 Ticker to use (after strip): `{ticker_to_use}`")
                
                debug_output.append("⏳ Calling get_risk_free_rate()...")
                _as_of = None if is_valuation_date_today(valuation_date_listed) else valuation_date_listed
                fetched_rate, fetch_debug = get_risk_free_rate(ticker_to_use, as_of_date=_as_of)
                
                # Add all debug messages from the function
                debug_output.extend(fetch_debug)
                
                debug_output.append(f"✅ Function returned: {fetched_rate}%")
                
                debug_output.append(f"💾 Updating session state...")
                debug_output.append(f"   - Before: {st.session_state.get('cached_rf_rate_listed', 'NOT SET')}")
                st.session_state.cached_rf_rate_listed = fetched_rate
                st.session_state.manual_rf_listed = fetched_rate  # CRITICAL: Update widget state too!
                debug_output.append(f"   - After: {st.session_state.cached_rf_rate_listed}")
                
                # No need to delete widget state - we're updating it directly
                debug_output.append(f"🔄 Widget state updated to: {fetched_rate}%")
                
                # Store debug output in session state
                st.session_state.rf_fetch_debug_listed = debug_output
                st.session_state.rf_fetch_success_listed = True
                
                # Rerun to update UI
                st.rerun()
        
        st.markdown("---")
        # ===== END RF RATE CONFIG =====
        
        # DEBUG: Show session state
        
        # ===== MARKET RETURN TICKER - SIMILAR TO RF RATE =====
        st.markdown("### 📈 Market Return Configuration")
        st.info("""
        💡 **Market Return (Rm)** is calculated from historical index returns.
        **Use `%5E` prefix** (e.g., `%5EBSESN` for Sensex, `%5ENSEI` for Nifty)
        
        **Default**: 12.0% | **Sensex avg**: ~12.5% | **Nifty avg**: ~12.8%
        """)
        
        # Initialize if not exists
        if 'cached_rm_rate_listed' not in st.session_state:
            st.session_state.cached_rm_rate_listed = 12.0
        
        rm_col1, rm_col2, rm_col3 = st.columns([3, 2, 1])
        with rm_col1:
            custom_rm_ticker_listed = st.text_input(
                "Yahoo Finance Ticker for Market Return",
                value="%5EBSESN",
                key='custom_rm_ticker_listed_top',
                help="Examples: %5EBSESN (Sensex), %5ENSEI (Nifty), %5ENSEBANK (Nifty Bank). Use %5E prefix!"
            )
        with rm_col2:
            st.metric("Current Market Return", f"{st.session_state.cached_rm_rate_listed:.2f}%")
        with rm_col3:
            st.write("")
            st.write("")
            
            # Add click counter
            if 'rm_fetch_click_count_listed' not in st.session_state:
                st.session_state.rm_fetch_click_count_listed = 0
            
            if st.button("🔄 Fetch", key='refresh_rm_listed_top'):
                # Increment click counter
                st.session_state.rm_fetch_click_count_listed += 1
                
                # Store debug output in session state
                debug_output = []
                debug_output.append(f"🔄 **FETCH BUTTON CLICKED #{st.session_state.rm_fetch_click_count_listed} - LISTED MODE**")
                debug_output.append(f"📝 Input ticker: `{custom_rm_ticker_listed}`")
                
                ticker_to_use = custom_rm_ticker_listed.strip() if custom_rm_ticker_listed.strip() else None
                debug_output.append(f"📝 Ticker to use (after strip): `{ticker_to_use}`")
                
                debug_output.append("⏳ Calling get_market_return()...")
                _as_of = None if is_valuation_date_today(valuation_date_listed) else valuation_date_listed
                fetched_rate, fetch_debug = get_market_return(ticker_to_use, as_of_date=_as_of)
                
                # Add all debug messages
                debug_output.extend(fetch_debug)
                debug_output.append(f"✅ Function returned: {fetched_rate}%")
                
                debug_output.append(f"💾 Updating session state...")
                debug_output.append(f"   - Before: {st.session_state.get('cached_rm_rate_listed', 'NOT SET')}")
                st.session_state.cached_rm_rate_listed = fetched_rate
                st.session_state.manual_rm_listed = fetched_rate  # CRITICAL: Update widget state too!
                debug_output.append(f"   - After: {st.session_state.cached_rm_rate_listed}")
                
                debug_output.append(f"🔄 Widget state updated to: {fetched_rate}%")
                
                # Store debug output in session state
                st.session_state.rm_fetch_debug_listed = debug_output
                st.session_state.rm_fetch_success_listed = True
                
                # Rerun to update UI
                st.rerun()
        
        st.markdown("---")
        # ===== END MARKET RETURN CONFIG =====
    
        col1, col2 = st.columns(2)
    
        with col1:
            # Exchange selection for COMPANY BEING VALUED
            exchange = st.radio(
                "Company Exchange:",
                ["NSE", "BSE", "NASDAQ/NYSE", "LSE", "SSE/HKEX", "Other"],
                horizontal=True,
                help="Select exchange. For NASDAQ/NYSE enter bare ticker (AAPL). LSE: TICKER.L. SSE: TICKER.SS. HKEX: TICKER.HK. Other: enter full YF ticker."
            )

            # Map exchange → yfinance suffix and Indian flag
            _EXCHANGE_MAP = {
                "NSE":       ("NS", True),
                "BSE":       ("BO", True),
                "NASDAQ/NYSE": ("",  False),
                "LSE":       ("L",  False),
                "SSE/HKEX":  ("",   False),   # user must include .SS / .HK in ticker
                "Other":     ("",   False),
            }
            exchange_suffix, _is_indian_exchange = _EXCHANGE_MAP.get(exchange, ("", False))

            _ticker_hints = {
                "NSE":         ("e.g., RELIANCE", "NSE Ticker (no suffix needed):"),
                "BSE":         ("e.g., RELIANCE",  "BSE Ticker (no suffix needed):"),
                "NASDAQ/NYSE": ("e.g., AAPL, MSFT, NVDA", "NASDAQ/NYSE Ticker:"),
                "LSE":         ("e.g., SHEL, HSBA", "LSE Ticker (no suffix — added automatically):"),
                "SSE/HKEX":    ("e.g., 600519.SS or 0700.HK", "SSE/HKEX Ticker (include .SS or .HK):"),
                "Other":       ("e.g., SIE.DE, 005930.KS", "Yahoo Finance Ticker (full, with suffix):"),
            }
            ticker_placeholder, ticker_label = _ticker_hints.get(exchange, ("Enter ticker", "Ticker:"))

            if _is_indian_exchange and TICKER_DROPDOWN_AVAILABLE:
                # NSE/BSE: let the user choose between picking the company by
                # name (dropdown) or typing the ticker/code directly — always
                # both available, not just as a fallback when the list fails
                # to load.
                if st.session_state.get('_listed_prev_exchange') != exchange:
                    st.session_state['listed_company_select'] = ""
                    st.session_state['_listed_prev_exchange'] = exchange

                _entry_mode = st.radio(
                    f"How do you want to select the {exchange} company?",
                    options=["🔍 Search by name", "⌨️ Enter ticker/code directly"],
                    key='listed_company_entry_mode',
                    horizontal=True
                )

                if _entry_mode == "🔍 Search by name":
                    _company_opts = get_company_display_options(exchange)
                    if _company_opts:
                        _company_select_label = f"{exchange} Company (search/select by name):"
                        _selected_company_display = st.selectbox(
                            _company_select_label,
                            options=[""] + _company_opts,
                            key='listed_company_select',
                            help="Start typing the company name to filter the list."
                        )
                        ticker = resolve_company_display(_selected_company_display, exchange)
                    else:
                        _err = get_load_error(exchange)
                        st.warning(f"⚠️ {exchange} company list unavailable — enter ticker manually below.{(' Details: ' + _err) if _err else ''}")
                        ticker = st.text_input(ticker_label, placeholder=ticker_placeholder, key='listed_ticker_manual')
                else:
                    ticker = st.text_input(ticker_label, placeholder=ticker_placeholder, key='listed_ticker_manual')
            else:
                ticker = st.text_input(ticker_label, placeholder=ticker_placeholder)
        
            # Reset analysis state when ticker changes
            if ticker and ticker != st.session_state.get('current_ticker'):
                st.session_state.current_ticker = ticker
                st.session_state.show_results_listed = False
        
            # Historical and projection years
            st.markdown("**⏱️ Time Periods**")
            hist_col1, hist_col2 = st.columns(2)
            with hist_col1:
                historical_years_listed = st.number_input(
                    "Historical Years",
                    min_value=2,
                    max_value=10,
                    value=3,
                    step=1,
                    help="Years of historical data to use for calculations"
                )
            with hist_col2:
                projection_years_listed = st.number_input(
                    "Projection Years",
                    min_value=1,
                    max_value=15,
                    value=5,
                    step=1,
                    help="Years to project into future"
                )
        
            st.markdown("---")
            st.markdown("**📊 Peer Companies (Both Exchanges)**")
        
            # Normalize ticker key (remove suffix for consistent storage)
            ticker_key = ticker.replace('.NS', '').replace('.BO', '') if ticker else ""
        
            # Initialize session state for fetch status
            if 'fetch_status' not in st.session_state:
                st.session_state.fetch_status = {}
        
            # Show current peer state (names for Indian exchanges, raw tickers for worldwide)
            _ww_val = st.session_state.get('worldwide_peers_listed', '')
            _indian_peer_displays_preview = [
                st.session_state.get(f'listed_peer_select_{_i}', '')
                for _i in range(int(st.session_state.get('listed_num_peers', 0) or 0))
            ]
            _indian_peer_displays_preview = [d for d in _indian_peer_displays_preview if d]
            if _indian_peer_displays_preview or _ww_val:
                with st.expander("👁️ Selected / Last Auto-Fetched Peers"):
                    if _indian_peer_displays_preview:
                        st.markdown("**Peers:**")
                        for _d in _indian_peer_displays_preview:
                            st.write(f"• {_d}")
                    if _ww_val:
                        st.markdown("**Worldwide Peers:**")
                        st.code(_ww_val)
        
            # ── Auto-fetch peers button ────────────────────────────────────
            # KEY FIX: Write directly to the widget's session-state key and
            # call st.rerun() — same pattern that works in Screener mode.
            if ticker_key:
                col_btn1, col_btn2, col_btn3 = st.columns([1, 1, 2])
                with col_btn1:
                    _fetch_btn = st.button(
                        "🔍 Auto-Fetch Peers",
                        key='listed_auto_fetch_peers',
                        help="Fetch worldwide industry peers from Yahoo Finance",
                        use_container_width=True,
                        disabled=not PEER_FETCHER_AVAILABLE
                    )
                with col_btn2:
                    _max_peers_listed = st.number_input(
                        "Max to fetch:", min_value=1, value=20, step=1,
                        key='listed_max_peers_fetch',
                        help="How many peers to try to fetch (no upper limit). Yahoo's own APIs may still return fewer if that's all they have."
                    )
                with col_btn3:
                    if PEER_FETCHER_AVAILABLE:
                        st.caption("💡 Fetches worldwide peers; for Indian exchanges also auto-fills the peer name dropdowns below.")
                    else:
                        st.caption("⚠️ utils_peer_fetcher module not found — enter peers manually below.")

                if _fetch_btn and PEER_FETCHER_AVAILABLE:
                    with st.spinner("🔍 Fetching industry peers from Yahoo Finance..."):
                        try:
                            full_ticker = get_ticker_with_exchange(ticker_key, exchange_suffix)
                            st.info(f"🔍 Searching peers for: {full_ticker}")
                            auto_peers = get_industry_peers(full_ticker, max_peers=int(_max_peers_listed), exclude_self=True)

                            if auto_peers and len(auto_peers) > 0:
                                st.success(f"✅ Found {len(auto_peers)} potential peers")

                                if _is_indian_exchange:
                                    # Map each auto-fetched bare ticker/code to its display
                                    # name and pre-fill that many peer-selection dropdowns.
                                    _peer_exch = "NSE" if exchange_suffix == "NS" else "BSE"
                                    if TICKER_DROPDOWN_AVAILABLE:
                                        _auto_displays = [bare_ticker_to_peer_display(p, _peer_exch) for p in auto_peers]
                                    else:
                                        _auto_displays = [f"{p} ({_peer_exch}: {p})" for p in auto_peers]
                                    st.session_state['listed_num_peers'] = len(_auto_displays)
                                    for _i, _d in enumerate(_auto_displays):
                                        st.session_state[f'listed_peer_select_{_i}'] = _d
                                    st.rerun()
                                else:
                                    # Non-Indian exchange: append the correct suffix to bare peer tickers.
                                    # NASDAQ/NYSE has no suffix (empty string) — peers stay bare.
                                    # LSE has suffix "L" — append .L to each peer.
                                    # SSE/HKEX and Other have empty suffix but user enters full ticker
                                    # (e.g. 600519.SS, 0700.HK, SIE.DE) — detect suffix from input ticker.
                                    if exchange_suffix:
                                        # e.g. LSE → suffix "L" → "SHEL.L"
                                        peer_suffix = f".{exchange_suffix}"
                                    else:
                                        # Detect from the user's ticker (e.g. "600519.SS" → ".SS")
                                        _dot = full_ticker.rfind('.')
                                        peer_suffix = full_ticker[_dot:] if _dot > 0 else ""

                                    if peer_suffix:
                                        formatted_peers = [
                                            p if p.endswith(peer_suffix) else f"{p}{peer_suffix}"
                                            for p in auto_peers
                                        ]
                                    else:
                                        formatted_peers = auto_peers  # NASDAQ/NYSE bare tickers

                                    st.session_state['worldwide_peers_listed'] = ",".join(formatted_peers)
                                    st.rerun()
                            else:
                                st.warning("⚠️ No peers found for this ticker. Enter manually below.")
                        except Exception as e:
                            st.error(f"❌ Error fetching peers: {str(e)}")
                            import traceback
                            with st.expander("🔍 Show error details"):
                                st.code(traceback.format_exc())
            # ── End auto-fetch ─────────────────────────────────────────────
        
            if _is_indian_exchange and TICKER_DROPDOWN_AVAILABLE:
                # Indian exchange: choose how many peers, then pick each —
                # by name (dropdown) or by typing the ticker/code directly,
                # your choice per peer, not forced into the dropdown.
                if 'listed_num_peers' not in st.session_state:
                    st.session_state['listed_num_peers'] = 0

                num_peers_listed = st.number_input(
                    "Number of Peers:",
                    min_value=0,
                    step=1,
                    key='listed_num_peers',
                    help="How many peer companies to include (no upper limit). Use Auto-Fetch above to populate automatically, or set this and pick each peer by name/ticker below."
                )

                _peer_entry_mode_listed = st.radio(
                    "How do you want to select peers?",
                    options=["🔍 Search by name", "⌨️ Enter ticker/code directly"],
                    key='listed_peer_entry_mode',
                    horizontal=True
                ) if int(num_peers_listed) > 0 else "🔍 Search by name"

                _peer_options_listed = get_peer_display_options() if _peer_entry_mode_listed == "🔍 Search by name" else []
                _selected_peer_displays_listed = []
                _manual_peer_tickers_listed = []
                for _i in range(int(num_peers_listed)):
                    if _peer_entry_mode_listed == "🔍 Search by name":
                        _pk = f'listed_peer_select_{_i}'
                        if _pk not in st.session_state:
                            st.session_state[_pk] = ""
                        _sel = st.selectbox(
                            f"Peer {_i + 1}:",
                            options=[""] + _peer_options_listed,
                            key=_pk,
                            help="Start typing the company name to filter the list."
                        )
                        _selected_peer_displays_listed.append(_sel)
                    else:
                        _pkm = f'listed_peer_manual_{_i}'
                        _tkr = st.text_input(
                            f"Peer {_i + 1} ticker/code:",
                            placeholder="e.g., TCS or TCS.NS",
                            key=_pkm,
                            help="Bare ticker/code is assumed to be on the same exchange selected above; a full Yahoo ticker (e.g. TCS.NS) also works."
                        )
                        if _tkr.strip():
                            _t = _tkr.strip().upper()
                            if '.' not in _t:
                                _t = f"{_t}.{exchange_suffix}"
                            _manual_peer_tickers_listed.append(_t)

                comp_tickers_listed = ",".join(
                    _manual_peer_tickers_listed if _peer_entry_mode_listed != "🔍 Search by name"
                    else [resolve_peer_display_to_full_ticker(d) for d in _selected_peer_displays_listed if d]
                )
            elif _is_indian_exchange:
                # Fallback if the dropdown utility module failed to load
                nse_peers = st.text_input(
                    "NSE Peer Tickers (comma-separated):",
                    placeholder="e.g., TCS, INFY, WIPRO",
                    key='nse_peers_listed',
                    help="Enter manually — name-based dropdown unavailable"
                )
                bse_peers = st.text_input(
                    "BSE Peer Tickers (comma-separated):",
                    placeholder="e.g., SUNPHARMA, TATAMOTORS",
                    key='bse_peers_listed',
                    help="Enter BSE-listed peer companies"
                )
                comp_tickers_listed = ""
                if nse_peers.strip():
                    nse_list = [f"{t.strip()}.NS" if not any(t.strip().endswith(s) for s in ('.NS','.BO','.SS')) else t.strip()
                                for t in nse_peers.split(',') if t.strip()]
                    comp_tickers_listed = ",".join(nse_list)
                if bse_peers.strip():
                    bse_list = [f"{t.strip()}.BO" if not any(t.strip().endswith(s) for s in ('.NS','.BO','.SS')) else t.strip()
                                for t in bse_peers.split(',') if t.strip()]
                    if comp_tickers_listed:
                        comp_tickers_listed += "," + ",".join(bse_list)
                    else:
                        comp_tickers_listed = ",".join(bse_list)
            else:
                # Non-Indian exchange: single worldwide peer box; tickers used as-is
                # (user can mix AAPL, MSFT, INFY.NS, 600519.SS etc.)
                worldwide_peers_raw = st.text_input(
                    "Peer Tickers — worldwide (comma-separated):",
                    placeholder="e.g., MSFT, GOOGL, INFY.NS, 600519.SS",
                    key='worldwide_peers_listed',
                    help="Enter tickers exactly as they appear on Yahoo Finance. NASDAQ/NYSE: bare ticker. NSE: TICKER.NS. BSE: TICKER.BO. SSE: TICKER.SS."
                )
                # Tickers are used as-is; no suffix is appended automatically
                comp_tickers_listed = ",".join(
                    [t.strip() for t in worldwide_peers_raw.split(',') if t.strip()]
                ) if worldwide_peers_raw.strip() else ""
    
        with col2:
            tax_rate = st.number_input("Tax Rate (%):", min_value=0.0, max_value=100.0, value=25.0, step=0.5, key='listed_tax')
            terminal_growth = st.number_input("Terminal Growth Rate (%):", min_value=0.0, max_value=10.0, value=4.0, step=0.5, key='listed_tg')
            
            # Risk-free rate override
            _RF_LABEL = {
                "NSE": "Risk-Free Rate (India G-Sec 10Y)",
                "BSE": "Risk-Free Rate (India G-Sec 10Y)",
                "NASDAQ/NYSE": "Risk-Free Rate (US Treasury 10Y)",
                "LSE": "Risk-Free Rate (UK Gilt 10Y)",
                "SSE/HKEX": "Risk-Free Rate (10Y Bond)",
                "Other": "Risk-Free Rate (10Y Bond)",
            }
            st.markdown(f"**🏛️ {_RF_LABEL.get(exchange, 'Risk-Free Rate (10Y)')}**")
            
            # Default Rf/Rm depend on market
            _DEFAULT_RF = {
                "NSE": 6.83, "BSE": 6.83,
                "NASDAQ/NYSE": 4.25, "LSE": 4.10,
                "SSE/HKEX": 2.50, "Other": 4.50,
            }
            _DEFAULT_RM = {
                "NSE": 12.0, "BSE": 12.0,
                "NASDAQ/NYSE": 10.5, "LSE": 9.5,
                "SSE/HKEX": 9.0, "Other": 10.0,
            }
            _rf_default = _DEFAULT_RF.get(exchange, 6.83)
            _rm_default = _DEFAULT_RM.get(exchange, 12.0)

            # Initialize widget state from session if not already set
            if 'manual_rf_listed' not in st.session_state:
                st.session_state.manual_rf_listed = st.session_state.get('cached_rf_rate_listed', _rf_default)
            
            manual_rf_rate = st.number_input(
                f"Risk-Free Rate (%)",
                min_value=0.0,
                max_value=20.0,
                value=st.session_state.manual_rf_listed,
                step=0.1,
                key='manual_rf_listed',
                help="Auto-fetched from ticker above. You can manually edit this value."
            )
            
            # Market Return input
            st.markdown("**📈 Market Return (Rm)**")
            
            # Initialize widget state from session if not already set
            if 'manual_rm_listed' not in st.session_state:
                st.session_state.manual_rm_listed = st.session_state.get('cached_rm_rate_listed', _rm_default)
            
            manual_rm_rate = st.number_input(
                f"Market Return (%)",
                min_value=0.0,
                max_value=50.0,
                value=st.session_state.manual_rm_listed,
                step=0.1,
                key='manual_rm_listed',
                help="Auto-fetched from ticker above. You can manually edit this value."
            )
        
            # Manual discount rate override
            st.markdown("**💰 Discount Rate Override (Optional)**")
            manual_discount_rate = st.number_input(
                "Manual Discount Rate (%)",
                min_value=0.0,
                max_value=50.0,
                value=0.0,
                step=0.5,
                key='manual_discount_listed',
                help="⚠️ Override WACC calculation. Leave at 0 to use auto-calculated WACC. Use this if you want to use a specific discount rate instead of WACC."
            )
            if manual_discount_rate > 0:
                st.info(f"💡 Using manual discount rate: {manual_discount_rate:.2f}% (Overriding WACC)")

            # ── Beta Calculation Period ────────────────────────────────────
            st.markdown("**📅 Beta Calculation Period**")
            st.caption("Daily returns only. Defaults to 3 years ending today if left unchanged.")
            _beta_default_end_listed   = datetime.now().date()
            _beta_default_start_listed = (_beta_default_end_listed.replace(year=_beta_default_end_listed.year - 3))
            _bc1, _bc2 = st.columns(2)
            with _bc1:
                beta_start_date_listed = st.date_input(
                    "From", value=_beta_default_start_listed,
                    key='beta_start_listed',
                    help="Start date for beta regression window"
                )
            with _bc2:
                beta_end_date_listed = st.date_input(
                    "To", value=_beta_default_end_listed,
                    key='beta_end_listed',
                    help="End date for beta regression window"
                )
            if beta_start_date_listed >= beta_end_date_listed:
                st.error("⚠️ Beta start date must be before end date.")
            # ── End Beta Period ────────────────────────────────────────────
    
        with st.expander("⚙️ Advanced Projection Assumptions - FULL CONTROL"):
            st.info("💡 **Complete Control:** Override ANY projection parameter below. Leave at 0 or blank for auto-calculation from historical data.")
        
            st.markdown("### 📊 Shares Outstanding Override")
            st.caption("For newly listed companies or when Yahoo Finance fails to fetch shares")
            manual_shares_override = st.number_input(
                "Manual Shares Outstanding (Absolute Number)",
                min_value=0,
                value=0,
                step=1000000,
                format="%d",
                help="⚠️ Use ONLY if Yahoo Finance fails to fetch shares. Enter absolute number (e.g., 50000000 for 5 crore shares). Leave at 0 to use auto-fetched value.",
                key="manual_shares_listed"
            )
            if manual_shares_override > 0:
                st.success(f"✅ Manual override active: **{manual_shares_override:,}** shares ({manual_shares_override/10000000:.2f} Crore)")
        
            st.markdown("---")
            st.markdown("### 📊 Revenue & Growth")
            col1, col2, col3 = st.columns(3)
            with col1:
                rev_growth_override_listed = st.number_input(
                    "Revenue Growth (%/year)", 
                    min_value=0.0, max_value=100.0, value=0.0, step=0.5,
                    key='listed_rev_growth',
                    help="0 = Auto from historical CAGR"
                )
            with col2:
                opex_margin_override_listed = st.number_input(
                    "Operating Expense Margin (%)", 
                    min_value=0.0, max_value=100.0, value=0.0, step=0.5,
                    key='listed_opex_margin',
                    help="0 = Auto from historical average"
                )
            with col3:
                ebitda_margin_override = st.number_input(
                    "EBITDA Margin (%)", 
                    min_value=0.0, max_value=100.0, value=0.0, step=0.5,
                    key='listed_ebitda',
                    help="0 = Calculated as Revenue - OpEx"
                )

            # ── PER-YEAR PARAMETERS (Listed) ──────────────────────────────
            st.markdown("### 📅 Year-by-Year Revenue Growth & EBITDA Margin")
            st.caption("Set individual values per projected year. Leave at 0 to fall back to the single overrides above (or auto-calc from history).")
            rev_growth_per_year_listed = []
            ebitda_margin_per_year_listed = []
            yr_cols_listed = st.columns(min(int(projection_years_listed), 5))
            for _yi in range(int(projection_years_listed)):
                _col = yr_cols_listed[_yi % len(yr_cols_listed)]
                with _col:
                    st.markdown(f"**Year {_yi + 1}**")
                    _rg = st.number_input(
                        f"Rev Growth %",
                        min_value=0.0, max_value=200.0, value=0.0, step=0.5,
                        key=f'listed_rg_yr{_yi+1}',
                        help="0 = use global override / auto"
                    )
                    _em = st.number_input(
                        f"EBITDA Margin %",
                        min_value=0.0, max_value=100.0, value=0.0, step=0.5,
                        key=f'listed_em_yr{_yi+1}',
                        help="0 = use global override / auto"
                    )
                    rev_growth_per_year_listed.append(_rg)
                    ebitda_margin_per_year_listed.append(_em)
            # ── END PER-YEAR ───────────────────────────────────────────────
            
            # ── MULTI-STAGE GROWTH PROFILE (optional, overrides the above) ──
            _ms_enabled_listed, _ms_growth_listed, _ms_margin_listed = render_multistage_growth_ui(
                key_prefix="ms_listed",
                total_years=projection_years_listed,
                default_stage1_growth=rev_growth_override_listed if rev_growth_override_listed > 0 else 15.0,
                terminal_growth=terminal_growth,
                default_stage1_margin=ebitda_margin_override if ebitda_margin_override > 0 else 20.0
            )
            if _ms_enabled_listed:
                rev_growth_per_year_listed = _ms_growth_listed
                if _ms_margin_listed:
                    ebitda_margin_per_year_listed = _ms_margin_listed
            # ── END MULTI-STAGE ──────────────────────────────────────────────

            st.markdown("### 🏗️ CapEx & Depreciation")
            col4, col5, col6 = st.columns(3)
            with col4:
                capex_ratio_override_listed = st.number_input(
                    "CapEx/Revenue (%)", 
                    min_value=0.0, max_value=50.0, value=0.0, step=0.5,
                    key='listed_capex_ratio',
                    help="0 = Auto from historical average"
                )
            with col5:
                depreciation_rate_override = st.number_input(
                    "Depreciation Rate (%)", 
                    min_value=0.0, max_value=30.0, value=0.0, step=0.5,
                    key='listed_dep_rate',
                    help="0 = Auto calculated"
                )
            with col6:
                depreciation_method = st.selectbox(
                    "Depreciation Method",
                    ["Auto", "% of Fixed Assets", "% of Revenue", "Absolute Value"],
                    key='listed_dep_method'
                )
        
            st.markdown("### 💰 Working Capital Management")
            col7, col8, col9 = st.columns(3)
            with col7:
                inventory_days_override = st.number_input(
                    "Inventory Days", 
                    min_value=0.0, max_value=365.0, value=0.0, step=1.0,
                    key='listed_inv_days',
                    help="0 = Auto from historical average"
                )
            with col8:
                debtor_days_override = st.number_input(
                    "Debtor/Receivables Days", 
                    min_value=0.0, max_value=365.0, value=0.0, step=1.0,
                    key='listed_deb_days',
                    help="0 = Auto from historical average"
                )
            with col9:
                creditor_days_override = st.number_input(
                    "Creditor/Payables Days", 
                    min_value=0.0, max_value=365.0, value=0.0, step=1.0,
                    key='listed_cred_days',
                    help="0 = Auto from historical average"
                )
        
            st.markdown("### 📈 Tax & Interest")
            col10, col11, col12 = st.columns(3)
            with col10:
                interest_rate_override = st.number_input(
                    "Interest Rate (%)", 
                    min_value=0.0, max_value=30.0, value=0.0, step=0.25,
                    key='listed_interest',
                    help="0 = Auto calculated from Debt"
                )
            with col11:
                tax_shield = st.checkbox(
                    "Include Tax Shield",
                    value=True,
                    key='listed_tax_shield',
                    help="Apply tax benefit on interest expense"
                )
            with col12:
                working_capital_as_pct_revenue = st.number_input(
                    "Working Capital % of Revenue", 
                    min_value=0.0, max_value=50.0, value=0.0, step=0.5,
                    key='listed_wc_pct',
                    help="0 = Calculate from Inv+Deb-Cred days"
                )
    
        # ================================
        # BANK/NBFC PARAMETERS (BEFORE FETCH)
        # ================================
        with st.expander("🏦 Bank/NBFC Valuation Parameters (Adjust if company is a Bank/NBFC)"):
            st.info("💡 **Note:** These parameters are used ONLY if the company is classified as a Bank/NBFC/Interest-Dominant entity after analysis. Adjust them before clicking 'Fetch & Analyze'.")
        
            col1, col2, col3, col4 = st.columns(4)
        
            with col1:
                terminal_growth_bank = st.number_input(
                    "Terminal Growth (%)",
                    min_value=0.0,
                    max_value=10.0,
                    value=3.5,
                    step=0.5,
                    key='terminal_growth_bank_pre',
                    help="Long-term perpetual growth rate for banks"
                )
        
            with col2:
                projection_years_bank = st.number_input(
                    "Projection Years",
                    min_value=3,
                    max_value=15,
                    value=5,
                    step=1,
                    key='projection_years_bank_pre',
                    help="Years to project for bank valuations"
                )
        
            with col3:
                assumed_roe_bank = st.number_input(
                    "Assumed ROE (%)",
                    min_value=5.0,
                    max_value=30.0,
                    value=15.0,
                    step=1.0,
                    key='assumed_roe_bank_pre',
                    help="Return on Equity assumption (15 = auto-calculated)"
                )
        
            with col4:
                cost_of_equity_override = st.number_input(
                    "Cost of Equity (%)",
                    min_value=0.0,
                    max_value=25.0,
                    value=0.0,
                    step=0.5,
                    key='cost_of_equity_override_pre',
                    help="Override Ke (0 = use calculated CAPM)"
                )
        
            col5, col6, col7 = st.columns(3)
        
            with col5:
                div_growth_bank = st.number_input(
                    "Dividend Growth (%)",
                    min_value=-10.0,
                    max_value=20.0,
                    value=0.0,
                    step=1.0,
                    key='div_growth_bank_pre',
                    help="0 = use calculated from history"
                )
        
            with col6:
                payout_ratio_bank = st.number_input(
                    "Payout Ratio (%)",
                    min_value=0.0,
                    max_value=90.0,
                    value=0.0,
                    step=5.0,
                    key='payout_ratio_bank_pre',
                    help="0 = use calculated from history"
                )
        
            with col7:
                st.metric("Info", "0 = Auto", help="Setting 0 uses auto-calculated values from historical data")
    
        # ================================
        # DDM & RIM PARAMETERS FOR NON-BANKING COMPANIES
        # ================================
        with st.expander("💎 DDM & RIM Parameters for Non-Banking Companies"):
            st.info("💡 **Note:** These parameters apply to Dividend Discount Model and Residual Income Model for non-banking companies. Leave at 0 for auto-calculation from actual data.")
        
            st.markdown("### 💰 Dividend Discount Model (DDM) Parameters")
            col1, col2 = st.columns(2)
        
            with col1:
                ddm_dividend_growth_override = st.number_input(
                    "Dividend Growth Rate (%)",
                    min_value=0.0,
                    max_value=30.0,
                    value=0.0,
                    step=0.5,
                    key='ddm_div_growth',
                    help="0 = Auto-calculate from historical dividend data. Used in Gordon Growth Model."
                )
        
            with col2:
                ddm_payout_ratio_override = st.number_input(
                    "Payout Ratio (%)",
                    min_value=0.0,
                    max_value=100.0,
                    value=0.0,
                    step=5.0,
                    key='ddm_payout',
                    help="0 = Auto-calculate from historical data. % of earnings paid as dividends."
                )
        
            st.markdown("### 🏢 Residual Income Model (RIM) Parameters")
            col3, col4, col5 = st.columns(3)
        
            with col3:
                rim_assumed_roe_override = st.number_input(
                    "Assumed ROE (%)",
                    min_value=0.0,
                    max_value=50.0,
                    value=0.0,
                    step=1.0,
                    key='rim_roe',
                    help="0 = Auto-calculate from historical data. Return on Equity assumption."
                )
        
            with col4:
                rim_terminal_growth_override = st.number_input(
                    "Terminal Growth (%)",
                    min_value=0.0,
                    max_value=10.0,
                    value=0.0,
                    step=0.5,
                    key='rim_terminal_growth',
                    help="0 = Use same as DCF terminal growth. Long-term perpetual growth rate."
                )
        
            with col5:
                rim_projection_years_override = st.number_input(
                    "Projection Years",
                    min_value=0,
                    max_value=15,
                    value=0,
                    step=1,
                    key='rim_proj_years',
                    help="0 = Use same as DCF projection years. Number of years to project."
                )
    
        if ticker:
            # Check if we're approaching rate limits
            if st.session_state.yahoo_request_count > 15:
                st.warning(f"""
                ⚠️ **High API Usage Warning**
            
                You've made {st.session_state.yahoo_request_count} Yahoo Finance requests this hour.
                You may encounter rate limits soon.
            
                **Recommendations:**
                - Wait a few minutes between requests
                - Use cached results when possible
                - Consider using Screener.in as alternative data source (checkbox below)
                """)
        
            # DATA SOURCE SELECTION CHECKBOXES
            st.markdown("---")
            st.markdown("### 📊 Data Source Options")
        
            col_src1, col_src2 = st.columns(2)
        
            with col_src1:
                use_screener_data = st.checkbox(
                    "🌐 Use Screener.in for company data",
                    value=False,
                    key='use_screener_data',
                    help="Fetch financial data from Screener.in instead of Yahoo Finance. Useful when Yahoo Finance is rate-limited or data is missing."
                )
                if use_screener_data:
                    st.info("📌 Company financials will be fetched from Screener.in (Indian stock screener)")
        
            with col_src2:
                use_screener_peers = st.checkbox(
                    "🔗 Use Screener.in for peer data",
                    value=False,
                    key='use_screener_peers',
                    help="Fetch peer company data from Screener.in for comparative valuation. Works independently of main data source."
                )
                if use_screener_peers:
                    st.info("📌 Peer comparison data will be fetched from Screener.in")
        
            st.markdown("---")
            
            # Stock Price Comparison Feature Toggle - compact, before fetch button
            enable_stock_comparison = st.checkbox(
                "📈 Show Stock Price vs Revenue & EPS chart",
                value=False,
                key='enable_stock_comparison_listed',
                help="Compare stock price with financials (max 4 years)"
            )
        
            # INPUT CHANGE DETECTION - Reset results if key inputs change
            current_inputs_listed = {
                'ticker': ticker,
                'projection_years': projection_years_listed,
                'terminal_growth': terminal_growth,
                'tax_rate': tax_rate,
                'manual_discount': manual_discount_rate if manual_discount_rate > 0 else None
            }
            
            # Check if inputs changed
            if 'previous_inputs_listed' in st.session_state:
                if st.session_state.previous_inputs_listed != current_inputs_listed:
                    # Inputs changed - clear results
                    st.session_state.show_results_listed = False
            
            # Store current inputs
            st.session_state.previous_inputs_listed = current_inputs_listed
        
            st.markdown("---")
            scenario_case_listed, scenario_pct_listed, scenario_drivers_listed = render_scenario_settings_ui("scenario_listed")
        
            st.markdown("---")
            st.markdown("### 🎯 Ready to Run Valuation")
            st.info("💡 **Click the button below to run valuation.** Results will appear only after clicking.")
            
            col_fetch1, col_fetch2, col_fetch3 = st.columns([2, 1, 1])
        
            with col_fetch1:
                if st.button("🚀 Fetch & Analyze", type="primary", key="fetch_analyze_listed"):
                    st.session_state.show_results_listed = True
        
            with col_fetch2:
                if st.button("🗑️ Clear Results", help="Clear current valuation results"):
                    st.session_state.show_results_listed = False
                    st.success("✅ Results cleared!")
                    st.rerun()
            
            with col_fetch3:
                if st.button("🔄 Reset Cache", help="Clear cached data to force fresh fetch"):
                    st.cache_data.clear()
                    st.success("✅ Cache cleared!")
                    st.rerun()
        
            if st.session_state.get('show_results_listed', False):
                # Safe defaults — prevent UnboundLocalError if any branch doesn't assign these
                current_price = 0
                _ticker_csym = '₹'  # Will be overwritten once yahoo_data/info is available

                # Determine data source based on checkbox
                use_screener = st.session_state.get('use_screener_data', False)
            
            
                if use_screener:
                    # SCREENER.IN DATA SOURCE
                    with st.spinner("🌐 Fetching data from Screener.in..."):
                        try:
                            # Remove exchange suffix for Screener.in
                            ticker_clean = ticker.replace('.NS', '').replace('.BO', '')
                        
                            screener_data = fetch_screener_financials(ticker_clean, num_years=5)
                        
                            if screener_data and screener_data.get('financials'):
                                st.success(f"✅ Data fetched from Screener.in for {screener_data.get('company_name', ticker_clean)}")
                            
                                # Try to get current price from Screener
                                current_price = screener_data.get('current_price', 0)
                                if current_price == 0:
                                    st.info("💡 Fetching current price from Yahoo Finance...")
                                    try:
                                        pass  # yf available via module-level shim (yf_ratelimit)
                                        yf_ticker = yf.Ticker(ticker)
                                        yf_info = yf_ticker.info if yf_ticker else None
                                        if yf_info:
                                            current_price = yf_info.get('currentPrice', 0) or yf_info.get('regularMarketPrice', 0)
                                        if current_price > 0:
                                            st.success(f"✅ Current Price: {_ticker_csym}{current_price:.2f} (Yahoo Finance)")
                                    except:
                                        st.warning("⚠️ Could not fetch current price - will use 0")
                            
                                # Convert to Yahoo Finance compatible format
                                yahoo_data = {
                                    'info': {
                                        'longName': screener_data.get('company_name', ticker_clean),
                                        'currentPrice': current_price,
                                        'sharesOutstanding': screener_data.get('shares', 0)
                                    },
                                    'income_statement': pd.DataFrame(),
                                    'balance_sheet': pd.DataFrame(),
                                    'cash_flow': pd.DataFrame(),
                                    'shares': screener_data.get('shares', 0),
                                    'shares_source': 'Screener.in (derived from EPS)',
                                    '_screener_financials': screener_data['financials'],  # Stash for extraction
                                    '_data_source': 'screener'
                                }
                            
                                shares = screener_data.get('shares', 0)
                                shares_source = 'Screener.in (EPS-based)'
                                company_name = screener_data.get('company_name', ticker_clean)
                                error = None
                                # Set currency symbol now that yahoo_data is ready
                                _ticker_csym = get_currency_symbol(yahoo_data.get('info'))
                            
                                # BUGFIX: Validate shares from Screener and provide feedback
                                if shares > 0:
                                    st.success(f"✅ Shares Outstanding fetched: **{shares:,}** from Screener.in (EPS-based)")
                                else:
                                    st.warning("⚠️ Screener.in returned shares outstanding = 0. This may be due to missing or incomplete EPS data.")
                                    st.info("💡 **Troubleshooting:**\n- The stock may not have sufficient historical data on Screener.in\n- Try using manual override above\n- Or uncheck 'Use Screener.in' to fetch from Yahoo Finance")
                            
                            else:
                                st.error("❌ Could not fetch data from Screener.in. Please check ticker symbol or try Yahoo Finance.")
                                st.stop()
                    
                        except Exception as e:
                            st.error(f"❌ Error fetching from Screener.in: {str(e)}")
                            st.info("💡 Try using Yahoo Finance instead (uncheck the Screener.in option)")
                            st.stop()
            
                if not use_screener:
                    # YAHOO FINANCE DATA SOURCE (DEFAULT)
                    with st.spinner("📊 Fetching data from Yahoo Finance..."):
                        yahoo_data, error = fetch_yahoo_financials(ticker, exchange_suffix)
                    
                        if error:
                            st.error(error)
                            st.info("💡 **Suggestion:** Try checking the 'Use Screener.in' option above to fetch from an alternative data source")
                            st.stop()
                    
                        shares = yahoo_data.get('shares', 0)
                        shares_source = yahoo_data.get('shares_source', 'Unknown')
                        company_name = yahoo_data['info'].get('longName', ticker)
                        yahoo_data['_data_source'] = 'yahoo'
                        # Set currency symbol and current price as soon as info is available
                        _ticker_csym = get_currency_symbol(yahoo_data.get('info'))
                        current_price = yahoo_data['info'].get('currentPrice', 0) or yahoo_data['info'].get('regularMarketPrice', 0) or 0
                    
                        # BUGFIX: Add validation and fallback for shares
                        if shares > 0:
                            st.success(f"✅ Shares Outstanding fetched: **{shares:,}** from {shares_source}")
                        else:
                            st.warning(f"⚠️ Shares outstanding = 0. Source attempted: {shares_source}")
                            # Try fallback calculation using Market Cap / Price
                            info = yahoo_data.get('info', {})
                            if 'marketCap' in info and 'currentPrice' in info:
                                market_cap = info.get('marketCap', 0)
                                current_price = info.get('currentPrice', 0)
                                if market_cap > 0 and current_price > 0:
                                    shares = int(market_cap / current_price)
                                    shares_source = "Calculated (Market Cap ÷ Current Price)"
                                    st.info(f"🔄 Recalculated shares: **{shares:,}** using {shares_source}")
            
                # Common processing for both Screener.in and Yahoo Finance
                # Apply manual override if provided
                if manual_shares_override > 0:
                    shares = manual_shares_override
                    shares_source = "Manual Override (User Input)"
                    st.warning(f"⚠️ Using manual shares override: **{shares:,}** shares")
                elif shares == 0:
                    st.error("❌ **CRITICAL:** Could not fetch shares outstanding!")
                    st.warning("🔧 **Action Required:** Please enter shares outstanding manually using the override field above.")
                    st.info("💡 **Where to find shares:**\n- Company's latest annual report\n- BSE/NSE company page\n- Screener.in or Moneycontrol")
                    st.stop()
            
                st.success(f"✅ Loaded data for **{company_name}**")
            
                # Show shares with source
                shares_in_crore = shares / 10000000
                col_sh1, col_sh2 = st.columns(2)
                with col_sh1:
                    st.metric("📊 Shares Outstanding", f"{shares:,}", help=f"Source: {shares_source}")
                with col_sh2:
                    st.metric("📊 Shares (Crore)", f"{shares_in_crore:.2f} Cr")
            
                if shares_source != "Manual Override (User Input)" and shares_source != "Direct (sharesOutstanding)":
                    st.caption(f"ℹ️ Shares source: {shares_source}")
            
                # Extract financials with user-selected historical years
                financials = extract_financials_listed(yahoo_data, num_years=historical_years_listed)
            
                if financials is None:
                    st.error("Failed to extract financial data")
                    st.stop()
            
                # ================================
                # BUSINESS MODEL CLASSIFICATION (RULEBOOK SECTION 2)
                # ================================
                st.markdown("---")
                st.subheader("🏢 Business Model Classification")
                
                classification = classify_business_model(
                    financials, 
                    income_stmt=yahoo_data['income_statement'], 
                    balance_sheet=yahoo_data['balance_sheet']
                )
                
                # Show classification and check if FCFF DCF is allowed
                should_stop = show_classification_warning(classification)
                
                if should_stop:
                    # For Interest-Dominant entities, run alternative valuation methods
                    st.markdown("---")
                    st.header("🏦 Bank/NBFC Valuation Methods")
                    st.info("Running alternative valuation methods appropriate for interest-dominant entities...")
                
                    # Initialize session state for bank parameters if not exists
                    if 'bank_params_applied' not in st.session_state:
                        st.session_state.bank_params_applied = {
                            'terminal_growth_bank': 3.5,
                            'projection_years_bank': 5,
                            'assumed_roe_bank': 15.0,
                            'cost_of_equity_override': 0.0,
                            'div_growth_bank': 0.0,
                            'payout_ratio_bank': 0.0
                        }
                
                    # Use parameters set BEFORE Fetch button (from the expandable section)
                    # No need for Apply Changes button - parameters are already set!
                
                    # Get current price (refresh from live data; fall back to value already fetched above)
                    stock = get_cached_ticker(get_ticker_with_exchange(ticker, exchange_suffix))
                    info = stock.info if stock else None
                    _bank_price = info.get('currentPrice', 0) if info else 0
                    if _bank_price > 0:
                        current_price = _bank_price
                    # _ticker_csym: update from live info if available, else keep value set earlier
                    if info:
                        _ticker_csym = get_currency_symbol(info)
                
                    # Calculate Ke for bank valuations
                    wacc_details = calculate_wacc(financials, tax_rate, peer_tickers=None, manual_rf_rate=manual_rf_rate, manual_rm_rate=manual_rm_rate,
                                                  beta_start_date=beta_start_date_listed, beta_end_date=beta_end_date_listed)
                    beta = get_stock_beta(get_ticker_with_exchange(ticker, exchange_suffix), period_years=3,
                                         beta_start_date=beta_start_date_listed, beta_end_date=beta_end_date_listed)
                    wacc_details['beta'] = beta
                    wacc_details['ke'] = wacc_details['rf'] + (beta * (wacc_details['rm'] - wacc_details['rf']))
                
                    # Use override or calculated Ke
                    if cost_of_equity_override > 0:
                        cost_of_equity = cost_of_equity_override
                    else:
                        cost_of_equity = wacc_details['ke']
                
                    # Run all bank valuation methods with user parameters
                    ri_model = calculate_residual_income_model(
                        financials, shares, cost_of_equity,
                        terminal_growth=terminal_growth_bank,
                        projection_years=projection_years_bank,
                        assumed_roe=assumed_roe_bank if assumed_roe_bank != 15 else None
                    )
                
                    ddm_model = calculate_dividend_discount_model(
                        financials, shares, cost_of_equity,
                        ticker=ticker,
                        div_growth_override=div_growth_bank if div_growth_bank != 0 else None,
                        payout_ratio_override=payout_ratio_bank if payout_ratio_bank != 0 else None,
                        exchange_suffix=exchange_suffix
                    )
                
                    pb_roe_model = calculate_pb_roe_valuation(
                        financials, shares, cost_of_equity,
                        assumed_roe=assumed_roe_bank if assumed_roe_bank != 15 else None
                    )
                
                    rel_val = calculate_relative_valuation(ticker, financials, shares, peer_tickers=comp_tickers_listed, exchange_suffix=exchange_suffix)
                
                    # ============================================
                    # BANK DCF OPTION - Using NII as Revenue
                    # ============================================
                    st.markdown("---")
                    st.markdown("### 🏦 **DCF for Banks (Experimental)**")
                    st.info("💡 **New Feature:** Calculate DCF for banks by treating Net Interest Income (NII) as revenue and interest expense as operating cost.")
                
                    bank_dcf_result = None
                    with st.expander("🔧 Enable DCF for Banks", expanded=False):
                        st.markdown("""
                        **How Bank DCF Works:**
                        - ✅ Revenue = Net Interest Income (Interest Income - Interest Expense)
                        - ✅ Operating Expenses = Non-Interest Expenses
                        - ✅ WACC = Bank-specific calculation (Cost of Funds methodology)
                        - ✅ FCFF = NII - OpEx - Taxes (no CapEx/WC for banks)
                        """)
                    
                        if st.checkbox("Calculate Bank DCF", key="enable_bank_dcf"):
                            try:
                                st.info("🏦 Calculating Bank FCFE Valuation (Equity DCF)...")
                            
                                # Get Ke (Cost of Equity) - NOT WACC!
                                # Calculate using CAPM
                                wacc_details_temp = calculate_wacc(financials, tax_rate, peer_tickers=comp_tickers_listed,
                                                                   beta_start_date=beta_start_date_listed,
                                                                   beta_end_date=beta_end_date_listed)
                                ke = wacc_details_temp['ke']  # Use only Ke, ignore WACC
                            
                                st.info(f"💡 Using **Ke = {ke:.2f}%** (NOT WACC - banks use equity DCF)")
                            
                                # Advanced bank parameters
                                with st.expander("🔧 Bank DCF Parameters"):
                                    col_bank1, col_bank2 = st.columns(2)
                                    with col_bank1:
                                        car_ratio = st.number_input(
                                            "Capital Adequacy Ratio (CAR) %",
                                            min_value=10.0, max_value=20.0, value=14.0, step=0.5,
                                            help="Target CAR - usually 13-15% in India"
                                        )
                                    with col_bank2:
                                        rwa_pct = st.number_input(
                                            "Risk Weight %",
                                            min_value=50.0, max_value=100.0, value=75.0, step=5.0,
                                            help="RWA as % of advances - usually 70-80%"
                                        )
                            
                                # Use bank-specific projections (FCFE)
                                projections_bank, drivers_bank = project_financials_bank(
                                    financials, 
                                    projection_years_bank, 
                                    tax_rate,
                                    car_ratio=car_ratio,
                                    rwa_percentage=rwa_pct
                                )
                            
                                if projections_bank:
                                    # Use bank FCFE valuation (NOT regular DCF!)
                                    bank_dcf_result, error_bank = calculate_bank_fcfe_valuation(
                                        projections_bank,
                                        ke,  # Cost of Equity, NOT WACC
                                        terminal_growth_bank,
                                        shares
                                    )
                                
                                    if error_bank:
                                        st.error(f"❌ Bank FCFE Error: {error_bank}")
                                    elif bank_dcf_result:
                                        st.success(f"✅ **Bank FCFE Fair Value:** {_ticker_csym}{bank_dcf_result['fair_value_per_share']:.2f}")
                                    
                                        # Display key metrics
                                        col_dcf1, col_dcf2, col_dcf3, col_dcf4 = st.columns(4)
                                        with col_dcf1:
                                            st.metric("FCFE Fair Value", f"{_ticker_csym}{bank_dcf_result['fair_value_per_share']:.2f}")
                                        with col_dcf2:
                                            st.metric("Cost of Equity (Ke)", f"{ke:.2f}%")
                                        with col_dcf3:
                                            upside = ((bank_dcf_result['fair_value_per_share'] - current_price) / current_price * 100) if current_price > 0 else 0
                                            st.metric("Upside/Downside", f"{upside:+.1f}%")
                                        with col_dcf4:
                                            st.metric("Sustainable Growth", f"{bank_dcf_result.get('sustainable_growth', 0):.2f}%")
                                    
                                        # Show projection drivers
                                        st.markdown("---")
                                        st.markdown("**📊 Bank FCFE Model Inputs:**")
                                        col_input1, col_input2 = st.columns(2)
                                        with col_input1:
                                            st.write(f"• **Revenue Growth:** {drivers_bank['revenue_growth']:.2f}%")
                                            st.write(f"• **ROE:** {drivers_bank['roe']:.2f}%")
                                            st.write(f"• **CAR Target:** {drivers_bank['car_ratio']:.1f}%")
                                        with col_input2:
                                            st.write(f"• **RWA Weight:** {drivers_bank['rwa_percentage']:.1f}%")
                                            st.write(f"• **Tax Rate:** {drivers_bank['tax_rate']:.1f}%")
                                            st.write(f"• **Terminal Growth:** {terminal_growth_bank:.2f}%")
                                    
                                        st.info("""
                                        💡 **FCFE Formula Used:**
                                    
                                        FCFE = PAT - (Growth in Advances × RWA% × CAR%)
                                    
                                        - Discounted at **Ke** (not WACC)
                                        - Values **Equity directly** (not Enterprise Value)
                                        - No debt adjustment (debt is raw material)
                                        """)
                                else:
                                    st.error("❌ Could not generate bank projections")
                                
                            except Exception as e:
                                st.error(f"❌ Bank DCF calculation failed: {str(e)}")
                                import traceback
                                with st.expander("🔍 Error Details"):
                                    st.code(traceback.format_exc())
                
                    # Display results
                    st.markdown("---")
                    st.success("✅ Bank Valuation Complete!")
                
                    # Current Price & Fair Value Display
                    col_price1, col_price2 = st.columns(2)
                    with col_price1:
                        st.metric("📊 Current Market Price", f"{_ticker_csym} {current_price:.2f}")
                
                    # Collect all fair values (initialize list first!)
                    fair_values = []
                    if ri_model:
                        fair_values.append(ri_model['value_per_share'])
                    if ddm_model:
                        fair_values.append(ddm_model['value_per_share'])
                    if pb_roe_model:
                        fair_values.append(pb_roe_model['value_per_share'])
                    if rel_val and rel_val.get('avg_fair_value') is not None:
                        fair_values.append(rel_val['avg_fair_value'])
                    if bank_dcf_result:
                        fair_values.append(bank_dcf_result['fair_value_per_share'])
                
                    avg_fair_value = np.mean(fair_values) if fair_values else 0
                
                    with col_price2:
                        st.metric("🎯 Average Fair Value", f"{_ticker_csym} {avg_fair_value:.2f}",
                                 delta=f"{((avg_fair_value - current_price) / current_price * 100):.1f}%")
                
                    # Price vs Value Gauge
                    if avg_fair_value > 0:
                        st.plotly_chart(create_price_vs_value_gauge(current_price, avg_fair_value), 
                                      use_container_width=True)
                
                    # Valuation Methods Tabs
                    tabs_list = [
                        "📊 Summary", 
                        "🏢 Residual Income", 
                        "💰 Dividend Discount",
                        "📈 P/B with ROE",
                        "🔄 Relative Valuation"
                    ]
                
                    # Add Stock Comparison tab if enabled
                    stock_comp_data = None
                    enable_stock_comparison_state = st.session_state.get('enable_stock_comparison_listed', False)
                    
                    if enable_stock_comparison_state and STOCK_COMPARISON_AVAILABLE:
                        tabs_list.append("📈 Stock vs Financials")
                
                    # Add Bank DCF tab if calculated
                    if bank_dcf_result:
                        tabs_list.append("🏦 Bank DCF")
                    
                    # Create tabs dynamically
                    tabs = st.tabs(tabs_list)
                    tab_idx = 0
                    tab1 = tabs[tab_idx]; tab_idx += 1
                    tab2 = tabs[tab_idx]; tab_idx += 1
                    tab3 = tabs[tab_idx]; tab_idx += 1
                    tab4 = tabs[tab_idx]; tab_idx += 1
                    tab5 = tabs[tab_idx]; tab_idx += 1
                    tab_stock = tabs[tab_idx] if enable_stock_comparison_state and STOCK_COMPARISON_AVAILABLE else None
                    if tab_stock: tab_idx += 1
                    tab6 = tabs[tab_idx] if bank_dcf_result else None
                
                    with tab1:
                        st.subheader("Valuation Summary - All Methods")
                    
                        # Comparison chart
                        valuations_dict = {}
                        if ri_model:
                            valuations_dict['Residual Income Model'] = ri_model
                        if ddm_model:
                            valuations_dict['Dividend Discount Model'] = ddm_model
                        if pb_roe_model:
                            valuations_dict['P/B with ROE'] = pb_roe_model
                        if rel_val and rel_val.get('avg_fair_value') is not None:
                            valuations_dict['Relative Valuation'] = {'value_per_share': rel_val['avg_fair_value']}
                    
                        if valuations_dict:
                            st.plotly_chart(create_bank_valuation_comparison_chart(valuations_dict), 
                                          use_container_width=True)
                    
                        # Summary table
                        summary_data = []
                        if ri_model:
                            summary_data.append(['Residual Income Model', f"{_ticker_csym}{ri_model['value_per_share']:.2f}", 
                                               f"{((ri_model['value_per_share'] - current_price) / current_price * 100):.1f}%"])
                        if ddm_model:
                            summary_data.append(['Dividend Discount Model', f"{_ticker_csym}{ddm_model['value_per_share']:.2f}",
                                               f"{((ddm_model['value_per_share'] - current_price) / current_price * 100):.1f}%"])
                        if pb_roe_model:
                            summary_data.append(['P/B with ROE Analysis', f"{_ticker_csym}{pb_roe_model['value_per_share']:.2f}",
                                               f"{((pb_roe_model['value_per_share'] - current_price) / current_price * 100):.1f}%"])
                        if rel_val and rel_val.get('avg_fair_value') is not None:
                            summary_data.append(['Relative Valuation (Avg)', f"{_ticker_csym}{rel_val['avg_fair_value']:.2f}",
                                               f"{((rel_val['avg_fair_value'] - current_price) / current_price * 100):.1f}%"])
                    
                        summary_df = pd.DataFrame(summary_data, columns=['Method', 'Fair Value', 'Upside/Downside'])
                        st.dataframe(summary_df, use_container_width=True, hide_index=True)
                    
                        # Historical charts
                        st.markdown("---")
                        st.subheader("📈 Historical Financial Analysis")
                        st.plotly_chart(create_historical_financials_chart(financials), use_container_width=True)
                
                    with tab2:
                        if ri_model:
                            st.subheader("Residual Income Model")
                            st.write(f"**Fair Value per Share:** {_ticker_csym}{ri_model['value_per_share']:.2f}")
                        
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.metric("Current Book Value", f"{_ticker_csym}{ri_model['current_book_value']:,.0f}")
                            with col2:
                                st.metric("ROE", f"{ri_model['roe']:.2f}%")
                            with col3:
                                st.metric("BV Growth Rate", f"{ri_model.get('bv_growth', 10):.1f}%",
                                         help="Historical book value growth rate")
                        
                            # Growth assumptions
                            st.markdown("---")
                            st.subheader("Growth Assumptions")
                            col4, col5 = st.columns(2)
                            with col4:
                                st.info(f"**Projection Phase Growth:** {ri_model.get('bv_growth', 10):.1f}%")
                                if ri_model.get('historical_bv_growth'):
                                    hist_growth = ri_model['historical_bv_growth']
                                    st.caption(f"Based on historical growth: {', '.join([f'{g:.1f}%' for g in hist_growth])}")
                                else:
                                    st.caption("Using default 10% growth")
                        
                            with col5:
                                st.info(f"**Terminal Growth:** {ri_model.get('terminal_growth', 4):.1f}%")
                                st.caption("Assumes gradual decline to sustainable long-term rate")
                        
                            # Projections table
                            st.markdown("---")
                            st.subheader("5-Year Residual Income Projections")
                            proj_df = pd.DataFrame(ri_model['projections'])
                            proj_df['book_value'] = proj_df['book_value'].apply(lambda x: f"{_ticker_csym}{x:,.0f}")
                            proj_df['net_income'] = proj_df['net_income'].apply(lambda x: f"{_ticker_csym}{x:,.0f}")
                            proj_df['residual_income'] = proj_df['residual_income'].apply(lambda x: f"{_ticker_csym}{x:,.0f}")
                            proj_df['pv_ri'] = proj_df['pv_ri'].apply(lambda x: f"{_ticker_csym}{x:,.0f}")
                            st.dataframe(proj_df, use_container_width=True)
                        
                            # Value breakdown
                            st.markdown("---")
                            st.subheader("Valuation Breakdown")
                            breakdown_df = pd.DataFrame({
                                'Component': ['Current Book Value', 'PV of 5Y Residual Income', 'PV of Terminal Value', 'Total Equity Value'],
                                'Value ({_ticker_csym})': [
                                    f"{_ticker_csym}{ri_model['current_book_value']:,.0f}",
                                    f"{_ticker_csym}{ri_model['sum_pv_ri']:,.0f}",
                                    f"{_ticker_csym}{ri_model['terminal_ri_pv']:,.0f}",
                                    f"{_ticker_csym}{ri_model['total_equity_value']:,.0f}"
                                ]
                            })
                            st.table(breakdown_df)
                        
                            # Formula explanation
                            with st.expander("📖 Residual Income Formula"):
                                st.latex(r"RI = NI - (K_e \times BV)")
                                st.latex(r"Value = BV_0 + \sum_{t=1}^{n} \frac{RI_t}{(1+K_e)^t} + \frac{TV}{(1+K_e)^n}")
                                st.write("Where:")
                                st.write("- RI = Residual Income")
                                st.write("- NI = Net Income")
                                st.write(f"- Kₑ = Cost of Equity = {cost_of_equity:.2f}%")
                                st.write("- BV = Book Value of Equity")
                                st.write("- TV = Terminal Value")
                        else:
                            st.warning("Residual Income Model calculation failed")
                
                    with tab3:
                        if ddm_model:
                            st.subheader("Dividend Discount Model")
                        
                            # Show data source
                            if ddm_model.get('using_actual_data'):
                                st.success("✅ Using actual historical dividend data")
                            else:
                                st.info("ℹ️ Using estimated dividend data (no historical dividends found)")
                        
                            st.write(f"**Fair Value per Share:** {_ticker_csym}{ddm_model['value_per_share']:.2f}")
                        
                            col1, col2 = st.columns(2)
                            with col1:
                                st.metric("Current DPS", f"{_ticker_csym}{ddm_model['current_dps']:.2f}")
                                st.metric("Payout Ratio", f"{ddm_model['payout_ratio']:.1f}%")
                        
                            with col2:
                                st.metric("Dividend Growth", f"{ddm_model['dividend_growth']:.1f}%")
                                st.metric("Next Year DPS (D1)", f"{_ticker_csym}{ddm_model['next_year_dps']:.2f}")
                        
                            # Historical dividends if available
                            if ddm_model.get('historical_dividends'):
                                st.markdown("---")
                                st.subheader("Historical Dividends (Annual)")
                                hist_divs = ddm_model['historical_dividends']
                                years_range = list(range(len(hist_divs), 0, -1))
                                hist_df = pd.DataFrame({
                                    'Year': [f"T-{y}" for y in years_range],
                                    'Dividend ({_ticker_csym})': hist_divs
                                })
                                st.dataframe(hist_df, use_container_width=True, hide_index=True)
                        
                            st.markdown("---")
                            st.subheader("5-Year Dividend Projection")
                        
                            # Dividend projections
                            div_df = pd.DataFrame(ddm_model['projections'])
                            div_df['dividend'] = div_df['dividend'].apply(lambda x: f"{_ticker_csym}{x:.2f}")
                            div_df['pv_dividend'] = div_df['pv_dividend'].apply(lambda x: f"{_ticker_csym}{x:.2f}")
                            st.dataframe(div_df, use_container_width=True)
                        
                            # DDM formula explanation
                            with st.expander("📖 DDM Formula & Assumptions"):
                                st.latex(r"Value = \frac{D_1}{K_e - g}")
                                st.write(f"Where:")
                                st.write(f"- D₁ = Next year dividend = {_ticker_csym}{ddm_model['next_year_dps']:.2f}")
                                st.write(f"- Kₑ = Cost of Equity = {cost_of_equity:.2f}%")
                                st.write(f"- g = Dividend Growth Rate = {ddm_model['dividend_growth']:.2f}%")
                        else:
                            st.warning("DDM calculation failed or not applicable (cost of equity ≤ growth rate)")
                
                    with tab4:
                        if pb_roe_model:
                            st.subheader("P/B with ROE Analysis")
                            st.write(f"**Fair Value per Share:** {_ticker_csym}{pb_roe_model['value_per_share']:.2f}")
                            st.write(f"**Book Value per Share:** {_ticker_csym}{pb_roe_model['book_value_per_share']:.2f}")
                            st.write(f"**ROE:** {pb_roe_model['roe']:.2f}%")
                            st.write(f"**Cost of Equity:** {pb_roe_model['cost_of_equity']:.2f}%")
                            st.write(f"**Fair P/B Ratio:** {pb_roe_model['fair_pb_ratio']:.2f}x")
                        
                            # Historical ROE chart
                            roe_years = financials['years']
                            fig = go.Figure()
                            fig.add_trace(go.Scatter(x=roe_years, y=pb_roe_model['historical_roe'],
                                                    mode='lines+markers', name='ROE',
                                                    line=dict(color='#06A77D', width=3)))
                            fig.add_hline(y=pb_roe_model['cost_of_equity'], line_dash="dash",
                                         annotation_text=f"Cost of Equity: {pb_roe_model['cost_of_equity']:.2f}%")
                            fig.update_layout(title="Historical ROE Trend", xaxis_title="Year",
                                            yaxis_title="ROE %", height=400)
                            st.plotly_chart(fig, use_container_width=True)
                        else:
                            st.warning("P/B ROE model calculation failed")
                
                    with tab5:
                        if rel_val:
                            st.subheader("Relative Valuation (Peer-Based)")

                            def _fmt_x(v):
                                return f"{v:.2f}x" if v is not None else "N/A"
                            def _fmt_cur(v):
                                return f"{_ticker_csym}{v:.2f}" if v is not None else "N/A"
                        
                            # Show peer data summary
                            if rel_val.get('peer_count', 0) > 0:
                                st.success(f"✅ Analyzed {rel_val['peer_count']} peer companies")
                            else:
                                st.warning("⚠️ No usable peer multiples were fetched — P/E and/or P/B based fair values below may show N/A rather than a fabricated estimate.")
                        
                            # Main metrics
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.metric("Current P/E", _fmt_x(rel_val.get('current_pe')))
                                st.metric("Peer Avg P/E", _fmt_x(rel_val.get('sector_avg_pe')))
                                st.caption(f"Range: {_fmt_x(rel_val.get('sector_low_pe'))} - {_fmt_x(rel_val.get('sector_high_pe'))}")
                        
                            with col2:
                                st.metric("Current P/B", _fmt_x(rel_val.get('current_pb')))
                                st.metric("Peer Avg P/B", _fmt_x(rel_val.get('sector_avg_pb')))
                                st.caption(f"Range: {_fmt_x(rel_val.get('sector_low_pb'))} - {_fmt_x(rel_val.get('sector_high_pb'))}")
                        
                            with col3:
                                st.metric("Fair Value (P/E)", _fmt_cur(rel_val.get('fair_value_pe_based')))
                                st.metric("Fair Value (P/B)", _fmt_cur(rel_val.get('fair_value_pb_based')))
                        
                            st.markdown("---")
                        
                            # Valuation ranges
                            st.subheader("Valuation Range Analysis")
                            col4, col5, col6 = st.columns(3)
                            with col4:
                                st.metric("Conservative", _fmt_cur(rel_val.get('conservative_value')),
                                         help="Based on 25th percentile peer P/E")
                            with col5:
                                st.metric("Fair Value", _fmt_cur(rel_val.get('avg_fair_value')),
                                         help="Average of P/E and P/B based valuations")
                            with col6:
                                st.metric("Aggressive", _fmt_cur(rel_val.get('aggressive_value')),
                                         help="Based on 75th percentile peer P/E")
                        
                            # Peer comparison table
                            if rel_val.get('peer_data') and len(rel_val['peer_data']) > 0:
                                st.markdown("---")
                                st.subheader("Peer Comparison")
                                peer_df = pd.DataFrame(rel_val['peer_data'])
                                peer_df.columns = ['Ticker', f'Price ({_ticker_csym})', 'P/E', 'P/B']
                                st.dataframe(peer_df, use_container_width=True, hide_index=True)
                        
                            # Interpretation — only when we have real peer averages to compare against
                            if rel_val.get('sector_avg_pe') or rel_val.get('sector_avg_pb'):
                                st.markdown("---")
                                st.subheader("Interpretation")
                                _sap = rel_val.get('sector_avg_pe')
                                _spb = rel_val.get('sector_avg_pb')
                                pe_premium = ((rel_val['current_pe'] - _sap) / _sap * 100) if (_sap and rel_val.get('current_pe')) else None
                                pb_premium = ((rel_val['current_pb'] - _spb) / _spb * 100) if (_spb and rel_val.get('current_pb')) else None
                        
                                if pe_premium is not None:
                                    if pe_premium > 20:
                                        st.warning(f"📊 **P/E Analysis:** Trading at {pe_premium:.1f}% premium to peers. May be overvalued unless justified by superior growth.")
                                    elif pe_premium < -20:
                                        st.success(f"📊 **P/E Analysis:** Trading at {abs(pe_premium):.1f}% discount to peers. Potential undervaluation.")
                                    else:
                                        st.info(f"📊 **P/E Analysis:** Trading in line with peers ({pe_premium:+.1f}% premium).")
                        
                                if pb_premium is not None:
                                    if pb_premium > 20:
                                        st.warning(f"📈 **P/B Analysis:** Trading at {pb_premium:.1f}% premium to peers. May indicate high growth expectations.")
                                    elif pb_premium < -20:
                                        st.success(f"📈 **P/B Analysis:** Trading at {abs(pb_premium):.1f}% discount to peers. Potential value opportunity.")
                                    else:
                                        st.info(f"📈 **P/B Analysis:** Trading in line with peers ({pb_premium:+.1f}% premium).")
                        else:
                            st.warning("Relative valuation calculation failed")
                
                    # Stock Price Comparison Tab (if enabled)
                    if tab_stock and enable_stock_comparison_state and STOCK_COMPARISON_AVAILABLE:
                        with tab_stock:
                            st.subheader("📈 Stock Price vs Revenue & EPS Analysis")
                            
                            with st.spinner("Fetching stock price data..."):
                                try:
                                    # Determine years to fetch (max 4 for Yahoo Finance)
                                    years_to_fetch = min(historical_years_listed, 4)
                                    
                                    # Fetch comparison data
                                    stock_comp_data = get_stock_comparison_data_listed(
                                        ticker=full_ticker,
                                        company_name=company_name,
                                        financials=financials,
                                        num_years=years_to_fetch
                                    )
                                    
                                    if stock_comp_data and stock_comp_data['chart_fig']:
                                        st.plotly_chart(stock_comp_data['chart_fig'], use_container_width=True)
                                        
                                        # Show data tables in expanders
                                        with st.expander("📊 View Raw Data"):
                                            col1, col2, col3 = st.columns(3)
                                            
                                            with col1:
                                                st.markdown("**Revenue Data**")
                                                if stock_comp_data['revenue_df'] is not None:
                                                    st.dataframe(stock_comp_data['revenue_df'], hide_index=True)
                                                else:
                                                    st.info("No revenue data available")
                                            
                                            with col2:
                                                st.markdown("**EPS Data**")
                                                if stock_comp_data['eps_df'] is not None:
                                                    st.dataframe(stock_comp_data['eps_df'], hide_index=True)
                                                else:
                                                    st.info("No EPS data available")
                                            
                                            with col3:
                                                st.markdown("**Stock Price Summary**")
                                                if stock_comp_data['stock_prices_df'] is not None:
                                                    price_df = stock_comp_data['stock_prices_df']
                                                    st.metric("Latest Price", f"{_ticker_csym}{price_df['Close'].iloc[-1]:.2f}")
                                                    st.metric("Period Return", f"{((price_df['Close'].iloc[-1] - price_df['Close'].iloc[0]) / price_df['Close'].iloc[0] * 100):.2f}%")
                                                    st.metric("Major Changes", f"{price_df['is_major'].sum()}" if 'is_major' in price_df.columns else "N/A")
                                                else:
                                                    st.info("No stock price data available")
                                        
                                        st.info(f"💡 Chart shows {years_to_fetch} years of data (Yahoo Finance limit: 4 years max)")
                                    else:
                                        st.warning("Could not generate stock comparison chart. Check if ticker and financial data are available.")
                                
                                except Exception as e:
                                    st.error(f"Error generating stock comparison: {str(e)}")
                
                    # Bank DCF Tab (if calculated)
                    if bank_dcf_result:
                        with tab6:
                            st.subheader("🏦 Bank FCFE Valuation (Equity DCF)")
                            st.success("✅ Using proper FCFE methodology - NOT FCFF!")
                        
                            # Fair value display
                            col_dcf1, col_dcf2, col_dcf3, col_dcf4 = st.columns(4)
                            with col_dcf1:
                                st.metric("Fair Value/Share", f"{_ticker_csym}{bank_dcf_result['fair_value_per_share']:.2f}",
                                        delta=f"{((bank_dcf_result['fair_value_per_share'] - current_price) / current_price * 100):.1f}%" if current_price > 0 else None)
                            with col_dcf2:
                                st.metric("Equity Value", f"{_ticker_csym}{bank_dcf_result['equity_value']:,.0f} Lacs")
                            with col_dcf3:
                                st.metric("Cost of Equity (Ke)", f"{bank_dcf_result['cost_of_equity']:.2f}%")
                            with col_dcf4:
                                st.metric("Terminal Growth", f"{bank_dcf_result['terminal_growth']:.2f}%")
                        
                            st.markdown("---")
                        
                            # Show projections
                            st.markdown("**📈 Bank FCFE Projection Details**")
                            col_proj1, col_proj2 = st.columns(2)
                        
                            with col_proj1:
                                st.markdown("**Model Inputs:**")
                                st.write(f"• Revenue Growth: {drivers_bank['revenue_growth']:.2f}%")
                                st.write(f"• ROE: {drivers_bank['roe']:.2f}%")
                                st.write(f"• CAR Target: {drivers_bank['car_ratio']:.1f}%")
                                st.write(f"• RWA Weight: {drivers_bank['rwa_percentage']:.1f}%")
                        
                            with col_proj2:
                                st.markdown("**Valuation Components:**")
                                st.write(f"• PV of FCFE (5Y): {_ticker_csym}{bank_dcf_result['sum_pv_fcfe']:,.0f} Lacs")
                                st.write(f"• PV of Terminal: {_ticker_csym}{bank_dcf_result['terminal_value_pv']:,.0f} Lacs")
                                st.write(f"• Sustainable Growth: {bank_dcf_result['sustainable_growth']:.2f}%")
                        
                            st.markdown("---")
                        
                            # Show FCFE projections
                            if 'fcfe' in projections_bank:
                                st.markdown("**💰 FCFE Projections (Free Cash Flow to Equity)**")
                            
                                # Create detailed projection table
                                proj_data = {
                                    'Year': projections_bank['year'],
                                    'Revenue': [f"{_ticker_csym}{rev:,.0f}" for rev in projections_bank['revenue']],
                                    'PAT': [f"{_ticker_csym}{pat:,.0f}" for pat in projections_bank['pat']],
                                    'FCFE': [f"{_ticker_csym}{fcfe:,.0f}" for fcfe in projections_bank['fcfe']],
                                }
                            
                                # Add PV of FCFE
                                if 'pv_fcfe_by_year' in bank_dcf_result:
                                    proj_data['PV of FCFE'] = [f"{_ticker_csym}{pv:,.0f}" for pv in bank_dcf_result['pv_fcfe_by_year']]
                            
                                proj_df = pd.DataFrame(proj_data)
                                st.dataframe(proj_df, use_container_width=True, hide_index=True)
                            
                                st.caption("💡 FCFE = PAT - Equity Required for Growth")
                        
                            st.markdown("---")
                            st.markdown("**💡 FCFE Methodology for Banks:**")
                            st.info("""
                            **Why FCFE (not FCFF):**
                            - For banks, debt is RAW MATERIAL (not financing)
                            - Interest expense is OPERATING COST (not financing cost)
                            - Cannot separate enterprise value from equity value
                            - Must value equity directly using FCFE
                        
                            **FCFE Formula:**
                            ```
                            FCFE = PAT - Equity Required for Growth
                        
                            Where:
                            Equity Required = Growth in Advances × RWA% × CAR%
                            ```
                        
                            **Discounting:**
                            - Uses Ke (Cost of Equity) - NOT WACC
                            - Values equity directly - NOT enterprise value
                            - No debt adjustment needed
                        
                            **Capital Adequacy:**
                            - Banks need equity capital to support loan growth
                            - Retained earnings fund asset expansion
                            - Only excess cash after capital requirements = FCFE
                            """)
                
                    st.stop()
            
                # ================================
                # STANDARD FCFF DCF FOR NON-BANK COMPANIES
                # ================================
                else:
                    st.markdown("---")
                
                # Calculate WC metrics
                wc_metrics = calculate_working_capital_metrics(financials)
                
                # Show Working Capital Data Status
                wc_status_parts = []
                if wc_metrics.get('has_inventory', False) or (inventory_days_override and inventory_days_override > 0):
                    inv_source = f"User Override: {inventory_days_override} days" if inventory_days_override and inventory_days_override > 0 else f"Historical: {wc_metrics['avg_inv_days']:.1f} days"
                    wc_status_parts.append(f"✅ Inventory ({inv_source})")
                else:
                    wc_status_parts.append("⚠️ Inventory (No data)")
                
                if wc_metrics.get('has_receivables', False) or (debtor_days_override and debtor_days_override > 0):
                    deb_source = f"User Override: {debtor_days_override} days" if debtor_days_override and debtor_days_override > 0 else f"Historical: {wc_metrics['avg_deb_days']:.1f} days"
                    wc_status_parts.append(f"✅ Debtors ({deb_source})")
                else:
                    wc_status_parts.append("⚠️ Debtors (No data)")
                
                if wc_metrics.get('has_payables', False) or (creditor_days_override and creditor_days_override > 0):
                    cred_source = f"User Override: {creditor_days_override} days" if creditor_days_override and creditor_days_override > 0 else f"Historical: {wc_metrics['avg_cred_days']:.1f} days"
                    wc_status_parts.append(f"✅ Creditors ({cred_source})")
                else:
                    wc_status_parts.append("⚠️ Creditors (No data)")
                
                st.info(f"🔍 **Working Capital Projection Status:** {' | '.join(wc_status_parts)}")
                
                # Project financials with ALL user overrides
                projections, drivers = project_financials(
                    financials, wc_metrics, projection_years_listed, tax_rate,
                    rev_growth_override_listed, opex_margin_override_listed, capex_ratio_override_listed,
                    # Per-year controls
                    rev_growth_per_year=rev_growth_per_year_listed if (_ms_enabled_listed or any(v > 0 for v in rev_growth_per_year_listed)) else None,
                    ebitda_margin_per_year=ebitda_margin_per_year_listed if (bool(_ms_margin_listed) or any(v > 0 for v in ebitda_margin_per_year_listed)) else None,
                    # Pass all advanced user controls
                    ebitda_margin_override=ebitda_margin_override if ebitda_margin_override > 0 else None,
                    depreciation_rate_override=depreciation_rate_override if depreciation_rate_override > 0 else None,
                    depreciation_method=depreciation_method,
                    inventory_days_override=inventory_days_override if inventory_days_override > 0 else None,
                    debtor_days_override=debtor_days_override if debtor_days_override > 0 else None,
                    creditor_days_override=creditor_days_override if creditor_days_override > 0 else None,
                    interest_rate_override=interest_rate_override if interest_rate_override > 0 else None,
                    working_capital_pct_override=working_capital_as_pct_revenue if working_capital_as_pct_revenue > 0 else None
                )
                
                # Calculate WACC (beta of the company itself)
                st.info("Calculating beta for the stock...")
                # Pass ticker WITH exchange suffix for proper beta calculation
                full_ticker = get_ticker_with_exchange(ticker, exchange_suffix)
                beta = get_stock_beta(full_ticker, period_years=3,
                                      beta_start_date=beta_start_date_listed,
                                      beta_end_date=beta_end_date_listed)
                st.success(f"✅ Beta calculated: {beta:.3f}")
                
                st.info(f"🏛️ Risk-Free Rate (India 10Y G-Sec): {manual_rf_rate:.2f}%")
                wacc_details = calculate_wacc(financials, tax_rate, peer_tickers=None,
                                              manual_rf_rate=manual_rf_rate, manual_rm_rate=manual_rm_rate,
                                              beta_start_date=beta_start_date_listed,
                                              beta_end_date=beta_end_date_listed)
                wacc_details['beta'] = beta  # Override with actual stock beta
                # Recalculate Ke and WACC with actual beta
                wacc_details['ke'] = wacc_details['rf'] + (beta * (wacc_details['rm'] - wacc_details['rf']))
                wacc_details['wacc'] = (wacc_details['we']/100 * wacc_details['ke']) + (wacc_details['wd']/100 * wacc_details['kd_after_tax'])
                
                # DCF Valuation
                # Extract cash balance
                cash_balance = financials['cash'][0] if financials['cash'][0] > 0 else 0
                
                valuation, error = calculate_dcf_valuation(
                    projections, wacc_details, terminal_growth, shares, cash_balance,
                    manual_discount_rate=manual_discount_rate if manual_discount_rate > 0 else None
                )
                
                if error:
                    st.error(error)
                    st.stop()
                
                # ================================
                # GET CURRENT PRICE EARLY
                # ================================
                current_price = 0
                try:
                    stock = get_cached_ticker(get_ticker_with_exchange(ticker, exchange_suffix))
                    info = stock.info if stock else None
                    if not is_valuation_date_today(valuation_date_listed):
                        # Valuation-date mode: price the company as of that date, not today
                        _hist_p = get_price_as_of(stock, valuation_date_listed)
                        if _hist_p:
                            current_price = _hist_p
                        else:
                            st.warning(f"⚠️ No {ticker} price found on/before {valuation_date_listed.strftime('%d-%b-%Y')} — using live price instead.")
                    if not current_price or current_price == 0:
                        # Try multiple methods to get the live/current price (Phase 1 approach)
                        if info:
                            current_price = info.get('currentPrice', 0)
                            if not current_price or current_price == 0:
                                current_price = info.get('regularMarketPrice', 0)
                        if not current_price or current_price == 0:
                            # Try getting from recent history
                            try:
                                hist = stock.history(period='1d')
                                if not hist.empty:
                                    current_price = hist['Close'].iloc[-1]
                            except:
                                pass
                except Exception as e:
                    st.warning(f"Could not fetch current price: {e}")
                    current_price = 0

                # Derive currency symbol from Yahoo Finance info (e.g. $ for GEV/TSLA, ₹ for Indian stocks)
                # 'info' may have been set in the try block above; fall back gracefully
                try:
                    _ticker_csym = get_currency_symbol(info)
                except Exception:
                    _ticker_csym = get_currency_symbol(yahoo_data.get('info') if 'yahoo_data' in dir() else None)

                # ================================
                # ADDITIONAL VALUATION MODELS FOR NON-BANKING COMPANIES
                # ================================
                
                # Calculate Dividend Discount Model (DDM)
                st.info("Calculating Dividend Discount Model...")
                ddm_result = calculate_dividend_discount_model(
                    financials, shares, wacc_details['ke'], 
                    ticker=ticker,
                    div_growth_override=ddm_dividend_growth_override if ddm_dividend_growth_override > 0 else None,
                    payout_ratio_override=ddm_payout_ratio_override if ddm_payout_ratio_override > 0 else None,
                    dcf_projections=projections,  # ✅ PASS EXISTING PROJECTIONS - NO DUPLICATION!
                    exchange_suffix=exchange_suffix
                )
                if ddm_result and isinstance(ddm_result, dict) and 'value_per_share' in ddm_result:
                    st.success(f"✅ DDM Fair Value: {_ticker_csym}{ddm_result['value_per_share']:.2f}")
                    if ddm_result.get('using_dcf_projections'):
                        st.caption("💡 Using DCF projected NOPAT for dividend projections")
                else:
                    st.warning("⚠️ DDM not applicable (company may not pay dividends)")
                
                # Calculate Residual Income Model (RIM)
                st.info("Calculating Residual Income Model...")
                # Use user overrides or defaults
                rim_terminal_g = rim_terminal_growth_override if rim_terminal_growth_override > 0 else terminal_growth
                rim_proj_years = rim_projection_years_override if rim_projection_years_override > 0 else projection_years_listed
                rim_roe = rim_assumed_roe_override if rim_assumed_roe_override > 0 else None
                
                rim_result = calculate_residual_income_model(
                    financials, shares, wacc_details['ke'], 
                    terminal_growth=rim_terminal_g, 
                    projection_years=rim_proj_years,
                    assumed_roe=rim_roe,
                    dcf_projections=projections  # ✅ PASS EXISTING PROJECTIONS - NO DUPLICATION!
                )
                if rim_result and isinstance(rim_result, dict) and 'value_per_share' in rim_result:
                    st.success(f"✅ RIM Fair Value: {_ticker_csym}{rim_result['value_per_share']:.2f}")
                    if rim_result.get('using_dcf_projections'):
                        st.caption("💡 Using DCF projected NOPAT as Net Income")
                else:
                    st.warning("⚠️ RIM calculation returned incomplete results")
                    if rim_result and isinstance(rim_result, dict):
                        st.caption(f"RIM result keys: {list(rim_result.keys())}")
                
                # ================================
                # DISPLAY RESULTS (SAME AS UNLISTED)
                # ================================
                
                st.success("✅ Valuation Complete!")
                
                # Calculate comparative valuation EARLY (needed for Forward P/E
                # display).
                comp_results = None
                if comp_tickers_listed and comp_tickers_listed.strip():
                    try:
                        # Get use_screener_peers flag from session state
                        use_screener_for_peers = st.session_state.get('use_screener_peers', False)
                    
                        comp_results = perform_comparative_valuation(
                            ticker, 
                            comp_tickers_listed, 
                            financials, 
                            shares, 
                            exchange_suffix, 
                            projections=projections,
                            use_screener_peers=use_screener_for_peers,
                            valuation_date=valuation_date_listed
                        )
                    except Exception as e:
                        st.warning(f"Could not calculate comparative valuation: {str(e)}")
                
                # Key Metrics with Current Price and P/E
                # Get current P/E if available
                current_pe = info.get('trailingPE', 0) if 'info' in locals() else 0
                current_eps = (financials['nopat'][0] * 100000) / shares if shares > 0 and financials['nopat'][0] > 0 else 0
                
                st.markdown("### 📊 Key Valuation Metrics")
                
                col1, col2, col3, col4, col5, col6 = st.columns(6)
                with col1:
                    st.metric("📊 Current Price", f"{_ticker_csym} {current_price:.2f}")
                with col2:
                    st.metric("🎯 Fair Value (DCF)", f"{_ticker_csym} {valuation['fair_value_per_share']:.2f}",
                             delta=f"{((valuation['fair_value_per_share'] - current_price) / current_price * 100):.1f}%")
                with col3:
                    st.metric("Current P/E", f"{current_pe:.2f}x" if current_pe > 0 else "N/A")
                with col4:
                    st.metric("Current EPS", f"{_ticker_csym} {current_eps:.2f}" if current_eps > 0 else "N/A")
                with col5:
                    st.metric("WACC", f"{wacc_details['wacc']:.2f}%")
                with col6:
                    st.metric("Terminal Growth", f"{terminal_growth:.1f}%")
                
                # Forward P/E Display (if available)
                if 'comp_results' in locals() and comp_results and 'forward_pe' in comp_results:
                    st.markdown("---")
                    st.markdown("### 📅 12-Month Forward P/E Valuation")
                
                    fpe = comp_results['forward_pe']
                
                    col_fpe1, col_fpe2, col_fpe3, col_fpe4, col_fpe5 = st.columns(5)
                
                    with col_fpe1:
                        st.metric("Current EPS", f"{_ticker_csym}{fpe.get('current_eps', 0):.2f}")
                
                    with col_fpe2:
                        st.metric("Forward EPS (12M)", f"{_ticker_csym}{fpe['forward_eps']:.2f}",
                                delta=f"+{fpe.get('earnings_growth_rate', 0):.1f}%",
                                help="Projected EPS for next 12 months")
                
                    with col_fpe3:
                        st.metric("Peer Avg P/E", f"{comp_results['multiples_stats']['pe']['average']:.2f}x" if 'pe' in comp_results['multiples_stats'] else "N/A")
                
                    with col_fpe4:
                        st.metric("Forward Fair Value (Avg)", f"{_ticker_csym}{fpe['fair_value_avg']:.2f}",
                                delta=f"{((fpe['fair_value_avg'] - current_price) / current_price * 100):.1f}%" if current_price > 0 else None)
                
                    with col_fpe5:
                        st.metric("Forward Fair Value (Median)", f"{_ticker_csym}{fpe['fair_value_median']:.2f}",
                                delta=f"{((fpe['fair_value_median'] - current_price) / current_price * 100):.1f}%" if current_price > 0 else None)
                
                    st.caption(f"💡 {fpe.get('calculation_note', 'Forward EPS projected using historical growth')}")
                
                st.markdown("---")
                
                # Price vs Value Gauge
                if valuation['fair_value_per_share'] > 0:
                    st.plotly_chart(create_price_vs_value_gauge(current_price, valuation['fair_value_per_share']), 
                                  use_container_width=True)
                
                # COMPARISON BAR CHART - All Valuation Methods vs Current Price
                with st.expander("📊 **All Valuation Methods Comparison**", expanded=False):
                    st.markdown("### Fair Value Comparison Across All Methods")
                
                    # Collect all available fair values
                    valuation_methods = []
                    fair_values = []
                    colors = []
                
                    # DCF
                    if valuation and valuation.get('fair_value_per_share', 0) > 0:
                        valuation_methods.append('DCF (FCFF)')
                        fair_values.append(valuation['fair_value_per_share'])
                        colors.append('#1f77b4')  # Blue
                
                    # DDM
                    if ddm_result and ddm_result.get('value_per_share', 0) > 0 and not ddm_result.get('error'):
                        valuation_methods.append('DDM (Gordon)')
                        fair_values.append(ddm_result['value_per_share'])
                        colors.append('#2ca02c')  # Green
                
                    # RIM
                    if rim_result and rim_result.get('value_per_share', 0) > 0 and not rim_result.get('error'):
                        valuation_methods.append('RIM')
                        fair_values.append(rim_result['value_per_share'])
                        colors.append('#ff7f0e')  # Orange
                
                    # Forward P/E
                    if comp_results and 'forward_pe' in comp_results:
                        fpe = comp_results['forward_pe']
                        if fpe.get('fair_value_avg', 0) > 0:
                            valuation_methods.append('Forward P/E (Avg)')
                            fair_values.append(fpe['fair_value_avg'])
                            colors.append('#d62728')  # Red
                    
                        if fpe.get('fair_value_median', 0) > 0:
                            valuation_methods.append('Forward P/E (Median)')
                            fair_values.append(fpe['fair_value_median'])
                            colors.append('#9467bd')  # Purple
                
                    # Comparative Valuation methods
                    if comp_results and 'valuations' in comp_results:
                        for method, val_data in comp_results['valuations'].items():
                            if val_data.get('fair_value_avg', 0) > 0:
                                method_name = val_data.get('method', method).replace('_', ' ').title()
                                valuation_methods.append(f"{method_name} (Avg)")
                                fair_values.append(val_data['fair_value_avg'])
                                colors.append('#8c564b')  # Brown
                
                    # Create comparison chart if we have data
                    if valuation_methods and current_price > 0 and len(fair_values) > 0:
                    
                        fig = go.Figure()
                    
                        # Add bars for each valuation method
                        fig.add_trace(go.Bar(
                            y=valuation_methods,
                            x=fair_values,
                            marker_color=colors,
                            text=[f"{_ticker_csym}{v:.2f}" for v in fair_values],
                            textposition='outside',
                            orientation='h',
                            name='Fair Value',
                            hovertemplate='<b>%{y}</b><br>Fair Value: {_ticker_csym}%{x:.2f}<extra></extra>'
                        ))
                    
                        # Add red line for current price
                        fig.add_vline(
                            x=current_price,
                            line_dash="dash",
                            line_color="red",
                            line_width=3,
                            annotation_text=f"Current Price: {_ticker_csym}{current_price:.2f}",
                            annotation_position="top right"
                        )
                    
                        # Add average fair value line
                        avg_fair_value = np.mean(fair_values)
                        fig.add_vline(
                            x=avg_fair_value,
                            line_dash="dot",
                            line_color="green",
                            line_width=2,
                            annotation_text=f"Average: {_ticker_csym}{avg_fair_value:.2f}",
                            annotation_position="bottom right"
                        )
                    
                        fig.update_layout(
                            title={
                                'text': "Fair Value Comparison: All Methods vs Current Price",
                                'x': 0.5,
                                'xanchor': 'center'
                            },
                            xaxis_title="Fair Value ({_ticker_csym})",
                            yaxis_title="Valuation Method",
                            height=max(400, len(valuation_methods) * 60),
                            showlegend=False,
                            xaxis=dict(gridcolor='lightgray', zeroline=True),
                            yaxis=dict(autorange="reversed"),  # Top to bottom
                            plot_bgcolor='rgba(240,240,240,0.5)'
                        )
                    
                        st.plotly_chart(fig, use_container_width=True)
                    
                        # Summary statistics
                        st.markdown("### 📊 Summary Statistics")
                        col_stat1, col_stat2, col_stat3, col_stat4 = st.columns(4)
                    
                        with col_stat1:
                            st.metric("Average Fair Value", f"{_ticker_csym}{avg_fair_value:.2f}",
                                    delta=f"{((avg_fair_value - current_price) / current_price * 100):.1f}%" if current_price > 0 else None)
                    
                        with col_stat2:
                            median_fv = np.median(fair_values)
                            st.metric("Median Fair Value", f"{_ticker_csym}{median_fv:.2f}",
                                    delta=f"{((median_fv - current_price) / current_price * 100):.1f}%" if current_price > 0 else None)
                    
                        with col_stat3:
                            max_fv = max(fair_values)
                            min_fv = min(fair_values)
                            st.metric("Range", f"{_ticker_csym}{min_fv:.2f} - {_ticker_csym}{max_fv:.2f}")
                    
                        with col_stat4:
                            consensus = "🔴 Overvalued" if current_price > avg_fair_value else "🟢 Undervalued"
                            deviation = abs((avg_fair_value - current_price) / current_price * 100)
                            st.metric("Consensus", consensus)
                            st.caption(f"Deviation: {deviation:.1f}%")
                    
                        # Methods breakdown
                        st.markdown("---")
                        st.markdown("**💡 Interpretation:**")
                        if current_price < min_fv:
                            st.success(f"✅ **Strong Buy Signal**: Current price ({_ticker_csym}{current_price:.2f}) is below ALL valuation methods. Significant upside potential.")
                        elif current_price > max_fv:
                            st.error(f"⚠️ **Overvalued**: Current price ({_ticker_csym}{current_price:.2f}) exceeds ALL valuation methods. Consider taking profits.")
                        elif current_price < avg_fair_value:
                            st.info(f"📈 **Undervalued**: Current price ({_ticker_csym}{current_price:.2f}) is below average fair value. Potential upside of {((avg_fair_value - current_price) / current_price * 100):.1f}%")
                        else:
                            st.warning(f"📉 **Fairly Valued to Overvalued**: Current price ({_ticker_csym}{current_price:.2f}) is at or above average fair value.")
                
                    else:
                        st.info("💡 Complete more valuation methods to see comprehensive comparison chart")
                
                # Tabs for detailed output - make dynamic
                tab_list_nonbank = [
                    "📊 Historical Analysis",
                    "📈 Projections",
                    "💰 FCF Working",
                    "🎯 WACC Breakdown",
                    "🏆 Valuation Summary",
                    "📉 Sensitivity Analysis",
                    "🎭 Scenario Analysis",
                    "🎲 Monte Carlo",
                    "📁 Comparative Valuation",
                    "🏢 Peer Comparison",
                    "💰 Dividend Discount Model",
                    "🏢 Residual Income Model",
                    "⚙️ Assumptions & Parameters",
                    "🏈 Football Field",
                    "⚖️ IBC Liquidation",
                    "📐 Ratio Analysis",
                    "🩺 Health Check",
                    "🧮 Altman Z-Score",
                    "💼 LBO Analysis"
                ]
                
                # Add Stock Comparison tab if enabled
                enable_stock_comparison_state = st.session_state.get('enable_stock_comparison_listed', False)
                if enable_stock_comparison_state and STOCK_COMPARISON_AVAILABLE:
                    tab_list_nonbank.append("📈 Stock vs Financials")
                
                tabs_nonbank = st.tabs(tab_list_nonbank)
                tab1 = tabs_nonbank[0]
                tab2 = tabs_nonbank[1]
                tab3 = tabs_nonbank[2]
                tab4 = tabs_nonbank[3]
                tab5 = tabs_nonbank[4]
                tab6 = tabs_nonbank[5]
                tab_scenario = tabs_nonbank[6]
                tab_montecarlo = tabs_nonbank[7]
                tab7 = tabs_nonbank[8]
                tab8 = tabs_nonbank[9]
                tab9 = tabs_nonbank[10]
                tab10 = tabs_nonbank[11]
                tab11 = tabs_nonbank[12]
                tab_football = tabs_nonbank[13]
                tab_ibc = tabs_nonbank[14]
                tab_ratios = tabs_nonbank[15]
                tab_health = tabs_nonbank[16]
                tab_zscore = tabs_nonbank[17]
                tab_lbo = tabs_nonbank[18]
                tab_stock_nonbank = tabs_nonbank[19] if len(tabs_nonbank) > 19 else None
                
                with tab1:
                    st.subheader("📊 Comprehensive Historical Financial Analysis")
                
                    # Use advanced charting function
                    st.plotly_chart(create_historical_financials_chart(financials), use_container_width=True)
                
                    # Data tables below charts
                    with st.expander("📋 View Raw Data Tables"):
                        st.subheader("Historical Financials (Last 3 Years)")
                    
                        hist_df = pd.DataFrame({
                            'Year': [str(y) for y in financials['years']],
                            'Revenue': financials['revenue'],
                            'Operating Expenses': financials['opex'],
                            'EBITDA': financials['ebitda'],
                            'Depreciation': financials['depreciation'],
                            'EBIT': financials['ebit'],
                            'Interest': financials['interest'],
                            'Tax': financials['tax'],
                            'NOPAT': financials['nopat']
                        })
                    
                        numeric_cols = hist_df.select_dtypes(include=[np.number]).columns.tolist()
                        format_dict = {col: '{:.2f}' for col in numeric_cols}
                        st.dataframe(hist_df.style.format(format_dict), use_container_width=True)
                    
                        st.subheader("Balance Sheet Metrics")
                        bs_df = pd.DataFrame({
                            'Year': [str(y) for y in financials['years']],
                            'Fixed Assets': financials['fixed_assets'],
                            'Inventory': financials['inventory'],
                            'Receivables': financials['receivables'],
                            'Payables': financials['payables'],
                            'Equity': financials['equity'],
                            'ST Debt': financials['st_debt'],
                            'LT Debt': financials['lt_debt']
                        })
                        numeric_cols = bs_df.select_dtypes(include=[np.number]).columns.tolist()
                        format_dict = {col: '{:.2f}' for col in numeric_cols}
                        st.dataframe(bs_df.style.format(format_dict), use_container_width=True)
                    
                        st.subheader("Working Capital Days")
                        wc_df = pd.DataFrame({
                            'Year': [str(y) for y in financials['years']],
                            'Inventory Days': wc_metrics['inventory_days'],
                            'Debtor Days': wc_metrics['debtor_days'],
                            'Creditor Days': wc_metrics['creditor_days']
                        })
                        numeric_cols = wc_df.select_dtypes(include=[np.number]).columns.tolist()
                        format_dict = {col: '{:.2f}' for col in numeric_cols}
                        st.dataframe(wc_df.style.format(format_dict), use_container_width=True)
                    
                        st.info(f"**Average Working Capital Days:** Inventory: {wc_metrics['avg_inv_days']:.1f} | Debtors: {wc_metrics['avg_deb_days']:.1f} | Creditors: {wc_metrics['avg_cred_days']:.1f}")
                
                with tab2:
                    st.subheader(f"📈 Projected Financials ({projection_years_listed} Years)")
                
                    # Use advanced charting function
                    st.plotly_chart(create_fcff_projection_chart(projections), use_container_width=True)
                
                    # Data table below
                    with st.expander("📋 View Projection Data Table"):
                        proj_df = pd.DataFrame({
                            'Year': [str(y) for y in projections['year']],
                            'Revenue': projections['revenue'],
                            'EBITDA': projections['ebitda'],
                            'Depreciation': projections['depreciation'],
                            'EBIT': projections['ebit'],
                            'NOPAT': projections['nopat'],
                            'Capex': projections['capex'],
                            'Δ WC': projections['delta_wc'],
                            'FCFF': projections['fcff']
                        })
                        numeric_cols = proj_df.select_dtypes(include=[np.number]).columns.tolist()
                        format_dict = {col: '{:.2f}' for col in numeric_cols}
                        st.dataframe(proj_df.style.format(format_dict), use_container_width=True)
                
                    st.info(f"**Key Drivers:** Revenue Growth: {drivers['avg_growth']:.2f}% | Opex Margin: {drivers['avg_opex_margin']:.2f}% | CapEx/Revenue: {drivers['avg_capex_ratio']:.2f}% | Depreciation Rate: {drivers['avg_dep_rate']:.2f}%")
                
                with tab3:
                    st.subheader("Free Cash Flow Working")
                
                    fcff_df = pd.DataFrame({
                        'Year': [str(y) for y in projections['year']],
                        'NOPAT': projections['nopat'],
                        '+ Depreciation': projections['depreciation'],
                        '- Δ WC': projections['delta_wc'],
                        '- Capex': projections['capex'],
                        '= FCFF': projections['fcff'],
                        'Discount Factor': [(1 + wacc_details['wacc']/100)**(-y) for y in projections['year']],
                        'PV(FCFF)': valuation['pv_fcffs']
                    })
                    numeric_cols = fcff_df.select_dtypes(include=[np.number]).columns.tolist()
                    format_dict = {col: '{:.4f}' for col in numeric_cols}
                    st.dataframe(fcff_df.style.format(format_dict), use_container_width=True)
                
                    st.metric("Sum of PV(FCFF)", f"{_ticker_csym} {valuation['sum_pv_fcff']:.2f} Lacs")
                
                with tab4:
                    st.subheader("🎯 WACC Calculation & Breakdown")
                
                    # Advanced WACC breakdown chart
                    st.plotly_chart(create_wacc_breakdown_chart(wacc_details), use_container_width=True)
                
                    st.markdown("---")
                
                    col1, col2 = st.columns(2)
                
                    with col1:
                        st.markdown("**Cost of Equity (Ke)**")
                        st.write(f"Risk-free Rate (Rf): **{wacc_details['rf']:.2f}%**")
                        st.write(f"Market Return (Rm): **{wacc_details['rm']:.2f}%**")
                        st.write(f"Beta (β) - {ticker}: **{wacc_details['beta']:.3f}**")
                        st.write(f"Ke = Rf + β × (Rm - Rf)")
                        st.write(f"Ke = {wacc_details['rf']:.2f}% + {wacc_details['beta']:.3f} × ({wacc_details['rm']:.2f}% - {wacc_details['rf']:.2f}%)")
                        st.write(f"**Ke = {wacc_details['ke']:.2f}%**")
                
                    with col2:
                        st.markdown("**Cost of Debt (Kd)**")
                        st.write(f"Interest Expense: **{_ticker_csym} {financials['interest'][0]:.2f} Lacs**")
                        st.write(f"Total Debt: **{_ticker_csym} {wacc_details['debt']:.2f} Lacs**")
                        st.write(f"Kd (pre-tax) = {wacc_details['kd']:.2f}%")
                        st.write(f"Tax Rate = {tax_rate}%")
                        st.write(f"**Kd (after-tax) = {wacc_details['kd_after_tax']:.2f}%**")
                
                    st.markdown("---")
                    st.markdown("**WACC Calculation**")
                
                    col3, col4 = st.columns(2)
                    with col3:
                        st.write(f"Equity (E): **{_ticker_csym} {wacc_details['equity']:.2f} Lacs** ({wacc_details['we']:.2f}%)")
                        st.write(f"Debt (D): **{_ticker_csym} {wacc_details['debt']:.2f} Lacs** ({wacc_details['wd']:.2f}%)")
                        st.write(f"Total Capital (V): **{_ticker_csym} {wacc_details['equity'] + wacc_details['debt']:.2f} Lacs**")
                
                    with col4:
                        st.write(f"WACC = (E/V × Ke) + (D/V × Kd × (1-Tax))")
                        st.write(f"WACC = ({wacc_details['we']:.2f}% × {wacc_details['ke']:.2f}%) + ({wacc_details['wd']:.2f}% × {wacc_details['kd_after_tax']:.2f}%)")
                        st.write(f"**WACC = {wacc_details['wacc']:.2f}%**")
                
                with tab5:
                    st.subheader("🏆 DCF Valuation Summary")
                
                    # Show FCFF adjustment notice if applicable
                    if valuation.get('fcff_adjusted', False):
                        st.warning("⚠️ **FCFF Adjustment Applied**")
                        adj_details = valuation.get('adjustment_details', {})
                        st.write(f"**Strategy Used:** {adj_details.get('strategy', 'N/A')}")
                        st.write(f"**Original Terminal FCFF:** {_ticker_csym}{projections['fcff'][-1]:.2f} Lacs")
                        st.write(f"**Adjusted Terminal FCFF:** {_ticker_csym}{valuation['adjusted_terminal_fcff']:.2f} Lacs")
                        st.caption("📌 Adjustment details shown during valuation run above")
                        st.markdown("---")
                
                    # Waterfall chart showing value buildup
                    st.plotly_chart(create_waterfall_chart(valuation), use_container_width=True)
                
                    st.markdown("### Terminal Value Calculation")
                
                    # Use adjusted FCFF if available
                    terminal_fcff = valuation.get('adjusted_terminal_fcff', projections['fcff'][-1])
                
                    st.write(f"FCFF (Year {projection_years_listed}): **{_ticker_csym} {terminal_fcff:.2f} Lacs**")
                    if valuation.get('fcff_adjusted', False):
                        st.caption(f"(Original: {_ticker_csym}{projections['fcff'][-1]:.2f} Lacs - Adjusted for sustainability)")
                
                    st.write(f"Terminal Growth Rate (g): **{terminal_growth}%**")
                    st.write(f"FCFF (Year {projection_years_listed + 1}) = FCFF{projection_years_listed} × (1 + g)")
                    st.write(f"FCFF (Year {projection_years_listed + 1}) = {_ticker_csym} {terminal_fcff:.2f} × (1 + {terminal_growth/100})")
                    st.write(f"FCFF (Year {projection_years_listed + 1}) = **{_ticker_csym} {terminal_fcff * (1 + terminal_growth/100):.2f} Lacs**")
                
                    st.write(f"\nTerminal Value = FCFF{projection_years_listed + 1} / (WACC - g)")
                    st.write(f"Terminal Value = {_ticker_csym} {terminal_fcff * (1 + terminal_growth/100):.2f} / ({wacc_details['wacc']:.2f}% - {terminal_growth}%)")
                    st.write(f"**Terminal Value = {_ticker_csym} {valuation['terminal_value']:.2f} Lacs**")
                
                    st.write(f"\nPV(Terminal Value) = TV / (1 + WACC)^{projection_years_listed}")
                    st.write(f"**PV(Terminal Value) = {_ticker_csym} {valuation['pv_terminal_value']:.2f} Lacs**")
                
                    st.markdown("---")
                    st.markdown("### Enterprise Value")

                    # Show growth phase adjustment if applied
                    if valuation.get('growth_phase_adjusted', False):
                        st.info("📊 **Growth-Phase Company:** Sum of PV(FCFF) was adjusted from negative to zero")
                        st.caption(f"Original: {_ticker_csym}{valuation['original_sum_pv_fcff']:.2f} Lacs → Adjusted: {_ticker_csym}{valuation['sum_pv_fcff']:.2f} Lacs")
                
                    _ev_col_label = f'Value ({_ticker_csym} Lacs)'
                    ev_df = pd.DataFrame({
                        'Component': ['Sum of PV(FCFF)', 'PV(Terminal Value)', 'Enterprise Value'],
                        _ev_col_label: [
                            valuation['sum_pv_fcff'],
                            valuation['pv_terminal_value'],
                            valuation['enterprise_value']
                        ]
                    })
                    st.dataframe(ev_df.style.format({_ev_col_label: '{:.2f}'}), use_container_width=True)
                
                    tv_pct = valuation['tv_percentage']
                    if tv_pct > 90:
                        st.warning(f"⚠️ Terminal Value represents {tv_pct:.1f}% of Enterprise Value (>90% is high)")
                    else:
                        st.info(f"Terminal Value represents {tv_pct:.1f}% of Enterprise Value")
                
                    st.markdown("---")
                    st.markdown("### Equity Value & Fair Value per Share")
                
                    equity_calc_df = pd.DataFrame({
                        'Item': ['Enterprise Value', 'Less: Total Debt', 'Add: Cash & Equivalents', '= Net Debt', 'Equity Value', f'Equity Value ({_ticker_csym})', 'Number of Shares', 'Fair Value per Share'],
                        'Value': [
                            f"{_ticker_csym} {valuation['enterprise_value']:.2f} Lacs",
                            f"{_ticker_csym} {valuation['total_debt']:.2f} Lacs",
                            f"{_ticker_csym} {valuation['cash']:.2f} Lacs",
                            f"{_ticker_csym} {valuation['net_debt']:.2f} Lacs",
                            f"{_ticker_csym} {valuation['equity_value']:.2f} Lacs",
                            f"{_ticker_csym} {valuation['equity_value_rupees']:,.0f}",
                            f"{shares:,}" if 'shares' in locals() else f"{num_shares:,}",
                            f"{_ticker_csym} {valuation['fair_value_per_share']:.2f}"
                        ]
                    })
                    st.table(equity_calc_df)
                
                    st.success(f"### 🎯 Fair Value per Share (DCF): {_ticker_csym} {valuation['fair_value_per_share']:.2f}")
                
                    # Show all valuation methods if available
                    if (ddm_result and ddm_result.get('value_per_share', 0) > 0) or (rim_result and rim_result.get('value_per_share', 0) > 0):
                        st.markdown("---")
                        st.markdown("### 💎 Additional Valuation Methods")
                    
                        all_methods = [
                            ("DCF (FCFF)", valuation['fair_value_per_share'])
                        ]
                    
                        if ddm_result and ddm_result.get('value_per_share', 0) > 0:
                            all_methods.append(("DDM (Gordon)", ddm_result['value_per_share']))
                    
                        if rim_result and rim_result.get('value_per_share', 0) > 0:
                            all_methods.append(("Residual Income", rim_result['value_per_share']))
                    
                        # Create comparison table
                        methods_df = pd.DataFrame(all_methods, columns=['Method', 'Fair Value'])
                        methods_df['Fair Value'] = methods_df['Fair Value'].apply(lambda x: f"{_ticker_csym}{x:.2f}")
                        methods_df['Upside/Downside'] = [
                            f"{((v - current_price) / current_price * 100):.1f}%" if current_price > 0 else "N/A"
                            for _, v in all_methods
                        ]
                    
                        st.dataframe(methods_df, use_container_width=True, hide_index=True)
                    
                        avg_all = np.mean([v for _, v in all_methods])
                        st.info(f"📊 **Average Fair Value (All Methods):** {_ticker_csym}{avg_all:.2f}")
                    
                        if current_price > 0:
                            upside_avg = ((avg_all - current_price) / current_price * 100)
                            st.metric("Overall Upside/Downside", f"{upside_avg:+.1f}%")
                
                with tab6:
                    st.subheader("📉 Advanced Sensitivity Analysis")
                
                    wacc_range = np.arange(max(1.0, wacc_details['wacc'] - 3), wacc_details['wacc'] + 3.5, 0.5)
                    g_range = np.arange(max(1.0, terminal_growth - 2), min(terminal_growth + 3, wacc_details['wacc'] - 1), 0.5)
                
                    if len(g_range) == 0:
                        g_range = np.array([terminal_growth])
                
                    # Interactive heatmap
                    st.plotly_chart(create_sensitivity_heatmap(projections, wacc_range, g_range, shares,
                                                              net_debt=valuation['net_debt'],
                                                              terminal_fcff=valuation.get('adjusted_terminal_fcff')),
                                  use_container_width=True)
                
                    # Traditional table below
                    with st.expander("📋 View Sensitivity Data Table"):
                        sensitivity_data = []
                    
                        for w in wacc_range:
                            row_data = {'WACC →': f"{w:.1f}%"}
                            for g_val in g_range:
                                if g_val >= w - 0.1:  # Need at least 0.1% gap
                                    row_data[f"g={g_val:.1f}%"] = "N/A"
                                else:
                                    try:
                                        fcff_n_plus_1 = valuation.get('adjusted_terminal_fcff', projections['fcff'][-1]) * (1 + g_val / 100)
                                        tv = fcff_n_plus_1 / ((w / 100) - (g_val / 100))
                                        pv_tv = tv / ((1 + w / 100) ** projection_years_listed)
                                        ev = sum(_f / ((1 + w / 100) ** (_i + 1)) for _i, _f in enumerate(projections['fcff'])) + pv_tv  # explicit-period PV re-discounted at THIS grid WACC (was stuck at the base WACC)
                                        eq_val = ev - valuation['net_debt']
                                        eq_val_rupees = eq_val * 100000
                                        fv = eq_val_rupees / shares if shares > 0 else 0
                                        row_data[f"g={g_val:.1f}%"] = f"{_ticker_csym}{fv:.2f}"
                                    except:
                                        row_data[f"g={g_val:.1f}%"] = "Error"
                            sensitivity_data.append(row_data)
                    
                        sensitivity_df = pd.DataFrame(sensitivity_data)
                        st.dataframe(sensitivity_df, use_container_width=True)
                    
                        st.caption("Sensitivity table shows Fair Value per Share for different WACC and terminal growth rate combinations")
                
                with tab_scenario:
                    _base_ebitda_margin = (projections['ebitda'][0] / projections['revenue'][0] * 100) \
                        if projections['revenue'] and projections['revenue'][0] > 0 else drivers['avg_opex_margin']
                    _base_wc_pct = (projections['wc'][0] / projections['revenue'][0] * 100) \
                        if projections['revenue'] and projections['revenue'][0] > 0 else 0.0
                    _base_drivers_listed = {
                        'growth': drivers['avg_growth'],
                        'ebitda_margin': _base_ebitda_margin,
                        'wacc': wacc_details['wacc'],
                        'capex_ratio': drivers['avg_capex_ratio'],
                        'wc_pct': _base_wc_pct,
                    }
                    render_scenario_analysis_tab(
                        financials, wc_metrics, projection_years_listed, tax_rate,
                        _base_drivers_listed, scenario_pct_listed, scenario_drivers_listed, scenario_case_listed,
                        terminal_growth, shares, cash_balance,
                        _ticker_csym, current_price=current_price, key_prefix="scenario_listed",
                        fixed_overrides={
                            'depreciation_rate_override': depreciation_rate_override if depreciation_rate_override > 0 else None,
                            'depreciation_method': depreciation_method,
                            'inventory_days_override': inventory_days_override if inventory_days_override > 0 else None,
                            'debtor_days_override': debtor_days_override if debtor_days_override > 0 else None,
                            'creditor_days_override': creditor_days_override if creditor_days_override > 0 else None,
                            'interest_rate_override': interest_rate_override if interest_rate_override > 0 else None,
                        },
                        base_fair_value=valuation['fair_value_per_share']
                    )
                
                with tab_montecarlo:
                    _base_ebitda_margin_mc = (projections['ebitda'][0] / projections['revenue'][0] * 100) \
                        if projections['revenue'] and projections['revenue'][0] > 0 else drivers['avg_opex_margin']
                    _base_wc_pct_mc = (projections['wc'][0] / projections['revenue'][0] * 100) \
                        if projections['revenue'] and projections['revenue'][0] > 0 else 0.0
                    _base_drivers_mc_listed = {
                        'growth': drivers['avg_growth'],
                        'ebitda_margin': _base_ebitda_margin_mc,
                        'wacc': wacc_details['wacc'],
                        'capex_ratio': drivers['avg_capex_ratio'],
                        'wc_pct': _base_wc_pct_mc,
                    }
                    render_monte_carlo_tab(
                        financials, wc_metrics, projection_years_listed, tax_rate,
                        _base_drivers_mc_listed, terminal_growth, shares, cash_balance,
                        _ticker_csym, current_price=current_price, key_prefix="mc_listed",
                        fixed_overrides={
                            'depreciation_rate_override': depreciation_rate_override if depreciation_rate_override > 0 else None,
                            'depreciation_method': depreciation_method,
                            'inventory_days_override': inventory_days_override if inventory_days_override > 0 else None,
                            'debtor_days_override': debtor_days_override if debtor_days_override > 0 else None,
                            'creditor_days_override': creditor_days_override if creditor_days_override > 0 else None,
                            'interest_rate_override': interest_rate_override if interest_rate_override > 0 else None,
                        },
                        base_fair_value=valuation['fair_value_per_share']
                    )
                
                with tab7:
                    st.subheader("🔍 Comparative (Relative) Valuation")
                
                    if comp_tickers_listed:
                        # Use already calculated comp_results from top of page
                        if not comp_results:
                            with st.spinner("Fetching comparable companies data..."):
                                comp_results = perform_comparative_valuation(ticker, comp_tickers_listed, financials, shares, exchange_suffix, projections=projections, valuation_date=valuation_date_listed)
                    
                        if comp_results:
                            # Currency symbol for this exchange
                            _csym = comp_results.get('_currency_symbol', '₹')

                            # Show comparables table
                            st.markdown("### Comparable Companies")
                            comp_df = pd.DataFrame(comp_results['comparables'])
                            if not comp_df.empty:
                                display_comp_df = comp_df[['ticker', 'name', 'price', 'pe', 'pb', 'ps', 'ev_ebitda', 'ev_sales']]
                                st.dataframe(display_comp_df.style.format({
                                    'price': f'{_csym}' + '{:.2f}',
                                    'pe': '{:.2f}x',
                                    'pb': '{:.2f}x',
                                    'ps': '{:.2f}x',
                                    'ev_ebitda': '{:.2f}x',
                                    'ev_sales': '{:.2f}x'
                                }), use_container_width=True)
                        
                            # Show multiples statistics
                            st.markdown("### Peer Multiples Statistics")
                            for multiple, stats in comp_results['multiples_stats'].items():
                                with st.expander(f"📊 {multiple.upper()} - Avg: {stats['average']:.2f}x, Median: {stats['median']:.2f}x"):
                                    st.write(f"**Range:** {stats['min']:.2f}x - {stats['max']:.2f}x")
                                    st.write(f"**Std Dev:** {stats['std']:.2f}x")
                                    st.write(f"**Peer Values:** {', '.join([f'{v:.2f}x' for v in stats['values']])}")
                        
                            # Show implied valuations
                            st.markdown("### Implied Fair Values")
                        
                            all_avg_values = []
                            all_median_values = []
                        
                            for method_key, val_data in comp_results['valuations'].items():
                                st.markdown(f"#### {val_data['method']}")
                            
                                col1, col2 = st.columns(2)
                            
                                with col1:
                                    st.markdown("**Using Average Multiple:**")
                                    st.write(val_data['formula_avg'])
                                    st.metric("Fair Value (Avg)", f"{_csym}{val_data['fair_value_avg']:.2f}", 
                                            f"{val_data['upside_avg']:.1f}%" if val_data['current_price'] else None)
                                    all_avg_values.append(val_data['fair_value_avg'])
                            
                                with col2:
                                    st.markdown("**Using Median Multiple:**")
                                    st.write(val_data['formula_median'])
                                    st.metric("Fair Value (Median)", f"{_csym}{val_data['fair_value_median']:.2f}",
                                            f"{val_data['upside_median']:.1f}%" if val_data['current_price'] else None)
                                    all_median_values.append(val_data['fair_value_median'])
                            
                                st.markdown("---")
                        
                            # Forward P/E if available
                            if 'forward_pe' in comp_results:
                                st.markdown("#### 📅 12-Month Forward P/E Valuation")
                                fpe = comp_results['forward_pe']
                            
                                st.info(f"**{fpe.get('calculation_note', '12-Month Forward EPS projection')}**")
                            
                                col1, col2, col3 = st.columns(3)
                                with col1:
                                    st.metric("Current EPS", f"{_csym}{fpe.get('current_eps', 0):.2f}")
                                with col2:
                                    st.metric("Forward EPS (12M)", f"{_csym}{fpe['forward_eps']:.2f}",
                                            delta=f"{fpe.get('earnings_growth_rate', 0):.1f}%",
                                            help="Projected EPS for next 12 months")
                                with col3:
                                    st.metric("Growth Rate", f"{fpe.get('earnings_growth_rate', 0):.1f}%")
                            
                                st.markdown("---")
                            
                                col1, col2 = st.columns(2)
                                with col1:
                                    st.write(fpe['formula_avg'])
                                    st.metric("Forward Fair Value (Avg)", f"{_csym}{fpe['fair_value_avg']:.2f}")
                                with col2:
                                    st.write(fpe['formula_median'])
                                    st.metric("Forward Fair Value (Median)", f"{_csym}{fpe['fair_value_median']:.2f}")
                            
                                st.markdown("---")
                            # Summary statistics
                            if all_avg_values and all_median_values:
                                st.markdown("### 📈 Comparative Valuation Summary")
                            
                                col1, col2, col3 = st.columns(3)
                            
                                with col1:
                                    st.metric("Average (All Methods)", f"{_csym}{np.mean(all_avg_values):.2f}")
                                    st.metric("Median (All Methods)", f"{_csym}{np.median(all_median_values):.2f}")
                            
                                with col2:
                                    st.metric("Min Fair Value", f"{_csym}{min(all_avg_values + all_median_values):.2f}")
                                    st.metric("Max Fair Value", f"{_csym}{max(all_avg_values + all_median_values):.2f}")
                            
                                with col3:
                                    if valuation['fair_value_per_share'] > 0:
                                        st.metric("DCF Fair Value", f"{_csym}{valuation['fair_value_per_share']:.2f}")
                                        combined_avg = (np.mean(all_avg_values) + valuation['fair_value_per_share']) / 2
                                        st.metric("DCF + Comp Avg", f"{_csym}{combined_avg:.2f}")
                        
                            # VISUAL ANALYSIS - CHARTS
                            st.markdown("---")
                            st.markdown("### 📊 Visual Analysis")
                        
                            comp_df_charts = pd.DataFrame(comp_results['comparables'])
                            if not comp_df_charts.empty:
                                from plotly.subplots import make_subplots
                            
                                # Chart 1: Multiples Comparison (4-in-1)
                                fig1 = make_subplots(rows=2, cols=2, subplot_titles=('P/E Ratio', 'P/B Ratio', 'EV/EBITDA', 'P/S Ratio'))
                                fig1.add_trace(go.Bar(x=comp_df_charts['name'], y=comp_df_charts['pe'], marker_color='steelblue', showlegend=False), row=1, col=1)
                                if 'pe' in comp_results['multiples_stats']:
                                    fig1.add_hline(y=comp_results['multiples_stats']['pe']['average'], line_dash="dash", line_color="red", row=1, col=1)
                                fig1.add_trace(go.Bar(x=comp_df_charts['name'], y=comp_df_charts['pb'], marker_color='lightcoral', showlegend=False), row=1, col=2)
                                if 'pb' in comp_results['multiples_stats']:
                                    fig1.add_hline(y=comp_results['multiples_stats']['pb']['average'], line_dash="dash", line_color="red", row=1, col=2)
                                fig1.add_trace(go.Bar(x=comp_df_charts['name'], y=comp_df_charts['ev_ebitda'], marker_color='mediumseagreen', showlegend=False), row=2, col=1)
                                if 'ev_ebitda' in comp_results['multiples_stats']:
                                    fig1.add_hline(y=comp_results['multiples_stats']['ev_ebitda']['average'], line_dash="dash", line_color="red", row=2, col=1)
                                fig1.add_trace(go.Bar(x=comp_df_charts['name'], y=comp_df_charts['ps'], marker_color='plum', showlegend=False), row=2, col=2)
                                if 'ps' in comp_results['multiples_stats']:
                                    fig1.add_hline(y=comp_results['multiples_stats']['ps']['average'], line_dash="dash", line_color="red", row=2, col=2)
                                fig1.update_layout(height=700, title_text="Peer Valuation Multiples Comparison")
                                fig1.update_xaxes(tickangle=-45)
                                st.plotly_chart(fig1, use_container_width=True)
                            
                                # Chart 2: Financial Metrics
                                fig2 = go.Figure()
                                # Use local currency label for chart
                                _fin_unit = "Crores" if _csym == '₹' else "Millions"
                                _fin_div  = 1e7 if _csym == '₹' else 1e6
                                fig2.add_trace(go.Bar(name='Revenue', x=comp_df_charts['name'], y=comp_df_charts['revenue']/_fin_div, marker_color='steelblue'))
                                fig2.add_trace(go.Bar(name='EBITDA', x=comp_df_charts['name'], y=comp_df_charts['ebitda']/_fin_div, marker_color='lightcoral'))
                                fig2.add_trace(go.Bar(name='Net Income', x=comp_df_charts['name'], y=comp_df_charts['net_income']/_fin_div, marker_color='mediumseagreen'))
                                fig2.update_layout(title=f"Financial Metrics ({_csym} {_fin_unit})", barmode='group', height=500, xaxis_tickangle=-45)
                                st.plotly_chart(fig2, use_container_width=True)
                            
                                # Chart 3: Implied Valuations
                                if comp_results['valuations']:
                                    val_methods = [m.upper().replace('_', ' ') for m in comp_results['valuations'].keys()]
                                    val_avg = [d['fair_value_avg'] for d in comp_results['valuations'].values()]
                                    val_median = [d['fair_value_median'] for d in comp_results['valuations'].values()]
                                    fig3 = go.Figure()
                                    fig3.add_trace(go.Bar(name='Avg', x=val_methods, y=val_avg, marker_color='steelblue'))
                                    fig3.add_trace(go.Bar(name='Median', x=val_methods, y=val_median, marker_color='lightcoral'))
                                    if comp_results['target'].get('current_price', 0) > 0:
                                        fig3.add_hline(y=comp_results['target']['current_price'], line_dash="dash", line_color="green", annotation_text="Current")
                                    if valuation.get('fair_value_per_share', 0) > 0:
                                        fig3.add_hline(y=valuation['fair_value_per_share'], line_dash="dot", line_color="purple", annotation_text="DCF")
                                    fig3.update_layout(title="Implied Valuations", barmode='group', yaxis_title=f"Price ({_csym})", height=500)
                                    st.plotly_chart(fig3, use_container_width=True)
                
                    else:
                        st.info("Enter comparable tickers above to see relative valuation")
                    
                    st.markdown("---")
                    render_implied_multiples_panel(
                        valuation, financials, shares,
                        comp_results=comp_results, current_price=current_price,
                        key_prefix="implied_listed"
                    )
                with tab8:
                    st.subheader("🏢 Advanced Peer Comparison Dashboard")
                
                    if comp_tickers_listed:
                        from peer_comparison_charts import create_peer_comparison_dashboard
                        create_peer_comparison_dashboard(ticker, comp_tickers_listed)
                    else:
                        st.info("💡 Click 'Auto-Fetch Peers' button above or enter peer tickers manually to see detailed peer comparison with 3D visualizations")
                
                with tab9:
                    st.subheader("💰 Dividend Discount Model (DDM)")
                    st.caption("Gordon Growth Model for dividend-paying companies")
                
                    # DDM Section
                    if ddm_result and ddm_result.get('value_per_share', 0) > 0:
                        # Top metrics
                        col1, col2, col3, col4 = st.columns(4)
                        with col1:
                            st.metric("Fair Value/Share", f"{_ticker_csym}{ddm_result.get('value_per_share', 0):.2f}",
                                    delta=f"{((ddm_result.get('value_per_share', 0) - current_price) / current_price * 100):.1f}%" if current_price > 0 else None)
                        with col2:
                            st.metric("Current DPS", f"{_ticker_csym}{ddm_result.get('current_dps', 0):.2f}")
                        with col3:
                            st.metric("Dividend Growth", f"{ddm_result.get('dividend_growth', 0):.1f}%")
                        with col4:
                            st.metric("Payout Ratio", f"{ddm_result.get('payout_ratio', 0):.1f}%")
                    
                        st.markdown("---")
                    
                        col_a, col_b = st.columns(2)
                    
                        with col_a:
                            st.markdown("**📊 Model Parameters**")
                            st.write(f"• **Current Dividend per Share:** {_ticker_csym}{ddm_result.get('current_dps', 0):.2f}")
                            st.write(f"• **Next Year DPS (D1):** {_ticker_csym}{ddm_result.get('next_year_dps', 0):.2f}")
                            st.write(f"• **Dividend Growth Rate:** {ddm_result.get('dividend_growth', 0):.2f}%")
                            st.write(f"• **Required Return (Ke):** {ddm_result.get('required_return', 0):.2f}%")
                            st.write(f"• **Payout Ratio:** {ddm_result.get('payout_ratio', 0):.1f}%")
                        
                            if ddm_result.get('using_actual_data'):
                                st.success("✅ Using actual dividend history from market data")
                            else:
                                st.info("ℹ️ Using estimated dividends based on earnings")
                    
                        with col_b:
                            st.markdown("**📐 Gordon Growth Model Formula**")
                            st.markdown("")  # Spacing
                        
                            # Larger, more readable formula
                            st.latex(r"\huge P_0 = \frac{D_1}{r - g}")
                        
                            st.markdown("")  # Spacing
                            st.markdown("**Where:**")
                        
                            # Better formatted definitions with boxes
                            st.markdown(f"""
                            - **P₀** (Fair Value per Share) = **{_ticker_csym}{ddm_result.get('value_per_share', 0):.2f}**
                            - **D₁** (Next Year Dividend) = {_ticker_csym}{ddm_result.get('next_year_dps', 0):.2f}
                            - **r** (Required Return/Cost of Equity) = {ddm_result.get('required_return', 0):.2f}%
                            - **g** (Dividend Growth Rate) = {ddm_result.get('dividend_growth', 0):.2f}%
                            """)
                        
                            st.markdown("")  # Spacing
                            calc_fv = ddm_result.get('next_year_dps', 0) / ((ddm_result.get('required_return', 10) - ddm_result.get('dividend_growth', 5)) / 100) if (ddm_result.get('required_return', 10) - ddm_result.get('dividend_growth', 5)) > 0 else 0
                        
                            st.markdown("**💡 Calculation:**")
                            st.code(f"""
    Fair Value = D₁ / (r - g)
               = {_ticker_csym}{ddm_result.get('next_year_dps', 0):.2f} / ({ddm_result.get('required_return', 0):.2f}% - {ddm_result.get('dividend_growth', 0):.2f}%)
               = {_ticker_csym}{ddm_result.get('next_year_dps', 0):.2f} / {(ddm_result.get('required_return', 10) - ddm_result.get('dividend_growth', 5)):.2f}%
               = {_ticker_csym}{calc_fv:.2f}
                                """, language="text")
                        
                            # 5-year dividend projections
                            if 'projections' in ddm_result and ddm_result.get('projections'):
                                st.markdown("---")
                                st.markdown("**📈 5-Year Dividend Projections**")
                            
                                proj_df = pd.DataFrame(ddm_result['projections'])
                            
                                # Handle source column if present
                                if 'source' in proj_df.columns:
                                    proj_display = pd.DataFrame({
                                        'Year': proj_df['year'],
                                        'Dividend per Share': proj_df['dividend'].apply(lambda x: f"{_ticker_csym}{x:.2f}"),
                                        'PV of Dividend': proj_df['pv_dividend'].apply(lambda x: f"{_ticker_csym}{x:.2f}"),
                                        'Source': proj_df['source']
                                    })
                                else:
                                    proj_display = pd.DataFrame({
                                        'Year': proj_df['year'],
                                        'Dividend per Share': proj_df['dividend'].apply(lambda x: f"{_ticker_csym}{x:.2f}"),
                                        'PV of Dividend': proj_df['pv_dividend'].apply(lambda x: f"{_ticker_csym}{x:.2f}")
                                    })
                            
                                st.dataframe(proj_display, use_container_width=True, hide_index=True)
                        
                            if ddm_result.get('historical_dividends'):
                                st.markdown("---")
                                st.markdown("**📊 Historical Dividends**")
                                hist_divs = ddm_result['historical_dividends']
                                st.write(f"Recent dividend history: {', '.join([f'{_ticker_csym}{d:.2f}' for d in hist_divs])}")
                        
                            # When to use DDM
                            st.markdown("---")
                            st.markdown("**💡 When to use DDM:**")
                            st.write("✅ Company pays regular dividends")
                            st.write("✅ Stable dividend payout history")
                            st.write("✅ Mature, established companies")
                            st.write("✅ Predictable dividend growth")
                    
                    elif ddm_result == None or not ddm_result.get('value_per_share'):
                        st.warning("⚠️ DDM not applicable for this company")
                        if ddm_result and 'reason' in ddm_result:
                            st.info(f"**Reason:** {ddm_result['reason']}")
                        st.caption("DDM requires the company to pay regular dividends. For non-dividend paying companies, focus on DCF and RIM models.")
                    
                        st.markdown("---")
                        st.markdown("**💡 When DDM is applicable:**")
                        st.write("✅ Company has a history of paying dividends")
                        st.write("✅ Dividend payments are stable and predictable")
                        st.write("✅ Company is in mature growth stage")
                        st.write("✅ Required return > dividend growth rate")
                
                    with tab10:
                        st.subheader("🏢 Residual Income Model (RIM)")
                        st.caption("Equity valuation based on book value and excess returns - FULL DISCLOSURE")
                    
                        # RIM Section with COMPLETE transparency
                        if rim_result and rim_result.get('value_per_share', 0) > 0:
                            # Top-level result
                            st.success(f"### 🎯 Fair Value per Share (RIM): {_ticker_csym}{rim_result['value_per_share']:.2f}")
                            
                            # Top metrics
                            col1, col2, col3, col4 = st.columns(4)
                            with col1:
                                st.metric("Fair Value/Share", f"{_ticker_csym}{rim_result['value_per_share']:.2f}",
                                        delta=f"{((rim_result['value_per_share'] - current_price) / current_price * 100):.1f}%" if current_price > 0 else None)
                            with col2:
                                st.metric("Book Value/Share", f"{_ticker_csym}{rim_result['book_value_per_share']:.2f}")
                            with col3:
                                st.metric("ROE", f"{rim_result['roe']:.2f}%")
                            with col4:
                                st.metric("BV Growth", f"{rim_result.get('bv_growth', 10):.1f}%")
                        
                            st.markdown("---")
                        
                            # SECTION 1: FORMULAS
                            st.markdown("### 📐 RIM Formulas & Methodology")
                            
                            col_formula1, col_formula2 = st.columns(2)
                            
                            with col_formula1:
                                st.markdown("**Main Valuation Formula:**")
                                st.latex(r"P_0 = BV_0 + \sum_{t=1}^{n} \frac{RI_t}{(1+r)^t} + \frac{TV}{(1+r)^n}")
                                
                                st.markdown("**Residual Income Formula:**")
                                st.latex(r"RI_t = NI_t - (r \times BV_{t-1})")
                                
                                st.markdown("**Terminal Value Formula:**")
                                st.latex(r"TV = \frac{RI_n \times (1+g)}{r - g}")
                            
                            with col_formula2:
                                st.markdown("**Where:**")
                                st.markdown(f"""
                                - **P₀** = Fair Value per Share = **{_ticker_csym}{rim_result['value_per_share']:.2f}**
                                - **BV₀** = Current Book Value = **{_ticker_csym}{rim_result['current_book_value']:,.0f}**
                                - **RI** = Residual Income (excess return)
                                - **NI** = Net Income (projected)
                                - **r** = Cost of Equity = **{rim_result['cost_of_equity']:.2f}%**
                                - **g** = Terminal Growth = **{rim_result['terminal_growth']:.2f}%**
                                - **n** = Projection Period = **5 years**
                                """)
                            
                            st.markdown("---")
                            
                            # SECTION 2: INPUT PARAMETERS
                            st.markdown("### 📊 Input Parameters & Assumptions")
                            
                            col_input1, col_input2, col_input3 = st.columns(3)
                            
                            with col_input1:
                                st.markdown("**Current State:**")
                                st.write(f"• Book Value (Total): {_ticker_csym}{rim_result['current_book_value']:,.0f}")
                                st.write(f"• Book Value/Share: {_ticker_csym}{rim_result['book_value_per_share']:.2f}")
                                st.write(f"• Current EPS: {_ticker_csym}{rim_result['current_eps']:.2f}")
                                st.write(f"• Number of Shares: {shares:,.0f}")
                            
                            with col_input2:
                                st.markdown("**Profitability:**")
                                st.write(f"• Return on Equity: {rim_result['roe']:.2f}%")
                                st.write(f"• Cost of Equity: {rim_result['cost_of_equity']:.2f}%")
                                excess_return = rim_result['roe'] - rim_result['cost_of_equity']
                                st.write(f"• Excess Return: {excess_return:.2f}%")
                            
                            with col_input3:
                                st.markdown("**Growth Rates:**")
                                st.write(f"• Book Value Growth: {rim_result.get('bv_growth', 10):.2f}%")
                                st.write(f"• Terminal Growth: {rim_result['terminal_growth']:.2f}%")
                                if rim_result.get('using_dcf_projections'):
                                    st.info("✅ Using DCF Projected NOPAT")
                                else:
                                    st.info("📊 Using ROE-based projection")
                            
                            st.markdown("---")
                            
                            # SECTION 3: YEAR-BY-YEAR PROJECTIONS WITH VISUALS
                            st.markdown("### 📈 Year-by-Year Residual Income Projections")
                            
                            projections = rim_result.get('projections', [])
                            if projections and len(projections) > 0:
                                # Create detailed projection table
                                proj_data = []
                                years_list = []
                                bv_list = []
                                ni_list = []
                                ri_list = []
                                pv_ri_list = []
                                
                                for proj in projections:
                                    year = proj.get('year', 0)
                                    bv_year = proj.get('book_value', 0) / 100000  # Convert to Lacs
                                    ni_year = proj.get('net_income', 0) / 100000
                                    ri_year = proj.get('residual_income', 0) / 100000
                                    pv_ri_year = proj.get('pv_ri', 0) / 100000
                                    req_return = bv_year * rim_result['cost_of_equity'] / 100
                                    
                                    years_list.append(f"Year {year}")
                                    bv_list.append(bv_year)
                                    ni_list.append(ni_year)
                                    ri_list.append(ri_year)
                                    pv_ri_list.append(pv_ri_year)
                                    
                                    proj_data.append({
                                        'Year': f"Year {year}",
                                        'Book Value ({_ticker_csym} Lacs)': f"{bv_year:.2f}",
                                        'Net Income ({_ticker_csym} Lacs)': f"{ni_year:.2f}",
                                        'Required Return ({_ticker_csym} Lacs)': f"{req_return:.2f}",
                                        'Residual Income ({_ticker_csym} Lacs)': f"{ri_year:.2f}",
                                        'Discount Factor': f"{1 / ((1 + rim_result['cost_of_equity']/100) ** year):.4f}",
                                        'PV of RI ({_ticker_csym} Lacs)': f"{pv_ri_year:.2f}"
                                    })
                                
                                # Display table
                                proj_df = pd.DataFrame(proj_data)
                                st.dataframe(proj_df, use_container_width=True, hide_index=True)
                                
                                # VISUAL 1: Residual Income by Year (Bar Chart)
                                st.markdown("#### 📊 Visual: Residual Income by Year")
                                
                                fig_ri = go.Figure()
                                
                                # Residual Income bars
                                colors = ['#06A77D' if ri > 0 else '#E63946' for ri in ri_list]
                                fig_ri.add_trace(go.Bar(
                                    x=years_list,
                                    y=ri_list,
                                    name='Residual Income',
                                    marker_color=colors,
                                    text=[f"{_ticker_csym}{ri:.2f}" for ri in ri_list],
                                    textposition='outside'
                                ))
                                
                                fig_ri.update_layout(
                                    title="Residual Income by Year ({_ticker_csym} Lacs)",
                                    xaxis_title="Year",
                                    yaxis_title="Residual Income ({_ticker_csym} Lacs)",
                                    height=320,
                                    showlegend=False,
                                    hovermode='x unified'
                                )
                                
                                st.plotly_chart(fig_ri, use_container_width=True)
                                
                                # VISUAL 2: Present Value Contribution (Bar Chart)
                                st.markdown("#### 💰 Visual: Present Value Contributions")
                                
                                fig_pv = go.Figure()
                                
                                fig_pv.add_trace(go.Bar(
                                    x=years_list,
                                    y=pv_ri_list,
                                    name='PV of RI',
                                    marker_color='#2E86AB',
                                    text=[f"{_ticker_csym}{pv:.2f}" for pv in pv_ri_list],
                                    textposition='outside'
                                ))
                                
                                fig_pv.update_layout(
                                    title="Present Value Contribution by Year ({_ticker_csym} Lacs)",
                                    xaxis_title="Year",
                                    yaxis_title="PV of RI ({_ticker_csym} Lacs)",
                                    height=320,
                                    showlegend=False
                                )
                                
                                st.plotly_chart(fig_pv, use_container_width=True)
                            else:
                                st.warning("⚠️ No year-by-year projection data available.")
                                
                                # Show ACTUAL calculations for ALL years
                                st.markdown("### 🔢 Detailed Calculations for Each Year")
                                
                                for idx, proj in enumerate(projections):
                                    year = proj.get('year', 0)
                                    bv_year = proj.get('book_value', 0) / 100000
                                    ni_year = proj.get('net_income', 0) / 100000
                                    req_ret_year = bv_year * rim_result['cost_of_equity'] / 100
                                    ri_year = proj.get('residual_income', 0) / 100000
                                    pv_ri_year = proj.get('pv_ri', 0) / 100000
                                    discount_factor = 1 / ((1 + rim_result['cost_of_equity']/100) ** year)
                                    
                                    with st.expander(f"📊 Year {year} Calculation", expanded=(idx==0)):
                                        st.code(f"""
YEAR {year} CALCULATIONS:
{'='*60}

Step 1: Book Value
    Book Value (Year {year}) = {_ticker_csym}{bv_year:.2f} Lacs

Step 2: Net Income  
    Net Income (Year {year}) = {_ticker_csym}{ni_year:.2f} Lacs

Step 3: Required Return (Equity Charge)
    Required Return = Book Value × Cost of Equity
                   = {_ticker_csym}{bv_year:.2f} × {rim_result['cost_of_equity']:.2f}%
                   = {_ticker_csym}{req_ret_year:.2f} Lacs

Step 4: Residual Income (Excess Profit)
    Residual Income = Net Income - Required Return
                   = {_ticker_csym}{ni_year:.2f} - {_ticker_csym}{req_ret_year:.2f}
                   = {_ticker_csym}{ri_year:.2f} Lacs
    
    {'✅ CREATING VALUE' if ri_year > 0 else '⚠️ DESTROYING VALUE'} - {'Company earns more than required return' if ri_year > 0 else 'Company earns less than required return'}

Step 5: Present Value (Discount to Today)
    Discount Factor = 1 / (1 + Ke)^{year}
                   = 1 / (1 + {rim_result['cost_of_equity']/100:.4f})^{year}
                   = {discount_factor:.4f}
    
    PV of RI = Residual Income × Discount Factor
            = {_ticker_csym}{ri_year:.2f} × {discount_factor:.4f}
            = {_ticker_csym}{pv_ri_year:.2f} Lacs

{'='*60}
CONTRIBUTION TO FAIR VALUE: {_ticker_csym}{pv_ri_year:.2f} Lacs
                                        """, language="text")
                                
                                # TERMINAL VALUE CALCULATION
                                st.markdown("### 🎯 Terminal Value Calculation")
                                
                                if len(projections) > 0:
                                    last_proj = projections[-1]
                                    last_ri = last_proj.get('residual_income', 0) / 100000
                                    
                                    st.code(f"""
TERMINAL VALUE (Beyond Year 5):
{'='*60}

Step 1: Terminal Year Residual Income
    RI (Year 5) = {_ticker_csym}{last_ri:.2f} Lacs

Step 2: Grow at Terminal Growth Rate
    RI (Year 6) = RI (Year 5) × (1 + g)
                = {_ticker_csym}{last_ri:.2f} × (1 + {rim_result['terminal_growth']/100:.4f})
                = {_ticker_csym}{last_ri * (1 + rim_result['terminal_growth']/100):.2f} Lacs

Step 3: Perpetuity Value (Gordon Growth Model)
    Terminal Value = RI (Year 6) / (Ke - g)
                  = {_ticker_csym}{last_ri * (1 + rim_result['terminal_growth']/100):.2f} / ({rim_result['cost_of_equity']:.2f}% - {rim_result['terminal_growth']:.2f}%)
                  = {_ticker_csym}{last_ri * (1 + rim_result['terminal_growth']/100):.2f} / {rim_result['cost_of_equity'] - rim_result['terminal_growth']:.2f}%
                  = {_ticker_csym}{(last_ri * (1 + rim_result['terminal_growth']/100)) / ((rim_result['cost_of_equity'] - rim_result['terminal_growth']) / 100):.2f} Lacs

Step 4: Discount to Present Value
    Discount Factor = 1 / (1 + Ke)^5
                   = 1 / (1 + {rim_result['cost_of_equity']/100:.4f})^5
                   = {1 / ((1 + rim_result['cost_of_equity']/100) ** 5):.4f}
    
    PV of Terminal Value = TV × Discount Factor
                        = {_ticker_csym}{(last_ri * (1 + rim_result['terminal_growth']/100)) / ((rim_result['cost_of_equity'] - rim_result['terminal_growth']) / 100):.2f} × {1 / ((1 + rim_result['cost_of_equity']/100) ** 5):.4f}
                        = {_ticker_csym}{rim_result['terminal_ri_pv'] / 100000:.2f} Lacs

{'='*60}
TERMINAL VALUE CONTRIBUTION: {_ticker_csym}{rim_result['terminal_ri_pv'] / 100000:.2f} Lacs
                                    """, language="text")
                            
                            st.markdown("---")
                            
                            # SECTION 4: VALUE BUILD-UP WITH VISUALS
                            st.markdown("### 💰 Fair Value Build-Up")
                            
                            pv_ri_per_share = (rim_result['sum_pv_ri'] / shares) if shares > 0 else 0
                            tv_per_share = (rim_result['terminal_ri_pv'] / shares) if shares > 0 else 0
                            
                            # VISUAL: Waterfall Chart for Value Build-up
                            st.markdown("#### 📊 Visual: Fair Value Waterfall")
                            
                            fig_waterfall = go.Figure(go.Waterfall(
                                name="Fair Value",
                                orientation="v",
                                measure=["absolute", "relative", "relative", "total"],
                                x=["Book Value<br>per Share", "PV of RI<br>(Years 1-5)", "Terminal<br>Value", "Fair Value<br>per Share"],
                                textposition="outside",
                                text=[f"{_ticker_csym}{rim_result['book_value_per_share']:.2f}", f"{_ticker_csym}{pv_ri_per_share:.2f}", f"{_ticker_csym}{tv_per_share:.2f}", f"{_ticker_csym}{rim_result['value_per_share']:.2f}"],
                                y=[rim_result['book_value_per_share'], pv_ri_per_share, tv_per_share, rim_result['value_per_share']],
                                connector={"line": {"color": "rgb(63, 63, 63)"}},
                                decreasing={"marker": {"color": "#E63946"}},
                                increasing={"marker": {"color": "#06A77D"}},
                                totals={"marker": {"color": "#2E86AB"}}
                            ))
                            
                            fig_waterfall.update_layout(
                                title="Fair Value Build-Up ({_ticker_csym} per Share)",
                                showlegend=False,
                                height=380,
                                yaxis_title="Value ({_ticker_csym})"
                            )
                            
                            st.plotly_chart(fig_waterfall, use_container_width=True)
                            
                            # VISUAL: Pie Chart for Value Composition
                            st.markdown("#### 🥧 Visual: Value Composition")
                            
                            # Check if we have negative components
                            has_negative = (pv_ri_per_share < 0 or tv_per_share < 0)
                            
                            if has_negative:
                                st.warning("⚠️ **Note:** Company has negative residual income (destroying value). Pie chart shows absolute values for visualization.")
                                
                                # Use absolute values for pie chart
                                fig_pie = go.Figure(data=[go.Pie(
                                    labels=['Book Value', 'PV of RI (5Y)', 'Terminal Value'],
                                    values=[abs(rim_result['book_value_per_share']), abs(pv_ri_per_share), abs(tv_per_share)],
                                    marker=dict(colors=['#2E86AB', '#E63946', '#F4D35E']),
                                    textinfo='label+percent',
                                    texttemplate='<b>%{label}</b><br>%{percent}',
                                    hovertemplate='<b>%{label}</b><br>{_ticker_csym}%{value:.2f} (absolute)<extra></extra>'
                                )])
                                
                                fig_pie.update_layout(
                                    title="Value Components (Absolute Values)",
                                    height=380
                                )
                                
                                st.plotly_chart(fig_pie, use_container_width=True, key="rim_pie_absolute")
                                
                                # Show actual signed values in a bar chart instead
                                st.markdown("#### 📊 Visual: Signed Value Components")
                                
                                fig_bar = go.Figure()
                                
                                components = ['Book Value', 'PV of RI (5Y)', 'Terminal Value', 'Total']
                                values = [rim_result['book_value_per_share'], pv_ri_per_share, tv_per_share, rim_result['value_per_share']]
                                colors_signed = ['#2E86AB', '#E63946' if pv_ri_per_share < 0 else '#06A77D', 
                                                '#E63946' if tv_per_share < 0 else '#06A77D',
                                                '#E63946' if rim_result['value_per_share'] < 0 else '#06A77D']
                                
                                fig_bar.add_trace(go.Bar(
                                    x=components,
                                    y=values,
                                    marker_color=colors_signed,
                                    text=[f"{_ticker_csym}{v:.2f}" for v in values],
                                    textposition='outside'
                                ))
                                
                                fig_bar.update_layout(
                                    title="Fair Value Components (Actual Signed Values)",
                                    yaxis_title="Value per Share ({_ticker_csym})",
                                    height=320,
                                    showlegend=False
                                )
                                
                                st.plotly_chart(fig_bar, use_container_width=True, key="rim_bar_signed")
                                
                            else:
                                # All positive - normal pie chart
                                fig_pie = go.Figure(data=[go.Pie(
                                    labels=['Book Value', 'PV of RI (5Y)', 'Terminal Value'],
                                    values=[rim_result['book_value_per_share'], pv_ri_per_share, tv_per_share],
                                    marker=dict(colors=['#2E86AB', '#06A77D', '#F4D35E']),
                                    textinfo='label+percent+value',
                                    texttemplate='<b>%{label}</b><br>%{percent}<br>{_ticker_csym}%{value:.2f}',
                                    hovertemplate='<b>%{label}</b><br>{_ticker_csym}%{value:.2f}<br>%{percent}<extra></extra>'
                                )])
                                
                                fig_pie.update_layout(
                                    title="Fair Value Composition ({_ticker_csym} per Share)",
                                    height=380
                                )
                                
                                st.plotly_chart(fig_pie, use_container_width=True, key="rim_pie_positive")
                            
                            # Detailed breakdown in columns
                            col_buildup1, col_buildup2 = st.columns([2, 1])
                            
                            with col_buildup1:
                                st.code(f"""
FAIR VALUE CALCULATION (Per Share Basis):
==========================================

Starting Point:
    Current Book Value/Share              = {_ticker_csym}{rim_result['book_value_per_share']:.2f}

Add: Present Value of Residual Income (Years 1-5)
    Sum of PV(RI) per Share               = {_ticker_csym}{pv_ri_per_share:.2f}

Add: Terminal Value
    PV of Terminal RI per Share           = {_ticker_csym}{tv_per_share:.2f}

==========================================
FAIR VALUE PER SHARE                      = {_ticker_csym}{rim_result['value_per_share']:.2f}
                                """, language="text")
                            
                            with col_buildup2:
                                st.markdown("**Total Equity Value:**")
                                st.write(f"• Book Value: {_ticker_csym}{rim_result['current_book_value']/100000:.2f} Lacs")
                                st.write(f"• PV of RI (5Y): {_ticker_csym}{rim_result['sum_pv_ri']/100000:.2f} Lacs")
                                st.write(f"• Terminal Value: {_ticker_csym}{rim_result['terminal_ri_pv']/100000:.2f} Lacs")
                                st.write(f"• **Total**: {_ticker_csym}{rim_result['total_equity_value']/100000:.2f} Lacs")
                                st.write("")
                                st.write(f"÷ Shares: {shares:,.0f}")
                                st.success(f"**= {_ticker_csym}{rim_result['value_per_share']:.2f} per share**")
                            
                            st.markdown("---")
                            
                            # SECTION 5: KEY INSIGHTS
                            st.markdown("### 💡 Key Insights")
                            
                            col_insight1, col_insight2 = st.columns(2)
                            
                            with col_insight1:
                                st.markdown("**Value Composition:**")
                                total_value = rim_result['book_value_per_share'] + pv_ri_per_share + tv_per_share
                                if total_value > 0:
                                    bv_pct = (rim_result['book_value_per_share'] / total_value) * 100
                                    ri_pct = (pv_ri_per_share / total_value) * 100
                                    tv_pct = (tv_per_share / total_value) * 100
                                    
                                    st.write(f"• Book Value: {bv_pct:.1f}% ({_ticker_csym}{rim_result['book_value_per_share']:.2f})")
                                    st.write(f"• RI (5 Years): {ri_pct:.1f}% ({_ticker_csym}{pv_ri_per_share:.2f})")
                                    st.write(f"• Terminal Value: {tv_pct:.1f}% ({_ticker_csym}{tv_per_share:.2f})")
                            
                            with col_insight2:
                                st.markdown("**Economic Profit:**")
                                if rim_result['roe'] > rim_result['cost_of_equity']:
                                    st.success(f"✅ Creating Value: ROE ({rim_result['roe']:.2f}%) > Ke ({rim_result['cost_of_equity']:.2f}%)")
                                    st.write(f"• Excess return: {excess_return:.2f}%")
                                elif rim_result['roe'] < rim_result['cost_of_equity']:
                                    st.error(f"⚠️ Destroying Value: ROE ({rim_result['roe']:.2f}%) < Ke ({rim_result['cost_of_equity']:.2f}%)")
                                    st.write(f"• Value deficit: {excess_return:.2f}%")
                                else:
                                    st.info(f"Earning exactly required return: ROE = Ke = {rim_result['roe']:.2f}%")
                            
                            st.markdown("---")
                            st.caption("📘 **RIM Model Note:** Residual Income Model values companies based on their ability to generate returns above the cost of equity. Positive residual income indicates value creation.")
                        
                            # 5-year RI projections
                            if 'projections' in rim_result and rim_result['projections']:
                                st.markdown("---")
                                st.markdown("**📈 5-Year Residual Income Projections**")
                            
                                proj_df = pd.DataFrame(rim_result['projections'])
                            
                                # Build display with source column if available
                                display_dict = {
                                    'Year': proj_df['year'],
                                    'Book Value': proj_df['book_value'].apply(lambda x: f"{_ticker_csym}{x:,.0f}"),
                                    'Net Income': proj_df['net_income'].apply(lambda x: f"{_ticker_csym}{x:,.0f}"),
                                    'Residual Income': proj_df['residual_income'].apply(lambda x: f"{_ticker_csym}{x:,.0f}"),
                                    'PV of RI': proj_df['pv_ri'].apply(lambda x: f"{_ticker_csym}{x:,.0f}")
                                }
                            
                                if 'source' in proj_df.columns:
                                    display_dict['Source'] = proj_df['source']
                            
                                proj_display = pd.DataFrame(display_dict)
                            
                                st.dataframe(proj_display, use_container_width=True, hide_index=True)
                            
                                st.caption(f"Sum of PV(RI): {_ticker_csym}{rim_result['sum_pv_ri']:,.0f} | Terminal Value PV: {_ticker_csym}{rim_result['terminal_ri_pv']:,.0f}")
                        
                            # When to use RIM
                            st.markdown("---")
                            st.markdown("**💡 When to use RIM:**")
                            st.write("✅ Any company with positive equity")
                            st.write("✅ Banks & financial institutions")
                            st.write("✅ Asset-intensive businesses")
                            st.write("✅ Companies that retain earnings")
                    
                        else:
                            # RIM calculation failed - show detailed error
                            st.error("⚠️ RIM Not Applicable for This Company")
                        
                            if rim_result and rim_result.get('error'):
                                st.markdown("### 🔍 Diagnostic Information")
                            
                                col_err1, col_err2 = st.columns(2)
                            
                                with col_err1:
                                    st.markdown("**❌ Reason:**")
                                    st.write(rim_result.get('reason', 'Unknown error'))
                                
                                    if 'technical_details' in rim_result:
                                        with st.expander("🔧 Technical Details"):
                                            st.code(rim_result['technical_details'])
                            
                                with col_err2:
                                    st.markdown("**💡 Suggestion:**")
                                    st.info(rim_result.get('suggestion', 'Use DCF or other valuation method'))
                            else:
                                st.warning("RIM calculation returned no results")
                        
                            st.markdown("---")
                            st.markdown("### ✅ When RIM Works Best")
                            col_req1, col_req2 = st.columns(2)
                        
                            with col_req1:
                                st.markdown("**Requirements:**")
                                st.write("✅ Positive book value (equity)")
                                st.write("✅ Positive and stable ROE")
                                st.write("✅ Profitable company")
                                st.write("✅ Complete financial data")
                        
                            with col_req2:
                                st.markdown("**Best for:**")
                                st.write("🏦 Banks & financial institutions")
                                st.write("🏗️ Asset-intensive businesses")
                                st.write("📈 Mature, stable companies")
                                st.write("💰 Companies retaining earnings")
                        
                            st.markdown("---")
                            st.markdown("### 🎯 Alternative Valuation Methods")
                            st.write("Since RIM is not applicable, please use:")
                            st.write("• **DCF (Tab 1)** - Primary valuation method")
                            st.write("• **Comparative Valuation (Tab 7)** - P/E, P/B multiples")
                            st.write("• **DDM (Tab 9)** - If company pays dividends")
                
                with tab11:
                    st.subheader("⚙️ Assumptions & Parameters")
                    
                    # Compact Data Config
                    years_display = [str(y).replace('_', '') for y in financials['years']]
                    years_range = f"{years_display[-1]} - {years_display[0]}" if len(years_display) > 1 else years_display[0]
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("📅 Historical Period", f"{len(financials['years'])} years", delta=years_range, delta_color="off")
                    with col2:
                        st.metric("🔮 Projection Years", f"{projection_years_listed} years")
                    with col3:
                        st.metric("🏢 Shares", f"{shares:,}")
                    
                    st.markdown("---")
                    
                    # WACC Bar Chart
                    st.markdown("### 💰 WACC Components")
                    fig_wacc = go.Figure()
                    fig_wacc.add_trace(go.Bar(
                        x=['WACC', 'Ke', 'Kd', 'Terminal g'],
                        y=[wacc_details['wacc'], wacc_details['ke'], wacc_details['kd_after_tax'], terminal_growth],
                        marker=dict(color=['#667eea', '#764ba2', '#f093fb', '#f5576c']),
                        text=[f"{wacc_details['wacc']:.2f}%", f"{wacc_details['ke']:.2f}%", 
                              f"{wacc_details['kd_after_tax']:.2f}%", f"{terminal_growth:.2f}%"],
                        textposition='auto'
                    ))
                    fig_wacc.update_layout(height=220, margin=dict(l=20, r=20, t=10, b=20), showlegend=False,
                                          yaxis=dict(title='%'), paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
                    st.plotly_chart(fig_wacc, use_container_width=True)
                    
                    col_a, col_b = st.columns(2)
                    with col_a:
                        rf_display = f"{wacc_details['rf']:.2f}%"
                        if abs(wacc_details['rf'] - 6.75) < 0.1:
                            rf_display += " (Fallback)"
                        st.caption(f"**Tax:** {tax_rate:.2f}% | **Beta:** {wacc_details['beta']:.3f} | **Rf:** {rf_display}")
                    with col_b:
                        st.caption(f"**Equity Wt:** {wacc_details['we']:.1f}% | **Debt Wt:** {wacc_details['wd']:.1f}%")
                    
                    st.markdown("---")
                    
                    # Projection Drivers - Compact Expander
                    st.markdown("### 📊 Projection Drivers")
                    with st.expander("View Parameters", expanded=False):
                        if drivers:
                            drivers_data = [
                                ['Revenue Growth', f"{drivers.get('revenue_growth_rate', 0):.2f}%", 'Historical CAGR'],
                                ['EBITDA Margin', f"{drivers.get('ebitda_margin', 0):.2f}%", 'Historical Avg'],
                                ['CapEx/Revenue', f"{drivers.get('capex_ratio', 0):.2f}%", 'Historical Avg']
                            ]
                            if wc_metrics.get('avg_inv_days', 0) > 0:
                                drivers_data.append(['Inventory Days', f"{wc_metrics['avg_inv_days']:.0f}", 'Historical'])
                            if wc_metrics.get('avg_deb_days', 0) > 0:
                                drivers_data.append(['Debtor Days', f"{wc_metrics['avg_deb_days']:.0f}", 'Historical'])
                            if wc_metrics.get('avg_cred_days', 0) > 0:
                                drivers_data.append(['Creditor Days', f"{wc_metrics['avg_cred_days']:.0f}", 'Historical'])
                            
                            st.dataframe(pd.DataFrame(drivers_data, columns=['Parameter', 'Value', 'Source']),
                                       use_container_width=True, hide_index=True, height=200)
                
                with tab_football:
                    st.subheader("🏈 Football Field")
                    st.caption("Compares the fair-value range from every valuation method run above, side by side.")
                    band_pct_listed = st.number_input(
                        "± Band for single-point methods (DCF / DDM / RIM) (%):",
                        min_value=1.0, max_value=50.0, value=12.0, step=1.0,
                        key="football_band_listed",
                        help="DCF, DDM and RIM produce one number, not a range — this sets how wide a band to "
                             "draw around that number on the chart. Comparable Valuation uses the actual "
                             "low/high peer multiples instead, so this doesn't affect it."
                    )
                    _ff_methods_listed = []
                    if valuation and valuation.get('fair_value_per_share', 0) > 0:
                        _fv = valuation['fair_value_per_share']
                        _ff_methods_listed.append({'method': 'DCF', 'low': _fv * (1 - band_pct_listed / 100), 'high': _fv * (1 + band_pct_listed / 100), 'mid': _fv})
                    if ddm_result and ddm_result.get('value_per_share', 0) > 0:
                        _fv = ddm_result['value_per_share']
                        _ff_methods_listed.append({'method': 'DDM', 'low': _fv * (1 - band_pct_listed / 100), 'high': _fv * (1 + band_pct_listed / 100), 'mid': _fv})
                    if rim_result and rim_result.get('value_per_share', 0) > 0:
                        _fv = rim_result['value_per_share']
                        _ff_methods_listed.append({'method': 'RIM', 'low': _fv * (1 - band_pct_listed / 100), 'high': _fv * (1 + band_pct_listed / 100), 'mid': _fv})
                    if comp_results:
                        for _key, _val_data in comp_results.get('valuations', {}).items():
                            _stats = comp_results.get('multiples_stats', {}).get(_key)
                            if _stats and _val_data.get('target_metric', 0) > 0 and _val_data.get('fair_value_avg', 0) > 0:
                                _low = _val_data['target_metric'] * _stats.get('min', 0)
                                _high = _val_data['target_metric'] * _stats.get('max', 0)
                                if _high > 0:
                                    _ff_methods_listed.append({'method': _val_data.get('method', _key.upper()), 'low': _low, 'high': _high, 'mid': _val_data['fair_value_avg']})
                    render_football_field_chart(_ff_methods_listed, current_price, _ticker_csym, key_prefix="football_listed")
                
                with tab_ibc:
                    render_ibc_liquidation_tab(
                        financials, _ticker_csym, key_prefix="ibc_listed",
                        dcf_equity_value=valuation.get('equity_value', 0) * 100000 if valuation else None,
                        dcf_fair_value_per_share=valuation.get('fair_value_per_share') if valuation else None,
                        num_shares=shares
                    )
                
                with tab_ratios:
                    render_ratio_dashboard(
                        financials, wc_metrics, _ticker_csym, key_prefix="ratios_listed",
                        current_price=current_price, shares=shares, tax_rate=tax_rate
                    )
                
                with tab_health:
                    render_valuation_health_check(
                        valuation, terminal_growth, wacc_details['wacc'], csym=_ticker_csym, key_prefix="health_listed",
                        ddm_result=ddm_result, rim_result=rim_result, comp_results=comp_results
                    )
                
                with tab_zscore:
                    render_altman_zscore_tab(
                        financials, _ticker_csym, key_prefix="zscore_listed",
                        current_price=current_price, shares=shares
                    )
                
                with tab_lbo:
                    try:
                        render_lbo_analysis_tab(financials, _ticker_csym, key_prefix="lbo_listed", tax_rate=tax_rate)
                    except Exception as _lbo_err:
                        st.error(f"LBO Analysis failed to render: {_lbo_err}")
                        st.code(traceback.format_exc())
                
                # Stock Price Comparison Tab (if enabled)
                if tab_stock_nonbank and enable_stock_comparison_state and STOCK_COMPARISON_AVAILABLE:
                    with tab_stock_nonbank:
                        st.subheader("📈 Stock Price vs Revenue & EPS Analysis")
                        
                        with st.spinner("Fetching stock price data..."):
                            try:
                                # Determine years to fetch (max 4 for Yahoo Finance)
                                years_to_fetch = min(historical_years_listed, 4)
                                
                                # Fetch comparison data
                                stock_comp_data = get_stock_comparison_data_listed(
                                    ticker=full_ticker,
                                    company_name=company_name,
                                    financials=financials,
                                    num_years=years_to_fetch
                                )
                                
                                if stock_comp_data and stock_comp_data['chart_fig']:
                                    st.plotly_chart(stock_comp_data['chart_fig'], use_container_width=True)
                                    
                                    # Show data tables in expanders
                                    with st.expander("📊 View Raw Data"):
                                        col1, col2, col3 = st.columns(3)
                                        
                                        with col1:
                                            st.markdown("**Revenue Data**")
                                            if stock_comp_data['revenue_df'] is not None:
                                                st.dataframe(stock_comp_data['revenue_df'], hide_index=True)
                                            else:
                                                st.info("No revenue data available")
                                        
                                        with col2:
                                            st.markdown("**EPS Data**")
                                            if stock_comp_data['eps_df'] is not None:
                                                st.dataframe(stock_comp_data['eps_df'], hide_index=True)
                                            else:
                                                st.info("No EPS data available")
                                        
                                        with col3:
                                            st.markdown("**Stock Price Summary**")
                                            if stock_comp_data['stock_prices_df'] is not None:
                                                price_df = stock_comp_data['stock_prices_df']
                                                st.metric("Latest Price", f"₹{price_df['Close'].iloc[-1]:.2f}")
                                                st.metric("Period Return", f"{((price_df['Close'].iloc[-1] - price_df['Close'].iloc[0]) / price_df['Close'].iloc[0] * 100):.2f}%")
                                                st.metric("Major Changes", f"{price_df['is_major'].sum()}" if 'is_major' in price_df.columns else "N/A")
                                            else:
                                                st.info("No stock price data available")
                                    
                                    st.info(f"💡 Chart shows {years_to_fetch} years of data (Yahoo Finance limit: 4 years max)")
                                else:
                                    st.warning("Could not generate stock comparison chart. Check if ticker and financial data are available.")
                            
                            except Exception as e:
                                st.error(f"Error generating stock comparison: {str(e)}")
    
    
    elif mode == "Unlisted Company (Excel Upload)":
        st.subheader("📄 Unlisted Company Valuation")

        # ===== VALUATION DATE — controls what "today" means for Rf, Rm, and =====
        # ===== peer prices used in comparative valuation (the target itself =====
        # ===== is unlisted and has no market price, but its peers do)       =====
        st.markdown("### 📅 Valuation Date")
        col_vdu1, col_vdu2 = st.columns([2, 3])
        with col_vdu1:
            valuation_date_unlisted = st.date_input(
                "Valuation Date",
                value=st.session_state.get('valuation_date_unlisted', datetime.now().date()),
                key='valuation_date_unlisted',
                help="Leave as today for a live valuation. Set an earlier date to price peer companies, Rf and Rm as of that date instead of today's market."
            )
        with col_vdu2:
            st.write("")
            if is_valuation_date_today(valuation_date_unlisted):
                st.caption("💡 Using **live/current** market data (today).")
            else:
                st.info(f"📌 Valuing **as of {valuation_date_unlisted.strftime('%d-%b-%Y')}** — peer prices/multiples, Rf and Rm will use data as of this date (not today).")
        st.markdown("---")
        
        # ===== RISK-FREE RATE TICKER - AT THE TOP, ALWAYS VISIBLE =====
        st.markdown("### 🏛️ Risk-Free Rate Configuration")
        st.info("💡 **Default: 6.83%**. Enter any Yahoo Finance ticker and click Fetch. Common options: ^TNX (US 10Y), ^IRX (US 3-month), or any stock/index for custom rates.")
        
        # Initialize if not exists
        if 'cached_rf_rate_unlisted' not in st.session_state:
            st.session_state.cached_rf_rate_unlisted = 6.83
        
        rf_col1, rf_col2, rf_col3 = st.columns([3, 2, 1])
        with rf_col1:
            custom_rf_ticker_unlisted = st.text_input(
                "Yahoo Finance Ticker for Risk-Free Rate",
                value="NIFTYGS10YR.NS",
                key='custom_rf_ticker_unlisted_top',
                help="Enter any ticker - will calculate CAGR for stocks/indices, or use value directly for yields. Examples: NIFTYGS10YR.NS, ^TNX, RELIANCE.NS"
            )
        with rf_col2:
            st.metric("Current RF Rate", f"{st.session_state.cached_rf_rate_unlisted:.2f}%")
        with rf_col3:
            st.write("")
            st.write("")
            
            # Add a counter to track button clicks
            if 'rf_fetch_click_count_unlisted' not in st.session_state:
                st.session_state.rf_fetch_click_count_unlisted = 0
            
            if st.button("🔄 Fetch", key='refresh_rf_unlisted_top'):
                # Increment click counter FIRST
                st.session_state.rf_fetch_click_count_unlisted += 1
                
                # Store debug output in session state so it persists across reruns
                debug_output = []
                debug_output.append(f"🔄 **FETCH BUTTON CLICKED #{st.session_state.rf_fetch_click_count_unlisted} - UNLISTED MODE**")
                debug_output.append(f"📝 Input ticker: `{custom_rf_ticker_unlisted}`")
                
                ticker_to_use = custom_rf_ticker_unlisted.strip() if custom_rf_ticker_unlisted.strip() else None
                debug_output.append(f"📝 Ticker to use (after strip): `{ticker_to_use}`")
                
                debug_output.append("⏳ Calling get_risk_free_rate()...")
                _as_of = None if is_valuation_date_today(valuation_date_unlisted) else valuation_date_unlisted
                fetched_rate, fetch_debug = get_risk_free_rate(ticker_to_use, as_of_date=_as_of)
                
                # Add all debug messages from the function
                debug_output.extend(fetch_debug)
                
                debug_output.append(f"✅ Function returned: {fetched_rate}%")
                
                debug_output.append(f"💾 Updating session state...")
                debug_output.append(f"   - Before: {st.session_state.get('cached_rf_rate_unlisted', 'NOT SET')}")
                st.session_state.cached_rf_rate_unlisted = fetched_rate
                st.session_state.manual_rf_unlisted = fetched_rate  # Update widget state
                debug_output.append(f"   - After: {st.session_state.cached_rf_rate_unlisted}")
                debug_output.append(f"🔄 Widget state updated to: {fetched_rate}%")
                
                # Store debug output in session state
                st.session_state.rf_fetch_debug_unlisted = debug_output
                st.session_state.rf_fetch_success_unlisted = True
                
                # Rerun to update UI
                st.rerun()
        
        st.markdown("---")
        # ===== END RF RATE CONFIG =====
        
        # ===== MARKET RETURN TICKER =====
        st.markdown("### 📈 Market Return Configuration")
        st.info("💡 Use `%5E` prefix (e.g., `%5EBSESN` for Sensex). **Default**: 12.0%")
        
        if 'cached_rm_rate_unlisted' not in st.session_state:
            st.session_state.cached_rm_rate_unlisted = 12.0
        
        rm_col1, rm_col2, rm_col3 = st.columns([3, 2, 1])
        with rm_col1:
            custom_rm_ticker_unlisted = st.text_input(
                "Yahoo Finance Ticker for Market Return",
                value="%5EBSESN",
                key='custom_rm_ticker_unlisted_top',
                help="Use %5E prefix. Examples: %5EBSESN, %5ENSEI"
            )
        with rm_col2:
            st.metric("Current Market Return", f"{st.session_state.cached_rm_rate_unlisted:.2f}%")
        with rm_col3:
            st.write("")
            st.write("")
            
            if 'rm_fetch_click_count_unlisted' not in st.session_state:
                st.session_state.rm_fetch_click_count_unlisted = 0
            
            if st.button("🔄 Fetch", key='refresh_rm_unlisted_top'):
                st.session_state.rm_fetch_click_count_unlisted += 1
                
                debug_output = []
                debug_output.append(f"🔄 **FETCH #{st.session_state.rm_fetch_click_count_unlisted} - UNLISTED MODE**")
                
                ticker_to_use = custom_rm_ticker_unlisted.strip() if custom_rm_ticker_unlisted.strip() else None
                _as_of = None if is_valuation_date_today(valuation_date_unlisted) else valuation_date_unlisted
                fetched_rate, fetch_debug = get_market_return(ticker_to_use, as_of_date=_as_of)
                
                debug_output.extend(fetch_debug)
                debug_output.append(f"✅ Returned: {fetched_rate}%")
                
                st.session_state.cached_rm_rate_unlisted = fetched_rate
                st.session_state.manual_rm_unlisted = fetched_rate  # Update widget state
                
                st.session_state.rm_fetch_debug_unlisted = debug_output
                st.rerun()
        
        st.markdown("---")
    
        # Template download section
        st.markdown("#### 📥 Download Excel Template")
        st.caption("Use this template to enter your company's financial data")
    
        # Create template in memory
        from io import BytesIO
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl import Workbook
    
        def create_template():
            wb = Workbook()
            if 'Sheet' in wb.sheetnames:
                del wb['Sheet']
        
            # Balance Sheet
            ws_bs = wb.create_sheet('BalanceSheet')
            ws_bs['A1'] = 'BALANCE SHEET'
            ws_bs['B1'] = 23
            ws_bs['C1'] = 24
            ws_bs['D1'] = 25
        
            for cell in ['A1', 'B1', 'C1', 'D1']:
                ws_bs[cell].font = Font(bold=True, color='FFFFFF')
                ws_bs[cell].fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
                ws_bs[cell].alignment = Alignment(horizontal='center')
        
            bs_items = [
                'Equity and Liabilities', 'Equity', 'Share Capital', 'Reserves and Surplus', 'Other Equity',
                'Total Equity', 'Liabilities', 'Non-current Liabilities', 'Long Term Borrowings',
                'Net Deferred Tax Liabilities', 'Other Long Term Liabilities', 'Long Term Provisions',
                'Total Non-current Liabilities', 'Current Liabilities', 'Short Term Borrowings',
                'Trade Payables', 'Other Current Liabilities', 'Short Term Provisions',
                'Total Current Liabilities', 'Total Equity and Liabilities', '',
                'Assets', 'Non-current Assets', 'Tangible Assets', 'Intangible Assets',
                'Capital Work in Progress', 'Non-current Investments', 'Long Term Loans and Advances',
                'Other Non-current Assets', 'Total Non-current Assets', 'Current Assets',
                'Inventories', 'Trade Receivables', 'Cash and Bank Balances',
                'Short Term Loans and Advances', 'Other Current Assets', 'Total Current Assets', 'Total Assets'
            ]
        
            row = 2
            for item in bs_items:
                ws_bs[f'A{row}'] = item
                if item in ['Equity and Liabilities', 'Equity', 'Liabilities', 'Non-current Liabilities',
                            'Current Liabilities', 'Assets', 'Non-current Assets', 'Current Assets']:
                    ws_bs[f'A{row}'].font = Font(bold=True)
                row += 1
        
            ws_bs.column_dimensions['A'].width = 35
            ws_bs.column_dimensions['B'].width = 15
            ws_bs.column_dimensions['C'].width = 15
            ws_bs.column_dimensions['D'].width = 15
        
            # Profit & Loss
            ws_pl = wb.create_sheet('Profit&Loss')
            ws_pl['A1'] = 'PROFIT & LOSS'
            ws_pl['B1'] = 23
            ws_pl['C1'] = 24
            ws_pl['D1'] = 25
        
            for cell in ['A1', 'B1', 'C1', 'D1']:
                ws_pl[cell].font = Font(bold=True, color='FFFFFF')
                ws_pl[cell].fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
                ws_pl[cell].alignment = Alignment(horizontal='center')
        
            pl_items = [
                'Net Revenue', 'Revenue Growth', 'Operating Cost', 'Cost of Materials Consumed',
                'Purchases of Stock-in-trade', 'Changes in Inventories / Finished Goods',
                'Employee Benefit Expense', 'Other Expenses', 'Total Operating Cost', 'EBITDA',
                'Other Income', 'Depreciation and Amortization Expense', 'Profit Before Interest and Tax',
                'Finance Costs', 'Profit Before Tax and Exceptional Items Before Tax',
                'Exceptional Items Before Tax', 'Profit Before Tax', 'Income Tax',
                'Profit for the Period from Continuing Operations'
            ]
        
            row = 2
            for item in pl_items:
                ws_pl[f'A{row}'] = item
                if item in ['Net Revenue', 'Total Operating Cost', 'EBITDA', 'Profit Before Interest and Tax',
                            'Profit Before Tax', 'Profit for the Period from Continuing Operations']:
                    ws_pl[f'A{row}'].font = Font(bold=True)
                row += 1
        
            ws_pl.column_dimensions['A'].width = 50
            ws_pl.column_dimensions['B'].width = 15
            ws_pl.column_dimensions['C'].width = 15
            ws_pl.column_dimensions['D'].width = 15
        
            buffer = BytesIO()
            wb.save(buffer)
            buffer.seek(0)
            return buffer.getvalue()
    
        template_data = create_template()
    
        st.download_button(
            label="📥 Download Excel Template",
            data=template_data,
            file_name="Financials_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Download this template and fill in your company's financial data"
        )
    
        st.markdown("---")
    
        col1, col2 = st.columns(2)
    
        with col1:
            company_name = st.text_input("Company Name:")
            excel_file = st.file_uploader("Upload Financial Excel Template", type=['xlsx', 'xls'])
        
            st.markdown("---")
            st.markdown("**📊 Peer Companies (Both Exchanges)**")
        
            # NSE Peers Box
            nse_peers = st.text_input(
                "NSE Peer Tickers (comma-separated):",
                placeholder="e.g., RELIANCE, TCS, INFY",
                key='nse_peers_unlisted',
                help="Enter NSE-listed peer companies for beta calculation"
            )
        
            # BSE Peers Box
            bse_peers = st.text_input(
                "BSE Peer Tickers (comma-separated):",
                placeholder="e.g., RELIANCE, TCS, INFY",
                key='bse_peers_unlisted',
                help="Enter BSE-listed peer companies for beta calculation"
            )
        
            # Combine peers with their exchange suffixes
            peer_tickers = ""
            if nse_peers.strip():
                nse_list = [f"{t.strip()}.NS" if '.NS' not in t and '.BO' not in t else t.strip() for t in nse_peers.split(',') if t.strip()]
                peer_tickers = ",".join(nse_list)
            if bse_peers.strip():
                bse_list = [f"{t.strip()}.BO" if '.NS' not in t and '.BO' not in t else t.strip() for t in bse_peers.split(',') if t.strip()]
                if peer_tickers:
                    peer_tickers += "," + ",".join(bse_list)
                else:
                    peer_tickers = ",".join(bse_list)
    
        with col2:
            num_shares = st.number_input("Number of Shares Outstanding:", min_value=1, value=100, step=1)
            tax_rate = st.number_input("Tax Rate (%):", min_value=0.0, max_value=100.0, value=25.0, step=0.5)
            terminal_growth = st.number_input("Terminal Growth Rate (%):", min_value=0.0, max_value=10.0, value=4.0, step=0.5)
            
            # Historical years input
            st.markdown("**⏱️ Historical Data Years**")
            historical_years_unlisted = st.number_input(
                "Historical Years to Use",
                min_value=2,
                max_value=10,
                value=3,
                step=1,
                key='unlisted_historical_years',
                help="Number of historical years to use for calculations. Excel must contain at least this many years of data."
            )
            
            # Risk-free rate manual input
            st.markdown("**🏛️ Risk-Free Rate (G-Sec 10Y)**")
            
            if 'manual_rf_unlisted' not in st.session_state:
                st.session_state.manual_rf_unlisted = st.session_state.get('cached_rf_rate_unlisted', 6.83)
            
            manual_rf_rate = st.number_input(
                f"Risk-Free Rate (%)",
                min_value=0.0,
                max_value=20.0,
                value=st.session_state.manual_rf_unlisted,
                step=0.1,
                key='manual_rf_unlisted',
                help="Auto-fetched from ticker above. You can manually edit this value."
            )
            
            # Market Return manual input
            st.markdown("**📈 Market Return (Rm)**")
            
            if 'manual_rm_unlisted' not in st.session_state:
                st.session_state.manual_rm_unlisted = st.session_state.get('cached_rm_rate_unlisted', 12.0)
            
            manual_rm_rate = st.number_input(
                f"Market Return (%)",
                min_value=0.0,
                max_value=50.0,
                value=st.session_state.manual_rm_unlisted,
                step=0.1,
                key='manual_rm_unlisted',
                help="Auto-fetched from ticker above. You can manually edit this value."
            )
            
            # Manual discount rate override
            st.markdown("---")
            st.markdown("**💰 Discount Rate Override (Optional)**")
            manual_discount_rate_unlisted = st.number_input(
                "Manual Discount Rate Override (%):",
                min_value=0.0,
                max_value=50.0,
                value=0.0,
                step=0.5,
                key='manual_discount_unlisted',
                help="⚠️ Override WACC calculation. Leave at 0 to use auto-calculated WACC."
            )
            if manual_discount_rate_unlisted > 0:
                st.info(f"💡 Using manual discount rate: {manual_discount_rate_unlisted:.2f}% (Overriding WACC)")
            
            # ── Beta Calculation Period ────────────────────────────────────
            st.markdown("**📅 Beta Calculation Period**")
            st.caption("Daily returns only. Defaults to 3 years ending today if left unchanged.")
            _beta_default_end_unlisted   = datetime.now().date()
            _beta_default_start_unlisted = (_beta_default_end_unlisted.replace(year=_beta_default_end_unlisted.year - 3))
            _uc1, _uc2 = st.columns(2)
            with _uc1:
                beta_start_date_unlisted = st.date_input(
                    "From", value=_beta_default_start_unlisted,
                    key='beta_start_unlisted',
                    help="Start date for beta regression window"
                )
            with _uc2:
                beta_end_date_unlisted = st.date_input(
                    "To", value=_beta_default_end_unlisted,
                    key='beta_end_unlisted',
                    help="End date for beta regression window"
                )
            if beta_start_date_unlisted >= beta_end_date_unlisted:
                st.error("⚠️ Beta start date must be before end date.")
            # ── End Beta Period ────────────────────────────────────────────

            # Valuation models to run (RIM only, no DDM for unlisted)
            st.markdown("**🎯 Valuation Models**")
            run_dcf_unlisted = st.checkbox("DCF (FCFF)", value=True, key='unlisted_dcf')
            run_rim_unlisted = st.checkbox("RIM (Residual Income)", value=True, key='unlisted_rim')
            run_comp_unlisted = st.checkbox("Comparative Valuation", value=True, key='unlisted_comp')
    
        with st.expander("⚙️ Advanced Projection Assumptions - FULL CONTROL"):
            st.info("💡 **Complete Control:** Override ANY projection parameter below. Leave at 0 or blank for auto-calculation from historical data.")
        
            st.markdown("### 📊 Revenue & Growth")
            col1, col2, col3 = st.columns(3)
            with col1:
                rev_growth_override_unlisted = st.number_input(
                    "Revenue Growth (%/year)", 
                    min_value=0.0, max_value=200.0, value=0.0, step=0.5,
                    key='unlisted_rev_growth',
                    help="0 = Auto from historical CAGR. Override to use custom growth rate."
                )
            with col2:
                opex_margin_override_unlisted = st.number_input(
                    "Operating Expense Margin (%)", 
                    min_value=0.0, max_value=100.0, value=0.0, step=0.5,
                    key='unlisted_opex_margin',
                    help="0 = Auto from historical average"
                )
            with col3:
                ebitda_margin_override_unlisted = st.number_input(
                    "EBITDA Margin (%)", 
                    min_value=0.0, max_value=100.0, value=0.0, step=0.5,
                    key='unlisted_ebitda',
                    help="0 = Calculated as Revenue - OpEx"
                )

            # ── PER-YEAR PARAMETERS (Unlisted) ────────────────────────────
            st.markdown("### 📅 Year-by-Year Revenue Growth & EBITDA Margin")
            st.caption("Set individual values per projected year. Leave at 0 to fall back to the single overrides above (or auto-calc from history).")
            rev_growth_per_year_unlisted = []
            ebitda_margin_per_year_unlisted = []
            _proj_yrs_unl = int(st.session_state.get('unlisted_proj_years', 5))
            yr_cols_unlisted = st.columns(min(_proj_yrs_unl, 5))
            for _yi in range(_proj_yrs_unl):
                _col = yr_cols_unlisted[_yi % len(yr_cols_unlisted)]
                with _col:
                    st.markdown(f"**Year {_yi + 1}**")
                    _rg = st.number_input(
                        f"Rev Growth %",
                        min_value=0.0, max_value=200.0, value=0.0, step=0.5,
                        key=f'unlisted_rg_yr{_yi+1}',
                        help="0 = use global override / auto"
                    )
                    _em = st.number_input(
                        f"EBITDA Margin %",
                        min_value=0.0, max_value=100.0, value=0.0, step=0.5,
                        key=f'unlisted_em_yr{_yi+1}',
                        help="0 = use global override / auto"
                    )
                    rev_growth_per_year_unlisted.append(_rg)
                    ebitda_margin_per_year_unlisted.append(_em)
            # ── END PER-YEAR ───────────────────────────────────────────────
            
            # ── MULTI-STAGE GROWTH PROFILE (optional, overrides the above) ──
            _ms_enabled_unlisted, _ms_growth_unlisted, _ms_margin_unlisted = render_multistage_growth_ui(
                key_prefix="ms_unlisted",
                total_years=_proj_yrs_unl,
                default_stage1_growth=rev_growth_override_unlisted if rev_growth_override_unlisted > 0 else 15.0,
                terminal_growth=terminal_growth,
                default_stage1_margin=ebitda_margin_override_unlisted if ebitda_margin_override_unlisted > 0 else 20.0
            )
            if _ms_enabled_unlisted:
                rev_growth_per_year_unlisted = _ms_growth_unlisted
                if _ms_margin_unlisted:
                    ebitda_margin_per_year_unlisted = _ms_margin_unlisted
            # ── END MULTI-STAGE ──────────────────────────────────────────────

            st.markdown("### 🏗️ CapEx & Depreciation")
            col4, col5, col6 = st.columns(3)
            with col4:
                capex_ratio_override_unlisted = st.number_input(
                    "CapEx/Revenue (%)", 
                    min_value=0.0, max_value=200.0, value=0.0, step=0.5,
                    key='unlisted_capex_ratio',
                    help="0 = Auto from historical average. Override for custom CapEx assumptions."
                )
            with col5:
                depreciation_rate_override_unlisted = st.number_input(
                    "Depreciation Rate (%)", 
                    min_value=0.0, max_value=30.0, value=0.0, step=0.5,
                    key='unlisted_dep_rate',
                    help="0 = Auto calculated"
                )
            with col6:
                depreciation_method_unlisted = st.selectbox(
                    "Depreciation Method",
                    ["Auto", "% of Fixed Assets", "% of Revenue", "Absolute Value"],
                    key='unlisted_dep_method'
                )
        
            st.markdown("### 💰 Working Capital Management")
            col7, col8, col9 = st.columns(3)
            with col7:
                inventory_days_override_unlisted = st.number_input(
                    "Inventory Days", 
                    min_value=0.0, max_value=365.0, value=0.0, step=1.0,
                    key='unlisted_inv_days',
                    help="0 = Auto from historical average"
                )
            with col8:
                debtor_days_override_unlisted = st.number_input(
                    "Debtor/Receivables Days", 
                    min_value=0.0, max_value=365.0, value=0.0, step=1.0,
                    key='unlisted_deb_days',
                    help="0 = Auto from historical average"
                )
            with col9:
                creditor_days_override_unlisted = st.number_input(
                    "Creditor/Payables Days", 
                    min_value=0.0, max_value=365.0, value=0.0, step=1.0,
                    key='unlisted_cred_days',
                    help="0 = Auto from historical average"
                )
        
            st.markdown("### 📈 Tax & Interest")
            col10, col11, col12 = st.columns(3)
            with col10:
                interest_rate_override_unlisted = st.number_input(
                    "Interest Rate (%)", 
                    min_value=0.0, max_value=30.0, value=0.0, step=0.25,
                    key='unlisted_interest',
                    help="0 = Auto calculated from Debt"
                )
            with col11:
                working_capital_as_pct_revenue_unlisted = st.number_input(
                    "Working Capital % of Revenue", 
                    min_value=0.0, max_value=50.0, value=0.0, step=0.5,
                    key='unlisted_wc_pct',
                    help="0 = Calculate from Inv+Deb-Cred days"
                )
            with col12:
                projection_years = st.number_input(
                    "Projection Years", 
                    min_value=3, max_value=15, value=5, step=1,
                    key='unlisted_proj_years',
                    help="Number of years to project forward"
                )
        
        # RIM PARAMETERS FOR UNLISTED MODE
        # ================================
        with st.expander("💎 RIM Parameters (Leave at 0 for Auto-Calculation)"):
            st.info("💡 **Note:** These parameters apply to Residual Income Model. Leave at 0 for auto-calculation from actual data.")
        
            st.markdown("### 🏢 Residual Income Model (RIM) Parameters")
            col_rim1, col_rim2, col_rim3, col_rim4 = st.columns(4)
        
            with col_rim1:
                rim_required_return_unlisted = st.number_input(
                    "Required Return (%)",
                    min_value=5.0,
                    max_value=30.0,
                    value=12.0,
                    step=0.5,
                    key='unlisted_rim_required_return',
                    help="Discount rate for RIM. Usually = Cost of Equity."
                )
        
            with col_rim2:
                rim_assumed_roe_unlisted = st.number_input(
                    "Assumed ROE (%)",
                    min_value=0.0,
                    max_value=50.0,
                    value=0.0,
                    step=1.0,
                    key='unlisted_rim_roe',
                    help="0 = Auto-calculate from historical data. Return on Equity assumption."
                )
        
            with col_rim3:
                rim_terminal_growth_unlisted = st.number_input(
                    "Terminal Growth (%)",
                    min_value=0.0,
                    max_value=20.0,
                    value=0.0,
                    step=0.5,
                    key='unlisted_rim_terminal_growth',
                    help="0 = Use same as DCF terminal growth. Long-term perpetual growth rate."
                )
        
            with col_rim4:
                rim_projection_years_unlisted = st.number_input(
                    "Projection Years",
                    min_value=0,
                    max_value=15,
                    value=0,
                    step=1,
                    key='unlisted_rim_proj_years',
                    help="0 = Use same as DCF projection years. Number of years to project."
                )
    
        if excel_file and company_name and num_shares:
            # INPUT CHANGE DETECTION - Reset results if key inputs change
            current_inputs_unlisted = {
                'excel_file': excel_file.name if excel_file else None,
                'company_name': company_name,
                'num_shares': num_shares,
                'terminal_growth': terminal_growth,
                'tax_rate': tax_rate,
                'wacc': manual_discount_rate_unlisted if manual_discount_rate_unlisted > 0 else None
            }
            
            # Check if inputs changed
            if 'previous_inputs_unlisted' in st.session_state:
                if st.session_state.previous_inputs_unlisted != current_inputs_unlisted:
                    # Inputs changed - clear results
                    st.session_state.show_results_unlisted = False
            
            # Store current inputs
            st.session_state.previous_inputs_unlisted = current_inputs_unlisted
            
            st.markdown("---")
            scenario_case_unlisted, scenario_pct_unlisted, scenario_drivers_unlisted = render_scenario_settings_ui("scenario_unlisted")
            
            st.markdown("---")
            st.markdown("### 🎯 Ready to Run Valuation")
            st.info("💡 **Click the button below to run valuation.** Results will appear only after clicking.")
            
            col_run1, col_run2 = st.columns([2, 1])
            
            with col_run1:
                if st.button("🚀 Run DCF Valuation", type="primary", key="run_dcf_btn"):
                    st.session_state.show_results_unlisted = True
            
            with col_run2:
                if st.button("🗑️ Clear Results", help="Clear current valuation results", key="clear_unlisted_results"):
                    st.session_state.show_results_unlisted = False
                    st.success("✅ Results cleared!")
                    st.rerun()
        
            if st.session_state.get('show_results_unlisted', False):
                with st.spinner("Processing..."):
                    # Currency symbol for display (unlisted companies are valued in INR/Lacs)
                    _ticker_csym = '₹'

                    # Parse Excel
                    df_bs, df_pl = parse_excel_to_dataframes(excel_file)
                
                    if df_bs is None or df_pl is None:
                        st.error("Failed to parse Excel file")
                        st.stop()
                
                    # Detect year columns
                    year_cols_all = detect_year_columns(df_bs)
                
                    if len(year_cols_all) < 2:
                        st.error("Need at least 2 years of historical data in Excel file")
                        st.stop()
                
                    st.success(f"✅ Excel contains {len(year_cols_all)} years of data")
                    
                    # Limit to user-selected historical years
                    if historical_years_unlisted > len(year_cols_all):
                        st.warning(f"⚠️ Requested {historical_years_unlisted} years but Excel has only {len(year_cols_all)} years. Using all available data.")
                        year_cols = year_cols_all
                    else:
                        year_cols = year_cols_all[-historical_years_unlisted:]
                        st.info(f"📊 Using {len(year_cols)} most recent years as selected")
                
                    # Extract financials
                    financials = extract_financials_unlisted(df_bs, df_pl, year_cols)
                
                    # ================================
                    # BUSINESS MODEL CLASSIFICATION (RULEBOOK SECTION 2)
                    # ================================
                    st.markdown("---")
                    st.subheader("🏢 Business Model Classification")
                
                    classification = classify_business_model(financials, income_stmt=None, balance_sheet=None)
                
                    # Show classification and check if FCFF DCF is allowed
                    should_stop = show_classification_warning(classification)
                
                    if should_stop:
                        st.stop()
                
                    st.markdown("---")
                
                    # Calculate WC metrics
                    wc_metrics = calculate_working_capital_metrics(financials)
                    
                    # Show Working Capital Data Status
                    wc_status_parts = []
                    if wc_metrics.get('has_inventory', False) or (inventory_days_override_unlisted and inventory_days_override_unlisted > 0):
                        inv_source = f"User Override: {inventory_days_override_unlisted} days" if inventory_days_override_unlisted and inventory_days_override_unlisted > 0 else f"Historical: {wc_metrics['avg_inv_days']:.1f} days"
                        wc_status_parts.append(f"✅ Inventory ({inv_source})")
                    else:
                        wc_status_parts.append("⚠️ Inventory (No data)")
                    
                    if wc_metrics.get('has_receivables', False) or (debtor_days_override_unlisted and debtor_days_override_unlisted > 0):
                        deb_source = f"User Override: {debtor_days_override_unlisted} days" if debtor_days_override_unlisted and debtor_days_override_unlisted > 0 else f"Historical: {wc_metrics['avg_deb_days']:.1f} days"
                        wc_status_parts.append(f"✅ Debtors ({deb_source})")
                    else:
                        wc_status_parts.append("⚠️ Debtors (No data)")
                    
                    if wc_metrics.get('has_payables', False) or (creditor_days_override_unlisted and creditor_days_override_unlisted > 0):
                        cred_source = f"User Override: {creditor_days_override_unlisted} days" if creditor_days_override_unlisted and creditor_days_override_unlisted > 0 else f"Historical: {wc_metrics['avg_cred_days']:.1f} days"
                        wc_status_parts.append(f"✅ Creditors ({cred_source})")
                    else:
                        wc_status_parts.append("⚠️ Creditors (No data)")
                    
                    st.info(f"🔍 **Working Capital Projection Status:** {' | '.join(wc_status_parts)}")
                
                    # Project financials
                    projections, drivers = project_financials(
                        financials, wc_metrics, projection_years, tax_rate,
                        rev_growth_override_unlisted if rev_growth_override_unlisted > 0 else None, 
                        opex_margin_override_unlisted if opex_margin_override_unlisted > 0 else None, 
                        capex_ratio_override_unlisted if capex_ratio_override_unlisted > 0 else None,
                        # Per-year controls
                        rev_growth_per_year=rev_growth_per_year_unlisted if (_ms_enabled_unlisted or any(v > 0 for v in rev_growth_per_year_unlisted)) else None,
                        ebitda_margin_per_year=ebitda_margin_per_year_unlisted if (bool(_ms_margin_unlisted) or any(v > 0 for v in ebitda_margin_per_year_unlisted)) else None,
                        # Pass all advanced user controls
                        ebitda_margin_override=ebitda_margin_override_unlisted if ebitda_margin_override_unlisted > 0 else None,
                        depreciation_rate_override=depreciation_rate_override_unlisted if depreciation_rate_override_unlisted > 0 else None,
                        depreciation_method=depreciation_method_unlisted,
                        inventory_days_override=inventory_days_override_unlisted if inventory_days_override_unlisted > 0 else None,
                        debtor_days_override=debtor_days_override_unlisted if debtor_days_override_unlisted > 0 else None,
                        creditor_days_override=creditor_days_override_unlisted if creditor_days_override_unlisted > 0 else None,
                        interest_rate_override=interest_rate_override_unlisted if interest_rate_override_unlisted > 0 else None,
                        working_capital_pct_override=working_capital_as_pct_revenue_unlisted if working_capital_as_pct_revenue_unlisted > 0 else None
                    )
                
                    # Calculate WACC (unlisted companies use auto-fetched risk-free rate)
                    wacc_details = calculate_wacc(financials, tax_rate, peer_tickers=peer_tickers, manual_rf_rate=manual_rf_rate, manual_rm_rate=manual_rm_rate,
                                                  beta_start_date=beta_start_date_unlisted, beta_end_date=beta_end_date_unlisted)
                
                    # DCF Valuation
                    # Extract cash balance
                    cash_balance = financials['cash'][0] if financials['cash'][0] > 0 else 0
                
                    valuation, error = calculate_dcf_valuation(
                        projections, wacc_details, terminal_growth, num_shares, cash_balance
                    )
                
                    if error:
                        st.error(error)
                        st.stop()
                
                    # ================================
                    # DISPLAY RESULTS
                    # ================================
                
                    st.success("✅ Valuation Complete!")
                
                
                    # Key Metrics
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Enterprise Value", f"{_ticker_csym} {valuation['enterprise_value']:.2f} Lacs")
                    with col2:
                        st.metric("Equity Value", f"{_ticker_csym} {valuation['equity_value']:.2f} Lacs")
                    with col3:
                        st.metric("Fair Value/Share", f"{_ticker_csym} {valuation['fair_value_per_share']:.2f}")
                    with col4:
                        st.metric("WACC", f"{wacc_details['wacc']:.2f}%")
                
                    # ============================================
                    # DISPLAY FAIR VALUES SUMMARY AT TOP
                    # ============================================
                    st.markdown("## 🎯 Valuation Summary")
                    
                    fair_values_dict = {}
                    
                    # DCF Value
                    if run_dcf_unlisted:
                        fair_values_dict['DCF (FCFF)'] = valuation['fair_value_per_share']
                    
                    # Calculate RIM if enabled
                    if run_rim_unlisted:
                        try:
                            # Calculate RIM using the correct function
                            rim_required_return = rim_required_return_unlisted if rim_required_return_unlisted > 0 else wacc_details['wacc']
                            rim_roe = rim_assumed_roe_unlisted if rim_assumed_roe_unlisted > 0 else None
                            rim_term_growth = rim_terminal_growth_unlisted if rim_terminal_growth_unlisted > 0 else terminal_growth
                            rim_proj_years = rim_projection_years_unlisted if rim_projection_years_unlisted > 0 else projection_years
                            
                            rim_results_unlisted = calculate_residual_income_model(
                                financials,
                                num_shares,
                                rim_required_return,
                                terminal_growth=rim_term_growth,
                                projection_years=rim_proj_years,
                                assumed_roe=rim_roe
                            )
                            
                            if rim_results_unlisted and 'value_per_share' in rim_results_unlisted:
                                fair_values_dict['RIM (Residual Income)'] = rim_results_unlisted.get('value_per_share', 0)
                        except Exception as e:
                            st.warning(f"RIM calculation encountered an error: {str(e)}")
                    
                    # Comparative valuations - calculate early if comparative is enabled
                    comp_avg = None
                    comp_median = None
                    comp_results = None
                    
                    if run_comp_unlisted and peer_tickers and peer_tickers.strip():
                        try:
                            comp_results = perform_comparative_valuation(
                                target_ticker=None,
                                comp_tickers_str=peer_tickers,
                                target_financials=financials,
                                target_shares=num_shares,
                                exchange_suffix="NS",
                                projections=projections if run_dcf_unlisted else None,
                                use_screener_peers=False,
                                valuation_date=valuation_date_unlisted
                            )
                            
                            if comp_results and 'valuations' in comp_results:
                                all_avg = []
                                all_median = []
                                for val_data in comp_results['valuations'].values():
                                    all_avg.append(val_data['fair_value_avg'])
                                    all_median.append(val_data['fair_value_median'])
                                
                                if all_avg:
                                    comp_avg = np.mean(all_avg)
                                    comp_median = np.median(all_median)
                        except Exception as e:
                            pass  # Silent fail
                
                    # Display fair values in columns
                    fv_cols = st.columns(len(fair_values_dict) + 2)  # +2 for comp avg/median
                    
                    col_idx = 0
                    for method, value in fair_values_dict.items():
                        with fv_cols[col_idx]:
                            st.metric(method, f"₹{value:.2f}")
                        col_idx += 1
                    
                    # Display comparative with actual values if calculated
                    with fv_cols[col_idx]:
                        if comp_avg and comp_avg > 0:
                            st.metric("Comp (Avg)", f"₹{comp_avg:.2f}")
                        else:
                            st.metric("Comp (Avg)", "N/A")
                    with fv_cols[col_idx + 1]:
                        if comp_median and comp_median > 0:
                            st.metric("Comp (Median)", f"₹{comp_median:.2f}")
                        else:
                            st.metric("Comp (Median)", "N/A")
                    
                    st.markdown("---")
                    
                    # ============================================
                    # BUILD DYNAMIC TAB LIST
                    # ============================================
                    tab_list = ["📊 Historical Financials"]
                    if run_dcf_unlisted:
                        tab_list.extend(["📋 Assumptions & Inputs", "📈 Projections", "💰 FCF Working", "🎯 WACC Calculation", "🏆 DCF Summary", "📉 Sensitivity Analysis", "🎭 Scenario Analysis", "🎲 Monte Carlo"])
                    if run_rim_unlisted:
                        tab_list.append("📚 RIM Valuation")
                    if run_comp_unlisted and peer_tickers:
                        tab_list.append("📊 Comparative Valuation")
                    tab_list.append("🏈 Football Field")
                    tab_list.append("⚖️ IBC Liquidation")
                    tab_list.append("📐 Ratio Analysis")
                    tab_list.append("🩺 Health Check")
                    tab_list.append("🧮 Altman Z-Score")
                    tab_list.append("💼 LBO Analysis")
                    
                    tabs = st.tabs(tab_list)
                    tab_idx = 0
                    
                    # ============================================
                    # TAB: HISTORICAL FINANCIALS
                    # ============================================
                    with tabs[tab_idx]:
                        st.subheader("Historical Financials")
                        
                        # Add Historical Financials Chart
                        try:
                            # Unlisted data is already oldest-first, so reverse_years=False
                            st.plotly_chart(create_historical_financials_chart(financials, reverse_years=False), use_container_width=True)
                        except Exception as e:
                            st.error(f"Historical chart error: {str(e)}")
                        
                        st.markdown("---")
                        st.markdown("### Income Statement")
                        
                        hist_df = pd.DataFrame({
                            'Year': [str(y) for y in financials['years']],
                            'Revenue': financials['revenue'],
                            'Operating Expenses': financials['opex'],
                            'EBITDA': financials['ebitda'],
                            'Depreciation': financials['depreciation'],
                            'EBIT': financials['ebit'],
                            'Interest': financials['interest'],
                            'Tax': financials['tax'],
                            'NOPAT': financials['nopat']
                        })
                        numeric_cols = hist_df.select_dtypes(include=[np.number]).columns.tolist()
                        format_dict = {col: '{:.2f}' for col in numeric_cols}
                        st.dataframe(hist_df.style.format(format_dict), use_container_width=True)
                        
                        st.markdown("---")
                        st.subheader("Balance Sheet Metrics")
                        bs_df = pd.DataFrame({
                            'Year': [str(y) for y in financials['years']],
                            'Fixed Assets': financials['fixed_assets'],
                            'Inventory': financials['inventory'],
                            'Receivables': financials['receivables'],
                            'Payables': financials['payables'],
                            'Equity': financials['equity'],
                            'ST Debt': financials['st_debt'],
                            'LT Debt': financials['lt_debt']
                        })
                        numeric_cols = bs_df.select_dtypes(include=[np.number]).columns.tolist()
                        format_dict = {col: '{:.2f}' for col in numeric_cols}
                        st.dataframe(bs_df.style.format(format_dict), use_container_width=True)
                    
                    tab_idx += 1
                    
                    # ============================================
                    # DCF TABS (only if DCF enabled)
                    # ============================================
                    if run_dcf_unlisted:
                        # TAB: Assumptions & Inputs
                        with tabs[tab_idx]:
                            st.subheader("📋 Assumptions & Projection Inputs")
                            
                            st.markdown("### 📊 Projection Drivers")
                            drivers_data = [
                                ['Revenue Growth', f"{drivers.get('avg_growth', 0):.2f}%", 'Historical CAGR'],
                                ['EBITDA Margin', f"{drivers.get('avg_ebitda_margin', 0):.2f}%", 'Historical Avg'],
                                ['CapEx/Revenue', f"{drivers.get('avg_capex_ratio', 0):.2f}%", 'Historical Avg'],
                                ['Depreciation Rate', f"{drivers.get('avg_depreciation_rate', 0):.2f}%", 'Historical Avg'],
                                ['Tax Rate', f"{tax_rate:.2f}%", 'User Input'],
                                ['Terminal Growth', f"{terminal_growth:.2f}%", 'User Input']
                            ]
                            st.dataframe(pd.DataFrame(drivers_data, columns=['Parameter', 'Value', 'Source']),
                                       use_container_width=True, hide_index=True)
                            
                            st.markdown("---")
                            st.markdown("### 💰 Working Capital Metrics")
                            wc_data = [
                                ['Inventory Days', f"{wc_metrics.get('avg_inv_days', 0):.0f}", 'Historical'],
                                ['Debtor Days', f"{wc_metrics.get('avg_deb_days', 0):.0f}", 'Historical'],
                                ['Creditor Days', f"{wc_metrics.get('avg_cred_days', 0):.0f}", 'Historical']
                            ]
                            st.dataframe(pd.DataFrame(wc_data, columns=['Metric', 'Value', 'Source']),
                                       use_container_width=True, hide_index=True)
                        
                        tab_idx += 1
                        
                        # TAB: Projections
                        with tabs[tab_idx]:
                            st.subheader("📈 Financial Projections")
                            
                            # Add Projections Chart
                            try:
                                st.plotly_chart(create_fcff_projection_chart(projections), use_container_width=True)
                            except Exception as e:
                                st.error(f"Projections chart error: {str(e)}")
                            
                            st.markdown("---")
                            
                            proj_df = pd.DataFrame({
                                'Year': [f"Year {y}" for y in projections['year']],
                                'Revenue': projections['revenue'],
                                'EBITDA': projections['ebitda'],
                                'NOPAT': projections['nopat'],
                                'CapEx': projections['capex'],
                                'Δ WC': projections['delta_wc'],
                                'FCFF': projections['fcff']
                            })
                            numeric_cols = proj_df.select_dtypes(include=[np.number]).columns.tolist()
                            format_dict = {col: '{:.2f}' for col in numeric_cols}
                            st.dataframe(proj_df.style.format(format_dict), use_container_width=True)
                        
                        tab_idx += 1
                        
                        # TAB: FCF Working
                        with tabs[tab_idx]:
                            st.subheader("💰 Free Cash Flow Calculation")
                            
                            fcf_df = pd.DataFrame({
                                'Year': [f"Year {y}" for y in projections['year']],
                                'NOPAT': projections['nopat'],
                                '- CapEx': [-c for c in projections['capex']],
                                '- Δ WC': [-wc for wc in projections['delta_wc']],
                                '= FCFF': projections['fcff']
                            })
                            numeric_cols = fcf_df.select_dtypes(include=[np.number]).columns.tolist()
                            format_dict = {col: '{:.2f}' for col in numeric_cols}
                            st.dataframe(fcf_df.style.format(format_dict).background_gradient(cmap='RdYlGn', subset=['= FCFF']), 
                                       use_container_width=True)
                            
                            st.markdown("---")
                            st.subheader("Terminal Value Calculation")
                            
                            fcff_final = valuation.get('adjusted_terminal_fcff', projections['fcff'][-1])
                            fcff_n_plus_1 = fcff_final * (1 + terminal_growth / 100)
                            
                            tv_calc = pd.DataFrame({
                                'Component': ['Final Year FCFF', 'Growth Rate', 'FCFF (Year N+1)', 'WACC', 'Terminal Growth', 'Terminal Value'],
                                'Value': [
                                    f"₹{fcff_final:.2f}L",
                                    f"{terminal_growth:.2f}%",
                                    f"₹{fcff_n_plus_1:.2f}L",
                                    f"{wacc_details['wacc']:.2f}%",
                                    f"{terminal_growth:.2f}%",
                                    f"₹{valuation['terminal_value']:.2f}L"
                                ]
                            })
                            st.dataframe(tv_calc, use_container_width=True, hide_index=True)
                        
                        tab_idx += 1
                        
                        # TAB: WACC Calculation
                        with tabs[tab_idx]:
                            st.subheader("🎯 WACC Calculation")
                            
                            # WACC Breakdown Chart
                            try:
                                st.plotly_chart(create_wacc_breakdown_chart(wacc_details), use_container_width=True)
                            except Exception as e:
                                st.error(f"WACC chart error: {str(e)}")
                            
                            st.markdown("---")
                            
                            wacc_df = pd.DataFrame({
                                'Component': ['Cost of Equity (Ke)', 'Cost of Debt (Kd)', 'Tax Rate', 'After-Tax Kd', 
                                            'Equity Weight', 'Debt Weight', 'WACC'],
                                'Value': [
                                    f"{wacc_details['ke']:.2f}%",
                                    f"{wacc_details['kd']:.2f}%",
                                    f"{tax_rate:.2f}%",
                                    f"{wacc_details['kd_after_tax']:.2f}%",
                                    f"{wacc_details['we']:.2f}%",
                                    f"{wacc_details['wd']:.2f}%",
                                    f"{wacc_details['wacc']:.2f}%"
                                ]
                            })
                            st.dataframe(wacc_df, use_container_width=True, hide_index=True)
                            
                            st.markdown("---")
                            st.markdown("### Cost of Equity (CAPM)")
                            st.latex(r"K_e = R_f + \beta \times (R_m - R_f)")
                            
                            capm_df = pd.DataFrame({
                                'Parameter': ['Risk-Free Rate (Rf)', 'Beta (β)', 'Market Return (Rm)', 'Market Risk Premium', 'Cost of Equity (Ke)'],
                                'Value': [
                                    f"{wacc_details['rf']:.2f}%",
                                    f"{wacc_details['beta']:.2f}",
                                    f"{wacc_details['rm']:.2f}%",
                                    f"{wacc_details['rm'] - wacc_details['rf']:.2f}%",
                                    f"{wacc_details['ke']:.2f}%"
                                ]
                            })
                            st.dataframe(capm_df, use_container_width=True, hide_index=True)
                        
                        tab_idx += 1
                        
                        # TAB: DCF Summary
                        with tabs[tab_idx]:
                            st.subheader("🏆 DCF Valuation Summary")
                            
                            # Waterfall Chart
                            try:
                                st.plotly_chart(create_waterfall_chart(valuation), use_container_width=True)
                            except Exception as e:
                                st.error(f"Waterfall chart error: {str(e)}")
                            
                            st.markdown("---")
                            
                            st.markdown("### Enterprise to Equity Value Bridge")
                            equity_calc_df = pd.DataFrame({
                                'Item': ['Enterprise Value', 'Less: Total Debt', 'Add: Cash & Equivalents', '= Net Debt', 
                                        'Equity Value', 'Equity Value (₹)', 'Number of Shares', 'Fair Value per Share'],
                                'Value (₹ Lacs)': [
                                    f"{valuation['enterprise_value']:.2f}",
                                    f"({valuation['total_debt']:.2f})",
                                    f"{cash_balance:.2f}",
                                    f"({valuation['net_debt']:.2f})",
                                    f"{valuation['equity_value']:.2f}",
                                    f"{valuation['equity_value_rupees']:,.0f}",
                                    f"{num_shares:,}",
                                    f"₹ {valuation['fair_value_per_share']:.2f}"
                                ]
                            })
                            st.table(equity_calc_df)
                            
                            st.success(f"### 🎯 Fair Value per Share (DCF): ₹{valuation['fair_value_per_share']:.2f}")
                        
                        tab_idx += 1
                        
                        # TAB: Sensitivity Analysis
                        with tabs[tab_idx]:
                            st.subheader("📉 Sensitivity Analysis")
                            
                            wacc_range = np.arange(max(1.0, wacc_details['wacc'] - 3), wacc_details['wacc'] + 3.5, 0.5)
                            g_range = np.arange(max(1.0, terminal_growth - 2), min(terminal_growth + 3, wacc_details['wacc'] - 1), 0.5)
                            
                            if len(g_range) == 0:
                                g_range = np.array([terminal_growth])
                            
                            sensitivity_data = []
                            
                            for w in wacc_range:
                                row_data = {'WACC →': f"{w:.1f}%"}
                                for g_val in g_range:
                                    if g_val >= w - 0.1:
                                        row_data[f"g={g_val:.1f}%"] = "N/A"
                                    else:
                                        try:
                                            fcff_n_plus_1 = valuation.get('adjusted_terminal_fcff', projections['fcff'][-1]) * (1 + g_val / 100)
                                            tv = fcff_n_plus_1 / ((w / 100) - (g_val / 100))
                                            pv_tv = tv / ((1 + w / 100) ** projection_years)
                                            ev = sum(_f / ((1 + w / 100) ** (_i + 1)) for _i, _f in enumerate(projections['fcff'])) + pv_tv  # explicit-period PV re-discounted at THIS grid WACC (was stuck at the base WACC)
                                            eq_val = ev - valuation['net_debt']
                                            eq_val_rupees = eq_val * 100000
                                            fv = eq_val_rupees / num_shares if num_shares > 0 else 0
                                            row_data[f"g={g_val:.1f}%"] = f"₹{fv:.2f}"
                                        except:
                                            row_data[f"g={g_val:.1f}%"] = "Error"
                                sensitivity_data.append(row_data)
                            
                            sensitivity_df = pd.DataFrame(sensitivity_data)
                            st.dataframe(sensitivity_df, use_container_width=True)
                            
                            st.caption("Sensitivity table shows Fair Value per Share for different WACC and terminal growth rate combinations")
                        
                        tab_idx += 1
                        
                        # TAB: Scenario Analysis
                        with tabs[tab_idx]:
                            _base_ebitda_margin_u = (projections['ebitda'][0] / projections['revenue'][0] * 100) \
                                if projections['revenue'] and projections['revenue'][0] > 0 else drivers['avg_opex_margin']
                            _base_wc_pct_u = (projections['wc'][0] / projections['revenue'][0] * 100) \
                                if projections['revenue'] and projections['revenue'][0] > 0 else 0.0
                            _base_drivers_unlisted = {
                                'growth': drivers['avg_growth'],
                                'ebitda_margin': _base_ebitda_margin_u,
                                'wacc': wacc_details['wacc'],
                                'capex_ratio': drivers['avg_capex_ratio'],
                                'wc_pct': _base_wc_pct_u,
                            }
                            render_scenario_analysis_tab(
                                financials, wc_metrics, projection_years, tax_rate,
                                _base_drivers_unlisted, scenario_pct_unlisted, scenario_drivers_unlisted, scenario_case_unlisted,
                                terminal_growth, num_shares, cash_balance,
                                _ticker_csym, current_price=0, key_prefix="scenario_unlisted",
                                fixed_overrides={
                                    'depreciation_rate_override': depreciation_rate_override_unlisted if depreciation_rate_override_unlisted > 0 else None,
                                    'depreciation_method': depreciation_method_unlisted,
                                    'inventory_days_override': inventory_days_override_unlisted if inventory_days_override_unlisted > 0 else None,
                                    'debtor_days_override': debtor_days_override_unlisted if debtor_days_override_unlisted > 0 else None,
                                    'creditor_days_override': creditor_days_override_unlisted if creditor_days_override_unlisted > 0 else None,
                                    'interest_rate_override': interest_rate_override_unlisted if interest_rate_override_unlisted > 0 else None,
                                },
                                base_fair_value=valuation['fair_value_per_share']
                            )
                        
                        tab_idx += 1
                        
                        # TAB: Monte Carlo Simulation
                        with tabs[tab_idx]:
                            _base_ebitda_margin_mc_u = (projections['ebitda'][0] / projections['revenue'][0] * 100) \
                                if projections['revenue'] and projections['revenue'][0] > 0 else drivers['avg_opex_margin']
                            _base_wc_pct_mc_u = (projections['wc'][0] / projections['revenue'][0] * 100) \
                                if projections['revenue'] and projections['revenue'][0] > 0 else 0.0
                            _base_drivers_mc_unlisted = {
                                'growth': drivers['avg_growth'],
                                'ebitda_margin': _base_ebitda_margin_mc_u,
                                'wacc': wacc_details['wacc'],
                                'capex_ratio': drivers['avg_capex_ratio'],
                                'wc_pct': _base_wc_pct_mc_u,
                            }
                            render_monte_carlo_tab(
                                financials, wc_metrics, projection_years, tax_rate,
                                _base_drivers_mc_unlisted, terminal_growth, num_shares, cash_balance,
                                _ticker_csym, current_price=0, key_prefix="mc_unlisted",
                                fixed_overrides={
                                    'depreciation_rate_override': depreciation_rate_override_unlisted if depreciation_rate_override_unlisted > 0 else None,
                                    'depreciation_method': depreciation_method_unlisted,
                                    'inventory_days_override': inventory_days_override_unlisted if inventory_days_override_unlisted > 0 else None,
                                    'debtor_days_override': debtor_days_override_unlisted if debtor_days_override_unlisted > 0 else None,
                                    'creditor_days_override': creditor_days_override_unlisted if creditor_days_override_unlisted > 0 else None,
                                    'interest_rate_override': interest_rate_override_unlisted if interest_rate_override_unlisted > 0 else None,
                                },
                                base_fair_value=valuation['fair_value_per_share']
                            )
                        
                        tab_idx += 1
                    
                    # ============================================
                    # TAB: RIM VALUATION (if enabled)
                    # ============================================
                    if run_rim_unlisted:
                        with tabs[tab_idx]:
                            st.subheader("📚 Residual Income Model (RIM) Valuation")
                            
                            if 'rim_results_unlisted' in locals() and rim_results_unlisted:
                                st.markdown("### 📖 Residual Income Formula")
                                st.latex(r"Value = Book\;Value + \sum_{t=1}^{n} \frac{RI_t}{(1+r)^t} + \frac{RI_{n+1}}{(r-g)(1+r)^n}")
                                st.latex(r"RI_t = NI_t - (r \times BV_{t-1})")
                                
                                st.markdown("---")
                                st.markdown("### 📊 RIM Parameters Used")
                                
                                rim_params_df = pd.DataFrame({
                                    'Parameter': ['Cost of Equity (r)', 'Terminal Growth (g)', 'Initial Book Value', 'ROE', 'BV Growth Rate'],
                                    'Value': [
                                        f"{rim_results_unlisted.get('cost_of_equity', 0):.2f}%",
                                        f"{rim_results_unlisted.get('terminal_growth', 0):.2f}%",
                                        f"₹{rim_results_unlisted.get('current_book_value', 0)/100000:.2f}L",
                                        f"{rim_results_unlisted.get('roe', 0):.2f}%",
                                        f"{rim_results_unlisted.get('bv_growth', 0):.2f}%"
                                    ]
                                })
                                st.dataframe(rim_params_df, use_container_width=True, hide_index=True)
                                
                                st.markdown("---")
                                st.markdown("### 📈 Projected Residual Income")
                                
                                if 'projections' in rim_results_unlisted and rim_results_unlisted['projections']:
                                    # Extract data from list of projection dicts
                                    proj_list = rim_results_unlisted['projections']
                                    rim_proj_df = pd.DataFrame({
                                        'Year': [f"Year {p['year']}" for p in proj_list],
                                        'Net Income': [p['net_income']/100000 for p in proj_list],  # Convert to Lacs
                                        'Book Value': [p['book_value']/100000 for p in proj_list],  # Convert to Lacs
                                        'Equity Charge': [p['book_value']/100000 * rim_results_unlisted['cost_of_equity']/100 for p in proj_list],
                                        'Residual Income': [p['residual_income']/100000 for p in proj_list],  # Convert to Lacs
                                        'PV of RI': [p['pv_ri']/100000 for p in proj_list]  # Convert to Lacs
                                    })
                                    numeric_cols = rim_proj_df.select_dtypes(include=[np.number]).columns.tolist()
                                    format_dict = {col: '{:.2f}' for col in numeric_cols}
                                    st.dataframe(rim_proj_df.style.format(format_dict), use_container_width=True)
                                
                                st.markdown("---")
                                st.markdown("### 🎯 RIM Valuation Summary")
                                
                                rim_summary_df = pd.DataFrame({
                                    'Component': ['Current Book Value', 'PV of Projected RI', 'PV of Terminal RI', 
                                                'Total Equity Value', 'Number of Shares', 'Fair Value per Share'],
                                    'Value': [
                                        f"₹{rim_results_unlisted.get('current_book_value', 0)/100000:.2f}L",
                                        f"₹{rim_results_unlisted.get('sum_pv_ri', 0)/100000:.2f}L",
                                        f"₹{rim_results_unlisted.get('terminal_ri_pv', 0)/100000:.2f}L",
                                        f"₹{rim_results_unlisted.get('total_equity_value', 0)/100000:.2f}L",
                                        f"{num_shares:,}",
                                        f"₹{rim_results_unlisted.get('value_per_share', 0):.2f}"
                                    ]
                                })
                                st.dataframe(rim_summary_df, use_container_width=True, hide_index=True)
                                
                                st.success(f"### 🎯 Fair Value per Share (RIM): ₹{rim_results_unlisted.get('value_per_share', 0):.2f}")
                            else:
                                st.warning("RIM valuation could not be calculated. Please check input parameters.")
                        
                        tab_idx += 1
                    
                    # ============================================
                    # TAB: COMPARATIVE VALUATION (if enabled)
                    # ============================================
                    if run_comp_unlisted and peer_tickers:
                        with tabs[tab_idx]:
                            st.subheader("📊 Comparative Valuation (Peer Multiples)")
                            
                            if peer_tickers and peer_tickers.strip():
                                with st.spinner("Fetching comparable companies data..."):
                                    try:
                                        comp_results_unlisted = perform_comparative_valuation(
                                            target_ticker=None,  # Unlisted company
                                            comp_tickers_str=peer_tickers,
                                            target_financials=financials,
                                            target_shares=num_shares,
                                            exchange_suffix="NS",
                                            projections=projections if run_dcf_unlisted else None,
                                            use_screener_peers=False,
                                            valuation_date=valuation_date_unlisted
                                        )
                                        
                                        if comp_results_unlisted and 'comparables' in comp_results_unlisted and len(comp_results_unlisted['comparables']) > 0:
                                            # Show comparables table
                                            st.markdown("### Comparable Companies")
                                            comp_df = pd.DataFrame(comp_results_unlisted['comparables'])
                                            if not comp_df.empty:
                                                display_comp_df = comp_df[['ticker', 'name', 'price', 'pe', 'pb', 'ps', 'ev_ebitda', 'ev_sales']]
                                                st.dataframe(display_comp_df.style.format({
                                                    'price': '₹{:.2f}',
                                                    'pe': '{:.2f}x',
                                                    'pb': '{:.2f}x',
                                                    'ps': '{:.2f}x',
                                                    'ev_ebitda': '{:.2f}x',
                                                    'ev_sales': '{:.2f}x'
                                                }), use_container_width=True)
                                            
                                            # Show multiples statistics
                                            st.markdown("### Peer Multiples Statistics")
                                            for multiple, stats in comp_results_unlisted['multiples_stats'].items():
                                                with st.expander(f"📊 {multiple.upper()} - Avg: {stats['average']:.2f}x, Median: {stats['median']:.2f}x"):
                                                    st.write(f"**Range:** {stats['min']:.2f}x - {stats['max']:.2f}x")
                                                    st.write(f"**Std Dev:** {stats['std']:.2f}x")
                                                    st.write(f"**Peer Values:** {', '.join([f'{v:.2f}x' for v in stats['values']])}")
                                            
                                            # Show implied valuations
                                            st.markdown("### Implied Fair Values")
                                            
                                            all_avg_values = []
                                            all_median_values = []
                                            
                                            for method_key, val_data in comp_results_unlisted['valuations'].items():
                                                st.markdown(f"#### {val_data['method']}")
                                                
                                                col1, col2 = st.columns(2)
                                                
                                                with col1:
                                                    st.markdown("**Using Average Multiple:**")
                                                    st.write(val_data['formula_avg'])
                                                    st.metric("Fair Value (Avg)", f"₹{val_data['fair_value_avg']:.2f}")
                                                    all_avg_values.append(val_data['fair_value_avg'])
                                                
                                                with col2:
                                                    st.markdown("**Using Median Multiple:**")
                                                    st.write(val_data['formula_median'])
                                                    st.metric("Fair Value (Median)", f"₹{val_data['fair_value_median']:.2f}")
                                                    all_median_values.append(val_data['fair_value_median'])
                                                
                                                st.markdown("---")
                                            
                                            # Summary statistics
                                            if all_avg_values and all_median_values:
                                                st.markdown("### 📈 Comparative Valuation Summary")
                                                
                                                col1, col2, col3 = st.columns(3)
                                                
                                                with col1:
                                                    avg_comp = np.mean(all_avg_values)
                                                    median_comp = np.median(all_median_values)
                                                    st.metric("Average (All Methods)", f"₹{avg_comp:.2f}")
                                                    st.metric("Median (All Methods)", f"₹{median_comp:.2f}")
                                                
                                                with col2:
                                                    st.metric("Min Fair Value", f"₹{min(all_avg_values + all_median_values):.2f}")
                                                    st.metric("Max Fair Value", f"₹{max(all_avg_values + all_median_values):.2f}")
                                                
                                                with col3:
                                                    if run_dcf_unlisted and valuation['fair_value_per_share'] > 0:
                                                        st.metric("DCF Fair Value", f"₹{valuation['fair_value_per_share']:.2f}")
                                                        combined_avg = (avg_comp + valuation['fair_value_per_share']) / 2
                                                        st.metric("DCF + Comp Avg", f"₹{combined_avg:.2f}")
                                        else:
                                            st.warning("Could not fetch comparable companies data or no valid comparables found")
                                    except Exception as e:
                                        st.error(f"Comparative valuation error: {str(e)}")
                                        import traceback
                                        st.error(traceback.format_exc())
                            else:
                                st.info("No peer tickers provided. Enter peer tickers to enable comparative valuation.")
                        
                            st.markdown("---")
                            render_implied_multiples_panel(
                                valuation, financials, num_shares,
                                comp_results=comp_results_unlisted if 'comp_results_unlisted' in locals() else None,
                                current_price=0,
                                key_prefix="implied_unlisted"
                            )
                        
                        tab_idx += 1
                    
                    # ============================================
                    # TAB: FOOTBALL FIELD
                    # ============================================
                    with tabs[tab_idx]:
                        st.subheader("🏈 Football Field")
                        st.caption("Compares the fair-value range from every valuation method run above, side by side.")
                        band_pct_unlisted = st.number_input(
                            "± Band for single-point methods (DCF / RIM) (%):",
                            min_value=1.0, max_value=50.0, value=12.0, step=1.0,
                            key="football_band_unlisted",
                            help="DCF and RIM produce one number, not a range — this sets how wide a band to "
                                 "draw around that number on the chart. Comparable Valuation uses the actual "
                                 "low/high peer multiples instead, so this doesn't affect it."
                        )
                        _ff_methods_unlisted = []
                        if run_dcf_unlisted and 'valuation' in locals() and valuation and valuation.get('fair_value_per_share', 0) > 0:
                            _fv = valuation['fair_value_per_share']
                            _ff_methods_unlisted.append({'method': 'DCF', 'low': _fv * (1 - band_pct_unlisted / 100), 'high': _fv * (1 + band_pct_unlisted / 100), 'mid': _fv})
                        if run_rim_unlisted and 'rim_results_unlisted' in locals() and rim_results_unlisted and rim_results_unlisted.get('value_per_share', 0) > 0:
                            _fv = rim_results_unlisted['value_per_share']
                            _ff_methods_unlisted.append({'method': 'RIM', 'low': _fv * (1 - band_pct_unlisted / 100), 'high': _fv * (1 + band_pct_unlisted / 100), 'mid': _fv})
                        if 'comp_results_unlisted' in locals() and comp_results_unlisted:
                            for _key, _val_data in comp_results_unlisted.get('valuations', {}).items():
                                _stats = comp_results_unlisted.get('multiples_stats', {}).get(_key)
                                if _stats and _val_data.get('target_metric', 0) > 0 and _val_data.get('fair_value_avg', 0) > 0:
                                    _low = _val_data['target_metric'] * _stats.get('min', 0)
                                    _high = _val_data['target_metric'] * _stats.get('max', 0)
                                    if _high > 0:
                                        _ff_methods_unlisted.append({'method': _val_data.get('method', _key.upper()), 'low': _low, 'high': _high, 'mid': _val_data['fair_value_avg']})
                        render_football_field_chart(_ff_methods_unlisted, 0, _ticker_csym if '_ticker_csym' in locals() else '₹', key_prefix="football_unlisted")
                    
                    tab_idx += 1
                    
                    # ============================================
                    # TAB: IBC LIQUIDATION VALUE
                    # ============================================
                    with tabs[tab_idx]:
                        render_ibc_liquidation_tab(
                            financials, _ticker_csym if '_ticker_csym' in locals() else '₹', key_prefix="ibc_unlisted",
                            dcf_equity_value=(valuation.get('equity_value', 0) * 100000 if run_dcf_unlisted and 'valuation' in locals() and valuation else None),
                            dcf_fair_value_per_share=(valuation.get('fair_value_per_share') if run_dcf_unlisted and 'valuation' in locals() and valuation else None),
                            num_shares=num_shares
                        )
                    
                    tab_idx += 1
                    
                    # ============================================
                    # TAB: RATIO ANALYSIS
                    # ============================================
                    with tabs[tab_idx]:
                        render_ratio_dashboard(
                            financials, wc_metrics, _ticker_csym if '_ticker_csym' in locals() else '₹', key_prefix="ratios_unlisted",
                            current_price=0, shares=num_shares, tax_rate=tax_rate
                        )
                    
                    tab_idx += 1
                    
                    # ============================================
                    # TAB: HEALTH CHECK
                    # ============================================
                    with tabs[tab_idx]:
                        render_valuation_health_check(
                            valuation if run_dcf_unlisted and 'valuation' in locals() else None,
                            terminal_growth, wacc_details['wacc'] if 'wacc_details' in locals() else 0,
                            csym=_ticker_csym if '_ticker_csym' in locals() else '₹', key_prefix="health_unlisted",
                            ddm_result=None,
                            rim_result=rim_results_unlisted if 'rim_results_unlisted' in locals() else None,
                            comp_results=comp_results_unlisted if 'comp_results_unlisted' in locals() else None
                        )
                    
                    tab_idx += 1
                    
                    # ============================================
                    # TAB: ALTMAN Z-SCORE
                    # ============================================
                    with tabs[tab_idx]:
                        render_altman_zscore_tab(
                            financials, _ticker_csym if '_ticker_csym' in locals() else '₹', key_prefix="zscore_unlisted",
                            current_price=0, shares=num_shares
                        )
                    
                    tab_idx += 1
                    
                    # ============================================
                    # TAB: LBO ANALYSIS
                    # ============================================
                    with tabs[tab_idx]:
                        try:
                            render_lbo_analysis_tab(
                                financials, _ticker_csym if '_ticker_csym' in locals() else '₹', key_prefix="lbo_unlisted", tax_rate=tax_rate
                            )
                        except Exception as _lbo_err:
                            st.error(f"LBO Analysis failed to render: {_lbo_err}")
                            st.code(traceback.format_exc())
                    
                    tab_idx += 1



    elif mode == "Screener Excel Mode (Screener.in Template)":

        # ===== VALUATION DATE — controls what "today" means for Rf, Rm, and =====
        # ===== the current price / multiples used in comparative valuation =====
        st.markdown("### 📅 Valuation Date")
        col_vds1, col_vds2 = st.columns([2, 3])
        with col_vds1:
            valuation_date_screener = st.date_input(
                "Valuation Date",
                value=st.session_state.get('valuation_date_screener', datetime.now().date()),
                key='valuation_date_screener',
                help="Leave as today for a live valuation. Set an earlier date to price the company, its peers, Rf and Rm as of that date instead of today's market."
            )
        with col_vds2:
            st.write("")
            if is_valuation_date_today(valuation_date_screener):
                st.caption("💡 Using **live/current** market data (today).")
            else:
                st.info(f"📌 Valuing **as of {valuation_date_screener.strftime('%d-%b-%Y')}** — current price, peer multiples, Rf and Rm will use data as of this date (not today).")
        st.markdown("---")
        
        # ===== RISK-FREE RATE TICKER - AT THE TOP, ALWAYS VISIBLE =====
        st.markdown("### 🏛️ Risk-Free Rate Configuration")
        st.info("💡 **Default: 6.83%**. Enter any Yahoo Finance ticker and click Fetch. Common options: ^TNX (US 10Y), ^IRX (US 3-month), or any stock/index for custom rates.")
        
        # Initialize if not exists
        if 'cached_rf_rate_screener' not in st.session_state:
            st.session_state.cached_rf_rate_screener = 6.83
        
        rf_col1, rf_col2, rf_col3 = st.columns([3, 2, 1])
        with rf_col1:
            custom_rf_ticker_screener = st.text_input(
                "Yahoo Finance Ticker for Risk-Free Rate",
                value="NIFTYGS10YR.NS",
                key='custom_rf_ticker_screener_top',
                help="Enter any ticker - will calculate CAGR for stocks/indices, or use value directly for yields. Examples: NIFTYGS10YR.NS, ^TNX, RELIANCE.NS"
            )
        with rf_col2:
            st.metric("Current RF Rate", f"{st.session_state.cached_rf_rate_screener:.2f}%")
        with rf_col3:
            st.write("")
            st.write("")
            
            # Add a counter to track button clicks
            if 'rf_fetch_click_count_screener' not in st.session_state:
                st.session_state.rf_fetch_click_count_screener = 0
            
            if st.button("🔄 Fetch", key='refresh_rf_screener_top'):
                # Increment click counter FIRST
                st.session_state.rf_fetch_click_count_screener += 1
                
                # Store debug output in session state so it persists across reruns
                debug_output = []
                debug_output.append(f"🔄 **FETCH BUTTON CLICKED #{st.session_state.rf_fetch_click_count_screener} - SCREENER MODE**")
                debug_output.append(f"📝 Input ticker: `{custom_rf_ticker_screener}`")
                
                ticker_to_use = custom_rf_ticker_screener.strip() if custom_rf_ticker_screener.strip() else None
                debug_output.append(f"📝 Ticker to use (after strip): `{ticker_to_use}`")
                
                debug_output.append("⏳ Calling get_risk_free_rate()...")
                _as_of = None if is_valuation_date_today(valuation_date_screener) else valuation_date_screener
                fetched_rate, fetch_debug = get_risk_free_rate(ticker_to_use, as_of_date=_as_of)
                
                # Add all debug messages from the function
                debug_output.extend(fetch_debug)
                
                debug_output.append(f"✅ Function returned: {fetched_rate}%")
                
                debug_output.append(f"💾 Updating session state...")
                debug_output.append(f"   - Before: {st.session_state.get('cached_rf_rate_screener', 'NOT SET')}")
                st.session_state.cached_rf_rate_screener = fetched_rate
                st.session_state.manual_rf_screener = fetched_rate  # Update widget state
                debug_output.append(f"   - After: {st.session_state.cached_rf_rate_screener}")
                debug_output.append(f"🔄 Widget state updated to: {fetched_rate}%")
                
                # Store debug output in session state
                st.session_state.rf_fetch_debug_screener = debug_output
                st.session_state.rf_fetch_success_screener = True
                
                # Rerun to update UI
                st.rerun()
        
        st.markdown("---")
        # ===== END RF RATE CONFIG =====
        
        # ===== MARKET RETURN TICKER =====
        st.markdown("### 📈 Market Return Configuration")
        st.info("💡 Use `%5E` prefix (e.g., `%5EBSESN` for Sensex). **Default**: 12.0%")
        
        if 'cached_rm_rate_screener' not in st.session_state:
            st.session_state.cached_rm_rate_screener = 12.0
        
        rm_col1, rm_col2, rm_col3 = st.columns([3, 2, 1])
        with rm_col1:
            custom_rm_ticker_screener = st.text_input(
                "Yahoo Finance Ticker for Market Return",
                value="%5EBSESN",
                key='custom_rm_ticker_screener_top',
                help="Use %5E prefix. Examples: %5EBSESN, %5ENSEI"
            )
        with rm_col2:
            st.metric("Current Market Return", f"{st.session_state.cached_rm_rate_screener:.2f}%")
        with rm_col3:
            st.write("")
            st.write("")
            
            if 'rm_fetch_click_count_screener' not in st.session_state:
                st.session_state.rm_fetch_click_count_screener = 0
            
            if st.button("🔄 Fetch", key='refresh_rm_screener_top'):
                st.session_state.rm_fetch_click_count_screener += 1
                
                debug_output = []
                debug_output.append(f"🔄 **FETCH #{st.session_state.rm_fetch_click_count_screener} - SCREENER MODE**")
                
                ticker_to_use = custom_rm_ticker_screener.strip() if custom_rm_ticker_screener.strip() else None
                _as_of = None if is_valuation_date_today(valuation_date_screener) else valuation_date_screener
                fetched_rate, fetch_debug = get_market_return(ticker_to_use, as_of_date=_as_of)
                
                debug_output.extend(fetch_debug)
                debug_output.append(f"✅ Returned: {fetched_rate}%")
                
                st.session_state.cached_rm_rate_screener = fetched_rate
                st.session_state.manual_rm_screener = fetched_rate  # Update widget state
                
                st.session_state.rm_fetch_debug_screener = debug_output
                st.rerun()
        
        st.markdown("---")
        
        # Add template download button
        st.markdown("#### 📥 Download Screener Template")
        st.caption("Download the pre-configured template that matches Screener.in format")
        
        try:
            with open('Screener_template.xlsx', 'rb') as f:
                template_bytes = f.read()
            
            st.download_button(
                label="📥 Download Screener Template",
                data=template_bytes,
                file_name="Screener_Template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Download this template matching Screener.in export format"
            )
        except FileNotFoundError:
            st.warning("⚠️ Screener_template.xlsx not found in the current directory")
        
        if not SCREENER_MODE_AVAILABLE:
            st.error(f"❌ Screener Mode module not available: {SCREENER_MODE_ERROR}")
            st.info("Please ensure screener_excel_mode.py is in the same directory as this file.")
            st.stop()
        
        st.markdown("---")
        
        col1, col2 = st.columns(2)
        
        with col1:
            company_name_screener = st.text_input("Company Name:", key='screener_company_name')
            
            # NEW: Ticker input for Yahoo Finance data (current price + beta)
            col_ticker1, col_ticker2 = st.columns([3, 1])
            with col_ticker2:
                exchange_screener = st.selectbox(
                    "Exchange",
                    ["NS", "BO"],
                    key='screener_exchange',
                    help="NS=NSE, BO=BSE"
                )
            with col_ticker1:
                _exch_label_screener = "NSE" if exchange_screener == "NS" else "BSE"
                manual_price_screener = None
                manual_beta_screener = None
                # Clear a stale company selection if the exchange was just switched
                if st.session_state.get('_screener_prev_exchange') != exchange_screener:
                    st.session_state['screener_company_select'] = ""
                    st.session_state['_screener_prev_exchange'] = exchange_screener

                if TICKER_DROPDOWN_AVAILABLE:
                    _entry_mode_screener = st.radio(
                        f"How do you want to select the {_exch_label_screener} company?",
                        options=["🔍 Search by name", "⌨️ Enter ticker/code directly", "📝 Not listed on NSE/BSE — enter price manually"],
                        key='screener_company_entry_mode',
                        horizontal=True
                    )
                    if _entry_mode_screener == "🔍 Search by name":
                        _company_opts_screener = get_company_display_options(_exch_label_screener)
                        if _company_opts_screener:
                            _selected_company_display_screener = st.selectbox(
                                f"{_exch_label_screener} Company (search/select by name, for current price & beta):",
                                options=[""] + _company_opts_screener,
                                key='screener_company_select',
                                help="Start typing the company name to filter the list."
                            )
                            ticker_symbol_screener = resolve_company_display(_selected_company_display_screener, _exch_label_screener)
                        else:
                            _err = get_load_error(_exch_label_screener)
                            st.warning(f"⚠️ {_exch_label_screener} company list unavailable — enter ticker manually below.{(' Details: ' + _err) if _err else ''}")
                            ticker_symbol_screener = st.text_input(
                                "Stock Ticker Symbol (for current price & beta):",
                                placeholder="e.g., RELIANCE, TCS, HDFCBANK",
                                key='screener_ticker_manual',
                                help="Enter stock ticker to fetch current market price and beta from Yahoo Finance"
                            )
                    elif _entry_mode_screener == "⌨️ Enter ticker/code directly":
                        ticker_symbol_screener = st.text_input(
                            "Stock Ticker Symbol (for current price & beta):",
                            placeholder="e.g., RELIANCE, TCS, HDFCBANK",
                            key='screener_ticker_manual',
                            help="Enter stock ticker to fetch current market price and beta from Yahoo Finance"
                        )
                    else:
                        # Company not yet listed on NSE/BSE (e.g. Screener.in
                        # profile pages of the form /company/id/<id>/ for
                        # pre-IPO / unlisted companies) — no ticker exists to
                        # fetch a price or beta from, so take both by hand.
                        ticker_symbol_screener = ""
                        st.caption("💡 For companies not yet trading on NSE/BSE (e.g. Screener.in company/id/ pages), enter the price and beta yourself below.")
                        _mcol1, _mcol2 = st.columns(2)
                        with _mcol1:
                            manual_price_screener = st.number_input(
                                "Manual Current Price (₹):",
                                min_value=0.0, value=0.0, step=0.01,
                                key='screener_manual_price',
                                help="Leave at 0 if a current price isn't available/applicable (e.g. no recent trade)."
                            )
                        with _mcol2:
                            manual_beta_screener = st.number_input(
                                "Manual Beta (optional):",
                                min_value=0.0, value=1.0, step=0.05,
                                key='screener_manual_beta',
                                help="Used for WACC/CAPM if provided. Defaults to 1.0 (market average) if unsure."
                            )
                else:
                    _entry_mode_screener = st.radio(
                        f"How do you want to provide {_exch_label_screener} market data?",
                        options=["⌨️ Enter ticker/code directly", "📝 Not listed on NSE/BSE — enter price manually"],
                        key='screener_company_entry_mode_nodropdown',
                        horizontal=True
                    )
                    if _entry_mode_screener == "⌨️ Enter ticker/code directly":
                        ticker_symbol_screener = st.text_input(
                            "Stock Ticker Symbol (for current price & beta):",
                            placeholder="e.g., RELIANCE, TCS, HDFCBANK",
                            key='screener_ticker',
                            help="Enter stock ticker to fetch current market price and beta from Yahoo Finance"
                        )
                    else:
                        ticker_symbol_screener = ""
                        st.caption("💡 For companies not yet trading on NSE/BSE (e.g. Screener.in company/id/ pages), enter the price and beta yourself below.")
                        _mcol1, _mcol2 = st.columns(2)
                        with _mcol1:
                            manual_price_screener = st.number_input(
                                "Manual Current Price (₹):",
                                min_value=0.0, value=0.0, step=0.01,
                                key='screener_manual_price',
                                help="Leave at 0 if a current price isn't available/applicable (e.g. no recent trade)."
                            )
                        with _mcol2:
                            manual_beta_screener = st.number_input(
                                "Manual Beta (optional):",
                                min_value=0.0, value=1.0, step=0.05,
                                key='screener_manual_beta',
                                help="Used for WACC/CAPM if provided. Defaults to 1.0 (market average) if unsure."
                            )
            
            # Shares Outstanding Source
            st.markdown("**📊 Shares Outstanding**")
            use_yahoo_shares_screener = st.checkbox(
                "Fetch shares outstanding from Yahoo Finance (use if Screener data is inaccurate)",
                value=False,
                key='use_yahoo_shares_screener',
                help="Requires a valid ticker above. If unchecked, shares are read from the uploaded Screener Excel."
            )
            manual_shares_override_screener = st.number_input(
                "Manual Override — leave at 0 to use auto",
                min_value=0,
                value=0,
                step=1000000,
                format="%d",
                key='manual_shares_screener',
                help="Takes priority over both Excel and Yahoo Finance. Enter absolute number (e.g., 50000000 for 5 crore)."
            )
            if manual_shares_override_screener > 0:
                st.success(f"✅ Manual override: **{manual_shares_override_screener:,}** shares ({manual_shares_override_screener/10000000:.2f} Crore)")

            st.markdown("---")

            # Integrated Auto Download or Manual Upload Section
            if AUTO_DOWNLOAD_AVAILABLE:
                use_auto_dl, excel_file_screener = integrate_with_existing_upload_section(
                    cookies_path="screener_cookies.pkl"
                )
            else:
                # Fallback to manual upload only
                st.markdown("---")
                st.markdown("### 📊 Screener Excel Data Source")
                excel_file_screener = st.file_uploader(
                    "Upload Screener Excel Template", 
                    type=['xlsx', 'xls'],
                    key='screener_excel_upload',
                    help="Upload the Excel export straight from Screener.in — no need to clean or reshape it first, that happens automatically. A pre-cleaned 'Balance Sheet'/'Profit and Loss Account' template also works."
                )
                if not AUTO_DOWNLOAD_AVAILABLE:
                    st.info("💡 Auto-download feature not available. Please upload Excel manually.")

            
            st.markdown("---")
            st.markdown("**📊 Peer Companies (Both Exchanges)**")
            
            # Auto-fetch button ABOVE the input boxes (so we can set state before inputs are created)
            if ticker_symbol_screener and ticker_symbol_screener.strip():
                col_btn1, col_btn2, col_btn3 = st.columns([1, 1, 2])
                with col_btn2:
                    _max_peers_screener = st.number_input(
                        "Max to fetch:", min_value=1, value=20, step=1,
                        key='screener_max_peers_fetch',
                        help="How many peers to try to fetch (no upper limit). Yahoo's own APIs may still return fewer if that's all they have."
                    )
                with col_btn1:
                    if st.button("🔍 Auto-Fetch Peers", key='screener_auto_fetch_peers', help="Automatically fetch industry peer companies", use_container_width=True):
                        with st.spinner("Fetching peers from Yahoo Finance..."):
                            try:
                                # Construct full ticker with exchange
                                full_ticker = f"{ticker_symbol_screener.strip().upper()}.{exchange_screener}"
                                st.info(f"🔍 Searching peers for {full_ticker}...")
                                
                                if PEER_FETCHER_AVAILABLE:
                                    # Fetch peers
                                    peers_list = get_industry_peers(full_ticker, max_peers=int(_max_peers_screener), exclude_self=True)
                                    
                                    if peers_list:
                                        # Map auto-fetched bare tickers/codes to display names
                                        # and pre-fill that many peer-selection dropdowns.
                                        _peer_exch_screener = "NSE" if exchange_screener == "NS" else "BSE"
                                        if TICKER_DROPDOWN_AVAILABLE:
                                            _auto_displays_screener = [bare_ticker_to_peer_display(p, _peer_exch_screener) for p in peers_list]
                                        else:
                                            _auto_displays_screener = [f"{p} ({_peer_exch_screener}: {p})" for p in peers_list]
                                        st.session_state['screener_num_peers'] = len(_auto_displays_screener)
                                        for _i, _d in enumerate(_auto_displays_screener):
                                            st.session_state[f'screener_peer_select_{_i}'] = _d

                                        # Force rerun to show peers immediately
                                        st.rerun()
                                    else:
                                        st.warning("⚠️ No peers found. Please enter manually.")
                                else:
                                    st.error("❌ Peer fetcher module not available (PEER_FETCHER_AVAILABLE = False). Please enter peers manually.")
                            except Exception as e:
                                st.error(f"❌ Error fetching peers: {str(e)}")
                                import traceback
                                st.code(traceback.format_exc())
                with col_btn2:
                    st.caption("💡 Click to automatically fetch industry peers from Yahoo Finance")
            
            if TICKER_DROPDOWN_AVAILABLE:
                # Choose how many peers, then pick each — by name (dropdown)
                # or by typing the ticker/code directly.
                if 'screener_num_peers' not in st.session_state:
                    st.session_state['screener_num_peers'] = 0

                num_peers_screener = st.number_input(
                    "Number of Peers:",
                    min_value=0,
                    step=1,
                    key='screener_num_peers',
                    help="How many peer companies to include (no upper limit). Use Auto-Fetch above to populate automatically, or set this and pick each peer by name/ticker below."
                )

                _peer_entry_mode_screener = st.radio(
                    "How do you want to select peers?",
                    options=["🔍 Search by name", "⌨️ Enter ticker/code directly"],
                    key='screener_peer_entry_mode',
                    horizontal=True
                ) if int(num_peers_screener) > 0 else "🔍 Search by name"

                _peer_options_screener = get_peer_display_options() if _peer_entry_mode_screener == "🔍 Search by name" else []
                _selected_peer_displays_screener = []
                _manual_peer_tickers_screener = []
                for _i in range(int(num_peers_screener)):
                    if _peer_entry_mode_screener == "🔍 Search by name":
                        _pk = f'screener_peer_select_{_i}'
                        if _pk not in st.session_state:
                            st.session_state[_pk] = ""
                        _sel = st.selectbox(
                            f"Peer {_i + 1}:",
                            options=[""] + _peer_options_screener,
                            key=_pk,
                            help="Start typing the company name to filter the list."
                        )
                        _selected_peer_displays_screener.append(_sel)
                    else:
                        _pkm = f'screener_peer_manual_{_i}'
                        _tkr = st.text_input(
                            f"Peer {_i + 1} ticker/code:",
                            placeholder="e.g., TCS or TCS.NS",
                            key=_pkm,
                            help="Bare ticker/code is assumed to be on the same exchange selected above; a full Yahoo ticker (e.g. TCS.NS) also works."
                        )
                        if _tkr.strip():
                            _t = _tkr.strip().upper()
                            if '.' not in _t:
                                _t = f"{_t}.{exchange_screener}"
                            _manual_peer_tickers_screener.append(_t)

                peer_tickers_screener = ",".join(
                    _manual_peer_tickers_screener if _peer_entry_mode_screener != "🔍 Search by name"
                    else [resolve_peer_display_to_full_ticker(d) for d in _selected_peer_displays_screener if d]
                )
            else:
                # Fallback if the dropdown utility module failed to load
                nse_peers_screener = st.text_input(
                    "NSE Peer Tickers (comma-separated):",
                    placeholder="e.g., RELIANCE, TCS, INFY",
                    key='nse_peers_input',
                    help="Enter manually — name-based dropdown unavailable."
                )
                bse_peers_screener = st.text_input(
                    "BSE Peer Tickers (comma-separated):",
                    placeholder="e.g., RELIANCE, TCS, INFY",
                    key='bse_peers_input',
                    help="Enter BSE-listed peer companies for beta calculation."
                )
                peer_tickers_screener = ""
                if nse_peers_screener.strip():
                    nse_list = [f"{t.strip()}.NS" if '.NS' not in t and '.BO' not in t else t.strip() 
                               for t in nse_peers_screener.split(',') if t.strip()]
                    peer_tickers_screener = ",".join(nse_list)
                if bse_peers_screener.strip():
                    bse_list = [f"{t.strip()}.BO" if '.NS' not in t and '.BO' not in t else t.strip() 
                               for t in bse_peers_screener.split(',') if t.strip()]
                    if peer_tickers_screener:
                        peer_tickers_screener += "," + ",".join(bse_list)
                    else:
                        peer_tickers_screener = ",".join(bse_list)
            
            st.markdown("---")
            st.markdown("**🎯 Data & Projection Configuration**")
            
            # Historical Years and Projection Years - PROMINENTLY PLACED
            col_years1, col_years2 = st.columns(2)
            with col_years1:
                historical_years_screener = st.number_input(
                    "📊 Historical Years to Use",
                    min_value=2,
                    max_value=15,
                    value=5,
                    step=1,
                    key='screener_historical_years',
                    help="Select how many years of historical data to use from your Excel file (e.g., if Excel has 10 years, you can use 3, 5, or all 10)"
                )
            with col_years2:
                projection_years_screener = st.number_input(
                    "📈 Projection Years",
                    min_value=3,
                    max_value=15,
                    value=5,
                    step=1,
                    key='screener_proj_years',
                    help="Number of years to project into the future"
                )
        
        with col2:
            tax_rate_screener = st.number_input(
                "Tax Rate (%):", 
                min_value=0.0, 
                max_value=100.0, 
                value=25.0, 
                step=0.5,
                key='screener_tax'
            )
            terminal_growth_screener = st.number_input(
                "Terminal Growth Rate (%):", 
                min_value=0.0, 
                max_value=20.0, 
                value=4.0, 
                step=0.5,
                key='screener_terminal',
                help="Long-term perpetual growth rate. Should be <= long-term GDP growth + inflation (typically 4-8% for India)"
            )
            
            # Risk-free rate override
            st.markdown("**🏛️ Risk-Free Rate (G-Sec 10Y)**")
            
            if 'manual_rf_screener' not in st.session_state:
                st.session_state.manual_rf_screener = st.session_state.get('cached_rf_rate_screener', 6.83)
            
            manual_rf_rate_screener = st.number_input(
                f"Risk-Free Rate (%)",
                min_value=0.0,
                max_value=20.0,
                value=st.session_state.manual_rf_screener,
                step=0.1,
                key='manual_rf_screener',
                help="Auto-fetched from ticker above. You can manually edit this value."
            )
            
            # Market Return manual input
            st.markdown("**📈 Market Return (Rm)**")
            
            if 'manual_rm_screener' not in st.session_state:
                st.session_state.manual_rm_screener = st.session_state.get('cached_rm_rate_screener', 12.0)
            
            manual_rm_rate_screener = st.number_input(
                f"Market Return (%)",
                min_value=0.0,
                max_value=50.0,
                value=st.session_state.manual_rm_screener,
                step=0.1,
                key='manual_rm_screener',
                help="Auto-fetched from ticker above. You can manually edit this value."
            )
            
            # Manual discount rate override
            st.markdown("---")
            st.markdown("**💰 Discount Rate Override (Optional)**")
            manual_discount_rate_screener = st.number_input(
                "Manual Discount Rate Override (%):",
                min_value=0.0,
                max_value=50.0,
                value=0.0,
                step=0.5,
                key='manual_discount_screener',
                help="⚠️ Override WACC calculation. Leave at 0 to use auto-calculated WACC. Use this if you want to use a specific discount rate instead of WACC."
            )
            if manual_discount_rate_screener > 0:
                st.info(f"💡 Using manual discount rate: {manual_discount_rate_screener:.2f}% (Overriding WACC)")
            
            # ── Beta Calculation Period ────────────────────────────────────
            st.markdown("**📅 Beta Calculation Period**")
            st.caption("Daily returns only. Defaults to 3 years ending today if left unchanged.")
            _beta_default_end_screener   = datetime.now().date()
            _beta_default_start_screener = (_beta_default_end_screener.replace(year=_beta_default_end_screener.year - 3))
            _sc1, _sc2 = st.columns(2)
            with _sc1:
                beta_start_date_screener = st.date_input(
                    "From", value=_beta_default_start_screener,
                    key='beta_start_screener',
                    help="Start date for beta regression window"
                )
            with _sc2:
                beta_end_date_screener = st.date_input(
                    "To", value=_beta_default_end_screener,
                    key='beta_end_screener',
                    help="End date for beta regression window"
                )
            if beta_start_date_screener >= beta_end_date_screener:
                st.error("⚠️ Beta start date must be before end date.")
            # ── End Beta Period ────────────────────────────────────────────

            # Valuation models to run
            st.markdown("**🎯 Valuation Models**")
            run_dcf_screener = st.checkbox("DCF (FCFF)", value=True, key='screener_dcf')
            run_ddm_screener = st.checkbox("DDM (Dividend Discount)", value=True, key='screener_ddm')
            run_rim_screener = st.checkbox("RIM (Residual Income)", value=True, key='screener_rim')
            run_comp_screener = st.checkbox("Comparative Valuation", value=True, key='screener_comp')
        
        with st.expander("⚙️ Advanced Projection Assumptions - FULL CONTROL"):
            st.info("💡 **Complete Control:** Override ANY projection parameter below. Leave at 0 or blank for auto-calculation from historical data.")
            
            st.markdown("### 📊 Revenue & Growth")
            col1, col2, col3 = st.columns(3)
            with col1:
                rev_growth_override_screener = st.number_input(
                    "Revenue Growth (%/year)", 
                    min_value=0.0, max_value=200.0, value=0.0, step=0.5,
                    key='screener_rev_growth',
                    help="0 = Auto from historical CAGR. Override to use custom growth rate."
                )
            with col2:
                opex_margin_override_screener = st.number_input(
                    "Operating Expense Margin (%)", 
                    min_value=0.0, max_value=100.0, value=0.0, step=0.5,
                    key='screener_opex_margin',
                    help="0 = Auto from historical average"
                )
            with col3:
                ebitda_margin_override_screener = st.number_input(
                    "EBITDA Margin (%)", 
                    min_value=0.0, max_value=100.0, value=0.0, step=0.5,
                    key='screener_ebitda',
                    help="0 = Calculated as Revenue - OpEx"
                )

            # ── PER-YEAR PARAMETERS (Screener) ────────────────────────────
            st.markdown("### 📅 Year-by-Year Revenue Growth & EBITDA Margin")
            st.caption("Set individual values per projected year. Leave at 0 to fall back to the single overrides above (or auto-calc from history).")
            rev_growth_per_year_screener = []
            ebitda_margin_per_year_screener = []
            _proj_yrs_scr = int(st.session_state.get('screener_proj_years', 5))
            yr_cols_screener = st.columns(min(_proj_yrs_scr, 5))
            for _yi in range(_proj_yrs_scr):
                _col = yr_cols_screener[_yi % len(yr_cols_screener)]
                with _col:
                    st.markdown(f"**Year {_yi + 1}**")
                    _rg = st.number_input(
                        f"Rev Growth %",
                        min_value=0.0, max_value=200.0, value=0.0, step=0.5,
                        key=f'screener_rg_yr{_yi+1}',
                        help="0 = use global override / auto"
                    )
                    _em = st.number_input(
                        f"EBITDA Margin %",
                        min_value=0.0, max_value=100.0, value=0.0, step=0.5,
                        key=f'screener_em_yr{_yi+1}',
                        help="0 = use global override / auto"
                    )
                    rev_growth_per_year_screener.append(_rg)
                    ebitda_margin_per_year_screener.append(_em)
            # ── END PER-YEAR ───────────────────────────────────────────────
            
            # ── MULTI-STAGE GROWTH PROFILE (optional, overrides the above) ──
            _ms_enabled_screener, _ms_growth_screener, _ms_margin_screener = render_multistage_growth_ui(
                key_prefix="ms_screener",
                total_years=_proj_yrs_scr,
                default_stage1_growth=rev_growth_override_screener if rev_growth_override_screener > 0 else 15.0,
                terminal_growth=terminal_growth_screener,
                default_stage1_margin=ebitda_margin_override_screener if ebitda_margin_override_screener > 0 else 20.0
            )
            if _ms_enabled_screener:
                rev_growth_per_year_screener = _ms_growth_screener
                if _ms_margin_screener:
                    ebitda_margin_per_year_screener = _ms_margin_screener
            # ── END MULTI-STAGE ──────────────────────────────────────────────
            
            st.markdown("### 🏗️ CapEx & Depreciation")
            col4, col5, col6 = st.columns(3)
            with col4:
                capex_ratio_override_screener = st.number_input(
                    "CapEx/Revenue (%)", 
                    min_value=0.0, max_value=200.0, value=0.0, step=0.5,
                    key='screener_capex_ratio',
                    help="0 = Auto from historical average. Override for custom CapEx assumptions."
                )
            with col5:
                depreciation_rate_override_screener = st.number_input(
                    "Depreciation Rate (%)", 
                    min_value=0.0, max_value=30.0, value=0.0, step=0.5,
                    key='screener_dep_rate',
                    help="0 = Auto calculated"
                )
            with col6:
                depreciation_method_screener = st.selectbox(
                    "Depreciation Method",
                    ["Auto", "% of Fixed Assets", "% of Revenue", "Absolute Value"],
                    key='screener_dep_method'
                )
            
            st.markdown("### 💰 Working Capital Management")
            col7, col8, col9 = st.columns(3)
            with col7:
                inventory_days_override_screener = st.number_input(
                    "Inventory Days", 
                    min_value=0.0, max_value=365.0, value=0.0, step=1.0,
                    key='screener_inv_days',
                    help="0 = Auto from historical average"
                )
            with col8:
                debtor_days_override_screener = st.number_input(
                    "Debtor/Receivables Days", 
                    min_value=0.0, max_value=365.0, value=0.0, step=1.0,
                    key='screener_deb_days',
                    help="0 = Auto from historical average"
                )
            with col9:
                creditor_days_override_screener = st.number_input(
                    "Creditor/Payables Days", 
                    min_value=0.0, max_value=365.0, value=0.0, step=1.0,
                    key='screener_cred_days',
                    help="0 = Auto from historical average"
                )
            
            st.markdown("### 📈 Tax & Interest")
            col10, col11, col12 = st.columns(3)
            with col10:
                interest_rate_override_screener = st.number_input(
                    "Interest Rate (%)", 
                    min_value=0.0, max_value=30.0, value=0.0, step=0.25,
                    key='screener_interest',
                    help="0 = Auto calculated from Debt"
                )
            with col11:
                working_capital_as_pct_revenue_screener = st.number_input(
                    "Working Capital % of Revenue", 
                    min_value=0.0, max_value=50.0, value=0.0, step=0.5,
                    key='screener_wc_pct',
                    help="0 = Calculate from Inv+Deb-Cred days"
                )
        
        # ================================
        # DDM & RIM PARAMETERS FOR SCREENER MODE
        # ================================
        with st.expander("💎 DDM & RIM Parameters (Leave at 0 for Auto-Calculation)"):
            st.info("💡 **Note:** These parameters apply to Dividend Discount Model and Residual Income Model. Leave at 0 for auto-calculation from actual data.")
        
            st.markdown("### 💰 Dividend Discount Model (DDM) Parameters")
            col_ddm1, col_ddm2, col_ddm3 = st.columns(3)
        
            with col_ddm1:
                ddm_dividend_growth_screener = st.number_input(
                    "Dividend Growth Rate (%)",
                    min_value=0.0,
                    max_value=200.0,
                    value=0.0,
                    step=0.5,
                    key='screener_ddm_div_growth',
                    help="0 = Auto-calculate from historical dividend data. Used in Gordon Growth Model."
                )
        
            with col_ddm2:
                ddm_required_return_screener = st.number_input(
                    "Required Return / Cost of Equity (%)",
                    min_value=5.0,
                    max_value=30.0,
                    value=12.0,
                    step=0.5,
                    key='screener_ddm_required_return',
                    help="Discount rate for DDM. Usually = Cost of Equity from CAPM."
                )
            
            with col_ddm3:
                ddm_payout_ratio_screener = st.number_input(
                    "Payout Ratio (%)",
                    min_value=0.0,
                    max_value=100.0,
                    value=0.0,
                    step=5.0,
                    key='screener_ddm_payout',
                    help="0 = Auto-calculate from historical data. % of earnings paid as dividends."
                )
        
            st.markdown("### 🏢 Residual Income Model (RIM) Parameters")
            col_rim1, col_rim2, col_rim3, col_rim4 = st.columns(4)
        
            with col_rim1:
                rim_required_return_screener = st.number_input(
                    "Required Return (%)",
                    min_value=5.0,
                    max_value=30.0,
                    value=12.0,
                    step=0.5,
                    key='screener_rim_required_return',
                    help="Discount rate for RIM. Usually = Cost of Equity."
                )
        
            with col_rim2:
                rim_assumed_roe_screener = st.number_input(
                    "Assumed ROE (%)",
                    min_value=0.0,
                    max_value=50.0,
                    value=0.0,
                    step=1.0,
                    key='screener_rim_roe',
                    help="0 = Auto-calculate from historical data. Return on Equity assumption."
                )
        
            with col_rim3:
                rim_terminal_growth_screener = st.number_input(
                    "Terminal Growth (%)",
                    min_value=0.0,
                    max_value=20.0,
                    value=0.0,
                    step=0.5,
                    key='screener_rim_terminal_growth',
                    help="0 = Use same as DCF terminal growth. Long-term perpetual growth rate."
                )
        
            with col_rim4:
                rim_projection_years_screener = st.number_input(
                    "Projection Years",
                    min_value=0,
                    max_value=15,
                    value=0,
                    step=1,
                    key='screener_rim_proj_years',
                    help="0 = Use same as DCF projection years (5). Number of years to project."
                )
        
        # Run valuation button
        if excel_file_screener and company_name_screener:
            # Stock Price Comparison Feature Toggle - compact, before run button
            enable_stock_comparison_screener = st.checkbox(
                "📈 Show Stock Price vs Revenue & EPS chart",
                value=False,
                key='screener_stock_comparison',
                help="Compare stock price with financials (max 10 years)"
            )
            
            # INPUT CHANGE DETECTION - Reset results if key inputs change
            current_inputs_screener = {
                'excel_file': excel_file_screener.name if excel_file_screener and hasattr(excel_file_screener, 'name') else None,
                'company_name': company_name_screener,
                'terminal_growth': terminal_growth_screener,
                'tax_rate': tax_rate_screener,
                'manual_discount': manual_discount_rate_screener if manual_discount_rate_screener > 0 else None,
                'rev_growth': rev_growth_override_screener
            }
            
            # Check if inputs changed
            if 'previous_inputs_screener' in st.session_state:
                if st.session_state.previous_inputs_screener != current_inputs_screener:
                    # Inputs changed - clear results
                    st.session_state.show_results_screener = False
            
            # Store current inputs
            st.session_state.previous_inputs_screener = current_inputs_screener
            
            st.markdown("---")
            scenario_case_screener, scenario_pct_screener, scenario_drivers_screener = render_scenario_settings_ui("scenario_screener")
            
            st.markdown("---")
            st.markdown("### 🎯 Ready to Run Valuation")
            st.info("💡 **Click the button below to run valuation.** Results will appear only after clicking.")
            
            col_run1, col_run2 = st.columns([2, 1])
            
            with col_run1:
                if st.button("🚀 Run Screener Mode Valuation", type="primary", key="run_screener_dcf_btn"):
                    st.session_state.show_results_screener = True
            
            with col_run2:
                if st.button("🗑️ Clear Results", help="Clear current valuation results", key="clear_screener_results"):
                    st.session_state.show_results_screener = False
                    st.success("✅ Results cleared!")
                    st.rerun()
            
            if st.session_state.get('show_results_screener', False):
                with st.spinner("Processing Screener Excel..."):
                    # Accept either a raw Screener.in export ("Data Sheet") or an
                    # already-cleaned template — auto-clean/reshape the raw case
                    # using the same pipeline the auto-download feature uses, so
                    # the user never has to pre-clean anything themselves.
                    prepared_path, was_auto_cleaned, _prep_diagnostic = prepare_uploaded_screener_excel(excel_file_screener)
                    if prepared_path is None:
                        st.error(
                            "Could not read this Excel file. It should either be a raw export from Screener.in "
                            "(containing a 'Data Sheet' tab) or a pre-cleaned template (containing 'Balance Sheet' "
                            "and 'Profit and Loss Account' tabs)."
                        )
                        with st.expander("🔍 Diagnostic details", expanded=True):
                            st.code(_prep_diagnostic)
                            st.caption(f"Input type received: {type(excel_file_screener)}")
                        st.stop()
                    if was_auto_cleaned:
                        st.success("✅ Detected a raw Screener.in export — automatically cleaned and reshaped it, no manual prep needed.")

                    # Parse Excel using Screener mode parser
                    df_bs_screener, df_pl_screener = parse_screener_excel_to_dataframes(prepared_path)
                    
                    if df_bs_screener is None or df_pl_screener is None:
                        st.error("Failed to parse Screener Excel file")
                        st.stop()
                    
                    # Detect year columns
                    year_cols_screener_all = detect_screener_year_columns(df_bs_screener)
                    
                    if len(year_cols_screener_all) < 2:
                        st.error("Need at least 3 years of historical data in Excel file")
                        st.stop()
                    
                    st.success(f"✅ Excel contains {len(year_cols_screener_all)} years of data: {', '.join([y.replace('_', '') for y in year_cols_screener_all])}")
                    
                    # Limit to user-selected historical years
                    if historical_years_screener > len(year_cols_screener_all):
                        st.warning(f"⚠️ Requested {historical_years_screener} years but Excel has only {len(year_cols_screener_all)} years. Using all available data.")
                        year_cols_screener = year_cols_screener_all
                    else:
                        year_cols_screener = year_cols_screener_all[-historical_years_screener:]
                        st.info(f"📊 Using {len(year_cols_screener)} most recent years as selected: {', '.join([y.replace('_', '') for y in year_cols_screener])}")
                    
                    # ================================
                    # RESOLVE SHARES OUTSTANDING
                    # ================================
                    num_shares_screener = 0
                    shares_final_source = "Unknown"

                    if manual_shares_override_screener > 0:
                        # Manual entry always takes top priority
                        num_shares_screener = manual_shares_override_screener
                        shares_final_source = "Manual Override (User Input)"
                        st.success(f"✅ Using manual override: {num_shares_screener:,} shares ({num_shares_screener/10000000:.2f} Crore)")

                    elif use_yahoo_shares_screener:
                        # Fetch shares from Yahoo Finance
                        if ticker_symbol_screener and ticker_symbol_screener.strip():
                            with st.spinner("Fetching shares outstanding from Yahoo Finance..."):
                                try:
                                    import yfinance as _yf
                                    full_ticker_yf = f"{ticker_symbol_screener.strip().upper()}.{exchange_screener}"
                                    yf_ticker = _yf.Ticker(full_ticker_yf)
                                    yf_info = yf_ticker.info
                                    yf_shares = yf_info.get('sharesOutstanding', 0) or yf_info.get('impliedSharesOutstanding', 0)
                                    if yf_shares and yf_shares > 0:
                                        num_shares_screener = int(yf_shares)
                                        shares_final_source = "Yahoo Finance (sharesOutstanding)"
                                        st.success(f"✅ Fetched from Yahoo Finance: {num_shares_screener:,} shares ({num_shares_screener/10000000:.2f} Crore)")
                                    else:
                                        st.warning("⚠️ Yahoo Finance returned 0 shares. Falling back to Screener Excel.")
                                        shares_from_excel = get_screener_shares_outstanding(df_bs_screener, year_cols_screener[-1])
                                        if shares_from_excel > 0:
                                            num_shares_screener = shares_from_excel
                                            shares_final_source = "Screener Excel (Yahoo Finance fallback)"
                                            st.info(f"📋 Fallback: {num_shares_screener:,} shares from Excel")
                                except Exception as e:
                                    st.warning(f"⚠️ Yahoo Finance fetch failed: {str(e)[:120]}. Falling back to Screener Excel.")
                                    shares_from_excel = get_screener_shares_outstanding(df_bs_screener, year_cols_screener[-1])
                                    if shares_from_excel > 0:
                                        num_shares_screener = shares_from_excel
                                        shares_final_source = "Screener Excel (Yahoo Finance fallback)"
                                        st.info(f"📋 Fallback: {num_shares_screener:,} shares from Excel")
                        else:
                            st.warning("⚠️ No ticker provided for Yahoo Finance fetch. Falling back to Screener Excel.")
                            shares_from_excel = get_screener_shares_outstanding(df_bs_screener, year_cols_screener[-1])
                            if shares_from_excel > 0:
                                num_shares_screener = shares_from_excel
                                shares_final_source = "Screener Excel (no ticker for Yahoo Finance)"
                                st.info(f"📋 Fallback: {num_shares_screener:,} shares from Excel")

                    else:
                        # Default: use Screener Excel
                        shares_from_excel = get_screener_shares_outstanding(df_bs_screener, year_cols_screener[-1])
                        if shares_from_excel > 0:
                            num_shares_screener = shares_from_excel
                            shares_final_source = "Screener Excel"
                            st.success(f"✅ Auto-detected {num_shares_screener:,} shares from Excel")
                        else:
                            st.warning("⚠️ Could not auto-detect shares outstanding from Excel")

                    # If still 0 after all sources, prompt manual entry as last resort
                    if num_shares_screener == 0:
                        st.error("❌ Could not determine shares outstanding from any source.")
                        num_shares_screener = st.number_input(
                            "Please enter Number of Shares Outstanding:",
                            min_value=1,
                            value=100000,
                            step=1000,
                            key='screener_shares_manual_fallback',
                            help="Enter the total number of shares outstanding for the company"
                        )
                        if num_shares_screener == 100000:
                            st.error("❌ Please enter the actual number of shares outstanding (default 100,000 is just a placeholder)")
                            st.stop()
                        shares_final_source = "Manual Fallback (User Input)"
                    
                    # Extract financials using Screener extraction
                    financials_screener = extract_screener_financials(df_bs_screener, df_pl_screener, year_cols_screener)
                    
                    # Confirm/update from financials extraction if we're using Screener Excel and haven't got a better source
                    if shares_final_source in ("Unknown", "Screener Excel") and manual_shares_override_screener == 0 and not use_yahoo_shares_screener:
                        if financials_screener.get('num_shares') and financials_screener['num_shares'] > 0:
                            num_shares_screener = financials_screener['num_shares']
                            shares_final_source = "Screener Excel (financials extraction)"
                            st.info(f"📋 Confirmed shares from Excel extraction: {num_shares_screener:,} shares")

                    # Add num_shares and source to financials dict
                    financials_screener['num_shares'] = num_shares_screener
                    st.caption(f"ℹ️ Shares source used: **{shares_final_source}**")
                    
                    # Display financial summary
                    display_screener_financial_summary(financials_screener)
                    
                    # ================================
                    # FETCH TICKER DATA (Current Price & Beta)
                    # ================================
                    st.markdown("---")
                    st.subheader("📈 Market Data from Yahoo Finance")
                    
                    current_price_screener = 0.0
                    beta_screener = 1.0
                    
                    if manual_price_screener is not None:
                        # Company not listed on NSE/BSE — use the manually entered price/beta
                        current_price_screener = manual_price_screener or 0.0
                        beta_screener = manual_beta_screener or 1.0
                        if current_price_screener > 0:
                            st.metric("📊 Current Market Price (manual entry)", f"₹{current_price_screener:.2f}")
                        else:
                            st.info(f"💡 No manual price entered — proceeding without a current-price comparison (beta = {beta_screener:.2f}).")
                    elif ticker_symbol_screener and ticker_symbol_screener.strip():
                        _as_of_screener = None if is_valuation_date_today(valuation_date_screener) else valuation_date_screener
                        ticker_data = fetch_ticker_data_for_screener(ticker_symbol_screener, exchange_screener, as_of_date=_as_of_screener)
                        
                        if ticker_data and not ticker_data.get('error'):
                            current_price_screener = ticker_data.get('current_price') or 0.0
                            beta_screener = ticker_data.get('beta') or 1.0
                            
                            # Only display current price (beta not needed for Screener mode)
                            st.metric("📊 Current Market Price", f"₹{current_price_screener:.2f}")
                        else:
                            st.warning(f"⚠️ Could not fetch ticker data: {ticker_data.get('error', 'Unknown error')}")
                            st.info("💡 Continuing with default beta (1.0) for WACC calculation")
                    else:
                        st.info("💡 No ticker provided - using default beta (1.0) for WACC calculation. Current price comparison will not be available.")
                    
                    # ================================
                    # BUSINESS MODEL CLASSIFICATION
                    # ================================
                    st.markdown("---")
                    st.subheader("🏢 Business Model Classification")
                    
                    classification = classify_business_model(financials_screener, income_stmt=None, balance_sheet=None)
                    
                    # Show classification and check if FCFF DCF is allowed
                    should_stop = show_classification_warning(classification)
                    
                    if should_stop and run_dcf_screener:
                        st.warning("⚠️ DCF valuation skipped due to business model classification. Other models will still run.")
                        run_dcf_screener = False
                    
                    st.markdown("---")
                    
                    # Calculate WC metrics
                    wc_metrics = calculate_working_capital_metrics(financials_screener)
                    
                    # Show Working Capital Data Status
                    wc_status_parts = []
                    if wc_metrics.get('has_inventory', False) or (inventory_days_override_screener and inventory_days_override_screener > 0):
                        inv_source = f"User Override: {inventory_days_override_screener} days" if inventory_days_override_screener and inventory_days_override_screener > 0 else f"Historical: {wc_metrics['avg_inv_days']:.1f} days"
                        wc_status_parts.append(f"✅ Inventory ({inv_source})")
                    else:
                        wc_status_parts.append("⚠️ Inventory (No data)")
                    
                    if wc_metrics.get('has_receivables', False) or (debtor_days_override_screener and debtor_days_override_screener > 0):
                        deb_source = f"User Override: {debtor_days_override_screener} days" if debtor_days_override_screener and debtor_days_override_screener > 0 else f"Historical: {wc_metrics['avg_deb_days']:.1f} days"
                        wc_status_parts.append(f"✅ Debtors ({deb_source})")
                    else:
                        wc_status_parts.append("⚠️ Debtors (No data)")
                    
                    if wc_metrics.get('has_payables', False) or (creditor_days_override_screener and creditor_days_override_screener > 0):
                        cred_source = f"User Override: {creditor_days_override_screener} days" if creditor_days_override_screener and creditor_days_override_screener > 0 else f"Historical: {wc_metrics['avg_cred_days']:.1f} days"
                        wc_status_parts.append(f"✅ Creditors ({cred_source})")
                    else:
                        wc_status_parts.append("⚠️ Creditors (No data)")
                    
                    st.info(f"🔍 **Working Capital Projection Status:** {' | '.join(wc_status_parts)}")
                    
                    # Initialize results containers
                    dcf_results_screener = None
                    ddm_results_screener = None
                    rim_results_screener = None
                    comp_val_results_screener = None
                    projections_screener = None
                    
                    # ================================
                    # RUN DCF VALUATION
                    # ================================
                    if run_dcf_screener:
                        st.markdown("---")
                        st.subheader("💰 DCF (FCFF) Valuation")
                        
                        # Project financials (same as unlisted mode)
                        projections_screener, drivers_screener = project_financials(
                            financials_screener, wc_metrics, projection_years_screener, tax_rate_screener,
                            rev_growth_override_screener if rev_growth_override_screener > 0 else None,
                            opex_margin_override_screener if opex_margin_override_screener > 0 else None,
                            capex_ratio_override_screener if capex_ratio_override_screener > 0 else None,
                            # Per-year controls
                            rev_growth_per_year=rev_growth_per_year_screener if (_ms_enabled_screener or any(v > 0 for v in rev_growth_per_year_screener)) else None,
                            ebitda_margin_per_year=ebitda_margin_per_year_screener if (bool(_ms_margin_screener) or any(v > 0 for v in ebitda_margin_per_year_screener)) else None,
                            ebitda_margin_override=ebitda_margin_override_screener if ebitda_margin_override_screener > 0 else None,
                            depreciation_rate_override=depreciation_rate_override_screener if depreciation_rate_override_screener > 0 else None,
                            depreciation_method=depreciation_method_screener,
                            inventory_days_override=inventory_days_override_screener if inventory_days_override_screener > 0 else None,
                            debtor_days_override=debtor_days_override_screener if debtor_days_override_screener > 0 else None,
                            creditor_days_override=creditor_days_override_screener if creditor_days_override_screener > 0 else None,
                            interest_rate_override=interest_rate_override_screener if interest_rate_override_screener > 0 else None,
                            working_capital_pct_override=working_capital_as_pct_revenue_screener if working_capital_as_pct_revenue_screener > 0 else None
                        )
                        
                        # Calculate WACC (using peer companies for beta)
                        # NOTE: tax_rate_screener is a percentage (e.g. 25.0), same
                        # convention calculate_wacc() expects — do NOT divide by 100
                        # here, it already does `tax_rate / 100` internally.
                        wacc_details = calculate_wacc(financials_screener, tax_rate_screener, peer_tickers=peer_tickers_screener, manual_rf_rate=manual_rf_rate_screener, manual_rm_rate=manual_rm_rate_screener,
                                                      beta_start_date=beta_start_date_screener, beta_end_date=beta_end_date_screener)
                        
                        # Display risk-free rate being used
                        st.info(f"🏛️ Risk-Free Rate (India 10Y G-Sec): {manual_rf_rate_screener:.2f}%")
                        
                        # Extract cash balance (index [0] is NEWEST)
                        cash_balance = financials_screener['cash'][0] if financials_screener['cash'][0] > 0 else 0
                        
                        # Calculate DCF
                        valuation, error = calculate_dcf_valuation(
                            projections_screener, wacc_details, terminal_growth_screener, num_shares_screener, cash_balance,
                            manual_discount_rate=manual_discount_rate_screener if manual_discount_rate_screener > 0 else None
                        )
                        
                        if error:
                            st.error(error)
                            st.stop()
                        
                        # Store results
                        dcf_results_screener = valuation
                        
                        # Display DCF results
                        st.success("✅ DCF Valuation Complete!")
                        
                        # Key Metrics
                        col1, col2, col3, col4 = st.columns(4)
                        with col1:
                            st.metric("Enterprise Value", f"₹ {dcf_results_screener['enterprise_value']:.2f} Lacs")
                        with col2:
                            st.metric("Equity Value", f"₹ {dcf_results_screener['equity_value']:.2f} Lacs")
                        with col3:
                            st.metric("Fair Value/Share", f"₹ {dcf_results_screener['fair_value_per_share']:.2f}")
                        with col4:
                            st.metric("WACC", f"{wacc_details['wacc']:.2f}%")
                    
                    # ================================
                    # RUN DDM VALUATION
                    # ================================
                    if run_ddm_screener:
                        # Use user-provided parameters or defaults
                        ddm_growth = ddm_dividend_growth_screener / 100 if ddm_dividend_growth_screener > 0 else 0.05
                        ddm_req_return = ddm_required_return_screener / 100
                        
                        ddm_results_screener = calculate_screener_ddm_valuation(
                            financials_screener,
                            num_shares_screener,
                            required_return=ddm_req_return,
                            growth_rate=ddm_growth
                        )
                    
                    # ================================
                    # RUN RIM VALUATION
                    # ================================
                    if run_rim_screener:
                        # Use user-provided parameters or defaults
                        rim_req_return = rim_required_return_screener / 100
                        rim_proj_yrs = rim_projection_years_screener if rim_projection_years_screener > 0 else 5
                        rim_term_growth = rim_terminal_growth_screener / 100 if rim_terminal_growth_screener > 0 else terminal_growth_screener / 100
                        rim_roe = rim_assumed_roe_screener / 100 if rim_assumed_roe_screener > 0 else None
                        
                        rim_results_screener = calculate_screener_rim_valuation(
                            financials_screener,
                            num_shares_screener,
                            required_return=rim_req_return,
                            projection_years=rim_proj_yrs,
                            terminal_growth=rim_term_growth,
                            assumed_roe=rim_roe
                        )
                    
                    # ================================
                    # RUN COMPARATIVE VALUATION
                    # ================================
                    if run_comp_screener and peer_tickers_screener:
                        comp_val_results_screener = perform_comparative_valuation(
                            target_ticker=None,  # Unlisted
                            comp_tickers_str=peer_tickers_screener,
                            target_financials=financials_screener,
                            target_shares=num_shares_screener,
                            exchange_suffix="NS",
                            projections=projections_screener if run_dcf_screener else None,
                            use_screener_peers=False,  # Use Yahoo Finance for peers
                            valuation_date=valuation_date_screener
                        )
                    
                    # ================================
                    # TABBED DISPLAY - CONSISTENT WITH UNLISTED MODE
                    # ================================
                    st.markdown("---")
                    
                    # ================================
                    # CURRENT PRICE VS FAIR VALUES COMPARISON
                    # ================================
                    if current_price_screener > 0:
                        st.subheader("💰 Current Price vs Fair Values")
                        
                        # Collect all fair values
                        fair_values_dict = {}
                        
                        if dcf_results_screener and dcf_results_screener.get('fair_value_per_share', 0) > 0:
                            fair_values_dict['DCF (FCFF)'] = dcf_results_screener['fair_value_per_share']
                        
                        if ddm_results_screener and ddm_results_screener.get('status') == 'Success':
                            fair_values_dict['DDM (Gordon)'] = ddm_results_screener.get('value_per_share', 0)
                        
                        if rim_results_screener and rim_results_screener.get('status') == 'Success':
                            fair_values_dict['RIM (Residual Income)'] = rim_results_screener.get('value_per_share', 0)
                        
                        if comp_val_results_screener and comp_val_results_screener.get('valuations'):
                            valuations = comp_val_results_screener['valuations']
                            all_avg_values = [v.get('fair_value_avg', 0) for v in valuations.values()]
                            if all_avg_values:
                                fair_values_dict['Comparative Valuation (Avg)'] = np.mean(all_avg_values)
                        
                        if fair_values_dict:
                            # Create comparison DataFrame
                            comparison_data = []
                            for method, fair_value in fair_values_dict.items():
                                upside = ((fair_value - current_price_screener) / current_price_screener * 100) if current_price_screener > 0 else 0
                                comparison_data.append({
                                    'Valuation Method': method,
                                    'Fair Value (₹)': fair_value,
                                    'Current Price (₹)': current_price_screener,
                                    'Upside/Downside (%)': upside
                                })
                            
                            comparison_df = pd.DataFrame(comparison_data)
                            
                            # Display comparison table
                            st.dataframe(
                                comparison_df.style.format({
                                    'Fair Value (₹)': '₹{:.2f}',
                                    'Current Price (₹)': '₹{:.2f}',
                                    'Upside/Downside (%)': '{:+.1f}%'
                                }).map(
                                    lambda x: 'background-color: #d4edda' if isinstance(x, (int, float)) and x > 0 else ('background-color: #f8d7da' if isinstance(x, (int, float)) and x < 0 else ''),
                                    subset=['Upside/Downside (%)']
                                ),
                                use_container_width=True,
                                hide_index=True
                            )
                            
                            # Calculate and display average
                            avg_fair_value = np.mean(list(fair_values_dict.values()))
                            avg_upside = ((avg_fair_value - current_price_screener) / current_price_screener * 100)
                            
                            col_avg1, col_avg2, col_avg3 = st.columns(3)
                            with col_avg1:
                                st.metric("Current Market Price", f"₹{current_price_screener:.2f}")
                            with col_avg2:
                                st.metric("Average Fair Value (All Methods)", f"₹{avg_fair_value:.2f}")
                            with col_avg3:
                                st.metric("Average Upside/Downside", f"{avg_upside:+.1f}%", 
                                         delta=f"₹{avg_fair_value - current_price_screener:+.2f}")
                            
                            # Add Gauge Chart - Only if fair value is valid (positive and reasonable)
                            if avg_fair_value > 0:
                                gauge_chart = create_price_vs_value_gauge(current_price_screener, avg_fair_value)
                                if gauge_chart:
                                    st.plotly_chart(gauge_chart, use_container_width=True)
                                if 'ticker_symbol_screener' in locals() and ticker_symbol_screener:
                                    render_screener_company_info(ticker_symbol_screener, key_prefix="company_info_screener")
                            else:
                                st.error("⚠️ **Invalid Fair Value**: The calculated fair value is negative or zero, which indicates a problem with the valuation inputs or methodology. Please review your financial data and assumptions.")
                        else:
                            st.info("No valid fair values calculated yet")
                        
                        st.markdown("---")
                    
                    # Build tab list based on what was run
                    tab_list = ["📊 Historical Financials"]
                    if run_dcf_screener:
                        tab_list.extend(["📋 Assumptions & Inputs", "📈 Projections", "💰 FCF Working", "🎯 WACC Calculation", "🏆 DCF Summary", "📉 Sensitivity Analysis", "🎭 Scenario Analysis", "🎲 Monte Carlo"])
                    if run_ddm_screener:
                        tab_list.append("💸 DDM Valuation")
                    if run_rim_screener:
                        tab_list.append("📚 RIM Valuation")
                    if run_comp_screener and peer_tickers_screener:
                        tab_list.append("📊 Comparative Valuation")
                    tab_list.append("🏈 Football Field")
                    tab_list.append("⚖️ IBC Liquidation")
                    tab_list.append("📐 Ratio Analysis")
                    tab_list.append("🩺 Health Check")
                    tab_list.append("🧮 Altman Z-Score")
                    tab_list.append("💼 LBO Analysis")
                    
                    # Add Stock Comparison tab if enabled
                    enable_stock_comparison_screener_state = st.session_state.get('screener_stock_comparison', False)
                    if enable_stock_comparison_screener_state and STOCK_COMPARISON_AVAILABLE:
                        tab_list.append("📈 Stock vs Financials")
                    
                    tabs = st.tabs(tab_list)
                    tab_idx = 0
                    
                    # Tab 1: Historical Financials
                    with tabs[tab_idx]:
                        st.subheader("Historical Financials")
                        
                        # Add Historical Financials Chart with error handling
                        try:
                            # Pass reverse_years=True for chronological display (Screener data is newest-first internally)
                            st.plotly_chart(create_historical_financials_chart(financials_screener, reverse_years=True), use_container_width=True)
                        except Exception as e:
                            st.error(f"Historical chart error: {str(e)}")
                        
                        st.markdown("---")
                        st.markdown("### Income Statement")
                        
                        # Reverse data for chronological display (oldest to newest)
                        years_display = list(reversed(financials_screener['years']))
                        
                        hist_df = pd.DataFrame({
                            'Year': [str(y) for y in years_display],
                            'Revenue': list(reversed(financials_screener['revenue'])),
                            'Operating Expenses': list(reversed(financials_screener['opex'])),
                            'EBITDA': list(reversed(financials_screener['ebitda'])),
                            'Depreciation': list(reversed(financials_screener['depreciation'])),
                            'EBIT': list(reversed(financials_screener['ebit'])),
                            'Interest': list(reversed(financials_screener['interest'])),
                            'Tax': list(reversed(financials_screener['tax'])),
                            'NOPAT': list(reversed(financials_screener['nopat']))
                        })
                        numeric_cols = hist_df.select_dtypes(include=[np.number]).columns.tolist()
                        format_dict = {col: '{:.2f}' for col in numeric_cols}
                        st.dataframe(hist_df.style.format(format_dict), use_container_width=True)
                        
                        st.markdown("---")
                        st.subheader("Balance Sheet Metrics")
                        bs_df = pd.DataFrame({
                            'Year': [str(y) for y in years_display],
                            'Fixed Assets': list(reversed(financials_screener['fixed_assets'])),
                            'Inventory': list(reversed(financials_screener['inventory'])),
                            'Receivables': list(reversed(financials_screener['receivables'])),
                            'Payables': list(reversed(financials_screener['payables'])),
                            'Equity': list(reversed(financials_screener['equity'])),
                            'ST Debt': list(reversed(financials_screener['st_debt'])),
                            'LT Debt': list(reversed(financials_screener['lt_debt']))
                        })
                        numeric_cols = bs_df.select_dtypes(include=[np.number]).columns.tolist()
                        format_dict = {col: '{:.2f}' for col in numeric_cols}
                        st.dataframe(bs_df.style.format(format_dict), use_container_width=True)
                    
                    tab_idx += 1
                    
                    # Tab 2: Assumptions & Inputs (NEW - for transparency)
                    if run_dcf_screener:
                        with tabs[tab_idx]:
                            st.subheader("📋 All Assumptions, Inputs & Parameters Used")
                            st.info("💡 This tab shows ALL parameters, assumptions, and ratios used in the valuation for complete transparency.")
                            
                            # Section 1: Data Configuration - COMPACT & CLEAN
                            st.markdown("### 🎯 Data Configuration")
                            
                            # Compact metrics row
                            years_display = [str(y).replace('_', '') for y in financials_screener['years']]
                            years_range = f"{years_display[-1]} - {years_display[0]}" if len(years_display) > 1 else years_display[0]
                            
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.metric("📅 Historical Years", f"{len(financials_screener['years'])} years", delta=years_range, delta_color="off")
                            with col2:
                                st.metric("🔮 Projection Period", f"{projection_years_screener} years", delta="Future", delta_color="off")
                            with col3:
                                st.metric("🏢 Shares Outstanding", f"{num_shares_screener:,}")
                            
                            st.markdown("---")
                            
                            # Section 2: WACC Components - COMPACT VISUAL CHART
                            st.markdown("### 💰 WACC & Discount Rate")
                            
                            # Create compact WACC breakdown chart
                            fig_wacc = go.Figure()
                            
                            fig_wacc.add_trace(go.Bar(
                                x=['WACC', 'Cost of Equity', 'Cost of Debt', 'Terminal Growth'],
                                y=[wacc_details['wacc'], wacc_details['ke'], wacc_details['kd_after_tax'], terminal_growth_screener],
                                marker=dict(color=['#667eea', '#764ba2', '#f093fb', '#f5576c']),
                                text=[f"{wacc_details['wacc']:.2f}%", f"{wacc_details['ke']:.2f}%", 
                                      f"{wacc_details['kd_after_tax']:.2f}%", f"{terminal_growth_screener:.2f}%"],
                                textposition='auto',
                                textfont=dict(size=14, color='white')
                            ))
                            
                            fig_wacc.update_layout(
                                height=250,
                                margin=dict(l=20, r=20, t=20, b=20),
                                showlegend=False,
                                paper_bgcolor='rgba(0,0,0,0)',
                                plot_bgcolor='rgba(0,0,0,0)',
                                yaxis=dict(title='Rate (%)', gridcolor='rgba(128,128,128,0.2)'),
                                xaxis=dict(title='')
                            )
                            
                            st.plotly_chart(fig_wacc, use_container_width=True)
                            
                            # Compact key assumptions in 2 columns
                            col_a, col_b = st.columns(2)
                            with col_a:
                                st.caption("**Tax Rate:** " + f"{tax_rate_screener:.2f}% | **Beta:** {wacc_details['beta']:.3f}")
                                st.caption("**Risk-Free Rate:** " + f"{wacc_details['rf']:.2f}% | **Market Return:** {wacc_details['rm']:.2f}%")
                            with col_b:
                                st.caption("**Equity Weight:** " + f"{wacc_details['we']:.2f}% | **Debt Weight:** {wacc_details['wd']:.2f}%")
                                st.caption("**Cost of Equity:** " + f"{wacc_details['ke']:.2f}% | **Cost of Debt:** {wacc_details['kd_after_tax']:.2f}%")
                            
                            st.markdown("---")
                            
                            # Section 3: Capital Structure - VISUAL PIE CHART
                            st.markdown("### 🏛️ Capital Structure (Latest Year)")
                            
                            equity = financials_screener['equity'][0]
                            total_debt = financials_screener['st_debt'][0] + financials_screener['lt_debt'][0]
                            cash = financials_screener['cash'][0]
                            net_debt = total_debt - cash
                            
                            fig_cap = go.Figure(data=[go.Pie(
                                labels=['Equity', 'Debt', 'Cash'],
                                values=[equity, total_debt, cash],
                                hole=.4,
                                marker=dict(colors=['#06A77D', '#D62828', '#F77F00']),
                                textinfo='label+percent',
                                textfont=dict(size=12)
                            )])
                            
                            fig_cap.update_layout(
                                height=300,
                                margin=dict(l=20, r=20, t=20, b=20),
                                showlegend=True,
                                legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5)
                            )
                            
                            st.plotly_chart(fig_cap, use_container_width=True)
                            
                            col_cap1, col_cap2, col_cap3 = st.columns(3)
                            with col_cap1:
                                st.metric("Equity", f"₹{equity:,.0f} L")
                            with col_cap2:
                                st.metric("Total Debt", f"₹{total_debt:,.0f} L")
                            with col_cap3:
                                st.metric("Net Debt", f"₹{net_debt:,.0f} L", delta="Cash adjusted")
                            
                            st.markdown("---")
                            
                            # Section 4: Projection Drivers - COMPACT TABLE
                            st.markdown("### 📊 Projection Drivers & Ratios")
                            
                            with st.expander("📋 View All Projection Parameters", expanded=False):
                                drivers_data = []
                                if drivers_screener:
                                    drivers_data.append(['Revenue Growth', f"{drivers_screener['avg_growth']:.2f}%", 'User Override' if rev_growth_override_screener > 0 else 'Historical CAGR'])
                                    drivers_data.append(['OpEx Margin', f"{drivers_screener['avg_opex_margin']:.2f}%", 'User Override' if opex_margin_override_screener > 0 else 'Historical Avg'])
                                    drivers_data.append(['CapEx/Revenue', f"{drivers_screener['avg_capex_ratio']:.2f}%", 'User Override' if capex_ratio_override_screener > 0 else 'Historical Avg'])
                                    drivers_data.append(['Depreciation', f"{'Auto' if depreciation_rate_override_screener == 0 else depreciation_rate_override_screener}", depreciation_method_screener])
                                    
                                    if inventory_days_override_screener > 0 or wc_metrics.get('avg_inv_days', 0) > 0:
                                        drivers_data.append(['Inventory Days', f"{inventory_days_override_screener if inventory_days_override_screener > 0 else wc_metrics.get('avg_inv_days', 0):.0f}", 'Override' if inventory_days_override_screener > 0 else 'Historical'])
                                    if debtor_days_override_screener > 0 or wc_metrics.get('avg_deb_days', 0) > 0:
                                        drivers_data.append(['Debtor Days', f"{debtor_days_override_screener if debtor_days_override_screener > 0 else wc_metrics.get('avg_deb_days', 0):.0f}", 'Override' if debtor_days_override_screener > 0 else 'Historical'])
                                    if creditor_days_override_screener > 0 or wc_metrics.get('avg_cred_days', 0) > 0:
                                        drivers_data.append(['Creditor Days', f"{creditor_days_override_screener if creditor_days_override_screener > 0 else wc_metrics.get('avg_cred_days', 0):.0f}", 'Override' if creditor_days_override_screener > 0 else 'Historical'])
                                
                                if drivers_data:
                                    drivers_compact = pd.DataFrame(drivers_data, columns=['Parameter', 'Value', 'Source'])
                                    st.dataframe(drivers_compact, use_container_width=True, hide_index=True, height=250)

                            
                            # Section 4: Capital Structure
                            st.markdown("### 🏦 Capital Structure (Latest Year)")
                            capital_df = pd.DataFrame({
                                'Item': [
                                    'Equity',
                                    'Short-term Debt',
                                    'Long-term Debt',
                                    'Total Debt',
                                    'Cash & Equivalents',
                                    'Net Debt',
                                    'Total Capital'
                                ],
                                'Value (₹ Lacs)': [
                                    f"{financials_screener['equity'][0]:.2f}",
                                    f"{financials_screener['st_debt'][0]:.2f}",
                                    f"{financials_screener['lt_debt'][0]:.2f}",
                                    f"{financials_screener['st_debt'][0] + financials_screener['lt_debt'][0]:.2f}",
                                    f"{financials_screener['cash'][0]:.2f}",
                                    f"{(financials_screener['st_debt'][0] + financials_screener['lt_debt'][0]) - financials_screener['cash'][0]:.2f}",
                                    f"{financials_screener['equity'][0] + financials_screener['st_debt'][0] + financials_screener['lt_debt'][0]:.2f}"
                                ]
                            })
                            st.dataframe(capital_df, use_container_width=True, hide_index=True)
                            
                            st.markdown("---")
                            
                            # Section 5: Valuation Model Selection
                            st.markdown("### 🎯 Valuation Models Executed")
                            models_run = []
                            if run_dcf_screener:
                                models_run.append("✅ DCF (FCFF) - Discounted Cash Flow")
                            if run_ddm_screener:
                                models_run.append("✅ DDM - Dividend Discount Model")
                            if run_rim_screener:
                                models_run.append("✅ RIM - Residual Income Model")
                            if run_comp_screener:
                                models_run.append("✅ Comparative Valuation - Peer Multiples")
                            
                            for model in models_run:
                                st.write(model)
                            
                            st.success("💡 **Transparency Note:** All assumptions and inputs used in this valuation are documented above for your review and validation.")
                    
                    tab_idx += 1
                    
                    # Tab 2: Projections (if DCF was run)
                    if run_dcf_screener:
                        with tabs[tab_idx]:
                            st.subheader(f"Projected Financials ({projection_years_screener} Years)")
                            
                            # Add Projection Chart with error handling
                            try:
                                st.plotly_chart(create_fcff_projection_chart(projections_screener), use_container_width=True)
                            except Exception as e:
                                st.error(f"Chart error: {str(e)}")
                            
                            st.markdown("---")
                            
                            proj_df = pd.DataFrame({
                                'Year': [f"Year {y}" for y in projections_screener['year']],
                                'Revenue': projections_screener['revenue'],
                                'EBITDA': projections_screener['ebitda'],
                                'EBIT': projections_screener['ebit'],
                                'NOPAT': projections_screener['nopat'],
                                'CapEx': projections_screener['capex'],
                                'Δ WC': projections_screener['delta_wc'],
                                'FCFF': projections_screener['fcff']
                            })
                            numeric_cols = proj_df.select_dtypes(include=[np.number]).columns.tolist()
                            format_dict = {col: '{:.2f}' for col in numeric_cols}
                            st.dataframe(proj_df.style.format(format_dict), use_container_width=True)
                            
                            st.info(f"**Key Drivers:** Revenue Growth: {drivers_screener['avg_growth']:.2f}% | Opex Margin: {drivers_screener['avg_opex_margin']:.2f}% | CapEx/Revenue: {drivers_screener['avg_capex_ratio']:.2f}%")
                        
                        tab_idx += 1
                        
                        # Tab 3: FCF Working
                        with tabs[tab_idx]:
                            st.subheader("Free Cash Flow Working")
                            
                            fcf_df = pd.DataFrame({
                                'Year': [f"Year {y}" for y in projections_screener['year']],
                                'NOPAT': projections_screener['nopat'],
                                '(+) Depreciation': projections_screener['depreciation'],
                                '(-) ΔWC': projections_screener['delta_wc'],
                                '(-) CapEx': projections_screener['capex'],
                                'FCFF': projections_screener['fcff'],
                                'PV(FCFF)': dcf_results_screener['pv_fcffs']
                            })
                            numeric_cols = fcf_df.select_dtypes(include=[np.number]).columns.tolist()
                            format_dict = {col: '{:.2f}' for col in numeric_cols}
                            st.dataframe(fcf_df.style.format(format_dict), use_container_width=True)
                            
                            st.info(f"**Sum of PV(FCFF):** ₹ {dcf_results_screener['sum_pv_fcff']:.2f} Lacs")
                        
                        tab_idx += 1
                        
                        # Tab 4: WACC Calculation
                        with tabs[tab_idx]:
                            st.subheader("🎯 WACC Calculation & Breakdown")
                            
                            # Add WACC Breakdown Chart with error handling
                            try:
                                st.plotly_chart(create_wacc_breakdown_chart(wacc_details), use_container_width=True)
                            except Exception as e:
                                st.error(f"WACC chart error: {str(e)}")
                            
                            st.markdown("---")
                            
                            # WACC Components
                            col1, col2 = st.columns(2)
                            
                            with col1:
                                st.markdown("#### Cost of Equity (Ke)")
                                ke_df = pd.DataFrame({
                                    'Component': ['Risk-Free Rate', 'Beta', 'Market Return', 'Market Risk Premium', 'Cost of Equity'],
                                    'Value': [
                                        f"{wacc_details['rf']:.2f}%",
                                        f"{wacc_details['beta']:.3f}",
                                        f"{wacc_details['rm']:.2f}%",
                                        f"{wacc_details['rm'] - wacc_details['rf']:.2f}%",
                                        f"{wacc_details['ke']:.2f}%"
                                    ]
                                })
                                st.table(ke_df)
                            
                            with col2:
                                st.markdown("#### Cost of Debt (Kd)")
                                kd_df = pd.DataFrame({
                                    'Component': ['Total Debt', 'Interest Expense', 'Cost of Debt (Pre-tax)', 'Tax Rate', 'Cost of Debt (After-tax)'],
                                    'Value': [
                                        f"₹ {wacc_details['debt']:.2f} Lacs",
                                        f"₹ {financials_screener['interest'][0]:.2f} Lacs",
                                        f"{wacc_details['kd']:.2f}%",
                                        f"{tax_rate_screener:.2f}%",
                                        f"{wacc_details['kd_after_tax']:.2f}%"
                                    ]
                                })
                                st.table(kd_df)
                            
                            st.markdown("#### WACC Calculation")
                            wacc_calc_df = pd.DataFrame({
                                'Component': ['Equity', 'Debt', 'Total Capital', 'Weight of Equity', 'Weight of Debt', 'WACC'],
                                'Value': [
                                    f"₹ {wacc_details['equity']:.2f} Lacs",
                                    f"₹ {wacc_details['debt']:.2f} Lacs",
                                    f"₹ {wacc_details['equity'] + wacc_details['debt']:.2f} Lacs",
                                    f"{wacc_details['we']:.2f}%",
                                    f"{wacc_details['wd']:.2f}%",
                                    f"{wacc_details['wacc']:.2f}%"
                                ]
                            })
                            st.table(wacc_calc_df)
                            
                            st.success(f"**WACC = (We × Ke) + (Wd × Kd) = ({wacc_details['we']:.2f}% × {wacc_details['ke']:.2f}%) + ({wacc_details['wd']:.2f}% × {wacc_details['kd_after_tax']:.2f}%) = {wacc_details['wacc']:.2f}%**")
                        
                        tab_idx += 1
                        
                        # Tab 5: DCF Summary
                        with tabs[tab_idx]:
                            st.subheader("🏆 DCF Valuation Summary")
                            
                            # CRITICAL: Check for negative fair value and display warning
                            if dcf_results_screener.get('negative_fair_value_warning'):
                                warning = dcf_results_screener['negative_fair_value_warning']
                                st.error("🚨 **CRITICAL ISSUE: Negative Fair Value Detected!**")
                                st.error(f"**Fair Value per Share: ₹{dcf_results_screener['fair_value_per_share']:.2f}**")
                                
                                st.markdown("### ⚠️ Problem Diagnosis:")
                                for reason in warning['reason']:
                                    st.warning(f"• {reason}")
                                
                                st.markdown("### 📊 Breakdown:")
                                problem_df = pd.DataFrame({
                                    'Metric': [
                                        'Enterprise Value',
                                        'Total Debt',
                                        'Cash & Equivalents',
                                        'Net Debt (Debt - Cash)',
                                        'Equity Value (EV - Net Debt)',
                                        'Number of Shares',
                                        'Fair Value per Share'
                                    ],
                                    'Value': [
                                        f"₹{warning['enterprise_value']:.2f} Lacs",
                                        f"₹{warning['total_debt']:.2f} Lacs",
                                        f"₹{warning['cash']:.2f} Lacs",
                                        f"₹{warning['net_debt']:.2f} Lacs",
                                        f"₹{warning['equity_value']:.2f} Lacs",
                                        f"{warning['num_shares']:,}",
                                        f"₹{dcf_results_screener['fair_value_per_share']:.2f}"
                                    ]
                                })
                                st.dataframe(problem_df, use_container_width=True, hide_index=True)
                                
                                st.markdown("### 💡 Possible Solutions:")
                                st.info("""
                                **1. Reduce Debt Assumptions:** If debt is too high, reduce debt levels or reclassify some debt.
                                
                                **2. Check Historical Data:** Verify that your Excel data is correct (especially debt and equity).
                                
                                **3. Increase Terminal Growth:** A higher terminal growth rate will increase terminal value and enterprise value.
                                
                                **4. Review WACC:** A lower WACC (discount rate) will increase present values.
                                
                                **5. Improve Projections:** Ensure FCFFs are positive and growing in projection years.
                                
                                **6. Use Alternative Models:** Consider DDM, RIM, or Comparative Valuation instead of DCF for this company.
                                """)
                                
                                st.markdown("---")
                            
                            # Add Waterfall Chart with error handling
                            try:
                                st.plotly_chart(create_waterfall_chart(dcf_results_screener), use_container_width=True)
                            except Exception as e:
                                st.error(f"Waterfall chart error: {str(e)}")
                            
                            st.markdown("---")
                            
                            st.markdown("### Enterprise Value Build-up")
                            ev_df = pd.DataFrame({
                                'Component': ['Sum of PV(FCFF)', 'PV(Terminal Value)', 'Enterprise Value'],
                                'Value (₹ Lacs)': [
                                    dcf_results_screener['sum_pv_fcff'],
                                    dcf_results_screener['pv_terminal_value'],
                                    dcf_results_screener['enterprise_value']
                                ]
                            })
                            st.dataframe(ev_df.style.format({'Value (₹ Lacs)': '{:.2f}'}), use_container_width=True)
                            
                            st.markdown("---")
                            st.markdown("### Equity Value Calculation")
                            equity_df = pd.DataFrame({
                                'Component': ['Enterprise Value', '(-) Total Debt', '(+) Cash', 'Equity Value', 'Number of Shares', 'Fair Value per Share'],
                                'Value': [
                                    f"₹ {dcf_results_screener['enterprise_value']:.2f} Lacs",
                                    f"₹ {dcf_results_screener['total_debt']:.2f} Lacs",
                                    f"₹ {dcf_results_screener['cash']:.2f} Lacs",
                                    f"₹ {dcf_results_screener['equity_value']:.2f} Lacs",
                                    f"{num_shares_screener:,.0f}",
                                    f"₹ {dcf_results_screener['fair_value_per_share']:.2f}"
                                ]
                            })
                            st.table(equity_df)
                            
                            st.info(f"**Terminal Value % of EV:** {dcf_results_screener['tv_percentage']:.1f}%")
                        
                        tab_idx += 1
                        
                        # Tab 6: Sensitivity Analysis (NEW)
                        with tabs[tab_idx]:
                            st.subheader("📉 Sensitivity Analysis")
                            
                            # Create WACC and growth ranges
                            wacc_range = np.arange(max(1.0, wacc_details['wacc'] - 3), wacc_details['wacc'] + 3.5, 0.5)
                            g_range = np.arange(max(1.0, terminal_growth_screener - 2), min(terminal_growth_screener + 3, wacc_details['wacc'] - 1), 0.5)
                            
                            if len(g_range) == 0:
                                g_range = np.array([terminal_growth_screener])
                            
                            # Add Sensitivity Heatmap with error handling
                            try:
                                st.plotly_chart(
                                    create_sensitivity_heatmap(projections_screener, wacc_range, g_range, num_shares_screener,
                                                              net_debt=dcf_results_screener['net_debt'],
                                                              terminal_fcff=dcf_results_screener.get('adjusted_terminal_fcff')),
                                    use_container_width=True
                                )
                            except Exception as e:
                                st.error(f"Sensitivity chart error: {str(e)}")
                            
                            st.info("💡 **How to Read:** Each cell shows the fair value per share for different combinations of WACC and terminal growth rate. Darker green = higher valuation, darker red = lower valuation.")
                        
                        tab_idx += 1
                        
                        # Tab: Scenario Analysis
                        with tabs[tab_idx]:
                            _base_ebitda_margin_s = (projections_screener['ebitda'][0] / projections_screener['revenue'][0] * 100) \
                                if projections_screener['revenue'] and projections_screener['revenue'][0] > 0 else drivers_screener['avg_opex_margin']
                            _base_wc_pct_s = (projections_screener['wc'][0] / projections_screener['revenue'][0] * 100) \
                                if projections_screener['revenue'] and projections_screener['revenue'][0] > 0 else 0.0
                            _base_drivers_screener = {
                                'growth': drivers_screener['avg_growth'],
                                'ebitda_margin': _base_ebitda_margin_s,
                                'wacc': wacc_details['wacc'],
                                'capex_ratio': drivers_screener['avg_capex_ratio'],
                                'wc_pct': _base_wc_pct_s,
                            }
                            render_scenario_analysis_tab(
                                financials_screener, wc_metrics, projection_years_screener, tax_rate_screener,
                                _base_drivers_screener, scenario_pct_screener, scenario_drivers_screener, scenario_case_screener,
                                terminal_growth_screener, num_shares_screener, cash_balance,
                                "₹", current_price=current_price_screener if 'current_price_screener' in locals() else 0,
                                key_prefix="scenario_screener",
                                fixed_overrides={
                                    'depreciation_rate_override': depreciation_rate_override_screener if depreciation_rate_override_screener > 0 else None,
                                    'depreciation_method': depreciation_method_screener,
                                    'inventory_days_override': inventory_days_override_screener if inventory_days_override_screener > 0 else None,
                                    'debtor_days_override': debtor_days_override_screener if debtor_days_override_screener > 0 else None,
                                    'creditor_days_override': creditor_days_override_screener if creditor_days_override_screener > 0 else None,
                                    'interest_rate_override': interest_rate_override_screener if interest_rate_override_screener > 0 else None,
                                },
                                base_fair_value=valuation['fair_value_per_share']
                            )
                        
                        tab_idx += 1
                        
                        # Tab: Monte Carlo Simulation
                        with tabs[tab_idx]:
                            _base_ebitda_margin_mc_s = (projections_screener['ebitda'][0] / projections_screener['revenue'][0] * 100) \
                                if projections_screener['revenue'] and projections_screener['revenue'][0] > 0 else drivers_screener['avg_opex_margin']
                            _base_wc_pct_mc_s = (projections_screener['wc'][0] / projections_screener['revenue'][0] * 100) \
                                if projections_screener['revenue'] and projections_screener['revenue'][0] > 0 else 0.0
                            _base_drivers_mc_screener = {
                                'growth': drivers_screener['avg_growth'],
                                'ebitda_margin': _base_ebitda_margin_mc_s,
                                'wacc': wacc_details['wacc'],
                                'capex_ratio': drivers_screener['avg_capex_ratio'],
                                'wc_pct': _base_wc_pct_mc_s,
                            }
                            render_monte_carlo_tab(
                                financials_screener, wc_metrics, projection_years_screener, tax_rate_screener,
                                _base_drivers_mc_screener, terminal_growth_screener, num_shares_screener, cash_balance,
                                "₹", current_price=current_price_screener if 'current_price_screener' in locals() else 0,
                                key_prefix="mc_screener",
                                fixed_overrides={
                                    'depreciation_rate_override': depreciation_rate_override_screener if depreciation_rate_override_screener > 0 else None,
                                    'depreciation_method': depreciation_method_screener,
                                    'inventory_days_override': inventory_days_override_screener if inventory_days_override_screener > 0 else None,
                                    'debtor_days_override': debtor_days_override_screener if debtor_days_override_screener > 0 else None,
                                    'creditor_days_override': creditor_days_override_screener if creditor_days_override_screener > 0 else None,
                                    'interest_rate_override': interest_rate_override_screener if interest_rate_override_screener > 0 else None,
                                },
                                base_fair_value=valuation['fair_value_per_share']
                            )
                        
                        tab_idx += 1
                    
                    # Tab: DDM Valuation
                    if run_ddm_screener:
                        with tabs[tab_idx]:
                            st.subheader("💸 Dividend Discount Model (DDM)")
                            if ddm_results_screener:
                                display_screener_ddm_results(ddm_results_screener)
                            else:
                                st.warning("DDM results not available")
                        tab_idx += 1
                    
                    # Tab: RIM Valuation
                    if run_rim_screener:
                        with tabs[tab_idx]:
                            st.subheader("📚 Residual Income Model (RIM)")
                            if rim_results_screener:
                                display_screener_rim_results(rim_results_screener)
                            else:
                                st.warning("RIM results not available")
                        tab_idx += 1
                    
                    # Tab: Comparative Valuation
                    if run_comp_screener and peer_tickers_screener:
                        with tabs[tab_idx]:
                            st.subheader("📊 Comparative Valuation (Peer Multiples)")
                            
                            if comp_val_results_screener:
                                # Show comparables table
                                st.markdown("### Comparable Companies")
                                comp_df = pd.DataFrame(comp_val_results_screener['comparables'])
                                if not comp_df.empty:
                                    display_comp_df = comp_df[['ticker', 'name', 'price', 'pe', 'pb', 'ps', 'ev_ebitda', 'ev_sales']]
                                    st.dataframe(display_comp_df.style.format({
                                        'price': '₹{:.2f}',
                                        'pe': '{:.2f}x',
                                        'pb': '{:.2f}x',
                                        'ps': '{:.2f}x',
                                        'ev_ebitda': '{:.2f}x',
                                        'ev_sales': '{:.2f}x'
                                    }), use_container_width=True)
                                
                                # Show multiples statistics
                                st.markdown("### Peer Multiples Statistics")
                                for multiple, stats in comp_val_results_screener['multiples_stats'].items():
                                    with st.expander(f"📊 {multiple.upper()} - Avg: {stats['average']:.2f}x, Median: {stats['median']:.2f}x"):
                                        st.write(f"**Range:** {stats['min']:.2f}x - {stats['max']:.2f}x")
                                        st.write(f"**Std Dev:** {stats['std']:.2f}x")
                                        st.write(f"**Peer Values:** {', '.join([f'{v:.2f}x' for v in stats['values']])}")
                                
                                # Show implied valuations
                                st.markdown("### Implied Fair Values")
                                
                                all_avg_values = []
                                all_median_values = []
                                
                                for method_key, val_data in comp_val_results_screener['valuations'].items():
                                    st.markdown(f"#### {val_data['method']}")
                                    
                                    col1, col2 = st.columns(2)
                                    
                                    with col1:
                                        st.markdown("**Using Average Multiple:**")
                                        st.write(val_data['formula_avg'])
                                        st.metric("Fair Value (Avg)", f"₹{val_data['fair_value_avg']:.2f}")
                                        all_avg_values.append(val_data['fair_value_avg'])
                                    
                                    with col2:
                                        st.markdown("**Using Median Multiple:**")
                                        st.write(val_data['formula_median'])
                                        st.metric("Fair Value (Median)", f"₹{val_data['fair_value_median']:.2f}")
                                        all_median_values.append(val_data['fair_value_median'])
                                    
                                    st.markdown("---")
                                
                                # Summary statistics
                                if all_avg_values and all_median_values:
                                    st.markdown("### 📈 Comparative Valuation Summary")
                                    
                                    col1, col2, col3 = st.columns(3)
                                    
                                    with col1:
                                        st.metric("Average (All Methods)", f"₹{np.mean(all_avg_values):.2f}")
                                        st.metric("Median (All Methods)", f"₹{np.median(all_median_values):.2f}")
                                    
                                    with col2:
                                        st.metric("Min Fair Value", f"₹{min(all_avg_values + all_median_values):.2f}")
                                        st.metric("Max Fair Value", f"₹{max(all_avg_values + all_median_values):.2f}")
                                    
                                    with col3:
                                        if dcf_results_screener and dcf_results_screener['fair_value_per_share'] > 0:
                                            st.metric("DCF Fair Value", f"₹{dcf_results_screener['fair_value_per_share']:.2f}")
                                            combined_avg = (np.mean(all_avg_values) + dcf_results_screener['fair_value_per_share']) / 2
                                            st.metric("DCF + Comp Avg", f"₹{combined_avg:.2f}")
                            else:
                                st.warning("Could not fetch comparable companies data")
                        
                            st.markdown("---")
                            render_implied_multiples_panel(
                                dcf_results_screener, financials_screener, num_shares_screener,
                                comp_results=comp_val_results_screener,
                                current_price=current_price_screener if 'current_price_screener' in locals() else 0,
                                key_prefix="implied_screener"
                            )
                        tab_idx += 1
                    
                    # Tab: Football Field
                    with tabs[tab_idx]:
                        st.subheader("🏈 Football Field")
                        st.caption("Compares the fair-value range from every valuation method run above, side by side.")
                        band_pct_screener = st.number_input(
                            "± Band for single-point methods (DCF / DDM / RIM) (%):",
                            min_value=1.0, max_value=50.0, value=12.0, step=1.0,
                            key="football_band_screener",
                            help="DCF, DDM and RIM produce one number, not a range — this sets how wide a band "
                                 "to draw around that number on the chart. Comparable Valuation uses the actual "
                                 "low/high peer multiples instead, so this doesn't affect it."
                        )
                        _ff_methods_screener = []
                        if run_dcf_screener and dcf_results_screener and dcf_results_screener.get('fair_value_per_share', 0) > 0:
                            _fv = dcf_results_screener['fair_value_per_share']
                            _ff_methods_screener.append({'method': 'DCF', 'low': _fv * (1 - band_pct_screener / 100), 'high': _fv * (1 + band_pct_screener / 100), 'mid': _fv})
                        if run_ddm_screener and ddm_results_screener and ddm_results_screener.get('status') == 'Success' and ddm_results_screener.get('value_per_share', 0) > 0:
                            _fv = ddm_results_screener['value_per_share']
                            _ff_methods_screener.append({'method': 'DDM', 'low': _fv * (1 - band_pct_screener / 100), 'high': _fv * (1 + band_pct_screener / 100), 'mid': _fv})
                        if run_rim_screener and rim_results_screener and rim_results_screener.get('status') == 'Success' and rim_results_screener.get('value_per_share', 0) > 0:
                            _fv = rim_results_screener['value_per_share']
                            _ff_methods_screener.append({'method': 'RIM', 'low': _fv * (1 - band_pct_screener / 100), 'high': _fv * (1 + band_pct_screener / 100), 'mid': _fv})
                        if comp_val_results_screener:
                            for _key, _val_data in comp_val_results_screener.get('valuations', {}).items():
                                _stats = comp_val_results_screener.get('multiples_stats', {}).get(_key)
                                if _stats and _val_data.get('target_metric', 0) > 0 and _val_data.get('fair_value_avg', 0) > 0:
                                    _low = _val_data['target_metric'] * _stats.get('min', 0)
                                    _high = _val_data['target_metric'] * _stats.get('max', 0)
                                    if _high > 0:
                                        _ff_methods_screener.append({'method': _val_data.get('method', _key.upper()), 'low': _low, 'high': _high, 'mid': _val_data['fair_value_avg']})
                        render_football_field_chart(_ff_methods_screener, current_price_screener if 'current_price_screener' in locals() else 0, "₹", key_prefix="football_screener")
                    tab_idx += 1
                    
                    # Tab: IBC Liquidation Value
                    with tabs[tab_idx]:
                        render_ibc_liquidation_tab(
                            financials_screener, "₹", key_prefix="ibc_screener",
                            dcf_equity_value=(dcf_results_screener.get('equity_value', 0) * 100000 if run_dcf_screener and dcf_results_screener else None),
                            dcf_fair_value_per_share=(dcf_results_screener.get('fair_value_per_share') if run_dcf_screener and dcf_results_screener else None),
                            num_shares=num_shares_screener
                        )
                    tab_idx += 1
                    
                    # Tab: Ratio Analysis
                    with tabs[tab_idx]:
                        render_ratio_dashboard(
                            financials_screener, wc_metrics, "₹", key_prefix="ratios_screener",
                            current_price=current_price_screener if 'current_price_screener' in locals() else 0,
                            shares=num_shares_screener, tax_rate=tax_rate_screener
                        )
                    tab_idx += 1
                    
                    # Tab: Health Check
                    with tabs[tab_idx]:
                        render_valuation_health_check(
                            dcf_results_screener if run_dcf_screener else None,
                            terminal_growth_screener, wacc_details['wacc'] if 'wacc_details' in locals() else 0,
                            csym="₹", key_prefix="health_screener",
                            ddm_result=ddm_results_screener if run_ddm_screener else None,
                            rim_result=rim_results_screener if run_rim_screener else None,
                            comp_results=comp_val_results_screener
                        )
                    tab_idx += 1
                    
                    # Tab: Altman Z-Score
                    with tabs[tab_idx]:
                        render_altman_zscore_tab(
                            financials_screener, "₹", key_prefix="zscore_screener",
                            current_price=current_price_screener if 'current_price_screener' in locals() else 0,
                            shares=num_shares_screener
                        )
                    tab_idx += 1
                    
                    # Tab: LBO Analysis
                    with tabs[tab_idx]:
                        try:
                            render_lbo_analysis_tab(
                                financials_screener, "₹", key_prefix="lbo_screener", tax_rate=tax_rate_screener
                            )
                        except Exception as _lbo_err:
                            st.error(f"LBO Analysis failed to render: {_lbo_err}")
                            st.code(traceback.format_exc())
                    tab_idx += 1
                    
                    # Stock Price Comparison Tab (if enabled)
                    if enable_stock_comparison_screener_state and STOCK_COMPARISON_AVAILABLE:
                        with tabs[tab_idx]:
                            st.subheader("📈 Stock Price vs Revenue & EPS Analysis")
                            
                            if ticker_symbol_screener and ticker_symbol_screener.strip():
                                with st.spinner("Fetching stock price data..."):
                                    try:
                                        # Construct full ticker
                                        full_ticker_screener = f"{ticker_symbol_screener.strip().upper()}.{exchange_screener}"
                                        
                                        # Determine years to fetch (max 10 for Screener)
                                        years_to_fetch = min(historical_years_screener, 10)
                                        
                                        # Fetch comparison data
                                        stock_comp_data_screener = get_stock_comparison_data_screener(
                                            ticker=full_ticker_screener,
                                            company_name=company_name_screener,
                                            balance_sheet_df=df_bs_screener,
                                            pnl_df=df_pl_screener,
                                            num_years=years_to_fetch
                                        )
                                        
                                        if stock_comp_data_screener and stock_comp_data_screener['chart_fig']:
                                            st.plotly_chart(stock_comp_data_screener['chart_fig'], use_container_width=True)
                                            
                                            # Show data tables in expanders
                                            with st.expander("📊 View Raw Data"):
                                                col1, col2, col3 = st.columns(3)
                                                
                                                with col1:
                                                    st.markdown("**Revenue Data**")
                                                    if stock_comp_data_screener['revenue_df'] is not None:
                                                        st.dataframe(stock_comp_data_screener['revenue_df'], hide_index=True)
                                                    else:
                                                        st.info("No revenue data available")
                                                
                                                with col2:
                                                    st.markdown("**EPS Data**")
                                                    if stock_comp_data_screener['eps_df'] is not None:
                                                        st.dataframe(stock_comp_data_screener['eps_df'], hide_index=True)
                                                    else:
                                                        st.info("No EPS data available")
                                                
                                                with col3:
                                                    st.markdown("**Stock Price Summary**")
                                                    if stock_comp_data_screener['stock_prices_df'] is not None:
                                                        price_df = stock_comp_data_screener['stock_prices_df']
                                                        st.metric("Latest Price", f"₹{price_df['Close'].iloc[-1]:.2f}")
                                                        st.metric("Period Return", f"{((price_df['Close'].iloc[-1] - price_df['Close'].iloc[0]) / price_df['Close'].iloc[0] * 100):.2f}%")
                                                        st.metric("Major Changes", f"{price_df['is_major'].sum()}" if 'is_major' in price_df.columns else "N/A")
                                                    else:
                                                        st.info("No stock price data available")
                                            
                                            st.info(f"💡 Chart shows {years_to_fetch} years of data (Screener Excel mode supports up to 10 years)")
                                        else:
                                            st.warning("Could not generate stock comparison chart. Check if ticker and Excel data are valid.")
                                    
                                    except Exception as e:
                                        st.error(f"Error generating stock comparison: {str(e)}")
                            else:
                                st.warning("Please enter a stock ticker symbol to view price comparison")
                        
                        tab_idx += 1
                    
                    # ================================
                    # SUMMARY
                    # ================================
                    st.markdown("---")
                    st.success("✅ Screener Mode Valuation Complete!")


        # Footer
        st.markdown("---")
        st.caption("💡 **Note:** All values in ₹ Lacs unless specified otherwise | Built with traditional DCF methodology")

# Run main function when executed directly
if __name__ == "__main__":
    main()
