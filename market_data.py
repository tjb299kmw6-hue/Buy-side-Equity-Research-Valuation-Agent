"""Market data layer: Yahoo Finance via yfinance (free, no API key).

Everything returned here is raw or lightly normalised data. No valuation maths
lives in this file (see engine.py) and no LLM calls (see gemini_layer.py).

Currency conventions used throughout the project
------------------------------------------------
* quote currency  : how the price is displayed on Yahoo (e.g. "GBp" = pence)
* trade currency  : major unit of the quote currency (e.g. "GBP")
* report currency : currency of the financial statements (e.g. "USD" for Shell)
* quote_factor    : quote units per major unit (100 for GBp, otherwise 1)
* fx_trade_to_report : 1 unit of trade currency = x units of report currency

All statement figures (revenue, EBITDA, debt, cash ...) are kept in the report
currency. Market cap is kept in the trade currency and converted by the engine.
"""
from __future__ import annotations

import html
import math
import re
import time
from dataclasses import dataclass, field
from datetime import datetime, timezone
from functools import lru_cache

import pandas as pd
import yfinance as yf

try:  # yfinance >= 0.2.50
    from yfinance.exceptions import YFRateLimitError
except Exception:  # pragma: no cover - older yfinance
    class YFRateLimitError(Exception):
        pass

# Quote currencies that are sub-units of a major currency.
SUBUNIT_CURRENCIES = {
    "GBp": ("GBP", 100.0),
    "GBX": ("GBP", 100.0),
    "ZAc": ("ZAR", 100.0),
    "ZAC": ("ZAR", 100.0),
    "ILA": ("ILS", 100.0),
}
NON_EQUITY_TYPES = {"ETF", "MUTUALFUND", "INDEX", "CRYPTOCURRENCY", "CURRENCY", "FUTURE", "OPTION"}
DEFAULT_RISK_FREE = 0.0425  # used only if the live US 10-year yield cannot be fetched
MAX_NEWS = 8


class DataError(Exception):
    """Raised when a company cannot be analysed (bad ticker, missing data)."""


class RateLimited(DataError):
    """Yahoo Finance throttled the request."""


@dataclass
class CompanyData:
    ticker: str
    name: str
    sector: str = ""
    industry: str = ""
    industry_key: str = ""
    country: str = ""
    exchange: str = ""
    website: str = ""
    summary: str = ""
    employees: int | None = None

    quote_currency: str = "USD"
    trade_currency: str = "USD"
    report_currency: str = "USD"
    quote_factor: float = 1.0
    fx_trade_to_report: float = 1.0

    price: float = 0.0          # quote units, as displayed on Yahoo
    price_major: float = 0.0    # trade currency, major units
    market_cap: float = 0.0     # trade currency
    shares: float = 0.0         # effective shares = market cap / price (handles ADRs, share classes)

    # Trailing-twelve-month flows and most-recent balance sheet (report currency)
    revenue_ltm: float | None = None
    ebitda_ltm: float | None = None
    net_income_ltm: float | None = None
    da_ltm: float | None = None
    capex_ltm: float | None = None
    ocf_ltm: float | None = None
    fcf_ltm: float | None = None
    interest_expense_ltm: float | None = None
    total_debt: float | None = None
    cash: float | None = None
    minority_interest: float | None = None
    preferred: float | None = None
    nwc: float | None = None
    ltm_basis: str = "TTM"
    balance_sheet_date: str = ""

    beta: float | None = None
    beta_source: str = ""
    tax_rate: float = 0.25
    tax_rate_source: str = ""
    pe_ttm_yahoo: float | None = None
    pe_fwd_yahoo: float | None = None
    eps_fwd: float | None = None
    consensus_growth: dict = field(default_factory=dict)   # {"0y": 0.08, "+1y": 0.07}
    target_mean: float | None = None
    target_low: float | None = None
    target_high: float | None = None
    analyst_count: int | None = None
    recommendation: str = ""
    week52_low: float | None = None
    week52_high: float | None = None
    revenue_growth_yoy: float | None = None   # Yahoo: most recent quarter, year on year

    hist: pd.DataFrame = field(default_factory=pd.DataFrame)  # rows = line items, cols = FYxxxx (oldest first)
    news: list = field(default_factory=list)
    notes: list = field(default_factory=list)
    as_of: str = ""


# --------------------------------------------------------------------------- helpers
def _num(x) -> float | None:
    try:
        v = float(x)
    except (TypeError, ValueError):
        return None
    return v if math.isfinite(v) else None


def _safe(fn, default=None):
    try:
        out = fn()
        return default if out is None else out
    except Exception:
        return default


def _frame(fn) -> pd.DataFrame:
    """Fetch a statement and return it with the newest period in the first column."""
    df = _safe(fn)
    if not isinstance(df, pd.DataFrame) or df.empty:
        return pd.DataFrame()
    try:
        df = df.sort_index(axis=1, ascending=False)
    except Exception:
        pass
    return df


def _row(df: pd.DataFrame, *labels: str) -> pd.Series | None:
    """First statement row found among `labels`, as a numeric Series."""
    if df is None or df.empty:
        return None
    for lab in labels:
        if lab in df.index:
            s = df.loc[lab]
            if isinstance(s, pd.DataFrame):
                s = s.iloc[0]
            s = pd.to_numeric(s, errors="coerce")
            if s.notna().any():
                return s
    return None


def _latest(df: pd.DataFrame, *labels: str) -> float | None:
    s = _row(df, *labels)
    if s is None:
        return None
    s = s.dropna()
    return _num(s.iloc[0]) if len(s) else None


def _first(*values):
    for v in values:
        v = _num(v)
        if v is not None:
            return v
    return None


def _currency_parts(code: str | None) -> tuple[str, str, float]:
    """-> (quote currency, trade currency, quote units per major unit)."""
    code = code or "USD"
    if code in SUBUNIT_CURRENCIES:
        major, factor = SUBUNIT_CURRENCIES[code]
        return code, major, factor
    return code, code.upper(), 1.0


@lru_cache(maxsize=256)
def _fx_cached(src: str, dst: str, _bucket: int) -> float | None:
    for pair, invert in ((f"{src}{dst}=X", False), (f"{dst}{src}=X", True)):
        rate = _last_price(pair)
        if rate and rate > 0:
            return 1.0 / rate if invert else rate
    return None


def fx_rate(src: str, dst: str) -> float | None:
    """Spot FX: 1 `src` = x `dst`. Cached for an hour."""
    if not src or not dst or src == dst:
        return 1.0
    return _fx_cached(src, dst, int(time.time() // 3600))


def _last_price(symbol: str) -> float | None:
    t = yf.Ticker(symbol)
    price = _safe(lambda: _num(t.fast_info["lastPrice"]))
    if price is None:
        price = _safe(lambda: _num(t.history(period="5d")["Close"].dropna().iloc[-1]))
    return price


def _get_info(t) -> dict:
    try:
        info = t.info
    except YFRateLimitError as exc:
        raise RateLimited("Yahoo Finance is rate-limiting requests right now. Wait a minute and try again.") from exc
    except Exception:
        info = {}
    return info if isinstance(info, dict) else {}


def _price_from(info: dict, t=None) -> float | None:
    price = _first(info.get("currentPrice"), info.get("regularMarketPrice"))
    if price is None and t is not None:
        price = _safe(lambda: _num(t.fast_info["lastPrice"]))
    if price is None:
        price = _first(info.get("previousClose"), info.get("regularMarketPreviousClose"))
    return price


def _market_cap(info: dict, price_major: float) -> tuple[float | None, float | None]:
    """-> (market cap in trade currency, effective share count)."""
    mcap = _num(info.get("marketCap"))
    shares = _first(info.get("sharesOutstanding"), info.get("impliedSharesOutstanding"))
    if mcap and shares and price_major:
        ratio = mcap / (shares * price_major)
        if 50 < ratio < 200:        # market cap delivered in sub-units (pence)
            mcap /= 100.0
        elif 0.005 < ratio < 0.02:  # price delivered in sub-units but flagged as major
            mcap *= 100.0
    if not mcap and shares and price_major:
        mcap = shares * price_major
    if not mcap or not price_major:
        return None, None
    return mcap, mcap / price_major


# --------------------------------------------------------------------------- lookups
def resolve_ticker(query: str) -> tuple[str, list[dict]]:
    """Turn a company name or ticker into a Yahoo symbol.

    Returns (symbol, candidates). Input typed as a ticker takes the exact symbol;
    otherwise the top equity result from Yahoo's search is used. The other
    matches are offered in the UI so the user can switch listing.
    """
    q = (query or "").strip()
    if not q:
        raise DataError("Enter a company name or ticker.")
    candidates: list[dict] = []
    try:
        res = yf.Search(q, max_results=8, news_count=0, lists_count=0, include_cb=False)
        for item in res.quotes or []:
            if item.get("quoteType") == "EQUITY" and item.get("symbol"):
                candidates.append({
                    "symbol": item["symbol"],
                    "name": item.get("longname") or item.get("shortname") or item["symbol"],
                    "exchange": item.get("exchDisp") or item.get("exchange") or "",
                })
    except Exception:
        candidates = []
    # Typed like a ticker (ALL CAPS, or with a suffix/digit such as shel.l or 7203.T): take the exact symbol.
    # Typed like a name ("Ford"): trust Yahoo's ranking, so a small company whose ticker is FORD does not win.
    looks_like_ticker = q == q.upper() or any(ch in q for ch in ".-") or any(ch.isdigit() for ch in q)
    if looks_like_ticker:
        for c in candidates:
            if c["symbol"].upper() == q.upper():
                return c["symbol"], candidates
    if candidates:
        return candidates[0]["symbol"], candidates
    if re.fullmatch(r"[A-Za-z0-9.\-]{1,12}", q):
        return q.upper(), []
    raise DataError(f"No listed company found for '{q}'. Try the ticker symbol (e.g. AAPL, SHEL.L).")


def fetch_risk_free() -> tuple[float, str]:
    """US 10-year Treasury yield (^TNX). Falls back to a labelled default."""
    y = _safe(lambda: _last_price("^TNX"))
    if y and 0.1 < y < 20:
        return y / 100.0, "US 10-year Treasury yield (Yahoo ^TNX, live)"
    return DEFAULT_RISK_FREE, "Default 4.25% (live US 10-year yield unavailable)"


def fetch_industry_peers(industry_key: str, exclude: str, limit: int = 6) -> list[str]:
    """Fallback peer list: largest companies in the same Yahoo industry."""
    if not industry_key:
        return []
    try:
        top = yf.Industry(industry_key).top_companies
        symbols = [str(s) for s in top.index]
    except Exception:
        return []
    return [s for s in symbols if s.upper() != exclude.upper()][:limit]


# Last-resort peer lists (large, liquid companies per Yahoo sector). Still validated against Yahoo before use.
SECTOR_PEERS = {
    "Technology": ["AAPL", "MSFT", "NVDA", "AVGO", "ORCL", "SAP", "CSCO", "IBM"],
    "Communication Services": ["GOOGL", "META", "NFLX", "DIS", "TMUS", "VZ", "T", "CMCSA"],
    "Consumer Cyclical": ["AMZN", "TSLA", "HD", "MCD", "NKE", "LOW", "TM", "BKNG"],
    "Consumer Defensive": ["WMT", "COST", "PG", "KO", "PEP", "PM", "MDLZ", "UL"],
    "Healthcare": ["LLY", "JNJ", "UNH", "ABBV", "MRK", "NVS", "AZN", "PFE"],
    "Financial Services": ["JPM", "BAC", "V", "MA", "WFC", "GS", "MS", "HSBC"],
    "Industrials": ["GE", "CAT", "RTX", "HON", "UNP", "DE", "BA", "SIE.DE"],
    "Energy": ["XOM", "CVX", "SHEL", "TTE", "COP", "BP", "EOG", "SLB"],
    "Basic Materials": ["LIN", "BHP", "RIO", "SHW", "FCX", "SCCO", "NEM", "APD"],
    "Utilities": ["NEE", "SO", "DUK", "CEG", "SRE", "AEP", "D", "NGG"],
    "Real Estate": ["PLD", "AMT", "EQIX", "WELL", "SPG", "PSA", "O", "CCI"],
}


# --------------------------------------------------------------------------- news
def _parse_date(value) -> str:
    if value is None:
        return ""
    try:
        if isinstance(value, (int, float)):
            return datetime.fromtimestamp(float(value), tz=timezone.utc).strftime("%Y-%m-%d")
        return pd.to_datetime(str(value), utc=True).strftime("%Y-%m-%d")
    except Exception:
        return ""


def _clean_text(text) -> str:
    text = re.sub(r"<[^>]+>", " ", html.unescape(str(text or "")))
    text = re.sub(r"\s+", " ", text).strip()
    return re.sub(r"\s+([.,;:!?])", r"\1", text)


def normalise_news(items, limit: int = MAX_NEWS) -> list[dict]:
    """Flatten yfinance news (old and new formats) into simple dicts."""
    out = []
    for item in items or []:
        if not isinstance(item, dict):
            continue
        c = item.get("content") if isinstance(item.get("content"), dict) else item
        title = _clean_text(c.get("title"))
        if not title:
            continue
        provider = c.get("provider") if isinstance(c.get("provider"), dict) else {}
        url = ""
        for key in ("canonicalUrl", "clickThroughUrl"):
            if isinstance(c.get(key), dict) and c[key].get("url"):
                url = c[key]["url"]
                break
        out.append({
            "title": title,
            "publisher": provider.get("displayName") or c.get("publisher") or "",
            "date": _parse_date(c.get("pubDate") or c.get("displayTime") or c.get("providerPublishTime")),
            "snippet": _clean_text(c.get("summary") or c.get("description"))[:320],
            "url": url or c.get("link") or "",
        })
        if len(out) >= limit:
            break
    return out


_LEGAL = re.compile(r"\b(inc|incorporated|corp|corporation|co|company|plc|ltd|limited|llc|ag|se|sa|nv|n\.v|"
                    r"spa|s\.p\.a|ab|asa|oyj|holdings?|group|the|class [a-z])\b\.?", re.I)
_GENERIC = {"bank", "first", "general", "american", "united", "national", "international", "global", "china",
            "royal", "new", "british", "japan", "the", "energy", "capital", "financial"}


def news_terms(symbol: str, name: str) -> tuple[list[str], str]:
    """-> (name phrases matched case-insensitively, ticker root matched case-sensitively)."""
    cleaned = re.sub(r"\s+", " ", _LEGAL.sub(" ", re.sub(r"[,()&]", " ", name or ""))).strip()
    terms = [cleaned] if len(cleaned) >= 3 else []
    first = cleaned.split(" ")[0] if cleaned else ""
    if len(first) >= 4 and first.lower() not in _GENERIC and first not in terms:
        terms.append(first)
    return terms, symbol.split(".")[0].upper()


def filter_news(items: list[dict], symbol: str, name: str, limit: int = MAX_NEWS) -> list[dict]:
    """Keep only headlines that mention the company. Yahoo's ticker feeds mix in general market stories."""
    terms, root = news_terms(symbol, name)
    pats = [re.compile(r"\b" + re.escape(t) + r"\b", re.I) for t in terms]
    if len(root) >= 2:
        pats.append(re.compile(r"\b" + re.escape(root) + r"\b"))
    seen, out = set(), []
    for n in sorted(items, key=lambda x: x.get("date") or "", reverse=True):
        key = n["title"].lower()
        if key in seen:
            continue
        if any(p.search(n["title"]) or p.search(n.get("snippet") or "") for p in pats):
            seen.add(key)
            out.append(n)
        if len(out) >= limit:
            break
    return out


def _company_news(t, symbol: str, info: dict) -> list[dict]:
    name = info.get("shortName") or info.get("longName") or symbol
    raw = normalise_news(_safe(lambda: t.get_news(count=25), None) or _safe(lambda: t.news, []), limit=25)
    try:
        raw += normalise_news(yf.Search(name, max_results=1, news_count=12, lists_count=0, include_cb=False).news, limit=12)
    except Exception:
        pass
    return filter_news(raw, symbol, info.get("longName") or name)


# --------------------------------------------------------------------------- statements
_HIST_SPEC = {
    # item: (statement, labels)
    "revenue": ("inc", ("Total Revenue", "Operating Revenue")),
    "ebitda": ("inc", ("EBITDA", "Normalized EBITDA")),
    "ebit": ("inc", ("EBIT", "Operating Income")),
    "net_income": ("inc", ("Net Income Common Stockholders", "Net Income")),
    "da_inc": ("inc", ("Reconciled Depreciation",)),
    "da_cf": ("cf", ("Depreciation And Amortization", "Depreciation Amortization Depletion")),
    "capex": ("cf", ("Capital Expenditure",)),
    "ocf": ("cf", ("Operating Cash Flow", "Cash Flow From Continuing Operating Activities")),
    "fcf": ("cf", ("Free Cash Flow",)),
    "tax": ("inc", ("Tax Provision",)),
    "pretax": ("inc", ("Pretax Income",)),
    "interest": ("inc", ("Interest Expense", "Interest Expense Non Operating")),
}


def build_history(inc: pd.DataFrame, cf: pd.DataFrame, max_years: int = 4) -> pd.DataFrame:
    """Annual history table: rows = line items, columns = FYxxxx (oldest first)."""
    rows = {}
    for item, (stmt, labels) in _HIST_SPEC.items():
        s = _row(inc if stmt == "inc" else cf, *labels)
        if s is not None:
            rows[item] = s
    if "revenue" not in rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows).T
    df = df.loc[:, df.loc["revenue"].notna()]
    df = df.sort_index(axis=1).iloc[:, -max_years:]

    def get(name):
        return df.loc[name] if name in df.index else pd.Series(float("nan"), index=df.columns)

    da = get("da_inc").fillna(get("da_cf"))
    capex = get("capex").abs()
    ebitda = get("ebitda").fillna(get("ebit") + da)
    fcf = get("fcf").fillna(get("ocf") - capex)
    out = pd.DataFrame({
        "revenue": get("revenue"), "ebitda": ebitda, "ebit": ebitda - da, "net_income": get("net_income"),
        "da": da, "capex": capex, "ocf": get("ocf"), "fcf": fcf,
        "tax": get("tax"), "pretax": get("pretax"), "interest": get("interest").abs(),
    }).T
    labels, seen = [], set()
    for col in out.columns:
        lab = f"FY{getattr(col, 'year', col)}"
        while lab in seen:
            lab += "'"
        seen.add(lab)
        labels.append(lab)
    out.columns = labels
    return out.astype(float)


def _effective_tax(hist: pd.DataFrame, country: str) -> tuple[float, str]:
    default = 0.21 if country == "United States" else 0.25
    fallback = (default, f"Statutory-style default {default:.0%} (effective rate unavailable or outside 10-40%)")
    if hist.empty or "tax" not in hist.index or "pretax" not in hist.index:
        return fallback
    tax, pretax = hist.loc["tax"], hist.loc["pretax"]
    mask = pretax > 0
    if not mask.any() or tax[mask].isna().all():
        return fallback
    rate = float((tax[mask] / pretax[mask]).dropna().tail(3).mean())
    if not math.isfinite(rate) or rate < 0.10 or rate > 0.40:
        return fallback
    return rate, "Effective rate: tax provision / pretax income, avg of last 3 fiscal years"


# Local benchmark for non-US listings. Yahoo's own beta for these is often measured against the
# S&P 500 across currencies, which can give nonsense (e.g. a negative beta for Shell on the LSE).
LOCAL_INDEX = {
    "L": ("^FTSE", "FTSE 100"), "IL": ("^FTSE", "FTSE 100"), "DE": ("^GDAXI", "DAX"), "F": ("^GDAXI", "DAX"),
    "PA": ("^FCHI", "CAC 40"), "AS": ("^AEX", "AEX"), "MC": ("^IBEX", "IBEX 35"), "MI": ("FTSEMIB.MI", "FTSE MIB"),
    "SW": ("^SSMI", "SMI"), "BR": ("^BFX", "BEL 20"), "ST": ("^OMX", "OMX Stockholm 30"), "CO": ("^OMXC25", "OMX Copenhagen 25"),
    "OL": ("OBX.OL", "OBX"), "HE": ("^OMXH25", "OMX Helsinki 25"), "T": ("^N225", "Nikkei 225"), "HK": ("^HSI", "Hang Seng"),
    "TO": ("^GSPTSE", "S&P/TSX"), "V": ("^GSPTSE", "S&P/TSX"), "AX": ("^AXJO", "ASX 200"), "NS": ("^NSEI", "Nifty 50"),
    "BO": ("^BSESN", "Sensex"), "KS": ("^KS11", "KOSPI"), "SS": ("000001.SS", "SSE Composite"), "SZ": ("399001.SZ", "SZSE Component"),
    "SA": ("^BVSP", "Bovespa"), "MX": ("^MXX", "IPC Mexico"), "SI": ("^STI", "Straits Times"), "JO": ("^J203.JO", "JSE All Share"),
}
BETA_PLAUSIBLE = (0.2, 3.0)


def benchmark_for(symbol: str) -> tuple[str, str]:
    suffix = symbol.rsplit(".", 1)[1].upper() if "." in symbol else ""
    return LOCAL_INDEX.get(suffix, ("^GSPC", "S&P 500"))


def choose_beta(symbol: str, yahoo_beta: float | None, regress) -> tuple[float, str, list[str]]:
    """Raw (unadjusted) beta with its source. `regress(index_symbol)` returns a regression beta or None.

    US listings: Yahoo beta if plausible, else a 5-year regression on the S&P 500.
    Non-US listings: 5-year regression on the local index first, because Yahoo's figure is unreliable there.
    """
    lo, hi = BETA_PLAUSIBLE
    ok = lambda b: b is not None and lo <= b <= hi   # noqa: E731
    index, index_name = benchmark_for(symbol)
    notes: list[str] = []
    local = index != "^GSPC"
    if not local and ok(yahoo_beta):
        return yahoo_beta, "Yahoo Finance (5-year monthly vs S&P 500)", notes
    reg = regress(index)
    if ok(reg):
        if local and yahoo_beta is not None and not ok(yahoo_beta):
            notes.append(f"Yahoo beta of {yahoo_beta:.2f} looks unreliable; regression beta vs the {index_name} used instead.")
        return reg, f"Calculated: 5-year monthly returns vs {index_name}", notes
    if ok(yahoo_beta):
        return yahoo_beta, "Yahoo Finance (5-year monthly)", notes
    shown = f"{yahoo_beta:.2f}" if yahoo_beta is not None else "none"
    notes.append(f"No reliable beta available (Yahoo: {shown}); market beta of 1.0 assumed.")
    return 1.0, "Default 1.0 (no reliable beta available)", notes


def _beta_from_prices(t, index: str = "^GSPC") -> float | None:
    """5-year monthly regression beta of the stock against `index`."""
    try:
        px = t.history(period="5y", interval="1mo")["Close"].pct_change().dropna()
        mk = yf.Ticker(index).history(period="5y", interval="1mo")["Close"].pct_change().dropna()
        px.index, mk.index = px.index.tz_localize(None).to_period("M"), mk.index.tz_localize(None).to_period("M")
        both = pd.concat([px, mk], axis=1, join="inner").dropna()
        if len(both) < 24:
            return None
        return _num(both.iloc[:, 0].cov(both.iloc[:, 1]) / both.iloc[:, 1].var())
    except Exception:
        return None


def _consensus_growth(t) -> dict:
    out = {}
    est = _safe(lambda: t.revenue_estimate)
    if isinstance(est, pd.DataFrame) and "growth" in est.columns:
        for key in ("0y", "+1y"):
            if key in est.index:
                g = _num(est.loc[key, "growth"])
                if g is not None and -0.9 < g < 3:
                    out[key] = g
    return out


# --------------------------------------------------------------------------- main fetch
def fetch_company(symbol: str) -> CompanyData:
    """Pull everything the engine needs for one company."""
    t = yf.Ticker(symbol)
    info = _get_info(t)
    qtype = str(info.get("quoteType") or "").upper()
    if qtype in NON_EQUITY_TYPES:
        raise DataError(f"{symbol} is a {qtype.lower()}, not a listed company. Enter a common-stock ticker.")

    price = _price_from(info, t)
    if not price:
        raise DataError(f"No market price found for '{symbol}'. Check the ticker (non-US tickers need a suffix, e.g. SHEL.L).")

    quote_ccy, trade_ccy, factor = _currency_parts(info.get("currency"))
    report_ccy = (info.get("financialCurrency") or trade_ccy).upper()
    price_major = price / factor
    notes: list[str] = []

    fx = fx_rate(trade_ccy, report_ccy)
    if fx is None:
        raise DataError(f"Could not fetch the {trade_ccy}/{report_ccy} exchange rate needed to value {symbol}.")
    if report_ccy != trade_ccy:
        notes.append(f"Shares trade in {trade_ccy} but the company reports in {report_ccy}; "
                     f"market cap converted at spot ({trade_ccy}/{report_ccy} = {fx:.4f}).")

    mcap, shares = _market_cap(info, price_major)
    if not mcap:
        raise DataError(f"No market capitalisation available for '{symbol}'.")

    # Statements --------------------------------------------------------------
    inc, bs, cf = _frame(lambda: t.income_stmt), _frame(lambda: t.balance_sheet), _frame(lambda: t.cash_flow)
    qbs = _frame(lambda: t.quarterly_balance_sheet)
    ttm_inc, ttm_cf = _frame(lambda: t.ttm_income_stmt), _frame(lambda: t.ttm_cash_flow)
    q_inc = _frame(lambda: t.quarterly_income_stmt) if ttm_inc.empty else pd.DataFrame()
    q_cf = _frame(lambda: t.quarterly_cash_flow) if ttm_cf.empty else pd.DataFrame()

    bases: set[str] = set()

    def ttm(stmt: str, *labels: str) -> float | None:
        ttm_df, q_df, a_df = (ttm_inc, q_inc, inc) if stmt == "inc" else (ttm_cf, q_cf, cf)
        v = _latest(ttm_df, *labels)
        if v is not None:
            bases.add("TTM")
            return v
        s = _row(q_df, *labels)
        if s is not None and len(s) >= 4 and s.iloc[:4].notna().all():
            bases.add("TTM")
            return float(s.iloc[:4].sum())
        v = _latest(a_df, *labels)
        if v is not None:
            bases.add("FY")
        return v

    revenue = _first(ttm("inc", "Total Revenue", "Operating Revenue"), info.get("totalRevenue"))
    if not revenue or revenue <= 0:
        raise DataError(f"No revenue data available for '{symbol}' on Yahoo Finance, so it cannot be valued.")

    da = _first(ttm("inc", "Reconciled Depreciation"),
                ttm("cf", "Depreciation And Amortization", "Depreciation Amortization Depletion"))
    ebitda = ttm("inc", "EBITDA", "Normalized EBITDA")
    if ebitda is None:
        ebit = ttm("inc", "EBIT", "Operating Income")
        ebitda = ebit + da if ebit is not None and da is not None else _num(info.get("ebitda"))
    net_income = _first(ttm("inc", "Net Income Common Stockholders", "Net Income"), info.get("netIncomeToCommon"))
    interest = ttm("inc", "Interest Expense", "Interest Expense Non Operating")
    ocf = _first(ttm("cf", "Operating Cash Flow", "Cash Flow From Continuing Operating Activities"),
                 info.get("operatingCashflow"))
    capex = ttm("cf", "Capital Expenditure")
    capex = abs(capex) if capex is not None else None
    fcf = ttm("cf", "Free Cash Flow")
    if fcf is None and ocf is not None and capex is not None:
        fcf = ocf - capex

    mrq = qbs if not qbs.empty else bs
    total_debt = _first(_latest(mrq, "Total Debt"), info.get("totalDebt"), _latest(bs, "Total Debt"))
    cash = _first(_latest(mrq, "Cash Cash Equivalents And Short Term Investments", "Cash And Cash Equivalents"),
                  info.get("totalCash"),
                  _latest(bs, "Cash Cash Equivalents And Short Term Investments", "Cash And Cash Equivalents"))
    minority = _first(_latest(mrq, "Minority Interest"), _latest(bs, "Minority Interest"))
    preferred = _first(_latest(mrq, "Preferred Stock", "Preferred Stock Equity"), _latest(bs, "Preferred Stock"))
    cur_assets, cur_liabs = _latest(mrq, "Current Assets"), _latest(mrq, "Current Liabilities")
    cur_debt = _first(_latest(mrq, "Current Debt And Capital Lease Obligation", "Current Debt"), 0.0)
    nwc = None
    if cur_assets is not None and cur_liabs is not None:
        nwc = (cur_assets - (cash or 0.0)) - (cur_liabs - cur_debt)
    bs_date = ""
    if not mrq.empty:
        bs_date = _safe(lambda: pd.Timestamp(mrq.columns[0]).strftime("%Y-%m-%d"), "")

    if total_debt is None:
        total_debt = 0.0
        notes.append("Total debt not reported on Yahoo; assumed zero.")
    if cash is None:
        cash = 0.0
        notes.append("Cash not reported on Yahoo; assumed zero.")
    if "FY" in bases:
        notes.append("Some trailing-twelve-month figures were unavailable; latest fiscal-year values used instead.")

    hist = build_history(inc, cf)
    tax_rate, tax_src = _effective_tax(hist, info.get("country") or "")

    beta, beta_src, beta_notes = choose_beta(symbol, _num(info.get("beta")), lambda idx: _beta_from_prices(t, idx))
    notes += beta_notes

    return CompanyData(
        ticker=symbol.upper(),
        name=info.get("longName") or info.get("shortName") or symbol.upper(),
        sector=info.get("sector") or "", industry=info.get("industry") or "",
        industry_key=info.get("industryKey") or "", country=info.get("country") or "",
        exchange=info.get("fullExchangeName") or info.get("exchange") or "",
        website=info.get("website") or "", summary=_clean_text(info.get("longBusinessSummary")),
        employees=info.get("fullTimeEmployees"),
        quote_currency=quote_ccy, trade_currency=trade_ccy, report_currency=report_ccy,
        quote_factor=factor, fx_trade_to_report=fx,
        price=price, price_major=price_major, market_cap=mcap, shares=shares,
        revenue_ltm=revenue, ebitda_ltm=ebitda, net_income_ltm=net_income, da_ltm=da, capex_ltm=capex,
        ocf_ltm=ocf, fcf_ltm=fcf, interest_expense_ltm=abs(interest) if interest is not None else None,
        total_debt=total_debt, cash=cash, minority_interest=minority or 0.0, preferred=preferred or 0.0, nwc=nwc,
        ltm_basis="Latest fiscal year" if bases == {"FY"} else "Trailing twelve months",
        balance_sheet_date=bs_date,
        beta=beta, beta_source=beta_src, tax_rate=tax_rate, tax_rate_source=tax_src,
        pe_ttm_yahoo=_num(info.get("trailingPE")), pe_fwd_yahoo=_num(info.get("forwardPE")),
        eps_fwd=_num(info.get("forwardEps")),
        consensus_growth=_consensus_growth(t),
        target_mean=_num(info.get("targetMeanPrice")), target_low=_num(info.get("targetLowPrice")),
        target_high=_num(info.get("targetHighPrice")), analyst_count=info.get("numberOfAnalystOpinions"),
        recommendation=str(info.get("recommendationKey") or "").replace("_", " ").title(),
        week52_low=_num(info.get("fiftyTwoWeekLow")), week52_high=_num(info.get("fiftyTwoWeekHigh")),
        revenue_growth_yoy=_num(info.get("revenueGrowth")),
        hist=hist, news=_company_news(t, symbol, info), notes=notes,
        as_of=datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC"),
    )


def fetch_peer(symbol: str) -> dict | None:
    """Light snapshot for one peer (single Yahoo call). None if the ticker is unusable."""
    symbol = symbol.strip().upper()
    if not symbol:
        return None
    try:
        info = _get_info(yf.Ticker(symbol))
    except RateLimited:
        return None
    if str(info.get("quoteType") or "").upper() in NON_EQUITY_TYPES:
        return None
    price = _price_from(info)
    if not price:
        return None
    _, trade_ccy, factor = _currency_parts(info.get("currency"))
    report_ccy = (info.get("financialCurrency") or trade_ccy).upper()
    fx = fx_rate(trade_ccy, report_ccy)
    mcap, _shares = _market_cap(info, price / factor)
    if not mcap or fx is None:
        return None
    return {
        "ticker": symbol,
        "name": info.get("shortName") or info.get("longName") or symbol,
        "currency": report_ccy,
        "market_cap": mcap * fx,
        "total_debt": _num(info.get("totalDebt")) or 0.0,
        "cash": _num(info.get("totalCash")) or 0.0,
        "revenue": _num(info.get("totalRevenue")),
        "ebitda": _num(info.get("ebitda")),
        "net_income": _num(info.get("netIncomeToCommon")),
        "pe_yahoo": _num(info.get("trailingPE")),
        "fwd_pe": _num(info.get("forwardPE")),
        "revenue_growth": _num(info.get("revenueGrowth")),
        "same_currency": report_ccy == trade_ccy,
    }
