"""Offline tests: financial maths against hand calculations, Yahoo parsing against a fake
yfinance, the Gemini layer against a fake client, and the Excel writer end to end.

Run:  python -m pytest -q
No network or API key needed.
"""
from __future__ import annotations

import io
import math
import os
import sys
from dataclasses import replace

import pandas as pd
import pytest
from openpyxl import load_workbook

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import demo_data  # noqa: E402
import engine  # noqa: E402
import excel_memo  # noqa: E402
import gemini_layer as gl  # noqa: E402
import market_data as md  # noqa: E402
from engine import Assumptions, dcf_value, lbo_case, run_all  # noqa: E402

M = 1e6


@pytest.fixture(autouse=True)
def _clear_gemini_cooldown():
    gl._COOLDOWN.clear()
    yield
    gl._COOLDOWN.clear()


@pytest.fixture
def demo():
    return demo_data.demo_company()


# ------------------------------------------------------------------ market metrics and WACC
def test_market_metrics_hand_calc(demo):
    m = engine.compute_market(demo)
    assert m["market_cap"] == pytest.approx(84.50 * 250 * M)
    assert m["net_debt"] == pytest.approx((4200 - 1150) * M)
    assert m["ev"] == pytest.approx((21125 + 4200 - 1150 + 60) * M)          # 24,235
    assert m["pe"] == pytest.approx(21125 / 1010)
    assert m["ev_ebitda"] == pytest.approx(24235 / 2060)
    assert m["fcf_yield"] == pytest.approx(1150 / 21125)
    assert m["eps_ltm"] == pytest.approx(1010 / 250)


def test_wacc_hand_calc(demo):
    a = Assumptions(risk_free=0.04, erp=0.05)
    w = engine.compute_wacc(demo, a, engine.compute_market(demo))
    ke = 0.04 + (0.67 * 1.08 + 0.33) * 0.05          # Blume-adjusted beta
    kd = 0.05                                   # implied 180/4200 = 4.29% is floored at risk-free + 1%
    we = 21125 / (21125 + 4200)
    assert w["cost_of_equity"] == pytest.approx(ke)
    assert w["pre_tax_cost_of_debt"] == pytest.approx(kd)
    assert w["wacc"] == pytest.approx(we * ke + (1 - we) * kd * (1 - 0.233))


def test_wacc_overrides(demo):
    a = Assumptions(risk_free=0.03, erp=0.06, beta=1.5, cost_of_debt=0.08, tax_rate=0.30)
    w = engine.compute_wacc(demo, a, engine.compute_market(demo))
    we = 21125 / 25325
    assert w["wacc"] == pytest.approx(we * (0.03 + 1.5 * 0.06) + (1 - we) * 0.08 * 0.70)


# ------------------------------------------------------------------ projections and DCF
def test_projection_growth_fade_and_fcf(demo):
    a = Assumptions(growth_y1=0.10, terminal_growth=0.02)
    p = engine.build_projections(demo, a, tax=0.25)
    t, drv = p["table"], p["drivers"]
    assert list(t["growth"].round(6)) == [0.10, 0.084, 0.068, 0.052, 0.036]   # straight-line fade
    assert t.loc[1, "revenue"] == pytest.approx(9800 * M * 1.10)
    row = t.loc[3]
    ebit = row["ebitda"] - row["da"]
    assert row["ufcf"] == pytest.approx(ebit * 0.75 + row["da"] - row["capex"] - row["d_nwc"])
    assert row["d_nwc"] == pytest.approx((row["revenue"] - t.loc[2, "revenue"]) * drv["nwc_pct"])
    assert drv["nwc_pct"] == pytest.approx(1180 / 9800)


def _flows(ufcf, ebitda=None):
    n = len(ufcf)
    return pd.DataFrame({"ufcf": ufcf, "ebitda": ebitda or [x * 2 for x in ufcf]}, index=range(1, n + 1), dtype=float)


def test_dcf_end_year_hand_calc():
    v = dcf_value(_flows([100, 110, 120, 130, 140]), wacc=0.10, g=0.02, mid_year=False)
    pv = sum(cf / 1.1 ** t for t, cf in enumerate([100, 110, 120, 130, 140], start=1))
    tv = 140 * 1.02 / 0.08
    assert v["sum_pv"] == pytest.approx(pv) == pytest.approx(447.697, abs=1e-3)
    assert v["tv"] == pytest.approx(1785.0)
    assert v["ev"] == pytest.approx(pv + tv / 1.1 ** 5) == pytest.approx(1556.04, abs=1e-2)


def test_dcf_mid_year_is_end_year_times_half_period():
    flows = _flows([100, 110, 120, 130, 140])
    end, mid = dcf_value(flows, 0.10, 0.02, False), dcf_value(flows, 0.10, 0.02, True)
    assert mid["ev"] == pytest.approx(end["ev"] * 1.1 ** 0.5)


def test_dcf_rejects_growth_at_or_above_wacc():
    assert dcf_value(_flows([100] * 5), wacc=0.03, g=0.03, mid_year=False) is None


def test_dcf_bridge_and_sensitivity(demo):
    r = run_all(demo, Assumptions(), demo_data.demo_peers())
    dcf, m = r["dcf"], r["market"]
    assert dcf["equity"] == pytest.approx(dcf["ev"] - m["net_debt"] - m["minority_interest"] - m["preferred"])
    assert dcf["price"] == pytest.approx(dcf["equity"] / (250 * M))
    grid = dcf["sensitivity"]
    assert grid.iloc[2, 2] == pytest.approx(dcf["price"])                 # centre = base case
    assert grid.iloc[0, 4] == grid.to_numpy().max()                       # lowest WACC, highest growth
    assert grid.iloc[4, 0] == grid.to_numpy().min()


# ------------------------------------------------------------------ LBO
def _flat_ops(ebitda=100.0, years=5):
    return pd.DataFrame({"ebitda": ebitda, "ebit": ebitda, "capex": 0.0, "d_nwc": 0.0}, index=range(1, years + 1))


def test_lbo_hand_calc():
    """Flat EBITDA 100, 4x debt at 10%, bought and sold at 10x, no tax: asset yield = cost of debt, so IRR = 10%."""
    c = lbo_case(_flat_ops(), ebitda0=100, entry_ev=1000, leverage=4.0, rate=0.10, exit_multiple=10.0, tax=0.0)
    assert c["debt0"] == 400 and c["equity0"] == 600
    assert list(c["schedule"]["debt_end"].round(3)) == [340.0, 274.0, 201.4, 121.54, 33.694]
    assert c["exit_equity"] == pytest.approx(1000 - 33.694)
    assert c["moic"] == pytest.approx(1.1 ** 5)
    assert c["irr"] == pytest.approx(0.10)


def test_lbo_interest_tax_shield():
    c = lbo_case(_flat_ops(), 100, 1000, 4.0, 0.10, 10.0, tax=0.25)
    y1 = c["schedule"].loc[1]
    assert y1["interest"] == pytest.approx(40.0)
    assert y1["taxes"] == pytest.approx((100 - 40) * 0.25)
    assert y1["fcf"] == pytest.approx(100 - 40 - 15)


def test_lbo_cash_builds_after_debt_is_repaid():
    c = lbo_case(_flat_ops(), 100, 1000, leverage=1.0, rate=0.10, exit_multiple=10.0, tax=0.0)
    assert c["exit_debt"] == pytest.approx(0.0)
    assert c["exit_cash"] > 0
    assert c["exit_equity"] == pytest.approx(1000 + c["exit_cash"])


def test_lbo_debt_capped_by_loan_to_value():
    c = lbo_case(_flat_ops(), 100, entry_ev=400, leverage=4.5, rate=0.07, exit_multiple=4.0, tax=0.0, max_ltv=0.75)
    assert c["debt0"] == pytest.approx(300) and c["leverage_used"] == pytest.approx(3.0)


def test_lbo_uses_specified_inputs_and_max_price_hits_target(demo):
    a = Assumptions()
    assert (a.lbo_leverage, a.lbo_rate, a.lbo_target_irr, engine.PROJECTION_YEARS) == (4.5, 0.07, 0.20, 5)
    r = run_all(demo, a, [])
    lbo = r["lbo"]
    assert lbo["debt0"] == pytest.approx(4.5 * 2060 * M)
    assert lbo["exit_multiple"] == pytest.approx(r["market"]["ev_ebitda"])       # exit at base multiple
    assert lbo["schedule"].loc[1, "interest"] == pytest.approx(lbo["debt0"] * 0.07)
    check = lbo_case(r["projections"]["table"], demo.ebitda_ltm, lbo["max_entry_ev"], 4.5, 0.07,
                     lbo["exit_multiple"], r["wacc"]["tax_rate"])
    assert check["irr"] == pytest.approx(0.20, abs=1e-6)
    assert lbo["meets_target"] == (lbo["irr"] >= 0.20)
    grid = lbo["sens_entry_exit"]
    assert grid.iloc[2, 2] == pytest.approx(lbo["irr"])
    assert (grid.diff(axis=1).iloc[:, 1:] > 0).all().all()       # higher exit multiple -> higher IRR
    assert (grid.diff(axis=0).iloc[1:, :] < 0).all().all()       # higher entry multiple -> lower IRR


# ------------------------------------------------------------------ peers and edge cases
def test_peer_multiples_and_implied_value(demo):
    r = run_all(demo, Assumptions(), demo_data.demo_peers())
    p = r["peers"]
    first = p["table"].iloc[0]
    assert first["ev"] == pytest.approx((31000 + 5200 - 1900) * M)
    assert first["ev_ebitda"] == pytest.approx(34300 / 2750)
    med = p["stats"]["ev_ebitda"]["median"]
    claims = (3050 + 60) * M
    assert p["implied"]["ev_ebitda"]["base"] == pytest.approx((med * 2060 * M - claims) / (250 * M))
    assert p["implied"]["pe"]["base"] == pytest.approx(p["stats"]["pe"]["median"] * 1010 / 250)


def test_pence_quote_with_usd_reporting(demo):
    """LSE-style listing: price in pence, statements in USD. Same business, same USD equity value."""
    fx = 1.25                                   # 1 GBP = 1.25 USD
    gbp_price = 84.50 / fx
    uk = replace(demo, quote_currency="GBp", trade_currency="GBP", quote_factor=100.0, fx_trade_to_report=fx,
                 price=gbp_price * 100, price_major=gbp_price, market_cap=gbp_price * 250 * M)
    us, gb = run_all(demo, Assumptions(), []), run_all(uk, Assumptions(), [])
    assert gb["market"]["market_cap"] == pytest.approx(us["market"]["market_cap"])
    assert gb["market"]["pe"] == pytest.approx(us["market"]["pe"])
    assert gb["dcf"]["price"] == pytest.approx(us["dcf"]["price"] / fx * 100)      # pence
    assert gb["dcf"]["upside"] == pytest.approx(us["dcf"]["upside"])
    assert gb["lbo"]["max_price"] == pytest.approx(us["lbo"]["max_price"] / fx * 100)


def test_company_without_ebitda_still_produces_a_memo(demo):
    bank = replace(demo, ebitda_ltm=None, da_ltm=None, capex_ltm=None, sector="Financial Services")
    r = run_all(bank, Assumptions(), [])
    assert not r["dcf"]["available"] and not r["lbo"]["available"]
    assert r["market"]["pe"] is not None and r["market"]["ev_ebitda"] is None
    xlsx = excel_memo.build_workbook(bank, r, Assumptions(), None, gl.merge_news(bank, None), {})
    assert load_workbook(io.BytesIO(xlsx)).sheetnames[0] == "Memo"


def test_loss_making_company(demo):
    loss = replace(demo, ebitda_ltm=-300 * M, net_income_ltm=-500 * M, fcf_ltm=-400 * M)
    r = run_all(loss, Assumptions(), demo_data.demo_peers())
    assert r["market"]["pe"] is None and r["market"]["ev_ebitda"] is None
    assert not r["lbo"]["available"]
    assert not r["dcf"]["available"]            # negative terminal cash flow
    excel_memo.build_workbook(loss, r, Assumptions(), None, [], {})


# ------------------------------------------------------------------ Excel output
def test_workbook_contents(demo):
    a = Assumptions()
    r = run_all(demo, a, demo_data.demo_peers())
    memo = demo_data.demo_memo()
    xlsx = excel_memo.build_workbook(demo, r, a, memo, gl.merge_news(demo, memo), {})
    wb = load_workbook(io.BytesIO(xlsx))
    assert wb.sheetnames == ["Memo", "Valuation", "DCF", "LBO", "Peers", "Projections", "News & Risks", "Assumptions"]
    values = {ws.title: [c.value for row in ws.iter_rows() for c in row if c.value is not None] for ws in wb}
    assert any(v == pytest.approx(r["dcf"]["price"]) for v in values["Memo"] if isinstance(v, float))
    assert any(v == pytest.approx(r["wacc"]["wacc"]) for v in values["Valuation"] if isinstance(v, float))
    assert any(v == pytest.approx(r["lbo"]["irr"]) for v in values["LBO"] if isinstance(v, float))
    assert 4.5 in values["LBO"] and 0.07 in values["LBO"] and 0.2 in values["LBO"]
    assert "PEERA" in values["Peers"]
    assert any("guidance" in str(v) for v in values["News & Risks"])
    fonts = {c.font.name for ws in wb for row in ws.iter_rows() for c in row if c.value is not None}
    assert fonts == {"Arial"}
    assert not any(isinstance(v, str) and v.startswith("=") for vs in values.values() for v in vs)


def test_text_starting_with_equals_is_not_a_formula(demo):
    memo = demo_data.demo_memo()
    memo["investment_thesis"] = "=HYPERLINK(\"http://example.com\")"
    r = run_all(demo, Assumptions(), [])
    wb = load_workbook(io.BytesIO(excel_memo.build_workbook(demo, r, Assumptions(), memo, [], {})))
    cell = next(c for row in wb["Memo"].iter_rows() for c in row if isinstance(c.value, str) and "HYPERLINK" in c.value)
    assert cell.data_type == "s"


# ------------------------------------------------------------------ Yahoo parsing (fake yfinance)
def _stmt(rows: dict, dates: list[str]) -> pd.DataFrame:
    return pd.DataFrame(rows, index=pd.to_datetime(dates)).T


class FakeTicker:
    """Mimics the parts of yfinance.Ticker this project reads, with realistic shapes and labels."""
    dates = ["2025-09-30", "2024-09-30", "2023-09-30", "2022-09-30"]

    def __init__(self, symbol):
        self.symbol = symbol
        self.info = {
            "quoteType": "EQUITY", "longName": "Acme Tools plc", "currency": "GBp", "financialCurrency": "USD",
            "currentPrice": 2000.0, "marketCap": 20_000_000_000, "sharesOutstanding": 1_000_000_000,
            "sector": "Industrials", "industry": "Tools", "industryKey": "tools", "country": "United Kingdom",
            "beta": 0.9, "trailingPE": 12.5, "forwardPE": 11.0, "targetMeanPrice": 2300.0,
            "fiftyTwoWeekLow": 1500.0, "fiftyTwoWeekHigh": 2200.0, "revenueGrowth": 0.05,
            "longBusinessSummary": "Makes <b>tools</b>.",
        }
        self.fast_info = {"lastPrice": {"GBPUSD=X": 1.25, "^TNX": 4.5}.get(symbol, 2000.0)}
        self.income_stmt = _stmt({
            "Total Revenue": [10e9, 9e9, 8e9, 7e9], "EBITDA": [2.0e9, 1.8e9, 1.6e9, 1.4e9],
            "EBIT": [1.6e9, 1.44e9, 1.28e9, 1.12e9], "Net Income Common Stockholders": [1.0e9, 0.9e9, 0.8e9, 0.7e9],
            "Reconciled Depreciation": [0.4e9, 0.36e9, 0.32e9, 0.28e9], "Tax Provision": [0.3e9, 0.27e9, 0.24e9, 0.21e9],
            "Pretax Income": [1.5e9, 1.35e9, 1.2e9, 1.05e9], "Interest Expense": [0.1e9, 0.09e9, 0.08e9, 0.07e9],
        }, self.dates)
        self.cash_flow = _stmt({
            "Operating Cash Flow": [1.5e9, 1.3e9, 1.2e9, 1.0e9], "Capital Expenditure": [-0.5e9, -0.45e9, -0.4e9, -0.35e9],
            "Free Cash Flow": [1.0e9, 0.85e9, 0.8e9, 0.65e9],
        }, self.dates)
        self.balance_sheet = _stmt({"Total Debt": [3.0e9, 3.1e9, 3.2e9, 3.3e9],
                                    "Cash And Cash Equivalents": [0.9e9, 0.8e9, 0.7e9, 0.6e9]}, self.dates)
        self.quarterly_balance_sheet = _stmt({
            "Total Debt": [2.8e9, 3.0e9], "Cash Cash Equivalents And Short Term Investments": [1.0e9, 0.9e9],
            "Minority Interest": [0.1e9, 0.1e9], "Current Assets": [4.0e9, 3.9e9], "Current Liabilities": [2.5e9, 2.4e9],
            "Current Debt": [0.5e9, 0.5e9],
        }, ["2025-12-31", "2025-09-30"])
        self.ttm_income_stmt = _stmt({"Total Revenue": [10.4e9], "EBITDA": [2.1e9], "Reconciled Depreciation": [0.42e9],
                                      "Net Income Common Stockholders": [1.05e9], "Interest Expense": [0.11e9]}, ["2025-12-31"])
        self.ttm_cash_flow = _stmt({"Operating Cash Flow": [1.6e9], "Capital Expenditure": [-0.52e9],
                                    "Free Cash Flow": [1.08e9]}, ["2025-12-31"])
        self.revenue_estimate = pd.DataFrame({"avg": [1, 1, 1, 1], "growth": [0.03, 0.03, 0.06, 0.08]},
                                             index=["0q", "+1q", "0y", "+1y"])
        self.news = [
            {"id": "1", "content": {"title": "Acme wins contract", "summary": "A <i>large</i> order.",
                                    "pubDate": "2026-11-01T09:30:00Z", "provider": {"displayName": "Wire"},
                                    "canonicalUrl": {"url": "https://example.com/a"}}},
            {"title": "Old format headline", "publisher": "Old Wire", "providerPublishTime": 1762000000,
             "link": "https://example.com/b"},
        ]

    def history(self, *a, **k):
        raise RuntimeError("not needed")


def test_fetch_company_parses_yahoo_shapes(monkeypatch):
    monkeypatch.setattr(md.yf, "Ticker", FakeTicker)
    md._fx_cached.cache_clear()
    d = md.fetch_company("acme.l")
    assert (d.ticker, d.quote_currency, d.trade_currency, d.report_currency) == ("ACME.L", "GBp", "GBP", "USD")
    assert d.quote_factor == 100 and d.fx_trade_to_report == pytest.approx(1.25)
    assert d.price == 2000 and d.price_major == 20 and d.market_cap == 20e9 and d.shares == pytest.approx(1e9)
    assert (d.revenue_ltm, d.ebitda_ltm, d.net_income_ltm) == (10.4e9, 2.1e9, 1.05e9)       # TTM beats annual
    assert (d.capex_ltm, d.fcf_ltm, d.da_ltm) == (0.52e9, 1.08e9, 0.42e9)
    assert (d.total_debt, d.cash, d.minority_interest) == (2.8e9, 1.0e9, 0.1e9)             # latest quarter
    assert d.nwc == pytest.approx((4.0e9 - 1.0e9) - (2.5e9 - 0.5e9))
    assert d.balance_sheet_date == "2025-12-31"
    assert list(d.hist.columns) == ["FY2022", "FY2023", "FY2024", "FY2025"]
    assert d.hist.loc["revenue", "FY2025"] == 10e9 and d.hist.loc["capex", "FY2025"] == 0.5e9
    assert d.tax_rate == pytest.approx(0.20)
    assert d.consensus_growth == {"0y": 0.06, "+1y": 0.08}
    assert d.summary == "Makes tools."
    assert [n["title"] for n in d.news] == ["Acme wins contract"]          # unrelated headline filtered out
    assert d.news[0]["date"] == "2026-11-01" and d.news[0]["snippet"] == "A large order."
    assert d.news[0]["url"] == "https://example.com/a"
    r = run_all(d, Assumptions(risk_free=0.045), [])
    assert r["market"]["market_cap"] == pytest.approx(25e9)                                   # GBP 20bn x 1.25
    assert r["market"]["ev"] == pytest.approx(25e9 + 2.8e9 - 1.0e9 + 0.1e9)
    assert r["dcf"]["available"] and r["lbo"]["available"]


def test_fetch_peer_and_risk_free(monkeypatch):
    monkeypatch.setattr(md.yf, "Ticker", FakeTicker)
    md._fx_cached.cache_clear()
    p = md.fetch_peer("acme.l")
    assert p["market_cap"] == pytest.approx(25e9) and p["currency"] == "USD" and not p["same_currency"]
    assert md.fetch_risk_free()[0] == pytest.approx(0.045)


def test_market_cap_delivered_in_pence_is_corrected():
    mcap, shares = md._market_cap({"marketCap": 2_000_000_000_000, "sharesOutstanding": 1e9}, 20.0)
    assert mcap == pytest.approx(20e9) and shares == pytest.approx(1e9)


def test_non_equity_and_missing_price_are_rejected(monkeypatch):
    class Etf(FakeTicker):
        def __init__(self, s):
            super().__init__(s)
            self.info["quoteType"] = "ETF"
    monkeypatch.setattr(md.yf, "Ticker", Etf)
    with pytest.raises(md.DataError):
        md.fetch_company("SPY")


# ------------------------------------------------------------------ Gemini layer (fake client)
class _FakeModels:
    def __init__(self, outcome):
        self.outcome, self.calls = outcome, []

    def generate_content(self, **kwargs):
        self.calls.append(kwargs)
        if isinstance(self.outcome, Exception):
            raise self.outcome
        return self.outcome


class _Resp:
    def __init__(self, text, parsed=None):
        self.text, self.parsed = text, parsed


def _patch_client(monkeypatch, outcome):
    models = _FakeModels(outcome)
    monkeypatch.setattr(gl.genai, "Client", lambda api_key, **kw: type("C", (), {"models": models})())
    return models


def test_gemini_structured_call(monkeypatch):
    memo = gl.QualitativeMemo(**demo_data.demo_memo())
    models = _patch_client(monkeypatch, _Resp(memo.model_dump_json()))
    out = gl._generate("key", "gemini-3.8-flash", "prompt", gl.QualitativeMemo)
    assert out.stance == "Neutral"
    call = models.calls[0]
    assert call["model"] == "gemini-3.8-flash"
    assert call["config"].response_mime_type == "application/json"
    assert call["config"].response_schema is gl.QualitativeMemo


def test_gemini_429_becomes_retry_message(monkeypatch):
    err = gl.genai_errors.ClientError(429, {"error": {"code": 429, "message": "quota", "status": "RESOURCE_EXHAUSTED",
                                                      "details": [{"@type": "x/RetryInfo", "retryDelay": "34s"}]}})
    _patch_client(monkeypatch, err)
    with pytest.raises(gl.GeminiRateLimit) as exc:
        gl._generate("key", "gemini-3.8-flash", "prompt", gl.PeerList)
    assert exc.value.retry_after == 35 and "429" in str(exc.value) and "35 seconds" in str(exc.value)


def test_gemini_bad_json_is_reported(monkeypatch):
    _patch_client(monkeypatch, _Resp('{"peers": "nope"}'))
    with pytest.raises(gl.GeminiError):
        gl._generate("key", "m", "prompt", gl.PeerList)


def test_api_key_resolution_order(monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY", "from-env")
    assert gl.resolve_api_key("  from-sidebar ") == ("from-sidebar", "sidebar")
    assert gl.resolve_api_key("") == ("from-env", "environment")      # no secrets file in the test run
    monkeypatch.delenv("GEMINI_API_KEY")
    assert gl.resolve_api_key("") == (None, "none")


def test_prompt_inputs_contain_no_raw_numbers_to_compute(demo):
    r = run_all(demo, Assumptions(), demo_data.demo_peers())
    payload = gl.metrics_payload(demo, r)
    assert payload["market"]["pe_ltm"] == "20.9x" and payload["dcf"]["value_per_share"].startswith("USD ")
    assert payload["lbo"]["target_irr"] == "20%"
    flat = gl.to_json(payload)
    assert len(flat) < 4000                                     # small prompt: metrics only, no filings
    heads = gl.headlines_payload(demo)
    assert len(heads) <= md.MAX_NEWS and set(heads[0]) == {"id", "date", "source", "title", "snippet"}


def test_merge_news_keeps_python_headlines(demo):
    memo = demo_data.demo_memo()
    memo["news"].append({"headline_id": 99, "summary": "invented", "sentiment": "Positive"})   # out of range: ignored
    merged = gl.merge_news(demo, memo)
    assert len(merged) == len(demo.news)
    assert merged[0]["title"] == demo.news[0]["title"] and merged[0]["sentiment"] == "Positive"
    assert all(m["summary"] != "invented" for m in merged)


def test_irr_identity():
    assert math.isclose(1.1 ** 5, 1.61051)


# ------------------------------------------------------------------ model fallback
class _PerModel:
    """Fake client whose models answer or fail individually, recording the order they were tried in."""
    def __init__(self, behaviour):
        self.behaviour, self.tried = behaviour, []
        self.models = self

    def generate_content(self, model, contents, config):
        self.tried.append(model)
        out = self.behaviour[model]
        if isinstance(out, Exception):
            raise out
        return out


def _quota():
    return gl.genai_errors.ClientError(429, {"error": {"code": 429, "message": "quota", "status": "RESOURCE_EXHAUSTED"}})


def _ok():
    return _Resp(gl.PeerList(peers=[gl.PeerPick(ticker="X", name="X", rationale="r")]).model_dump_json())


CHAIN = ("gemini-3.5-flash-lite", "gemini-3.8-flash", "gemini-3.6-flash")


def test_model_chain_order_and_override(monkeypatch):
    monkeypatch.delenv("GEMINI_MODEL", raising=False)
    assert gl.model_chain() == CHAIN
    monkeypatch.setenv("GEMINI_MODEL", "gemini-3.6-flash")
    assert gl.model_chain() == ("gemini-3.6-flash", "gemini-3.8-flash")      # no duplicates


def test_fallback_moves_on_only_when_quota_is_used_up(monkeypatch):
    client = _PerModel({CHAIN[0]: _quota(), CHAIN[1]: _quota(), CHAIN[2]: _ok()})
    monkeypatch.setattr(gl.genai, "Client", lambda api_key, **kw: client)
    result, used = gl.generate_with_fallback("k", CHAIN, "p", gl.PeerList)
    assert used == "gemini-3.6-flash" and client.tried == list(CHAIN) and result.peers[0].ticker == "X"


def test_primary_success_never_touches_fallbacks(monkeypatch):
    client = _PerModel({CHAIN[0]: _ok(), CHAIN[1]: _quota(), CHAIN[2]: _quota()})
    monkeypatch.setattr(gl.genai, "Client", lambda api_key, **kw: client)
    assert gl.generate_with_fallback("k", CHAIN, "p", gl.PeerList)[1] == CHAIN[0] and client.tried == [CHAIN[0]]


def test_all_models_exhausted_gives_one_clear_429_message(monkeypatch):
    client = _PerModel({m: _quota() for m in CHAIN})
    monkeypatch.setattr(gl.genai, "Client", lambda api_key, **kw: client)
    with pytest.raises(gl.GeminiRateLimit) as exc:
        gl.generate_with_fallback("k", CHAIN, "p", gl.PeerList)
    assert all(m in str(exc.value) for m in CHAIN) and "own free API key" in str(exc.value)


def test_other_errors_do_not_trigger_fallback(monkeypatch):
    bad_key = gl.genai_errors.ClientError(400, {"error": {"code": 400, "message": "API key not valid", "status": "INVALID_ARGUMENT"}})
    client = _PerModel({CHAIN[0]: bad_key, CHAIN[1]: _ok(), CHAIN[2]: _ok()})
    monkeypatch.setattr(gl.genai, "Client", lambda api_key, **kw: client)
    with pytest.raises(gl.GeminiError):
        gl.generate_with_fallback("k", CHAIN, "p", gl.PeerList)
    assert client.tried == [CHAIN[0]]


def test_unavailable_model_is_skipped(monkeypatch):
    gone = gl.genai_errors.ClientError(404, {"error": {"code": 404, "message": "not found", "status": "NOT_FOUND"}})
    client = _PerModel({CHAIN[0]: gone, CHAIN[1]: _ok(), CHAIN[2]: _ok()})
    monkeypatch.setattr(gl.genai, "Client", lambda api_key, **kw: client)
    assert gl.generate_with_fallback("k", CHAIN, "p", gl.PeerList)[1] == CHAIN[1]


# ------------------------------------------------------------------ live-run fixes: beta and news relevance
def test_non_us_listing_ignores_implausible_yahoo_beta():
    beta, src, notes = md.choose_beta("SHEL.L", -0.25, lambda idx: 0.85 if idx == "^FTSE" else None)
    assert beta == 0.85 and "FTSE 100" in src and "unreliable" in notes[0]


def test_us_listing_keeps_plausible_yahoo_beta_and_defaults_safely():
    assert md.choose_beta("AAPL", 1.1, lambda idx: 0.5)[0] == 1.1
    assert md.choose_beta("AAPL", -0.3, lambda idx: 1.25)[0] == 1.25
    beta, src, notes = md.choose_beta("XYZ.L", -0.3, lambda idx: None)
    assert beta == 1.0 and notes


def test_blume_adjustment_and_user_override(demo):
    m = engine.compute_market(demo)
    w = engine.compute_wacc(replace(demo, beta=0.4), Assumptions(), m)
    assert w["beta"] == pytest.approx(0.67 * 0.4 + 0.33) and w["beta_raw"] == 0.4
    assert engine.compute_wacc(demo, Assumptions(beta=0.4), m)["beta"] == 0.4


def test_news_filter_keeps_only_company_stories():
    items = [
        {"title": "Jensen Huang vs. Steve Jobs: can Nvidia thrive?", "snippet": "", "date": "2026-11-12"},
        {"title": "Tony Robbins: 4 rules investors follow", "snippet": "", "date": "2026-11-12"},
        {"title": "Apple unveils new iPhone", "snippet": "", "date": "2026-11-10"},
        {"title": "Big tech earnings recap", "snippet": "AAPL beat estimates", "date": "2026-11-11"},
        {"title": "Apple unveils new iPhone", "snippet": "duplicate", "date": "2026-11-09"},
        {"title": "Pineapple prices rise", "snippet": "", "date": "2026-11-13"},
    ]
    kept = md.filter_news(items, "AAPL", "Apple Inc.")
    assert [n["title"] for n in kept] == ["Big tech earnings recap", "Apple unveils new iPhone"]


def test_news_terms_avoid_generic_words():
    assert md.news_terms("SHEL.L", "Shell plc") == (["Shell"], "SHEL")
    terms, _ = md.news_terms("BAC", "Bank of America Corporation")
    assert terms == ["Bank of America"]


def test_overloaded_model_falls_back(monkeypatch):
    busy = gl.genai_errors.ServerError(503, {"error": {"code": 503, "message": "high demand", "status": "UNAVAILABLE"}})
    client = _PerModel({CHAIN[0]: busy, CHAIN[1]: _ok(), CHAIN[2]: _ok()})
    monkeypatch.setattr(gl.genai, "Client", lambda api_key, **kw: client)
    assert gl.generate_with_fallback("k", CHAIN, "p", gl.PeerList)[1] == CHAIN[1]


def test_all_models_overloaded_says_retry(monkeypatch):
    busy = gl.genai_errors.ServerError(503, {"error": {"code": 503, "message": "high demand", "status": "UNAVAILABLE"}})
    client = _PerModel({m: busy for m in CHAIN})
    monkeypatch.setattr(gl.genai, "Client", lambda api_key, **kw: client)
    with pytest.raises(gl.GeminiBusy) as exc:
        gl.generate_with_fallback("k", CHAIN, "p", gl.PeerList)
    assert "Retry" in str(exc.value)


def test_timeout_falls_back(monkeypatch):
    class ReadTimeout(Exception):
        pass
    client = _PerModel({CHAIN[0]: ReadTimeout("The read operation timed out"), CHAIN[1]: _ok(), CHAIN[2]: _ok()})
    monkeypatch.setattr(gl.genai, "Client", lambda api_key, **kw: client)
    assert gl.generate_with_fallback("k", CHAIN, "p", gl.PeerList)[1] == CHAIN[1]


def test_busy_model_is_skipped_on_the_next_call(monkeypatch):
    busy = gl.genai_errors.ServerError(503, {"error": {"code": 503, "message": "high demand", "status": "UNAVAILABLE"}})
    client = _PerModel({CHAIN[0]: busy, CHAIN[1]: _ok(), CHAIN[2]: _ok()})
    monkeypatch.setattr(gl.genai, "Client", lambda api_key, **kw: client)
    gl.generate_with_fallback("k", CHAIN, "p", gl.PeerList)
    gl.generate_with_fallback("k", CHAIN, "p", gl.PeerList)
    assert client.tried == [CHAIN[0], CHAIN[1], CHAIN[1]]          # primary not retried within the cooldown
