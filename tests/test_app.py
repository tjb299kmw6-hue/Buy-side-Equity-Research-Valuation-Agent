"""End-to-end app tests with Streamlit's AppTest. Yahoo and Gemini are replaced by fakes, so no network is used."""
from __future__ import annotations

import io
import os
import sys

import pytest
from openpyxl import load_workbook
from streamlit.testing.v1 import AppTest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import demo_data  # noqa: E402
import gemini_layer as gl  # noqa: E402
import market_data as md  # noqa: E402
from test_agent import FakeTicker  # noqa: E402

APP = os.path.join(ROOT, "app.py")


class FakeSearch:
    def __init__(self, query, **kwargs):
        self.quotes = [
            {"symbol": "ACME.L", "longname": "Acme Tools plc", "quoteType": "EQUITY", "exchDisp": "London"},
            {"symbol": "ACME", "longname": "Acme Tools ADR", "quoteType": "EQUITY", "exchDisp": "NYSE"},
            {"symbol": "ACMX", "longname": "Acme ETF", "quoteType": "ETF", "exchDisp": "NYSE"},
        ]


class FakeGemini:
    """Stands in for genai.Client: answers by schema, or raises the configured error."""
    error = None            # raised for every model
    fail_models: set = set()  # raise a 429 only for these model ids
    models_used: list = []
    calls = 0

    def __init__(self, api_key, **kwargs):
        self.models = self

    def generate_content(self, model, contents, config):
        FakeGemini.calls += 1
        FakeGemini.models_used.append(model)
        if FakeGemini.error:
            raise FakeGemini.error
        if model in FakeGemini.fail_models:
            raise gl.genai_errors.ClientError(429, {"error": {"code": 429, "message": "quota", "status": "RESOURCE_EXHAUSTED"}})
        if config.response_schema is gl.PeerList:
            body = gl.PeerList(peers=[gl.PeerPick(ticker=t, name=t, rationale="Similar tools business.")
                                      for t in ("AAA.L", "BBB.L", "ACME.L")])
        else:
            body = gl.QualitativeMemo(**demo_data.demo_memo())
        return type("R", (), {"text": body.model_dump_json(), "parsed": None})()


@pytest.fixture(autouse=True)
def fakes(monkeypatch):
    import streamlit as st
    st.cache_data.clear()
    md._fx_cached.cache_clear()
    FakeGemini.error, FakeGemini.calls = None, 0
    gl._COOLDOWN.clear()
    FakeGemini.fail_models, FakeGemini.models_used = set(), []
    monkeypatch.delenv("GEMINI_MODEL", raising=False)
    monkeypatch.setattr(md.yf, "Ticker", FakeTicker)
    monkeypatch.setattr(md.yf, "Search", FakeSearch)
    monkeypatch.setattr(md.yf, "Industry", lambda key: (_ for _ in ()).throw(RuntimeError("offline")))
    monkeypatch.setattr(gl.genai, "Client", FakeGemini)
    monkeypatch.delenv("GEMINI_API_KEY", raising=False)


def _button(at, label):
    return next(b for b in at.button if b.label == label)


def _run_company(at, query="Acme", ai=False):
    if ai:
        at.checkbox(key="use_ai").check()
    at.text_input(key="query").input(query)
    at.run()
    _button(at, "Run analysis").click()
    at.run()
    return at


def _metrics(at):
    return {m.label: m.value for m in at.metric}


def test_landing_page_renders():
    at = AppTest.from_file(APP, default_timeout=60).run()
    assert not at.exception
    assert at.title[0].value == "Equity Memo Agent"


def test_demo_mode_needs_no_network():
    at = AppTest.from_file(APP, default_timeout=60).run()
    _button(at, "Try the synthetic demo (no API calls)").click()
    at.run()
    assert not at.exception
    m = _metrics(at)
    assert m["Share price (USD)"] == "84.50" and m["P/E (LTM)"] == "20.9x" and m["EV / EBITDA"] == "11.8x"
    assert FakeGemini.calls == 0
    wb = load_workbook(io.BytesIO(at.session_state["result"]["xlsx"]))
    assert len(wb.sheetnames) == 8


def test_full_run_with_gemini(monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY", "test-key")
    at = _run_company(AppTest.from_file(APP, default_timeout=60).run(), ai=True)
    assert not at.exception
    res = at.session_state["result"]
    assert res["d"].ticker == "ACME.L" and res["memo"]["stance"] == "Neutral" and res["issue"] is None
    assert [p["ticker"] for p in res["r"]["peers"]["table"].to_dict("records")] == ["AAA.L", "BBB.L"]   # self removed
    assert "Gemini" in res["meta"]["peer_source"]
    assert FakeGemini.calls == 2                                   # one peer call, one memo call
    assert _metrics(at)["Share price (GBp)"] == "2,000.00"
    # Same company again inside the hour: served from cache, no further Gemini calls.
    _button(at, "Run analysis").click()
    at.run()
    assert FakeGemini.calls == 2 and not at.exception


def test_name_versus_ticker_resolution():
    assert md.resolve_ticker("Acme")[0] == "ACME.L"       # a name: Yahoo's top equity hit
    assert md.resolve_ticker("ACME")[0] == "ACME"         # typed as a ticker: exact symbol
    assert md.resolve_ticker("acme.l")[0] == "ACME.L"
    assert len(md.resolve_ticker("Acme")[1]) == 2         # the ETF is filtered out


def test_manual_peers_skip_the_gemini_peer_call(monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY", "test-key")
    at = AppTest.from_file(APP, default_timeout=60).run()
    at.text_input(key="peers_text").input("ccc.l, ddd.l")
    at = _run_company(at, ai=True)
    res = at.session_state["result"]
    assert list(res["r"]["peers"]["table"]["ticker"]) == ["CCC.L", "DDD.L"]
    assert res["meta"]["peer_source"].startswith("Entered manually") and FakeGemini.calls == 1


def test_ai_is_off_by_default_even_if_a_key_exists(monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY", "test-key")
    at = _run_company(AppTest.from_file(APP, default_timeout=60).run())
    assert not at.exception
    res = at.session_state["result"]
    assert res["memo"] is None and FakeGemini.calls == 0 and len(res["xlsx"]) > 10_000
    assert res["meta"]["llm_status"].startswith("Not used") and not at.info and not at.warning
    assert "similar size" in res["meta"]["peer_source"]
    assert res["news"] and all(n["summary"] == n["snippet"] for n in res["news"])


def test_ai_on_without_key_explains_and_still_delivers():
    at = _run_company(AppTest.from_file(APP, default_timeout=60).run(), ai=True)
    assert not at.exception
    res = at.session_state["result"]
    assert res["memo"] is None and FakeGemini.calls == 0 and len(res["xlsx"]) > 10_000
    assert any("no Gemini API key" in i.value for i in at.info)


def test_rate_limit_shows_retry_and_recovers(monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY", "test-key")
    FakeGemini.error = gl.genai_errors.ClientError(429, {"error": {"code": 429, "message": "quota", "status": "RESOURCE_EXHAUSTED"}})
    at = _run_company(AppTest.from_file(APP, default_timeout=60).run(), ai=True)
    assert not at.exception
    res = at.session_state["result"]
    assert isinstance(res["issue"], gl.GeminiRateLimit) and res["memo"] is None and len(res["xlsx"]) > 10_000
    assert any("429" in w.value for w in at.warning)
    FakeGemini.error = None
    _button(at, "Retry Gemini").click()
    at.run()
    assert at.session_state["result"]["memo"] is not None and not at.warning


def test_sidebar_overrides_reach_the_engine():
    at = AppTest.from_file(APP, default_timeout=60).run()
    at.number_input(key="lev").set_value(3.0)
    at.number_input(key="target_irr").set_value(15.0)
    at.number_input(key="rf").set_value(5.0)
    at = _run_company(at)
    res = at.session_state["result"]
    assert res["a"].lbo_leverage == 3.0 and res["a"].lbo_target_irr == pytest.approx(0.15)
    assert res["r"]["wacc"]["risk_free"] == pytest.approx(0.05)
    assert res["r"]["lbo"]["debt0"] == pytest.approx(3.0 * 2.1e9)


def test_unknown_ticker_gives_clean_error(monkeypatch):
    class NoData(FakeTicker):
        def __init__(self, s):
            super().__init__(s)
            self.info, self.fast_info = {}, {}
    monkeypatch.setattr(md.yf, "Ticker", NoData)
    monkeypatch.setattr(md.yf, "Search", lambda q, **k: type("S", (), {"quotes": []})())
    at = _run_company(AppTest.from_file(APP, default_timeout=60).run(), "ZZZZ")
    assert not at.exception
    assert any("No market price" in e.value for e in at.error) and "result" not in at.session_state


def test_reset_clears_the_company_but_keeps_the_key():
    at = AppTest.from_file(APP, default_timeout=60).run()
    at.text_input(key="gemini_key").input("my-key")
    at = _run_company(at)
    assert "result" in at.session_state
    _button(at, "Reset").click()
    at.run()
    assert "result" not in at.session_state and at.text_input(key="query").value == ""
    assert at.session_state["gemini_key"] == "my-key"


def test_primary_out_of_quota_falls_back_through_the_chain(monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY", "test-key")
    FakeGemini.fail_models = {"gemini-3.5-flash-lite", "gemini-3.8-flash"}
    at = _run_company(AppTest.from_file(APP, default_timeout=60).run(), ai=True)
    assert not at.exception and not at.warning
    res = at.session_state["result"]
    assert res["memo"] is not None and res["model"] == "gemini-3.6-flash"
    assert "gemini-3.6-flash" in res["meta"]["llm_status"] and "gemini-3.6-flash" in res["meta"]["footer"]
    assert FakeGemini.models_used[:3] == ["gemini-3.5-flash-lite", "gemini-3.8-flash", "gemini-3.6-flash"]


def test_all_three_models_out_of_quota_still_delivers_workbook(monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY", "test-key")
    FakeGemini.fail_models = {"gemini-3.8-flash", "gemini-3.5-flash-lite", "gemini-3.6-flash"}
    at = _run_company(AppTest.from_file(APP, default_timeout=60).run(), ai=True)
    assert not at.exception
    res = at.session_state["result"]
    assert res["memo"] is None and len(res["xlsx"]) > 10_000
    assert any("all models tried" in w.value for w in at.warning)




def test_no_ai_peers_are_size_matched(monkeypatch):
    """Tiny names from Yahoo's industry list are dropped for companies of comparable size."""
    import pandas as pd

    def snap(sym, cap):
        return {"ticker": sym, "name": sym, "currency": "USD", "market_cap": cap, "total_debt": 1e9, "cash": 1e9,
                "revenue": 1e10, "ebitda": 2e9, "net_income": 1e9, "pe_yahoo": 15.0, "fwd_pe": 14.0,
                "revenue_growth": 0.05, "same_currency": True}
    caps = {"TINY1": 1e6, "TINY2": 5e6, "BIG1": 2.2e10, "BIG2": 2.8e10, "BIG3": 3.1e10, "BIG4": 1.9e10}
    monkeypatch.setattr(md.yf, "Industry", lambda key: type("I", (), {"top_companies": pd.DataFrame(index=list(caps))})())
    monkeypatch.setattr(md, "fetch_peer", lambda sym: snap(sym.upper(), caps[sym.upper()]) if sym.upper() in caps else None)
    at = _run_company(AppTest.from_file(APP, default_timeout=60).run())
    assert not at.exception
    got = list(at.session_state["result"]["r"]["peers"]["table"]["ticker"])
    assert got and not any(t.startswith("TINY") for t in got) and {"BIG1", "BIG2"} <= set(got)
    assert "similar size" in at.session_state["result"]["meta"]["peer_source"]
