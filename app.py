"""Equity Memo Agent: Streamlit front end.

Type a company name or ticker -> Python pulls market data (Yahoo Finance), runs
WACC / DCF / LBO / peer multiples, Gemini optionally writes commentary, and
the result is a formatted Excel memo you can download.

Run locally:  streamlit run app.py
"""
from __future__ import annotations

import math

import altair as alt
import pandas as pd
import streamlit as st

import demo_data
import excel_memo
import formatting as fmt
import gemini_layer as gl
import market_data as md
from engine import PROJECTION_YEARS, Assumptions, run_all

st.set_page_config(page_title="Equity Memo Agent", layout="wide", initial_sidebar_state="expanded")

NAVY, BLUE, GOLD, INK, MUTED = "#1F3864", "#2F5597", "#C9A227", "#1A1A1A", "#5F6B7A"
DATA_TTL = 3600   # one hour: keeps Yahoo calls low and lets identical Gemini prompts hit the cache
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

st.markdown(
    f"""<style>
    .block-container {{padding-top: 2.2rem;}}
    h1, h2, h3 {{color: {NAVY};}}
    [data-testid="stMetricValue"] {{font-size: 1.45rem; color: {NAVY};}}
    [data-testid="stMetricLabel"] {{color: {MUTED};}}
    .stance {{display:inline-block; padding:2px 12px; border-radius:4px; color:#fff; font-weight:600;}}
    </style>""",
    unsafe_allow_html=True,
)


# ----------------------------------------------------------------------------- cached data access
@st.cache_data(ttl=DATA_TTL, show_spinner=False)
def find_ticker(query: str):
    return md.resolve_ticker(query)


@st.cache_data(ttl=DATA_TTL, show_spinner=False)
def load_company(symbol: str):
    return md.fetch_company(symbol)


@st.cache_data(ttl=DATA_TTL, show_spinner=False)
def load_peer(symbol: str):
    return md.fetch_peer(symbol)


@st.cache_data(ttl=DATA_TTL, show_spinner=False)
def load_risk_free():
    return md.fetch_risk_free()


@st.cache_data(ttl=3600, show_spinner=False)
def load_industry_peers(industry_key: str, exclude: str):
    return md.fetch_industry_peers(industry_key, exclude, limit=10)


# ----------------------------------------------------------------------------- callbacks
def _reset():
    """Clear everything except the user's own Gemini key, ready for the next company."""
    keep = {k: st.session_state.get(k) for k in ("gemini_key", "use_ai")}
    st.session_state.clear()
    st.session_state.update({k: v for k, v in keep.items() if v is not None})


def _rerun_with(symbol: str | None = None):
    if symbol and "job" in st.session_state:
        st.session_state["job"] = {"query": symbol, "demo": False}
    st.session_state["pending"] = True


def _use_peers(tickers: list[str]):
    st.session_state["peers_text"] = ", ".join(tickers)


def _pct_input(label, key, default=None, placeholder=None, lo=-50.0, hi=100.0, step=0.5, help=None):
    v = st.number_input(label, value=default, min_value=lo, max_value=hi, step=step, key=key,
                        placeholder=placeholder, format="%.2f", help=help)
    return None if v is None else v / 100.0


# ----------------------------------------------------------------------------- sidebar
with st.sidebar:
    st.markdown("## Equity Memo Agent")
    st.caption("Structured valuation memo for any listed company, exported to Excel.")
    query = st.text_input("Company name or ticker", key="query", placeholder="e.g. AAPL, Shell, SAP.DE, 7203.T")
    c1, c2 = st.columns(2)
    run_clicked = c1.button("Run analysis", type="primary", width="stretch")
    c2.button("Reset", on_click=_reset, width="stretch", help="Clear the current company and all overrides")
    demo_clicked = st.button("Try the synthetic demo (no API calls)", width="stretch")

    with st.expander("Valuation assumptions"):
        rf = _pct_input("Risk-free rate (%)", "rf", placeholder="Auto: US 10-year yield", lo=0.0, hi=25.0, step=0.1,
                        help="For non-US companies, enter the local 10-year government bond yield.")
        erp = _pct_input("Equity risk premium (%)", "erp", 5.0, lo=0.0, hi=15.0, step=0.25)
        tg = _pct_input("Terminal growth (%)", "tg", 2.5, lo=0.0, hi=5.0, step=0.25)
        beta = st.number_input("Beta", value=None, min_value=0.0, max_value=4.0, step=0.05, key="beta",
                               placeholder="Auto: Yahoo beta")
        tax = _pct_input("Tax rate (%)", "tax", placeholder="Auto: effective rate", lo=0.0, hi=50.0)
        kd = _pct_input("Pre-tax cost of debt (%)", "kd", placeholder="Auto: interest / debt", lo=0.0, hi=25.0, step=0.25)
        g1 = _pct_input("Year-1 revenue growth (%)", "g1", placeholder="Auto: consensus or history", lo=-30.0, hi=60.0)
        margin = _pct_input("Year-5 EBITDA margin (%)", "margin", placeholder="Auto: hold LTM margin", lo=-20.0, hi=90.0)
        mid_year = st.checkbox("Mid-year discounting", value=True, key="mid_year")

    with st.expander("LBO assumptions"):
        lev = st.number_input("Senior debt / EBITDA (x)", value=4.5, min_value=0.5, max_value=10.0, step=0.25, key="lev")
        lbo_rate = _pct_input("Interest rate (%)", "lbo_rate", 7.0, lo=0.5, hi=25.0, step=0.25)
        target_irr = _pct_input("Minimum target IRR (%)", "target_irr", 20.0, lo=0.0, hi=60.0, step=1.0)
        exit_mult = st.number_input("Exit EV/EBITDA (x)", value=None, min_value=1.0, max_value=60.0, step=0.5,
                                    key="exit_mult", placeholder="Auto: current multiple")
        st.caption(f"Exit horizon: {PROJECTION_YEARS} years.")

    with st.expander("Peers"):
        peers_text = st.text_input("Peer tickers, comma-separated", key="peers_text",
                                   placeholder="Blank = largest companies in the same Yahoo industry")
        st.caption("Use Yahoo tickers, e.g. MSFT, GOOGL, SAP.DE")

    with st.expander("Optional: AI commentary (Gemini)"):
        st.checkbox("Add AI-written thesis, risks and news summaries", value=False, key="use_ai",
                    help="Off by default. All numbers come from the Python engine either way; Gemini only writes text.")
        st.text_input("Gemini API key", type="password", key="gemini_key",
                      help="Overrides the app's key for your session. Free key: https://aistudio.google.com/apikey")

    st.caption("Data: Yahoo Finance (yfinance). All maths: Python. Not investment advice.")

inputs = dict(rf=rf, erp=erp, tg=tg, beta=beta, tax=tax, kd=kd, g1=g1, margin=margin, mid_year=mid_year,
              lev=lev, lbo_rate=lbo_rate, target_irr=target_irr, exit_mult=exit_mult,
              peers_text=peers_text, use_ai=bool(st.session_state.get("use_ai")),
              gemini_key=st.session_state.get("gemini_key", ""))


# ----------------------------------------------------------------------------- pipeline
def _parse_tickers(text: str) -> list[str]:
    seen, out = set(), []
    for raw in text.replace(";", ",").replace("\n", ",").split(","):
        t = raw.strip().upper()
        if t and t not in seen:
            seen.add(t)
            out.append(t)
    return out[:10]


def _size_matched_peers(d) -> tuple[list[dict], str]:
    """No-AI peers: Yahoo industry list, then large caps in the same sector, keeping those of comparable size."""
    my_cap = d.market_cap * d.fx_trade_to_report
    exclude = d.ticker.upper()

    def ratio(snap):
        fx = md.fx_rate(snap["currency"], d.report_currency)
        return snap["market_cap"] * fx / my_cap if fx and snap.get("market_cap") and my_cap > 0 else None

    pool, seen, stages = [], {exclude}, []
    sector_list = md.SECTOR_PEERS.get(d.sector or "", [])
    for stage, tickers in (("industry", load_industry_peers(d.industry_key, d.ticker)), ("sector", sector_list)):
        for t in tickers:
            if t.upper() in seen:
                continue
            seen.add(t.upper())
            snap = load_peer(t)
            if snap and ratio(snap):
                pool.append((stage, ratio(snap), snap))
        stages.append(stage)
        if sum(1 for _, r_, _ in pool if 0.1 <= r_ <= 10) >= 4:
            break
    in_band = [x for x in pool if 0.1 <= x[1] <= 10]
    chosen = in_band if len(in_band) >= 3 else pool
    chosen = sorted(chosen, key=lambda x: abs(math.log(x[1])))[:6]
    kinds = {x[0] for x in chosen}
    if not chosen:
        return [], "No peer list available"
    label = ("Companies of similar size in the same Yahoo Finance industry" if kinds == {"industry"} else
             f"Companies of similar size in the same industry or sector ({d.sector})")
    return [x[2] for x in chosen], label + "; enter closer peers under 'Peers' for a tighter comparison"


def _choose_peers(d, inp, key, models, demo):
    """-> (peer snapshots, meta dict, Gemini issue or None). Manual > Gemini > Yahoo industry list."""
    meta, issue = {"peer_rationale": [], "peers_dropped": []}, None
    model = models[0]
    if demo:
        meta["peer_source"] = "Synthetic demo peers (invented companies)"
        return demo_data.demo_peers(), meta, None
    manual = _parse_tickers(inp["peers_text"])
    tickers: list[str] = []
    if manual:
        tickers, meta["peer_source"] = manual, "Entered manually in the app sidebar"
    elif key:
        try:
            picks, model = gl.suggest_peers(d.ticker, d.name, d.sector, d.industry, d.country, models, key)
            tickers = _parse_tickers(",".join(p["ticker"] for p in picks))
            meta["peer_rationale"] = picks
            meta["peer_source"] = f"Suggested by Google Gemini ({model}) and validated against Yahoo Finance"
        except gl.GeminiError as exc:
            issue = exc
    if not tickers:
        snapshots, label = _size_matched_peers(d)
        meta["peer_source"] = label + (" (Gemini unavailable)" if issue else "")
        return snapshots, meta, issue
    tickers = [t for t in tickers if t != d.ticker.upper()]
    snapshots = []
    for t in tickers:
        snap = load_peer(t)
        if snap:
            snapshots.append(snap)
        else:
            meta["peers_dropped"].append(t)
    used = {s["ticker"] for s in snapshots}
    meta["peer_rationale"] = [p for p in meta["peer_rationale"] if p["ticker"].strip().upper() in used]
    return snapshots, meta, issue


def analyse(job: dict, inp: dict) -> dict | None:
    demo = job.get("demo", False)
    status = st.status("Running analysis", expanded=True)
    try:
        with status:
            if demo:
                d, candidates = demo_data.demo_company(), []
            else:
                st.write("Finding the company on Yahoo Finance")
                symbol, candidates = find_ticker(job["query"])
                st.write(f"Loading market data and financial statements for {symbol}")
                d = load_company(symbol)
            if inp["rf"] is not None:
                risk_free, rf_source = inp["rf"], "User override"
            elif demo:
                risk_free, rf_source = md.DEFAULT_RISK_FREE, "Demo default"
            else:
                risk_free, rf_source = load_risk_free()
            a = Assumptions(
                risk_free=risk_free, risk_free_source=rf_source, erp=inp["erp"], beta=inp["beta"],
                cost_of_debt=inp["kd"], tax_rate=inp["tax"], terminal_growth=inp["tg"], growth_y1=inp["g1"],
                ebitda_margin_target=inp["margin"], mid_year=inp["mid_year"], lbo_leverage=inp["lev"],
                lbo_rate=inp["lbo_rate"], lbo_target_irr=inp["target_irr"], lbo_exit_multiple=inp["exit_mult"],
            )
            key, _key_source = gl.resolve_api_key(inp["gemini_key"]) if inp["use_ai"] else (None, "off")
            models = gl.model_chain()
            model = models[0]

            st.write("Selecting peers" + (" with Gemini" if key and not demo and not inp["peers_text"] else ""))
            peers, meta, peer_issue = _choose_peers(d, inp, key, models, demo)
            issue = None

            st.write("Running WACC, DCF, LBO and peer models in Python")
            r = run_all(d, a, peers)

            memo = None
            if demo:
                memo = demo_data.demo_memo()
                meta["llm_status"] = "Canned synthetic text (demo mode, Gemini not called)"
            elif not key and not inp["use_ai"]:
                meta["llm_status"] = "Not used (Python-only memo)"
                meta["llm_note"] = ("Python-only memo: every figure is computed from Yahoo Finance data. "
                                    "AI commentary (thesis, opportunities, threats) was not requested.")
            elif not key:
                meta["llm_status"] = "Not generated: no Gemini API key supplied"
                meta["llm_note"] = ("AI commentary was requested but no Gemini API key was supplied. "
                                    "Add a free key under 'Optional: AI commentary' in the sidebar and re-run.")
            else:
                st.write("Writing thesis, risks and news summaries with Gemini (can take up to a minute when Google is busy)")
                try:
                    memo, model = gl.write_memo(d.ticker, gl.to_json(gl.metrics_payload(d, r)),
                                                gl.to_json(gl.headlines_payload(d)), d.summary[:900], models, key)
                    meta["llm_status"] = (f"Written by Google Gemini ({model}) from Python-computed metrics and "
                                          "Yahoo Finance headlines; schema-validated JSON")
                except gl.GeminiError as exc:
                    issue = exc
            if issue is not None:
                meta["llm_status"] = f"Not generated: {issue}"
                meta["llm_note"] = f"Qualitative analysis was not generated. {issue}"
            meta["footer"] = ("Generated by Equity Memo Agent. Market data: Yahoo Finance. Financial maths: Python engine. "
                              + (f"Qualitative text: Google Gemini ({model})." if memo and not demo else ""))
            news = gl.merge_news(d, memo)

            st.write("Building the Excel memo")
            xlsx = excel_memo.build_workbook(d, r, a, memo, news, meta)
        status.update(label=f"Analysis complete: {d.name}", state="complete", expanded=False)
    except md.DataError as exc:
        status.update(label="Analysis stopped", state="error", expanded=False)
        st.error(str(exc))
        return None
    except Exception as exc:   # keep the app alive and say what went wrong
        status.update(label="Analysis failed", state="error", expanded=False)
        st.error(f"Unexpected error: {exc}")
        return None
    return {"d": d, "r": r, "a": a, "memo": memo, "news": news, "meta": meta, "xlsx": xlsx,
            "candidates": candidates, "issue": issue, "peer_issue": peer_issue, "demo": demo, "model": model,
            "file_name": excel_memo.filename(d)}


# ----------------------------------------------------------------------------- display helpers
def _table(rows: list[tuple], columns: list[str]):
    st.dataframe(pd.DataFrame(rows, columns=columns), hide_index=True, width="stretch")


def _grid(frame: pd.DataFrame, row_fmt, col_fmt, cell_fmt, corner: str):
    out = pd.DataFrame(
        [[cell_fmt(v) for v in frame.iloc[i]] for i in range(len(frame))],
        index=[row_fmt(i) for i in frame.index], columns=[col_fmt(c) for c in frame.columns])
    out.index.name = corner
    st.dataframe(out, width="stretch")


def _football_field(summary: list[dict], price: float, qc: str):
    rows = [s for s in summary if s["low"] is not None and s["high"] is not None]
    if len(rows) < 2:
        return
    df = pd.DataFrame(rows)[["method", "low", "base", "high"]]
    lo, hi = min(df["low"].min(), price), max(df["high"].max(), price)
    pad = (hi - lo) * 0.08 or 1
    x = alt.X("low:Q", title=f"Value per share ({qc})", axis=alt.Axis(tickCount=8, format=",.0f"),
              scale=alt.Scale(domain=[lo - pad, hi + pad], nice=False))
    y = alt.Y("method:N", sort=None, title=None, axis=alt.Axis(labelLimit=260, labelColor=INK))
    tip = [alt.Tooltip("method:N", title="Method"), alt.Tooltip("low:Q", format=",.2f", title="Low"),
           alt.Tooltip("base:Q", format=",.2f", title="Base"), alt.Tooltip("high:Q", format=",.2f", title="High")]
    base = alt.Chart(df)
    bars = base.mark_bar(height=16, cornerRadius=4, color=BLUE).encode(x=x, x2="high:Q", y=y, tooltip=tip)
    ticks = base.mark_tick(color="#FFFFFF", thickness=3, size=16).encode(x="base:Q", y=y, tooltip=tip)
    pdf = pd.DataFrame({"price": [price], "label": [f"Current price {price:,.2f}"]})
    rule = alt.Chart(pdf).mark_rule(color=INK, strokeDash=[5, 4], strokeWidth=1.5).encode(x="price:Q")
    label = alt.Chart(pdf).mark_text(align="left", dx=6, dy=-4, baseline="top", color=INK, fontSize=11).encode(
        x="price:Q", y=alt.value(0), text="label:N")
    chart = (bars + ticks + rule + label).properties(height=44 * len(df) + 40).configure_view(stroke=None) \
        .configure_axis(gridColor="#E3E7ED", domainColor="#BFC8D6", tickColor="#BFC8D6", titleColor=MUTED, labelColor=MUTED)
    st.altair_chart(chart, width="stretch")
    st.caption("Bar = low to high; white mark = base case; dashed line = current share price.")


def _revenue_chart(d, proj: pd.DataFrame, rc: str):
    rows = []
    if d.hist is not None and not d.hist.empty:
        for col in d.hist.columns:
            for item, name in (("revenue", "Revenue"), ("ebitda", "EBITDA")):
                v = d.hist.loc[item, col] if item in d.hist.index else None
                if v is not None and pd.notna(v):
                    rows.append((col, name, v / 1e6))
    rows += [("LTM", "Revenue", d.revenue_ltm / 1e6)]
    if d.ebitda_ltm is not None:
        rows += [("LTM", "EBITDA", d.ebitda_ltm / 1e6)]
    for t, row in proj.iterrows():
        rows += [(f"Year {t}E", "Revenue", row["revenue"] / 1e6), (f"Year {t}E", "EBITDA", row["ebitda"] / 1e6)]
    df = pd.DataFrame(rows, columns=["period", "metric", "value"])
    chart = alt.Chart(df).mark_bar(cornerRadiusTopLeft=3, cornerRadiusTopRight=3).encode(
        x=alt.X("period:N", sort=None, title=None, axis=alt.Axis(labelAngle=0)),
        xOffset=alt.XOffset("metric:N", sort=["Revenue", "EBITDA"]),
        y=alt.Y("value:Q", title=f"{rc} millions"),
        color=alt.Color("metric:N", sort=["Revenue", "EBITDA"], legend=alt.Legend(orient="top", title=None),
                        scale=alt.Scale(domain=["Revenue", "EBITDA"], range=[BLUE, GOLD])),
        tooltip=[alt.Tooltip("period:N", title="Period"), alt.Tooltip("metric:N", title="Metric"),
                 alt.Tooltip("value:Q", format=",.0f", title=f"{rc} m")],
    ).properties(height=300).configure_view(stroke=None).configure_axis(
        gridColor="#E3E7ED", domainColor="#BFC8D6", tickColor="#BFC8D6", titleColor=MUTED, labelColor=MUTED)
    st.altair_chart(chart, width="stretch")


# ----------------------------------------------------------------------------- results page
def render(res: dict):
    d, r, a, memo, meta = res["d"], res["r"], res["a"], res["memo"], res["meta"]
    m, w, dcf, lbo, peers = r["market"], r["wacc"], r["dcf"], r["lbo"], r["peers"]
    qc, rc = d.quote_currency, d.report_currency
    mm = lambda v: fmt.NA if v is None else f"{v / 1e6:,.0f}"   # noqa: E731

    left, right = st.columns([3, 1], vertical_alignment="bottom")
    with left:
        st.title(f"{d.name} ({d.ticker})")
        st.caption(" | ".join(x for x in (d.exchange, d.sector, d.industry, f"As of {d.as_of}") if x))
    with right:
        st.download_button("Download Excel memo", data=res["xlsx"], file_name=res["file_name"], mime=XLSX_MIME,
                           type="primary", width="stretch")

    if len(res["candidates"]) > 1:
        options = {f"{c['symbol']}  |  {c['name']}  |  {c['exchange']}": c["symbol"] for c in res["candidates"]}
        current = next((k for k, v in options.items() if v.upper() == d.ticker), None)
        with st.expander("Not the right company or listing?"):
            pick = st.selectbox("Other matches on Yahoo Finance", list(options),
                                index=list(options).index(current) if current else 0)
            st.button("Analyse this one instead", on_click=_rerun_with, args=(options[pick],),
                      disabled=options[pick].upper() == d.ticker)

    issue = res["issue"]
    if isinstance(issue, gl.GeminiRateLimit):
        st.warning(f"{issue} The Excel memo below is complete on the quantitative side; qualitative sections are blank.")
        st.button("Retry Gemini", on_click=_rerun_with)
    elif issue is not None:
        st.warning(f"{issue} Quantitative sections are unaffected.")
        st.button("Retry Gemini", on_click=_rerun_with)
    elif res.get("peer_issue") is not None:
        st.warning(f"Gemini could not suggest peers, so the largest companies in the same Yahoo industry were used. {res['peer_issue']}")
        st.button("Retry Gemini", on_click=_rerun_with)
    elif memo is None and res["meta"].get("llm_status", "").startswith("Not generated"):
        st.info("AI commentary was requested but no Gemini API key was found. Add one under "
                "'Optional: AI commentary' in the sidebar and run again.")
    peer_rows = res["r"]["peers"].get("table")
    if not res["demo"] and (peer_rows is None or len(peer_rows) < 2):
        st.warning("Fewer than two peers could be loaded from Yahoo Finance, so peer multiples are thin or missing. "
                   "Enter peer tickers under 'Peers' in the sidebar (e.g. MSFT, GOOGL, SAP.DE) and run again.")
    if res["demo"]:
        st.info("Synthetic demo: this company, its peers, its headlines and the memo text are invented.")

    cols = st.columns(5)
    cols[0].metric(f"Share price ({qc})", fmt.num(m["price"]), border=True)
    cols[1].metric(f"Market cap ({rc})", fmt.money(m["market_cap"]), border=True)
    cols[2].metric(f"Enterprise value ({rc})", fmt.money(m["ev"]), border=True)
    cols[3].metric("P/E (LTM)", fmt.mult(m["pe"]), border=True)
    cols[4].metric("EV / EBITDA", fmt.mult(m["ev_ebitda"]), border=True)
    cols = st.columns(5)
    cols[0].metric("FCF yield", fmt.pct(m["fcf_yield"]), border=True)
    cols[1].metric("WACC", fmt.pct(w["wacc"]), border=True)
    cols[2].metric(f"DCF value ({qc})", fmt.num(dcf.get("price")), border=True,
                   help="DCF value per share. The figure underneath is the upside or downside versus the share price.",
                   delta=f"{dcf['upside']:+.0%}" if dcf.get("available") else None)
    cols[3].metric("LBO IRR", fmt.pct(lbo.get("irr")), border=True,
                   help=f"Sponsor IRR if bought at today's price. The figure underneath is versus the {lbo.get('target_irr', 0.2):.0%} target.",
                   delta=(f"{(lbo['irr'] - lbo['target_irr']) * 100:+.1f} pts"
                          if lbo.get("available") and lbo.get("irr") is not None else None))
    cols[4].metric(f"Max LBO price ({qc})", fmt.num(lbo.get("max_price")), border=True,
                   help="Highest price per share a buyer could pay and still earn the target IRR, versus today's price.",
                   delta=f"{lbo['max_price_premium']:+.0%}" if lbo.get("max_price_premium") is not None else None)

    tabs = st.tabs(["Memo", "Valuation", "DCF", "LBO", "Peers", "Projections", "News and risks", "Assumptions"])

    with tabs[0]:
        if memo:
            colour = {"Constructive": "#1E7B4B", "Cautious": "#B03A2E"}.get(memo["stance"], "#9A6700")
            st.markdown(f'<span class="stance" style="background:{colour}">Stance: {memo["stance"]}</span>',
                        unsafe_allow_html=True)
            st.markdown("#### Thesis")
            st.write(memo["investment_thesis"])
            st.markdown("#### Variant perception")
            st.write(memo["variant_perception"])
            st.markdown("#### Valuation view")
            st.write(memo["valuation_view"])
            a_col, b_col = st.columns(2)
            with a_col:
                st.markdown("#### Key opportunities")
                for p in memo["opportunities"]:
                    st.markdown(f"**{p['title']}.** {p['detail']}")
            with b_col:
                st.markdown("#### Key threats")
                for p in memo["threats"]:
                    st.markdown(f"**{p['title']}.** {p['detail']}")
        else:
            st.write(meta.get("llm_note", "Qualitative analysis not generated."))
        if d.summary:
            with st.expander("Business description"):
                st.write(memo["business_summary"] if memo else d.summary)
        if r["notes"]:
            st.markdown("#### Model notes")
            for note in r["notes"]:
                st.markdown(f"- {note}")

    with tabs[1]:
        st.markdown(f"#### Valuation summary, per share ({qc})")
        _football_field(r["summary"], d.price, qc)
        _table([(s["method"], fmt.num(s["low"]) if s["low"] is not None else "", fmt.num(s["base"]),
                 fmt.num(s["high"]) if s["high"] is not None else "", fmt.pct(s["vs_price"], signed=True), s["basis"])
                for s in r["summary"]], ["Method", "Low", "Base", "High", "Base vs price", "Basis"])
        st.markdown("#### WACC")
        _table([
            ("Risk-free rate", fmt.pct(w["risk_free"], 2), w["risk_free_source"]),
            ("Equity risk premium", fmt.pct(w["erp"], 2), "Assumption"),
            ("Beta (adjusted)", fmt.num(w["beta"]), w["beta_source"]),
            ("Cost of equity", fmt.pct(w["cost_of_equity"], 2), "Risk-free + beta x equity risk premium"),
            ("Pre-tax cost of debt", fmt.pct(w["pre_tax_cost_of_debt"], 2), w["cost_of_debt_source"]),
            ("Tax rate", fmt.pct(w["tax_rate"]), w["tax_source"]),
            ("Weight of equity / debt", f"{w['weight_equity']:.1%} / {w['weight_debt']:.1%}", "Market cap and total debt"),
            ("WACC", fmt.pct(w["wacc"], 2), "Weighted average"),
        ], ["Component", "Value", "Basis"])

    proj = r["projections"]["table"]
    with tabs[2]:
        if not dcf.get("available"):
            st.warning(f"DCF not produced. {dcf.get('reason', '')}")
        else:
            k = st.columns(4)
            k[0].metric(f"Enterprise value ({rc})", fmt.money(dcf["ev"]))
            k[1].metric(f"Equity value ({rc})", fmt.money(dcf["equity"]))
            k[2].metric(f"Value per share ({qc})", fmt.num(dcf["price"]))
            k[3].metric("Terminal value share of EV", fmt.pct(dcf["tv_share"], 0))
            st.markdown(f"#### Unlevered free cash flow ({rc} millions)")
            lines = [("Revenue", "revenue", mm), ("Growth", "growth", fmt.pct), ("EBITDA", "ebitda", mm),
                     ("EBITDA margin", "ebitda_margin", fmt.pct), ("EBIT", "ebit", mm), ("Taxes on EBIT", "taxes", mm),
                     ("D&A", "da", mm), ("Capex", "capex", mm), ("Increase in NWC", "d_nwc", mm),
                     ("Unlevered FCF", "ufcf", mm)]
            table = pd.DataFrame({f"Year {t}": [f(proj.loc[t, c]) for _, c, f in lines] for t in proj.index},
                                 index=[n for n, _, _ in lines])
            table.loc["PV of UFCF"] = [mm(v) for v in dcf["pv_ufcf"]]
            st.dataframe(table, width="stretch")
            st.markdown(f"#### Sensitivity: value per share ({qc})")
            _grid(dcf["sensitivity"], lambda v: f"WACC {v:.1%}", lambda v: f"g {v:.1%}", fmt.num, "WACC / terminal growth")

    with tabs[3]:
        if not lbo.get("available"):
            st.warning(f"LBO not produced. {lbo.get('reason', '')}")
        else:
            st.caption(f"{lbo['leverage_used']:.1f}x senior debt / EBITDA at {lbo['rate']:.1%}; entry at {lbo['entry_multiple']:.1f}x, "
                       f"{lbo['years']}-year exit at {lbo['exit_multiple']:.1f}x EV/EBITDA; target IRR {lbo['target_irr']:.0%}.")
            k = st.columns(4)
            k[0].metric("Sponsor IRR", fmt.pct(lbo["irr"]))
            k[1].metric("MOIC", fmt.mult(lbo["moic"]))
            k[2].metric("Meets target IRR?", "Yes" if lbo["meets_target"] else "No")
            k[3].metric(f"Max price for target IRR ({qc})", fmt.num(lbo["max_price"]))
            st.markdown(f"#### Debt schedule ({rc} millions)")
            sch = lbo["schedule"]
            lines = [("EBITDA", "ebitda", mm), ("Cash interest", "interest", mm), ("Cash taxes", "taxes", mm),
                     ("FCF for debt repayment", "fcf", mm), ("Senior debt, end of year", "debt_end", mm),
                     ("Debt / EBITDA", "leverage", fmt.mult)]
            st.dataframe(pd.DataFrame({f"Year {t}": [f(sch.loc[t, c]) for _, c, f in lines] for t in sch.index},
                                      index=[n for n, _, _ in lines]), width="stretch")
            st.markdown("#### IRR sensitivity: entry multiple vs exit multiple")
            _grid(lbo["sens_entry_exit"], lambda v: f"Entry {v:.1f}x", lambda v: f"Exit {v:.1f}x", fmt.pct, "Entry / exit")
            st.markdown("#### IRR sensitivity: leverage vs interest rate")
            _grid(lbo["sens_leverage_rate"], lambda v: f"{v:.1f}x debt", lambda v: f"{v:.1%}", fmt.pct, "Leverage / rate")

    with tabs[4]:
        st.caption(f"Peers: {meta.get('peer_source', '')}.")
        rows = [peers["subject"]] + peers["table"].to_dict("records")
        _table([(p["ticker"], p["name"], p["currency"], fmt.money(p["market_cap"]), fmt.money(p["ev"]), fmt.mult(p["pe"]),
                 fmt.mult(p["fwd_pe"]), fmt.mult(p["ev_ebitda"]), fmt.mult(p["ev_revenue"]),
                 fmt.pct(p["ebitda_margin"]), fmt.pct(p["revenue_growth"])) for p in rows],
               ["Ticker", "Company", "Ccy", "Market cap", "EV", "P/E", "Fwd P/E", "EV/EBITDA", "EV/Revenue",
                "EBITDA margin", "Rev growth"])
        if peers["available"]:
            stats = peers["stats"]
            _table([(lab, *(fmt.mult(stats.get(k, {}).get(s)) for k in ("pe", "fwd_pe", "ev_ebitda", "ev_revenue")))
                    for lab, s in (("Peer median", "median"), ("Peer mean", "mean"), ("Peer high", "high"), ("Peer low", "low"))]
                   + [(f"{d.ticker} premium to median", *(fmt.pct(peers["premium"].get(k), signed=True)
                                                           for k in ("pe", "fwd_pe", "ev_ebitda", "ev_revenue")))],
                   ["", "P/E", "Fwd P/E", "EV/EBITDA", "EV/Revenue"])
            if meta.get("peer_rationale") and not inputs["peers_text"]:
                st.button("Copy these peers to the sidebar to edit them", on_click=_use_peers,
                          args=(list(peers["table"]["ticker"]),))
        if meta.get("peers_dropped"):
            st.caption("Excluded (no usable Yahoo data): " + ", ".join(meta["peers_dropped"]))

    with tabs[5]:
        if proj.empty:
            st.warning("Projections not produced for this company.")
        else:
            _revenue_chart(d, proj, rc)
            drv = r["projections"]["drivers"]
            _table([
                ("Year-1 revenue growth", fmt.pct(drv["growth_y1"]), drv["growth_source"]),
                ("Terminal growth", fmt.pct(drv["terminal_growth"]), "Growth fades in a straight line toward this rate"),
                ("EBITDA margin, year 5", fmt.pct(drv["ebitda_margin_target"]), drv["margin_source"]),
                ("D&A, % of revenue", fmt.pct(drv["da_pct"]), drv["da_source"]),
                ("Capex, % of revenue", fmt.pct(drv["capex_pct"]), drv["capex_source"]),
                ("NWC, % of revenue", fmt.pct(drv["nwc_pct"]), drv["nwc_source"]),
            ], ["Driver", "Value", "Basis"])
            lines = [("Revenue", "revenue", mm), ("EBITDA", "ebitda", mm), ("EBIT", "ebit", mm),
                     ("Net income (illustrative)", "net_income", mm), (f"EPS ({qc}, illustrative)", "eps", fmt.num),
                     ("Unlevered FCF", "ufcf", mm)]
            st.dataframe(pd.DataFrame({f"Year {t}E": [f(proj.loc[t, c]) for _, c, f in lines] for t in proj.index},
                                      index=[n for n, _, _ in lines]), width="stretch")

    with tabs[6]:
        st.markdown("#### Recent news")
        if not res["news"]:
            st.write("No recent headlines returned by Yahoo Finance.")
        for n in res["news"]:
            head = f"**{n['title']}**"
            if str(n.get("url", "")).startswith("http"):
                head = f"**[{n['title']}]({n['url']})**"
            tag = " | ".join(x for x in (n.get("date"), n.get("publisher"), n.get("sentiment")) if x)
            st.markdown(f"{head}  \n{tag}  \n{n.get('summary', '')}")
        if memo:
            st.markdown("#### Risk register")
            _table([(x["risk"], x["severity"], x["detail"], x["mitigant"]) for x in memo["risks"]],
                   ["Risk", "Severity", "How it hurts", "Mitigant / what to watch"])
            st.markdown("#### What to monitor")
            for item in memo["monitor"]:
                st.markdown(f"- {item}")

    with tabs[7]:
        st.caption("The same detail, with sources, is on the Assumptions tab of the Excel memo.")
        _table([
            ("Flow figures", d.ltm_basis, f"Balance sheet date {d.balance_sheet_date or 'latest'}"),
            ("Quote / reporting currency", f"{qc} / {rc}", f"FX {d.trade_currency}/{rc}: {d.fx_trade_to_report:.4f}"),
            ("Risk-free rate", fmt.pct(a.risk_free, 2), a.risk_free_source),
            ("Terminal growth", fmt.pct(a.terminal_growth, 2), "Sidebar"),
            ("Discounting", "Mid-year" if a.mid_year else "End of year", "Sidebar"),
            ("LBO structure", f"{a.lbo_leverage:.1f}x at {a.lbo_rate:.1%}", f"{PROJECTION_YEARS}-year exit, target IRR {a.lbo_target_irr:.0%}"),
            ("Peers", f"{len(peers['table'])} companies", meta.get("peer_source", "")),
            ("Qualitative text", "Gemini" if memo and not res["demo"] else "None", meta.get("llm_status", "")),
        ], ["Item", "Value", "Basis"])

    st.caption(excel_memo.DISCLAIMER)


def landing():
    st.title("Equity Memo Agent")
    st.write("A repeatable buy-side style valuation memo for any listed company, delivered as a formatted Excel workbook.")
    a_col, b_col, c_col = st.columns(3)
    a_col.markdown("**1. Enter a company**  \nName or Yahoo ticker in the sidebar. Non-US listings work with a suffix, such as `SHEL.L` or `SAP.DE`.")
    b_col.markdown("**2. Python does all the maths**  \nPrice, market cap, EV, P/E, EV/EBITDA, FCF yield, WACC, 5-year DCF, LBO returns and peer multiples.")
    c_col.markdown("**3. Download the memo**  \nA formatted Excel workbook with Yahoo headlines, assumptions you can edit and a football-field chart. AI commentary is optional and off by default.")
    st.caption("Reset in the sidebar clears the current company so you can run the next one. " + excel_memo.DISCLAIMER)


# ----------------------------------------------------------------------------- controller
trigger = st.session_state.pop("pending", False)
if run_clicked:
    if query.strip():
        st.session_state["job"] = {"query": query.strip(), "demo": False}
        trigger = True
    else:
        st.sidebar.error("Enter a company name or ticker first.")
if demo_clicked:
    st.session_state["job"] = {"query": "DEMO", "demo": True}
    trigger = True

if trigger and "job" in st.session_state:
    outcome = analyse(st.session_state["job"], inputs)
    if outcome is not None:
        st.session_state["result"] = outcome
    else:
        st.session_state.pop("result", None)

if "result" in st.session_state:
    render(st.session_state["result"])
elif not trigger:
    landing()
