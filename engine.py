"""Valuation engine: pure Python financial maths. No network, no LLM.

Inputs : a CompanyData record (market_data.py) and an Assumptions record.
Outputs: one results dict consumed by the Excel writer, the UI and the Gemini prompt.

Money values are in the company's reporting currency (absolute units).
Per-share values are in the quote currency exactly as shown on Yahoo (e.g. pence for LSE).
"""
from __future__ import annotations

import math
from dataclasses import dataclass

import numpy as np
import pandas as pd

from market_data import CompanyData

PROJECTION_YEARS = 5


@dataclass
class Assumptions:
    risk_free: float = 0.0425
    risk_free_source: str = "Default"
    erp: float = 0.05                         # equity risk premium
    beta: float | None = None                 # None -> Yahoo beta
    cost_of_debt: float | None = None         # None -> implied from interest expense
    tax_rate: float | None = None             # None -> effective tax rate
    terminal_growth: float = 0.025
    growth_y1: float | None = None            # None -> consensus, else historical CAGR
    ebitda_margin_target: float | None = None  # None -> hold LTM margin flat
    mid_year: bool = True
    # LBO
    lbo_leverage: float = 4.5                 # senior debt / LTM EBITDA
    lbo_rate: float = 0.07
    lbo_target_irr: float = 0.20
    lbo_exit_multiple: float | None = None    # None -> base (current) EV/EBITDA
    lbo_max_ltv: float = 0.75                 # sanity cap on debt / entry EV


# --------------------------------------------------------------------------- utils
def _clip(x: float, lo: float, hi: float) -> float:
    return max(lo, min(hi, x))


def _pos(x) -> bool:
    return x is not None and isinstance(x, (int, float)) and math.isfinite(x) and x > 0


def per_share(d: CompanyData, equity_value: float) -> float:
    """Equity value in reporting currency -> price per share in quote units."""
    return float(equity_value) / d.fx_trade_to_report / d.shares * d.quote_factor


def _hist_ratio(d: CompanyData, item: str, years: int = 3) -> float | None:
    h = d.hist
    if h is None or h.empty or item not in h.index or "revenue" not in h.index:
        return None
    ratio = (h.loc[item] / h.loc["revenue"]).replace([np.inf, -np.inf], np.nan).dropna().tail(years)
    return float(ratio.mean()) if len(ratio) else None


def _hist_cagr(d: CompanyData) -> float | None:
    h = d.hist
    if h is None or h.empty or "revenue" not in h.index:
        return None
    rev = h.loc["revenue"].dropna()
    if len(rev) < 2 or rev.iloc[0] <= 0 or rev.iloc[-1] <= 0:
        return None
    return float((rev.iloc[-1] / rev.iloc[0]) ** (1 / (len(rev) - 1)) - 1)


# --------------------------------------------------------------------------- market snapshot
def compute_market(d: CompanyData) -> dict:
    fx = d.fx_trade_to_report
    mcap = d.market_cap * fx
    net_debt = d.total_debt - d.cash
    ev = mcap + net_debt + d.minority_interest + d.preferred
    ni = d.net_income_ltm
    eps = per_share(d, ni) if ni is not None else None

    if _pos(ni):
        pe, pe_src = mcap / ni, "Market cap / LTM net income to common"
    elif ni is None and _pos(d.pe_ttm_yahoo):
        pe, pe_src = d.pe_ttm_yahoo, "Yahoo trailing P/E"
    else:
        pe, pe_src = None, "Not meaningful (LTM earnings negative or unavailable)"
    fwd_pe = d.pe_fwd_yahoo if _pos(d.pe_fwd_yahoo) and d.pe_fwd_yahoo < 500 else None

    ebitda, rev = d.ebitda_ltm, d.revenue_ltm
    return {
        "price": d.price, "market_cap": mcap, "net_debt": net_debt, "ev": ev,
        "total_debt": d.total_debt, "cash": d.cash,
        "minority_interest": d.minority_interest, "preferred": d.preferred,
        "pe": pe, "pe_source": pe_src, "fwd_pe": fwd_pe, "eps_ltm": eps,
        "ev_ebitda": ev / ebitda if _pos(ebitda) and ev > 0 else None,
        "ev_revenue": ev / rev if _pos(rev) and ev > 0 else None,
        "fcf_yield": d.fcf_ltm / mcap if d.fcf_ltm is not None and mcap > 0 else None,
        "ebitda_margin": ebitda / rev if ebitda is not None and _pos(rev) else None,
        "net_margin": ni / rev if ni is not None and _pos(rev) else None,
        "net_debt_ebitda": net_debt / ebitda if _pos(ebitda) else None,
        "revenue_ltm": rev, "ebitda_ltm": ebitda, "net_income_ltm": ni, "fcf_ltm": d.fcf_ltm,
    }


# --------------------------------------------------------------------------- WACC
def compute_wacc(d: CompanyData, a: Assumptions, mkt: dict) -> dict:
    notes = []
    if a.beta is not None:
        beta, beta_raw, beta_src = a.beta, a.beta, "User override (used as entered)"
    else:
        beta_raw = _clip(d.beta, 0.2, 3.0)
        # Blume adjustment (as Bloomberg's adjusted beta): raw betas drift toward 1 over time.
        beta = 0.67 * beta_raw + 0.33
        beta_src = f"{d.beta_source}; raw {beta_raw:.2f}, Blume-adjusted (0.67 x raw + 0.33)"
    ke = a.risk_free + beta * a.erp

    if a.cost_of_debt is not None:
        kd, kd_src = a.cost_of_debt, "User override"
    else:
        implied = None
        if _pos(d.interest_expense_ltm) and _pos(d.total_debt):
            implied = d.interest_expense_ltm / d.total_debt
        lo, hi = a.risk_free + 0.01, a.risk_free + 0.08
        if implied is None:
            kd, kd_src = a.risk_free + 0.02, "Risk-free + 2.0% spread (no interest expense data)"
        else:
            kd = _clip(implied, lo, hi)
            kd_src = "LTM interest expense / total debt"
            if kd != implied:
                kd_src += f" ({implied:.1%}), bounded to risk-free +1% to +8%"

    tax = a.tax_rate if a.tax_rate is not None else d.tax_rate
    tax_src = "User override" if a.tax_rate is not None else d.tax_rate_source
    equity, debt = mkt["market_cap"], max(d.total_debt, 0.0)
    we = equity / (equity + debt)
    wd = 1.0 - we
    wacc = we * ke + wd * kd * (1 - tax)
    return {
        "risk_free": a.risk_free, "risk_free_source": a.risk_free_source, "erp": a.erp,
        "beta": beta, "beta_raw": beta_raw, "beta_source": beta_src, "cost_of_equity": ke,
        "pre_tax_cost_of_debt": kd, "cost_of_debt_source": kd_src,
        "tax_rate": tax, "tax_source": tax_src, "after_tax_cost_of_debt": kd * (1 - tax),
        "equity_value": equity, "debt_value": debt, "weight_equity": we, "weight_debt": wd,
        "wacc": wacc, "notes": notes,
    }


# --------------------------------------------------------------------------- projections
def _pick_growth(d: CompanyData, a: Assumptions) -> tuple[float, str, list]:
    notes = []
    if a.growth_y1 is not None:
        return a.growth_y1, "User override", notes
    cons = [g for g in (d.consensus_growth.get("0y"), d.consensus_growth.get("+1y")) if g is not None]
    cagr = _hist_cagr(d)
    if cons:
        g, src = float(np.mean(cons)), "Analyst consensus revenue growth (Yahoo), avg of current and next fiscal year"
    elif cagr is not None:
        g, src = cagr, f"Historical revenue CAGR over {d.hist.shape[1] - 1} fiscal years"
    else:
        g, src = a.terminal_growth, "No growth data available; terminal growth used"
    bounded = _clip(g, -0.10, 0.30)
    if bounded != g:
        notes.append(f"Year-1 revenue growth of {g:.1%} bounded to {bounded:.1%}.")
    return bounded, src, notes


def build_projections(d: CompanyData, a: Assumptions, tax: float) -> dict:
    n = PROJECTION_YEARS
    g1, g_src, notes = _pick_growth(d, a)
    gt = a.terminal_growth
    m0 = d.ebitda_ltm / d.revenue_ltm
    mt = a.ebitda_margin_target if a.ebitda_margin_target is not None else m0

    da_pct = _hist_ratio(d, "da")
    da_src = "D&A / revenue, avg of last 3 fiscal years"
    if da_pct is None:
        if d.da_ltm is not None:
            da_pct, da_src = d.da_ltm / d.revenue_ltm, "LTM D&A / revenue"
        else:
            da_pct, da_src = 0.03, "Default 3.0% of revenue (no D&A data)"
            notes.append("No D&A data; 3% of revenue assumed.")
    capex_pct = _hist_ratio(d, "capex")
    capex_src = "Capex / revenue, avg of last 3 fiscal years"
    if capex_pct is None:
        if d.capex_ltm is not None:
            capex_pct, capex_src = d.capex_ltm / d.revenue_ltm, "LTM capex / revenue"
        else:
            capex_pct, capex_src = da_pct, "Set equal to D&A (no capex data)"
            notes.append("No capex data; capex set equal to D&A.")
    if d.nwc is not None:
        raw = d.nwc / d.revenue_ltm
        nwc_pct = _clip(raw, -0.15, 0.30)
        nwc_src = "Non-cash net working capital / LTM revenue (latest balance sheet)"
        if nwc_pct != raw:
            nwc_src += f", {raw:.1%} bounded to -15%..30%"
    else:
        nwc_pct, nwc_src = 0.0, "No working-capital data; no change in NWC assumed"

    interest = d.interest_expense_ltm or 0.0
    rows, rev_prev = [], d.revenue_ltm
    for t in range(1, n + 1):
        g = g1 + (gt - g1) * (t - 1) / n
        margin = m0 + (mt - m0) * t / n
        rev = rev_prev * (1 + g)
        ebitda = rev * margin
        da = rev * da_pct
        ebit = ebitda - da
        taxes = max(ebit, 0.0) * tax
        capex = rev * capex_pct
        d_nwc = (rev - rev_prev) * nwc_pct
        pretax = ebit - interest
        net_income = pretax - max(pretax, 0.0) * tax
        rows.append({
            "year": t, "growth": g, "revenue": rev, "ebitda": ebitda, "ebitda_margin": margin,
            "da": da, "ebit": ebit, "taxes": taxes, "nopat": ebit - taxes, "capex": capex,
            "d_nwc": d_nwc, "ufcf": ebit - taxes + da - capex - d_nwc,
            "net_income": net_income, "eps": per_share(d, net_income),
        })
        rev_prev = rev
    proj = pd.DataFrame(rows).set_index("year")
    return {
        "table": proj,
        "drivers": {
            "growth_y1": g1, "growth_source": g_src, "terminal_growth": gt,
            "ebitda_margin_ltm": m0, "ebitda_margin_target": mt,
            "margin_source": "User override" if a.ebitda_margin_target is not None else "LTM EBITDA margin held flat",
            "da_pct": da_pct, "da_source": da_src, "capex_pct": capex_pct, "capex_source": capex_src,
            "nwc_pct": nwc_pct, "nwc_source": nwc_src, "tax_rate": tax,
            "interest_expense": interest,
        },
        "notes": notes,
    }


# --------------------------------------------------------------------------- DCF
def _claims(mkt: dict) -> float:
    """Non-equity claims deducted from enterprise value."""
    return mkt["net_debt"] + mkt["minority_interest"] + mkt["preferred"]


def dcf_value(proj: pd.DataFrame, wacc: float, g: float, mid_year: bool) -> dict | None:
    """5-year unlevered DCF with a Gordon-growth terminal value."""
    if wacc - g < 0.005:
        return None
    n = len(proj)
    periods = np.array([t - 0.5 if mid_year else t for t in proj.index], dtype=float)
    factors = 1.0 / (1.0 + wacc) ** periods
    pv = proj["ufcf"].to_numpy() * factors
    tv = proj["ufcf"].iloc[-1] * (1 + g) / (wacc - g)
    # Perpetuity cash flows arrive mid-year too, hence N-0.5 under the mid-year convention.
    tv_factor = 1.0 / (1.0 + wacc) ** (n - 0.5 if mid_year else n)
    return {"periods": periods, "factors": factors, "pv_ufcf": pv, "sum_pv": float(pv.sum()),
            "tv": float(tv), "pv_tv": float(tv * tv_factor), "ev": float(pv.sum() + tv * tv_factor)}


def run_dcf(d: CompanyData, a: Assumptions, mkt: dict, wacc: float, proj: pd.DataFrame) -> dict:
    g, claims = a.terminal_growth, _claims(mkt)
    notes = []
    base = dcf_value(proj, wacc, g, a.mid_year)
    if base is None:
        return {"available": False,
                "reason": f"WACC ({wacc:.1%}) is not sufficiently above terminal growth ({g:.1%})."}
    if base["tv"] <= 0:
        return {"available": False,
                "reason": "Projected terminal-year free cash flow is negative, so a perpetuity DCF is not meaningful."}

    equity = base["ev"] - claims
    price = per_share(d, equity)
    tv_share = base["pv_tv"] / base["ev"] if base["ev"] else None
    if tv_share is not None and tv_share > 0.85:
        notes.append(f"Terminal value is {tv_share:.0%} of DCF enterprise value; treat the DCF as highly assumption-sensitive.")
    if equity <= 0:
        notes.append("DCF equity value is negative: net debt exceeds DCF enterprise value.")

    # Cross-check: terminal value from the current EV/EBITDA multiple.
    exit_check = None
    if mkt["ev_ebitda"]:
        n = len(proj)
        tv_x = proj["ebitda"].iloc[-1] * mkt["ev_ebitda"]
        ev_x = base["sum_pv"] + tv_x / (1 + wacc) ** n
        f_n = proj["ufcf"].iloc[-1]
        exit_check = {"exit_multiple": mkt["ev_ebitda"], "tv": float(tv_x), "ev": float(ev_x),
                      "equity": float(ev_x - claims), "price": per_share(d, ev_x - claims),
                      "implied_growth": float((tv_x * wacc - f_n) / (tv_x + f_n)) if tv_x + f_n else None}

    steps = [-0.01, -0.005, 0.0, 0.005, 0.01]
    waccs, growths = [wacc + s for s in steps], [g + s for s in steps]
    grid = pd.DataFrame(index=waccs, columns=growths, dtype=float)
    for w in waccs:
        for gg in growths:
            v = dcf_value(proj, w, gg, a.mid_year)
            grid.loc[w, gg] = per_share(d, v["ev"] - claims) if v else np.nan
    inner = grid.iloc[1:4, 1:4].to_numpy(dtype=float)

    return {
        "available": True, "wacc": wacc, "terminal_growth": g, "mid_year": a.mid_year,
        "periods": base["periods"], "factors": base["factors"], "pv_ufcf": base["pv_ufcf"],
        "sum_pv": base["sum_pv"], "tv": base["tv"], "pv_tv": base["pv_tv"], "tv_share": tv_share,
        "implied_exit_multiple": float(base["tv"] / proj["ebitda"].iloc[-1]) if proj["ebitda"].iloc[-1] > 0 else None,
        "ev": base["ev"], "claims": claims, "equity": equity, "price": price,
        "upside": price / d.price - 1 if d.price else None,
        "range_low": float(np.nanmin(inner)), "range_high": float(np.nanmax(inner)),
        "exit_check": exit_check, "sensitivity": grid, "notes": notes,
    }


# --------------------------------------------------------------------------- LBO
def lbo_case(proj: pd.DataFrame, ebitda0: float, entry_ev: float, leverage: float, rate: float,
             exit_multiple: float, tax: float, max_ltv: float = 0.75) -> dict:
    """Sponsor returns for one scenario: senior debt only, 100% cash sweep, exit in the final projection year."""
    debt0 = min(leverage * ebitda0, max_ltv * entry_ev)
    equity0 = entry_ev - debt0
    debt, cash, rows = debt0, 0.0, []
    for t, r in proj.iterrows():
        interest = debt * rate                      # on opening balance (no circularity)
        taxes = max(r["ebit"] - interest, 0.0) * tax
        fcf = r["ebitda"] - r["capex"] - r["d_nwc"] - taxes - interest
        if fcf >= 0:
            paydown = min(fcf, debt)
            cash += fcf - paydown
        else:
            paydown = fcf                           # shortfall is funded with additional debt
        debt -= paydown
        rows.append({"year": t, "ebitda": r["ebitda"], "interest": interest, "taxes": taxes,
                     "capex": r["capex"], "d_nwc": r["d_nwc"], "fcf": fcf, "paydown": paydown,
                     "debt_end": debt, "cash_end": cash,
                     "leverage": debt / r["ebitda"] if r["ebitda"] > 0 else np.nan})
    years = len(proj)
    exit_ev = float(proj["ebitda"].iloc[-1]) * exit_multiple
    exit_equity = exit_ev - debt + cash
    moic = exit_equity / equity0 if equity0 > 0 else None
    irr = moic ** (1 / years) - 1 if moic is not None and moic > 0 else None
    return {"entry_ev": float(entry_ev), "debt0": float(debt0), "equity0": float(equity0),
            "ltv": float(debt0 / entry_ev), "leverage_used": float(debt0 / ebitda0), "exit_ev": float(exit_ev),
            "exit_debt": float(debt), "exit_cash": float(cash), "exit_equity": float(exit_equity),
            "moic": float(moic) if moic is not None else None, "irr": float(irr) if irr is not None else None,
            "years": years, "schedule": pd.DataFrame(rows).set_index("year")}


def _max_entry_ev(proj, ebitda0, a: Assumptions, exit_multiple, tax, hi_seed: float) -> float | None:
    """Highest purchase EV at which the sponsor still earns the target IRR (bisection)."""
    def gap(ev):
        irr = lbo_case(proj, ebitda0, ev, a.lbo_leverage, a.lbo_rate, exit_multiple, tax, a.lbo_max_ltv)["irr"]
        return (irr if irr is not None else -1.0) - a.lbo_target_irr
    lo, hi = hi_seed * 1e-4, hi_seed * 50
    if gap(lo) < 0 or gap(hi) > 0:
        return None
    for _ in range(100):
        mid = (lo + hi) / 2
        if gap(mid) >= 0:
            lo = mid
        else:
            hi = mid
    return lo


def run_lbo(d: CompanyData, a: Assumptions, mkt: dict, proj: pd.DataFrame, tax: float) -> dict:
    ebitda0, base_mult = d.ebitda_ltm, mkt["ev_ebitda"]
    if not _pos(ebitda0) or not base_mult:
        return {"available": False, "reason": "LTM EBITDA is negative or unavailable, so EBITDA-based leverage cannot be sized."}
    exit_mult = a.lbo_exit_multiple if a.lbo_exit_multiple else base_mult
    claims, notes = _claims(mkt), []
    kw = dict(leverage=a.lbo_leverage, rate=a.lbo_rate, tax=tax, max_ltv=a.lbo_max_ltv)

    base = lbo_case(proj, ebitda0, mkt["ev"], exit_multiple=exit_mult, **kw)
    if base["leverage_used"] < a.lbo_leverage - 1e-9:
        notes.append(f"{a.lbo_leverage:.1f}x EBITDA of debt exceeds {a.lbo_max_ltv:.0%} of entry EV; "
                     f"debt capped at {base['leverage_used']:.1f}x.")
    if base_mult > 20:
        notes.append(f"At {base_mult:.1f}x EV/EBITDA a leveraged buyout is illustrative only; "
                     f"{a.lbo_leverage:.1f}x of debt funds just {base['ltv']:.0%} of the price.")

    def max_price(mult):
        ev = _max_entry_ev(proj, ebitda0, a, mult, tax, mkt["ev"])
        return (ev, per_share(d, ev - claims)) if ev is not None else (None, None)

    max_ev, max_px = max_price(exit_mult)
    _, px_low = max_price(exit_mult * 0.9)
    _, px_high = max_price(exit_mult * 1.1)

    # Sensitivity 1: entry multiple x exit multiple
    mults = [0.8, 0.9, 1.0, 1.1, 1.2]
    entry_axis, exit_axis = [base_mult * m for m in mults], [exit_mult * m for m in mults]
    grid1 = pd.DataFrame(index=entry_axis, columns=exit_axis, dtype=float)
    for em in entry_axis:
        for xm in exit_axis:
            irr = lbo_case(proj, ebitda0, em * ebitda0, exit_multiple=xm, **kw)["irr"]
            grid1.loc[em, xm] = irr if irr is not None else np.nan
    # Sensitivity 2: leverage x interest rate
    lev_step = 0.5 if a.lbo_leverage >= 1.0 else a.lbo_leverage / 4     # keep the axis positive for tiny inputs
    rate_step = 0.005 if a.lbo_rate >= 0.01 else a.lbo_rate / 4
    levs = [a.lbo_leverage + k * lev_step for k in (-2, -1, 0, 1, 2)]
    rates = [a.lbo_rate + k * rate_step for k in (-2, -1, 0, 1, 2)]
    grid2 = pd.DataFrame(index=levs, columns=rates, dtype=float)
    for lv in levs:
        for rt in rates:
            irr = lbo_case(proj, ebitda0, mkt["ev"], lv, rt, exit_mult, tax, a.lbo_max_ltv)["irr"]
            grid2.loc[lv, rt] = irr if irr is not None else np.nan

    return {
        "available": True, "entry_multiple": base_mult, "exit_multiple": exit_mult,
        "leverage": a.lbo_leverage, "rate": a.lbo_rate, "target_irr": a.lbo_target_irr,
        "ebitda_ltm": ebitda0, **{k: base[k] for k in (
            "entry_ev", "debt0", "equity0", "ltv", "leverage_used", "exit_ev", "exit_debt", "exit_cash",
            "exit_equity", "moic", "irr", "years", "schedule")},
        "meets_target": bool(base["irr"] is not None and base["irr"] >= a.lbo_target_irr),
        "max_entry_ev": max_ev, "max_entry_multiple": max_ev / ebitda0 if max_ev else None,
        "max_price": max_px, "max_price_low": px_low, "max_price_high": px_high,
        "max_price_premium": max_px / d.price - 1 if max_px is not None and d.price else None,
        "sens_entry_exit": grid1, "sens_leverage_rate": grid2, "notes": notes,
    }


# --------------------------------------------------------------------------- peers
PEER_METRICS = ("pe", "fwd_pe", "ev_ebitda", "ev_revenue", "ebitda_margin", "revenue_growth")


def _peer_row(p: dict) -> dict:
    ev = p["market_cap"] + p["total_debt"] - p["cash"]
    pe = p["market_cap"] / p["net_income"] if _pos(p.get("net_income")) else None
    if pe is None and p.get("same_currency") and _pos(p.get("pe_yahoo")):
        pe = p["pe_yahoo"]
    fwd = p.get("fwd_pe") if p.get("same_currency") else None   # cross-currency forward P/E is unreliable
    ev_ebitda = ev / p["ebitda"] if _pos(p.get("ebitda")) and ev > 0 else None
    return {
        "ticker": p["ticker"], "name": p["name"], "currency": p["currency"],
        "market_cap": p["market_cap"], "ev": ev,
        "pe": pe if _pos(pe) and pe < 200 else None,
        "fwd_pe": fwd if _pos(fwd) and fwd < 200 else None,
        "ev_ebitda": ev_ebitda if _pos(ev_ebitda) and ev_ebitda < 100 else None,
        "ev_revenue": ev / p["revenue"] if _pos(p.get("revenue")) and ev > 0 else None,
        "ebitda_margin": p["ebitda"] / p["revenue"] if p.get("ebitda") is not None and _pos(p.get("revenue")) else None,
        "revenue_growth": p.get("revenue_growth"),
    }


def run_peers(d: CompanyData, mkt: dict, peers: list[dict]) -> dict:
    subject = {
        "ticker": d.ticker, "name": d.name, "currency": d.report_currency,
        "market_cap": mkt["market_cap"], "ev": mkt["ev"], "pe": mkt["pe"], "fwd_pe": mkt["fwd_pe"],
        "ev_ebitda": mkt["ev_ebitda"], "ev_revenue": mkt["ev_revenue"],
        "ebitda_margin": mkt["ebitda_margin"], "revenue_growth": d.revenue_growth_yoy,
    }
    rows = [_peer_row(p) for p in peers if p and p["ticker"].upper() != d.ticker.upper()]
    table = pd.DataFrame(rows, columns=list(subject.keys()))
    stats = {}
    for m in PEER_METRICS:
        vals = pd.to_numeric(table[m], errors="coerce").dropna() if len(table) else pd.Series(dtype=float)
        if len(vals):
            stats[m] = {"median": float(vals.median()), "mean": float(vals.mean()), "low": float(vals.min()),
                        "high": float(vals.max()), "q1": float(vals.quantile(0.25)),
                        "q3": float(vals.quantile(0.75)), "n": int(len(vals))}
    claims, implied = _claims(mkt), {}
    if "ev_ebitda" in stats and _pos(d.ebitda_ltm):
        s = stats["ev_ebitda"]
        implied["ev_ebitda"] = {k: per_share(d, s[q] * d.ebitda_ltm - claims)
                                for k, q in (("low", "q1"), ("base", "median"), ("high", "q3"))}
    if "pe" in stats and _pos(mkt["eps_ltm"]):
        s = stats["pe"]
        implied["pe"] = {k: s[q] * mkt["eps_ltm"] for k, q in (("low", "q1"), ("base", "median"), ("high", "q3"))}
    premium = {}
    for m in ("pe", "fwd_pe", "ev_ebitda", "ev_revenue"):
        if m in stats and subject[m] and stats[m]["median"]:
            premium[m] = subject[m] / stats[m]["median"] - 1
    return {"available": len(table) > 0, "subject": subject, "table": table, "stats": stats,
            "implied": implied, "premium": premium}


# --------------------------------------------------------------------------- orchestration
def _summary_rows(d: CompanyData, dcf: dict, lbo: dict, peers: dict) -> list[dict]:
    rows = []

    def add(method, low, base, high, basis):
        if base is None or not math.isfinite(base):
            return
        rows.append({"method": method, "low": low, "base": base, "high": high,
                     "vs_price": base / d.price - 1 if d.price else None, "basis": basis})

    if dcf.get("available"):
        add("DCF: perpetuity growth", dcf["range_low"], dcf["price"], dcf["range_high"],
            "Range: WACC and terminal growth each +/-0.5%")
        if dcf.get("exit_check"):
            add("DCF: exit multiple cross-check", None, dcf["exit_check"]["price"], None,
                f"Terminal value at current {dcf['exit_check']['exit_multiple']:.1f}x EV/EBITDA")
    for key, label in (("ev_ebitda", "Peers: EV/EBITDA"), ("pe", "Peers: P/E")):
        imp = peers.get("implied", {}).get(key)
        if imp:
            add(label, imp["low"], imp["base"], imp["high"], "Peer median; range = 25th to 75th percentile")
    if lbo.get("available") and lbo.get("max_price") is not None:
        add(f"LBO: max price at {lbo['target_irr']:.0%} IRR", lbo["max_price_low"], lbo["max_price"],
            lbo["max_price_high"], "Range: exit multiple +/-10%")
    if d.target_mean:
        n = f" ({d.analyst_count} analysts)" if d.analyst_count else ""
        add("Analyst consensus target", d.target_low, d.target_mean, d.target_high, f"Yahoo Finance{n}")
    if d.week52_low and d.week52_high:
        add("52-week trading range", d.week52_low, d.price, d.week52_high, "Low / current / high")
    return rows


def run_all(d: CompanyData, a: Assumptions, peers: list[dict] | None = None) -> dict:
    """Full quantitative run. Never raises for a thin data set: unavailable sections carry a reason."""
    notes = list(d.notes)
    mkt = compute_market(d)
    wacc = compute_wacc(d, a, mkt)
    notes += wacc["notes"]

    if d.ebitda_ltm is None:
        reason = ("EBITDA is not reported for this company (typical for banks and insurers), so an "
                  "unlevered DCF and an LBO are not meaningful. Use P/E and book-value measures instead.")
        proj = {"table": pd.DataFrame(), "drivers": {}, "notes": []}
        dcf = {"available": False, "reason": reason}
        lbo = {"available": False, "reason": reason}
        notes.append(reason)
    else:
        proj = build_projections(d, a, wacc["tax_rate"])
        notes += proj["notes"]
        dcf = run_dcf(d, a, mkt, wacc["wacc"], proj["table"])
        lbo = run_lbo(d, a, mkt, proj["table"], wacc["tax_rate"])
        for section, label in ((dcf, "DCF"), (lbo, "LBO")):
            if section["available"]:
                notes += section["notes"]
            else:
                notes.append(f"{label} not produced: {section['reason']}")
    if d.sector == "Financial Services" and d.ebitda_ltm is not None:
        notes.append("Financial-sector company: EV-based measures (EV/EBITDA, DCF, LBO) can be distorted by "
                     "operating debt and customer balances.")

    peer_out = run_peers(d, mkt, peers or [])
    if not peer_out["available"]:
        notes.append("No usable peer data; peer-implied valuation omitted.")
    return {
        "market": mkt, "wacc": wacc, "projections": proj, "dcf": dcf, "lbo": lbo, "peers": peer_out,
        "summary": _summary_rows(d, dcf, lbo, peer_out), "notes": notes,
    }
