"""Gemini layer: qualitative work only (peer ideas, news summaries, thesis, risks).

Financial maths never happens here. Gemini receives numbers that Python has already
computed and formatted, plus the top headlines, and returns JSON that is validated
against a Pydantic schema so it maps straight into Excel cells.
"""
from __future__ import annotations

import json
import os
import re
import time
from typing import Literal

import streamlit as st
from google import genai
from google.genai import errors as genai_errors
from google.genai import types
from pydantic import BaseModel, Field, ValidationError

import formatting as fmt
from market_data import CompanyData

DEFAULT_MODEL = "gemini-3.5-flash-lite"   # primary (fast, rarely overloaded); free-tier compatible; override with GEMINI_MODEL in secrets/env
FALLBACK_MODELS = ("gemini-3.8-flash", "gemini-3.6-flash")   # tried in order if the primary's quota runs out
CACHE_TTL = 3600
REQUEST_TIMEOUT_MS = 40_000   # per model; a slow model is skipped in favour of the next one
COOLDOWN_SECONDS = 300        # a model that was busy or out of quota is skipped for this long
_COOLDOWN: dict[str, float] = {}

SYSTEM_INSTRUCTION = (
    "You are a senior buy-side equity analyst writing an internal investment memo. "
    "You are rigorous, sceptical and concise.\n"
    "Rules:\n"
    "1. Never perform, re-derive, round or adjust financial calculations. Any figure you quote must be "
    "copied exactly as it appears in the METRICS block.\n"
    "2. Base news summaries only on the headline and snippet supplied. Do not add facts, figures, dates "
    "or quotes that are not in the input.\n"
    "3. If the data does not support a claim, say so rather than guessing.\n"
    "4. Plain professional English. No markdown, no bullet characters, no emojis, no disclaimers."
)


# --------------------------------------------------------------------------- schemas
class PeerPick(BaseModel):
    ticker: str = Field(description="Yahoo Finance ticker, with exchange suffix for non-US listings (SAP.DE, SHEL.L, 7203.T)")
    name: str = Field(description="Company name")
    rationale: str = Field(description="One short sentence on why it is comparable")


class PeerList(BaseModel):
    peers: list[PeerPick]


class NewsSummary(BaseModel):
    headline_id: int = Field(description="The id of the supplied headline this summary refers to")
    summary: str = Field(description="One or two sentences: what happened and why it matters for the equity case")
    sentiment: Literal["Positive", "Neutral", "Negative"]


class Point(BaseModel):
    title: str = Field(description="Three to six words")
    detail: str = Field(description="One or two sentences")


class Risk(BaseModel):
    risk: str = Field(description="Three to six words")
    severity: Literal["High", "Medium", "Low"]
    detail: str = Field(description="One sentence on how it would hurt the equity value")
    mitigant: str = Field(description="One sentence on what offsets it or what to watch")


class QualitativeMemo(BaseModel):
    business_summary: str = Field(description="Two or three sentences on what the company does and how it makes money")
    investment_thesis: str = Field(description="Three or four sentences: the core argument for or against owning the shares")
    variant_perception: str = Field(description="Two or three sentences: where this view may differ from consensus and why")
    valuation_view: str = Field(description="Two or three sentences interpreting the supplied DCF, peer and LBO outputs versus the share price")
    stance: Literal["Constructive", "Neutral", "Cautious"]
    news: list[NewsSummary]
    opportunities: list[Point]
    threats: list[Point]
    risks: list[Risk]
    monitor: list[str] = Field(description="Three short items to track over the next two quarters")


# --------------------------------------------------------------------------- errors
class GeminiError(Exception):
    """Any Gemini failure, with a message safe to show in the UI."""


class GeminiRateLimit(GeminiError):
    def __init__(self, message: str, retry_after: int | None = None):
        super().__init__(message)
        self.retry_after = retry_after


class GeminiBusy(GeminiError):
    """Temporary server-side problem (500/503/504, e.g. 'model is experiencing high demand')."""


# --------------------------------------------------------------------------- configuration
def resolve_api_key(sidebar_key: str = "") -> tuple[str | None, str]:
    """Sidebar override -> st.secrets["GEMINI_API_KEY"] -> GEMINI_API_KEY env var."""
    if sidebar_key and sidebar_key.strip():
        return sidebar_key.strip(), "sidebar"
    try:
        key = st.secrets["GEMINI_API_KEY"]
        if key:
            return str(key).strip(), "st.secrets"
    except Exception:   # no secrets file, or key not set
        pass
    key = os.environ.get("GEMINI_API_KEY")
    return (key.strip(), "environment") if key else (None, "none")


def resolve_model() -> str:
    try:
        model = st.secrets["GEMINI_MODEL"]
        if model:
            return str(model)
    except Exception:
        pass
    return os.environ.get("GEMINI_MODEL") or DEFAULT_MODEL


def model_chain() -> tuple[str, ...]:
    """Primary model followed by the two fallbacks (duplicates removed, order kept)."""
    chain = []
    for m in (resolve_model(), *FALLBACK_MODELS):
        if m not in chain:
            chain.append(m)
    return tuple(chain)


# --------------------------------------------------------------------------- API call
def _retry_seconds(exc) -> int | None:
    try:
        for item in (exc.details or {}).get("error", {}).get("details", []):
            delay = item.get("retryDelay")
            if delay:
                return int(float(re.sub(r"[^0-9.]", "", str(delay)))) + 1
    except Exception:
        pass
    return None


def _generate(api_key: str, model: str, prompt: str, schema: type[BaseModel]) -> BaseModel:
    try:
        client = genai.Client(api_key=api_key, http_options=types.HttpOptions(
            timeout=REQUEST_TIMEOUT_MS,                               # never leave the user waiting indefinitely
            retry_options=types.HttpRetryOptions(attempts=1)))         # we fall back to the next model instead
        response = client.models.generate_content(
            model=model,
            contents=prompt,
            config=types.GenerateContentConfig(
                system_instruction=SYSTEM_INSTRUCTION,
                response_mime_type="application/json",
                response_schema=schema,
                temperature=0.3,
            ),
        )
    except genai_errors.APIError as exc:
        code = getattr(exc, "code", None)
        if code == 429:
            wait = _retry_seconds(exc)
            raise GeminiRateLimit(
                "Gemini free-tier rate limit reached (429). "
                + (f"Retry in about {wait} seconds" if wait else "Wait about a minute and retry")
                + ", or paste your own free API key in the sidebar.", wait) from exc
        if code in (400, 401, 403) and "api key" in str(exc).lower():
            raise GeminiError("Gemini rejected the API key. Check GEMINI_API_KEY or the sidebar key.") from exc
        if code in (500, 502, 503, 504):
            raise GeminiBusy(f"Gemini model '{model}' is temporarily overloaded ({code}).") from exc
        if code == 404:
            raise GeminiError(f"Gemini model '{model}' was not found for this key. Set GEMINI_MODEL to an available model.") from exc
        raise GeminiError(f"Gemini API error {code}: {getattr(exc, 'message', None) or exc}") from exc
    except Exception as exc:
        if "timeout" in type(exc).__name__.lower() or "timed out" in str(exc).lower():
            raise GeminiBusy(f"Gemini model '{model}' did not respond within {REQUEST_TIMEOUT_MS // 1000} seconds.") from exc
        raise GeminiError(f"Could not reach Gemini: {exc}") from exc

    parsed = getattr(response, "parsed", None)
    if isinstance(parsed, schema):
        return parsed
    text = getattr(response, "text", None)
    if not text:
        raise GeminiError("Gemini returned an empty response.")
    try:
        return schema.model_validate_json(text)
    except ValidationError as exc:
        raise GeminiError("Gemini returned JSON that did not match the expected schema.") from exc


def generate_with_fallback(api_key: str, models: tuple[str, ...], prompt: str,
                           schema: type[BaseModel]) -> tuple[BaseModel, str]:
    """Try each model in order. Move on when a model's quota is exhausted (429), it is temporarily
    overloaded (500/503/504), or it is unavailable (404).

    Free-tier quotas are counted per model, so a later model usually still has capacity.
    Returns (result, model that answered). Any other error stops immediately.
    """
    skipped: list[str] = []
    last_limit: GeminiRateLimit | None = None
    busy = False
    now = time.time()
    fresh = [m for m in models if _COOLDOWN.get(m, 0) <= now]
    # Skip models that just failed (saves 30-60 s of waiting per call); if all are cooling down, try them all.
    for model in (fresh or list(models)):
        try:
            return _generate(api_key, model, prompt, schema), model
        except GeminiRateLimit as exc:
            last_limit = exc
            _COOLDOWN[model] = time.time() + COOLDOWN_SECONDS
            skipped.append(f"{model} (quota used up)")
        except GeminiBusy as exc:
            busy = True
            _COOLDOWN[model] = time.time() + COOLDOWN_SECONDS
            skipped.append(f"{model} ({'timed out' if 'did not respond' in str(exc) else 'overloaded'})")
        except GeminiError as exc:
            if "was not found" not in str(exc):
                raise
            skipped.append(f"{model} (not available)")
    if busy and last_limit is None:
        raise GeminiBusy("Gemini is busy on all models tried: " + ", ".join(skipped)
                         + ". This is usually temporary; click Retry Gemini in a minute.")
    waits = [e for e in (last_limit.retry_after,) if e] if last_limit else []
    raise GeminiRateLimit(
        "Gemini free-tier rate limit reached (429) on all models tried: " + ", ".join(skipped) + ". "
        + (f"Retry in about {waits[0]} seconds" if waits else "Wait a minute (or until tomorrow if the daily limit is used)")
        + ", or paste your own free API key in the sidebar.", waits[0] if waits else None)


# --------------------------------------------------------------------------- cached entry points
# Exceptions are never cached by st.cache_data, so a 429 can simply be retried.
# `_api_key` starts with an underscore so it is not part of the cache key (or stored).
@st.cache_data(ttl=CACHE_TTL, show_spinner=False)
def suggest_peers(ticker: str, name: str, sector: str, industry: str, country: str,
                  models: tuple, _api_key: str) -> tuple[list[dict], str]:
    prompt = (
        f"Suggest 6 publicly listed companies that are the closest operating and valuation comparables for "
        f"{name} (Yahoo ticker {ticker}; sector: {sector or 'n/a'}; industry: {industry or 'n/a'}; "
        f"country: {country or 'n/a'}).\n"
        "Prefer similar business models, end markets and scale. Use each company's primary, liquid listing "
        "and give the exact Yahoo Finance ticker. Exclude the company itself, private companies, ETFs and funds."
    )
    result, used = generate_with_fallback(_api_key, models, prompt, PeerList)
    return [p.model_dump() for p in result.peers], used


@st.cache_data(ttl=CACHE_TTL, show_spinner=False)
def write_memo(ticker: str, metrics_json: str, headlines_json: str, description: str,
               models: tuple, _api_key: str) -> tuple[dict, str]:
    prompt = (
        f"Write the qualitative sections of an equity memo on {ticker}.\n\n"
        f"METRICS (computed in Python; ground truth, do not recalculate):\n{metrics_json}\n\n"
        f"BUSINESS DESCRIPTION:\n{description or 'Not available.'}\n\n"
        f"HEADLINES (JSON list; summarise only from title and snippet):\n{headlines_json}\n\n"
        "Requirements:\n"
        "- news: one entry for each headline that is relevant to the investment case, most material first, "
        "at most 6, using the headline's id.\n"
        "- opportunities: 3 or 4. threats: 3 or 4. risks: 3 to 5, each with severity and mitigant.\n"
        "- valuation_view and variant_perception must refer to the supplied valuation outputs "
        "(price versus DCF value, peer multiples, LBO return versus the hurdle, analyst target).\n"
        "- stance reflects the balance of valuation and fundamentals in the metrics."
    )
    result, used = generate_with_fallback(_api_key, models, prompt, QualitativeMemo)
    return result.model_dump(), used


# --------------------------------------------------------------------------- prompt inputs
def metrics_payload(d: CompanyData, r: dict) -> dict:
    """Compact, pre-formatted metrics for the prompt. Strings, so the model quotes rather than computes."""
    m, w, dcf, lbo, peers = r["market"], r["wacc"], r["dcf"], r["lbo"], r["peers"]
    rc, qc = d.report_currency, d.quote_currency
    drivers = r["projections"]["drivers"]
    out = {
        "company": {"name": d.name, "ticker": d.ticker, "sector": d.sector, "industry": d.industry,
                    "country": d.country},
        "market": {
            "share_price": fmt.price(m["price"], qc), "market_cap": fmt.money(m["market_cap"], rc),
            "enterprise_value": fmt.money(m["ev"], rc), "net_debt": fmt.money(m["net_debt"], rc),
            "pe_ltm": fmt.mult(m["pe"]), "pe_forward": fmt.mult(m["fwd_pe"]),
            "ev_ebitda_ltm": fmt.mult(m["ev_ebitda"]), "ev_revenue_ltm": fmt.mult(m["ev_revenue"]),
            "fcf_yield": fmt.pct(m["fcf_yield"]), "net_debt_to_ebitda": fmt.mult(m["net_debt_ebitda"]),
            "week52_range": f"{fmt.num(d.week52_low)} - {fmt.num(d.week52_high)} {qc}",
        },
        "fundamentals_ltm": {
            "revenue": fmt.money(m["revenue_ltm"], rc), "ebitda": fmt.money(m["ebitda_ltm"], rc),
            "ebitda_margin": fmt.pct(m["ebitda_margin"]), "net_income": fmt.money(m["net_income_ltm"], rc),
            "free_cash_flow": fmt.money(m["fcf_ltm"], rc),
            "revenue_growth_latest_quarter_yoy": fmt.pct(d.revenue_growth_yoy),
        },
        "wacc": {"wacc": fmt.pct(w["wacc"]), "cost_of_equity": fmt.pct(w["cost_of_equity"]),
                 "beta": fmt.num(w["beta"]), "after_tax_cost_of_debt": fmt.pct(w["after_tax_cost_of_debt"])},
    }
    if drivers:
        out["projection_assumptions"] = {
            "year1_revenue_growth": fmt.pct(drivers["growth_y1"]), "growth_basis": drivers["growth_source"],
            "terminal_growth": fmt.pct(drivers["terminal_growth"]),
            "ebitda_margin": fmt.pct(drivers["ebitda_margin_target"]),
        }
    if dcf.get("available"):
        out["dcf"] = {"value_per_share": fmt.price(dcf["price"], qc),
                      "upside_vs_price": fmt.pct(dcf["upside"], signed=True),
                      "terminal_value_share_of_ev": fmt.pct(dcf["tv_share"], 0),
                      "sensitivity_range": f"{fmt.num(dcf['range_low'])} - {fmt.num(dcf['range_high'])} {qc}"}
    else:
        out["dcf"] = {"not_available": dcf.get("reason", "")}
    if lbo.get("available"):
        out["lbo"] = {
            "structure": f"{lbo['leverage']:.1f}x senior debt / EBITDA at {lbo['rate']:.1%}, "
                         f"{lbo['years']}-year exit at {lbo['exit_multiple']:.1f}x EV/EBITDA",
            "sponsor_irr_at_current_price": fmt.pct(lbo["irr"]), "moic": fmt.mult(lbo["moic"]),
            "target_irr": fmt.pct(lbo["target_irr"], 0), "meets_target": "yes" if lbo["meets_target"] else "no",
            "max_price_for_target_irr": fmt.price(lbo["max_price"], qc),
            "max_price_vs_current": fmt.pct(lbo["max_price_premium"], signed=True)}
    else:
        out["lbo"] = {"not_available": lbo.get("reason", "")}
    if peers.get("available"):
        st_ = peers["stats"]
        out["peers"] = {
            "tickers": list(peers["table"]["ticker"]),
            "median_pe": fmt.mult(st_.get("pe", {}).get("median")),
            "median_ev_ebitda": fmt.mult(st_.get("ev_ebitda", {}).get("median")),
            "median_ebitda_margin": fmt.pct(st_.get("ebitda_margin", {}).get("median")),
            "subject_premium_to_median": {k: fmt.pct(v, signed=True) for k, v in peers["premium"].items()},
        }
    if d.target_mean:
        out["analyst_consensus"] = {"mean_target": fmt.price(d.target_mean, qc), "rating": d.recommendation,
                                    "analysts": d.analyst_count}
    if r["notes"]:
        out["model_notes"] = r["notes"][:6]
    return out


def headlines_payload(d: CompanyData) -> list[dict]:
    return [{"id": i, "date": n["date"], "source": n["publisher"], "title": n["title"], "snippet": n["snippet"]}
            for i, n in enumerate(d.news)]


def to_json(obj) -> str:
    return json.dumps(obj, ensure_ascii=False, separators=(",", ":"), default=str)


def merge_news(d: CompanyData, memo: dict | None) -> list[dict]:
    """Attach Gemini summaries to the Python-held headlines (titles, dates and links stay exact)."""
    items = [dict(n, summary=n["snippet"], sentiment="") for n in d.news]
    if not memo:
        return items
    ordered, used = [], set()
    for s in memo.get("news", []):
        i = s.get("headline_id")
        if isinstance(i, int) and 0 <= i < len(items) and i not in used:
            used.add(i)
            ordered.append(dict(items[i], summary=s.get("summary") or items[i]["snippet"],
                                sentiment=s.get("sentiment", "")))
    return ordered + [n for i, n in enumerate(items) if i not in used]
