"""Shared number formatting for the UI, the Gemini prompt and Excel labels."""
from __future__ import annotations

import math

NA = "n.m."


def _ok(v) -> bool:
    return v is not None and isinstance(v, (int, float)) and math.isfinite(v)


def money(v, ccy: str = "") -> str:
    """Compact money: 'USD 21.1bn'."""
    if not _ok(v):
        return NA
    a, sign = abs(v), "-" if v < 0 else ""
    if a >= 1e12:
        body = f"{a / 1e12:,.2f}tn"
    elif a >= 1e9:
        body = f"{a / 1e9:,.1f}bn"
    elif a >= 1e6:
        body = f"{a / 1e6:,.1f}m"
    else:
        body = f"{a:,.0f}"
    return f"{ccy} {sign}{body}".strip()


def pct(v, digits: int = 1, signed: bool = False) -> str:
    if not _ok(v):
        return NA
    return f"{v:+.{digits}%}" if signed else f"{v:.{digits}%}"


def mult(v) -> str:
    return f"{v:.1f}x" if _ok(v) else NA


def price(v, ccy: str = "") -> str:
    return f"{ccy} {v:,.2f}".strip() if _ok(v) else NA


def num(v, digits: int = 2) -> str:
    return f"{v:,.{digits}f}" if _ok(v) else NA
