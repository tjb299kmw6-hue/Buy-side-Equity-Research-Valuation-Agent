"""Excel memo writer (openpyxl). Turns engine results + Gemini text into a formatted workbook.

All numbers are values computed by the Python engine; the Assumptions tab documents
every input and its source. Blue font = input/assumption, black = engine output.
"""
from __future__ import annotations

import io
import math
import re
from datetime import datetime, timezone

import numpy as np
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from engine import PROJECTION_YEARS, Assumptions
from market_data import CompanyData

# Corporate palette --------------------------------------------------------------
NAVY, BLUE, ICE, GREY, LINE = "1F3864", "2F5597", "DCE6F2", "F3F5F8", "BFC8D6"
TEXT, MUTED, WHITE = "1A1A1A", "5F6B7A", "FFFFFF"
GREEN, RED, AMBER = "1E7B4B", "B03A2E", "9A6700"
GREEN_FILL, RED_FILL = "E3F1E8", "F8E4E1"
GOLD = "C9A227"    # second chart series (validated against BLUE for colour-blind separation)
INPUT = "0000FF"   # finance convention: blue = hard-coded input
FONT = "Arial"

F_MM = '#,##0;(#,##0);"-"'
F_PCT = '0.0%;(0.0%);"-"'
F_PCT2 = '0.00%;(0.00%);"-"'
F_X = '0.0"x"'
F_PX = '#,##0.00;(#,##0.00);"-"'
F_NUM2 = '0.00'
F_FACTOR = '0.000'
NM = "n.m."
MM = 1e6
_ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
DISCLAIMER = ("For information and education only. Not investment advice or a recommendation to buy or sell any "
              "security. Data from Yahoo Finance may be delayed or wrong; verify against company filings before relying on it.")


def _clean(v):
    if v is None:
        return None
    if isinstance(v, np.generic):
        v = v.item()
    if isinstance(v, float) and not math.isfinite(v):
        return None
    if isinstance(v, str):
        v = _ILLEGAL.sub("", v)
    return v


class Sheet:
    """Thin wrapper that keeps styling in one place."""

    def __init__(self, wb: Workbook, title: str, widths: list[float], landscape: bool = True):
        self.ws = wb.create_sheet(title)
        self.ws.sheet_view.showGridLines = False
        self.ws.sheet_properties.tabColor = NAVY
        for i, w in enumerate(widths, start=1):
            self.ws.column_dimensions[get_column_letter(i)].width = w
        self.widths = widths
        self.ws.page_setup.orientation = "landscape" if landscape else "portrait"
        self.ws.page_setup.fitToWidth, self.ws.page_setup.fitToHeight = 1, 0
        self.ws.sheet_properties.pageSetUpPr.fitToPage = True
        self.ws.print_options.horizontalCentered = True
        self.row = 1

    # -- primitives
    def cell(self, row, col, value=None, *, fmt=None, bold=False, italic=False, color=TEXT, size=10, fill=None,
             h=None, v="center", wrap=False, line=False, link=None, indent=0):
        c = self.ws.cell(row=row, column=col)
        value = _clean(value)
        if fmt and value is None:
            value, h, color = NM, "right", MUTED
        c.value = value
        if isinstance(value, str) and value.startswith("="):
            c.data_type = "s"   # never let text become a formula
        c.font = Font(name=FONT, size=size, bold=bold, italic=italic, color=color,
                      underline="single" if link else None)
        if fill:
            c.fill = PatternFill("solid", start_color=fill, end_color=fill)
        c.alignment = Alignment(horizontal=h or ("left" if indent else None), vertical=v, wrap_text=wrap, indent=indent)
        if fmt and not isinstance(value, str):
            c.number_format = fmt
        if line:
            c.border = Border(bottom=Side(style="thin", color=LINE))
        if link:
            c.hyperlink = link
        return c

    def merge(self, row, c1, c2, value=None, **style):
        rest = {k: v for k, v in style.items() if k not in ("fmt", "link")}
        self.cell(row, c1, value, **style)
        for col in range(c1 + 1, c2 + 1):
            self.cell(row, col, None, **rest)
        if c2 > c1:
            self.ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)

    def span(self, c1, c2) -> float:
        return sum(self.widths[c - 1] for c in range(c1, c2 + 1))

    def fit(self, row, blocks, size=10, minimum=15.0):
        """Set row height for wrapped text. blocks = [(text, c1, c2), ...]."""
        lines = 1
        for text, c1, c2 in blocks:
            per_line = max(8.0, self.span(c1, c2) * (1.15 if size <= 10 else 1.0))
            n = sum(max(1, math.ceil(len(p) / per_line)) for p in str(text or "").split("\n"))
            lines = max(lines, n)
        self.ws.row_dimensions[row].height = max(minimum, lines * (size + 3.2) + 4)

    # -- building blocks
    def title(self, text, subtitle, c1, c2):
        self.merge(1, c1, c2, text, bold=True, size=16, color=WHITE, fill=NAVY, h="left")
        self.ws.row_dimensions[1].height = 32
        self.merge(2, c1, c2, subtitle, size=10, color=NAVY, fill=ICE, h="left")
        self.ws.row_dimensions[2].height = 20
        self.row = 4

    def band(self, text, c1, c2):
        self.merge(self.row, c1, c2, text, bold=True, size=11, color=WHITE, fill=BLUE, h="left")
        self.ws.row_dimensions[self.row].height = 20
        self.row += 1

    def header(self, labels, c1, left=(0,)):
        """Column header row. `left` = indices of left-aligned (text) columns; the rest are right-aligned."""
        for i, lab in enumerate(labels):
            self.cell(self.row, c1 + i, lab, bold=True, size=9, color=NAVY, fill=ICE,
                      h="left" if i in left else "right", wrap=True)
        self.ws.row_dimensions[self.row].height = 18
        self.row += 1

    def kv_header(self, labels, c1, c2):
        """Header for label / value / basis tables, with the basis column running to c2."""
        self.cell(self.row, c1, labels[0], bold=True, size=9, color=NAVY, fill=ICE)
        self.cell(self.row, c1 + 1, labels[1], bold=True, size=9, color=NAVY, fill=ICE, h="right")
        self.merge(self.row, c1 + 2, c2, labels[2], bold=True, size=9, color=NAVY, fill=ICE, indent=1)
        self.ws.row_dimensions[self.row].height = 18
        self.row += 1

    def kv(self, label, value, fmt=None, basis="", *, c=2, basis_to=None, bold=False, is_input=False):
        """Label / value / basis row."""
        r = self.row
        self.cell(r, c, label, bold=bold, line=True)
        color = INPUT if is_input else TEXT
        if fmt is None:
            self.cell(r, c + 1, value, bold=bold, color=color, h="right", line=True, wrap=True)
        else:
            self.cell(r, c + 1, value, fmt=fmt, bold=bold, color=color, line=True)
        end = basis_to or c + 2
        self.merge(r, c + 2, end, basis, size=9, color=MUTED, wrap=True, line=True, indent=1)
        self.fit(r, [(basis, c + 2, end), (value if fmt is None else "", c + 1, c + 1)], size=9)
        self.row += 1

    def series(self, label, values, fmt, c=2, *, bold=False, italic=False, color=TEXT, indent=False, fill=None):
        r = self.row
        self.cell(r, c, ("   " + label) if indent else label, bold=bold, italic=italic, line=True, fill=fill,
                  color=MUTED if italic else TEXT)
        for i, v in enumerate(values):
            if v == "":
                self.cell(r, c + 1 + i, None, line=True, fill=fill)
            else:
                self.cell(r, c + 1 + i, v, fmt=fmt, bold=bold, italic=italic, line=True, fill=fill,
                          color=MUTED if italic else color)
        self.row += 1

    def text(self, value, c1, c2, *, size=10, color=TEXT, italic=False, bold=False, fill=None):
        self.merge(self.row, c1, c2, value, size=size, color=color, italic=italic, bold=bold, wrap=True, v="top", fill=fill)
        self.fit(self.row, [(value, c1, c2)], size=size)
        self.row += 1

    def grid(self, frame, row_fmt, col_fmt, cell_fmt, corner, c=2, *, base=None, colour=None):
        """Sensitivity table. colour(value) -> (font colour, fill) or None."""
        r0 = self.row
        self.cell(r0, c, corner, bold=True, size=9, color=NAVY, fill=ICE, wrap=True)
        for j, col in enumerate(frame.columns):
            self.cell(r0, c + 1 + j, float(col), fmt=col_fmt, bold=True, color=NAVY, fill=ICE, h="right")
        for i, idx in enumerate(frame.index):
            self.cell(r0 + 1 + i, c, float(idx), fmt=row_fmt, bold=True, color=NAVY, fill=ICE, h="right")
            for j, col in enumerate(frame.columns):
                v = _clean(frame.iloc[i, j])
                style = colour(v) if (colour and v is not None) else None
                is_base = base == (i, j)
                self.cell(r0 + 1 + i, c + 1 + j, v, fmt=cell_fmt, bold=is_base, line=True,
                          color=style[0] if style else TEXT, fill=(ICE if is_base and not (style and style[1]) else (style[1] if style else None)))
        self.row = r0 + len(frame.index) + 2


def _mm(v):
    return None if v is None else v / MM


# =============================================================================== sheets
def _memo_sheet(wb, d: CompanyData, r: dict, memo: dict | None, news: list, meta: dict):
    s = Sheet(wb, "Memo", [2, 21, 15, 15, 15, 21, 15, 15, 28])
    m, w, dcf, lbo = r["market"], r["wacc"], r["dcf"], r["lbo"]
    qc, rc = d.quote_currency, d.report_currency
    L, R = 2, 9
    sub = " | ".join(x for x in ("Equity valuation memo", d.exchange, d.sector, d.industry) if x)
    s.title(f"{d.name} ({d.ticker})", sub, L, R)
    s.merge(3, L, R, f"As of {d.as_of}.  Share prices in {qc}; financials in {rc} millions unless stated.",
            size=9, color=MUTED, italic=True)

    s.band("Key metrics", L, R)
    tiles = [
        (f"Share price ({qc})", m["price"], F_PX, NAVY),
        (f"Market cap ({rc} m)", _mm(m["market_cap"]), F_MM, NAVY),
        (f"Enterprise value ({rc} m)", _mm(m["ev"]), F_MM, NAVY),
        ("P/E (LTM)", m["pe"], F_X, NAVY),
        ("EV / EBITDA (LTM)", m["ev_ebitda"], F_X, NAVY),
        ("FCF yield (LTM FCF / market cap)", m["fcf_yield"], F_PCT, NAVY),
        ("WACC", w["wacc"], F_PCT, NAVY),
        (f"DCF value per share ({qc})", dcf.get("price"), F_PX, NAVY),
        ("DCF upside / (downside)", dcf.get("upside"), F_PCT, None),
        ("LBO sponsor IRR at current price", lbo.get("irr"), F_PCT, None),
        (f"Max LBO price at {meta['target_irr']:.0%} IRR ({qc})", lbo.get("max_price"), F_PX, NAVY),
        ("Net debt / EBITDA", m["net_debt_ebitda"], F_X, NAVY),
    ]
    for i, (label, value, fmt, colour) in enumerate(tiles):
        row, col = s.row + 2 * (i // 4), L + 2 * (i % 4)
        if colour is None:   # signed metrics: green when favourable
            hurdle = meta["target_irr"] if "IRR" in label else 0.0
            colour = MUTED if value is None else (GREEN if value >= hurdle else RED)
        s.merge(row, col, col + 1, label, size=8, color=MUTED, fill=GREY, h="left")
        s.merge(row + 1, col, col + 1, value, fmt=fmt, bold=True, size=13, color=colour, fill=GREY, h="left")
        s.ws.row_dimensions[row + 1].height = 22
    s.row += 7

    s.band(f"Valuation summary (per share, {qc})", L, R)
    _summary_table(s, r, d, L, R)
    s.row += 1

    s.band("Investment view", L, R)
    if memo:
        stance_colour = {"Constructive": GREEN, "Cautious": RED}.get(memo.get("stance"), AMBER)
        s.cell(s.row, L, "Stance", bold=True)
        s.merge(s.row, L + 1, R, memo.get("stance"), bold=True, color=stance_colour)
        s.row += 1
        for label, key in (("Business", "business_summary"), ("Thesis", "investment_thesis"),
                           ("Variant perception", "variant_perception"), ("Valuation view", "valuation_view")):
            s.cell(s.row, L, label, bold=True, v="top")
            s.merge(s.row, L + 1, R, memo.get(key), wrap=True, v="top")
            s.fit(s.row, [(memo.get(key), L + 1, R)])
            s.row += 1
    else:
        s.text(meta["llm_note"], L, R, color=MUTED, italic=True)
    s.row += 1

    if memo:
        mid = 5
        s.merge(s.row, L, mid, "Key opportunities", bold=True, size=11, color=WHITE, fill=BLUE)
        s.merge(s.row, mid + 1, R, "Key threats", bold=True, size=11, color=WHITE, fill=BLUE)
        s.ws.row_dimensions[s.row].height = 20
        s.row += 1
        opps, threats = memo.get("opportunities", []), memo.get("threats", [])
        for i in range(max(len(opps), len(threats))):
            blocks = []
            for items, c1, c2 in ((opps, L, mid), (threats, mid + 1, R)):
                if i < len(items):
                    s.cell(s.row, c1, items[i].get("title"), bold=True, wrap=True, v="top", line=True)
                    s.merge(s.row, c1 + 1, c2, items[i].get("detail"), wrap=True, v="top", line=True)
                    blocks += [(items[i].get("title"), c1, c1), (items[i].get("detail"), c1 + 1, c2)]
            s.fit(s.row, blocks)
            s.row += 1
        s.row += 1

    s.band("Recent news", L, R)
    if news:
        r0 = s.row
        s.cell(r0, L, "Date", bold=True, size=9, color=NAVY, fill=ICE)
        s.cell(r0, L + 1, "Sentiment", bold=True, size=9, color=NAVY, fill=ICE)
        s.merge(r0, L + 2, L + 4, "Headline", bold=True, size=9, color=NAVY, fill=ICE)
        s.merge(r0, L + 5, R, "Summary", bold=True, size=9, color=NAVY, fill=ICE)
        s.row += 1
        for n in news[:5]:
            sent = n.get("sentiment") or ""
            s.cell(s.row, L, n.get("date"), v="top", line=True)
            s.cell(s.row, L + 1, sent, v="top", line=True, bold=True,
                   color={"Positive": GREEN, "Negative": RED}.get(sent, MUTED))
            s.merge(s.row, L + 2, L + 4, n.get("title"), wrap=True, v="top", line=True)
            s.merge(s.row, L + 5, R, n.get("summary"), wrap=True, v="top", line=True)
            s.fit(s.row, [(n.get("title"), L + 2, L + 4), (n.get("summary"), L + 5, R)])
            s.row += 1
    else:
        s.text("No recent headlines returned by Yahoo Finance.", L, R, color=MUTED, italic=True)
    s.row += 1

    if r["notes"]:
        s.band("Model notes", L, R)
        for note in r["notes"]:
            s.text(note, L, R, size=9)
        s.row += 1
    s.text(f"{meta['footer']} {DISCLAIMER}", L, R, size=8, color=MUTED, italic=True)
    s.ws.freeze_panes = "A4"


def _summary_table(s: Sheet, r: dict, d: CompanyData, L: int, R: int):
    """Valuation summary. Layout: method (2 cols) | low | base | high | vs price | basis (to R)."""
    r0 = s.row
    s.merge(r0, L, L + 1, "Method", bold=True, size=9, color=NAVY, fill=ICE)
    for i, lab in enumerate(("Low", "Base", "High", "Base vs price")):
        s.cell(r0, L + 2 + i, lab, bold=True, size=9, color=NAVY, fill=ICE, h="right")
    s.merge(r0, L + 6, R, "Basis", bold=True, size=9, color=NAVY, fill=ICE, indent=1)
    s.row += 1
    for row in r["summary"]:
        s.merge(s.row, L, L + 1, row["method"], line=True)
        for col, key in ((L + 2, "low"), (L + 3, "base"), (L + 4, "high")):
            if row[key] is None:
                s.cell(s.row, col, None, line=True)
            else:
                s.cell(s.row, col, row[key], fmt=F_PX, bold=(key == "base"), line=True,
                       color=TEXT if key == "base" else MUTED)
        vs = row["vs_price"]
        s.cell(s.row, L + 5, vs, fmt=F_PCT, line=True, color=GREEN if (vs or 0) >= 0 else RED)
        s.merge(s.row, L + 6, R, row["basis"], size=9, color=MUTED, line=True, indent=1)
        s.row += 1
    s.merge(s.row, L, L + 1, "Current share price", bold=True, fill=GREY)
    s.cell(s.row, L + 2, None, fill=GREY)
    s.cell(s.row, L + 3, d.price, fmt=F_PX, bold=True, fill=GREY)
    for col in range(L + 4, R + 1):
        s.cell(s.row, col, None, fill=GREY)
    s.row += 1


def _valuation_sheet(wb, d: CompanyData, r: dict, meta: dict):
    s = Sheet(wb, "Valuation", [2, 34, 16, 14, 14, 14, 14, 44])
    m, w = r["market"], r["wacc"]
    qc, rc = d.quote_currency, d.report_currency
    L, R = 2, 8
    s.title("Market snapshot, WACC and valuation summary", f"{d.name} ({d.ticker})", L, R)

    s.band("Market snapshot", L, R)
    s.kv_header(["Item", "Value", "Basis"], L, R)
    kv = lambda *a, **k: s.kv(*a, basis_to=R, **k)   # noqa: E731
    kv(f"Share price ({qc})", m["price"], F_PX, "Yahoo Finance last price")
    kv("Effective shares (m)", _mm(d.shares), F_MM, "Market cap / share price (captures all share classes and ADR ratios)")
    kv(f"Market capitalisation ({rc} m)", _mm(m["market_cap"]), F_MM,
       "Yahoo Finance" + ("" if d.trade_currency == rc else f", converted from {d.trade_currency} at {d.fx_trade_to_report:.4f}"), bold=True)
    kv("Plus: total debt", _mm(m["total_debt"]), F_MM, f"Most recent balance sheet ({d.balance_sheet_date or 'latest'}), includes lease liabilities")
    kv("Less: cash and short-term investments", _mm(m["cash"]), F_MM, "Most recent balance sheet")
    kv("Plus: minority interest", _mm(m["minority_interest"]), F_MM, "Most recent balance sheet")
    kv("Plus: preferred equity", _mm(m["preferred"]), F_MM, "Most recent balance sheet")
    kv(f"Enterprise value ({rc} m)", _mm(m["ev"]), F_MM, "Market cap + debt - cash + minority interest + preferred", bold=True)
    kv("LTM revenue", _mm(m["revenue_ltm"]), F_MM, d.ltm_basis)
    kv("LTM EBITDA", _mm(m["ebitda_ltm"]), F_MM, d.ltm_basis)
    kv("LTM net income to common", _mm(m["net_income_ltm"]), F_MM, d.ltm_basis)
    kv("LTM free cash flow", _mm(m["fcf_ltm"]), F_MM, "Operating cash flow - capital expenditure")
    kv("P/E (LTM)", m["pe"], F_X, m["pe_source"], bold=True)
    kv("P/E (forward)", m["fwd_pe"], F_X, "Yahoo Finance consensus forward P/E")
    kv("EV / EBITDA (LTM)", m["ev_ebitda"], F_X, "Enterprise value / LTM EBITDA", bold=True)
    kv("EV / revenue (LTM)", m["ev_revenue"], F_X, "Enterprise value / LTM revenue")
    kv("FCF yield", m["fcf_yield"], F_PCT, "LTM free cash flow / market cap", bold=True)
    kv("EBITDA margin (LTM)", m["ebitda_margin"], F_PCT, "LTM EBITDA / LTM revenue")
    kv("Net debt / EBITDA", m["net_debt_ebitda"], F_X, "(Total debt - cash) / LTM EBITDA")
    s.row += 1

    s.band("Weighted average cost of capital", L, R)
    s.kv_header(["Component", "Value", "Basis"], L, R)
    kv("Risk-free rate", w["risk_free"], F_PCT2, w["risk_free_source"], is_input=True)
    kv("Equity risk premium", w["erp"], F_PCT2, "Assumption (adjustable in the app sidebar)", is_input=True)
    kv("Beta (adjusted)", w["beta"], F_NUM2, w["beta_source"], is_input=True)
    kv("Cost of equity", w["cost_of_equity"], F_PCT2, "CAPM: risk-free + beta x equity risk premium", bold=True)
    kv("Pre-tax cost of debt", w["pre_tax_cost_of_debt"], F_PCT2, w["cost_of_debt_source"], is_input=True)
    kv("Tax rate", w["tax_rate"], F_PCT, w["tax_source"], is_input=True)
    kv("After-tax cost of debt", w["after_tax_cost_of_debt"], F_PCT2, "Pre-tax cost of debt x (1 - tax rate)")
    kv(f"Equity value ({rc} m)", _mm(w["equity_value"]), F_MM, "Market capitalisation")
    kv(f"Debt value ({rc} m)", _mm(w["debt_value"]), F_MM, "Total debt (book value as proxy for market value)")
    kv("Weight of equity", w["weight_equity"], F_PCT, "E / (D + E)")
    kv("Weight of debt", w["weight_debt"], F_PCT, "D / (D + E)")
    kv("WACC", w["wacc"], F_PCT2, "E/(D+E) x cost of equity + D/(D+E) x after-tax cost of debt", bold=True)
    s.row += 1

    s.band(f"Valuation summary (per share, {qc})", L, R)
    _summary_table(s, r, d, L, R)
    ranged = [(i, x) for i, x in enumerate(r["summary"]) if x["low"] is not None and x["high"] is not None]
    if len(ranged) >= 2:
        # Football field: transparent "low" bar stacked under a visible "high - low" bar.
        s.row += 1
        h0 = s.row
        for j, lab in enumerate(("Chart data", "Low", "Range (high - low)")):
            s.cell(h0, L + j, lab, size=8, color=MUTED, italic=True, h="left" if j == 0 else "right")
        for k, (_, x) in enumerate(ranged, start=1):
            s.cell(h0 + k, L, x["method"], size=8, color=MUTED)
            s.cell(h0 + k, L + 1, x["low"], fmt=F_PX, size=8, color=MUTED)
            s.cell(h0 + k, L + 2, x["high"] - x["low"], fmt=F_PX, size=8, color=MUTED)
        n = len(ranged)
        chart = BarChart()
        chart.type, chart.grouping, chart.overlap = "bar", "stacked", 100
        chart.title = f"Valuation range per share ({qc}) - current price {d.price:,.2f}"
        chart.add_data(Reference(s.ws, min_col=L + 1, max_col=L + 2, min_row=h0, max_row=h0 + n), titles_from_data=True)
        chart.set_categories(Reference(s.ws, min_col=L, min_row=h0 + 1, max_row=h0 + n))
        chart.series[0].graphicalProperties.noFill = True
        chart.series[0].graphicalProperties.line.noFill = True
        chart.series[1].graphicalProperties.solidFill = BLUE
        chart.series[1].graphicalProperties.line.solidFill = NAVY
        chart.legend = None
        chart.y_axis.numFmt = "#,##0"
        chart.y_axis.delete, chart.x_axis.delete = False, False
        chart.x_axis.scaling.orientation = "maxMin"   # first method at the top
        chart.y_axis.crosses = "max"                  # keep the value axis at the bottom
        chart.height, chart.width = 7.5, 17
        s.ws.add_chart(chart, f"E{h0}")
        s.row = h0 + max(n + 2, 17)


def _dcf_sheet(wb, d: CompanyData, r: dict):
    s = Sheet(wb, "DCF", [2, 38, 14, 14, 14, 14, 14, 14, 30])
    dcf, proj, m = r["dcf"], r["projections"]["table"], r["market"]
    qc, rc = d.quote_currency, d.report_currency
    L, R = 2, 9
    s.title("Discounted cash flow: 5-year unlevered FCF", f"{d.name} ({d.ticker}) | {rc} millions unless stated", L, R)
    if proj.empty:
        s.text(dcf.get("reason", "DCF not available."), L, R, color=RED)
        return

    years = [f"Year {t}" for t in proj.index]
    s.band(f"Free cash flow projection ({rc} m)", L, R - 1)
    s.header(["", "LTM"] + years, L)
    col = lambda name, scale=MM: [v / scale for v in proj[name]]   # noqa: E731
    ebit_ltm = (d.ebitda_ltm - d.da_ltm) if d.da_ltm is not None else None
    blank = ""
    s.series("Revenue", [_mm(d.revenue_ltm)] + col("revenue"), F_MM, bold=True)
    s.series("growth", [blank] + list(proj["growth"]), F_PCT, italic=True, indent=True)
    s.series("EBITDA", [_mm(d.ebitda_ltm)] + col("ebitda"), F_MM, bold=True)
    s.series("margin", [m["ebitda_margin"]] + list(proj["ebitda_margin"]), F_PCT, italic=True, indent=True)
    s.series("Less: depreciation and amortisation", [_mm(d.da_ltm) if d.da_ltm is not None else blank] + col("da"), F_MM)
    s.series("EBIT", [_mm(ebit_ltm) if ebit_ltm is not None else blank] + col("ebit"), F_MM, bold=True)
    s.series("Less: taxes on EBIT", [blank] + col("taxes"), F_MM)
    s.series("NOPAT", [blank] + col("nopat"), F_MM)
    s.series("Plus: depreciation and amortisation", [blank] + col("da"), F_MM)
    s.series("Less: capital expenditure", [_mm(d.capex_ltm) if d.capex_ltm is not None else blank] + col("capex"), F_MM)
    s.series("Less: increase in net working capital", [blank] + col("d_nwc"), F_MM)
    s.series("Unlevered free cash flow", [blank] + col("ufcf"), F_MM, bold=True, fill=ICE)
    if dcf.get("available"):
        s.series("Discount period (years)", [blank] + list(dcf["periods"]), F_NUM2, italic=True)
        s.series("Discount factor", [blank] + list(dcf["factors"]), F_FACTOR, italic=True)
        s.series("Present value of UFCF", [blank] + [v / MM for v in dcf["pv_ufcf"]], F_MM, bold=True)
    s.row += 1
    if not dcf.get("available"):
        s.text(f"DCF valuation not produced: {dcf.get('reason', '')}", L, R, color=RED)
        return

    s.band("Valuation bridge", L, R)
    s.kv_header(["Item", "Value", "Basis"], L, R)
    kv = lambda *a, **k: s.kv(*a, basis_to=R, **k)   # noqa: E731
    kv("WACC", dcf["wacc"], F_PCT2, "See Valuation tab", is_input=True)
    kv("Terminal growth rate", dcf["terminal_growth"], F_PCT2, "Assumption (adjustable in the app sidebar)", is_input=True)
    kv("Sum of PV of UFCF (years 1-5)", _mm(dcf["sum_pv"]), F_MM,
       "Mid-year convention" if dcf["mid_year"] else "End-of-year discounting")
    kv("Terminal value (year 5)", _mm(dcf["tv"]), F_MM, "Gordon growth: UFCF5 x (1 + g) / (WACC - g)")
    kv("PV of terminal value", _mm(dcf["pv_tv"]), F_MM,
       "Discounted 4.5 years (mid-year convention)" if dcf["mid_year"] else "Discounted 5 years")
    kv(f"Enterprise value ({rc} m)", _mm(dcf["ev"]), F_MM, "PV of UFCF + PV of terminal value", bold=True)
    kv("Less: net debt", _mm(m["net_debt"]), F_MM, "Total debt - cash")
    kv("Less: minority interest and preferred", _mm(m["minority_interest"] + m["preferred"]), F_MM, "Most recent balance sheet")
    kv(f"Equity value ({rc} m)", _mm(dcf["equity"]), F_MM, "", bold=True)
    kv("Effective shares (m)", _mm(d.shares), F_MM, "Market cap / share price")
    if d.trade_currency != rc or d.quote_factor != 1:
        kv(f"{d.trade_currency}/{rc} rate; quote units per {d.trade_currency}", d.fx_trade_to_report, F_FACTOR,
           f"Spot FX; price quoted in {qc} ({d.quote_factor:.0f} per {d.trade_currency})")
    kv(f"DCF value per share ({qc})", dcf["price"], F_PX, "Equity value / effective shares", bold=True)
    kv(f"Current share price ({qc})", d.price, F_PX, "Yahoo Finance")
    kv("Upside / (downside)", dcf["upside"], F_PCT, "DCF value / price - 1", bold=True)
    kv("Terminal value as % of EV", dcf["tv_share"], F_PCT, "Higher = more sensitive to terminal assumptions")
    kv("Implied terminal EV/EBITDA", dcf["implied_exit_multiple"], F_X, "Terminal value / year-5 EBITDA")
    s.row += 1

    x = dcf.get("exit_check")
    if x:
        s.band("Cross-check: exit multiple method", L, R)
        s.kv_header(["Item", "Value", "Basis"], L, R)
        kv("Exit EV/EBITDA multiple", x["exit_multiple"], F_X, "Current LTM EV/EBITDA")
        kv("Terminal value (year 5)", _mm(x["tv"]), F_MM, "Year-5 EBITDA x exit multiple")
        kv(f"Enterprise value ({rc} m)", _mm(x["ev"]), F_MM, "PV of UFCF + terminal value discounted 5 years")
        kv(f"Value per share ({qc})", x["price"], F_PX, "", bold=True)
        kv("Implied perpetual growth", x["implied_growth"], F_PCT, "Growth rate that reconciles this terminal value to Gordon growth")
        s.row += 1

    s.band(f"Sensitivity: DCF value per share ({qc}). Rows = WACC, columns = terminal growth. Green = above current price", L, R)
    s.grid(dcf["sensitivity"], F_PCT, F_PCT, F_PX, "WACC / growth", base=(2, 2),
           colour=lambda v: (GREEN if v >= d.price else RED, None))
    s.ws.freeze_panes = "C4"


def _lbo_sheet(wb, d: CompanyData, r: dict):
    s = Sheet(wb, "LBO", [2, 38, 14, 14, 14, 14, 14, 14, 30])
    lbo, m = r["lbo"], r["market"]
    qc, rc = d.quote_currency, d.report_currency
    L, R = 2, 9
    s.title("Leveraged buyout: sponsor returns", f"{d.name} ({d.ticker}) | {rc} millions unless stated", L, R)
    if not lbo.get("available"):
        s.text(f"LBO not produced: {lbo.get('reason', '')}", L, R, color=RED)
        return
    kv = lambda *a, **k: s.kv(*a, basis_to=R, **k)   # noqa: E731
    tgt = lbo["target_irr"]

    s.band("Transaction assumptions", L, R)
    s.kv_header(["Item", "Value", "Basis"], L, R)
    kv("LTM EBITDA", _mm(lbo["ebitda_ltm"]), F_MM, d.ltm_basis)
    kv("Entry EV/EBITDA", lbo["entry_multiple"], F_X, "Purchase at current enterprise value (no control premium)")
    kv("Senior debt / EBITDA", lbo["leverage"], F_X, "Input", is_input=True)
    kv("Interest rate on senior debt", lbo["rate"], F_PCT2, "Input; cash interest on opening balance", is_input=True)
    kv("Exit EV/EBITDA", lbo["exit_multiple"], F_X, "Input; base case = entry multiple", is_input=True)
    kv("Holding period (years)", lbo["years"], "0", "Input", is_input=True)
    kv("Minimum target IRR", tgt, F_PCT, "Input", is_input=True)
    kv("Tax rate", r["wacc"]["tax_rate"], F_PCT, "As WACC; interest is tax-deductible")
    s.row += 1

    s.band("Sources and uses at entry", L, R)
    s.kv_header(["Item", f"{rc} m", "Basis"], L, R)
    kv("Senior debt", _mm(lbo["debt0"]), F_MM, f"{lbo['leverage_used']:.1f}x LTM EBITDA = {lbo['ltv']:.0%} of entry EV")
    kv("Sponsor equity", _mm(lbo["equity0"]), F_MM, "Entry EV - senior debt")
    kv("Total sources", _mm(lbo["entry_ev"]), F_MM, "", bold=True)
    kv("Purchase of equity", _mm(m["market_cap"]), F_MM, "At current market capitalisation")
    kv("Refinance net debt and other claims", _mm(lbo["entry_ev"] - m["market_cap"]), F_MM, "Net debt + minority interest + preferred")
    kv("Total uses", _mm(lbo["entry_ev"]), F_MM, "No transaction fees assumed", bold=True)
    s.row += 1

    sch = lbo["schedule"]
    s.band(f"Debt schedule ({rc} m): 100% cash sweep", L, R - 1)
    s.header(["", "Entry"] + [f"Year {t}" for t in sch.index], L)
    c = lambda name: [v / MM for v in sch[name]]   # noqa: E731
    blank = ""
    s.series("EBITDA", [_mm(lbo["ebitda_ltm"])] + c("ebitda"), F_MM, bold=True)
    s.series("Less: cash interest", [blank] + c("interest"), F_MM)
    s.series("Less: cash taxes", [blank] + c("taxes"), F_MM)
    s.series("Less: capital expenditure", [blank] + c("capex"), F_MM)
    s.series("Less: increase in net working capital", [blank] + c("d_nwc"), F_MM)
    s.series("Free cash flow for debt repayment", [blank] + c("fcf"), F_MM, bold=True, fill=ICE)
    s.series("Senior debt, end of year", [_mm(lbo["debt0"])] + c("debt_end"), F_MM, bold=True)
    s.series("Cash, end of year", [0.0] + c("cash_end"), F_MM)
    s.series("Debt / EBITDA", [lbo["leverage_used"]] + list(sch["leverage"]), F_X, italic=True)
    s.row += 1

    s.band("Returns", L, R)
    s.kv_header(["Item", "Value", "Basis"], L, R)
    kv("Exit enterprise value", _mm(lbo["exit_ev"]), F_MM, f"Year-{lbo['years']} EBITDA x {lbo['exit_multiple']:.1f}x")
    kv("Less: senior debt at exit", _mm(lbo["exit_debt"]), F_MM, "")
    kv("Plus: cash at exit", _mm(lbo["exit_cash"]), F_MM, "")
    kv("Exit equity value", _mm(lbo["exit_equity"]), F_MM, "", bold=True)
    kv("Sponsor equity invested", _mm(lbo["equity0"]), F_MM, "")
    kv("Multiple of invested capital", lbo["moic"], F_X, "Exit equity / equity invested", bold=True)
    kv("Sponsor IRR at current price", lbo["irr"], F_PCT, f"MOIC ^ (1/{lbo['years']}) - 1; no interim distributions", bold=True)
    kv(f"Meets {tgt:.0%} target IRR?", "Yes" if lbo["meets_target"] else "No", None, "At the current enterprise value", bold=True)
    kv(f"Max entry EV for {tgt:.0%} IRR ({rc} m)", _mm(lbo["max_entry_ev"]), F_MM, "Highest purchase price that still earns the target IRR")
    kv("Implied max entry EV/EBITDA", lbo["max_entry_multiple"], F_X, "")
    kv(f"Implied max price per share ({qc})", lbo["max_price"], F_PX, "(Max entry EV - net debt and other claims) / shares", bold=True)
    kv("Premium / (discount) to current price", lbo["max_price_premium"], F_PCT, "Negative = a sponsor could not pay today's price and earn the target")
    s.row += 1

    colour = lambda v: (GREEN, GREEN_FILL) if v >= tgt else (RED, RED_FILL)   # noqa: E731
    s.band(f"Sensitivity: sponsor IRR. Rows = entry EV/EBITDA, columns = exit EV/EBITDA. Green = meets {tgt:.0%} target", L, R)
    s.grid(lbo["sens_entry_exit"], F_X, F_X, F_PCT, "Entry / exit", base=(2, 2), colour=colour)
    s.band(f"Sensitivity: sponsor IRR. Rows = senior debt / EBITDA, columns = interest rate. Green = meets {tgt:.0%} target", L, R)
    s.grid(lbo["sens_leverage_rate"], F_X, F_PCT, F_PCT, "Leverage / rate", base=(2, 2), colour=colour)
    s.ws.freeze_panes = "C4"


def _peers_sheet(wb, d: CompanyData, r: dict, meta: dict):
    s = Sheet(wb, "Peers", [2, 12, 36, 8, 16, 16, 10, 10, 12, 12, 14, 14])
    peers = r["peers"]
    qc = d.quote_currency
    L, R = 2, 12
    s.title("Peer valuation multiples", f"{d.name} ({d.ticker}) versus comparable companies", L, R)
    s.band("Trading multiples", L, R)
    s.header(["Ticker", "Company", "Ccy", "Market cap (m)", "EV (m)", "P/E", "Fwd P/E", "EV/EBITDA", "EV/Revenue",
              "EBITDA margin", "Rev growth"], L, left=(0, 1, 2))
    fmts = [None, None, None, F_MM, F_MM, F_X, F_X, F_X, F_X, F_PCT, F_PCT]
    keys = ["ticker", "name", "currency", "market_cap", "ev", "pe", "fwd_pe", "ev_ebitda", "ev_revenue",
            "ebitda_margin", "revenue_growth"]

    def write(row: dict, bold=False, fill=None):
        for i, (k, f) in enumerate(zip(keys, fmts)):
            v = row.get(k)
            if k in ("market_cap", "ev") and v is not None:
                v = v / MM
            if f is None:
                s.cell(s.row, L + i, v, bold=bold, fill=fill, line=True)
            else:
                s.cell(s.row, L + i, v, fmt=f, bold=bold, fill=fill, line=True)
        s.row += 1

    write(peers["subject"], bold=True, fill=ICE)
    for _, row in peers["table"].iterrows():
        write({k: (None if isinstance(v, float) and math.isnan(v) else v) for k, v in row.to_dict().items()})
    if not peers["available"]:
        s.text("No usable peer data was returned. Enter peer tickers in the app sidebar and re-run.", L, R, color=MUTED, italic=True)
        return
    s.row += 1
    for label, stat in (("Peer median", "median"), ("Peer mean", "mean"), ("Peer high", "high"), ("Peer low", "low")):
        row = {"ticker": label, "name": "", "currency": ""}
        row.update({k: peers["stats"].get(k, {}).get(stat) for k in keys[5:]})
        for i, (k, f) in enumerate(zip(keys, fmts)):
            if k in ("market_cap", "ev"):
                s.cell(s.row, L + i, None, fill=GREY)
            elif f is None:
                s.cell(s.row, L + i, row.get(k), bold=True, fill=GREY)
            else:
                s.cell(s.row, L + i, row.get(k), fmt=f, bold=(stat == "median"), fill=GREY)
        s.row += 1
    s.cell(s.row, L, f"{d.ticker} premium / (discount) to peer median", bold=True)
    for i, k in enumerate(keys):
        if k in peers["premium"]:
            v = peers["premium"][k]
            s.cell(s.row, L + i, v, fmt=F_PCT, bold=True, color=RED if v > 0 else GREEN)
    s.row += 2

    if peers["implied"]:
        s.band(f"Peer-implied value per share ({qc})", L, R)
        s.header(["Method", "", "", "25th pct", "Median", "75th pct", "", "vs price"], L)
        for key, label in (("ev_ebitda", "EV/EBITDA x LTM EBITDA, less net debt"), ("pe", "P/E x LTM EPS")):
            imp = peers["implied"].get(key)
            if not imp:
                continue
            s.merge(s.row, L, L + 2, label, line=True)
            for j, k in enumerate(("low", "base", "high")):
                s.cell(s.row, L + 3 + j, imp[k], fmt=F_PX, bold=(k == "base"), line=True)
            s.cell(s.row, L + 6, None, line=True)
            vs = imp["base"] / d.price - 1
            s.cell(s.row, L + 7, vs, fmt=F_PCT, line=True, color=GREEN if vs >= 0 else RED)
            s.row += 1
        s.row += 1

    s.band("Peer selection", L, R)
    s.text(f"Source: {meta['peer_source']}. Multiples are calculated in Python from Yahoo Finance data: "
           "EV = market cap + total debt - cash; P/E = market cap / LTM net income; revenue growth = latest quarter "
           "year on year. Market cap and EV are shown in each company's reporting currency. Multiples above 200x P/E "
           "or 100x EV/EBITDA are treated as not meaningful.", L, R, size=9, color=MUTED)
    for p in meta.get("peer_rationale", []):
        s.cell(s.row, L, p.get("ticker"), bold=True, v="top")
        s.merge(s.row, L + 1, R, p.get("rationale"), wrap=True, v="top", size=9)
        s.fit(s.row, [(p.get("rationale"), L + 1, R)], size=9)
        s.row += 1
    if meta.get("peers_dropped"):
        s.text("Suggested but excluded (no usable Yahoo data): " + ", ".join(meta["peers_dropped"]), L, R, size=9, color=MUTED)
    s.ws.freeze_panes = "D6"


def _projections_sheet(wb, d: CompanyData, r: dict):
    h, proj, drivers = d.hist, r["projections"]["table"], r["projections"]["drivers"]
    hist_cols = list(h.columns) if h is not None and not h.empty else []
    n_proj = len(proj)
    widths = [2, 36] + [13] * (len(hist_cols) + 1 + max(n_proj, PROJECTION_YEARS)) + [3]
    s = Sheet(wb, "Projections", widths)
    qc, rc = d.quote_currency, d.report_currency
    L = 2
    R = L + len(hist_cols) + 1 + max(n_proj, 1)
    s.title("Historical financials and 5-year projections", f"{d.name} ({d.ticker}) | {rc} millions unless stated", L, R)

    s.band("Income and cash flow", L, R)
    r_group = s.row
    s.cell(r_group, L, None, fill=GREY)
    for i in range(len(hist_cols)):
        s.cell(r_group, L + 1 + i, "Actual", size=8, color=MUTED, fill=GREY, h="right")
    s.cell(r_group, L + 1 + len(hist_cols), "Actual", size=8, color=MUTED, fill=GREY, h="right")
    for i in range(n_proj):
        s.cell(r_group, L + 2 + len(hist_cols) + i, "Projected", size=8, color=INPUT, fill=GREY, h="right")
    s.row += 1
    header_row = s.row
    s.header(["Period"] + hist_cols + ["LTM"] + [f"Year {t}" for t in proj.index], L)

    def hist(item, scale=MM):
        if not hist_cols or item not in h.index:
            return [None] * len(hist_cols)
        return [None if math.isnan(v) else v / scale for v in h.loc[item]]

    def growth(vals):
        out = [""]
        for a, b in zip(vals[:-1], vals[1:]):
            out.append(b / a - 1 if a and b is not None and a > 0 else None)
        return out

    def ratio(num, den):
        return [(a / b if a is not None and b else None) for a, b in zip(num, den)]

    p = lambda name, scale=MM: [v / scale for v in proj[name]] if n_proj else []   # noqa: E731
    rev_h, ebitda_h = hist("revenue"), hist("ebitda")
    rev = rev_h + [_mm(d.revenue_ltm)] + p("revenue")
    ebitda = ebitda_h + [_mm(d.ebitda_ltm)] + p("ebitda")
    da = hist("da") + [_mm(d.da_ltm)] + p("da")
    ebit_ltm = (d.ebitda_ltm - d.da_ltm) if (d.ebitda_ltm is not None and d.da_ltm is not None) else None
    ebit = hist("ebit") + [_mm(ebit_ltm)] + p("ebit")
    capex = hist("capex") + [_mm(d.capex_ltm)] + p("capex")
    g_hist = growth(rev_h) if rev_h else []
    revenue_row = s.row
    s.series("Revenue", rev, F_MM, bold=True)
    s.series("growth", g_hist + [""] + (list(proj["growth"]) if n_proj else []), F_PCT, italic=True, indent=True)
    ebitda_row = s.row
    s.series("EBITDA", ebitda, F_MM, bold=True)
    s.series("margin", ratio(ebitda, rev), F_PCT, italic=True, indent=True)
    s.series("Depreciation and amortisation", da, F_MM)
    s.series("EBIT", ebit, F_MM, bold=True)
    s.series("margin", ratio(ebit, rev), F_PCT, italic=True, indent=True)
    s.series("Net income", hist("net_income") + [_mm(d.net_income_ltm)] + p("net_income"), F_MM, bold=True)
    s.series(f"EPS ({qc})", [""] * len(hist_cols) + [r["market"]["eps_ltm"]] + p("eps", 1), F_PX)
    s.series("Capital expenditure", capex, F_MM)
    s.series("% of revenue", ratio(capex, rev), F_PCT, italic=True, indent=True)
    s.series("Free cash flow: levered (OCF - capex)", hist("fcf") + [_mm(d.fcf_ltm)] + [""] * n_proj, F_MM)
    if n_proj:
        s.series("Free cash flow: unlevered (DCF basis)", [""] * (len(hist_cols) + 1) + p("ufcf"), F_MM, bold=True, fill=ICE)
    s.row += 1
    s.text("Projected net income and EPS are illustrative: (EBIT - LTM interest expense) x (1 - tax rate), on a constant "
           "share count. LTM = trailing twelve months and overlaps the latest fiscal year.", L, R, size=9, color=MUTED, italic=True)
    s.row += 1

    if drivers:
        s.band("Projection drivers", L, R)
        s.kv_header(["Driver", "Value", "Basis"], L, R)
        kv = lambda *a, **k: s.kv(*a, basis_to=R, is_input=True, **k)   # noqa: E731
        kv("Year-1 revenue growth", drivers["growth_y1"], F_PCT, drivers["growth_source"])
        kv("Terminal growth", drivers["terminal_growth"], F_PCT, "Growth fades in a straight line from year 1 to this rate in the year after year 5")
        kv("EBITDA margin: LTM", drivers["ebitda_margin_ltm"], F_PCT, "LTM EBITDA / LTM revenue")
        kv("EBITDA margin: year 5", drivers["ebitda_margin_target"], F_PCT, drivers["margin_source"])
        kv("D&A, % of revenue", drivers["da_pct"], F_PCT, drivers["da_source"])
        kv("Capex, % of revenue", drivers["capex_pct"], F_PCT, drivers["capex_source"])
        kv("Net working capital, % of revenue", drivers["nwc_pct"], F_PCT, drivers["nwc_source"] + "; applied to the change in revenue")
        kv("Tax rate", drivers["tax_rate"], F_PCT, r["wacc"]["tax_source"])
        s.row += 1

    n_periods = len(rev)
    if n_periods >= 3:
        chart = BarChart()
        chart.type, chart.grouping = "col", "clustered"
        chart.title = f"Revenue and EBITDA ({rc} m)"
        for row_idx in (revenue_row, ebitda_row):
            chart.add_data(Reference(s.ws, min_col=L, max_col=L + n_periods, min_row=row_idx), from_rows=True, titles_from_data=True)
        chart.set_categories(Reference(s.ws, min_col=L + 1, max_col=L + n_periods, min_row=header_row))
        chart.series[0].graphicalProperties.solidFill = BLUE
        chart.series[1].graphicalProperties.solidFill = GOLD
        chart.y_axis.numFmt = "#,##0"
        chart.y_axis.delete, chart.x_axis.delete = False, False
        chart.legend.position = "b"
        chart.height, chart.width = 8, 22
        s.ws.add_chart(chart, f"B{s.row}")
    s.ws.freeze_panes = s.ws.cell(row=header_row + 1, column=L + 1)


def _news_sheet(wb, d: CompanyData, memo: dict | None, news: list, meta: dict):
    s = Sheet(wb, "News & Risks", [2, 12, 22, 48, 72, 12, 10])
    L, R = 2, 7
    s.title("Recent news, opportunities, threats and risks", f"{d.name} ({d.ticker})", L, R)

    s.band("Recent news headlines and summaries", L, R)
    if news:
        s.header(["Date", "Source", "Headline", "Summary", "Sentiment", "Link"], L, left=(0, 1, 2, 3))
        for n in news:
            sent = n.get("sentiment") or ""
            s.cell(s.row, L, n.get("date"), v="top", line=True)
            s.cell(s.row, L + 1, n.get("publisher"), v="top", wrap=True, line=True)
            s.cell(s.row, L + 2, n.get("title"), v="top", wrap=True, bold=True, line=True)
            s.cell(s.row, L + 3, n.get("summary"), v="top", wrap=True, line=True)
            s.cell(s.row, L + 4, sent or "-", v="top", bold=True, line=True, h="right",
                   color={"Positive": GREEN, "Negative": RED}.get(sent, MUTED))
            url = n.get("url") or ""
            if url.startswith("http"):
                s.cell(s.row, L + 5, "Open", v="top", color=BLUE, link=url, line=True, h="right")
            else:
                s.cell(s.row, L + 5, None, line=True)
            s.fit(s.row, [(n.get("title"), L + 2, L + 2), (n.get("summary"), L + 3, L + 3), (n.get("publisher"), L + 1, L + 1)])
            s.row += 1
        s.text("Headlines, dates, sources and links are taken directly from Yahoo Finance. "
               + ("Summaries and sentiment are written by Gemini from the headline and snippet only."
                  if memo else "Summaries are Yahoo Finance snippets (no Gemini summary generated)."),
               L, R, size=9, color=MUTED, italic=True)
    else:
        s.text("No recent headlines returned by Yahoo Finance.", L, R, color=MUTED, italic=True)
    s.row += 1

    if not memo:
        s.band("Opportunities, threats and risks", L, R)
        s.text(meta["llm_note"], L, R, color=MUTED, italic=True)
        return
    for title, key in (("Key opportunities", "opportunities"), ("Key threats", "threats")):
        s.band(title, L, R)
        for item in memo.get(key, []):
            s.merge(s.row, L, L + 1, item.get("title"), bold=True, wrap=True, v="top", line=True)
            s.merge(s.row, L + 2, R, item.get("detail"), wrap=True, v="top", line=True)
            s.fit(s.row, [(item.get("title"), L, L + 1), (item.get("detail"), L + 2, R)])
            s.row += 1
        s.row += 1
    s.band("Risk register", L, R)
    r0 = s.row
    s.merge(r0, L, L + 1, "Risk", bold=True, size=9, color=NAVY, fill=ICE)
    s.cell(r0, L + 2, "How it hurts", bold=True, size=9, color=NAVY, fill=ICE)
    s.cell(r0, L + 3, "Mitigant / what to watch", bold=True, size=9, color=NAVY, fill=ICE)
    s.merge(r0, L + 4, R, "Severity", bold=True, size=9, color=NAVY, fill=ICE, h="right")
    s.row += 1
    for item in memo.get("risks", []):
        sev = item.get("severity", "")
        s.merge(s.row, L, L + 1, item.get("risk"), bold=True, wrap=True, v="top", line=True)
        s.cell(s.row, L + 2, item.get("detail"), wrap=True, v="top", line=True)
        s.cell(s.row, L + 3, item.get("mitigant"), wrap=True, v="top", line=True)
        s.merge(s.row, L + 4, R, sev, bold=True, v="top", h="right", line=True,
                color={"High": RED, "Medium": AMBER, "Low": GREEN}.get(sev, TEXT))
        s.fit(s.row, [(item.get("risk"), L, L + 1), (item.get("detail"), L + 2, L + 2), (item.get("mitigant"), L + 3, L + 3)])
        s.row += 1
    s.row += 1
    if memo.get("monitor"):
        s.band("What to monitor", L, R)
        for item in memo["monitor"]:
            s.text(item, L, R)


def _assumptions_sheet(wb, d: CompanyData, r: dict, a: Assumptions, meta: dict):
    s = Sheet(wb, "Assumptions", [2, 36, 30, 90])
    w, drv = r["wacc"], r["projections"]["drivers"]
    L, R = 2, 4
    s.title("Assumptions, sources and methodology", "Blue = input or assumption. Black = calculated by the Python engine.", L, R)
    kv = s.kv

    s.band("Run details", L, R)
    kv("Company", d.name, None, f"Ticker {d.ticker}; {d.exchange}")
    kv("Data as of", d.as_of, None, "Market data: Yahoo Finance via the yfinance library (unofficial; may be delayed)")
    kv("Flow figures", d.ltm_basis, None, f"Balance sheet date: {d.balance_sheet_date or 'latest available'}")
    kv("Quote currency", d.quote_currency, None, "Currency of the share price and all per-share values")
    kv("Reporting currency", d.report_currency, None, "Currency of financial statements, market cap and EV in this workbook")
    kv("FX rate", d.fx_trade_to_report, "0.0000", f"1 {d.trade_currency} = x {d.report_currency} (spot)")
    kv("Units", "Millions", None, "Unless stated; per-share values in quote currency")
    kv("Workbook generated", meta["generated"], None, "Numbers are static values from the Python engine; re-run the app to refresh")
    s.row += 1

    s.band("WACC inputs", L, R)
    s.kv_header(["Input", "Value", "Source / basis"], L, R)
    kv("Risk-free rate", w["risk_free"], F_PCT2, w["risk_free_source"]
       + ("" if d.report_currency == "USD" else f". Note: US yield used for a {d.report_currency} reporter; override in the sidebar for a local rate"), is_input=True)
    kv("Equity risk premium", w["erp"], F_PCT2, "Default 5.0% unless changed in the sidebar", is_input=True)
    kv("Beta (adjusted)", w["beta"], F_NUM2, w["beta_source"], is_input=True)
    kv("Pre-tax cost of debt", w["pre_tax_cost_of_debt"], F_PCT2, w["cost_of_debt_source"], is_input=True)
    kv("Tax rate", w["tax_rate"], F_PCT, w["tax_source"], is_input=True)
    s.row += 1

    s.band("DCF and projection inputs", L, R)
    s.kv_header(["Input", "Value", "Source / basis"], L, R)
    kv("Projection period (years)", PROJECTION_YEARS, "0", "Fixed", is_input=True)
    kv("Terminal growth", a.terminal_growth, F_PCT2, "Default 2.5% unless changed in the sidebar", is_input=True)
    kv("Discounting convention", "Mid-year" if a.mid_year else "End of year", None, "Adjustable in the sidebar", is_input=True)
    if drv:
        kv("Year-1 revenue growth", drv["growth_y1"], F_PCT, drv["growth_source"], is_input=True)
        kv("EBITDA margin, year 5", drv["ebitda_margin_target"], F_PCT, drv["margin_source"], is_input=True)
        kv("D&A, % of revenue", drv["da_pct"], F_PCT, drv["da_source"], is_input=True)
        kv("Capex, % of revenue", drv["capex_pct"], F_PCT, drv["capex_source"], is_input=True)
        kv("Net working capital, % of revenue", drv["nwc_pct"], F_PCT, drv["nwc_source"], is_input=True)
    s.row += 1

    s.band("LBO inputs", L, R)
    s.kv_header(["Input", "Value", "Source / basis"], L, R)
    kv("Senior debt / EBITDA", a.lbo_leverage, F_X, "Specified input", is_input=True)
    kv("Interest rate", a.lbo_rate, F_PCT2, "Specified input", is_input=True)
    kv("Exit horizon (years)", PROJECTION_YEARS, "0", "Specified input", is_input=True)
    kv("Exit EV/EBITDA", r["lbo"].get("exit_multiple"), F_X,
       "Base (current) EV/EBITDA" if a.lbo_exit_multiple is None else "User override", is_input=True)
    kv("Minimum target IRR", a.lbo_target_irr, F_PCT, "Specified input", is_input=True)
    kv("Maximum debt / entry EV", a.lbo_max_ltv, F_PCT, "Sanity cap; only binds for companies trading on low EBITDA multiples", is_input=True)
    s.row += 1

    s.band("Peers and qualitative analysis", L, R)
    kv("Peer selection", "", None, meta["peer_source"])
    kv("Qualitative text", "", None, meta["llm_status"])
    s.row += 1

    s.band("Methodology", L, R)
    for line in (
        "Division of labour: all financial maths is computed by the Python engine. Gemini only suggests peer tickers and writes "
        "qualitative text from the Python-computed metrics and the top headlines. It is never given raw filings and never does arithmetic.",
        "Enterprise value = market cap + total debt - cash and short-term investments + minority interest + preferred equity.",
        "WACC: cost of equity from CAPM; cost of debt from LTM interest expense / total debt, bounded to risk-free +1% to +8%; "
        "weights use market cap and book debt.",
        "Projections: revenue starts from LTM and grows at the year-1 rate, fading in a straight line toward terminal growth. "
        "EBITDA margin, D&A, capex and working-capital ratios as listed above.",
        "DCF: unlevered FCF = EBIT x (1 - tax) + D&A - capex - increase in NWC, discounted at WACC; Gordon-growth terminal value; "
        "equity value = EV - net debt - minority interest - preferred. Stock-based compensation is treated as a cash cost.",
        "LBO: purchase at current EV, senior debt as specified, cash interest on opening balances, 100% cash sweep, no fees, "
        "exit at the base multiple in year 5. Max price = highest entry EV that still earns the target IRR.",
        "Peers: multiples computed from Yahoo Finance data for each peer; implied values use the peer median and 25th/75th percentiles.",
    ):
        s.text(line, L, R, size=9)
    s.row += 1
    if r["notes"]:
        s.band("Data and model notes", L, R)
        for note in r["notes"]:
            s.text(note, L, R, size=9)
        s.row += 1
    s.band("Disclaimer", L, R)
    s.text(DISCLAIMER, L, R, size=9, color=MUTED, italic=True)


# =============================================================================== entry point
def build_workbook(d: CompanyData, r: dict, a: Assumptions, memo: dict | None, news: list, meta: dict) -> bytes:
    """Return the finished .xlsx as bytes (ready for st.download_button)."""
    meta = dict(meta)
    meta.setdefault("generated", datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC"))
    meta.setdefault("target_irr", a.lbo_target_irr)
    meta.setdefault("peer_source", "Not specified")
    meta.setdefault("llm_status", "Not generated")
    meta.setdefault("llm_note", "Qualitative analysis was not generated for this run.")
    meta.setdefault("footer", "Generated by Equity Memo Agent. Market data: Yahoo Finance. Financial maths: Python engine.")

    wb = Workbook()
    wb.remove(wb.active)
    _memo_sheet(wb, d, r, memo, news, meta)
    _valuation_sheet(wb, d, r, meta)
    _dcf_sheet(wb, d, r)
    _lbo_sheet(wb, d, r)
    _peers_sheet(wb, d, r, meta)
    _projections_sheet(wb, d, r)
    _news_sheet(wb, d, memo, news, meta)
    _assumptions_sheet(wb, d, r, a, meta)
    wb.properties.title = f"{d.name} ({d.ticker}) equity valuation memo"
    wb.properties.creator = "Equity Memo Agent"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def filename(d: CompanyData) -> str:
    safe = re.sub(r"[^A-Za-z0-9]+", "-", d.ticker).strip("-")
    return f"{safe}_Equity_Memo_{datetime.now(timezone.utc):%Y-%m-%d}.xlsx"
