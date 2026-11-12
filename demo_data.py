"""Synthetic demo company so the app can be tried with no API keys and no network.

Everything in this file is invented. "Demo Industrials" is not a real company and
none of these figures, peers or headlines describe a real business.
"""
from __future__ import annotations

import pandas as pd

from market_data import CompanyData

M = 1e6


def demo_company() -> CompanyData:
    hist = pd.DataFrame(
        {
            "FY2022": [8100, 1580, 1230, 760, 350, 400, 1290, 890, 235, 1010, 175],
            "FY2023": [8650, 1730, 1360, 850, 370, 420, 1420, 1000, 262, 1125, 178],
            "FY2024": [9150, 1880, 1485, 930, 395, 440, 1530, 1090, 288, 1230, 180],
            "FY2025": [9600, 2010, 1600, 990, 410, 465, 1600, 1135, 305, 1310, 182],
        },
        index=["revenue", "ebitda", "ebit", "net_income", "da", "capex", "ocf", "fcf", "tax", "pretax", "interest"],
        dtype=float,
    ) * M
    news = [
        ("Demo Industrials lifts full-year margin guidance after automation orders beat", "Synthetic Newswire",
         "2026-11-05", "Management raised its EBITDA margin outlook, citing a stronger order book in factory automation."),
        ("Demo Industrials to buy sensor maker in bolt-on deal", "Synthetic Newswire", "2026-10-28",
         "The company agreed to acquire a privately held industrial sensor business, funded from cash."),
        ("Union talks at Demo Industrials' largest plant enter second month", "Synthetic Newswire", "2026-10-20",
         "Negotiations over a new three-year wage agreement continue with no stoppage announced."),
        ("Demo Industrials expands buyback programme", "Synthetic Newswire", "2026-10-09",
         "The board authorised an additional repurchase tranche to run through next year."),
        ("Input costs ease for industrial manufacturers, survey shows", "Synthetic Newswire", "2026-09-30",
         "A sector survey pointed to lower steel and energy costs across the industry in the third quarter."),
    ]
    return CompanyData(
        ticker="DEMO", name="Demo Industrials Inc. (synthetic)", sector="Industrials",
        industry="Specialty Industrial Machinery", industry_key="specialty-industrial-machinery",
        country="United States", exchange="Synthetic demo data",
        summary="Demo Industrials is an invented company used to demonstrate this tool. It is modelled as a "
                "diversified maker of factory automation equipment, industrial sensors and aftermarket services.",
        employees=24000,
        quote_currency="USD", trade_currency="USD", report_currency="USD", quote_factor=1.0, fx_trade_to_report=1.0,
        price=84.50, price_major=84.50, market_cap=84.50 * 250 * M, shares=250 * M,
        revenue_ltm=9800 * M, ebitda_ltm=2060 * M, net_income_ltm=1010 * M, da_ltm=420 * M, capex_ltm=470 * M,
        ocf_ltm=1620 * M, fcf_ltm=1150 * M, interest_expense_ltm=180 * M,
        total_debt=4200 * M, cash=1150 * M, minority_interest=60 * M, preferred=0.0, nwc=1180 * M,
        ltm_basis="Trailing twelve months", balance_sheet_date="2026-09-30",
        beta=1.08, beta_source="Synthetic demo value", tax_rate=0.233,
        tax_rate_source="Effective rate: tax provision / pretax income, avg of last 3 fiscal years",
        pe_ttm_yahoo=20.9, pe_fwd_yahoo=18.6, eps_fwd=4.54, consensus_growth={"0y": 0.060, "+1y": 0.055},
        target_mean=96.0, target_low=78.0, target_high=112.0, analyst_count=14, recommendation="Buy",
        week52_low=66.20, week52_high=91.80, revenue_growth_yoy=0.064, hist=hist,
        news=[{"title": t, "publisher": p, "date": dt, "snippet": s, "url": ""} for t, p, dt, s in news],
        notes=["SYNTHETIC DEMO DATA: this company, its peers and its headlines are invented."],
        as_of="Synthetic demo data",
    )


def demo_peers() -> list[dict]:
    rows = [
        ("PEERA", "Peer A Automation (synthetic)", 31000, 5200, 1900, 13400, 2750, 1290, 0.071),
        ("PEERB", "Peer B Controls (synthetic)", 12400, 3100, 700, 6900, 1420, 690, 0.038),
        ("PEERC", "Peer C Machinery (synthetic)", 18800, 2600, 2100, 8200, 1720, 870, 0.052),
        ("PEERD", "Peer D Sensors (synthetic)", 9200, 1500, 450, 3600, 760, 350, 0.094),
        ("PEERE", "Peer E Industrial (synthetic)", 26500, 7400, 1600, 15800, 3350, 1560, 0.021),
    ]
    return [{"ticker": t, "name": n, "currency": "USD", "market_cap": mc * M, "total_debt": debt * M,
             "cash": cash * M, "revenue": rev * M, "ebitda": ebitda * M, "net_income": ni * M,
             "pe_yahoo": None, "fwd_pe": mc / ni * 0.9, "revenue_growth": g, "same_currency": True}
            for t, n, mc, debt, cash, rev, ebitda, ni, g in rows]


def demo_memo() -> dict:
    """Canned qualitative text in the same shape Gemini returns."""
    return {
        "business_summary": "Synthetic example. Demo Industrials is an invented diversified industrial that sells "
                            "factory automation equipment and sensors, with a recurring aftermarket service base.",
        "investment_thesis": "Synthetic example text. The illustrative case rests on mid-single-digit growth, stable "
                             "margins near 21% and steady cash conversion. With the shares close to the DCF value and "
                             "in line with peers, returns depend on execution rather than re-rating.",
        "variant_perception": "Synthetic example text. Consensus assumes margins hold; the model shows value is "
                              "sensitive to that assumption, so evidence of sustained pricing power would be the "
                              "differentiating call.",
        "valuation_view": "Synthetic example text. The DCF, peer multiples and analyst target bracket the current "
                          "price, while the LBO return falls short of the 20% hurdle at today's valuation.",
        "stance": "Neutral",
        "news": [
            {"headline_id": 0, "summary": "Synthetic example: guidance raised on stronger automation orders, "
                                          "supportive of the margin assumption.", "sentiment": "Positive"},
            {"headline_id": 1, "summary": "Synthetic example: small cash-funded acquisition adds sensor capability; "
                                          "integration is the watch point.", "sentiment": "Neutral"},
            {"headline_id": 2, "summary": "Synthetic example: prolonged wage talks raise the risk of cost inflation "
                                          "or disruption.", "sentiment": "Negative"},
        ],
        "opportunities": [
            {"title": "Automation demand", "detail": "Synthetic example: order growth in factory automation could lift revenue growth above the base case."},
            {"title": "Aftermarket mix", "detail": "Synthetic example: a larger service base would raise margins and cash conversion."},
            {"title": "Bolt-on acquisitions", "detail": "Synthetic example: balance sheet capacity allows small deals in sensors."},
        ],
        "threats": [
            {"title": "Industrial cycle", "detail": "Synthetic example: a capex downturn would hit orders and operating leverage."},
            {"title": "Labour costs", "detail": "Synthetic example: an expensive wage settlement would compress margins."},
            {"title": "Pricing pressure", "detail": "Synthetic example: larger peers could compete on price in core lines."},
        ],
        "risks": [
            {"risk": "Margin compression", "severity": "High", "detail": "Synthetic example: each point of margin moves value materially.",
             "mitigant": "Synthetic example: service revenue is less cyclical."},
            {"risk": "Acquisition integration", "severity": "Medium", "detail": "Synthetic example: poor integration would dilute returns.",
             "mitigant": "Synthetic example: deals are small relative to the group."},
            {"risk": "Rate sensitivity", "severity": "Low", "detail": "Synthetic example: higher rates raise the discount rate.",
             "mitigant": "Synthetic example: leverage is moderate."},
        ],
        "monitor": ["Synthetic example: order intake", "Synthetic example: wage settlement", "Synthetic example: margin guidance"],
    }
