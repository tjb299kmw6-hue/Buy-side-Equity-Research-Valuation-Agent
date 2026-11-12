# Equity Memo Agent

**Try it here: https://buyside-equity-memo-app.streamlit.app** (free, nothing to install, no sign-up)

I built this to speed up the first pass on a stock. You type in a company name or ticker, and it pulls the market data, runs a standard set of valuation models and gives you a formatted Excel memo you can download and work from.

It works for any company listed on Yahoo Finance, including non-US listings like `SHEL.L`, `SAP.DE` or `7203.T`.

All the numbers (market cap, EV, P/E, EV/EBITDA, FCF yield, WACC, a 5-year DCF, an LBO returns test and peer multiples) are calculated in Python. Recent news headlines come straight from Yahoo Finance. There's an optional Gemini add-on that writes a short thesis and risk section, but it's off by default and it never touches the numbers.

> This is for learning and research only, not investment advice. Yahoo data can be delayed or wrong, so check anything important against the company's filings.

## What's in the memo

The Excel file has eight tabs:

| Tab | What it shows |
|---|---|
| Memo | One-page summary: key metrics, valuation range, recent news and model notes |
| Valuation | Market snapshot, EV bridge, WACC build and a football-field chart |
| DCF | 5-year unlevered free cash flow, terminal value, exit-multiple cross-check and a WACC vs growth sensitivity |
| LBO | Sources and uses, debt paydown, sponsor IRR and MOIC, the max price that still hits the target IRR, and two sensitivity tables |
| Peers | Trading multiples for comparable companies and the value they imply |
| Projections | Historical financials, LTM and the 5-year forecast |
| News & Risks | Headlines with links (plus AI commentary if you turn it on) |
| Assumptions | Every input, where it came from, and how each model works |

The LBO defaults to 4.5x senior debt / EBITDA at 7% interest, a 5-year hold, exit at today's EV/EBITDA multiple and a 20% target IRR. You can change all of these, and the DCF assumptions, from the sidebar.

## How to use it

Open the link above, type a company, hit **Run analysis**, then **Download Excel memo**. Hit **Reset** to do another one. If you want to see what the output looks like first, **Try the synthetic demo** runs on a made-up company.

If you'd like the AI commentary, open *Optional: AI commentary* in the sidebar, tick the box and paste in your own free Gemini key from <https://aistudio.google.com/apikey>.

## How the numbers are worked out

- **Enterprise value**: market cap + debt − cash + minority interest + preferred equity.
- **WACC**: cost of equity from CAPM (US 10-year yield, 5% equity risk premium, Blume-adjusted beta). For non-US companies the beta comes from a 5-year regression against the local index, because Yahoo's own beta for these is often unreliable. Cost of debt is interest expense / total debt, kept within 1% to 8% above the risk-free rate.
- **Projections**: revenue starts at analyst consensus growth (or the historical growth rate if there's no consensus) and fades towards terminal growth. EBITDA margin is held at today's level, and D&A, capex and working capital follow their recent share of revenue.
- **DCF**: unlevered free cash flow discounted at WACC, Gordon-growth terminal value (2.5% by default), mid-year discounting, and an exit-multiple version as a sense check.
- **LBO**: buy at today's EV, all spare cash goes to paying down debt, exit at the starting multiple. The "max price" is the highest price a sponsor could pay and still hit the target IRR.
- **Peers**: either the tickers you type in, or picked automatically from companies of a similar size (0.1x to 10x market cap) in the same industry, topped up from the same sector if needed. Typing your own peers will usually give a tighter comparison.

## Limitations

- Banks and insurers don't report EBITDA, so the DCF and LBO don't apply to them and are marked as such.
- The risk-free rate is the US 10-year for every company. For a non-US company you should put the local government bond yield in the sidebar.
- Commodity producers (oil, mining) often come out looking very cheap on a DCF, because today's cash flows get projected forward without the commodity cycle. The memo flags this.
- The equity bridge doesn't adjust for pension deficits or long-term investments, which matters for a handful of companies.
- Yahoo Finance is an unofficial source and can occasionally throttle requests. If you get an error, wait a minute and try again.
- It uses the same assumptions for every company. It's a starting point for analysis, not a replacement for a proper company-specific model.

## Running it yourself

You need Python 3.10 or newer.

```bash
git clone https://github.com/tjb299kmw6-hue/Buy-side-Equity-Research-Valuation-Agent.git
cd Buy-side-Equity-Research-Valuation-Agent
pip install -r requirements.txt
streamlit run app.py
```

No Git? Click **Code > Download ZIP** on this page, unzip it, open a terminal in the folder and run the last two lines. On Windows, use `py -m pip install -r requirements.txt` and `py -m streamlit run app.py`.

For the optional Gemini commentary when running locally, copy `.streamlit/secrets.toml.example` to `.streamlit/secrets.toml` and add your key (that file is ignored by Git, so it won't be uploaded). The free tier is limited to roughly 20 requests a day per model, and each company uses two. If one model runs out, the app falls back to `gemini-3.8-flash` and then `gemini-3.6-flash`. If all three are busy, you still get the full numbers.

## Code layout

```
app.py            the Streamlit app
market_data.py    pulls prices, financials, news and FX from Yahoo Finance
engine.py         all the valuation maths (WACC, DCF, LBO, peers)
excel_memo.py     builds the Excel workbook
gemini_layer.py   the optional Gemini commentary
demo_data.py      the made-up company for the demo
tests/            offline tests
```

To run the tests: `pip install pytest`, then `python -m pytest -q`. They check the maths against hand calculations and run the whole app end to end without needing internet or an API key.
