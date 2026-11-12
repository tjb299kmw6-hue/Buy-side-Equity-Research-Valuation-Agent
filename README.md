# Equity Memo Agent

Type a company name or ticker and get a structured, buy-side style valuation memo as a formatted Excel workbook. It works for any listed company on Yahoo Finance, including non-US listings (`SHEL.L`, `SAP.DE`, `7203.T`).

- **Python does all the maths**: share price, market cap, EV, P/E, EV/EBITDA, FCF yield, WACC, 5-year DCF, LBO returns and peer multiples.
- **Real headlines, no AI needed**: recent Yahoo Finance headlines with snippets and links, filtered for relevance.
- **Optional AI commentary (off by default)**: tick it in the sidebar and Google Gemini adds a thesis, opportunities, threats, risk register and news summaries. It is only given Python-computed metrics and the top headlines, and it never does arithmetic.
- **Free to run**: Yahoo Finance data needs no key; the optional Gemini commentary has a free tier.

> For information and education only. Not investment advice. Yahoo Finance data can be delayed or wrong; check company filings before relying on any number.

## What you get

An `.xlsx` memo with eight tabs:

| Tab | Contents |
|---|---|
| Memo | One-page summary: key metrics, valuation summary, stance, thesis, variant perception, opportunities and threats, recent news, model notes |
| Valuation | Market snapshot, EV bridge, WACC build, valuation range chart |
| DCF | 5-year unlevered free cash flow, valuation bridge, exit-multiple cross-check, WACC x terminal growth sensitivity |
| LBO | Sources and uses, debt schedule with cash sweep, sponsor IRR and MOIC, maximum price for the target IRR, two IRR sensitivities |
| Peers | Trading multiples for the company and its peers, peer statistics, peer-implied value per share |
| Projections | Historical financials, LTM, 5-year projections, drivers, chart |
| News & Risks | Headlines with summaries and links, opportunities, threats, risk register, what to monitor |
| Assumptions | Every input with its source, the methodology, and data notes |

Default LBO case: 4.5x senior debt / EBITDA at 7.0% interest, 5-year exit at the base (current) EV/EBITDA multiple, 20% minimum target IRR. All adjustable in the sidebar.

## Use it online

No install needed: open the live app (link to be added once deployed on Streamlit Community Cloud), type a company name or ticker, and download the Excel memo. For optional AI commentary, open *Optional: AI commentary* in the sidebar, tick the box and paste your own free Gemini key.

## Run it on your own computer

Requires Python 3.10 or newer (3.12 or 3.13 recommended).

```bash
git clone https://github.com/tjb299kmw6-hue/Buy-side-Equity-Research-Valuation-Agent.git
cd Buy-side-Equity-Research-Valuation-Agent
pip install -r requirements.txt
streamlit run app.py
```

No Git? On the GitHub page click **Code > Download ZIP**, unzip it, open a terminal in that folder, and run the last two commands. On Windows, use `py -m pip ...` and `py -m streamlit run app.py` if `pip` is not recognised.

Click **Try the synthetic demo** to see the full output with no keys and no network calls (the demo company is invented).

### Add your free Gemini key

1. Create a key at <https://aistudio.google.com/apikey>.
2. Give it to the app in any one of three ways (checked in this order):
   - tick **Add AI-written thesis...** under *Optional: AI commentary* and paste it into **Gemini API key** (your session only), or
   - copy `.streamlit/secrets.toml.example` to `.streamlit/secrets.toml` and set `GEMINI_API_KEY`, or
   - set an environment variable: `export GEMINI_API_KEY=...`

AI commentary is off by default. Without it the app still produces the full quantitative workbook with real Yahoo headlines; only the thesis, opportunities and threats sections are left out, and peers are chosen by size (below).

## Deploy free on Streamlit Community Cloud

1. Push this folder to a public GitHub repository. `.streamlit/secrets.toml` is git-ignored, so your key is not committed.
2. Go to <https://share.streamlit.io>, choose **Create app**, pick the repository and set the main file to `app.py`.
3. Open **Advanced settings > Secrets** and paste:
   ```toml
   GEMINI_API_KEY = "your-key"
   ```
4. Deploy and share the link.

### Free-tier quota: read this before sharing a link

Gemini's free tier is small. At the time of writing the Flash models were reported at about 5 requests per minute and 20 requests per day per key (check the rate-limit page in AI Studio for your own limits). Each new company uses two requests, so a single shared key covers roughly ten companies a day. The app is built around that:

- Responses are cached for an hour (`@st.cache_data(ttl=3600)`), so repeat runs of the same company cost nothing.
- Entering peers manually skips one of the two requests.
- On a `429` the app shows a retry message and still delivers the quantitative workbook.
- Every visitor can paste their own free key in the sidebar, which is the practical way to share this widely.
- **Automatic fallback:** if a model's quota runs out (HTTP 429) the app tries the next one, in this order: `gemini-3.5-flash-lite` (primary), `gemini-3.8-flash`, then `gemini-3.6-flash`. Quotas are counted per model, so each fallback has its own allowance. The Assumptions tab and the memo footer record which model actually wrote the text. Only quota errors (and an unavailable model) trigger a fallback; a bad API key stops straight away.
- To change the primary model, set `GEMINI_MODEL` in secrets or the environment. The two fallbacks stay in place.

## How it works

```
app.py            Streamlit UI, caching, orchestration
market_data.py    Yahoo Finance via yfinance: prices, statements, news, FX, peers
engine.py         Pure Python maths: market metrics, WACC, projections, DCF, LBO, peers
gemini_layer.py   google-genai SDK, Pydantic schemas, key handling, 429 handling
excel_memo.py     openpyxl workbook writer
demo_data.py      Synthetic company for the offline demo
tests/            Offline tests (no network, no key)
```

Gemini calls use the official `google-genai` SDK with `GenerateContentConfig(response_mime_type="application/json", response_schema=<Pydantic model>)`, so the reply is validated JSON that maps straight into Excel cells.

### Method in brief

- **Enterprise value** = market cap + total debt - cash and short-term investments + minority interest + preferred.
- **WACC**: CAPM cost of equity (live US 10-year yield, 5.0% equity risk premium, Blume-adjusted beta: Yahoo beta for US listings, a 5-year regression against the local index for non-US listings, where Yahoo's own beta is unreliable); cost of debt from interest expense / total debt, bounded to risk-free +1% to +8%; market-value equity and book debt weights.
- **Projections**: revenue grows from LTM at the consensus rate (or the historical CAGR if there is no consensus), fading in a straight line toward terminal growth. EBITDA margin held at LTM; D&A and capex at their 3-year average share of revenue; working capital at its current share of revenue.
- **DCF**: unlevered FCF discounted at WACC with a Gordon-growth terminal value (2.5% default) and mid-year convention; cross-checked with an exit-multiple terminal value.
- **LBO**: purchase at the current EV, senior debt as specified, cash interest on opening balances, 100% cash sweep, no fees, exit at the base multiple. The "max price" is the highest entry price that still earns the target IRR.
- **Peers**: typed in, or chosen automatically with no AI: companies of similar market cap (0.1x to 10x) from the company's Yahoo industry, topped up with large caps from the same sector (Gemini suggests them instead if AI commentary is on). All are validated against Yahoo and the multiples are computed in Python. For a tighter comparison, enter your own peers.

### Known limits

- Banks and insurers do not report EBITDA, so the DCF and LBO are marked not applicable for them.
- The risk-free rate is the US 10-year yield for every company. For non-US companies, enter the local government bond yield in the sidebar.
- Yahoo Finance is an unofficial data source and can throttle shared cloud servers. If that happens, wait a minute and retry.
- For companies that report in a different currency from their listing (ADRs, some LSE names), market cap is converted at spot FX.
- The equity bridge deducts net debt, minority interest and preferred only. Long-term investments, pension deficits and other non-operating items are not adjusted, which matters for a few cash-rich or pension-heavy companies.
- This is a systematic first-pass screen with uniform assumptions. It is a starting point for analysis, not a substitute for a company-specific model.

## Tests

```bash
pip install pytest
python -m pytest -q
```

The tests check the financial maths against hand calculations, the Yahoo parsing against a fake `yfinance`, the Gemini layer against a fake client (including the 429 path), and the whole app end to end.
