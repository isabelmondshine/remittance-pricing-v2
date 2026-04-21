#  Remittance Competitive Pricing Tracker
## Setup & Usage Guide

---

## What This Does

Pulls **live, bank-to-bank transfer pricing** from each provider every time it runs,
and writes a formatted Excel file with:
- Fee (USD), FX Rate, FX Markup %, and Amount Received
- For send amounts: $100, $500, $1,000, $5,000, $10,000
- Across corridors: USD→EUR, GBP, MXN, INR, PHP, CAD, AUD, JPY
- Providers: Wise, Remitly, Western Union, MoneyGram, Revolut, Euronet (+ Xoom, Ria, WorldRemit via Wise API)

---

## How Each Provider Is Fetched

| Provider       | Method                          | Auth Required | Notes |
|----------------|---------------------------------|---------------|-------|
| **Wise**       | Public Comparison API           | None          | `api.wise.com/v4/comparisons` — also returns Remitly, WU, MGram data |
| **Remitly**    | Public price calculator API     | None          | `api.remitly.io/v3/pricing` — Economy rate, bank deposit |
| **Western Union** | Public price estimator API   | Session cookie | May need first request via browser |
| **MoneyGram**  | Public fee estimator API        | None          | `moneygram.com/mgo/us/en/fee-estimator/api` |
| **Revolut**    | Headless browser (Playwright)   | None          | JS-rendered page — no public API |
| **Euronet**    | Headless browser (Playwright)   | None          | JS-rendered page — coverage varies |

> **Note on Wise's Comparison API:** Wise maintains this as a public endpoint that powers
> their own comparison page at wise.com/compare. It scrapes competitor pricing ~hourly
> and covers bank transfer in/out only. This is the single most reliable source for
> cross-provider data and often makes direct calls redundant for the providers it covers.

---

## Installation

```bash
# 1. Install Python dependencies
pip install requests openpyxl playwright

# 2. Install Playwright browser (for Revolut + Euronet scraping)
playwright install chromium

# 3. Run it
python remittance_pricer.py
```

---

## Usage Examples

```bash
# Run all corridors and amounts (default)
python remittance_pricer.py

# Single corridor
python remittance_pricer.py --corridors USD-EUR

# Multiple specific corridors
python remittance_pricer.py --corridors USD-EUR USD-INR USD-MXN

# Specific amounts only
python remittance_pricer.py --amounts 500 1000 5000

# Skip browser scraping (faster — skips Revolut & Euronet)
python remittance_pricer.py --no-browser

# Custom output filename
python remittance_pricer.py --output pricing_daily.xlsx
```

---

## Running Daily (Automated)

### macOS / Linux — cron
```bash
# Open crontab
crontab -e

# Run every weekday at 8:00 AM, save to a dated file
0 8 * * 1-5 cd /path/to/script && python remittance_pricer.py >> logs/pricer.log 2>&1
```

### Windows — Task Scheduler
1. Open Task Scheduler → Create Basic Task
2. Set trigger: Daily, 8:00 AM
3. Action: Start a Program
   - Program: `C:\Python311\python.exe`
   - Arguments: `C:\path\to\remittance_pricer.py`
   - Start in: `C:\path\to\script\`

### GitHub Actions (cloud, free)
Create `.github/workflows/daily_pricing.yml`:
```yaml
name: Daily Remittance Pricing
on:
  schedule:
    - cron: '0 12 * * 1-5'  # 8 AM ET on weekdays
  workflow_dispatch:          # also allows manual trigger

jobs:
  fetch-pricing:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v3
      - uses: actions/setup-python@v4
        with:
          python-version: '3.11'
      - run: pip install requests openpyxl playwright
      - run: playwright install chromium
      - run: python remittance_pricer.py --output pricing_$(date +%Y%m%d).xlsx
      - uses: actions/upload-artifact@v3
        with:
          name: pricing-report
          path: pricing_*.xlsx
```

---

## Troubleshooting

### "HTTP 403" or "HTTP 401" from Remitly/Western Union/MoneyGram
Some providers add anti-bot headers on their calculator APIs. If this happens:
1. Open that provider's website in Chrome DevTools → Network tab
2. Find the JSON request their calculator makes
3. Copy the exact URL and any required headers into the script's provider function

### Revolut/Euronet not loading
- Make sure you ran `playwright install chromium`
- Try `--no-browser` to skip those providers
- Check if their website structure changed (update the CSS selectors in the browser functions)

### Wise API returns no competitors
- Wise's comparison API coverage varies by corridor — some corridors have fewer competitors
- Try a different corridor to verify the API is working

---

## Output Excel Structure

| Sheet | Contents |
|-------|----------|
| **Pricing Summary** | Main comparison table: all providers × amounts × corridors |
| **USD→EUR**, **USD→INR**, etc. | Raw data per corridor with auto-filter |
| **Legend & Notes** | Color coding guide and data source notes |

### Color Coding
- 🟢 **Green FX markup**: < 0.5% above mid-market (excellent)
- 🟡 **Amber FX markup**: 0.5–1.5% above mid-market (moderate)
- 🔴 **Red FX markup**: > 1.5% above mid-market (expensive)
- 🟢 **Green background**: Best received amount / lowest fee for that amount

---

## Extending the Script

To add a new provider, create a function following this pattern:

```python
def get_newprovider_quote(from_ccy: str, to_ccy: str, amount: float) -> Quote:
    try:
        resp = requests.get("https://newprovider.com/api/price", params={
            "from": from_ccy, "to": to_ccy, "amount": amount
        }, headers=HEADERS, timeout=20)
        resp.raise_for_status()
        data = resp.json()
        return Quote(
            provider="New Provider",
            from_ccy=from_ccy, to_ccy=to_ccy, send_amount=amount,
            fee_usd=data["fee"],
            fx_rate=data["rate"],
            fx_markup_pct=data.get("markup"),
            received_amount=data["receivedAmount"],
            note="Bank transfer",
        )
    except Exception as e:
        return Quote("New Provider", from_ccy, to_ccy, amount, error=str(e))
```

Then add it to the `fetch_all_quotes()` orchestrator function.
