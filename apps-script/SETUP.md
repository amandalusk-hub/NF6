# MNW4 Net Worth Sheet — Apps Script Setup

## Files to add to your Apps Script project

| File | Purpose |
|------|---------|
| `PropertyValues.gs` | Zillow Zestimate auto-pull for US properties |
| `CurrencyConversion.gs` | Foreign currency → USD conversion for international assets |

---

## Step 1: Add the scripts

1. Open your Google Sheet
2. Go to **Extensions > Apps Script**
3. For each `.gs` file above:
   - Click **+** next to "Files"
   - Choose **Script**
   - Name it (e.g. `PropertyValues`, `CurrencyConversion`)
   - Paste the contents
4. Click **Save** (Ctrl+S)

---

## Step 2: Set up Zillow (US Properties)

1. Go to [RapidAPI](https://rapidapi.com) → sign up free
2. Search for **"Zillow Com"** API → subscribe (free tier: 20 req/month)
3. Copy your **RapidAPI Key**
4. In Apps Script: **Project Settings** (gear icon) → **Script Properties**
5. Add: `RAPIDAPI_KEY` = `<your key>`

**In your sheet**, use the formula:
```
=ZESTIMATE("123 Main St, Austin, TX 78701")
```

To refresh all US properties at once, run `refreshAllZestimates()` from the editor,
or install the daily trigger by running `installZilowTrigger()` once.

---

## Step 3: Set up Foreign Currency (International Properties)

No API key required — uses a free exchange rate API automatically.

**In your sheet**, use these formulas:
```
=FX_RATE("COP")           → today's Colombian Peso to USD rate
=TO_USD(500000000, "COP") → converts 500M COP to USD
=TO_USD(250000, "EUR")    → converts 250K EUR to USD
```

To set up currency dropdowns in column C, run `setupCurrencyDropdowns()` once.

To install a daily auto-refresh trigger, run `installFxTrigger()` once.

---

## Supported Currencies (sample)

| Code | Currency |
|------|----------|
| COP | Colombian Peso |
| EUR | Euro |
| GBP | British Pound |
| MXN | Mexican Peso |
| CAD | Canadian Dollar |
| AUD | Australian Dollar |
| JPY | Japanese Yen |
| CHF | Swiss Franc |
| BRL | Brazilian Real |
| AED | UAE Dirham |

Full list of 30+ currencies in `CurrencyConversion.gs`.

---

## Recommended Sheet Layout for International Properties

| Column A | Column B | Column C | Column D | Column E | Column F |
|----------|----------|----------|----------|----------|----------|
| Asset Name | Address/Description | Currency | Local Value | USD Value | Last Updated |
| Casa Bianca | Cartagena, Colombia | COP | 2,000,000,000 | =TO_USD(D2,C2) | auto |

---

## Optional: Higher-volume FX API

For frequent refreshes, get a free key at [exchangerate-api.com](https://www.exchangerate-api.com)
and add it to Script Properties as `EXCHANGERATE_API_KEY`.
