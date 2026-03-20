# Family Office Wealth Tracker — Apps Script Setup

## Files in this project

| File | Purpose |
|------|---------|
| `Code.gs` | Core backend: assets CRUD, FX rates, Plaid API, property refresh, web app entry |
| `PropertyValues.gs` | US property auto-valuation via Rentcast + interactive lookup |
| `CurrencyConversion.gs` | Foreign currency → USD helpers (`TO_USD`, `FX_RATE` formulas) |
| `Plaid.gs` | Plaid Link sidebar UI (credential setup, open sidebar, remove connection) |
| `PlaidLink.html` | Plaid Link embedded iframe |

---

## Step 1: Add files to your Apps Script project

1. Open your Google Sheet
2. Go to **Extensions > Apps Script**
3. For each `.gs` file above, click **+** next to "Files" → Script → paste contents
4. Click **Save** (Ctrl+S)

---

## Step 2: Set Script Properties

Go to **Project Settings** (gear icon) → **Script Properties**, then add:

| Property | Value | Required? |
|----------|-------|-----------|
| `RENTCAST_API_KEY` | Your key from rentcast.io | For US property values |
| `PLAID_CLIENT_ID` | Your Plaid client ID | For bank sync |
| `PLAID_SECRET` | Your Plaid secret | For bank sync |
| `PLAID_ENV` | `sandbox` or `production` | For bank sync |
| `EXCHANGERATE_API_KEY` | Key from exchangerate-api.com | Optional (higher FX volume) |

---

## Step 3: US Property Values (Rentcast)

Zillow deprecated their public API. This tracker uses **Rentcast** instead.

1. Sign up at [rentcast.io](https://rentcast.io) — free tier: **50 requests/month**
2. Copy your API key → add as `RENTCAST_API_KEY` in Script Properties

**How to tag a property for auto-valuation:**

In the asset's **Notes** field, include the address like this:
```
address: 709 Kuhlman Road, Houston, TX 77024
```

Then run **Tracker > Refresh US Property Values** (or `refreshUSPropertyValues()`) to update all tagged Real Estate assets.

For a one-off lookup, run **Tracker > Lookup Single Property** (or `lookupSingleProperty()`).

---

## Step 4: Foreign Currency (International Properties)

No separate setup required — `CurrencyConversion.gs` uses [open.er-api.com](https://open.er-api.com) automatically.

Use these formulas in your sheet:
```
=FX_RATE("COP")              → today's COP → USD rate
=TO_USD(2000000000, "COP")   → converts 2B COP to USD
=TO_USD(250000, "EUR")       → converts 250K EUR to USD
```

For higher request volume, add a free key from [exchangerate-api.com](https://www.exchangerate-api.com) as `EXCHANGERATE_API_KEY`.

---

## Step 5: Bank Sync (Plaid)

1. Set `PLAID_CLIENT_ID`, `PLAID_SECRET`, `PLAID_ENV` in Script Properties
2. Run **Tracker > Connect Bank Account** to open the Plaid Link sidebar
3. Connect your bank — accounts are automatically added to the Assets sheet
4. Use **Tracker > Sync Plaid Accounts** to refresh balances on demand
   (also runs automatically in the daily 7 AM sync trigger)

---

## Supported Currencies

COP · EUR · GBP · MXN · CAD · AUD · JPY · CHF · BRL · AED · DOP · and 20+ more.
Full list in `CurrencyConversion.gs`.

---

## Daily Auto-Sync

Run `installTriggers()` once from the Apps Script editor to install a daily 7 AM trigger that refreshes:
- FX exchange rates
- Plaid account balances
- US property values (Rentcast)
