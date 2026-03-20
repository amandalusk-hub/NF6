/**
 * CurrencyConversion.gs
 * Convert international property/asset values to USD.
 *
 * Features:
 *  - =TO_USD(amount, "COP")   → converts Colombian Pesos to USD
 *  - =TO_USD(amount, "EUR")   → converts Euros to USD
 *  - =FX_RATE("COP")          → returns today's COP/USD rate
 *  - refreshAllFxRates()      → bulk-updates all foreign currency rows
 *
 * Uses exchangerate-api.com (free, no key required for basic use).
 * For higher volume, add FREE_KEY from https://www.exchangerate-api.com
 */

// ─────────────────────────────────────────────
// Custom Sheet Functions
// ─────────────────────────────────────────────

/**
 * Convert an amount in a foreign currency to USD.
 * @param {number} amount - The value in the foreign currency
 * @param {string} currencyCode - ISO 4217 currency code (e.g. "COP", "EUR", "MXN")
 * @return {number} USD equivalent
 * @customfunction
 */
function TO_USD(amount, currencyCode) {
  if (!amount || !currencyCode) return '';
  currencyCode = currencyCode.toString().trim().toUpperCase();
  if (currencyCode === 'USD') return Number(amount);

  var rate = FX_RATE(currencyCode);
  if (typeof rate !== 'number') return rate; // Pass through errors
  return Number(amount) * rate;
}

/**
 * Get today's exchange rate for a currency to USD.
 * @param {string} currencyCode - ISO 4217 currency code (e.g. "COP", "EUR")
 * @return {number} Rate (1 unit of currencyCode = X USD)
 * @customfunction
 */
function FX_RATE(currencyCode) {
  if (!currencyCode) return '';
  currencyCode = currencyCode.toString().trim().toUpperCase();
  if (currencyCode === 'USD') return 1;

  var cache = CacheService.getScriptCache();
  var cacheKey = 'fx_' + currencyCode + '_USD';
  var cached = cache.get(cacheKey);
  if (cached) return Number(cached);

  try {
    var apiKey = PropertiesService.getScriptProperties().getProperty('EXCHANGERATE_API_KEY');
    var baseUrl = apiKey
      ? 'https://v6.exchangerate-api.com/v6/' + apiKey + '/latest/' + currencyCode
      : 'https://open.er-api.com/v6/latest/' + currencyCode;

    var resp = UrlFetchApp.fetch(baseUrl, { muteHttpExceptions: true });

    if (resp.getResponseCode() !== 200) {
      // Fallback: try Google Finance GOOGLEFINANCE workaround
      return getFxRateFromGoogleFinance(currencyCode);
    }

    var data = JSON.parse(resp.getContentText());
    if (!data.rates || !data.rates['USD']) {
      return getFxRateFromGoogleFinance(currencyCode);
    }

    var rate = data.rates['USD'];
    // Cache for 4 hours
    cache.put(cacheKey, rate.toString(), 14400);
    return rate;

  } catch (e) {
    console.error('FX_RATE error for ' + currencyCode + ':', e);
    return 'Error: ' + e.message;
  }
}

/**
 * Fallback: use Google's built-in Finance data via a dummy spreadsheet query.
 * Less reliable but requires no external API.
 */
function getFxRateFromGoogleFinance(currencyCode) {
  try {
    // Google Finance currency pair format: CURRENCY:USDXXX
    var pair = 'CURRENCY:' + currencyCode + 'USD';
    var url = 'https://finance.google.com/finance?q=' + pair + '&output=json';
    // This is a known-tricky endpoint; use the Sheets formula approach instead
    // Best fallback is to return an error and let the user know
    return 'No rate: check ' + currencyCode;
  } catch(e) {
    return 'Rate unavailable';
  }
}

// ─────────────────────────────────────────────
// Supported currencies reference (for the sheet dropdown)
// ─────────────────────────────────────────────
var SUPPORTED_CURRENCIES = {
  'AED': 'UAE Dirham',
  'ARS': 'Argentine Peso',
  'AUD': 'Australian Dollar',
  'BRL': 'Brazilian Real',
  'CAD': 'Canadian Dollar',
  'CHF': 'Swiss Franc',
  'CLP': 'Chilean Peso',
  'CNY': 'Chinese Yuan',
  'COP': 'Colombian Peso',
  'EUR': 'Euro',
  'GBP': 'British Pound',
  'HKD': 'Hong Kong Dollar',
  'INR': 'Indian Rupee',
  'JPY': 'Japanese Yen',
  'KRW': 'South Korean Won',
  'MXN': 'Mexican Peso',
  'MYR': 'Malaysian Ringgit',
  'NOK': 'Norwegian Krone',
  'NZD': 'New Zealand Dollar',
  'PEN': 'Peruvian Sol',
  'PHP': 'Philippine Peso',
  'PLN': 'Polish Zloty',
  'SAR': 'Saudi Riyal',
  'SEK': 'Swedish Krona',
  'SGD': 'Singapore Dollar',
  'THB': 'Thai Baht',
  'TRY': 'Turkish Lira',
  'TWD': 'Taiwan Dollar',
  'USD': 'US Dollar',
  'VND': 'Vietnamese Dong',
  'ZAR': 'South African Rand'
};

// ─────────────────────────────────────────────
// Bulk refresh: updates all foreign currency property rows
// Looks for rows where column C has a currency code != "USD"
// and column D has a local-currency value, then writes USD
// equivalent to column E.
//
// Adjust column indices to match your sheet layout.
// ─────────────────────────────────────────────
function refreshAllFxRates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Net Worth') || ss.getSheets()[0];
  var data = sheet.getDataRange().getValues();

  var updatedCount = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    // Expected layout (0-indexed):
    //   col A (0) = Asset name
    //   col B (1) = Address / description
    //   col C (2) = Currency code (e.g. "COP", "EUR")
    //   col D (3) = Value in local currency
    //   col E (4) = USD Value (auto-calculated here)
    //   col F (5) = Last updated

    var currencyCode = (row[2] || '').toString().trim().toUpperCase();
    var localValue = row[3];

    if (!currencyCode || currencyCode === 'USD' || !localValue || isNaN(localValue)) continue;

    var usdValue = TO_USD(Number(localValue), currencyCode);
    if (typeof usdValue === 'number') {
      sheet.getRange(i + 1, 5).setValue(usdValue);       // Column E: USD value
      sheet.getRange(i + 1, 6).setValue(new Date());      // Column F: Last updated
      updatedCount++;
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Updated ' + updatedCount + ' foreign currency values to USD.',
    'FX Rates Refreshed', 5
  );
}

// ─────────────────────────────────────────────
// Install daily FX refresh trigger
// Run once manually from Apps Script editor
// ─────────────────────────────────────────────
function installFxTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'refreshAllFxRates') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('refreshAllFxRates')
    .timeBased()
    .everyDays(1)
    .atHour(7) // 7 AM daily
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Daily FX rate refresh scheduled for 7 AM', 'Trigger Installed', 5
  );
}

// ─────────────────────────────────────────────
// Add currency dropdown validation to a range
// Call once to set up the sheet
// ─────────────────────────────────────────────
function setupCurrencyDropdowns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Net Worth') || ss.getSheets()[0];

  var codes = Object.keys(SUPPORTED_CURRENCIES).sort();
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(codes, true)
    .setAllowInvalid(false)
    .build();

  // Apply to column C (currency code column), rows 2 through 200
  // Adjust range as needed for your sheet layout
  var range = sheet.getRange('C2:C200');
  range.setDataValidation(rule);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Currency dropdowns added to column C (rows 2-200)', 'Setup Complete', 5
  );
}
