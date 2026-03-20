// ============================================================
// PROPERTY VALUES — Auto-pull US property valuations via Rentcast
// ============================================================
// Zillow deprecated their public API (moved to Bridge Interactive/paid).
// Rentcast.io provides free property AVM (Automated Valuation Model).
// Free tier: 50 requests/month — sufficient for this use case.
//
// SETUP:
//   1. Sign up at https://rentcast.io → free account
//   2. Get your API key from the dashboard
//   3. In Apps Script: Extensions > Apps Script > Project Settings > Script Properties
//      Add: RENTCAST_API_KEY = <your key>
//
// ADD TO Code.gs onOpen() menu (under "Update Exchange Rates"):
//   .addItem('Refresh US Property Values (Rentcast)', 'refreshUSPropertyValues')
//
// ADD TO Code.gs dailySync():
//   try { refreshUSPropertyValues(); } catch(e) { Logger.log('Property values error: ' + e); }
// ============================================================

/**
 * Refresh all US Real Estate assets using Rentcast AVM.
 * Targets rows where Category = "Real Estate" and Subcategory = "United States".
 * Uses the asset Name field as the address.
 */
function refreshUSPropertyValues() {
  var apiKey = PropertiesService.getScriptProperties().getProperty('RENTCAST_API_KEY');
  if (!apiKey) {
    SpreadsheetApp.getUi().alert(
      'Rentcast API key not set.\n\n' +
      'Add RENTCAST_API_KEY in:\nExtensions > Apps Script > Project Settings > Script Properties'
    );
    return;
  }

  var sheet = getSheet_(SHEET.ASSETS);
  if (!sheet) return;

  var data = sheet.getDataRange().getValues();
  var updated = 0;
  var errors  = [];

  for (var i = 1; i < data.length; i++) {
    var row      = data[i];
    var category = row[ASSET_COL.CATEGORY    - 1];
    var subcat   = row[ASSET_COL.SUBCATEGORY - 1];
    var name     = row[ASSET_COL.NAME        - 1];
    var source   = row[ASSET_COL.SOURCE      - 1];

    // Only process US Real Estate with Manual source
    if (category !== 'Real Estate' || subcat !== 'United States' || source !== 'Manual') continue;
    if (!name) continue;

    var result = getRentcastValue_(name, apiKey);

    if (!result.success) {
      errors.push(name + ': ' + result.error);
      Logger.log('Rentcast error for "' + name + '": ' + result.error);
      Utilities.sleep(500);
      continue;
    }

    var oldValue = row[ASSET_COL.LOCAL_VALUE - 1];
    var newValue = result.value;

    // Update the sheet row directly
    var sheetRow = i + 1;
    sheet.getRange(sheetRow, ASSET_COL.LOCAL_VALUE).setValue(newValue);
    sheet.getRange(sheetRow, ASSET_COL.LAST_UPDATED).setValue(new Date());

    // Add Rentcast range to Notes if not already there
    var currentNotes = (row[ASSET_COL.NOTES - 1] || '').toString();
    var rangeNote = 'Rentcast: $' + formatNumber_(result.lowValue) + ' – $' + formatNumber_(result.highValue);
    // Replace existing Rentcast note or append
    if (currentNotes.indexOf('Rentcast:') !== -1) {
      currentNotes = currentNotes.replace(/Rentcast:[^|]*/g, rangeNote);
    } else {
      currentNotes = currentNotes ? currentNotes + ' | ' + rangeNote : rangeNote;
    }
    sheet.getRange(sheetRow, ASSET_COL.NOTES).setValue(currentNotes);

    // Log the change
    if (typeof oldValue === 'number' && Math.abs(newValue - oldValue) > 0.01) {
      logToMasterHistory_(name, oldValue, newValue, 'USD');
      Logger.log('Updated ' + name + ': $' + oldValue + ' → $' + newValue);
    }

    updated++;
    Utilities.sleep(600); // Stay within free tier rate limits
  }

  // Recalculate all USD values and refresh dashboard
  recalcUsdValues_();

  // Report results
  var msg = 'Property refresh complete.\n' + updated + ' properties updated.';
  if (errors.length > 0) {
    msg += '\n\nCould not fetch:\n' + errors.slice(0, 5).join('\n');
    if (errors.length > 5) msg += '\n...and ' + (errors.length - 5) + ' more (see logs)';
  }
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'Rentcast Update', 10);
  Logger.log(msg);
}

/**
 * Fetch property AVM from Rentcast for a given address.
 * @param {string} address - Full property address
 * @param {string} apiKey  - Rentcast API key
 * @return {Object} { success, value, lowValue, highValue, error }
 */
function getRentcastValue_(address, apiKey) {
  try {
    var url = 'https://api.rentcast.io/v1/avm/value?address=' + encodeURIComponent(address);
    var resp = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: { 'X-Api-Key': apiKey },
      muteHttpExceptions: true
    });

    var code = resp.getResponseCode();
    if (code === 404) return { success: false, error: 'Address not found' };
    if (code === 429) return { success: false, error: 'Rate limit reached (50 req/month on free tier)' };
    if (code !== 200) return { success: false, error: 'HTTP ' + code };

    var data = JSON.parse(resp.getContentText());
    var value = data.price || data.value || data.priceRangeMid || null;
    if (!value) return { success: false, error: 'No valuation returned' };

    return {
      success:   true,
      value:     Math.round(value),
      lowValue:  Math.round(data.priceLow  || value * 0.95),
      highValue: Math.round(data.priceHigh || value * 1.05)
    };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

/**
 * Look up a single property value on demand (callable from menu or custom function).
 * Shows result in a dialog.
 */
function lookupSingleProperty() {
  var ui = SpreadsheetApp.getUi();
  var apiKey = PropertiesService.getScriptProperties().getProperty('RENTCAST_API_KEY');
  if (!apiKey) {
    ui.alert('RENTCAST_API_KEY not set in Script Properties.');
    return;
  }

  var resp = ui.prompt(
    'Property Lookup',
    'Enter full US address (e.g. 709 Kuhlman Road Houston TX):',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  var address = resp.getResponseText().trim();
  if (!address) return;

  var result = getRentcastValue_(address, apiKey);
  if (result.success) {
    ui.alert(
      'Rentcast Estimate: ' + address,
      'Estimated Value: $' + formatNumber_(result.value) +
      '\nRange: $' + formatNumber_(result.lowValue) + ' – $' + formatNumber_(result.highValue),
      ui.ButtonSet.OK
    );
  } else {
    ui.alert('Could not get estimate: ' + result.error);
  }
}

// ============================================================
// INTERNATIONAL PROPERTY CURRENCY NOTES
// ============================================================
// The Colombian and European properties in the Assets sheet currently
// have currency set to COP/EUR but values stored in USD (per the Notes column).
//
// To enable automatic currency conversion for international properties:
//   1. Update Local Value column to the actual local currency amount
//      e.g. Vingt Paris: change Local Value from 3,400,000 (USD) to
//           actual EUR value (e.g. 3,148,148 EUR at 1.08 rate)
//   2. Make sure Currency column is set correctly (EUR, COP, etc.)
//   3. Run "Update Exchange Rates" — recalcUsdValues_() will auto-convert
//
// The ExchangeRates.gs already handles this correctly.
// ============================================================

// Helper: format a number with commas
function formatNumber_(n) {
  if (!n) return '0';
  return Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}
