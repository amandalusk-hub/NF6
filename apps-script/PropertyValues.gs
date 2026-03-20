// ============================================================
// PROPERTY VALUES — US property auto-valuation via Rentcast
// Companion to Code.gs (Family Office Wealth Tracker)
//
// SETUP:
//   1. Sign up at https://rentcast.io → free account (50 req/month)
//   2. Copy your API key
//   3. In Apps Script: Project Settings > Script Properties
//      Add: RENTCAST_API_KEY = <your key>
//
// HOW IT FINDS PROPERTIES:
//   Any asset with Category = "Real Estate" whose Notes field
//   contains an address in the format:
//     address: 123 Main St, Houston, TX 77001
//   will be auto-updated when refreshUSPropertyValues() runs.
//
// USAGE:
//   - Run refreshUSPropertyValues() to bulk-update all US properties
//   - Run lookupSingleProperty() for a one-off estimate
// ============================================================

/**
 * Refresh Rentcast AVM values for all US Real Estate assets.
 * Reads Assets sheet, finds Real Estate rows with a US address
 * in the Notes field, calls Rentcast, and writes updated values.
 *
 * Asset sheet columns (as defined in Code.gs COL.ASSETS):
 *   1  ID            6  Local Value   11 Date Added
 *   2  Name          7  USD Rate      12 Last Updated
 *   3  Category      8  USD Value     13 Notes
 *   4  Entity        9  My Share %    14 Plaid Account ID
 *   5  Currency     10  My Share USD
 */
function refreshUSPropertyValues() {
  var apiKey = PropertiesService.getScriptProperties().getProperty('RENTCAST_API_KEY');
  if (!apiKey) {
    SpreadsheetApp.getUi().alert(
      'Rentcast API key not set.\n\n' +
      'Go to: Extensions > Apps Script > Project Settings > Script Properties\n' +
      'Add property: RENTCAST_API_KEY = your key from rentcast.io'
    );
    return;
  }

  var sheet = getSheet_('ASSETS');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Assets sheet not found.');
    return;
  }

  var rows    = sheet.getDataRange().getValues();
  var updated = 0;
  var errors  = [];

  for (var i = 1; i < rows.length; i++) {
    var category = String(rows[i][2] || '').trim();   // col 3
    var currency = String(rows[i][4] || '').trim();   // col 5
    var notes    = String(rows[i][12] || '').trim();  // col 13

    if (category !== 'Real Estate') continue;
    if (currency !== 'USD') continue;

    // Extract address from notes: "address: 123 Main St, City, TX 12345"
    var addrMatch = notes.match(/address:\s*([^\|]+?)(?:\s*\||$)/i);
    if (!addrMatch) continue;
    var address = addrMatch[1].trim();

    // Only US addresses (state abbreviation present)
    if (!/\b(AL|AK|AZ|AR|CA|CO|CT|DE|FL|GA|HI|ID|IL|IN|IA|KS|KY|LA|ME|MD|MA|MI|MN|MS|MO|MT|NE|NV|NH|NJ|NM|NY|NC|ND|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VT|VA|WA|WV|WI|WY|DC)\b/i.test(address)) continue;

    var result = getRentcastEstimate_(address, apiKey);

    if (!result.success) {
      errors.push(rows[i][1] + ': ' + result.error);
      Utilities.sleep(500);
      continue;
    }

    var sheetRow = i + 1;
    var oldUsd   = Number(rows[i][7]) || 0;  // col 8 USD Value
    var newUsd   = result.value;
    var sharePct = Number(rows[i][8]) || 100; // col 9 My Share %

    sheet.getRange(sheetRow, 6).setValue(newUsd);              // Local Value
    sheet.getRange(sheetRow, 7).setValue(1);                   // USD Rate
    sheet.getRange(sheetRow, 8).setValue(newUsd);              // USD Value
    sheet.getRange(sheetRow, 10).setValue(newUsd * sharePct / 100); // My Share USD
    sheet.getRange(sheetRow, 12).setValue(new Date());         // Last Updated

    // Append Rentcast range note
    var rangeNote = 'Rentcast ' + formatDate_(new Date()) +
                    ': $' + formatNumber_(result.lowValue) +
                    '–$' + formatNumber_(result.highValue);
    var newNotes = notes.replace(/Rentcast [^\|]*/g, rangeNote);
    if (newNotes === notes) {
      newNotes = notes ? notes + ' | ' + rangeNote : rangeNote;
    }
    sheet.getRange(sheetRow, 13).setValue(newNotes);           // Notes

    if (Math.abs(newUsd - oldUsd) > 0.01) {
      logHistory_(rows[i][1], oldUsd, newUsd, 'USD', 'Auto-updated via Rentcast');
    }

    updated++;
    Utilities.sleep(600); // Stay within free tier rate limits
  }

  var msg = updated + ' US propert' + (updated === 1 ? 'y' : 'ies') + ' updated.';
  if (errors.length > 0) msg += '\n\nNot updated:\n' + errors.join('\n');
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'Rentcast Update', 10);
}

/**
 * Interactive: prompt for a US address and show the Rentcast estimate.
 */
function lookupSingleProperty() {
  var ui     = SpreadsheetApp.getUi();
  var apiKey = PropertiesService.getScriptProperties().getProperty('RENTCAST_API_KEY');
  if (!apiKey) {
    ui.alert(
      'RENTCAST_API_KEY not set.\n\n' +
      'Go to: Extensions > Apps Script > Project Settings > Script Properties\n' +
      'Add property: RENTCAST_API_KEY = your key from rentcast.io'
    );
    return;
  }

  var resp = ui.prompt('Property Lookup', 'Enter full US address:', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  var address = resp.getResponseText().trim();
  if (!address) return;

  var result = getRentcastEstimate_(address, apiKey);
  if (result.success) {
    ui.alert(
      'Rentcast Estimate',
      address + '\n\n' +
      'Value: $' + formatNumber_(result.value) + '\n' +
      'Range: $' + formatNumber_(result.lowValue) + ' – $' + formatNumber_(result.highValue),
      ui.ButtonSet.OK
    );
  } else {
    ui.alert('Could not get estimate: ' + result.error);
  }
}

// ── Rentcast API ──────────────────────────────────────────────────────────────

function getRentcastEstimate_(address, apiKey) {
  try {
    var url  = 'https://api.rentcast.io/v1/avm/value?address=' + encodeURIComponent(address);
    var resp = UrlFetchApp.fetch(url, {
      method:  'GET',
      headers: { 'X-Api-Key': apiKey },
      muteHttpExceptions: true
    });

    var code = resp.getResponseCode();
    if (code === 404) return { success: false, error: 'Address not found' };
    if (code === 429) return { success: false, error: 'Rate limit (50/month on free tier)' };
    if (code !== 200) return { success: false, error: 'HTTP ' + code };

    var data  = JSON.parse(resp.getContentText());
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

// ── Helpers ───────────────────────────────────────────────────────────────────

function formatNumber_(n) {
  if (!n) return '0';
  return Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function formatDate_(d) {
  return (d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getFullYear();
}
