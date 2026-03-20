// ============================================================
// PROPERTY VALUES — Auto-pull US property valuations via Rentcast
// Compatible with MNW4 Tiller-based Code.gs
// ============================================================
// SETUP:
//   1. Sign up at https://rentcast.io → free account (50 req/month)
//   2. Copy your API key
//   3. In Apps Script: Project Settings > Script Properties
//      Add: RENTCAST_API_KEY = <your key>
// ============================================================

// These are the US property names as they appear in the Assets/Balances sheet.
// Update this list if names change.
var US_PROPERTY_NAMES = [
  '709 Kuhlman Road Houston TX',
  'ASC 1167 McBride Ave Woodland Park NJ',
  'Clinic Clifton NJ 1117 US-46 Suite 205',
  'Clinic San Diego CA 5330 Carroll Canyon',
  'Sovereign Property Orlando FL',
  'Sovereign Property San Antonio Alamo',
  '1405 Plantation Vlg Dorado PR',
  'Otium 65 Condado Ave San Juan',
  'Otium Villa Internacional #22 San Juan'
];

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

  var sheet = getSheet_(SHEET.ASSETS);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Assets sheet not found.');
    return;
  }

  var data = sheet.getDataRange().getValues();
  var updated = 0;
  var errors  = [];

  for (var i = 1; i < data.length; i++) {
    var row      = data[i];
    var name     = String(row[ASSET_COL.NAME - 1] || '').trim();
    var currency = String(row[ASSET_COL.CURRENCY - 1] || '').trim();

    // Only process USD assets whose name is in our US property list
    if (!name) continue;
    if (currency !== 'USD') continue;
    if (!isUSProperty_(name)) continue;

    var result = getRentcastValue_(name, apiKey);

    if (!result.success) {
      errors.push(name + ': ' + result.error);
      Utilities.sleep(500);
      continue;
    }

    var oldValue = Number(row[ASSET_COL.LOCAL_VALUE - 1]) || 0;
    var newValue = result.value;
    var sheetRow = i + 1;

    // Update Local Value, USD Rate, USD Value, Notes
    sheet.getRange(sheetRow, ASSET_COL.LOCAL_VALUE).setValue(newValue);
    sheet.getRange(sheetRow, ASSET_COL.USD_RATE).setValue(1);
    sheet.getRange(sheetRow, ASSET_COL.USD_VALUE).setValue(newValue);

    // Update delta
    if (ASSET_COL.DELTA) {
      sheet.getRange(sheetRow, ASSET_COL.DELTA).setValue(newValue - oldValue);
    }

    // Append Rentcast range to notes
    var currentNotes = String(row[ASSET_COL.NOTES - 1] || '');
    var rangeNote = 'Rentcast ' + formatDate_(new Date()) + ': $' +
                    formatNumber_(result.lowValue) + '–$' + formatNumber_(result.highValue);
    if (currentNotes.indexOf('Rentcast') !== -1) {
      currentNotes = currentNotes.replace(/Rentcast[^|]*/g, rangeNote);
    } else {
      currentNotes = currentNotes ? currentNotes + ' | ' + rangeNote : rangeNote;
    }
    sheet.getRange(sheetRow, ASSET_COL.NOTES).setValue(currentNotes);

    updated++;
    Utilities.sleep(600); // Stay within free tier rate limits
  }

  // Also update Balances sheet if values changed
  if (updated > 0) {
    syncAssetsToBalances_();
    recalcUsdValues_();
  }

  var msg = updated + ' US properties updated.';
  if (errors.length > 0) msg += '\n\nNot found:\n' + errors.join('\n');
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'Rentcast Update', 10);
}

// ── Rentcast API call ────────────────────────────────────────
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
    if (code === 429) return { success: false, error: 'Rate limit hit (50/month on free tier)' };
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

// ── Lookup a single property on demand ───────────────────────
function lookupSingleProperty() {
  var ui = SpreadsheetApp.getUi();
  var apiKey = PropertiesService.getScriptProperties().getProperty('RENTCAST_API_KEY');
  if (!apiKey) { ui.alert('RENTCAST_API_KEY not set in Script Properties.'); return; }

  var resp = ui.prompt('Property Lookup', 'Enter full US address:', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  var address = resp.getResponseText().trim();
  if (!address) return;

  var result = getRentcastValue_(address, apiKey);
  if (result.success) {
    ui.alert(
      'Rentcast Estimate',
      address + '\n\nValue: $' + formatNumber_(result.value) +
      '\nRange: $' + formatNumber_(result.lowValue) + ' – $' + formatNumber_(result.highValue),
      ui.ButtonSet.OK
    );
  } else {
    ui.alert('Could not get estimate: ' + result.error);
  }
}

// ── Sync updated asset values back to Balances sheet ─────────
function syncAssetsToBalances_() {
  var assetsSheet = getSheet_(SHEET.ASSETS);
  var balSheet    = getSheet_(SHEET.BALANCES);
  if (!assetsSheet || !balSheet) return;

  var assetData = assetsSheet.getDataRange().getValues();
  var balData   = balSheet.getDataRange().getValues();

  for (var a = 1; a < assetData.length; a++) {
    var assetName = String(assetData[a][ASSET_COL.NAME - 1] || '').trim();
    var usdValue  = assetData[a][ASSET_COL.USD_VALUE - 1];
    if (!assetName || !usdValue) continue;

    for (var b = 0; b < balData.length; b++) {
      if (String(balData[b][0] || '').trim() === assetName) {
        balSheet.getRange(b + 1, 4).setValue(usdValue); // Column D = value
        break;
      }
    }
  }
}

// ── Check if asset name matches a known US property ──────────
function isUSProperty_(name) {
  for (var i = 0; i < US_PROPERTY_NAMES.length; i++) {
    if (name.toLowerCase().indexOf(US_PROPERTY_NAMES[i].toLowerCase()) !== -1 ||
        US_PROPERTY_NAMES[i].toLowerCase().indexOf(name.toLowerCase()) !== -1) {
      return true;
    }
  }
  // Also catch anything that looks like a US address (state abbreviation)
  return /\b(TX|NJ|CA|FL|NY|PR|HI|AZ|CO|WA|IL|GA|MA|PA|OH)\b/.test(name);
}

// ── Helpers ───────────────────────────────────────────────────
function formatNumber_(n) {
  if (!n) return '0';
  return Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function formatDate_(d) {
  return (d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getFullYear();
}
