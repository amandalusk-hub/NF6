/**
 * PropertyValues.gs
 * Zillow Zestimate lookup for US properties.
 * Uses the Zillow API via RapidAPI.
 *
 * SETUP:
 *   1. Get a free RapidAPI key at https://rapidapi.com
 *   2. Subscribe to "Zillow Com" API (free tier available)
 *   3. In Apps Script: Project Settings > Script Properties
 *      Add property: RAPIDAPI_KEY = <your key>
 */

// ─────────────────────────────────────────────
// Custom Sheet Function
// Usage in a cell: =ZESTIMATE("123 Main St, Austin, TX 78701")
// ─────────────────────────────────────────────
function ZESTIMATE(address) {
  if (!address) return '';

  var cache = CacheService.getScriptCache();
  var cacheKey = 'zestimate_' + address.toString().replace(/\s+/g, '_');
  var cached = cache.get(cacheKey);
  if (cached) return Number(cached);

  var apiKey = PropertiesService.getScriptProperties().getProperty('RAPIDAPI_KEY');
  if (!apiKey) return 'ERROR: Set RAPIDAPI_KEY in Script Properties';

  try {
    // Step 1: Search for property to get zpid
    var searchUrl = 'https://zillow-com1.p.rapidapi.com/propertyExtendedSearch?location='
      + encodeURIComponent(address) + '&home_type=Houses,Condos,MultiFamily,Townhomes';

    var searchResp = UrlFetchApp.fetch(searchUrl, {
      method: 'GET',
      headers: {
        'X-RapidAPI-Key': apiKey,
        'X-RapidAPI-Host': 'zillow-com1.p.rapidapi.com'
      },
      muteHttpExceptions: true
    });

    if (searchResp.getResponseCode() !== 200) {
      return 'API Error: ' + searchResp.getResponseCode();
    }

    var searchData = JSON.parse(searchResp.getContentText());
    if (!searchData.props || searchData.props.length === 0) {
      return 'Not found';
    }

    var prop = searchData.props[0];
    var zestimate = prop.zestimate;

    // Step 2: If no zestimate in search, fetch property detail
    if (!zestimate && prop.zpid) {
      var detailUrl = 'https://zillow-com1.p.rapidapi.com/property?zpid=' + prop.zpid;
      var detailResp = UrlFetchApp.fetch(detailUrl, {
        method: 'GET',
        headers: {
          'X-RapidAPI-Key': apiKey,
          'X-RapidAPI-Host': 'zillow-com1.p.rapidapi.com'
        },
        muteHttpExceptions: true
      });

      if (detailResp.getResponseCode() === 200) {
        var detailData = JSON.parse(detailResp.getContentText());
        zestimate = detailData.zestimate || detailData.price;
      }
    }

    if (!zestimate) return 'No Zestimate';

    // Cache for 6 hours
    cache.put(cacheKey, zestimate.toString(), 21600);
    return Number(zestimate);

  } catch (e) {
    console.error('ZESTIMATE error:', e);
    return 'Error: ' + e.message;
  }
}

// ─────────────────────────────────────────────
// Bulk refresh: call this via a time-based trigger (e.g. daily)
// Updates all cells in the "REAL ESTATE" section that have a
// Zillow address in column B.
// ─────────────────────────────────────────────
function refreshAllZestimates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Net Worth') || ss.getSheets()[0];
  var data = sheet.getDataRange().getValues();

  var inRealEstate = false;

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var label = (row[0] || '').toString().toUpperCase();

    // Detect section header
    if (label.indexOf('REAL ESTATE') !== -1 || label.indexOf('PROPERTY') !== -1) {
      inRealEstate = true;
      continue;
    }
    // Stop at next major section
    if (inRealEstate && label && label === label.toUpperCase() && label.length > 3 &&
        label.indexOf('REAL ESTATE') === -1 && row[1] === '') {
      inRealEstate = false;
    }

    if (!inRealEstate) continue;

    var address = (row[1] || '').toString().trim(); // Column B = address
    var isUSAddress = isUSBasedAddress(address);

    if (address && isUSAddress) {
      var value = ZESTIMATE(address);
      if (typeof value === 'number') {
        // Write to column H (USD Value) — adjust column index as needed
        sheet.getRange(i + 1, 8).setValue(value);
        sheet.getRange(i + 1, 9).setValue(new Date()); // Last updated timestamp
      }
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('Zestimates refreshed!', 'Property Values', 5);
}

// Simple heuristic: US addresses contain a state abbreviation or zip code
function isUSBasedAddress(address) {
  if (!address) return false;
  var usStatePattern = /\b(AL|AK|AZ|AR|CA|CO|CT|DE|FL|GA|HI|ID|IL|IN|IA|KS|KY|LA|ME|MD|MA|MI|MN|MS|MO|MT|NE|NV|NH|NJ|NM|NY|NC|ND|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VT|VA|WA|WV|WI|WY|DC)\b/i;
  var usZipPattern = /\b\d{5}(-\d{4})?\b/;
  return usStatePattern.test(address) || usZipPattern.test(address);
}

// ─────────────────────────────────────────────
// Install daily refresh trigger
// Run once manually from Apps Script editor
// ─────────────────────────────────────────────
function installZilowTrigger() {
  // Remove existing triggers for this function
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'refreshAllZestimates') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('refreshAllZestimates')
    .timeBased()
    .everyDays(1)
    .atHour(6) // 6 AM daily
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Daily Zillow refresh scheduled for 6 AM', 'Trigger Installed', 5
  );
}
