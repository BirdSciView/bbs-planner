This is web-app that can help Breeding Bird Surveyors complete their routes.  

First, create a google sheet.  

Select 'Extensions ->Apps Script -> Paste the code below as a code.gs file.  Make sure the sheet ID corresponds with your sheet.

Deploy

Second, Host the Index.html file (github pages is a good option).

Visit the associated webpage.  Connect the webpage to your spreadsheet by pasting the deployments url in the setup screen and select 'Test Connection'.

code.gs:
// BBS Route Planner — Code.gs
// Deploy as: Web App | Execute as: Me | Access: Anyone
// This script acts as a pure data API — the HTML is hosted on GitHub Pages.

// ── CONFIG ────────────────────────────────────────────────────────────────────
var SPREADSHEET_ID = ''; // Paste your Sheet ID here, e.g. '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms'

// ── CORS HEADERS ──────────────────────────────────────────────────────────────
function corsResponse(data) {
  var output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ── ROUTING ───────────────────────────────────────────────────────────────────
function doGet(e) {
  var action = e && e.parameter && e.parameter.action ? e.parameter.action : '';
  try {
    if (action === 'getRoutes')            return corsResponse(getRoutes());
    if (action === 'getStopData')          return corsResponse(getAllStopData(e.parameter.route));
    if (action === 'getRouteSheetSummary') return corsResponse(getRouteSheetSummary(e.parameter.route));
    return corsResponse({ ok: true, message: 'BBS API ready' });
  } catch(err) {
    return corsResponse({ ok: false, error: err.message });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action || '';
    if (action === 'saveRoutes')   return corsResponse(saveRoutes(body.routes));
    if (action === 'saveStopData') return corsResponse(saveStopData(body.entry));
    return corsResponse({ ok: false, error: 'Unknown action: ' + action });
  } catch(err) {
    return corsResponse({ ok: false, error: err.message });
  }
}

// ── SPREADSHEET ───────────────────────────────────────────────────────────────
function getSpreadsheet() {
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);
  return SpreadsheetApp.getActiveSpreadsheet();
}

// ── ROUTES ────────────────────────────────────────────────────────────────────
function getRoutes() {
  try {
    var props = PropertiesService.getUserProperties();
    var stored = props.getProperty('bbs_routes');
    if (!stored) return { ok: true, data: {} };
    var parsed = JSON.parse(stored);
    return { ok: true, data: (parsed && typeof parsed === 'object') ? parsed : {} };
  } catch(e) {
    return { ok: false, error: e.message, data: {} };
  }
}

function saveRoutes(routes) {
  try {
    var props = PropertiesService.getUserProperties();
    props.setProperty('bbs_routes', JSON.stringify(routes || {}));
    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── STOP DATA ─────────────────────────────────────────────────────────────────
function saveStopData(entry) {
  try {
    if (!entry || !entry.route) return { ok: false, error: 'No entry data' };
    var ss = getSpreadsheet();
    if (!ss) return { ok: false, error: 'No spreadsheet found' };

    var sheetName = String(entry.route).substring(0, 100);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(['Route','Stop','Date','Cars','Excessive Noise','Sky','Wind','Temp (C)','Precipitation','Latitude','Longitude','Saved At']);
      sheet.getRange(1,1,1,12).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    var data = sheet.getDataRange().getValues();
    var existingRow = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(entry.stop)) { existingRow = i + 1; break; }
    }

    var row = [
      entry.route || '', entry.stop || '', entry.date || '',
      entry.cars || 0, entry.noise ? 'Yes' : 'No',
      entry.sky || '', entry.wind || '',
      (entry.temp !== undefined && entry.temp !== '') ? entry.temp : '',
      entry.precip || '',
      (entry.lat != null) ? entry.lat : '',
      (entry.lng != null) ? entry.lng : '',
      new Date().toISOString()
    ];

    if (existingRow > 0) {
      sheet.getRange(existingRow, 1, 1, 12).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function getAllStopData(routeName) {
  try {
    if (!routeName) return { ok: true, data: [] };
    var ss = getSpreadsheet();
    if (!ss) return { ok: true, data: [] };
    var sheet = ss.getSheetByName(routeName);
    if (!sheet || sheet.getLastRow() < 2) return { ok: true, data: [] };
    var rows = sheet.getDataRange().getValues();
    var headers = rows[0];
    var data = rows.slice(1).map(function(row) {
      var obj = {};
      headers.forEach(function(h, i) { obj[String(h)] = row[i]; });
      return obj;
    });
    return { ok: true, data: data };
  } catch(e) {
    return { ok: false, error: e.message, data: [] };
  }
}

function getRouteSheetSummary(routeName) {
  try {
    var ss = getSpreadsheet();
    if (!ss) return { ok: false, error: 'No spreadsheet', count: 0, url: '' };
    var url = ss.getUrl() || '';
    var sheet = routeName ? ss.getSheetByName(routeName) : null;
    var count = sheet ? Math.max(0, sheet.getLastRow() - 1) : 0;
    return { ok: true, count: count, url: url };
  } catch(e) {
    return { ok: false, error: e.message, count: 0, url: '' };
  }
}
