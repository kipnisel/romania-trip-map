// Google Apps Script - Backend for the Romania trip map.
// Stores & returns (1) reviews per place, (2) new suggestions, in a Google Sheet.
// Setup: see README.md. Deploy as Web App (Execute as: Me, Access: Anyone),
// then paste the /exec URL into index.html (SCRIPT_URL).

// SHEET_ID: paste your spreadsheet ID here (the long code in the sheet URL,
// between /d/ and /edit). Needed for a standalone script (e.g. set up from mobile).
// Leave '' if the script is bound to the sheet (Extensions > Apps Script).
var SHEET_ID = '';

function ss_() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

// Column headers per sheet
var SHEETS = {
  Reviews:     ['timestamp', 'poi', 'name', 'rating', 'comment'],
  Suggestions: ['timestamp', 'name', 'place', 'where', 'link', 'reason']
};

function getSheet_(name) {
  var ss = ss_();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(SHEETS[name]);
  }
  return sh;
}

// Read a sheet and return an array of objects keyed by the header row
function readSheet_(name) {
  var sh = getSheet_(name);
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (!row.join('')) continue;
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var key = String(headers[j]);
      var val = row[j];
      if (key === 'timestamp' && val) val = new Date(val).toISOString();
      obj[key] = val;
    }
    out.push(obj);
  }
  return out;
}

function clip_(v, n) { return String(v == null ? '' : v).slice(0, n); }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Read - returns reviews + suggestions. Supports JSONP (?callback=...)
function doGet(e) {
  var data = { reviews: readSheet_('Reviews'), suggestions: readSheet_('Suggestions') };
  var json = JSON.stringify(data);
  var cb = e && e.parameter && e.parameter.callback;
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// Write - review or suggestion (by d.type). Sent as text/plain to bypass CORS.
function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);

    if (d.type === 'suggestion') {
      if (!d.place) return json_({ ok: false, error: 'missing place' });
      getSheet_('Suggestions').appendRow([
        new Date(), clip_(d.name, 40) || 'anon', clip_(d.place, 80),
        clip_(d.where, 60), clip_(d.link, 200), clip_(d.reason, 500)
      ]);
      return json_({ ok: true });
    }

    var poi = clip_(d.poi, 60);
    if (!poi || !d.comment) return json_({ ok: false, error: 'missing' });
    getSheet_('Reviews').appendRow([
      new Date(), poi, clip_(d.name, 40) || 'anon',
      parseInt(d.rating) || '', clip_(d.comment, 600)
    ]);
    return json_({ ok: true });

  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}
