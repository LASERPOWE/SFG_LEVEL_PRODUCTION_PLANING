/**
 * Apps Script — writes the live "ERP DUMP_PROD_ORD_OPEN MC" JSON
 * into a separate Google Sheet (one chunk per cell in column A).
 * Claude reads that Sheet via the Drive connector and parses it.
 *
 * SETUP (one-time):
 * 1. Paste this whole file into Apps Script editor, Save (Ctrl+S)
 * 2. Top of editor: select function "syncToSheet" -> click Run
 *    Authorize when prompted.
 * 3. Look at the Execution log. You'll see a line like:
 *      SHEET_URL: https://docs.google.com/spreadsheets/d/<ID>/edit
 *    Copy that whole line (or just the ID) and paste into chat.
 *
 * REFRESH later: just run syncToSheet again. URL stays the same — the
 * same Sheet gets overwritten with fresh data each time.
 */

var SPREADSHEET_ID    = '1eN0atckz_PjEQB_7SH05eC4y6EdehGXO0MxMefNWPnE';
var DEFAULT_TAB       = 'ERP DUMP_PROD_ORD_OPEN MC';
var FEED_SHEET_NAME   = 'ERP_LIVE_JSON_FEED';
var FEED_SHEET_PROP   = 'ERP_LIVE_FEED_SHEET_ID';
var CHUNK_SIZE        = 40000;   // safely under Google Sheets' 50k cell limit

function buildPayload(tabName) {
  tabName = tabName || DEFAULT_TAB;
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(tabName);
  if (!sh) {
    return { error: 'Tab not found: ' + tabName,
             available: ss.getSheets().map(function(s){return s.getName();}) };
  }
  var range  = sh.getDataRange();
  var values = range.getValues();
  var rows = [];
  var headers = [];
  if (values.length > 0) {
    headers = values[0].map(function(h){ return String(h || '').trim(); });
    for (var i = 1; i < values.length; i++) {
      var raw = values[i];
      var hasAny = false;
      for (var j = 0; j < raw.length; j++) {
        if (raw[j] !== '' && raw[j] !== null && raw[j] !== undefined) { hasAny = true; break; }
      }
      if (!hasAny) continue;
      var obj = {};
      for (var j = 0; j < headers.length; j++) {
        if (!headers[j]) continue;
        var v = raw[j];
        if (Object.prototype.toString.call(v) === '[object Date]') {
          v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        obj[headers[j]] = v;
      }
      rows.push(obj);
    }
  }
  return {
    tab: tabName,
    columns: headers,
    rows: rows,
    row_count: rows.length,
    exported_at: new Date().toISOString()
  };
}

function syncToSheet() {
  var payload = buildPayload(DEFAULT_TAB);
  var json = JSON.stringify(payload);
  Logger.log('Built payload. Rows: ' + payload.row_count + '  JSON length: ' + json.length);

  // Get or create the feed Spreadsheet
  var props = PropertiesService.getScriptProperties();
  var feedId = props.getProperty(FEED_SHEET_PROP);
  var feedSs;
  if (feedId) {
    try { feedSs = SpreadsheetApp.openById(feedId); } catch (e) { feedSs = null; }
  }
  if (!feedSs) {
    feedSs = SpreadsheetApp.create(FEED_SHEET_NAME);
    props.setProperty(FEED_SHEET_PROP, feedSs.getId());
    Logger.log('Created new feed Spreadsheet.');
  }

  var sh = feedSs.getSheets()[0];
  sh.clear();

  // Split JSON into chunks of CHUNK_SIZE characters
  var chunks = [];
  for (var i = 0; i < json.length; i += CHUNK_SIZE) {
    chunks.push([json.substr(i, CHUNK_SIZE)]);
  }
  // Write header row + chunks
  sh.getRange(1, 1).setValue('JSON_CHUNK');
  sh.getRange(1, 2).setValue('META');
  sh.getRange(1, 2).setNote('Rows: ' + payload.row_count + ' | exported: ' + payload.exported_at);
  if (chunks.length > 0) {
    sh.getRange(2, 1, chunks.length, 1).setValues(chunks);
  }
  // Add terminator row to make it easy to detect end of data
  sh.getRange(chunks.length + 2, 1).setValue('END_OF_DATA');

  var url = feedSs.getUrl();
  Logger.log('SHEET_URL: ' + url);
  Logger.log('SHEET_ID: ' + feedSs.getId());
  Logger.log('Chunks written: ' + chunks.length);
  Logger.log('Rows: ' + payload.row_count + '  Exported: ' + payload.exported_at);
  return url;
}

/** Optional: schedule auto-refresh every 5 minutes. */
function scheduleSync() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'syncToSheet') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncToSheet').timeBased().everyMinutes(5).create();
  Logger.log('Scheduled syncToSheet every 5 minutes.');
}

/** Web-app endpoint (kept for compatibility, not needed for the new flow). */
function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var tabName = params.tab || DEFAULT_TAB;
  var payload;
  try { payload = buildPayload(tabName); }
  catch (err) { payload = { error: String(err && err.message ? err.message : err) }; }
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
