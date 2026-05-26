/**
 * Planning Inputs + Users Sync API
 * --------------------------------
 * Single Apps Script Web App for:
 *   - Per-line planning inputs (EQPT, HOD, Done, Production Done)
 *   - User accounts + role-based login (User / Admin / Super Admin)
 *
 * Sheets used (auto-created on first run):
 *   PLANNING_INPUTS  - line_id, done, prod_done, eqpt, hod, updated_at, updated_by,
 *                      commit_start, commit_end, commit_by
 *   USERS            - username, password_plain, password_hash, role, created_at, last_login, view_filter
 *
 *   commit_start / commit_end are YYYY-MM-DD strings entered from the drill-down
 *   table by the "user" or "super_admin" role. commit_by stamps the username who
 *   last edited the commitment dates so the frontend can show "already booked by
 *   X" conflicts. The "admin" role can view but cannot edit these dates.
 *
 *   view_filter is a comma-separated list of APPROVE_STATUS values the user
 *   is allowed to see. Only enforced for the "user" role; Admins and Super
 *   Admins see all data regardless. Empty/blank for a "user" = no data view.
 *
 * Password handling:
 *   Both the PLAIN password and its SHA-256 hash are stored side by side.
 *   The login flow still matches against the hash. The Super Admin panel
 *   simply reads back the plain password so resets / sharing are easy.
 *   Trade-off: anyone with the spreadsheet open can read passwords. That
 *   is intentional for this trust-based internal tool.
 *
 * Setup:
 * 1. Container-bound Apps Script (Extensions -> Apps Script in your Sheet),
 *    OR a standalone script where SpreadsheetApp.openById(SHEET_ID) works.
 * 2. Paste this file, Save.
 * 3. Deploy -> Manage deployments -> pencil icon on existing deployment ->
 *    Version: New version -> Deploy. Same URL stays unchanged.
 * 4. First-time bootstrap: an empty USERS sheet auto-seeds a default Super
 *    Admin row -> username: "samrat", password: "samrat123". Log in once,
 *    change the password from the User Control Panel, and add your team.
 */

const INPUT_SHEET_NAME = 'PLANNING_INPUTS';
const USERS_SHEET_NAME = 'USERS';
const DEFAULT_USERNAME = 'samrat';
const DEFAULT_PASSWORD_PLAIN = 'samrat123';

// SHA-256 helper using Apps Script's built-in crypto.
function sha256_(text) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text || ''));
  return raw.map(function(b){ return ('00' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

// ---------- PLANNING_INPUTS sheet ----------
// Canonical column order (1-indexed):
//   1 line_id | 2 done | 3 prod_done | 4 eqpt | 5 hod | 6 updated_at |
//   7 updated_by | 8 commit_start | 9 commit_end | 10 commit_by |
//   11 reschedule_count | 12 reschedule_reason
//
// reschedule_count tracks how many times this line's commit_end has been
// pushed FORWARD via the Reschedule modal. Pure reorder doesn't increment.
// reschedule_reason holds the latest free-text reason entered when Apply
// New Schedule was clicked (mandatory for every reschedule â€" forward push
// or interchange).
var PLANNING_HEADERS = [
  'line_id', 'done', 'prod_done', 'eqpt', 'hod',
  'updated_at', 'updated_by',
  'commit_start', 'commit_end', 'commit_by',
  'reschedule_count', 'reschedule_reason'
];

function getInputSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(INPUT_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(INPUT_SHEET_NAME);
    sh.appendRow(PLANNING_HEADERS.slice());
    return sh;
  }
  // Auto-migrate: append any missing canonical column at its expected index.
  var hdr = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
  PLANNING_HEADERS.forEach(function(name, i){
    if (hdr[i] !== name) sh.getRange(1, i + 1).setValue(name);
  });
  return sh;
}

function readRecords_() {
  var sh = getInputSheet_();
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0].map(String);
  var tz = Session.getScriptTimeZone();
  return values.slice(1)
    .filter(function(r){ return r[0]; })
    .map(function(r){
      var obj = {};
      headers.forEach(function(h, i){
        var v = r[i];
        if ((h === 'commit_start' || h === 'commit_end') &&
            Object.prototype.toString.call(v) === '[object Date]') {
          v = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
        }
        obj[h] = v;
      });
      return obj;
    });
}

function doPlanningUpsert_(p) {
  var lineId = String(p.line_id || '').trim();
  if (!lineId) return { status: 'error', message: 'line_id required' };
  var sh = getInputSheet_();
  var lastRow = sh.getLastRow();
  var targetRow = 0;
  // Pull existing row (if any) so we only overwrite the fields that were
  // actually sent in this POST. This lets the front-end save just the
  // commitment dates (or just done/eqpt/etc) without wiping the other fields.
  var existing = null;
  if (lastRow >= 2) {
    var range = sh.getRange(2, 1, lastRow - 1, PLANNING_HEADERS.length).getValues();
    for (var i = 0; i < range.length; i++) {
      if (String(range[i][0]) === lineId) {
        targetRow = i + 2;
        existing = range[i];
        break;
      }
    }
  }
  if (!targetRow) targetRow = lastRow + 1;

  function pick(paramKey, existingIdx, fallback) {
    if (p[paramKey] != null && String(p[paramKey]) !== '') return p[paramKey];
    if (Object.prototype.hasOwnProperty.call(p, paramKey)) return p[paramKey]; // explicit blank clears
    if (existing) return existing[existingIdx];
    return fallback;
  }

  var doneVal;
  if (Object.prototype.hasOwnProperty.call(p, 'done')) {
    doneVal = String(p.done || '') === '1';
  } else {
    doneVal = existing ? (existing[1] === true || String(existing[1]).toLowerCase() === 'true') : false;
  }

  var prodDone = Object.prototype.hasOwnProperty.call(p, 'prod_done')
    ? (p.prod_done || '')
    : (existing ? existing[2] : '');
  var eqptVal = Object.prototype.hasOwnProperty.call(p, 'eqpt')
    ? (p.eqpt || '')
    : (existing ? existing[3] : '');
  var hodVal = Object.prototype.hasOwnProperty.call(p, 'hod')
    ? (p.hod || '')
    : (existing ? existing[4] : '');
  var updatedAt = p.updated_at || new Date().toISOString();
  var updatedBy = Object.prototype.hasOwnProperty.call(p, 'updated_by')
    ? (p.updated_by || '')
    : (existing ? existing[6] : '');
  var commitStart = Object.prototype.hasOwnProperty.call(p, 'commit_start')
    ? (p.commit_start || '')
    : (existing ? existing[7] : '');
  var commitEnd = Object.prototype.hasOwnProperty.call(p, 'commit_end')
    ? (p.commit_end || '')
    : (existing ? existing[8] : '');
  var commitBy = Object.prototype.hasOwnProperty.call(p, 'commit_by')
    ? (p.commit_by || '')
    : (existing ? existing[9] : '');
  var rescheduleCount;
  if (Object.prototype.hasOwnProperty.call(p, 'reschedule_count')) {
    var rcRaw = parseInt(p.reschedule_count, 10);
    rescheduleCount = isNaN(rcRaw) ? 0 : rcRaw;
  } else {
    rescheduleCount = existing ? (parseInt(existing[10], 10) || 0) : 0;
  }
  var rescheduleReason = Object.prototype.hasOwnProperty.call(p, 'reschedule_reason')
    ? (p.reschedule_reason || '')
    : (existing ? (existing[11] || '') : '');

  sh.getRange(targetRow, 1, 1, PLANNING_HEADERS.length).setValues([[
    lineId, doneVal, prodDone, eqptVal, hodVal,
    updatedAt, updatedBy,
    commitStart, commitEnd, commitBy,
    rescheduleCount, rescheduleReason
  ]]);
  return { status: 'ok', line_id: lineId };
}

// ---------- USERS sheet ----------
// New schema: username | password_plain | password_hash | role | created_at | last_login
// Old schema: username | password_hash | role | created_at | last_login
// On first call after upgrade, we migrate the old schema in place.
function getUsersSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(USERS_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(USERS_SHEET_NAME);
    sh.appendRow(['username', 'password_plain', 'password_hash', 'role', 'created_at', 'last_login', 'view_filter']);
    sh.appendRow([
      DEFAULT_USERNAME,
      DEFAULT_PASSWORD_PLAIN,
      sha256_(DEFAULT_PASSWORD_PLAIN),
      'super_admin',
      new Date().toISOString(),
      '',
      ''
    ]);
    return sh;
  }
  // Migration: if the second column is `password_hash` (old schema), insert a
  // new `password_plain` column between username and password_hash.
  var hdr = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
  if (hdr.length >= 2 && hdr[1] === 'password_hash' && hdr.indexOf('password_plain') < 0) {
    sh.insertColumnBefore(2);
    sh.getRange(1, 2).setValue('password_plain');
  }
  // Ensure view_filter column exists at position 7.
  hdr = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
  if (hdr.indexOf('view_filter') < 0) {
    sh.getRange(1, 7).setValue('view_filter');
  }
  // Ensure all expected headers exist in canonical order.
  hdr = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
  var want = ['username', 'password_plain', 'password_hash', 'role', 'created_at', 'last_login', 'view_filter'];
  want.forEach(function(name, i){
    if (hdr[i] !== name) sh.getRange(1, i + 1).setValue(name);
  });
  return sh;
}

function readUsers_() {
  var sh = getUsersSheet_();
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0].map(String);
  return values.slice(1)
    .filter(function(r){ return r[0]; })
    .map(function(r){
      var obj = {};
      headers.forEach(function(h, i){ obj[h] = r[i]; });
      return obj;
    });
}

function findUserRow_(username) {
  var sh = getUsersSheet_();
  var last = sh.getLastRow();
  if (last < 2) return { sheet: sh, row: 0, user: null };
  var data = sh.getRange(2, 1, last - 1, 7).getValues();
  var needle = String(username || '').toLowerCase().trim();
  if (!needle) return { sheet: sh, row: 0, user: null };
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase().trim() === needle) {
      return {
        sheet: sh, row: i + 2,
        user: {
          username: data[i][0],
          password_plain: data[i][1],
          password_hash: data[i][2],
          role: data[i][3],
          created_at: data[i][4],
          last_login: data[i][5],
          view_filter: data[i][6] || ''
        }
      };
    }
  }
  return { sheet: sh, row: 0, user: null };
}

function authUser_(username, password_hash) {
  if (!username || !password_hash) return { status: 'error', message: 'username and password required' };
  var found = findUserRow_(username);
  if (!found.user) return { status: 'error', message: 'User not found' };
  if (String(found.user.password_hash) !== String(password_hash)) {
    return { status: 'error', message: 'Wrong password' };
  }
  found.sheet.getRange(found.row, 6).setValue(new Date().toISOString());
  return {
    status: 'ok',
    username: String(found.user.username),
    role: String(found.user.role || 'user'),
    view_filter: String(found.user.view_filter || '')
  };
}

// Upsert handles both create (new row) and update. For password updates we
// store BOTH the plain value (column 2) and the SHA-256 hash (column 3) so
// the Super Admin panel can read the plain value back.
function upsertUser_(p) {
  var username = String(p.username || '').trim();
  if (!username) return { status: 'error', message: 'username required' };
  var role = String(p.role || 'user').toLowerCase();
  if (['user', 'admin', 'super_admin'].indexOf(role) < 0) role = 'user';
  var newPlain = (p.password_plain == null) ? '' : String(p.password_plain);
  var newHash = (p.password_hash == null) ? '' : String(p.password_hash);
  // If the caller sent a plain password but not a hash (shouldn't normally
  // happen â€" client should hash), compute the hash server-side.
  if (newPlain && !newHash) newHash = sha256_(newPlain);
  // view_filter is always overwritten (even if empty) on every save, because
  // empty intentionally means "no view access" for users.
  var viewFilter = (p.view_filter == null) ? '' : String(p.view_filter).trim();
  var found = findUserRow_(username);
  if (found.user) {
    if (newPlain) {
      found.sheet.getRange(found.row, 2).setValue(newPlain);
      found.sheet.getRange(found.row, 3).setValue(newHash);
    }
    found.sheet.getRange(found.row, 4).setValue(role);
    found.sheet.getRange(found.row, 7).setValue(viewFilter);
    return { status: 'ok', updated: username };
  }
  if (!newPlain || !newHash) return { status: 'error', message: 'password required for new user' };
  var sh = getUsersSheet_();
  sh.appendRow([username, newPlain, newHash, role, new Date().toISOString(), '', viewFilter]);
  return { status: 'ok', created: username };
}

function deleteUser_(username) {
  if (!username) return { status: 'error', message: 'username required' };
  var found = findUserRow_(username);
  if (!found.user) return { status: 'error', message: 'User not found' };
  var users = readUsers_();
  var supers = users.filter(function(u){ return String(u.role).toLowerCase() === 'super_admin'; });
  if (supers.length <= 1 && String(found.user.role).toLowerCase() === 'super_admin') {
    return { status: 'error', message: 'Cannot delete the last super_admin' };
  }
  found.sheet.deleteRow(found.row);
  return { status: 'ok', deleted: username };
}

// Return ALL fields including password_plain so the Super Admin panel can
// display the original passwords. Hashes are also included but the panel
// doesn't render them.
function listUsersPublic_() {
  return readUsers_().map(function(u){
    return {
      username: u.username,
      password_plain: u.password_plain || '',
      role: u.role,
      created_at: u.created_at,
      last_login: u.last_login,
      view_filter: u.view_filter || ''
    };
  });
}

// ---------- HTTP entrypoints ----------
function doGet(e) {
  var params = e && e.parameter ? e.parameter : {};
  var action = String(params.action || 'list').toLowerCase();
  var payload;
  try {
    if (action === 'auth') {
      payload = authUser_(params.username, params.password_hash);
    } else if (action === 'list_users') {
      payload = { status: 'ok', users: listUsersPublic_() };
    } else {
      payload = {
        status: 'ok',
        records: readRecords_(),
        exported_at: new Date().toISOString()
      };
    }
  } catch (err) {
    payload = { status: 'error', message: String(err) };
  }
  if (params.callback) {
    return ContentService
      .createTextOutput(params.callback + '(' + JSON.stringify(payload) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var p = e && e.parameter ? e.parameter : {};
  var action = String(p.action || 'upsert').toLowerCase();
  var payload;
  try {
    if (action === 'upsert_user') {
      payload = upsertUser_(p);
    } else if (action === 'delete_user') {
      payload = deleteUser_(String(p.username || ''));
    } else {
      payload = doPlanningUpsert_(p);
    }
  } catch (err) {
    payload = { status: 'error', message: String(err) };
  }
  return json_(payload);
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
