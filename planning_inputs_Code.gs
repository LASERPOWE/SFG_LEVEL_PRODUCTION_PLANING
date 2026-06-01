/**
 * Planning Inputs + Users Sync API
 * Version: 2026-05-27 duplicate-header cleanup + consume_from
 * --------------------------------
 * Single Apps Script Web App for:
 *   - Per-line planning inputs (EQPT, HOD, Done, Production Done)
 *   - User accounts + role-based login (User / Admin / Super Admin)
 *
 * Sheets used (auto-created on first run):
 *   PLANNING_INPUTS  - line_id, done, prod_done, eqpt, hod, consume_from,
 *                      updated_at, updated_by, commit_start, commit_end, commit_by
 *   USERS            - username, password_plain, password_hash, role, created_at, last_login, view_filter, input_rights
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
//   1 line_id | 2 done | 3 prod_done | 4 eqpt | 5 hod | 6 consume_from |
//   7 updated_at | 8 updated_by | 9 commit_start | 10 commit_end |
//   11 commit_by | 12 reschedule_count | 13 reschedule_reason
//
// reschedule_count tracks how many times this line's commit_end has been
// pushed FORWARD via the Reschedule modal. Pure reorder doesn't increment.
// reschedule_reason holds the latest free-text reason entered when Apply
// New Schedule was clicked (mandatory for every reschedule â€" forward push
// or interchange).
var PLANNING_HEADERS = [
  'line_id', 'done', 'prod_done', 'eqpt', 'hod', 'consume_from',
  'updated_at', 'updated_by',
  'commit_start', 'commit_end', 'commit_by',
  'reschedule_count', 'reschedule_reason'
];

function isConsumeFromValue_(v) {
  var s = String(v == null ? '' : v).trim();
  return !s || s === 'KHP' || s === 'PP-U1' || s === 'PP-U2';
}

function looksIsoDateOnly_(v) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(v == null ? '' : v).trim());
}

function looksIsoDateTime_(v) {
  return /^\d{4}-\d{2}-\d{2}T/.test(String(v == null ? '' : v).trim());
}

function maybeRepairShiftedPlanningData_(sh) {
  var lastRow = sh.getLastRow();
  var lastCol = Math.max(sh.getLastColumn(), PLANNING_HEADERS.length);
  if (lastRow < 2) return;
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  var canonical = PLANNING_HEADERS.every(function(name, i){ return hdr[i] === name; });
  if (!canonical) return;
  var sampleCount = Math.min(lastRow - 1, 8);
  if (sampleCount <= 0) return;
  var rows = sh.getRange(2, 1, sampleCount, PLANNING_HEADERS.length).getValues();
  var shiftedVotes = 0;
  rows.forEach(function(r){
    var consume = r[5];
    var updatedAt = r[6];
    var updatedBy = r[7];
    var commitStart = r[8];
    var commitEnd = r[9];
    if (!consume && !updatedAt && !updatedBy && !commitStart && !commitEnd) return;
    if (!isConsumeFromValue_(consume) && (looksIsoDateTime_(consume) || looksIsoDateOnly_(updatedBy))) shiftedVotes++;
  });
  if (!shiftedVotes) return;
  var all = sh.getRange(2, 1, lastRow - 1, PLANNING_HEADERS.length).getValues();
  var fixed = all.map(function(r){
    var out = r.slice();
    for (var i = PLANNING_HEADERS.length - 1; i >= 6; i--) out[i] = out[i - 1];
    out[5] = '';
    return out;
  });
  sh.getRange(2, 1, fixed.length, PLANNING_HEADERS.length).setValues(fixed);
}

function normalizePlanningHeaderName_(name) {
  var s = String(name == null ? '' : name).trim();
  if (PLANNING_HEADERS.indexOf(s) >= 0) return s;
  if (s.indexOf('reschedule_reas') === 0) return 'reschedule_reason';
  if (s.indexOf('reschedule_cou') === 0) return 'reschedule_count';
  if (s.indexOf('commit_start') === 0) return 'commit_start';
  if (s.indexOf('commit_end') === 0) return 'commit_end';
  if (s.indexOf('commit_by') === 0) return 'commit_by';
  if (s.indexOf('updated_at') === 0) return 'updated_at';
  if (s.indexOf('updated_by') === 0) return 'updated_by';
  if (s.indexOf('consume_from') === 0) return 'consume_from';
  return '';
}

function isBlankCell_(v) {
  return String(v == null ? '' : v).trim() === '';
}

function recoverExtraPlanningData_(sh) {
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol <= PLANNING_HEADERS.length) return;
  var data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var changed = false;
  data.forEach(function(row){
    var extra = row.slice(PLANNING_HEADERS.length);
    for (var j = 0; j < extra.length; j++) {
      var v = extra[j];
      var next = extra[j + 1];
      var next2 = extra[j + 2];
      if (isBlankCell_(v)) continue;
      if (isBlankCell_(row[6]) && looksIsoDateTime_(v)) {
        row[6] = v;
        changed = true;
        if (isBlankCell_(row[7]) && !isBlankCell_(next) && !looksIsoDateOnly_(next) && !looksIsoDateTime_(next)) {
          row[7] = next;
          changed = true;
        }
        continue;
      }
      if (looksIsoDateOnly_(v)) {
        if (isBlankCell_(row[8])) {
          row[8] = v;
          changed = true;
          if (!isBlankCell_(next) && looksIsoDateOnly_(next) && isBlankCell_(row[9])) {
            row[9] = next;
            changed = true;
            if (!isBlankCell_(next2) && !looksIsoDateOnly_(next2) && !looksIsoDateTime_(next2) && isBlankCell_(row[10])) {
              row[10] = next2;
              changed = true;
            }
          }
        } else if (isBlankCell_(row[9])) {
          row[9] = v;
          changed = true;
        }
      }
    }
  });
  if (changed) {
    var fixed = data.map(function(row){ return row.slice(0, PLANNING_HEADERS.length); });
    sh.getRange(2, 1, fixed.length, PLANNING_HEADERS.length).setValues(fixed);
  }
}

function cleanupDuplicatePlanningHeaderColumns_(sh) {
  var lastCol = sh.getLastColumn();
  var lastRow = sh.getLastRow();
  if (lastCol <= PLANNING_HEADERS.length) return;
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  for (var c = lastCol; c > PLANNING_HEADERS.length; c--) {
    var name = normalizePlanningHeaderName_(hdr[c - 1]);
    if (!name) continue;
    var targetCol = PLANNING_HEADERS.indexOf(name) + 1;
    if (targetCol <= 0 || targetCol === c) continue;
    if (lastRow >= 2) {
      var dupVals = sh.getRange(2, c, lastRow - 1, 1).getValues();
      var targetVals = sh.getRange(2, targetCol, lastRow - 1, 1).getValues();
      var changed = false;
      for (var i = 0; i < dupVals.length; i++) {
        var dup = dupVals[i][0];
        var cur = targetVals[i][0];
        if (String(cur == null ? '' : cur).trim() === '' &&
            String(dup == null ? '' : dup).trim() !== '') {
          targetVals[i][0] = dup;
          changed = true;
        }
      }
      if (changed) sh.getRange(2, targetCol, targetVals.length, 1).setValues(targetVals);
    }
    sh.deleteColumn(c);
  }
}

function getInputSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(INPUT_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(INPUT_SHEET_NAME);
    sh.appendRow(PLANNING_HEADERS.slice());
    return sh;
  }
  // Safe migration from the old 12-column schema:
  // insert consume_from before updated_at so existing commit dates do not shift.
  var hdr = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
  if (hdr[0] === 'line_id' && hdr[1] === 'done' && hdr[2] === 'prod_done' &&
      hdr[3] === 'eqpt' && hdr[4] === 'hod' && hdr[5] === 'updated_at' &&
      hdr.indexOf('consume_from') < 0) {
    sh.insertColumnBefore(6);
  }
  // Auto-migrate / normalize headers to the canonical order.
  hdr = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), PLANNING_HEADERS.length)).getValues()[0].map(String);
  PLANNING_HEADERS.forEach(function(name, i){
    if (hdr[i] !== name) sh.getRange(1, i + 1).setValue(name);
  });
  maybeRepairShiftedPlanningData_(sh);
  recoverExtraPlanningData_(sh);
  cleanupDuplicatePlanningHeaderColumns_(sh);
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
  var consumeFromVal = Object.prototype.hasOwnProperty.call(p, 'consume_from')
    ? (p.consume_from || '')
    : (existing ? existing[5] : '');
  var updatedAt = p.updated_at || new Date().toISOString();
  var updatedBy = Object.prototype.hasOwnProperty.call(p, 'updated_by')
    ? (p.updated_by || '')
    : (existing ? existing[7] : '');
  var commitStart = Object.prototype.hasOwnProperty.call(p, 'commit_start')
    ? (p.commit_start || '')
    : (existing ? existing[8] : '');
  var commitEnd = Object.prototype.hasOwnProperty.call(p, 'commit_end')
    ? (p.commit_end || '')
    : (existing ? existing[9] : '');
  var commitBy = Object.prototype.hasOwnProperty.call(p, 'commit_by')
    ? (p.commit_by || '')
    : (existing ? existing[10] : '');
  var rescheduleCount;
  if (Object.prototype.hasOwnProperty.call(p, 'reschedule_count')) {
    var rcRaw = parseInt(p.reschedule_count, 10);
    rescheduleCount = isNaN(rcRaw) ? 0 : rcRaw;
  } else {
    rescheduleCount = existing ? (parseInt(existing[11], 10) || 0) : 0;
  }
  var rescheduleReason = Object.prototype.hasOwnProperty.call(p, 'reschedule_reason')
    ? (p.reschedule_reason || '')
    : (existing ? (existing[12] || '') : '');

  sh.getRange(targetRow, 1, 1, PLANNING_HEADERS.length).setValues([[
    lineId, doneVal, prodDone, eqptVal, hodVal, consumeFromVal,
    updatedAt, updatedBy,
    commitStart, commitEnd, commitBy,
    rescheduleCount, rescheduleReason
  ]]);
  return { status: 'ok', line_id: lineId };
}

// ---------- USERS sheet ----------
// New schema: username | password_plain | password_hash | role | created_at | last_login | view_filter | input_rights
// Old schema: username | password_hash | role | created_at | last_login
// On first call after upgrade, we migrate the old schema in place.
function defaultInputRights_(role) {
  role = String(role || 'user').toLowerCase();
  var flat;
  if (role === 'super_admin') flat = { view: true, entry: true, edit: true, delete: true };
  else if (role === 'admin') flat = { view: true, entry: true, edit: true, delete: false };
  else flat = { view: true, entry: true, edit: false, delete: false };
  return makeInputRightsFromFlat_(flat);
}

function inputRightColumns_() {
  return ['done', 'prod_done', 'eqpt', 'hod', 'commit_start', 'commit_end', 'consume_from'];
}

function makeInputRightsFromFlat_(flat) {
  var out = { columns: {} };
  inputRightColumns_().forEach(function(k){
    out.columns[k] = {
      view: flat.view !== false,
      entry: !!flat.entry,
      edit: !!flat.edit,
      delete: !!flat.delete
    };
  });
  return out;
}

function normalizeInputRights_(raw, role) {
  var obj = null;
  if (raw && typeof raw === 'object') obj = raw;
  else if (raw) {
    try { obj = JSON.parse(String(raw)); } catch(e) { obj = null; }
  }
  if (!obj) obj = defaultInputRights_(role);
  function truthy(v, def) {
    if (v == null) return !!def;
    return v === true || v === 'true' || v === 1 || v === '1';
  }
  if (!obj.columns) {
    obj = makeInputRightsFromFlat_({
      view: obj.view !== false && obj.view !== 'false' && obj.view !== 0 && obj.view !== '0',
      entry: truthy(obj.entry, false),
      edit: truthy(obj.edit, false),
      delete: truthy(obj.delete, false)
    });
  }
  var base = defaultInputRights_(role);
  var out = { columns: {} };
  inputRightColumns_().forEach(function(k){
    var src = obj.columns && obj.columns[k] ? obj.columns[k] : null;
    var fallback = base.columns[k];
    out.columns[k] = {
      view: src ? (src.view !== false && src.view !== 'false' && src.view !== 0 && src.view !== '0') : fallback.view,
      entry: src ? truthy(src.entry, false) : fallback.entry,
      edit: src ? truthy(src.edit, false) : fallback.edit,
      delete: src ? truthy(src.delete, false) : fallback.delete
    };
  });
  return JSON.stringify(out);
}

function getUsersSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(USERS_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(USERS_SHEET_NAME);
    sh.appendRow(['username', 'password_plain', 'password_hash', 'role', 'created_at', 'last_login', 'view_filter', 'input_rights']);
    sh.appendRow([
      DEFAULT_USERNAME,
      DEFAULT_PASSWORD_PLAIN,
      sha256_(DEFAULT_PASSWORD_PLAIN),
      'super_admin',
      new Date().toISOString(),
      '',
      '',
      JSON.stringify(defaultInputRights_('super_admin'))
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
  // Ensure input_rights column exists at position 8.
  hdr = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
  if (hdr.indexOf('input_rights') < 0) {
    sh.getRange(1, 8).setValue('input_rights');
  }
  // Ensure all expected headers exist in canonical order.
  hdr = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
  var want = ['username', 'password_plain', 'password_hash', 'role', 'created_at', 'last_login', 'view_filter', 'input_rights'];
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
  var data = sh.getRange(2, 1, last - 1, 8).getValues();
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
          view_filter: data[i][6] || '',
          input_rights: data[i][7] || ''
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
    view_filter: String(found.user.view_filter || ''),
    input_rights: normalizeInputRights_(found.user.input_rights, found.user.role)
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
  var inputRights = normalizeInputRights_(p.input_rights, role);
  var found = findUserRow_(username);
  if (found.user) {
    if (newPlain) {
      found.sheet.getRange(found.row, 2).setValue(newPlain);
      found.sheet.getRange(found.row, 3).setValue(newHash);
    }
    found.sheet.getRange(found.row, 4).setValue(role);
    found.sheet.getRange(found.row, 7).setValue(viewFilter);
    found.sheet.getRange(found.row, 8).setValue(inputRights);
    return { status: 'ok', updated: username };
  }
  if (!newPlain || !newHash) return { status: 'error', message: 'password required for new user' };
  var sh = getUsersSheet_();
  sh.appendRow([username, newPlain, newHash, role, new Date().toISOString(), '', viewFilter, inputRights]);
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
      view_filter: u.view_filter || '',
      input_rights: normalizeInputRights_(u.input_rights, u.role)
    };
  });
}

// ---------- MANPOWER_ALLOCATIONS sheet ----------
// One row per (date, equipment, emp_code) allocation. Conflict rule
// enforced at upsert time: an emp_code can only sit on ONE equipment per
// date â€" attempting to allocate the same emp_code to a different
// equipment on the same date returns a "conflict" error so the front-end
// can show "already allocated to X" inline.
var ALLOC_SHEET_NAME = 'MANPOWER_ALLOCATIONS';
var ALLOC_HEADERS = [
  'id',                // composite key: date|equipment|emp_code|shift
  'date',              // YYYY-MM-DD
  'equipment',
  'booking_line_key',  // line_id of the Reschedule booking that triggered this allocation
  'emp_code',
  'emp_name',
  'from_team',         // team currently allocated from (can differ from original_team)
  'original_team',     // team this employee belongs to per TEAM MASTER_MANPOWER
  'allocated_at',
  'allocated_by',
  'rate',              // member day-rate captured AT ALLOCATION TIME (so cost is historically accurate even if Team Master rate changes later)
  'shift'              // 'day_shift' (8am-8pm) | 'night_shift' (8pm-8am) | '' (legacy rows)
];

function getAllocSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ALLOC_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(ALLOC_SHEET_NAME);
    sh.appendRow(ALLOC_HEADERS.slice());
    return sh;
  }
  var hdr = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
  ALLOC_HEADERS.forEach(function(name, i){
    if (hdr[i] !== name) sh.getRange(1, i + 1).setValue(name);
  });
  return sh;
}

function _allocId_(date, equipment, empCode, shift) {
  // Shift is part of the primary key so the same employee can legitimately
  // hold separate Day and Night allocations on the same date + equipment
  // (relevant during peak hand-offs) without overwriting each other.
  return String(date || '').trim() + '|' +
         String(equipment || '').trim() + '|' +
         String(empCode || '').trim() + '|' +
         String(shift || '').trim();
}

function readAllocations_() {
  var sh = getAllocSheet_();
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0].map(String);
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (!r[0]) continue;
    var obj = {};
    headers.forEach(function(h, idx){ obj[h] = r[idx]; });
    // Date normalization: Sheets coerces YYYY-MM-DD into Date when shown.
    if (obj.date) {
      if (obj.date instanceof Date) {
        var d = obj.date;
        obj.date = d.getFullYear() + '-' +
                   String(d.getMonth() + 1).padStart(2, '0') + '-' +
                   String(d.getDate()).padStart(2, '0');
      } else {
        var s = String(obj.date).trim();
        var m = s.match(/^(\\d{4})-(\\d{2})-(\\d{2})/);
        if (m) obj.date = m[1] + '-' + m[2] + '-' + m[3];
      }
    }
    // Rate normalization: stored as number; frontend expects num.
    if (obj.rate != null && obj.rate !== '') {
      var rr = parseFloat(obj.rate);
      obj.rate = isNaN(rr) ? 0 : rr;
    } else {
      obj.rate = 0;
    }
    out.push(obj);
  }
  return out;
}

function listAllocations_(date) {
  var rows = readAllocations_();
  if (date) {
    var d = String(date).trim();
    rows = rows.filter(function(r){ return String(r.date || '').trim() === d; });
  }
  return { status: 'ok', allocations: rows };
}

function _findAllocRow_(sh, date, equipment, empCode, shift) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var data = sh.getRange(2, 1, last - 1, ALLOC_HEADERS.length).getValues();
  var wantId = _allocId_(date, equipment, empCode, shift);
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0] || '').trim() === wantId) return i + 2;
  }
  return 0;
}

function _findAllocConflict_(sh, date, equipment, empCode, shift) {
  // Conflict = same date + same emp + same shift on a DIFFERENT equipment.
  // Different shifts (Day vs Night) on different equipment for the same
  // employee are allowed (a hand-off scenario) â€" so we require shifts to
  // match before flagging a clash. Legacy rows without a shift value are
  // treated as 'any' to preserve old behaviour for un-migrated data.
  var last = sh.getLastRow();
  if (last < 2) return null;
  var data = sh.getRange(2, 1, last - 1, ALLOC_HEADERS.length).getValues();
  var d = String(date || '').trim();
  var eq = String(equipment || '').trim();
  var emp = String(empCode || '').trim();
  var sh_ = String(shift || '').trim();
  var shiftIdx = ALLOC_HEADERS.indexOf('shift');
  for (var i = 0; i < data.length; i++) {
    var rowDate = data[i][1];
    if (rowDate instanceof Date) {
      rowDate = rowDate.getFullYear() + '-' +
                String(rowDate.getMonth() + 1).padStart(2, '0') + '-' +
                String(rowDate.getDate()).padStart(2, '0');
    } else {
      rowDate = String(rowDate || '').trim().slice(0, 10);
    }
    var rowEq = String(data[i][2] || '').trim();
    var rowEmp = String(data[i][4] || '').trim();
    var rowShift = shiftIdx >= 0 ? String(data[i][shiftIdx] || '').trim() : '';
    if (rowDate !== d || rowEmp !== emp || rowEq === eq) continue;
    // Shift-aware clash: only flag when both rows declare the same shift
    // OR when either side is blank (legacy rows still block to be safe).
    if (sh_ && rowShift && sh_ !== rowShift) continue;
    return {
      id: data[i][0],
      date: rowDate,
      equipment: rowEq,
      booking_line_key: data[i][3],
      emp_code: rowEmp,
      emp_name: data[i][5],
      from_team: data[i][6],
      original_team: data[i][7],
      shift: rowShift
    };
  }
  return null;
}

function upsertAllocation_(p) {
  var date = String(p.date || '').trim();
  var equipment = String(p.equipment || '').trim();
  var empCode = String(p.emp_code || '').trim();
  var shift = String(p.shift || '').trim();
  if (!date || !equipment || !empCode) {
    return { status: 'error', message: 'date, equipment and emp_code are required' };
  }
  var sh = getAllocSheet_();
  var conflict = _findAllocConflict_(sh, date, equipment, empCode, shift);
  if (conflict) {
    var shiftLabel = conflict.shift ? ' (' + conflict.shift + ')' : '';
    return { status: 'conflict', message: 'Already allocated to ' + conflict.equipment + shiftLabel, conflict: conflict };
  }
  var row = _findAllocRow_(sh, date, equipment, empCode, shift);
  if (!row) row = sh.getLastRow() + 1;
  var id = _allocId_(date, equipment, empCode, shift);
  var rateVal = 0;
  if (p.rate != null && p.rate !== '') {
    rateVal = parseFloat(p.rate);
    if (isNaN(rateVal)) rateVal = 0;
  }
  var values = [
    id, date, equipment,
    String(p.booking_line_key || ''),
    empCode,
    String(p.emp_name || ''),
    String(p.from_team || ''),
    String(p.original_team || ''),
    p.allocated_at || new Date().toISOString(),
    String(p.allocated_by || ''),
    rateVal,
    shift
  ];
  sh.getRange(row, 1, 1, ALLOC_HEADERS.length).setValues([values]);
  return { status: 'ok', id: id, row_number: row };
}

function deleteAllocation_(p) {
  var date = String(p.date || '').trim();
  var equipment = String(p.equipment || '').trim();
  var empCode = String(p.emp_code || '').trim();
  var shift = String(p.shift || '').trim();
  if (!date || !equipment || !empCode) {
    return { status: 'error', message: 'date, equipment and emp_code are required' };
  }
  var sh = getAllocSheet_();
  var row = _findAllocRow_(sh, date, equipment, empCode, shift);
  if (!row) return { status: 'ok', message: 'No matching allocation; nothing to delete' };
  sh.deleteRow(row);
  return { status: 'ok', deleted: _allocId_(date, equipment, empCode, shift) };
}

// Bulk replace: wipe all allocations for (date, equipment) and re-insert
// the supplied list. items[] each carry emp_code, emp_name, from_team,
// original_team, rate.
function bulkSetAllocations_(p) {
  var date = String(p.date || '').trim();
  var equipment = String(p.equipment || '').trim();
  var shift = String(p.shift || '').trim();
  if (!date || !equipment) {
    return { status: 'error', message: 'date and equipment are required' };
  }
  var items = [];
  try { items = JSON.parse(String(p.items || '[]')); } catch(e) { items = []; }
  if (!Array.isArray(items)) items = [];
  var sh = getAllocSheet_();
  var shiftIdx = ALLOC_HEADERS.indexOf('shift');
  var last = sh.getLastRow();
  if (last >= 2) {
    var data = sh.getRange(2, 1, last - 1, ALLOC_HEADERS.length).getValues();
    for (var i = data.length - 1; i >= 0; i--) {
      var rd = data[i][1];
      if (rd instanceof Date) {
        rd = rd.getFullYear() + '-' +
             String(rd.getMonth() + 1).padStart(2, '0') + '-' +
             String(rd.getDate()).padStart(2, '0');
      } else {
        rd = String(rd || '').trim().slice(0, 10);
      }
      var req = String(data[i][2] || '').trim();
      var rowShift = (shiftIdx >= 0) ? String(data[i][shiftIdx] || '').trim() : '';
      // Wipe only (date + equipment + shift) â€" leaves the other shift's
      // allocations untouched so Day and Night can be saved independently.
      if (rd === date && req === equipment && rowShift === shift) {
        sh.deleteRow(i + 2);
      }
    }
  }
  var conflicts = [];
  var inserted = [];
  for (var k = 0; k < items.length; k++) {
    var it = items[k] || {};
    var ec = String(it.emp_code || '').trim();
    if (!ec) continue;
    var itShift = String(it.shift || shift || '').trim();
    var conf = _findAllocConflict_(sh, date, equipment, ec, itShift);
    if (conf) {
      conflicts.push({ emp_code: ec, emp_name: it.emp_name || '', conflict: conf });
      continue;
    }
    var id = _allocId_(date, equipment, ec, itShift);
    var row = sh.getLastRow() + 1;
    var itRate = 0;
    if (it.rate != null && it.rate !== '') {
      itRate = parseFloat(it.rate);
      if (isNaN(itRate)) itRate = 0;
    }
    sh.getRange(row, 1, 1, ALLOC_HEADERS.length).setValues([[
      id, date, equipment,
      String(p.booking_line_key || it.booking_line_key || ''),
      ec,
      String(it.emp_name || ''),
      String(it.from_team || ''),
      String(it.original_team || ''),
      new Date().toISOString(),
      String(p.allocated_by || ''),
      itRate,
      itShift
    ]]);
    inserted.push(id);
  }
  return {
    status: conflicts.length ? 'partial' : 'ok',
    inserted: inserted.length,
    conflicts: conflicts
  };
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
    } else if (action === 'upsert_user') {
      payload = upsertUser_(params);
    } else if (action === 'delete_user') {
      payload = deleteUser_(String(params.username || ''));
    } else if (action === 'list_allocations') {
      payload = listAllocations_(params.date);
    } else if (action === 'upsert_allocation') {
      payload = upsertAllocation_(params);
    } else if (action === 'delete_allocation') {
      payload = deleteAllocation_(params);
    } else if (action === 'bulk_set_allocations') {
      payload = bulkSetAllocations_(params);
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
    } else if (action === 'upsert_allocation') {
      payload = upsertAllocation_(p);
    } else if (action === 'delete_allocation') {
      payload = deleteAllocation_(p);
    } else if (action === 'bulk_set_allocations') {
      payload = bulkSetAllocations_(p);
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
