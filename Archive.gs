// ── Booking archive & retention ──────────────────────────────────────────
// Put this in its own script file (Archive.gs) alongside Code.gs. Apps Script
// shares one global scope across files, so everything here can use
// readRecord()/writeRecord() from Code.gs and vice versa.
//
// What it does:
//   archiveOldBookings()   monthly trigger (1st, ~2am) — copies bookings older
//                          than ARCHIVE_AFTER_MONTHS into the archive spreadsheet,
//                          checks they arrived, THEN removes them from the live sheet.
//   purgeExpiredRecords()  monthly trigger (2nd, ~2am) — deletes a client's whole
//                          archived history once their retention period has passed.
//   searchArchive_()       backs action=searchArchive — searches archive + live.
//
// Retention rules (per client, counted from their LAST session, live or archived):
//   Adult:  6 years after last session.
//   Child (under 18 at their first session, DOB listed on the Children tab):
//           whichever is LATER of their 25th birthday or 10 years after last session.
//
// One-off setup: run setupArchive() from the editor (authorises the new
// permissions, creates the archive spreadsheet and the two monthly triggers).

const ARCHIVE_AFTER_MONTHS = 3;
const RETENTION_ADULT_YEARS = 6;
const RETENTION_CHILD_YEARS_AFTER_END = 10;
const RETENTION_CHILD_UNTIL_AGE = 25;
const RETENTION_ADULT_AGE = 18;
const ARCHIVE_TZ = 'Europe/London';
const ARCHIVE_HEADERS = ['Key', 'Type', 'Date', 'Start', 'End', 'Practitioner', 'Client',
  'Client 1', 'Client 2', 'Category', 'Payment status', 'Amount', 'Archived at', 'DNA'];
// 'DNA' was added as a 14th column at the end so rows archived before it
// existed still line up — on those older rows the column is simply blank.
const ARCHIVE_COLS = ARCHIVE_HEADERS.length;
const ARCHIVE_CHILDREN_SHEET = 'Children';
const ARCHIVE_LOG_SHEET = 'Activity log';

// ── Setup (run once from the editor) ──────────────────────────────────────
function setupArchive() {
  const props = PropertiesService.getScriptProperties();
  var id = props.getProperty('ARCHIVE_SS_ID');
  var ss;
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    ss = SpreadsheetApp.create('Room Booking Archive');
    props.setProperty('ARCHIVE_SS_ID', ss.getId());
  }
  ss.setSpreadsheetLocale('en_GB');        // so 05/03/2010 means 5 March
  ss.setSpreadsheetTimeZone(ARCHIVE_TZ);

  if (!ss.getSheetByName(ARCHIVE_CHILDREN_SHEET)) {
    const ch = ss.insertSheet(ARCHIVE_CHILDREN_SHEET);
    ch.getRange(1, 1, 1, 3)
      .setValues([['Client name (exactly as written in bookings)', 'Date of birth (dd/mm/yyyy)', 'Notes']])
      .setFontWeight('bold');
    ch.setFrozenRows(1);
    ch.setColumnWidth(1, 280);
    ch.setColumnWidth(2, 180);
  }
  if (!ss.getSheetByName(ARCHIVE_LOG_SHEET)) {
    const lg = ss.insertSheet(ARCHIVE_LOG_SHEET);
    lg.getRange(1, 1, 1, 5)
      .setValues([['Run at', 'Job', 'Bookings archived', 'Bookings purged', 'Clients purged']])
      .setFontWeight('bold');
    lg.setFrozenRows(1);
  }
  const blank = ss.getSheetByName('Sheet1');
  if (blank && ss.getSheets().length > 1 && blank.getLastRow() === 0) ss.deleteSheet(blank);

  ensureMonthlyTrigger_('archiveOldBookings', 1);
  ensureMonthlyTrigger_('purgeExpiredRecords', 2);
  Logger.log('Archive ready: ' + ss.getUrl());
}

function ensureMonthlyTrigger_(fnName, dayOfMonth) {
  const exists = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === fnName;
  });
  if (exists) { Logger.log('Trigger already exists: ' + fnName); return; }
  ScriptApp.newTrigger(fnName).timeBased()
    .onMonthDay(dayOfMonth).atHour(2).inTimezone(ARCHIVE_TZ).create();
  Logger.log('Trigger created: ' + fnName + ' (day ' + dayOfMonth + ', ~2am)');
}

// ── Helpers ───────────────────────────────────────────────────────────────
function getArchiveSs_() {
  const id = PropertiesService.getScriptProperties().getProperty('ARCHIVE_SS_ID');
  if (!id) throw new Error('Archive not set up yet — run setupArchive() once from the Apps Script editor');
  return SpreadsheetApp.openById(id);
}

function archiveYearSheets_(ss) {
  return ss.getSheets().filter(function (s) { return /^\d{4}$/.test(s.getName()); });
}

function getArchiveYearSheet_(ss, year) {
  var sh = ss.getSheetByName(year);
  if (sh) {
    // Year sheets created before the DNA column existed: add its heading once
    const hdr = sh.getRange(1, ARCHIVE_COLS);
    if (!hdr.getDisplayValue()) hdr.setValue(ARCHIVE_HEADERS[ARCHIVE_COLS - 1]).setFontWeight('bold');
    return sh;
  }
  sh = ss.insertSheet(year);
  sh.getRange(1, 1, 1, ARCHIVE_COLS).setValues([ARCHIVE_HEADERS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.getRange(1, 1, sh.getMaxRows(), ARCHIVE_COLS).setNumberFormat('@'); // keep dates/times as plain text
  return sh;
}

// All data rows of a year sheet as plain strings (no Date objects).
function readArchiveSheet_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, ARCHIVE_COLS).getDisplayValues();
}

function archivedKeys_(ss) {
  const keys = {};
  archiveYearSheets_(ss).forEach(function (sh) {
    const last = sh.getLastRow();
    if (last < 2) return;
    sh.getRange(2, 1, last - 1, 1).getDisplayValues().forEach(function (r) {
      if (r[0]) keys[r[0]] = true;
    });
  });
  return keys;
}

function archiveKey_(type, b) { return type + '-' + b.id; }

function archiveRowFromBooking_(b, type, archivedAt) {
  return [
    archiveKey_(type, b),
    type === 'Z' ? 'Zoom' : 'Room',
    String(b.date || ''),
    String(b.start || ''),
    String(b.end || ''),
    String(b.who || ''),
    String(b.note || ''),
    String(b.client1 || ''),
    String(b.client2 || ''),
    String(b.category || ''),
    String(b.paymentStatus || ''),
    (b.expectedAmount !== undefined && b.expectedAmount !== null) ? String(b.expectedAmount) : '',
    archivedAt,
    b.dna ? 'DNA' : ''
  ];
}

function archiveRowToObj_(v, sheetName) {
  return {
    sheet: sheetName, key: v[0], type: v[1], date: v[2], start: v[3], end: v[4],
    who: v[5], client: v[6], client1: v[7], client2: v[8], category: v[9],
    paymentStatus: v[10], amount: v[11],
    dna: String(v[13] || '').toUpperCase() === 'DNA'
  };
}

function normName_(s) {
  return String(s || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

// The individual clients a booking belongs to (couples count as two people).
function bookingClients_(r) {
  var list = [r.client1, r.client2].filter(function (x) { return x && String(x).trim(); });
  if (!list.length && r.client) list.push(r.client);
  return list.map(normName_).filter(Boolean);
}

function addYears_(ymd, n) {
  return String(parseInt(ymd.slice(0, 4), 10) + n) + ymd.slice(4);
}

function todayYmd_() {
  return Utilities.formatDate(new Date(), ARCHIVE_TZ, 'yyyy-MM-dd');
}

// Accepts a Date (typed into a UK-locale sheet), 'yyyy-mm-dd' or 'dd/mm/yyyy'.
function parseDob_(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, ARCHIVE_TZ, 'yyyy-MM-dd');
  const s = String(v || '').trim();
  var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return '';
}

function logArchiveActivity_(ss, job, archived, purgedRows, purgedClients) {
  const lg = ss.getSheetByName(ARCHIVE_LOG_SHEET);
  if (!lg) return;
  lg.appendRow([Utilities.formatDate(new Date(), ARCHIVE_TZ, 'yyyy-MM-dd HH:mm'), job,
    archived, purgedRows, purgedClients]);
}

// ── Monthly archive ───────────────────────────────────────────────────────
function archiveOldBookings() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var record = readRecord();
    const cutoffDate = new Date();
    cutoffDate.setMonth(cutoffDate.getMonth() - ARCHIVE_AFTER_MONTHS);
    const cutoff = Utilities.formatDate(cutoffDate, ARCHIVE_TZ, 'yyyy-MM-dd');

    const old = [];
    var skippedNoId = 0;
    [['R', record.bookings || []], ['Z', record.zoomBookings || []]].forEach(function (pair) {
      pair[1].forEach(function (b) {
        if (!b.date || b.date >= cutoff) return;
        if (b.id === undefined || b.id === null || b.id === '') { skippedNoId++; return; }
        old.push({ type: pair[0], b: b });
      });
    });
    if (skippedNoId) Logger.log('archiveOldBookings: ' + skippedNoId + ' old bookings have no id — left in live sheet');
    if (!old.length) { Logger.log('archiveOldBookings: nothing older than ' + cutoff); return; }

    const ss = getArchiveSs_();
    const existing = archivedKeys_(ss);
    const archivedAt = new Date().toISOString();
    const byYear = {};
    var added = 0;
    old.forEach(function (o) {
      const key = archiveKey_(o.type, o.b);
      if (existing[key]) return;         // already archived (e.g. re-added by a stale save)
      existing[key] = true;
      const y = o.b.date.slice(0, 4);
      (byYear[y] = byYear[y] || []).push(archiveRowFromBooking_(o.b, o.type, archivedAt));
      added++;
    });

    Object.keys(byYear).sort().forEach(function (y) {
      const rows = byYear[y];
      const sh = getArchiveYearSheet_(ss, y);
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, ARCHIVE_COLS)
        .setNumberFormat('@').setValues(rows);
    });
    SpreadsheetApp.flush();

    // Verify every booking is really in the archive before removing anything.
    const after = archivedKeys_(ss);
    const missing = old.filter(function (o) { return !after[archiveKey_(o.type, o.b)]; });
    if (missing.length) {
      throw new Error('Archive check failed: ' + missing.length +
        ' bookings not found in the archive — live data left untouched');
    }

    const moved = {};
    old.forEach(function (o) { moved[archiveKey_(o.type, o.b)] = true; });
    if (record.bookings) record.bookings = record.bookings.filter(function (b) { return !moved['R-' + b.id]; });
    if (record.zoomBookings) record.zoomBookings = record.zoomBookings.filter(function (b) { return !moved['Z-' + b.id]; });
    writeRecord(record);

    logArchiveActivity_(ss, 'Archive', added, 0, 0);
    Logger.log('archiveOldBookings: ' + old.length + ' bookings before ' + cutoff +
      ' moved out of live sheet (' + added + ' newly added to archive)');
  } finally {
    lock.releaseLock();
  }
}

// ── Retention / purge ─────────────────────────────────────────────────────
// Pure decision logic (no Sheets calls) so it can be tested on its own.
// archiveRows: [{key, date, client, client1, client2}]
// liveRows:    [{date, client, client1, client2}]
// dobByClient: {normalisedName: 'yyyy-mm-dd'}
function planPurge_(archiveRows, liveRows, dobByClient, today) {
  const first = {}, last = {};
  archiveRows.concat(liveRows).forEach(function (r) {
    if (!r.date) return;
    bookingClients_(r).forEach(function (c) {
      if (!first[c] || r.date < first[c]) first[c] = r.date;
      if (!last[c] || r.date > last[c]) last[c] = r.date;
    });
  });

  const due = {}, rule = {};
  Object.keys(last).forEach(function (c) {
    const dob = dobByClient[c];
    if (dob && addYears_(dob, RETENTION_ADULT_AGE) > first[c]) {
      const a = addYears_(dob, RETENTION_CHILD_UNTIL_AGE);
      const b = addYears_(last[c], RETENTION_CHILD_YEARS_AFTER_END);
      due[c] = a > b ? a : b;
      rule[c] = 'child';
    } else {
      due[c] = addYears_(last[c], RETENTION_ADULT_YEARS);
      rule[c] = 'adult';
    }
  });

  const purgeKeys = {};
  const keptClients = {};
  archiveRows.forEach(function (r) {
    const clients = bookingClients_(r);
    var purge;
    if (!clients.length) purge = !!r.date && today >= addYears_(r.date, RETENTION_ADULT_YEARS);
    else purge = clients.every(function (c) { return today >= due[c]; });
    if (purge) purgeKeys[r.key] = true;
    else clients.forEach(function (c) { keptClients[c] = true; });
  });
  liveRows.forEach(function (r) { bookingClients_(r).forEach(function (c) { keptClients[c] = true; }); });

  const purgedClients = Object.keys(due).filter(function (c) {
    return today >= due[c] && !keptClients[c];
  });
  return { purgeKeys: purgeKeys, purgedClients: purgedClients, due: due, rule: rule, last: last };
}

function gatherPurgeInputs_(ss) {
  const archiveRows = [];
  archiveYearSheets_(ss).forEach(function (sh) {
    readArchiveSheet_(sh).forEach(function (v) { archiveRows.push(archiveRowToObj_(v, sh.getName())); });
  });
  const record = readRecord();
  const liveRows = (record.bookings || []).concat(record.zoomBookings || []).map(function (b) {
    return { date: b.date, client: b.note, client1: b.client1, client2: b.client2 };
  });
  const dobByClient = {};
  const ch = ss.getSheetByName(ARCHIVE_CHILDREN_SHEET);
  if (ch && ch.getLastRow() > 1) {
    ch.getRange(2, 1, ch.getLastRow() - 1, 2).getValues().forEach(function (r) {
      const name = normName_(r[0]);
      const dob = parseDob_(r[1]);
      if (name && dob) dobByClient[name] = dob;
      else if (name) Logger.log('Children tab: could not read date of birth for "' + r[0] + '"');
    });
  }
  return { archiveRows: archiveRows, liveRows: liveRows, dobByClient: dobByClient };
}

function purgeExpiredRecords() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = getArchiveSs_();
    const inp = gatherPurgeInputs_(ss);
    const plan = planPurge_(inp.archiveRows, inp.liveRows, inp.dobByClient, todayYmd_());
    const purgeCount = Object.keys(plan.purgeKeys).length;
    if (!purgeCount) {
      logArchiveActivity_(ss, 'Purge', 0, 0, 0);
      Logger.log('purgeExpiredRecords: nothing due');
      return;
    }

    // Rewrite each affected year sheet: kept rows first, then clear the tail.
    // If anything fails midway the worst case is leftover duplicates, never lost rows.
    archiveYearSheets_(ss).forEach(function (sh) {
      const vals = readArchiveSheet_(sh);
      const keep = vals.filter(function (v) { return !plan.purgeKeys[v[0]]; });
      if (keep.length === vals.length) return;
      if (!keep.length) { ss.deleteSheet(sh); return; }
      sh.getRange(2, 1, keep.length, ARCHIVE_COLS).setNumberFormat('@').setValues(keep);
      sh.getRange(2 + keep.length, 1, vals.length - keep.length, ARCHIVE_COLS).clearContent();
    });

    // Remove purged children's names and dates of birth too.
    const gone = {};
    plan.purgedClients.forEach(function (c) { gone[c] = true; });
    const ch = ss.getSheetByName(ARCHIVE_CHILDREN_SHEET);
    if (ch && ch.getLastRow() > 1) {
      const n = ch.getLastRow() - 1;
      const rows = ch.getRange(2, 1, n, 3).getValues();
      const keepRows = rows.filter(function (r) { return !gone[normName_(r[0])]; });
      if (keepRows.length !== rows.length) {
        ch.getRange(2, 1, n, 3).clearContent();
        if (keepRows.length) ch.getRange(2, 1, keepRows.length, 3).setValues(keepRows);
      }
    }

    logArchiveActivity_(ss, 'Purge', 0, purgeCount, plan.purgedClients.length);
    Logger.log('purgeExpiredRecords: removed ' + purgeCount + ' bookings, ' +
      plan.purgedClients.length + ' clients');
  } finally {
    lock.releaseLock();
  }
}

// Run from the editor to see each client's purge date without deleting anything.
function previewPurge() {
  const ss = getArchiveSs_();
  const inp = gatherPurgeInputs_(ss);
  const today = todayYmd_();
  const plan = planPurge_(inp.archiveRows, inp.liveRows, inp.dobByClient, today);
  Object.keys(plan.due).sort().forEach(function (c) {
    Logger.log(c + ' — ' + plan.rule[c] + ', last session ' + plan.last[c] +
      ', purge due ' + plan.due[c] + (today >= plan.due[c] ? '  <-- DUE NOW' : ''));
  });
  Logger.log('Would purge ' + Object.keys(plan.purgeKeys).length + ' archived bookings today.');
}

// ── Search (action=searchArchive) ─────────────────────────────────────────
// Params: client (required, 2+ letters), from, to (yyyy-mm-dd, optional), who (optional)
function searchArchive_(p) {
  const q = normName_(p.client);
  if (q.length < 2) throw new Error('Type at least 2 letters of the client name');
  const from = p.from || '', to = p.to || '', who = p.who || '';
  const matches = function (r) {
    const hay = [r.client, r.client1, r.client2].map(normName_).join(' | ');
    return hay.indexOf(q) > -1 &&
      (!from || r.date >= from) && (!to || r.date <= to) && (!who || r.who === who);
  };

  const results = [];
  const seen = {};
  const record = readRecord();
  [['R', 'Room', record.bookings || []], ['Z', 'Zoom', record.zoomBookings || []]].forEach(function (t) {
    t[2].forEach(function (b) {
      const r = {
        source: 'live', type: t[1], date: b.date, start: b.start, end: b.end, who: b.who,
        client: b.note, client1: b.client1, client2: b.client2, category: b.category || '',
        paymentStatus: b.paymentStatus || '', amount: b.expectedAmount != null ? String(b.expectedAmount) : '',
        dna: !!b.dna
      };
      if (!matches(r)) return;
      seen[t[0] + '-' + b.id] = true;
      results.push(r);
    });
  });

  if (PropertiesService.getScriptProperties().getProperty('ARCHIVE_SS_ID')) {
    const ss = getArchiveSs_();
    archiveYearSheets_(ss).forEach(function (sh) {
      readArchiveSheet_(sh).forEach(function (v) {
        const r = archiveRowToObj_(v, sh.getName());
        if (seen[r.key] || !matches(r)) return;
        r.source = 'archive';
        delete r.sheet;
        results.push(r);
      });
    });
  }

  results.sort(function (a, b) {
    return (b.date + (b.start || '')).localeCompare(a.date + (a.start || ''));
  });
  return { total: results.length, results: results.slice(0, 500) };
}

// ── Correct a misspelt client name ────────────────────────────────────────
// Edit the three values below, save, then pick fixClientName in the
// function dropdown and click Run. Fixes live room + Zoom bookings and the
// archive. Matching ignores capitals and extra spaces but must otherwise be
// the whole name exactly as it was typed.
function fixClientName() {
  const WRONG_NAME = 'Kate owens';
  const RIGHT_NAME = 'Kate Owen';
  const ONLY_ON_DATE = '2026-07-28'; // yyyy-mm-dd, or '' to fix every date
  renameClient_(WRONG_NAME, RIGHT_NAME, ONLY_ON_DATE);
}

function renameClient_(wrong, right, onlyDate) {
  const w = normName_(wrong);
  right = String(right || '').trim();
  if (!w || !right) throw new Error('Fill in both WRONG_NAME and RIGHT_NAME first');
  if (w === normName_(right)) throw new Error('WRONG_NAME and RIGHT_NAME are the same');

  // Returns true if it changed anything. Works on {note|client, client1, client2, couples?}.
  function fix(obj, noteField) {
    var changed = false;
    if (normName_(obj.client1) === w) { obj.client1 = right; changed = true; }
    if (normName_(obj.client2) === w) { obj.client2 = right; changed = true; }
    if (normName_(obj[noteField]) === w) { obj[noteField] = right; changed = true; }
    else if (changed && obj.client1 && obj.client2) { obj[noteField] = obj.client1 + ' & ' + obj.client2; }
    return changed;
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // Live sheet
    var record = readRecord();
    var live = 0;
    [['Room', record.bookings || []], ['Zoom', record.zoomBookings || []]].forEach(function (t) {
      t[1].forEach(function (b) {
        if (onlyDate && b.date !== onlyDate) return;
        if (fix(b, 'note')) {
          if (t[0] === 'Zoom' && b.topic && normName_(b.topic).indexOf(w) > -1) {
            b.topic = String(b.topic).replace(new RegExp(String(wrong).trim().replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'i'), right);
          }
          live++;
          Logger.log('Fixed live ' + t[0] + ' booking: ' + b.date + ' ' + (b.start || '') + ' (' + b.who + ')');
        }
      });
    });
    if (live) writeRecord(record);

    // Archive
    var archived = 0;
    if (PropertiesService.getScriptProperties().getProperty('ARCHIVE_SS_ID')) {
      const ss = getArchiveSs_();
      archiveYearSheets_(ss).forEach(function (sh) {
        const vals = readArchiveSheet_(sh);
        var touched = false;
        vals.forEach(function (v) {
          if (onlyDate && v[2] !== onlyDate) return;
          const o = { client: v[6], client1: v[7], client2: v[8] };
          if (fix(o, 'client')) {
            v[6] = o.client; v[7] = o.client1; v[8] = o.client2;
            touched = true; archived++;
            Logger.log('Fixed archived booking: ' + v[2] + ' ' + v[3] + ' (' + v[5] + ')');
          }
        });
        if (touched) sh.getRange(2, 1, vals.length, ARCHIVE_COLS).setNumberFormat('@').setValues(vals);
      });
    }

    Logger.log('Done: ' + live + ' live and ' + archived + ' archived bookings renamed to "' + right + '"' +
      (onlyDate ? ' on ' + onlyDate : ''));
    if (!live && !archived) Logger.log('Nothing matched — check the spelling of WRONG_NAME (and ONLY_ON_DATE).');
  } finally {
    lock.releaseLock();
  }
}
