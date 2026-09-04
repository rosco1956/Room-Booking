// Room Booking — Google Apps Script
// Deploy as: Web app / Execute as: Me / Who has access: Anyone

const SHEET_NAME = 'Data';
const TM = 'https://rest.textmagic.com/api/v2';

function getSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.getSheetByName('Sheet1');
    if (!sheet) {
        sheet = ss.insertSheet(SHEET_NAME);
        sheet.getRange('A1').setValue(JSON.stringify({
            bookings: [], people: ['Rosswell', 'Donna', 'Helen']
        }));
    }
    return sheet;
}

// ── Chunked record storage ────────────────────────────────────────────────
const CHUNK_SIZE = 45000;   // safely under the 50,000-char cell limit
const MAX_CHUNK_ROWS = 300; // upper bound on chunks ever cleared/written

function readRecord() {
    const sheet = getSheet();
    const values = sheet.getRange(1, 1, MAX_CHUNK_ROWS, 1).getValues();
    var combined = '';
    for (var i = 0; i < values.length; i++) {
        const cell = values[i][0];
        if (cell === '' || cell === null || cell === undefined) break;
        combined += String(cell);
    }
    try { return JSON.parse(combined || '{}'); }
    catch (e) { return {}; }
}

function writeRecord(record) {
    const sheet = getSheet();
    const json = JSON.stringify(record);
    const chunks = [];
    for (var i = 0; i < json.length; i += CHUNK_SIZE) {
        chunks.push(json.slice(i, i + CHUNK_SIZE));
    }
    if (!chunks.length) chunks.push('{}');
    if (chunks.length > MAX_CHUNK_ROWS) {
        throw new Error('Record too large: needs ' + chunks.length + ' chunks, max is ' + MAX_CHUNK_ROWS);
    }

    // Find how many rows currently hold data, so we only need to touch that many
    // rows (or the new chunk count, whichever is larger) instead of always writing
    // all MAX_CHUNK_ROWS. Keeps normal writes small and fast.
    const existing = sheet.getRange(1, 1, MAX_CHUNK_ROWS, 1).getValues();
    var previousUsedRows = 0;
    for (var p = 0; p < existing.length; p++) {
        const cell = existing[p][0];
        if (cell === '' || cell === null || cell === undefined) break;
        previousUsedRows++;
    }
    const rowsToWrite = Math.max(chunks.length, previousUsedRows);

    // Single atomic write covering just the rows that need to change: real data in
    // the rows we need, blank strings in any old leftover rows beyond that. This
    // replaces any leftover chunks from a previous, larger write in the SAME call as
    // writing the new data — so there is no separate "clear" step and therefore no
    // gap where a concurrent read can see an empty sheet. Sheets applies one
    // setValues() range write as a single operation, so a read always sees either
    // the complete old data or the complete new data, never a partial/cleared state.
    const out = [];
    for (var r = 0; r < rowsToWrite; r++) {
        out.push([r < chunks.length ? chunks[r] : '']);
    }
    sheet.getRange(1, 1, rowsToWrite, 1).setValues(out);
}

function corsOutput(data) {
    const output = ContentService.createTextOutput(JSON.stringify(data));
    output.setMimeType(ContentService.MimeType.JSON);
    return output;
}

function ok(data) { return corsOutput(Object.assign({ ok: true }, data)); }
function fail(msg) { return corsOutput({ ok: false, error: msg }); }

// ── Trim helpers ─────────────────────────────────────────────────────────
function trimRecord(record) {
    const todayStr = Utilities.formatDate(new Date(), 'Europe/London', 'yyyy-MM-dd');
    const oneYearAgo = new Date();
    oneYearAgo.setFullYear(oneYearAgo.getFullYear() - 1);
    const cutoff = Utilities.formatDate(oneYearAgo, 'Europe/London', 'yyyy-MM-dd');
    if (record.zoomBookings) {
        record.zoomBookings = record.zoomBookings.filter(function (b) { return b.date >= cutoff; });
    }
    if (record.bookings) {
        record.bookings = record.bookings.filter(function (b) { return b.date >= cutoff; });
    }
    if (record.smsLog) record.smsLog = record.smsLog.slice(0, 20);
    return record;
}

// ── MONTHLY INVOICE FEATURE — top-level functions ──────────────────────────
// Fixed rate per booking, per practitioner. Edit these to your real rates.
const INVOICE_RATES = {
    'Rosswell': 12,
    'Donna': 12,
    'Helen': 12
};
function getAllBookings_() {
    const record = readRecord();
    return record.bookings || [];
}

function getBookingsForInvoice(practitioner, year, month, client, category) {
    const allBookings = getAllBookings_();
    return allBookings.filter(function (b) {
        const d = new Date(b.date);
        const matchesPractitioner = b.who === practitioner;
        const matchesDate = d.getFullYear() === Number(year) && (d.getMonth() + 1) === Number(month);
        const matchesClient = !client || client === 'All' || b.note === client;
        const matchesCategory = !category || category === 'All' || b.category === category;
        return matchesPractitioner && matchesDate && matchesClient && matchesCategory;
    }).sort(function (a, b) { return new Date(a.date) - new Date(b.date); });
}

function handleInvoiceMetaRequest(e) {
    const practitioner = e.parameter.practitioner;
    const allBookings = getAllBookings_();
    const filtered = practitioner ? allBookings.filter(function (b) { return b.who === practitioner; }) : allBookings;
    const clients = Array.from(new Set(filtered.map(function (b) { return b.note; }).filter(Boolean))).sort();
    const categories = Array.from(new Set(filtered.map(function (b) { return b.category; }).filter(Boolean))).sort();
    return ContentService.createTextOutput(JSON.stringify({ clients: clients, categories: categories }))
        .setMimeType(ContentService.MimeType.JSON);
}

function handleInvoiceRequest(e) {
    const practitioner = e.parameter.practitioner;
    const year = e.parameter.year;
    const month = e.parameter.month;
    const client = e.parameter.client;
    const category = e.parameter.category;
    const rate = INVOICE_RATES[practitioner] || 0;

    const bookings = getBookingsForInvoice(practitioner, year, month, client, category);
    const html = buildInvoiceHtml_(practitioner, year, month, bookings, rate, client, category);

    let filename = practitioner + '-invoice-' + year + '-' + String(month).padStart(2, '0');
    if (client && client !== 'All') filename += '-' + client.replace(/[^a-z0-9]+/gi, '');
    if (category && category !== 'All') filename += '-' + category.replace(/[^a-z0-9]+/gi, '');
    filename += '.pdf';

    const pdfBlob = htmlToPdfBlob_(html, filename);

    const folder = getOrCreateInvoiceFolder_();
    const savedFile = folder.createFile(pdfBlob);

    MailApp.sendEmail({
        to: INVOICE_RECIPIENT,
        subject: 'Invoice: ' + filename,
        body: 'Attached is the invoice for ' + practitioner + ' (' + year + '-' + month + ').',
        attachments: [pdfBlob]
    });

    return ok({
        message: 'Invoice created and emailed',
        fileUrl: savedFile.getUrl(),
        filename: filename
    });
}

// ── END MONTHLY INVOICE FEATURE top-level functions ────────────────────────

function calcDurationHours_(start, end) {
    var s = start.split(':'), e = end.split(':');
    var startMins = parseInt(s[0]) * 60 + parseInt(s[1]);
    var endMins = parseInt(e[0]) * 60 + parseInt(e[1]);
    return (endMins - startMins) / 60;
}

function buildInvoiceHtml_(practitioner, year, month, bookings, rate, client, category) {
    const monthName = Utilities.formatDate(new Date(year, month - 1, 1), Session.getScriptTimeZone(), 'MMMM yyyy');
    const filterBits = [];
    if (client && client !== 'All') filterBits.push('Client: ' + client);
    if (category && category !== 'All') filterBits.push('Category: ' + category);
    const filterLine = filterBits.length ? '<div style="color:#666;">' + filterBits.join(' &middot; ') + '</div>' : '';

    var runningTotal = 0;
    const rows = bookings.map(function (b) {
        const dateStr = Utilities.formatDate(new Date(b.date), Session.getScriptTimeZone(), 'dd MMM yyyy');
        const hours = calcDurationHours_(b.start, b.end);
        const cost = hours * rate;
        runningTotal += cost;
        return '<tr>' +
            '<td style="padding:6px 10px;border-bottom:1px solid #eee;">' + dateStr + '</td>' +
            '<td style="padding:6px 10px;border-bottom:1px solid #eee;">' + (b.note || '') + '</td>' +
            '<td style="padding:6px 10px;border-bottom:1px solid #eee;">' + (b.category || '') + '</td>' +
            '<td style="padding:6px 10px;border-bottom:1px solid #eee;text-align:right;">' + b.start + '\u2013' + b.end + ' (' + hours.toFixed(2) + 'h)</td>' +
            '<td style="padding:6px 10px;border-bottom:1px solid #eee;text-align:right;">£' + cost.toFixed(2) + '</td>' +
            '</tr>';
    }).join('');

    return '<html><body style="font-family:Arial,sans-serif;color:#222;padding:24px;">' +
        '<h2 style="margin-bottom:0;">Invoice — ' + practitioner + '</h2>' +
        '<div style="color:#666;">' + monthName + '</div>' +
        filterLine +
        '<div style="margin-bottom:20px;"></div>' +
        '<table style="border-collapse:collapse;width:100%;">' +
        '<thead><tr>' +
        '<th style="text-align:left;padding:6px 10px;border-bottom:2px solid #333;">Date</th>' +
        '<th style="text-align:left;padding:6px 10px;border-bottom:2px solid #333;">Client</th>' +
        '<th style="text-align:left;padding:6px 10px;border-bottom:2px solid #333;">Category</th>' +
        '<th style="text-align:right;padding:6px 10px;border-bottom:2px solid #333;">Time</th>' +
        '<th style="text-align:right;padding:6px 10px;border-bottom:2px solid #333;">Cost</th>' +
        '</tr></thead><tbody>' + rows + '</tbody>' +
        '<tfoot><tr>' +
        '<td colspan="4" style="padding:10px 10px 0;text-align:right;font-weight:bold;border-top:2px solid #333;">Total (' + bookings.length + ' sessions):</td>' +
        '<td style="padding:10px 10px 0;text-align:right;font-weight:bold;border-top:2px solid #333;">£' + runningTotal.toFixed(2) + '</td>' +
        '</tr></tfoot>' +
        '</table>' +
        '<div style="margin-top:32px;padding-top:16px;border-top:1px solid #ccc;font-size:13px;color:#444;">' +
        '<div style="font-weight:bold;margin:0;">Payment details</div>' +
        '<div style="margin:0;">Account Name: GADSDEN R &amp; D</div>' +
        '<div style="margin:0;">Account Number: 40791181</div>' +
        '<div style="margin:0;">Sort Code: 52-41-07</div>' +
        '</div>' +
        '</body></html>';
}

function htmlToPdfBlob_(html, filename) {
    const htmlBlob = Utilities.newBlob(html, 'text/html', filename + '.html');
    const tempFile = Drive.Files.create(
        { name: filename, mimeType: MimeType.GOOGLE_DOCS },
        htmlBlob
    );
    const pdfBlob = DriveApp.getFileById(tempFile.id)
        .getAs('application/pdf')
        .setName(filename);
    DriveApp.getFileById(tempFile.id).setTrashed(true);
    return pdfBlob;
}


// ── Main GET handler ─────────────────────────────────────────────────────
function doGet(e) {
    try {
        return doGetInner(e);
    } catch (outerErr) {
        try {
            PropertiesService.getScriptProperties().setProperty(
                'lastFatalError',
                new Date().toISOString() + ' | ' + outerErr.message + ' | ' + (outerErr.stack || '')
            );
        } catch (logErr) { }
        return fail('FATAL: ' + outerErr.message);
    }
}

// ── Config ──────────────────────────────────────────────
const INVOICE_FOLDER_NAME = 'Room Booking Invoices';
const INVOICE_RECIPIENT = 'rosswell.gadsden@tuta.io';

/**
 * action=invoiceMeta — cheap preview, no PDF/Drive/email work.
 * Params: year, month (1-12)
 */

/**
 * action=invoice — generates PDFs, saves to Drive, emails combined to INVOICE_RECIPIENT.
 * Params: year, month (1-12)
 */

function getOrCreateInvoiceFolder_() {
    const folders = DriveApp.getFoldersByName(INVOICE_FOLDER_NAME);
    return folders.hasNext() ? folders.next() : DriveApp.createFolder(INVOICE_FOLDER_NAME);
}

function doGetInner(e) {
    if (!e || !e.parameter) return ok({ status: 'alive' });
    const action = (e.parameter.action || 'load');

    if (action === 'lastError') {
        const msg = PropertiesService.getScriptProperties().getProperty('lastFatalError') || 'No fatal error recorded';
        return ok({ lastFatalError: msg });
    }

    if (action === 'load') {
        try {
            var record = readRecord();
            if (!record.bookings) record.bookings = [];
            if (!record.people) record.people = [];
            return ok({ record: record });
        } catch (err) { return fail(err.message); }
    }

    if (action === 'addBooking') {
        try {
            const lock = LockService.getScriptLock();
            lock.waitLock(10000);
            try {
                var record = readRecord();
                if (!record.bookings) record.bookings = [];
                const newBooking = JSON.parse(decodeURIComponent(e.parameter.booking || '{}'));
                const clash = record.bookings.find(function (b) {
                    return b.date === newBooking.date &&
                        !(newBooking.end <= b.start || newBooking.start >= b.end);
                });
                if (clash) return fail('Clash with ' + clash.who + ' (' + clash.start + '-' + clash.end + ')');
                record.bookings.push(newBooking);
                record = trimRecord(record);
                writeRecord(record);
                return ok({ count: record.bookings.length, record: { bookings: record.bookings } });
            } finally {
                lock.releaseLock();
            }
        } catch (err) { return fail(err.message); }
    }

    if (action === 'deleteBooking') {
        try {
            const lock = LockService.getScriptLock();
            lock.waitLock(10000);
            try {
                var record = readRecord();
                if (!record.bookings) record.bookings = [];
                const delId = parseInt(e.parameter.id || '0');
                const before = (record.bookings || []).length;
                record.bookings = (record.bookings || []).filter(function (b) { return b.id !== delId; });
                writeRecord(record);
                return ok({ count: record.bookings.length, deleted: before - record.bookings.length });
            } finally {
                lock.releaseLock();
            }
        } catch (err) { return fail(err.message); }
    }

    if (action === 'updateBooking') {
        try {
            const lock = LockService.getScriptLock();
            lock.waitLock(10000);
            try {
                var record = readRecord();
                if (!record.bookings) record.bookings = [];
                const updated = JSON.parse(decodeURIComponent(e.parameter.booking || '{}'));
                const clash = (record.bookings || []).find(function (b) {
                    return b.id !== updated.id && b.date === updated.date &&
                        !(updated.end <= b.start || updated.start >= b.end);
                });
                if (clash) return fail('Clash with ' + clash.who + ' (' + clash.start + '-' + clash.end + ')');
                record.bookings = (record.bookings || []).map(function (b) {
                    return b.id === updated.id ? updated : b;
                });
                writeRecord(record);
                return ok({ count: record.bookings.length });
            } finally {
                lock.releaseLock();
            }
        } catch (err) { return fail(err.message); }
    }

    if (action === 'saveSettings') {
        try {
            const lock = LockService.getScriptLock();
            lock.waitLock(10000);
            try {
                var record = readRecord();
                const settings = JSON.parse(decodeURIComponent(e.parameter.data || '{}'));
                if (settings.people) record.people = settings.people;
                if (settings.rate != null) record.rate = settings.rate;
                if (settings.smsLog) record.smsLog = settings.smsLog;
                if (settings.waitingList) record.waitingList = settings.waitingList;
                if (settings.appPin) record.appPin = settings.appPin;
                if (settings.zoomBookings) {
                    const existing = record.zoomBookings || [];
                    const existingById = {};
                    existing.forEach(function (b) { existingById[String(b.id)] = b; });
                    settings.zoomBookings.forEach(function (b) { existingById[String(b.id)] = b; });
                    record.zoomBookings = Object.values(existingById);
                }
                record = trimRecord(record);
                writeRecord(record);
                return ok({});
            } finally {
                lock.releaseLock();
            }
        } catch (err) { return fail(err.message); }
    }

    if (action === 'save') {
        try {
            const lock = LockService.getScriptLock();
            lock.waitLock(10000);
            try {
                const body = JSON.parse(decodeURIComponent(e.parameter.data || '{}'));
                writeRecord(body);
                return ok({});
            } finally {
                lock.releaseLock();
            }
        } catch (err) { return fail(err.message); }
    }

    if (action === 'tmMe') {
        try {
            const r = tmGet(e, '/user');
            const d = JSON.parse(r.getContentText());
            const code = r.getResponseCode();
            return code === 200
                ? ok({ firstName: d.firstName, lastName: d.lastName, balance: d.balance })
                : fail('HTTP ' + code + ': ' + JSON.stringify(d));
        } catch (err) { return fail(err.message); }
    }

    if (action === 'tmContacts') {
        try {
            const search = e.parameter.search || '';
            const limit = e.parameter.limit || 50;
            var path = '/contacts/search?limit=' + limit + '&page=1';
            if (search) path += '&query=' + encodeURIComponent(search);
            const r = tmGet(e, path);
            const d = JSON.parse(r.getContentText());
            return r.getResponseCode() === 200
                ? ok({ resources: d.resources || [] })
                : fail(d.message || 'Could not fetch contacts');
        } catch (err) { return fail(err.message); }
    }

    if (action === 'tmSend') {
        try {
            const text = e.parameter.text || '';
            const phones = e.parameter.phones || '';
            if (!text || !phones) return fail('Missing text or phone');
            const r = tmPost(e, '/messages', { text: text, phones: phones });
            const d = JSON.parse(r.getContentText());
            const code = r.getResponseCode();
            Logger.log('tmSend HTTP ' + code + ': ' + r.getContentText().slice(0, 200));
            return code === 201
                ? ok({ id: d.id })
                : fail('HTTP ' + code + ': ' + (d.message || d.error || JSON.stringify(d)));
        } catch (err) { return fail(err.message); }
    }

    if (action === 'tmContactEmail') {
        try {
            const search = e.parameter.search || '';
            const limit = e.parameter.limit || 5;
            var path = '/contacts/search?limit=' + limit + '&page=1';
            if (search) path += '&query=' + encodeURIComponent(search);
            const r = tmGet(e, path);
            const d = JSON.parse(r.getContentText());
            if (r.getResponseCode() !== 200) return fail(d.message || 'Failed');
            const contacts = (d.resources || [])
                .filter(function (c) { return nameMatches(c, search); })
                .map(function (c) {
                    return {
                        name: ((c.firstName || '') + ' ' + (c.lastName || '')).trim(),
                        phone: c.phone || '',
                        email: c.email || ''
                    };
                });
            return ok({ contacts: contacts });
        } catch (err) { return fail(err.message); }
    }

    if (action === 'tmIncrementAppts') {
        try {
            const contactId = e.parameter.contactId || '';
            if (!contactId) return fail('No contactId');
            const r = tmGet(e, '/contacts/' + contactId);
            if (r.getResponseCode() !== 200) return fail('Contact not found');
            const contact = JSON.parse(r.getContentText());
            const fields = contact.customFields || [];
            const apptField = fields.find(function (f) { return f.name === 'Appointments'; });
            const current = apptField ? parseInt(apptField.value || '0') : 0;
            const newCount = current + 1;
            if (apptField) {
                var apptPayload = {};
                apptPayload['customFieldValues[' + apptField.userCustomFieldId + ']'] = String(newCount);
                tmPost(e, '/contacts/' + contactId, apptPayload);
            }
            return ok({ appointments: newCount });
        } catch (err) { return fail(err.message); }
    }

    // ── MONTHLY INVOICE FEATURE — doGet wiring ──────────────────────────
    if (action === 'invoice') {
        try {
            return handleInvoiceRequest(e);
        } catch (err) { return fail(err.message); }
    }

    if (action === 'invoiceMeta') {
        try {
            return handleInvoiceMetaRequest(e);
        } catch (err) { return fail(err.message); }
    }
    // ── END MONTHLY INVOICE FEATURE doGet wiring ────────────────────────

    if (action === 'ical') {
        try {
            const who = e.parameter.who || '';
            var record = readRecord();
            const bookings = (record.bookings || []).filter(function (b) {
                return !who || b.who === who;
            });
            const lines = [
                'BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Room Booking//EN',
                'CALSCALE:GREGORIAN', 'METHOD:PUBLISH',
                'X-WR-CALNAME:Room Booking' + (who ? ' - ' + who : ''),
                'X-WR-TIMEZONE:Europe/London',
                'REFRESH-INTERVAL;VALUE=DURATION:PT1H', 'X-PUBLISHED-TTL:PT1H'
            ];
            bookings.forEach(function (b) {
                lines.push('BEGIN:VEVENT');
                lines.push('UID:' + b.id + '@roombooking.gas');
                lines.push('DTSTAMP:' + icsDateGAS(
                    Utilities.formatDate(new Date(), 'Europe/London', 'yyyy-MM-dd'),
                    Utilities.formatDate(new Date(), 'Europe/London', 'HH:mm')
                ));
                lines.push('DTSTART;TZID=Europe/London:' + icsDateGAS(b.date, b.start));
                lines.push('DTEND;TZID=Europe/London:' + icsDateGAS(b.date, b.end));
                lines.push('SUMMARY:' + icsEscapeGAS(b.who + ' - Room 34' + (b.note ? ' (' + b.note + ')' : '')));
                lines.push('LOCATION:' + icsEscapeGAS('Room 34, Canute Suite, Royal Mail House, Terminus Ter, Southampton SO14 3FD'));
                lines.push('DESCRIPTION:' + icsEscapeGAS(b.note ? 'Client: ' + b.note : 'Room booking for ' + b.who));
                lines.push('STATUS:CONFIRMED');
                lines.push('END:VEVENT');
            });
            lines.push('END:VCALENDAR');
            return ContentService.createTextOutput(lines.join('\r\n'))
                .setMimeType(ContentService.MimeType.ICAL);
        } catch (err) { return fail(err.message); }
    }

    if (action === 'shortenUrl') {
        try {
            const who = e.parameter.who || '';
            const longUrl = decodeURIComponent(e.parameter.url || '');
            if (!longUrl) return fail('No URL provided');
            const shortUrl = shortenUrl(who, longUrl);
            return ok({ shortUrl: shortUrl });
        } catch (err) { return fail(err.message); }
    }


if (action === 'createZoom') {
    try {
      const who       = e.parameter.who||'';
      const topic     = decodeURIComponent(e.parameter.topic||'Session');
      const date      = e.parameter.date||'';
      const start     = e.parameter.start||'09:00';
      const duration  = parseInt(e.parameter.duration||'60');
      const pendingId = parseInt(e.parameter.pendingId||'0');
      const client    = decodeURIComponent(e.parameter.client||'');
      const client2   = decodeURIComponent(e.parameter.client2||'');
      const couples   = e.parameter.couples==='1';
      const requirePrepay  = e.parameter.requirePrepay==='1';
      const expectedAmount = requirePrepay ? parseFloat(e.parameter.expectedAmount||'0') : null;
      const clientEmail    = requirePrepay ? decodeURIComponent(e.parameter.clientEmail||'') : '';
      if (!who||!date) return fail('Missing who or date');
      if (requirePrepay && (!expectedAmount || expectedAmount<=0)) return fail('Missing expectedAmount for prepayment booking');
      if (requirePrepay && !clientEmail) return fail('Missing clientEmail for prepayment booking');

      const lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        var record2 = readRecord();
        if (!record2.zoomBookings) record2.zoomBookings = [];

        // Idempotency guard: if this pendingId already has a meeting
        // (e.g. this is a retry after a dropped or garbled response),
        // return the existing booking instead of creating a duplicate
        // Zoom meeting.
        if (pendingId) {
          var existing = record2.zoomBookings.find(function(b){ return b.pendingId === pendingId; });
          if (existing) {
            Logger.log('createZoom: pendingId '+pendingId+' already has a meeting, returning existing (no new Zoom meeting created)');
            return ok(existing);
          }
        }

        Logger.log('Creating Zoom meeting for '+who+' on '+date+' at '+start);
        const meeting = createZoomMeeting(who, topic, date, start, duration);
        Logger.log('Meeting created: '+meeting.joinUrl);

        var timeParts = start.split(':');
        var startMins = parseInt(timeParts[0])*60+parseInt(timeParts[1]);
        var endMins   = startMins+duration;
        var endTime   = ('0'+Math.floor(endMins/60)%24).slice(-2)+':'+('0'+endMins%60).slice(-2);

        const note = couples&&client2 ? client+' & '+client2 : client;
        const zb = {
          id:           Date.now(),
          pendingId:    pendingId,
          who:          who,
          date:         date,
          start:        start,
          end:          endTime,
          note:         note,
          client1:      client,
          client2:      client2,
          couples:      couples&&!!client2,
          zoomMeetingId:meeting.meetingId,
          zoomJoinUrl:  meeting.joinUrl,
          zoomShortUrl: meeting.shortUrl,
          zoomPassword: meeting.password,
          zoomStartUrl: meeting.startUrl,
          isZoom:       true,
          updated:      new Date().toISOString(),
          // Prepayment fields — omitted entirely when requirePrepay is off,
          // so existing bookings/behaviour are completely unaffected.
          ...(requirePrepay ? {
            requirePrepay:  true,
            expectedAmount: expectedAmount,
            clientEmail:    clientEmail,
            paymentStatus:  'awaiting'
          } : {})
        };

        record2.zoomBookings.push(zb);
        record2 = trimRecord(record2);
        writeRecord(record2);
        Logger.log('Saved to sheet, pendingId: '+pendingId+(requirePrepay?' (awaiting payment, £'+expectedAmount+')':''));
        return ok(zb);
      } finally {
        lock.releaseLock();
      }
    } catch(err) {
      Logger.log('createZoom error: '+err.message);
      return fail(err.message);
    }
  }

  if (action === 'sendPaymentLinkEmail') {
    try {
      const id        = parseInt(e.parameter.id||'0');
      const to        = e.parameter.to||'';
      const firstName = e.parameter.firstName||'there';
      const who       = e.parameter.who||'';
      const date      = e.parameter.date||'';
      const start     = e.parameter.start||'';
      const paymentLinkRaw = e.parameter.paymentLink||'';
      const linkMatch = paymentLinkRaw.match(/https?:\/\/\S+/);
      const paymentLink = linkMatch ? linkMatch[0].replace(/[.,;:!?)\]]+$/,'') : '';
      const amount    = e.parameter.amount||'';
      const mobile    = e.parameter.mobile||'';
      if (!to) return fail('No email address');
      if (!paymentLink) return fail('No payment link');

      sendPaymentRequestEmailInternal({to, firstName, who, date, start, paymentLink, amount, mobile, reminderNumber: 0});

      // Persist the link + first-sent timestamp on the booking so the
      // automated chase (checkZettlePayments) can reuse the exact same
      // link later without needing you to paste it again.
      if (id) {
        const lock = LockService.getScriptLock();
        lock.waitLock(10000);
        try {
          var record4 = readRecord();
          var zb4 = (record4.zoomBookings||[]).find(function(b){ return b.id === id; });
          if (zb4) {
            zb4.paymentLink = paymentLink;
            if (!zb4.paymentLinkSentAt) zb4.paymentLinkSentAt = new Date().toISOString();
            zb4.chaseCount = zb4.chaseCount || 0;
            writeRecord(trimRecord(record4));
          }
        } finally {
          lock.releaseLock();
        }
      }
      return ok({sent:true});
    } catch(err) { return fail(err.message); }
  }

  if (action === 'markZoomPaid') {
    // Manual override — used for the ambiguous-match case, or any payment
    // taken outside the automatic Zettle matching (cash, bank transfer, etc).
    try {
      const id = parseInt(e.parameter.id||'0');
      if (!id) return fail('Missing id');
      const lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        var record3 = readRecord();
        var zb2 = (record3.zoomBookings||[]).find(function(b){ return b.id === id; });
        if (!zb2) return fail('Booking not found');

        zb2.paymentStatus = 'paid';
        zb2.paidAt = new Date().toISOString();
        zb2.manualPaymentOverride = true;
        delete zb2.paymentAmbiguous;
        delete zb2.paymentCandidates;

        if (zb2.clientEmail) {
          try {
            sendZoomConfirmationEmailInternal(zb2);
            zb2.confirmationSentAt = new Date().toISOString();
          } catch (e2) {
            Logger.log('markZoomPaid: confirmation email failed for '+id+': '+e2.message);
          }
        }

        record3 = trimRecord(record3);
        writeRecord(record3);
        return ok(zb2);
      } finally {
        lock.releaseLock();
      }
    } catch(err) { return fail(err.message); }
  }

    if (action === 'deleteZoom') {
        try {
            const who = e.parameter.who || '';
            const meetingId = e.parameter.meetingId || '';
            if (!who || !meetingId) return fail('Missing who or meetingId');
            deleteZoomMeeting(who, meetingId);
            return ok({});
        } catch (err) { return fail(err.message); }
    }

    if (action === 'sendZoomEmail') {
        try {
            const to = e.parameter.to || '';
            const firstName = e.parameter.firstName || 'there';
            const who = e.parameter.who || '';
            const date = e.parameter.date || '';
            const start = e.parameter.start || '';
            const joinUrl = e.parameter.joinUrl || '';
            const mobile = e.parameter.mobile || '';
            if (!to) return fail('No email address');
            const subject = 'Your Zoom session with ' + who + ' is confirmed';
            const html =
                '<div style="font-family:Arial,sans-serif;max-width:560px;margin:0 auto;color:#1A1814;">' +
                '<div style="background:#1A1814;padding:24px 28px;border-radius:8px 8px 0 0;">' +
                '<h2 style="color:#fff;margin:0;font-size:18px;">📹 Zoom Session Confirmed</h2>' +
                '</div>' +
                '<div style="background:#f9f8f6;padding:28px;border-radius:0 0 8px 8px;border:1px solid #E5E0D8;border-top:none;">' +
                '<p style="margin:0 0 8px;">Hi ' + firstName + ',</p>' +
                '<p style="margin:0 0 20px;">Your Zoom session with <strong>' + who + '</strong> is confirmed for:</p>' +
                '<div style="background:#fff;border:1px solid #E5E0D8;border-radius:8px;padding:16px 20px;margin-bottom:20px;">' +
                '<div style="font-size:20px;font-weight:bold;color:#1A1814;">' + date + '</div>' +
                '<div style="font-size:16px;color:#5C564E;margin-top:4px;">at ' + start + '</div>' +
                '</div>' +
                '<div style="background:#FFF8E7;border:1px solid #C9A84C;border-radius:8px;padding:14px 18px;margin-bottom:20px;">' +
                '<p style="margin:0;font-size:14px;color:#6B4C00;">⏰ <strong>Please join 5 minutes early</strong> so we can test your audio, microphone and speaker before we begin.</p>' +
                '</div>' +
                '<p style="margin:0 0 8px;font-size:14px;color:#5C564E;">Join the meeting here:</p>' +
                '<a href="' + joinUrl + '" style="display:inline-block;background:#1A1814;color:#fff;text-decoration:none;padding:12px 24px;border-radius:8px;font-weight:bold;font-size:14px;margin-bottom:20px;">Join Zoom Meeting</a>' +
                '<p style="margin:0 0 20px;font-size:13px;color:#8A8479;">Or copy this link: <a href="' + joinUrl + '" style="color:#1A1814;">' + joinUrl + '</a></p>' +
                '<hr style="border:none;border-top:1px solid #E5E0D8;margin:20px 0;">' +
                '<p style="margin:0;font-size:13px;color:#8A8479;">If you need to cancel please call <strong>' + mobile + '</strong>.</p>' +
                '<p style="margin:16px 0 0;font-size:14px;">Kind regards,<br><strong>' + who + '</strong></p>' +
                '</div>' +
                '</div>';
            const plain =
                'Hi ' + firstName + ',\n\n' +
                'Your Zoom session with ' + who + ' is confirmed for ' + date + ' at ' + start + '.\n\n' +
                'Please join 5 minutes early so we can test your audio, microphone and speaker before we begin.\n\n' +
                'Join the meeting here:\n' + joinUrl + '\n\n' +
                'If you need to cancel please call ' + mobile + '.\n\n' +
                'Kind regards,\n' + who;
            GmailApp.sendEmail(to, subject, plain, { htmlBody: html });
            return ok({ sent: true });
        } catch (err) { return fail(err.message); }
    }

    if (action === 'zoomical') {
        try {
            const who = e.parameter.who || '';
            var record = readRecord();
            const zoomBookings = (record.zoomBookings || []).filter(function (b) {
                return !who || b.who === who;
            });
            const lines = [
                'BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Room Booking Zoom//EN',
                'CALSCALE:GREGORIAN', 'METHOD:PUBLISH',
                'X-WR-CALNAME:Zoom Sessions' + (who ? ' - ' + who : ''),
                'X-WR-TIMEZONE:Europe/London',
                'REFRESH-INTERVAL;VALUE=DURATION:PT1H', 'X-PUBLISHED-TTL:PT1H'
            ];
            zoomBookings.forEach(function (b) {
                var joinUrl = b.zoomShortUrl || b.zoomJoinUrl || '';
                var pwd = b.zoomPassword ? '\nPassword: ' + b.zoomPassword : '';
                lines.push('BEGIN:VEVENT');
                lines.push('UID:zoom-' + b.id + '@roombooking.gas');
                lines.push('DTSTAMP:' + icsDateGAS(
                    Utilities.formatDate(new Date(), 'Europe/London', 'yyyy-MM-dd'),
                    Utilities.formatDate(new Date(), 'Europe/London', 'HH:mm')
                ));
                lines.push('DTSTART;TZID=Europe/London:' + icsDateGAS(b.date, b.start));
                lines.push('DTEND;TZID=Europe/London:' + icsDateGAS(b.date, b.end));
                lines.push('SUMMARY:' + icsEscapeGAS(b.topic || b.note || 'Zoom Session'));
                lines.push('DESCRIPTION:' + icsEscapeGAS('Client: ' + (b.note || '') + (joinUrl ? '\nJoin: ' + joinUrl : '') + pwd));
                lines.push('LOCATION:' + icsEscapeGAS(joinUrl));
                lines.push('URL:' + icsEscapeGAS(joinUrl));
                lines.push('STATUS:CONFIRMED');
                lines.push('END:VEVENT');
            });
            lines.push('END:VCALENDAR');
            return ContentService.createTextOutput(lines.join('\r\n'))
                .setMimeType(ContentService.MimeType.ICAL);
        } catch (err) { return fail(err.message); }
    }

    if (action === 'clearLog') {
        try {
            const lock = LockService.getScriptLock();
            lock.waitLock(10000);
            try {
                var rec = readRecord();
                const logId = parseInt(e.parameter.logId || '0');
                if (rec.smsLog) rec.smsLog = rec.smsLog.filter(function (l) { return l.id !== logId; });
                writeRecord(rec);
                return ok({});
            } finally {
                lock.releaseLock();
            }
        } catch (err) { return fail(err.message); }
    }

    return fail('Unknown action: ' + action);
}

function doPost(e) { return doGet(e); }

// ── TextMagic helpers ────────────────────────────────────────────────────
function tmAuth(e) {
    return { 'X-TM-Username': e.parameter.tmUser || '', 'X-TM-Key': e.parameter.tmKey || '' };
}
function tmGet(e, path) {
    return UrlFetchApp.fetch(TM + path, { method: 'GET', headers: tmAuth(e), muteHttpExceptions: true, followRedirects: true });
}
function tmPost(e, path, payload) {
    return UrlFetchApp.fetch(TM + path, {
        method: 'POST', headers: tmAuth(e),
        contentType: 'application/x-www-form-urlencoded', payload: payload,
        muteHttpExceptions: true, followRedirects: true
    });
}

// ── iCalendar helpers ────────────────────────────────────────────────────
function icsDateGAS(dateStr, timeStr) {
    var p = dateStr.split('-'), t = timeStr.split(':');
    return p[0] + p[1] + p[2] + 'T' + t[0] + t[1] + '00';
}
function icsEscapeGAS(str) {
    return String(str || '').split('\\').join('\\\\').split(';').join('\\;')
        .split(',').join('\\,').split('\n').join('\\n');
}

// ── Name matching ────────────────────────────────────────────────────────
function nameMatches(contact, clientName) {
    const searchNorm = clientName.trim().toLowerCase().replace(/\s+/g, ' ');
    const contactFull = ((contact.firstName || '') + ' ' + (contact.middleName || '') + ' ' + (contact.lastName || '')).trim().toLowerCase().replace(/\s+/g, ' ');
    const contactFirst = (contact.firstName || '').trim().toLowerCase();
    const contactLast = (contact.lastName || '').trim().toLowerCase();
    const searchParts = searchNorm.split(' ');
    const searchFirst = searchParts[0] || '';
    const searchLast = searchParts[searchParts.length - 1] || '';
    if (contactFull === searchNorm) return true;
    if (searchParts.length > 1 && contactFirst === searchFirst && contactLast === searchLast) return true;
    if (contactFull.includes(searchNorm) || searchNorm.includes(contactFull)) return true;
    return false;
}

// ── Daily auto-reminder trigger ──────────────────────────────────────────
function dailyReminders() {
    var record = readRecord();

    const props = PropertiesService.getScriptProperties();
    const bookings = record.bookings || [];
    const zoomBookings = record.zoomBookings || [];

    const today = new Date();
    const target = new Date(today);
    target.setDate(today.getDate() + 2);
    const targetStr = Utilities.formatDate(target, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    Logger.log('Daily reminders running for target date: ' + targetStr);

    const due = bookings.filter(function (b) {
        return b.date === targetStr && !b.smsSent && !b.autoSent && b.note && !b.isOnlineSession;
    });
    Logger.log('Room bookings due: ' + due.length);

    due.forEach(function (b) {
        const tmUser = props.getProperty('tm_user_' + b.who) || '';
        const tmKey = props.getProperty('tm_key_' + b.who) || '';
        const tmMobile = props.getProperty('tm_mobile_' + b.who) || '07498218609';
        if (!tmUser || !tmKey) { Logger.log('Skipping ' + b.who + ' - no TM credentials'); return; }

        var clients = (b.couples && b.client1 && b.client2) ? [b.client1, b.client2] : [b.note || ''];
        const d = Utilities.formatDate(new Date(b.date + 'T12:00:00'), Session.getScriptTimeZone(), 'EEE d MMM yyyy');

        clients.forEach(function (clientName) {
            if (!clientName) return;
            var clientPhone = '';
            var clientFirstName = clientName.split(' ')[0];
            try {
                const searchR = UrlFetchApp.fetch(TM + '/contacts/search?limit=5&page=1&query=' + encodeURIComponent(clientName), {
                    method: 'GET', headers: { 'X-TM-Username': tmUser, 'X-TM-Key': tmKey }, muteHttpExceptions: true
                });
                if (searchR.getResponseCode() === 200) {
                    const contacts = JSON.parse(searchR.getContentText()).resources || [];
                    const matchedContact = contacts.find(function (c) { return nameMatches(c, clientName); });
                    if (matchedContact) {
                        clientPhone = matchedContact.phone || '';
                        Logger.log('Matched: ' + clientPhone + ' for: ' + clientName);
                    } else {
                        Logger.log('No match for: ' + clientName + ' (' + contacts.length + ' results)');
                        return;
                    }
                }
            } catch (err) { Logger.log('Contact search error: ' + err.message); return; }
            if (!clientPhone) return;
            const text = 'Note: Hi ' + clientFirstName + ', you have an appt with ' + b.who +
                ' in Room 34, Canute Suite, Royal Mail House, Terminus Ter, Southampton SO14 3FD on ' +
                d + ' at ' + b.start + '. Call ' + tmMobile + '. Reply STOP to opt out.';
            try {
                const sendR = UrlFetchApp.fetch(TM + '/messages', {
                    method: 'POST', headers: { 'X-TM-Username': tmUser, 'X-TM-Key': tmKey },
                    contentType: 'application/x-www-form-urlencoded',
                    payload: { text: text, phones: clientPhone }, muteHttpExceptions: true
                });
                if (sendR.getResponseCode() === 201) {
                    b.autoSent = new Date().toISOString();
                    if (!record.smsLog) record.smsLog = [];
                    record.smsLog.unshift({
                        id: Date.now(), who: b.who, date: b.date, start: b.start,
                        clientName: clientName, phone: clientPhone, sentAt: new Date().toISOString(), auto: true, type: 'remind'
                    });
                    Logger.log('Sent reminder to ' + clientName);
                } else {
                    Logger.log('SMS failed for ' + clientName + ': HTTP ' + sendR.getResponseCode() + ' ' + sendR.getContentText().slice(0, 100));
                }
            } catch (err) { Logger.log('SMS error: ' + err.message); }
        });
        b.autoSent = new Date().toISOString();
    });

    // Prepayment bookings must not get their join link leaked via this
    // day-before reminder before payment is confirmed — the link is only
    // meant to go out once checkZettlePayments matches payment (or via
    // markZoomPaid). Non-prepay bookings, and prepay bookings already
    // paid, are unaffected by this check.
    const zoomDue = zoomBookings.filter(function (b) {
        const unpaidPrepay = b.requirePrepay && b.paymentStatus !== 'paid';
        return b.date === targetStr && !b.zoomReminderSent && b.zoomJoinUrl && !unpaidPrepay;
    });
    Logger.log('Zoom bookings due: ' + zoomDue.length);

    zoomDue.forEach(function (b) {
        const tmUser = props.getProperty('tm_user_' + b.who) || '';
        const tmKey = props.getProperty('tm_key_' + b.who) || '';
        const tmMobile = props.getProperty('tm_mobile_' + b.who) || '07498218609';
        if (!tmUser || !tmKey) { Logger.log('Skipping Zoom for ' + b.who); return; }

        var clients = (b.couples && b.client1 && b.client2) ? [b.client1, b.client2] : [b.note || ''];
        const d = Utilities.formatDate(new Date(b.date + 'T12:00:00'), Session.getScriptTimeZone(), 'EEE d MMM yyyy');
        const joinUrl = b.zoomShortUrl || b.zoomJoinUrl;

        clients.forEach(function (clientName) {
            if (!clientName) return;
            try {
                const searchR = UrlFetchApp.fetch(TM + '/contacts/search?limit=5&page=1&query=' + encodeURIComponent(clientName), {
                    method: 'GET', headers: { 'X-TM-Username': tmUser, 'X-TM-Key': tmKey }, muteHttpExceptions: true
                });
                if (searchR.getResponseCode() !== 200) return;
                const contacts = JSON.parse(searchR.getContentText()).resources || [];
                const contact = contacts.find(function (c) { return nameMatches(c, clientName); });
                if (!contact) { Logger.log('No match for Zoom client: ' + clientName); return; }
                const first = clientName.split(' ')[0];
                const phone = contact.phone || '';
                const email = contact.email || '';
                if (phone) {
                    const smsText = 'Note: Hi ' + first + ', your Zoom session with ' + b.who + ' is on ' + d + ' at ' + b.start +
                        '. Join: ' + joinUrl + '. Call ' + tmMobile + ' to cancel. Reply STOP to opt out.';
                    const smsR = UrlFetchApp.fetch(TM + '/messages', {
                        method: 'POST', headers: { 'X-TM-Username': tmUser, 'X-TM-Key': tmKey },
                        contentType: 'application/x-www-form-urlencoded',
                        payload: { text: smsText, phones: phone }, muteHttpExceptions: true
                    });
                    Logger.log('Zoom SMS HTTP ' + smsR.getResponseCode() + ' to ' + clientName);
                }
                if (email) {
                    var NL2 = '\n';
                    const subject = 'Reminder: Zoom session with ' + b.who + ' on ' + d;
                    const body2 = 'Hi ' + first + ',' + NL2 + NL2 +
                        'This is a reminder that your Zoom session with ' + b.who + ' is on ' + d + ' at ' + b.start + '.' + NL2 + NL2 +
                        'Join here: ' + joinUrl + NL2 + NL2 +
                        'To cancel call ' + tmMobile + '.' + NL2 + NL2 +
                        'Kind regards,' + NL2 + b.who;
                    try { GmailApp.sendEmail(email, subject, body2); Logger.log('Zoom email sent to ' + email); }
                    catch (ge) { Logger.log('Email error: ' + ge.message); }
                }
            } catch (err) { Logger.log('Zoom reminder error: ' + err.message); }
        });
        b.zoomReminderSent = new Date().toISOString();
    });

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
        record = trimRecord(record);
        writeRecord(record);
        Logger.log('Sheet updated and trimmed');
    } finally {
        lock.releaseLock();
    }
}

// ── Zettle payment integration ───────────────────────────────────────────
// Requires Script Properties: ZETTLE_CLIENT_ID, ZETTLE_API_KEY
// (Project Settings → Script Properties in the Apps Script editor)

// Fee schedule for prepay categories, in pounds. Concessionary is
// per-booking (client-specific), so it is NOT matched by this table —
// see zb.expectedAmount handling in checkZettlePayments().
const ZETTLE_PREPAY_FEES = {
  'Supervision':      65,
  'NHS Supervision':  52,
  'Private Practice': 59,
  'Couples':          94
};

function getZettleAccessToken() {
  const props = PropertiesService.getScriptProperties();
  const clientId = props.getProperty('ZETTLE_CLIENT_ID');
  const apiKey   = props.getProperty('ZETTLE_API_KEY');
  if (!clientId || !apiKey) throw new Error('Missing ZETTLE_CLIENT_ID or ZETTLE_API_KEY in Script Properties');

  const resp = UrlFetchApp.fetch('https://oauth.zettle.com/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: {
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      client_id: clientId,
      assertion: apiKey
    },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) {
    throw new Error('Zettle token request failed: ' + resp.getResponseCode() + ' ' + resp.getContentText().slice(0, 300));
  }
  const data = JSON.parse(resp.getContentText());
  return data.access_token;
}

// Fetches recent Zettle purchases (last 3 days is plenty for same-week matching)
function fetchRecentZettlePurchases(accessToken) {
  const startDate = Utilities.formatDate(
    new Date(Date.now() - 3 * 24 * 60 * 60 * 1000),
    'UTC', "yyyy-MM-dd'T'HH:mm:ss"
  );
  const resp = UrlFetchApp.fetch(
    'https://purchase.izettle.com/purchases/v2?limit=200&descending=true&startDate=' + encodeURIComponent(startDate),
    {
      method: 'get',
      headers: { 'Authorization': 'Bearer ' + accessToken },
      muteHttpExceptions: true
    }
  );
  if (resp.getResponseCode() !== 200) {
    throw new Error('Zettle purchase fetch failed: ' + resp.getResponseCode() + ' ' + resp.getContentText().slice(0, 300));
  }
  return JSON.parse(resp.getContentText()).purchases || [];
}

// Main poller — run this on a time-driven trigger every 10-15 minutes.
function checkZettlePayments() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var record = readRecord();
    var zoomBookings = record.zoomBookings || [];

    // Bookings still waiting on payment
    var awaiting = zoomBookings.filter(function(b) {
      return b.isZoom && b.paymentStatus === 'awaiting';
    });
    if (!awaiting.length) { Logger.log('No bookings awaiting payment.'); return; }

    // ── Chase reminders (max 3, ~24h apart, 8am-8pm only) ────────────
    // Only for bookings where the initial payment-link email has
    // actually been sent (paymentLinkSentAt set) and the link itself
    // was saved. Chases stop permanently after 3 — no further emails,
    // the booking just sits flagged for you to handle manually.
    var currentHour = parseInt(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'H'));
    var withinSendingHours = currentHour >= 8 && currentHour < 20;

    awaiting.forEach(function(b) {
      if (!b.paymentLinkSentAt || !b.paymentLink) return; // link never sent yet
      var chaseCount = b.chaseCount || 0;
      if (chaseCount >= 3) return; // exhausted — no more automatic emails
      var lastSent = b.lastChaseSentAt || b.paymentLinkSentAt;
      var hoursSince = (Date.now() - new Date(lastSent).getTime()) / 3600000;
      if (hoursSince < 24) return; // not due yet
      if (!withinSendingHours) return; // due, but outside 8am-8pm — wait for a run inside the window

      var tmMobile = PropertiesService.getScriptProperties().getProperty('tm_mobile_' + b.who) || '07498218609';
      var firstName = (b.note || b.client1 || '').split(' ')[0] || 'there';
      var d = Utilities.formatDate(new Date(b.date + 'T12:00:00'), Session.getScriptTimeZone(), 'EEE d MMM yyyy');
      var nextChaseNumber = chaseCount + 1;

      try {
        sendPaymentRequestEmailInternal({
          to: b.clientEmail, firstName: firstName, who: b.who,
          date: d, start: b.start, paymentLink: b.paymentLink,
          amount: b.expectedAmount, mobile: tmMobile,
          reminderNumber: nextChaseNumber
        });
        b.chaseCount = nextChaseNumber;
        b.lastChaseSentAt = new Date().toISOString();
        Logger.log('Sent chase reminder '+nextChaseNumber+'/3 for booking '+b.id);
        if (nextChaseNumber >= 3) {
          b.paymentChaseExhausted = true;
          Logger.log('Booking '+b.id+' has now had 3 reminders with no payment — no further automatic emails will be sent.');
        }
      } catch (e) {
        Logger.log('Chase reminder failed for booking '+b.id+': '+e.message);
      }
    });

    var accessToken = getZettleAccessToken();
    var purchases = fetchRecentZettlePurchases(accessToken);

    // Only look at real, non-refund payments
    var candidates = purchases.filter(function(p) { return !p.refund && p.amount > 0; });

    // Track which purchase UUIDs already matched a booking this run,
    // so two bookings never claim the same purchase.
    var claimed = {};

    awaiting.forEach(function(b) {
      var expected = b.expectedAmount; // pounds, set at booking creation
      if (!expected) { Logger.log('Booking ' + b.id + ' has no expectedAmount, skipping'); return; }
      var expectedMinor = Math.round(expected * 100);

      var matches = candidates.filter(function(p) {
        return !claimed[p.purchaseUUID1 || p.purchaseUUID] && p.amount === expectedMinor;
      });

      if (matches.length === 0) {
        return; // no payment yet, keep waiting
      }
      if (matches.length > 1) {
        // Ambiguous: two+ unclaimed purchases at the same amount.
        // Flag for manual confirmation rather than guessing.
        b.paymentAmbiguous = true;
        b.paymentCandidates = matches.map(function(p) {
          return { uuid: p.purchaseUUID1 || p.purchaseUUID, timestamp: p.timestamp, amount: p.amount };
        });
        Logger.log('Ambiguous payment match for booking ' + b.id + ' (' + matches.length + ' candidates)');
        return;
      }

      // Exactly one match — confirm it
      var purchase = matches[0];
      var uuid = purchase.purchaseUUID1 || purchase.purchaseUUID;
      claimed[uuid] = true;

      b.paymentStatus = 'paid';
      b.paidAt = new Date().toISOString();
      b.zettlePurchaseUuid = uuid;
      delete b.paymentAmbiguous;
      delete b.paymentCandidates;

      Logger.log('Matched payment for booking ' + b.id + ' -> purchase ' + uuid);

      // Fire the Zoom confirmation email now that payment is confirmed
      if (b.clientEmail) {
        try {
          sendZoomConfirmationEmailInternal(b);
          b.confirmationSentAt = new Date().toISOString();
        } catch (e) {
          Logger.log('Failed to send confirmation for booking ' + b.id + ': ' + e.message);
        }
      } else {
        Logger.log('No clientEmail on booking ' + b.id + ', cannot auto-send confirmation');
      }
    });

    record = trimRecord(record);
    writeRecord(record);
  } finally {
    lock.releaseLock();
  }
}

// Shared confirmation-email sender, usable both from the createZoom action
// (invoice-later categories, sent immediately) and from checkZettlePayments
// (prepay categories, sent once payment is matched).
function sendZoomConfirmationEmailInternal(b) {
  var firstName = (b.note || '').split(' ')[0] || 'there';
  var d = Utilities.formatDate(new Date(b.date + 'T12:00:00'), Session.getScriptTimeZone(), 'EEE d MMM yyyy');
  var joinUrl = b.zoomShortUrl || b.zoomJoinUrl;
  var tmMobile = PropertiesService.getScriptProperties().getProperty('tm_mobile_' + b.who) || '07498218609';

  var subject = 'Your Zoom session with ' + b.who + ' is confirmed';
  var html =
    '<div style="font-family:Arial,sans-serif;max-width:560px;margin:0 auto;color:#1A1814;">' +
    '<div style="background:#1A1814;padding:24px 28px;border-radius:8px 8px 0 0;">' +
    '<h2 style="color:#fff;margin:0;font-size:18px;">📹 Zoom Session Confirmed</h2>' +
    '</div>' +
    '<div style="background:#f9f8f6;padding:28px;border-radius:0 0 8px 8px;border:1px solid #E5E0D8;border-top:none;">' +
    '<p style="margin:0 0 8px;">Hi ' + firstName + ',</p>' +
    '<p style="margin:0 0 20px;">Your Zoom session with <strong>' + b.who + '</strong> is confirmed for:</p>' +
    '<div style="background:#fff;border:1px solid #E5E0D8;border-radius:8px;padding:16px 20px;margin-bottom:20px;">' +
    '<div style="font-size:20px;font-weight:bold;color:#1A1814;">' + d + '</div>' +
    '<div style="font-size:16px;color:#5C564E;margin-top:4px;">at ' + b.start + '</div>' +
    '</div>' +
    '<div style="background:#FFF8E7;border:1px solid #C9A84C;border-radius:8px;padding:14px 18px;margin-bottom:20px;">' +
    '<p style="margin:0;font-size:14px;color:#6B4C00;">⏰ <strong>Please join 5 minutes early</strong> so we can test your audio, microphone and speaker before we begin.</p>' +
    '</div>' +
    '<p style="margin:0 0 8px;font-size:14px;color:#5C564E;">Join the meeting here:</p>' +
    '<a href="' + joinUrl + '" style="display:inline-block;background:#1A1814;color:#fff;text-decoration:none;padding:12px 24px;border-radius:8px;font-weight:bold;font-size:14px;margin-bottom:20px;">Join Zoom Meeting</a>' +
    '<p style="margin:0 0 20px;font-size:13px;color:#8A8479;">Or copy this link: <a href="' + joinUrl + '" style="color:#1A1814;">' + joinUrl + '</a></p>' +
    '<hr style="border:none;border-top:1px solid #E5E0D8;margin:20px 0;">' +
    '<p style="margin:0;font-size:13px;color:#8A8479;">If you need to cancel please call <strong>' + tmMobile + '</strong>.</p>' +
    '<p style="margin:16px 0 0;font-size:14px;">Kind regards,<br><strong>' + b.who + '</strong></p>' +
    '</div>' +
    '</div>';
  var plain =
    'Hi ' + firstName + ',\n\n' +
    'Your Zoom session with ' + b.who + ' is confirmed for ' + d + ' at ' + b.start + '.\n\n' +
    'Please join 5 minutes early so we can test your audio, microphone and speaker before we begin.\n\n' +
    'Join the meeting here:\n' + joinUrl + '\n\n' +
    'If you need to cancel please call ' + tmMobile + '.\n\n' +
    'Kind regards,\n' + b.who;

  GmailApp.sendEmail(b.clientEmail, subject, plain, { htmlBody: html });
}

// Shared HTML/plain-text builder for the payment-request email, used both
// by the manual "Send payment link" action (sendPaymentLinkEmail, inside
// doGetInner) and by the automated chase reminders in checkZettlePayments.
// Must live at the TOP LEVEL (not nested inside doGetInner) — otherwise it
// is out of scope for checkZettlePayments, which runs as its own top-level
// function from the time-driven trigger, not through doGetInner at all.
// reminderNumber: 0 = initial request, 1-3 = chase reminders (wording
// escalates slightly).
function sendPaymentRequestEmailInternal(p) {
  const isChase = p.reminderNumber > 0;
  const subject = isChase
    ? 'Reminder '+p.reminderNumber+' of 3: payment still required for your session with '+p.who
    : 'Payment required for your session with '+p.who;
  const intro = isChase
    ? 'This is a reminder ('+p.reminderNumber+' of 3) that we still haven\'t received payment for your session with <strong>'+p.who+'</strong> on <strong>'+p.date+' at '+p.start+'</strong>'+(p.amount?' (£'+p.amount+')':'')+'. Please complete payment using the link below — your Zoom joining details will follow as soon as payment is received.'
    : 'To confirm your session with <strong>'+p.who+'</strong> on <strong>'+p.date+' at '+p.start+'</strong>'+(p.amount?' (£'+p.amount+')':'')+', please complete payment using the link below. Your Zoom joining details will follow as soon as payment is received.';
  const introPlain = isChase
    ? 'This is a reminder ('+p.reminderNumber+' of 3) that we still haven\'t received payment for your session with '+p.who+' on '+p.date+' at '+p.start+(p.amount?' (£'+p.amount+')':'')+'. Please complete payment using the link below.'
    : 'To confirm your session with '+p.who+' on '+p.date+' at '+p.start+(p.amount?' (£'+p.amount+')':'')+', please complete payment using the link below. Your Zoom joining details will follow as soon as payment is received.';
  const html =
    '<div style="font-family:Arial,sans-serif;max-width:560px;margin:0 auto;color:#1A1814;">'+
    '<div style="background:#1A1814;padding:24px 28px;border-radius:8px 8px 0 0;">'+
    '<h2 style="color:#fff;margin:0;font-size:18px;">'+(isChase?'⏰ Payment Reminder':'💳 Payment Required')+'</h2>'+
    '</div>'+
    '<div style="background:#f9f8f6;padding:28px;border-radius:0 0 8px 8px;border:1px solid #E5E0D8;border-top:none;">'+
    '<p style="margin:0 0 8px;">Hi '+p.firstName+',</p>'+
    '<p style="margin:0 0 20px;">'+intro+'</p>'+
    '<a href="'+p.paymentLink+'" style="display:inline-block;background:#1A1814;color:#fff;text-decoration:none;padding:12px 24px;border-radius:8px;font-weight:bold;font-size:14px;margin-bottom:20px;">Pay now</a>'+
    '<p style="margin:0 0 20px;font-size:13px;color:#8A8479;">Or copy this link: <a href="'+p.paymentLink+'" style="color:#1A1814;">'+p.paymentLink+'</a></p>'+
    '<hr style="border:none;border-top:1px solid #E5E0D8;margin:20px 0;">'+
    '<p style="margin:0;font-size:13px;color:#8A8479;">If you need to cancel please call <strong>'+p.mobile+'</strong> — please note cancellations within 24 hours of the appointment are non-refundable.</p>'+
    '<p style="margin:16px 0 0;font-size:14px;">Kind regards,<br><strong>'+p.who+'</strong></p>'+
    '</div>'+
    '</div>';
  const plain =
    'Hi '+p.firstName+',\n\n'+
    introPlain+'\n\n'+
    p.paymentLink+'\n\n'+
    'If you need to cancel please call '+p.mobile+' — please note cancellations within 24 hours of the appointment are non-refundable.\n\n'+
    'Kind regards,\n'+p.who;
  GmailApp.sendEmail(p.to, subject, plain, {htmlBody: html});
}

// ── Zoom helpers ─────────────────────────────────────────────────────────
function setZoomCredentials(name, accountId, clientId, clientSecret) {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('zoom_account_' + name, accountId);
    props.setProperty('zoom_client_' + name, clientId);
    props.setProperty('zoom_secret_' + name, clientSecret);
    Logger.log('Zoom credentials saved for ' + name);
}

function setZoomEmail(name, email) {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('zoom_email_' + name, email);
    Logger.log('Zoom email saved for ' + name + ': ' + email);
}

function setBitlyToken(name, token) {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('bitly_token_' + name, token);
    Logger.log('Bitly token saved for ' + name);
}

function formatZoomDateTime(date, startTime) {
    var parts = date.split('-');
    var timeParts = startTime.split(':');
    return parts[0] + '-' + ('0' + parts[1]).slice(-2) + '-' + ('0' + parts[2]).slice(-2) + 'T' +
        ('0' + parseInt(timeParts[0])).slice(-2) + ':' + ('0' + parseInt(timeParts[1])).slice(-2) + ':00';
}

function getZoomToken(name) {
    const props = PropertiesService.getScriptProperties();
    const accountId = props.getProperty('zoom_account_' + name) || '';
    const clientId = props.getProperty('zoom_client_' + name) || '';
    const clientSecret = props.getProperty('zoom_secret_' + name) || '';
    if (!accountId || !clientId || !clientSecret) throw new Error('No Zoom credentials for ' + name);
    const encoded = Utilities.base64Encode(clientId + ':' + clientSecret);
    const r = UrlFetchApp.fetch(
        'https://zoom.us/oauth/token?grant_type=account_credentials&account_id=' + accountId,
        { method: 'POST', headers: { 'Authorization': 'Basic ' + encoded }, muteHttpExceptions: true }
    );
    if (r.getResponseCode() !== 200) throw new Error('Zoom token failed: ' + r.getContentText());
    return JSON.parse(r.getContentText()).access_token;
}

function shortenUrl(name, longUrl) {
    const props = PropertiesService.getScriptProperties();
    const token = props.getProperty('bitly_token_' + name) || '';
    if (!token) return longUrl;
    try {
        const r = UrlFetchApp.fetch('https://api-ssl.bitly.com/v4/shorten', {
            method: 'POST',
            headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
            payload: JSON.stringify({ long_url: longUrl }),
            muteHttpExceptions: true
        });
        if (r.getResponseCode() === 200 || r.getResponseCode() === 201) {
            return JSON.parse(r.getContentText()).link || longUrl;
        }
    } catch (e) { Logger.log('Bitly error: ' + e.message); }
    return longUrl;
}

function createZoomMeeting(name, topic, date, startTime, durationMins) {
    const token = getZoomToken(name);
    const dt = formatZoomDateTime(date, startTime);
    const payload = {
        topic: topic, type: 2,
        start_time: dt, duration: durationMins,
        timezone: 'Europe/London',
        settings: { host_video: true, participant_video: true, waiting_room: true, auto_recording: 'none' }
    };
    const props2 = PropertiesService.getScriptProperties();
    const userEmail = props2.getProperty('zoom_email_' + name) || 'me';
    const r = UrlFetchApp.fetch('https://api.zoom.us/v2/users/' + userEmail + '/meetings', {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    });
    const code = r.getResponseCode();
    if (code !== 201) throw new Error('Zoom creation failed (' + code + '): ' + r.getContentText());
    const meeting = JSON.parse(r.getContentText());
    const shortUrl = shortenUrl(name, meeting.join_url);
    return {
        meetingId: meeting.id,
        joinUrl: meeting.join_url,
        shortUrl: shortUrl,
        password: meeting.password || '',
        startUrl: meeting.start_url
    };
}

function deleteZoomMeeting(name, meetingId) {
    try {
        const token = getZoomToken(name);
        UrlFetchApp.fetch('https://api.zoom.us/v2/meetings/' + meetingId, {
            method: 'DELETE', headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true
        });
        Logger.log('Zoom meeting ' + meetingId + ' deleted');
    } catch (e) { Logger.log('Zoom delete error: ' + e.message); }
}

// ── Credentials ──────────────────────────────────────────────────────────
function setCredentials(name, tmUser, tmKey, mobileNumber) {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('tm_user_' + name, tmUser);
    props.setProperty('tm_key_' + name, tmKey);
    props.setProperty('tm_mobile_' + name, mobileNumber);
    Logger.log('Credentials saved for ' + name);
}

// ── Keep alive ───────────────────────────────────────────────────────────
function keepAlive() {
    getSheet();
}

// ── Manual trim ──────────────────────────────────────────────────────────
function trimSheetData() {
    var record = readRecord();
    const beforeBookings = (record.bookings || []).length;
    const beforeZoom = (record.zoomBookings || []).length;
    record = trimRecord(record);
    Logger.log('Bookings: ' + beforeBookings + ' -> ' + (record.bookings || []).length);
    Logger.log('Zoom bookings: ' + beforeZoom + ' -> ' + (record.zoomBookings || []).length);
    const json = JSON.stringify(record);
    Logger.log('New total size: ' + json.length + ' chars (now split across cells, no longer capped at 50,000)');
    writeRecord(record);
    Logger.log('Done');
}

// ── Test helpers ─────────────────────────────────────────────────────────
function testDailyReminders() {
    Logger.log('=== Testing dailyReminders ===');
    dailyReminders();
    Logger.log('=== Done ===');
}

function testZoom() {
    try {
        const token = getZoomToken('Rosswell');
        Logger.log('Zoom token: ' + token.substring(0, 20) + '...');
        const meeting = createZoomMeeting('Rosswell', 'Test Meeting',
            Utilities.formatDate(new Date(), 'Europe/London', 'yyyy-MM-dd'), '09:00', 30);
        Logger.log('Test meeting: ' + meeting.joinUrl);
    } catch (e) { Logger.log('Error: ' + e.message); }
}

function debugNameMatch() {
    const props = PropertiesService.getScriptProperties();
    const tmUser = props.getProperty('tm_user_Rosswell') || '';
    const tmKey = props.getProperty('tm_key_Rosswell') || '';
    const names = ['Vanessa Loubier', 'Heather Walker', 'Maggie Lomax'];
    names.forEach(function (clientName) {
        const r = UrlFetchApp.fetch(
            TM + '/contacts/search?limit=5&page=1&query=' + encodeURIComponent(clientName),
            { method: 'GET', headers: { 'X-TM-Username': tmUser, 'X-TM-Key': tmKey }, muteHttpExceptions: true }
        );
        Logger.log('HTTP ' + r.getResponseCode() + ' for: "' + clientName + '"');
        Logger.log(r.getContentText());
    });
}
function testInvoiceFilter() {
  const all = getAllBookings_();
  Logger.log('Total bookings: ' + all.length);
  Logger.log('Sample booking: ' + JSON.stringify(all[0]));
  Logger.log('Helen bookings (any month): ' + JSON.stringify(all.filter(function(b){ return b.practitioner === 'Helen'; }).slice(0,3)));
  const julyHelen = getBookingsForInvoice('Helen', 2026, 7, 'All', 'All');
  Logger.log('July 2026 Helen via getBookingsForInvoice: ' + julyHelen.length);
}
