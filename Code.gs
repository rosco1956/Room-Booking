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
      bookings:[], people:['Rosswell','Donna','Helen']
    }));
  }
  return sheet;
}

function corsOutput(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function ok(data)  { return corsOutput(Object.assign({ok:true}, data)); }
function fail(msg) { return corsOutput({ok:false, error:msg}); }

// ── Trim helpers ─────────────────────────────────────────────────────────
function trimRecord(record) {
  const todayStr = new Date().toISOString().slice(0,10);
  if (record.zoomBookings) {
    record.zoomBookings = record.zoomBookings.filter(function(b){ return b.date >= todayStr; });
  }
  if (record.smsLog) record.smsLog = record.smsLog.slice(0,20);
  return record;
}

// ── Main GET handler ─────────────────────────────────────────────────────
function doGet(e) {
  if (!e || !e.parameter) return ok({status:'alive'});
  const action = (e.parameter.action || 'load');

  if (action === 'load') {
    try {
      const raw = getSheet().getRange('A1').getValue();
      var record = {};
      try { record = JSON.parse(raw||'{}'); } catch(err) { record = {bookings:[],people:[]}; }
      return ok({record: record});
    } catch(err) { return fail(err.message); }
  }

  if (action === 'addBooking') {
    try {
      const sheet = getSheet();
      const raw = sheet.getRange('A1').getValue();
      var record = {};
      try { record = JSON.parse(raw||'{}'); } catch(e2) { record = {bookings:[]}; }
      if (!record.bookings) record.bookings = [];
      const newBooking = JSON.parse(decodeURIComponent(e.parameter.booking || '{}'));
      const clash = record.bookings.find(function(b) {
        return b.date === newBooking.date &&
               !(newBooking.end <= b.start || newBooking.start >= b.end);
      });
      if (clash) return fail('Clash with '+clash.who+' ('+clash.start+'-'+clash.end+')');
      record.bookings.push(newBooking);
      sheet.getRange('A1').setValue(JSON.stringify(record));
      return ok({count: record.bookings.length, record: {bookings: record.bookings}});
    } catch(err) { return fail(err.message); }
  }

  if (action === 'deleteBooking') {
    try {
      const sheet = getSheet();
      const raw = sheet.getRange('A1').getValue();
      var record = {};
      try { record = JSON.parse(raw||'{}'); } catch(e2) { record = {bookings:[]}; }
      const delId = parseInt(e.parameter.id || '0');
      const before = (record.bookings||[]).length;
      record.bookings = (record.bookings||[]).filter(function(b){ return b.id !== delId; });
      sheet.getRange('A1').setValue(JSON.stringify(record));
      return ok({count: record.bookings.length, deleted: before - record.bookings.length});
    } catch(err) { return fail(err.message); }
  }

  if (action === 'updateBooking') {
    try {
      const sheet = getSheet();
      const raw = sheet.getRange('A1').getValue();
      var record = {};
      try { record = JSON.parse(raw||'{}'); } catch(e2) { record = {bookings:[]}; }
      const updated = JSON.parse(decodeURIComponent(e.parameter.booking || '{}'));
      const clash = (record.bookings||[]).find(function(b) {
        return b.id !== updated.id && b.date === updated.date &&
               !(updated.end <= b.start || updated.start >= b.end);
      });
      if (clash) return fail('Clash with '+clash.who+' ('+clash.start+'-'+clash.end+')');
      record.bookings = (record.bookings||[]).map(function(b){
        return b.id === updated.id ? updated : b;
      });
      sheet.getRange('A1').setValue(JSON.stringify(record));
      return ok({count: record.bookings.length});
    } catch(err) { return fail(err.message); }
  }

  if (action === 'saveSettings') {
    try {
      const sheet = getSheet();
      const raw = sheet.getRange('A1').getValue();
      var record = {};
      try { record = JSON.parse(raw||'{}'); } catch(e2) { record = {}; }
      const settings = JSON.parse(decodeURIComponent(e.parameter.data || '{}'));
      if (settings.people)       record.people       = settings.people;
      if (settings.rate!=null)   record.rate         = settings.rate;
      if (settings.smsLog)       record.smsLog       = settings.smsLog;
      if (settings.waitingList)  record.waitingList  = settings.waitingList;
      if (settings.appPin)       record.appPin       = settings.appPin;
      if (settings.zoomBookings) record.zoomBookings = settings.zoomBookings;
      record = trimRecord(record);
      sheet.getRange('A1').setValue(JSON.stringify(record));
      return ok({});
    } catch(err) { return fail(err.message); }
  }

  if (action === 'save') {
    try {
      const body = JSON.parse(decodeURIComponent(e.parameter.data || '{}'));
      getSheet().getRange('A1').setValue(JSON.stringify(body));
      return ok({});
    } catch(err) { return fail(err.message); }
  }

  if (action === 'tmMe') {
    try {
      const r = tmGet(e, '/user');
      const d = JSON.parse(r.getContentText());
      const code = r.getResponseCode();
      return code===200
        ? ok({firstName:d.firstName, lastName:d.lastName, balance:d.balance})
        : fail('HTTP '+code+': '+JSON.stringify(d));
    } catch(err) { return fail(err.message); }
  }

  if (action === 'tmContacts') {
    try {
      const search = e.parameter.search || '';
      const limit  = e.parameter.limit  || 50;
      var path = '/contacts/search?limit='+limit+'&page=1';
      if (search) path += '&query='+encodeURIComponent(search);
      const r = tmGet(e, path);
      const d = JSON.parse(r.getContentText());
      return r.getResponseCode()===200
        ? ok({resources: d.resources||[]})
        : fail(d.message||'Could not fetch contacts');
    } catch(err) { return fail(err.message); }
  }

  if (action === 'tmSend') {
    try {
      const text   = e.parameter.text   || '';
      const phones = e.parameter.phones || '';
      if (!text||!phones) return fail('Missing text or phone');
      const r = tmPost(e, '/messages', {text:text, phones:phones});
      const d = JSON.parse(r.getContentText());
      const code = r.getResponseCode();
      Logger.log('tmSend HTTP '+code+': '+r.getContentText().slice(0,200));
      return code===201
        ? ok({id:d.id})
        : fail('HTTP '+code+': '+(d.message||d.error||JSON.stringify(d)));
    } catch(err) { return fail(err.message); }
  }

  if (action === 'tmContactEmail') {
    try {
      const search = e.parameter.search||'';
      const limit  = e.parameter.limit||5;
      var path = '/contacts/search?limit='+limit+'&page=1';
      if (search) path += '&query='+encodeURIComponent(search);
      const r = tmGet(e, path);
      const d = JSON.parse(r.getContentText());
      if (r.getResponseCode()!==200) return fail(d.message||'Failed');
      const contacts = (d.resources||[])
        .filter(function(c){ return nameMatches(c, search); })
        .map(function(c){
          return {
            name:  ((c.firstName||'')+' '+(c.lastName||'')).trim(),
            phone: c.phone||'',
            email: c.email||''
          };
        });
      return ok({contacts:contacts});
    } catch(err) { return fail(err.message); }
  }

  if (action === 'tmIncrementAppts') {
    try {
      const contactId = e.parameter.contactId||'';
      if (!contactId) return fail('No contactId');
      const r = tmGet(e, '/contacts/'+contactId);
      if (r.getResponseCode()!==200) return fail('Contact not found');
      const contact = JSON.parse(r.getContentText());
      const fields = contact.customFields||[];
      const apptField = fields.find(function(f){ return f.name==='Appointments'; });
      const current = apptField ? parseInt(apptField.value||'0') : 0;
      const newCount = current+1;
      if (apptField) {
        var apptPayload = {};
        apptPayload['customFieldValues['+apptField.userCustomFieldId+']'] = String(newCount);
        tmPost(e, '/contacts/'+contactId, apptPayload);
      }
      return ok({appointments: newCount});
    } catch(err) { return fail(err.message); }
  }

  if (action === 'ical') {
    try {
      const who = e.parameter.who || '';
      const raw = getSheet().getRange('A1').getValue();
      var record = {};
      try { record = JSON.parse(raw||'{}'); } catch(err) { record = {bookings:[]}; }
      const bookings = (record.bookings||[]).filter(function(b){
        return !who || b.who === who;
      });
      const lines = [
        'BEGIN:VCALENDAR','VERSION:2.0','PRODID:-//Room Booking//EN',
        'CALSCALE:GREGORIAN','METHOD:PUBLISH',
        'X-WR-CALNAME:Room Booking'+(who?' - '+who:''),
        'X-WR-TIMEZONE:Europe/London',
        'REFRESH-INTERVAL;VALUE=DURATION:PT1H','X-PUBLISHED-TTL:PT1H'
      ];
      bookings.forEach(function(b){
        lines.push('BEGIN:VEVENT');
        lines.push('UID:'+b.id+'@roombooking.gas');
        lines.push('DTSTAMP:'+icsDateGAS(
          Utilities.formatDate(new Date(),'Europe/London','yyyy-MM-dd'),
          Utilities.formatDate(new Date(),'Europe/London','HH:mm')
        ));
        lines.push('DTSTART;TZID=Europe/London:'+icsDateGAS(b.date,b.start));
        lines.push('DTEND;TZID=Europe/London:'+icsDateGAS(b.date,b.end));
        lines.push('SUMMARY:'+icsEscapeGAS(b.who+' - Room 34'+(b.note?' ('+b.note+')':'')));
        lines.push('LOCATION:'+icsEscapeGAS('Room 34, Canute Suite, Royal Mail House, Terminus Ter, Southampton SO14 3FD'));
        lines.push('DESCRIPTION:'+icsEscapeGAS(b.note?'Client: '+b.note:'Room booking for '+b.who));
        lines.push('STATUS:CONFIRMED');
        lines.push('END:VEVENT');
      });
      lines.push('END:VCALENDAR');
      return ContentService.createTextOutput(lines.join('\r\n'))
        .setMimeType(ContentService.MimeType.ICAL);
    } catch(err) { return fail(err.message); }
  }

  if (action === 'shortenUrl') {
    try {
      const who     = e.parameter.who||'';
      const longUrl = decodeURIComponent(e.parameter.url||'');
      if (!longUrl) return fail('No URL provided');
      const shortUrl = shortenUrl(who, longUrl);
      return ok({shortUrl: shortUrl});
    } catch(err) { return fail(err.message); }
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
      if (!who||!date) return fail('Missing who or date');

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
        updated:      new Date().toISOString()
      };

      const lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        const sheet = getSheet();
        const raw2 = sheet.getRange('A1').getValue();
        var record2 = {};
        try { record2 = JSON.parse(raw2||'{}'); } catch(pe) { record2 = {}; }
        if (!record2.zoomBookings) record2.zoomBookings = [];
        record2.zoomBookings.push(zb);
        record2 = trimRecord(record2);
        sheet.getRange('A1').setValue(JSON.stringify(record2));
        Logger.log('Saved to sheet, pendingId: '+pendingId);
      } finally {
        lock.releaseLock();
      }
      return ok(zb);
    } catch(err) {
      Logger.log('createZoom error: '+err.message);
      return fail(err.message);
    }
  }

  if (action === 'deleteZoom') {
    try {
      const who       = e.parameter.who||'';
      const meetingId = e.parameter.meetingId||'';
      if (!who||!meetingId) return fail('Missing who or meetingId');
      deleteZoomMeeting(who, meetingId);
      return ok({});
    } catch(err) { return fail(err.message); }
  }

  if (action === 'sendZoomEmail') {
    try {
      const to        = e.parameter.to||'';
      const firstName = e.parameter.firstName||'there';
      const who       = e.parameter.who||'';
      const date      = e.parameter.date||'';
      const start     = e.parameter.start||'';
      const joinUrl   = e.parameter.joinUrl||'';
      const mobile    = e.parameter.mobile||'';
      if (!to) return fail('No email address');
      const subject = 'Your Zoom session with '+who+' is confirmed';
      const NL = '\n';
      const body = 'Hi '+firstName+','+NL+NL+
        'Your Zoom session with '+who+' is confirmed for '+date+' at '+start+'.'+NL+NL+
        'Join the meeting here:'+NL+joinUrl+NL+NL+
        'If you need to cancel please call '+mobile+'.'+NL+NL+
        'Kind regards,'+NL+who;
      GmailApp.sendEmail(to, subject, body);
      return ok({sent:true});
    } catch(err) { return fail(err.message); }
  }

  if (action === 'zoomical') {
    try {
      const who = e.parameter.who||'';
      const raw = getSheet().getRange('A1').getValue();
      var record = {};
      try { record = JSON.parse(raw||'{}'); } catch(err) { record = {zoomBookings:[]}; }
      const zoomBookings = (record.zoomBookings||[]).filter(function(b){
        return !who || b.who === who;
      });
      const lines = [
        'BEGIN:VCALENDAR','VERSION:2.0','PRODID:-//Room Booking Zoom//EN',
        'CALSCALE:GREGORIAN','METHOD:PUBLISH',
        'X-WR-CALNAME:Zoom Sessions'+(who?' - '+who:''),
        'X-WR-TIMEZONE:Europe/London',
        'REFRESH-INTERVAL;VALUE=DURATION:PT1H','X-PUBLISHED-TTL:PT1H'
      ];
      zoomBookings.forEach(function(b){
        var joinUrl = b.zoomShortUrl||b.zoomJoinUrl||'';
        var pwd = b.zoomPassword?'\nPassword: '+b.zoomPassword:'';
        lines.push('BEGIN:VEVENT');
        lines.push('UID:zoom-'+b.id+'@roombooking.gas');
        lines.push('DTSTAMP:'+icsDateGAS(
          Utilities.formatDate(new Date(),'Europe/London','yyyy-MM-dd'),
          Utilities.formatDate(new Date(),'Europe/London','HH:mm')
        ));
        lines.push('DTSTART;TZID=Europe/London:'+icsDateGAS(b.date,b.start));
        lines.push('DTEND;TZID=Europe/London:'+icsDateGAS(b.date,b.end));
        lines.push('SUMMARY:'+icsEscapeGAS(b.topic||b.note||'Zoom Session'));
        lines.push('DESCRIPTION:'+icsEscapeGAS('Client: '+(b.note||'')+(joinUrl?'\nJoin: '+joinUrl:'')+pwd));
        lines.push('LOCATION:'+icsEscapeGAS(joinUrl));
        lines.push('URL:'+icsEscapeGAS(joinUrl));
        lines.push('STATUS:CONFIRMED');
        lines.push('END:VEVENT');
      });
      lines.push('END:VCALENDAR');
      return ContentService.createTextOutput(lines.join('\r\n'))
        .setMimeType(ContentService.MimeType.ICAL);
    } catch(err) { return fail(err.message); }
  }

  if (action === 'clearLog') {
    try {
      const sheet = getSheet();
      const raw = sheet.getRange('A1').getValue();
      var rec = {};
      try { rec = JSON.parse(raw||'{}'); } catch(e2) { rec = {}; }
      const logId = parseInt(e.parameter.logId||'0');
      if (rec.smsLog) rec.smsLog = rec.smsLog.filter(function(l){ return l.id !== logId; });
      sheet.getRange('A1').setValue(JSON.stringify(rec));
      return ok({});
    } catch(err) { return fail(err.message); }
  }

  return fail('Unknown action: '+action);
}

function doPost(e) { return doGet(e); }

// ── TextMagic helpers ────────────────────────────────────────────────────
function tmAuth(e) {
  return {'X-TM-Username':e.parameter.tmUser||'','X-TM-Key':e.parameter.tmKey||''};
}
function tmGet(e, path) {
  return UrlFetchApp.fetch(TM+path,{method:'GET',headers:tmAuth(e),muteHttpExceptions:true,followRedirects:true});
}
function tmPost(e, path, payload) {
  return UrlFetchApp.fetch(TM+path,{method:'POST',headers:tmAuth(e),
    contentType:'application/x-www-form-urlencoded',payload:payload,
    muteHttpExceptions:true,followRedirects:true});
}

// ── iCalendar helpers ────────────────────────────────────────────────────
function icsDateGAS(dateStr, timeStr){
  var p=dateStr.split('-'), t=timeStr.split(':');
  return p[0]+p[1]+p[2]+'T'+t[0]+t[1]+'00';
}
function icsEscapeGAS(str){
  return String(str||'').split('\\').join('\\\\').split(';').join('\\;')
    .split(',').join('\\,').split('\n').join('\\n');
}

// ── Name matching ────────────────────────────────────────────────────────
function nameMatches(contact, clientName) {
  const searchNorm = clientName.trim().toLowerCase().replace(/\s+/g, ' ');
  const contactFull = ((contact.firstName||'')+' '+(contact.middleName||'')+' '+(contact.lastName||'')).trim().toLowerCase().replace(/\s+/g, ' ');
  const contactFirst = (contact.firstName||'').trim().toLowerCase();
  const contactLast = (contact.lastName||'').trim().toLowerCase();
  const searchParts = searchNorm.split(' ');
  const searchFirst = searchParts[0]||'';
  const searchLast = searchParts[searchParts.length-1]||'';
  if (contactFull === searchNorm) return true;
  if (searchParts.length>1 && contactFirst===searchFirst && contactLast===searchLast) return true;
  if (contactFull.includes(searchNorm) || searchNorm.includes(contactFull)) return true;
  return false;
}

// ── Daily auto-reminder trigger ──────────────────────────────────────────
function dailyReminders() {
  const sheet = getSheet();
  const raw = sheet.getRange('A1').getValue();
  var record = {};
  try { record = JSON.parse(raw||'{}'); } catch(e) { return; }

  const props = PropertiesService.getScriptProperties();
  const bookings = record.bookings || [];
  const zoomBookings = record.zoomBookings || [];

  const today = new Date();
  const target = new Date(today);
  target.setDate(today.getDate()+2);
  const targetStr = Utilities.formatDate(target, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  Logger.log('Daily reminders running for target date: '+targetStr);

  const due = bookings.filter(function(b){
    return b.date===targetStr && !b.smsSent && !b.autoSent && b.note && !b.isOnlineSession;
  });
  Logger.log('Room bookings due: '+due.length);

  due.forEach(function(b){
    const tmUser   = props.getProperty('tm_user_'+b.who)||'';
    const tmKey    = props.getProperty('tm_key_'+b.who)||'';
    const tmMobile = props.getProperty('tm_mobile_'+b.who)||'07498218609';
    if (!tmUser||!tmKey){ Logger.log('Skipping '+b.who+' - no TM credentials'); return; }

    var clients = (b.couples&&b.client1&&b.client2) ? [b.client1,b.client2] : [b.note||''];
    const d = Utilities.formatDate(new Date(b.date+'T12:00:00'), Session.getScriptTimeZone(), 'EEE d MMM yyyy');

    clients.forEach(function(clientName){
      if (!clientName) return;
      var clientPhone = '';
      var clientFirstName = clientName.split(' ')[0];
      try {
        const searchR = UrlFetchApp.fetch(TM+'/contacts/search?limit=5&page=1&query='+encodeURIComponent(clientName),{
          method:'GET', headers:{'X-TM-Username':tmUser,'X-TM-Key':tmKey}, muteHttpExceptions:true
        });
        if (searchR.getResponseCode()===200){
          const contacts = JSON.parse(searchR.getContentText()).resources||[];
          const matchedContact = contacts.find(function(c){ return nameMatches(c, clientName); });
          if (matchedContact){
            clientPhone = matchedContact.phone||'';
            Logger.log('Matched: '+clientPhone+' for: '+clientName);
          } else {
            Logger.log('No match for: '+clientName+' ('+contacts.length+' results)');
            return;
          }
        }
      } catch(err) { Logger.log('Contact search error: '+err.message); return; }
      if (!clientPhone) return;
      const text = 'Note: Hi '+clientFirstName+', you have an appt with '+b.who+
        ' in Room 34, Canute Suite, Royal Mail House, Terminus Ter, Southampton SO14 3FD on '+
        d+' at '+b.start+'. Call '+tmMobile+'. Reply STOP to opt out.';
      try {
        const sendR = UrlFetchApp.fetch(TM+'/messages',{
          method:'POST', headers:{'X-TM-Username':tmUser,'X-TM-Key':tmKey},
          contentType:'application/x-www-form-urlencoded',
          payload:{text:text, phones:clientPhone}, muteHttpExceptions:true
        });
        if (sendR.getResponseCode()===201){
          b.autoSent = new Date().toISOString();
          if (!record.smsLog) record.smsLog = [];
          record.smsLog.unshift({id:Date.now(),who:b.who,date:b.date,start:b.start,
            clientName:clientName,phone:clientPhone,sentAt:new Date().toISOString(),auto:true,type:'remind'});
          Logger.log('Sent reminder to '+clientName);
        } else {
          Logger.log('SMS failed for '+clientName+': HTTP '+sendR.getResponseCode()+' '+sendR.getContentText().slice(0,100));
        }
      } catch(err) { Logger.log('SMS error: '+err.message); }
    });
    b.autoSent = new Date().toISOString();
  });

  const zoomDue = zoomBookings.filter(function(b){
    return b.date===targetStr && !b.zoomReminderSent && b.zoomJoinUrl;
  });
  Logger.log('Zoom bookings due: '+zoomDue.length);

  zoomDue.forEach(function(b){
    const tmUser   = props.getProperty('tm_user_'+b.who)||'';
    const tmKey    = props.getProperty('tm_key_'+b.who)||'';
    const tmMobile = props.getProperty('tm_mobile_'+b.who)||'07498218609';
    if (!tmUser||!tmKey){ Logger.log('Skipping Zoom for '+b.who); return; }

    var clients = (b.couples&&b.client1&&b.client2) ? [b.client1,b.client2] : [b.note||''];
    const d = Utilities.formatDate(new Date(b.date+'T12:00:00'), Session.getScriptTimeZone(), 'EEE d MMM yyyy');
    const joinUrl = b.zoomShortUrl||b.zoomJoinUrl;

    clients.forEach(function(clientName){
      if (!clientName) return;
      try {
        const searchR = UrlFetchApp.fetch(TM+'/contacts/search?limit=5&page=1&query='+encodeURIComponent(clientName),{
          method:'GET', headers:{'X-TM-Username':tmUser,'X-TM-Key':tmKey}, muteHttpExceptions:true
        });
        if (searchR.getResponseCode()!==200) return;
        const contacts = JSON.parse(searchR.getContentText()).resources||[];
        const contact = contacts.find(function(c){ return nameMatches(c, clientName); });
        if (!contact){ Logger.log('No match for Zoom client: '+clientName); return; }
        const first = clientName.split(' ')[0];
        const phone = contact.phone||'';
        const email = contact.email||'';
        if (phone){
          const smsText = 'Note: Hi '+first+', your Zoom session with '+b.who+' is on '+d+' at '+b.start+
            '. Join: '+joinUrl+'. Call '+tmMobile+' to cancel. Reply STOP to opt out.';
          const smsR = UrlFetchApp.fetch(TM+'/messages',{
            method:'POST', headers:{'X-TM-Username':tmUser,'X-TM-Key':tmKey},
            contentType:'application/x-www-form-urlencoded',
            payload:{text:smsText,phones:phone}, muteHttpExceptions:true
          });
          Logger.log('Zoom SMS HTTP '+smsR.getResponseCode()+' to '+clientName);
        }
        if (email){
          var NL2 = '\n';
          const subject = 'Reminder: Zoom session with '+b.who+' on '+d;
          const body2 = 'Hi '+first+','+NL2+NL2+
            'This is a reminder that your Zoom session with '+b.who+' is on '+d+' at '+b.start+'.'+NL2+NL2+
            'Join here: '+joinUrl+NL2+NL2+
            'To cancel call '+tmMobile+'.'+NL2+NL2+
            'Kind regards,'+NL2+b.who;
          try { GmailApp.sendEmail(email, subject, body2); Logger.log('Zoom email sent to '+email); }
          catch(ge){ Logger.log('Email error: '+ge.message); }
        }
      } catch(err){ Logger.log('Zoom reminder error: '+err.message); }
    });
    b.zoomReminderSent = new Date().toISOString();
  });

  // Always trim BEFORE writing — regardless of whether anything was sent
  record = trimRecord(record);
  sheet.getRange('A1').setValue(JSON.stringify(record));
  Logger.log('Sheet updated and trimmed');
}

// ── Zoom helpers ─────────────────────────────────────────────────────────
function setZoomCredentials(name, accountId, clientId, clientSecret) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('zoom_account_'+name, accountId);
  props.setProperty('zoom_client_'+name,  clientId);
  props.setProperty('zoom_secret_'+name,  clientSecret);
  Logger.log('Zoom credentials saved for '+name);
}

function setZoomEmail(name, email) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('zoom_email_'+name, email);
  Logger.log('Zoom email saved for '+name+': '+email);
}

function setBitlyToken(name, token) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('bitly_token_'+name, token);
  Logger.log('Bitly token saved for '+name);
}

function formatZoomDateTime(date, startTime) {
  var parts = date.split('-');
  var timeParts = startTime.split(':');
  return parts[0]+'-'+('0'+parts[1]).slice(-2)+'-'+('0'+parts[2]).slice(-2)+'T'+
    ('0'+parseInt(timeParts[0])).slice(-2)+':'+('0'+parseInt(timeParts[1])).slice(-2)+':00';
}

function getZoomToken(name) {
  const props = PropertiesService.getScriptProperties();
  const accountId    = props.getProperty('zoom_account_'+name)||'';
  const clientId     = props.getProperty('zoom_client_'+name)||'';
  const clientSecret = props.getProperty('zoom_secret_'+name)||'';
  if (!accountId||!clientId||!clientSecret) throw new Error('No Zoom credentials for '+name);
  const encoded = Utilities.base64Encode(clientId+':'+clientSecret);
  const r = UrlFetchApp.fetch(
    'https://zoom.us/oauth/token?grant_type=account_credentials&account_id='+accountId,
    { method:'POST', headers:{'Authorization':'Basic '+encoded}, muteHttpExceptions:true }
  );
  if (r.getResponseCode()!==200) throw new Error('Zoom token failed: '+r.getContentText());
  return JSON.parse(r.getContentText()).access_token;
}

function shortenUrl(name, longUrl) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('bitly_token_'+name)||'';
  if (!token) return longUrl;
  try {
    const r = UrlFetchApp.fetch('https://api-ssl.bitly.com/v4/shorten', {
      method:'POST',
      headers:{'Authorization':'Bearer '+token,'Content-Type':'application/json'},
      payload: JSON.stringify({long_url: longUrl}),
      muteHttpExceptions:true
    });
    if (r.getResponseCode()===200||r.getResponseCode()===201){
      return JSON.parse(r.getContentText()).link||longUrl;
    }
  } catch(e) { Logger.log('Bitly error: '+e.message); }
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
  const userEmail = props2.getProperty('zoom_email_'+name)||'me';
  const r = UrlFetchApp.fetch('https://api.zoom.us/v2/users/'+userEmail+'/meetings', {
    method:'POST',
    headers:{'Authorization':'Bearer '+token,'Content-Type':'application/json'},
    payload: JSON.stringify(payload),
    muteHttpExceptions:true
  });
  const code = r.getResponseCode();
  if (code!==201) throw new Error('Zoom creation failed ('+code+'): '+r.getContentText());
  const meeting = JSON.parse(r.getContentText());
  const shortUrl = shortenUrl(name, meeting.join_url);
  return {
    meetingId:  meeting.id,
    joinUrl:    meeting.join_url,
    shortUrl:   shortUrl,
    password:   meeting.password||'',
    startUrl:   meeting.start_url
  };
}

function deleteZoomMeeting(name, meetingId) {
  try {
    const token = getZoomToken(name);
    UrlFetchApp.fetch('https://api.zoom.us/v2/meetings/'+meetingId, {
      method:'DELETE', headers:{'Authorization':'Bearer '+token}, muteHttpExceptions:true
    });
    Logger.log('Zoom meeting '+meetingId+' deleted');
  } catch(e) { Logger.log('Zoom delete error: '+e.message); }
}

// ── Credentials ──────────────────────────────────────────────────────────
function setCredentials(name, tmUser, tmKey, mobileNumber) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('tm_user_'+name,   tmUser);
  props.setProperty('tm_key_'+name,    tmKey);
  props.setProperty('tm_mobile_'+name, mobileNumber);
  Logger.log('Credentials saved for '+name);
}

// ── Keep alive ───────────────────────────────────────────────────────────
function keepAlive() {
  getSheet();
}

// ── Manual trim ──────────────────────────────────────────────────────────
function trimSheetData() {
  const sheet = getSheet();
  const raw = sheet.getRange('A1').getValue();
  var record = JSON.parse(raw||'{}');
  const before = (record.zoomBookings||[]).length;
  record = trimRecord(record);
  Logger.log('Zoom bookings: '+before+' -> '+(record.zoomBookings||[]).length);
  const json = JSON.stringify(record);
  Logger.log('New size: '+json.length+' chars');
  sheet.getRange('A1').setValue(json);
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
    Logger.log('Zoom token: '+token.substring(0,20)+'...');
    const meeting = createZoomMeeting('Rosswell','Test Meeting',
      Utilities.formatDate(new Date(),'Europe/London','yyyy-MM-dd'), '09:00', 30);
    Logger.log('Test meeting: '+meeting.joinUrl);
  } catch(e) { Logger.log('Error: '+e.message); }
}

function debugNameMatch() {
  const props = PropertiesService.getScriptProperties();
  const tmUser = props.getProperty('tm_user_Rosswell')||'';
  const tmKey  = props.getProperty('tm_key_Rosswell')||'';
  const names = ['Vanessa Loubier', 'Heather Walker', 'Maggie Lomax'];
  names.forEach(function(clientName) {
    const r = UrlFetchApp.fetch(
      TM+'/contacts/search?limit=5&page=1&query='+encodeURIComponent(clientName),
      {method:'GET', headers:{'X-TM-Username':tmUser,'X-TM-Key':tmKey}, muteHttpExceptions:true}
    );
    Logger.log('HTTP '+r.getResponseCode()+' for: "'+clientName+'"');
    Logger.log(r.getContentText());
  });
}
