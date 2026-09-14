// ============================================================
// ⚙️ تنظیمات این نسخه — برای هر شرکت/مشتری جدید فقط همین یک خط رو عوض کن
// (خط دومی که باید عوض بشه، API_URL توی خود index.html هست)
// ============================================================
var MAX_EMPLOYEES = 0; // 0 = بدون محدودیت. برای محدود کردن، یه عدد بذار مثلاً 10 یا 20.
var ADVANCED_TOOLS_PASSWORD = '2067230'; // رمزِ جداگانه برای «تعمیرِ داده‌های قدیمی» تو تنظیمات — هر وقت خواستی عوض کن، فقط همین عدد رو تغییر بده.
// «تنظیماتِ ابرِ مدیریت» — با زدنِ آیکونِ چرخ‌دنده‌یِ پایینِ صفحه‌یِ ورودِ مدیریت باز می‌شه. کاملاً
// جدا از یوزرنیم/پسوردِ ادمینِ معمولی؛ اینجا فقط برایِ محدودیتِ اشتراک و تعدادِ پرسنل استفاده
// می‌شه، پس بهتره فقط خودت (صاحبِ برنامه) اینا رو بدونی، نه هر ادمینِ معمولی.
var SUPER_ADMIN_USERNAME = 'trddyar';
var SUPER_ADMIN_PASSWORD = 'Moh@mm@d2067230';

// ============================================================
// Self-healing sheet getter: creates the sheet if missing, and
// adds any newly-introduced header columns to existing sheets
// (so upgrading this script never breaks old data).
// ============================================================
function getSheet(name, headers){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  var justCreated = false;
  var justAddedCols = false;
  if(!sh){
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    justCreated = true;
  } else {
    var lastCol = Math.max(sh.getLastColumn(), 1);
    var existing = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    var missing = headers.filter(function(h){ return existing.indexOf(h) === -1; });
    if(missing.length > 0){
      sh.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
      justAddedCols = true;
    }
  }
  // Force plain-text formatting on the whole working area. Without this, Sheets
  // silently reinterprets strings like "14:00:00" as its own internal time type,
  // which then comes back from the API as a garbled 1899-12-30T...Z value.
  // This is a fairly expensive write (up to 1000x20 cells), and the format sticks once
  // applied — so it only needs to run the FIRST time a sheet is created or gains new
  // columns, not on every single read. Running it unconditionally on every getSheet()
  // call (as before) meant every action — even plain reads — paid this cost once per
  // sheet it touched, which was the single biggest drag on the whole site's speed.
  if(justCreated || justAddedCols){
    var wideCols = Math.max(headers.length, sh.getLastColumn(), 20);
    sh.getRange(1, 1, 1000, wideCols).setNumberFormat('@');
  }
  return sh;
}

// A malformed/unparseable isoTimestamp must never silently corrupt sort order and hide a
// genuinely recent, valid entry (NaN comparisons in Array.sort are undefined behavior) —
// treat anything unparseable as "oldest possible" so it always sorts to the very start,
// never masking real data near the end of the list.
function isoCompare(isoA, isoB){
  var ta = new Date(isoA).getTime();
  var tb = new Date(isoB).getTime();
  if(isNaN(ta)) ta = -Infinity;
  if(isNaN(tb)) tb = -Infinity;
  return ta - tb;
}
// Pairs a chronologically-sorted list of one employee's log rows into in/out cycles — same
// state-machine approach as the frontend's pairLogEntriesForDisplay: each check-in pairs with
// the next check-out; a check-in with no closing check-out before the next check-in is left
// as its own open/orphan pair. Used for things like absence detection where we need to know
// which calendar day each cycle actually started on.
function pairLogsIntoCycles(sortedEntries){
  var pairs = [];
  var openIn = null;
  sortedEntries.forEach(function(e){
    if(e.type === 'ورود'){
      if(openIn) pairs.push({inE:openIn, outE:null});
      openIn = e;
    } else {
      if(openIn){ pairs.push({inE:openIn, outE:e}); openIn = null; }
      else pairs.push({inE:null, outE:e});
    }
  });
  if(openIn) pairs.push({inE:openIn, outE:null});
  return pairs;
}
// Returns true if inserting a new 'خروج' (checkout) at the given timestamp would leave it
// without a matching open check-in before it — an orphan checkout. Used to block exactly
// that everywhere admin tools can create/edit a log entry. Pairing is purely by chronological
// order (via pairLogsIntoCycles), never same-calendar-day matching, so a legitimate overnight
// shift — check-in one day, checkout the next — is still correctly matched and never flagged.
function wouldBeOrphanCheckout(existingLogsForUser, newIsoTimestamp){
  var combined = existingLogsForUser.slice();
  combined.push({type:'خروج', isoTimestamp: newIsoTimestamp, _isProbe: true});
  combined.sort(function(a,b){ return isoCompare(a.isoTimestamp, b.isoTimestamp); });
  var pairs = pairLogsIntoCycles(combined);
  var match = pairs.filter(function(p){ return p.outE && p.outE._isProbe; })[0];
  return !!(match && !match.inE);
}
function readRows(sheet){
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  // Fields that MUST always be compared as strings, no matter how Sheets happened to type
  // that particular cell. A purely-numeric username/id (e.g. "880196963") is exactly the kind
  // of value Sheets loves to silently auto-store as a NUMBER instead of text — and a number
  // never strictly-equals the string version of the same value in JS. Every "===" comparison
  // against username/id/employeeCode/project id anywhere in this file quietly relied on both
  // sides already being the same type, which broke for numeric-looking values: the affected
  // employee's own history became invisible to every lookup, so nothing ever recognized them
  // as already checked in.
  var FORCE_STRING_FIELDS = {username:1, id:1, employeeCode:1, project_id:1, deviceId:1};
  var rows = [];
  for(var i=1;i<values.length;i++){
    var obj = {};
    for(var j=0;j<headers.length;j++){
      var v = values[i][j];
      // A handful of older cells may have gotten auto-converted by Sheets into its own
      // internal date/time type before this sheet was reliably kept as plain text —
      // normalize those back into a sane string instead of leaking a raw
      // "1899-12-30T17:34:16.000Z" value out to the app.
      if(v instanceof Date) v = normalizeLegacyDateCell(v);
      if(FORCE_STRING_FIELDS[headers[j]] && v !== '' && v !== null && v !== undefined) v = String(v);
      obj[headers[j]] = v;
    }
    obj._row = i+1;
    rows.push(obj);
  }
  return rows;
}
function normalizeLegacyDateCell(d){
  function pad(n){ return (n<10?'0':'')+n; }
  var isTimeOnly = d.getFullYear() === 1899 && d.getMonth() === 11 && d.getDate() === 30;
  if(isTimeOnly) return pad(d.getHours())+':'+pad(d.getMinutes())+':'+pad(d.getSeconds());
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
}

var _colIndexCache = {};
function colIndex(sheet, headerName){
  var sheetName = sheet.getName();
  if(!_colIndexCache[sheetName]){
    var lastCol = Math.max(sheet.getLastColumn(), 1);
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var map = {};
    headers.forEach(function(h, i){ map[h] = i+1; });
    _colIndexCache[sheetName] = map;
  }
  return _colIndexCache[sheetName][headerName] || 0;
}

// Appends a row using an object keyed by header name, so it's always correct
// regardless of the PHYSICAL column order in the sheet (which can drift after
// schema upgrades, since getSheet() appends new columns at the end).
function appendRowByHeader(sheet, dataObj){
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var row = headers.map(function(h){ return dataObj.hasOwnProperty(h) ? dataObj[h] : ''; });
  var targetRow = sheet.getLastRow() + 1;
  // Plain-text format on just this new row BEFORE writing — prevents Sheets from auto-
  // converting strings like "17:34:16" into its own internal date/time type (which then
  // comes back from the API as a garbled 1899-12-30T...Z value). Formatting only the one
  // row being written is cheap, unlike reformatting the whole sheet on every read, and
  // correctly covers rows beyond whatever range got pre-formatted when the sheet was made.
  sheet.getRange(targetRow, 1, 1, lastCol).setNumberFormat('@');
  sheet.appendRow(row);
  // Force this write to commit immediately. Apps Script can otherwise reuse a "warm"
  // execution context for the very next incoming request (e.g. the app's own follow-up
  // fetch a moment after a check-in) whose read of the sheet may not yet reflect a write
  // that technically hasn't been flushed — this is what let someone check in more than
  // once in a row: the immediate re-fetch still saw the old "last entry" and the button
  // flipped back to available. Flushing here makes every write visible to the very next
  // read, no matter how soon it happens.
  SpreadsheetApp.flush();
}

// Backfill missing ids on old rows so every row can be addressed (edit/delete).
function ensureIds(sheet, rows){
  var idCol = colIndex(sheet, 'id');
  rows.forEach(function(r){
    if(!r.id){
      r.id = Utilities.getUuid();
      sheet.getRange(r._row, idCol).setValue(r.id);
    }
  });
  return rows;
}

function getAdminPin(){
  var sh = getSheet('Settings', ['key','value']);
  var rows = readRows(sh);
  var pinRow = rows.filter(function(r){ return r.key === 'admin_pin'; })[0];
  if(!pinRow){ appendRowByHeader(sh, {key:'admin_pin', value:'1234'}); return '1234'; }
  return String(pinRow.value);
}
function getAdminUsername(){
  var u = getSetting('admin_username');
  return u || 'admin';
}
function getSetting(key){
  var sh = getSheet('Settings', ['key','value']);
  var row = readRows(sh).filter(function(r){ return r.key === key; })[0];
  return row ? String(row.value) : '';
}
function setSetting(key, value){
  var sh = getSheet('Settings', ['key','value']);
  var row = readRows(sh).filter(function(r){ return r.key === key; })[0];
  if(row){ esh_set(sh, row._row, 'value', value); }
  else{ appendRowByHeader(sh, {key:key, value:value}); }
}
function notifyAdmin(subject, body){
  var email = getSetting('notify_email');
  if(!email) return;
  try{
    MailApp.sendEmail(email, subject, body);
    setSetting('last_email_status', 'ok — ' + new Date().toISOString());
  }catch(e){
    setSetting('last_email_status', 'error — ' + (e && e.message ? e.message : e));
  }
}
// Run this manually from the Apps Script editor (select it in the function dropdown, click Run)
// the first time you set up email notifications. This forces Google to show the
// "Send email as you" permission prompt, which a web app request cannot trigger itself.
function testEmailSetup(){
  var to = getSetting('notify_email') || Session.getActiveUser().getEmail();
  MailApp.sendEmail(to, 'تست اعلان ثبت تردد', 'اگر این ایمیل رو دریافت کردید، اعلان‌های ایمیلی درست تنظیم شده‌اند.');
}

function distMeters(lat1, lng1, lat2, lng2){
  var R = 6371000;
  var dLat = (lat2-lat1) * Math.PI/180;
  var dLng = (lng2-lng1) * Math.PI/180;
  var a = Math.sin(dLat/2)*Math.sin(dLat/2) +
          Math.cos(lat1*Math.PI/180)*Math.cos(lat2*Math.PI/180)*
          Math.sin(dLng/2)*Math.sin(dLng/2);
  var c = 2*Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
  return R*c;
}

function jsonOut(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// Server-authoritative date/time (Asia/Tehran) + Jalali conversion.
// The client's device clock is NEVER trusted for the recorded time.
// ============================================================
var TZ = 'Asia/Tehran';

function toJalaali(gy, gm, gd) {
  var g_d_m = [0,31,59,90,120,151,181,212,243,273,304,334];
  var jy = (gy <= 1600) ? 0 : 979;
  gy -= (gy <= 1600) ? 621 : 1600;
  var gy2 = (gm > 2) ? (gy + 1) : gy;
  var days = (365*gy) + (Math.floor((gy2+3)/4)) - (Math.floor((gy2+99)/100)) + (Math.floor((gy2+399)/400)) - 80 + gd + g_d_m[gm-1];
  jy += 33*Math.floor(days/12053);
  days %= 12053;
  jy += 4*Math.floor(days/1461);
  days %= 1461;
  if (days > 365) { jy += Math.floor((days-1)/365); days = (days-1)%365; }
  var jm = (days < 186) ? 1+Math.floor(days/31) : 7+Math.floor((days-186)/30);
  var jd = 1 + ((days < 186) ? (days%31) : ((days-186)%30));
  return [jy, jm, jd];
}
var J_MONTHS = ['فروردین','اردیبهشت','خرداد','تیر','مرداد','شهریور','مهر','آبان','آذر','دی','بهمن','اسفند'];
var FA_DIGITS = ['۰','۱','۲','۳','۴','۵','۶','۷','۸','۹'];
function toFaDigits(s){
  return String(s).replace(/[0-9]/g, function(d){ return FA_DIGITS[+d]; });
}
// Employee-entered <input type="time"> values come back as plain "HH:MM" (Latin digits, no seconds).
// Normalize them to the same "HH:mm:ss" Persian-digit format the server itself produces,
// so times are never inconsistent depending on where they originated.
function normalizeTimeInput(t){
  if(!t) return '';
  var parts = String(t).split(':');
  var hh = (parts[0]||'00'); if(hh.length<2) hh='0'+hh;
  var mm = (parts[1]||'00'); if(mm.length<2) mm='0'+mm;
  return hh + ':' + mm + ':00';
}
function pad2(n){ n = String(n); return n.length < 2 ? '0'+n : n; }
function timeToMinutes(hhmm){
  var parts = String(hhmm||'0:0').split(':');
  return (parseInt(parts[0],10)||0)*60 + (parseInt(parts[1],10)||0);
}
function jalaliDateSortKey(dateStr){
  var p = String(dateStr||'').trim().split(' ');
  if(p.length < 3) return 0;
  var day = parseInt(p[0],10)||0;
  var monthIdx = J_MONTHS.indexOf(p[1]) + 1;
  var year = parseInt(p[p.length-1],10)||0;
  return year*10000 + monthIdx*100 + day;
}
// Looks up one row from the Shifts sheet by id — shared by getEffectiveShift and login so
// both read the exact same full shift-definition fields (including the newer advanced-settings
// and weekly-schedule columns), rather than each hand-rolling a slightly different subset.
// Fetches every Shifts row once and returns it as an id->row map — used by all three
// computeEmployeePerformance call sites to avoid each one hand-rolling the same sheet read,
// and critically avoids re-reading the Shifts sheet once per employee in the bulk report.
function buildShiftsByIdMap(){
  var shdefsAll = getSheet('Shifts', ['id','name','type','shiftStart','shiftEnd','requiredMinutes','lateTolerance','earlyTolerance','shiftYear','floatingMinutes','dailyOvertimeCapMinutes','weeklyScheduleJson','createdAt','otStartMinutes','otMiddleMinutes','otEndMinutes','lateAllowedMinutes','earlyLeaveAllowedMinutes','floatBeforeShiftMinutes','attendanceWindowStart','attendanceWindowEnd','followOfficialHolidays']);
  var map = {};
  readRows(shdefsAll).forEach(function(r){ map[r.id] = r; });
  return map;
}
// Fetches every ShiftDays row once and returns it as a "shiftId|jalaliDate" -> row map — used
// by all three computeEmployeePerformance call sites the same way buildShiftsByIdMap is, so
// the sheet is never re-read per employee/per day inside the bulk report's loop.
// شیفتِ عادی's configurable defaults — read once per computeEmployeePerformance call site
// (never inside its per-day loop) and passed in as normalSettings. Falls back to the exact
// original hardcoded values (7:20 required, 60min break, 22:00–06:00 night) for any field the
// admin hasn't explicitly saved yet, so an untouched installation's calculations never change
// just because this feature exists.
// License/subscription window — set via the hidden ⚙️ gear on the admin login screen (super-
// admin settings, gated by its own separate username/password, NOT the regular admin PIN).
// Empty start/end (the default, untouched state) means no restriction at all — this only ever
// blocks logins once BOTH dates have actually been configured.
// Configurable via the super-admin settings panel (⚙️ gear on the admin login screen) — falls
// back to the MAX_EMPLOYEES code constant above until the admin actually saves a value there,
// so nothing changes for an installation that never touches this setting.
function getMaxEmployees(){
  var v = getSetting('max_employees');
  return v !== '' ? (parseInt(v,10)||0) : MAX_EMPLOYEES;
}
function isLicenseValid(){
  var startStr = getSetting('license_start_date');
  var endStr = getSetting('license_end_date');
  if(!startStr || !endStr) return true;
  var todayKey = jalaliDateSortKey(serverNow().jalaliDate);
  return todayKey >= jalaliDateSortKey(startStr) && todayKey <= jalaliDateSortKey(endStr);
}
function getNormalShiftSettings(){
  return {
    requiredMinutes: getSetting('normal_required_minutes') || String(7*60+20),
    breakMinutes: getSetting('normal_break_minutes') || '60',
    nightStart: getSetting('normal_night_start') || '22:00',
    nightEnd: getSetting('normal_night_end') || '06:00'
  };
}
function buildShiftDaysMap(){
  var sdAll = getSheet('ShiftDays', ['id','shiftId','jalaliDate','dayType','start1','end1','hasSecondPart','start2','end2',
    'dailyRequiredMinutes','dayFloatingMinutes','dayOvertimeCapMinutes','dayOtStartMinutes','dayOtMiddleMinutes','dayOtEndMinutes',
    'dayLateAllowedMinutes','dayEarlyAllowedMinutes','dayFloatBeforeShiftMinutes','dayAttendanceWindowStart','dayAttendanceWindowEnd',
    'isNightShift','nightShiftCrossesNextDay','nightShiftStart','nightShiftDuration','createdAt']);
  var map = {};
  readRows(sdAll).forEach(function(r){ map[r.shiftId + '|' + r.jalaliDate] = r; });
  return map;
}
function lookupShiftDefById(shiftId){
  if(!shiftId) return null;
  var shdefs = getSheet('Shifts', ['id','name','type','shiftStart','shiftEnd','requiredMinutes','lateTolerance','earlyTolerance','shiftYear','floatingMinutes','dailyOvertimeCapMinutes','weeklyScheduleJson','createdAt','otStartMinutes','otMiddleMinutes','otEndMinutes','lateAllowedMinutes','earlyLeaveAllowedMinutes','floatBeforeShiftMinutes','attendanceWindowStart','attendanceWindowEnd','followOfficialHolidays']);
  return readRows(shdefs).filter(function(r){ return r.id === shiftId; })[0] || null;
}
// Resolves which shift's rules actually apply to a given employee on a given day. Always
// returns an object (never null) so existing callers that read .type/.shiftStart/etc keep
// working unchanged — "شیفتِ عادی" (no special shift at all) is represented the same way it
// always was, as a plain {type:'simple', shiftStart: emp.shiftStart, ...} built from the
// employee's own flat fields.
//
// Resolution order for a given day:
//   1. A date-RANGED assignment (has an endDate) covering that day — this is the "فقط برایِ
//      این بازه" case and always wins, since that's the entire point of a temporary override.
//   2. Otherwise, the most recent GENERAL assignment (no endDate) whose startDate is on or
//      before this day — the "کلی، از این تاریخ به بعد" case. If the admin has changed the
//      general shift more than once over time, whichever change's startDate is latest (while
//      still not being in the future relative to this day) is the one that applies — so past
//      days keep calculating under whatever was genuinely in effect on them.
//   3. Otherwise, the employee's plain shiftId field (kept for the simplest "این پرسنل این
//      شیفت رو داره، همیشه" case with no date logic involved at all).
//   4. Otherwise, شیفتِ عادی — the plain historical default this whole function used to
//      always return.
function getEffectiveShift(username, jalaliDateStr, emp){
  var sash = getSheet('ShiftAssignments', ['id','username','startDate','endDate','shiftStart','shiftEnd','shiftName','shiftId','createdAt']);
  var key = jalaliDateSortKey(jalaliDateStr);
  var rowsForUser = readRows(sash).filter(function(r){ return r.username === username; });
  var rangedMatch = rowsForUser.filter(function(r){
    return r.endDate && jalaliDateSortKey(r.startDate) <= key && key <= jalaliDateSortKey(r.endDate);
  }).sort(function(a,b){ return new Date(b.createdAt||0).getTime() - new Date(a.createdAt||0).getTime(); })[0];

  if(rangedMatch){
    if(rangedMatch.shiftId){
      var shdefR = lookupShiftDefById(rangedMatch.shiftId);
      if(shdefR) return shdefR;
      // shiftId pointed at a since-deleted shift — fall through below rather than silently
      // applying no rules at all for the rest of this function.
    } else if(rangedMatch.shiftStart || rangedMatch.shiftEnd){
      // Old-style flat assignment (predates the full Shift-definition system) — still honored.
      return {type:'simple', shiftStart: rangedMatch.shiftStart, shiftEnd: rangedMatch.shiftEnd, name: rangedMatch.shiftName||'', requiredMinutes:'', lateTolerance:0, earlyTolerance:0};
    }
    // else: an explicit "بازگشت به شیفتِ عادی برای این بازه" row (shiftId and flat fields all
    // empty on purpose) — falls through to the general/normal resolution below, same as if no
    // ranged assignment existed for this day at all.
  }

  var generalMatches = rowsForUser.filter(function(r){
    return !r.endDate && jalaliDateSortKey(r.startDate) <= key;
  }).sort(function(a,b){ return jalaliDateSortKey(b.startDate) - jalaliDateSortKey(a.startDate); });
  if(generalMatches.length){
    var gMatch = generalMatches[0];
    if(gMatch.shiftId){
      var shdefGen = lookupShiftDefById(gMatch.shiftId);
      if(shdefGen) return shdefGen;
    } else {
      return {type:'simple', shiftStart: emp.shiftStart, shiftEnd: emp.shiftEnd, name: '', requiredMinutes:'', lateTolerance:0, earlyTolerance:0};
    }
  }

  if(emp.shiftId){
    var shdefG = lookupShiftDefById(emp.shiftId);
    if(shdefG) return shdefG;
  }
  return {type:'simple', shiftStart: emp.shiftStart, shiftEnd: emp.shiftEnd, name: '', requiredMinutes:'', lateTolerance:0, earlyTolerance:0};
}
function jalaliToGregorianServer(jy, jm, jd){
  var gy = (jy <= 979) ? 621 : 1600;
  jy -= (jy <= 979) ? 0 : 979;
  var days = (jy*365) + (Math.floor(jy/33)*8) + Math.floor(((jy%33)+3)/4) + 78 + jd + ((jm<7) ? (jm-1)*31 : ((jm-7)*30)+186);
  gy += 400*Math.floor(days/146097); days %= 146097;
  if(days > 36524){ gy += 100*Math.floor(--days/36524); days %= 36524; if(days>=365) days++; }
  gy += 4*Math.floor(days/1461); days %= 1461;
  if(days > 365){ gy += Math.floor((days-1)/365); days=(days-1)%365; }
  var gd = days+1;
  var sal_a=[0,31,((gy%4===0&&gy%100!==0)||(gy%400===0))?29:28,31,30,31,30,31,31,30,31,30,31];
  var gm; for(gm=0; gm<13; gm++){ var v=sal_a[gm]; if(gd<=v) break; gd-=v; }
  return [gy, gm, gd];
}
// Number of days in a given Jalali month/year, found by walking the Gregorian calendar
// day-by-day from the 1st of that month until the Jalali month changes — correctly handles
// Esfand's leap-year length without a separate leap-year lookup table.
function jalaliMonthLengthServer(jy, jm){
  var g = jalaliToGregorianServer(jy, jm, 1);
  var dt = new Date(g[0], g[1]-1, g[2]);
  var count = 0;
  while(true){
    var j = toJalaali(dt.getFullYear(), dt.getMonth()+1, dt.getDate());
    if(j[0] !== jy || j[1] !== jm) break;
    count++;
    dt.setDate(dt.getDate()+1);
  }
  return count;
}
function parseLeaveHoursToMinutesShared(s){
  if(!s) return 0;
  var parts = String(s).split(':');
  var h = parseInt(parts[0],10)||0;
  var m = parseInt(parts[1],10)||0;
  return h*60+m;
}
// Resolves which shift applies to one specific day, using ONLY pre-fetched data (no sheet
// reads) — this is what computeEmployeePerformance actually calls in its per-day loop, since
// getEffectiveShift's own sheet reads would otherwise happen 30+ times per employee in the
// bulk monthly report. Returns null for "شیفتِ عادی" (no override — use the plain default
// rules exactly as before), or the full shift-definition row otherwise. Mirrors
// getEffectiveShift's resolution order (ranged, then most-recent general, then emp.shiftId).
function resolveEffectiveShiftCached(jalaliDateStr, emp, assignmentsForUser, shiftsById){
  var key = jalaliDateSortKey(jalaliDateStr);
  var rangedMatch = assignmentsForUser.filter(function(r){
    return r.endDate && jalaliDateSortKey(r.startDate) <= key && key <= jalaliDateSortKey(r.endDate);
  }).sort(function(a,b){ return new Date(b.createdAt||0).getTime() - new Date(a.createdAt||0).getTime(); })[0];
  if(rangedMatch){
    if(rangedMatch.shiftId && shiftsById[rangedMatch.shiftId]) return shiftsById[rangedMatch.shiftId];
    if(!rangedMatch.shiftId && !rangedMatch.shiftStart && !rangedMatch.shiftEnd){
      return null; // explicit "شیفتِ عادی برایِ این بازه"
    }
    // else: old-style flat assignment (no weeklyScheduleJson) — doesn't carry the newer
    // per-day/advanced fields this function uses, so it's outside this integration's scope
    // and treated the same as شیفتِ عادی here (still handled correctly by getEffectiveShift
    // for the note-adding logic elsewhere).
  }
  var generalCandidates = assignmentsForUser.filter(function(r){
    return !r.endDate && jalaliDateSortKey(r.startDate) <= key;
  }).sort(function(a,b){ return jalaliDateSortKey(b.startDate) - jalaliDateSortKey(a.startDate); });
  if(generalCandidates.length){
    var gm = generalCandidates[0];
    if(gm.shiftId && shiftsById[gm.shiftId]) return shiftsById[gm.shiftId];
    if(!gm.shiftId) return null; // explicit "شیفتِ عادی از این تاریخ به بعد"
  }
  if(emp.shiftId && shiftsById[emp.shiftId]) return shiftsById[emp.shiftId];
  return null;
}
// Minutes between two "HH:MM" times, 0 if malformed or non-positive (never lets a bad shift
// definition manufacture negative required-minutes).
function shiftSegmentMinutes(startStr, endStr){
  if(!startStr || !endStr) return 0;
  var m = timeToMinutes(endStr) - timeToMinutes(startStr);
  return m > 0 ? m : 0;
}
// Per-day override helper: a specific calendar date's own advanced-setting field (set via the
// day-edit modal) takes precedence over the shift-level default of the same name — left blank
// on the day, the shift-level value is used instead.
function dayOrShiftValue(dayVal, shiftVal){
  return (dayVal !== undefined && dayVal !== null && dayVal !== '') ? dayVal : shiftVal;
}
// Breaks a day's raw worked pairs into how many minutes fall strictly BEFORE the shift's
// first segment start, strictly BETWEEN segment 1's end and segment 2's start (only
// meaningful for a split shift), and strictly AFTER the last segment's end — used to apply
// the shift's three separate overtime caps (start/middle/end) to only the relevant portion of
// a day's overtime, rather than one lump sum.
function computeShiftOvertimeZones(dayShiftPairs, sched){
  var seg1StartMin = timeToMinutes(sched.start1);
  var seg1EndMin = timeToMinutes(sched.end1);
  var seg2StartMin = sched.hasSecondPart ? timeToMinutes(sched.start2) : null;
  var lastEndMin = sched.hasSecondPart ? timeToMinutes(sched.end2) : seg1EndMin;
  var beforeMin = 0, middleMin = 0, afterMin = 0;
  dayShiftPairs.forEach(function(p){
    var inMin = timeToMinutes(p.inTime);
    var outMin = timeToMinutes(p.outTime);
    if(outMin < inMin) outMin += 24*60; // overnight pair
    if(inMin < seg1StartMin){ beforeMin += Math.min(outMin, seg1StartMin) - inMin; }
    if(outMin > lastEndMin){ afterMin += outMin - Math.max(inMin, lastEndMin); }
    if(seg2StartMin !== null && seg1EndMin < seg2StartMin){
      var overlapStart = Math.max(inMin, seg1EndMin);
      var overlapEnd = Math.min(outMin, seg2StartMin);
      if(overlapEnd > overlapStart) middleMin += (overlapEnd - overlapStart);
    }
  });
  return {beforeMin: Math.max(0,beforeMin), middleMin: Math.max(0,middleMin), afterMin: Math.max(0,afterMin)};
}
// Shared payroll-policy calculation, used by both the admin monthly summary and the
// employee's own monthly/daily views — one pass computes per-day rows AND the monthly totals
// together, so the two views can never drift out of sync with each other.
// Policy: 7:20 required daily hours (none on Fridays/holidays), 1-hour break deducted whenever
// raw worked time exceeds 7:20, night work = minutes worked after 22:00, and Friday/holiday
// work counts entirely as its own category rather than normal/overtime hours.
// Shift integration: when a "شیفتِ ساده" applies to a given day (via emp.shiftId or a
// date-ranged ShiftAssignments row — see resolveEffectiveShiftCached), that day's required
// minutes come from the shift's own weekly schedule instead of the flat 7:20 default, and a
// late arrival/early leave within the shift's allowed minutes is credited back rather than
// counted as a shortfall. Employees with no shift at all (شیفتِ عادی, the default for
// everyone) are completely unaffected — this only ever activates when shiftsById actually
// resolves something for that specific day.
function computeEmployeePerformance(username, yearMP, monthMP, allLogsForUser, markersForUser, holidaySetMP, emp, assignmentsForUser, shiftsById, shiftDaysMap, normalSettings){
  // شیفتِ عادی — the default rules used for every employee who has no specific shift assigned
  // at all. Configurable now (via the "شیفتِ عادی" tab in the شیفت admin menu) instead of being
  // permanently hardcoded — normalSettings carries whatever the admin has saved, falling back
  // to these exact original values if nothing was ever saved (so an untouched installation
  // behaves identically to before this became editable).
  normalSettings = normalSettings || {};
  var REQUIRED_MIN = parseInt(normalSettings.requiredMinutes, 10) || (7*60+20);
  var BREAK_MIN = parseInt(normalSettings.breakMinutes, 10) || 60;
  var NIGHT_START_MIN = normalSettings.nightStart ? timeToMinutes(normalSettings.nightStart) : (22*60);
  // 06:00 the NEXT day, expressed in "minutes since this day's midnight" — night work is only
  // ever nightStart–nightEnd, not open-ended. nightEnd is always interpreted as belonging to
  // the day AFTER nightStart (adding 24h) since a night-work window inherently crosses
  // midnight — a same-day end time earlier than the start wouldn't make sense here.
  var NIGHT_END_MIN = normalSettings.nightEnd ? (24*60 + timeToMinutes(normalSettings.nightEnd)) : (24*60 + 6*60);
  var daysInMonthMP = jalaliMonthLengthServer(yearMP, monthMP);
  // Older/other callers may not pass these — default to "no shift ever applies", which
  // reproduces the exact pre-shift-integration behavior.
  emp = emp || {};
  assignmentsForUser = assignmentsForUser || [];
  shiftsById = shiftsById || {};
  shiftDaysMap = shiftDaysMap || {};

  var pairsMP = pairLogsIntoCycles(allLogsForUser);
  var empMarkersMP = {};
  markersForUser.forEach(function(m){ empMarkersMP[m.jalaliDate] = m; });
  var pairsByDateMP = {};
  pairsMP.forEach(function(p){
    if(p.inE && p.outE){
      var key = p.inE.jalaliDate;
      if(!pairsByDateMP[key]) pairsByDateMP[key] = [];
      pairsByDateMP[key].push(p);
    }
  });

  var totals = {workedMin:0, requiredMin:0, leaveMin:0, nightMin:0, fridayMin:0, holidayMin:0, overtimeMin:0, breakMin:0, holidayBreakMin:0, usefulMin:0, missionDaysCount:0, extraPresenceMin:0};
  var dailyRows = [];

  for(var dMP=1; dMP<=daysInMonthMP; dMP++){
    var dateStrMP = dMP + ' ' + J_MONTHS[monthMP-1] + ' ' + yearMP;
    var isFridayMP = isFridayFromJalaliDateStr(dateStrMP);
    var isHolidayNonFridayMP = !isFridayMP && !!holidaySetMP[dateStrMP];

    // Resolve the shift BEFORE deciding whether today is "off" — whether Friday AND official
    // holidays actually count as off now both depend on the shift's own "تبعیتِ از تعطیلاتِ
    // رسمی" setting (followOfficialHolidays), per the user's explicit correction: when it's
    // unticked, Friday is no longer automatically off either — the shift's own per-day
    // schedule applies to every calendar date with no built-in weekly exception at all.
    var effShiftToday = resolveEffectiveShiftCached(dateStrMP, emp, assignmentsForUser, shiftsById);
    // شیفتِ عادی (no shift assigned) always gets Friday/holidays off, exactly as before. A
    // shift that explicitly sets followOfficialHolidays to 'no' gets NEITHER off — its own
    // per-day schedule (or بدونِ شیفت rules) apply on that date just like any other, and not
    // showing up counts as an absence rather than a day off.
    var shiftFollowsHolidaysMP = !effShiftToday || String(effShiftToday.followOfficialHolidays||'').toLowerCase() !== 'no';
    var isOffDayMP = shiftFollowsHolidaysMP && (isFridayMP || isHolidayNonFridayMP);
    if(isOffDayMP) effShiftToday = null; // a genuinely off day never has a shift-specific schedule, matching the old behavior exactly

    // Resolve today's shift (if any), and if the employee has one, look up THIS SPECIFIC
    // calendar date's own defined hours (ShiftDays — a per-date table, not a repeating weekly
    // template). If the shift has no entry for this exact date, the day is "بدونِ شیفت": not
    // a required-hours day, but any time actually worked still counts as overtime, capped by
    // the shift's own "سقفِ اضافه‌کاریِ روزِ بدونِ شیفت" (dailyOvertimeCapMinutes) if set.
    // Everything here only ever activates when the employee actually has a shift assigned —
    // شیفتِ عادی (effShiftToday === null) is completely untouched, exactly as before.
    var shiftScheduleToday = null;
    var isNoShiftDefinedDay = false;
    var sdEntryMP = null; // the raw ShiftDays row for today, if any — used below for per-day override lookups
    if(effShiftToday){
      var sdKeyMP = effShiftToday.id + '|' + dateStrMP;
      sdEntryMP = shiftDaysMap[sdKeyMP];
      if(sdEntryMP){
        var dayTypeMP = sdEntryMP.dayType || 'simple';
        if(dayTypeMP === 'floating'){
          // شناور day: no fixed check-in/out window at all, just a required total for the
          // day — the employee can work it whenever they want within the day.
          shiftScheduleToday = { isWorkingDay: true, isFloating: true, dailyRequiredMinutes: parseInt(sdEntryMP.dailyRequiredMinutes,10) || 0 };
        } else {
          shiftScheduleToday = {
            isWorkingDay: true, isFloating: false, start1: sdEntryMP.start1, end1: sdEntryMP.end1,
            hasSecondPart: String(sdEntryMP.hasSecondPart||'').toLowerCase() === 'yes',
            start2: sdEntryMP.start2, end2: sdEntryMP.end2
          };
        }
      } else {
        isNoShiftDefinedDay = true;
      }
    }
    // Per-day override helper: a specific calendar date's own advanced-setting fields (set via
    // the day-edit modal) take precedence over the shift-level default of the same name — left
    // blank on the day, the shift-level value is used instead. Falls straight to the
    // shift-level value when there's no day entry at all (بدونِ شیفت still uses the shift's
    // own dailyOvertimeCapMinutes, exactly as before this change).

    var dayWorkedMin = 0, dayNightMin = 0, dayShiftPairs = [];
    // Night-work window for today: the default 22:00–06:00(next day) rule, UNLESS today has a
    // shift day entry that explicitly defines its own custom night-shift window (isNightShift
    // ticked, with its own start + duration) — in which case that custom window entirely
    // replaces the default for this specific day, exactly as described: ticked → uses its own
    // times; not ticked → falls straight back to the plain 22:00–06:00 rule.
    var nightWinStartMP = NIGHT_START_MIN, nightWinEndMP = NIGHT_END_MIN;
    if(sdEntryMP && String(sdEntryMP.isNightShift||'').toLowerCase() === 'yes' && sdEntryMP.nightShiftStart){
      var customNightStart = timeToMinutes(sdEntryMP.nightShiftStart);
      var customNightDuration = timeToMinutes(sdEntryMP.nightShiftDuration||'00:00');
      if(customNightDuration > 0){
        nightWinStartMP = customNightStart;
        nightWinEndMP = customNightStart + customNightDuration;
      }
    }
    (pairsByDateMP[dateStrMP] || []).forEach(function(p){
      var inMin = timeToMinutes(p.inE.time);
      var outMin = timeToMinutes(p.outE.time);
      var durMin = outMin - inMin;
      if(p.outE.jalaliDate !== p.inE.jalaliDate || outMin < inMin){ durMin += 24*60; }
      dayWorkedMin += durMin;
      var outAbs = inMin + durMin;
      // Night minutes are only ever the portion of this pair that overlaps the night window —
      // capped at the window's own end, not left open-ended past it.
      var nightOverlapStart = Math.max(inMin, nightWinStartMP);
      var nightOverlapEnd = Math.min(outAbs, nightWinEndMP);
      if(nightOverlapEnd > nightOverlapStart) dayNightMin += (nightOverlapEnd - nightOverlapStart);
      dayShiftPairs.push({inTime: p.inE.time, outTime: p.outE.time, inProject: p.inE.project||'', outProject: p.outE.project||''});
    });

    // Late-arrival/early-leave allowance — credited back into dayWorkedMin (as if the person
    // had arrived/left exactly on the shift's scheduled time) up to the allowed minutes (this
    // day's own override if set, else the shift-level default), so a small, permitted
    // deviation never shows up as a shortfall. Only applies on a day the shift actually marks
    // as a working, non-floating day (a شناور day has no fixed start/end to be "late" against
    // at all) — using the FIRST pair's check-in against the day's first segment start, and the
    // LAST pair's check-out against the day's last segment end (segment 2's end if this day
    // has one, else segment 1's).
    if(shiftScheduleToday && shiftScheduleToday.isWorkingDay && !shiftScheduleToday.isFloating && dayShiftPairs.length){
      var lateAllowedMP = parseInt(dayOrShiftValue(sdEntryMP && sdEntryMP.dayLateAllowedMinutes, effShiftToday.lateAllowedMinutes), 10) || 0;
      var earlyAllowedMP = parseInt(dayOrShiftValue(sdEntryMP && sdEntryMP.dayEarlyAllowedMinutes, effShiftToday.earlyLeaveAllowedMinutes), 10) || 0;
      if(lateAllowedMP > 0 && shiftScheduleToday.start1){
        var expectedStartMP = timeToMinutes(shiftScheduleToday.start1);
        var actualInMP = timeToMinutes(dayShiftPairs[0].inTime);
        var lateByMP = actualInMP - expectedStartMP;
        if(lateByMP > 0 && lateByMP <= lateAllowedMP) dayWorkedMin += lateByMP;
      }
      var lastSegEndMP = shiftScheduleToday.hasSecondPart ? shiftScheduleToday.end2 : shiftScheduleToday.end1;
      if(earlyAllowedMP > 0 && lastSegEndMP){
        var expectedEndMP = timeToMinutes(lastSegEndMP);
        var actualOutMP = timeToMinutes(dayShiftPairs[dayShiftPairs.length-1].outTime);
        var earlyByMP = expectedEndMP - actualOutMP;
        if(earlyByMP > 0 && earlyByMP <= earlyAllowedMP) dayWorkedMin += earlyByMP;
      }
    }

    // Required minutes: when the employee has a shift AND today has a specifically-defined
    // entry, that entry's hours replace the flat 7:20 default. A "بدونِ شیفت" day (employee
    // has a shift, but this exact date has no entry defined for it) has ZERO required minutes
    // — it's simply not expected to be worked at all, same spirit as Friday/holiday, but any
    // time actually worked on it still counts (capped below by dailyOvertimeCapMinutes rather
    // than the three shift-segment caps, since there's no defined schedule to split against).
    // Friday/holiday days are untouched either way — isOffDayMP already forced effShiftToday
    // to null for them above. شیفتِ عادی (no shift assigned at all) still gets the flat
    // REQUIRED_MIN fallback, exactly as before.
    var dayRequiredMin;
    if(isOffDayMP){
      dayRequiredMin = 0;
    } else if(shiftScheduleToday){
      if(shiftScheduleToday.isFloating){
        dayRequiredMin = shiftScheduleToday.dailyRequiredMinutes;
      } else {
        dayRequiredMin = shiftScheduleToday.isWorkingDay
          ? (shiftSegmentMinutes(shiftScheduleToday.start1, shiftScheduleToday.end1) + (shiftScheduleToday.hasSecondPart ? shiftSegmentMinutes(shiftScheduleToday.start2, shiftScheduleToday.end2) : 0))
          : 0;
      }
    } else if(isNoShiftDefinedDay){
      dayRequiredMin = 0;
    } else {
      dayRequiredMin = REQUIRED_MIN;
    }
    var dayBreakMin = (dayWorkedMin > REQUIRED_MIN) ? BREAK_MIN : 0;
    var dayUsefulMin = dayWorkedMin - dayBreakMin;
    // Per-day value shown on a single day's row — kept clamped at zero so an individual short
    // day never displays as a confusing negative number in the daily breakdown.
    var dayOvertimeMin = Math.max(0, dayUsefulMin - dayRequiredMin);
    // Monthly-total contribution — deliberately NOT clamped. A short day (useful < required)
    // must actually pull the month's overtime total down, or a real shortfall would silently
    // vanish instead of netting against surplus from other days — the total is only correct
    // as (useful − required) added up plainly across every normal day; per-day clamping was
    // exactly what let a shortfall disappear instead of being reflected in the sum.
    var dayOvertimeContribution = isOffDayMP ? 0 : (dayUsefulMin - dayRequiredMin);

    // Three separate overtime caps (start/middle/end of shift) — only relevant on a day a
    // shift actually applies and is a working day, and only when there's genuine overtime
    // (a shortfall day is untouched, exactly as before). The raw before/middle/after minutes
    // are computed from the actual worked pairs, then scaled down proportionally so their sum
    // matches dayOvertimeContribution exactly (which is already net of the day's break) before
    // each zone's own cap is applied — this keeps the break deduction's effect distributed
    // fairly across zones instead of arbitrarily assigning it to just one of them. Whatever a
    // cap trims away is tracked separately as "مازادِ حضور" (extra presence), never simply
    // discarded.
    var dayExtraPresenceMin = 0;
    if(shiftScheduleToday && shiftScheduleToday.isWorkingDay && !shiftScheduleToday.isFloating && dayOvertimeContribution > 0 && dayShiftPairs.length){
      var otStartCap = parseInt(dayOrShiftValue(sdEntryMP && sdEntryMP.dayOtStartMinutes, effShiftToday.otStartMinutes), 10) || 0;
      var otMiddleCap = parseInt(dayOrShiftValue(sdEntryMP && sdEntryMP.dayOtMiddleMinutes, effShiftToday.otMiddleMinutes), 10) || 0;
      var otEndCap = parseInt(dayOrShiftValue(sdEntryMP && sdEntryMP.dayOtEndMinutes, effShiftToday.otEndMinutes), 10) || 0;
      if(otStartCap > 0 || otMiddleCap > 0 || otEndCap > 0){
        var zonesMP = computeShiftOvertimeZones(dayShiftPairs, shiftScheduleToday);
        var rawZoneSumMP = zonesMP.beforeMin + zonesMP.middleMin + zonesMP.afterMin;
        if(rawZoneSumMP > 0){
          var scaleMP = dayOvertimeContribution / rawZoneSumMP;
          var scaledBeforeMP = zonesMP.beforeMin * scaleMP;
          var scaledMiddleMP = zonesMP.middleMin * scaleMP;
          var scaledAfterMP = zonesMP.afterMin * scaleMP;
          var cappedBeforeMP = otStartCap > 0 ? Math.min(scaledBeforeMP, otStartCap) : scaledBeforeMP;
          var cappedMiddleMP = otMiddleCap > 0 ? Math.min(scaledMiddleMP, otMiddleCap) : scaledMiddleMP;
          var cappedAfterMP = otEndCap > 0 ? Math.min(scaledAfterMP, otEndCap) : scaledAfterMP;
          var cappedTotalMP = cappedBeforeMP + cappedMiddleMP + cappedAfterMP;
          dayExtraPresenceMin = Math.max(0, dayOvertimeContribution - cappedTotalMP);
          dayOvertimeContribution = cappedTotalMP;
        }
      }
    } else if((isNoShiftDefinedDay || (shiftScheduleToday && shiftScheduleToday.isFloating)) && dayOvertimeContribution > 0){
      // "بدونِ شیفت" day, OR a شناور (floating) day — neither has a fixed start/end schedule
      // to split overtime zones against, so a single flat cap applies to the whole day's
      // worked time instead: this day's own override if set, else the shift-level
      // dailyOvertimeCapMinutes ("سقفِ اضافه‌کاریِ روزانه/روزِ بدونِ شیفت").
      var noShiftCapMP = parseInt(dayOrShiftValue(sdEntryMP && sdEntryMP.dayOvertimeCapMinutes, effShiftToday.dailyOvertimeCapMinutes), 10) || 0;
      if(noShiftCapMP > 0 && dayOvertimeContribution > noShiftCapMP){
        dayExtraPresenceMin = dayOvertimeContribution - noShiftCapMP;
        dayOvertimeContribution = noShiftCapMP;
      }
    }
    // Friday/holiday work uses the RAW worked minutes, not net of break — the break those
    // days accrue doesn't get taken off their own Friday/holiday-work total; instead the
    // whole off-day break total comes off the regular-day overtime figure below, once, in
    // aggregate. Either way nets out to the same grand total, but this matches how the two
    // pieces are meant to be read individually.
    // Only bucketed as Friday-work when Friday is genuinely being treated as off for this
    // employee — a shift with followOfficialHolidays='no' treats Friday as an ordinary working
    // day instead, via the normal shiftScheduleToday/بدونِ شیفت path below, not this bucket.
    var dayFridayMin = (isFridayMP && shiftFollowsHolidaysMP) ? dayWorkedMin : 0;
    // Only bucketed as holiday-work when the holiday is genuinely being treated as off for
    // this employee (isHolidayNonFridayMP alone isn't enough — a shift with
    // followOfficialHolidays='no' treats this same calendar date as an ordinary working day
    // instead, via the normal shiftScheduleToday/بدونِ شیفت path below, not this bucket).
    var dayHolidayMin = (isHolidayNonFridayMP && shiftFollowsHolidaysMP) ? dayWorkedMin : 0;

    var markerMP = empMarkersMP[dateStrMP];
    var dayLeaveMin = (markerMP && markerMP.status === 'مرخصی') ? parseLeaveHoursToMinutesShared(markerMP.leaveHours) : 0;
    var dayStatusLabel = markerMP ? (markerMP.status === 'مرخصی' ? 'مرخصی' : (markerMP.status === 'ماموریت' ? 'ماموریت' : (String(markerMP.status).toLowerCase() === 'off' ? 'تعطیل/آف' : (markerMP.status === 'غیبت' ? 'غیبت' : '')))) : '';
    if(dayStatusLabel === 'ماموریت') totals.missionDaysCount++;

    totals.workedMin += dayWorkedMin;
    totals.requiredMin += dayRequiredMin;
    totals.leaveMin += dayLeaveMin;
    totals.nightMin += dayNightMin;
    totals.fridayMin += dayFridayMin;
    totals.holidayMin += dayHolidayMin;
    totals.overtimeMin += dayOvertimeContribution;
    totals.extraPresenceMin += dayExtraPresenceMin;
    // Break is now split: regular working-day break stays in breakMin (it's what dayUsefulMin
    // was net of when computing normal-day overtime above), while Friday/holiday break gets
    // its own separate total — and that total is subtracted from the regular-day overtime
    // figure once, below, rather than from each Friday/holiday day's own work total.
    if(isOffDayMP){ totals.holidayBreakMin += dayBreakMin; } else { totals.breakMin += dayBreakMin; }
    totals.usefulMin += dayUsefulMin;

    dailyRows.push({
      day: dMP, jalaliDate: dateStrMP, isFriday: isFridayMP, isHoliday: isHolidayNonFridayMP, statusLabel: dayStatusLabel,
      workedMin: dayWorkedMin, requiredMin: dayRequiredMin, usefulMin: dayUsefulMin, overtimeMin: dayOvertimeMin,
      leaveMin: dayLeaveMin, nightMin: dayNightMin, fridayMin: dayFridayMin, holidayMin: dayHolidayMin, breakMin: dayBreakMin,
      extraPresenceMin: dayExtraPresenceMin,
      shiftPairs: dayShiftPairs
    });
  }

  // The whole month's Friday/holiday break comes off the regular-day overtime figure here,
  // once, in aggregate — rather than off each Friday/holiday day's own work total (which now
  // holds raw worked minutes, unreduced by break). This still nets out to the exact same
  // grand total either way, but keeps "اضافه‌کاری" and "جمعه‌کاری/تعطیل‌کاری" each reading as
  // the specific thing they're meant to represent.
  totals.overtimeMin -= totals.holidayBreakMin;
  totals.grandOvertimeMin = totals.overtimeMin + totals.fridayMin + totals.holidayMin;
  return {totals: totals, daily: dailyRows};
}
var PERSIAN_WEEKDAYS_SERVER = ['شنبه','يكشنبه','دوشنبه','سه‌شنبه','چهارشنبه','پنجشنبه','جمعه'];
function jalaliWeekdayNameServer(jy, jm, jd){
  var g = jalaliToGregorianServer(jy, jm, jd);
  var dt = new Date(g[0], g[1]-1, g[2]);
  var idx = (dt.getDay() + 1) % 7; // JS getDay(): 0=Sun..6=Sat; شنبه(Sat) should be index 0
  return PERSIAN_WEEKDAYS_SERVER[idx];
}
// Same index calculation as jalaliWeekdayNameServer, but returns the 0-6 index directly
// instead of a name string — used to look up a shift's weeklyScheduleJson entry (built
// frontend-side as an array in شنبه..جمعه order) by position rather than by matching day-name
// strings. Matching by name is fragile here: the frontend's day names use Persian ye/kaf
// (یکشنبه) while PERSIAN_WEEKDAYS_SERVER uses Arabic ye/kaf (يكشنبه) — visually identical,
// different Unicode code points, so a === comparison between them silently never matches.
function jalaliWeekdayIndexServer(jy, jm, jd){
  var g = jalaliToGregorianServer(jy, jm, jd);
  var dt = new Date(g[0], g[1]-1, g[2]);
  return (dt.getDay() + 1) % 7;
}
function parseJalaliDateStrServer(s){
  var parts = String(s||'').trim().split(' ');
  var d = parseInt(parts[0],10);
  var mIdx = J_MONTHS.indexOf(parts[1]) + 1;
  var y = parseInt(parts[parts.length-1],10);
  return {y:y, m:mIdx, d:d};
}
function iterateJalaliDaysServer(startDateStr, endDateStr){
  var sp = parseJalaliDateStrServer(startDateStr);
  var ep = parseJalaliDateStrServer(endDateStr);
  var sg = jalaliToGregorianServer(sp.y, sp.m, sp.d);
  var eg = jalaliToGregorianServer(ep.y, ep.m, ep.d);
  var curDate = new Date(sg[0], sg[1]-1, sg[2]);
  var endDateObj = new Date(eg[0], eg[1]-1, eg[2]);
  var result = [];
  var guard = 0;
  while(curDate <= endDateObj && guard < 400){
    var j = toJalaali(curDate.getFullYear(), curDate.getMonth()+1, curDate.getDate());
    result.push(j[2] + ' ' + J_MONTHS[j[1]-1] + ' ' + j[0]);
    curDate.setDate(curDate.getDate() + 1);
    guard++;
  }
  return result;
}
// Marks yesterday as "آف" (holiday, no work required) in DayStatusMarkers for every active
// employee who had no attendance on it and doesn't already have a leave/mission marker for
// that day. This only runs anything on days that were actually holidays (Fridays or entries
// in the Holidays sheet) — on ordinary working days it does nothing.
//
// IMPORTANT — this function does NOT run on its own. Deploying updated code never installs
// a schedule; a time-driven trigger must be created once, manually, from the Apps Script
// editor: Triggers (clock icon in the left sidebar) → Add Trigger → choose this function
// (runDailyOffMarkerJob) → Event source: Time-driven → Day timer → pick a time after
// midnight (e.g. 00:30–01:00) → Save. Once that one-time setup is done, it runs
// automatically every day without any further action.
function runDailyOffMarkerJob(){
  var yesterdayD = new Date();
  yesterdayD.setDate(yesterdayD.getDate() - 1);
  var yj = toJalaali(yesterdayD.getFullYear(), yesterdayD.getMonth()+1, yesterdayD.getDate());
  var yesterdayStr = yj[2] + ' ' + J_MONTHS[yj[1]-1] + ' ' + yj[0];

  var holInfo = getHolidayInfo(yesterdayStr);
  // No early return here anymore — an ordinary working day still needs checking, since an
  // employee with no attendance and no marker on a real work day should be auto-marked
  // absent, not silently skipped. Only the resulting STATUS differs (Off vs غیبت).

  var esh = getSheet('Employees', ['id','fullName','username','password','deviceId','allowedProjects','employeeCode','freeZone','fingerprintRequired','jobTitle','shiftType','shiftStart','shiftEnd','nationalId','jobGroup','photoUrl','shiftId','active','sessionVersion','noDeviceLimit']);
  var employees = readRows(esh).filter(function(e){ return String(e.active) !== 'false'; });

  var lsh = getSheet('Logs', ['id','jalaliDate','time','employee','username','project','type','lat','lng','accuracy','isoTimestamp','note','employeeCode','isHoliday','holidayReason','source']);
  var allLogs = readRows(lsh);

  var dsmsh = getSheet('DayStatusMarkers', ['id','username','employee','employeeCode','jalaliDate','status','leaveHours','createdAt','createdTime']);
  var existingMarkers = readRows(dsmsh);
  var alreadyMarked = {};
  existingMarkers.filter(function(m){ return m.jalaliDate === yesterdayStr; }).forEach(function(m){ alreadyMarked[m.username] = true; });

  var t = serverNow();
  employees.forEach(function(emp){
    if(alreadyMarked[emp.username]) return; // already has a marker (leave, mission, or a previous off/absence) for this day

    var empLogs = allLogs.filter(function(r){ return r.username === emp.username; }).sort(function(a,b){ return isoCompare(a.isoTimestamp, b.isoTimestamp); });
    var pairs = pairLogsIntoCycles(empLogs);
    var present = false;
    pairs.forEach(function(p){
      if(p.inE && p.inE.jalaliDate === yesterdayStr) present = true;
      if(p.inE && p.outE && p.outE.jalaliDate === yesterdayStr && p.inE.jalaliDate !== yesterdayStr){
        var startDay = p.inE.jalaliDate;
        var firstCycle = pairs.filter(function(p2){ return p2.inE && p2.inE.jalaliDate === startDay; })[0];
        if(firstCycle === p) present = true;
      }
    });
    if(present) return; // actually attended — leave it as real attendance, don't mark anything

    appendRowByHeader(dsmsh, {
      id: Utilities.getUuid(), username: emp.username, employee: emp.fullName, employeeCode: emp.employeeCode||'',
      jalaliDate: yesterdayStr, status: holInfo.isHoliday ? 'Off' : 'غیبت', leaveHours: '', createdAt: t.jalaliDate, createdTime: t.time
    });
  });
}
function syncLeaveDayStatus(username, employee, startDate, endDate, leaveType, startTime, endTime, kind, employeeCode){
  var statusLabel = (kind === 'mission') ? 'ماموریت' : 'مرخصی';
  var dsmshL = getSheet('DayStatusMarkers', ['id','username','employee','employeeCode','jalaliDate','status','leaveHours','createdAt','createdTime']);
  var dates = iterateJalaliDaysServer(startDate, endDate);
  var leaveHoursStr;
  if(leaveType === 'hourly'){
    var stMin = timeToMinutes(startTime);
    var enMin = timeToMinutes(endTime);
    var durMin = enMin - stMin;
    if(durMin <= 0) durMin += 24*60;
    leaveHoursStr = pad2(Math.floor(durMin/60)) + ':' + pad2(durMin%60);
  } else {
    leaveHoursStr = '7:20';
  }
  var tL2 = serverNow();
  var allRows = readRows(dsmshL);
  dates.forEach(function(dateStr){
    var existing = allRows.filter(function(r){ return r.username === username && r.jalaliDate === dateStr; })[0];
    if(existing){
      esh_set(dsmshL, existing._row, 'status', statusLabel);
      esh_set(dsmshL, existing._row, 'leaveHours', leaveHoursStr);
    } else {
      var idL = Utilities.getUuid();
      appendRowByHeader(dsmshL, {id:idL, username:username, employee:employee, employeeCode: employeeCode||'', jalaliDate:dateStr, status:statusLabel, leaveHours:leaveHoursStr, createdAt: tL2.jalaliDate, createdTime: tL2.time});
    }
  });
}
function addDaysToJalaliServer(dateStr, days){
  var p = parseJalaliDateStrServer(dateStr);
  var g = jalaliToGregorianServer(p.y, p.m, p.d);
  var dt = new Date(g[0], g[1]-1, g[2]);
  dt.setDate(dt.getDate() + days);
  var j = toJalaali(dt.getFullYear(), dt.getMonth()+1, dt.getDate());
  return j[2] + ' ' + J_MONTHS[j[1]-1] + ' ' + j[0];
}
function deleteOneLogRow(id){
  var lsh6 = getSheet('Logs', ['id','jalaliDate','time','employee','username','project','type','lat','lng','accuracy','isoTimestamp','note','employeeCode','isHoliday','holidayReason','source']);
  var row7 = ensureIds(lsh6, readRows(lsh6)).filter(function(r){ return r.id === id; })[0];
  if(row7) lsh6.deleteRow(row7._row);
  return {ok:true};
}
function setRowField(sheet, rowValues, headerName, value){
  var idx = colIndex(sheet, headerName) - 1;
  if(idx >= 0) rowValues[idx] = value;
}
function updateOneLogRow(body){
  var lsh5 = getSheet('Logs', ['id','jalaliDate','time','employee','username','project','type','lat','lng','accuracy','isoTimestamp','note','employeeCode','isHoliday','holidayReason','source']);
  var rows7 = ensureIds(lsh5, readRows(lsh5));
  var row6 = rows7.filter(function(r){ return r.id === body.id; })[0];
  if(!row6) return {ok:false, error:'not found'};

  var newIsoCheck = jalaliDateTimeToIso(body.jalaliDate, body.time);
  if(newIsoCheck){
    // Fully free-form editing (any date/time, any order) — the only real requirement is that
    // it doesn't land at the exact same instant as another existing entry for this employee.
    var conflictRow = rows7.filter(function(r){ return r.id !== body.id && r.username === row6.username && new Date(r.isoTimestamp).getTime() === new Date(newIsoCheck).getTime(); })[0];
    if(conflictRow){
      return {ok:false, error:'time_conflict', at: conflictRow.jalaliDate + ' ' + conflictRow.time};
    }
  }
  // A checkout must always have an open check-in before it, chronologically — this was the
  // one gap left: adminAddLog (both the plain single-entry form and the report edit view's
  // "create a new side" path) already had this check, but EDITING an existing row straight to
  // 'خروج' (or moving its time earlier than the check-in it was supposed to close) never went
  // through that check at all. The row itself is excluded from the comparison set since we're
  // changing it, not adding a new one alongside it — and this is still purely chronological,
  // never same-calendar-day matching, so a legitimate overnight shift is never blocked by it.
  var newTypeCheck = body.type === 'in' ? 'ورود' : (body.type === 'out' ? 'خروج' : body.type);
  if(newTypeCheck === 'خروج' && newIsoCheck){
    var empRowsForOrphanCheck = rows7.filter(function(r){ return r.id !== body.id && r.username === row6.username; });
    if(wouldBeOrphanCheckout(empRowsForOrphanCheck, newIsoCheck)){
      return {ok:false, error:'no_open_checkin'};
    }
  }

  var lastCol = lsh5.getLastColumn();
  var rowRange = lsh5.getRange(row6._row, 1, 1, lastCol);
  var rowValues = rowRange.getValues()[0];

  setRowField(lsh5, rowValues, 'jalaliDate', body.jalaliDate);
  setRowField(lsh5, rowValues, 'time', body.time);
  setRowField(lsh5, rowValues, 'employee', body.employee);
  setRowField(lsh5, rowValues, 'project', body.project);
  setRowField(lsh5, rowValues, 'type', body.type === 'in' ? 'ورود' : (body.type === 'out' ? 'خروج' : body.type));
  var newIso = jalaliDateTimeToIso(body.jalaliDate, body.time);
  if(newIso) setRowField(lsh5, rowValues, 'isoTimestamp', newIso);

  var eshFind = getSheet('Employees', ['id','fullName','username','password','deviceId','allowedProjects','employeeCode','freeZone','fingerprintRequired','jobTitle','shiftType','shiftStart','shiftEnd','nationalId','jobGroup','photoUrl','shiftId','active','sessionVersion','noDeviceLimit']);
  var empMatch = readRows(eshFind).filter(function(r){ return r.fullName === body.employee; })[0];
  if(empMatch){
    setRowField(lsh5, rowValues, 'username', empMatch.username);
    setRowField(lsh5, rowValues, 'employeeCode', empMatch.employeeCode||'');
  }
  var holInfo4 = getHolidayInfo(body.jalaliDate);
  setRowField(lsh5, rowValues, 'isHoliday', holInfo4.isHoliday ? 'تعطیل' : '');
  setRowField(lsh5, rowValues, 'holidayReason', holInfo4.reason);

  var noteStr = String(row6.note || '');
  var editMatch = noteStr.match(/ویرایش‌شده توسط مدیر \((\d+)\)/);
  if(editMatch){
    var newCount = parseInt(editMatch[1], 10) + 1;
    noteStr = noteStr.replace(/ویرایش‌شده توسط مدیر \(\d+\)/, 'ویرایش‌شده توسط مدیر ('+newCount+')');
  } else {
    noteStr = (noteStr ? noteStr + ' | ' : '') + 'ویرایش‌شده توسط مدیر (1)';
  }
  setRowField(lsh5, rowValues, 'note', noteStr);
  rowRange.setNumberFormat('@'); // plain text — prevents Sheets from auto-converting the
  // time/date strings we're about to write into its own internal date/time type, which
  // would otherwise come back from the API as a garbled 1899-12-30T...Z value.
  rowRange.setValues([rowValues]);
  SpreadsheetApp.flush();
  return {ok:true};
}
function jalaliDateTimeToIso(jalaliDateStr, timeStr){
  var p = String(jalaliDateStr||'').trim().split(' ');
  if(p.length < 3) return null;
  var day = parseInt(p[0],10);
  var monthIdx = J_MONTHS.indexOf(p[1]) + 1;
  var year = parseInt(p[p.length-1],10);
  if(!day || !monthIdx || !year) return null;
  var g = jalaliToGregorianServer(year, monthIdx, day);
  var tParts = String(timeStr||'00:00:00').split(':');
  var hh = parseInt(tParts[0],10)||0, mm = parseInt(tParts[1],10)||0, ss = parseInt(tParts[2],10)||0;
  var dateStr = g[0]+'-'+pad2(g[1])+'-'+pad2(g[2])+' '+pad2(hh)+':'+pad2(mm)+':'+pad2(ss);
  var d = Utilities.parseDate(dateStr, TZ, 'yyyy-MM-dd HH:mm:ss');
  return d.toISOString();
}
function isFridayFromJalaliDateStr(jalaliDateStr){
  var p = String(jalaliDateStr||'').trim().split(' ');
  if(p.length < 3) return false;
  var day = parseInt(p[0],10);
  var monthIdx = J_MONTHS.indexOf(p[1]) + 1;
  var year = parseInt(p[p.length-1],10);
  if(!day || !monthIdx || !year) return false;
  var g = jalaliToGregorianServer(year, monthIdx, day);
  return new Date(g[0], g[1]-1, g[2]).getDay() === 5;
}
function getHolidayInfo(jalaliDateStr){
  if(isFridayFromJalaliDateStr(jalaliDateStr)) return {isHoliday:true, reason:'جمعه'};
  var hsh = getSheet('Holidays', ['id','jalaliDate','reason','status']);
  var match = readRows(hsh).filter(function(r){ return r.jalaliDate === jalaliDateStr; })[0];
  if(match) return {isHoliday:true, reason: match.reason || 'تعطیل رسمی'};
  return {isHoliday:false, reason:''};
}
function serverNow(){
  var now = new Date();
  var gy = parseInt(Utilities.formatDate(now, TZ, 'yyyy'), 10);
  var gm = parseInt(Utilities.formatDate(now, TZ, 'MM'), 10);
  var gd = parseInt(Utilities.formatDate(now, TZ, 'dd'), 10);
  var hh = parseInt(Utilities.formatDate(now, TZ, 'HH'), 10);
  var mi = parseInt(Utilities.formatDate(now, TZ, 'mm'), 10);
  var ss = parseInt(Utilities.formatDate(now, TZ, 'ss'), 10);
  var hms = pad2(hh)+':'+pad2(mi)+':'+pad2(ss);
  var j = toJalaali(gy, gm, gd);
  return {
    jalaliDate: j[2] + ' ' + J_MONTHS[j[1]-1] + ' ' + j[0],
    time: hms,
    isoTimestamp: now.toISOString()
  };
}

// ============================================================
// doGet — public reads (project list) + PIN-gated report export
// ============================================================
function doGet(e){
  var action = e.parameter.action;

  if(action === 'projects'){
    var sh = getSheet('Projects', ['id','name','lat','lng','radius','city','province','restricted','active']);
    // Inactive projects are hidden from this public list entirely — an employee should never
    // even see one to pick from, let alone log time against it. Existing rows with no active
    // value at all (from before this feature existed) are treated as active, so nothing that
    // was already visible silently disappears just because this column is new.
    var activeProjectsOnly = readRows(sh).filter(function(r){ return String(r.active||'yes') !== 'no'; });
    return jsonOut({ok:true, projects: activeProjectsOnly});
  }
  if(action === 'holidaysPublic'){
    var hshp = getSheet('Holidays', ['id','jalaliDate','reason','status']);
    return jsonOut({ok:true, holidays: readRows(hshp).map(function(r){ delete r._row; return r; })});
  }

  if(action === 'logs'){
    if(String(e.parameter.pin) !== getAdminPin()){
      return jsonOut({ok:false, error:'invalid pin'});
    }
    var lsh = getSheet('Logs', ['id','jalaliDate','time','employee','username','project','type','lat','lng','accuracy','isoTimestamp','note','employeeCode','isHoliday','holidayReason','source']);
    var rows = ensureIds(lsh, readRows(lsh)).map(function(r){ delete r._row; return r; });
    return jsonOut({ok:true, logs: rows});
  }

  return jsonOut({ok:false, error:'unknown action'});
}

// ============================================================
// doPost — everything else
// ============================================================
// Wraps the real handler in a try/catch so any hidden runtime exception (a null reference,
// a broken sheet read, etc.) comes back as a visible, specific error message instead of
// whatever generic fallback text Apps Script's own infrastructure might otherwise produce.
function doPost(e){
  try {
    return doPostInner(e);
  } catch(errTop){
    return jsonOut({ok:false, error: 'exception: ' + (errTop && errTop.message ? errTop.message : String(errTop))});
  }
}
function doPostInner(e){
  var body = JSON.parse(e.postData.contents);
  var action = body.action;
  // Defensively normalize username/id-like fields on the way in too — if a purely-numeric
  // username was ever fetched by the app from a cell Sheets had auto-typed as a number, the
  // client would hold (and send back) a JS number, not a string. Coercing here means every
  // comparison against readRows() output (which is now always string-typed — see readRows)
  // stays reliable regardless of what type happened to arrive.
  if(body.username !== undefined && body.username !== null) body.username = String(body.username);
  if(body.entry && body.entry.username !== undefined && body.entry.username !== null) body.entry.username = String(body.entry.username);
  if(body.deviceId !== undefined && body.deviceId !== null) body.deviceId = String(body.deviceId);

  if(action === 'ping'){
    return jsonOut({ok:true, version: 'v-2026-09-14-security-fix-projects-active-reset', hasUpdateOneLogRow: (typeof updateOneLogRow === 'function'), hasGetLeaveRequestsForUser: true});
  }

  // Public — no PIN required, since this needs to display on the splash screen before login.
  if(action === 'getCompanyBranding'){
    return jsonOut({ok:true, companyName: getSetting('company_name'), companyLogo: getSetting('company_logo')});
  }
  if(action === 'setCompanyBranding'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    if(body.companyName !== undefined) setSetting('company_name', body.companyName);
    if(body.companyLogo) setSetting('company_logo', body.companyLogo);
    return jsonOut({ok:true});
  }

  // ---------------- login (device-locked) ----------------
  // ---------------- super-admin settings (license window + max employees) ----------------
  // Completely separate gate from the regular admin PIN — reached via the hidden ⚙️ gear on
  // the admin login screen, authenticated with its own username/password (SUPER_ADMIN_USERNAME
  // / SUPER_ADMIN_PASSWORD at the top of this file). Re-checks the credentials on every call,
  // stateless, matching how ADVANCED_TOOLS_PASSWORD already works elsewhere in this file.
  function checkSuperAdminAuth(b){
    return String(b.suUsername||'').toLowerCase() === SUPER_ADMIN_USERNAME.toLowerCase() && String(b.suPassword||'') === SUPER_ADMIN_PASSWORD;
  }
  if(action === 'superAdminLogin'){
    return jsonOut({ok: checkSuperAdminAuth(body)});
  }
  if(action === 'getSuperAdminSettings'){
    if(!checkSuperAdminAuth(body)) return jsonOut({ok:false, error:'invalid_credentials'});
    return jsonOut({ok:true, licenseStart: getSetting('license_start_date'), licenseEnd: getSetting('license_end_date'), maxEmployees: getSetting('max_employees')});
  }
  if(action === 'saveSuperAdminSettings'){
    if(!checkSuperAdminAuth(body)) return jsonOut({ok:false, error:'invalid_credentials'});
    // Both dates must be set together or both left empty — a lone start or end date would be
    // ambiguous (unlimited on one side, expired forever on the other) and isLicenseValid()
    // only treats "both present" as an actual restriction anyway.
    if((body.licenseStart && !body.licenseEnd) || (!body.licenseStart && body.licenseEnd)){
      return jsonOut({ok:false, error:'both_dates_required'});
    }
    setSetting('license_start_date', body.licenseStart || '');
    setSetting('license_end_date', body.licenseEnd || '');
    setSetting('max_employees', body.maxEmployees !== undefined ? String(parseInt(body.maxEmployees,10)||0) : '');
    return jsonOut({ok:true});
  }
  if(action === 'resetAllData'){
    // Wipes every data sheet down to just its header row — for spinning up a genuinely fresh
    // install for a new client, not for routine use. Gated by the same super-admin
    // credentials as the rest of this settings panel, and the frontend already requires typing
    // a confirmation phrase plus a second styled confirm dialog before this is ever called.
    if(!checkSuperAdminAuth(body)) return jsonOut({ok:false, error:'invalid_credentials'});
    var sheetsToWipe = [
      'Logs','Employees','Projects','Requests','LeaveRequests','PendingOfflineLogs','DayStatusMarkers',
      'Holidays','Shifts','ShiftDays','ShiftAssignments','AdminSessions','Settings'
    ];
    sheetsToWipe.forEach(function(name){
      var sh;
      try { sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); } catch(e){ sh = null; }
      if(!sh) return;
      var lastRow = sh.getLastRow();
      if(lastRow > 1) sh.deleteRows(2, lastRow - 1);
    });
    // Settings was wiped along with everything else above, including admin_pin/admin_username
    // — explicitly restored to the plain admin/admin default right after, so the fresh install
    // isn't accidentally left with no way to log into the admin panel at all.
    setSetting('admin_username', 'admin');
    setSetting('admin_pin', 'admin');
    return jsonOut({ok:true});
  }
  if(action === 'login'){
    if(!isLicenseValid()) return jsonOut({ok:false, error:'license_expired'});
    var esh = getSheet('Employees', ['id','fullName','username','password','deviceId','allowedProjects','employeeCode','freeZone','fingerprintRequired','jobTitle','shiftType','shiftStart','shiftEnd','nationalId','jobGroup','photoUrl','shiftId','active','sessionVersion','noDeviceLimit']);
    var erows = readRows(esh);
    var match = erows.filter(function(r){
      return String(r.username).toLowerCase() === String(body.username).toLowerCase() &&
             String(r.password) === String(body.password);
    })[0];
    if(!match) return jsonOut({ok:false, error:'invalid_credentials'});
    if(String(match.active) === 'false'){
      return jsonOut({ok:false, error:'account_disabled'});
    }

    var incomingDevice = String(body.deviceId || '');
    var noDeviceLimitLogin = String(match.noDeviceLimit||'').toLowerCase() === 'yes';
    if(!noDeviceLimitLogin){
      if(!match.deviceId){
        // first login ever — bind this device to the account
        esh.getRange(match._row, colIndex(esh,'deviceId')).setValue(incomingDevice);
      } else if(String(match.deviceId) !== incomingDevice){
        return jsonOut({ok:false, error:'device_mismatch'});
      }
    }
    // when noDeviceLimit is on, deviceId is never checked or updated — this employee can log
    // in from any phone, every time, with no binding at all.

    // sessionVersion — a stamp the employee's app caches at login and re-checks periodically
    // (see checkSessionVersion below). Bumped by the admin editing this employee or resetting
    // their device, so an already-logged-in phone gets forced back to the login screen soon
    // after any such change, rather than silently staying "logged in" under stale info.
    var sessionVerLogin = match.sessionVersion || '';
    if(!sessionVerLogin){
      sessionVerLogin = String(Date.now());
      esh.getRange(match._row, colIndex(esh,'sessionVersion')).setValue(sessionVerLogin);
    }

    var effShiftLogin = {type:'simple', shiftStart: match.shiftStart||'', shiftEnd: match.shiftEnd||'', requiredMinutes:''};
    if(match.shiftId){
      var shdefsLogin = getSheet('Shifts', ['id','name','type','shiftStart','shiftEnd','requiredMinutes','lateTolerance','earlyTolerance','shiftYear','floatingMinutes','dailyOvertimeCapMinutes','weeklyScheduleJson','createdAt']);
      var shdefLogin = readRows(shdefsLogin).filter(function(r){ return r.id === match.shiftId; })[0];
      if(shdefLogin){
        effShiftLogin = {type: shdefLogin.type||'simple', shiftStart: shdefLogin.shiftStart||'', shiftEnd: shdefLogin.shiftEnd||'', requiredMinutes: shdefLogin.requiredMinutes||''};
      }
    }

    return jsonOut({
      ok:true,
      fullName: match.fullName,
      username: match.username,
      employeeCode: match.employeeCode || '',
      freeZone: !!match.freeZone,
      fingerprintRequired: !!match.fingerprintRequired,
      jobTitle: match.jobTitle||'', shiftType: effShiftLogin.type, shiftStart: effShiftLogin.shiftStart, shiftEnd: effShiftLogin.shiftEnd, shiftRequiredMinutes: effShiftLogin.requiredMinutes,
      allowedProjects: match.allowedProjects || '',  // '' or 'ALL' = unrestricted; else comma-separated project ids
      sessionVersion: sessionVerLogin,
      photoUrl: match.photoUrl || ''
    });
  }

  // Lightweight periodic check the employee's app calls while already "logged in" (from its
  // own cached localStorage session, without re-entering credentials) — confirms the admin
  // hasn't edited this employee or reset their device since login. A mismatch means: log out.
  if(action === 'checkSessionVersion'){
    var eshSV = getSheet('Employees', ['id','fullName','username','password','deviceId','allowedProjects','employeeCode','freeZone','fingerprintRequired','jobTitle','shiftType','shiftStart','shiftEnd','nationalId','jobGroup','photoUrl','shiftId','active','sessionVersion','noDeviceLimit']);
    var matchSV = readRows(eshSV).filter(function(r){ return r.username === body.username; })[0];
    if(!matchSV) return jsonOut({ok:true, valid:false, reason:'not_found'});
    if(String(matchSV.active) === 'false') return jsonOut({ok:true, valid:false, reason:'account_disabled'});
    var currentSV = matchSV.sessionVersion || '';
    var stillValid = currentSV && String(currentSV) === String(body.sessionVersion||'');
    return jsonOut({ok:true, valid: !!stillValid});
  }

  if(action === 'checkPin'){
    if(!isLicenseValid()) return jsonOut({ok:false, error:'license_expired'});
    var userOk = String(body.username||'').toLowerCase() === getAdminUsername().toLowerCase();
    var passOk = String(body.pin) === getAdminPin();
    if(!userOk || !passOk) return jsonOut({ok:false});

    // Limit concurrent admin devices to 2. The same device logging in again just refreshes
    // its own session (doesn't count as a new one); a genuinely new 3rd device gets rejected
    // until the admin frees up a slot from Settings.
    var deviceIdCP = String(body.deviceId||'').trim();
    if(deviceIdCP){
      var assh = getSheet('AdminSessions', ['id','deviceId','deviceLabel','loginTime','lastActiveTime']);
      var sessions = readRows(assh);
      var mySession = sessions.filter(function(s){ return s.deviceId === deviceIdCP; })[0];
      var tCP = serverNow();
      if(mySession){
        assh.getRange(mySession._row, colIndex(assh,'lastActiveTime')).setValue(tCP.jalaliDate + ' ' + tCP.time);
      } else {
        if(sessions.length >= 2){
          return jsonOut({ok:false, error:'admin_device_limit_reached'});
        }
        appendRowByHeader(assh, {
          id: Utilities.getUuid(), deviceId: deviceIdCP, deviceLabel: body.deviceLabel || 'دستگاه ناشناس',
          loginTime: tCP.jalaliDate + ' ' + tCP.time, lastActiveTime: tCP.jalaliDate + ' ' + tCP.time
        });
      }
    }
    return jsonOut({ok:true});
  }
  if(action === 'getAdminSessions'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var asshG = getSheet('AdminSessions', ['id','deviceId','deviceLabel','loginTime','lastActiveTime']);
    var rowsG = readRows(asshG).map(function(r){ delete r._row; return r; });
    return jsonOut({ok:true, sessions: rowsG, thisDeviceId: String(body.deviceId||'')});
  }
  if(action === 'removeAdminSession'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var asshR = getSheet('AdminSessions', ['id','deviceId','deviceLabel','loginTime','lastActiveTime']);
    var rowR = readRows(asshR).filter(function(r){ return r.id === body.id; })[0];
    if(!rowR) return jsonOut({ok:false, error:'not found'});
    asshR.deleteRow(rowR._row);
    return jsonOut({ok:true});
  }

  // ---------------- attendance logging (server time is authoritative) ----------------
  if(action === 'addLog'){
    var esh2 = getSheet('Employees', ['id','fullName','username','password','deviceId','allowedProjects','employeeCode','freeZone','fingerprintRequired','jobTitle','shiftType','shiftStart','shiftEnd','nationalId','jobGroup','photoUrl','shiftId','active','sessionVersion','noDeviceLimit']);
    var emp = readRows(esh2).filter(function(r){ return r.username === body.entry.username; })[0];
    if(!emp) return jsonOut({ok:false, error:'unknown_employee'});
    // The periodic sessionVersion poll usually force-logs-out a just-deactivated employee
    // within ~20 seconds, but that's not instant — without this check, a check-in submitted
    // during that short window would still silently go through and get recorded.
    if(String(emp.active) === 'false'){
      return jsonOut({ok:false, error:'account_disabled'});
    }
    if(emp.deviceId && String(emp.deviceId) !== String(body.deviceId||'')){
      return jsonOut({ok:false, error:'device_mismatch'});
    }
    if(emp.allowedProjects && emp.allowedProjects !== 'ALL'){
      var allowed = String(emp.allowedProjects).split(',').map(function(s){ return s.trim(); });
      if(allowed.indexOf(String(body.project_id||'')) === -1){
        return jsonOut({ok:false, error:'project_not_allowed'});
      }
    }
    // Enforced here too, not just by hiding inactive projects from the frontend's picker — a
    // direct API call (or a stale cached project list on the employee's device) must never be
    // able to still log time against a project the admin has deactivated.
    if(body.project_id){
      var pshAL = getSheet('Projects', ['id','name','lat','lng','radius','city','province','restricted','active']);
      var projRowAL = readRows(pshAL).filter(function(r){ return r.id === String(body.project_id); })[0];
      if(projRowAL && String(projRowAL.active||'yes') === 'no'){
        return jsonOut({ok:false, error:'project_inactive'});
      }
    }
    if(!emp.freeZone){
      var psh0 = getSheet('Projects', ['id','name','lat','lng','radius','city','province','restricted','active']);
      var proj0 = readRows(psh0).filter(function(r){ return r.id === body.project_id; })[0];
      // Only skip the geofence when a project is EXPLICITLY marked unrestricted ('no') — an
      // empty/missing value (all existing projects, before this field existed) defaults to
      // restricted so nothing silently loses its geofence just because this column is new.
      if(proj0 && proj0.restricted !== 'no'){
        var dist0 = distMeters(Number(body.entry.lat), Number(body.entry.lng), Number(proj0.lat), Number(proj0.lng));
        if(dist0 > Number(proj0.radius)){
          return jsonOut({ok:false, error:'outside_geofence'});
        }
      }
    }
    if(emp.fingerprintRequired && body.entry.supportsFingerprint && !body.entry.fingerprintVerified){
      return jsonOut({ok:false, error:'fingerprint_required'});
    }

    var d = body.entry;
    var isOffline = !!body.offline;
    var t;
    var note = '';
    if(isOffline && d.clientIso){
      // Trust the device clock ONLY for offline-queued entries, and flag it clearly for the admin.
      var offDate = new Date(d.clientIso);
      var ogy = parseInt(Utilities.formatDate(offDate, TZ, 'yyyy'), 10);
      var ogm = parseInt(Utilities.formatDate(offDate, TZ, 'MM'), 10);
      var ogd = parseInt(Utilities.formatDate(offDate, TZ, 'dd'), 10);
      var ohh = parseInt(Utilities.formatDate(offDate, TZ, 'HH'), 10);
      var omi = parseInt(Utilities.formatDate(offDate, TZ, 'mm'), 10);
      var oss = parseInt(Utilities.formatDate(offDate, TZ, 'ss'), 10);
      var ohms = pad2(ohh)+':'+pad2(omi)+':'+pad2(oss);
      var oj = toJalaali(ogy, ogm, ogd);
      t = {
        jalaliDate: oj[2] + ' ' + J_MONTHS[oj[1]-1] + ' ' + oj[0],
        time: ohms,
        isoTimestamp: offDate.toISOString()
      };
      note = '⚠ ثبت آفلاین — بر اساس ساعت دستگاه (نه سرور)';
    } else {
      t = serverNow(); // <-- authoritative; client-sent date/time is ignored
      if(d.clientIso){
        var drift = Math.abs(new Date(t.isoTimestamp).getTime() - new Date(d.clientIso).getTime());
        if(drift > 10*60*1000){ note = '⚠ اختلاف ساعت گوشی با سرور بیش از ۱۰ دقیقه بود'; }
      }
    }
    if(d.accuracy && Number(d.accuracy) > 100){
      note = (note ? note + ' | ' : '') + '⚠ دقت موقعیت پایین بود (±' + Math.round(Number(d.accuracy)) + 'م)';
    }

    var lsh2 = getSheet('Logs', ['id','jalaliDate','time','employee','username','project','type','lat','lng','accuracy','isoTimestamp','note','employeeCode','isHoliday','holidayReason','source']);
    var empRows = readRows(lsh2).filter(function(r){ return r.username === d.username; });
    empRows.sort(function(a,b){ return isoCompare(a.isoTimestamp, b.isoTimestamp); });
    var lastEntry = empRows.length ? empRows[empRows.length-1] : null;
    var wantIn = (d.type === 'in');
    // Up to TWO check-ins per calendar day are allowed — but the limit is specifically on
    // check-INS started that day, not on raw row count. This matters for night shifts: a
    // check-out that lands on the NEXT day (crossing midnight) must not eat into that next
    // day's own two-check-in allowance — it's just closing out the previous day's session,
    // not something new happening "today". A check-out never needs this check at all, since
    // it can only ever happen right after a legitimate, already-validated check-in anyway.
    if(wantIn){
      var todaysCheckins = empRows.filter(function(r){ return r.jalaliDate === t.jalaliDate && r.type === 'ورود'; }).length;
      if(todaysCheckins >= 2){
        return jsonOut({ok:false, error:'daily_limit_reached'});
      }
    }
    if(wantIn && lastEntry && lastEntry.type === 'ورود'){
      return jsonOut({ok:false, error:'already_checked_in'});
    }
    if(!wantIn){
      if(!lastEntry || lastEntry.type === 'خروج'){
        return jsonOut({ok:false, error:'not_checked_in'});
      }
      var minutesSinceIn = (new Date(t.isoTimestamp).getTime() - new Date(lastEntry.isoTimestamp).getTime()) / 60000;
      if(minutesSinceIn < 5){
        return jsonOut({ok:false, error:'checkout_too_soon'});
      }
      if(minutesSinceIn > 24*60){
        return jsonOut({ok:false, error:'checkout_expired'});
      }
    }

    var effShift = getEffectiveShift(d.username, t.jalaliDate, emp);
    // This specific date's own ShiftDays entry (if any) — used below so a per-day attendance
    // window override takes precedence over the shift-level default, matching how
    // computeEmployeePerformance itself resolves the same override for other per-day settings.
    var sdEntryLive = null;
    if(effShift && effShift.id){
      var sdShLive = getSheet('ShiftDays', ['id','shiftId','jalaliDate','dayType','start1','end1','hasSecondPart','start2','end2',
        'dailyRequiredMinutes','dayFloatingMinutes','dayOvertimeCapMinutes','dayOtStartMinutes','dayOtMiddleMinutes','dayOtEndMinutes',
        'dayLateAllowedMinutes','dayEarlyAllowedMinutes','dayFloatBeforeShiftMinutes','dayAttendanceWindowStart','dayAttendanceWindowEnd',
        'isNightShift','nightShiftCrossesNextDay','nightShiftStart','nightShiftDuration','createdAt']);
      sdEntryLive = readRows(sdShLive).filter(function(r){ return r.shiftId === effShift.id && r.jalaliDate === t.jalaliDate; })[0];
    }
    var effWindowStart = dayOrShiftValue(sdEntryLive && sdEntryLive.dayAttendanceWindowStart, effShift.attendanceWindowStart);
    var effWindowEnd = dayOrShiftValue(sdEntryLive && sdEntryLive.dayAttendanceWindowEnd, effShift.attendanceWindowEnd);

    // Attendance window — "پرسنل قبل از این ساعت مجاز به ثبتِ تردد نمی‌باشند": if the shift
    // sets either boundary, check-in/out is rejected outright outside that window. Only
    // applies when the shift actually configures one of these fields — شیفتِ عادی and any
    // shift that leaves them blank are completely unaffected. Assumes a same-day window
    // (start time before end time); an overnight window (e.g. 22:00–06:00) isn't handled here.
    if(effWindowStart || effWindowEnd){
      var nowMinAW = timeToMinutes(t.time);
      if(effWindowStart && nowMinAW < timeToMinutes(effWindowStart)){
        return jsonOut({ok:false, error:'outside_attendance_window', windowStart: effWindowStart, windowEnd: effWindowEnd});
      }
      if(effWindowEnd && nowMinAW > timeToMinutes(effWindowEnd)){
        return jsonOut({ok:false, error:'outside_attendance_window', windowStart: effWindowStart, windowEnd: effWindowEnd});
      }
    }

    if(effShift.type === 'floating'){
      if(!wantIn && effShift.requiredMinutes && lastEntry){
        var workedMin = Math.round((new Date(t.isoTimestamp).getTime() - new Date(lastEntry.isoTimestamp).getTime()) / 60000);
        var reqMin = Number(effShift.requiredMinutes) || 0;
        var diffFloat = workedMin - reqMin;
        if(diffFloat > 0){ note = (note ? note + ' | ' : '') + '⏰ اضافه‌کاری: ' + diffFloat + ' دقیقه (شناور — کارکرد ' + workedMin + ' از ' + reqMin + ' دقیقه موظف)'; }
        else if(diffFloat < 0){ note = (note ? note + ' | ' : '') + '⏰ کسری کار: ' + (-diffFloat) + ' دقیقه (شناور — کارکرد ' + workedMin + ' از ' + reqMin + ' دقیقه موظف)'; }
      }
    } else {
      if(effShift.shiftStart && wantIn){
        var diffIn = timeToMinutes(t.time) - timeToMinutes(effShift.shiftStart) - (effShift.lateTolerance||0);
        if(diffIn > 0){ note = (note ? note + ' | ' : '') + '⏰ تاخیر ورود: ' + diffIn + ' دقیقه'; }
      }
      if(effShift.shiftEnd && !wantIn){
        var diffOut = timeToMinutes(effShift.shiftEnd) - timeToMinutes(t.time) - (effShift.earlyTolerance||0);
        if(diffOut > 0){ note = (note ? note + ' | ' : '') + '⏰ تعجیل خروج: ' + diffOut + ' دقیقه'; }
        else if(diffOut < 0){ note = (note ? note + ' | ' : '') + '⏰ اضافه‌کاری: ' + (-diffOut) + ' دقیقه'; }
      }
    }

    var holInfo = getHolidayInfo(t.jalaliDate);
    var id = Utilities.getUuid();
    appendRowByHeader(lsh2, {
      id:id, jalaliDate:t.jalaliDate, time:t.time, employee:d.employee, username:d.username, project:d.project,
      type: d.type === 'in' ? 'ورود' : 'خروج', lat:d.lat, lng:d.lng, accuracy:d.accuracy,
      isoTimestamp:t.isoTimestamp, note:note, employeeCode: emp.employeeCode || '',
      isHoliday: holInfo.isHoliday ? 'تعطیل' : '', holidayReason: holInfo.reason, source:'اتومات'
    });
    return jsonOut({ok:true, jalaliDate:t.jalaliDate, time:t.time, isoTimestamp:t.isoTimestamp});
  }

  if(action === 'getMyLogs'){
    var lsh3 = getSheet('Logs', ['id','jalaliDate','time','employee','username','project','type','lat','lng','accuracy','isoTimestamp','note','employeeCode','isHoliday','holidayReason','source']);
    var rows3 = ensureIds(lsh3, readRows(lsh3)).filter(function(r){ return r.username === body.username; })
      .map(function(r){ delete r._row; return r; });
    var posh3 = getSheet('PendingOfflineLogs', ['id','username','employee','employeeCode','project','type','jalaliDate','time','clientIso','lat','lng','accuracy','reason','status','createdAt']);
    var pendingOffline3 = readRows(posh3).filter(function(r){ return r.username === body.username && r.status === 'pending'; })
      .map(function(r){ return {jalaliDate:r.jalaliDate, time:r.time, type:r.type, project:r.project}; });
    // photoUrl is looked up fresh here too (not just at login) — this call already runs
    // frequently (app open, periodic polling), so a photo the admin adds/changes shows up
    // for the employee well before their next full login/force-logout cycle would.
    var eshGML = getSheet('Employees', ['id','fullName','username','password','deviceId','allowedProjects','employeeCode','freeZone','fingerprintRequired','jobTitle','shiftType','shiftStart','shiftEnd','nationalId','jobGroup','photoUrl','shiftId','active','sessionVersion','noDeviceLimit']);
    var empGML = readRows(eshGML).filter(function(r){ return r.username === body.username; })[0];
    return jsonOut({ok:true, logs: rows3, pendingOffline: pendingOffline3, today: serverNow().jalaliDate, photoUrl: empGML ? (empGML.photoUrl||'') : ''});
  }

  // ---------------- manual entry requests ----------------
  if(action === 'createRequest'){
    var esh3 = getSheet('Employees', ['id','fullName','username','password','deviceId','allowedProjects','employeeCode','freeZone','fingerprintRequired','jobTitle','shiftType','shiftStart','shiftEnd','nationalId','jobGroup','photoUrl','shiftId','active','sessionVersion','noDeviceLimit']);
    var emp2 = readRows(esh3).filter(function(r){ return r.username === body.username; })[0];
    if(!emp2) return jsonOut({ok:false, error:'unknown_employee'});
    if(emp2.deviceId && String(emp2.deviceId) !== String(body.deviceId||'')){
      return jsonOut({ok:false, error:'device_mismatch'});
    }
    var rsh = getSheet('Requests', ['id','username','employee','employeeCode','project','type','reason','status','createdAt','createdTime','requestedDate','requestedTime']);
    var t2 = serverNow();
    var normalizedReqTime = normalizeTimeInput(body.requestedTime);
    if(body.requestedDate && normalizedReqTime){
      var reqIso = jalaliDateTimeToIso(body.requestedDate, normalizedReqTime);
      if(reqIso){
        var lshChk = getSheet('Logs', ['id','jalaliDate','time','employee','username','project','type','lat','lng','accuracy','isoTimestamp','note','employeeCode','isHoliday','holidayReason','source']);
        var empRowsChk = readRows(lshChk).filter(function(r){ return r.username === body.username; });
        var lastChk = empRowsChk.length ? empRowsChk.sort(function(a,b){ return isoCompare(a.isoTimestamp, b.isoTimestamp); })[empRowsChk.length-1] : null;
        if(lastChk && String(reqIso) <= String(lastChk.isoTimestamp)){
          return jsonOut({ok:false, error:'must_be_after_last_entry'});
        }
      }
    }
    var rid = Utilities.getUuid();
    appendRowByHeader(rsh, {
      id:rid, username:body.username, employee:emp2.fullName, employeeCode: emp2.employeeCode||'', project:body.project, type:body.type,
      reason: body.reason||'', status:'pending', createdAt:t2.jalaliDate, createdTime:t2.time,
      requestedDate: body.requestedDate||'', requestedTime: normalizedReqTime
    });
    notifyAdmin(
      'درخواست ثبت دستی جدید — ' + emp2.fullName,
      emp2.fullName + ' یک درخواست ثبت ' + (body.type==='in'?'ورود':'خروج') + ' برای پروژه «' + body.project + '» ارسال کرد.\n' +
      'دلیل: ' + (body.reason||'-') +
      (body.requestedDate ? ('\nزمان درخواستی: ' + body.requestedDate + ' ' + (body.requestedTime||'')) : '') +
      '\n\nبرای تایید یا رد، وارد پنل مدیریت شوید.'
    );
    return jsonOut({ok:true});
  }

  if(action === 'getMyRequests'){
    var rsh2 = getSheet('Requests', ['id','username','employee','employeeCode','project','type','reason','status','createdAt','createdTime','requestedDate','requestedTime']);
    var rows4 = readRows(rsh2).filter(function(r){ return r.username === body.username; })
      .map(function(r){ delete r._row; return r; })
      .sort(function(a,b){ return String(b.createdAt+b.createdTime).localeCompare(String(a.createdAt+a.createdTime)); });
    return jsonOut({ok:true, requests: rows4});
  }

  // ---------------- leave requests ----------------
  if(action === 'createLeaveRequest'){
    var eshL = getSheet('Employees', ['id','fullName','username','password','deviceId','allowedProjects','employeeCode','freeZone','fingerprintRequired','jobTitle','shiftType','shiftStart','shiftEnd','nationalId','jobGroup','photoUrl','shiftId','active','sessionVersion','noDeviceLimit']);
    var empL = readRows(eshL).filter(function(r){ return r.username === body.username; })[0];
    if(!empL) return jsonOut({ok:false, error:'unknown_employee'});
    if(empL.deviceId && String(empL.deviceId) !== String(body.deviceId||'')){
      return jsonOut({ok:false, error:'device_mismatch'});
    }
    if(body.leaveType === 'hourly'){
      var stMin = timeToMinutes(normalizeTimeInput(body.startTime));
      var enMin = timeToMinutes(normalizeTimeInput(body.endTime));
      var durMin = enMin - stMin;
      if(durMin <= 0) durMin += 24*60;
      if(durMin > 240){
        return jsonOut({ok:false, error:'leave_too_long'});
      }
    }
    var lrsh = getSheet('LeaveRequests', ['id','username','employee','employeeCode','startDate','endDate','reason','status','createdAt','createdTime','leaveType','startTime','endTime','requestType','leaveCategory','project']);
    var tL = serverNow();
    var lrid = Utilities.getUuid();
    appendRowByHeader(lrsh, {
      id:lrid, username:body.username, employee:empL.fullName, employeeCode: empL.employeeCode||'', startDate:body.startDate, endDate:body.endDate,
      reason: body.reason||'', status:'pending', createdAt:tL.jalaliDate, createdTime:tL.time,
      leaveType: body.leaveType || 'daily', startTime: normalizeTimeInput(body.startTime), endTime: normalizeTimeInput(body.endTime),
      requestType: body.requestType || 'leave', leaveCategory: body.leaveCategory || '', project: body.project || ''
    });
    var when = (body.leaveType === 'hourly')
      ? (body.startDate + ' ساعت ' + (body.startTime||'') + ' تا ' + (body.endTime||''))
      : (body.startDate + ' تا ' + body.endDate);
    notifyAdmin(
      'درخواست مرخصی جدید — ' + empL.fullName,
      empL.fullName + ' درخواست مرخصی ' + (body.leaveType==='hourly'?'ساعتی':'روزانه') + ' ثبت کرد.\n' +
      'بازه: ' + when + '\nدلیل: ' + (body.reason||'-') +
      '\n\nبرای تایید یا رد، وارد پنل مدیریت شوید.'
    );
    return jsonOut({ok:true});
  }
  if(action === 'getMyLeaveRequests'){
    var lrsh2 = getSheet('LeaveRequests', ['id','username','employee','employeeCode','startDate','endDate','reason','status','createdAt','createdTime','leaveType','startTime','endTime','requestType','leaveCategory','project']);
    var lrows = readRows(lrsh2).filter(function(r){ return r.username === body.username; })
      .map(function(r){ delete r._row; return r; })
      .sort(function(a,b){ return String(b.createdAt+b.createdTime).localeCompare(String(a.createdAt+a.createdTime)); });
    return jsonOut({ok:true, leaves: lrows});
  }
  if(action === 'submitOfflineLog'){
    var d2 = body.entry;
    var eshO = getSheet('Employees', ['id','fullName','username','password','deviceId','allowedProjects','employeeCode','freeZone','fingerprintRequired','jobTitle','shiftType','shiftStart','shiftEnd','nationalId','jobGroup','photoUrl','shiftId','active','sessionVersion','noDeviceLimit']);
    var empO = readRows(eshO).filter(function(r){ return r.username === d2.username; })[0];
    if(!empO) return jsonOut({ok:false, error:'unknown_employee'});
    if(empO.deviceId && String(empO.deviceId) !== String(body.deviceId||'')){
      return jsonOut({ok:false, error:'device_mismatch'});
    }
    var posh = getSheet('PendingOfflineLogs', ['id','username','employee','employeeCode','project','type','jalaliDate','time','clientIso','lat','lng','accuracy','reason','status','createdAt']);
    // Block a second offline submission while an earlier one from this employee is still
    // awaiting admin review — otherwise a shaky connection can lead to several confusing
    // duplicate submissions piling up before the admin even sees the first one.
    var alreadyPending = readRows(posh).filter(function(r){ return r.username === d2.username && r.status === 'pending'; })[0];
    if(alreadyPending){
      return jsonOut({ok:false, error:'pending_offline_exists', at: alreadyPending.jalaliDate + ' ' + alreadyPending.time});
    }
    var clientDate = new Date(d2.clientIso);
    var pgy = clientDate.getFullYear(), pgm = clientDate.getMonth()+1, pgd = clientDate.getDate();
    var pj = toJalaali(pgy, pgm, pgd);
    var pjalaliDate = pj[2] + ' ' + J_MONTHS[pj[1]-1] + ' ' + pj[0];
    var ptime = pad2(clientDate.getHours())+':'+pad2(clientDate.getMinutes())+':'+pad2(clientDate.getSeconds());
    var poid = Utilities.getUuid();
    var t4 = serverNow();
    appendRowByHeader(posh, {
      id:poid, username:d2.username, employee:empO.fullName, employeeCode: empO.employeeCode||'',
      project:d2.project, type:d2.type, jalaliDate:pjalaliDate, time:ptime, clientIso:d2.clientIso,
      lat:d2.lat||'', lng:d2.lng||'', accuracy:d2.accuracy||'', reason:'', status:'pending', createdAt:t4.jalaliDate
    });
    notifyAdmin(
      'ثبت آفلاین جدید در انتظار بررسی — ' + empO.fullName,
      empO.fullName + ' یک ' + (d2.type==='in'?'ورود':'خروج') + ' آفلاین (بر اساس ساعت گوشی: ' + pjalaliDate + ' ' + ptime + ') ثبت کرد که نیاز به بررسی و تایید شما داره.\n\nبرای بررسی، وارد پنل مدیریت → «تردد ناقص» شوید.'
    );
    return jsonOut({ok:true});
  }

  // Employee-facing: "شیفت‌هایِ من" — their currently-effective shift (resolved exactly the
  // same way computeEmployeePerformance would for today), shown as a read-only weekly table.
  // Must stay ABOVE the blanket admin-PIN gate below, same reasoning as getMyPerformance next
  // to it — employees never send a PIN.
  if(action === 'getMyShiftInfo'){
    var eshMSI = getSheet('Employees', ['id','fullName','username','password','deviceId','allowedProjects','employeeCode','freeZone','fingerprintRequired','jobTitle','shiftType','shiftStart','shiftEnd','nationalId','jobGroup','photoUrl','shiftId','active','sessionVersion','noDeviceLimit']);
    var empMSI = readRows(eshMSI).filter(function(r){ return r.username === String(body.username||'').trim(); })[0];
    if(!empMSI) return jsonOut({ok:false, error:'unknown_employee'});
    var nowJ = serverNow();
    var yearMSI = parseInt(body.year, 10) || parseJalaliDateStrServer(nowJ.jalaliDate).y;
    var monthMSI = parseInt(body.month, 10) || parseJalaliDateStrServer(nowJ.jalaliDate).m;
    var daysInMonthMSI = jalaliMonthLengthServer(yearMSI, monthMSI);

    var assignSheetMSI = getSheet('ShiftAssignments', ['id','username','startDate','endDate','shiftStart','shiftEnd','shiftName','shiftId','createdAt']);
    var assignmentsMSI = readRows(assignSheetMSI).filter(function(r){ return r.username === empMSI.username; });
    var shiftsByIdMSI = buildShiftsByIdMap();
    var shiftDaysMapMSI = buildShiftDaysMap();

    var daysOutMSI = [];
    var overallShiftNameMSI = null; // for the header hint text only
    for(var dMSI=1; dMSI<=daysInMonthMSI; dMSI++){
      var dateStrMSI = dMSI + ' ' + J_MONTHS[monthMSI-1] + ' ' + yearMSI;
      var effShiftMSI = resolveEffectiveShiftCached(dateStrMSI, empMSI, assignmentsMSI, shiftsByIdMSI);
      if(effShiftMSI){
        overallShiftNameMSI = effShiftMSI.name || 'شیفتِ اختصاصی';
        // Attached directly on THIS day's own object, not a single month-wide cache — a
        // "ساعتی" (شیفتِ عادی) day elsewhere in the same month must never show these, and a
        // month-wide "last shift seen" cache was exactly the bug that let it happen.
        var shiftDetailsMSI = {
          shiftName: overallShiftNameMSI,
          floatingMinutes: effShiftMSI.floatingMinutes||'', dailyOvertimeCapMinutes: effShiftMSI.dailyOvertimeCapMinutes||'',
          otStartMinutes: effShiftMSI.otStartMinutes||'', otMiddleMinutes: effShiftMSI.otMiddleMinutes||'', otEndMinutes: effShiftMSI.otEndMinutes||'',
          lateAllowedMinutes: effShiftMSI.lateAllowedMinutes||'', earlyLeaveAllowedMinutes: effShiftMSI.earlyLeaveAllowedMinutes||'',
          floatBeforeShiftMinutes: effShiftMSI.floatBeforeShiftMinutes||'',
          attendanceWindowStart: effShiftMSI.attendanceWindowStart||'', attendanceWindowEnd: effShiftMSI.attendanceWindowEnd||''
        };
        var sdEntryMSI = shiftDaysMapMSI[effShiftMSI.id + '|' + dateStrMSI];
        if(sdEntryMSI){
          daysOutMSI.push({day: dMSI, jalaliDate: dateStrMSI, hasShift: true, start1: sdEntryMSI.start1, end1: sdEntryMSI.end1, hasSecondPart: String(sdEntryMSI.hasSecondPart||'').toLowerCase()==='yes', start2: sdEntryMSI.start2, end2: sdEntryMSI.end2, shiftDetails: shiftDetailsMSI});
        } else {
          daysOutMSI.push({day: dMSI, jalaliDate: dateStrMSI, hasShift: false, shiftDetails: shiftDetailsMSI});
        }
      } else {
        // شیفتِ عادی for this specific day — no fixed check-in/out window is actually
        // enforced under the old/default rules, just a required daily total, so this is
        // flagged as flexible ("ساعتی") rather than claiming a specific window like 07:30–14:50
        // that was never really true. No shiftDetails at all here — no custom shift applies
        // to this day, so none of those settings are relevant to it.
        var isFridayMSI = isFridayFromJalaliDateStr(dateStrMSI);
        daysOutMSI.push(isFridayMSI
          ? {day: dMSI, jalaliDate: dateStrMSI, hasShift: false, isNormalOff: true}
          : {day: dMSI, jalaliDate: dateStrMSI, hasShift: true, isFlexibleHours: true});
      }
    }
    return jsonOut({ok:true, shiftName: overallShiftNameMSI || 'شیفتِ عادی', hasCustomShift: !!overallShiftNameMSI, days: daysOutMSI});
  }

  // Employee-facing: their own monthly totals + full daily breakdown in one call, authenticated
  // by username only (matching getMyLogs above) — must stay ABOVE the blanket admin-PIN gate
  // right below, or every employee call to it fails with "invalid pin" despite never sending
  // one (exactly the bug this comment is here to prevent happening again).
  if(action === 'getMyPerformance'){
    var yearMYP = parseInt(body.year, 10);
    var monthMYP = parseInt(body.month, 10);
    if(!yearMYP || !monthMYP) return jsonOut({ok:false, error:'year and month are required'});
    var eshMYP = getSheet('Employees', ['id','fullName','username','password','deviceId','allowedProjects','employeeCode','freeZone','fingerprintRequired','jobTitle','shiftType','shiftStart','shiftEnd','nationalId','jobGroup','photoUrl','shiftId','active','sessionVersion','noDeviceLimit']);
    var empMYP = readRows(eshMYP).filter(function(r){ return r.username === String(body.username||'').trim(); })[0];
    if(!empMYP) return jsonOut({ok:false, error:'unknown_employee'});

    var daysInMonthMYP = jalaliMonthLengthServer(yearMYP, monthMYP);
    var firstDayGMYP = jalaliToGregorianServer(yearMYP, monthMYP, 1);
    var startBoundMYP = new Date(firstDayGMYP[0], firstDayGMYP[1]-1, firstDayGMYP[2]);
    startBoundMYP.setDate(startBoundMYP.getDate() - 2);
    var lastDayGMYP = jalaliToGregorianServer(yearMYP, monthMYP, daysInMonthMYP);
    var endBoundMYP = new Date(lastDayGMYP[0], lastDayGMYP[1]-1, lastDayGMYP[2]);
    endBoundMYP.setDate(endBoundMYP.getDate() + 2);

    var lshMYP = getSheet('Logs', ['id','jalaliDate','time','employee','username','project','type','lat','lng','accuracy','isoTimestamp','note','employeeCode','isHoliday','holidayReason','source']);
    var empLogsMYP = readRows(lshMYP).filter(function(r){
      var tms = new Date(r.isoTimestamp).getTime();
      return r.username === empMYP.username && !isNaN(tms) && tms >= startBoundMYP.getTime() && tms <= endBoundMYP.getTime();
    }).sort(function(a,b){ return isoCompare(a.isoTimestamp, b.isoTimestamp); });

    var dsmshMYP = getSheet('DayStatusMarkers', ['id','username','employee','employeeCode','jalaliDate','status','leaveHours','createdAt','createdTime']);
    var empMarkersMYP = readRows(dsmshMYP).filter(function(m){ return m.username === empMYP.username; });

    var hshMYP = getSheet('Holidays', ['id','jalaliDate','reason','status']);
    var holidaySetMYP = {};
    readRows(hshMYP).forEach(function(h){ holidaySetMYP[h.jalaliDate] = true; });

    var shiftsByIdMYP = buildShiftsByIdMap();
    var assignSheetMYP = getSheet('ShiftAssignments', ['id','username','startDate','endDate','shiftStart','shiftEnd','shiftName','shiftId','createdAt']);
    var assignmentsMYP = readRows(assignSheetMYP).filter(function(r){ return r.username === empMYP.username; });
    var shiftDaysMapMYP = buildShiftDaysMap();
    var normalSettingsMYP = getNormalShiftSettings();

    var calcMYP = computeEmployeePerformance(empMYP.username, yearMYP, monthMYP, empLogsMYP, empMarkersMYP, holidaySetMYP, empMYP, assignmentsMYP, shiftsByIdMYP, shiftDaysMapMYP, normalSettingsMYP);
    return jsonOut({ok:true, totals: calcMYP.totals, daily: calcMYP.daily});
  }

  // ==================== everything below requires the admin PIN ====================
  if(String(body.pin) !== getAdminPin()){
    return jsonOut({ok:false, error:'invalid pin'});
  }

  // ---------------- employees ----------------
  if(action === 'getEmployees'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var esh4 = getSheet('Employees', ['id','fullName','username','password','deviceId','allowedProjects','employeeCode','freeZone','fingerprintRequired','jobTitle','shiftType','shiftStart','shiftEnd','nationalId','jobGroup','photoUrl','shiftId','active','sessionVersion','noDeviceLimit']);
    var rows5 = readRows(esh4).map(function(r){ delete r._row; return r; });
    return jsonOut({ok:true, employees: rows5});
  }
  if(action === 'addEmployee'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    if(!body.employeeCode || !String(body.employeeCode).trim()){
      return jsonOut({ok:false, error:'employee_code_required'});
    }
    var esh5 = getSheet('Employees', ['id','fullName','username','password','deviceId','allowedProjects','employeeCode','freeZone','fingerprintRequired','jobTitle','shiftType','shiftStart','shiftEnd','nationalId','jobGroup','photoUrl','shiftId','active','sessionVersion','noDeviceLimit']);
    if(getMaxEmployees() > 0 && readRows(esh5).length >= getMaxEmployees()){
      return jsonOut({ok:false, error:'employee_limit_reached'});
    }
    var newUsernameA = String(body.username||'').trim().toLowerCase();
    var dupUserA = readRows(esh5).filter(function(r){ return String(r.username||'').trim().toLowerCase() === newUsernameA; })[0];
    if(dupUserA) return jsonOut({ok:false, error:'username_taken'});
    var id2 = Utilities.getUuid();
    appendRowByHeader(esh5, {
      id:id2, fullName:body.fullName, username:body.username, password:body.password,
      deviceId:'', allowedProjects: body.allowedProjects || '', employeeCode: body.employeeCode || '',
      freeZone: body.freeZone ? 'yes' : '', fingerprintRequired: body.fingerprintRequired ? 'yes' : '',
      jobTitle: body.jobTitle||'', shiftType: body.shiftType||'', shiftStart: body.shiftStart||'', shiftEnd: body.shiftEnd||'', shiftId: body.shiftId||'',
      nationalId: body.nationalId||'', jobGroup: body.jobGroup||'', photoUrl: body.photoUrl||'',
      noDeviceLimit: body.noDeviceLimit ? 'yes' : ''
    });
    return jsonOut({ok:true, id:id2});
  }
  if(action === 'updateEmployee'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    if(!body.employeeCode || !String(body.employeeCode).trim()){
      return jsonOut({ok:false, error:'employee_code_required'});
    }
    var esh6 = getSheet('Employees', ['id','fullName','username','password','deviceId','allowedProjects','employeeCode','freeZone','fingerprintRequired','jobTitle','shiftType','shiftStart','shiftEnd','nationalId','jobGroup','photoUrl','shiftId','active','sessionVersion','noDeviceLimit']);
    var row = readRows(esh6).filter(function(r){ return r.id === body.id; })[0];
    if(!row) return jsonOut({ok:false, error:'not found'});
    var newUsernameU = String(body.username||'').trim().toLowerCase();
    var dupUserU = readRows(esh6).filter(function(r){ return r.id !== body.id && String(r.username||'').trim().toLowerCase() === newUsernameU; })[0];
    if(dupUserU) return jsonOut({ok:false, error:'username_taken'});
    esh6.getRange(row._row, colIndex(esh6,'fullName')).setValue(body.fullName);
    esh6.getRange(row._row, colIndex(esh6,'username')).setValue(body.username);
    esh6.getRange(row._row, colIndex(esh6,'password')).setValue(body.password);
    esh6.getRange(row._row, colIndex(esh6,'allowedProjects')).setValue(body.allowedProjects || '');
    esh6.getRange(row._row, colIndex(esh6,'employeeCode')).setValue(body.employeeCode || '');
    esh6.getRange(row._row, colIndex(esh6,'freeZone')).setValue(body.freeZone ? 'yes' : '');
    esh6.getRange(row._row, colIndex(esh6,'fingerprintRequired')).setValue(body.fingerprintRequired ? 'yes' : '');
    esh6.getRange(row._row, colIndex(esh6,'jobTitle')).setValue(body.jobTitle||'');
    esh6.getRange(row._row, colIndex(esh6,'shiftType')).setValue(body.shiftType||'');
    esh6.getRange(row._row, colIndex(esh6,'shiftStart')).setValue(body.shiftStart||'');
    esh6.getRange(row._row, colIndex(esh6,'shiftEnd')).setValue(body.shiftEnd||'');
    // Only touched if the caller explicitly sent a shiftId — the main "save employee" form
    // never does (it has no shift field of its own), so without this check every ordinary
    // edit (renaming someone, changing their password, etc.) would silently wipe out whatever
    // شیفتِ اختصاصی had been assigned via the separate shift-assignment section.
    if(body.shiftId !== undefined) esh6.getRange(row._row, colIndex(esh6,'shiftId')).setValue(body.shiftId||'');
    esh6.getRange(row._row, colIndex(esh6,'nationalId')).setValue(body.nationalId||'');
    esh6.getRange(row._row, colIndex(esh6,'jobGroup')).setValue(body.jobGroup||'');
    esh6.getRange(row._row, colIndex(esh6,'noDeviceLimit')).setValue(body.noDeviceLimit ? 'yes' : '');
    if(body.photoUrl) esh6.getRange(row._row, colIndex(esh6,'photoUrl')).setValue(body.photoUrl);
    // Bumping this forces any already-logged-in phone for this employee back to the login
    // screen soon (see checkSessionVersion) — an edit here (new password, username, allowed
    // projects, etc.) shouldn't silently leave a stale session active on their device.
    esh_bumpSessionVersion(esh6, row._row);
    return jsonOut({ok:true});
  }
  if(action === 'deleteEmployee'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var esh7 = getSheet('Employees', ['id','fullName','username','password','deviceId','allowedProjects','employeeCode','freeZone','fingerprintRequired','jobTitle','shiftType','shiftStart','shiftEnd','nationalId','jobGroup','photoUrl','shiftId','active','sessionVersion','noDeviceLimit']);
    var row2 = readRows(esh7).filter(function(r){ return r.id === body.id; })[0];
    if(row2) esh7.deleteRow(row2._row);
    return jsonOut({ok:true});
  }
  if(action === 'setEmployeeActive'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var eshAct = getSheet('Employees', ['id','fullName','username','password','deviceId','allowedProjects','employeeCode','freeZone','fingerprintRequired','jobTitle','shiftType','shiftStart','shiftEnd','nationalId','jobGroup','photoUrl','shiftId','active','sessionVersion','noDeviceLimit']);
    var rowAct = readRows(eshAct).filter(function(r){ return r.id === body.id; })[0];
    if(!rowAct) return jsonOut({ok:false, error:'not found'});
    eshAct.getRange(rowAct._row, colIndex(eshAct,'active')).setValue(body.active ? 'true' : 'false');
    esh_bumpSessionVersion(eshAct, rowAct._row);
    return jsonOut({ok:true});
  }
  if(action === 'resetDevice'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var esh8 = getSheet('Employees', ['id','fullName','username','password','deviceId','allowedProjects','employeeCode','freeZone','fingerprintRequired','jobTitle','shiftType','shiftStart','shiftEnd','nationalId','jobGroup','photoUrl','shiftId','active','sessionVersion','noDeviceLimit']);
    var row3 = readRows(esh8).filter(function(r){ return r.id === body.id; })[0];
    if(!row3) return jsonOut({ok:false, error:'not found'});
    esh8.getRange(row3._row, colIndex(esh8,'deviceId')).setValue('');
    esh_bumpSessionVersion(esh8, row3._row);
    return jsonOut({ok:true});
  }

  // ---------------- projects ----------------
  if(action === 'getProjectsAdmin'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var psh = getSheet('Projects', ['id','name','lat','lng','radius','city','province','restricted','active']);
    return jsonOut({ok:true, projects: readRows(psh).map(function(r){ delete r._row; return r; })});
  }
  if(action === 'addProject'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var psh2 = getSheet('Projects', ['id','name','lat','lng','radius','city','province','restricted','active']);
    var id3 = Utilities.getUuid();
    appendRowByHeader(psh2, {id:id3, name:body.name, lat:body.lat, lng:body.lng, radius:body.radius, city: body.city||'', province: body.province||'', restricted: body.restricted ? 'yes' : 'no', active: 'yes'});
    return jsonOut({ok:true, id:id3});
  }
// Bulk find-and-replace a plain string value in one column across a sheet — used to keep
// project names in sync everywhere after a rename, since Logs/LeaveRequests/etc. store the
// project as a plain text string rather than referencing the Projects row by id.
function replaceColumnValueInSheet(sheetName, headers, columnName, oldValue, newValue){
  if(oldValue === newValue) return 0;
  var sh = getSheet(sheetName, headers);
  var lastRow = sh.getLastRow();
  if(lastRow < 2) return 0;
  var lastCol = sh.getLastColumn();
  var colIdx = colIndex(sh, columnName);
  if(!colIdx) return 0;
  var range = sh.getRange(2, 1, lastRow-1, lastCol);
  var values = range.getValues();
  var changedCount = 0;
  for(var i=0; i<values.length; i++){
    if(String(values[i][colIdx-1]||'') === oldValue){
      values[i][colIdx-1] = newValue;
      changedCount++;
    }
  }
  if(changedCount > 0) range.setValues(values);
  return changedCount;
}
  if(action === 'updateProject'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var psh3 = getSheet('Projects', ['id','name','lat','lng','radius','city','province','restricted','active']);
    var row4 = readRows(psh3).filter(function(r){ return r.id === body.id; })[0];
    if(!row4) return jsonOut({ok:false, error:'not found'});
    var oldProjectName = row4.name;
    esh_set(psh3, row4._row, 'name', body.name);
    esh_set(psh3, row4._row, 'lat', body.lat);
    esh_set(psh3, row4._row, 'lng', body.lng);
    esh_set(psh3, row4._row, 'radius', body.radius);
    esh_set(psh3, row4._row, 'city', body.city||'');
    esh_set(psh3, row4._row, 'province', body.province||'');
    esh_set(psh3, row4._row, 'restricted', body.restricted ? 'yes' : 'no');
    // A renamed project is stored as plain text in every other sheet that references it (no
    // stable id link) — this is the one moment we reliably know both the old and new name, so
    // propagate the rename everywhere right now rather than leaving old rows stuck showing a
    // name that no longer exists anywhere in the Projects list.
    if(oldProjectName && body.name && oldProjectName !== body.name){
      replaceColumnValueInSheet('Logs', ['id','jalaliDate','time','employee','username','project','type','lat','lng','accuracy','isoTimestamp','note','employeeCode','isHoliday','holidayReason','source'], 'project', oldProjectName, body.name);
      replaceColumnValueInSheet('LeaveRequests', ['id','username','employee','employeeCode','startDate','endDate','reason','status','createdAt','createdTime','leaveType','startTime','endTime','requestType','leaveCategory','project'], 'project', oldProjectName, body.name);
      replaceColumnValueInSheet('Requests', ['id','username','employee','employeeCode','project','type','reason','status','createdAt','createdTime','requestedDate','requestedTime'], 'project', oldProjectName, body.name);
      replaceColumnValueInSheet('PendingOfflineLogs', ['id','username','employee','employeeCode','project','type','jalaliDate','time','clientIso','lat','lng','accuracy','reason','status','createdAt'], 'project', oldProjectName, body.name);
    }
    return jsonOut({ok:true});
  }
  if(action === 'deleteProject'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var psh4 = getSheet('Projects', ['id','name','lat','lng','radius','city','province','restricted','active']);
    var row5 = readRows(psh4).filter(function(r){ return r.id === body.id; })[0];
    if(row5) psh4.deleteRow(row5._row);
    return jsonOut({ok:true});
  }
  if(action === 'setProjectActive'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var pshSA = getSheet('Projects', ['id','name','lat','lng','radius','city','province','restricted','active']);
    var rowSA = readRows(pshSA).filter(function(r){ return r.id === body.id; })[0];
    if(!rowSA) return jsonOut({ok:false, error:'not found'});
    esh_set(pshSA, rowSA._row, 'active', body.active ? 'yes' : 'no');
    return jsonOut({ok:true});
  }

  // ---------------- shifts ----------------
  var SHIFTS_SCHEMA = ['id','name','type','shiftStart','shiftEnd','requiredMinutes','lateTolerance','earlyTolerance','shiftYear','floatingMinutes','dailyOvertimeCapMinutes','weeklyScheduleJson','createdAt','otStartMinutes','otMiddleMinutes','otEndMinutes','lateAllowedMinutes','earlyLeaveAllowedMinutes','floatBeforeShiftMinutes','attendanceWindowStart','attendanceWindowEnd','followOfficialHolidays'];
  if(action === 'getNormalShiftSettings'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    return jsonOut({ok:true, settings: getNormalShiftSettings()});
  }
  if(action === 'saveNormalShiftSettings'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var reqMinSN = parseInt(body.requiredMinutes, 10);
    var breakMinSN = parseInt(body.breakMinutes, 10);
    if(!reqMinSN || reqMinSN <= 0) return jsonOut({ok:false, error:'invalid_required_minutes'});
    if(!breakMinSN || breakMinSN < 0) return jsonOut({ok:false, error:'invalid_break_minutes'});
    if(!body.nightStart || !body.nightEnd) return jsonOut({ok:false, error:'night_times_required'});
    setSetting('normal_required_minutes', String(reqMinSN));
    setSetting('normal_break_minutes', String(breakMinSN));
    setSetting('normal_night_start', body.nightStart);
    setSetting('normal_night_end', body.nightEnd);
    return jsonOut({ok:true});
  }
  if(action === 'getShifts'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var shSh = getSheet('Shifts', SHIFTS_SCHEMA);
    var shRows = readRows(shSh).map(function(r){ delete r._row; return r; });
    return jsonOut({ok:true, shifts: shRows});
  }
  if(action === 'addShift'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    if(!body.name || !String(body.name).trim()) return jsonOut({ok:false, error:'name_required'});
    var shSh2 = getSheet('Shifts', SHIFTS_SCHEMA);
    var idSh = Utilities.getUuid();
    appendRowByHeader(shSh2, {
      id: idSh, name: body.name, type: body.type || 'simple',
      shiftYear: body.shiftYear || '', floatingMinutes: body.floatingMinutes || '',
      dailyOvertimeCapMinutes: body.dailyOvertimeCapMinutes || '',
      weeklyScheduleJson: body.weeklyScheduleJson || '[]',
      createdAt: new Date().toISOString(),
      otStartMinutes: body.otStartMinutes || '', otMiddleMinutes: body.otMiddleMinutes || '', otEndMinutes: body.otEndMinutes || '',
      lateAllowedMinutes: body.lateAllowedMinutes || '', earlyLeaveAllowedMinutes: body.earlyLeaveAllowedMinutes || '',
      floatBeforeShiftMinutes: body.floatBeforeShiftMinutes || '',
      attendanceWindowStart: body.attendanceWindowStart || '', attendanceWindowEnd: body.attendanceWindowEnd || '',
      followOfficialHolidays: body.followOfficialHolidays ? 'yes' : 'no'
    });
    return jsonOut({ok:true, id:idSh});
  }
  if(action === 'updateShift'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    if(!body.name || !String(body.name).trim()) return jsonOut({ok:false, error:'name_required'});
    var shSh3 = getSheet('Shifts', SHIFTS_SCHEMA);
    var shRow = readRows(shSh3).filter(function(r){ return r.id === body.id; })[0];
    if(!shRow) return jsonOut({ok:false, error:'not found'});
    esh_set(shSh3, shRow._row, 'name', body.name);
    esh_set(shSh3, shRow._row, 'type', body.type || 'simple');
    esh_set(shSh3, shRow._row, 'shiftYear', body.shiftYear || '');
    esh_set(shSh3, shRow._row, 'floatingMinutes', body.floatingMinutes || '');
    esh_set(shSh3, shRow._row, 'dailyOvertimeCapMinutes', body.dailyOvertimeCapMinutes || '');
    esh_set(shSh3, shRow._row, 'weeklyScheduleJson', body.weeklyScheduleJson || '[]');
    esh_set(shSh3, shRow._row, 'otStartMinutes', body.otStartMinutes || '');
    esh_set(shSh3, shRow._row, 'otMiddleMinutes', body.otMiddleMinutes || '');
    esh_set(shSh3, shRow._row, 'otEndMinutes', body.otEndMinutes || '');
    esh_set(shSh3, shRow._row, 'lateAllowedMinutes', body.lateAllowedMinutes || '');
    esh_set(shSh3, shRow._row, 'earlyLeaveAllowedMinutes', body.earlyLeaveAllowedMinutes || '');
    esh_set(shSh3, shRow._row, 'floatBeforeShiftMinutes', body.floatBeforeShiftMinutes || '');
    esh_set(shSh3, shRow._row, 'attendanceWindowStart', body.attendanceWindowStart || '');
    esh_set(shSh3, shRow._row, 'attendanceWindowEnd', body.attendanceWindowEnd || '');
    esh_set(shSh3, shRow._row, 'followOfficialHolidays', body.followOfficialHolidays ? 'yes' : 'no');
    return jsonOut({ok:true});
  }
  if(action === 'deleteShift'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var shSh4 = getSheet('Shifts', SHIFTS_SCHEMA);
    var shRowD = readRows(shSh4).filter(function(r){ return r.id === body.id; })[0];
    if(!shRowD) return jsonOut({ok:false, error:'not found'});
    // Guard against deleting a shift that's still assigned to someone — checks BOTH the old
    // flat shiftId field on Employees AND the ShiftAssignments history table, since an
    // employee's actual current shift may now come from either.
    var eshCheckD = getSheet('Employees', ['id','fullName','username','password','deviceId','allowedProjects','employeeCode','freeZone','fingerprintRequired','jobTitle','shiftType','shiftStart','shiftEnd','nationalId','jobGroup','photoUrl','shiftId','active','sessionVersion','noDeviceLimit']);
    var inUseCount = readRows(eshCheckD).filter(function(r){ return r.shiftId === body.id; }).length;
    var assignSh4 = getSheet('ShiftAssignments', ['id','username','startDate','endDate','shiftStart','shiftEnd','shiftName','shiftId','createdAt']);
    inUseCount += readRows(assignSh4).filter(function(r){ return r.shiftId === body.id; }).length;
    if(inUseCount > 0) return jsonOut({ok:false, error:'shift_in_use', count: inUseCount});
    shSh4.deleteRow(shRowD._row);
    // Clean up this shift's own per-day calendar entries too — otherwise they'd sit around
    // forever as orphaned rows referencing a shiftId that no longer exists anywhere.
    var sdSchemaDel = ['id','shiftId','jalaliDate','start1','end1','hasSecondPart','start2','end2','createdAt'];
    var sdShDel = getSheet('ShiftDays', sdSchemaDel);
    var sdRowsDel = readRows(sdShDel).filter(function(r){ return r.shiftId === body.id; });
    // delete from the bottom up so earlier row numbers don't shift under us mid-loop
    sdRowsDel.sort(function(a,b){ return b._row - a._row; }).forEach(function(r){ sdShDel.deleteRow(r._row); });
    return jsonOut({ok:true});
  }

  // ---------------- shift calendar days (per-specific-date shift hours) ----------------
  // Unlike the earlier weeklyScheduleJson (a repeating شنبه..جمعه template), this is the model
  // actually requested: every individual calendar date can have its own independent start/end
  // times, entered one at a time (or copied from another already-defined day onto several
  // selected days at once, or copied wholesale to the same dates next year). A date with no
  // row here is simply "بدونِ شیفت" for that shift — not an error, just unfilled-in so far.
  // dayType 'simple' uses start1/end1(/start2/end2) like before. dayType 'floating' has no
  // fixed check-in/out window at all — instead dailyRequiredMinutes sets how many minutes must
  // be worked that day, whenever the employee chooses to work them (checked at addLog time is
  // out of scope for this pass; computeEmployeePerformance uses it as this day's required
  // total). Every day* field is a PER-DAY override of the shift-level setting of the same
  // name — left blank, the shift-level value is used instead; set, it takes precedence for
  // this specific calendar date only. Night-shift fields are stored for display/reference for
  // now (an admin might want to record a shift's official night-shift window) but are not yet
  // wired into a separate calculation pathway — dayNightMin's own after-22:00 threshold
  // already covers actual worked night minutes regardless of shift type.
  var SHIFT_DAYS_SCHEMA = ['id','shiftId','jalaliDate','dayType','start1','end1','hasSecondPart','start2','end2',
    'dailyRequiredMinutes','dayFloatingMinutes','dayOvertimeCapMinutes','dayOtStartMinutes','dayOtMiddleMinutes','dayOtEndMinutes',
    'dayLateAllowedMinutes','dayEarlyAllowedMinutes','dayFloatBeforeShiftMinutes','dayAttendanceWindowStart','dayAttendanceWindowEnd',
    'isNightShift','nightShiftCrossesNextDay','nightShiftStart','nightShiftDuration','createdAt'];
  if(action === 'getShiftDays'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    if(!body.shiftId) return jsonOut({ok:false, error:'shift_id_required'});
    var sdShG = getSheet('ShiftDays', SHIFT_DAYS_SCHEMA);
    var sdRowsG = readRows(sdShG).filter(function(r){ return r.shiftId === body.shiftId; }).map(function(r){ delete r._row; return r; });
    return jsonOut({ok:true, days: sdRowsG});
  }
  if(action === 'setShiftDay'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    if(!body.shiftId || !body.jalaliDate) return jsonOut({ok:false, error:'shift_id_and_date_required'});
    var sdShS = getSheet('ShiftDays', SHIFT_DAYS_SCHEMA);
    var dayFieldsS = {
      dayType: body.dayType || 'simple',
      start1: body.start1||'', end1: body.end1||'', hasSecondPart: body.hasSecondPart ? 'yes' : '',
      start2: body.start2||'', end2: body.end2||'',
      dailyRequiredMinutes: body.dailyRequiredMinutes||'',
      dayFloatingMinutes: body.dayFloatingMinutes||'', dayOvertimeCapMinutes: body.dayOvertimeCapMinutes||'',
      dayOtStartMinutes: body.dayOtStartMinutes||'', dayOtMiddleMinutes: body.dayOtMiddleMinutes||'', dayOtEndMinutes: body.dayOtEndMinutes||'',
      dayLateAllowedMinutes: body.dayLateAllowedMinutes||'', dayEarlyAllowedMinutes: body.dayEarlyAllowedMinutes||'',
      dayFloatBeforeShiftMinutes: body.dayFloatBeforeShiftMinutes||'',
      dayAttendanceWindowStart: body.dayAttendanceWindowStart||'', dayAttendanceWindowEnd: body.dayAttendanceWindowEnd||'',
      isNightShift: body.isNightShift ? 'yes' : '', nightShiftCrossesNextDay: body.nightShiftCrossesNextDay ? 'yes' : '',
      nightShiftStart: body.nightShiftStart||'', nightShiftDuration: body.nightShiftDuration||''
    };
    var existingS = readRows(sdShS).filter(function(r){ return r.shiftId === body.shiftId && r.jalaliDate === body.jalaliDate; })[0];
    if(existingS){
      Object.keys(dayFieldsS).forEach(function(k){ esh_set(sdShS, existingS._row, k, dayFieldsS[k]); });
      return jsonOut({ok:true, id: existingS.id});
    }
    var idS = Utilities.getUuid();
    appendRowByHeader(sdShS, Object.assign({id: idS, shiftId: body.shiftId, jalaliDate: body.jalaliDate, createdAt: new Date().toISOString()}, dayFieldsS));
    return jsonOut({ok:true, id:idS});
  }
  if(action === 'deleteShiftDay'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var sdShD = getSheet('ShiftDays', SHIFT_DAYS_SCHEMA);
    var rowD2 = readRows(sdShD).filter(function(r){ return r.shiftId === body.shiftId && r.jalaliDate === body.jalaliDate; })[0];
    if(rowD2) sdShD.deleteRow(rowD2._row);
    return jsonOut({ok:true});
  }
  if(action === 'bulkDeleteShiftDays'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    if(!body.dates || !body.dates.length) return jsonOut({ok:true, deleted:0});
    var sdShBD = getSheet('ShiftDays', SHIFT_DAYS_SCHEMA);
    var dateSetBD = {}; body.dates.forEach(function(d){ dateSetBD[d]=true; });
    var toDeleteBD = readRows(sdShBD).filter(function(r){ return r.shiftId === body.shiftId && dateSetBD[r.jalaliDate]; });
    toDeleteBD.sort(function(a,b){ return b._row - a._row; }).forEach(function(r){ sdShBD.deleteRow(r._row); });
    return jsonOut({ok:true, deleted: toDeleteBD.length});
  }
  if(action === 'copyShiftDayToMultiple'){
    // Copies one already-defined day's hours onto several selected target dates — the
    // "کپی" workflow: admin defines one day manually, ticks the checkboxes on the days they
    // want the same hours applied to, and this fills all of them in one call.
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    if(!body.shiftId || !body.sourceDate || !body.targetDates || !body.targetDates.length) return jsonOut({ok:false, error:'missing_params'});
    var sdShC = getSheet('ShiftDays', SHIFT_DAYS_SCHEMA);
    var allRowsC = readRows(sdShC);
    var sourceRowC = allRowsC.filter(function(r){ return r.shiftId === body.shiftId && r.jalaliDate === body.sourceDate; })[0];
    if(!sourceRowC) return jsonOut({ok:false, error:'source_day_not_defined'});
    var rowsByDateC = {};
    allRowsC.forEach(function(r){ if(r.shiftId === body.shiftId) rowsByDateC[r.jalaliDate] = r; });
    body.targetDates.forEach(function(targetDate){
      if(targetDate === body.sourceDate) return; // copying onto itself is a no-op
      var existingC = rowsByDateC[targetDate];
      if(existingC){
        esh_set(sdShC, existingC._row, 'start1', sourceRowC.start1);
        esh_set(sdShC, existingC._row, 'end1', sourceRowC.end1);
        esh_set(sdShC, existingC._row, 'hasSecondPart', sourceRowC.hasSecondPart);
        esh_set(sdShC, existingC._row, 'start2', sourceRowC.start2);
        esh_set(sdShC, existingC._row, 'end2', sourceRowC.end2);
      } else {
        appendRowByHeader(sdShC, {
          id: Utilities.getUuid(), shiftId: body.shiftId, jalaliDate: targetDate,
          start1: sourceRowC.start1, end1: sourceRowC.end1, hasSecondPart: sourceRowC.hasSecondPart,
          start2: sourceRowC.start2, end2: sourceRowC.end2, createdAt: new Date().toISOString()
        });
      }
    });
    return jsonOut({ok:true, copied: body.targetDates.length});
  }
  if(action === 'copyShiftToNextYear'){
    // Copies every day this shift has defined onto the same month/day one Jalali year later —
    // "کپیِ شیفت در سالِ بعد". Purely additive/overwriting on the target dates; days the shift
    // doesn't have defined this year are simply not touched next year either.
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    if(!body.shiftId) return jsonOut({ok:false, error:'shift_id_required'});
    var sdShY = getSheet('ShiftDays', SHIFT_DAYS_SCHEMA);
    var allRowsY = readRows(sdShY);
    var thisShiftRowsY = allRowsY.filter(function(r){ return r.shiftId === body.shiftId; });
    var rowsByDateY = {};
    allRowsY.forEach(function(r){ if(r.shiftId === body.shiftId) rowsByDateY[r.jalaliDate] = r; });
    var copiedCountY = 0;
    thisShiftRowsY.forEach(function(r){
      var pj = parseJalaliDateStrServer(r.jalaliDate);
      if(!pj) return;
      var nextYearDate = pj.d + ' ' + J_MONTHS[pj.m-1] + ' ' + (pj.y+1);
      var existingY = rowsByDateY[nextYearDate];
      if(existingY){
        esh_set(sdShY, existingY._row, 'start1', r.start1);
        esh_set(sdShY, existingY._row, 'end1', r.end1);
        esh_set(sdShY, existingY._row, 'hasSecondPart', r.hasSecondPart);
        esh_set(sdShY, existingY._row, 'start2', r.start2);
        esh_set(sdShY, existingY._row, 'end2', r.end2);
      } else {
        appendRowByHeader(sdShY, {
          id: Utilities.getUuid(), shiftId: body.shiftId, jalaliDate: nextYearDate,
          start1: r.start1, end1: r.end1, hasSecondPart: r.hasSecondPart,
          start2: r.start2, end2: r.end2, createdAt: new Date().toISOString()
        });
      }
      copiedCountY++;
    });
    return jsonOut({ok:true, copied: copiedCountY});
  }
  if(action === 'bulkSaveShiftDays'){
    // Replaces the previous "generateShiftCalendarFromWeekly" workflow: the admin now stages
    // the ENTIRE calendar client-side (whether generated from a weekly pattern for شیفتِ ساده,
    // or built up day-by-day for شیفتِ سفارشی) and nothing touches the backend until this one
    // final "ثبت" — at which point the whole staged set is written in one shot, replacing
    // whatever ShiftDays rows already existed for this shift. Deleting-then-reinserting (rather
    // than diffing) correctly handles the admin having deleted or edited days locally too.
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    if(!body.shiftId || !body.days) return jsonOut({ok:false, error:'missing_params'});
    var sdShBS = getSheet('ShiftDays', SHIFT_DAYS_SCHEMA);
    var existingBS = readRows(sdShBS).filter(function(r){ return r.shiftId === body.shiftId; });
    existingBS.sort(function(a,b){ return b._row - a._row; }).forEach(function(r){ sdShBS.deleteRow(r._row); });
    var nowIsoBS = new Date().toISOString();
    var newRowsBS = body.days.map(function(d){
      return [
        Utilities.getUuid(), body.shiftId, d.jalaliDate, d.dayType||'simple',
        d.start1||'', d.end1||'', d.hasSecondPart?'yes':'', d.start2||'', d.end2||'',
        d.dailyRequiredMinutes||'', d.dayFloatingMinutes||'', d.dayOvertimeCapMinutes||'',
        d.dayOtStartMinutes||'', d.dayOtMiddleMinutes||'', d.dayOtEndMinutes||'',
        d.dayLateAllowedMinutes||'', d.dayEarlyAllowedMinutes||'', d.dayFloatBeforeShiftMinutes||'',
        d.dayAttendanceWindowStart||'', d.dayAttendanceWindowEnd||'',
        d.isNightShift?'yes':'', d.nightShiftCrossesNextDay?'yes':'', d.nightShiftStart||'', d.nightShiftDuration||'',
        nowIsoBS
      ];
    });
    if(newRowsBS.length > 0){
      var startRowBS = sdShBS.getLastRow() + 1;
      sdShBS.getRange(startRowBS, 1, newRowsBS.length, SHIFT_DAYS_SCHEMA.length).setValues(newRowsBS);
    }
    return jsonOut({ok:true, saved: newRowsBS.length});
  }

  // ---------------- per-employee shift assignments (history, general or date-ranged) ----------------
  // Every employee implicitly starts on "شیفتِ عادی" (the normal/default calculation rules —
  // represented here as shiftId:'' ). An assignment record layers a specific shift on top of
  // that from startDate onward — either open-ended (endDate empty, "کلی") or bounded (endDate
  // set, "فقط برایِ این بازه"). getEffectiveShift (used by computeEmployeePerformance) resolves
  // which one wins for any given day: a date-ranged assignment covering that day beats the
  // employee's general shiftId, and once a ranged assignment's window ends the employee falls
  // back to whatever general shift (or "شیفتِ عادی") was already in effect — never to "no rules
  // at all". This is a full history, not a single overwritten field, so past days keep
  // calculating under whatever was actually in effect on that day even after the admin changes
  // the assignment later. Reuses the ShiftAssignments sheet that already existed (extended
  // here with a shiftId column) rather than introducing a second, competing table.
  var SHIFT_ASSIGNMENTS_SCHEMA = ['id','username','startDate','endDate','shiftStart','shiftEnd','shiftName','shiftId','createdAt'];
  if(action === 'getEmployeeShiftAssignments'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var asSh = getSheet('ShiftAssignments', SHIFT_ASSIGNMENTS_SCHEMA);
    var asRows = readRows(asSh).filter(function(r){ return r.username === body.username; }).map(function(r){ delete r._row; return r; });
    asRows.sort(function(a,b){ return jalaliDateSortKey(b.startDate) - jalaliDateSortKey(a.startDate); });
    return jsonOut({ok:true, assignments: asRows});
  }
  if(action === 'assignEmployeeShift'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    if(!body.username) return jsonOut({ok:false, error:'username_required'});
    if(!body.startDate) return jsonOut({ok:false, error:'start_date_required'});
    var asSh2 = getSheet('ShiftAssignments', SHIFT_ASSIGNMENTS_SCHEMA);
    var idAs = Utilities.getUuid();
    appendRowByHeader(asSh2, {
      id: idAs, username: body.username, shiftId: body.shiftId || '',
      startDate: body.startDate, endDate: body.endDate || '',
      shiftStart: '', shiftEnd: '', shiftName: '', // old flat-style fields — unused by new,
      // shiftId-based assignments, kept only so the schema stays compatible with any older row
      createdAt: new Date().toISOString()
    });
    return jsonOut({ok:true, id:idAs});
  }
  if(action === 'deleteEmployeeShiftAssignment'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var asSh3 = getSheet('ShiftAssignments', SHIFT_ASSIGNMENTS_SCHEMA);
    var asRowD = readRows(asSh3).filter(function(r){ return r.id === body.id; })[0];
    if(asRowD) asSh3.deleteRow(asRowD._row);
    return jsonOut({ok:true});
  }

  // ---------------- logs (admin edit/delete) ----------------
  if(action === 'getLogs'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var lsh4 = getSheet('Logs', ['id','jalaliDate','time','employee','username','project','type','lat','lng','accuracy','isoTimestamp','note','employeeCode','isHoliday','holidayReason','source']);
    var yearGL = parseInt(body.year, 10);
    var monthGL = parseInt(body.month, 10);
    if(yearGL && monthGL){
      // Scope to just the requested month (with a margin for night-shift spillover across
      // the boundary) instead of reading/returning the entire sheet every time — this is
      // the same fix that made the status report fast, applied here too since normal report
      // loads were just as slow once enough history had accumulated.
      var daysInMonthGL = jalaliMonthLengthServer(yearGL, monthGL);
      var firstDayGL = jalaliToGregorianServer(yearGL, monthGL, 1);
      var startBoundGL = new Date(firstDayGL[0], firstDayGL[1]-1, firstDayGL[2]);
      startBoundGL.setDate(startBoundGL.getDate() - 2);
      var lastDayGL = jalaliToGregorianServer(yearGL, monthGL, daysInMonthGL);
      var endBoundGL = new Date(lastDayGL[0], lastDayGL[1]-1, lastDayGL[2]);
      endBoundGL.setDate(endBoundGL.getDate() + 2);
      var startMsGL = startBoundGL.getTime();
      var endMsGL = endBoundGL.getTime();
      var rows6 = ensureIds(lsh4, readRows(lsh4).filter(function(r){
        var tms = new Date(r.isoTimestamp).getTime();
        return !isNaN(tms) && tms >= startMsGL && tms <= endMsGL;
      })).map(function(r){ delete r._row; return r; });
      return jsonOut({ok:true, logs: rows6});
    }
    var rows6 = ensureIds(lsh4, readRows(lsh4)).map(function(r){ delete r._row; return r; });
    return jsonOut({ok:true, logs: rows6});
  }
  if(action === 'getDashboardData'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var eshDD = getSheet('Employees', ['id','fullName','username','password','deviceId','allowedProjects','employeeCode','freeZone','fingerprintRequired','jobTitle','shiftType','shiftStart','shiftEnd','nationalId','jobGroup','photoUrl','shiftId','active','sessionVersion','noDeviceLimit']);
    var empsDD = readRows(eshDD).map(function(r){ delete r._row; return r; });

    var lshDD = getSheet('Logs', ['id','jalaliDate','time','employee','username','project','type','lat','lng','accuracy','isoTimestamp','note','employeeCode','isHoliday','holidayReason','source']);
    var daysBackDD = body.daysBack || 3;
    var cutoffDD = new Date();
    cutoffDD.setDate(cutoffDD.getDate() - daysBackDD);
    var cutoffMsDD = cutoffDD.getTime();
    var recentLogsDD = readRows(lshDD).filter(function(r){ return new Date(r.isoTimestamp).getTime() >= cutoffMsDD; }).map(function(r){ delete r._row; return r; });

    var lrshDD = getSheet('LeaveRequests', ['id','username','employee','employeeCode','startDate','endDate','reason','status','createdAt','createdTime','leaveType','startTime','endTime','requestType','leaveCategory','project']);
    var allLeavesDD = readRows(lrshDD).map(function(r){ delete r._row; return r; });

    var rshDD = getSheet('Requests', ['id','username','employee','employeeCode','project','type','reason','status','createdAt','createdTime','requestedDate','requestedTime']);
    var pendingReqDD = readRows(rshDD).filter(function(r){ return r.status === 'pending'; }).map(function(r){ delete r._row; return r; });

    var hshDD = getSheet('Holidays', ['id','jalaliDate','reason','status']);
    var holidaysDD = readRows(hshDD).map(function(r){ delete r._row; return r; });

    return jsonOut({ok:true, employees: empsDD, logs: recentLogsDD, leaves: allLeavesDD, requests: pendingReqDD, holidays: holidaysDD, today: serverNow().jalaliDate});
  }
  // Classifies every employee's status for every day of a given Jalali month: normal
  // attendance, absence, approved leave, approved mission, or holiday/off. Used by the
  // report page's status filter. Priority per day: leave/mission marker > holiday > has
  // attendance (normal) > absence. A night-shift check-in counts as covering its own
  // start day AND the day its matching check-out lands on, but only for the FIRST cycle
  // that started on a given day — a second same-day cycle spilling into the next day does
  // not also cover that next day.
  if(action === 'getAttendanceStatusReport'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var yearASR = parseInt(body.year, 10);
    var monthASR = parseInt(body.month, 10);
    if(!yearASR || !monthASR) return jsonOut({ok:false, error:'year and month are required'});
    var statusFilterASR = body.status || 'all';

    var eshASR = getSheet('Employees', ['id','fullName','username','password','deviceId','allowedProjects','employeeCode','freeZone','fingerprintRequired','jobTitle','shiftType','shiftStart','shiftEnd','nationalId','jobGroup','photoUrl','shiftId','active','sessionVersion','noDeviceLimit']);
    var employeesASR = readRows(eshASR).filter(function(e){ return String(e.active) !== 'false'; });

    var daysInMonthASR = jalaliMonthLengthServer(yearASR, monthASR);

    // Only read/process log rows anywhere near this month — reading and re-pairing an
    // employee's ENTIRE history on every request gets slower and slower as data accumulates,
    // and was leaving this report stuck on "loading" for large datasets. A 2-day margin on
    // each side is enough to correctly detect night-shift spillover across the month boundary.
    var firstDayG = jalaliToGregorianServer(yearASR, monthASR, 1);
    var startBoundASR = new Date(firstDayG[0], firstDayG[1]-1, firstDayG[2]);
    startBoundASR.setDate(startBoundASR.getDate() - 2);
    var lastDayG = jalaliToGregorianServer(yearASR, monthASR, daysInMonthASR);
    var endBoundASR = new Date(lastDayG[0], lastDayG[1]-1, lastDayG[2]);
    endBoundASR.setDate(endBoundASR.getDate() + 2);
    var startMsASR = startBoundASR.getTime();
    var endMsASR = endBoundASR.getTime();

    var lshASR = getSheet('Logs', ['id','jalaliDate','time','employee','username','project','type','lat','lng','accuracy','isoTimestamp','note','employeeCode','isHoliday','holidayReason','source']);
    var allLogsASR = readRows(lshASR).filter(function(r){
      var tms = new Date(r.isoTimestamp).getTime();
      return !isNaN(tms) && tms >= startMsASR && tms <= endMsASR;
    });

    var dsmshASR = getSheet('DayStatusMarkers', ['id','username','employee','employeeCode','jalaliDate','status','leaveHours','createdAt','createdTime']);
    var markersASR = readRows(dsmshASR);

    var resultsASR = [];
    var todayJ = parseJalaliDateStrServer(serverNow().jalaliDate); // never classify future days — nothing has happened yet

    employeesASR.forEach(function(emp){
      var empLogsASR = allLogsASR.filter(function(r){ return r.username === emp.username; }).sort(function(a,b){ return isoCompare(a.isoTimestamp, b.isoTimestamp); });
      var pairsASR = pairLogsIntoCycles(empLogsASR);
      var empMarkersASR = {};
      markersASR.filter(function(m){ return m.username === emp.username; }).forEach(function(m){ empMarkersASR[m.jalaliDate] = m.status; });

      for(var d=1; d<=daysInMonthASR; d++){
        var isFutureASR = (yearASR > todayJ.y) || (yearASR === todayJ.y && monthASR > todayJ.m) || (yearASR === todayJ.y && monthASR === todayJ.m && d > todayJ.d);
        if(isFutureASR) continue;
        var dateStrASR = d + ' ' + J_MONTHS[monthASR-1] + ' ' + yearASR;
        var statusASR;
        var markerValASR = empMarkersASR[dateStrASR];
        // Only trust the stored marker for leave/mission/holiday — no independently
        // re-derived holiday guess for unmarked days. That fallback used to call
        // getHolidayInfo() (a full sheet re-read) once per employee per day, which was both
        // the source of "way more Off days than actually exist" (it marked every Friday
        // going forward, even future ones, regardless of whether the daily job had actually
        // run) AND the main reason this report could take minutes to load.
        if(markerValASR === 'مرخصی'){
          statusASR = 'leave';
        } else if(markerValASR === 'ماموریت'){
          statusASR = 'mission';
        } else if(markerValASR && String(markerValASR).toLowerCase() === 'off'){
          statusASR = 'holiday';
        } else {
          var presentASR = false;
          pairsASR.forEach(function(p){
            if(p.inE && p.inE.jalaliDate === dateStrASR) presentASR = true;
            if(p.inE && p.outE && p.outE.jalaliDate === dateStrASR && p.inE.jalaliDate !== dateStrASR){
              var startDayASR = p.inE.jalaliDate;
              var firstCycleASR = pairsASR.filter(function(p2){ return p2.inE && p2.inE.jalaliDate === startDayASR; })[0];
              if(firstCycleASR === p) presentASR = true;
            }
          });
          statusASR = presentASR ? 'normal' : 'absence';
        }
        if(statusFilterASR === 'all' || statusFilterASR === statusASR){
          resultsASR.push({username: emp.username, fullName: emp.fullName, employeeCode: emp.employeeCode||'', jalaliDate: dateStrASR, status: statusASR});
        }
      }
    });

    return jsonOut({ok:true, results: resultsASR});
  }
  // Monthly payroll-style performance report (worked/required/useful/overtime/night/Friday/
  // holiday/leave hours per employee) — replicates the exact policy from the reference
  // spreadsheet: 7:30 standard check-in, 14:50 standard check-out, 7:20 required daily hours,
  // 1-hour break deducted whenever raw worked time exceeds 7:20, and night work counted for
  // any minutes worked after 22:00. Friday/holiday work is NOT required time — all hours
  // worked that day count entirely as "جمعه‌کاری"/"تعطیل‌کاری" instead of normal/overtime.
  if(action === 'getMonthlyPerformanceReport'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var yearMP = parseInt(body.year, 10);
    var monthMP = parseInt(body.month, 10);
    if(!yearMP || !monthMP) return jsonOut({ok:false, error:'year and month are required'});

    var eshMP = getSheet('Employees', ['id','fullName','username','password','deviceId','allowedProjects','employeeCode','freeZone','fingerprintRequired','jobTitle','shiftType','shiftStart','shiftEnd','nationalId','jobGroup','photoUrl','shiftId','active','sessionVersion','noDeviceLimit']);
    var employeesMP = readRows(eshMP).filter(function(e){ return String(e.active) !== 'false'; });

    var daysInMonthMP = jalaliMonthLengthServer(yearMP, monthMP);
    var firstDayGMP = jalaliToGregorianServer(yearMP, monthMP, 1);
    var startBoundMP = new Date(firstDayGMP[0], firstDayGMP[1]-1, firstDayGMP[2]);
    startBoundMP.setDate(startBoundMP.getDate() - 2);
    var lastDayGMP = jalaliToGregorianServer(yearMP, monthMP, daysInMonthMP);
    var endBoundMP = new Date(lastDayGMP[0], lastDayGMP[1]-1, lastDayGMP[2]);
    endBoundMP.setDate(endBoundMP.getDate() + 2);
    var startMsMP = startBoundMP.getTime();
    var endMsMP = endBoundMP.getTime();

    var lshMP = getSheet('Logs', ['id','jalaliDate','time','employee','username','project','type','lat','lng','accuracy','isoTimestamp','note','employeeCode','isHoliday','holidayReason','source']);
    var allLogsMP = readRows(lshMP).filter(function(r){
      var tms = new Date(r.isoTimestamp).getTime();
      return !isNaN(tms) && tms >= startMsMP && tms <= endMsMP;
    });

    var dsmshMP = getSheet('DayStatusMarkers', ['id','username','employee','employeeCode','jalaliDate','status','leaveHours','createdAt','createdTime']);
    var markersMP = readRows(dsmshMP);

    var hshMP = getSheet('Holidays', ['id','jalaliDate','reason','status']);
    var holidaySetMP = {};
    readRows(hshMP).forEach(function(h){ holidaySetMP[h.jalaliDate] = true; });

    // Fetched once, outside the per-employee loop below (same reasoning as holidaySetMP
    // above) — computeEmployeePerformance itself never re-reads these sheets per day/employee.
    var shiftsByIdMP = buildShiftsByIdMap();
    var assignSheetMP = getSheet('ShiftAssignments', ['id','username','startDate','endDate','shiftStart','shiftEnd','shiftName','shiftId','createdAt']);
    var allAssignmentsMP = readRows(assignSheetMP);
    var shiftDaysMapMP = buildShiftDaysMap();
    var normalSettingsMP = getNormalShiftSettings();

    var resultsMP = [];
    employeesMP.forEach(function(emp){
      var empLogsMP = allLogsMP.filter(function(r){ return r.username === emp.username; }).sort(function(a,b){ return isoCompare(a.isoTimestamp, b.isoTimestamp); });
      var empMarkersMP = markersMP.filter(function(m){ return m.username === emp.username; });
      var empAssignmentsMP = allAssignmentsMP.filter(function(r){ return r.username === emp.username; });
      var calc = computeEmployeePerformance(emp.username, yearMP, monthMP, empLogsMP, empMarkersMP, holidaySetMP, emp, empAssignmentsMP, shiftsByIdMP, shiftDaysMapMP, normalSettingsMP);
      resultsMP.push({
        username: emp.username, fullName: emp.fullName, employeeCode: emp.employeeCode||'',
        workedMin: calc.totals.workedMin, requiredMin: calc.totals.requiredMin, leaveMin: calc.totals.leaveMin,
        nightMin: calc.totals.nightMin, fridayMin: calc.totals.fridayMin, holidayMin: calc.totals.holidayMin,
        overtimeMin: calc.totals.overtimeMin, breakMin: calc.totals.breakMin, holidayBreakMin: calc.totals.holidayBreakMin, usefulMin: calc.totals.usefulMin,
        grandOvertimeMin: calc.totals.grandOvertimeMin, missionDaysCount: calc.totals.missionDaysCount, extraPresenceMin: calc.totals.extraPresenceMin
      });
    });

    return jsonOut({ok:true, results: resultsMP});
  }
  // Admin-only: one specific employee's full daily breakdown (including raw shift in/out
  // times) — used to generate the official print-form PDF export. Unlike getMyPerformance,
  // this is authenticated by the admin PIN, not the employee's own credentials.
  if(action === 'getEmployeePerformanceDetail'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var yearEPD = parseInt(body.year, 10);
    var monthEPD = parseInt(body.month, 10);
    if(!yearEPD || !monthEPD) return jsonOut({ok:false, error:'year and month are required'});
    var eshEPD = getSheet('Employees', ['id','fullName','username','password','deviceId','allowedProjects','employeeCode','freeZone','fingerprintRequired','jobTitle','shiftType','shiftStart','shiftEnd','nationalId','jobGroup','photoUrl','shiftId','active','sessionVersion','noDeviceLimit']);
    var empEPD = readRows(eshEPD).filter(function(r){ return r.username === String(body.username||'').trim(); })[0];
    if(!empEPD) return jsonOut({ok:false, error:'unknown_employee'});

    var daysInMonthEPD = jalaliMonthLengthServer(yearEPD, monthEPD);
    var firstDayGEPD = jalaliToGregorianServer(yearEPD, monthEPD, 1);
    var startBoundEPD = new Date(firstDayGEPD[0], firstDayGEPD[1]-1, firstDayGEPD[2]);
    startBoundEPD.setDate(startBoundEPD.getDate() - 2);
    var lastDayGEPD = jalaliToGregorianServer(yearEPD, monthEPD, daysInMonthEPD);
    var endBoundEPD = new Date(lastDayGEPD[0], lastDayGEPD[1]-1, lastDayGEPD[2]);
    endBoundEPD.setDate(endBoundEPD.getDate() + 2);

    var lshEPD = getSheet('Logs', ['id','jalaliDate','time','employee','username','project','type','lat','lng','accuracy','isoTimestamp','note','employeeCode','isHoliday','holidayReason','source']);
    var empLogsEPD = readRows(lshEPD).filter(function(r){
      var tms = new Date(r.isoTimestamp).getTime();
      return r.username === empEPD.username && !isNaN(tms) && tms >= startBoundEPD.getTime() && tms <= endBoundEPD.getTime();
    }).sort(function(a,b){ return isoCompare(a.isoTimestamp, b.isoTimestamp); });

    var dsmshEPD = getSheet('DayStatusMarkers', ['id','username','employee','employeeCode','jalaliDate','status','leaveHours','createdAt','createdTime']);
    var empMarkersEPD = readRows(dsmshEPD).filter(function(m){ return m.username === empEPD.username; });

    var hshEPD = getSheet('Holidays', ['id','jalaliDate','reason','status']);
    var holidaySetEPD = {};
    readRows(hshEPD).forEach(function(h){ holidaySetEPD[h.jalaliDate] = true; });

    var shiftsByIdEPD = buildShiftsByIdMap();
    var assignSheetEPD = getSheet('ShiftAssignments', ['id','username','startDate','endDate','shiftStart','shiftEnd','shiftName','shiftId','createdAt']);
    var assignmentsEPD = readRows(assignSheetEPD).filter(function(r){ return r.username === empEPD.username; });
    var shiftDaysMapEPD = buildShiftDaysMap();
    var normalSettingsEPD = getNormalShiftSettings();

    var calcEPD = computeEmployeePerformance(empEPD.username, yearEPD, monthEPD, empLogsEPD, empMarkersEPD, holidaySetEPD, empEPD, assignmentsEPD, shiftsByIdEPD, shiftDaysMapEPD, normalSettingsEPD);
    var dailyWithWeekday = calcEPD.daily.map(function(d){
      var pj = parseJalaliDateStrServer(d.jalaliDate);
      d.weekday = jalaliWeekdayNameServer(pj.y, pj.m, pj.d);
      return d;
    });
    return jsonOut({ok:true, fullName: empEPD.fullName, employeeCode: empEPD.employeeCode||'', totals: calcEPD.totals, daily: dailyWithWeekday});
  }
  // Employee-facing: their own monthly totals + full daily breakdown in one call, authenticated
  // by username+password (same credentials they already log in with) rather than the admin PIN.
  if(action === 'getLogsSince'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    // A SHARED, server-side cursor — not one each client tracks for itself — so a
    // notification for any single check-in/checkout fires exactly once in total, no matter
    // how many admin devices/browser sessions happen to be polling. Whichever device polls
    // first after a new event effectively "claims" it by advancing this shared cursor; any
    // other device polling moments later simply won't see that same entry anymore, instead
    // of every device independently re-notifying about everything it personally missed.
    var sharedCursorLS = getSetting('admin_notify_cursor');
    var sinceMsLS = sharedCursorLS ? new Date(sharedCursorLS).getTime() : (Date.now() - 5*60000);
    var lshLS = getSheet('Logs', ['id','jalaliDate','time','employee','username','project','type','lat','lng','accuracy','isoTimestamp','note','employeeCode','isHoliday','holidayReason','source']);
    var rowsLS = ensureIds(lshLS, readRows(lshLS))
      // Only genuine employee self-check-ins (GPS-based, source 'اتومات') are worth notifying
      // the admin about — a manual/bulk entry the admin JUST typed in themselves (source
      // 'دستی') doesn't need to notify them about their own just-performed action. Missing
      // this was exactly why a bulk save of many days (which creates many rows at once)
      // produced a burst of pointless notifications spread across the next several poll
      // cycles, one per row, as if a dozen people had just checked in.
      .filter(function(r){ return new Date(r.isoTimestamp).getTime() > sinceMsLS && r.source === 'اتومات'; })
      .map(function(r){ return {id:r.id, username:r.username, employee:r.employee, project:r.project, type:r.type, time:r.time, jalaliDate:r.jalaliDate, isoTimestamp:r.isoTimestamp}; })
      .sort(function(a,b){ return isoCompare(a.isoTimestamp, b.isoTimestamp); });
    var nowIsoLS = new Date().toISOString();
    setSetting('admin_notify_cursor', nowIsoLS);
    return jsonOut({ok:true, logs: rowsLS, serverNowIso: nowIsoLS});
  }
  if(action === 'getRecentLogs'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var lsh4r = getSheet('Logs', ['id','jalaliDate','time','employee','username','project','type','lat','lng','accuracy','isoTimestamp','note','employeeCode','isHoliday','holidayReason','source']);
    var daysBack = body.daysBack || 3;
    var cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - daysBack);
    var cutoffMs = cutoff.getTime();
    var rows6r = readRows(lsh4r).filter(function(r){ return new Date(r.isoTimestamp).getTime() >= cutoffMs; }).map(function(r){ delete r._row; return r; });
    return jsonOut({ok:true, logs: rows6r});
  }
  if(action === 'getLogsForUser'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var lsh4u = getSheet('Logs', ['id','jalaliDate','time','employee','username','project','type','lat','lng','accuracy','isoTimestamp','note','employeeCode','isHoliday','holidayReason','source']);
    var rows6u = ensureIds(lsh4u, readRows(lsh4u)).filter(function(r){
      return r.username === body.username || (body.employeeName && r.employee === body.employeeName);
    }).map(function(r){ delete r._row; return r; });
    return jsonOut({ok:true, logs: rows6u});
  }
  if(action === 'getOpenSessions'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var lshOS = getSheet('Logs', ['id','jalaliDate','time','employee','username','project','type','lat','lng','accuracy','isoTimestamp','note','employeeCode','isHoliday','holidayReason','source']);
    var eshOS = getSheet('Employees', ['id','fullName','username','password','deviceId','allowedProjects','employeeCode','freeZone','fingerprintRequired','jobTitle','shiftType','shiftStart','shiftEnd','nationalId','jobGroup','photoUrl','shiftId','active','sessionVersion','noDeviceLimit']);
    var empByUserOS = {};
    readRows(eshOS).forEach(function(e){ empByUserOS[e.username] = e; });
    var allLogsOS = readRows(lshOS);
    var nowMsOS = new Date().getTime();
    var byUserOS = {};
    // Group by username only (not project) — an employee can check in on one project and
    // check out on a different one (e.g. moved sites mid-shift); that check-out should
    // still count as closing the session, not leave it "still open" for the first project.
    allLogsOS.forEach(function(r){ (byUserOS[r.username] = byUserOS[r.username] || []).push(r); });
    var openList = [];
    Object.keys(byUserOS).forEach(function(usernameOS){
      var userLogsOS = byUserOS[usernameOS].slice().sort(function(a,b){ return isoCompare(a.isoTimestamp, b.isoTimestamp); });
      var latestOS = userLogsOS[userLogsOS.length-1];
      if(!latestOS || latestOS.type !== 'ورود') return;
      // Judge by elapsed TIME since the check-in, not by calendar date. A plain date
      // comparison ("checked in today = not open yet") breaks for night shifts: an
      // employee who checks in this evening and is due to check out tomorrow morning
      // would wrongly look "still today" — and stay hidden — for hours after they
      // were actually already overdue. Instead we compare against how long their
      // actual assigned shift runs (handles overnight shifts, since shiftEnd < shiftStart
      // just means the duration wraps past midnight), plus a grace buffer.
      var hoursSinceOS = (nowMsOS - new Date(latestOS.isoTimestamp).getTime()) / 3600000;
      var expectedHoursOS = 16; // fallback if the employee has no configured shift
      var empOS = empByUserOS[latestOS.username];
      if(empOS){
        var effShiftOS = getEffectiveShift(latestOS.username, latestOS.jalaliDate, empOS);
        if(effShiftOS.shiftStart && effShiftOS.shiftEnd){
          var durMinOS = timeToMinutes(effShiftOS.shiftEnd) - timeToMinutes(effShiftOS.shiftStart);
          if(durMinOS <= 0) durMinOS += 24*60; // overnight shift: end time is on the next day
          expectedHoursOS = (durMinOS/60) + 2; // +2h grace before flagging as overdue
        }
      }
      if(hoursSinceOS > expectedHoursOS){
        openList.push({
          id: latestOS.id, username: latestOS.username, employee: latestOS.employee, employeeCode: latestOS.employeeCode||'',
          project: latestOS.project, jalaliDate: latestOS.jalaliDate, time: latestOS.time, isoTimestamp: latestOS.isoTimestamp
        });
      }
    });
    openList.sort(function(a,b){ return isoCompare(a.isoTimestamp, b.isoTimestamp); });
    return jsonOut({ok:true, openSessions: openList});
  }
  if(action === 'adminCloseOpenSession'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var lshCS = getSheet('Logs', ['id','jalaliDate','time','employee','username','project','type','lat','lng','accuracy','isoTimestamp','note','employeeCode','isHoliday','holidayReason','source']);
    var allRowsCS = readRows(lshCS);
    var openRowCS = allRowsCS.filter(function(r){ return r.id === body.id; })[0];
    if(!openRowCS) return jsonOut({ok:false, error:'not found'});
    if(openRowCS.type !== 'ورود') return jsonOut({ok:false, error:'not_an_open_checkin'});

    var newIsoCS = jalaliDateTimeToIso(body.jalaliDate, body.time);
    if(!newIsoCS) return jsonOut({ok:false, error:'invalid_date_time'});
    if(new Date(newIsoCS).getTime() <= new Date(openRowCS.isoTimestamp).getTime()){
      return jsonOut({ok:false, error:'must_be_after_checkin'});
    }
    // only real restriction: no OTHER log entry for this employee may fall strictly between the open check-in and this new checkout —
    // that would mean the employee already did something else in between (e.g. checked in again), a genuine conflict
    var openMsCS = new Date(openRowCS.isoTimestamp).getTime();
    var newMsCS = new Date(newIsoCS).getTime();
    var conflictCS = allRowsCS.filter(function(r){
      var rMsCS = new Date(r.isoTimestamp).getTime();
      return r.username === openRowCS.username && r.id !== openRowCS.id && rMsCS > openMsCS && rMsCS < newMsCS;
    })[0];
    if(conflictCS){
      return jsonOut({ok:false, error:'conflict', at: conflictCS.jalaliDate + ' ' + conflictCS.time + ' (' + conflictCS.type + ')'});
    }

    var holInfoCS = getHolidayInfo(body.jalaliDate);
    var lidCS = Utilities.getUuid();
    appendRowByHeader(lshCS, {
      id:lidCS, jalaliDate:body.jalaliDate, time:body.time, employee:openRowCS.employee, username:openRowCS.username, project:body.project||openRowCS.project,
      type:'خروج', lat:'', lng:'', accuracy:'', isoTimestamp:newIsoCS,
      note: 'ثبت مستقیم توسط مدیر (تکمیل تردد ناقص)' + (body.reason ? (' — ' + body.reason) : ''),
      employeeCode: openRowCS.employeeCode || '',
      isHoliday: holInfoCS.isHoliday ? 'تعطیل' : '', holidayReason: holInfoCS.reason, source:'دستی'
    });
    return jsonOut({ok:true});
  }
  if(action === 'adminAddLog'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var eshA = getSheet('Employees', ['id','fullName','username','password','deviceId','allowedProjects','employeeCode','freeZone','fingerprintRequired','jobTitle','shiftType','shiftStart','shiftEnd','nationalId','jobGroup','photoUrl','shiftId','active','sessionVersion','noDeviceLimit']);
    var empA = readRows(eshA).filter(function(r){ return r.username === body.username; })[0];
    if(!empA) return jsonOut({ok:false, error:'unknown_employee'});
    var lshA = getSheet('Logs', ['id','jalaliDate','time','employee','username','project','type','lat','lng','accuracy','isoTimestamp','note','employeeCode','isHoliday','holidayReason','source']);
    var isoA = jalaliDateTimeToIso(body.jalaliDate, body.time) || serverNow().isoTimestamp;
    var empRowsA = readRows(lshA).filter(function(r){ return r.username === empA.username; });
    // A checkout must always have an open check-in before it, chronologically — regardless of
    // which admin tool created it. This is checked by pairing order, not same-calendar-day
    // matching, so a legitimate overnight shift (check-in one day, checkout early the next)
    // is never blocked by it.
    if(body.type === 'out' && wouldBeOrphanCheckout(empRowsA, isoA)){
      return jsonOut({ok:false, error:'no_open_checkin'});
    }
    if(body.relaxedOrder){
      // Used from the report's ✏️ edit view — that view exists specifically to fix/backfill
      // older records, so we don't require strict "after last entry" ordering or in/out
      // alternation here. The only real requirement left is that it doesn't land at the exact
      // same instant as another existing entry for this employee — a genuine time conflict.
      var conflictA = empRowsA.filter(function(r){ return new Date(r.isoTimestamp).getTime() === new Date(isoA).getTime(); })[0];
      if(conflictA){
        return jsonOut({ok:false, error:'time_conflict', at: conflictA.jalaliDate + ' ' + conflictA.time});
      }
    } else {
      var lastA = empRowsA.length ? empRowsA.sort(function(a,b){ return isoCompare(a.isoTimestamp, b.isoTimestamp); })[empRowsA.length-1] : null;
      if(lastA && new Date(isoA).getTime() <= new Date(lastA.isoTimestamp).getTime()){
        return jsonOut({ok:false, error:'must_be_after_last_entry'});
      }
      // Types must alternate — can't add a second check-in in a row (or a check-out with
      // no open check-in before it), even if it lands on a later calendar date.
      var newTypeA = body.type === 'in' ? 'ورود' : 'خروج';
      if(lastA && lastA.type === newTypeA){
        return jsonOut({ok:false, error: newTypeA === 'ورود' ? 'expected_checkout_first' : 'expected_checkin_first'});
      }
    }
    var holInfoA = getHolidayInfo(body.jalaliDate);
    var lidA = Utilities.getUuid();
    appendRowByHeader(lshA, {
      id:lidA, jalaliDate:body.jalaliDate, time:body.time, employee:empA.fullName, username:empA.username, project:body.project||'',
      type: body.type === 'in' ? 'ورود' : 'خروج', lat:'', lng:'', accuracy:'', isoTimestamp:isoA,
      note: 'ثبت مستقیم توسط مدیر' + (body.reason ? (' — ' + body.reason) : ''),
      employeeCode: empA.employeeCode || '',
      isHoliday: holInfoA.isHoliday ? 'تعطیل' : '', holidayReason: holInfoA.reason, source:'دستی'
    });
    return jsonOut({ok:true});
  }
  if(action === 'getDayStatusForUser'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var dsmshG = getSheet('DayStatusMarkers', ['id','username','employee','employeeCode','jalaliDate','status','leaveHours','createdAt','createdTime']);
    var rowsG = readRows(dsmshG).filter(function(r){ return r.username === body.username; }).map(function(r){ delete r._row; return r; });
    return jsonOut({ok:true, statuses: rowsG});
  }
  if(action === 'syncDayStatusForMonth'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var eshDS = getSheet('Employees', ['id','fullName','username','password','deviceId','allowedProjects','employeeCode','freeZone','fingerprintRequired','jobTitle','shiftType','shiftStart','shiftEnd','nationalId','jobGroup','photoUrl','shiftId','active','sessionVersion','noDeviceLimit']);
    var empDS = readRows(eshDS).filter(function(r){ return r.username === body.username; })[0];
    if(!empDS) return jsonOut({ok:false, error:'unknown_employee'});
    var dsmsh2 = getSheet('DayStatusMarkers', ['id','username','employee','employeeCode','jalaliDate','status','leaveHours','createdAt','createdTime']);
    var datesToClear = body.allDatesInMonth || [];
    var existingRows = readRows(dsmsh2).filter(function(r){
      return r.username === body.username && datesToClear.indexOf(r.jalaliDate) !== -1 && r.status !== 'مرخصی' && r.status !== 'ماموریت';
    });
    existingRows.sort(function(a,b){ return b._row - a._row; });
    existingRows.forEach(function(r){ dsmsh2.deleteRow(r._row); });
    var tDS = serverNow();
    var newEntries = body.entries || [];
    newEntries.forEach(function(e){
      var idDS = Utilities.getUuid();
      appendRowByHeader(dsmsh2, {id:idDS, username:empDS.username, employee:empDS.fullName, employeeCode: empDS.employeeCode||'', jalaliDate:e.jalaliDate, status:e.status, createdAt: tDS.jalaliDate, createdTime: tDS.time});
    });
    return jsonOut({ok:true, count: newEntries.length});
  }
  if(action === 'adminAddLeaveBatch'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var eshLB = getSheet('Employees', ['id','fullName','username','password','deviceId','allowedProjects','employeeCode','freeZone','fingerprintRequired','jobTitle','shiftType','shiftStart','shiftEnd','nationalId','jobGroup','photoUrl','shiftId','active','sessionVersion','noDeviceLimit']);
    var empLB = readRows(eshLB).filter(function(r){ return r.username === body.username; })[0];
    if(!empLB) return jsonOut({ok:false, error:'unknown_employee'});
    var lrshLB = getSheet('LeaveRequests', ['id','username','employee','employeeCode','startDate','endDate','reason','status','createdAt','createdTime','leaveType','startTime','endTime','requestType','leaveCategory','project']);
    var tLB = serverNow();
    var entriesLB = body.entries || (body.days||[]).map(function(d){ return {date:d, requestType: body.requestType||'leave', leaveType:'daily'}; });
    entriesLB.forEach(function(entry){
      var lidLB = Utilities.getUuid();
      var reqTypeLB = entry.requestType || 'leave';
      var leaveTypeLB = entry.leaveType || 'daily';
      appendRowByHeader(lrshLB, {
        id:lidLB, username:empLB.username, employee:empLB.fullName, employeeCode: empLB.employeeCode||'',
        startDate:entry.date, endDate:entry.date, reason: body.reason||'ثبت مستقیم توسط مدیر',
        status:'approved', createdAt:tLB.jalaliDate, createdTime:tLB.time,
        leaveType:leaveTypeLB, startTime: entry.startTime||'', endTime: entry.endTime||'',
        requestType: reqTypeLB, leaveCategory: body.leaveCategory || 'استحقاقی', project: entry.project || body.project || ''
      });
      syncLeaveDayStatus(empLB.username, empLB.fullName, entry.date, entry.date, leaveTypeLB, entry.startTime||'', entry.endTime||'', reqTypeLB, empLB.employeeCode);
    });
    return jsonOut({ok:true, count: entriesLB.length});
  }
  if(action === 'adminAddLogBatch'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var eshB = getSheet('Employees', ['id','fullName','username','password','deviceId','allowedProjects','employeeCode','freeZone','fingerprintRequired','jobTitle','shiftType','shiftStart','shiftEnd','nationalId','jobGroup','photoUrl','shiftId','active','sessionVersion','noDeviceLimit']);
    var empB = readRows(eshB).filter(function(r){ return r.username === body.username; })[0];
    if(!empB) return jsonOut({ok:false, error:'unknown_employee'});
    var entriesB = body.entries || [];
    if(!entriesB.length) return jsonOut({ok:false, error:'no_entries'});

    var lshB = getSheet('Logs', ['id','jalaliDate','time','employee','username','project','type','lat','lng','accuracy','isoTimestamp','note','employeeCode','isHoliday','holidayReason','source']);
    // NOTE: intentionally NOT checked against the employee's overall existing history for ordering —
    // this bulk feature is for backfilling any historical month (e.g. an employee without a phone),
    // and admin may fill months in any order. Only internal same-batch ordering, and exact duplicates
    // against existing rows, are validated below.
    var empRowsB = readRows(lshB).filter(function(r){ return r.username === body.username; });
    var existingIsoSet = {};
    var existingKeySet = {};
    empRowsB.forEach(function(r){
      existingIsoSet[String(r.isoTimestamp)] = true;
      var normType = r.type === 'ورود' ? 'in' : (r.type === 'خروج' ? 'out' : r.type);
      existingKeySet[String(r.jalaliDate).trim() + '|' + String(r.time).trim() + '|' + normType] = true;
    });

    var computedB = [];
    for(var bi=0; bi<entriesB.length; bi++){
      var eb = entriesB[bi];
      var isoB = jalaliDateTimeToIso(eb.jalaliDate, eb.time);
      if(!isoB) return jsonOut({ok:false, error:'invalid_date_time', at: eb.jalaliDate + ' ' + eb.time});
      var keyB = String(eb.jalaliDate).trim() + '|' + String(eb.time).trim() + '|' + eb.type;
      if(existingIsoSet[String(isoB)] || existingKeySet[keyB]){
        return jsonOut({ok:false, error:'duplicate_entry', at: eb.jalaliDate + ' ' + eb.time});
      }
      computedB.push({jalaliDate:eb.jalaliDate, time:eb.time, type:eb.type, iso:isoB});
    }
    computedB.sort(function(a,c){ return String(a.iso).localeCompare(String(c.iso)); });

    // The one thing still worth checking against existing history, even in backfill mode:
    // if this batch's very FIRST entry is a checkout, there must be a genuinely open ورود
    // (from existing data) before it — otherwise it's a floating checkout with nothing to
    // close, exactly the case this is meant to catch. Everything else about ordering within
    // the batch stays as flexible as before; this doesn't touch that. Purely chronological
    // (via wouldBeOrphanCheckout), so a legitimate overnight shift is never blocked by it.
    if(computedB.length && computedB[0].type === 'out' && wouldBeOrphanCheckout(empRowsB, computedB[0].iso)){
      return jsonOut({ok:false, error:'no_open_checkin', at: computedB[0].jalaliDate + ' ' + computedB[0].time});
    }

    var runningLast = null;
    var runningLastType = null;
    for(var bj=0; bj<computedB.length; bj++){
      if(runningLast && String(computedB[bj].iso) === String(runningLast)){
        return jsonOut({ok:false, error:'duplicate_entry', at: computedB[bj].jalaliDate + ' ' + computedB[bj].time});
      }
      if(runningLast && String(computedB[bj].iso) < String(runningLast)){
        return jsonOut({ok:false, error:'must_be_after_last_entry', at: computedB[bj].jalaliDate + ' ' + computedB[bj].time});
      }
      // Types must alternate (in, out, in, out, ...) — a new check-in can't come before
      // the check-out that should close the previous one, even if it lands on a later
      // calendar date (e.g. a night-shift check-in the same day an earlier shift's
      // check-out was recorded).
      if(runningLastType && runningLastType === computedB[bj].type){
        return jsonOut({ok:false, error: computedB[bj].type === 'in' ? 'expected_checkout_first' : 'expected_checkin_first', at: computedB[bj].jalaliDate + ' ' + computedB[bj].time});
      }
      runningLast = computedB[bj].iso;
      runningLastType = computedB[bj].type;
    }

    var t5 = serverNow();
    computedB.forEach(function(eb){
      var holInfoB = getHolidayInfo(eb.jalaliDate);
      var lidB = Utilities.getUuid();
      appendRowByHeader(lshB, {
        id:lidB, jalaliDate:eb.jalaliDate, time:eb.time, employee:empB.fullName, username:empB.username, project:body.project||'',
        type: eb.type === 'in' ? 'ورود' : 'خروج', lat:'', lng:'', accuracy:'', isoTimestamp:eb.iso,
        note: 'ثبت مستقیم توسط مدیر (ماهانه)' + (body.reason ? (' — ' + body.reason) : ''),
        employeeCode: empB.employeeCode || '',
        isHoliday: holInfoB.isHoliday ? 'تعطیل' : '', holidayReason: holInfoB.reason, source:'دستی'
      });
    });

    return jsonOut({ok:true, count: computedB.length});
  }
  if(action === 'getPendingOfflineLogs'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var poshG = getSheet('PendingOfflineLogs', ['id','username','employee','employeeCode','project','type','jalaliDate','time','clientIso','lat','lng','accuracy','reason','status','createdAt']);
    var rowsG = readRows(poshG).filter(function(r){ return r.status === 'pending'; }).map(function(r){ delete r._row; return r; });
    return jsonOut({ok:true, pending: rowsG});
  }
  if(action === 'approveOfflineLog'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var poshA = getSheet('PendingOfflineLogs', ['id','username','employee','employeeCode','project','type','jalaliDate','time','clientIso','lat','lng','accuracy','reason','status','createdAt']);
    var prowA = readRows(poshA).filter(function(r){ return r.id === body.id; })[0];
    if(!prowA) return jsonOut({ok:false, error:'not found'});
    var useDateO = body.jalaliDate || prowA.jalaliDate;
    var useTimeO = body.time || prowA.time;
    var isoO = jalaliDateTimeToIso(useDateO, useTimeO) || serverNow().isoTimestamp;
    var lshO = getSheet('Logs', ['id','jalaliDate','time','employee','username','project','type','lat','lng','accuracy','isoTimestamp','note','employeeCode','isHoliday','holidayReason','source']);
    var empRowsO = readRows(lshO).filter(function(r){ return r.username === prowA.username; });
    var lastO = empRowsO.length ? empRowsO.sort(function(a,b){ return isoCompare(a.isoTimestamp, b.isoTimestamp); })[empRowsO.length-1] : null;
    if(lastO && String(isoO) <= String(lastO.isoTimestamp)){
      return jsonOut({ok:false, error:'must_be_after_last_entry'});
    }
    if(prowA.type === 'out' && wouldBeOrphanCheckout(empRowsO, isoO)){
      return jsonOut({ok:false, error:'no_open_checkin'});
    }
    var holInfoO = getHolidayInfo(useDateO);
    var lidO = Utilities.getUuid();
    appendRowByHeader(lshO, {
      id:lidO, jalaliDate:useDateO, time:useTimeO, employee:prowA.employee, username:prowA.username, project:prowA.project,
      type: prowA.type === 'in' ? 'ورود' : 'خروج', lat:prowA.lat||'', lng:prowA.lng||'', accuracy:prowA.accuracy||'',
      isoTimestamp:isoO, note: '⚠ ثبت آفلاین (تاییدشده توسط مدیر)', employeeCode: prowA.employeeCode || '',
      isHoliday: holInfoO.isHoliday ? 'تعطیل' : '', holidayReason: holInfoO.reason, source:'اتومات'
    });
    esh_set(poshA, prowA._row, 'status', 'approved');
    return jsonOut({ok:true});
  }
  if(action === 'rejectOfflineLog'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var poshR = getSheet('PendingOfflineLogs', ['id','username','employee','employeeCode','project','type','jalaliDate','time','clientIso','lat','lng','accuracy','reason','status','createdAt']);
    var prowR = readRows(poshR).filter(function(r){ return r.id === body.id; })[0];
    if(!prowR) return jsonOut({ok:false, error:'not found'});
    esh_set(poshR, prowR._row, 'status', 'rejected');
    return jsonOut({ok:true});
  }
  if(action === 'rejectOfflineLogBatch'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var poshRB = getSheet('PendingOfflineLogs', ['id','username','employee','employeeCode','project','type','jalaliDate','time','clientIso','lat','lng','accuracy','reason','status','createdAt']);
    var idsSet = {};
    (body.ids||[]).forEach(function(x){ idsSet[x] = true; });
    var rowsRB = readRows(poshRB).filter(function(r){ return r.status === 'pending' && idsSet[r.id]; });
    rowsRB.forEach(function(r){ esh_set(poshRB, r._row, 'status', 'rejected'); });
    return jsonOut({ok:true, count: rowsRB.length});
  }
  if(action === 'recalcAllLogTimestamps'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var lshR = getSheet('Logs', ['id','jalaliDate','time','employee','username','project','type','lat','lng','accuracy','isoTimestamp','note','employeeCode','isHoliday','holidayReason','source']);
    var lastRowR = lshR.getLastRow();
    var lastColR = lshR.getLastColumn();
    if(lastRowR < 2) return jsonOut({ok:true, total:0, fixed:0});
    var headerRowR = lshR.getRange(1,1,1,lastColR).getValues()[0];
    var idxMapR = {};
    headerRowR.forEach(function(h,i){ idxMapR[h] = i; });

    var hshR2 = getSheet('Holidays', ['id','jalaliDate','reason','status']);
    var holidayMapR = {};
    readRows(hshR2).forEach(function(h){ holidayMapR[h.jalaliDate] = h.reason || 'تعطیل رسمی'; });
    function getHolidayInfoFast(jalaliDateStr){
      if(isFridayFromJalaliDateStr(jalaliDateStr)) return {isHoliday:true, reason:'جمعه'};
      if(holidayMapR.hasOwnProperty(jalaliDateStr)) return {isHoliday:true, reason: holidayMapR[jalaliDateStr]};
      return {isHoliday:false, reason:''};
    }

    var dataRangeR = lshR.getRange(2,1,lastRowR-1,lastColR);
    var valuesR = dataRangeR.getValues();
    var fixedCount = 0;
    for(var ri=0; ri<valuesR.length; ri++){
      var rowV = valuesR[ri];
      var jalaliDateV = rowV[idxMapR['jalaliDate']];
      var timeV = rowV[idxMapR['time']];
      var newIsoR = jalaliDateTimeToIso(jalaliDateV, timeV);
      var holInfoR = getHolidayInfoFast(jalaliDateV);
      var curIso = rowV[idxMapR['isoTimestamp']];
      if(newIsoR && newIsoR !== curIso){
        rowV[idxMapR['isoTimestamp']] = newIsoR;
        fixedCount++;
      }
      var wantHoliday = holInfoR.isHoliday ? 'تعطیل' : '';
      if(String(rowV[idxMapR['isHoliday']]||'') !== wantHoliday || String(rowV[idxMapR['holidayReason']]||'') !== holInfoR.reason){
        rowV[idxMapR['isHoliday']] = wantHoliday;
        rowV[idxMapR['holidayReason']] = holInfoR.reason;
      }
    }
    dataRangeR.setNumberFormat('@'); // keep everything plain text before writing back
    dataRangeR.setValues(valuesR);
    return jsonOut({ok:true, total: valuesR.length, fixed: fixedCount});
  }
  if(action === 'runOffMarkerJobManually'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    runDailyOffMarkerJob();
    return jsonOut({ok:true});
  }
  if(action === 'backfillMissionLogs'){
    var lshBF = getSheet('Logs', ['id','jalaliDate','time','employee','username','project','type','lat','lng','accuracy','isoTimestamp','note','employeeCode','isHoliday','holidayReason','source']);
    var lrshBF = getSheet('LeaveRequests', ['id','username','employee','employeeCode','startDate','endDate','reason','status','createdAt','createdTime','leaveType','startTime','endTime','requestType','leaveCategory','project']);
    var missionsBF = readRows(lrshBF).filter(function(r){
      return r.requestType === 'mission' && r.status === 'approved' && r.leaveType === 'hourly' && r.startTime && r.endTime;
    });
    if(missionsBF.length === 0) return jsonOut({ok:true, missionLogsAdded: 0});

    var existingLogsBF = readRows(lshBF);
    var existingIndexBF = {};
    existingLogsBF.forEach(function(r){ existingIndexBF[r.username+'|'+r.isoTimestamp+'|'+r.type] = true; });

    var missionLogsAdded = 0;
    var newRowsToAppend = [];
    missionsBF.forEach(function(m){
      var projBF = m.project || '';
      var noteBF = projBF
        ? 'ثبت مستقیم توسط مدیر (ماموریت) — بازیابی‌شده'
        : 'ثبت مستقیم توسط مدیر (ماموریت) — بازیابی‌شده، پروژه رو دستی وارد کن';
      var startIsoBF = jalaliDateTimeToIso(m.startDate, m.startTime);
      var keyIn = m.username+'|'+startIsoBF+'|ورود';
      if(startIsoBF && !existingIndexBF[keyIn]){
        newRowsToAppend.push({
          id: Utilities.getUuid(), jalaliDate: m.startDate, time: m.startTime, employee: m.employee, username: m.username, project: projBF,
          type: 'ورود', lat:'', lng:'', accuracy:'', isoTimestamp: startIsoBF,
          note: noteBF, employeeCode: m.employeeCode || '',
          isHoliday: '', holidayReason: '', source: 'دستی'
        });
        existingIndexBF[keyIn] = true;
        missionLogsAdded++;
      }
      var stMinBF = timeToMinutes(m.startTime);
      var enMinBF = timeToMinutes(m.endTime);
      var outDateBF = (enMinBF <= stMinBF) ? addDaysToJalaliServer(m.startDate, 1) : m.startDate;
      var endIsoBF = jalaliDateTimeToIso(outDateBF, m.endTime);
      var keyOut = m.username+'|'+endIsoBF+'|خروج';
      if(endIsoBF && !existingIndexBF[keyOut]){
        newRowsToAppend.push({
          id: Utilities.getUuid(), jalaliDate: outDateBF, time: m.endTime, employee: m.employee, username: m.username, project: projBF,
          type: 'خروج', lat:'', lng:'', accuracy:'', isoTimestamp: endIsoBF,
          note: noteBF, employeeCode: m.employeeCode || '',
          isHoliday: '', holidayReason: '', source: 'دستی'
        });
        existingIndexBF[keyOut] = true;
        missionLogsAdded++;
      }
    });

    if(newRowsToAppend.length){
      var headersBF = lshBF.getRange(1,1,1,lshBF.getLastColumn()).getValues()[0];
      var rowsAsArraysBF = newRowsToAppend.map(function(obj){
        return headersBF.map(function(h){ return obj.hasOwnProperty(h) ? obj[h] : ''; });
      });
      var targetRangeBF = lshBF.getRange(lshBF.getLastRow()+1, 1, rowsAsArraysBF.length, headersBF.length);
      targetRangeBF.setNumberFormat('@'); // plain text before writing — see note in appendRowByHeader
      targetRangeBF.setValues(rowsAsArraysBF);
    }

    return jsonOut({ok:true, missionLogsAdded: missionLogsAdded});
  }
  // Separate, fixed password (ADVANCED_TOOLS_PASSWORD at the top of this file) gating the
  // "تعمیرِ داده‌های قدیمی" tools — deliberately independent of the regular admin PIN, since
  // that PIN may be known to several day-to-day admins while these rarely-needed/riskier
  // tools should stay restricted to whoever knows this separate password.
  if(action === 'checkAdvancedToolsPassword'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    return jsonOut({ok: String(body.password) === String(ADVANCED_TOOLS_PASSWORD)});
  }
  if(action === 'backfillLogSource'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var lshBF = getSheet('Logs', ['id','jalaliDate','time','employee','username','project','type','lat','lng','accuracy','isoTimestamp','note','employeeCode','isHoliday','holidayReason','source']);
    var rowsBF = readRows(lshBF);
    var filledCount = 0;
    rowsBF.forEach(function(r){
      if(String(r.source||'').trim() !== '') return; // already set — leave it alone
      var noteBF = String(r.note || '');
      var guessed;
      if(noteBF.indexOf('ثبت مستقیم توسط مدیر') !== -1) guessed = 'دستی';
      else if(noteBF.indexOf('ثبت دستی') !== -1) guessed = 'دستی';
      else if(noteBF.indexOf('ثبت آفلاین') !== -1) guessed = 'اتومات';
      else guessed = 'اتومات'; // default: a real GPS check-in with no special note
      lshBF.getRange(r._row, colIndex(lshBF,'source')).setValue(guessed);
      filledCount++;
    });
    return jsonOut({ok:true, total: rowsBF.length, filled: filledCount});
  }
  // Every sheet that references an employee stores BOTH their username (stable) and their
  // plain-text full name (a snapshot from whenever that row was created). Renaming someone
  // in Employees never touches those old snapshots, so older rows keep showing the name they
  // had at the time — this walks every such sheet and brings every row's name field back in
  // sync with each employee's CURRENT name, matched reliably by username.
  if(action === 'syncEmployeeNamesEverywhere'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var eshSyncN = getSheet('Employees', ['id','fullName','username','password','deviceId','allowedProjects','employeeCode','freeZone','fingerprintRequired','jobTitle','shiftType','shiftStart','shiftEnd','nationalId','jobGroup','photoUrl','shiftId','active','sessionVersion','noDeviceLimit']);
    var nameByUsernameN = {};
    readRows(eshSyncN).forEach(function(e){ if(e.username) nameByUsernameN[e.username] = e.fullName; });

    var sheetsToSyncN = [
      {name:'Logs', headers:['id','jalaliDate','time','employee','username','project','type','lat','lng','accuracy','isoTimestamp','note','employeeCode','isHoliday','holidayReason','source']},
      {name:'DayStatusMarkers', headers:['id','username','employee','employeeCode','jalaliDate','status','leaveHours','createdAt','createdTime']},
      {name:'LeaveRequests', headers:['id','username','employee','employeeCode','startDate','endDate','reason','status','createdAt','createdTime','leaveType','startTime','endTime','requestType','leaveCategory','project']},
      {name:'Requests', headers:['id','username','employee','employeeCode','project','type','reason','status','createdAt','createdTime','requestedDate','requestedTime']},
      {name:'PendingOfflineLogs', headers:['id','username','employee','employeeCode','project','type','jalaliDate','time','clientIso','lat','lng','accuracy','reason','status','createdAt']}
    ];
    var totalFixedN = 0;
    sheetsToSyncN.forEach(function(sInfo){
      var sh = getSheet(sInfo.name, sInfo.headers);
      var lastRow = sh.getLastRow();
      if(lastRow < 2) return;
      var lastCol = sh.getLastColumn();
      var empColIdx = colIndex(sh, 'employee');
      var userColIdx = colIndex(sh, 'username');
      if(!empColIdx || !userColIdx) return;
      var range = sh.getRange(2, 1, lastRow-1, lastCol);
      var values = range.getValues();
      var changed = false;
      for(var i=0; i<values.length; i++){
        var uname = String(values[i][userColIdx-1]||'');
        var correctName = nameByUsernameN[uname];
        if(correctName && String(values[i][empColIdx-1]||'') !== correctName){
          values[i][empColIdx-1] = correctName;
          changed = true;
          totalFixedN++;
        }
      }
      if(changed) range.setValues(values);
    });
    return jsonOut({ok:true, fixed: totalFixedN});
  }
  if(action === 'setupSheetFilters'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var sheetNames = ['Logs','Employees','Projects','Requests','LeaveRequests','PendingOfflineLogs','DayStatusMarkers','Holidays'];
    var doneList = [];
    sheetNames.forEach(function(name){
      var sh;
      try { sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); } catch(e){ sh = null; }
      if(!sh) return;
      var lastRow = sh.getLastRow();
      var lastCol = sh.getLastColumn();
      if(lastRow < 1 || lastCol < 1) return;
      var existingFilter = sh.getFilter();
      if(existingFilter) existingFilter.remove();
      var range = sh.getRange(1, 1, lastRow, lastCol);
      range.createFilter();
      doneList.push(name);
    });
    // Real chronological sort — only Logs has a genuinely sortable date key (isoTimestamp).
    // Other sheets store dates as Jalali text (e.g. "3 مرداد 1405"), which does not sort correctly
    // as plain text, so they are left in their natural (insertion) order with a filter applied,
    // letting the person sort/filter interactively from the header dropdown.
    var lshSort = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Logs');
    if(lshSort && lshSort.getLastRow() > 1){
      var isoColIdx = colIndex(lshSort, 'isoTimestamp');
      if(isoColIdx > 0){
        lshSort.getRange(2, 1, lshSort.getLastRow()-1, lshSort.getLastColumn()).sort({column: isoColIdx, ascending: true});
      }
    }
    return jsonOut({ok:true, sheets: doneList});
  }
  if(action === 'getHolidays'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var hsh = getSheet('Holidays', ['id','jalaliDate','reason','status']);
    return jsonOut({ok:true, holidays: readRows(hsh).map(function(r){ delete r._row; return r; })});
  }
  if(action === 'addHoliday'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var hsh2 = getSheet('Holidays', ['id','jalaliDate','reason','status']);
    var hid = Utilities.getUuid();
    appendRowByHeader(hsh2, {id:hid, jalaliDate: body.jalaliDate, reason: body.reason||'', status:'تعطیل'});
    return jsonOut({ok:true});
  }
  if(action === 'deleteHoliday'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var hsh3 = getSheet('Holidays', ['id','jalaliDate','reason','status']);
    var hrow = readRows(hsh3).filter(function(r){ return r.id === body.id; })[0];
    if(hrow) hsh3.deleteRow(hrow._row);
    return jsonOut({ok:true});
  }
  if(action === 'updateLog'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var res1 = updateOneLogRow(body);
    return jsonOut(res1);
  }
  if(action === 'updateLogBatch'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var updates = body.updates || [];
    for(var ui=0; ui<updates.length; ui++){
      var r1 = updateOneLogRow(updates[ui]);
      if(!r1.ok) return jsonOut({ok:false, error:r1.error, at: updates[ui].jalaliDate + ' ' + updates[ui].time});
    }
    return jsonOut({ok:true, count: updates.length});
  }
  if(action === 'deleteLog'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var r2 = deleteOneLogRow(body.id);
    return jsonOut(r2);
  }
  if(action === 'deleteLogBatch'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var idsToDelete = body.ids || [];
    for(var di=0; di<idsToDelete.length; di++){ deleteOneLogRow(idsToDelete[di]); }
    return jsonOut({ok:true, count: idsToDelete.length});
  }

  // ---------------- manual entry requests (admin decision) ----------------
  if(action === 'getRequests'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var rsh3 = getSheet('Requests', ['id','username','employee','employeeCode','project','type','reason','status','createdAt','createdTime','requestedDate','requestedTime']);
    var rows8 = readRows(rsh3).map(function(r){ delete r._row; return r; })
      .sort(function(a,b){ return String(b.createdAt+b.createdTime).localeCompare(String(a.createdAt+a.createdTime)); });
    return jsonOut({ok:true, requests: rows8});
  }
  if(action === 'deleteRequests'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var rshDel = getSheet('Requests', ['id','username','employee','employeeCode','project','type','reason','status','createdAt','createdTime','requestedDate','requestedTime']);
    var idsDel = body.ids || [];
    var toDel = readRows(rshDel).filter(function(r){ return idsDel.indexOf(r.id) !== -1; });
    toDel.sort(function(a,b){ return b._row - a._row; }); // bottom-up so row numbers don't shift mid-delete
    toDel.forEach(function(r){ rshDel.deleteRow(r._row); });
    return jsonOut({ok:true, deleted: toDel.length});
  }
  if(action === 'approveRequest' || action === 'rejectRequest'){
    var rsh4 = getSheet('Requests', ['id','username','employee','employeeCode','project','type','reason','status','createdAt','createdTime','requestedDate','requestedTime']);
    var rrows = readRows(rsh4);
    var rrow = rrows.filter(function(r){ return r.id === body.id; })[0];
    if(!rrow) return jsonOut({ok:false, error:'not found'});
    var newStatus = (action === 'approveRequest') ? 'approved' : 'rejected';
    esh_set(rsh4, rrow._row, 'status', newStatus);

    if(action === 'approveRequest'){
      if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
      var t3 = serverNow();
      var useDate = rrow.requestedDate || t3.jalaliDate;
      var useTime = rrow.requestedTime || t3.time;
      var lsh7 = getSheet('Logs', ['id','jalaliDate','time','employee','username','project','type','lat','lng','accuracy','isoTimestamp','note','employeeCode','isHoliday','holidayReason','source']);
      var iso3 = rrow.requestedDate ? (jalaliDateTimeToIso(useDate, useTime) || t3.isoTimestamp) : t3.isoTimestamp;
      var empRowsChk2 = readRows(lsh7).filter(function(r){ return r.username === rrow.username; });
      var lastChk2 = empRowsChk2.length ? empRowsChk2.sort(function(a,b){ return isoCompare(a.isoTimestamp, b.isoTimestamp); })[empRowsChk2.length-1] : null;
      if(lastChk2 && String(iso3) <= String(lastChk2.isoTimestamp)){
        esh_set(rsh4, rrow._row, 'status', 'pending'); // revert — do not silently approve an invalid entry
        return jsonOut({ok:false, error:'must_be_after_last_entry'});
      }
      if(rrow.type === 'out' && wouldBeOrphanCheckout(empRowsChk2, iso3)){
        esh_set(rsh4, rrow._row, 'status', 'pending'); // revert — same as above, don't leave this approved
        return jsonOut({ok:false, error:'no_open_checkin'});
      }
      var lid = Utilities.getUuid();
      var holInfo3 = getHolidayInfo(useDate);
      appendRowByHeader(lsh7, {
        id:lid, jalaliDate:useDate, time:useTime, employee:rrow.employee, username:rrow.username, project:rrow.project,
        type: rrow.type === 'in' ? 'ورود' : 'خروج', lat:'', lng:'', accuracy:'', isoTimestamp:iso3,
        note: 'ثبت دستی (تاییدشده) — دلیل: ' + (rrow.reason || '-') + (rrow.requestedDate ? ' — زمان درخواستی توسط پرسنل' : ''),
        employeeCode: rrow.employeeCode || '',
        isHoliday: holInfo3.isHoliday ? 'تعطیل' : '', holidayReason: holInfo3.reason, source:'دستی'
      });
    }
    return jsonOut({ok:true});
  }

  if(action === 'getLeaveRequests'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var lrsh3 = getSheet('LeaveRequests', ['id','username','employee','employeeCode','startDate','endDate','reason','status','createdAt','createdTime','leaveType','startTime','endTime','requestType','leaveCategory','project']);
    var lrows2 = readRows(lrsh3).map(function(r){ delete r._row; return r; })
      .sort(function(a,b){ return String(b.createdAt+b.createdTime).localeCompare(String(a.createdAt+a.createdTime)); });
    return jsonOut({ok:true, leaves: lrows2});
  }
  if(action === 'getLeaveRequestsForUser'){
    var lrshU = getSheet('LeaveRequests', ['id','username','employee','employeeCode','startDate','endDate','reason','status','createdAt','createdTime','leaveType','startTime','endTime','requestType','leaveCategory','project']);
    var lrowsU = readRows(lrshU).filter(function(r){ return r.username === body.username; }).map(function(r){ delete r._row; return r; });
    return jsonOut({ok:true, leaves: lrowsU});
  }
  if(action === 'deleteLeaveRequests'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    var lrshDel = getSheet('LeaveRequests', ['id','username','employee','employeeCode','startDate','endDate','reason','status','createdAt','createdTime','leaveType','startTime','endTime','requestType','leaveCategory','project']);
    var idsDel2 = body.ids || [];
    var toDel2 = readRows(lrshDel).filter(function(r){ return idsDel2.indexOf(r.id) !== -1; });
    var datesToCleanup = [];
    var usernamesToCleanup = {};
    toDel2.forEach(function(r){
      if(r.status === 'approved'){
        iterateJalaliDaysServer(r.startDate, r.endDate).forEach(function(dt){ datesToCleanup.push(dt); });
        usernamesToCleanup[r.username] = true;
      }
    });
    toDel2.sort(function(a,b){ return b._row - a._row; });
    toDel2.forEach(function(r){ lrshDel.deleteRow(r._row); });
    if(datesToCleanup.length){
      var dsmshCleanup = getSheet('DayStatusMarkers', ['id','username','employee','employeeCode','jalaliDate','status','leaveHours','createdAt','createdTime']);
      var markerRows = readRows(dsmshCleanup).filter(function(r){
        return usernamesToCleanup[r.username] && (r.status === 'مرخصی' || r.status === 'ماموریت') && datesToCleanup.indexOf(r.jalaliDate) !== -1;
      });
      markerRows.sort(function(a,b){ return b._row - a._row; });
      markerRows.forEach(function(r){ dsmshCleanup.deleteRow(r._row); });
    }
    return jsonOut({ok:true, deleted: toDel2.length});
  }
  if(action === 'approveLeaveRequest' || action === 'rejectLeaveRequest'){
    var lrsh4 = getSheet('LeaveRequests', ['id','username','employee','employeeCode','startDate','endDate','reason','status','createdAt','createdTime','leaveType','startTime','endTime','requestType','leaveCategory','project']);
    var lrow2 = readRows(lrsh4).filter(function(r){ return r.id === body.id; })[0];
    if(!lrow2) return jsonOut({ok:false, error:'not found'});
    esh_set(lrsh4, lrow2._row, 'status', (action === 'approveLeaveRequest') ? 'approved' : 'rejected');
    if(action === 'approveLeaveRequest' && (lrow2.requestType === 'leave' || lrow2.requestType === 'mission' || !lrow2.requestType)){
      syncLeaveDayStatus(lrow2.username, lrow2.employee, lrow2.startDate, lrow2.endDate, lrow2.leaveType, lrow2.startTime, lrow2.endTime, lrow2.requestType, lrow2.employeeCode);
    }
    return jsonOut({ok:true});
  }

  if(action === 'setPin'){
    // Was completely unauthenticated before this fix — anyone who knew the API URL could call
    // this directly and take over the admin account without ever knowing the current
    // password. Now requires the current PIN, exactly like every other admin action.
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    // Both fields are now required together (not independently optional) — asked for
    // explicitly: changing admin credentials should always set a genuinely new username AND
    // password at once, never leave one silently unchanged.
    if(!body.newUsername || !body.newPin) return jsonOut({ok:false, error:'both_fields_required'});
    var ssh = getSheet('Settings', ['key','value']);
    var row8 = readRows(ssh).filter(function(r){ return r.key === 'admin_pin'; })[0];
    if(row8){ ssh.getRange(row8._row, 2).setValue(body.newPin); }
    else{ appendRowByHeader(ssh, {key:'admin_pin', value:body.newPin}); }
    setSetting('admin_username', body.newUsername);
    return jsonOut({ok:true});
  }
  if(action === 'getAdminUsername'){
    return jsonOut({ok:true, username: getAdminUsername()});
  }
  if(action === 'setNotifyEmail'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    setSetting('notify_email', body.email || '');
    return jsonOut({ok:true});
  }
  if(action === 'getNotifyEmail'){
    if(String(body.pin) !== getAdminPin()) return jsonOut({ok:false, error:'invalid pin'});
    return jsonOut({ok:true, email: getSetting('notify_email'), lastStatus: getSetting('last_email_status')});
  }

  return jsonOut({ok:false, error:'unknown action'});
}

// small helper: set a single named-column cell for a given row
function esh_set(sheet, rowNum, headerName, value){
  sheet.getRange(rowNum, colIndex(sheet, headerName)).setValue(value);
}
// Forces any already-logged-in phone for this employee back to the login screen soon (see
// checkSessionVersion) — called after any admin action that changes what that employee's
// active session should be trusting (edits, activation toggles, device resets).
function esh_bumpSessionVersion(sheet, rowNum){
  esh_set(sheet, rowNum, 'sessionVersion', String(Date.now()));
}
