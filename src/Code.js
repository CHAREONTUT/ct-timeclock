/**
 * Code.js — ฝั่งเซิร์ฟเวอร์ (Google Apps Script)
 * ชีต: Employees, Holidays, Logs, EmpOffDays, Settings
 */

var SHEETS = { EMP: 'Employees', HOL: 'Holidays', LOG: 'Logs', OFF: 'EmpOffDays', SET: 'Settings' };
var MAX_PIN_TRIES = 5;

/* ---------- เว็บแอป ---------- */

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('ลงเวลาทำงาน')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- เมนู ---------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('⏱ ระบบลงเวลา')
    .addItem('สรุปเดือนนี้', 'menuSummaryThisMonth')
    .addItem('สรุปเดือนที่แล้ว', 'menuSummaryLastMonth')
    .addItem('สรุปเดือนที่ระบุ…', 'menuSummaryPick')
    .addSeparator()
    .addItem('ตั้งค่าเริ่มต้น (ครั้งแรก)', 'setup')
    .addToUi();
}

function setup() {
  var ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone('Asia/Bangkok');
  ensureSheet_(ss, SHEETS.EMP, ['EmpID', 'ชื่อ', 'PIN', 'วันหยุดประจำสัปดาห์', 'ใช้งาน'], [
    ['E01', 'พนักงาน 1', '1111', '', true],
    ['E02', 'พนักงาน 2', '2222', '', true],
    ['E03', 'พนักงาน 3', '3333', '', true]
  ]);
  ensureSheet_(ss, SHEETS.HOL, ['วันที่', 'ชื่อวันหยุด'], []);
  ensureSheet_(ss, SHEETS.LOG, ['เวลา', 'EmpID', 'ชื่อ', 'ประเภท', 'หมายเหตุ', 'บันทึกเมื่อ', 'Lat', 'Lng'], []);
  ensureSheet_(ss, SHEETS.OFF, ['EmpID', 'เดือน', 'วันที่ลาหยุด', 'บันทึกเมื่อ'], []);
  ensureSheet_(ss, SHEETS.SET, ['Key', 'Value', 'คำอธิบาย'], [
    ['STANDARD_HOURS', 9, 'ชั่วโมงทำงานปกติต่อวัน'],
    ['OT_ON_WORKDAYS', true, 'TRUE = ทำเกินชั่วโมงปกติในวันทำงาน นับเป็น OT ×1.5'],
    ['MIN_OT_MINUTES', 0, 'OT ต่อวันที่น้อยกว่านี้ (นาที) ไม่นับ'],
    ['HR_PIN', '0000', 'PIN 4 หลักสำหรับหน้า HR (เปลี่ยนด่วน!)']
  ]);
  ss.getSheetByName(SHEETS.LOG).getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm');
  ss.getSheetByName(SHEETS.LOG).getRange('F:F').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  ss.getSheetByName(SHEETS.HOL).getRange('A:A').setNumberFormat('yyyy-mm-dd');
  ss.getSheetByName(SHEETS.EMP).getRange('C:C').setNumberFormat('@');
  SpreadsheetApp.getUi().alert(
    'ตั้งค่าเรียบร้อย\n\n' +
    '1. แก้ชื่อ/PIN พนักงานในชีต Employees\n' +
    '2. เปลี่ยน HR_PIN ในชีต Settings\n' +
    '3. ใส่วันหยุดบริษัทในชีต Holidays'
  );
}

function ensureSheet_(ss, name, header, sample) {
  var sh = ss.getSheetByName(name);
  if (sh) return sh;
  sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  sh.setFrozenRows(1);
  if (sample.length) sh.getRange(2, 1, sample.length, header.length).setValues(sample);
  return sh;
}

/* ---------- อ่านข้อมูล ---------- */

function readEmployees_() {
  var rows = SpreadsheetApp.getActive().getSheetByName(SHEETS.EMP).getDataRange().getValues().slice(1);
  var map = {};
  rows.forEach(function (r) {
    var id = String(r[0]).trim();
    if (!id) return;
    map[id] = {
      id: id, name: String(r[1]).trim(), pin: String(r[2]).trim(),
      offDays: parseOffDays_(r[3]),
      active: r[4] === true || String(r[4]).toUpperCase() === 'TRUE'
    };
  });
  return map;
}

function readHolidays_() {
  var rows = SpreadsheetApp.getActive().getSheetByName(SHEETS.HOL).getDataRange().getValues().slice(1);
  var map = {};
  rows.forEach(function (r) {
    if (r[0] instanceof Date) map[dayKey_(r[0])] = String(r[1] || 'วันหยุด');
    else if (/^\d{4}-\d{2}-\d{2}$/.test(String(r[0]).trim())) map[String(r[0]).trim()] = String(r[1] || 'วันหยุด');
  });
  return map;
}

function readEmpOffDays_(month) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.OFF);
  if (!sh || sh.getLastRow() < 2) return {};
  var rows = sh.getDataRange().getValues().slice(1);
  var result = {};
  rows.forEach(function (r) {
    var empId = String(r[0]).trim(), m = String(r[1]).trim();
    if (!empId || m !== month) return;
    result[empId] = String(r[2]).split(',').map(function (d) { return d.trim(); })
      .filter(function (d) { return /^\d{4}-\d{2}-\d{2}$/.test(d); });
  });
  return result;
}

function readSettings_() {
  var rows = SpreadsheetApp.getActive().getSheetByName(SHEETS.SET).getDataRange().getValues().slice(1);
  var s = { standardHours: 9, otOnWorkdays: true, minOtMinutes: 0, hrPin: '0000' };
  rows.forEach(function (r) {
    var k = String(r[0]).trim(), v = r[1];
    if (k === 'STANDARD_HOURS') s.standardHours = Number(v) || 9;
    if (k === 'OT_ON_WORKDAYS') s.otOnWorkdays = !(v === false || String(v).toUpperCase() === 'FALSE');
    if (k === 'MIN_OT_MINUTES') s.minOtMinutes = Number(v) || 0;
    if (k === 'HR_PIN') s.hrPin = String(v).trim();
  });
  return s;
}

function readLogs_() {
  var rows = SpreadsheetApp.getActive().getSheetByName(SHEETS.LOG).getDataRange().getValues().slice(1);
  return rows.filter(function (r) { return r[0] instanceof Date && r[1]; }).map(function (r) {
    return {
      time: r[0], empId: String(r[1]).trim(),
      type: String(r[3]).trim().toUpperCase(),
      lat: r[6] || null, lng: r[7] || null
    };
  });
}

function computeDays_(monthPrefix) {
  var emps = readEmployees_();
  var settings = readSettings_();
  var empOffDays = monthPrefix ? readEmpOffDays_(monthPrefix) : {};
  Object.keys(emps).forEach(function (id) { emps[id].monthlyOffDays = empOffDays[id] || []; });

  var days = summarizeDays_(buildSessions_(readLogs_()), {
    employees: emps, holidays: readHolidays_(),
    standardHours: settings.standardHours, otOnWorkdays: settings.otOnWorkdays
  });
  if (settings.minOtMinutes > 0) {
    var min = settings.minOtMinutes / 60;
    days.forEach(function (d) {
      if (d.ot15 > 0 && d.ot15 < min) d.ot15 = 0;
      if (d.ot3 > 0 && d.ot3 < min) d.ot3 = 0;
    });
  }
  if (monthPrefix) days = days.filter(function (d) { return d.date.indexOf(monthPrefix) === 0; });
  return { days: days, emps: emps };
}

/* ---------- Auth ---------- */

function auth_(empId, pin) {
  var emps = readEmployees_();
  var emp = emps[String(empId)];
  if (!emp || !emp.active) throw new Error('ไม่พบพนักงาน');
  var cache = CacheService.getScriptCache();
  var key = 'tries_' + emp.id;
  var tries = Number(cache.get(key) || 0);
  if (tries >= MAX_PIN_TRIES) throw new Error('ใส่ PIN ผิดหลายครั้ง กรุณารอ 15 นาที');
  if (String(pin) !== emp.pin) {
    cache.put(key, String(tries + 1), 900);
    throw new Error('PIN ไม่ถูกต้อง (' + (MAX_PIN_TRIES - tries - 1) + ' ครั้งที่เหลือ)');
  }
  cache.remove(key);
  return emp;
}

function authHr_(pin) {
  var settings = readSettings_();
  var cache = CacheService.getScriptCache();
  var key = 'hr_tries';
  var tries = Number(cache.get(key) || 0);
  if (tries >= MAX_PIN_TRIES) throw new Error('ใส่ PIN ผิดหลายครั้ง กรุณารอ 15 นาที');
  if (String(pin) !== settings.hrPin) {
    cache.put(key, String(tries + 1), 900);
    throw new Error('PIN HR ไม่ถูกต้อง');
  }
  cache.remove(key);
  return true;
}

/* ---------- API ---------- */

function apiEmployees() {
  var emps = readEmployees_();
  return Object.keys(emps).filter(function (k) { return emps[k].active; })
    .map(function (k) { return { id: emps[k].id, name: emps[k].name }; });
}

function apiStatus(empId, pin) {
  var emp = auth_(empId, pin);
  var now = new Date();
  var today = dayKey_(now);
  var month = today.slice(0, 7);
  var mine = readLogs_().filter(function (l) { return l.empId === emp.id; });
  var sessions = buildSessions_(mine);
  var last = sessions[sessions.length - 1];
  var openSince = last && last.open ? last.start : null;

  var holidays = readHolidays_();
  var empOffDays = readEmpOffDays_(month);
  var myOffDays = empOffDays[emp.id] || [];
  var todayHoliday = holidays[today] || (myOffDays.indexOf(today) >= 0 ? 'วันหยุดของคุณ' : '');

  var settings = readSettings_();
  var days = computeDays_(month).days.filter(function (d) { return d.empId === emp.id; });
  var total = summarizeEmployees_(days)[0] || null;
  return {
    name: emp.name, now: now.toISOString(),
    openSince: openSince ? openSince.toISOString() : null,
    todayHoliday: todayHoliday, standardHours: settings.standardHours,
    month: month, total: total, days: days.slice().reverse(),
    myOffDays: myOffDays,
    offDaysSet: myOffDays.length > 0
  };
}

function apiPunch(empId, pin, type, note, lat, lng) {
  var emp = auth_(empId, pin);
  type = String(type).toUpperCase();
  if (type !== 'IN' && type !== 'OUT') throw new Error('ประเภทไม่ถูกต้อง');

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var mine = readLogs_().filter(function (l) { return l.empId === emp.id; })
      .sort(function (a, b) { return a.time - b.time; });
    var lastType = mine.length ? mine[mine.length - 1].type : 'OUT';
    if (type === 'IN' && lastType === 'IN') throw new Error('คุณกดเข้างานไว้แล้ว');
    if (type === 'OUT' && lastType !== 'IN') throw new Error('ยังไม่ได้กดเข้างาน');
    var now = new Date();
    var latVal = (lat !== null && lat !== undefined && lat !== '') ? Number(lat) : '';
    var lngVal = (lng !== null && lng !== undefined && lng !== '') ? Number(lng) : '';
    SpreadsheetApp.getActive().getSheetByName(SHEETS.LOG)
      .appendRow([now, emp.id, emp.name, type, String(note || '').slice(0, 200), now, latVal, lngVal]);
  } finally { lock.releaseLock(); }
  return apiStatus(empId, pin);
}

function apiGetMonthlyOffDays(empId, pin, month) {
  auth_(empId, pin);
  if (!/^\d{4}-\d{2}$/.test(month)) throw new Error('รูปแบบเดือนไม่ถูกต้อง');
  var offDays = readEmpOffDays_(month);
  var holidays = readHolidays_();
  var monthHols = {};
  Object.keys(holidays).forEach(function (d) { if (d.indexOf(month) === 0) monthHols[d] = holidays[d]; });
  return { month: month, dates: offDays[empId] || [], companyHolidays: monthHols };
}

function apiSetMonthlyOffDays(empId, pin, month, dates) {
  auth_(empId, pin);
  if (!/^\d{4}-\d{2}$/.test(month)) throw new Error('รูปแบบเดือนไม่ถูกต้อง');
  var valid = (dates || []).filter(function (d) {
    return /^\d{4}-\d{2}-\d{2}$/.test(d) && d.indexOf(month) === 0;
  });
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.OFF);
  var rows = sh.getLastRow() > 1 ? sh.getDataRange().getValues() : [[]];
  var found = -1;
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(empId) && String(rows[i][1]).trim() === month) {
      found = i + 1; break;
    }
  }
  var now = new Date();
  if (found > 0) {
    sh.getRange(found, 3, 1, 2).setValues([[valid.join(','), now]]);
  } else {
    sh.appendRow([empId, month, valid.join(','), now]);
  }
  return { ok: true, month: month, dates: valid };
}

function apiHrLogin(pin) {
  authHr_(pin);
  return { ok: true };
}

function apiHrData(pin, month) {
  authHr_(pin);
  if (!month || !/^\d{4}-\d{2}$/.test(month)) month = dayKey_(new Date()).slice(0, 7);
  var r = computeDays_(month);
  var nameOf = function (id) { return (r.emps[id] && r.emps[id].name) || id; };

  var summary = summarizeEmployees_(r.days).map(function (e) {
    return {
      empId: e.empId, name: nameOf(e.empId),
      daysWorked: e.daysWorked, hours: e.hours, regular: e.regular,
      ot15: e.ot15, holidayDays: e.holidayDays, holidayWork: e.holidayWork,
      ot3: e.ot3, problems: e.problems
    };
  });

  // เติมพนักงานที่ยังไม่มี record ในเดือนนี้
  var seen = {};
  summary.forEach(function (e) { seen[e.empId] = true; });
  Object.keys(r.emps).forEach(function (id) {
    if (r.emps[id].active && !seen[id]) {
      summary.push({
        empId: id, name: nameOf(id),
        daysWorked: 0, hours: 0, regular: 0, ot15: 0,
        holidayDays: 0, holidayWork: 0, ot3: 0, problems: 0
      });
    }
  });

  var byEmp = {};
  r.days.forEach(function (d) { (byEmp[d.empId] = byEmp[d.empId] || []).push(d); });

  return { month: month, summary: summary, detail: byEmp };
}

/* ---------- รายงานรายเดือน ---------- */

function menuSummaryThisMonth() { buildMonthlyReport(dayKey_(new Date()).slice(0, 7)); }

function menuSummaryLastMonth() {
  var d = new Date(dayKey_(new Date()).slice(0, 7) + '-01T00:00:00Z');
  d.setUTCMonth(d.getUTCMonth() - 1);
  buildMonthlyReport(d.toISOString().slice(0, 7));
}

function menuSummaryPick() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('สรุปรายเดือน', 'พิมพ์เดือน เช่น 2026-10', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var m = res.getResponseText().trim();
  if (!/^\d{4}-\d{2}$/.test(m)) { ui.alert('รูปแบบต้องเป็น YYYY-MM'); return; }
  buildMonthlyReport(m);
}

function buildMonthlyReport(month) {
  var ss = SpreadsheetApp.getActive();
  var r = computeDays_(month);
  var nameOf = function (id) { return (r.emps[id] && r.emps[id].name) || id; };
  var title = 'สรุป ' + month;
  var sh = ss.getSheetByName(title) || ss.insertSheet(title);
  sh.clear();

  var head1 = ['EmpID', 'ชื่อ', 'วันที่มาทำงาน', 'ชม.ทำงานรวม', 'ชม.ปกติ', 'OT วันธรรมดา (×1.5)',
    'วันหยุดที่มาทำ', 'ชม.ทำวันหยุด (×1)', 'OT วันหยุด (×3)', 'ต้องตรวจ'];
  var sum = summarizeEmployees_(r.days).map(function (e) {
    return [e.empId, nameOf(e.empId), e.daysWorked, e.hours, e.regular, e.ot15,
      e.holidayDays, e.holidayWork, e.ot3, e.problems];
  });
  sh.getRange(1, 1).setValue('สรุปเดือน ' + month + ' (สร้างเมื่อ ' +
    Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm') + ')').setFontWeight('bold');
  sh.getRange(2, 1, 1, head1.length).setValues([head1]).setFontWeight('bold').setBackground('#e8f0fe');
  if (sum.length) sh.getRange(3, 1, sum.length, head1.length).setValues(sum);

  var start = 3 + Math.max(sum.length, 1) + 2;
  var head2 = ['วันที่', 'EmpID', 'ชื่อ', 'ประเภทวัน', 'ชม.รวม', 'ชม.ปกติ', 'OT วันธรรมดา',
    'ชม.ทำวันหยุด', 'OT วันหยุด', 'ต้องตรวจ'];
  sh.getRange(start - 1, 1).setValue('รายละเอียดรายวัน').setFontWeight('bold');
  sh.getRange(start, 1, 1, head2.length).setValues([head2]).setFontWeight('bold').setBackground('#e8f0fe');
  var det = r.days.map(function (d) {
    return [d.date, d.empId, nameOf(d.empId),
      d.dayType === 'HOLIDAY' ? 'วันหยุด: ' + d.holidayName : 'วันธรรมดา',
      d.hours, d.regular, d.ot15, d.holidayWork, d.ot3, d.problems.join(', ')];
  });
  if (det.length) {
    sh.getRange(start + 1, 1, det.length, head2.length).setValues(det);
    det.forEach(function (row, i) {
      if (row[9]) sh.getRange(start + 1 + i, 1, 1, head2.length).setBackground('#fde7e9');
      else if (row[3] !== 'วันธรรมดา') sh.getRange(start + 1 + i, 1, 1, head2.length).setBackground('#fff4e5');
    });
  }
  sh.autoResizeColumns(1, head2.length);
  ss.setActiveSheet(sh);
}
