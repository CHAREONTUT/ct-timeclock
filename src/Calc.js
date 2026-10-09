/**
 * Calc.js — ตรรกะคำนวณชั่วโมงทำงาน/OT (ไม่พึ่ง Google API จึงทดสอบด้วย Node ได้)
 *
 * - ทำงานไม่เป็นเวลา ไม่คิดสาย นับเฉพาะชั่วโมงที่ทำจริง
 * - กะที่ลากข้ามเที่ยงคืน นับเป็นของ "วันที่เข้างาน"
 * - วันหยุด = ชีต Holidays (บริษัท) | วันหยุดประจำสัปดาห์รายคน | วันหยุดรายเดือนที่ตั้งเอง
 * - วันทำงาน : ≤ std = ปกติ, > std = OT ×1.5
 * - วันหยุด  : ≤ std = ทำวันหยุด, > std = OT ×3
 */

var TZ_OFFSET_MS = 7 * 60 * 60 * 1000;
var DAY_CODES = ['SUN', 'MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT'];
var THAI_DAY_ALIASES = {
  'อา': 'SUN', 'จ': 'MON', 'อ': 'TUE', 'พ': 'WED', 'พฤ': 'THU', 'ศ': 'FRI', 'ส': 'SAT'
};

function dayKey_(date) {
  return new Date(date.getTime() + TZ_OFFSET_MS).toISOString().slice(0, 10);
}

function weekdayCode_(key) {
  return DAY_CODES[new Date(key + 'T00:00:00Z').getUTCDay()];
}

function parseOffDays_(text) {
  if (!text) return [];
  return String(text).split(/[,\s/]+/).map(function (s) {
    s = s.trim().replace(/\.$/, '');
    if (!s) return null;
    if (THAI_DAY_ALIASES[s]) return THAI_DAY_ALIASES[s];
    var up = s.toUpperCase().slice(0, 3);
    return DAY_CODES.indexOf(up) >= 0 ? up : null;
  }).filter(Boolean);
}

function round2_(n) { return Math.round(n * 100) / 100; }

/**
 * จับคู่ IN/OUT เป็น sessions
 * logs: [{empId, time: Date, type: 'IN'|'OUT'}]
 */
function buildSessions_(logs) {
  var byEmp = {};
  logs.forEach(function (l) {
    (byEmp[l.empId] = byEmp[l.empId] || []).push(l);
  });
  var sessions = [];
  Object.keys(byEmp).forEach(function (emp) {
    var list = byEmp[emp].slice().sort(function (a, b) { return a.time - b.time; });
    var open = null;
    list.forEach(function (l) {
      if (l.type === 'IN') {
        if (open) sessions.push({ empId: emp, start: open.time, end: null, problem: 'ลืมกดเลิกงาน' });
        open = l;
      } else if (l.type === 'OUT') {
        if (open) {
          sessions.push({ empId: emp, start: open.time, end: l.time, problem: null });
          open = null;
        } else {
          sessions.push({ empId: emp, start: null, end: l.time, problem: 'กดเลิกงานโดยไม่ได้กดเข้างาน' });
        }
      }
    });
    if (open) sessions.push({ empId: emp, start: open.time, end: null, problem: null, open: true });
  });
  return sessions;
}

/**
 * สรุปรายวัน
 * opts.employees[id].monthlyOffDays = ['YYYY-MM-DD', ...] (วันหยุดส่วนตัวของเดือนนั้น)
 * opts.employees[id].offDays = ['MON', ...] (วันหยุดประจำสัปดาห์ ยังรองรับไว้)
 */
function summarizeDays_(sessions, opts) {
  var std = opts.standardHours || 9;
  var days = {};
  sessions.forEach(function (s) {
    var anchor = s.start || s.end;
    var key = dayKey_(anchor);
    var id = s.empId + '|' + key;
    var d = days[id] || (days[id] = { empId: s.empId, date: key, hours: 0, problems: [], sessions: [], working: false });
    d.sessions.push({ start: s.start ? s.start.toISOString() : null, end: s.end ? s.end.toISOString() : null });
    if (s.open) d.working = true;
    if (s.start && s.end) d.hours += (s.end - s.start) / 3600000;
    if (s.problem) d.problems.push(s.problem);
  });

  return Object.keys(days).sort().map(function (id) {
    var d = days[id];
    var emp = opts.employees[d.empId] || { offDays: [], monthlyOffDays: [] };
    var globalHol = opts.holidays[d.date];
    var isWeeklyOff = (emp.offDays || []).indexOf(weekdayCode_(d.date)) >= 0;
    var isPersonalOff = (emp.monthlyOffDays || []).indexOf(d.date) >= 0;
    var isHoliday = !!globalHol || isWeeklyOff || isPersonalOff;
    var holName = globalHol || (isPersonalOff ? 'วันหยุดส่วนตัว' : (isWeeklyOff ? 'วันหยุดประจำสัปดาห์' : ''));
    var h = d.hours;
    var r = {
      empId: d.empId, date: d.date, hours: round2_(h),
      dayType: isHoliday ? 'HOLIDAY' : 'WORK',
      holidayName: holName,
      regular: 0, ot15: 0, holidayWork: 0, ot3: 0,
      problems: d.problems, sessions: d.sessions, working: d.working
    };
    if (isHoliday) {
      r.holidayWork = round2_(Math.min(h, std));
      r.ot3 = round2_(Math.max(0, h - std));
    } else if (opts.otOnWorkdays === false) {
      r.regular = round2_(h);
    } else {
      r.regular = round2_(Math.min(h, std));
      r.ot15 = round2_(Math.max(0, h - std));
    }
    return r;
  });
}

function summarizeEmployees_(dayRows) {
  var out = {};
  dayRows.forEach(function (d) {
    var e = out[d.empId] || (out[d.empId] = {
      empId: d.empId, daysWorked: 0, hours: 0, regular: 0, ot15: 0,
      holidayDays: 0, holidayWork: 0, ot3: 0, problems: 0
    });
    if (d.hours > 0) {
      e.daysWorked++;
      if (d.dayType === 'HOLIDAY') e.holidayDays++;
    }
    e.hours += d.hours; e.regular += d.regular; e.ot15 += d.ot15;
    e.holidayWork += d.holidayWork; e.ot3 += d.ot3;
    e.problems += d.problems.length;
  });
  return Object.keys(out).map(function (k) {
    var e = out[k];
    ['hours', 'regular', 'ot15', 'holidayWork', 'ot3'].forEach(function (f) { e[f] = round2_(e[f]); });
    return e;
  });
}

if (typeof module !== 'undefined') {
  module.exports = {
    dayKey_: dayKey_, weekdayCode_: weekdayCode_, parseOffDays_: parseOffDays_,
    buildSessions_: buildSessions_, summarizeDays_: summarizeDays_,
    summarizeEmployees_: summarizeEmployees_
  };
}
