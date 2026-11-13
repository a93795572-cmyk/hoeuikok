/**
 * 회의콕 - 등록된 일정으로 회의 시간과 참석자를 자동 추천하는 Apps Script 앱
 *
 * 시트 구성 (setupSheets()가 자동 생성)
 *  - 직원        : 사번, 이름, 부서, 직급, 담당, 이메일, 출근, 퇴근
 *  - 일정        : 일정ID, 사번, 날짜, 시작, 종료, 유형, 메모, 회의ID
 *  - 회의        : 회의ID, 안건유형, 제목, 주관자사번, 날짜, 시작, 종료, 회의실, 참석자사번, 상태, 등록시각
 *  - 안건템플릿  : 안건유형, 부서, 최소직급, 인원
 *  - 회의실      : 회의실, 수용인원
 *  - 설정        : 항목, 값  (직급 순서, 탐색 간격 등)
 *
 * 날짜는 "yyyy-MM-dd", 시간은 "HH:mm" 문자열로 저장한다(시트 서식: 일반 텍스트).
 */

// ===================== 기본값 =====================
var SHEETS = {
  EMP: '직원',
  SCH: '일정',
  MTG: '회의',
  TPL: '안건템플릿',
  ROOM: '회의실',
  CFG: '설정'
};

var HEADERS = {
  '직원': ['사번', '이름', '부서', '직급', '담당', '이메일', '출근', '퇴근'],
  '일정': ['일정ID', '사번', '날짜', '시작', '종료', '유형', '메모', '회의ID'],
  '회의': ['회의ID', '안건유형', '제목', '주관자사번', '날짜', '시작', '종료', '회의실', '참석자사번', '상태', '등록시각'],
  '안건템플릿': ['안건유형', '부서', '최소직급', '인원'],
  '회의실': ['회의실', '수용인원'],
  '설정': ['항목', '값']
};

var DEFAULT_RANKS = ['사원', '주임', '대리', '과장', '차장', '부장'];
var SCHEDULE_TYPES = ['개인업무', '외근', '교육', '휴무', '회의', '기타'];

// ===================== 메뉴 / 진입점 =====================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('회의콕')
    .addItem('앱 열기', 'openApp')
    .addSeparator()
    .addItem('시트 처음 만들기 (예시 데이터 포함)', 'setupSheetsWithSample')
    .addItem('시트 처음 만들기 (빈 시트)', 'setupSheetsEmpty')
    .addToUi();
}

function openApp() {
  var html = HtmlService.createTemplateFromFile('index').evaluate()
    .setWidth(1280).setHeight(820);
  SpreadsheetApp.getUi().showModalDialog(html, '회의콕');
}

/** 웹 앱으로 배포했을 때 진입점 */
function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('회의콕')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ===================== 시트 준비 =====================
function setupSheetsWithSample() { setupSheets_(true); }
function setupSheetsEmpty() { setupSheets_(false); }

function setupSheets_(withSample) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HEADERS).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      var h = HEADERS[name];
      sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold')
        .setBackground('#002C5F').setFontColor('#FFFFFF');
      sh.setFrozenRows(1);
      sh.getRange(1, 1, sh.getMaxRows(), h.length).setNumberFormat('@'); // 일반 텍스트
    }
  });
  if (withSample) writeSample_(ss);
  return '시트 준비 완료';
}

function writeSample_(ss) {
  var emp = ss.getSheetByName(SHEETS.EMP);
  if (emp.getLastRow() <= 1) {
    var rows = [
      ['E001', '김서진', '생산관리', '사원', '의장 1라인 생산계획', 'seojin@example.com', '08:00', '17:00'],
      ['E002', '박도현', '생산관리', '과장', '생산계획 총괄', '', '08:00', '17:00'],
      ['E003', '이하은', '생산관리', '대리', '차체 라인', '', '08:00', '17:00'],
      ['E004', '최민준', '구매', '대리', '전장 부품 조달', '', '08:30', '17:30'],
      ['E005', '정유나', '구매', '과장', '구매 파트장', '', '08:30', '17:30'],
      ['E006', '강태윤', '품질', '대리', '입고 품질', '', '08:00', '17:00'],
      ['E007', '윤서아', '품질', '차장', '품질 팀장', '', '08:00', '17:00'],
      ['E008', '한지호', '물류', '주임', '부품 입고 물류', '', '07:30', '16:30'],
      ['E009', '오민서', '물류', '과장', '물류 파트장', '', '07:30', '16:30'],
      ['E010', '서준우', '생산기술', '대리', '설비 보전', '', '08:00', '17:00']
    ];
    emp.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
  var tpl = ss.getSheetByName(SHEETS.TPL);
  if (tpl.getLastRow() <= 1) {
    var t = [
      ['부품 공급 차질', '생산관리', '대리', '1'],
      ['부품 공급 차질', '구매', '대리', '1'],
      ['부품 공급 차질', '물류', '주임', '1'],
      ['품질 이슈', '품질', '대리', '1'],
      ['품질 이슈', '생산관리', '대리', '1'],
      ['품질 이슈', '생산기술', '대리', '1'],
      ['생산계획 변경', '생산관리', '과장', '1'],
      ['생산계획 변경', '구매', '대리', '1'],
      ['생산계획 변경', '물류', '주임', '1'],
      ['설비 고장', '생산기술', '대리', '1'],
      ['설비 고장', '생산관리', '대리', '1']
    ];
    tpl.getRange(2, 1, t.length, 4).setValues(t);
  }
  var room = ss.getSheetByName(SHEETS.ROOM);
  if (room.getLastRow() <= 1) {
    room.getRange(2, 1, 4, 2).setValues([['본관 3층 대회의실', '12'], ['본관 3층 소회의실 A', '6'], ['본관 3층 소회의실 B', '6'], ['공장동 회의실', '8']]);
  }
  var cfg = ss.getSheetByName(SHEETS.CFG);
  if (cfg.getLastRow() <= 1) {
    cfg.getRange(2, 1, 3, 2).setValues([
      ['직급순서', DEFAULT_RANKS.join(',')],
      ['탐색간격(분)', '30'],
      ['추천개수', '5']
    ]);
  }
  var sch = ss.getSheetByName(SHEETS.SCH);
  if (sch.getLastRow() <= 1) {
    var d0 = nextWorkday_(today_());
    var d1 = nextWorkday_(addDays_(d0, 1));
    var s = [
      [newId_('S'), 'E001', d0, '09:00', '11:00', '개인업무', '주간 생산계획 작성', ''],
      [newId_('S'), 'E002', d0, '10:00', '12:00', '회의', '팀 주간회의', ''],
      [newId_('S'), 'E004', d0, '13:00', '15:00', '외근', '협력사 방문', ''],
      [newId_('S'), 'E005', d0, '00:00', '23:59', '휴무', '연차', ''],
      [newId_('S'), 'E006', d0, '14:00', '16:00', '교육', '품질 교육', ''],
      [newId_('S'), 'E008', d0, '09:00', '10:30', '개인업무', '입고 검수', ''],
      [newId_('S'), 'E003', d1, '09:00', '12:00', '개인업무', '라인 순회', ''],
      [newId_('S'), 'E009', d1, '13:00', '14:00', '회의', '물류 협력사 회의', '']
    ];
    sch.getRange(2, 1, s.length, 8).setValues(s);
  }
}

// ===================== 공통 유틸 =====================
function sheet_(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error("'" + name + "' 시트가 없습니다. 메뉴 회의콕 > 시트 처음 만들기를 먼저 실행하세요.");
  return sh;
}

function readTable_(name) {
  var sh = sheet_(name);
  var last = sh.getLastRow();
  var h = HEADERS[name];
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, h.length).getDisplayValues();
  return values
    .map(function (r, i) {
      var o = { _row: i + 2 };
      h.forEach(function (k, j) { o[k] = String(r[j] || '').trim(); });
      return o;
    })
    .filter(function (o) { return o[h[0]] !== ''; });
}

function appendRow_(name, obj) {
  var h = HEADERS[name];
  sheet_(name).appendRow(h.map(function (k) { return obj[k] == null ? '' : String(obj[k]); }));
}

function updateRow_(name, row, obj) {
  var h = HEADERS[name];
  sheet_(name).getRange(row, 1, 1, h.length).setValues([h.map(function (k) { return obj[k] == null ? '' : String(obj[k]); })]);
}

function tz_() { return Session.getScriptTimeZone() || 'Asia/Seoul'; }
function today_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'); }
function nowHM_() { return Utilities.formatDate(new Date(), tz_(), 'HH:mm'); }
function newId_(p) { return p + Utilities.formatDate(new Date(), tz_(), 'yyMMddHHmmss') + Math.floor(Math.random() * 900 + 100); }

function toMin_(hm) {
  var m = /^(\d{1,2}):(\d{2})$/.exec(String(hm).trim());
  if (!m) throw new Error('시간 형식이 올바르지 않습니다: ' + hm + ' (예: 09:30)');
  return Number(m[1]) * 60 + Number(m[2]);
}
function toHM_(min) {
  var h = Math.floor(min / 60), m = min % 60;
  return (h < 10 ? '0' : '') + h + ':' + (m < 10 ? '0' : '') + m;
}
function parseDate_(s) {
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s).trim());
  if (!m) throw new Error('날짜 형식이 올바르지 않습니다: ' + s + ' (예: 2026-11-10)');
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}
function fmtDate_(d) {
  var y = d.getFullYear(), mo = d.getMonth() + 1, da = d.getDate();
  return y + '-' + (mo < 10 ? '0' : '') + mo + '-' + (da < 10 ? '0' : '') + da;
}
function addDays_(s, n) { var d = parseDate_(s); d.setDate(d.getDate() + n); return fmtDate_(d); }
function isWeekend_(s) { var w = parseDate_(s).getDay(); return w === 0 || w === 6; }
function nextWorkday_(s) { while (isWeekend_(s)) s = addDays_(s, 1); return s; }

function getConfig_() {
  var cfg = {};
  try { readTable_(SHEETS.CFG).forEach(function (r) { cfg[r['항목']] = r['값']; }); } catch (e) {}
  var ranks = (cfg['직급순서'] || DEFAULT_RANKS.join(',')).split(',').map(function (s) { return s.trim(); }).filter(String);
  return {
    ranks: ranks,
    step: Math.max(10, Number(cfg['탐색간격(분)']) || 30),
    topN: Math.max(1, Number(cfg['추천개수']) || 5)
  };
}
function rankIndex_(ranks, r) { var i = ranks.indexOf(r); return i < 0 ? -1 : i; }

// ===================== 화면 초기 데이터 =====================
function getBootstrap() {
  var cfg = getConfig_();
  var emps = readTable_(SHEETS.EMP);
  var tpl = readTable_(SHEETS.TPL);
  var rooms = readTable_(SHEETS.ROOM);
  var depts = {};
  emps.forEach(function (e) { if (e['부서']) depts[e['부서']] = true; });
  tpl.forEach(function (t) { if (t['부서']) depts[t['부서']] = true; });
  var agenda = {};
  tpl.forEach(function (t) {
    (agenda[t['안건유형']] = agenda[t['안건유형']] || []).push({ dept: t['부서'], minRank: t['최소직급'], count: Number(t['인원']) || 1 });
  });
  var me = '';
  try { me = Session.getActiveUser().getEmail() || ''; } catch (e) {}
  return {
    today: today_(),
    now: nowHM_(),
    userEmail: me,
    ranks: cfg.ranks,
    scheduleTypes: SCHEDULE_TYPES,
    employees: emps.map(empOut_),
    depts: Object.keys(depts).sort(),
    agenda: agenda,
    rooms: rooms.map(function (r) { return { name: r['회의실'], cap: Number(r['수용인원']) || 0 }; })
  };
}

function empOut_(e) {
  return { id: e['사번'], name: e['이름'], dept: e['부서'], rank: e['직급'], role: e['담당'], email: e['이메일'], start: e['출근'], end: e['퇴근'] };
}

// ===================== 직원 =====================
function saveEmployee(p) {
  if (!p || !p.name || !p.dept || !p.rank) throw new Error('이름, 부서, 직급은 꼭 입력해야 합니다.');
  var s = toMin_(p.start || '08:00'), e = toMin_(p.end || '17:00');
  if (e <= s) throw new Error('퇴근 시간은 출근 시간보다 늦어야 합니다.');
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var rows = readTable_(SHEETS.EMP);
    var obj = { '사번': p.id, '이름': p.name.trim(), '부서': p.dept.trim(), '직급': p.rank, '담당': p.role || '', '이메일': p.email || '', '출근': toHM_(s), '퇴근': toHM_(e) };
    var found = p.id ? rows.filter(function (r) { return r['사번'] === p.id; })[0] : null;
    if (found) {
      updateRow_(SHEETS.EMP, found._row, obj);
    } else {
      var max = 0;
      rows.forEach(function (r) { var n = Number(String(r['사번']).replace(/\D/g, '')); if (n > max) max = n; });
      obj['사번'] = 'E' + ('000' + (max + 1)).slice(-3);
      appendRow_(SHEETS.EMP, obj);
    }
    return getBootstrap();
  } finally { lock.releaseLock(); }
}

function deleteEmployee(id) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var r = readTable_(SHEETS.EMP).filter(function (x) { return x['사번'] === id; })[0];
    if (r) sheet_(SHEETS.EMP).deleteRow(r._row);
    return getBootstrap();
  } finally { lock.releaseLock(); }
}

// ===================== 일정 =====================
/** 한 사람의 기간 일정 (주간 보기용) */
function getSchedules(empId, from, to) {
  return readTable_(SHEETS.SCH)
    .filter(function (r) { return r['사번'] === empId && r['날짜'] >= from && r['날짜'] <= to; })
    .map(schOut_)
    .sort(function (a, b) { return (a.date + a.start).localeCompare(b.date + b.start); });
}
function schOut_(r) { return { id: r['일정ID'], emp: r['사번'], date: r['날짜'], start: r['시작'], end: r['종료'], type: r['유형'], memo: r['메모'], meetingId: r['회의ID'] }; }

function saveSchedule(p) {
  if (!p || !p.emp || !p.date || !p.type) throw new Error('사람, 날짜, 유형은 꼭 입력해야 합니다.');
  parseDate_(p.date);
  var allDay = p.type === '휴무' || p.allDay;
  var s = allDay ? 0 : toMin_(p.start), e = allDay ? 1439 : toMin_(p.end);
  if (e <= s) throw new Error('끝나는 시간은 시작 시간보다 늦어야 합니다.');
  var dates = [p.date];
  if (p.repeatUntil && p.repeatUntil > p.date) {
    var d = p.date;
    while (d < p.repeatUntil) { d = addDays_(d, 1); if (!p.skipWeekend || !isWeekend_(d)) dates.push(d); }
    if (dates.length > 62) throw new Error('반복 등록은 최대 62일까지 가능합니다.');
  }
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    if (p.id) {
      var r = readTable_(SHEETS.SCH).filter(function (x) { return x['일정ID'] === p.id; })[0];
      if (!r) throw new Error('수정할 일정을 찾지 못했습니다.');
      if (r['회의ID']) throw new Error('회의로 잡힌 일정은 회의 목록에서 취소해 주세요.');
      updateRow_(SHEETS.SCH, r._row, { '일정ID': p.id, '사번': p.emp, '날짜': p.date, '시작': toHM_(s), '종료': toHM_(e), '유형': p.type, '메모': p.memo || '', '회의ID': '' });
    } else {
      var sh = sheet_(SHEETS.SCH);
      var rows = dates.map(function (d) { return [newId_('S'), p.emp, d, toHM_(s), toHM_(e), p.type, p.memo || '', '']; });
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, 8).setNumberFormat('@').setValues(rows);
    }
    return true;
  } finally { lock.releaseLock(); }
}

function deleteSchedule(id) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var r = readTable_(SHEETS.SCH).filter(function (x) { return x['일정ID'] === id; })[0];
    if (!r) return true;
    if (r['회의ID']) throw new Error('회의로 잡힌 일정은 회의 목록에서 취소해 주세요.');
    sheet_(SHEETS.SCH).deleteRow(r._row);
    return true;
  } finally { lock.releaseLock(); }
}

// ===================== 자동 추천 핵심 =====================
/**
 * req = {
 *   agenda, title, duration(분), from: 'yyyy-MM-dd', deadlineDate, deadlineTime,
 *   requirements: [{dept, minRank, count}], requiredIds: [사번], organizerId,
 *   room: '' (자동) | 회의실 이름, skipWeekend: true
 * }
 */
function findSlots(req) {
  var cfg = getConfig_();
  var emps = readTable_(SHEETS.EMP).map(empOut_);
  var sch = readTable_(SHEETS.SCH).map(schOut_);
  var rooms = readTable_(SHEETS.ROOM).map(function (r) { return { name: r['회의실'], cap: Number(r['수용인원']) || 0 }; });
  var mtgs = readTable_(SHEETS.MTG).filter(function (m) { return m['상태'] !== '취소'; });
  return computeSlots_(req, { emps: emps, sch: sch, rooms: rooms, mtgs: mtgs, cfg: cfg, today: today_(), now: nowHM_() });
}

function computeSlots_(req, db) {
  var dur = Number(req.duration) || 60;
  if (dur < 10 || dur > 480) throw new Error('회의 길이는 10분~8시간 사이로 입력하세요.');
  if (!req.deadlineDate) throw new Error('마감 날짜를 입력하세요.');
  var deadlineTime = req.deadlineTime || '18:00';
  var from = req.from && req.from > db.today ? req.from : db.today;
  if (req.deadlineDate < from) throw new Error('마감이 이미 지났습니다. 마감 날짜를 확인하세요.');
  var dayCount = Math.round((parseDate_(req.deadlineDate) - parseDate_(from)) / 86400000);
  if (dayCount > 31) throw new Error('탐색 기간은 31일 이내로 잡아 주세요.');

  var empById = {};
  db.emps.forEach(function (e) { empById[e.id] = e; });
  var required = (req.requiredIds || []).filter(function (id) { return empById[id]; });
  var reqs = (req.requirements || []).filter(function (r) { return r.dept; }).map(function (r) {
    return { dept: r.dept, minRank: r.minRank || '', count: Math.max(1, Number(r.count) || 1) };
  });
  if (!required.length && !reqs.length) throw new Error('필수 참석자나 필요 부서를 한 명/한 줄 이상 넣어 주세요.');

  // 사람별 바쁜 시간 (날짜 → [[s,e],...])
  var busy = {};
  db.sch.forEach(function (s) {
    var k = s.emp + '|' + s.date;
    (busy[k] = busy[k] || []).push([toMin_(s.start), toMin_(s.end)]);
  });
  // 주간 회의 부담(공평하게 배정하기 위함)
  var load = {};
  db.sch.forEach(function (s) { if (s.type === '회의' && s.date >= from && s.date <= req.deadlineDate) load[s.emp] = (load[s.emp] || 0) + 1; });

  function freeReason(id, date, s, e) {
    var emp = empById[id];
    if (!emp) return '등록되지 않은 사람';
    var ws = toMin_(emp.start || '08:00'), we = toMin_(emp.end || '17:00');
    if (s < ws || e > we) return '근무시간 밖';
    var list = busy[id + '|' + date] || [];
    for (var i = 0; i < list.length; i++) if (list[i][0] < e && s < list[i][1]) return '일정 있음';
    return '';
  }

  // 회의실 점유
  var roomBusy = {};
  db.mtgs.forEach(function (m) {
    var k = m['회의실'] + '|' + m['날짜'];
    (roomBusy[k] = roomBusy[k] || []).push([toMin_(m['시작']), toMin_(m['종료'])]);
  });
  function roomFree(name, date, s, e) {
    var l = roomBusy[name + '|' + date] || [];
    for (var i = 0; i < l.length; i++) if (l[i][0] < e && s < l[i][1]) return false;
    return true;
  }

  // 탐색 범위: 관련된 사람들의 가장 이른 출근 ~ 가장 늦은 퇴근
  var involved = required.slice();
  reqs.forEach(function (r) { db.emps.forEach(function (e) { if (e.dept === r.dept) involved.push(e.id); }); });
  var dayStart = 24 * 60, dayEnd = 0;
  involved.forEach(function (id) {
    var e = empById[id]; if (!e) return;
    dayStart = Math.min(dayStart, toMin_(e.start || '08:00'));
    dayEnd = Math.max(dayEnd, toMin_(e.end || '17:00'));
  });
  if (dayStart >= dayEnd) { dayStart = 480; dayEnd = 1020; }

  var step = db.cfg.step;
  var ranks = db.cfg.ranks;
  var full = [], partial = [];
  var d = from;
  for (var day = 0; day <= dayCount; day++, d = addDays_(from, day)) {
    if (req.skipWeekend !== false && isWeekend_(d)) continue;
    var startMin = Math.ceil(dayStart / step) * step;
    if (d === db.today) startMin = Math.max(startMin, Math.ceil((toMin_(db.now) + 1) / step) * step);
    for (var s = startMin; s + dur <= dayEnd; s += step) {
      var e = s + dur;
      if (d === req.deadlineDate && e > toMin_(deadlineTime)) break;

      // 1) 필수 참석자
      var missing = [];
      required.forEach(function (id) { var why = freeReason(id, d, s, e); if (why) missing.push({ who: empById[id].name, why: why }); });

      // 2) 부서·직급 조건
      var chosen = required.slice();
      var lines = [];
      var shortCount = 0;
      reqs.forEach(function (r) {
        var minIdx = r.minRank ? rankIndex_(ranks, r.minRank) : -1;
        // 필수 참석자 중 조건에 맞는 사람이 있으면 그 사람으로 채운 것으로 본다
        var already = chosen.filter(function (id) {
          var e2 = empById[id];
          return e2.dept === r.dept && rankIndex_(ranks, e2.rank) >= minIdx && !freeReason(id, d, s, e);
        }).length;
        var cands = db.emps.filter(function (x) {
          return x.dept === r.dept && chosen.indexOf(x.id) < 0 && rankIndex_(ranks, x.rank) >= minIdx && !freeReason(x.id, d, s, e);
        }).sort(function (a, b) {
          // 조건에 가장 가까운 직급 먼저(윗사람을 불필요하게 부르지 않기), 같은 직급이면 회의가 적은 사람
          var ra = rankIndex_(ranks, a.rank) - minIdx, rb = rankIndex_(ranks, b.rank) - minIdx;
          if (ra !== rb) return ra - rb;
          return (load[a.id] || 0) - (load[b.id] || 0);
        });
        var need = Math.max(0, r.count - already);
        var pick = cands.slice(0, need);
        pick.forEach(function (x) { chosen.push(x.id); });
        var short = need - pick.length;
        shortCount += short;
        lines.push({
          dept: r.dept, minRank: r.minRank, count: r.count,
          picked: pick.map(function (x) { return x.id; }),
          alternates: cands.slice(need, need + 3).map(function (x) { return x.id; }),
          short: short
        });
      });

      // 3) 회의실
      var room = '';
      if (req.room) {
        room = roomFree(req.room, d, s, e) ? req.room : '';
      } else {
        var fit = db.rooms.filter(function (rm) { return rm.cap >= chosen.length && roomFree(rm.name, d, s, e); })
          .sort(function (a, b) { return a.cap - b.cap; });
        room = fit.length ? fit[0].name : '';
      }

      var slot = {
        date: d, start: toHM_(s), end: toHM_(e),
        attendees: chosen,
        lines: lines,
        room: room,
        missing: missing,
        short: shortCount,
        noRoom: !room
      };
      if (!missing.length && !shortCount && room) full.push(slot);
      else partial.push(slot);
      if (full.length >= db.cfg.topN) break;
    }
    if (full.length >= db.cfg.topN) break;
  }

  partial.sort(function (a, b) {
    var pa = a.missing.length * 10 + a.short * 3 + (a.noRoom ? 1 : 0);
    var pb = b.missing.length * 10 + b.short * 3 + (b.noRoom ? 1 : 0);
    if (pa !== pb) return pa - pb;
    return (a.date + a.start).localeCompare(b.date + b.start);
  });

  var people = {};
  db.emps.forEach(function (e) { people[e.id] = { name: e.name, dept: e.dept, rank: e.rank }; });
  return {
    slots: full,
    nearMisses: full.length ? [] : partial.slice(0, 3),
    people: people,
    searched: { from: from, to: req.deadlineDate, deadlineTime: deadlineTime, duration: dur }
  };
}

/** 추천 시간 확정: 회의 기록 + 참석자 일정에 회의 추가 */
function confirmMeeting(p) {
  if (!p || !p.date || !p.start || !p.end || !p.attendees || !p.attendees.length) throw new Error('확정할 회의 정보가 부족합니다.');
  var lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    // 그 사이 다른 사람이 일정을 바꿨는지 다시 확인
    var s = toMin_(p.start), e = toMin_(p.end);
    var sch = readTable_(SHEETS.SCH);
    var clash = [];
    var empById = {};
    readTable_(SHEETS.EMP).forEach(function (x) { empById[x['사번']] = x; });
    p.attendees.forEach(function (id) {
      sch.forEach(function (r) {
        if (r['사번'] === id && r['날짜'] === p.date && toMin_(r['시작']) < e && s < toMin_(r['종료'])) clash.push(empById[id] ? empById[id]['이름'] : id);
      });
    });
    if (clash.length) throw new Error('그 사이 일정이 바뀌었습니다: ' + clash.filter(function (v, i, a) { return a.indexOf(v) === i; }).join(', ') + '. 다시 추천받아 주세요.');
    if (p.room) {
      var rc = readTable_(SHEETS.MTG).filter(function (m) {
        return m['상태'] !== '취소' && m['회의실'] === p.room && m['날짜'] === p.date && toMin_(m['시작']) < e && s < toMin_(m['종료']);
      });
      if (rc.length) throw new Error(p.room + '이(가) 그 사이 예약되었습니다. 다시 추천받아 주세요.');
    }

    var id = newId_('M');
    var title = p.title || p.agenda || '부서 간 회의';
    appendRow_(SHEETS.MTG, {
      '회의ID': id, '안건유형': p.agenda || '', '제목': title, '주관자사번': p.organizerId || '',
      '날짜': p.date, '시작': p.start, '종료': p.end, '회의실': p.room || '',
      '참석자사번': p.attendees.join(','), '상태': '확정',
      '등록시각': Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm')
    });
    var sh = sheet_(SHEETS.SCH);
    var rows = p.attendees.map(function (a) { return [newId_('S'), a, p.date, p.start, p.end, '회의', title, id]; });
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, 8).setNumberFormat('@').setValues(rows);

    if (p.sendInvite) sendInvites_(title, p, empById);
    return { id: id };
  } finally { lock.releaseLock(); }
}

/** 선택: 이메일이 등록된 참석자에게 구글 캘린더 초대 발송 */
function sendInvites_(title, p, empById) {
  var guests = p.attendees.map(function (a) { return empById[a] ? empById[a]['이메일'] : ''; }).filter(String);
  if (!guests.length) return;
  var d = parseDate_(p.date);
  var st = new Date(d), en = new Date(d);
  st.setHours(Math.floor(toMin_(p.start) / 60), toMin_(p.start) % 60);
  en.setHours(Math.floor(toMin_(p.end) / 60), toMin_(p.end) % 60);
  CalendarApp.getDefaultCalendar().createEvent('[회의콕] ' + title, st, en, {
    location: p.room || '', guests: guests.join(','), sendInvites: true
  });
}

// ===================== 회의 목록 =====================
function listMeetings() {
  var emps = {};
  readTable_(SHEETS.EMP).forEach(function (e) { emps[e['사번']] = e['이름'] + ' ' + e['직급'] + '(' + e['부서'] + ')'; });
  return readTable_(SHEETS.MTG).map(function (m) {
    return {
      id: m['회의ID'], agenda: m['안건유형'], title: m['제목'], organizer: emps[m['주관자사번']] || '',
      date: m['날짜'], start: m['시작'], end: m['종료'], room: m['회의실'], status: m['상태'],
      attendeeIds: String(m['참석자사번']).split(',').filter(String),
      attendees: String(m['참석자사번']).split(',').filter(String).map(function (id) { return emps[id] || id; })
    };
  }).sort(function (a, b) { return (b.date + b.start).localeCompare(a.date + a.start); });
}

function cancelMeeting(id) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var m = readTable_(SHEETS.MTG).filter(function (x) { return x['회의ID'] === id; })[0];
    if (!m) throw new Error('회의를 찾지 못했습니다.');
    m['상태'] = '취소';
    updateRow_(SHEETS.MTG, m._row, m);
    var rows = readTable_(SHEETS.SCH).filter(function (r) { return r['회의ID'] === id; }).map(function (r) { return r._row; });
    rows.sort(function (a, b) { return b - a; }).forEach(function (r) { sheet_(SHEETS.SCH).deleteRow(r); });
    return listMeetings();
  } finally { lock.releaseLock(); }
}