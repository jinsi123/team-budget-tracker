/**
 * 예산 사용 관리 (Google Sheets + Apps Script)
 *
 * 구조: 팀 → 단위사업 → 세부사업. 예산은 세부사업·연도마다 정한다.
 * 첫 화면에 팀 폴더가 보이고, 누르면 그 팀의 단위사업·세부사업별 예산과 집행 내역이 나온다.
 * 이호조 [지출집행현황] 엑셀을 올리면 품의 단위로 집행내역에 들어간다(같은 품의는 덮어씀).
 *
 * 시트 구성
 *   팀구성   : 팀 | 단위사업                (단위사업이 어느 팀 것인지)
 *   예산     : 연도 | 단위사업 | 세부사업 | 예산
 *   집행내역 : 번호 | 출처 | 연도 | 단위사업 | 세부사업 | 통계목 | 일시 | 시간 | 담당자 | 거래처 | 금액 | 남은금액 | 입력시각
 *
 * 이호조 파일에서 저장하는 항목은 단위사업·세부사업·담당자·일시·거래처(앞 두 글자)·금액·통계목뿐이다.
 * (품의번호·적요·결재자 등은 저장하지 않는다. [번호]는 같은 품의를 알아보기 위한 해시값이다.)
 * 금액     = 원인행위(계약) 금액. 원인행위 전이면 품의금액.
 * 남은금액 = 그 세부사업·그 연도 예산 − (일시 순으로 그 줄까지의 금액 합)
 */

var SHEET_TEAMS = '팀구성';
var SHEET_BUDGET = '예산';
var SHEET_LOG = '집행내역';
var TEAM_HEADERS = ['팀', '단위사업'];
var BUDGET_HEADERS = ['연도', '단위사업', '세부사업', '예산'];
var LOG_HEADERS = ['번호', '출처', '연도', '단위사업', '세부사업', '통계목', '일시', '시간', '담당자', '거래처', '금액', '남은금액', '입력시각'];
var C = {}; LOG_HEADERS.forEach(function (h, i) { C[h] = i; });
var SRC_EHOJO = '이호조';
var SRC_MANUAL = '직접입력';
var TZ = 'Asia/Seoul';

/**
 * 이 스프레드시트. 시트 메뉴에서 실행될 때는 열린 시트를 쓰고 그 ID를 기억해 둔다.
 * 웹 화면에서는 열린 시트가 없으므로(getActive가 null) 기억해 둔 ID로 연다.
 */
function ss_() {
  var props = PropertiesService.getScriptProperties();
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    if (props.getProperty('SHEET_ID') !== active.getId()) props.setProperty('SHEET_ID', active.getId());
    return active;
  }
  var id = props.getProperty('SHEET_ID');
  if (!id) throw new Error('예산 시트를 아직 찾지 못했습니다. 구글 시트를 한 번 열어 위쪽 [예산 관리] 메뉴에서 [처음 설정하기]를 눌러 주세요.');
  return SpreadsheetApp.openById(id);
}

/* ------------------------------------------------------------------ */
/* 메뉴                                                                */
/* ------------------------------------------------------------------ */

function onOpen() {
  try { ss_(); } catch (e) {}
  SpreadsheetApp.getUi()
    .createMenu('예산 관리')
    .addItem('처음 설정하기', 'setup')
    .addItem('웹 앱 주소 등록', 'setWebAppUrl')
    .addItem('사이트 주소·비밀번호 보기', 'showSiteUrl')
    .addItem('비밀번호 정하기', 'setPassword')
    .addItem('남은금액 다시 계산', 'recalc')
    .addToUi();
}

/** 시트에서 집행내역이나 예산을 직접 고치면 남은금액을 다시 계산한다. */
function onEdit(e) {
  try {
    var name = e && e.range && e.range.getSheet().getName();
    if (name === SHEET_LOG || name === SHEET_BUDGET) recalc();
  } catch (err) {}
}

function setup() {
  var ss = ss_();
  ensureSheet_(ss, SHEET_TEAMS, TEAM_HEADERS);
  ensureSheet_(ss, SHEET_BUDGET, BUDGET_HEADERS).getRange('D:D').setNumberFormat('#,##0');
  var log = ensureSheet_(ss, SHEET_LOG, LOG_HEADERS);
  log.getRange('G:H').setNumberFormat('@');
  log.getRange('K:L').setNumberFormat('#,##0');
  var first = ss.getSheets()[0];
  if ([SHEET_TEAMS, SHEET_BUDGET, SHEET_LOG].indexOf(first.getName()) < 0 && first.getLastRow() === 0) ss.deleteSheet(first);
  password_();
  SpreadsheetApp.getUi().alert(
    '설정을 마쳤습니다.\n\n다음으로 웹 앱을 배포하고 [예산 관리 > 웹 앱 주소 등록]을 해 주세요.\n그다음 사이트에서 [이호조 엑셀 올리기]로 지출집행현황 파일을 올리면 단위사업·세부사업이 자동으로 채워집니다.');
}

function ensureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (sh) return sh;
  sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#EEF1F7');
  sh.setFrozenRows(1);
  return sh;
}

function webAppUrl_() {
  return PropertiesService.getScriptProperties().getProperty('WEB_APP_URL') || '';
}

function setWebAppUrl() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('웹 앱 주소 등록', '배포 후 받은 웹 앱 URL(…/exec)을 붙여넣으세요.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var url = r.getResponseText().trim();
  if (!/^https:\/\/script\.google\.com\/.+\/exec$/.test(url)) {
    ui.alert('주소가 …/exec 로 끝나야 합니다. 다시 확인해 주세요.');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('WEB_APP_URL', url);
  showSiteUrl();
}

function password_() {
  var props = PropertiesService.getScriptProperties();
  var k = props.getProperty('PASSWORD');
  if (!k) {
    k = String(Math.floor(100000 + Math.random() * 900000));
    props.setProperty('PASSWORD', k);
  }
  return k;
}

function check_(k) {
  if (!k || String(k) !== password_()) throw new Error('비밀번호가 맞지 않습니다. 화면을 새로고침해 다시 들어가 주세요.');
}

function showSiteUrl() {
  var url = webAppUrl_();
  if (!url) {
    SpreadsheetApp.getUi().alert('먼저 웹 앱으로 배포하고 [예산 관리 > 웹 앱 주소 등록]을 해 주세요.');
    return;
  }
  var html = '<div style="font-family:sans-serif;font-size:14px;line-height:1.6">' +
    '<p>아래 주소를 열고 비밀번호를 넣으면 팀 폴더 화면이 나옵니다. 휴대폰에서도 됩니다.</p>' +
    '<p><a href="' + url + '" target="_blank">사이트 열기</a></p>' +
    '<p><input style="width:100%;font-size:12px" readonly value="' + esc_(url) + '" onclick="this.select()"></p>' +
    '<p>비밀번호: <b>' + esc_(password_()) + '</b><br><span style="color:#5B6478">[예산 관리 > 비밀번호 정하기]에서 바꿀 수 있습니다.</span></p></div>';
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(460).setHeight(300), '사이트 주소');
}

function setPassword() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('비밀번호 정하기', '사이트에 들어갈 때 쓸 비밀번호를 적으세요 (4자 이상).', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var pw = r.getResponseText().trim();
  if (pw.length < 4) { ui.alert('4자 이상으로 정해 주세요.'); return; }
  PropertiesService.getScriptProperties().setProperty('PASSWORD', pw);
  ui.alert('비밀번호를 바꿨습니다. 다른 기기에서도 새 비밀번호로 다시 들어가야 합니다.');
}

function esc_(s) {
  return String(s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

/* ------------------------------------------------------------------ */
/* 시트 읽기·쓰기                                                       */
/* ------------------------------------------------------------------ */

function num_(v) {
  if (typeof v === 'number') return v;
  var n = Number(String(v || '').replace(/[^\d.-]/g, ''));
  return isNaN(n) ? 0 : n;
}

function str_(v) { return String(v == null ? '' : v).trim(); }

function dateText_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  var s = str_(v).replace(/[./]/g, '-').replace(/-$/, '');
  if (/^\d{8}$/.test(s)) return s.slice(0, 4) + '-' + s.slice(4, 6) + '-' + s.slice(6, 8);
  var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  return m ? m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2) : s;
}

function timeText_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'HH:mm');
  var m = str_(v).match(/^(\d{1,2}):(\d{2})/);
  return m ? ('0' + m[1]).slice(-2) + ':' + m[2] : '';
}

function body_(sh, width) {
  return sh && sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues() : [];
}

function readTeams_() {
  return body_(ss_().getSheetByName(SHEET_TEAMS), 2).map(function (r) {
    return { team: str_(r[0]), unit: str_(r[1]) };
  }).filter(function (x) { return x.team; });
}

function readBudgets_() {
  return body_(ss_().getSheetByName(SHEET_BUDGET), 4).map(function (r) {
    return { year: Number(str_(r[0]).slice(0, 4)) || 0, unit: str_(r[1]), sub: str_(r[2]), budget: num_(r[3]) };
  }).filter(function (b) { return b.year && b.sub; });
}

function readLogRaw_() {
  var sh = ss_().getSheetByName(SHEET_LOG);
  if (sh && sh.getLastColumn() > 0) {
    var head = sh.getRange(1, 1, 1, LOG_HEADERS.length).getValues()[0].map(str_);
    if (head.join('|') !== LOG_HEADERS.join('|')) throw new Error('[집행내역] 시트의 열 구성이 예전 형식입니다. 그 시트 이름을 바꾸거나 지운 뒤 [예산 관리 > 처음 설정하기]를 다시 눌러 주세요.');
  }
  return body_(sh, LOG_HEADERS.length);
}

function toRow_(r) {
  var date = dateText_(r[C['일시']]);
  return {
    id: str_(r[C['번호']]), src: str_(r[C['출처']]) || SRC_MANUAL,
    year: Number(str_(r[C['연도']]).slice(0, 4)) || Number(date.slice(0, 4)) || 0,
    unit: str_(r[C['단위사업']]), sub: str_(r[C['세부사업']]), item: str_(r[C['통계목']]),
    date: date, time: timeText_(r[C['시간']]),
    person: str_(r[C['담당자']]), vendor: str_(r[C['거래처']]), amount: num_(r[C['금액']])
  };
}

function readLog_() {
  return readLogRaw_().map(toRow_).filter(function (x) { return x.sub || x.amount; });
}

/** 세부사업·연도마다 일시·시간 순으로 남은금액을 계산해 [남은금액] 칸에 쓴다. */
function recalc() {
  var sh = ss_().getSheetByName(SHEET_LOG);
  if (!sh || sh.getLastRow() < 2) return;
  var bm = {};
  readBudgets_().forEach(function (b) { var k = b.year + '|' + b.unit + '|' + b.sub; bm[k] = (bm[k] || 0) + b.budget; });
  var raw = readLogRaw_();
  var rows = raw.map(function (r, i) { var x = toRow_(r); x.i = i; return x; });
  var left = {}, out = raw.map(function () { return ['']; });
  rows.slice().sort(function (a, b) {
    var ka = a.date + ' ' + a.time, kb = b.date + ' ' + b.time;
    return ka < kb ? -1 : ka > kb ? 1 : a.i - b.i;
  }).forEach(function (r) {
    if (!r.sub) return;
    var k = r.year + '|' + r.unit + '|' + r.sub;
    if (!(k in left)) left[k] = bm[k] || 0;
    left[k] -= r.amount;
    out[r.i][0] = left[k];
  });
  sh.getRange(2, C['남은금액'] + 1, out.length, 1).setValues(out);
}

function newId_() {
  return 'M' + Utilities.formatDate(new Date(), TZ, 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
}

function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/** 시트 본문을 통째로 다시 쓴다. */
function writeBody_(sh, headers, rows) {
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, headers.length).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

/* ------------------------------------------------------------------ */
/* 웹 앱                                                               */
/* ------------------------------------------------------------------ */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('App')
    .setTitle('예산 사용 관리')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 로그인. 틀리면 잠깐 기다린 뒤 false. */
function login(pw) {
  if (str_(pw) === password_()) return true;
  Utilities.sleep(1500);
  return false;
}

/** { teams: [{team, unit}], budgets: [{year, unit, sub, budget}], rows: [...] } */
function getData(key) {
  check_(key);
  return { teams: readTeams_(), budgets: readBudgets_(), rows: readLog_() };
}

function cleanManual_(x) {
  x = x || {};
  var r = {
    unit: str_(x.unit), sub: str_(x.sub), item: str_(x.item), date: dateText_(x.date), time: timeText_(x.time),
    person: str_(x.person), vendor: str_(x.vendor), amount: num_(x.amount)
  };
  if (!r.unit || !r.sub) throw new Error('단위사업과 세부사업을 골라 주세요.');
  if (!r.person) throw new Error('담당자를 적어 주세요.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(r.date)) throw new Error('일시(날짜)를 확인해 주세요.');
  if (!r.vendor) throw new Error('거래처를 적어 주세요.');
  if (!r.amount) throw new Error('금액을 적어 주세요.');
  return r;
}

function manualLine_(id, r) {
  var line = LOG_HEADERS.map(function () { return ''; });
  line[C['번호']] = id; line[C['출처']] = SRC_MANUAL; line[C['연도']] = Number(r.date.slice(0, 4));
  line[C['단위사업']] = r.unit; line[C['세부사업']] = r.sub; line[C['통계목']] = r.item;
  line[C['일시']] = r.date; line[C['시간']] = r.time; line[C['담당자']] = r.person;
  line[C['거래처']] = r.vendor; line[C['금액']] = r.amount; line[C['입력시각']] = now_();
  return line;
}

function addRow(key, x) {
  check_(key);
  var r = cleanManual_(x);
  withLock_(function () {
    ss_().getSheetByName(SHEET_LOG).appendRow(manualLine_(newId_(), r));
    recalc();
  });
  return getData(key);
}

function findRow_(id) {
  var raw = readLogRaw_();
  for (var i = 0; i < raw.length; i++) if (str_(raw[i][C['번호']]) === String(id)) return { row: i + 2, src: str_(raw[i][C['출처']]) };
  throw new Error('그 내역을 찾지 못했습니다. 다른 사람이 먼저 지웠을 수 있습니다. 새로고침해 주세요.');
}

function updateRow(key, id, x) {
  check_(key);
  var r = cleanManual_(x);
  withLock_(function () {
    var hit = findRow_(id);
    if (hit.src === SRC_EHOJO) throw new Error('이호조에서 올린 내역은 여기서 고칠 수 없습니다. 이호조 엑셀을 다시 올리면 바뀐 내용이 반영됩니다.');
    ss_().getSheetByName(SHEET_LOG).getRange(hit.row, 1, 1, LOG_HEADERS.length).setValues([manualLine_(id, r)]);
    recalc();
  });
  return getData(key);
}

function deleteRow(key, id) {
  check_(key);
  withLock_(function () {
    ss_().getSheetByName(SHEET_LOG).deleteRow(findRow_(id).row);
    recalc();
  });
  return getData(key);
}

/**
 * 이호조 엑셀에서 읽은 품의 목록을 넣는다. 같은 번호는 덮어쓰고, 새 번호는 뒤에 붙인다.
 * 화면(브라우저)에서 엑셀을 읽어 품의 단위로 묶은 뒤 보낸다.
 * items: [{id, year, unit, sub, item, date, person, vendor, amount}]  (vendor는 앞 두 글자만 남긴다)
 */
function importEhojo(key, items) {
  check_(key);
  items = items || [];
  var added = 0, updated = 0;
  withLock_(function () {
    var sh = ss_().getSheetByName(SHEET_LOG);
    var raw = readLogRaw_();
    var at = {}; raw.forEach(function (r, i) { at[str_(r[C['번호']])] = i; });
    items.forEach(function (x) {
      var id = str_(x.id);
      if (!id) return;
      var line = LOG_HEADERS.map(function () { return ''; });
      line[C['번호']] = id; line[C['출처']] = SRC_EHOJO; line[C['연도']] = Number(x.year) || Number(String(x.date).slice(0, 4));
      line[C['단위사업']] = str_(x.unit); line[C['세부사업']] = str_(x.sub); line[C['통계목']] = str_(x.item);
      line[C['일시']] = dateText_(x.date); line[C['시간']] = '';
      line[C['담당자']] = str_(x.person); line[C['거래처']] = Array.from(str_(x.vendor)).slice(0, 2).join('');
      line[C['금액']] = num_(x.amount);
      line[C['입력시각']] = now_();
      if (id in at) { raw[at[id]] = line; updated++; }
      else { at[id] = raw.length; raw.push(line); added++; }
    });
    writeBody_(sh, LOG_HEADERS, raw);
    sh.getRange('G:H').setNumberFormat('@');
    recalc();
  });
  var d = getData(key);
  d.result = { added: added, updated: updated };
  return d;
}

/** 단위사업을 팀에 배정한다. map: { 단위사업: 팀 } (팀이 빈 값이면 배정 해제) */
function assignUnits(key, map) {
  check_(key);
  withLock_(function () {
    var cur = readTeams_();
    var units = Object.keys(map || {});
    var keep = cur.filter(function (t) { return units.indexOf(t.unit) < 0; });
    units.forEach(function (u) { var t = str_(map[u]); if (t) keep.push({ team: t, unit: u }); });
    // 단위사업이 하나도 없는 팀도 폴더로 남긴다
    cur.forEach(function (t) { if (!keep.some(function (k) { return k.team === t.team; })) keep.push({ team: t.team, unit: '' }); });
    keep = keep.filter(function (k) { return k.unit || !keep.some(function (o) { return o.team === k.team && o.unit; }); });
    writeBody_(ss_().getSheetByName(SHEET_TEAMS), TEAM_HEADERS, keep.map(function (k) { return [k.team, k.unit]; }));
  });
  return getData(key);
}

/** 새 팀 폴더 */
function addTeam(key, team) {
  check_(key);
  team = str_(team);
  if (!team) throw new Error('팀 이름을 적어 주세요.');
  withLock_(function () {
    if (readTeams_().some(function (t) { return t.team === team; })) throw new Error('이미 있는 팀입니다.');
    ss_().getSheetByName(SHEET_TEAMS).appendRow([team, '']);
  });
  return getData(key);
}

/**
 * 한 연도의 세부사업 예산을 저장한다. 목록에 있는 세부사업만 바꾸고 나머지는 그대로 둔다.
 * list: [{unit, sub, budget}]  (budget이 빈 값이면 그 줄을 지운다)
 */
function saveBudgets(key, year, list) {
  check_(key);
  year = Number(year);
  if (!(year >= 2000 && year <= 2100)) throw new Error('연도를 확인해 주세요.');
  withLock_(function () {
    var cur = readBudgets_();
    var keyOf = function (b) { return b.unit + '|' + b.sub; };
    var touched = {}; (list || []).forEach(function (b) { touched[keyOf(b)] = b; });
    var out = cur.filter(function (b) { return b.year !== year || !touched[keyOf(b)]; });
    Object.keys(touched).forEach(function (k) {
      var b = touched[k];
      if (str_(b.budget) !== '' && str_(b.sub)) out.push({ year: year, unit: str_(b.unit), sub: str_(b.sub), budget: num_(b.budget) });
    });
    out.sort(function (a, b) { return a.year - b.year || (a.unit < b.unit ? -1 : a.unit > b.unit ? 1 : 0); });
    var sh = ss_().getSheetByName(SHEET_BUDGET);
    writeBody_(sh, BUDGET_HEADERS, out.map(function (b) { return [b.year, b.unit, b.sub, b.budget]; }));
    sh.getRange('D:D').setNumberFormat('#,##0');
    recalc();
  });
  return getData(key);
}
