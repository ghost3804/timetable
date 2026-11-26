'use strict';
/* =====================================================================
   시간표 자동 편성 (1차) — 브라우저 안에서만 동작
   입력: 운영계획서(.xls/.xlsx) + 코드표(.xlsx)
   출력: 전자시간표(업로드용) / 출력용시간표 / 평가일정(평가날출력-4번 B~W) / 편성 저장(JSON)
   ===================================================================== */
const $ = s => document.querySelector(s);
const $$ = s => [...document.querySelectorAll(s)];
const WD = '일월화수목금토';
const norm = s => String(s ?? '').replace(/\s+/g, '');
const esc = s => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
  .replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g, m => '_x' + m.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0') + '_');
const pad = n => String(n).padStart(2, '0');
const ymd = d => `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}`;
const dkey = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const parseD = s => { const [y, m, d] = s.split('-').map(Number); return new Date(y, m - 1, d); };
const addD = (d, n) => { const x = new Date(d); x.setDate(x.getDate() + n); return x; };
const serial = d => (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) - Date.UTC(1899, 11, 30)) / 86400000;
const hhmm = (t, add) => { let m = +t.slice(0, 2) * 60 + +t.slice(2) + add; return pad(Math.floor(m / 60)) + pad(m % 60); };

/* ---------- 공휴일 (편집 가능) ---------- */
const DEFAULT_HOLIDAYS = `2026-01-01 신정
2026-02-16 설날
2026-02-17 설날
2026-02-18 설날
2026-03-02 대체공휴일(삼일절)
2026-05-05 어린이날
2026-05-25 대체공휴일(부처님오신날)
2026-06-03 지방선거
2026-07-17 제헌절
2026-08-17 대체공휴일(광복절)
2026-09-24 추석
2026-09-25 추석
2026-10-05 대체공휴일(개천절)
2026-10-09 한글날
2026-12-25 성탄절
2027-01-01 신정
2027-02-08 설날
2027-02-09 대체공휴일(설날)
2027-03-01 삼일절
2027-05-03 대체공휴일(노동절)
2027-05-05 어린이날
2027-05-13 부처님오신날
2027-07-19 대체공휴일(제헌절)
2027-08-16 대체공휴일(광복절)
2027-09-14 추석
2027-09-15 추석
2027-09-16 추석
2027-10-04 대체공휴일(개천절)
2027-10-11 대체공휴일(한글날)
2027-12-27 대체공휴일(성탄절)`;

/* ---------- 교시표 ---------- */
// lunchAfter: 4 = 4교시 뒤 점심, 3 = 3교시 뒤 점심
function periodTable(lunchAfter) {
  const base = [
    { p: 1, s: '0900', disp: '09:00-09:50' }, { p: 2, s: '1000', disp: '10:00-10:50' }, { p: 3, s: '1100', disp: '11:00-11:50' },
    lunchAfter === 4 ? { p: 4, s: '1200', disp: '12:00-12:50' } : { p: 4, s: '1240', disp: '12:40-13:30' },
    { p: 5, s: '1340', disp: '13:40-14:30' }, { p: 6, s: '1440', disp: '14:40-15:30' }, { p: 7, s: '1540', disp: '15:40-16:30' }, { p: 8, s: '1640', disp: '16:40-17:30' }];
  const lunch = lunchAfter === 4 ? { p: '점심', s: '1250', disp: '12:50-13:40' } : { p: '점심', s: '1150', disp: '11:50-12:40' };
  const rows = [...base]; rows.splice(lunchAfter, 0, lunch);
  return { rows, lunch, lunchAfter, byP: Object.fromEntries(base.map(x => [x.p, x])) };
}

/* ---------- 상태 ---------- */
const S = { plan: null, codes: null, subjects: [], lessons: {}, lt: {}, lr: {}, lock: {}, order: [], fixedRows: [], ref: [], refIdx: {}, days: [], endDate: null, holidays: {}, files: {} };

/* ---------- 파일 읽기 ---------- */
async function readRows(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: 'array' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  // 시트 범위가 B1처럼 A1이 아닌 곳에서 시작해도 칸 번호가 밀리지 않도록 항상 A1부터 읽음
  const rg = XLSX.utils.decode_range(ws['!ref'] || 'A1'); rg.s.c = 0; rg.s.r = 0;
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true, blankrows: true, range: rg });
}
const cellS = v => String(v ?? '').trim();

function parsePlan(rows) {
  const get = (r, c) => cellS((rows[r] || [])[c]);
  const nonEmpty = r => (rows[r] || []).map((v, i) => [i, cellS(v)]).filter(x => x[1]);
  const rowText = r => nonEmpty(r).map(x => x[1]).join(' ');
  const isNoise = r => /출력일자/.test(rowText(r));
  const plan = { name: '', subjects: [], teachers: [], rooms: [], methods: {}, totalHours: 0 };
  // 과정명
  for (let r = 0; r < rows.length; r++) {
    const ne = nonEmpty(r);
    const i = ne.findIndex(x => x[1] === '훈련과정명');
    if (i >= 0 && ne[i + 1]) { plan.name = ne[i + 1][1]; break; }
  }
  // 교과편성연계표 구간
  // 형식 1: '■ 교과편성연계표' 아래 소양/전공/비NCS(이론)/비NCS(실기)
  // 형식 2: NCS 비적용 과정 — 'NCS 비적용 교과' 표 하나, 시간은 이론+실기
  const SEC = { 'NCS소양교과': '소양', 'NCS전공교과': '전공', '비NCS교과(이론)': '비NCS 이론', '비NCS교과(실기)': '비NCS 실기', 'NCS비적용교과': '비NCS' };
  let sec = null, hdr = null, started = false;
  for (let r = 0; r < rows.length; r++) {
    const t = rowText(r), n = norm(t);
    if (n.includes('교과편성연계표')) { started = true; continue; }
    if (started && n.startsWith('■')) break;
    if (n.startsWith('■교수학습방법과평가방법')) break;
    if (SEC[n]) { sec = SEC[n]; hdr = null; started = true; continue; }
    if (!sec || isNoise(r) || !t) continue;
    if (norm(get(r, 1)) === '교과목') {
      hdr = {};
      nonEmpty(r).forEach(([i, v]) => { hdr[norm(v)] = i; });
      continue;
    }
    if (!hdr) continue;
    const first = get(r, 1);
    if (!first || /^총\s*(훈련|이론|실기)시간|조회된|등록된/.test(first)) { if (/^총\s*(훈련|이론|실기)시간/.test(first)) { sec = null; hdr = null; } continue; }
    const col = k => { const key = Object.keys(hdr).find(h => h.startsWith(k)); return key === undefined ? -1 : hdr[key]; };
    const th = col('이론'), ph = col('실기');
    const hours = col('편성') >= 0 ? (parseFloat(get(r, col('편성'))) || 0) : (parseFloat(get(r, th)) || 0) + (parseFloat(get(r, ph)) || 0);
    const tRaw = get(r, col('훈련교'));
    const teachers = tRaw.split(/\n/).map(x => x.replace(/\[.*?\]/g, '').trim()).filter(Boolean);
    let kind = sec;
    if (sec === '비NCS') kind = (parseFloat(get(r, ph)) || 0) > 0 ? '비NCS 실기' : '비NCS 이론';
    const s = { kind, name: first, hours, teachers, ncs: '', unit: first, level: '' };
    if (sec === '전공') {
      const u = get(r, col('NCS능력단위'));
      const m = u.match(/\[([^\]]+)\]\s*(.*)/);
      if (m) { s.ncs = m[1].trim(); s.unit = m[2].trim(); }
      s.level = String(parseInt(get(r, col('수준'))) || '');
    } else {
      s.unitDetail = get(r, col('단원'));
    }
    plan.subjects.push(s);
  }
  // 평가방법
  let ev = null;
  for (let r = 0; r < rows.length; r++) {
    const n = norm(rowText(r));
    if (n.includes('교수학습방법과평가방법')) { ev = {}; continue; }
    if (!ev) continue;
    if (n.startsWith('■')) break;
    if (isNoise(r)) continue;
    if (norm(get(r, 1)) === '교과구분') { nonEmpty(r).forEach(([i, v]) => { ev[norm(v)] = i; }); continue; }
    const kName = Object.keys(ev).find(k => k.startsWith('능력단위또는')), kM = Object.keys(ev).find(k => k === '평가방법');
    if (kName && kM) {
      const nm = get(r, ev[kName]), m = get(r, ev[kM]);
      if (nm) plan.methods[norm(nm)] = m;
    }
  }
  // 훈련교·강사, 훈련시설
  const listAfter = (title, colTitle) => {
    const out = []; let on = false, c = -1;
    for (let r = 0; r < rows.length; r++) {
      const n = norm(rowText(r));
      if (n.startsWith('■' + title)) { on = true; continue; }
      if (!on) continue;
      if (n.startsWith('■')) break;
      if (isNoise(r)) continue;
      if (norm(get(r, 1)) === '연번') { nonEmpty(r).forEach(([i, v]) => { if (norm(v) === colTitle) c = i; }); continue; }
      if (c >= 0 && /^\d/.test(get(r, 1)) && get(r, c)) out.push(get(r, c));
    }
    return out;
  };
  plan.teachers = listAfter('훈련교·강사', '훈련교·강사');
  plan.rooms = listAfter('훈련시설', '훈련시설명');
  plan.totalHours = plan.subjects.reduce((a, s) => a + s.hours, 0);
  return plan;
}

function parseCodes(rows) {
  const c = { teachers: {}, rooms: [], subjects: {}, subjList: [] };
  let cat = '';
  rows.forEach((r, i) => {
    if (i === 0) return;
    const a = cellS(r[0]), b = cellS(r[1]), e = cellS(r[4]);
    if (a) cat = a;
    if (!b || !e || e === '-') return;
    if (cat === '강사') c.teachers[b.replace(/\[.*?\]/g, '').trim()] = e;
    else if (cat === '시설') c.rooms.push({ name: b, code: e });
    else if (cat === '교과목') {
      const m = b.match(/^\((.*?)\)\s*(.*)$/), name = (m ? m[2] : b).trim();
      c.subjects[norm(name)] = e; c.subjList.push({ name, code: e, kind: m ? m[1].trim() : '' });   // 전자시간표 불러오기에서 코드 → 이름
    }
  });
  return c;
}

/* ---------- 과목 표 (설정) ---------- */
function buildSubjects() {
  const p = S.plan;
  p.raw = { t: [...p.teachers], r: [...p.rooms] };
  S.subjects = p.subjects.map((s, i) => ({
    ...s, idx: i,
    teacher: s.teachers[0] || p.teachers[0] || '',
    tsplit: [{ t: s.teachers[0] || p.teachers[0] || '', h: s.hours }],
    code: '',
    method: (k => k === undefined ? '평가자 체크리스트' : p.methods[k])([norm(s.unit), norm(s.name)].find(k => k in p.methods)),
    half: i < Math.ceil(p.subjects.length / 2) ? 'am' : 'pm', color: COLORS[i % COLORS.length]
  }));
  S.order = S.subjects.map(s => s.idx); S.orderPm = null; S.trules = []; S.manT = {}; S.subst = {};
  canonAll();
}
const COLORS = ['#cfe2ff', '#ffe0c2', '#d4f0d4', '#f5d0e6', '#fff1b8', '#d9d2f5', '#c9eef0', '#f2d6c9', '#e2e8c0', '#f0c8c8'];
/* ---------- 코드 맞추기: 코드표가 기준 ----------
   전자시간표에는 코드표의 코드만 들어가므로, 운영계획서 이름을 코드표 항목에 맞춰 둠.
   찾는 순서: 직접 지정(S.cmap) → 이름 같음 → 정리한 이름 같음 → (강의실) 앞부분 같음 → (강의실) 건물·호수 같음
   직접 지정은 브라우저에 저장해 다음 과정에서도 그대로 씀 */
const CMAP_KEY = 'tt_codemap_v1';
const loadCmap = () => { try { const m = JSON.parse(localStorage.getItem(CMAP_KEY) || '{}'); return { t: m.t || {}, r: m.r || {}, s: m.s || {} }; } catch (e) { return { t: {}, r: {}, s: {} }; } };
const saveCmap = () => { try { localStorage.setItem(CMAP_KEY, JSON.stringify(S.cmap)); } catch (e) { /* 저장 못 해도 이번 작업엔 문제없음 */ } };
const kT = x => String(x || '').replace(/\[.*?\]/g, '').replace(/_.*$/, '').replace(/\s+/g, '');
const kR = x => String(x || '').replace(/^\s*\(\d+\)\s*/, '').replace(/_.*$/, '').replace(/\(.*?\)/g, '').replace(/[\s·.,-]+/g, '');
const rNo = x => (String(x || '').replace(/^\s*\(\d+\)/, '').replace(/_.*$/, '').match(/\d+(?=\s*호)/g) || []).join('-');
const codeEntries = cat => {
  const C = S.codes; if (!C) return [];
  if (cat === 't') return Object.entries(C.teachers).map(([name, code]) => ({ name, code }));
  if (cat === 'r') return C.rooms;
  return C.subjList && C.subjList.length ? C.subjList : Object.entries(C.subjects).map(([name, code]) => ({ name, code }));
};
function findCode(cat, name) {
  if (!name || !S.codes) return null;
  const E = codeEntries(cat), man = ((S.cmap || {})[cat] || {})[name];
  if (man === '-') return null;                                                     // '코드표에 없음'으로 직접 정함
  if (man) { const e = E.find(x => x.code === man); if (e) return { e, how: '직접' }; }
  let e = E.find(x => x.name === name) || E.find(x => x.code === name);
  if (e) return { e, how: '같음' };
  if (cat === 't') { e = E.filter(x => kT(x.name) === kT(name)); return e.length === 1 ? { e: e[0], how: '자동' } : null; }
  if (cat === 's') { e = E.filter(x => norm(x.name) === norm(name)); return e.length === 1 ? { e: e[0], how: '자동' } : null; }
  const k = kR(name); if (!k) return null;
  e = E.filter(x => kR(x.name) === k); if (e.length === 1) return { e: e[0], how: '자동' };
  e = E.filter(x => kR(x.name).startsWith(k) || k.startsWith(kR(x.name))); if (e.length === 1) return { e: e[0], how: '자동' };
  const n = rNo(name); if (n) { e = E.filter(x => rNo(x.name) === n); if (e.length === 1) return { e: e[0], how: '자동(호수)' }; }
  return null;
}
const canonName = (cat, name) => { const f = findCode(cat, name); return f ? f.e.name : name; };
const subjCodeOf = s => { const f = findCode('s', s.name) || (s.unit && s.unit !== s.name ? findCode('s', s.unit) : null); return f ? f.e.code : ''; };
// 이름 하나를 앱 전체에서 다른 이름으로 바꿈 (운영계획서 이름 → 코드표 이름)
function replaceName(cat, from, to) {
  if (!from || from === to) return;
  const sw = v => v === from ? to : v, uniq = a => [...new Set(a.map(sw))].filter(Boolean);
  if (cat === 't') {
    if (S.plan) S.plan.teachers = uniq(S.plan.teachers || []);
    (S.subjects || []).forEach(s => {
      s.teachers = uniq(s.teachers || []); s.teacher = sw(s.teacher);
      if (s.tsplit) { const m = []; s.tsplit.forEach(x => { x.t = sw(x.t); const o = m.find(y => y.t === x.t); if (o) o.h += +x.h || 0; else m.push(x); }); s.tsplit = m; }
    });
    Object.keys(S.lt || {}).forEach(k => S.lt[k] = sw(S.lt[k]));
    (S.trules || []).forEach(r => r.t = sw(r.t));
  } else {
    if (S.plan) S.plan.rooms = uniq(S.plan.rooms || []);
    (S.subjects || []).forEach(s => { if (s.room) s.room = sw(s.room); });
    Object.keys(S.lr || {}).forEach(k => S.lr[k] = sw(S.lr[k]));
    const rv = $('#room').value; if (rv === from) { if (![...$('#room').options].some(o => o.value === to)) $('#room').insertAdjacentHTML('beforeend', `<option>${esc(to)}</option>`); $('#room').value = to; }
  }
}
// 운영계획서에서 읽은 이름(raw)을 모두 코드표 이름으로 통일 + 교과목 코드 다시 찾기
function canonAll() {
  if (!S.plan || !S.codes) return;
  const raw = S.plan.raw || (S.plan.raw = { t: [...(S.plan.teachers || [])], r: [...(S.plan.rooms || [])] });
  const usedT = new Set([...(raw.t || []), ...(S.subjects || []).flatMap(s => [...(s.teachers || []), ...splitOf(s).map(x => x.t)])]);
  const usedR = new Set([...(raw.r || []), ...(S.subjects || []).map(s => s.room).filter(Boolean)]);
  usedT.forEach(n => replaceName('t', n, canonName('t', n)));
  usedR.forEach(n => replaceName('r', n, canonName('r', n)));
  (S.subjects || []).forEach(s => s.code = subjCodeOf(s) || s.code || '');
}
const roomCode = name => { const f = findCode('r', name); return f ? f.e.code : ''; };
const allRooms = () => [...new Set([...(S.plan?.rooms || []), ...(S.codes?.rooms || []).map(r => r.name)])].filter(Boolean);
const allTeachers = () => [...new Set([...(S.plan?.teachers || []), ...Object.keys(S.codes?.teachers || {})])].filter(Boolean);
const subjRoom = s => (s && s.room) || $('#room').value;   // 과목에 강의실을 따로 정하지 않으면 과정 기본 강의실
// 출력용시간표 등에 쓰는 짧은 강의실 이름: 앞의 (23) 번호, 뒤의 _코드, 괄호 설명을 뺌 (다 빠지면 원래 이름)
const roomShort = name => { const o = String(name || '').trim(), v = o.replace(/^\(\d+\)\s*/, '').replace(/_.*$/, '').replace(/\s*\(.*?\)\s*/g, ' ').replace(/\s+/g, ' ').trim(); return v || o; };
const teacherCode = n => { const f = findCode('t', n); return f ? f.e.code : ''; };
/* ---------- 한 과목을 여러 교사(최대 3명)가 나눠 맡기 ---------- */
const MAX_T = 3;
const splitOf = s => (s.tsplit && s.tsplit.length ? s.tsplit : [{ t: s.teacher, h: s.hours }]);
const lessonT = k => (S.lt && S.lt[k]) || (S.subjects[S.lessons[k]] || {}).teacher || '';
const evenSplit = (H, n) => [...Array(n)].map((_, j) => Math.floor(H / n) + (j < H % n ? 1 : 0));
const lessonKeysSorted = () => Object.keys(S.lessons).sort((a, b) => a.split('|')[0] === b.split('|')[0] ? +a.split('|')[1] - +b.split('|')[1] : a < b ? -1 : 1);
// 과목의 편성 칸을 날짜·교시 순으로 훑으며 1번 교사 시간 → 2번 교사 시간 → 3번 교사 시간 순으로 배정
function assignTeachers(only) {
  S.lt = S.lt || {};
  const keys = lessonKeysSorted();
  S.subjects.forEach(s => {
    if (only !== undefined && s.idx !== only) return;
    const sp = splitOf(s); let j = 0, used = 0;
    keys.filter(k => S.lessons[k] === s.idx).forEach(k => {
      while (j < sp.length - 1 && used >= (+sp[j].h || 0)) { j++; used = 0; }
      S.lt[k] = sp[j].t; used++;
    });
  });
}
function syncTeacher(s) { s.tsplit = splitOf(s); s.teacher = s.tsplit[0].t; }

function renderSetup() {
  const p = S.plan;
  $('#planInfo').innerHTML = `<b>${esc(p.name)}</b> · 교과목 ${p.subjects.length}개 · 총 ${p.totalHours}시간 · 교사 ${p.teachers.map(esc).join(', ') || '-'} · 시설 ${p.rooms.map(esc).join(', ') || '-'}`;
  const tlist = [...new Set([...p.teachers, ...Object.keys(S.codes?.teachers || {})])];
  const tsel = cur => (tlist.length ? tlist : [cur]).map(t => `<option ${t === cur ? 'selected' : ''}>${esc(t)}</option>`).join('');
  S.sel = S.sel || new Set();
  const rlist = allRooms();
  const rsel = cur => `<option value="" ${cur ? '' : 'selected'}>(과정 기본 강의실)</option>` + rlist.map(r => `<option ${r === cur ? 'selected' : ''}>${esc(r)}</option>`).join('');
  const bt = $('#bT').value, br = $('#bR').value;
  $('#bT').innerHTML = tlist.map(t => `<option>${esc(t)}</option>`).join('');
  $('#bR').innerHTML = rsel('');
  if (bt && tlist.includes(bt)) $('#bT').value = bt; if (br && rlist.includes(br)) $('#bR').value = br;
  $('#subjBody').innerHTML = S.subjects.map((s, i) => `<tr>
    <td><input type="checkbox" class="sChk" data-i="${i}" ${S.sel.has(i) ? 'checked' : ''}></td>
    <td><span class="sw" style="background:${s.color}"></span>${i + 1}</td><td>${esc(s.kind)}</td><td>${esc(s.name)}</td>
    <td>${esc(s.ncs ? s.unit : '')}</td><td class="num">${s.hours}</td>
    <td class="tcell">${teacherCell(s, i, tsel)}</td>
    <td><select data-i="${i}" class="rSel">${rsel(s.room || '')}</select></td>
    <td class="split-only"><select data-i="${i}" class="hSel"><option value="am" ${s.half === 'am' ? 'selected' : ''}>오전</option><option value="pm" ${s.half === 'pm' ? 'selected' : ''}>오후</option><option value="both" ${s.half === 'both' ? 'selected' : ''}>나눠서</option></select>${s.half === 'both'
      ? `<div class="amh">오전 <input type="number" min="0" max="${s.hours}" value="${amH(s)}" data-i="${i}" class="aH">시간<br><small>오후 ${s.hours - amH(s)}시간</small></div>` : ''}</td>
    <td class="code ${s.code ? '' : 'miss'}">${s.code ? esc(s.code) : '코드 없음'}</td></tr>`).join('');
  bindTeacherCells(tlist);
  $$('.rSel').forEach(e => e.onchange = () => { S.subjects[+e.dataset.i].room = e.value; checkCodes(); if (S.settings) refreshAll(false); });
  $$('.sChk').forEach(e => e.onchange = () => { e.checked ? S.sel.add(+e.dataset.i) : S.sel.delete(+e.dataset.i); $('#bAll').checked = S.sel.size === S.subjects.length; });
  $('#bAll').checked = S.subjects.length > 0 && S.sel.size === S.subjects.length;
  $$('.hSel').forEach(e => e.onchange = () => { const s = S.subjects[+e.dataset.i]; s.half = e.value; if (e.value === 'both' && !(s.amH >= 0)) s.amH = Math.ceil(s.hours / 2); renderSetup(); });
  $$('.aH').forEach(e => e.onchange = () => { const s = S.subjects[+e.dataset.i]; s.amH = Math.max(0, Math.min(s.hours, Math.round(+e.value || 0))); renderSetup(); });
  const rv = $('#room').value;
  $('#room').innerHTML = (p.rooms.length ? p.rooms : ['']).map(r => `<option>${esc(r)}</option>`).join('');
  if (rv && p.rooms.includes(rv)) $('#room').value = rv;
  $('#setup').hidden = false;
  checkCodes();
  if (typeof renderOrder === 'function') renderOrder();
}
/* ---------- 교사·강의실 일괄 등록: 체크한 과목에 한 번에 넣기 ---------- */
function bulkApply(kind) {
  if (!S.sel || !S.sel.size) return msg('먼저 왼쪽 칸에서 과목을 체크해 주세요. (맨 위 체크 = 전체 선택)');
  const n = S.sel.size, val = kind === 't' ? $('#bT').value : ($('#bR').value || '(과정 기본)');
  S.sel.forEach(i => {
    const s = S.subjects[i];
    if (kind === 't') { s.tsplit = [{ t: $('#bT').value, h: s.hours }]; syncTeacher(s); }
    else s.room = $('#bR').value;
  });
  renderSetup();
  if (kind === 't' && S.settings && Object.keys(S.lessons).length) S.sel.forEach(i => assignTeachers(i));
  if (S.settings) refreshAll(false);
  msg(`체크한 과목 ${n}개에 ${kind === 't' ? '교사' : '강의실'} '${val}'을(를) 넣었습니다.`);
}
function teacherCell(s, i, tsel) {
  const sp = splitOf(s), H = s.hours, sum = sp.reduce((a, x) => a + (+x.h || 0), 0);
  const rows = sp.map((x, j) => `<div class="trow"><select data-i="${i}" data-j="${j}" class="tSel">${tsel(x.t)}</select>${sp.length > 1
    ? `<input type="number" min="0" max="${H}" value="${x.h}" data-i="${i}" data-j="${j}" class="tH">시간<button type="button" class="tDel" data-i="${i}" data-j="${j}" title="이 교사 빼기">×</button>` : ''}</div>`).join('');
  const add = sp.length < MAX_T ? `<button type="button" class="tAdd" data-i="${i}">${sp.length === 1 ? '+ 교사 나누기' : '+ 교사 추가'}</button>` : '';
  const chk = sp.length > 1 ? `<span class="tsum ${sum === H ? 'ok' : 'bad'}">합계 ${sum}/${H}시간</span>` : '';
  return rows + `<div class="tfoot">${add}${chk}</div>`;
}
function bindTeacherCells(tlist) {
  const changed = i => { const s = S.subjects[i]; syncTeacher(s); renderSetup(); if (S.settings && Object.keys(S.lessons).length) { assignTeachers(i); refreshAll(false); } };
  $$('.tSel').forEach(e => e.onchange = () => { const s = S.subjects[+e.dataset.i]; s.tsplit = splitOf(s).map(x => ({ ...x })); s.tsplit[+e.dataset.j].t = e.value; changed(+e.dataset.i); });
  $$('.tH').forEach(e => e.onchange = () => {
    const i = +e.dataset.i, j = +e.dataset.j, s = S.subjects[i], sp = splitOf(s).map(x => ({ ...x })), H = s.hours;
    const v = Math.max(0, Math.min(H, Math.round(+e.value || 0))); sp[j].h = v;
    // 합계가 총시간이 되도록 다음 교사 시간을 맞춤 (안 되면 합계를 빨간색으로 표시)
    const o = (j + 1) % sp.length, rest = H - sp.reduce((a, x, q) => q === o ? a : a + (+x.h || 0), 0);
    if (rest >= 0) sp[o].h = rest;
    s.tsplit = sp; changed(i);
  });
  $$('.tDel').forEach(e => e.onclick = () => {
    const i = +e.dataset.i, s = S.subjects[i], sp = splitOf(s).filter((_, j) => j !== +e.dataset.j);
    const hs = evenSplit(s.hours, sp.length); s.tsplit = sp.map((x, j) => ({ t: x.t, h: hs[j] })); changed(i);
  });
  $$('.tAdd').forEach(e => e.onclick = () => {
    const i = +e.dataset.i, s = S.subjects[i], sp = splitOf(s).map(x => ({ ...x }));
    const have = new Set(sp.map(x => x.t));
    // 운영계획서에 그 과목 교사로 적힌 사람을 먼저, 없으면 전체 교사 목록에서
    const next = [...(s.teachers || []), ...tlist].find(t => t && !have.has(t)) || sp[0].t;
    sp.push({ t: next, h: 0 });
    const hs = evenSplit(s.hours, sp.length); s.tsplit = sp.map((x, j) => ({ t: x.t, h: hs[j] })); changed(i);
  });
}
function checkCodes() {
  const miss = [];
  if (!S.codes) { $('#codeWarn').textContent = '코드표를 넣으면 강사·시설·교과목 코드가 자동으로 붙습니다.'; return; }
  S.subjects.forEach(s => { if (!s.code) miss.push(`교과목 '${s.name}'`); splitOf(s).forEach(x => { if (!teacherCode(x.t)) miss.push(`교사 '${x.t}'`); }); });
  [...new Set([$('#room').value, ...S.subjects.map(subjRoom), ...Object.values(S.lr || {})])].forEach(r => { if (!roomCode(r)) miss.push(`시설 '${r}'`); });
  $('#codeWarn').innerHTML = (miss.length ? '⚠ 코드표에서 찾지 못한 항목: ' + [...new Set(miss)].map(esc).join(', ') + ' — <b>[코드 맞추기]</b>에서 직접 골라 주세요.' : '✔ 모든 코드를 찾았습니다.')
    + ` <button type="button" class="bbtn" id="btnCmap">코드 맞추기 ▾${miss.length ? ` <span class="cmapN">${[...new Set(miss)].length}</span>` : ''}</button>`;
  $('#btnCmap').onclick = () => { const el = $('#cmapPanel'); el.hidden = !el.hidden; renderCmap(); };
  renderCmap();
}

/* ---------- 편성 ---------- */
function readHolidays() {
  const h = {};
  $('#holidays').value.split('\n').forEach(l => { const m = l.trim().match(/^(\d{4}-\d{2}-\d{2})\s*(.*)$/); if (m) h[m[1]] = m[2] || '휴일'; });
  return h;
}
function settings() {
  return {
    start: $('#start').value, lunchAfter: +$('#lunch').value,
    periods: $$('.pchk').filter(e => e.checked).map(e => +e.value).sort((a, b) => a - b),
    weekdays: $$('.wchk').filter(e => e.checked).map(e => +e.value),
    mode: $('#mode').value, spill: $('#spill').checked, extra: $('#extra').checked
  };
}
function buildDays() {
  const st = S.settings; S.days = [];
  const e0 = S.rangeEnd || S.endDate, e = addD(e0, Math.max(0, 5 - e0.getDay()));   // 마지막 주 금요일까지
  for (let d = parseD(st.start); d <= e; d = addD(d, 1)) if (d.getDay() >= 1 && d.getDay() <= 5) S.days.push(new Date(d));
}
function refreshAll(scroll) {
  if (!S.settings) return;
  recalcEnd();
  renderPreview(); renderTotals(); renderEval(); if (typeof renderOverview === 'function') renderOverview();
  $('#result').hidden = false;
  if (scroll) $('#result').scrollIntoView({ behavior: 'smooth' });
}
const extraPeriods = () => { const ps = S.settings.periods; if (S.settings.extra === false || !ps.length) return []; return [ps[0] - 1, ps[ps.length - 1] + 1].filter(p => p >= 1 && p <= 8 && !ps.includes(p)); };
const isSlot = (d, p) => S.lessons[dkey(d) + '|' + p] !== undefined || d >= parseD(S.settings.start) && S.settings.weekdays.includes(d.getDay()) && !S.holidays[dkey(d)] && (S.settings.periods.includes(p) || extraPeriods().includes(p));
const usedPeriods = () => { const set = new Set(S.settings.periods); Object.keys(S.lessons).forEach(k => set.add(+k.split('|')[1])); return [...set].sort((a, b) => a - b); };
const outDays = () => S.days.filter(d => d <= S.endDate);
// 출력용: 개강 주 월요일 ~ 수료 주 금요일의 모든 월~금
const weekDays = () => { const s0 = parseD(S.settings.start), out = []; let d = addD(s0, -((s0.getDay() + 6) % 7)); const e = addD(S.endDate, 5 - S.endDate.getDay());
  for (; d <= e; d = addD(d, 1)) if (d.getDay() >= 1 && d.getDay() <= 5) out.push(new Date(d)); return out; };
const teacherNames = () => { const set = new Set(); lessonKeysSorted().forEach(k => { const t = lessonT(k); if (t) set.add(t); }); return [...set]; };
function recalcEnd() { const ks = Object.keys(S.lessons).map(k => k.split('|')[0]).sort(); if (ks.length) S.endDate = parseD(ks[ks.length - 1]); }

function renderPreview() {
  const T = periodTable(S.settings.lunchAfter);
  const weeks = []; let cur = [];
  // 주차 달력은 항상 월~금: 개강 주 월요일부터 채움
  const pre = []; for (let d = addD(S.days[0], -((S.days[0].getDay() + 6) % 7)); d < S.days[0]; d = addD(d, 1)) pre.push(new Date(d));
  [...pre, ...S.days].forEach(d => { if (cur.length && d.getDay() === 1) { weeks.push(cur); cur = []; } cur.push(d); });
  if (cur.length) weeks.push(cur);
  const sDate = S.days[0], eDate = S.endDate;
  $('#summary').innerHTML = `훈련기간 <b>${dkey(sDate)} ~ ${dkey(eDate)}</b>${S.rangeEnd && S.endDate < S.rangeEnd ? ' <span class="bad">(옮긴 칸 때문에 수료일이 앞당겨짐)</span>' : ''} · 수업일 <b>${new Set(Object.keys(S.lessons).map(k => k.split('|')[0])).size}일</b> · 편성 <b>${Object.keys(S.lessons).length}시간</b>${(() => { const nl = Object.keys(S.lock || {}).filter(k => S.lock[k] && S.lessons[k] !== undefined).length, nc = (S.ref && S.ref.length) ? Object.keys(S.lessons).filter(k => lessonClash(k).length).length : 0;
    return (nl ? ` · 고정 <b>${nl}칸</b>` : '') + (S.ref && S.ref.length ? ` · 기준 시간표 ${S.ref.length}개와 겹침 <b class="${nc ? 'bad' : ''}">${nc}칸</b>` : ''); })()}`;
  $('#preview').innerHTML = weeks.map((w, wi) => {
    const s0 = parseD(S.settings.start);
    const head = w.map(d => `<th class="${S.holidays[dkey(d)] ? 'hol' : ''}${d < s0 ? ' pre' : ''}">${d.getMonth() + 1}/${d.getDate()}(${WD[d.getDay()]})${S.holidays[dkey(d)] ? '<br><small>' + esc(S.holidays[dkey(d)]) + '</small>' : ''}</th>`).join('');
    const ps = [...new Set([...usedPeriods(), ...extraPeriods()])].sort((a, b) => a - b), showLunch = ps[0] <= T.lunchAfter && ps[ps.length - 1] > T.lunchAfter;
    const ex = extraPeriods();
    const body = T.rows.filter(r => r.p === '점심' ? showLunch : ps.includes(r.p)).map(r => {
      if (r.p === '점심') return `<tr class="lunchrow"><td class="pc">점심<br><small>${r.disp}</small></td>${w.map(() => '<td></td>').join('')}</tr>`;
      return `<tr class="${ex.includes(r.p) ? 'extrarow' : ''}"><td class="pc">${r.p}교시${ex.includes(r.p) ? ' <em>여유</em>' : ''}<br><small>${r.disp}</small></td>${w.map(d => {
        const k = dkey(d) + '|' + r.p, v = S.lessons[k];
        if (!isSlot(d, r.p)) return `<td class="off"></td>`;
        if (v === undefined) return `<td class="slot empty${ex.includes(r.p) ? ' extra' : ''}" data-k="${k}"></td>`;
        const s = S.subjects[v];
        const t = lessonT(k), sp = splitOf(s), tj = sp.findIndex(x => x.t === t), R = lessonR(k), rm = R !== $('#room').value ? `<br><small class="rm">${esc(roomShort(R))}</small>` : '';
        const lk = isLocked(k), cl = lessonClash(k), un = isUnreg(s, t), sb = !!(S.subst && S.subst[k]);
        const tip = `${s.name} / ${t} / ${R}${lk ? '\n[고정 칸] 드래그로 움직이지 않습니다' : ''}${un ? '\n[등록 외 교사] 운영계획서에 이 과목 교사로 등록되지 않음' : ''}${sb ? '\n[대체] 기준 시간표와 겹쳐 다른 교사로 자동 대체됨' : ''}${cl.length ? '\n⚠ 겹침: ' + cl.map(c => c.kind + ' — ' + c.label).join('\n   ') : ''}\n(두 번 클릭하면 이 칸 교사 바꾸기)`;
        return `<td class="slot${lk ? ' locked' : ''}${cl.length ? ' clash' : ''}${un ? ' unreg' : ''}" data-k="${k}" draggable="${!lk}" style="background:${s.color}" title="${esc(tip)}">${lk ? '<span class="lk">고정</span>' : ''}${sb ? '<span class="sb">대체</span>' : ''}${esc(s.name)}<br><small class="${sp.length > 1 ? 'tm t' + Math.max(0, tj) : ''}">${esc(t)}</small>${rm}</td>`;
      }).join('')}</tr>`;
    }).join('');
    return `<table class="week" id="wk${wi}"><caption>${wi + 1}주차</caption><thead><tr><th></th>${head}</tr></thead><tbody>${body}</tbody></table>`;
  }).join('');
  bindDrag();
}
let dragK = null;
function bindDrag() {
  $$('#preview td.slot').forEach(td => {
    td.addEventListener('click', () => {
      if (!S.lockMode) return; const k = td.dataset.k; if (S.lessons[k] === undefined) return;
      S.lock = S.lock || {}; if (S.lock[k]) delete S.lock[k]; else S.lock[k] = 1; S.edited = true; refreshAll(false);
    });
    td.addEventListener('dragstart', e => { if (isLocked(td.dataset.k) || S.lockMode) { e.preventDefault(); return; } dragK = td.dataset.k; e.dataTransfer.effectAllowed = 'move'; td.classList.add('dragging'); });
    td.addEventListener('dragend', () => td.classList.remove('dragging'));
    td.addEventListener('dragover', e => { e.preventDefault(); td.classList.add('over'); });
    td.addEventListener('dragleave', () => td.classList.remove('over'));
    td.addEventListener('drop', e => {
      e.preventDefault(); td.classList.remove('over');
      const k2 = td.dataset.k; if (!dragK || dragK === k2) return;
      if (isLocked(k2)) { dragK = null; return msg('고정 칸에는 놓을 수 없습니다. [칸 고정하기]로 먼저 풀어 주세요.'); }
      S.lr = S.lr || {}; const ra = S.lr[dragK], rb = S.lr[k2];
      if (rb === undefined) delete S.lr[dragK]; else S.lr[dragK] = rb;
      if (ra === undefined) delete S.lr[k2]; else S.lr[k2] = ra;
      const a = S.lessons[dragK], b = S.lessons[k2], ta = S.lt[dragK], tb = S.lt[k2];
      if (b === undefined) delete S.lessons[dragK]; else S.lessons[dragK] = b;
      if (a === undefined) delete S.lessons[k2]; else S.lessons[k2] = a;
      if (tb === undefined) delete S.lt[dragK]; else S.lt[dragK] = tb;
      if (ta === undefined) delete S.lt[k2]; else S.lt[k2] = ta;
      ['manT', 'subst'].forEach(n => { const m = S[n] = S[n] || {}, x = m[dragK], y = m[k2]; if (y) m[dragK] = y; else delete m[dragK]; if (x) m[k2] = x; else delete m[k2]; });
      dragK = null; S.edited = true; refreshAll(false);
    });
    // 두 번 클릭: 이 칸만 교사 바꾸기 (그 과목 담당 교사 → 다른 교사 순)
    td.addEventListener('dblclick', () => {
      const k = td.dataset.k, v = S.lessons[k]; if (v === undefined) return;
      const s = S.subjects[v], cur = lessonT(k);
      const all = [...new Set([cur, ...regT(s), ...splitOf(s).map(x => x.t), ...(allowOut() ? allTeachers() : [])])].filter(Boolean);
      const sel = document.createElement('select'); sel.className = 'cellSel';
      sel.innerHTML = all.map(t => `<option ${t === cur ? 'selected' : ''}>${esc(t)}</option>`).join('');
      td.draggable = false; td.querySelector('small').replaceWith(sel); sel.focus();
      const done = () => { if (sel.value !== cur) { S.lt[k] = sel.value; (S.manT = S.manT || {})[k] = 1; if (S.subst) delete S.subst[k]; S.edited = true; } refreshAll(false); };
      sel.onchange = done; sel.onblur = done;
    });
  });
}
function renderTotals() {
  const cnt = {}, ct = {}, tt = {};
  Object.keys(S.lessons).forEach(k => { const v = S.lessons[k], t = lessonT(k); cnt[v] = (cnt[v] || 0) + 1; ct[v + '|' + t] = (ct[v + '|' + t] || 0) + 1; tt[t] = (tt[t] || 0) + 1; });
  const dd = d => `<td class="num ${d ? 'bad' : ''}">${d > 0 ? '+' + d : d}</td>`;
  $('#totals').innerHTML = `<tr><th>교과목 / 교사</th><th>계획</th><th>편성</th><th>과부족</th></tr>` + S.subjects.map(s => {
    const n = cnt[s.idx] || 0, sp = splitOf(s);
    let h = `<tr><td><span class="sw" style="background:${s.color}"></span>${esc(s.name)}</td><td class="num">${s.hours}</td><td class="num">${n}</td>${dd(n - s.hours)}</tr>`;
    const names = [...new Set([...(sp.length > 1 ? sp.map(x => x.t) : []), ...Object.keys(ct).filter(x => x.startsWith(s.idx + '|')).map(x => x.slice(String(s.idx).length + 1))])];
    if (names.length > 1) h += names.map(t => { const plan = sp.length > 1 ? sp.filter(x => x.t === t).reduce((a, x) => a + (+x.h || 0), 0) : (t === sp[0].t ? s.hours : 0), m = ct[s.idx + '|' + t] || 0;
      return `<tr class="sub"><td>└ ${esc(t)}</td><td class="num">${plan}</td><td class="num">${m}</td>${dd(m - plan)}</tr>`; }).join('');
    return h;
  }).join('');
  const tn = Object.keys(tt);
  $('#tTotals').innerHTML = `<tr><th>교사</th><th>맡은 교과목</th><th>편성</th></tr>` + tn.map(t => `<tr><td>${esc(t)}</td><td>${S.subjects.filter(s => ct[s.idx + '|' + t]).map(s => `${esc(s.name)} ${ct[s.idx + '|' + t]}`).join(', ')}</td><td class="num">${tt[t]}</td></tr>`).join('');
}

/* ---------- 평가일정 ---------- */
function evalPlan() {
  const keys = Object.keys(S.lessons).sort((a, b) => a.split('|')[0] === b.split('|')[0] ? +a.split('|')[1] - +b.split('|')[1] : a < b ? -1 : 1);
  const days = [...new Set(keys.map(k => k.split('|')[0]))].sort();
  const end = days[days.length - 1];
  const nextDay = k => days.find(x => x > k);
  const firstOnOrAfter = k => days.find(x => x >= k);
  const plus7 = k => dkey(addD(parseD(k), 7));
  return S.subjects.map(s => {
    const hrs = keys.filter(k => S.lessons[k] === s.idx).map(k => k.split('|')[0]);
    if (!hrs.length || s.method === '') return { s, none: true };   // 평가방법이 비어 있으면(예: 재량교과) 평가일정에서 뺌
    const first = hrs[0];
    let main = hrs[Math.max(0, Math.ceil(hrs.length * 0.8) - 1)];
    const rest = days.filter(x => x > main);
    if (rest.length < 2) main = days[Math.max(0, days.length - 3)];
    const step = prev => { const c = plus7(prev); if (c > end) return nextDay(prev) || prev; return firstOnOrAfter(c) || nextDay(prev) || prev; };
    let e2 = step(main), e3 = step(e2);
    if (!(main < e2 && e2 < e3)) {                 // 남은 날이 모자라 같은 날로 겹치면: 다음 수업일로 연달아
      e2 = nextDay(main); e3 = e2 && nextDay(e2);
      if (!e3) { main = days[Math.max(0, days.length - 3)]; e2 = days[Math.max(0, days.length - 2)]; e3 = days[days.length - 1]; }
    }
    // 평가 교사 = 본평가일에 그 과목을 가르치는 교사 (그날 수업이 없으면 그 과목을 가장 많이 맡은 교사)
    const sk = keys.filter(k => S.lessons[k] === s.idx), onMain = sk.filter(k => k.startsWith(main + '|'));
    const most = Object.entries(sk.reduce((a, k) => (a[lessonT(k)] = (a[lessonT(k)] || 0) + 1, a), {})).sort((a, b) => b[1] - a[1])[0];
    const teacher = onMain.length ? lessonT(onMain[onMain.length - 1]) : (most ? most[0] : s.teacher);
    return { s, first, main, e2, e3, hours: hrs.length, teacher };
  }).filter(x => !x.none).sort((a, b) => a.main < b.main ? -1 : a.main > b.main ? 1 : a.s.idx - b.s.idx);
}
const fmtK = k => { const d = parseD(k); return `${d.getMonth() + 1}/${d.getDate()}(${WD[d.getDay()]})`; };
function renderEval() {
  const ev = evalPlan();
  const clash = {}; ev.forEach(e => [e.main, e.e2, e.e3].forEach(k => clash[k] = (clash[k] || 0) + 1));
  $('#evalBody').innerHTML = ev.map(e => `<tr><td>${esc(e.s.name)}</td><td>${esc(e.teacher)}</td><td class="num">${e.hours}</td><td>${fmtK(e.first)}</td>${[e.main, e.e2, e.e3].map(k => `<td class="${clash[k] > 1 ? 'warn' : ''}">${fmtK(k)}</td>`).join('')}</tr>`).join('');
}

/* ---------- 내려받기 공통 ---------- */
function b64ToBytes(b) { const s = atob(b); const a = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) a[i] = s.charCodeAt(i); return a; }
async function loadTpl(name) { return JSZip.loadAsync(b64ToBytes(TPL[name])); }
async function download(zip, filename) {
  const blob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  saveBlob(blob, filename);
}
function saveBlob(blob, filename) {
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = filename.replace(/[\\/:*?"<>|]/g, '_');
  document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
}
const col = n => { let s = ''; while (n) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; };
const cStr = (ref, s, v) => v === '' || v == null ? `<c r="${ref}"${s ? ` s="${s}"` : ''}/>` : `<c r="${ref}"${s ? ` s="${s}"` : ''} t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`;
const cNum = (ref, s, v) => v === '' || v == null ? `<c r="${ref}"${s ? ` s="${s}"` : ''}/>` : `<c r="${ref}"${s ? ` s="${s}"` : ''}><v>${v}</v></c>`;
const courseTag = () => `${S.plan.name} (${dkey(S.days[0]).slice(2).replace(/-/g, '.')}-${dkey(S.endDate).slice(2).replace(/-/g, '.')})`;
function needResult() { if (!S.settings || !Object.keys(S.lessons).length) { msg('먼저 [자동 편성]을 눌러 주세요.'); return false; } return true; }

/* ---------- ① 전자시간표 (샘플 서식 그대로) ---------- */
async function dlElec() {
  if (!needResult()) return;
  const zip = await loadTpl('elec');
  let sh = await zip.file('xl/worksheets/sheet1.xml').async('string');
  let ss = await zip.file('xl/sharedStrings.xml').async('string');
  const head = sh.slice(0, sh.indexOf('<row r="3"')), tail = sh.slice(sh.indexOf('</sheetData>'));
  const sis = [...ss.matchAll(/<si>([\s\S]*?)<\/si>/g)].map(m => m[1].replace(/<[^>]+>/g, ''));
  const used = [...head.matchAll(/<v>(\d+)<\/v>/g)].map(m => +m[1]);
  const table = sis.slice(0, Math.max(...used) + 1).map(x => x.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&'));
  const idx = new Map(table.map((x, i) => [x, i]));
  const sid = v => { v = String(v); if (!idx.has(v)) { idx.set(v, table.length); table.push(v); } return idx.get(v); };
  const T = periodTable(S.settings.lunchAfter);
  const rows = [];
  outDays().forEach(d => {
    const ps = usedPeriods().filter(p => S.lessons[dkey(d) + '|' + p] !== undefined);
    if (!ps.length) return;
    const st = T.byP[ps[0]].s, en = hhmm(T.byP[ps[ps.length - 1]].s, 60);
    const items = ps.map(p => ({ t: T.byP[p].s, p }));
    if (ps[0] <= T.lunchAfter && ps[ps.length - 1] > T.lunchAfter) items.push({ t: T.lunch.s, lunch: true });
    items.sort((a, b) => a.t.localeCompare(b.t));
    items.forEach(it => {
      if (it.lunch) rows.push([ymd(d), st, en, 'N', it.t, '2', '', '', '']);
      else { const s = S.subjects[S.lessons[dkey(d) + '|' + it.p]]; rows.push([ymd(d), st, en, 'N', it.t, '1', teacherCode(lessonT(dkey(d) + '|' + it.p)), roomCode(lessonR(dkey(d) + '|' + it.p)), s.code]); }
    });
  });
  let n = 0;
  const body = rows.map((r, i) => `<row r="${i + 3}" spans="1:9" ht="40.5" x14ac:dyDescent="0.3">` + r.map((v, j) => { n++; return `<c r="${'ABCDEFGHI'[j]}${i + 3}" s="2" t="s"><v>${sid(v)}</v></c>`; }).join('') + '</row>').join('');
  let headX = head.replace(/<dimension ref="[^"]+"\/>/, `<dimension ref="A1:I${rows.length + 2}"/>`);
  const total = n + (head.match(/t="s"/g) || []).length;
  ss = ss.slice(0, ss.indexOf('<si>')).replace(/count="\d+"/, `count="${total}"`).replace(/uniqueCount="\d+"/, `uniqueCount="${table.length}"`)
    + table.map(x => `<si><t xml:space="preserve">${esc(x)}</t></si>`).join('') + '</sst>';
  zip.file('xl/worksheets/sheet1.xml', headX + body + tail);
  zip.file('xl/sharedStrings.xml', ss);
  await download(zip, `전자시간표_${courseTag()}.xlsx`);
}

/* ---------- ② 출력용시간표 ---------- */
async function dlPrint() {
  if (!needResult()) return;
  const zip = await loadTpl('print');
  const sh = await zip.file('xl/worksheets/sheet1.xml').async('string');
  const T = periodTable(S.settings.lunchAfter);
  const rowXml = r => sh.match(new RegExp(`<row r="${r}"[^>]*>[\\s\\S]*?</row>`))[0];
  let top = [1, 2, 3, 4, 5].map(rowXml).join('');
  top = top.replace(/(<c r="A1"[^>]*><is><t>)[\s\S]*?(<\/t>)/, `$1${esc(S.plan.name)} 연간훈련 세부계획표$2`)
    .replace(/(<c r="B2"[^>]*><is><t>)[\s\S]*?(<\/t>)/, `$1훈련기간 : ${dkey(S.days[0]).replace(/-/g, '.')} ~ ${dkey(S.endDate).replace(/-/g, '.')}$2`)
    .replace(/<c r="H2"( s="\d+")?[^>]*?(?:\/>|>[\s\S]*?<\/c>)/, m => cStr('H2', (m.match(/ s="(\d+)"/) || [])[1], `담당교사 : ${teacherNames().join(', ')}`));
  const tplRow = i => { const x = rowXml(6 + i); const st = {}; [...x.matchAll(/<c r="([A-Z]+)\d+" s="(\d+)"/g)].forEach(m => st[m[1]] = m[2]); return { ht: (x.match(/ht="([\d.]+)"/) || [])[1] || '27.95', st }; };
  const tpl = [...Array(9)].map((_, i) => tplRow(i));         // 틀은 4교시 뒤 점심형: 4번째(0부터)가 점심
  const styleFor = (pos, isLunch) => isLunch ? tpl[4] : (pos === 4 && S.settings.lunchAfter === 3 ? tpl[2] : tpl[pos]);
  // 수업 없는 날용 회색 서식: 블록 행에서 쓰는 서식마다 회색 채우기를 더한 복사본을 만듦
  let styles = await zip.file('xl/styles.xml').async('string');
  const fills = styles.match(/<fills count="(\d+)">([\s\S]*?)<\/fills>/);
  const grayFill = +fills[1];
  styles = styles.replace(fills[0], `<fills count="${grayFill + 1}">${fills[2]}<fill><patternFill patternType="solid"><fgColor rgb="FFEDEDED"/><bgColor indexed="64"/></patternFill></fill></fills>`);
  const xfsM = styles.match(/<cellXfs count="(\d+)">([\s\S]*?)<\/cellXfs>/);
  const xfs = xfsM[2].match(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g);
  const grayOf = {}; let extra = '';
  const used = new Set(tpl.flatMap(t => Object.values(t.st)));
  [...used].forEach(id => {
    let x = xfs[+id].replace(/fillId="\d+"/, `fillId="${grayFill}"`);
    x = /applyFill=/.test(x) ? x.replace(/applyFill="\d"/, 'applyFill="1"') : x.replace('<xf ', '<xf applyFill="1" ');
    grayOf[id] = String(xfs.length + Object.keys(grayOf).length); extra += x;
  });
  styles = styles.replace(xfsM[0], `<cellXfs count="${xfs.length + Object.keys(grayOf).length}">${xfsM[2]}${extra}</cellXfs>`);
  zip.file('xl/styles.xml', styles);
  const grayStyle = st => Object.fromEntries(Object.entries(st).map(([c, v]) => [c, c === 'A' ? v : grayOf[v]]));
  let r = 6, cum = 0, out = '', merges = [], breaks = [];
  const OD = weekDays();
  OD.forEach((d, di) => {
    const k = dkey(d), hol = S.holidays[k];
    const dayTot = usedPeriods().filter(p => S.lessons[k + '|' + p] !== undefined).length;
    cum += dayTot;
    T.rows.forEach((pr, pos) => {
      const isL = pr.p === '점심', rr = r + pos, sty = dayTot ? styleFor(pos, isL).st : grayStyle(styleFor(pos, isL).st);
      const v = isL ? undefined : S.lessons[k + '|' + pr.p];
      const s = v !== undefined ? S.subjects[v] : null;
      let c = '';
      c += cNum('A' + rr, sty.A, s ? 1 : '');
      c += pos === 0 ? cNum('B' + rr, sty.B, serial(d)) : cNum('B' + rr, sty.B, '');
      c += pos === 0 ? cStr('C' + rr, sty.C, WD[d.getDay()]) : cStr('C' + rr, sty.C, '');
      c += isL ? cStr('D' + rr, sty.D, '점심') : cNum('D' + rr, sty.D, pr.p);
      c += cStr('E' + rr, sty.E, pr.disp);
      c += cStr('F' + rr, sty.F, s ? s.kind : '');
      c += cStr('G' + rr, sty.G, s ? s.name : (pos === 0 && hol ? hol : ''));
      c += cStr('H' + rr, sty.H, s ? (s.ncs ? s.unit : s.name) : '');
      c += cStr('I' + rr, sty.I, s ? (s.ncs ? s.unit : (s.unitDetail || s.name)) : '');
      c += cStr('J' + rr, sty.J, s ? roomShort(lessonR(k + '|' + pr.p)) : '');
      c += cStr('K' + rr, sty.K, s ? lessonT(k + '|' + pr.p) : '');
      c += pos === 0 ? cNum('L' + rr, sty.L, dayTot || '') : cNum('L' + rr, sty.L, '');
      c += pos === 0 ? cNum('M' + rr, sty.M, dayTot ? cum : '') : cNum('M' + rr, sty.M, '');
      out += `<row r="${rr}" ht="${styleFor(pos, isL).ht}" customHeight="1">${c}</row>`;
    });
    ['B', 'C', 'L', 'M'].forEach(cl => merges.push(`${cl}${r}:${cl}${r + 8}`));
    r += 9;
    if (d.getDay() === 5 && di < OD.length - 1) breaks.push(r - 1);
  });
  const last = r - 1;
  const headMerges = [...sh.matchAll(/<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"\/>/g)].filter(m => +m[4] <= 5).map(m => m[0]);
  let x = sh.replace(/<sheetData>[\s\S]*<\/sheetData>/, '<sheetData>' + top + out + '</sheetData>')
    .replace(/<mergeCells[^>]*>[\s\S]*?<\/mergeCells>/, `<mergeCells count="${headMerges.length + merges.length}">${headMerges.join('')}${merges.map(m => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>`)
    .replace(/<dimension ref="[^"]+"\/>/, `<dimension ref="A1:M${last}"/>`);
  x = x.replace(/(<pageSetup[^>]*\/>)/, `$1<headerFooter><oddFooter>&amp;C&amp;"맑은 고딕,굵게"&amp;14&amp;P주차 / 전체 &amp;N주</oddFooter></headerFooter>`);
  if (breaks.length) x = x.replace('</worksheet>', `<rowBreaks count="${breaks.length}" manualBreakCount="${breaks.length}">${breaks.map(b => `<brk id="${b}" max="16383" man="1"/>`).join('')}</rowBreaks></worksheet>`);
  zip.file('xl/worksheets/sheet1.xml', x);
  const sheetName = `출력용시간표(${S.settings.lunchAfter}교시)`;
  let wb = await zip.file('xl/workbook.xml').async('string');
  wb = wb.replace(/name="출력용시간표\(\d교시\)"/, `name="${sheetName}"`).replace(/'출력용시간표\(\d교시\)'!\$1:\$5/, `'${sheetName}'!$1:$5`)
    .replace(/'출력용시간표\(\d교시\)'!\$A\$1:\$M\$\d+/, `'${sheetName}'!$A$1:$M$${last}`);
  zip.file('xl/workbook.xml', wb);
  await download(zip, `출력용시간표_${courseTag()}.xlsx`);
}

/* ---------- ③ 평가일정 (평가날출력-4번 B~W) ---------- */
async function dlEval() {
  if (!needResult()) return;
  const zip = await loadTpl('eval');
  const sh = await zip.file('xl/worksheets/sheet1.xml').async('string');
  const rowXml = r => sh.match(new RegExp(`<row r="${r}"[^>]*>[\\s\\S]*?</row>`))[0];
  const st5 = {}; [...rowXml(5).matchAll(/<c r="([A-Z]+)5" s="(\d+)"/g)].forEach(m => st5[m[1]] = m[2]);
  const st3 = (rowXml(3).match(/s="(\d+)"/) || [])[1];
  const y = S.days[0].getFullYear(), teachers = teacherNames();
  let r2 = rowXml(2).replace(/(<t>)[\s\S]*?(<\/t>)/, `$1${y}년   훈 련 실 시 계 획 서$2`);
  const r3 = `<row r="3">${cStr('B3', st3, '과정번호')}${cStr('D3', st3, $('#courseNo').value)}${cStr('E3', st3, $('#courseType').value)}${cStr('F3', st3, S.plan.name)}${cStr('J3', st3, `훈련기간 : ${dkey(S.days[0]).replace(/-/g, '.')} ~ ${dkey(S.endDate).replace(/-/g, '.')} `)}${cStr('O3', st3, teachers[0] || '')}${cStr('R3', st3, teachers[1] || '')}</row>`;
  const body = evalPlan().map((e, i) => {
    const s = e.s, r = 5 + i, ncs = !!s.ncs;
    const vals = [
      ['B', 'n', i + 1], ['C', 'n', s.idx + 1], ['D', 's', s.code], ['E', 's', s.kind], ['F', 's', s.name],
      ['G', 's', norm(ncs ? s.unit : s.name)], ['H', 's', ncs ? `[${s.ncs}]` : `[${s.kind}]`], ['I', 's', ncs ? s.unit : s.name],
      ['J', 's', ncs && s.level ? `${s.level}수준` : '1수준'], ['K', 's', s.method], ['L', 's', ''], ['M', 'n', 0], ['N', 's', ''], ['O', 'n', 0],
      ['P', 'n', parseD(e.main).getMonth() + 1], ['Q', 'n', e.hours], ['R', 's', roomShort(subjRoom(s))], ['S', 's', e.teacher],
      ['T', 'n', serial(parseD(e.first))], ['U', 'n', serial(parseD(e.main))], ['V', 'n', serial(parseD(e.e2))], ['W', 'n', serial(parseD(e.e3))]];
    return `<row r="${r}">` + vals.map(([c, t, v]) => t === 'n' ? cNum(c + r, st5[c], v) : cStr(c + r, st5[c], v)).join('') + '</row>';
  }).join('');
  const x = sh.replace(/<sheetData>[\s\S]*<\/sheetData>/, '<sheetData>' + rowXml(1) + r2 + r3 + rowXml(4) + body + '</sheetData>')
    .replace(/<dimension ref="[^"]+"\/>/, `<dimension ref="A1:W${4 + evalPlan().length}"/>`);
  zip.file('xl/worksheets/sheet1.xml', x);
  await download(zip, `평가일정_${courseTag()}.xlsx`);
}

/* ---------- 저장 / 불러오기 ---------- */
function saveJson() {
  if (!needResult()) return;
  const data = { version: 1, plan: S.plan, codes: S.codes, subjects: S.subjects, settings: S.settings, holidays: $('#holidays').value, lessons: S.lessons, lt: S.lt, lr: S.lr || {}, lock: S.lock || {}, order: S.order, orderPm: S.orderPm, trules: S.trules || [], manT: S.manT || {}, subst: S.subst || {}, allowOut: allowOut(), autoSub: $('#autoSub').checked, cmap: S.cmap, fixedRows: S.fixedRows || [], fixedName: S.fixedName || '', rangeEnd: dkey(S.rangeEnd || S.endDate), room: $('#room').value, courseNo: $('#courseNo').value, courseType: $('#courseType').value };
  saveBlob(new Blob([JSON.stringify(data)], { type: 'application/json' }), `편성저장_${courseTag()}.json`);
}
async function loadJson(file) {
  const d = JSON.parse(await file.text());
  S.plan = d.plan; S.codes = d.codes; S.subjects = d.subjects; S.order = d.order;
  if (d.cmap) ['t', 'r', 's'].forEach(c => S.cmap[c] = { ...S.cmap[c], ...(d.cmap[c] || {}) });
  renderSetup();
  const st = d.settings;
  $('#start').value = st.start; $('#lunch').value = st.lunchAfter; $('#mode').value = st.mode; $('#spill').checked = st.spill;
  $$('.pchk').forEach(e => e.checked = st.periods.includes(+e.value)); $$('.wchk').forEach(e => e.checked = st.weekdays.includes(+e.value));
  $('#holidays').value = d.holidays; $('#room').value = d.room; $('#courseNo').value = d.courseNo || ''; $('#courseType').value = d.courseType || '일반';
  S.holidays = readHolidays(); S.settings = st; S.lessons = d.lessons; S.lt = d.lt || {}; S.lr = d.lr || {}; S.lock = d.lock || {};
  S.order = d.order; S.orderPm = d.orderPm; ensureOrder();
  S.trules = d.trules || []; S.manT = d.manT || {}; S.subst = d.subst || {}; $('#allowOut').checked = !!d.allowOut; $('#autoSub').checked = d.autoSub !== false; renderRules(); S.fixedRows = d.fixedRows || []; S.fixedName = d.fixedName || ''; renderFixedInfo();
  S.subjects.forEach(syncTeacher); if (!d.lt) assignTeachers();
  const ks = Object.keys(S.lessons).map(k => k.split('|')[0]).sort(); S.endDate = parseD(ks[ks.length - 1]); S.rangeEnd = d.rangeEnd ? parseD(d.rangeEnd) : S.endDate; $('#extra').checked = st.extra !== false;
  buildDays(); toggleMode(); refreshAll(true); msg('저장한 편성을 불러왔습니다.');
}

/* ---------- 화면 연결 ---------- */
function msg(t) { const m = $('#msg'); m.textContent = t; m.hidden = false; clearTimeout(msg.t); msg.t = setTimeout(() => m.hidden = true, 4500); }
function toggleMode() { document.body.classList.toggle('split', $('#mode').value === 'split'); }
async function onPlan(f) {
  try { S.plan = parsePlan(await readRows(f)); }
  catch (e) { return msg('운영계획서를 읽지 못했습니다: ' + e.message); }
  if (!S.plan.subjects.length) return msg('운영계획서에서 교과목을 찾지 못했습니다. HRD 통합심사 운영계획서 파일인지 확인해 주세요.');
  $('#planName').textContent = f.name; buildSubjects(); renderSetup(); S.settings = null; $('#result').hidden = true;
}
async function onCodes(f) {
  try { S.codes = parseCodes(await readRows(f)); } catch (e) { return msg('코드표를 읽지 못했습니다: ' + e.message); }
  $('#codeName').textContent = f.name;
  if (S.plan) { S.subjects.forEach(s => s.code = ''); canonAll(); renderSetup(); if (S.settings) refreshAll(false); }
  const n = unmatchedCount(); if (n) msg(`코드표에서 찾지 못한 이름이 ${n}개 있습니다. [코드 맞추기]에서 직접 골라 주세요.`);
}
function bindDrop(zone, input, fn) {
  zone.addEventListener('click', () => input.click());
  input.addEventListener('change', () => input.files[0] && fn(input.files[0]));
  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('over'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('over'));
  zone.addEventListener('drop', e => { e.preventDefault(); zone.classList.remove('over'); e.dataTransfer.files[0] && fn(e.dataTransfer.files[0]); });
}
window.addEventListener('DOMContentLoaded', () => {
  $('#holidays').value = DEFAULT_HOLIDAYS;
  S.cmap = loadCmap();
  bindDrop($('#dzPlan'), $('#fPlan'), onPlan);
  bindDrop($('#dzCodes'), $('#fCodes'), onCodes);
  $('#fJson').addEventListener('change', e => e.target.files[0] && loadJson(e.target.files[0]));
  $('#btnGen').onclick = () => { if (S.edited && !confirmBox()) return; S.edited = false; generate(); };
  $('#mode').onchange = toggleMode;
  $('#room').onchange = () => { checkCodes(); if (S.settings) refreshAll(false); };
  $('#btnElec').onclick = () => dlElec().catch(e => msg('오류: ' + e.message));
  $('#btnPrint').onclick = () => dlPrint().catch(e => msg('오류: ' + e.message));
  $('#btnEval').onclick = () => dlEval().catch(e => msg('오류: ' + e.message));
  $('#btnJson').onclick = saveJson;
  $('#btnShift').onclick = shiftStart;
  $('#bTgo').onclick = () => bulkApply('t');
  $('#bRgo').onclick = () => bulkApply('r');
  $('#bAll').onchange = e => { S.sel = new Set(e.target.checked ? S.subjects.map((_, i) => i) : []); $$('.sChk').forEach(c => c.checked = e.target.checked); };
  $('#btnLoad').onclick = () => $('#fJson').click();
});
function confirmBox() { return window.confirm('드래그로 고친 내용이 처음부터 다시 편성됩니다. 계속할까요?'); }
