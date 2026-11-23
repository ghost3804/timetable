/* =====================================================================
   2차 기능
   - 세부설정: 편성 순서(오전·오후) 드래그
   - 고정일정: 양식으로 넣은 칸은 먼저 박아 두고(자물쇠), 나머지만 순서대로
   - 기준 시간표: 다른 과정 편성(.json / 전자시간표 .xlsx)과 교사·강의실이 겹치지 않게
   ===================================================================== */
const lessonR = k => (S.lr && S.lr[k]) || subjRoom(S.subjects[S.lessons[k]]);
const isLocked = k => !!(S.lock && S.lock[k]);
const findSubject = name => {
  const n = norm(name).replace(/[()（）\[\]]/g, ''); if (!n) return -1;
  let i = S.subjects.findIndex(s => norm(s.name).replace(/[()（）\[\]]/g, '') === n || s.code === name.trim());
  if (i < 0) i = S.subjects.findIndex(s => { const m = norm(s.name).replace(/[()（）\[\]]/g, ''); return m.includes(n) || n.includes(m); });
  return i;
};

/* ---------- 세부설정: 편성 순서 ---------- */
function ensureOrder() {
  const n = S.subjects.length;
  if (!Array.isArray(S.order) || S.order.length !== n || new Set(S.order).size !== n) S.order = S.subjects.map(s => s.idx);
}
let dragO = null;
function renderOrder() {
  const el = $('#orderPanel'); if (!el || el.hidden || !S.subjects.length) return;
  ensureOrder();
  const st = settings(), split = st.mode === 'split';
  const amN = st.periods.filter(p => p <= st.lunchAfter).length, pmN = st.periods.length - amN;
  const card = (i, n) => { const s = S.subjects[i];
    return `<div class="ocard" draggable="true" data-i="${i}" style="background:${s.color}"><span class="og">⋮⋮</span><b>${n}.</b><span class="on">${esc(s.name)}</span><span class="oh">${s.hours}시간 · ${esc(s.teacher)}</span></div>`; };
  const list = (key, title, arr) => `<div class="ocol"><div class="ot">${title} <small>${arr.reduce((a, i) => a + S.subjects[i].hours, 0)}시간</small></div><div class="olist" data-h="${key}">${arr.map((i, n) => card(i, n + 1)).join('')}<div class="oend" data-h="${key}">여기에 놓으면 맨 뒤로</div></div></div>`;
  el.innerHTML = `<p class="hint" style="margin:0 0 8px">카드를 끌어서 순서를 바꾸세요. ${split ? '오전↔오후 칸 사이로 옮기면 그 시간대로 바뀝니다. ' : ''}바꾼 뒤 <b>[자동 편성]</b>을 누르면 이 순서대로 채웁니다. (고정일정 칸은 먼저 박히고 남은 시간만 채움)</p>
    <div class="ocols${split ? '' : ' one'}">${split
      ? list('am', `오전 (${amN}교시)`, S.order.filter(i => S.subjects[i].half !== 'pm')) + list('pm', `오후 (${pmN}교시)`, S.order.filter(i => S.subjects[i].half === 'pm'))
      : list('any', '편성 순서 (위에서부터 차례로)', S.order)}</div>`;
  const move = (i, beforeI, h) => {
    S.order = S.order.filter(x => x !== i);
    if (beforeI === undefined) {
      // 그 목록의 마지막 과목 뒤로
      const same = S.order.filter(x => h === 'any' || (h === 'pm') === (S.subjects[x].half === 'pm'));
      const last = same[same.length - 1];
      S.order.splice(last === undefined ? S.order.length : S.order.indexOf(last) + 1, 0, i);
    } else S.order.splice(S.order.indexOf(beforeI), 0, i);
    if (h !== 'any') S.subjects[i].half = h;
    renderSetup();
  };
  $$('#orderPanel .ocard').forEach(c => {
    c.ondragstart = e => { dragO = +c.dataset.i; e.dataTransfer.effectAllowed = 'move'; c.classList.add('dragging'); };
    c.ondragend = () => c.classList.remove('dragging');
    c.ondragover = e => { e.preventDefault(); c.classList.add('over'); };
    c.ondragleave = () => c.classList.remove('over');
    c.ondrop = e => { e.preventDefault(); e.stopPropagation(); if (dragO === null || dragO === +c.dataset.i) return; const i = dragO; dragO = null; move(i, +c.dataset.i, c.closest('.olist').dataset.h); };
  });
  $$('#orderPanel .olist').forEach(l => {
    l.ondragover = e => e.preventDefault();
    l.ondrop = e => { e.preventDefault(); if (dragO === null) return; const i = dragO; dragO = null; move(i, undefined, l.dataset.h); };
  });
}

/* ---------- 고정일정 ---------- */
const FIX_HEAD = ['날짜', '교시', '교과목', '교사', '강의실', '적용 과정', '비고'];
function xlsxDownload(aoa, filename, widths) {
  const ws = XLSX.utils.aoa_to_sheet(aoa); ws['!cols'] = widths.map(w => ({ wch: w }));
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, '고정일정');
  saveBlob(new Blob([XLSX.write(wb, { bookType: 'xlsx', type: 'array' })], { type: 'application/octet-stream' }), filename);
}
function dlFixedTemplate() {
  xlsxDownload([
    ['고정일정 양식 — 3행부터 한 줄에 하나씩 적으세요. 교시: 오전 / 오후 / 종일 / 1-4 / 5,6 / 3 처럼. 적용 과정을 비우면 모든 과정에 적용(과정명 일부나 과정번호를 적으면 그 과정만). 교사·강의실을 비우면 그 과목의 기본값.'],
    FIX_HEAD,
    ['2027-03-03', '오전', '산업안전보건', '', '', '', '예시: 9개 과정 공통'],
    ['2027-03-03', '5-7', '직업윤리', '김진호', '본관 204호_DS', '', '예시'],
    ['2027-03-10', '종일', '취업역량강화', '', '', '컴퓨터활용', '예시: 과정명에 "컴퓨터활용"이 들어간 과정만']
  ], '고정일정_양식.xlsx', [12, 9, 26, 10, 18, 18, 30]);
}
function dlFixedExport() {
  if (!S.lock || !Object.keys(S.lock).some(k => S.lock[k] && S.lessons[k] !== undefined)) return msg('고정(자물쇠) 칸이 없습니다. 미리보기에서 [칸 고정하기]로 칸을 고정하거나 고정일정을 넣어 주세요.');
  const rows = [], keys = lessonKeysSorted().filter(isLocked);
  keys.forEach(k => {
    const [d, p] = k.split('|'), s = S.subjects[S.lessons[k]], t = lessonT(k), r = (S.lr && S.lr[k]) || (s.room || '');
    const last = rows[rows.length - 1];
    if (last && last.d === d && last.i === S.lessons[k] && last.t === t && last.r === r && last.p2 === +p - 1) { last.p2 = +p; return; }
    rows.push({ d, i: S.lessons[k], t, r, p1: +p, p2: +p });
  });
  xlsxDownload([['고정일정 — ' + S.plan.name + ' 에서 내보냄. 다른 과정에 그대로 넣을 수 있어요.'], FIX_HEAD,
    ...rows.map(x => [x.d, x.p1 === x.p2 ? String(x.p1) : `${x.p1}-${x.p2}`, S.subjects[x.i].name, x.t, x.r, '', ''])], `고정일정_${courseTag()}.xlsx`, [12, 9, 26, 10, 18, 18, 30]);
}
const toKeyDate = v => {
  if (typeof v === 'number' && v > 20000) { const d = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 864e5); return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`; }
  const s = String(v ?? '').trim(); let m = s.match(/^(\d{4})\D+(\d{1,2})\D+(\d{1,2})/); if (m) return `${m[1]}-${pad(+m[2])}-${pad(+m[3])}`;
  m = s.match(/^(\d{4})(\d{2})(\d{2})$/); if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  return '';
};
async function onFixed(file) {
  const rows = await readRows(file);
  const h = rows.findIndex(r => r.some(c => cellS(c) === '날짜') && r.some(c => /교과목/.test(cellS(c))));
  if (h < 0) return msg('고정일정 양식의 머리글(날짜·교시·교과목…)을 찾지 못했습니다. [고정일정 양식 받기]로 받은 파일을 써 주세요.');
  const ci = name => rows[h].findIndex(c => cellS(c).replace(/\s/g, '').startsWith(name));
  const c = { d: ci('날짜'), p: Math.max(ci('교시'), ci('시간')), s: ci('교과목'), t: ci('교사'), r: ci('강의실'), g: ci('적용'), n: ci('비고') };
  const list = rows.slice(h + 1).map(r => ({ date: toKeyDate(r[c.d]), time: cellS(r[c.p]), subj: cellS(r[c.s]), teacher: c.t >= 0 ? cellS(r[c.t]) : '', room: c.r >= 0 ? cellS(r[c.r]) : '', course: c.g >= 0 ? cellS(r[c.g]) : '', note: c.n >= 0 ? cellS(r[c.n]) : '' }))
    .filter(x => x.date && x.subj && !/^예시/.test(x.note));
  S.fixedRows = list; S.fixedName = file.name;
  renderFixedInfo();
  msg(`고정일정 ${list.length}줄을 읽었습니다. [자동 편성]을 누르면 먼저 박아 두고 나머지를 채웁니다.`);
}
function fixedPeriods(spec, st) {
  const s = String(spec).replace(/\s/g, '').replace(/교시/g, '');
  if (/^오전/.test(s)) return st.periods.filter(p => p <= st.lunchAfter);
  if (/^오후/.test(s)) return st.periods.filter(p => p > st.lunchAfter);
  if (/^(종일|전일|하루)/.test(s) || s === '') return st.periods.slice();
  const out = [];
  s.split(',').forEach(part => { const m = part.match(/^(\d)(?:[-~](\d))?$/); if (m) for (let p = +m[1]; p <= +(m[2] || m[1]); p++) out.push(p); });
  return [...new Set(out)].filter(p => p >= 1 && p <= 8).sort((a, b) => a - b);
}
// 이 과정에 해당하는 고정일정 → { 칸키: {i, t, r} }, 경고 목록
function placeFixed(st) {
  const placed = {}, warn = [], hol = S.holidays, start = st.start;
  const me = norm(S.plan.name), no = norm($('#courseNo').value);
  (S.fixedRows || []).forEach(x => {
    if (x.course) { const g = norm(x.course); if (!(me.includes(g) || (no && no === g))) return; }
    const i = findSubject(x.subj);
    if (i < 0) { warn.push(`${x.date} '${x.subj}': 이 과정에 없는 교과목`); return; }
    const d = parseD(x.date);
    if (x.date < start) { warn.push(`${x.date} '${x.subj}': 개강일 전`); return; }
    if (d.getDay() === 0 || d.getDay() === 6 || hol[x.date]) { warn.push(`${x.date} '${x.subj}': 주말·공휴일`); return; }
    const ps = fixedPeriods(x.time, st);
    if (!ps.length) { warn.push(`${x.date} '${x.subj}': 교시 '${x.time}'를 읽지 못함`); return; }
    const s = S.subjects[i];
    ps.forEach(p => { const k = x.date + '|' + p; if (placed[k]) warn.push(`${x.date} ${p}교시: 고정일정이 두 번 적힘`); placed[k] = { i, t: x.teacher || s.teacher, r: x.room && x.room !== subjRoom(s) ? x.room : '' }; });
  });
  const cnt = {}; Object.values(placed).forEach(v => cnt[v.i] = (cnt[v.i] || 0) + 1);
  Object.entries(cnt).forEach(([i, n]) => { if (n > S.subjects[i].hours) warn.push(`'${S.subjects[i].name}': 고정 ${n}시간이 계획 ${S.subjects[i].hours}시간보다 많음`); });
  return { placed, warn };
}
function renderFixedInfo() {
  const el = $('#fixedInfo'); if (!el) return;
  const n = (S.fixedRows || []).length;
  el.innerHTML = n ? `고정일정 <b>${n}줄</b> (${esc(S.fixedName || '')}) <button type="button" class="tDel" id="fixClear" title="고정일정 빼기">×</button>` : '';
  const b = $('#fixClear'); if (b) b.onclick = () => { S.fixedRows = []; S.fixedName = ''; renderFixedInfo(); msg('고정일정을 뺐습니다. 이미 편성된 자물쇠 칸은 [칸 고정하기]로 풀 수 있어요.'); };
}

/* ---------- 기준 시간표 (다른 과정) ---------- */
// S.ref = [{ name, items: [{ d, s, e, t:[키], r:[키], label }] }]
const tKeys = t => [t, teacherCode(t), String(teacherCode(t) || '').split('_')[0]].filter(Boolean);
const rKeys = r => [r, roomCode(r)].filter(Boolean);
function refFromJson(d, fname) {
  const T = periodTable(d.settings.lunchAfter), items = [];
  const tc = n => (d.codes?.teachers || {})[n] || '';
  const rc = n => { const x = (d.codes?.rooms || []).find(r => r.name.startsWith(n) || n.startsWith(r.name)); return x ? x.code : ''; };
  Object.entries(d.lessons || {}).forEach(([k, v]) => {
    const [dk, p] = k.split('|'), s = d.subjects[v]; if (!s || !T.byP[p]) return;
    const t = (d.lt && d.lt[k]) || s.teacher || '', r = (d.lr && d.lr[k]) || s.room || d.room || '';
    items.push({ d: dk, s: T.byP[p].s, e: hhmm(T.byP[p].s, 50), t: [t, tc(t), tc(t).split('_')[0]].filter(Boolean), r: [r, rc(r)].filter(Boolean), label: `${d.plan.name} · ${s.name} (${t})` });
  });
  return { name: d.plan?.name || fname, file: fname, start: d.settings.start, items };
}
function refFromElec(rows, fname) {
  const items = [];
  rows.slice(2).forEach(r => {
    const dt = cellS(r[0]), tm = cellS(r[4]); if (!/^\d{8}$/.test(dt) || cellS(r[5]) !== '1' || !/^\d{4}$/.test(tm)) return;
    const t = cellS(r[6]), rc = cellS(r[7]);
    items.push({ d: `${dt.slice(0, 4)}-${dt.slice(4, 6)}-${dt.slice(6)}`, s: tm, e: hhmm(tm, 50), t: [t, t.split('_')[0]].filter(Boolean), r: [rc].filter(Boolean), label: `${fname.replace(/\.xlsx?$/, '')} (${t.split('_')[0]})` });
  });
  return { name: fname.replace(/^전자시간표_/, '').replace(/\.xlsx?$/, ''), file: fname, items };
}
async function onRefFiles(files) {
  S.ref = S.ref || []; let added = 0, skip = [];
  for (const f of files) {
    try {
      let ref = null;
      if (/\.json$/i.test(f.name)) {
        const d = JSON.parse(await f.text()); if (!d.lessons || !d.settings) { skip.push(f.name); continue; }
        if (S.plan && d.plan?.name === S.plan.name && d.settings.start === $('#start').value) { skip.push(f.name + '(지금 과정)'); continue; }
        ref = refFromJson(d, f.name);
      } else if (/\.xlsx?$/i.test(f.name)) {
        const rows = await readRows(f); if (!/훈련일자|일자/.test(rows.slice(0, 2).flat().join(''))) { skip.push(f.name); continue; }
        ref = refFromElec(rows, f.name);
      } else continue;
      if (!ref.items.length) { skip.push(f.name); continue; }
      S.ref = S.ref.filter(x => x.file !== ref.file); S.ref.push(ref); added++;
    } catch (e) { skip.push(f.name); }
  }
  buildRefIdx(); renderRefInfo();
  if (S.settings) refreshAll(false);
  msg(`기준 시간표 ${added}개를 넣었습니다.${skip.length ? ' 건너뜀: ' + skip.slice(0, 4).join(', ') + (skip.length > 4 ? ' …' : '') : ''}`);
}
function buildRefIdx() { S.refIdx = {}; (S.ref || []).forEach(r => r.items.forEach(it => (S.refIdx[it.d] = S.refIdx[it.d] || []).push(it))); }
// 그 시간에 교사나 강의실이 다른 과정과 겹치면 [{kind, label}]
function conflictAt(dk, s, t, r) {
  const list = (S.refIdx || {})[dk]; if (!list) return [];
  const e = hhmm(s, 50), tk = tKeys(t), rk = rKeys(r), out = [];
  list.forEach(it => {
    if (!(it.s < e && s < it.e)) return;
    if (tk.some(x => it.t.includes(x))) out.push({ kind: '교사', label: it.label });
    else if (rk.some(x => it.r.includes(x))) out.push({ kind: '강의실', label: it.label });
  });
  return out;
}
const lessonClash = k => { const [dk, p] = k.split('|'), T = periodTable(S.settings.lunchAfter); return T.byP[p] ? conflictAt(dk, T.byP[p].s, lessonT(k), lessonR(k)) : []; };
function renderRefInfo() {
  const el = $('#refInfo'); if (!el) return;
  const R = S.ref || [];
  el.innerHTML = R.length ? `<b>기준 시간표 ${R.length}개</b> — 같은 시간에 교사·강의실이 겹치지 않게 편성합니다: ` + R.map((r, i) => `<span class="rchip">${esc(r.name)} <small>${r.items.length}시간</small><button type="button" class="tDel" data-r="${i}" title="빼기">×</button></span>`).join(' ') + ` <button type="button" class="tDel" id="refClear">모두 빼기</button>` : '';
  $$('#refInfo button[data-r]').forEach(b => b.onclick = () => { S.ref.splice(+b.dataset.r, 1); buildRefIdx(); renderRefInfo(); if (S.settings) refreshAll(false); });
  const c = $('#refClear'); if (c) c.onclick = () => { S.ref = []; buildRefIdx(); renderRefInfo(); if (S.settings) refreshAll(false); };
}

/* ---------- 편성 (고정일정 → 순서대로, 기준 시간표 피하기) ---------- */
function generate() {
  const st = settings();
  if (!S.plan) return msg('운영계획서를 먼저 넣어 주세요.');
  if (!st.start) return msg('개강일을 입력해 주세요.');
  if (!st.periods.length) return msg('수업 교시를 하나 이상 골라 주세요.');
  if (!st.weekdays.length) return msg('수업 요일을 하나 이상 골라 주세요.');
  S.holidays = readHolidays(); ensureOrder();
  const T = periodTable(st.lunchAfter), avoid = !!(S.ref && S.ref.length) && $('#avoid').checked;
  const fx = placeFixed(st);
  const lessons = {}, lt = {}, lr = {}, lock = {};
  Object.entries(fx.placed).forEach(([k, v]) => { lessons[k] = v.i; lt[k] = v.t; if (v.r) lr[k] = v.r; lock[k] = 1; });
  // 과목별 남은 시간과 교사 차례 (고정일정에서 쓴 교사 시간은 빼고)
  const used = {}, byT = {};
  Object.values(fx.placed).forEach(v => { used[v.i] = (used[v.i] || 0) + 1; byT[v.i + '|' + v.t] = (byT[v.i + '|' + v.t] || 0) + 1; });
  const tq = {};
  S.subjects.forEach(s => {
    const n = Math.max(0, s.hours - (used[s.idx] || 0)), arr = [];
    splitOf(s).forEach(x => { for (let j = 0; j < Math.max(0, (+x.h || 0) - (byT[s.idx + '|' + x.t] || 0)); j++) arr.push(x.t); });
    while (arr.length < n) arr.push(splitOf(s)[splitOf(s).length - 1].t);
    tq[s.idx] = arr.slice(0, n);
  });
  const q = { any: [], am: [], pm: [] };
  S.order.forEach(i => { const s = S.subjects[i]; for (let k = 0; k < tq[i].length; k++) (st.mode === 'split' ? q[s.half === 'pm' ? 'pm' : 'am'] : q.any).push(i); });
  const ok = (dk, p, i) => !avoid || !conflictAt(dk, T.byP[p].s, tq[i][0] || S.subjects[i].teacher, subjRoom(S.subjects[i])).length;
  const pick = (Q, dk, p) => { const seen = new Set(); for (let j = 0; j < Q.length; j++) { if (seen.has(Q[j])) continue; seen.add(Q[j]); if (ok(dk, p, Q[j])) return Q.splice(j, 1)[0]; } return undefined; };
  let d = parseD(st.start), guard = 0, last = null, skipped = 0;
  const left = () => q.any.length + q.am.length + q.pm.length;
  while (left() && guard++ < 1500) {
    const dk = dkey(d);
    if (st.weekdays.includes(d.getDay()) && !S.holidays[dk]) {
      for (const p of st.periods) {
        const k = dk + '|' + p;
        if (lessons[k] !== undefined) continue;          // 고정 칸
        let v;
        if (st.mode === 'split') {
          const isAm = p <= st.lunchAfter;
          v = pick(isAm ? q.am : q.pm, dk, p);
          if (v === undefined && st.spill && !(isAm ? q.am : q.pm).length) v = pick(isAm ? q.pm : q.am, dk, p);
        } else v = pick(q.any, dk, p);
        if (v !== undefined) { lessons[k] = v; lt[k] = tq[v].shift() || S.subjects[v].teacher; last = new Date(d); }
        else if (left()) skipped++;
      }
    }
    d = addD(d, 1);
  }
  if (left()) return msg('편성 기간이 너무 깁니다. 교시·요일 설정이나 기준 시간표 겹침을 확인해 주세요.');
  const fxLast = Object.keys(fx.placed).map(k => k.split('|')[0]).sort().pop();
  if (fxLast && (!last || parseD(fxLast) > last)) last = parseD(fxLast);
  S.lessons = lessons; S.lt = lt; S.lr = lr; S.lock = lock; S.endDate = last; S.rangeEnd = last; S.settings = st;
  buildDays(); refreshAll(true);
  const notes = [];
  if (Object.keys(fx.placed).length) notes.push(`고정일정 ${Object.keys(fx.placed).length}칸을 먼저 넣었습니다`);
  if (avoid && skipped) notes.push(`기준 시간표와 겹쳐 비워 둔 칸 ${skipped}개`);
  if (fx.warn.length) notes.push('⚠ ' + fx.warn.slice(0, 3).join(' / ') + (fx.warn.length > 3 ? ` 외 ${fx.warn.length - 3}건` : ''));
  S.fixWarn = fx.warn;
  if (notes.length) msg(notes.join(' · '));
}

/* ---------- 개강일만 바꾸기 (고정 칸은 제자리, 나머지는 수업일 순서대로) ---------- */
function shiftStart() {
  if (!S.settings || !Object.keys(S.lessons).length) return msg('먼저 편성을 만들거나 저장한 편성(.json)을 불러와 주세요.');
  const ns = $('#start').value, st = S.settings, os = st.start;
  if (!ns) return msg('새 개강일을 입력해 주세요.');
  if (ns === os) return msg('개강일이 지금과 같습니다. 위의 개강일 칸을 새 날짜로 바꾼 뒤 눌러 주세요.');
  const isClass = (d, hol) => st.weekdays.includes(d.getDay()) && !hol[dkey(d)];
  const oldHol = S.holidays, lock = S.lock || {};
  const keys = lessonKeysSorted(), free = keys.filter(k => !lock[k]);
  const endO = [keys[keys.length - 1].split('|')[0], dkey(S.rangeEnd || S.endDate)].sort().pop();
  const ord = {}; let n = 0;
  for (let d = parseD(os); dkey(d) <= endO; d = addD(d, 1)) if (isClass(d, oldHol)) ord[dkey(d)] = n++;
  const miss = free.filter(k => ord[k.split('|')[0]] === undefined);
  if (miss.length) return msg('수업일이 아닌 날에 있는 칸이 있어 옮기지 못했습니다: ' + miss.slice(0, 3).join(', '));
  const newHol = readHolidays(), nd = []; let guard = 0;
  for (let d = parseD(ns); nd.length < n + 400 && guard++ < 3000; d = addD(d, 1)) if (isClass(d, newHol)) nd.push(dkey(d));
  const ps = [...new Set([...st.periods, ...free.map(k => +k.split('|')[1])])].sort((a, b) => a - b);
  const L = {}, Tt = {}, R = {}, K = {};
  keys.filter(k => lock[k]).forEach(k => { L[k] = S.lessons[k]; if (S.lt[k] !== undefined) Tt[k] = S.lt[k]; if (S.lr && S.lr[k]) R[k] = S.lr[k]; K[k] = 1; });
  // 새 자리: 수업일 순번·교시 그대로. 고정 칸과 부딪히면 다음 빈 칸으로 밀어냄
  const seq = []; nd.forEach(dk => ps.forEach(p => seq.push(dk + '|' + p)));
  const pos = Object.fromEntries(seq.map((k, j) => [k, j]));
  let floor = 0, bumped = 0;
  for (const k of free) {
    const [dk, p] = k.split('|'); let j = pos[nd[ord[dk]] + '|' + p];
    if (j === undefined) return msg('새 개강일로 옮길 자리를 찾지 못했습니다.');
    if (j < floor) j = floor;
    const j0 = j; while (j < seq.length && L[seq[j]] !== undefined) j++;
    if (j >= seq.length) return msg('새 개강일로 옮길 자리를 찾지 못했습니다.');
    if (j !== j0 || pos[nd[ord[dk]] + '|' + p] < floor) bumped++;
    const k2 = seq[j]; floor = j + 1;
    L[k2] = S.lessons[k]; if (S.lt[k] !== undefined) Tt[k2] = S.lt[k]; if (S.lr && S.lr[k]) R[k2] = S.lr[k];
  }
  const rOrd = Object.keys(ord).filter(x => x <= dkey(S.rangeEnd || S.endDate)).length - 1;
  S.lessons = L; S.lt = Tt; S.lr = R; S.lock = K; S.holidays = newHol; S.settings = { ...st, start: ns };
  recalcEnd(); S.rangeEnd = rOrd >= 0 && nd[rOrd] && parseD(nd[rOrd]) > S.endDate ? parseD(nd[rOrd]) : S.endDate;
  buildDays(); refreshAll(true);
  msg(`편성은 그대로 두고 개강일을 ${os} → ${ns}(으)로 옮겼습니다.${Object.keys(K).length ? ` 고정 칸 ${Object.keys(K).length}개는 제자리${bumped ? `, 고정 칸을 피해 밀린 칸 ${bumped}개` : ''}.` : ''} 수료일 ${dkey(S.endDate)}`);
}

/* ---------- 미리보기: 칸 고정 모드 ---------- */
S.lockMode = false;
function toggleLockMode() {
  S.lockMode = !S.lockMode;
  $('#btnLockMode').classList.toggle('on', S.lockMode);
  $('#btnLockMode').textContent = S.lockMode ? '칸 고정하기: 켜짐 (칸을 누르면 고정/풀기)' : '칸 고정하기';
  document.body.classList.toggle('lockmode', S.lockMode);
}
function bindExtra() {
  $('#btnOrder').onclick = () => { const el = $('#orderPanel'); el.hidden = !el.hidden; $('#btnOrder').classList.toggle('on', !el.hidden); renderOrder(); };
  $('#btnFixTpl').onclick = dlFixedTemplate;
  $('#btnFixLoad').onclick = () => $('#fFixed').click();
  $('#fFixed').addEventListener('change', e => { e.target.files[0] && onFixed(e.target.files[0]); e.target.value = ''; });
  $('#btnRef').onclick = () => $('#fRef').click();
  $('#btnRefDir').onclick = () => $('#fRefDir').click();
  $('#fRef').addEventListener('change', e => { onRefFiles([...e.target.files]); e.target.value = ''; });
  $('#fRefDir').addEventListener('change', e => { onRefFiles([...e.target.files].filter(f => /\.(json|xlsx)$/i.test(f.name) && !/^~\$/.test(f.name))); e.target.value = ''; });
  $('#btnLockMode').onclick = toggleLockMode;
  $('#btnFixExport').onclick = dlFixedExport;
  $('#mode').addEventListener('change', renderOrder);
  $('#lunch').addEventListener('change', renderOrder);
  $$('.pchk').forEach(e => e.addEventListener('change', renderOrder));
}
window.addEventListener('DOMContentLoaded', bindExtra);
