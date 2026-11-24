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

/* ---------- 교사 등록 여부 · 오전/오후 시간 ---------- */
// 운영계획서에 그 과목 훈련교·강사로 등록된 교사 (과목에 없으면 과정 전체 교사)
const regT = s => (s && s.teachers && s.teachers.length ? s.teachers : (S.plan?.teachers || [])).filter(Boolean);
const allowOut = () => !!($('#allowOut') && $('#allowOut').checked);
const isUnreg = (s, t) => { const r = regT(s); return !!t && r.length > 0 && !r.includes(t); };
// 오전·오후 나눠서: half = 'am' | 'pm' | 'both'(오전 amH시간 + 오후 나머지)
const amH = s => s.half === 'both' ? Math.max(0, Math.min(s.hours, Math.round(s.amH >= 0 ? +s.amH : Math.ceil(s.hours / 2)))) : s.half === 'pm' ? 0 : s.hours;
const pmH = s => s.hours - amH(s);
const inAm = s => s.half !== 'pm', inPm = s => s.half === 'pm' || s.half === 'both';

/* ---------- 세부설정: 편성 순서 (오전 목록 = S.order, 오후 목록 = S.orderPm) ---------- */
function ensureOrder() {
  const n = S.subjects.length, ok = a => Array.isArray(a) && a.length === n && new Set(a).size === n && a.every(i => i >= 0 && i < n);
  if (!ok(S.order)) S.order = S.subjects.map(s => s.idx);
  if (!ok(S.orderPm)) S.orderPm = S.order.slice();
}
let dragO = null;
function renderOrder() {
  const el = $('#orderPanel'); if (!el || el.hidden || !S.subjects.length) return;
  ensureOrder();
  const st = settings(), split = st.mode === 'split';
  const amP = st.periods.filter(p => p <= st.lunchAfter), pmP = st.periods.filter(p => p > st.lunchAfter);
  const rng = a => a.length ? (a.length > 1 ? `${a[0]}~${a[a.length - 1]}교시` : `${a[0]}교시`) : '없음';
  const hh = (key, i) => { const s = S.subjects[i]; return key === 'am' ? amH(s) : key === 'pm' ? pmH(s) : s.hours; };
  const card = (key, i, n) => { const s = S.subjects[i], both = split && s.half === 'both';
    const ctl = !split ? '' : both
      ? (key === 'am' ? `<label class="ohx">오전 <input type="number" class="oAm" data-i="${i}" min="0" max="${s.hours}" value="${amH(s)}" draggable="false">/${s.hours}시간</label>` : `<span class="ohx">오후 ${pmH(s)}/${s.hours}시간</span>`)
        + `<button type="button" class="obtn oJoin" data-i="${i}" data-h="${key}" title="나누지 않고 이 목록(${key === 'am' ? '오전' : '오후'})에만 넣기">합치기</button>`
      : `<button type="button" class="obtn oSplit" data-i="${i}" title="이 과목을 오전·오후에 나눠 넣기">나누기</button>`;
    return `<div class="ocard${both ? ' both' : ''}" draggable="true" data-i="${i}" style="background:${s.color}"><span class="og">⋮⋮</span><b>${n}.</b><span class="on">${esc(s.name)}${both ? ' <em>나눔</em>' : ''}</span>${both ? '' : `<span class="oh">${hh(key, i)}시간</span>`}<span class="oh">${esc(s.teacher)}</span>${ctl}</div>`; };
  const list = (key, title, arr) => `<div class="ocol"><div class="ot">${title} <small>· 합계 ${arr.reduce((a, i) => a + hh(key, i), 0)}시간</small></div><div class="olist" data-h="${key}">${arr.map((i, n) => card(key, i, n + 1)).join('')}<div class="oend" data-h="${key}">여기에 놓으면 맨 뒤로</div></div></div>`;
  el.innerHTML = split
    ? `<p class="hint" style="margin:0 0 8px">카드를 끌어 <b>뿌리는 순서</b>를 바꾸고, 오전↔오후 목록 사이로 옮기면 그 시간대로 바뀝니다. <b>[나누기]</b>를 누르면 한 과목을 오전 몇 시간 + 오후 나머지로 나눠 양쪽 목록에 넣습니다. [자동 편성]을 누르면 오전 띠는 오전 목록 순서대로, 오후 띠는 오후 목록 순서대로 채웁니다. (고정일정 칸은 먼저 박히고 남은 시간만 채움)</p>
    <div class="ocols">${list('am', `오전 하루 ${amP.length}시간 <small>(${rng(amP)})</small>`, S.order.filter(i => inAm(S.subjects[i]))) + list('pm', `오후 하루 ${pmP.length}시간 <small>(${rng(pmP)})</small>`, S.orderPm.filter(i => inPm(S.subjects[i])))}</div>`
    : `<p class="hint" style="margin:0 0 8px">지금은 편성 방식이 <b>시간순</b>이라 오전·오후 구분 없이 아래 순서대로 이어서 채웁니다. <button type="button" class="bbtn" id="oToSplit">오전·오후 순서 따로 정하기</button></p>
    <div class="ocols one">${list('any', '편성 순서 (위에서부터 차례로)', S.order)}</div>`;
  const move = (i, beforeI, h) => {
    const s = S.subjects[i];
    if (h === 'am' && s.half === 'pm') s.half = 'am';
    if (h === 'pm' && s.half === 'am') s.half = 'pm';
    const key = h === 'pm' ? 'orderPm' : 'order', inL = h === 'pm' ? inPm : h === 'am' ? inAm : () => true;
    let arr = S[key].filter(x => x !== i);
    if (beforeI === undefined) {
      // 그 목록의 마지막 과목 뒤로
      const same = arr.filter(x => inL(S.subjects[x])), last = same[same.length - 1];
      arr.splice(last === undefined ? arr.length : arr.indexOf(last) + 1, 0, i);
    } else arr.splice(arr.indexOf(beforeI), 0, i);
    S[key] = arr;
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
  // 카드 안 버튼·입력칸: 끌기와 섞이지 않게
  $$('#orderPanel .obtn, #orderPanel .oAm').forEach(x => { x.onmousedown = e => e.stopPropagation(); x.ondragstart = e => { e.preventDefault(); e.stopPropagation(); }; });
  $$('#orderPanel .oSplit').forEach(b => b.onclick = () => { const s = S.subjects[+b.dataset.i]; s.half = 'both'; if (!(s.amH >= 0) || s.amH >= s.hours || s.amH <= 0) s.amH = Math.ceil(s.hours / 2); renderSetup(); });
  $$('#orderPanel .oJoin').forEach(b => b.onclick = () => { S.subjects[+b.dataset.i].half = b.dataset.h; renderSetup(); });
  $$('#orderPanel .oAm').forEach(x => x.onchange = () => { const s = S.subjects[+x.dataset.i]; s.amH = Math.max(0, Math.min(s.hours, Math.round(+x.value || 0))); renderSetup(); });
  const ts = $('#oToSplit'); if (ts) ts.onclick = () => { $('#mode').value = 'split'; toggleMode(); renderSetup(); };
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
const rKeys = r => r ? [r, roomCode(r)].filter(Boolean) : [];
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
  if (st.mode === 'split') {
    // 고정일정이 오전/오후 어느 쪽에서 시간을 썼는지 빼고 나눔
    const usedAm = {}; Object.entries(fx.placed).forEach(([k, v]) => { if (+k.split('|')[1] <= st.lunchAfter) usedAm[v.i] = (usedAm[v.i] || 0) + 1; });
    const nAm = {}; S.subjects.forEach(s => { const n = tq[s.idx].length, uA = usedAm[s.idx] || 0;
      // 나눠서: 오전 몫에서 오전 고정 시간을 뺀 만큼 오전 목록, 나머지는 오후 목록
      nAm[s.idx] = s.half === 'both' ? Math.min(n, Math.max(0, amH(s) - uA)) : s.half === 'pm' ? 0 : n; });
    S.order.forEach(i => { if (inAm(S.subjects[i])) for (let k = 0; k < nAm[i]; k++) q.am.push(i); });
    S.orderPm.forEach(i => { if (inPm(S.subjects[i])) for (let k = 0; k < tq[i].length - nAm[i]; k++) q.pm.push(i); });
  } else S.order.forEach(i => { for (let k = 0; k < tq[i].length; k++) q.any.push(i); });
  const autoSub = avoid && $('#autoSub').checked, rules = liveRules();
  const ok = (dk, p, i) => {
    if (!avoid) return true;
    const s = S.subjects[i], t0 = T.byP[p].s;
    if (conflictAt(dk, t0, '', subjRoom(s)).length) return false;                 // 강의실이 겹치면 안 됨
    const r = matchRule(i, dk, p, st.lunchAfter, rules), t = r ? r.t : (tq[i][0] || s.teacher);
    if (!teacherClash(dk, t0, t)) return true;
    return autoSub && subCands(s, t).some(x => !teacherClash(dk, t0, x));          // 다른 등록 교사로 대체 가능하면 통과
  };
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
  S.manT = {}; const ar = applyRules();
  buildDays(); refreshAll(true);
  const notes = [];
  if (ar.nr) notes.push(`교사 규칙으로 ${ar.nr}칸 교사 지정`);
  if (ar.ns) notes.push(`겹쳐서 다른 교사로 대체 ${ar.ns}칸`);
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
  if (miss.length) return msg('수업일이 아닌 날(공휴일 목록·수업 요일 밖)에 있는 칸이 있어 옮기지 못했습니다: ' + miss.slice(0, 3).join(', ') + ' — 그 칸을 다른 날로 옮기거나, 실제로 수업한 날이면 공휴일 목록에서 지운 뒤 다시 눌러 주세요.');
  const newHol = readHolidays(), nd = []; let guard = 0;
  for (let d = parseD(ns); nd.length < n + 400 && guard++ < 3000; d = addD(d, 1)) if (isClass(d, newHol)) nd.push(dkey(d));
  const ps = [...new Set([...st.periods, ...free.map(k => +k.split('|')[1])])].sort((a, b) => a - b);
  const L = {}, Tt = {}, R = {}, K = {}, M = {}, oM = S.manT || {};
  keys.filter(k => lock[k]).forEach(k => { L[k] = S.lessons[k]; if (S.lt[k] !== undefined) Tt[k] = S.lt[k]; if (S.lr && S.lr[k]) R[k] = S.lr[k]; K[k] = 1; if (oM[k]) M[k] = 1; });
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
    L[k2] = S.lessons[k]; if (S.lt[k] !== undefined) Tt[k2] = S.lt[k]; if (S.lr && S.lr[k]) R[k2] = S.lr[k]; if (oM[k]) M[k2] = 1;
  }
  const rOrd = Object.keys(ord).filter(x => x <= dkey(S.rangeEnd || S.endDate)).length - 1;
  S.lessons = L; S.lt = Tt; S.lr = R; S.lock = K; S.manT = M; S.subst = {}; S.holidays = newHol; S.settings = { ...st, start: ns };
  recalcEnd(); S.rangeEnd = rOrd >= 0 && nd[rOrd] && parseD(nd[rOrd]) > S.endDate ? parseD(nd[rOrd]) : S.endDate;
  buildDays(); refreshAll(true);
  msg(`편성은 그대로 두고 개강일을 ${os} → ${ns}(으)로 옮겼습니다.${Object.keys(K).length ? ` 고정 칸 ${Object.keys(K).length}개는 제자리${bumped ? `, 고정 칸을 피해 밀린 칸 ${bumped}개` : ''}.` : ''} 수료일 ${dkey(S.endDate)}`);
}

/* ---------- 전자시간표 불러와서 수정하기 ----------
   이미 만든(또는 HRD에 올린) 전자시간표 .xlsx → 미리보기 편집기로 되살림
   - 점심 위치: 시간구분 '2' 행(11:50 → 3교시 뒤, 12:50 → 4교시 뒤) 또는 4교시 시작 시각으로 판단
   - 코드 → 이름: 코드표가 있으면 이름으로, 없으면 코드를 그대로 이름처럼 씀
   - 운영계획서가 이미 들어 있으면 교과목 코드로 그 과목에 맞춤(계획 시간·평가방법 유지) */
const elecCourseName = fname => fname.replace(/\.xlsx?$/i, '').replace(/^전자시간표_?/, '').replace(/\s*\(\d{2}\.\d{2}\.\d{2}-\d{2}\.\d{2}\.\d{2}\)\s*$/, '').trim() || '불러온 과정';
function readElec(rows) {
  const recs = [], other = { lunch: [], vac: 0, bad: [] };
  rows.slice(2).forEach((r, i) => {
    const dt = cellS(r[0]).replace(/\D/g, ''); if (dt.length !== 8) return;
    const tm = cellS(r[4]).replace(/\D/g, '').padStart(4, '0'), kind = cellS(r[5]);
    if (cellS(r[3]) && cellS(r[3]).toUpperCase() !== 'N') other.vac++;
    const d = `${dt.slice(0, 4)}-${dt.slice(4, 6)}-${dt.slice(6)}`;
    if (kind === '2') other.lunch.push(tm);
    else if (kind === '1') recs.push({ d, tm, t: cellS(r[6]), r: cellS(r[7]), s: cellS(r[8]), row: i + 3 });
  });
  return { recs, other };
}
async function onElecImport(file) {
  let rows; try { rows = await readRows(file); } catch (e) { return msg('전자시간표를 읽지 못했습니다: ' + e.message); }
  if (!/훈련일자/.test(rows.slice(0, 3).flat().map(cellS).join(''))) return msg('전자시간표 형식(2행에 훈련일자·시작시간·시간구분…)이 아닙니다.');
  const { recs, other } = readElec(rows);
  if (!recs.length) return msg('전자시간표에서 수업(시간구분 1) 행을 찾지 못했습니다.');
  if (S.settings && Object.keys(S.lessons).length && !window.confirm('지금 화면의 편성을 지우고 이 전자시간표 내용으로 바꿀까요?')) return;
  const warn = [];
  // 점심 위치
  const lt0 = other.lunch.sort((a, b) => other.lunch.filter(x => x === b).length - other.lunch.filter(x => x === a).length)[0];
  const times = new Set(recs.map(x => x.tm));
  const la = lt0 === '1150' ? 3 : lt0 === '1250' ? 4 : times.has('1240') ? 3 : 4;
  const T = periodTable(la), pOf = Object.fromEntries(Object.values(T.byP).map(x => [x.s, x.p]));
  // 코드표 (없으면 빈 표를 만들고 모르는 코드는 코드 그대로 등록 → 다시 내려받을 때 같은 코드가 나감)
  S.codes = S.codes || { teachers: {}, rooms: [], subjects: {}, subjList: [] };
  const C = S.codes, unk = { t: new Set(), r: new Set(), s: new Set() };
  const tName = code => { if (!code) return ''; let e = Object.entries(C.teachers).find(([, c]) => c === code);
    if (!e) e = Object.entries(C.teachers).find(([, c]) => String(c).split('_')[0] === code.split('_')[0]);
    if (e) return e[0]; unk.t.add(code); C.teachers[code] = code; return code; };
  const rName = code => { if (!code) return ''; const x = C.rooms.find(r => r.code === code); if (x) return x.name; unk.r.add(code); C.rooms.push({ name: code, code }); return code; };
  const sInfo = code => { const x = (C.subjList || []).find(s => s.code === code); if (x) return x;
    const e = Object.entries(C.subjects).find(([, c]) => c === code); if (e) return { name: e[0], code, kind: '' };
    unk.s.add(code); C.subjects[norm(code)] = code; return { name: code, code, kind: '' }; };
  // 칸 만들기
  const cells = [], skipT = new Set(), dup = [];
  const seen = {};
  recs.forEach(x => {
    const p = pOf[x.tm]; if (!p) { skipT.add(x.tm); return; }
    const k = x.d + '|' + p; if (seen[k]) { dup.push(k); return; } seen[k] = 1;
    cells.push({ k, sc: x.s, t: tName(x.t), r: rName(x.r) });
  });
  if (!cells.length) return msg('교시 시각을 읽지 못했습니다: ' + [...skipT].join(', '));
  // 과목: 운영계획서가 있으면 코드로 맞추고, 없거나 못 맞춘 코드는 새 과목으로
  const order = [...new Set(cells.map(c => c.sc))];
  const fromPlan = !!(S.plan && S.subjects.length && !S.plan.fromElec);
  const cnt = c => cells.filter(x => x.sc === c).length;
  if (!fromPlan) {
    S.plan = { name: elecCourseName(file.name), subjects: [], teachers: [], rooms: [], methods: {}, totalHours: 0, fromElec: true };
    S.subjects = [];
  }
  const idxOf = {};
  if (fromPlan) S.subjects.forEach(s => { if (s.code && idxOf[s.code] === undefined) idxOf[s.code] = s.idx; });
  const added = [];
  order.forEach(code => {
    if (idxOf[code] !== undefined) return;
    const inf = sInfo(code), i = S.subjects.length;
    const s = { kind: inf.kind || (fromPlan ? '추가' : ''), name: inf.name, hours: cnt(code), teachers: [], ncs: '', unit: inf.name, level: '', idx: i, code,
      teacher: '', tsplit: [], method: '평가자 체크리스트', half: 'am', color: COLORS[i % COLORS.length] };
    S.subjects.push(s); S.plan.subjects.push(s); idxOf[code] = i; if (fromPlan) added.push(inf.name);
  });
  // 과정 기본 강의실 = 가장 많이 쓴 강의실
  const most = arr => Object.entries(arr.reduce((a, v) => (v && (a[v] = (a[v] || 0) + 1), a), {})).sort((a, b) => b[1] - a[1]);
  const room0 = (most(cells.map(c => c.r))[0] || [''])[0];
  const tAll = most(cells.map(c => c.t)).map(x => x[0]), rAll = most(cells.map(c => c.r)).map(x => x[0]);
  S.plan.teachers = [...new Set([...(S.plan.teachers || []), ...tAll])];
  S.plan.rooms = [...new Set([room0, ...(S.plan.rooms || []), ...rAll])].filter(Boolean);
  // 과목별 교사(많이 맡은 순, 최대 3명)·강의실
  const used = new Set(order.map(c => idxOf[c]));
  S.subjects.forEach(s => {
    if (!used.has(s.idx)) return;
    const mine = cells.filter(c => idxOf[c.sc] === s.idx), ts = most(mine.map(c => c.t));
    if (ts.length) {
      const sp = ts.slice(0, MAX_T).map(([t, n]) => ({ t, h: n }));
      if (ts.length > MAX_T) sp[MAX_T - 1].h += ts.slice(MAX_T).reduce((a, x) => a + x[1], 0);
      s.tsplit = sp; s.teacher = sp[0].t; if (!s.teachers || !s.teachers.length) s.teachers = sp.map(x => x.t);
    }
    const r1 = (most(mine.map(c => c.r))[0] || [''])[0]; s.room = r1 && r1 !== room0 ? r1 : '';
  });
  if (!fromPlan) S.plan.totalHours = S.subjects.reduce((a, s) => a + s.hours, 0);
  if (!fromPlan) S.subjects.forEach((s, j) => s.half = j < Math.ceil(S.subjects.length / 2) ? 'am' : 'pm');   // 운영계획서와 같은 기본값: 앞쪽 절반 오전, 나머지 오후
  S.order = S.subjects.map(s => s.idx); S.orderPm = null; S.manT = {}; S.subst = {}; if (!fromPlan) S.trules = [];
  // 화면 조건 맞추기
  const days = [...new Set(cells.map(c => c.k.split('|')[0]))].sort();
  const perDay = {}; cells.forEach(c => { const [d, p] = c.k.split('|'); (perDay[d] = perDay[d] || new Set()).add(+p); });
  const pc = {}; Object.values(perDay).forEach(set => set.forEach(p => pc[p] = (pc[p] || 0) + 1));
  const periods = Object.keys(pc).map(Number).filter(p => pc[p] >= days.length / 2).sort((a, b) => a - b);
  const wds = new Set(days.map(d => parseD(d).getDay()));
  $('#start').value = days[0]; $('#lunch').value = String(la); $('#mode').value = 'seq';
  $$('.pchk').forEach(e => e.checked = (periods.length ? periods : Object.keys(pc).map(Number)).includes(+e.value));
  $$('.wchk').forEach(e => e.checked = wds.has(+e.value));
  S.sel = new Set(); renderSetup(); $('#room').value = room0; toggleMode();
  S.holidays = readHolidays(); S.settings = settings();
  // 편성 칸 채우기
  const L = {}, Tt = {}, R = {};
  cells.forEach(c => { const i = idxOf[c.sc], s = S.subjects[i]; L[c.k] = i; if (c.t) Tt[c.k] = c.t; if (c.r && c.r !== subjRoom(s)) R[c.k] = c.r; });
  S.lessons = L; S.lt = Tt; S.lr = R; S.lock = {}; recalcEnd(); S.rangeEnd = S.endDate; S.edited = true; S.elecName = file.name;
  buildDays(); checkCodes(); refreshAll(true);
  // 알림
  const hol = days.filter(d => S.holidays[d]);
  if (unk.s.size || unk.t.size || unk.r.size) warn.push(`코드표에 없는 코드는 코드 그대로 표시: ${[...unk.s, ...unk.t, ...unk.r].slice(0, 4).join(', ')}${unk.s.size + unk.t.size + unk.r.size > 4 ? ' …' : ''}`);
  if (added.length) warn.push(`운영계획서에 없는 교과목 코드 ${added.length}개를 과목으로 추가`);
  if (skipT.size) warn.push(`읽지 못한 시작 시각 ${[...skipT].join(', ')} (야간 등) 칸은 뺌`);
  if (dup.length) warn.push(`같은 날·같은 교시 중복 ${dup.length}칸은 첫 줄만 씀`);
  if (hol.length) warn.push(`공휴일 목록에 있는 날 수업 ${hol.length}일 (${hol.slice(0, 2).join(', ')})`);
  if (other.vac) warn.push(`방학/원격 'N'이 아닌 행 ${other.vac}개 — 다시 내려받으면 'N'으로 나감`);
  S.importWarn = warn;
  msg(`전자시간표를 불러왔습니다: ${days.length}일 · ${cells.length}시간 · 교과목 ${order.length}개 · 점심 ${la}교시 뒤${warn.length ? ' · ⚠ ' + warn.join(' / ') : ''}`);
  if (warn.length) setTimeout(() => msg('⚠ ' + warn.join(' / ')), 4600);
}

/* ---------- 교사 규칙 표 · 겹침 자동 대체 ----------
   규칙: { i: 과목 번호(-1 = 모든 과목), days: [요일 번호] (비면 모든 요일), h: 'all'|'am'|'pm', t: 교사 }
   위에서부터 처음 맞는 규칙 하나를 씀. 우선순위: 고정 칸 > 직접(두 번 클릭) 바꾼 칸 > 규칙 > 과목 기본 교사 */
const liveRules = () => (S.trules || []).filter(r => r.t && (r.i === -1 || S.subjects[r.i]));
function matchRule(i, dk, p, la, rules) {
  const wd = parseD(dk).getDay();
  return (rules || liveRules()).find(r => (r.i === -1 || r.i === i) && (!r.days || !r.days.length || r.days.includes(wd)) && (r.h === 'all' || (r.h === 'am') === (+p <= la)));
}
const teacherClash = (dk, t0, t) => !!t && conflictAt(dk, t0, t, '').length > 0;
const subCands = (s, cur) => [...new Set([...regT(s), ...splitOf(s).map(x => x.t), ...(allowOut() ? allTeachers() : [])])].filter(x => x && x !== cur);
// 지금 편성에 규칙을 적용하고, 기준 시간표와 교사가 겹치면 다른 교사로 대체
function applyRules() {
  const st = S.settings, rules = liveRules(), T = periodTable(st.lunchAfter);
  const autoSub = !!(S.ref && S.ref.length) && $('#autoSub').checked && $('#avoid').checked;
  S.subst = {}; S.manT = S.manT || {}; let nr = 0, ns = 0, nf = 0;
  lessonKeysSorted().forEach(k => {
    if (isLocked(k) || S.manT[k]) return;
    const [dk, p] = k.split('|'), i = S.lessons[k], s = S.subjects[i];
    const r = matchRule(i, dk, p, st.lunchAfter, rules);
    if (r) { if (S.lt[k] !== r.t) nr++; S.lt[k] = r.t; }
    if (autoSub && T.byP[p] && teacherClash(dk, T.byP[p].s, S.lt[k])) {
      const alt = subCands(s, S.lt[k]).find(x => !teacherClash(dk, T.byP[p].s, x));
      if (alt) { S.lt[k] = alt; S.subst[k] = 1; ns++; } else nf++;
    }
  });
  return { nr, ns, nf };
}
// [지금 편성에 규칙 적용]: 고정 칸·직접 바꾼 칸은 두고, 나머지는 과목 기본 교사로 되돌린 뒤 규칙·대체를 다시 적용
function reapplyRules() {
  if (!S.settings || !Object.keys(S.lessons).length) return msg('먼저 [자동 편성]을 하거나 편성을 불러와 주세요.');
  const keep = {}; Object.keys(S.lessons).forEach(k => { if (isLocked(k) || (S.manT || {})[k]) keep[k] = S.lt[k]; });
  assignTeachers(); Object.assign(S.lt, keep);
  const r = applyRules(); S.edited = true; refreshAll(false);
  msg(`규칙 적용: 교사 지정 ${r.nr}칸${r.ns ? ` · 겹쳐서 대체 ${r.ns}칸` : ''}${r.nf ? ` · 대체할 교사가 없어 겹친 채 남은 칸 ${r.nf}개` : ''} (고정 칸·직접 바꾼 칸은 그대로)`);
}
function renderRules() {
  const body = $('#ruleBody'); if (!body) return;
  const R = S.trules = S.trules || [];
  const subjOpt = cur => `<option value="-1"${cur === -1 ? ' selected' : ''}>(모든 과목)</option>` + S.subjects.map(s => `<option value="${s.idx}"${s.idx === cur ? ' selected' : ''}>${esc(s.name)}</option>`).join('');
  const tOpt = (i, cur) => { const s = S.subjects[i], base = i < 0 || allowOut() ? allTeachers() : [...regT(s), ...splitOf(s).map(x => x.t)];
    return [...new Set([...base, cur].filter(Boolean))].map(t => `<option${t === cur ? ' selected' : ''}>${esc(t)}</option>`).join(''); };
  body.innerHTML = R.length ? R.map((r, j) => `<tr>
    <td class="num">${j + 1}</td><td><select class="rS" data-j="${j}">${subjOpt(r.i)}</select></td>
    <td class="chks">${[1, 2, 3, 4, 5].map(d => `<label><input type="checkbox" class="rD" data-j="${j}" value="${d}" ${(r.days || []).includes(d) ? 'checked' : ''}>${WD[d]}</label>`).join('')}</td>
    <td><select class="rH" data-j="${j}">${[['all', '오전·오후'], ['am', '오전'], ['pm', '오후']].map(([v, n]) => `<option value="${v}"${r.h === v ? ' selected' : ''}>${n}</option>`).join('')}</select></td>
    <td><select class="rT" data-j="${j}">${tOpt(r.i, r.t)}</select>${r.i >= 0 && isUnreg(S.subjects[r.i], r.t) ? ' <span class="unregTag">등록 외</span>' : ''}</td>
    <td><button type="button" class="tDel" data-j="${j}" title="이 규칙 빼기">×</button> <button type="button" class="tAdd rUp" data-j="${j}" title="위로">▲</button></td></tr>`).join('')
    : '<tr><td colspan="6" class="hint">아직 규칙이 없습니다. [+ 규칙 추가]를 눌러 예) 스프레드시트기초 · 화·목 · 오전·오후 · 최혜진 처럼 정하세요.</td></tr>';
  $$('#ruleBody .rS').forEach(e => e.onchange = () => { const r = R[+e.dataset.j]; r.i = +e.value; const s = S.subjects[r.i]; if (r.i >= 0 && !allowOut() && !regT(s).includes(r.t)) r.t = regT(s)[0] || s.teacher; renderRules(); });
  $$('#ruleBody .rD').forEach(e => e.onchange = () => { const r = R[+e.dataset.j]; r.days = $$(`#ruleBody .rD[data-j="${e.dataset.j}"]`).filter(x => x.checked).map(x => +x.value); });
  $$('#ruleBody .rH').forEach(e => e.onchange = () => { R[+e.dataset.j].h = e.value; });
  $$('#ruleBody .rT').forEach(e => e.onchange = () => { R[+e.dataset.j].t = e.value; renderRules(); });
  $$('#ruleBody .tDel').forEach(e => e.onclick = () => { R.splice(+e.dataset.j, 1); renderRules(); });
  $$('#ruleBody .rUp').forEach(e => e.onclick = () => { const j = +e.dataset.j; if (j > 0) { [R[j - 1], R[j]] = [R[j], R[j - 1]]; renderRules(); } });
}
function addRule() {
  if (!S.subjects.length) return msg('운영계획서나 전자시간표를 먼저 넣어 주세요.');
  const s = S.subjects[0]; (S.trules = S.trules || []).push({ i: s.idx, days: [], h: 'all', t: regT(s)[0] || s.teacher }); renderRules();
}

/* ---------- 전체 보기: 과목별 기간 막대(주 단위) + 평가일 ---------- */
function renderOverview() {
  const el = $('#overview'); if (!el || el.hidden || !S.settings || !S.days.length) return;
  const mon = d => addD(d, -((d.getDay() + 6) % 7));
  const w0 = mon(S.days[0]), weeks = []; for (let d = w0; d <= S.endDate; d = addD(d, 7)) weeks.push(d);
  const wOf = dk => Math.floor((parseD(dk) - w0) / 6048e5);
  const cnt = {}, first = {}, lastD = {}, tot = {};
  lessonKeysSorted().forEach(k => { const i = S.lessons[k], dk = k.split('|')[0], w = wOf(dk); cnt[i + '|' + w] = (cnt[i + '|' + w] || 0) + 1; tot[i] = (tot[i] || 0) + 1; if (!first[i]) first[i] = dk; lastD[i] = dk; });
  const ev = {}; evalPlan().forEach(e => { ev[e.s.idx] = e; });
  const rows = S.subjects.filter(s => tot[s.idx]).sort((a, b) => first[a.idx] < first[b.idx] ? -1 : first[a.idx] > first[b.idx] ? 1 : a.idx - b.idx);
  const maxW = Math.max(1, ...Object.values(cnt));
  // 달 머리글
  const mh = []; weeks.forEach(d => { const m = `${d.getFullYear()}.${d.getMonth() + 1}`; if (mh.length && mh[mh.length - 1].m === m) mh[mh.length - 1].n++; else mh.push({ m, n: 1 }); });
  const mark = (s, w) => { const e = ev[s.idx]; if (!e) return ''; return [[e.main, '●', '본평가'], [e.e2, '○', '2차'], [e.e3, '○', '3차']].filter(x => wOf(x[0]) === w).map(x => `<i class="${x[1] === '●' ? 'em' : 'e2'}" title="${x[2]} ${fmtK(x[0])}">${x[1]}</i>`).join(''); };
  el.innerHTML = `<div class="ovwrap"><table class="ov"><thead><tr><th rowspan="2" class="ovn">교과목 <small>(시간)</small></th>${mh.map(x => `<th colspan="${x.n}" class="ovm">${x.m.split('.')[1]}월</th>`).join('')}</tr>
    <tr>${weeks.map((d, w) => `<th class="ovw" title="${w + 1}주차 ${d.getMonth() + 1}/${d.getDate()}~">${w + 1}</th>`).join('')}</tr></thead><tbody>
    ${rows.map(s => { const a = wOf(first[s.idx]), z = wOf(lastD[s.idx]);
      return `<tr><td class="ovn"><span class="sw" style="background:${s.color}"></span>${esc(s.name)} <small>${tot[s.idx]}시간 · ${fmtK(first[s.idx])}~${fmtK(lastD[s.idx])}</small></td>${weeks.map((d, w) => {
      const n = cnt[s.idx + '|' + w] || 0, m = mark(s, w), inBar = w >= a && w <= z;
      const cls = inBar ? `bar${w === a ? ' bs' : ''}${w === z ? ' be' : ''}${n ? '' : ' gap'}` : '';
      return `<td data-w="${w}"${cls ? ` class="${cls}"` : ''}><div class="ovc"${inBar ? ` style="background:${s.color};opacity:${n ? (0.55 + 0.45 * n / maxW).toFixed(2) : '0.3'}"` : ''} title="${esc(s.name)} · ${w + 1}주차(${d.getMonth() + 1}/${d.getDate()}~)${n ? ` · ${n}시간` : inBar ? ' · 이 주는 수업 없음' : ''}"></div>${m}</td>`; }).join('')}</tr>`; }).join('')}
    </tbody></table></div><p class="hint" style="margin:4px 0 0">막대 = 첫 수업일부터 마지막 수업일까지(칸 하나 = 한 주). 색이 진할수록 그 주에 수업이 많고, 흐린 칸은 기간 중 수업이 없는 주입니다. <b>●</b> 본평가 · <b>○</b> 2차·3차. 칸을 누르면 아래 주간 시간표의 그 주로 갑니다.</p>`;
  $$('#overview td[data-w]').forEach(td => td.onclick = () => { const t = $('#wk' + td.dataset.w); if (t) { $('#preview').scrollTop = t.offsetTop - $('#preview').offsetTop; t.classList.add('flash'); setTimeout(() => t.classList.remove('flash'), 1200); } });
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
  $('#btnOrder').onclick = () => { const el = $('#orderPanel'); el.hidden = !el.hidden; $('#btnOrder').classList.toggle('on', !el.hidden);
    if (!el.hidden && $('#mode').value !== 'split') { $('#mode').value = 'split'; toggleMode(); if (S.subjects.length) renderSetup(); msg('세부설정: 편성 방식을 \'오전·오후 나눠서\'로 바꿨습니다. 오전 목록·오후 목록 순서대로 뿌립니다.'); }
    renderOrder(); };
  $('#btnFixTpl').onclick = dlFixedTemplate;
  $('#btnFixLoad').onclick = () => $('#fFixed').click();
  $('#fFixed').addEventListener('change', e => { e.target.files[0] && onFixed(e.target.files[0]); e.target.value = ''; });
  $('#btnRef').onclick = () => $('#fRef').click();
  $('#btnRefDir').onclick = () => $('#fRefDir').click();
  $('#fRef').addEventListener('change', e => { onRefFiles([...e.target.files]); e.target.value = ''; });
  $('#fRefDir').addEventListener('change', e => { onRefFiles([...e.target.files].filter(f => /\.(json|xlsx)$/i.test(f.name) && !/^~\$/.test(f.name))); e.target.value = ''; });
  $('#btnLockMode').onclick = toggleLockMode;
  $('#btnRules').onclick = () => { const el = $('#rulePanel'); el.hidden = !el.hidden; $('#btnRules').classList.toggle('on', !el.hidden); renderRules(); };
  $('#ruleAdd').onclick = addRule;
  $('#ruleApply').onclick = reapplyRules;
  $('#allowOut').onchange = () => { renderRules(); if (S.settings) refreshAll(false); };
  $('#btnOverview').onclick = () => { const el = $('#overview'); el.hidden = !el.hidden; $('#btnOverview').classList.toggle('on', !el.hidden); renderOverview(); };
  bindDrop($('#dzElec'), $('#fElec'), f => { $('#elecName').textContent = f.name; onElecImport(f).catch(er => msg('오류: ' + er.message)); $('#fElec').value = ''; });
  $('#btnFixExport').onclick = dlFixedExport;
  $('#mode').addEventListener('change', renderOrder);
  $('#lunch').addEventListener('change', renderOrder);
  $$('.pchk').forEach(e => e.addEventListener('change', renderOrder));
}
window.addEventListener('DOMContentLoaded', bindExtra);
