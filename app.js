const $ = id => document.getElementById(id);
const fmt = x => Number(x).toLocaleString('ko-KR', { maximumFractionDigits: 3 });
const fmt0 = x => Number(x).toLocaleString('ko-KR', { maximumFractionDigits: 0 });
const fx = (x, d = 4) => Number(x).toFixed(d);
const LEVEL_KR = { basic: '기본부문', small: '소분류', mid: '중분류', large: '대분류' };
const LEVEL_FILE = { basic: 'io-basic.xlsx', small: 'io-small.xlsx', mid: 'io-mid.xlsx', large: 'io-large.xlsx' };

let worker, nextId = 0, requests = new Map();
let basicSectors = [], mappingReady = false;
let levelsLoaded = {}, validationState = {};
let selBasic = null;

function text(tag, value) { const e = document.createElement(tag); e.textContent = value; return e; }
function call(type, payload) {
  return new Promise((resolve, reject) => {
    const id = ++nextId; requests.set(id, { resolve, reject });
    worker.postMessage({ id, type, payload });
  });
}
function startWorker() {
  worker = new Worker('./worker.js');
  worker.onmessage = ({ data: d }) => {
    if (d.type === 'status') { $('status').textContent = d.text; return; }
    const p = requests.get(d.id); if (!p) return;
    requests.delete(d.id); d.ok ? p.resolve(d.result) : p.reject(Error(d.error));
  };
  worker.onerror = e => { for (const p of requests.values()) p.reject(Error(e.message || 'Worker 실행 오류')); requests.clear(); };
}
function options(el, list) {
  el.replaceChildren(...list.map(x => { const o = document.createElement('option'); o.value = x.code; o.textContent = x.code + ' ' + x.name; return o; }));
}

// ---- 데이터 로딩 ----
async function ensureLevel(level) {
  if (levelsLoaded[level]) return levelsLoaded[level];
  $('status').textContent = LEVEL_KR[level] + ' 자료 내려받는 중… (' + LEVEL_FILE[level] + ')';
  const buffer = await (await fetch('./data/' + LEVEL_FILE[level])).arrayBuffer();
  const info = await call('init', { level, levelKr: LEVEL_KR[level], bytes: buffer });
  levelsLoaded[level] = info;
  validationState[level] = info.residuals;
  renderValidation();
  return info;
}
function renderValidation() {
  const lines = Object.keys(validationState).map(lv =>
    LEVEL_KR[lv] + ': ' + JSON.stringify(validationState[lv]));
  $('validation').textContent = lines.join('\n') + '\n(최대 절대 잔차 허용값 1e-6 미만이면 통과)';
}

async function loadBasic() {
  for (const id of ['load-basic']) $(id).disabled = true;
  $('status').textContent = '분류 연결표 내려받는 중…';
  if (worker) worker.terminate();
  for (const p of requests.values()) p.reject(Error('다시 로드')); requests.clear();
  startWorker(); levelsLoaded = {}; validationState = {};
  try {
    const mapBuf = await (await fetch('./data/classification.xlsx')).arrayBuffer();
    await call('mapping', mapBuf); mappingReady = true;
    const info = await ensureLevel('basic');
    basicSectors = info.sectors;
    // 공유 대상 선택
    $('target-panel').hidden = false;
    options($('target'), basicSectors);
    $('target-search').value = '';
    selBasic = basicSectors[0].code;
    $('target').value = selBasic;
    updateChain();
    // 탐색/시뮬레이터 활성화
    $('sheet').replaceChildren(...info.sheets.map(s => { const o = text('option', s); o.value = s; return o; }));
    options($('sector'), basicSectors);
    for (const id of ['sheet', 'row', 'row-search', 'col-search', 'sector', 'sector-search', 'amount', 'add', 'run', 'clear', 'run-a', 'run-b', 'run-c', 'run-d']) $(id).disabled = false;
    await explore();
    $('status').textContent = '준비 완료 · 기본부문 ' + basicSectors.length + '개 · 계수 검증 통과. 분석 대상 산업을 고르세요.';
  } catch (e) {
    $('status').textContent = '준비 실패: ' + e.message + ' — 새로고침 후 재시도하세요.';
    $('load-basic').disabled = false;
  }
}

// ---- 공유 대상 ----
function filterTargets() {
  const q = $('target-search').value.trim();
  const list = basicSectors.filter(s => (s.code + ' ' + s.name).includes(q));
  options($('target'), list);
  // 걸러진 목록에 기존 선택이 남아 있으면 유지, 없으면 맨 위 항목으로 선택을 갱신한다.
  if (list.some(s => s.code === selBasic)) {
    $('target').value = selBasic;
  } else if (list.length) {
    selBasic = list[0].code;
    $('target').value = selBasic;
  }
  updateChain();
}
function updateChain() {
  const s = basicSectors.find(x => x.code === selBasic);
  $('target-chain').textContent = s ? ('선택: ' + s.code + ' ' + s.name + ' — 분류 연결은 분석 실행 시 표시됩니다.') : '';
}

// ---- 원자료 탐색 (기본부문) ----
let exploreSeq = 0;
async function explore() {
  const seq = ++exploreSeq;
  try {
    const d = await call('explore', { level: 'basic', sheet: $('sheet').value, row: $('row').value });
    if (seq !== exploreSeq) return;
    rowChoices = d.rows; filterRows(d.selected); cells = d.cells;
    $('unit').textContent = d.unit;
    $('explore-caption').textContent = $('sheet').value + ' / 행: ' + d.selected + ' ' + d.name;
    renderCells();
  } catch (e) { $('status').textContent = e.message; }
}
let rowChoices = [], cells = [];
function filterRows(selected) {
  const q = $('row-search').value.trim();
  options($('row'), rowChoices.filter(r => (r.code + ' ' + r.name).includes(q)));
  if (selected && [...$('row').options].some(o => o.value === selected)) $('row').value = selected;
}
function renderCells() {
  const q = $('col-search').value.trim();
  $('cells').replaceChildren(...cells.filter(c => (c.code + ' ' + c.name).includes(q)).map(c => {
    const tr = document.createElement('tr');
    [c.code, c.name, typeof c.value === 'number' ? fmt(c.value) : (c.value ?? '')].forEach(v => tr.append(text('td', v)));
    return tr;
  }));
}

// ---- 시뮬레이터 (기본부문) ----
let basket = [], simResult = null;
function renderBasket() {
  $('basket').replaceChildren(...basket.map((b, idx) => {
    const li = document.createElement('li');
    li.append(b.code + ' ' + b.name + ' / ' + fmt(b.amount) + '백만원 ');
    const del = text('button', '삭제'); del.onclick = () => { basket.splice(idx, 1); renderBasket(); };
    li.append(del); return li;
  }));
}
function sectorName(code) { const s = basicSectors.find(x => x.code === code); return s ? s.name : code; }

// ---- 분석 A ----
async function runA() {
  $('run-a').disabled = true; $('status').textContent = '① 수준별 비교 준비 중…';
  try {
    for (const lv of ['basic', 'small', 'mid', 'large']) await ensureLevel(lv);
    const d = await call('analysis', { which: 'A', args: { basic: selBasic } });
    $('a-out').hidden = false;
    $('a-caption').textContent = '대상: 기본부문 ' + selBasic + ' ' + sectorName(selBasic) + ' — 각 분류에서 이 산업이 속한 부문으로 계산';
    $('a-rows').replaceChildren(...d.rows.map(r => {
      const tr = document.createElement('tr');
      [r.levelKr, fmt0(r.sectors), r.code + ' ' + r.name, fx(r.prod, 4), fx(r.imp, 4), fx(r.va, 4)].forEach(v => tr.append(text('td', v)));
      return tr;
    }));
    const max = Math.max(...d.rows.map(r => r.prod), 1);
    $('a-bars').replaceChildren(...d.rows.map(r => bar(r.levelKr + ' · ' + fx(r.prod, 4), r.prod / max)));
    const b = d.rows[0], l = d.rows[3];
    $('a-read').textContent = '같은 산업이라도 분류 수준에 따라 승수가 ' + fx(Math.min(...d.rows.map(r => r.prod)), 3) + ' ~ ' + fx(Math.max(...d.rows.map(r => r.prod)), 3) + ' 로 달라집니다. 통합 분류는 그 부문 자체가 아니라 같은 묶음 산업들의 평균 구조를 반영하기 때문입니다. 분석 목적에 맞는 분류 수준을 고르는 것이 중요합니다.';
    $('status').textContent = '① 완료.';
  } catch (e) { $('status').textContent = '① 실패: ' + e.message; }
  $('run-a').disabled = false;
}
function bar(label, frac) {
  const d = text('div', label); d.className = 'bar';
  const t = document.createElement('div'); t.className = 'track';
  const f = document.createElement('div'); f.className = 'fill'; f.style.width = Math.max(1, frac * 100) + '%';
  t.append(f); d.append(t); return d;
}

// ---- 분석 B : 영향력 × 감응도 산점도 ----
async function runB() {
  $('run-b').disabled = true; const level = $('b-level').value;
  $('status').textContent = '② 핵심산업 지도 준비 중…';
  try {
    await ensureLevel(level);
    const d = await call('analysis', { which: 'B', args: { level, basic: selBasic } });
    $('b-out').hidden = false;
    $('b-caption').textContent = d.levelKr + ' ' + d.points.length + '개 부문 · 점선 교차점이 전산업 평균(1.0)';
    $('b-chart').innerHTML = buildScatter(d.points, d.selected);
    renderQuadrants(d.points, d.selected);
    $('status').textContent = '② 완료.';
  } catch (e) { $('status').textContent = '② 실패: ' + e.message; }
  $('run-b').disabled = false;
}
function buildScatter(pts, sel) {
  const W = 640, H = 460, m = 56;
  const xs = pts.map(p => p.infl), ys = pts.map(p => p.sens);
  const xmin = Math.min(...xs, 0.6), xmax = Math.max(...xs, 1.4);
  const ymin = Math.min(...ys, 0.4), ymax = Math.max(...ys, 1.6);
  const sx = v => m + (v - xmin) / (xmax - xmin) * (W - 2 * m);
  const sy = v => H - m - (v - ymin) / (ymax - ymin) * (H - 2 * m);
  let s = `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="영향력계수 대 감응도계수 산점도" style="width:100%;height:auto;background:#fbfdfe;border:1px solid #dce5eb;border-radius:8px">`;
  s += `<line x1="${m}" y1="${H - m}" x2="${W - m}" y2="${H - m}" stroke="#94a3b8"/>`;
  s += `<line x1="${m}" y1="${m}" x2="${m}" y2="${H - m}" stroke="#94a3b8"/>`;
  s += `<line x1="${sx(1)}" y1="${m}" x2="${sx(1)}" y2="${H - m}" stroke="#b65c00" stroke-dasharray="5 4"/>`;
  s += `<line x1="${m}" y1="${sy(1)}" x2="${W - m}" y2="${sy(1)}" stroke="#b65c00" stroke-dasharray="5 4"/>`;
  s += `<text x="${W - m}" y="${H - m + 34}" text-anchor="end" font-size="13" fill="#475569">영향력계수(후방연쇄) →</text>`;
  s += `<text x="${m - 40}" y="${m - 16}" font-size="13" fill="#475569">↑ 감응도계수(전방연쇄)</text>`;
  for (const p of pts) {
    const on = p.code === sel;
    s += `<circle cx="${fx(sx(p.infl), 1)}" cy="${fx(sy(p.sens), 1)}" r="${on ? 7 : 3}" fill="${on ? '#b91c1c' : '#168179'}" fill-opacity="${on ? 1 : 0.45}"><title>${p.code} ${p.name} (영향력 ${fx(p.infl, 3)}, 감응도 ${fx(p.sens, 3)})</title></circle>`;
    if (on) s += `<text x="${fx(sx(p.infl) + 9, 1)}" y="${fx(sy(p.sens) + 4, 1)}" font-size="13" font-weight="700" fill="#b91c1c">${p.name}</text>`;
  }
  return s + '</svg>';
}
function renderQuadrants(pts, sel) {
  const q = { a: 0, b: 0, c: 0, d: 0 };
  for (const p of pts) {
    if (p.infl >= 1 && p.sens >= 1) q.a++;
    else if (p.infl >= 1 && p.sens < 1) q.b++;
    else if (p.infl < 1 && p.sens >= 1) q.c++;
    else q.d++;
  }
  const defs = [
    ['a', '전·후방 모두 강', '영향력>1, 감응도>1', q.a],
    ['b', '최종재형(후방 강)', '영향력>1, 감응도<1', q.b],
    ['c', '소재·중간재형(전방 강)', '영향력<1, 감응도>1', q.c],
    ['d', '독립형', '영향력<1, 감응도<1', q.d],
  ];
  const selP = pts.find(p => p.code === sel);
  let selQ = null;
  if (selP) selQ = selP.infl >= 1 ? (selP.sens >= 1 ? 'a' : 'b') : (selP.sens >= 1 ? 'c' : 'd');
  $('b-quad').replaceChildren(...defs.map(([k, t, cond, n]) => {
    const d = document.createElement('div'); d.className = 'card';
    if (k === selQ) d.style.outline = '2px solid #b91c1c';
    d.append(text('strong', fmt0(n) + '개'), text('span', t), text('span', cond));
    return d;
  }));
  if (selP) $('b-read').textContent = '선택 산업(' + selP.name + ')은 영향력 ' + fx(selP.infl, 3) + ', 감응도 ' + fx(selP.sens, 3) + ' — ' +
    (selQ === 'c' ? '소재·중간재형입니다. 스스로 전방으로 미치는 힘은 평균 이하지만, 다른 산업이 생산을 늘릴 때 끌려 올라가는 민감도가 높습니다.' :
     selQ === 'a' ? '전·후방 연쇄가 모두 강한 핵심 산업입니다.' :
     selQ === 'b' ? '최종재형으로, 다른 산업을 끌어올리는 힘은 크지만 전방 수요 민감도는 낮습니다.' :
     '전·후방 연쇄가 모두 약한 독립형입니다.');
  else $('b-read').textContent = '';
}

// ---- 분석 C : 수요 원천 ----
async function runC() {
  $('run-c').disabled = true; const level = $('c-level').value;
  $('status').textContent = '③ 수요 원천 준비 중…';
  try {
    await ensureLevel(level);
    const d = await call('analysis', { which: 'C', args: { level, basic: selBasic } });
    $('c-out').hidden = false;
    $('c-caption').textContent = d.levelKr + ' · ' + d.code + ' ' + d.name + ' — 총 생산유발액 ' + fmt0(d.total) + '백만원';
    const max = Math.max(...d.groups.map(g => g.value), 1);
    $('c-groups').replaceChildren(...d.groups.map(g => bar(g.label + ' · ' + fmt0(g.value) + '백만원 (' + fx(g.value / d.total * 100, 1) + '%)', g.value / max)));
    $('c-detail').replaceChildren(...d.items.map(it => {
      const tr = document.createElement('tr');
      [it.label, it.group, fmt0(it.value), fx(it.value / d.total * 100, 1) + '%'].forEach(v => tr.append(text('td', v)));
      return tr;
    }));
    $('status').textContent = '③ 완료.';
  } catch (e) { $('status').textContent = '③ 실패: ' + e.message; }
  $('run-c').disabled = false;
}

// ---- 분석 D : 수출의 국내기여 ----
async function runD() {
  $('run-d').disabled = true; const level = $('d-level').value;
  const amount = Math.max(0, Number($('d-amount').value) || 0);
  $('status').textContent = '④ 수출의 국내기여 준비 중…';
  try {
    await ensureLevel(level);
    const d = await call('analysis', { which: 'D', args: { level, basic: selBasic, amount } });
    $('d-out').hidden = false;
    $('d-caption').textContent = d.levelKr + ' · ' + d.code + ' ' + d.name + ' 최종수요 ' + fmt0(d.amount) + '백만원 증가 시';
    const domestic = d.valueAdded, leak = d.imports, tot = domestic + leak;
    $('d-split').replaceChildren(stacked([
      ['국내 부가가치', domestic, '#168179'],
      ['수입 누출', leak, '#b45309'],
    ], tot));
    $('d-cards').replaceChildren(
      card('생산유발액', fmt0(d.production), '백만원'),
      card('국내 부가가치', fmt0(domestic), '백만원 · ' + fx(d.vaCoef * 100, 1) + '%'),
      card('수입 누출', fmt0(leak), '백만원 · ' + fx(d.impCoef * 100, 1) + '%'),
    );
    $('d-read').textContent = '최종수요 ' + fmt0(d.amount) + '백만원 중 국내에 남는 부가가치는 ' + fmt0(domestic) + '백만원(' + fx(d.vaCoef * 100, 1) + '%), 수입으로 해외에 새는 금액은 ' + fmt0(leak) + '백만원(' + fx(d.impCoef * 100, 1) + '%)입니다. 수입유발계수 + 부가가치유발계수 = ' + fx(d.impCoef + d.vaCoef, 4) + '.';
    $('status').textContent = '④ 완료.';
  } catch (e) { $('status').textContent = '④ 실패: ' + e.message; }
  $('run-d').disabled = false;
}
function card(title, v, unit) {
  const d = document.createElement('div'); d.className = 'card';
  d.append(text('strong', v), text('span', unit || '')); const t = text('span', title); d.prepend(t); return d;
}
function stacked(parts, tot) {
  const wrap = document.createElement('div'); wrap.className = 'bar';
  const track = document.createElement('div'); track.className = 'track'; track.style.display = 'flex'; track.style.height = '26px';
  for (const [label, v, color] of parts) {
    const seg = document.createElement('div'); seg.style.width = (v / tot * 100) + '%'; seg.style.background = color;
    seg.style.color = '#fff'; seg.style.fontSize = '.8rem'; seg.style.display = 'flex'; seg.style.alignItems = 'center'; seg.style.justifyContent = 'center'; seg.style.overflow = 'hidden';
    seg.textContent = fx(v / tot * 100, 0) + '%'; seg.title = label + ' ' + fmt0(v);
    track.append(seg);
  }
  wrap.append(track);
  const legend = text('div', parts.map(p => p[0] + ' ' + fmt0(p[1]) + '백만원').join('   ·   ')); legend.className = 'note';
  wrap.append(legend); return wrap;
}

// ---- 이벤트 ----
function wire() {
  $('load-basic').onclick = loadBasic;
  $('target-search').oninput = filterTargets;
  $('target').onchange = () => { selBasic = $('target').value; updateChain(); };
  $('sheet').onchange = explore; $('row').onchange = explore;
  $('row-search').oninput = () => filterRows($('row').value);
  $('col-search').oninput = renderCells;
  $('sector-search').oninput = () => options($('sector'), basicSectors.filter(s => (s.code + ' ' + s.name).includes($('sector-search').value.trim())));
  $('add').onclick = () => {
    const code = $('sector').value; if (!code) return;
    const amount = Number($('amount').value);
    if (!(amount > 0)) { $('status').textContent = '0보다 큰 증가액을 입력하세요.'; return; }
    basket.push({ code, name: sectorName(code), amount }); renderBasket();
  };
  $('clear').onclick = () => { basket = []; renderBasket(); $('results').hidden = true; };
  $('run').onclick = runSim;
  $('download').onclick = downloadCsv;
  $('run-a').onclick = runA;
  $('run-b').onclick = runB; $('b-level').onchange = () => { if (!$('b-out').hidden) runB(); };
  $('run-c').onclick = runC; $('c-level').onchange = () => { if (!$('c-out').hidden) runC(); };
  $('run-d').onclick = runD; $('d-level').onchange = () => { if (!$('d-out').hidden) runD(); };
}
async function runSim() {
  if (!basket.length) { $('status').textContent = '시나리오에 산업을 추가하세요.'; return; }
  $('run').disabled = true;
  try {
    const d = await call('simulate', { level: 'basic', shocks: basket.map(b => ({ code: b.code, amount: b.amount })) });
    simResult = d; $('results').hidden = false;
    $('cards').replaceChildren(
      card('수요 증가액', fmt0(d.demand), '백만원'),
      card('생산유발액', fmt0(d.production), '백만원'),
      card('수입유발액', fmt0(d.imports), '백만원'),
      card('부가가치유발액', fmt0(d.valueAdded), '백만원'),
      card('생산유발 승수', fx(d.multiplier, 4), ''),
    );
    const max = Math.max(...d.rows.map(r => r.production), 1);
    $('bars').replaceChildren(...d.rows.slice(0, 15).map(r => bar(r.name + ' · ' + fmt(r.production) + '백만원', r.production / max)));
    $('result-rows').replaceChildren(...d.rows.map(r => {
      const tr = document.createElement('tr');
      [r.code, r.name, fmt(r.production), fmt(r.imports), fmt(r.valueAdded)].forEach(v => tr.append(text('td', v)));
      return tr;
    }));
    $('status').textContent = '시뮬레이션 완료.';
  } catch (e) { $('status').textContent = '시뮬레이션 실패: ' + e.message; }
  $('run').disabled = false;
}
function csvCell(v) { let s = String(v); return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; }
function downloadCsv() {
  if (!simResult) return;
  const rows = [['산업코드', '산업명', '생산유발액_백만원', '수입유발액_백만원', '부가가치유발액_백만원'],
    ...simResult.rows.map(r => [r.code, r.name, r.production, r.imports, r.valueAdded])];
  const url = URL.createObjectURL(new Blob(['﻿' + rows.map(r => r.map(csvCell).join(',')).join('\n')], { type: 'text/csv' }));
  const a = document.createElement('a'); a.href = url; a.download = 'io-simulation.csv'; a.click(); URL.revokeObjectURL(url);
}

wire();
