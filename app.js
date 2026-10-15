import { Wllama } from './wllama.js';

// 믿:음 2.0 Mini 2.3B, 4비트 GGUF 약 1.4GB. 시험판은 허깅페이스 공개본에서 받는다.
// 정식판은 gs://axbot 버킷의 서명 URL로 바꾼다.
const MODEL = { repo: 'DevQuasar/K-intelligence.Midm-2.0-Mini-Instruct-GGUF',
                file: 'K-intelligence.Midm-2.0-Mini-Instruct.Q4_K_M.gguf' };
const N_CTX = 6144;          // 한 번에 다루는 글자 그릇
const CTX_CHARS = 3600;      // 모델에 넘기는 자료 글자 수 상한
const CHUNK = 600;           // 자료를 나누는 조각 크기

const $ = (id) => document.getElementById(id);
const show = (id) => { document.querySelectorAll('.screen').forEach(s => s.classList.remove('on')); $(id).classList.add('on'); };

let picked = [];   // {path, file}
let docs = [];     // {path, kind, text, facts}
let chunks = [];   // {path, text}
let overviewText = '';
let history = [];  // {q, a}
let wllama = null, modelReady = false, modelErr = null;

/* ---------- 1. 작업에 사용할 폴더 선택 ---------- */
window.addEventListener('dragover', e => e.preventDefault());
window.addEventListener('drop', e => e.preventDefault());
$('pickBtn').addEventListener('click', () => $('pickInput').click());
$('pickInput').addEventListener('change', e => {
  for (const f of e.target.files) picked.push({ path: f.webkitRelativePath || f.name, file: f });
  e.target.value = '';
  $('picked').textContent = `자료 ${picked.length}개를 받았습니다.`;
  $('go').disabled = picked.length === 0;
});

/* ---------- 2. 진도: 여섯 조각 ---------- */
const STEPS = ['자료 읽기', '빈 자료·중복 검사', '내용 나누기', '한눈에 정리', '내장된 AI 불러오는 중입니다', '준비 완료'];
const COLORS = ['#1A237E', '#7B1F2A', '#00695C', '#E65100', '#4A148C', '#1B5E20'];
function pie(done) {
  const svg = $('pie'); let h = '';
  for (let i = 0; i < 6; i++) {
    const a0 = (i / 6) * 2 * Math.PI - Math.PI / 2, a1 = ((i + 1) / 6) * 2 * Math.PI - Math.PI / 2;
    const p = (a) => `${(90 * Math.cos(a)).toFixed(2)} ${(90 * Math.sin(a)).toFixed(2)}`;
    const fill = i < done ? COLORS[i] : '#fff';
    h += `<path d="M0 0 L${p(a0)} A90 90 0 0 1 ${p(a1)} Z" fill="${fill}" stroke="${COLORS[i]}" stroke-width="2"/>`;
  }
  svg.innerHTML = h;
}
let curStep = 0;
function step(i, note = '') { curStep = i; pie(i); $('stepName').textContent = STEPS[Math.min(i, 5)]; $('stepNote').textContent = note; }

/* ---------- 자료 읽기 ---------- */
const ext = (p) => (p.split('.').pop() || '').toLowerCase();
async function readOne({ path, file }) {
  const e = ext(path);
  try {
    if (['txt', 'md', 'log', 'json', 'xml', 'html', 'htm'].includes(e)) {
      let t = await file.text();
      if (e === 'html' || e === 'htm') t = new DOMParser().parseFromString(t, 'text/html').body.innerText;
      return { path, kind: '글', text: t };
    }
    if (['csv', 'tsv', 'xlsx', 'xls'].includes(e)) {
      const wb = e === 'csv' || e === 'tsv'
        ? XLSX.read(await file.text(), { type: 'string', FS: e === 'tsv' ? '\t' : ',', sheetRows: 3000 })
        : XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true, sheetRows: 3000 });
      let text = '', facts = [];
      for (const name of wb.SheetNames) {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: '' });
        if (!rows.length) continue;
        facts.push(tableFacts(name, rows));
        text += `[${name}]\n` + rows.slice(0, 400).map(r => r.map(cell).join(' | ')).join('\n') + '\n';
      }
      return { path, kind: '표', text, facts };
    }
    if (e === 'docx') {
      const r = await mammoth.extractRawText({ arrayBuffer: await file.arrayBuffer() });
      return { path, kind: '문서', text: r.value };
    }
    if (e === 'pdf') {
      pdfjsLib.GlobalWorkerOptions.workerSrc = 'pdf.worker.min.js';
      const pdf = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
      let t = '';
      for (let i = 1; i <= Math.min(pdf.numPages, 30); i++) {
        const c = await (await pdf.getPage(i)).getTextContent();
        t += c.items.map(x => x.str).join(' ') + '\n';
      }
      return { path, kind: 'PDF', text: t };
    }
  } catch (err) { return { path, kind: '읽지 못함', text: '', err: String(err) }; }
  return { path, kind: '읽지 못함', text: '' };
}
function cell(v) {
  if (!(v instanceof Date)) return String(v).trim();
  const d = new Date(v.getTime() + 12 * 3600e3 - v.getTimezoneOffset() * 60e3); // 엑셀 날짜의 시간대 밀림 보정
  return d.toISOString().slice(0, 10);
}

// 숫자는 규칙으로 계산한다. 모델은 이 값을 옮겨 쓰기만 한다.
function tableFacts(sheet, rows) {
  const head = rows[0].map(cell), body = rows.slice(1).filter(r => r.some(v => String(v).trim() !== ''));
  const cols = [];
  head.forEach((h, j) => {
    const vals = body.map(r => r[j]).filter(v => v !== '' && v != null);
    const nums = vals.map(v => typeof v === 'number' ? v : Number(String(v).replace(/,/g, ''))).filter(n => Number.isFinite(n));
    const dates = vals.filter(v => v instanceof Date || /^\d{4}[-./]\d{1,2}[-./]\d{1,2}/.test(String(v))).map(cell).sort();
    if (dates.length >= vals.length * 0.8 && dates.length) cols.push(`${h || j + 1}열: 날짜 ${dates[0]} ~ ${dates[dates.length - 1]}`);
    else if (nums.length >= vals.length * 0.8 && nums.length) {
      const sum = nums.reduce((a, b) => a + b, 0);
      cols.push(`${h || j + 1}열: 최소 ${fmt(Math.min(...nums))}, 최대 ${fmt(Math.max(...nums))}, 평균 ${fmt(sum / nums.length)}, 합계 ${fmt(sum)}`);
    } else if (vals.length) {
      const cnt = {}; vals.forEach(v => { const k = cell(v); cnt[k] = (cnt[k] || 0) + 1; });
      const top = Object.entries(cnt).sort((a, b) => b[1] - a[1]).slice(0, 3).map(([k, n]) => `${k} ${n}건`).join(', ');
      cols.push(`${h || j + 1}열: 많은 값 ${top}`);
    }
  });
  return `[${sheet}] ${body.length}행, 항목 ${head.length}개 (${head.filter(Boolean).join(', ')})\n  ` + cols.join('\n  ');
}
function fmt(n) { return Number.isInteger(n) ? n.toLocaleString('ko-KR') : n.toLocaleString('ko-KR', { maximumFractionDigits: 2 }); }

async function hash(file) {
  const buf = await crypto.subtle.digest('SHA-256', await file.arrayBuffer());
  return [...new Uint8Array(buf)].map(b => b.toString(16).padStart(2, '0')).join('');
}

/* ---------- 작업 시키기 ---------- */
$('go').addEventListener('click', async () => {
  show('s2');
  loadModel();                         // 모델은 자료 정리와 함께 받는다
  step(0, `자료 ${picked.length}개`);
  docs = [];
  for (let i = 0; i < picked.length; i++) {
    const p = picked[i];
    $('stepNote').textContent = `${i + 1} / ${picked.length} · ${p.path.split('/').pop()}`;
    if (p.file.size > MAX_BYTES) { docs.push({ path: p.path, kind: '읽지 못함', text: '', big: true }); continue; }
    docs.push(await withTimeout(readOne(p), 20000, { path: p.path, kind: '읽지 못함', text: '' }));
  }
  step(1);

  const seen = new Map(); const dup = [], empty = [], fail = [];
  for (let i = 0; i < picked.length; i++) {
    const d = docs[i];
    if (d.kind === '읽지 못함') { fail.push(d.path); continue; }
    if (!d.text.trim()) { empty.push(d.path); d.skip = true; continue; }
    const h = picked[i].file.size > MAX_BYTES ? picked[i].path : await hash(picked[i].file);
    if (seen.has(h)) { dup.push(`${d.path} = ${seen.get(h)}`); d.skip = true; } else seen.set(h, d.path);
  }
  await pause(); step(2);

  chunks = [];
  for (const d of docs.filter(d => !d.skip && d.text)) {
    const t = d.text.replace(/[ \t]+/g, ' ').replace(/\n{3,}/g, '\n\n');
    for (let i = 0; i < t.length; i += CHUNK) chunks.push({ path: d.path, text: t.slice(i, i + CHUNK) });
  }
  await pause(); step(3);

  buildOverview({ dup, empty, fail });
  await pause(); step(4, 'AX봇이 작업할 폴더를 준비 중입니다');
  await modelPromise;
  if (modelErr) {
    $('stepName').textContent = '내장된 AI를 불러오지 못했습니다';
    $('stepNote').textContent = '인터넷 연결을 확인하고 새로 고침해 주세요';
    return;
  }
  step(6);
  await pause(600);
  show('s3');
  $('ask').focus();
});
const MAX_BYTES = 30 * 1048576; // 30MB 넘는 파일은 건너뜀
const withTimeout = (pr, ms, fb) => Promise.race([pr, new Promise(r => setTimeout(() => r(fb), ms))]);
const pause = (ms = 350) => new Promise(r => setTimeout(r, ms));

function buildOverview({ dup, empty, fail }) {
  const ok = docs.filter(d => !d.skip && d.kind !== '읽지 못함');
  const kinds = {}; ok.forEach(d => kinds[d.kind] = (kinds[d.kind] || 0) + 1);
  const lines = ok.map(d => {
    const first = d.facts ? d.facts.map(f => f.split('\n')[0]).join(' / ') : firstLine(d.text);
    return `${d.path} (${d.kind}) — ${first}`;
  });
  overviewText = `불러온 자료 ${picked.length}개 중 읽은 자료 ${ok.length}개\n` +
    lines.join('\n') + '\n' +
    ok.filter(d => d.facts).map(d => `${d.path} 계산값\n${d.facts.join('\n')}`).join('\n');

  const esc = (s) => s.replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
  $('overview').innerHTML =
    `<h2>폴더 안 내용 한눈에</h2>` +
    `<div>자료 ${picked.length}개 · 읽은 자료 ${ok.length}개 · ` +
    Object.entries(kinds).map(([k, n]) => `${k} ${n}`).join(' · ') + `</div>` +
    `<ul>${lines.slice(0, 30).map(l => `<li>${esc(l)}</li>`).join('')}${lines.length > 30 ? `<li>외 ${lines.length - 30}개</li>` : ''}</ul>` +
    (dup.length ? `<div style="margin-top:8px">같은 자료 ${dup.length}개는 한 번만 썼습니다.</div>` : '') +
    (empty.length + fail.length ? `<div style="margin-top:4px;color:#7B1F2A">읽지 못한 자료 ${empty.length + fail.length}개: ${esc([...fail, ...empty].slice(0, 5).join(', '))}${empty.length + fail.length > 5 ? ' 등' : ''}</div>` : '');
}
function firstLine(t) { const s = t.trim().split('\n').find(x => x.trim()) || ''; return s.trim().slice(0, 60); }

/* ---------- 믿:음 불러오기 ---------- */
let modelPromise = null, onCpu = false;
const progressCallback = ({ loaded, total }) => {
  if ($('s2').classList.contains('on') && total && curStep === 4)
    $('stepNote').textContent = `AX봇이 작업할 폴더를 준비 중입니다 · ${Math.round(loaded / 1048576).toLocaleString()}MB / ${Math.round(total / 1048576).toLocaleString()}MB`;
};
async function loadCpu() {
  try { await wllama.exit(); } catch {}
  wllama = new Wllama({ default: './wllama.wasm' });
  await wllama.loadModelFromHF(MODEL, { n_ctx: N_CTX, n_gpu_layers: 0, progressCallback });
  onCpu = true;
}
function loadModel() {
  if (modelPromise) return modelPromise;
  wllama = new Wllama({ default: './wllama.wasm' }, { parallelDownloads: 3 });
  modelPromise = (async () => {
    try {
      await wllama.loadModelFromHF(MODEL, { n_ctx: N_CTX, progressCallback });
    } catch (e1) {
      console.warn('그래픽 장치로 불러오기 실패, CPU로 다시 불러온다', e1);
      await loadCpu();
    }
    modelReady = true;
  })().catch(e => { modelErr = e; console.error(e); });
  return modelPromise;
}

/* ---------- 3. 프롬프트 → 결과 ---------- */
document.querySelectorAll('.chip').forEach(c => c.addEventListener('click', () => { $('ask').value = c.textContent; $('ask').focus(); }));
$('send').addEventListener('click', ask);
$('ask').addEventListener('keydown', e => { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) ask(); });

const SYSTEM = `너는 AX봇이다. 아래 [자료]에 적힌 내용만으로 답한다.
- 자료에 없는 사실, 숫자, 이름, 날짜를 지어내지 않는다.
- 숫자는 [자료]와 [계산값]에 적힌 그대로 옮긴다. 새로 계산하지 않는다.
- 자료로 답할 수 없으면 "불러온 자료에서는 찾을 수 없습니다"라고만 쓴다.
- 보고서를 요청받으면 제목, 요약, 주요 내용, 맺음 순서로 쓴다.
- 문장은 "~한다", "~이다"로 끝낸다.
- 내용마다 나온 자료 이름을 [자료 이름]으로 밝힌다.`;

function bigrams(s) { const t = s.replace(/\s+/g, ''); const g = new Set(); for (let i = 0; i < t.length - 1; i++) g.add(t.slice(i, i + 2)); return g; }
const WIDE = /정리|보고서|전체|요약|모두|전부|한눈|무엇|뭐가|중요/;

function pickContext(q, limit = CTX_CHARS) {
  const qg = bigrams(q);
  const scored = chunks.map((c, i) => {
    const cg = bigrams(c.text); let s = 0; qg.forEach(g => { if (cg.has(g)) s++; });
    return { i, s };
  }).sort((a, b) => b.s - a.s);
  const out = [], used = new Set(); let len = 0;
  const add = (i) => { if (used.has(i)) return; const t = `[${chunks[i].path}]\n${chunks[i].text}`; if (len + t.length > limit) return; used.add(i); out.push(t); len += t.length; };
  // 넓은 요청이면 자료마다 앞 조각을 먼저 고르게 넣는다
  if (WIDE.test(q) || scored[0]?.s === 0) {
    const firsts = new Map(); chunks.forEach((c, i) => { if (!firsts.has(c.path)) firsts.set(c.path, i); });
    firsts.forEach(i => add(i));
  }
  scored.forEach(({ i, s }) => { if (s > 0) add(i); });
  return out.join('\n\n');
}

async function ask() {
  const q = $('ask').value.trim();
  if (!q) return;
  if (!modelReady) { $('modelState').textContent = modelErr ? '내장된 AI를 불러오지 못했습니다' : '내장된 AI를 불러오는 중입니다'; return; }
  $('send').disabled = true; $('modelState').textContent = '작성 중';
  const turn = document.createElement('div'); turn.className = 'turn';
  turn.innerHTML = `<div class="q"></div><div class="a"></div>`;
  turn.querySelector('.q').textContent = q;
  $('talk').prepend(turn);
  const aEl = turn.querySelector('.a');

  const build = (k) => [
    { role: 'system', content: SYSTEM },
    { role: 'user', content: `[계산값]\n${overviewText.slice(0, Math.round(1800 * k))}\n\n[자료]\n${pickContext(q, CTX_CHARS * k)}` },
    { role: 'assistant', content: '자료를 읽었다. 요청을 말해 주면 자료 안에서만 답한다.' },
  ];
  const withTurn = (k) => {
    const m = build(k);
    history.slice(-2).forEach(h => m.push({ role: 'user', content: h.q }, { role: 'assistant', content: h.a.slice(0, 600) }));
    m.push({ role: 'user', content: q });
    return m;
  };

  let a = '', lastErr = null;
  const run = async (k) => {
    a = '';
    const stream = await wllama.createChatCompletion({ messages: withTurn(k), max_tokens: 1024, temperature: 0.2, top_p: 0.9, stream: true });
    for await (const ch of stream) { a += ch.choices?.[0]?.delta?.content ?? ''; aEl.textContent = a; }
    if (!a.trim()) throw new Error('빈 답변');
  };
  // 1차: 그대로 → 2차: CPU로 바꿔서 → 3차: 자료를 절반으로 줄여서
  const tries = [async () => run(1), async () => { if (!onCpu) { $('modelState').textContent = '다시 작성 중'; await loadCpu(); } return run(1); }, async () => run(0.5)];
  let ok = false;
  for (const t of tries) {
    try { await t(); ok = true; break; } catch (e) { lastErr = e; console.error(e); }
  }
  if (!ok) {
    a = '결과를 만들지 못했습니다.';
    aEl.textContent = a + '\n원인: ' + (lastErr?.message || String(lastErr));
  }
  history.push({ q, a });
  $('ask').value = ''; $('send').disabled = false; $('modelState').textContent = '';
  $('send').textContent = '이어서 묻기';
}

/* ---------- 저장 · 처음으로 ---------- */
$('save').addEventListener('click', () => {
  if (!history.length) return;
  const esc = (s) => s.replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
  const body = history.map(h => `<h2 style="color:#1A237E;font-size:16pt">${esc(h.q)}</h2><p style="white-space:pre-wrap;font-size:12pt">${esc(h.a)}</p>`).join('');
  const html = `<html><head><meta charset="utf-8"></head><body style="font-family:'맑은 고딕';color:#000">${body}</body></html>`;
  const d = new Date(), ymd = `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob(['﻿' + html], { type: 'application/msword' }));
  a.download = `AXbot_result_${ymd}.doc`; a.click();
});
$('again').addEventListener('click', () => {
  $('ask').value = ''; $('ask').focus();
  $('ask').scrollIntoView({ behavior: 'smooth', block: 'center' });
});
$('home').addEventListener('click', () => {
  picked = []; docs = []; chunks = []; history = [];
  $('picked').textContent = ''; $('talk').innerHTML = ''; $('go').disabled = true; $('send').textContent = '결과 만들기';
  show('s1');
});

