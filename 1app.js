(() => {
  'use strict';

  const XLSX_URL = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
  const state = { data26: [], data25: [], all: [], filtered: [] };
  const COLORS = ['#39a7ff', '#7b61ff', '#ff5c8a', '#2ee879', '#ffd166', '#ff9f43', '#2ff3ff', '#a78bfa'];

  const COLS = {
    month: ['출발 년월', '출발년월'], pax: ['모객'], sales: ['매출(판매가+추가판매)', '매출'], profit: ['공헌이익'],
    hq: ['운영본부'], dept: ['운영부서'], region: ['상품지역'], country: ['상품국가'], channel: ['예약경로 기준 분류'], grade: ['상품등급']
  };

  const $ = id => document.getElementById(id);
  const num = v => (v === null || v === undefined || v === '') ? 0 : (Number(String(v).replace(/,/g, '')) || 0);
  const txt = v => (v === null || v === undefined || String(v).trim() === '') ? '미분류' : String(v).trim();
  const fmtNum = v => Math.round(v || 0).toLocaleString('ko-KR');
  const fmtEok = v => `${(Number(v || 0) / 100000000).toLocaleString('ko-KR', { maximumFractionDigits: 1 })}억`;
  const log = m => { const el = $('logArea'); if (el) { el.textContent += `\n${m}`; el.scrollTop = el.scrollHeight; } };
  const setStatus = m => { const el = $('statusPill'); if (el) el.textContent = m; };

  function pick(row, keys) { for (const k of keys) if (Object.prototype.hasOwnProperty.call(row, k)) return row[k]; return ''; }
  function normalizeMonth(v, isBase) {
    let s = txt(v);
    if (/^\d{6}$/.test(s)) s = `${s.slice(0, 4)}-${s.slice(4, 6)}`;
    if (/^\d{4}[./]\d{1,2}/.test(s)) s = s.replace(/[.]/g, '-').slice(0, 7);
    if (/^\d{4}-\d{1,2}/.test(s)) { const p = s.split('-'); s = `${p[0]}-${String(p[1]).padStart(2, '0')}`; }
    if (isBase && /^2025-\d{2}$/.test(s)) return `2026-${s.slice(5, 7)}`;
    return s;
  }

  function normalizeRows(rows, isBase) {
    return rows.map(r => ({
      month: normalizeMonth(pick(r, COLS.month), isBase), pax: num(pick(r, COLS.pax)), sales: num(pick(r, COLS.sales)), profit: num(pick(r, COLS.profit)),
      hq: txt(pick(r, COLS.hq)), dept: txt(pick(r, COLS.dept)), region: txt(pick(r, COLS.region)), country: txt(pick(r, COLS.country)), channel: txt(pick(r, COLS.channel)), grade: txt(pick(r, COLS.grade)),
      year: isBase ? '2025' : '2026'
    })).filter(r => r.month && (r.pax || r.sales || r.profit));
  }

  function loadXlsxLib() { return new Promise((resolve, reject) => { if (window.XLSX) return resolve(); const s = document.createElement('script'); s.src = XLSX_URL; s.onload = () => window.XLSX ? resolve() : reject(new Error('XLSX 객체 없음')); s.onerror = () => reject(new Error('XLSX 로드 실패')); document.head.appendChild(s); }); }
  function totals(rows) { return rows.reduce((a, r) => (a.pax += r.pax, a.sales += r.sales, a.profit += r.profit, a), { pax: 0, sales: 0, profit: 0 }); }
  function groupSum(rows, key) { const m = new Map(); rows.forEach(r => { if (!m.has(r[key])) m.set(r[key], { pax: 0, sales: 0, profit: 0 }); const x = m.get(r[key]); x.pax += r.pax; x.sales += r.sales; x.profit += r.profit; }); return m; }

  function fitCanvas(canvas) { const rect = canvas.getBoundingClientRect(); const dpr = window.devicePixelRatio || 1; canvas.width = Math.max(320, rect.width) * dpr; canvas.height = Math.max(220, rect.height) * dpr; const ctx = canvas.getContext('2d'); ctx.setTransform(dpr, 0, 0, dpr, 0, 0); return { ctx, w: Math.max(320, rect.width), h: Math.max(220, rect.height) }; }

  function drawBars(id, labels, valuesA, valuesB) {
    const c = $(id); if (!c) return; const { ctx, w, h } = fitCanvas(c); ctx.clearRect(0, 0, w, h);
    const pad = { l: 46, r: 16, t: 16, b: 42 }, max = Math.max(1, ...valuesA, ...valuesB), innerW = w - pad.l - pad.r, innerH = h - pad.t - pad.b;
    const gW = innerW / Math.max(labels.length, 1), bw = Math.max(4, gW * 0.32);
    ctx.strokeStyle = 'rgba(255,255,255,.12)'; for (let i = 0; i < 4; i++) { const y = pad.t + innerH * i / 3; ctx.beginPath(); ctx.moveTo(pad.l, y); ctx.lineTo(w - pad.r, y); ctx.stroke(); }
    labels.forEach((lab, i) => {
      const x = pad.l + gW * i + gW * 0.18;
      const hA = innerH * (valuesA[i] / max), hB = innerH * (valuesB[i] / max);
      ctx.fillStyle = COLORS[0]; ctx.fillRect(x, h - pad.b - hA, bw, hA);
      ctx.fillStyle = COLORS[2]; ctx.fillRect(x + bw + 3, h - pad.b - hB, bw, hB);
      ctx.fillStyle = '#cdd7ef'; ctx.font = '11px sans-serif'; ctx.textAlign = 'center'; ctx.fillText(lab.slice(5), x + bw, h - 12);
    });
  }

  function drawDonut(id, labels, values) {
    const c = $(id); if (!c) return; const { ctx, w, h } = fitCanvas(c); ctx.clearRect(0, 0, w, h);
    const cx = w / 2, cy = h / 2, r = Math.min(w, h) * 0.35, inner = r * 0.6, sum = values.reduce((a, b) => a + b, 0) || 1;
    let start = -Math.PI / 2;
    values.forEach((v, i) => { const end = start + (v / sum) * Math.PI * 2; ctx.beginPath(); ctx.moveTo(cx, cy); ctx.arc(cx, cy, r, start, end); ctx.closePath(); ctx.fillStyle = COLORS[i % COLORS.length]; ctx.fill(); start = end; });
    ctx.globalCompositeOperation = 'destination-out'; ctx.beginPath(); ctx.arc(cx, cy, inner, 0, Math.PI * 2); ctx.fill(); ctx.globalCompositeOperation = 'source-over';
  }

  function buildFilters() {
    const keys = [['filterHQ', 'hq'], ['filterDept', 'dept'], ['filterRegion', 'region'], ['filterCountry', 'country'], ['filterChannel', 'channel'], ['filterGrade', 'grade'], ['filterMonth', 'month']];
    for (const [id, key] of keys) { const vals = [...new Set(state.data26.map(x => x[key]).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'ko')); $(id).innerHTML = '<option value="ALL">전체</option>' + vals.map(v => `<option value="${v}">${v}</option>`).join(''); }
  }

  function applyFilters() { const f = { hq: $('filterHQ').value, dept: $('filterDept').value, region: $('filterRegion').value, country: $('filterCountry').value, channel: $('filterChannel').value, grade: $('filterGrade').value, month: $('filterMonth').value }; state.filtered = state.all.filter(r => Object.entries(f).every(([k, v]) => v === 'ALL' || r[k] === v)); render(); }

  function render() {
    const cur = state.filtered.filter(x => x.year === '2026'); const base = state.filtered.filter(x => x.year === '2025'); const c = totals(cur), b = totals(base);
    $('dataScope').textContent = `자동 로드 완료: 2026 ${cur.length.toLocaleString()}건 / 2025 ${base.length.toLocaleString()}건`;
    $('kpiGrid').innerHTML = `<div class="kpi-card"><div class="kpi-label">총 모객</div><div class="kpi-value">${fmtNum(c.pax)}명</div><div class="kpi-delta">기준 ${fmtNum(b.pax)}명</div></div><div class="kpi-card"><div class="kpi-label">총 매출</div><div class="kpi-value">${fmtEok(c.sales)}</div><div class="kpi-delta">기준 ${fmtEok(b.sales)}</div></div><div class="kpi-card"><div class="kpi-label">공헌이익</div><div class="kpi-value">${fmtEok(c.profit)}</div><div class="kpi-delta">기준 ${fmtEok(b.profit)}</div></div><div class="kpi-card"><div class="kpi-label">분석 건수</div><div class="kpi-value">${fmtNum(cur.length + base.length)}건</div><div class="kpi-delta">필터 반영</div></div>`;

    const m26 = groupSum(cur, 'month'); const m25 = groupSum(base, 'month'); const months = [...new Set([...m26.keys(), ...m25.keys()])].sort();
    drawBars('chartMonthlyOverview', months, months.map(m => (m26.get(m) || {}).pax || 0), months.map(m => (m25.get(m) || {}).pax || 0));
    drawBars('chartPaxMonthly', months, months.map(m => (m26.get(m) || {}).pax || 0), months.map(m => (m25.get(m) || {}).pax || 0));
    drawBars('chartRevenueMonthly', months, months.map(m => (m26.get(m) || {}).sales || 0), months.map(m => (m25.get(m) || {}).sales || 0));

    const hq = [...groupSum(cur, 'hq').entries()].map(([k, v]) => ({ k, v: v.sales })).sort((a, b2) => b2.v - a.v).slice(0, 8);
    drawDonut('chartHQDonut', hq.map(x => x.k), hq.map(x => x.v));

    const region = [...groupSum(cur, 'region').entries()].map(([k, v]) => ({ k, v: v.sales })).sort((a, b2) => b2.v - a.v).slice(0, 10);
    drawBars('chartRegionBar', region.map(x => x.k), region.map(x => x.v), region.map(() => 0));

    const country = [...groupSum(cur, 'country').entries()].map(([k, v]) => ({ k, v: v.pax })).sort((a, b2) => b2.v - a.v).slice(0, 12);
    drawBars('chartCountryTop', country.map(x => x.k), country.map(x => x.v), country.map(() => 0));
  }

  async function loadWorkbookFromArrayBuffer(ab, label) {
    await loadXlsxLib();
    const wb = XLSX.read(ab, { type: 'array' });
    const ws26 = wb.Sheets['RAW_기준일'] || wb.Sheets['RAW_기준년'];
    const ws25 = wb.Sheets['RAW_YOY'];
    if (!ws26 || !ws25) throw new Error('필수 시트 누락: RAW_기준일(또는 RAW_기준년), RAW_YOY');
    state.data26 = normalizeRows(XLSX.utils.sheet_to_json(ws26, { defval: '' }), false);
    state.data25 = normalizeRows(XLSX.utils.sheet_to_json(ws25, { defval: '' }), true);
    state.all = state.data26.concat(state.data25);
    buildFilters(); state.filtered = state.all; render(); setStatus('분석 완료'); log(`${label} 로드 완료`);
  }

  async function autoLoad() {
    try { setStatus('기본 엑셀 자동 로드 중...'); const r = await fetch('./data.xlsx', { cache: 'no-store' }); if (!r.ok) throw new Error(`HTTP ${r.status}`); await loadWorkbookFromArrayBuffer(await r.arrayBuffer(), 'data.xlsx'); setStatus('자동 로드 성공'); }
    catch (e) { setStatus('자동 로드 실패: 파일 첨부 사용'); log(`자동 로드 실패: ${e.message}`); }
  }

  function boot() {
    $('logArea').textContent = 'app.js 로드 완료'; setStatus('대기 중');
    $('analyzeBtn').addEventListener('click', async () => { const f = $('fileInput').files && $('fileInput').files[0]; if (!f) return alert('먼저 파일을 선택하세요.'); await loadWorkbookFromArrayBuffer(await f.arrayBuffer(), f.name); });
    $('sampleBtn').addEventListener('click', autoLoad);
    ['filterHQ', 'filterDept', 'filterRegion', 'filterCountry', 'filterChannel', 'filterGrade', 'filterMonth'].forEach(id => $(id).addEventListener('change', applyFilters));
    $('resetFilters').addEventListener('click', () => { ['filterHQ', 'filterDept', 'filterRegion', 'filterCountry', 'filterChannel', 'filterGrade', 'filterMonth'].forEach(id => $(id).value = 'ALL'); applyFilters(); });
    window.addEventListener('resize', () => state.all.length && render());
    autoLoad();
  }

  document.addEventListener('DOMContentLoaded', boot);
})();
