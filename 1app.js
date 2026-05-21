(() => {
  'use strict';

  const COLS = {
    month: '출발 년월', pax: '모객', sales: '매출(판매가+추가판매)', profit: '공헌이익',
    hq: '운영본부', dept: '운영부서', region: '상품지역', city: '도착도시', country: '상품국가',
    channel: '예약경로 기준 분류', grade: '상품등급', date: '기준일시'
  };
  const COLORS = ['#39a7ff','#7b61ff','#ff5c8a','#2ee879','#ffd166','#ff9f43','#2ff3ff','#a78bfa','#fb7185','#34d399','#f472b6','#60a5fa'];
  let state = { raw26: [], raw25: [], data: [], filtered: [], filters: {}, slide: 0, heatMetric: 'pax' };

  const $ = id => document.getElementById(id);
  const qs = sel => document.querySelector(sel);
  const qsa = sel => [...document.querySelectorAll(sel)];
  const log = msg => { const el = $('logArea'); if (!el) return; el.textContent += `\n${new Date().toLocaleTimeString()}  ${msg}`; el.scrollTop = el.scrollHeight; };
  const setStatus = (txt, type='neutral') => { const el=$('statusPill'); if(el){ el.textContent=txt; el.className=`status-pill ${type}`; } };
  const num = v => { if (v === null || v === undefined || v === '') return 0; if (typeof v === 'number') return Number.isFinite(v) ? v : 0; const n = Number(String(v).replace(/,/g,'')); return Number.isFinite(n) ? n : 0; };
  const str = v => (v === null || v === undefined || v === '') ? '미지정' : String(v).trim();
  const fmt = (n, unit='') => (Math.abs(n) >= 100000000 ? `${(n/100000000).toLocaleString('ko-KR',{maximumFractionDigits:1})}억${unit}` : Math.abs(n) >= 10000 ? `${(n/10000).toLocaleString('ko-KR',{maximumFractionDigits:1})}만${unit}` : `${Math.round(n).toLocaleString('ko-KR')}${unit}`);
  const pct = v => Number.isFinite(v) ? `${(v*100).toFixed(1)}%` : '-';
  const safeDiv = (a,b) => b ? a/b : 0;
  const yoy = (a,b) => b ? a/b : 0;
  const shiftYear = m => /^\d{4}-\d{2}$/.test(m) ? `${Number(m.slice(0,4))+1}-${m.slice(5,7)}` : m;

  function boot(){
    $('logArea').textContent = 'app.js 로드 완료';
    setStatus('app.js 로드 완료');
    $('analyzeBtn').addEventListener('click', analyzeSelectedFile);
    $('sampleBtn').addEventListener('click', loadSample);
    $('resetFilters').addEventListener('click', resetFilters);
    $('fullscreenBtn').addEventListener('click', () => { document.body.classList.toggle('fullscreen'); document.fullscreenElement ? document.exitFullscreen?.() : document.documentElement.requestFullscreen?.(); });
    $('prevSlide').addEventListener('click', () => goSlide(state.slide - 1));
    $('nextSlide').addEventListener('click', () => goSlide(state.slide + 1));
    $('toggleHeatMetric').addEventListener('click', () => { state.heatMetric = state.heatMetric === 'pax' ? 'sales' : 'pax'; $('toggleHeatMetric').textContent = state.heatMetric === 'pax' ? '매출 YoY로 전환' : '모객 YoY로 전환'; renderAll(); });
    qsa('.tab-btn').forEach(b => b.addEventListener('click', () => goSlide(Number(b.dataset.slide))));
    ['filterHQ','filterDept','filterRegion','filterCountry','filterChannel','filterGrade','filterMonth'].forEach(id => $(id).addEventListener('change', applyFiltersAndRender));
    renderEmpty();
    autoLoadBundledExcel();
  }



  async function autoLoadBundledExcel(){
    try{
      setStatus('?? ?? ?? ?? ?...');
      log('?? ?? ?? ?? ??: ./YoY_0518_??.xlsx');
      const res = await fetch('./YoY_0518_??.xlsx', { cache: 'no-store' });
      if(!res.ok) throw new Error(`HTTP ${res.status}`);
      const ab = await res.arrayBuffer();
      const wb = XLSX.read(ab, {type:'array', cellDates:false, raw:false});
      const ws26 = wb.Sheets['RAW_???'] || wb.Sheets['RAW_???'];
      const ws25 = wb.Sheets['RAW_YOY'];
      if(!ws26 || !ws25) throw new Error('?? ?? ??: RAW_???(?? RAW_???), RAW_YOY');
      const raw26 = XLSX.utils.sheet_to_json(ws26, {defval:''});
      const raw25 = XLSX.utils.sheet_to_json(ws25, {defval:''});
      buildDataset(raw26, raw25, 'YoY_0518_??.xlsx (?? ??)');
      setStatus('?? ?? ?? ?? ??', 'ok');
      log('?? ?? ??');
    }catch(err){
      console.warn(err);
      setStatus('?? ?? ??: ?? ??? ??', 'warn');
      log('?? ?? ??: ' + err.message);
    }
  }

  function goSlide(i){
    const max = qsa('.slide').length - 1; state.slide = (i + max + 1) % (max + 1);
    qsa('.slide').forEach(s => s.classList.toggle('active', Number(s.dataset.slide) === state.slide));
    qsa('.tab-btn').forEach(b => b.classList.toggle('active', Number(b.dataset.slide) === state.slide));
  }

  async function analyzeSelectedFile(){
    const file = $('fileInput').files?.[0];
    if(!file){ alert('먼저 엑셀 파일을 선택하세요.'); return; }
    $('logArea').textContent = '파일 선택 확인: ' + file.name;
    setStatus('엑셀 읽는 중');
    if(!window.XLSX){ setStatus('SheetJS 로드 실패'); log('XLSX 라이브러리가 로드되지 않았습니다. 인터넷/CDN 차단 가능성이 있습니다. 샘플 검증은 가능합니다.'); alert('SheetJS CDN이 로드되지 않았습니다. 회사망/CDN 차단이면 xlsx.full.min.js 로컬 파일이 필요합니다.'); return; }
    try{
      log('ArrayBuffer 읽기 시작');
      const buffer = await file.arrayBuffer();
      log(`ArrayBuffer 읽기 완료: ${Math.round(buffer.byteLength/1024).toLocaleString()} KB`);
      const wb = XLSX.read(buffer, {type:'array', cellDates:false, raw:false});
      log('시트 목록: ' + wb.SheetNames.join(', '));
      const ws26 = wb.Sheets['RAW_기준일'];
      const ws25 = wb.Sheets['RAW_YOY'];
      if(!ws26 || !ws25) throw new Error('RAW_기준일 또는 RAW_YOY 시트를 찾을 수 없습니다.');
      const raw26 = XLSX.utils.sheet_to_json(ws26, {defval:''});
      const raw25 = XLSX.utils.sheet_to_json(ws25, {defval:''});
      log(`RAW_기준일 행수: ${raw26.length.toLocaleString()} / RAW_YOY 행수: ${raw25.length.toLocaleString()}`);
      buildDataset(raw26, raw25, file.name);
    }catch(err){ console.error(err); setStatus('분석 실패'); log('오류: ' + err.message); alert('분석 실패: ' + err.message); }
  }

  function loadSample(){
    $('logArea').textContent = '샘플 데이터 생성 시작';
    const months26 = ['2026-05','2026-06','2026-07','2026-08','2026-09','2026-10','2026-11','2026-12','2027-01'];
    const hqs = ['상품1본부','상품2본부','상품3본부'];
    const depts = ['동남아사업1부','일본사업부','유럽사업부','미주사업부','중국사업부','남태평양사업부'];
    const regions = ['나트랑','일본','서유럽','미주','장가계','괌/사이판'];
    const countries = ['베트남','일본','프랑스','미국','중국','괌'];
    const channels = ['대리점','모두닷컴','제휴사','웹예약(영업공통)','모두웨어(전시공통)'];
    const grades = ['없음','시그니처','시그니처블랙','하이클래스'];
    const make = (yearShift=0) => {
      const rows=[];
      for(let i=0;i<months26.length;i++) for(let j=0;j<depts.length;j++) for(let k=0;k<channels.length;k++){
        const m26 = months26[i]; const yyyy = Number(m26.slice(0,4))-yearShift; const m = `${yyyy}-${m26.slice(5,7)}`;
        const base = (depts.length-j)*8 + (i<3?40:18) + k*3;
        const factor = yearShift ? (1.25 + (j%3)*.08) : (0.85 + (i%4)*.1 + (j%2)*.08);
        const pax = Math.max(1, Math.round(base*factor));
        const asp = 750000 + j*180000 + i*25000 + k*40000;
        rows.push({[COLS.date]: `${yyyy}-05-18`, [COLS.month]: m, [COLS.pax]: pax, [COLS.sales]: pax*asp, [COLS.profit]: pax*asp*(.08+(j%4)*.025), [COLS.hq]: hqs[j%hqs.length], [COLS.dept]: depts[j], [COLS.region]: regions[j], [COLS.city]: regions[j], [COLS.country]: countries[j], [COLS.channel]: channels[k], [COLS.grade]: grades[(j+k)%grades.length], '판매속성':'01.해외PKG'});
      }
      return rows;
    };
    buildDataset(make(0), make(1), '샘플 데이터');
  }

  function normalizeRows(rows, yearLabel){
    return rows.map((r, idx) => {
      const month = str(r[COLS.month]);
      return {
        _idx: idx, year: yearLabel, displayMonth: yearLabel === '2025' ? shiftYear(month) : month, month,
        pax: num(r[COLS.pax]), sales: num(r[COLS.sales]), profit: num(r[COLS.profit]),
        hq: str(r[COLS.hq]), dept: str(r[COLS.dept]), region: str(r[COLS.region]), city: str(r[COLS.city]), country: str(r[COLS.country]), channel: str(r[COLS.channel]), grade: str(r[COLS.grade])
      };
    }).filter(r => r.displayMonth !== '미지정' && (r.pax || r.sales || r.profit));
  }

  function buildDataset(raw26, raw25, filename){
    state.raw26 = normalizeRows(raw26, '2026');
    state.raw25 = normalizeRows(raw25, '2025');
    state.data = [...state.raw26, ...state.raw25];
    log(`정규화 완료: 2026 ${state.raw26.length.toLocaleString()}행 / 2025 ${state.raw25.length.toLocaleString()}행`);
    if(!state.raw26.length || !state.raw25.length) log('주의: 한쪽 연도 데이터가 비어 있습니다. YoY가 제한됩니다.');
    $('dataScope').textContent = `${filename} · RAW_기준일 ${state.raw26.length.toLocaleString()}행 / RAW_YOY ${state.raw25.length.toLocaleString()}행`;
    buildFilters();
    applyFiltersAndRender();
    setStatus('분석 완료');
    log('대시보드 렌더링 완료');
  }

  function unique(data, key){ return [...new Set(data.map(d=>d[key]).filter(Boolean))].sort((a,b)=>String(a).localeCompare(String(b),'ko')); }
  function fillSelect(id, values, selected='ALL'){
    const el=$(id); const label=el.options[0]?.textContent || '전체'; el.innerHTML = `<option value="ALL">${label}</option>` + values.map(v=>`<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join(''); el.value = values.includes(selected) ? selected : 'ALL';
  }
  function buildFilters(){
    const d = state.raw26;
    fillSelect('filterHQ', unique(d,'hq')); fillSelect('filterDept', unique(d,'dept')); fillSelect('filterRegion', unique(d,'region')); fillSelect('filterCountry', unique(d,'country'));
    fillSelect('filterChannel', unique(d,'channel')); fillSelect('filterGrade', unique(d,'grade')); fillSelect('filterMonth', unique(d,'displayMonth'));
  }
  function resetFilters(){ ['filterHQ','filterDept','filterRegion','filterCountry','filterChannel','filterGrade','filterMonth'].forEach(id=>$(id).value='ALL'); applyFiltersAndRender(); }
  function currentFilter(){ return { hq:$('filterHQ').value, dept:$('filterDept').value, region:$('filterRegion').value, country:$('filterCountry').value, channel:$('filterChannel').value, grade:$('filterGrade').value, displayMonth:$('filterMonth').value }; }
  function pass(d,f){ return Object.keys(f).every(k => f[k]==='ALL' || d[k]===f[k]); }
  function applyFiltersAndRender(){ state.filters = currentFilter(); state.filtered = state.data.filter(d => pass(d,state.filters)); renderAll(); }

  function aggregate(data, key){
    const m = new Map();
    data.forEach(d => { const k = d[key] || '미지정'; if(!m.has(k)) m.set(k,{key:k,pax26:0,pax25:0,sales26:0,sales25:0,profit26:0,profit25:0,count:0}); const o=m.get(k); if(d.year==='2026'){o.pax26+=d.pax;o.sales26+=d.sales;o.profit26+=d.profit;} else {o.pax25+=d.pax;o.sales25+=d.sales;o.profit25+=d.profit;} o.count++; });
    return [...m.values()].map(o=>({...o, paxYoY:yoy(o.pax26,o.pax25), salesYoY:yoy(o.sales26,o.sales25), asp26:safeDiv(o.sales26,o.pax26), margin26:safeDiv(o.profit26,o.sales26)}));
  }
  function total(data){ return data.reduce((a,d)=>{ if(d.year==='2026'){a.pax26+=d.pax;a.sales26+=d.sales;a.profit26+=d.profit;} else {a.pax25+=d.pax;a.sales25+=d.sales;a.profit25+=d.profit;} return a; },{pax26:0,pax25:0,sales26:0,sales25:0,profit26:0,profit25:0}); }

  function renderAll(){
    const t = total(state.filtered);
    renderKpis(t); renderMonthly(); renderDonuts(); renderHeatmap(); renderRanks(); renderRegion(); renderChannelGrade(); renderProfit(); renderQuality(); renderInsights();
  }
  function renderEmpty(){ $('kpiGrid').innerHTML = '<div class="empty-state" style="grid-column:1/-1">엑셀 파일을 선택하거나 샘플 데이터로 화면을 검증하세요.</div>'; }

  function renderKpis(t){
    const cards = [
      ['총 모객', fmt(t.pax26,'명'), `YoY ${pct(yoy(t.pax26,t.pax25))}`, yoy(t.pax26,t.pax25)-1],
      ['총 매출', fmt(t.sales26), `YoY ${pct(yoy(t.sales26,t.sales25))}`, yoy(t.sales26,t.sales25)-1],
      ['평균판매가', fmt(safeDiv(t.sales26,t.pax26)), '매출 ÷ 모객 기준', 0],
      ['공헌이익률', pct(safeDiv(t.profit26,t.sales26)), `공헌이익 ${fmt(t.profit26)}`, safeDiv(t.profit26,t.sales26)-safeDiv(t.profit25,t.sales25)],
      ['2025 모객', fmt(t.pax25,'명'), 'RAW_YOY 동월 기준', 0],
      ['모객 증감', fmt(t.pax26-t.pax25,'명'), `${t.pax26>=t.pax25?'증가':'감소'}`, t.pax26-t.pax25],
      ['매출 증감', fmt(t.sales26-t.sales25), `${t.sales26>=t.sales25?'증가':'감소'}`, t.sales26-t.sales25],
      ['데이터 행수', state.filtered.length.toLocaleString(), '필터 적용 후', 0]
    ];
    $('kpiGrid').innerHTML = cards.map((c,i)=>`<div class="kpi-card" style="animation-delay:${i*45}ms"><div class="kpi-label">${c[0]}</div><div class="kpi-value" data-target="${stripNum(c[1])}">${c[1]}</div><div class="kpi-delta ${c[3]>0?'up':c[3]<0?'down':'neutral'}">${c[2]}</div></div>`).join('');
  }

  function renderMonthly(){
    const m = aggregate(state.filtered, 'displayMonth').sort((a,b)=>a.key.localeCompare(b.key));
    drawCombo($('chartMonthlyOverview'), m.map(x=>x.key), m.map(x=>x.pax26), m.map(x=>x.pax25), {title:'모객', unit:'명'});
    drawCombo($('chartPaxMonthly'), m.map(x=>x.key), m.map(x=>x.pax26), m.map(x=>x.pax25), {title:'모객', unit:'명'});
    drawCombo($('chartRevenueMonthly'), m.map(x=>x.key), m.map(x=>x.sales26), m.map(x=>x.sales25), {title:'매출', unit:'원'});
    drawLine($('chartASP'), m.map(x=>x.key), m.map(x=>safeDiv(x.sales26,x.pax26)), {unit:'원', color:COLORS[6]});
    drawLine($('chartMargin'), m.map(x=>x.key), m.map(x=>safeDiv(x.profit26,x.sales26)*100), {unit:'%', color:COLORS[3]});
    const best = [...m].sort((a,b)=>b.paxYoY-a.paxYoY)[0], worst = [...m].filter(x=>x.pax25).sort((a,b)=>a.paxYoY-b.paxYoY)[0];
    $('monthlyInsights').innerHTML = [best&&`<div class="rank-item"><b>최고 회복월: ${best.key}</b><small>모객 YoY ${pct(best.paxYoY)}</small></div>`, worst&&`<div class="rank-item"><b>주의월: ${worst.key}</b><small>모객 YoY ${pct(worst.paxYoY)}</small></div>`, `<div class="rank-item"><b>월수</b><small>${m.length}개 출발월 비교</small></div>`].filter(Boolean).join('');
  }
  function renderDonuts(){
    const hq = aggregate(state.filtered.filter(d=>d.year==='2026'), 'hq').sort((a,b)=>b.sales26-a.sales26).slice(0,8);
    drawDonut($('chartHQDonut'), hq.map(x=>({label:x.key,value:x.sales26})), '매출'); renderLegend('hqLegend', hq.map(x=>x.key));
  }
  function renderHeatmap(){
    const months = unique(state.filtered,'displayMonth');
    const depts = aggregate(state.filtered,'dept').sort((a,b)=>b.pax26-a.pax26).slice(0,18).map(x=>x.key);
    const by = new Map();
    aggregateBy2(state.filtered,'dept','displayMonth').forEach(o=>by.set(`${o.k1}||${o.k2}`,o));
    const metric = state.heatMetric === 'pax' ? 'paxYoY' : 'salesYoY';
    let html = '<table class="heat-table"><thead><tr><th>운영부서</th>' + months.map(m=>`<th>${m}</th>`).join('') + '</tr></thead><tbody>';
    depts.forEach(d => { html += `<tr><td class="heat-row-name">${escapeHtml(d)}</td>`; months.forEach(m=>{ const o=by.get(`${d}||${m}`)||{}; const v=o[metric]||0; html += `<td title="${escapeHtml(d)} ${m} ${state.heatMetric==='pax'?'모객':'매출'} YoY ${pct(v)}" style="background:${heatColor(v)};color:${v<.75?'#fff':'#08101f'}">${o.pax25||o.sales25?pct(v):'-'}</td>`; }); html+='</tr>'; });
    $('heatmap').innerHTML = html + '</tbody></table>';
  }
  function aggregateBy2(data,k1,k2){
    const map = new Map();
    data.forEach(d=>{ const key = `${d[k1]}||${d[k2]}`; if(!map.has(key)) map.set(key,{k1:d[k1],k2:d[k2],pax26:0,pax25:0,sales26:0,sales25:0,profit26:0,profit25:0}); const o=map.get(key); if(d.year==='2026'){o.pax26+=d.pax;o.sales26+=d.sales;o.profit26+=d.profit}else{o.pax25+=d.pax;o.sales25+=d.sales;o.profit25+=d.profit}; });
    return [...map.values()].map(o=>({...o,paxYoY:yoy(o.pax26,o.pax25),salesYoY:yoy(o.sales26,o.sales25)}));
  }
  function renderRanks(){
    const a = aggregate(state.filtered,'dept').filter(x=>x.pax25>0 && x.pax26>10);
    renderRank('deptTop', [...a].sort((x,y)=>y.paxYoY-x.paxYoY).slice(0,7), 'paxYoY');
    renderRank('deptBottom', [...a].sort((x,y)=>x.paxYoY-y.paxYoY).slice(0,7), 'paxYoY', true);
  }
  function renderRegion(){
    const regions = aggregate(state.filtered,'region').sort((a,b)=>b.sales26-a.sales26).slice(0,10);
    drawBar($('chartRegionBar'), regions.map(x=>x.key), regions.map(x=>x.sales26), {horizontal:false, unit:'원', color:COLORS[0]});
    const countries = aggregate(state.filtered,'country').sort((a,b)=>b.pax26-a.pax26).slice(0,15);
    drawBar($('chartCountryTop'), countries.map(x=>x.key), countries.map(x=>x.pax26), {horizontal:true, unit:'명', color:COLORS[2]});
    renderTreemap(regions);
  }
  function renderChannelGrade(){
    const ch = aggregate(state.filtered,'channel').sort((a,b)=>b.sales26-a.sales26).slice(0,8);
    const gr = aggregate(state.filtered,'grade').sort((a,b)=>b.pax26-a.pax26).slice(0,8);
    drawDonut($('chartChannelDonut'), ch.map(x=>({label:x.key,value:x.sales26})), '채널'); renderLegend('channelLegend', ch.map(x=>x.key));
    drawDonut($('chartGradeDonut'), gr.map(x=>({label:x.key,value:x.pax26})), '등급'); renderLegend('gradeLegend', gr.map(x=>x.key));
    drawBar($('chartChannelProfit'), ch.map(x=>x.key), ch.map(x=>x.margin26*100), {horizontal:true, unit:'%', color:COLORS[3]});
    renderRank('channelRank', ch.filter(x=>x.sales25>0).sort((a,b)=>b.salesYoY-a.salesYoY).slice(0,8), 'salesYoY');
  }
  function renderProfit(){
    const country = aggregate(state.filtered,'country').filter(x=>x.pax26>0).sort((a,b)=>b.sales26-a.sales26).slice(0,24);
    drawBubble($('chartBubble'), country);
    const top = aggregate(state.filtered,'country').sort((a,b)=>b.profit26-a.profit26).slice(0,12);
    drawBar($('chartProfitTop'), top.map(x=>x.key), top.map(x=>x.profit26), {horizontal:true, unit:'원', color:COLORS[4]});
  }
  function renderQuality(){
    const keys=['month','hq','dept','region','country','channel','grade'];
    const missing = keys.map(k=>[k, state.filtered.filter(d=>d[k]==='미지정').length]);
    const negProfit = state.filtered.filter(d=>d.year==='2026' && d.profit<0).length;
    const zeroPaxSales = state.filtered.filter(d=>d.year==='2026' && d.pax===0 && d.sales>0).length;
    $('qualityPanel').innerHTML = [
      ['2026 RAW', state.filtered.filter(d=>d.year==='2026').length.toLocaleString()+'행'], ['2025 RAW', state.filtered.filter(d=>d.year==='2025').length.toLocaleString()+'행'],
      ['음수 공헌이익', negProfit.toLocaleString()+'행'], ['모객 0 / 매출 존재', zeroPaxSales.toLocaleString()+'행'],
      ['미지정 컬럼 최대', (missing.sort((a,b)=>b[1]-a[1])[0]?.[1]||0).toLocaleString()+'건'], ['월 범위', unique(state.filtered,'displayMonth').join(' ~ ') || '-'],
      ['운영부서 수', unique(state.filtered,'dept').length.toLocaleString()], ['국가 수', unique(state.filtered,'country').length.toLocaleString()]
    ].map(x=>`<div class="quality-item"><span>${x[0]}</span><b>${x[1]}</b></div>`).join('');
  }
  function renderInsights(){
    const t=total(state.filtered); const depts=aggregate(state.filtered,'dept').filter(x=>x.pax25>0); const countries=aggregate(state.filtered,'country').filter(x=>x.sales26>0); const channels=aggregate(state.filtered,'channel').filter(x=>x.sales25>0);
    const bestDept=[...depts].sort((a,b)=>b.paxYoY-a.paxYoY)[0], riskDept=[...depts].sort((a,b)=>a.paxYoY-b.paxYoY)[0];
    const profitCountry=[...countries].sort((a,b)=>b.profit26-a.profit26)[0], aspCountry=[...countries].sort((a,b)=>b.asp26-a.asp26)[0], channel=[...channels].sort((a,b)=>b.salesYoY-a.salesYoY)[0];
    const cards = [
      ['전체 성과', `모객 YoY ${pct(yoy(t.pax26,t.pax25))}`, `현재 필터 기준 2026 모객은 ${fmt(t.pax26,'명')}, 2025 동월은 ${fmt(t.pax25,'명')}입니다. 매출 YoY는 ${pct(yoy(t.sales26,t.sales25))}입니다.`],
      ['회복 기회', bestDept?.key || '-', bestDept ? `${bestDept.key}의 모객 YoY가 ${pct(bestDept.paxYoY)}로 상대적으로 우수합니다. 유사 상품/채널 확장 후보입니다.` : '비교 가능한 부서 데이터가 부족합니다.'],
      ['리스크 영역', riskDept?.key || '-', riskDept ? `${riskDept.key}의 모객 YoY가 ${pct(riskDept.paxYoY)}입니다. 가격, 좌석, 노출, 채널 믹스 점검이 필요합니다.` : '비교 가능한 부서 데이터가 부족합니다.'],
      ['수익 기여', profitCountry?.key || '-', profitCountry ? `${profitCountry.key}가 공헌이익 ${fmt(profitCountry.profit26)}로 가장 큽니다. 수익성 방어 관점에서 우선 관리 대상입니다.` : '국가 데이터가 부족합니다.'],
      ['고가 상품군', aspCountry?.key || '-', aspCountry ? `${aspCountry.key}의 평균판매가가 ${fmt(aspCountry.asp26)}로 높습니다. 고단가 패키지/프리미엄 상품 검토가 가능합니다.` : 'ASP 산출 데이터가 부족합니다.'],
      ['채널 인사이트', channel?.key || '-', channel ? `${channel.key} 채널의 매출 YoY가 ${pct(channel.salesYoY)}입니다. 효율 채널이면 예산/노출 확대 후보입니다.` : '채널 비교 데이터가 부족합니다.']
    ];
    $('insightCards').innerHTML = cards.map((c,i)=>`<article class="insight-card" style="animation-delay:${i*70}ms"><span class="tag">${escapeHtml(c[0])}</span><h3>${escapeHtml(c[1])}</h3><p>${escapeHtml(c[2])}</p></article>`).join('');
  }

  function renderRank(id, arr, metric, reverse=false){
    const max = Math.max(...arr.map(x=>Math.abs((x[metric]||0)-1)), .01);
    $(id).innerHTML = arr.map(x=>{ const v=x[metric]||0; return `<div class="rank-item"><div><b>${escapeHtml(x.key)}</b><small> 2026 ${fmt(x.pax26,'명')} / 2025 ${fmt(x.pax25,'명')}</small><div class="rank-bar" style="width:${Math.max(4,Math.min(100,Math.abs(v-1)/max*100))}%;background:${v>=1?'linear-gradient(90deg,#2ee879,#39a7ff)':'linear-gradient(90deg,#ff4d5f,#ff9f43)'}"></div></div><strong class="${v>=1?'up':'down'}">${pct(v)}</strong></div>`; }).join('') || '<div class="empty-state">데이터 없음</div>';
  }
  function renderTreemap(arr){
    const totalVal = arr.reduce((s,x)=>s+x.sales26,0) || 1;
    $('treemap').innerHTML = arr.map((x,i)=>`<div class="tree-box" style="flex:${Math.max(.18,x.sales26/totalVal)};background:linear-gradient(135deg,${COLORS[i%COLORS.length]}99,rgba(123,97,255,.25))"><b>${escapeHtml(x.key)}</b><small>매출 ${fmt(x.sales26)}<br>모객 ${fmt(x.pax26,'명')}<br>YoY ${pct(x.salesYoY)}</small></div>`).join('');
  }
  function renderLegend(id, labels){ $(id).innerHTML = labels.map((l,i)=>`<span><i style="background:${COLORS[i%COLORS.length]}"></i>${escapeHtml(l)}</span>`).join(''); }

  function setupCanvas(canvas){ const dpr=window.devicePixelRatio||1; const rect=canvas.getBoundingClientRect(); canvas.width=Math.max(300,rect.width)*dpr; canvas.height=Math.max(220,rect.height)*dpr; const ctx=canvas.getContext('2d'); ctx.scale(dpr,dpr); return {ctx,w:canvas.width/dpr,h:canvas.height/dpr}; }
  function axes(ctx,w,h,pad){ ctx.strokeStyle='rgba(255,255,255,.12)'; ctx.lineWidth=1; for(let i=0;i<5;i++){ const y=pad.t+(h-pad.t-pad.b)*i/4; ctx.beginPath(); ctx.moveTo(pad.l,y); ctx.lineTo(w-pad.r,y); ctx.stroke(); } ctx.fillStyle='#9aa8c7'; ctx.font='12px sans-serif'; }
  function drawCombo(canvas, labels, a, b){ const {ctx,w,h}=setupCanvas(canvas); ctx.clearRect(0,0,w,h); const pad={l:56,r:20,t:26,b:48}; axes(ctx,w,h,pad); const max=Math.max(...a,...b,1); const bw=(w-pad.l-pad.r)/Math.max(labels.length,1)*.34; labels.forEach((lab,i)=>{ const x=pad.l+(i+.5)*(w-pad.l-pad.r)/labels.length; const yA=h-pad.b-(a[i]/max)*(h-pad.t-pad.b); const yB=h-pad.b-(b[i]/max)*(h-pad.t-pad.b); ctx.fillStyle='rgba(57,167,255,.72)'; roundRect(ctx,x-bw-2,yA,bw,h-pad.b-yA,4,true); ctx.fillStyle='rgba(255,92,138,.72)'; roundRect(ctx,x+2,yB,bw,h-pad.b-yB,4,true); ctx.save(); ctx.translate(x,h-pad.b+18); ctx.rotate(-.25); ctx.fillStyle='#cbd6f3'; ctx.font='11px sans-serif'; ctx.textAlign='center'; ctx.fillText(lab,0,0); ctx.restore(); }); legendCanvas(ctx,w,['2026','2025'],[COLORS[0],COLORS[2]]); }
  function drawLine(canvas, labels, values, opt={}){ const {ctx,w,h}=setupCanvas(canvas); ctx.clearRect(0,0,w,h); const pad={l:54,r:22,t:26,b:48}; axes(ctx,w,h,pad); const max=Math.max(...values,1), min=Math.min(...values,0); const range=max-min||1; ctx.strokeStyle=opt.color||COLORS[0]; ctx.lineWidth=3; ctx.beginPath(); values.forEach((v,i)=>{ const x=pad.l+(i+.5)*(w-pad.l-pad.r)/values.length; const y=h-pad.b-((v-min)/range)*(h-pad.t-pad.b); i?ctx.lineTo(x,y):ctx.moveTo(x,y); }); ctx.stroke(); values.forEach((v,i)=>{ const x=pad.l+(i+.5)*(w-pad.l-pad.r)/values.length; const y=h-pad.b-((v-min)/range)*(h-pad.t-pad.b); ctx.fillStyle=opt.color||COLORS[0]; ctx.beginPath(); ctx.arc(x,y,4,0,Math.PI*2); ctx.fill(); ctx.fillStyle='#cbd6f3'; ctx.font='11px sans-serif'; ctx.textAlign='center'; ctx.fillText(labels[i],x,h-pad.b+18); }); }
  function drawBar(canvas, labels, values, opt={}){ const {ctx,w,h}=setupCanvas(canvas); ctx.clearRect(0,0,w,h); const pad={l: opt.horizontal?115:50,r:20,t:20,b: opt.horizontal?24:70}; axes(ctx,w,h,pad); const max=Math.max(...values,1); if(opt.horizontal){ const gap=(h-pad.t-pad.b)/labels.length; labels.forEach((lab,i)=>{ const y=pad.t+i*gap+gap*.18, bh=gap*.62, len=(values[i]/max)*(w-pad.l-pad.r); ctx.fillStyle=opt.color||COLORS[i%COLORS.length]; roundRect(ctx,pad.l,y,len,bh,7,true); ctx.fillStyle='#dce4ff'; ctx.font='12px sans-serif'; ctx.textAlign='right'; ctx.fillText(lab,pad.l-8,y+bh*.7); ctx.textAlign='left'; ctx.fillText(fmt(values[i], opt.unit==='%'?'%':''),pad.l+len+6,y+bh*.7); }); } else { const gap=(w-pad.l-pad.r)/labels.length; labels.forEach((lab,i)=>{ const x=pad.l+i*gap+gap*.18, bw=gap*.62, y=h-pad.b-(values[i]/max)*(h-pad.t-pad.b); ctx.fillStyle=opt.color||COLORS[i%COLORS.length]; roundRect(ctx,x,y,bw,h-pad.b-y,7,true); ctx.save(); ctx.translate(x+bw/2,h-pad.b+16); ctx.rotate(-.45); ctx.fillStyle='#dce4ff'; ctx.font='11px sans-serif'; ctx.textAlign='right'; ctx.fillText(lab,0,0); ctx.restore(); }); } }
  function drawDonut(canvas, data, center){ const {ctx,w,h}=setupCanvas(canvas); ctx.clearRect(0,0,w,h); const cx=w/2, cy=h/2, r=Math.min(w,h)*.34, inner=r*.58; const sum=data.reduce((s,d)=>s+d.value,0)||1; let start=-Math.PI/2; data.forEach((d,i)=>{ const end=start+d.value/sum*Math.PI*2; ctx.beginPath(); ctx.moveTo(cx,cy); ctx.arc(cx,cy,r,start,end); ctx.closePath(); ctx.fillStyle=COLORS[i%COLORS.length]; ctx.fill(); start=end; }); ctx.globalCompositeOperation='destination-out'; ctx.beginPath(); ctx.arc(cx,cy,inner,0,Math.PI*2); ctx.fill(); ctx.globalCompositeOperation='source-over'; ctx.fillStyle='#eef4ff'; ctx.textAlign='center'; ctx.font='800 18px sans-serif'; ctx.fillText(center,cx,cy-4); ctx.font='12px sans-serif'; ctx.fillStyle='#9aa8c7'; ctx.fillText(fmt(sum),cx,cy+18); }
  function drawBubble(canvas, arr){ const {ctx,w,h}=setupCanvas(canvas); ctx.clearRect(0,0,w,h); const pad={l:60,r:28,t:26,b:48}; axes(ctx,w,h,pad); const maxX=Math.max(...arr.map(x=>x.asp26),1), maxY=Math.max(...arr.map(x=>x.margin26),.01), maxS=Math.max(...arr.map(x=>x.sales26),1); arr.forEach((d,i)=>{ const x=pad.l+(d.asp26/maxX)*(w-pad.l-pad.r); const y=h-pad.b-(d.margin26/maxY)*(h-pad.t-pad.b); const r=5+Math.sqrt(d.sales26/maxS)*22; ctx.fillStyle=COLORS[i%COLORS.length]+'aa'; ctx.beginPath(); ctx.arc(x,y,r,0,Math.PI*2); ctx.fill(); ctx.fillStyle='#eef4ff'; ctx.font='11px sans-serif'; ctx.textAlign='center'; if(r>12) ctx.fillText(d.key,x,y+3); }); ctx.fillStyle='#9aa8c7'; ctx.font='12px sans-serif'; ctx.fillText('ASP →',w-60,h-16); ctx.save(); ctx.translate(16,50); ctx.rotate(-Math.PI/2); ctx.fillText('이익률 →',0,0); ctx.restore(); }
  function legendCanvas(ctx,w,labels,colors){ ctx.font='12px sans-serif'; let x=w/2-70; labels.forEach((l,i)=>{ ctx.fillStyle=colors[i]; ctx.fillRect(x,10,12,12); ctx.fillStyle='#cbd6f3'; ctx.fillText(l,x+17,20); x+=70; }); }
  function roundRect(ctx,x,y,w,h,r,fill){ if(h<0){y+=h;h=Math.abs(h)} ctx.beginPath(); ctx.moveTo(x+r,y); ctx.arcTo(x+w,y,x+w,y+h,r); ctx.arcTo(x+w,y+h,x,y+h,r); ctx.arcTo(x,y+h,x,y,r); ctx.arcTo(x,y,x+w,y,r); if(fill)ctx.fill(); else ctx.stroke(); }
  function heatColor(v){ if(!v) return 'rgba(255,255,255,.05)'; if(v>=1.15) return 'linear-gradient(135deg,#2ee879,#39a7ff)'; if(v>=1) return 'linear-gradient(135deg,#8ee6a9,#d1fae5)'; if(v>=.85) return 'linear-gradient(135deg,#ffd166,#ffefb0)'; if(v>=.7) return 'linear-gradient(135deg,#ff9f43,#ffd0a1)'; return 'linear-gradient(135deg,#ff4d5f,#9f1239)'; }
  function escapeHtml(s){ return String(s).replace(/[&<>"]/g, m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[m])); }
  function stripNum(s){ return String(s).replace(/[^0-9.-]/g,''); }

  window.addEventListener('resize', () => state.data.length && renderAll());
  document.addEventListener('DOMContentLoaded', boot);
})();
