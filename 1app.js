(() => {
  'use strict';

  const COLS = {
    month: '異쒕컻 ?꾩썡', pax: '紐④컼', sales: '留ㅼ텧(?먮ℓ媛+異붽??먮ℓ)', profit: '怨듯뿄?댁씡',
    hq: '?댁쁺蹂몃?', dept: '?댁쁺遺??, region: '?곹뭹吏??, city: '?꾩갑?꾩떆', country: '?곹뭹援??',
    channel: '?덉빟寃쎈줈 湲곗? 遺꾨쪟', grade: '?곹뭹?깃툒', date: '湲곗??쇱떆'
  };
  const COLORS = ['#39a7ff','#7b61ff','#ff5c8a','#2ee879','#ffd166','#ff9f43','#2ff3ff','#a78bfa','#fb7185','#34d399','#f472b6','#60a5fa'];
  let state = { raw26: [], raw25: [], data: [], filtered: [], filters: {}, slide: 0, heatMetric: 'pax' };

  const $ = id => document.getElementById(id);
  const qs = sel => document.querySelector(sel);
  const qsa = sel => [...document.querySelectorAll(sel)];
  const log = msg => { const el = $('logArea'); if (!el) return; el.textContent += `\n${new Date().toLocaleTimeString()}  ${msg}`; el.scrollTop = el.scrollHeight; };
  const setStatus = (txt, type='neutral') => { const el=$('statusPill'); if(el){ el.textContent=txt; el.className=`status-pill ${type}`; } };
  const num = v => { if (v === null || v === undefined || v === '') return 0; if (typeof v === 'number') return Number.isFinite(v) ? v : 0; const n = Number(String(v).replace(/,/g,'')); return Number.isFinite(n) ? n : 0; };
  const str = v => (v === null || v === undefined || v === '') ? '誘몄??? : String(v).trim();
  const fmt = (n, unit='') => (Math.abs(n) >= 100000000 ? `${(n/100000000).toLocaleString('ko-KR',{maximumFractionDigits:1})}??{unit}` : Math.abs(n) >= 10000 ? `${(n/10000).toLocaleString('ko-KR',{maximumFractionDigits:1})}留?{unit}` : `${Math.round(n).toLocaleString('ko-KR')}${unit}`);
  const pct = v => Number.isFinite(v) ? `${(v*100).toFixed(1)}%` : '-';
  const safeDiv = (a,b) => b ? a/b : 0;
  const yoy = (a,b) => b ? a/b : 0;
  const shiftYear = m => /^\d{4}-\d{2}$/.test(m) ? `${Number(m.slice(0,4))+1}-${m.slice(5,7)}` : m;

  function boot(){
    $('logArea').textContent = 'app.js 濡쒕뱶 ?꾨즺';
    setStatus('app.js 濡쒕뱶 ?꾨즺');
    $('analyzeBtn').addEventListener('click', analyzeSelectedFile);
    $('sampleBtn').addEventListener('click', loadSample);
    $('resetFilters').addEventListener('click', resetFilters);
    $('fullscreenBtn').addEventListener('click', () => { document.body.classList.toggle('fullscreen'); document.fullscreenElement ? document.exitFullscreen?.() : document.documentElement.requestFullscreen?.(); });
    $('prevSlide').addEventListener('click', () => goSlide(state.slide - 1));
    $('nextSlide').addEventListener('click', () => goSlide(state.slide + 1));
    $('toggleHeatMetric').addEventListener('click', () => { state.heatMetric = state.heatMetric === 'pax' ? 'sales' : 'pax'; $('toggleHeatMetric').textContent = state.heatMetric === 'pax' ? '留ㅼ텧 YoY濡??꾪솚' : '紐④컼 YoY濡??꾪솚'; renderAll(); });
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
    if(!file){ alert('癒쇱? ?묒? ?뚯씪???좏깮?섏꽭??'); return; }
    $('logArea').textContent = '?뚯씪 ?좏깮 ?뺤씤: ' + file.name;
    setStatus('?묒? ?쎈뒗 以?);
    if(!window.XLSX){ setStatus('SheetJS 濡쒕뱶 ?ㅽ뙣'); log('XLSX ?쇱씠釉뚮윭由ш? 濡쒕뱶?섏? ?딆븯?듬땲?? ?명꽣??CDN 李⑤떒 媛?μ꽦???덉뒿?덈떎. ?섑뵆 寃利앹? 媛?ν빀?덈떎.'); alert('SheetJS CDN??濡쒕뱶?섏? ?딆븯?듬땲?? ?뚯궗留?CDN 李⑤떒?대㈃ xlsx.full.min.js 濡쒖뺄 ?뚯씪???꾩슂?⑸땲??'); return; }
    try{
      log('ArrayBuffer ?쎄린 ?쒖옉');
      const buffer = await file.arrayBuffer();
      log(`ArrayBuffer ?쎄린 ?꾨즺: ${Math.round(buffer.byteLength/1024).toLocaleString()} KB`);
      const wb = XLSX.read(buffer, {type:'array', cellDates:false, raw:false});
      log('?쒗듃 紐⑸줉: ' + wb.SheetNames.join(', '));
      const ws26 = wb.Sheets['RAW_湲곗???];
      const ws25 = wb.Sheets['RAW_YOY'];
      if(!ws26 || !ws25) throw new Error('RAW_湲곗????먮뒗 RAW_YOY ?쒗듃瑜?李얠쓣 ???놁뒿?덈떎.');
      const raw26 = XLSX.utils.sheet_to_json(ws26, {defval:''});
      const raw25 = XLSX.utils.sheet_to_json(ws25, {defval:''});
      log(`RAW_湲곗????됱닔: ${raw26.length.toLocaleString()} / RAW_YOY ?됱닔: ${raw25.length.toLocaleString()}`);
      buildDataset(raw26, raw25, file.name);
    }catch(err){ console.error(err); setStatus('遺꾩꽍 ?ㅽ뙣'); log('?ㅻ쪟: ' + err.message); alert('遺꾩꽍 ?ㅽ뙣: ' + err.message); }
  }

  function loadSample(){
    $('logArea').textContent = '?섑뵆 ?곗씠???앹꽦 ?쒖옉';
    const months26 = ['2026-05','2026-06','2026-07','2026-08','2026-09','2026-10','2026-11','2026-12','2027-01'];
    const hqs = ['?곹뭹1蹂몃?','?곹뭹2蹂몃?','?곹뭹3蹂몃?'];
    const depts = ['?숇궓?꾩궗??遺','?쇰낯?ъ뾽遺','?좊읇?ъ뾽遺','誘몄＜?ъ뾽遺','以묎뎅?ъ뾽遺','?⑦깭?됱뼇?ъ뾽遺'];
    const regions = ['?섑듃??,'?쇰낯','?쒖쑀??,'誘몄＜','?κ?怨?,'愿??ъ씠??];
    const countries = ['踰좏듃??,'?쇰낯','?꾨옉??,'誘멸뎅','以묎뎅','愿?];
    const channels = ['?由ъ젏','紐⑤몢?룹뺨','?쒗쑕??,'?뱀삁???곸뾽怨듯넻)','紐⑤몢?⑥뼱(?꾩떆怨듯넻)'];
    const grades = ['?놁쓬','?쒓렇?덉쿂','?쒓렇?덉쿂釉붾옓','?섏씠?대옒??];
    const make = (yearShift=0) => {
      const rows=[];
      for(let i=0;i<months26.length;i++) for(let j=0;j<depts.length;j++) for(let k=0;k<channels.length;k++){
        const m26 = months26[i]; const yyyy = Number(m26.slice(0,4))-yearShift; const m = `${yyyy}-${m26.slice(5,7)}`;
        const base = (depts.length-j)*8 + (i<3?40:18) + k*3;
        const factor = yearShift ? (1.25 + (j%3)*.08) : (0.85 + (i%4)*.1 + (j%2)*.08);
        const pax = Math.max(1, Math.round(base*factor));
        const asp = 750000 + j*180000 + i*25000 + k*40000;
        rows.push({[COLS.date]: `${yyyy}-05-18`, [COLS.month]: m, [COLS.pax]: pax, [COLS.sales]: pax*asp, [COLS.profit]: pax*asp*(.08+(j%4)*.025), [COLS.hq]: hqs[j%hqs.length], [COLS.dept]: depts[j], [COLS.region]: regions[j], [COLS.city]: regions[j], [COLS.country]: countries[j], [COLS.channel]: channels[k], [COLS.grade]: grades[(j+k)%grades.length], '?먮ℓ?띿꽦':'01.?댁쇅PKG'});
      }
      return rows;
    };
    buildDataset(make(0), make(1), '?섑뵆 ?곗씠??);
  }

  function normalizeRows(rows, yearLabel){
    return rows.map((r, idx) => {
      const month = str(r[COLS.month]);
      return {
        _idx: idx, year: yearLabel, displayMonth: yearLabel === '2025' ? shiftYear(month) : month, month,
        pax: num(r[COLS.pax]), sales: num(r[COLS.sales]), profit: num(r[COLS.profit]),
        hq: str(r[COLS.hq]), dept: str(r[COLS.dept]), region: str(r[COLS.region]), city: str(r[COLS.city]), country: str(r[COLS.country]), channel: str(r[COLS.channel]), grade: str(r[COLS.grade])
      };
    }).filter(r => r.displayMonth !== '誘몄??? && (r.pax || r.sales || r.profit));
  }

  function buildDataset(raw26, raw25, filename){
    state.raw26 = normalizeRows(raw26, '2026');
    state.raw25 = normalizeRows(raw25, '2025');
    state.data = [...state.raw26, ...state.raw25];
    log(`?뺢퇋???꾨즺: 2026 ${state.raw26.length.toLocaleString()}??/ 2025 ${state.raw25.length.toLocaleString()}??);
    if(!state.raw26.length || !state.raw25.length) log('二쇱쓽: ?쒖そ ?곕룄 ?곗씠?곌? 鍮꾩뼱 ?덉뒿?덈떎. YoY媛 ?쒗븳?⑸땲??');
    $('dataScope').textContent = `${filename} 쨌 RAW_湲곗???${state.raw26.length.toLocaleString()}??/ RAW_YOY ${state.raw25.length.toLocaleString()}??;
    buildFilters();
    applyFiltersAndRender();
    setStatus('遺꾩꽍 ?꾨즺');
    log('??쒕낫???뚮뜑留??꾨즺');
  }

  function unique(data, key){ return [...new Set(data.map(d=>d[key]).filter(Boolean))].sort((a,b)=>String(a).localeCompare(String(b),'ko')); }
  function fillSelect(id, values, selected='ALL'){
    const el=$(id); const label=el.options[0]?.textContent || '?꾩껜'; el.innerHTML = `<option value="ALL">${label}</option>` + values.map(v=>`<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join(''); el.value = values.includes(selected) ? selected : 'ALL';
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
    data.forEach(d => { const k = d[key] || '誘몄???; if(!m.has(k)) m.set(k,{key:k,pax26:0,pax25:0,sales26:0,sales25:0,profit26:0,profit25:0,count:0}); const o=m.get(k); if(d.year==='2026'){o.pax26+=d.pax;o.sales26+=d.sales;o.profit26+=d.profit;} else {o.pax25+=d.pax;o.sales25+=d.sales;o.profit25+=d.profit;} o.count++; });
    return [...m.values()].map(o=>({...o, paxYoY:yoy(o.pax26,o.pax25), salesYoY:yoy(o.sales26,o.sales25), asp26:safeDiv(o.sales26,o.pax26), margin26:safeDiv(o.profit26,o.sales26)}));
  }
  function total(data){ return data.reduce((a,d)=>{ if(d.year==='2026'){a.pax26+=d.pax;a.sales26+=d.sales;a.profit26+=d.profit;} else {a.pax25+=d.pax;a.sales25+=d.sales;a.profit25+=d.profit;} return a; },{pax26:0,pax25:0,sales26:0,sales25:0,profit26:0,profit25:0}); }

  function renderAll(){
    const t = total(state.filtered);
    renderKpis(t); renderMonthly(); renderDonuts(); renderHeatmap(); renderRanks(); renderRegion(); renderChannelGrade(); renderProfit(); renderQuality(); renderInsights();
  }
  function renderEmpty(){ $('kpiGrid').innerHTML = '<div class="empty-state" style="grid-column:1/-1">?묒? ?뚯씪???좏깮?섍굅???섑뵆 ?곗씠?곕줈 ?붾㈃??寃利앺븯?몄슂.</div>'; }

  function renderKpis(t){
    const cards = [
      ['珥?紐④컼', fmt(t.pax26,'紐?), `YoY ${pct(yoy(t.pax26,t.pax25))}`, yoy(t.pax26,t.pax25)-1],
      ['珥?留ㅼ텧', fmt(t.sales26), `YoY ${pct(yoy(t.sales26,t.sales25))}`, yoy(t.sales26,t.sales25)-1],
      ['?됯퇏?먮ℓ媛', fmt(safeDiv(t.sales26,t.pax26)), '留ㅼ텧 첨 紐④컼 湲곗?', 0],
      ['怨듯뿄?댁씡瑜?, pct(safeDiv(t.profit26,t.sales26)), `怨듯뿄?댁씡 ${fmt(t.profit26)}`, safeDiv(t.profit26,t.sales26)-safeDiv(t.profit25,t.sales25)],
      ['2025 紐④컼', fmt(t.pax25,'紐?), 'RAW_YOY ?숈썡 湲곗?', 0],
      ['紐④컼 利앷컧', fmt(t.pax26-t.pax25,'紐?), `${t.pax26>=t.pax25?'利앷?':'媛먯냼'}`, t.pax26-t.pax25],
      ['留ㅼ텧 利앷컧', fmt(t.sales26-t.sales25), `${t.sales26>=t.sales25?'利앷?':'媛먯냼'}`, t.sales26-t.sales25],
      ['?곗씠???됱닔', state.filtered.length.toLocaleString(), '?꾪꽣 ?곸슜 ??, 0]
    ];
    $('kpiGrid').innerHTML = cards.map((c,i)=>`<div class="kpi-card" style="animation-delay:${i*45}ms"><div class="kpi-label">${c[0]}</div><div class="kpi-value" data-target="${stripNum(c[1])}">${c[1]}</div><div class="kpi-delta ${c[3]>0?'up':c[3]<0?'down':'neutral'}">${c[2]}</div></div>`).join('');
  }

  function renderMonthly(){
    const m = aggregate(state.filtered, 'displayMonth').sort((a,b)=>a.key.localeCompare(b.key));
    drawCombo($('chartMonthlyOverview'), m.map(x=>x.key), m.map(x=>x.pax26), m.map(x=>x.pax25), {title:'紐④컼', unit:'紐?});
    drawCombo($('chartPaxMonthly'), m.map(x=>x.key), m.map(x=>x.pax26), m.map(x=>x.pax25), {title:'紐④컼', unit:'紐?});
    drawCombo($('chartRevenueMonthly'), m.map(x=>x.key), m.map(x=>x.sales26), m.map(x=>x.sales25), {title:'留ㅼ텧', unit:'??});
    drawLine($('chartASP'), m.map(x=>x.key), m.map(x=>safeDiv(x.sales26,x.pax26)), {unit:'??, color:COLORS[6]});
    drawLine($('chartMargin'), m.map(x=>x.key), m.map(x=>safeDiv(x.profit26,x.sales26)*100), {unit:'%', color:COLORS[3]});
    const best = [...m].sort((a,b)=>b.paxYoY-a.paxYoY)[0], worst = [...m].filter(x=>x.pax25).sort((a,b)=>a.paxYoY-b.paxYoY)[0];
    $('monthlyInsights').innerHTML = [best&&`<div class="rank-item"><b>理쒓퀬 ?뚮났?? ${best.key}</b><small>紐④컼 YoY ${pct(best.paxYoY)}</small></div>`, worst&&`<div class="rank-item"><b>二쇱쓽?? ${worst.key}</b><small>紐④컼 YoY ${pct(worst.paxYoY)}</small></div>`, `<div class="rank-item"><b>?붿닔</b><small>${m.length}媛?異쒕컻??鍮꾧탳</small></div>`].filter(Boolean).join('');
  }
  function renderDonuts(){
    const hq = aggregate(state.filtered.filter(d=>d.year==='2026'), 'hq').sort((a,b)=>b.sales26-a.sales26).slice(0,8);
    drawDonut($('chartHQDonut'), hq.map(x=>({label:x.key,value:x.sales26})), '留ㅼ텧'); renderLegend('hqLegend', hq.map(x=>x.key));
  }
  function renderHeatmap(){
    const months = unique(state.filtered,'displayMonth');
    const depts = aggregate(state.filtered,'dept').sort((a,b)=>b.pax26-a.pax26).slice(0,18).map(x=>x.key);
    const by = new Map();
    aggregateBy2(state.filtered,'dept','displayMonth').forEach(o=>by.set(`${o.k1}||${o.k2}`,o));
    const metric = state.heatMetric === 'pax' ? 'paxYoY' : 'salesYoY';
    let html = '<table class="heat-table"><thead><tr><th>?댁쁺遺??/th>' + months.map(m=>`<th>${m}</th>`).join('') + '</tr></thead><tbody>';
    depts.forEach(d => { html += `<tr><td class="heat-row-name">${escapeHtml(d)}</td>`; months.forEach(m=>{ const o=by.get(`${d}||${m}`)||{}; const v=o[metric]||0; html += `<td title="${escapeHtml(d)} ${m} ${state.heatMetric==='pax'?'紐④컼':'留ㅼ텧'} YoY ${pct(v)}" style="background:${heatColor(v)};color:${v<.75?'#fff':'#08101f'}">${o.pax25||o.sales25?pct(v):'-'}</td>`; }); html+='</tr>'; });
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
    drawBar($('chartRegionBar'), regions.map(x=>x.key), regions.map(x=>x.sales26), {horizontal:false, unit:'??, color:COLORS[0]});
    const countries = aggregate(state.filtered,'country').sort((a,b)=>b.pax26-a.pax26).slice(0,15);
    drawBar($('chartCountryTop'), countries.map(x=>x.key), countries.map(x=>x.pax26), {horizontal:true, unit:'紐?, color:COLORS[2]});
    renderTreemap(regions);
  }
  function renderChannelGrade(){
    const ch = aggregate(state.filtered,'channel').sort((a,b)=>b.sales26-a.sales26).slice(0,8);
    const gr = aggregate(state.filtered,'grade').sort((a,b)=>b.pax26-a.pax26).slice(0,8);
    drawDonut($('chartChannelDonut'), ch.map(x=>({label:x.key,value:x.sales26})), '梨꾨꼸'); renderLegend('channelLegend', ch.map(x=>x.key));
    drawDonut($('chartGradeDonut'), gr.map(x=>({label:x.key,value:x.pax26})), '?깃툒'); renderLegend('gradeLegend', gr.map(x=>x.key));
    drawBar($('chartChannelProfit'), ch.map(x=>x.key), ch.map(x=>x.margin26*100), {horizontal:true, unit:'%', color:COLORS[3]});
    renderRank('channelRank', ch.filter(x=>x.sales25>0).sort((a,b)=>b.salesYoY-a.salesYoY).slice(0,8), 'salesYoY');
  }
  function renderProfit(){
    const country = aggregate(state.filtered,'country').filter(x=>x.pax26>0).sort((a,b)=>b.sales26-a.sales26).slice(0,24);
    drawBubble($('chartBubble'), country);
    const top = aggregate(state.filtered,'country').sort((a,b)=>b.profit26-a.profit26).slice(0,12);
    drawBar($('chartProfitTop'), top.map(x=>x.key), top.map(x=>x.profit26), {horizontal:true, unit:'??, color:COLORS[4]});
  }
  function renderQuality(){
    const keys=['month','hq','dept','region','country','channel','grade'];
    const missing = keys.map(k=>[k, state.filtered.filter(d=>d[k]==='誘몄???).length]);
    const negProfit = state.filtered.filter(d=>d.year==='2026' && d.profit<0).length;
    const zeroPaxSales = state.filtered.filter(d=>d.year==='2026' && d.pax===0 && d.sales>0).length;
    $('qualityPanel').innerHTML = [
      ['2026 RAW', state.filtered.filter(d=>d.year==='2026').length.toLocaleString()+'??], ['2025 RAW', state.filtered.filter(d=>d.year==='2025').length.toLocaleString()+'??],
      ['?뚯닔 怨듯뿄?댁씡', negProfit.toLocaleString()+'??], ['紐④컼 0 / 留ㅼ텧 議댁옱', zeroPaxSales.toLocaleString()+'??],
      ['誘몄???而щ읆 理쒕?', (missing.sort((a,b)=>b[1]-a[1])[0]?.[1]||0).toLocaleString()+'嫄?], ['??踰붿쐞', unique(state.filtered,'displayMonth').join(' ~ ') || '-'],
      ['?댁쁺遺????, unique(state.filtered,'dept').length.toLocaleString()], ['援?? ??, unique(state.filtered,'country').length.toLocaleString()]
    ].map(x=>`<div class="quality-item"><span>${x[0]}</span><b>${x[1]}</b></div>`).join('');
  }
  function renderInsights(){
    const t=total(state.filtered); const depts=aggregate(state.filtered,'dept').filter(x=>x.pax25>0); const countries=aggregate(state.filtered,'country').filter(x=>x.sales26>0); const channels=aggregate(state.filtered,'channel').filter(x=>x.sales25>0);
    const bestDept=[...depts].sort((a,b)=>b.paxYoY-a.paxYoY)[0], riskDept=[...depts].sort((a,b)=>a.paxYoY-b.paxYoY)[0];
    const profitCountry=[...countries].sort((a,b)=>b.profit26-a.profit26)[0], aspCountry=[...countries].sort((a,b)=>b.asp26-a.asp26)[0], channel=[...channels].sort((a,b)=>b.salesYoY-a.salesYoY)[0];
    const cards = [
      ['?꾩껜 ?깃낵', `紐④컼 YoY ${pct(yoy(t.pax26,t.pax25))}`, `?꾩옱 ?꾪꽣 湲곗? 2026 紐④컼? ${fmt(t.pax26,'紐?)}, 2025 ?숈썡? ${fmt(t.pax25,'紐?)}?낅땲?? 留ㅼ텧 YoY??${pct(yoy(t.sales26,t.sales25))}?낅땲??`],
      ['?뚮났 湲고쉶', bestDept?.key || '-', bestDept ? `${bestDept.key}??紐④컼 YoY媛 ${pct(bestDept.paxYoY)}濡??곷??곸쑝濡??곗닔?⑸땲?? ?좎궗 ?곹뭹/梨꾨꼸 ?뺤옣 ?꾨낫?낅땲??` : '鍮꾧탳 媛?ν븳 遺???곗씠?곌? 遺議깊빀?덈떎.'],
      ['由ъ뒪???곸뿭', riskDept?.key || '-', riskDept ? `${riskDept.key}??紐④컼 YoY媛 ${pct(riskDept.paxYoY)}?낅땲?? 媛寃? 醫뚯꽍, ?몄텧, 梨꾨꼸 誘뱀뒪 ?먭????꾩슂?⑸땲??` : '鍮꾧탳 媛?ν븳 遺???곗씠?곌? 遺議깊빀?덈떎.'],
      ['?섏씡 湲곗뿬', profitCountry?.key || '-', profitCountry ? `${profitCountry.key}媛 怨듯뿄?댁씡 ${fmt(profitCountry.profit26)}濡?媛???쎈땲?? ?섏씡??諛⑹뼱 愿?먯뿉???곗꽑 愿由???곸엯?덈떎.` : '援?? ?곗씠?곌? 遺議깊빀?덈떎.'],
      ['怨좉? ?곹뭹援?, aspCountry?.key || '-', aspCountry ? `${aspCountry.key}???됯퇏?먮ℓ媛媛 ${fmt(aspCountry.asp26)}濡??믪뒿?덈떎. 怨좊떒媛 ?⑦궎吏/?꾨━誘몄뾼 ?곹뭹 寃?좉? 媛?ν빀?덈떎.` : 'ASP ?곗텧 ?곗씠?곌? 遺議깊빀?덈떎.'],
      ['梨꾨꼸 ?몄궗?댄듃', channel?.key || '-', channel ? `${channel.key} 梨꾨꼸??留ㅼ텧 YoY媛 ${pct(channel.salesYoY)}?낅땲?? ?⑥쑉 梨꾨꼸?대㈃ ?덉궛/?몄텧 ?뺣? ?꾨낫?낅땲??` : '梨꾨꼸 鍮꾧탳 ?곗씠?곌? 遺議깊빀?덈떎.']
    ];
    $('insightCards').innerHTML = cards.map((c,i)=>`<article class="insight-card" style="animation-delay:${i*70}ms"><span class="tag">${escapeHtml(c[0])}</span><h3>${escapeHtml(c[1])}</h3><p>${escapeHtml(c[2])}</p></article>`).join('');
  }

  function renderRank(id, arr, metric, reverse=false){
    const max = Math.max(...arr.map(x=>Math.abs((x[metric]||0)-1)), .01);
    $(id).innerHTML = arr.map(x=>{ const v=x[metric]||0; return `<div class="rank-item"><div><b>${escapeHtml(x.key)}</b><small> 2026 ${fmt(x.pax26,'紐?)} / 2025 ${fmt(x.pax25,'紐?)}</small><div class="rank-bar" style="width:${Math.max(4,Math.min(100,Math.abs(v-1)/max*100))}%;background:${v>=1?'linear-gradient(90deg,#2ee879,#39a7ff)':'linear-gradient(90deg,#ff4d5f,#ff9f43)'}"></div></div><strong class="${v>=1?'up':'down'}">${pct(v)}</strong></div>`; }).join('') || '<div class="empty-state">?곗씠???놁쓬</div>';
  }
  function renderTreemap(arr){
    const totalVal = arr.reduce((s,x)=>s+x.sales26,0) || 1;
    $('treemap').innerHTML = arr.map((x,i)=>`<div class="tree-box" style="flex:${Math.max(.18,x.sales26/totalVal)};background:linear-gradient(135deg,${COLORS[i%COLORS.length]}99,rgba(123,97,255,.25))"><b>${escapeHtml(x.key)}</b><small>留ㅼ텧 ${fmt(x.sales26)}<br>紐④컼 ${fmt(x.pax26,'紐?)}<br>YoY ${pct(x.salesYoY)}</small></div>`).join('');
  }
  function renderLegend(id, labels){ $(id).innerHTML = labels.map((l,i)=>`<span><i style="background:${COLORS[i%COLORS.length]}"></i>${escapeHtml(l)}</span>`).join(''); }

  function setupCanvas(canvas){ const dpr=window.devicePixelRatio||1; const rect=canvas.getBoundingClientRect(); canvas.width=Math.max(300,rect.width)*dpr; canvas.height=Math.max(220,rect.height)*dpr; const ctx=canvas.getContext('2d'); ctx.scale(dpr,dpr); return {ctx,w:canvas.width/dpr,h:canvas.height/dpr}; }
  function axes(ctx,w,h,pad){ ctx.strokeStyle='rgba(255,255,255,.12)'; ctx.lineWidth=1; for(let i=0;i<5;i++){ const y=pad.t+(h-pad.t-pad.b)*i/4; ctx.beginPath(); ctx.moveTo(pad.l,y); ctx.lineTo(w-pad.r,y); ctx.stroke(); } ctx.fillStyle='#9aa8c7'; ctx.font='12px sans-serif'; }
  function drawCombo(canvas, labels, a, b){ const {ctx,w,h}=setupCanvas(canvas); ctx.clearRect(0,0,w,h); const pad={l:56,r:20,t:26,b:48}; axes(ctx,w,h,pad); const max=Math.max(...a,...b,1); const bw=(w-pad.l-pad.r)/Math.max(labels.length,1)*.34; labels.forEach((lab,i)=>{ const x=pad.l+(i+.5)*(w-pad.l-pad.r)/labels.length; const yA=h-pad.b-(a[i]/max)*(h-pad.t-pad.b); const yB=h-pad.b-(b[i]/max)*(h-pad.t-pad.b); ctx.fillStyle='rgba(57,167,255,.72)'; roundRect(ctx,x-bw-2,yA,bw,h-pad.b-yA,4,true); ctx.fillStyle='rgba(255,92,138,.72)'; roundRect(ctx,x+2,yB,bw,h-pad.b-yB,4,true); ctx.save(); ctx.translate(x,h-pad.b+18); ctx.rotate(-.25); ctx.fillStyle='#cbd6f3'; ctx.font='11px sans-serif'; ctx.textAlign='center'; ctx.fillText(lab,0,0); ctx.restore(); }); legendCanvas(ctx,w,['2026','2025'],[COLORS[0],COLORS[2]]); }
  function drawLine(canvas, labels, values, opt={}){ const {ctx,w,h}=setupCanvas(canvas); ctx.clearRect(0,0,w,h); const pad={l:54,r:22,t:26,b:48}; axes(ctx,w,h,pad); const max=Math.max(...values,1), min=Math.min(...values,0); const range=max-min||1; ctx.strokeStyle=opt.color||COLORS[0]; ctx.lineWidth=3; ctx.beginPath(); values.forEach((v,i)=>{ const x=pad.l+(i+.5)*(w-pad.l-pad.r)/values.length; const y=h-pad.b-((v-min)/range)*(h-pad.t-pad.b); i?ctx.lineTo(x,y):ctx.moveTo(x,y); }); ctx.stroke(); values.forEach((v,i)=>{ const x=pad.l+(i+.5)*(w-pad.l-pad.r)/values.length; const y=h-pad.b-((v-min)/range)*(h-pad.t-pad.b); ctx.fillStyle=opt.color||COLORS[0]; ctx.beginPath(); ctx.arc(x,y,4,0,Math.PI*2); ctx.fill(); ctx.fillStyle='#cbd6f3'; ctx.font='11px sans-serif'; ctx.textAlign='center'; ctx.fillText(labels[i],x,h-pad.b+18); }); }
  function drawBar(canvas, labels, values, opt={}){ const {ctx,w,h}=setupCanvas(canvas); ctx.clearRect(0,0,w,h); const pad={l: opt.horizontal?115:50,r:20,t:20,b: opt.horizontal?24:70}; axes(ctx,w,h,pad); const max=Math.max(...values,1); if(opt.horizontal){ const gap=(h-pad.t-pad.b)/labels.length; labels.forEach((lab,i)=>{ const y=pad.t+i*gap+gap*.18, bh=gap*.62, len=(values[i]/max)*(w-pad.l-pad.r); ctx.fillStyle=opt.color||COLORS[i%COLORS.length]; roundRect(ctx,pad.l,y,len,bh,7,true); ctx.fillStyle='#dce4ff'; ctx.font='12px sans-serif'; ctx.textAlign='right'; ctx.fillText(lab,pad.l-8,y+bh*.7); ctx.textAlign='left'; ctx.fillText(fmt(values[i], opt.unit==='%'?'%':''),pad.l+len+6,y+bh*.7); }); } else { const gap=(w-pad.l-pad.r)/labels.length; labels.forEach((lab,i)=>{ const x=pad.l+i*gap+gap*.18, bw=gap*.62, y=h-pad.b-(values[i]/max)*(h-pad.t-pad.b); ctx.fillStyle=opt.color||COLORS[i%COLORS.length]; roundRect(ctx,x,y,bw,h-pad.b-y,7,true); ctx.save(); ctx.translate(x+bw/2,h-pad.b+16); ctx.rotate(-.45); ctx.fillStyle='#dce4ff'; ctx.font='11px sans-serif'; ctx.textAlign='right'; ctx.fillText(lab,0,0); ctx.restore(); }); } }
  function drawDonut(canvas, data, center){ const {ctx,w,h}=setupCanvas(canvas); ctx.clearRect(0,0,w,h); const cx=w/2, cy=h/2, r=Math.min(w,h)*.34, inner=r*.58; const sum=data.reduce((s,d)=>s+d.value,0)||1; let start=-Math.PI/2; data.forEach((d,i)=>{ const end=start+d.value/sum*Math.PI*2; ctx.beginPath(); ctx.moveTo(cx,cy); ctx.arc(cx,cy,r,start,end); ctx.closePath(); ctx.fillStyle=COLORS[i%COLORS.length]; ctx.fill(); start=end; }); ctx.globalCompositeOperation='destination-out'; ctx.beginPath(); ctx.arc(cx,cy,inner,0,Math.PI*2); ctx.fill(); ctx.globalCompositeOperation='source-over'; ctx.fillStyle='#eef4ff'; ctx.textAlign='center'; ctx.font='800 18px sans-serif'; ctx.fillText(center,cx,cy-4); ctx.font='12px sans-serif'; ctx.fillStyle='#9aa8c7'; ctx.fillText(fmt(sum),cx,cy+18); }
  function drawBubble(canvas, arr){ const {ctx,w,h}=setupCanvas(canvas); ctx.clearRect(0,0,w,h); const pad={l:60,r:28,t:26,b:48}; axes(ctx,w,h,pad); const maxX=Math.max(...arr.map(x=>x.asp26),1), maxY=Math.max(...arr.map(x=>x.margin26),.01), maxS=Math.max(...arr.map(x=>x.sales26),1); arr.forEach((d,i)=>{ const x=pad.l+(d.asp26/maxX)*(w-pad.l-pad.r); const y=h-pad.b-(d.margin26/maxY)*(h-pad.t-pad.b); const r=5+Math.sqrt(d.sales26/maxS)*22; ctx.fillStyle=COLORS[i%COLORS.length]+'aa'; ctx.beginPath(); ctx.arc(x,y,r,0,Math.PI*2); ctx.fill(); ctx.fillStyle='#eef4ff'; ctx.font='11px sans-serif'; ctx.textAlign='center'; if(r>12) ctx.fillText(d.key,x,y+3); }); ctx.fillStyle='#9aa8c7'; ctx.font='12px sans-serif'; ctx.fillText('ASP ??,w-60,h-16); ctx.save(); ctx.translate(16,50); ctx.rotate(-Math.PI/2); ctx.fillText('?댁씡瑜???,0,0); ctx.restore(); }
  function legendCanvas(ctx,w,labels,colors){ ctx.font='12px sans-serif'; let x=w/2-70; labels.forEach((l,i)=>{ ctx.fillStyle=colors[i]; ctx.fillRect(x,10,12,12); ctx.fillStyle='#cbd6f3'; ctx.fillText(l,x+17,20); x+=70; }); }
  function roundRect(ctx,x,y,w,h,r,fill){ if(h<0){y+=h;h=Math.abs(h)} ctx.beginPath(); ctx.moveTo(x+r,y); ctx.arcTo(x+w,y,x+w,y+h,r); ctx.arcTo(x+w,y+h,x,y+h,r); ctx.arcTo(x,y+h,x,y,r); ctx.arcTo(x,y,x+w,y,r); if(fill)ctx.fill(); else ctx.stroke(); }
  function heatColor(v){ if(!v) return 'rgba(255,255,255,.05)'; if(v>=1.15) return 'linear-gradient(135deg,#2ee879,#39a7ff)'; if(v>=1) return 'linear-gradient(135deg,#8ee6a9,#d1fae5)'; if(v>=.85) return 'linear-gradient(135deg,#ffd166,#ffefb0)'; if(v>=.7) return 'linear-gradient(135deg,#ff9f43,#ffd0a1)'; return 'linear-gradient(135deg,#ff4d5f,#9f1239)'; }
  function escapeHtml(s){ return String(s).replace(/[&<>"]/g, m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[m])); }
  function stripNum(s){ return String(s).replace(/[^0-9.-]/g,''); }

  window.addEventListener('resize', () => state.data.length && renderAll());
  document.addEventListener('DOMContentLoaded', boot);
})();

