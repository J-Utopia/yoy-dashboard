(() => {
  'use strict';

  const XLSX_URL = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
  const COLORS = ['#39a7ff','#7b61ff','#ff5c8a','#2ee879','#ffd166','#ff9f43','#2ff3ff','#a78bfa'];
  const state = { r26: [], r25: [], all: [], filtered: [], slide: 0 };

  const COL = {
    month: ['출발 년월','출발년월'], pax: ['모객'], sales: ['매출(판매가+추가판매)','매출'], profit: ['공헌이익'],
    hq:['운영본부'], dept:['운영부서'], region:['상품지역'], country:['상품국가'], channel:['예약경로 기준 분류'], grade:['상품등급']
  };

  const $ = id => document.getElementById(id);
  const txt = v => (v==null||String(v).trim()==='') ? '미분류' : String(v).trim();
  const num = v => (v==null||v==='') ? 0 : (Number(String(v).replace(/,/g,''))||0);
  const fmtNum = v => Math.round(v||0).toLocaleString('ko-KR');
  const fmtEok = v => `${(Number(v||0)/100000000).toLocaleString('ko-KR',{maximumFractionDigits:1})}억`;

  function log(m){ const el=$('logArea'); if(el){ el.textContent += `\n${m}`; el.scrollTop = el.scrollHeight; } }
  function status(m){ const el=$('statusPill'); if(el) el.textContent = m; log(m); }

  function pick(r, keys){ for(const k of keys){ if(Object.prototype.hasOwnProperty.call(r,k)) return r[k]; } return ''; }
  function normMonth(v, isBase){
    let s = txt(v);
    if(/^\d{6}$/.test(s)) s = `${s.slice(0,4)}-${s.slice(4,6)}`;
    if(/^\d{4}[./]\d{1,2}/.test(s)) s = s.replace(/[.]/g,'-').slice(0,7);
    if(/^\d{4}-\d{1,2}/.test(s)){ const p=s.split('-'); s=`${p[0]}-${String(p[1]).padStart(2,'0')}`; }
    if(isBase && /^2025-\d{2}$/.test(s)) s = `2026-${s.slice(5,7)}`;
    return s;
  }

  function normalize(rows, isBase){
    return rows.map(r=>({
      month:normMonth(pick(r,COL.month),isBase), pax:num(pick(r,COL.pax)), sales:num(pick(r,COL.sales)), profit:num(pick(r,COL.profit)),
      hq:txt(pick(r,COL.hq)), dept:txt(pick(r,COL.dept)), region:txt(pick(r,COL.region)), country:txt(pick(r,COL.country)), channel:txt(pick(r,COL.channel)), grade:txt(pick(r,COL.grade)),
      year:isBase?'2025':'2026'
    })).filter(r=>r.month && (r.pax||r.sales||r.profit));
  }

  function totals(rows){ return rows.reduce((a,r)=>(a.pax+=r.pax,a.sales+=r.sales,a.profit+=r.profit,a),{pax:0,sales:0,profit:0}); }
  function group(rows,key){ const m=new Map(); rows.forEach(r=>{ if(!m.has(r[key])) m.set(r[key],[]); m.get(r[key]).push(r); }); return m; }
  function sumBy(rows,key){ return [...group(rows,key)].map(([k,rs])=>({k,...totals(rs)})); }

  function fit(canvas){ const r=canvas.getBoundingClientRect(); const dpr=window.devicePixelRatio||1; const w=Math.max(320,r.width),h=Math.max(220,r.height); canvas.width=w*dpr; canvas.height=h*dpr; const ctx=canvas.getContext('2d'); ctx.setTransform(dpr,0,0,dpr,0,0); return {ctx,w,h}; }
  function drawAxes(ctx,w,h,p){ ctx.strokeStyle='rgba(255,255,255,.12)'; for(let i=0;i<4;i++){ const y=p.t+(h-p.t-p.b)*i/3; ctx.beginPath(); ctx.moveTo(p.l,y); ctx.lineTo(w-p.r,y); ctx.stroke(); } }

  function drawBars(id, labels, a, b){
    const c=$(id); if(!c) return; const {ctx,w,h}=fit(c); ctx.clearRect(0,0,w,h);
    const p={l:44,r:12,t:12,b:40}; drawAxes(ctx,w,h,p);
    const max=Math.max(1,...a,...b), iw=w-p.l-p.r, ih=h-p.t-p.b, gw=iw/Math.max(labels.length,1), bw=Math.max(4,gw*0.3);
    labels.forEach((lab,i)=>{
      const x=p.l+gw*i+gw*0.15; const ha=ih*(a[i]/max), hb=ih*(b[i]/max);
      ctx.fillStyle=COLORS[0]; ctx.fillRect(x,h-p.b-ha,bw,ha);
      if(b.some(v=>v>0)){ ctx.fillStyle=COLORS[2]; ctx.fillRect(x+bw+2,h-p.b-hb,bw,hb); }
      ctx.fillStyle='#cbd6ef'; ctx.font='11px sans-serif'; ctx.textAlign='center'; ctx.fillText(String(lab).slice(5),x+bw,h-12);
    });
  }

  function drawDonut(id, vals){
    const c=$(id); if(!c) return; const {ctx,w,h}=fit(c); ctx.clearRect(0,0,w,h);
    const cx=w/2, cy=h/2, r=Math.min(w,h)*0.34, ir=r*0.58; const sum=vals.reduce((s,v)=>s+v,0)||1; let st=-Math.PI/2;
    vals.forEach((v,i)=>{ const en=st+(v/sum)*Math.PI*2; ctx.beginPath(); ctx.moveTo(cx,cy); ctx.arc(cx,cy,r,st,en); ctx.closePath(); ctx.fillStyle=COLORS[i%COLORS.length]; ctx.fill(); st=en; });
    ctx.globalCompositeOperation='destination-out'; ctx.beginPath(); ctx.arc(cx,cy,ir,0,Math.PI*2); ctx.fill(); ctx.globalCompositeOperation='source-over';
  }

  function buildFilters(){
    const map=[['filterHQ','hq'],['filterDept','dept'],['filterRegion','region'],['filterCountry','country'],['filterChannel','channel'],['filterGrade','grade'],['filterMonth','month']];
    map.forEach(([id,key])=>{ const vals=[...new Set(state.r26.map(x=>x[key]).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'ko')); $(id).innerHTML='<option value="ALL">전체</option>'+vals.map(v=>`<option value="${v}">${v}</option>`).join(''); });
  }

  function applyFilters(){
    const f={hq:$('filterHQ').value,dept:$('filterDept').value,region:$('filterRegion').value,country:$('filterCountry').value,channel:$('filterChannel').value,grade:$('filterGrade').value,month:$('filterMonth').value};
    state.filtered = state.all.filter(r=>Object.entries(f).every(([k,v])=>v==='ALL'||r[k]===v));
    renderAll();
  }

  function renderHeat(cur,base){
    const months=[...new Set(cur.concat(base).map(x=>x.month))].sort();
    const depts=[...new Set(cur.map(x=>x.dept))].slice(0,15);
    const gc=new Map(), gb=new Map();
    cur.forEach(r=>{ const k=`${r.dept}|${r.month}`; if(!gc.has(k)) gc.set(k,[]); gc.get(k).push(r); });
    base.forEach(r=>{ const k=`${r.dept}|${r.month}`; if(!gb.has(k)) gb.set(k,[]); gb.get(k).push(r); });
    let html='<table class="heat-table"><thead><tr><th>부서</th>'+months.map(m=>`<th>${m.slice(5)}</th>`).join('')+'</tr></thead><tbody>';
    depts.forEach(d=>{ html+=`<tr><td class="heat-row-name">${d}</td>`; months.forEach(m=>{ const c=totals(gc.get(`${d}|${m}`)||[]).pax; const b=totals(gb.get(`${d}|${m}`)||[]).pax; const y=b?c/b:0; html+=`<td>${b?`${(y*100).toFixed(0)}%`:'-'}</td>`;}); html+='</tr>';});
    $('heatmap').innerHTML=html+'</tbody></table>';
  }

  function renderInsights(cur,base){
    const tc=totals(cur), tb=totals(base);
    const yoyP=tb.pax?tc.pax/tb.pax:0, yoyS=tb.sales?tc.sales/tb.sales:0;
    $('insightCards').innerHTML=`<article class="insight-card"><span class="tag">요약</span><h3>모객 YoY ${(yoyP*100).toFixed(1)}%</h3><p>매출 YoY ${(yoyS*100).toFixed(1)}%, 2026 ${fmtNum(tc.pax)}명 / 기준 ${fmtNum(tb.pax)}명</p></article>`;
  }

  function renderAll(){
    const cur=state.filtered.filter(x=>x.year==='2026');
    const base=state.filtered.filter(x=>x.year==='2025');
    const c=totals(cur), b=totals(base);
    $('dataScope').textContent=`자동 로드 완료: 2026 ${cur.length.toLocaleString()}건 / 2025 ${base.length.toLocaleString()}건`;
    $('kpiGrid').innerHTML=`<div class="kpi-card"><div class="kpi-label">총 모객</div><div class="kpi-value">${fmtNum(c.pax)}명</div><div class="kpi-delta">기준 ${fmtNum(b.pax)}명</div></div><div class="kpi-card"><div class="kpi-label">총 매출</div><div class="kpi-value">${fmtEok(c.sales)}</div><div class="kpi-delta">기준 ${fmtEok(b.sales)}</div></div><div class="kpi-card"><div class="kpi-label">공헌이익</div><div class="kpi-value">${fmtEok(c.profit)}</div><div class="kpi-delta">기준 ${fmtEok(b.profit)}</div></div><div class="kpi-card"><div class="kpi-label">분석 건수</div><div class="kpi-value">${fmtNum(cur.length+base.length)}건</div><div class="kpi-delta">필터 반영</div></div>`;

    const m26=group(cur,'month'), m25=group(base,'month');
    const months=[...new Set(cur.concat(base).map(x=>x.month))].sort();
    drawBars('chartMonthlyOverview',months,months.map(m=>totals(m26.get(m)||[]).pax),months.map(m=>totals(m25.get(m)||[]).pax));
    drawBars('chartPaxMonthly',months,months.map(m=>totals(m26.get(m)||[]).pax),months.map(m=>totals(m25.get(m)||[]).pax));
    drawBars('chartRevenueMonthly',months,months.map(m=>totals(m26.get(m)||[]).sales),months.map(m=>totals(m25.get(m)||[]).sales));

    const hq=sumBy(cur,'hq').sort((x,y)=>y.sales-x.sales).slice(0,8); drawDonut('chartHQDonut',hq.map(x=>x.sales));
    const region=sumBy(cur,'region').sort((x,y)=>y.sales-x.sales).slice(0,10); drawBars('chartRegionBar',region.map(x=>x.k),region.map(x=>x.sales),region.map(()=>0));
    const country=sumBy(cur,'country').sort((x,y)=>y.pax-x.pax).slice(0,12); drawBars('chartCountryTop',country.map(x=>x.k),country.map(x=>x.pax),country.map(()=>0));
    const ch=sumBy(cur,'channel').sort((x,y)=>y.sales-x.sales).slice(0,8); drawDonut('chartChannelDonut',ch.map(x=>x.sales));
    const gr=sumBy(cur,'grade').sort((x,y)=>y.pax-x.pax).slice(0,8); drawDonut('chartGradeDonut',gr.map(x=>x.pax));
    const pt=sumBy(cur,'country').sort((x,y)=>y.profit-x.profit).slice(0,12); drawBars('chartProfitTop',pt.map(x=>x.k),pt.map(x=>x.profit),pt.map(()=>0));

    renderHeat(cur,base); renderInsights(cur,base);
  }

  function goSlide(i){ const slides=[...document.querySelectorAll('.slide')], tabs=[...document.querySelectorAll('.tab-btn')]; state.slide=((i%slides.length)+slides.length)%slides.length; slides.forEach((s,ix)=>s.classList.toggle('active',ix===state.slide)); tabs.forEach((t,ix)=>t.classList.toggle('active',ix===state.slide)); }

  async function ensureXlsx(){ if(window.XLSX) return; await new Promise((res,rej)=>{ const s=document.createElement('script'); s.src=XLSX_URL; s.onload=()=>res(); s.onerror=()=>rej(new Error('XLSX 로드 실패')); document.head.appendChild(s);}); }
  async function loadBuffer(ab,label){
    await ensureXlsx();
    const wb=XLSX.read(ab,{type:'array'});
    const ws26=wb.Sheets['RAW_기준일']||wb.Sheets['RAW_기준년'];
    const ws25=wb.Sheets['RAW_YOY'];
    if(!ws26||!ws25) throw new Error('필수 시트 누락: RAW_기준일(또는 RAW_기준년), RAW_YOY');
    state.r26=normalize(XLSX.utils.sheet_to_json(ws26,{defval:''}),false);
    state.r25=normalize(XLSX.utils.sheet_to_json(ws25,{defval:''}),true);
    state.all=state.r26.concat(state.r25);
    buildFilters(); state.filtered=state.all; renderAll(); status(`${label} 분석 완료`);
  }

  async function autoLoad(){ try{ status('기본 엑셀 자동 로드 중...'); const r=await fetch('./data.xlsx',{cache:'no-store'}); if(!r.ok) throw new Error(`HTTP ${r.status}`); await loadBuffer(await r.arrayBuffer(),'data.xlsx'); }catch(e){ status(`자동 로드 실패: ${e.message}`);} }

  function bind(){
    $('analyzeBtn').addEventListener('click', async()=>{ const f=$('fileInput').files?.[0]; if(!f) return alert('먼저 파일을 선택하세요.'); await loadBuffer(await f.arrayBuffer(),f.name);});
    $('sampleBtn').addEventListener('click', autoLoad);
    ['filterHQ','filterDept','filterRegion','filterCountry','filterChannel','filterGrade','filterMonth'].forEach(id=>$(id).addEventListener('change',applyFilters));
    $('resetFilters').addEventListener('click',()=>{ ['filterHQ','filterDept','filterRegion','filterCountry','filterChannel','filterGrade','filterMonth'].forEach(id=>$(id).value='ALL'); applyFilters();});
    document.querySelectorAll('.tab-btn').forEach(b=>b.addEventListener('click',()=>goSlide(Number(b.dataset.slide||0))));
    $('prevSlide').addEventListener('click',()=>goSlide(state.slide-1)); $('nextSlide').addEventListener('click',()=>goSlide(state.slide+1));
    $('fullscreenBtn').addEventListener('click',()=>document.documentElement.requestFullscreen&&document.documentElement.requestFullscreen());
    window.addEventListener('resize',()=>state.all.length&&renderAll());
  }

  document.addEventListener('DOMContentLoaded',()=>{ $('logArea').textContent='app.js 로드 완료'; bind(); goSlide(0); autoLoad(); });
})();
