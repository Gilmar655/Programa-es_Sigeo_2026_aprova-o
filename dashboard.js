/* Programações 2026 | Visualização independente de bibliotecas externas.
   Os registros são sempre lidos do Excel pelo servidor.py quando executado localmente.
*/
(() => {
  'use strict';
  const $ = (id) => document.getElementById(id);
  const BR = new Intl.NumberFormat('pt-BR');
  const COLLATOR = new Intl.Collator('pt-BR', {numeric:true, sensitivity:'base'});
  const MONTHS=['jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez'];
  const COLORS=['#1587df','#10c3cd','#8b73e4','#f2ac4e','#21b783','#df7190','#7d9abf'];
  const TYPE_COLOR={FI:'#1587df',BRA:'#10c3cd',PIE:'#8b73e4'};
  const fields={responsavel:'f-responsavel',tipo_intervencao:'f-tipo',regional:'f-regional',status_programacao:'f-status',status_gilmar:'f-gilmar'};
  const state={all:[],filtered:[],revision:null,mode:'week',sortKey:'data',sortDir:1,page:1,pageSize:10,isLive:false,snapshot:true};
  let toastTimer=null;
  const safe=(x)=>String(x??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const clean=(x)=>String(x??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().trim();
  const unique=(arr)=>[...new Set(arr.filter(v=>String(v??'').trim()!==''))].sort(COLLATOR.compare);
  const countFmt=(n)=>BR.format(n);
  function dateLocal(iso){if(!/^\d{4}-\d{2}-\d{2}$/.test(String(iso??'')))return null;const d=new Date(iso+'T12:00:00');return Number.isNaN(+d)?null:d;}
  function isoLocal(d){return [d.getFullYear(),String(d.getMonth()+1).padStart(2,'0'),String(d.getDate()).padStart(2,'0')].join('-');}
  function todayISO(){return isoLocal(new Date());}
  function dateBR(value){const d=dateLocal(value);return d?d.toLocaleDateString('pt-BR'):(value||'—');}
  function niceMonth(iso){const d=dateLocal(iso+'-01');return d?`${MONTHS[d.getMonth()].toUpperCase()}/${String(d.getFullYear()).slice(2)}`:iso;}
  function fmtTimestamp(value){if(!value)return '—';let d=new Date(value);return Number.isNaN(+d)?String(value):d.toLocaleString('pt-BR',{day:'2-digit',month:'2-digit',year:'numeric',hour:'2-digit',minute:'2-digit',second:'2-digit'});}
  function toast(message){const node=$('toast');node.textContent=message;node.classList.add('show');if(toastTimer)clearTimeout(toastTimer);toastTimer=setTimeout(()=>node.classList.remove('show'),3600);}
  function clock(){const now=new Date();$('relogio').textContent=now.toLocaleTimeString('pt-BR');$('data-hoje').textContent=now.toLocaleDateString('pt-BR',{weekday:'long',day:'2-digit',month:'long',year:'numeric'});}
  function setChip(kind,text){const c=$('connection-chip');c.className='connection-chip '+(kind?'is-'+kind:'');$('connection-text').textContent=text;}
  function setMeta(payload,isLive){
    state.isLive=!!isLive;
    $('footer-source').textContent=(payload.source||'Programações_2026.xlsx').split(/[\\/]/).pop();
    if(isLive&&!payload.stale){
      const warnings=Array.isArray(payload.warnings)?payload.warnings:[];
      setChip(warnings.length?'warn':'live',warnings.length?'Atenção às abas':'Excel sincronizado');
      $('source-label').textContent=warnings.length?'Planilha lida com avisos: confira as abas':'Conectado ao arquivo Excel do OneDrive';
      $('source-description').textContent=warnings.length?warnings.slice(0,3).join(' · '):payload.source+' · Nova leitura automática a cada 5 segundos.';
      $('hero-mode').textContent='Atualização automática ativa';
    }else if(isLive&&payload.stale){
      setChip('warn','Última versão válida');
      $('source-label').textContent='Aguardando o Excel ficar disponível';
      $('source-description').textContent=(payload.error||'Falha temporária na leitura.')+' · Exibindo a última base lida.';
      $('hero-mode').textContent='Repetindo leitura automaticamente';
    }else{
      setChip('warn','Prévia de referência');
      $('source-label').textContent='Prévia baseada no Excel anexado';
      $('source-description').textContent='Para atualizar diretamente do caminho do OneDrive, execute INICIAR_DASHBOARD.bat no computador onde está a planilha.';
      $('hero-mode').textContent='Inicie o servidor para sincronizar';
    }
    $('last-sync').textContent='LEITURA: '+(payload.read_at?fmtTimestamp(payload.read_at):'base de referência');
  }
  function loadData(payload,isLive=false){
    if(!payload||!Array.isArray(payload.records))return;
    const changed=payload.revision!==state.revision || state.all.length!==payload.records.length;
    if(changed){
      const wasLoaded=state.revision!==null;
      state.all=payload.records.map(r=>({
        projeto:String(r.projeto||''),data:String(r.data||''),responsavel:String(r.responsavel||''),
        status_programacao:String(r.status_programacao||''),tipo_intervencao:String(r.tipo_intervencao||''),
        status_gilmar:String(r.status_gilmar||''),regional:String(r.regional||'Não identificada'),
        aba:String(r.aba||''),linha:r.linha||'',
      }));
      state.revision=payload.revision||'snapshot';
      refreshFilterOptions();applyFilters();
      if(wasLoaded&&isLive)toast('Planilha atualizada. Indicadores e gráficos recalculados.');
    }
    setMeta(payload,isLive);
  }
  function refreshFilterOptions(){
    const selectors=[...Object.entries(fields)];
    for(const [field,domId] of selectors){
      const select=$(domId);
      const old=select.value;
      const values=unique(state.all.map(r=>r[field]));
      // Deixar as opções "Todos" e "Não preenchido" no caso de Status Gilmar.
      while(select.options.length>(field==='status_gilmar'?2:1))select.remove(select.options.length-1);
      for(const value of values){const op=document.createElement('option');op.value=value;op.textContent=value;select.add(op);}
      select.value=[...select.options].some(op=>op.value===old)?old:'';
    }
    const monthSelect=$('f-mes');const old=monthSelect.value;
    while(monthSelect.options.length>1)monthSelect.remove(1);
    const months=unique(state.all.map(r=>/^\d{4}-\d{2}/.test(r.data)?r.data.slice(0,7):''));
    for(const month of months){const op=document.createElement('option');op.value=month;op.textContent=niceMonth(month);monthSelect.add(op);}
    monthSelect.value=[...monthSelect.options].some(op=>op.value===old)?old:'';
  }
  function applyFilters(){
    const q=clean($('f-search').value),month=$('f-mes').value,windowType=$('f-janela').value;
    const chosen=Object.fromEntries(Object.entries(fields).map(([key,id])=>[key,$(id).value]));
    const today=todayISO();
    const limit=windowType==='7'||windowType==='30'?(()=>{const d=new Date();d.setDate(d.getDate()+Number(windowType));return isoLocal(d);})():null;
    state.filtered=state.all.filter(r=>{
      if(q&&!clean([r.projeto,r.data,dateBR(r.data),r.responsavel,r.status_programacao,r.tipo_intervencao,r.status_gilmar,r.regional,r.aba].join(' ')).includes(q))return false;
      for(const [field,value] of Object.entries(chosen)){
        if(!value)continue;
        if(field==='status_gilmar'&&value==='__VAZIO__'){if(r.status_gilmar.trim())return false;}
        else if(r[field]!==value)return false;
      }
      if(month&&!r.data.startsWith(month))return false;
      if(windowType==='past'&&(!dateLocal(r.data)||r.data>=today))return false;
      if(limit&&(!dateLocal(r.data)||r.data<today||r.data>=limit))return false;
      return true;
    });
    state.page=1;
    $('filter-count').textContent=`${countFmt(state.filtered.length)} de ${countFmt(state.all.length)} registros`;
    render();
  }
  function render(){renderKpis();renderTrend();renderTypeDonut();renderPeople();renderAgenda();renderRegions();renderAlerts();renderTable();}
  function renderKpis(){
    const rows=state.filtered;
    const distinct=unique(rows.map(r=>r.projeto));
    const pending=rows.filter(r=>clean(r.status_programacao).includes('pendent')).length;
    const approved=rows.filter(r=>clean(r.status_gilmar).includes('aprov')).length;
    const start=todayISO(),end=new Date();end.setDate(end.getDate()+7);const seven=isoLocal(end);
    const upcoming=rows.filter(r=>dateLocal(r.data)&&r.data>=start&&r.data<seven).length;
    const types=unique(rows.map(r=>r.tipo_intervencao));
    $('k-total').textContent=countFmt(rows.length);
    $('k-unique').textContent=countFmt(distinct.length);
    $('k-pending').textContent=countFmt(pending);
    $('k-approved').textContent=countFmt(approved);
    $('k-next').textContent=countFmt(upcoming);
    $('k-types').textContent=countFmt(types.length);
    $('k-total-sub').textContent='Registros de programação';
    $('k-unique-sub').textContent=`${rows.length-distinct.length} programação(ões) adicionais`;
    $('k-pending-sub').textContent=`${rows.length?BR.format(pending/rows.length*100):0}% da seleção`;
    $('k-approved-sub').textContent=`${rows.length?BR.format(approved/rows.length*100):0}% da seleção`;
    $('k-types-sub').textContent=types.length?types.join(' · '):'Nenhum tipo na seleção';
    $('hero-projects').textContent=countFmt(distinct.length);
    $('hero-types').textContent=countFmt(types.length);
  }
  function mondayISO(iso){const d=dateLocal(iso);if(!d)return null;const offset=(d.getDay()+6)%7;d.setDate(d.getDate()-offset);return isoLocal(d);}
  function trendGroup(rows){
    const buckets=new Map();
    for(const r of rows){if(!dateLocal(r.data))continue;const key=state.mode==='week'?mondayISO(r.data):r.data.slice(0,7);buckets.set(key,(buckets.get(key)||0)+1);}
    return [...buckets.entries()].sort((a,b)=>a[0].localeCompare(b[0])).map(([key,count])=>({key,count,label:state.mode==='week'?dateBR(key).slice(0,5):niceMonth(key)}));
  }
  function renderTrend(){
    const data=trendGroup(state.filtered),holder=$('trend-chart');
    if(!data.length){holder.innerHTML='<div class="empty-state">Nenhuma data válida para construir o gráfico.</div>';$('trend-foot').textContent='Nenhum período na seleção';return;}
    const W=730,H=215,left=40,right=15,top=24,bottom=35,chartH=H-top-bottom,chartW=W-left-right;
    const maxVal=Math.max(...data.map(d=>d.count),1);const yMax=Math.max(2,Math.ceil(maxVal/4)*4);
    const x=(i)=>data.length===1?left+chartW/2:left+(i*chartW/(data.length-1));
    const y=(num)=>top+chartH-(num/yMax)*chartH;
    const pts=data.map((d,i)=>[x(i),y(d.count)]);
    const poly=pts.map((p,i)=>(i?'L':'M')+p[0].toFixed(1)+','+p[1].toFixed(1)).join(' ');
    const fill=poly+` L${pts[pts.length-1][0].toFixed(1)},${top+chartH} L${pts[0][0].toFixed(1)},${top+chartH} Z`;
    const grid=[0,1,2,3,4].map(i=>{const val=yMax*i/4,yy=y(val).toFixed(1);return `<line x1="${left}" x2="${W-right}" y1="${yy}" y2="${yy}" stroke="#e9eff6" stroke-dasharray="4 5"/><text class="axis-label" text-anchor="end" x="${left-12}" y="${Number(yy)+4}">${Math.round(val*10)/10}</text>`;}).join('');
    const points=pts.map((p,i)=>`<circle cx="${p[0]}" cy="${p[1]}" r="5.2" fill="#fff" stroke="#178bdc" stroke-width="3"><title>${safe(data[i].label)}: ${data[i].count} programação(ões)</title></circle><text class="value-label" x="${p[0]}" y="${p[1]-13}" text-anchor="middle">${data[i].count}</text>`).join('');
    const xlabels=data.map((d,i)=>`<text class="axis-label" x="${x(i)}" y="${H-11}" text-anchor="middle">${safe(d.label)}</text>`).join('');
    holder.innerHTML=`<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Linha temporal com ${data.length} períodos e ${state.filtered.length} registros" preserveAspectRatio="none"><defs><linearGradient id="fillTrend" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#45b5ec" stop-opacity=".22"/><stop offset="100%" stop-color="#56c4ee" stop-opacity=".008"/></linearGradient></defs>${grid}<path d="${fill}" fill="url(#fillTrend)"/><path d="${poly}" stroke="#148adc" stroke-width="3" fill="none" stroke-linejoin="round" stroke-linecap="round"/>${points}${xlabels}</svg>`;
    $('trend-foot').textContent=`${data.length} período(s) · ${countFmt(state.filtered.length)} programações`;
  }
  function renderTypeDonut(){
    const totals=new Map();for(const r of state.filtered){const key=r.tipo_intervencao||'Não informado';totals.set(key,(totals.get(key)||0)+1);}
    const entries=[...totals.entries()].sort((a,b)=>b[1]-a[1]);const total=state.filtered.length;
    let cursor=0;const gradient=[];
    $('type-legend').innerHTML=entries.map(([label,count],i)=>{
      const color=TYPE_COLOR[label]||COLORS[i%COLORS.length];const start=cursor;cursor+=(count/(total||1))*100;
      gradient.push(`${color} ${start.toFixed(2)}% ${cursor.toFixed(2)}%`);
      return `<div class="donut-item"><span class="donut-item-title"><i class="donut-color" style="background:${color}"></i>${safe(label)}</span><strong>${count}</strong><small>${(count/(total||1)*100).toLocaleString('pt-BR',{maximumFractionDigits:1})}% das atividades</small></div>`;
    }).join('')||'<div class="empty-state">Nenhum tipo na seleção</div>';
    $('type-donut').style.setProperty('--donut',gradient.length?'conic-gradient('+gradient.join(',')+')':'conic-gradient(#eaf0f8 0 100%)');
    $('donut-total').textContent=countFmt(total);
  }
  function renderPeople(){
    const map=new Map();
    for(const r of state.filtered){const name=r.responsavel||'Não informado',type=r.tipo_intervencao||'Outros';if(!map.has(name))map.set(name,{});const count=map.get(name);count[type]=(count[type]||0)+1;}
    const rows=[...map].map(([person,types])=>({person,types,total:Object.values(types).reduce((a,b)=>a+b,0)})).sort((a,b)=>b.total-a.total||COLLATOR.compare(a.person,b.person));
    if(!rows.length){$('responsavel-bars').innerHTML='<div class="empty-state">Nenhum responsável na seleção.</div>';return;}
    const max=Math.max(...rows.map(r=>r.total),1);
    $('responsavel-bars').innerHTML=rows.map(row=>{
      const segments=Object.entries(row.types).sort((a,b)=>{const order=['FI','BRA','PIE'];return order.indexOf(a[0])-order.indexOf(b[0]);}).map(([type,count])=>{
        const width=(count/max*100).toFixed(2),color=TYPE_COLOR[type]||'#a3b8d0';return `<span style="background:${color};width:${width}%" title="${safe(type)}: ${count} programação(ões)"></span>`;
      }).join('');
      return `<div class="bars-row"><div class="bars-row-label" title="${safe(row.person)}">${safe(row.person)}</div><div class="bars-row-track">${segments}</div><span class="bars-row-count">${row.total}</span></div>`;
    }).join('');
  }
  function renderAgenda(){
    const today=todayISO();const upcoming=state.filtered.filter(r=>dateLocal(r.data)&&r.data>=today).sort((a,b)=>a.data.localeCompare(b.data)||COLLATOR.compare(a.projeto,b.projeto)).slice(0,5);
    $('agenda-list').innerHTML=upcoming.map(r=>{
      const date=dateLocal(r.data);return `<div class="agenda-item"><div class="agenda-date"><strong>${String(date.getDate()).padStart(2,'0')}</strong><small>${MONTHS[date.getMonth()]}</small></div><div class="agenda-info"><strong title="${safe(r.projeto)}">${safe(r.projeto||'Projeto não informado')}</strong><small>${safe(r.responsavel||'Sem responsável')} · ${safe(dateBR(r.data))}</small></div><span class="agenda-type">${safe(r.tipo_intervencao||'—')}</span></div>`;
    }).join('')||'<div class="empty-state">Não há programações futuras na seleção.</div>';
  }
  function renderRegions(){
    const total=state.filtered.length;const regions=new Map();for(const r of state.filtered){const key=r.regional||'Não identificada';regions.set(key,(regions.get(key)||0)+1);}
    const sorted=[...regions].sort((a,b)=>b[1]-a[1]);
    $('regional-bars').innerHTML=sorted.map(([name,count],i)=>`<div class="regional-item"><div class="regional-item-top"><strong>${safe(name)}</strong><span>${count} <small style="color:#9cacbd;font-weight:500">(${(count/(total||1)*100).toLocaleString('pt-BR',{maximumFractionDigits:1})}%)</small></span></div><div class="region-track"><div class="region-fill ${i%2?'other':''}" style="width:${(count/(total||1)*100).toFixed(2)}%"></div></div></div>`).join('')||'<div class="empty-state">Nenhuma regional identificada nesta seleção.</div>';
  }
  function renderAlerts(){
    const rows=state.filtered;
    const pending=rows.filter(r=>clean(r.status_programacao).includes('pendent')).length;
    const emptyGilmar=rows.filter(r=>!r.status_gilmar.trim()).length;
    const projCount=new Map();rows.forEach(r=>{if(r.projeto)projCount.set(r.projeto,(projCount.get(r.projeto)||0)+1);});
    const repeated=[...projCount.values()].filter(n=>n>1).length;
    const late=rows.filter(r=>dateLocal(r.data)&&r.data<todayISO()).length;
    const alerts=[
      {name:'Programações pendentes de aprovação',sub:'Verificar o campo Status Programação',val:pending,type:pending?'warning':'positive',icon:'◷'},
      {name:'Status Gilmar sem preenchimento',sub:'Registros que ainda não possuem status',val:emptyGilmar,type:emptyGilmar?'warning':'positive',icon:'!'},
      {name:'Projetos com mais de uma programação',sub:'Pode ser esperado: datas ou intervenções diferentes',val:repeated,type:'',icon:'⟳'},
      {name:'Datas de programação anteriores a hoje',sub:'Conferir andamento; não implica execução atrasada',val:late,type:late?'negative':'positive',icon:late?'!':'✓'},
    ];
    $('alert-list').innerHTML=alerts.map(a=>`<div class="alert-item ${a.type}"><div class="alert-icon">${a.icon}</div><div class="alert-text"><strong>${safe(a.name)}</strong><small>${safe(a.sub)}</small></div><strong class="alert-count">${a.val}</strong></div>`).join('');
  }
  function sortRecords(rows){const key=state.sortKey,dir=state.sortDir;return [...rows].sort((a,b)=>{const left=key==='aba'?`${a.aba}#${String(a.linha).padStart(6,'0')}`:a[key]||'',right=key==='aba'?`${b.aba}#${String(b.linha).padStart(6,'0')}`:b[key]||'';let cmp=COLLATOR.compare(String(left),String(right));return cmp*dir || (a.data||'').localeCompare(b.data||'') || Number(a.linha)-Number(b.linha);});}
  function classStatus(text){if(!text)return 'blank';const n=clean(text);if(n.includes('aprov')&&!n.includes('pendent'))return 'good';if(n.includes('pendent'))return '';return 'other';}
  function renderTable(){
    const all=sortRecords(state.filtered),total=all.length,pages=Math.max(1,Math.ceil(total/state.pageSize));state.page=Math.min(Math.max(state.page,1),pages);
    const first=(state.page-1)*state.pageSize,rows=all.slice(first,first+state.pageSize);
    $('table-count').textContent=`${countFmt(total)} registro(s)`;
    $('data-tbody').innerHTML=rows.map(r=>{
      const status=r.status_programacao||'Não informado';const gilmar=r.status_gilmar||'Não preenchido';
      const typeClass=clean(r.tipo_intervencao)==='bra'?'bra':clean(r.tipo_intervencao)==='pie'?'pie':'';
      return `<tr><td class="project-code">${safe(r.projeto||'—')}</td><td class="date-cell">${safe(dateBR(r.data))}</td><td class="person-cell">${safe(r.responsavel||'—')}</td><td><span class="status-tag ${classStatus(r.status_programacao)}">${safe(status)}</span></td><td><span class="type-tag ${typeClass}">${safe(r.tipo_intervencao||'—')}</span></td><td><span class="status-tag ${classStatus(r.status_gilmar)}">${safe(gilmar)}</span></td><td class="regional-tag">${safe(r.regional)}</td><td class="record-small">${safe(r.aba)} · ${safe(r.linha)}</td></tr>`;
    }).join('')||'<tr><td colspan="8" class="empty-cell">Nenhum registro corresponde aos filtros selecionados.</td></tr>';
    $('page-info').textContent=`${total?first+1:0}–${Math.min(first+state.pageSize,total)} de ${total} · Página ${state.page}/${pages}`;
    $('page-prev').disabled=state.page===1;$('page-next').disabled=state.page>=pages;
    document.querySelectorAll('thead th[data-key]').forEach(th=>{const arrow=th.querySelector('span');if(arrow)arrow.textContent=th.dataset.key===state.sortKey?(state.sortDir>0?'↑':'↓'):'↕';});
  }
  function exportCsv(){
    const cols=[['projeto','Projeto'],['data','Data Programacao'],['responsavel','Responsável'],['status_programacao','Status Programacao'],['tipo_intervencao','Tipo Intervencao'],['status_gilmar','Status Gilmar'],['regional','Regional (via código)'],['aba','Aba'],['linha','Linha Excel']];
    const quote=v=>'"'+String(v??'').replace(/"/g,'""')+'"';
    const lines=[cols.map(x=>quote(x[1])).join(';'),...sortRecords(state.filtered).map(r=>cols.map(([key])=>quote(key==='data'?dateBR(r.data):r[key])).join(';'))];
    const blob=new Blob(['\uFEFF'+lines.join('\r\n')],{type:'text/csv;charset=utf-8;'});
    const url=URL.createObjectURL(blob);const link=document.createElement('a');link.href=url;link.download=`Programacoes_2026_filtrado_${todayISO()}.csv`;document.body.appendChild(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),1500);toast(`${state.filtered.length} registros exportados para CSV.`);
  }
  function bind(){
    for(const id of ['f-search','f-responsavel','f-tipo','f-regional','f-status','f-gilmar','f-mes','f-janela']){$(id).addEventListener(id==='f-search'?'input':'change',applyFilters);}
    $('reset-filters').addEventListener('click',()=>{for(const id of ['f-search','f-responsavel','f-tipo','f-regional','f-status','f-gilmar','f-mes','f-janela'])$(id).value='';applyFilters();});
    $('btn-week').addEventListener('click',()=>switchPeriod('week'));$('btn-month').addEventListener('click',()=>switchPeriod('month'));
    $('page-size').addEventListener('change',e=>{state.pageSize=Number(e.target.value)||10;state.page=1;renderTable();});
    $('page-prev').addEventListener('click',()=>{state.page--;renderTable();});$('page-next').addEventListener('click',()=>{state.page++;renderTable();});
    document.querySelectorAll('thead th[data-key]').forEach(th=>th.addEventListener('click',()=>{if(state.sortKey===th.dataset.key)state.sortDir*=-1;else{state.sortKey=th.dataset.key;state.sortDir=1;}state.page=1;renderTable();}));
    $('export-csv').addEventListener('click',exportCsv);$('sync-button').addEventListener('click',()=>poll(true));
  }
  function switchPeriod(mode){state.mode=mode;$('btn-week').classList.toggle('is-active',mode==='week');$('btn-month').classList.toggle('is-active',mode==='month');$('btn-week').setAttribute('aria-pressed',String(mode==='week'));$('btn-month').setAttribute('aria-pressed',String(mode==='month'));renderTrend();}
  const localServer=['localhost','127.0.0.1'].includes(location.hostname)&&location.protocol.startsWith('http');
  let polling=false;
  async function poll(clicked=false){
    if(!localServer){if(clicked)toast('Para ler o Excel automaticamente, abra o dashboard pelo INICIAR_DASHBOARD.bat.');return;}
    if(polling)return;polling=true;
    try{
      const response=await fetch('/api/data',{cache:'no-store',signal:AbortSignal.timeout?AbortSignal.timeout(4800):undefined});
      if(!response.ok)throw Error(`Serviço local: ${response.status}`);
      const payload=await response.json();
      if(payload.ok){
        const updated=payload.revision!==state.revision;
        loadData(payload,true);
        if(clicked&&!updated&&!payload.stale)toast('Excel conferido. Nenhuma alteração desde a última leitura.');
        if(clicked&&updated)toast('Dados atualizados diretamente do Excel.');
      }else{
        setChip('error','Excel indisponível');
        $('source-label').textContent='Não foi possível abrir a planilha Excel';
        $('source-description').textContent=payload.error||'Verifique a localização e o nome do Excel.';
        $('hero-mode').textContent='Aguardando arquivo do OneDrive';
        if(clicked)toast('O arquivo Excel não foi encontrado no local configurado.');
      }
    }catch(error){setChip('warn','Prévia sem conexão');$('source-label').textContent='Servidor local não está respondendo';$('source-description').textContent='Reabra INICIAR_DASHBOARD.bat para habilitar a leitura automática. '+error.message;if(clicked)toast('Não foi possível conectar ao servidor local.');}
    finally{polling=false;}
  }
  function init(){bind();clock();setInterval(clock,1000);
    if(window.__DASHBOARD_SNAPSHOT__&&Array.isArray(window.__DASHBOARD_SNAPSHOT__.records))loadData(window.__DASHBOARD_SNAPSHOT__,false);
    else{state.all=[];refreshFilterOptions();applyFilters();setChip('error','Sem base inicial');}
    if(localServer){poll();setInterval(poll,5000);}
  }
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
