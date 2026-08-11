
var SENHA_HASH = 'c5d6126c1f876f8a2a1ae01486dd6d9578c837f98f8a2cf4a71a985ffa85cb9a';
var SP_SITE = 'https://gerencianet.sharepoint.com/sites/contabilidade2';
var SP_FOLDER = '/sites/contabilidade2/Shared Documents/GEST\u00c3\u0083O CONT\u00c3\u0081BIL/Cronogramas e agendas/2026/Fechamento mensal 2026';
var EMPRESAS = ['IP','SCFI','EVA','Lesta','Holding'];
var SC = {
  'No Prazo':         {bg:'#1a3a2a',fg:'#4ade80',bd:'#276221'},
  'Em atraso':        {bg:'#3a2e0a',fg:'#fbbf24',bd:'#9C6500'},
  'Atrasado':         {bg:'#3a0a0a',fg:'#f87171',bd:'#9C0006'},
  'A iniciar':        {bg:'#0a1a3a',fg:'#60a5fa',bd:'#2E75B6'},
  'Fechamento previo':{bg:'#1a2e0a',fg:'#86efac',bd:'#375623'},
  'N/A':              {bg:'#1e1e2e',fg:'#8b92b8',bd:'#595959'}
};
var ALL_STATUSES=['No Prazo','Em atraso','Atrasado','A iniciar','Fechamento previo','N/A'];
var data=[], mesRef='', dashEmp=EMPRESAS.slice(), fEmp=EMPRESAS.slice(), fSt=ALL_STATUSES.slice();

function toggleInArray(arr,val){
  var idx=arr.indexOf(val);
  if(idx>=0)arr.splice(idx,1);else arr.push(val);
  return arr;
}

async function sha256(msg){
  var buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(msg));
  return Array.from(new Uint8Array(buf)).map(function(b){return b.toString(16).padStart(2,'0');}).join('');
}

async function verificarSenha(){
  var s = document.getElementById('senhaInput').value;
  if(!s){document.getElementById('loginErr').textContent='Digite a senha.';return;}
  var h = await sha256(s);
  if(h === SENHA_HASH){
    sessionStorage.setItem('efi_auth','1');
    document.getElementById('loginScreen').style.display='none';
    iniciar();
  } else {
    document.getElementById('loginErr').textContent='Senha incorreta.';
    document.getElementById('senhaInput').value='';
    document.getElementById('senhaInput').focus();
  }
}

window.addEventListener('load', function(){
  if(sessionStorage.getItem('efi_auth')==='1'){
    document.getElementById('loginScreen').style.display='none';
    iniciar();
  } else {
    setTimeout(function(){document.getElementById('senhaInput').focus();},100);
  }
});

function iniciar(){
  document.getElementById('loadScreen').style.display='flex';
  autoLoadSharePoint();
}

async function autoLoadSharePoint(){
  try{
    setStatus('Conectando ao SharePoint...');
    var apiUrl = SP_SITE+'/_api/web/GetFolderByServerRelativeUrl(@p)/Files?@p=\''+encodeURIComponent(SP_FOLDER)+'\'&$orderby=TimeLastModified%20desc';
    var res = await fetch(apiUrl,{credentials:'include',headers:{'Accept':'application/json;odata=verbose'}});
    if(!res.ok) throw new Error('SP '+res.status);
    var json = await res.json();
    var files = json.d.results.filter(function(f){return f.Name.indexOf('Acompanhamento')===0;});
    if(!files.length) throw new Error('Arquivo nao encontrado');
    var latest = files[0];
    setStatus('Carregando '+latest.Name+'...');
    var fUrl = SP_SITE+'/_api/web/GetFileByServerRelativeUrl(@p)/$value?@p=\''+encodeURIComponent(latest.ServerRelativeUrl)+'\'';
    var fRes = await fetch(fUrl,{credentials:'include'});
    if(!fRes.ok) throw new Error('Erro ao baixar');
    var buffer = await fRes.arrayBuffer();
    processWorkbook(XLSX.read(buffer,{type:'array',cellDates:true}), latest.Name);
  } catch(err){
    console.warn('Auto-fetch:',err.message);
    document.getElementById('lsAuto').style.display='none';
    document.getElementById('lsErro').textContent='Auto-carregamento indisponivel. Use a URL abaixo ou selecione o arquivo.';
    document.getElementById('lsManual').style.display='block';
  }
}

function setStatus(m){var e=document.getElementById('lsStatus');if(e)e.textContent=m;}

async function loadFromUrl(){
  var url=document.getElementById('spUrlInput').value.trim();
  if(!url){alert('Cole a URL.');return;}
  document.getElementById('lsAuto').style.display='flex';
  document.getElementById('lsManual').style.display='none';
  setStatus('Carregando...');
  try{
    var res=await fetch(url,{credentials:'include'});
    if(!res.ok) throw new Error('Erro '+res.status);
    var buf=await res.arrayBuffer();
    var name=decodeURIComponent(url.split('/').pop().split('?')[0]);
    processWorkbook(XLSX.read(buf,{type:'array',cellDates:true}),name);
  } catch(err){
    document.getElementById('lsAuto').style.display='none';
    document.getElementById('lsManual').style.display='block';
    document.getElementById('lsErro').textContent='Nao foi possivel carregar: '+err.message;
  }
}

function loadExcel(inp){
  var f=inp.files[0];if(!f)return;
  var r=new FileReader();
  r.onload=function(e){
    try{processWorkbook(XLSX.read(e.target.result,{type:'binary',cellDates:true}),f.name);}
    catch(err){alert('Erro: '+err.message);}
  };
  r.readAsBinaryString(f);
}

function xlDate(v){
  if(!v)return'';
  if(typeof v==='string')return v;
  if(v instanceof Date){return String(v.getDate()).padStart(2,'0')+'/'+String(v.getMonth()+1).padStart(2,'0')+'/'+v.getFullYear();}
  if(typeof v==='number'){var d=new Date(Math.round((v-25569)*86400*1000));return String(d.getUTCDate()).padStart(2,'0')+'/'+String(d.getUTCMonth()+1).padStart(2,'0')+'/'+d.getUTCFullYear();}
  return String(v);
}

function parseBR(d){
  if(!d)return null;
  var p=String(d).split('/');
  if(p.length!==3)return null;
  var dt=new Date(p[2],p[1]-1,p[0]);
  return isNaN(dt.getTime())?null:dt;
}

function calcStatus(ent,prazo){
  if(!prazo)return'A iniciar';
  if(ent==='N/A')return'N/A';
  if(ent==='Previo')return'Fechamento previo';
  if(!ent){var p=prazo.split('/');return new Date(p[2],p[1]-1,p[0])<new Date()?'Em atraso':'A iniciar';}
  var e=ent.split('/'),p=prazo.split('/');
  return new Date(e[2],e[1]-1,e[0])>new Date(p[2],p[1]-1,p[0])?'Atrasado':'No Prazo';
}

function normalizeHeader(s){
  return String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,' ').trim().toLowerCase();
}

function acharColunas(rows){
  var cols={obs:19,resp:5,impacta:2,prazoFch:17,entregaFch:18,situacaoFch:19};
  for(var hi=0;hi<3&&hi<rows.length;hi++){
    var header=rows[hi]||[];
    var prazoM=[],entregaM=[],situM=[];
    for(var c=0;c<header.length;c++){
      var h=normalizeHeader(header[c]);
      if(h==='observacoes')cols.obs=c;
      if(h==='responsavel')cols.resp=c;
      if(h.indexOf('impacta resultado')===0)cols.impacta=c;
      if(h==='data do prazo')prazoM.push(c);
      if(h==='data da entrega')entregaM.push(c);
      if(h==='situacao da entrega')situM.push(c);
    }
    if(prazoM.length)cols.prazoFch=prazoM[prazoM.length-1];
    if(entregaM.length)cols.entregaFch=entregaM[entregaM.length-1];
    if(situM.length)cols.situacaoFch=situM[situM.length-1];
  }
  return cols;
}

function escHtml(s){
  return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function getSit(raw){
  if(!raw)return'';
  var s=String(raw).trim();
  if(s==='Fechamento previo'||s==='Fechamento pr\u00e9vio')return'Fechamento previo';
  return s;
}

function processWorkbook(wb,fileName){
  data=[];var id=0;mesRef='';
  EMPRESAS.forEach(function(emp){
    if(wb.SheetNames.indexOf(emp)<0)return;
    var ws=wb.Sheets[emp];
    var rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
    if(!mesRef&&rows[0]&&rows[0][3])mesRef=xlDate(rows[0][3]);
    var cols=acharColunas(rows);
    for(var i=2;i<rows.length;i++){
      var r=rows[i];
      var dem=r[0];
      if(!dem||String(dem).trim()===''||String(dem).trim()==='Demandas')continue;
      var dpf=xlDate(r[cols.prazoFch]);
      var def=r[cols.entregaFch]?xlDate(r[cols.entregaFch]):'';
      var sf=r[cols.situacaoFch]?getSit(r[cols.situacaoFch]):calcStatus(def,dpf);
      var obsRaw=r[cols.obs];
      var obs=(obsRaw===undefined||obsRaw===null)?'':String(obsRaw).trim();
      var impactaRaw=r[cols.impacta];
      var impacta=String(impactaRaw||'').trim();
      if(impacta!=='Sim'&&impacta!=='Não')impacta='Não informado';
      data.push({id:emp+'-'+id++,empresa:emp,
        demanda:String(dem).trim(),descricao:String(r[1]||'').trim(),
        responsavel:String(r[cols.resp]||'').trim(),
        impacta:impacta,
        dataPrazoFch:dpf,dataEntregaFch:def,situacaoFch:sf,
        observacoes:obs});
    }
  });
  document.getElementById('mesBadge').textContent=mesRef?'Fechamento '+mesRef:fileName.replace('.xlsx','');
  document.getElementById('loadScreen').style.display='none';
  document.getElementById('app').style.display='block';
  dashEmp=EMPRESAS.slice();fEmp=EMPRESAS.slice();fSt=ALL_STATUSES.slice();
  renderAll();
}

function badge(s){
  var k=s==='Fechamento previo'?'Fechamento previo':s;
  var c=SC[k]||{bg:'#1e1e2e',fg:'#8b92b8'};
  return '<span class="badge" style="background:'+c.bg+';color:'+c.fg+'">'+s+'</span>';
}

function dashData(){return data.filter(function(d){return dashEmp.indexOf(d.empresa)>=0;});}

function renderAll(){renderDashEmpFilter();kpis();empresaBars();collabTable();renderEmpresaFilter();renderStatusFilter();renderOp();setTimeout(function(){statusChart();perfCard();impactoCard();},80);}

function renderDashEmpFilter(){
  var todasBtn='<button class="ef-btn'+(dashEmp.length===EMPRESAS.length?' active':'')+'" onclick="setDE(\'Todas\')">Todas</button>';
  var emps=EMPRESAS.map(function(e){
    return '<button class="ef-btn'+(dashEmp.indexOf(e)>=0?' active':'')+'" onclick="setDE(\''+e+'\')">'+e+'</button>';
  }).join('');
  document.getElementById('dashEmpFilter').innerHTML=todasBtn+emps;
}
function setDE(e){
  if(e==='Todas'){dashEmp=EMPRESAS.slice();}else{toggleInArray(dashEmp,e);}
  renderDashEmpFilter();kpis();empresaBars();collabTable();setTimeout(function(){statusChart();perfCard();impactoCard();},80);
}

function kpis(){
  var dd=dashData();
  var tot=dd.length;
  var prazo=dd.filter(function(d){return d.situacaoFch==='No Prazo';}).length;
  var atras=dd.filter(function(d){return d.situacaoFch==='Atrasado';}).length;
  var emA=dd.filter(function(d){return d.situacaoFch==='Em atraso';}).length;
  var ainit=dd.filter(function(d){return d.situacaoFch==='A iniciar';}).length;
  var prev=dd.filter(function(d){return d.situacaoFch==='Fechamento previo';}).length;
  var na=dd.filter(function(d){return d.situacaoFch==='N/A';}).length;
  var pct=tot?Math.round(prazo/tot*100):0;
  var cards=[
    {l:'Total de demandas',v:tot,  s:dashEmp.length===EMPRESAS.length?'5 empresas':(dashEmp.length?dashEmp.join(', '):'Nenhuma empresa'),c:'#3d7dd4',vc:'#7fb3f5'},
    {l:'Concluido no prazo',v:prazo,s:pct+'% do total',c:'#276221',vc:'#4ade80'},
    {l:'A iniciar',v:ainit,s:'Prazo nao venceu',c:'#2E75B6',vc:'#60a5fa'},
    {l:'Em atraso',v:emA,s:'Prazo vencido s/ entrega',c:'#9C6500',vc:'#fbbf24'},
    {l:'Concl. c/ atraso',v:atras,s:'Entregue apos prazo',c:'#9C0006',vc:'#f87171'},
    {l:'Fechamento previo',v:prev,s:'Acao paliativa',c:'#375623',vc:'#86efac'},
    {l:'N/A',v:na,s:'Nao se aplica',c:'#595959',vc:'#8b92b8'}
  ];
  document.getElementById('kpiGrid').innerHTML=cards.map(function(k){
    return '<div class="kpi" style="border-left-color:'+k.c+'"><div class="kpi-label">'+k.l+'</div><div class="kpi-val" style="color:'+k.vc+'">'+k.v+'</div><div class="kpi-sub">'+k.s+'</div></div>';
  }).join('');
}

function empresaBars(){
  var emps=EMPRESAS.filter(function(e){return dashEmp.indexOf(e)>=0;});
  document.getElementById('empresaBars').innerHTML=emps.map(function(emp){
    var rows=data.filter(function(d){return d.empresa===emp;});
    if(!rows.length)return'';
    var concl=rows.filter(function(d){return d.situacaoFch==='No Prazo'||d.situacaoFch==='Atrasado';}).length;
    var pct=Math.round(concl/rows.length*100);
    var hasE=rows.some(function(d){return d.situacaoFch==='Em atraso';});
    var dc=hasE?'#fbbf24':'#4ade80';
    return '<div class="empresa-row"><div class="empresa-dot" style="background:'+dc+'"></div><div class="empresa-name">'+emp+'</div><div class="bar-bg"><div class="bar-fill" style="width:'+pct+'%;background:'+dc+'"></div></div><div class="empresa-pct" style="color:'+dc+'">'+pct+'%</div><div class="empresa-cnt">'+concl+'/'+rows.length+' concl.</div></div>';
  }).join('');
}

var chartInst=null;
function statusChart(){
  var ctx=document.getElementById('statusChart').getContext('2d');
  var dd=dashData();
  var labels=['No Prazo','Em atraso','Atrasado','A iniciar','Fch. previo','N/A'];
  var keys=['No Prazo','Em atraso','Atrasado','A iniciar','Fechamento previo','N/A'];
  var vals=keys.map(function(s){return dd.filter(function(d){return d.situacaoFch===s;}).length;});
  var bgs=['#276221','#9C6500','#9C0006','#2E75B6','#375623','#595959'];
  if(chartInst)chartInst.destroy();
  chartInst=new Chart(ctx,{type:'bar',
    data:{labels:labels,datasets:[{data:vals,backgroundColor:bgs,borderRadius:4,borderSkipped:false}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false}},
      scales:{
        x:{ticks:{color:'#8b92b8',font:{size:10}},grid:{color:'rgba(255,255,255,.04)'}},
        y:{ticks:{color:'#8b92b8',font:{size:10}},grid:{color:'rgba(255,255,255,.06)'},beginAtZero:true}
      }}
  });
}

var perfChartInst=null;
function perfCard(){
  var dd=dashData();
  var tot=dd.length;
  var noPrazo=dd.filter(function(d){return d.situacaoFch==='No Prazo';}).length;
  var atrasado=dd.filter(function(d){return d.situacaoFch==='Atrasado';}).length;
  var andamento=tot-noPrazo-atrasado; // A iniciar + Em atraso + Fechamento previo + N/A
  var pPrazo=tot?Math.round(noPrazo/tot*100):0;
  var pAtraso=tot?Math.round(atrasado/tot*100):0;

  document.getElementById('perfStats').innerHTML=
    '<div class="perf-row"><span class="perf-label">Total de demandas</span><span class="perf-val">'+tot+'</span></div>'+
    '<div class="perf-row"><span class="perf-label" style="color:#4ade80">Concluidas no prazo</span><span class="perf-val" style="color:#4ade80">'+noPrazo+' ('+pPrazo+'%)</span></div>'+
    '<div class="perf-row"><span class="perf-label" style="color:#f87171">Entregues fora do prazo</span><span class="perf-val" style="color:#f87171">'+atrasado+' ('+pAtraso+'%)</span></div>';

  var canvas=document.getElementById('perfChart');
  if(!canvas)return;
  var ctx=canvas.getContext('2d');
  if(perfChartInst)perfChartInst.destroy();
  if(!tot){
    canvas.style.display='none';
    return;
  }
  canvas.style.display='block';
  perfChartInst=new Chart(ctx,{type:'pie',
    data:{labels:['No Prazo','Atrasado','Em andamento'],
      datasets:[{data:[noPrazo,atrasado,andamento],backgroundColor:['#276221','#9C0006','#3d7dd4'],borderColor:'#1a1d2e',borderWidth:2}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{
        legend:{position:'bottom',labels:{color:'#8b92b8',font:{size:9},boxWidth:8,padding:8}},
        tooltip:{callbacks:{label:function(c){
          var v=c.parsed,pc=tot?Math.round(v/tot*100):0;
          return c.label+': '+v+' ('+pc+'%)';
        }}}
      }}
  });
}

function collabTable(){
  var counts={};
  dashData().forEach(function(d){
    var names=d.responsavel.split(/\s+e\s+/i).map(function(n){return n.trim();}).filter(Boolean);
    if(!names.length)names=['(Sem responsavel)'];
    names.forEach(function(n){
      if(!counts[n])counts[n]={noPrazo:0,atrasado:0,emAtraso:0,total:0};
      counts[n].total++;
      if(d.situacaoFch==='No Prazo')counts[n].noPrazo++;
      else if(d.situacaoFch==='Atrasado')counts[n].atrasado++;
      else if(d.situacaoFch==='Em atraso')counts[n].emAtraso++;
    });
  });
  var names=Object.keys(counts).sort(function(a,b){
    var pa=counts[a].emAtraso+counts[a].atrasado, pb=counts[b].emAtraso+counts[b].atrasado;
    if(pb!==pa)return pb-pa;
    return counts[b].total-counts[a].total;
  });
  if(!names.length){
    document.getElementById('collabTable').innerHTML='<tbody><tr><td colspan="4" style="text-align:center;padding:20px;color:#8b92b8">Sem dados</td></tr></tbody>';
    return;
  }
  var rows=names.map(function(n){
    var c=counts[n];
    return '<tr><td>'+escHtml(n)+'</td>'+
      '<td style="color:#4ade80;font-weight:600">'+c.noPrazo+'</td>'+
      '<td style="color:#f87171;font-weight:600">'+c.atrasado+'</td>'+
      '<td style="color:#fbbf24;font-weight:600">'+c.emAtraso+'</td></tr>';
  }).join('');
  document.getElementById('collabTable').innerHTML='<thead><tr><th>Colaborador</th><th>No Prazo</th><th>Atrasado</th><th>Em atraso</th></tr></thead><tbody>'+rows+'</tbody>';
}

var impactoChartInst=null;
function impactoCard(){
  var dd=dashData();
  var sim=dd.filter(function(d){return d.impacta==='Sim';});
  var nao=dd.filter(function(d){return d.impacta==='Não';});
  var naoInf=dd.filter(function(d){return d.impacta==='Não informado';});
  var simTot=sim.length;
  var byStatus={};
  ALL_STATUSES.forEach(function(s){byStatus[s]=sim.filter(function(d){return d.situacaoFch===s;}).length;});
  var pNoPrazo=simTot?Math.round(byStatus['No Prazo']/simTot*100):0;
  var pendentes=byStatus['Em atraso']+byStatus['A iniciar']+byStatus['Fechamento previo']+byStatus['N/A'];

  document.getElementById('impactoStats').innerHTML=
    '<div class="perf-row"><span class="perf-label">Impactam resultado (Sim)</span><span class="perf-val">'+simTot+'</span></div>'+
    '<div class="perf-row"><span class="perf-label" style="color:#4ade80">Concluidas no prazo</span><span class="perf-val" style="color:#4ade80">'+byStatus['No Prazo']+' ('+pNoPrazo+'%)</span></div>'+
    '<div class="perf-row"><span class="perf-label" style="color:#f87171">Entregues com atraso</span><span class="perf-val" style="color:#f87171">'+byStatus['Atrasado']+'</span></div>'+
    '<div class="perf-row"><span class="perf-label" style="color:#fbbf24">Pendentes</span><span class="perf-val" style="color:#fbbf24">'+pendentes+'</span></div>'+
    '<div class="perf-row"><span class="perf-label">Nao impactam (Nao)</span><span class="perf-val">'+nao.length+'</span></div>'+
    '<div class="perf-row"><span class="perf-label">Nao informado</span><span class="perf-val">'+naoInf.length+'</span></div>';

  var canvas=document.getElementById('impactoChart');
  if(!canvas)return;
  var ctx=canvas.getContext('2d');
  if(impactoChartInst)impactoChartInst.destroy();
  if(!simTot){
    canvas.style.display='none';
    return;
  }
  canvas.style.display='block';
  var labels=ALL_STATUSES;
  var vals=labels.map(function(s){return byStatus[s];});
  var bgs=labels.map(function(s){return (SC[s]&&SC[s].fg)||'#8b92b8';});
  impactoChartInst=new Chart(ctx,{type:'bar',
    data:{labels:labels,datasets:[{data:vals,backgroundColor:bgs,borderRadius:4,borderSkipped:false}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false}},
      scales:{
        x:{ticks:{color:'#8b92b8',font:{size:9}},grid:{color:'rgba(255,255,255,.04)'}},
        y:{ticks:{color:'#8b92b8',font:{size:9}},grid:{color:'rgba(255,255,255,.06)'},beginAtZero:true}
      }}
  });
}

function renderEmpresaFilter(){
  var todasBtn='<button class="ef-btn'+(fEmp.length===EMPRESAS.length?' active':'')+'" onclick="setE(\'Todas\')">Todas</button>';
  var emps=EMPRESAS.map(function(e){
    return '<button class="ef-btn'+(fEmp.indexOf(e)>=0?' active':'')+'" onclick="setE(\''+e+'\')">'+e+'</button>';
  }).join('');
  document.getElementById('empresaFilter').innerHTML=todasBtn+emps;
}
function setE(e){
  if(e==='Todas'){fEmp=EMPRESAS.slice();}else{toggleInArray(fEmp,e);}
  renderEmpresaFilter();renderOp();
}

function renderStatusFilter(){
  var todosAtivo=fSt.length===ALL_STATUSES.length;
  var cTodos=SC['Todos']||{bg:'#242740',fg:'#8b92b8',bd:'#8b92b8'};
  var stTodos=todosAtivo?'background:'+cTodos.bg+';color:'+cTodos.fg+';border-color:'+(cTodos.bd||cTodos.fg)+';font-weight:700':'';
  var todosBtn='<button class="sf-btn" style="'+stTodos+'" onclick="setSt(\'Todos\')">Todos</button>';
  var chips=ALL_STATUSES.map(function(s){
    var c=SC[s]||{bg:'#242740',fg:'#8b92b8',bd:'#8b92b8'};
    var a=fSt.indexOf(s)>=0;
    var st=a?'background:'+c.bg+';color:'+c.fg+';border-color:'+(c.bd||c.fg)+';font-weight:700':'';
    return '<button class="sf-btn" style="'+st+'" onclick="setSt(\''+s+'\')">'+s+'</button>';
  }).join('');
  document.getElementById('statusFilter').innerHTML=todosBtn+chips;
}
function setSt(s){
  if(s==='Todos'){fSt=ALL_STATUSES.slice();}else{toggleInArray(fSt,s);}
  renderStatusFilter();renderOp();
}

function renderOp(){
  var q=(document.getElementById('searchOp')||{value:''}).value.toLowerCase();
  var rows=data.filter(function(d){
    return(fEmp.indexOf(d.empresa)>=0)&&(fSt.indexOf(d.situacaoFch)>=0)&&(!q||d.demanda.toLowerCase().indexOf(q)>=0||d.responsavel.toLowerCase().indexOf(q)>=0);
  });
  var trs=rows.map(function(r){
    var entCell=r.dataEntregaFch||'-';
    var obsCell='<span class="obs-cell">'+(r.observacoes?escHtml(r.observacoes):'')+'</span>';
    return '<tr><td><span class="emp-tag">'+r.empresa+'</span></td><td style="max-width:130px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">'+r.demanda+'</td><td style="max-width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#8b92b8">'+r.descricao+'</td><td>'+r.responsavel+'</td><td>'+r.dataPrazoFch+'</td><td>'+entCell+'</td><td>'+badge(r.situacaoFch)+'</td><td>'+obsCell+'</td></tr>';
  }).join('');
  document.getElementById('opTable').innerHTML='<thead><tr><th>Emp.</th><th>Demanda</th><th>Descricao</th><th>Responsavel</th><th>Prazo FCH</th><th>Data entrega</th><th>Status</th><th>Observações</th></tr></thead><tbody>'+trs+'</tbody>';
  document.getElementById('opCount').textContent=rows.length+' demanda'+(rows.length!==1?'s':'')+' exibida'+(rows.length!==1?'s':'');
}

function switchTab(tab,el){
  document.querySelectorAll('.tab').forEach(function(t){t.classList.remove('active');});
  el.classList.add('active');
  ['dashView','opView'].forEach(function(id){document.getElementById(id).style.display='none';});
  document.getElementById(tab==='dash'?'dashView':'opView').style.display='block';
}
