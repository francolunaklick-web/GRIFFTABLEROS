<?php require_once __DIR__ . "/../auth_check.php"; ?>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Control de Aportes — Griff Salud</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
:root{
  --green:#1E8449;--green-lt:#e8f5e9;--blue:#1565C0;--blue-lt:#e3f2fd;
  --yellow:#F57F17;--yellow-lt:#fff8e1;--red:#C62828;--red-lt:#ffebee;
  --border:#e0e0e0;--muted:#757575;--bg:#f5f6f8;--card:#fff;--text:#1a1a2e;
  --radius:10px;--shadow:0 2px 8px rgba(0,0,0,.08);
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',sans-serif;background:var(--bg);color:var(--text);font-size:14px}
.header{background:#fff;border-bottom:2px solid var(--green);padding:14px 24px;display:flex;align-items:center;gap:16px}
.header-logo{width:36px;height:36px;background:var(--green);border-radius:8px;display:flex;align-items:center;justify-content:center}
.header-title{font-size:18px;font-weight:700;color:var(--green)}
.header-sub{font-size:12px;color:var(--muted)}
.header-back{margin-left:auto;font-size:12px;color:var(--blue);text-decoration:none}
.badge-update{margin-left:12px;font-size:11px;color:var(--muted);background:#f0f0f0;padding:3px 8px;border-radius:20px}
.main{max-width:1440px;margin:0 auto;padding:20px 24px}
/* kpis */
.kpi-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(155px,1fr));gap:12px;margin-bottom:18px}
.kpi{background:var(--card);border-radius:var(--radius);box-shadow:var(--shadow);padding:14px 16px;border-top:3px solid var(--green)}
.kpi.blue{border-top-color:var(--blue)}.kpi.yellow{border-top-color:var(--yellow)}.kpi.red{border-top-color:var(--red)}
.kpi-label{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px}
.kpi-value{font-size:21px;font-weight:800}
.kpi-sub{font-size:11px;color:var(--muted);margin-top:2px}
/* filtros */
.filtros{background:var(--card);border-radius:var(--radius);box-shadow:var(--shadow);padding:16px 20px;margin-bottom:16px}
.f-title{font-size:11px;font-weight:700;text-transform:uppercase;color:var(--muted);letter-spacing:.5px;margin-bottom:10px}
.fila{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:10px;align-items:center}
.fila:last-child{margin-bottom:0}
.fila-lbl{font-size:12px;font-weight:600;color:var(--muted);min-width:95px}
.btn{padding:5px 13px;border:1px solid var(--border);border-radius:20px;background:#fff;cursor:pointer;font-size:12px;transition:all .15s}
.btn:hover{border-color:var(--green);color:var(--green)}
.btn.active{background:var(--green);color:#fff;border-color:var(--green)}
.btn.r-menos25.active{background:var(--red);border-color:var(--red);color:#fff}
.btn.r-25y50.active{background:var(--yellow);border-color:var(--yellow);color:#fff}
.btn.preset{border-color:var(--blue);color:var(--blue)}
.btn.preset:hover,.btn.preset.active{background:var(--blue);color:#fff;border-color:var(--blue)}
.btn.dl{border-color:var(--green);color:var(--green);font-weight:700}
.btn.dl:hover{background:var(--green);color:#fff}
select.sel{padding:5px 10px;border:1px solid var(--border);border-radius:6px;font-size:12px;cursor:pointer}
input.srch{padding:6px 12px;border:1px solid var(--border);border-radius:6px;font-size:13px;width:240px;outline:none}
input.srch:focus{border-color:var(--green)}
/* panel/tabs */
.panel{background:var(--card);border-radius:var(--radius);box-shadow:var(--shadow);padding:20px;margin-bottom:16px}
.panel.notab{padding-bottom:8px}
.tabs{display:flex;gap:0;border-bottom:2px solid var(--border);margin:-20px -20px 16px}
.tab-btn{padding:11px 20px;background:none;border:none;border-bottom:3px solid transparent;cursor:pointer;font-size:13px;font-weight:600;color:var(--muted);margin-bottom:-2px;transition:all .15s}
.tab-btn.active{color:var(--green);border-bottom-color:var(--green)}
.tab-pane{display:none}.tab-pane.active{display:block}
.p-title{font-size:14px;font-weight:700;margin-bottom:12px;display:flex;align-items:center;gap:8px}
.cnt{font-size:12px;font-weight:600;background:var(--green-lt);color:var(--green);padding:2px 8px;border-radius:12px}
/* tabla */
table{width:100%;border-collapse:collapse;font-size:13px}
th{text-align:left;padding:8px 10px;border-bottom:2px solid var(--border);font-size:11px;text-transform:uppercase;letter-spacing:.4px;color:var(--muted);white-space:nowrap;cursor:pointer;user-select:none}
th:hover{color:var(--green)}
th.asc::after{content:" ↑"}th.desc::after{content:" ↓"}
td{padding:7px 10px;border-bottom:1px solid #f0f0f0;vertical-align:middle}
tr:hover td{background:#fafafa}
.r,.right{text-align:right}
.bold{font-weight:700}
.tr-total td{font-weight:700;background:#f8f8f8;border-top:2px solid var(--border)}
/* badges */
.badge{display:inline-block;padding:2px 7px;border-radius:10px;font-size:10px;font-weight:700;white-space:nowrap}
.b-m25{background:var(--red-lt);color:var(--red)}
.b-25y50{background:var(--yellow-lt);color:var(--yellow)}
.b-mas50{background:var(--green-lt);color:var(--green)}
.b-miem{background:#e3f2fd;color:#1565C0}
/* paginación */
.pag{display:flex;gap:5px;flex-wrap:wrap;margin-top:10px;align-items:center}
.pb{padding:4px 9px;border:1px solid var(--border);border-radius:4px;background:#fff;cursor:pointer;font-size:12px}
.pb.active{background:var(--green);color:#fff;border-color:var(--green)}
.pb:hover:not(.active){border-color:var(--green)}
.pag-info{font-size:12px;color:var(--muted);margin-left:6px}
.loading{text-align:center;padding:50px;color:var(--muted);font-size:15px}
.nota{font-size:11px;color:var(--muted);margin-top:10px;padding:8px 12px;background:#f9f9f9;border-radius:6px;border-left:3px solid var(--border)}
</style>
</head>
<body>

<div class="header">
  <div class="header-logo">
    <svg viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2.2"><circle cx="12" cy="12" r="4"/><path d="M12 2v3M12 19v3M4.22 4.22l2.12 2.12M17.66 17.66l2.12 2.12M2 12h3M19 12h3M4.22 19.78l2.12-2.12M17.66 6.34l2.12-2.12"/></svg>
  </div>
  <div>
    <div class="header-title">Control de Aportes</div>
    <div class="header-sub">Aporte per cápita por grupo familiar · Filtros por delegación y período</div>
  </div>
  <span id="upd" class="badge-update">Cargando...</span>
  <a class="header-back" href="../INICIO.php">← Inicio</a>
</div>

<div class="main">
  <div id="loading" class="loading">⏳ Cargando datos...</div>
  <div id="app" style="display:none">

    <!-- KPIs -->
    <div class="kpi-row">
      <div class="kpi blue"><div class="kpi-label">Grupos familiares</div><div class="kpi-value" id="k-grupos">—</div><div class="kpi-sub">Únicos en período</div></div>
      <div class="kpi blue"><div class="kpi-label">Total miembros</div><div class="kpi-value" id="k-miem">—</div><div class="kpi-sub">Personas cubiertas</div></div>
      <div class="kpi red"><div class="kpi-label">Grupos &lt; $25k pc</div><div class="kpi-value" id="k-m25">—</div><div class="kpi-sub">Per cápita bajo</div></div>
      <div class="kpi yellow"><div class="kpi-label">Grupos $25k–$50k pc</div><div class="kpi-value" id="k-r25">—</div><div class="kpi-sub">Per cápita medio</div></div>
      <div class="kpi"><div class="kpi-label">Total aportado</div><div class="kpi-value" id="k-tot">—</div><div class="kpi-sub">Aportes + contribuciones</div></div>
      <div class="kpi"><div class="kpi-label">Per cápita promedio</div><div class="kpi-value" id="k-pc">—</div><div class="kpi-sub">Total ÷ miembros</div></div>
    </div>

    <!-- Filtros -->
    <div class="filtros">
      <div class="f-title">Filtros — el rango se calcula sobre: Aporte total del grupo ÷ Cantidad de miembros</div>
      <div class="fila">
        <span class="fila-lbl">Período</span>
        <div id="mes-btns"></div>
      </div>
      <div class="fila">
        <span class="fila-lbl">Delegación</span>
        <select id="sel-deleg" class="sel" onchange="render()"><option value="">Todas</option></select>
      </div>
      <div class="fila" style="align-items:center">
        <span class="fila-lbl">Per cápita</span>
        <span style="font-size:12px;color:var(--muted)">Desde $</span>
        <input type="number" id="pc-min" placeholder="0" min="0" step="1000"
          style="width:110px;padding:5px 8px;border:1px solid var(--border);border-radius:6px;font-size:13px;outline:none"
          oninput="limpiarAtajos();render()">
        <span style="font-size:12px;color:var(--muted)">Hasta $</span>
        <input type="number" id="pc-max" placeholder="Sin límite" min="0" step="1000"
          style="width:120px;padding:5px 8px;border:1px solid var(--border);border-radius:6px;font-size:13px;outline:none"
          oninput="limpiarAtajos();render()">
        <span style="font-size:11px;color:var(--muted);margin-left:4px">— atajos:</span>
        <button class="btn r-menos25" id="b-m25"   onclick="setAtajo(0,25000,this)">– $25k</button>
        <button class="btn r-25y50"   id="b-r25"   onclick="setAtajo(25000,50000,this)">$25k – $50k</button>
        <button class="btn"           id="b-mas50" onclick="setAtajo(50000,'',this)">+ $50k</button>
      </div>
      <div class="fila">
        <span class="fila-lbl">Tamaño grupo</span>
        <button class="btn active" id="b-tam-todos" onclick="setTam(0,this)">Todos</button>
        <button class="btn" onclick="setTam(1,this)">Solo titular</button>
        <button class="btn" onclick="setTam(2,this)">2 miembros</button>
        <button class="btn" onclick="setTam(3,this)">3+ miembros</button>
      </div>
      <div class="fila">
        <span class="fila-lbl">Presets</span>
        <button class="btn preset" onclick="presetBahia()">Bahía Blanca $25k–$50k</button>
        <button class="btn" onclick="resetFiltros()">↺ Limpiar</button>
        <button class="btn dl" onclick="descargarExcel()" style="margin-left:auto">⬇ Descargar Excel</button>
      </div>
    </div>

    <!-- Tabs -->
    <div class="panel notab">
      <div class="tabs">
        <button class="tab-btn active" onclick="setTab('resumen',this)">Por Delegación</button>
        <button class="tab-btn"        onclick="setTab('grupos',this)">Grupos Familiares</button>
        <button class="tab-btn"        onclick="setTab('mensual',this)">Por Mes</button>
      </div>

      <!-- Resumen por delegación -->
      <div id="tab-resumen" class="tab-pane active">
        <div class="p-title">Resumen por Delegación <span class="cnt" id="cnt-res">—</span></div>
        <table>
          <thead><tr>
            <th onclick="srt('res','deleg')">Delegación</th>
            <th class="r" onclick="srt('res','grupos')">Grupos</th>
            <th class="r" onclick="srt('res','miem')">Miembros</th>
            <th class="r" onclick="srt('res','m25')">Menos $25k</th>
            <th class="r" onclick="srt('res','r25')">$25k–$50k</th>
            <th class="r" onclick="srt('res','mas50')">Más $50k</th>
            <th class="r" onclick="srt('res','tot')">Total aportado</th>
            <th class="r" onclick="srt('res','pc')">Per cápita prom.</th>
          </tr></thead>
          <tbody id="tb-res"></tbody>
        </table>
        <div class="nota">* Grupos únicos en el período seleccionado. Per cápita = total del grupo ÷ miembros del grupo familiar.</div>
      </div>

      <!-- Grupos familiares -->
      <div id="tab-grupos" class="tab-pane">
        <div style="display:flex;gap:12px;align-items:center;flex-wrap:wrap;margin-bottom:12px">
          <div class="p-title" style="margin:0">Grupos <span class="cnt" id="cnt-grp">—</span></div>
          <input class="srch" id="srch" placeholder="Buscar por titular, delegación..." oninput="renderGrupos()">
          <div class="pag-info" id="grp-pag-info"></div>
        </div>
        <table>
          <thead><tr>
            <th onclick="srt('grp','titular')">Titular</th>
            <th onclick="srt('grp','deleg')">Delegación</th>
            <th onclick="srt('grp','mes')">Mes</th>
            <th class="r" onclick="srt('grp','miembros')">Miembros</th>
            <th onclick="srt('grp','paren')">A cargo</th>
            <th class="r" onclick="srt('grp','ap')">Aportes</th>
            <th class="r" onclick="srt('grp','co')">Contribuc.</th>
            <th class="r" onclick="srt('grp','tot')">Total grupo</th>
            <th class="r" onclick="srt('grp','pc')">Per cápita</th>
            <th>Rango</th>
          </tr></thead>
          <tbody id="tb-grp"></tbody>
        </table>
        <div class="pag" id="pag-grp"></div>
      </div>

      <!-- Por mes -->
      <div id="tab-mensual" class="tab-pane">
        <div class="p-title">Evolución Mensual</div>
        <table>
          <thead><tr>
            <th>Mes</th>
            <th class="r">Grupos</th>
            <th class="r">Miembros</th>
            <th class="r">Menos $25k</th>
            <th class="r">$25k–$50k</th>
            <th class="r">Más $50k</th>
            <th class="r">Total aportado</th>
            <th class="r">Per cápita prom.</th>
          </tr></thead>
          <tbody id="tb-mes"></tbody>
        </table>
      </div>
    </div>

  </div><!-- #app -->
</div>

<script>
(function(){
  const s = document.createElement('script');
  s.src = 'datos.js?v=' + Date.now();
  s.onload = () => init();
  s.onerror = () => { document.getElementById('loading').textContent = '❌ Error — corré actualizar.py primero.'; };
  document.head.appendChild(s);
})();

// ── estado ────────────────────────────────────────────────────────────────────
let F = { meses: [], deleg: '', pcMin: null, pcMax: null, tam: 0 };
let SS = { res:{c:'deleg',a:true}, grp:{c:'pc',a:true} };
let grpPage = 0;
const PS = 60;

// ── init ──────────────────────────────────────────────────────────────────────
function init() {
  document.getElementById('loading').style.display = 'none';
  document.getElementById('app').style.display = 'block';
  document.getElementById('upd').textContent = 'Actualizado: ' + APORTES_META.actualizado;

  // Mes buttons
  const mb = document.getElementById('mes-btns');
  mb.innerHTML = '<button class="btn active" onclick="setMes(\'\',this)">Todos</button>' +
    APORTES_META.meses.map(m => {
      const [mm,yy] = m.split('-');
      const L = ['','Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'][+mm]+' '+yy.slice(2);
      return `<button class="btn" onclick="setMes('${m}',this)">${L}</button>`;
    }).join('');

  // Delegación
  const sel = document.getElementById('sel-deleg');
  APORTES_META.delegaciones.forEach(d => {
    const o = document.createElement('option'); o.value = d; o.textContent = d; sel.appendChild(o);
  });

  render();
}

// ── datos filtrados ───────────────────────────────────────────────────────────
function getData() {
  const pcMin = F.pcMin !== null ? F.pcMin : -Infinity;
  const pcMax = F.pcMax !== null ? F.pcMax :  Infinity;
  return APORTES_DATA.filter(r => {
    if (F.meses.length && !F.meses.includes(r.mes)) return false;
    if (F.deleg && r.deleg !== F.deleg) return false;
    if (r.pc < pcMin || r.pc > pcMax) return false;
    if (F.tam === 1 && r.miembros !== 1) return false;
    if (F.tam === 2 && r.miembros !== 2) return false;
    if (F.tam === 3 && r.miembros < 3) return false;
    return true;
  });
}

// ── render principal ──────────────────────────────────────────────────────────
function render() {
  const data = getData();
  renderKpis(data);
  renderResumen(data);
  renderGrupos();
  renderMensual(data);
}

// ── KPIs ──────────────────────────────────────────────────────────────────────
function renderKpis(d) {
  const gids = new Set(d.map(r=>r.gid));
  const miem = d.reduce((s,r)=>s+r.miembros,0);
  const tot  = d.reduce((s,r)=>s+r.tot,0);
  const m25  = d.filter(r=>r.rango==='menos25').length;
  const r25  = d.filter(r=>r.rango==='entre25y50').length;
  const pcProm = miem > 0 ? tot/miem : 0;
  set('k-grupos', gids.size.toLocaleString('es-AR'));
  set('k-miem',   miem.toLocaleString('es-AR'));
  set('k-m25',    m25.toLocaleString('es-AR'));
  set('k-r25',    r25.toLocaleString('es-AR'));
  set('k-tot',    arsMM(tot));
  set('k-pc',     ars(pcProm));
}

// ── Resumen por delegación ────────────────────────────────────────────────────
function renderResumen(data) {
  const map = {};
  data.forEach(r => {
    const k = r.deleg || 'Sin delegación';
    if (!map[k]) map[k] = {deleg:k,grupos:new Set(),miem:0,m25:0,r25:0,mas50:0,tot:0,pc_sum:0,pc_n:0};
    const g = map[k];
    g.grupos.add(r.gid); g.miem+=r.miembros;
    if(r.rango==='menos25') g.m25++; else if(r.rango==='entre25y50') g.r25++; else g.mas50++;
    g.tot+=r.tot; g.pc_sum+=r.pc; g.pc_n++;
  });
  let rows = Object.values(map).map(g=>({...g, grupos:g.grupos.size, pc:g.pc_n?g.pc_sum/g.pc_n:0}));
  const ss = SS.res;
  rows.sort((a,b)=>{ const v=typeof a[ss.c]==='string'?a[ss.c].localeCompare(b[ss.c]):a[ss.c]-b[ss.c]; return ss.a?v:-v; });
  set('cnt-res', rows.length + ' delegaciones');
  const tot = rows.reduce((a,r)=>({grupos:a.grupos+r.grupos,miem:a.miem+r.miem,m25:a.m25+r.m25,r25:a.r25+r.r25,mas50:a.mas50+r.mas50,tot:a.tot+r.tot,pc:0}),
    {grupos:0,miem:0,m25:0,r25:0,mas50:0,tot:0,pc:0});
  document.getElementById('tb-res').innerHTML =
    rows.map(r=>`<tr>
      <td>${r.deleg}</td>
      <td class="r">${r.grupos.toLocaleString('es-AR')}</td>
      <td class="r">${r.miem.toLocaleString('es-AR')}</td>
      <td class="r"><span class="badge b-m25">${r.m25}</span></td>
      <td class="r"><span class="badge b-25y50">${r.r25}</span></td>
      <td class="r"><span class="badge b-mas50">${r.mas50}</span></td>
      <td class="r">${ars(r.tot)}</td>
      <td class="r bold">${ars(r.pc)}</td>
    </tr>`).join('') +
    `<tr class="tr-total">
      <td>TOTAL</td>
      <td class="r">${tot.grupos.toLocaleString('es-AR')}</td>
      <td class="r">${tot.miem.toLocaleString('es-AR')}</td>
      <td class="r">${tot.m25.toLocaleString('es-AR')}</td>
      <td class="r">${tot.r25.toLocaleString('es-AR')}</td>
      <td class="r">${tot.mas50.toLocaleString('es-AR')}</td>
      <td class="r">${ars(tot.tot)}</td><td></td>
    </tr>`;
}

// ── Grupos familiares ─────────────────────────────────────────────────────────
const RBADGE = {
  menos25:    '<span class="badge b-m25">< $25k</span>',
  entre25y50: '<span class="badge b-25y50">$25k–$50k</span>',
  mas50:      '<span class="badge b-mas50">> $50k</span>',
};
function renderGrupos() {
  const srch = (document.getElementById('srch')?.value||'').toLowerCase();
  let data = getData();
  if (srch) data = data.filter(r=>(r.titular||'').toLowerCase().includes(srch)||(r.deleg||'').toLowerCase().includes(srch));
  const ss = SS.grp;
  data.sort((a,b)=>{ const v=typeof a[ss.c]==='string'?a[ss.c].localeCompare(b[ss.c]):a[ss.c]-b[ss.c]; return ss.a?v:-v; });
  const total=data.length, pages=Math.ceil(total/PS);
  if(grpPage>=pages) grpPage=0;
  const sl = data.slice(grpPage*PS,(grpPage+1)*PS);
  set('cnt-grp', total.toLocaleString('es-AR')+' grupos');
  set('grp-pag-info', `${grpPage*PS+1}–${Math.min((grpPage+1)*PS,total)} de ${total}`);
  document.getElementById('tb-grp').innerHTML = sl.map(r=>`<tr>
    <td><strong>${r.titular||'—'}</strong></td>
    <td>${r.deleg}</td>
    <td>${r.mes}</td>
    <td class="r"><span class="badge b-miem">${r.miembros} ${r.miembros===1?'persona':'personas'}</span></td>
    <td style="font-size:11px;color:var(--muted);max-width:160px">${r.paren||'—'}</td>
    <td class="r">${ars(r.ap)}</td>
    <td class="r">${ars(r.co)}</td>
    <td class="r">${ars(r.tot)}</td>
    <td class="r bold">${ars(r.pc)}</td>
    <td>${RBADGE[r.rango]||''}</td>
  </tr>`).join('') || '<tr><td colspan="10" style="text-align:center;color:var(--muted);padding:20px">Sin resultados</td></tr>';
  // Paginación
  let ph='';
  if(pages>1){
    if(grpPage>0) ph+=`<button class="pb" onclick="goP(${grpPage-1})">‹</button>`;
    const s=Math.max(0,grpPage-2),e=Math.min(pages-1,grpPage+2);
    for(let i=s;i<=e;i++) ph+=`<button class="pb${i===grpPage?' active':''}" onclick="goP(${i})">${i+1}</button>`;
    if(grpPage<pages-1) ph+=`<button class="pb" onclick="goP(${grpPage+1})">›</button>`;
  }
  document.getElementById('pag-grp').innerHTML = ph;
}
function goP(p){grpPage=p;renderGrupos();}

// ── Por mes ───────────────────────────────────────────────────────────────────
function renderMensual(data) {
  const map = {};
  data.forEach(r=>{
    if(!map[r.mes]) map[r.mes]={mes:r.mes,grupos:0,miem:0,m25:0,r25:0,mas50:0,tot:0,pc_s:0,pc_n:0};
    const g=map[r.mes]; g.grupos++; g.miem+=r.miembros;
    if(r.rango==='menos25') g.m25++; else if(r.rango==='entre25y50') g.r25++; else g.mas50++;
    g.tot+=r.tot; g.pc_s+=r.pc; g.pc_n++;
  });
  const meses=Object.keys(map).sort((a,b)=>{const[am,ay]=a.split('-'),[bm,by]=b.split('-');return(+ay*100+ +am)-(+by*100+ +bm);});
  document.getElementById('tb-mes').innerHTML = meses.map(m=>{
    const g=map[m];
    return `<tr>
      <td><strong>${m}</strong></td>
      <td class="r">${g.grupos.toLocaleString('es-AR')}</td>
      <td class="r">${g.miem.toLocaleString('es-AR')}</td>
      <td class="r"><span class="badge b-m25">${g.m25}</span></td>
      <td class="r"><span class="badge b-25y50">${g.r25}</span></td>
      <td class="r"><span class="badge b-mas50">${g.mas50}</span></td>
      <td class="r bold">${ars(g.tot)}</td>
      <td class="r">${ars(g.pc_n?g.pc_s/g.pc_n:0)}</td>
    </tr>`;
  }).join('')||'<tr><td colspan="8" style="text-align:center;color:var(--muted)">Sin datos</td></tr>';
}

// ── filtros ───────────────────────────────────────────────────────────────────
function setMes(m,btn){
  if(m===''){
    F.meses=[];
    document.querySelectorAll('#mes-btns .btn').forEach(b=>b.classList.remove('active'));
    btn.classList.add('active');
  } else {
    document.querySelector('#mes-btns .btn').classList.remove('active');
    btn.classList.toggle('active');
    F.meses=[...(document.querySelectorAll('#mes-btns .btn.active'))].map(b=>b.getAttribute('onclick').match(/'([^']+)'/)?.[1]).filter(Boolean);
    if(!F.meses.length){ document.querySelector('#mes-btns .btn').classList.add('active'); }
  }
  render();
}
function setAtajo(min, max, btn){
  document.getElementById('pc-min').value = min || '';
  document.getElementById('pc-max').value = max || '';
  F.pcMin = min || null;
  F.pcMax = max || null;
  ['b-m25','b-r25','b-mas50'].forEach(id=>document.getElementById(id).classList.remove('active'));
  btn.classList.add('active');
  render();
}
function limpiarAtajos(){
  ['b-m25','b-r25','b-mas50'].forEach(id=>document.getElementById(id).classList.remove('active'));
  const mn = parseFloat(document.getElementById('pc-min').value);
  const mx = parseFloat(document.getElementById('pc-max').value);
  F.pcMin = isNaN(mn) ? null : mn;
  F.pcMax = isNaN(mx) ? null : mx;
}
function setTam(t,btn){
  F.tam=t;
  document.querySelectorAll('.fila .btn').forEach(b=>{ if(b.getAttribute('onclick')&&b.getAttribute('onclick').includes('setTam')) b.classList.remove('active'); });
  btn.classList.add('active');
  render();
}
function presetBahia(){
  resetFiltros(true);
  F.deleg='OSPIF BAHIA BLANCA'; F.pcMin=25000; F.pcMax=50000;
  document.getElementById('sel-deleg').value='OSPIF BAHIA BLANCA';
  document.getElementById('pc-min').value='25000';
  document.getElementById('pc-max').value='50000';
  document.getElementById('b-r25').classList.add('active');
  render();
}
function descargarExcel(){
  const data = getData();
  const RBADGE_TXT = {menos25:'< $25k', entre25y50:'$25k–$50k', mas50:'> $50k'};
  // Hoja 1 — Grupos (detalle completo)
  const filas = data.map(r=>({
    'Titular':          r.titular || '',
    'Delegación':       r.deleg   || '',
    'Mes':              r.mes     || '',
    'Miembros':         r.miembros,
    'A cargo':          r.a_cargo,
    'Parentescos':      r.paren   || '',
    'Aportes ($)':      r.ap,
    'Contribuciones ($)':r.co,
    'Total grupo ($)':  r.tot,
    'Per cápita ($)':   r.pc,
    'Rango':            RBADGE_TXT[r.rango] || r.rango,
    'Grupo ID':         r.gid,
  }));
  // Hoja 2 — Resumen por delegación
  const mapD = {};
  data.forEach(r=>{
    const k = r.deleg||'Sin delegación';
    if(!mapD[k]) mapD[k]={deleg:k,grupos:new Set(),miem:0,m25:0,r25:0,mas50:0,tot:0};
    mapD[k].grupos.add(r.gid); mapD[k].miem+=r.miembros;
    if(r.rango==='menos25') mapD[k].m25++; else if(r.rango==='entre25y50') mapD[k].r25++; else mapD[k].mas50++;
    mapD[k].tot+=r.tot;
  });
  const filasDel = Object.values(mapD).map(g=>({
    'Delegación':      g.deleg,
    'Grupos únicos':   g.grupos.size,
    'Miembros totales':g.miem,
    'Menos $25k':      g.m25,
    '$25k–$50k':       g.r25,
    'Más $50k':        g.mas50,
    'Total aportado ($)': g.tot,
  }));

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(filas),    'Grupos');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(filasDel), 'Por Delegación');

  // nombre de archivo con filtros aplicados
  const parts = ['aportes'];
  if(F.meses.length===1) parts.push(F.meses[0]);
  if(F.deleg) parts.push(F.deleg.replace(/\s+/g,'_'));
  if(F.pcMin!==null||F.pcMax!==null) parts.push('pc'+(F.pcMin||0)+'a'+(F.pcMax||'max'));
  XLSX.writeFile(wb, parts.join('-')+'.xlsx');
}
function resetFiltros(silent){
  F={meses:[],deleg:'',pcMin:null,pcMax:null,tam:0};
  document.getElementById('sel-deleg').value='';
  document.getElementById('pc-min').value='';
  document.getElementById('pc-max').value='';
  document.querySelectorAll('#mes-btns .btn').forEach(b=>b.classList.remove('active'));
  document.querySelector('#mes-btns .btn').classList.add('active');
  ['b-m25','b-r25','b-mas50'].forEach(id=>document.getElementById(id).classList.remove('active'));
  document.querySelectorAll('.fila .btn').forEach(b=>{ if(b.getAttribute('onclick')&&b.getAttribute('onclick').includes('setTam(0')) b.classList.add('active'); });
  if(document.getElementById('srch')) document.getElementById('srch').value='';
  if(!silent) render();
}

// ── tabs ──────────────────────────────────────────────────────────────────────
function setTab(name,btn){
  document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
  document.querySelectorAll('.tab-pane').forEach(p=>p.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById('tab-'+name).classList.add('active');
  if(name==='grupos') renderGrupos();
}

// ── sort ──────────────────────────────────────────────────────────────────────
function srt(tbl,col){
  const ss=SS[tbl==='res'?'res':'grp'];
  if(ss.c===col) ss.a=!ss.a; else {ss.c=col;ss.a=col!=='tot'&&col!=='pc'&&col!=='grupos';}
  tbl==='res'?renderResumen(getData()):renderGrupos();
}

// ── helpers ───────────────────────────────────────────────────────────────────
function ars(n){ return n?'$ '+Math.round(n).toLocaleString('es-AR'):'$ 0'; }
function arsMM(n){ return n>=1e6?'$ '+(n/1e6).toFixed(1)+' M':ars(n); }
function set(id,v){ const el=document.getElementById(id); if(el) el.textContent=v; }
</script>
</body>
</html>
