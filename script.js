/* ─── CONSTANTS ─────────────────────────── */
const PALETTE = [
  '#F15A22','#38bdf8','#fbbf24','#10b981',
  '#f43f5e','#a855f7','#06b6d4','#84cc16',
  '#fb923c','#e879f9','#f97316','#14b8a6'
];

let rawData   = [];
let chartInst = null;
let curPais   = '';

/* ─── DRAG & DROP ───────────────────────── */
const upzone = document.getElementById('upzone');
upzone.addEventListener('dragover',  e=>{ e.preventDefault(); upzone.style.borderColor='#F15A22'; });
upzone.addEventListener('dragleave', ()=>{ upzone.style.borderColor=''; });
upzone.addEventListener('drop', e=>{
  e.preventDefault(); upzone.style.borderColor='';
  if(e.dataTransfer.files[0]) readFile(e.dataTransfer.files[0]);
});

/* ─── FILE INPUT ────────────────────────── */
document.getElementById('excelInput').addEventListener('change', function(e){
  if(e.target.files[0]) readFile(e.target.files[0]);
});

function readFile(file){
  const reader = new FileReader();
  reader.onload = function(evt){
    try{
      const wb   = XLSX.read(new Uint8Array(evt.target.result), {type:'array'});
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
      if(!rows.length){ alert('El archivo está vacío.'); return; }
      rawData = rows;
      upzone.innerHTML = `<div class="upzone-ico">✅</div><p style="color:rgba(255,255,255,.65);">${rows.length} registros cargados<br><b style="color:#F15A22;">${file.name}</b></p>`;
      document.getElementById('dyn-ctrl').style.display='block';
      initPaises();
    } catch(err){ alert('Error al leer. Verifica que sea .xlsx o .xls'); }
  };
  reader.readAsArrayBuffer(file);
}

/* ─── INIT PAÍSES ───────────────────────── */
function initPaises(){
  const paises = [...new Set(rawData.map(r=> r.Pais||r.pais||r.PAIS||'—'))];
  const sel = document.getElementById('paisFilter');
  sel.innerHTML = paises.map(p=>`<option value="${p}">${p}</option>`).join('');
  curPais = paises[0];
  updateStores();
}

/* ─── UPDATE TIENDAS ────────────────────── */
function updateStores(){
  curPais = document.getElementById('paisFilter').value;
  const tiendas = [...new Set(
    rawData
      .filter(r=>(r.Pais||r.pais||r.PAIS||'—')===curPais)
      .map(r=>r.Tienda||r.tienda||r.TIENDA||'—')
  )];

  /* checklist */
  document.getElementById('tiendaChecklist').innerHTML = tiendas.map((t,i)=>`
    <label>
      <input type="checkbox" class="store-cb" value="${t}" onchange="renderChart()" checked>
      <span style="display:inline-flex;align-items:center;gap:6px;">
        <span style="width:9px;height:9px;border-radius:50%;background:${PALETTE[i%PALETTE.length]};flex-shrink:0;"></span>
        ${t}
      </span>
    </label>`).join('');

  /* areas */
  const areas = [...new Set(
    rawData
      .filter(r=>(r.Pais||r.pais||r.PAIS||'—')===curPais)
      .map(r=>r.Area_Comercial||r.Area||r.area||'—')
  )];
  const aSel = document.getElementById('areaFilter');
  aSel.innerHTML = '<option value="__all__">— Todas las áreas —</option>' +
    areas.map(a=>`<option value="${a}">${a}</option>`).join('');

  renderChart();
}

function selectAll(v){
  document.querySelectorAll('.store-cb').forEach(cb=>cb.checked=v);
  renderChart();
}

/* ─── RENDER CHART ──────────────────────── */
function renderChart(){
  const selected  = Array.from(document.querySelectorAll('.store-cb:checked')).map(cb=>cb.value);
  const metrica   = document.getElementById('metricaFilter').value;
  const areaVal   = document.getElementById('areaFilter').value;
  const chartType = document.getElementById('chartTypeFilter').value;

  const legendEl  = document.getElementById('legend');
  const emptyEl   = document.getElementById('emptyMsg');
  const canvas    = document.getElementById('mainChart');

  if(!selected.length){
    if(chartInst) chartInst.destroy();
    canvas.style.display='none'; emptyEl.style.display='flex';
    legendEl.innerHTML='<span style="font-size:12px;color:rgba(255,255,255,.2);font-family:JetBrains Mono,monospace;">Selecciona al menos una tienda</span>';
    return;
  }

  /* filter rows */
  let subset = rawData.filter(r=>
    (r.Pais||r.pais||r.PAIS||'—')===curPais &&
    selected.includes(r.Tienda||r.tienda||r.TIENDA||'—')
  );
  if(areaVal!=='__all__'){
    subset = subset.filter(r=>(r.Area_Comercial||r.Area||r.area||'—')===areaVal);
  }

  const areasX = [...new Set(subset.map(r=>r.Area_Comercial||r.Area||r.area||'—'))];
  const metLabel = metrica==='m2'?'Metros Cuadrados (m²)':'Metros Cúbicos (m³)';

  document.getElementById('ctitle').textContent = `Comparativo ${metLabel} — ${curPais}`;
  document.getElementById('csub').textContent   = `${selected.length} tienda${selected.length>1?'s':''} · ${areasX.length} área${areasX.length!==1?'s':''}`;

  /* datasets */
  const datasets = selected.map((tienda,i)=>{
    const color = PALETTE[i%PALETTE.length];
    const rows  = subset.filter(r=>(r.Tienda||r.tienda||r.TIENDA||'—')===tienda);
    const vals  = areasX.map(area=>{
      const row = rows.find(r=>(r.Area_Comercial||r.Area||r.area||'—')===area);
      return row ? parseFloat(row[metrica]||row[metrica.toUpperCase()]||0) : 0;
    });
    const isLine = chartType==='line';
    return{
      label:tienda, data:vals,
      borderColor:color,
      backgroundColor: isLine ? color+'22' : color+'CC',
      borderWidth:2.5, tension:0.35,
      pointRadius:4, pointHoverRadius:7,
      fill:isLine,
    };
  });

  /* legend */
  legendEl.innerHTML = datasets.map(ds=>`
    <div class="legend-item">
      <div class="ldot" style="background:${ds.borderColor}"></div>
      ${ds.label}
    </div>`).join('');

  if(chartInst) chartInst.destroy();
  emptyEl.style.display='none';
  canvas.style.display='block';

  chartInst = new Chart(canvas.getContext('2d'),{
    type: chartType,
    data:{ labels:areasX, datasets },
    options:{
      responsive:true, maintainAspectRatio:false,
      interaction:{mode:'index',intersect:false},
      plugins:{
        legend:{display:false},
        tooltip:{
          backgroundColor:'#111',titleColor:'#fff',
          bodyColor:'rgba(255,255,255,.7)',
          borderColor:'rgba(241,90,34,.4)',borderWidth:1,padding:12,
          callbacks:{
            label: ctx=>` ${ctx.dataset.label}: ${ctx.parsed.y.toLocaleString('es-PE')} ${metrica==='m2'?'m²':'m³'}`
          }
        }
      },
      scales:{
        x:{
          grid:{display:false},
          ticks:{color:'rgba(255,255,255,.42)',font:{size:11,family:'Barlow'},maxRotation:36}
        },
        y:{
          grid:{color:'rgba(255,255,255,.05)'},
          ticks:{
            color:'rgba(255,255,255,.42)',font:{size:11,family:'JetBrains Mono'},
            callback: v=>v.toLocaleString('es-PE')+(metrica==='m2'?' m²':' m³')
          }
        }
      },
      animation:{duration:500,easing:'easeOutQuart'}
    }
  });
}

/* ─── CLEAR ─────────────────────────────── */
function clearData(){
  rawData=[];
  if(chartInst){chartInst.destroy();chartInst=null;}
  document.getElementById('dyn-ctrl').style.display='none';
  document.getElementById('mainChart').style.display='none';
  document.getElementById('emptyMsg').style.display='flex';
  document.getElementById('excelInput').value='';
  document.getElementById('legend').innerHTML='<span style="font-size:12px;color:rgba(255,255,255,.2);font-family:JetBrains Mono,monospace;">La leyenda aparecerá aquí con el color de cada tienda</span>';
  document.getElementById('ctitle').textContent='Comparativo por área comercial';
  document.getElementById('csub').textContent='Carga un Excel para comenzar';
  upzone.innerHTML='<div class="upzone-ico">📂</div><p>Arrastra o <b>haz clic</b> para cargar<br><span style="font-size:11px;color:rgba(255,255,255,.25)">.xlsx · .xls · .csv</span></p>';
}

/* ─── DOWNLOAD TEMPLATE ─────────────────── */
function downloadTemplate(){
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ['Pais','Tienda','Area_Comercial','m2','m3'],
    ['Peru','Promart Chorrillos','Herramientas',1850,7400],
    ['Peru','Promart Chorrillos','Ferretería',920,3680],
    ['Peru','Promart Chorrillos','Jardín y Mascotas',640,2560],
    ['Peru','Promart Miraflores','Herramientas',1200,4800],
    ['Peru','Promart Miraflores','Ferretería',680,2720],
    ['Peru','Promart Miraflores','Jardín y Mascotas',480,1920],
    ['Peru','Promart Lima Norte','Herramientas',2200,8800],
    ['Peru','Promart Lima Norte','Ferretería',1100,4400],
    ['Peru','Promart Lima Norte','Jardín y Mascotas',750,3000],
    ['Chile','Promart Santiago','Herramientas',2100,8400],
    ['Chile','Promart Santiago','Ferretería',980,3920],
    ['Chile','Promart Concepción','Herramientas',1650,6600],
    ['Chile','Promart Concepción','Ferretería',820,3280],
  ]);
  ws['!cols']=[12,22,22,10,10].map(wch=>({wch}));
  XLSX.utils.book_append_sheet(wb,ws,'Datos');
  XLSX.writeFile(wb,'plantilla_analisis_promart.xlsx');
}

/* ─── NAV ACTIVE ON SCROLL ──────────────── */
const navLinks = document.querySelectorAll('.menu a');
const observer = new IntersectionObserver(entries=>{
  entries.forEach(entry=>{
    if(entry.isIntersecting){
      const id=entry.target.id;
      navLinks.forEach(a=>a.classList.toggle('active', a.getAttribute('href')==='#'+id));
    }
  });
},{threshold:0.3});
document.querySelectorAll('section[id]').forEach(s=>observer.observe(s));
