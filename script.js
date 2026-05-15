let dataMaestra = [];
let miGrafico;

document.getElementById('excelInput').addEventListener('change', function(e) {
    const reader = new FileReader();
    reader.onload = function(evt) {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        dataMaestra = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
        
        if(dataMaestra.length > 0) {
            document.getElementById('controls').style.display = 'block';
            inicializarFiltros();
        }
    };
    reader.readAsArrayBuffer(e.target.files[0]);
});

function inicializarFiltros() {
    // Cargar Países
    const paises = [...new Set(dataMaestra.map(x => x.Pais))];
    const paisSel = document.getElementById('paisFilter');
    paisSel.innerHTML = paises.map(p => `<option value="${p}">${p}</option>`).join('');

    // Cargar Áreas
    const areas = [...new Set(dataMaestra.map(x => x.Area_Comercial))];
    const areaSel = document.getElementById('areaFilter');
    areaSel.innerHTML = '<option value="all">TODAS LAS AREAS</option>' + 
                        areas.map(a => `<option value="${a}">${a}</option>`).join('');

    actualizarListaTiendas();
}

function actualizarListaTiendas() {
    const pais = document.getElementById('paisFilter').value;
    const tiendas = [...new Set(dataMaestra.filter(x => x.Pais === pais).map(x => x.Tienda))];
    const list = document.getElementById('tiendaChecklist');
    
    list.innerHTML = tiendas.map((t, i) => `
        <label><input type="checkbox" class="t-check" value="${t}" checked onchange="renderizarGrafico()"> ${t}</label>
    `).join('');
    
    renderizarGrafico();
}

function renderizarGrafico() {
    const tiendasSel = Array.from(document.querySelectorAll('.t-check:checked')).map(x => x.value);
    const areaSel = document.getElementById('areaFilter').value;
    const metrica = document.getElementById('metricaFilter').value;

    let filtrado = dataMaestra.filter(x => tiendasSel.includes(x.Tienda));
    if(areaSel !== 'all') filtrado = filtrado.filter(x => x.Area_Comercial === areaSel);

    if (miGrafico) miGrafico.destroy();

    const colores = ['#F15A22', '#38bdf8', '#fbbf24', '#10b981', '#f43f5e', '#a855f7'];
    const datasets = tiendasSel.map((tienda, i) => {
        const puntos = filtrado.filter(x => x.Tienda === tienda);
        return {
            label: tienda,
            data: puntos.map(x => x[metrica]),
            borderColor: colores[i % colores.length],
            borderWidth: 4,
            tension: 0.3,
            pointRadius: 5
        };
    });

    const labels = [...new Set(filtrado.map(x => x.Area_Comercial))];

    const ctx = document.getElementById('lineChart').getContext('2d');
    miGrafico = new Chart(ctx, {
        type: 'line',
        data: { labels, datasets },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: true, labels: { color: '#fff' } } },
            scales: {
                y: { grid: { color: '#334155' }, ticks: { color: '#94a3b8' } },
                x: { grid: { display: false }, ticks: { color: '#94a3b8' } }
            }
        }
    });
}
