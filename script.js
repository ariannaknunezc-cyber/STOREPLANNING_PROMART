let db = [];
let chart;

document.getElementById('excelFile').addEventListener('change', function(e) {
    const reader = new FileReader();
    reader.onload = function(evt) {
        const data = new Uint8Array(evt.target.result);
        const wb = XLSX.read(data, {type: 'array'});
        db = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
        
        if(db.length > 0) {
            document.getElementById('dynamic-controls').style.display = 'block';
            initFilters();
        }
    };
    reader.readAsArrayBuffer(e.target.files[0]);
});

function initFilters() {
    // 1. Cargar Países
    const paises = [...new Set(db.map(x => x.Pais))];
    const paisSel = document.getElementById('paisSel');
    paisSel.innerHTML = paises.map(p => `<option value="${p}">${p}</option>`).join('');
    
    // 2. Cargar Áreas
    const areas = [...new Set(db.map(x => x.Area_Comercial))];
    const areaSel = document.getElementById('areaSel');
    areaSel.innerHTML = `<option value="all">TODAS LAS AREAS</option>` + 
                        areas.map(a => `<option value="${a}">${a}</option>`).join('');

    filterStores();
}

function filterStores() {
    const pSel = document.getElementById('paisSel').value;
    const tiendas = [...new Set(db.filter(x => x.Pais === pSel).map(x => x.Tienda))];
    const checklist = document.getElementById('store-checklist');
    
    checklist.innerHTML = tiendas.map(t => `
        <label><input type="checkbox" class="st-check" value="${t}" checked> ${t}</label>
    `).join('');
    
    updateChart();
}

function updateChart() {
    const selectedStores = Array.from(document.querySelectorAll('.st-check:checked')).map(x => x.value);
    const selectedArea = document.getElementById('areaSel').value;
    const metric = document.getElementById('metricSel').value;

    let filtered = db.filter(x => selectedStores.includes(x.Tienda));
    
    if(selectedArea !== 'all') {
        filtered = filtered.filter(x => x.Area_Comercial === selectedArea);
    }

    // Preparar Datos para Líneas por Tienda
    if (chart) chart.destroy();
    
    const datasets = selectedStores.map((store, i) => {
        const storeData = filtered.filter(x => x.Tienda === store);
        return {
            label: store,
            data: storeData.map(x => x[metric]),
            borderColor: i % 2 === 0 ? '#F15A22' : '#ffffff',
            borderWidth: 4,
            tension: 0.4,
            pointRadius: 6
        };
    });

    const labels = [...new Set(filtered.map(x => x.Area_Comercial))];

    chart = new Chart(document.getElementById('mainChart'), {
        type: 'line',
        data: { labels, datasets },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: { grid: { color: '#333' }, ticks: { color: '#fff' } },
                x: { grid: { display: false }, ticks: { color: '#fff' } }
            },
            plugins: { legend: { labels: { color: '#fff' } } }
        }
    });
}
