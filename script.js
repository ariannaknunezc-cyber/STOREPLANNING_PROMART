let rawData = [];
let chartInstance = null;

document.getElementById('excelInput').addEventListener('change', function(e) {
    const reader = new FileReader();
    reader.onload = function(evt) {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        rawData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
        if(rawData.length > 0) {
            document.getElementById('dynamic-controls').style.display = 'block';
            logic_initPaises();
        }
    };
    reader.readAsArrayBuffer(e.target.files[0]);
});

function logic_initPaises() {
    const paises = [...new Set(rawData.map(item => item.Pais))];
    const select = document.getElementById('paisFilter');
    select.innerHTML = paises.map(p => `<option value="${p}">${p}</option>`).join('');
    logic_updateStores();
}

function logic_updateStores() {
    const selectedPais = document.getElementById('paisFilter').value;
    const tiendas = [...new Set(rawData.filter(item => item.Pais === selectedPais).map(item => item.Tienda))];
    const checklist = document.getElementById('tiendaChecklist');
    
    checklist.innerHTML = tiendas.map(t => `
        <label><input type="checkbox" class="store-cb" value="${t}" onchange="logic_renderChart()" checked> ${t}</label>
    `).join('');
    logic_renderChart();
}

function logic_renderChart() {
    const selectedStores = Array.from(document.querySelectorAll('.store-cb:checked')).map(cb => cb.value);
    const metrica = document.getElementById('metricaFilter').value;
    const legendContainer = document.getElementById('chart-legend');
    
    if (selectedStores.length === 0) {
        if(chartInstance) chartInstance.destroy();
        legendContainer.innerHTML = '';
        return;
    }

    const colorPalette = ['#F15A22', '#38bdf8', '#fbbf24', '#10b981', '#f43f5e', '#a855f7'];
    const dataForChart = rawData.filter(item => selectedStores.includes(item.Tienda));
    const areasX = [...new Set(dataForChart.map(item => item.Area_Comercial))];

    const datasets = selectedStores.map((tienda, i) => {
        const color = colorPalette[i % colorPalette.length];
        const storeRows = dataForChart.filter(r => r.Tienda === tienda);
        const values = areasX.map(area => {
            const row = storeRows.find(r => r.Area_Comercial === area);
            return row ? row[metrica] : 0;
        });
        return { label: tienda, data: values, borderColor: color, backgroundColor: color, tension: 0.3, borderWidth: 3 };
    });

    legendContainer.innerHTML = datasets.map(ds => `
        <div style="display:flex; align-items:center; gap:5px; font-size:12px; color:#94a3b8">
            <div style="width:10px; height:10px; border-radius:50%; background:${ds.borderColor}"></div> ${ds.label}
        </div>
    `).join('');

    if (chartInstance) chartInstance.destroy();
    const ctx = document.getElementById('mainChart').getContext('2d');
    chartInstance = new Chart(ctx, {
        type: 'line',
        data: { labels: areasX, datasets: datasets },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: {
                y: { grid: { color: '#1e293b' }, ticks: { color: '#94a3b8' } },
                x: { grid: { display: false }, ticks: { color: '#94a3b8' } }
            }
        }
    });
}
