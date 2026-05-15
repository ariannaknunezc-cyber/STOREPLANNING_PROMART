let rawData = [];
let chart;

document.getElementById('excelFile').addEventListener('change', function(e) {
    const reader = new FileReader();
    reader.onload = function(evt) {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        rawData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
        
        if(rawData.length > 0) {
            document.getElementById('controls').style.display = 'block';
            document.getElementById('noData').style.display = 'none';
            inicializarFiltroPais();
        }
    };
    reader.readAsArrayBuffer(e.target.files[0]);
});

function inicializarFiltroPais() {
    const paises = [...new Set(rawData.map(item => item.Pais))];
    const select = document.getElementById('paisFilter');
    select.innerHTML = paises.map(p => `<option value="${p}">${p}</option>`).join('');
    actualizarTiendas();
}

function actualizarTiendas() {
    const paisSeleccionado = document.getElementById('paisFilter').value;
    const tiendas = rawData.filter(item => item.Pais === paisSeleccionado);
    const nombresTiendas = [...new Set(tiendas.map(item => item.Tienda))];
    
    const select = document.getElementById('tiendaFilter');
    select.innerHTML = nombresTiendas.map(t => `<option value="${t}">${t}</option>`).join('');
    actualizarGrafico();
}

function actualizarGrafico() {
    const tiendaSel = document.getElementById('tiendaFilter').value;
    const metrica = document.getElementById('metricaFilter').value;
    
    const dataFiltrada = rawData.filter(item => item.Tienda === tiendaSel);
    
    const labels = dataFiltrada.map(item => item.Area_Comercial);
    const values = dataFiltrada.map(item => item[metrica]);

    if (chart) chart.destroy();

    const ctx = document.getElementById('lineChart').getContext('2d');
    chart = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [{
                label: metrica === 'm2' ? 'Metros Cuadrados (m²)' : 'Metros Cúbicos (m³)',
                data: values,
                borderColor: '#F15A22',
                backgroundColor: 'rgba(241, 90, 34, 0.2)',
                borderWidth: 4,
                fill: true,
                tension: 0.4,
                pointRadius: 6,
                pointBackgroundColor: '#fff'
            }]
        },
        options: {
            responsive: true,
            scales: {
                y: { beginAtZero: true, grid: { color: '#333' }, ticks: { color: '#aaa' } },
                x: { grid: { display: false }, ticks: { color: '#aaa' } }
            },
            plugins: {
                legend: { labels: { color: '#fff', font: { size: 14 } } }
            }
        }
    });
}
