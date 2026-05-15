let storeData = [];
let chart;

const ctx = document.getElementById('storeChart').getContext('2d');

function initChart(labels, values, labelName) {
    if (chart) chart.destroy();
    chart = new Chart(ctx, {
        type: 'bar', // Cambiado a barras para mejor comparación de tiendas individuales
        data: {
            labels: labels,
            datasets: [{
                label: labelName,
                data: values,
                backgroundColor: '#F15A22',
                borderRadius: 5
            }]
        },
        options: {
            responsive: true,
            plugins: { legend: { labels: { color: '#fff' } } },
            scales: {
                y: { ticks: { color: '#fff' }, grid: { color: 'rgba(255,255,255,0.1)' } },
                x: { ticks: { color: '#fff' } }
            }
        }
    });
}

document.getElementById('excelFile').addEventListener('change', function(e) {
    const reader = new FileReader();
    reader.onload = function(evt) {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        storeData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);

        // Llenar selector de Áreas
        const areas = [...new Set(storeData.map(item => item.Area_Comercial))];
        const filterArea = document.getElementById('filterArea');
        filterArea.innerHTML = '<option value="todas">Todas las Áreas</option>';
        areas.forEach(area => {
            filterArea.innerHTML += `<option value="${area}">${area}</option>`;
        });

        document.getElementById('controls').style.display = 'flex';
        updateVisualization();
    };
    reader.readAsArrayBuffer(e.target.files[0]);
});

function updateVisualization() {
    const areaFilter = document.getElementById('filterArea').value;
    const metric = document.getElementById('metricType').value;
    
    let filtered = storeData;
    if (areaFilter !== 'todas') {
        filtered = storeData.filter(item => item.Area_Comercial === areaFilter);
    }

    const labels = filtered.map(item => item.Tienda);
    const values = filtered.map(item => item[metric]);
    const labelName = metric === 'm2' ? 'Metros Cuadrados (m²)' : 'Metros Cúbicos (m³)';

    initChart(labels, values, labelName);
}

// Eventos de los filtros
document.getElementById('filterArea').addEventListener('change', updateVisualization);
document.getElementById('metricType').addEventListener('change', updateVisualization);
