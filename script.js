let storeData = [];
let chartInstance = null;

document.getElementById('excelFile').addEventListener('change', function(e) {
    const reader = new FileReader();
    reader.onload = function(evt) {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        storeData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
        
        if(storeData.length > 0) {
            initControls();
            document.getElementById('filter-section').style.display = 'block';
            updateChart();
        }
    };
    reader.readAsArrayBuffer(e.target.files[0]);
});

function initControls() {
    const storeList = document.getElementById('store-list');
    const areaFilter = document.getElementById('areaFilter');
    
    // Limpiar
    storeList.innerHTML = '';
    areaFilter.innerHTML = '<option value="all">Todas las Áreas</option>';

    // Tiendas Únicas para Checkboxes
    storeData.forEach((item, index) => {
        const div = document.createElement('div');
        div.className = 'store-item';
        div.innerHTML = `<input type="checkbox" checked value="${index}" class="store-cb"> <span>${item.Tienda}</span>`;
        storeList.appendChild(div);
    });

    // Áreas Únicas para el Selector
    const areas = [...new Set(storeData.map(item => item.Area_Comercial))];
    areas.forEach(area => {
        if(area) areaFilter.innerHTML += `<option value="${area}">${area}</option>`;
    });
}

function updateChart() {
    const selectedIndices = Array.from(document.querySelectorAll('.store-cb:checked')).map(cb => parseInt(cb.value));
    const areaVal = document.getElementById('areaFilter').value;

    let filtered = storeData.filter((_, index) => selectedIndices.includes(index));
    
    if (areaVal !== 'all') {
        filtered = filtered.filter(item => item.Area_Comercial === areaVal);
    }

    const labels = filtered.map(item => item.Area_Comercial || 'S/N');
    const m2Data = filtered.map(item => item.m2 || 0);
    const m3Data = filtered.map(item => item.m3 || 0);
    const tiendas = filtered.map(item => item.Tienda);

    if (chartInstance) chartInstance.destroy();

    const ctx = document.getElementById('capexChart').getContext('2d');
    chartInstance = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [
                {
                    label: 'm²',
                    data: m2Data,
                    borderColor: '#ec4899',
                    backgroundColor: 'rgba(236, 72, 153, 0.1)',
                    borderWidth: 4,
                    tension: 0.4,
                    fill: true,
                    pointRadius: 6
                },
                {
                    label: 'm³',
                    data: m3Data,
                    borderColor: '#14b8a6',
                    backgroundColor: 'rgba(20, 184, 166, 0.1)',
                    borderWidth: 4,
                    tension: 0.4,
                    fill: true,
                    pointRadius: 6
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: {
                    callbacks: {
                        afterLabel: (context) => `Tienda: ${tiendas[context.dataIndex]}`
                    }
                }
            },
            scales: {
                x: { 
                    ticks: { color: '#94a3b8' },
                    grid: { display: false },
                    title: { display: true, text: 'ÁREA COMERCIAL', color: '#64748b' }
                },
                y: { 
                    ticks: { color: '#94a3b8' },
                    grid: { color: 'rgba(255,255,255,0.1)' }
                }
            }
        }
    });
}
