let rawData = [];
let myChart = null;

document.getElementById('excelFile').addEventListener('change', function(e) {
    const reader = new FileReader();
    reader.onload = function(evt) {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        rawData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
        
        displayStoreCheckboxes();
        document.getElementById('filter-panel').style.display = 'block';
        updateChart();
    };
    reader.readAsArrayBuffer(e.target.files[0]);
});

function displayStoreCheckboxes() {
    const container = document.getElementById('store-checkboxes');
    container.innerHTML = '';
    rawData.forEach((item, index) => {
        const label = document.createElement('label');
        label.innerHTML = `<input type="checkbox" checked value="${index}" class="store-cb"> ${item.Tienda} (${item.Area_Comercial})`;
        container.appendChild(label);
    });
}

function updateChart() {
    const selectedIndices = Array.from(document.querySelectorAll('.store-cb:checked')).map(cb => cb.value);
    const metric = document.getElementById('metricSelector').value;
    
    const filteredData = rawData.filter((_, index) => selectedIndices.includes(index.toString()));
    
    // Eje X: Área Comercial | Eje Y: m2 o m3
    const labels = filteredData.map(item => item.Area_Comercial);
    const values = filteredData.map(item => item[metric]);
    const tiendaNames = filteredData.map(item => item.Tienda);

    if (myChart) myChart.destroy();
    
    const ctx = document.getElementById('storeChart').getContext('2d');
    myChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: metric === 'm2' ? 'Superficie (m²)' : 'Volumen (m³)',
                data: values,
                backgroundColor: '#F15A22'
            }]
        },
        options: {
            plugins: {
                tooltip: {
                    callbacks: {
                        afterLabel: (context) => `Tienda: ${tiendaNames[context.dataIndex]}`
                    }
                }
            },
            scales: {
                x: { title: { display: true, text: 'ÁREA COMERCIAL' } }
            }
        }
    });
}
