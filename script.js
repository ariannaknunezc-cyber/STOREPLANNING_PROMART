let rawStoreData = [];
let myChart = null;

const excelInput = document.getElementById('excelFile');

excelInput.addEventListener('change', function(e) {
    const file = e.target.files[0];
    const reader = new FileReader();

    reader.onload = function(evt) {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        const sheetName = workbook.SheetNames[0];
        rawStoreData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

        if(rawStoreData.length > 0) {
            renderSidebar();
            document.getElementById('dashboard').style.display = 'grid';
            updateChart(); // Primera carga automática
        }
    };
    reader.readAsArrayBuffer(file);
});

function renderSidebar() {
    const container = document.getElementById('store-list');
    container.innerHTML = '';
    
    rawStoreData.forEach((item, index) => {
        const div = document.createElement('label');
        div.className = 'store-item';
        div.innerHTML = `
            <input type="checkbox" checked value="${index}" class="store-check">
            <span>${item.Tienda}</span>
        `;
        container.appendChild(div);
    });
}

function updateChart() {
    const selectedIndices = Array.from(document.querySelectorAll('.store-check:checked')).map(cb => parseInt(cb.value));
    const filteredData = rawStoreData.filter((_, index) => selectedIndices.includes(index));

    // Eje X: Áreas Comerciales
    const labels = filteredData.map(item => item.Area_Comercial || 'S/N');
    const m2Values = filteredData.map(item => item.m2 || 0);
    const m3Values = filteredData.map(item => item.m3 || 0);
    const tiendas = filteredData.map(item => item.Tienda);

    if (myChart) myChart.destroy();

    const ctx = document.getElementById('storeChart').getContext('2d');
    myChart = new Chart(ctx, {
        type: 'line', // Tipo de gráfico: Línea
        data: {
            labels: labels,
            datasets: [
                {
                    label: 'Metros Cuadrados (m²)',
                    data: m2Values,
                    borderColor: '#F15A22',
                    backgroundColor: 'rgba(241, 90, 34, 0.2)',
                    fill: true,
                    tension: 0.3,
                    borderWidth: 3,
                    pointRadius: 5
                },
                {
                    label: 'Metros Cúbicos (m³)',
                    data: m3Values,
                    borderColor: '#111111',
                    backgroundColor: 'rgba(17, 17, 17, 0.1)',
                    fill: true,
                    tension: 0.3,
                    borderWidth: 3,
                    pointRadius: 5
                }
            ]
        },
        options: {
            responsive: true,
            plugins: {
                legend: { position: 'top' },
                tooltip: {
                    callbacks: {
                        afterLabel: (context) => `Tienda: ${tiendas[context.dataIndex]}`
                    }
                }
            },
            scales: {
                x: { 
                    title: { display: true, text: 'ÁREAS COMERCIALES' },
                    grid: { display: false }
                },
                y: { 
                    title: { display: true, text: 'VALORES METRICOS' },
                    beginAtZero: true 
                }
            }
        }
    });
}
