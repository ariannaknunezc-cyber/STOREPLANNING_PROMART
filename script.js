const ctx = document.getElementById('lineChart').getContext('2d');
let lineChart = new Chart(ctx, {
    type: 'line',
    data: {
        labels: ['Tienda A', 'Tienda B', 'Tienda C'],
        datasets: [
            { label: 'm²', data: [1200, 1900, 1500], borderColor: '#F15A22', tension: 0.3 },
            { label: 'm³', data: [4000, 7000, 5500], borderColor: '#FFFFFF', tension: 0.3 }
        ]
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

document.getElementById('excelFile').addEventListener('change', function(e) {
    const file = e.target.files[0];
    const reader = new FileReader();
    reader.onload = function(evt) {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet);

        if(json.length > 0) {
            lineChart.data.labels = json.map(item => item.Tienda || item.tienda || 'Tienda');
            lineChart.data.datasets[0].data = json.map(item => item.m2 || 0);
            lineChart.data.datasets[1].data = json.map(item => item.m3 || 0);
            lineChart.update();
        }
    };
    reader.readAsArrayBuffer(file);
});
