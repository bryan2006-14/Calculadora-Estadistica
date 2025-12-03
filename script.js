// Variables globales
let chartInstances = {};
let processedData = null;
let expandedChartInstance = null;

// Event listeners
document.getElementById('dataType').addEventListener('change', function() {
    const noAgrupadosInput = document.getElementById('noAgrupadosInput');
    const agrupadosInput = document.getElementById('agrupadosInput');
    
    if (this.value === 'agrupados') {
        agrupadosInput.style.display = 'block';
    } else {
        agrupadosInput.style.display = 'none';
    }
});

document.getElementById('processBtn').addEventListener('click', processData);
document.getElementById('clearBtn').addEventListener('click', clearAll);
document.getElementById('exportExcelBtn').addEventListener('click', exportToExcel);

// Función principal de procesamiento
function processData() {
    const input = document.getElementById('dataInput').value.trim();
    if (!input) {
        showNotification('Por favor ingrese los datos', 'warning');
        return;
    }

    try {
        // Parsear datos
        const data = input.split(/[\s,]+/).map(Number).filter(n => !isNaN(n));
        
        if (data.length === 0) {
            showNotification('No se encontraron datos válidos', 'danger');
            return;
        }

        // Ordenar datos
        const sortedData = [...data].sort((a, b) => a - b);
        displaySortedData(sortedData);

        // Procesar según el tipo
        const dataType = document.getElementById('dataType').value;
        
        if (dataType === 'no-agrupados') {
            processNoAgrupados(sortedData);
        } else {
            const numClases = parseInt(document.getElementById('numClases').value);
            processAgrupados(sortedData, numClases);
        }

        // Mostrar sección de resultados y botón de exportar
        document.getElementById('resultsSection').style.display = 'block';
        document.getElementById('exportExcelBtn').style.display = 'block';
        
        showNotification('Datos procesados exitosamente', 'success');
        
        // Scroll suave a resultados
        setTimeout(() => {
            document.getElementById('resultsSection').scrollIntoView({ 
                behavior: 'smooth', 
                block: 'start' 
            });
        }, 300);
        
    } catch (error) {
        showNotification('Error al procesar los datos: ' + error.message, 'danger');
        console.error(error);
    }
}

// Mostrar datos ordenados
function displaySortedData(data) {
    const container = document.getElementById('sortedData');
    container.innerHTML = `
        <div class="mb-2">
            <span class="badge bg-primary">n = ${data.length}</span>
            <span class="badge bg-info">Min = ${Math.min(...data)}</span>
            <span class="badge bg-warning text-dark">Max = ${Math.max(...data)}</span>
        </div>
        <div style="word-wrap: break-word;">${data.join(', ')}</div>
    `;
}

// Procesar datos no agrupados
function processNoAgrupados(data) {
    const freqMap = {};
    data.forEach(val => {
        freqMap[val] = (freqMap[val] || 0) + 1;
    });

    const freqData = Object.entries(freqMap).map(([val, freq]) => ({
        value: parseFloat(val),
        frequency: freq
    })).sort((a, b) => a.value - b.value);

    let cumFreq = 0;
    freqData.forEach(item => {
        cumFreq += item.frequency;
        item.cumFrequency = cumFreq;
        item.relativeFreq = (item.frequency / data.length * 100).toFixed(2);
        item.cumRelativeFreq = (item.cumFrequency / data.length * 100).toFixed(2);
    });

    displayFrequencyTable(freqData, false);
    
    const measures = calculateMeasures(data);
    displayMeasures(measures);
    
    const quantiles = calculateQuantiles(data);
    displayQuantiles(quantiles);
    
    createCharts(freqData, data, false);
    
    processedData = { data, freqData, measures, quantiles, isGrouped: false };
}

// Procesar datos agrupados
function processAgrupados(data, numClases) {
    const min = Math.min(...data);
    const max = Math.max(...data);
    const range = max - min;
    const classWidth = Math.ceil(range / numClases);

    const intervals = [];
    for (let i = 0; i < numClases; i++) {
        const lower = min + (i * classWidth);
        const upper = lower + classWidth;
        intervals.push({
            lower,
            upper,
            midpoint: (lower + upper) / 2,
            frequency: 0,
            values: []
        });
    }

    data.forEach(val => {
        for (let i = 0; i < intervals.length; i++) {
            if (i === intervals.length - 1) {
                if (val >= intervals[i].lower && val <= intervals[i].upper) {
                    intervals[i].frequency++;
                    intervals[i].values.push(val);
                    break;
                }
            } else {
                if (val >= intervals[i].lower && val < intervals[i].upper) {
                    intervals[i].frequency++;
                    intervals[i].values.push(val);
                    break;
                }
            }
        }
    });

    let cumFreq = 0;
    intervals.forEach(interval => {
        cumFreq += interval.frequency;
        interval.cumFrequency = cumFreq;
        interval.relativeFreq = (interval.frequency / data.length * 100).toFixed(2);
        interval.cumRelativeFreq = (interval.cumFrequency / data.length * 100).toFixed(2);
    });

    displayFrequencyTable(intervals, true);
    
    const measures = calculateMeasuresGrouped(intervals, data.length);
    displayMeasures(measures);
    
    const quantiles = calculateQuantiles(data);
    displayQuantiles(quantiles);
    
    createCharts(intervals, data, true);
    
    processedData = { data, intervals, measures, quantiles, isGrouped: true };
}

// Mostrar tabla de frecuencias
function displayFrequencyTable(data, isGrouped) {
    const table = document.getElementById('frequencyTable');
    let html = '<thead><tr>';
    
    if (isGrouped) {
        html += '<th><i class="bi bi-columns"></i> Intervalo</th>';
        html += '<th><i class="bi bi-circle"></i> Marca de Clase</th>';
    } else {
        html += '<th><i class="bi bi-123"></i> Valor (Xi)</th>';
    }
    
    html += '<th><i class="bi bi-bar-chart"></i> Frecuencia (fi)</th>';
    html += '<th><i class="bi bi-arrow-up"></i> Frec. Acumulada (Fi)</th>';
    html += '<th><i class="bi bi-percent"></i> Frec. Relativa (%)</th>';
    html += '<th><i class="bi bi-graph-up"></i> Frec. Rel. Acum. (%)</th>';
    html += '</tr></thead><tbody>';

    data.forEach((item, index) => {
        html += `<tr style="animation-delay: ${index * 0.05}s">`;
        if (isGrouped) {
            html += `<td><strong>[${item.lower.toFixed(2)} - ${item.upper.toFixed(2)})</strong></td>`;
            html += `<td><span class="badge bg-primary">${item.midpoint.toFixed(2)}</span></td>`;
        } else {
            html += `<td><strong>${item.value}</strong></td>`;
        }
        html += `<td><span class="badge bg-success">${item.frequency}</span></td>`;
        html += `<td>${item.cumFrequency}</td>`;
        html += `<td>${item.relativeFreq}%</td>`;
        html += `<td><strong>${item.cumRelativeFreq}%</strong></td>`;
        html += '</tr>';
    });

    html += '</tbody>';
    table.innerHTML = html;
}

// Calcular medidas estadísticas para datos no agrupados
function calculateMeasures(data) {
    const n = data.length;
    const sorted = [...data].sort((a, b) => a - b);
    
    const mean = data.reduce((a, b) => a + b, 0) / n;
    
    let median;
    if (n % 2 === 0) {
        median = (sorted[n/2 - 1] + sorted[n/2]) / 2;
    } else {
        median = sorted[Math.floor(n/2)];
    }
    
    const freqMap = {};
    data.forEach(val => freqMap[val] = (freqMap[val] || 0) + 1);
    const maxFreq = Math.max(...Object.values(freqMap));
    const modes = Object.keys(freqMap).filter(key => freqMap[key] === maxFreq).map(Number);
    const mode = modes.length === n ? 'No hay moda' : modes.join(', ');
    
    const variance = data.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / n;
    const stdDev = Math.sqrt(variance);
    
    const range = Math.max(...data) - Math.min(...data);
    const cv = (stdDev / mean) * 100;
    
    return { mean, median, mode, variance, stdDev, range, cv, min: Math.min(...data), max: Math.max(...data) };
}

// Calcular medidas estadísticas para datos agrupados
function calculateMeasuresGrouped(intervals, n) {
    let sumFx = 0;
    intervals.forEach(int => {
        sumFx += int.midpoint * int.frequency;
    });
    const mean = sumFx / n;
    
    const halfN = n / 2;
    let medianInterval = intervals.find(int => int.cumFrequency >= halfN);
    const L = medianInterval.lower;
    const F = medianInterval === intervals[0] ? 0 : intervals[intervals.indexOf(medianInterval) - 1].cumFrequency;
    const fm = medianInterval.frequency;
    const c = medianInterval.upper - medianInterval.lower;
    const median = L + ((halfN - F) / fm) * c;
    
    const modalInterval = intervals.reduce((a, b) => a.frequency > b.frequency ? a : b);
    const L0 = modalInterval.lower;
    const idx = intervals.indexOf(modalInterval);
    const d1 = modalInterval.frequency - (idx > 0 ? intervals[idx - 1].frequency : 0);
    const d2 = modalInterval.frequency - (idx < intervals.length - 1 ? intervals[idx + 1].frequency : 0);
    const mode = L0 + (d1 / (d1 + d2)) * c;
    
    let sumFx2 = 0;
    intervals.forEach(int => {
        sumFx2 += Math.pow(int.midpoint - mean, 2) * int.frequency;
    });
    const variance = sumFx2 / n;
    const stdDev = Math.sqrt(variance);
    
    const range = intervals[intervals.length - 1].upper - intervals[0].lower;
    const cv = (stdDev / mean) * 100;
    
    return { 
        mean, 
        median, 
        mode: mode.toFixed(2), 
        variance, 
        stdDev, 
        range, 
        cv,
        min: intervals[0].lower,
        max: intervals[intervals.length - 1].upper
    };
}

// Mostrar medidas estadísticas
function displayMeasures(measures) {
    const centralHTML = `
        <div class="col-md-4 mb-3">
            <div class="stat-card">
                <div class="stat-label"><i class="bi bi-calculator"></i> Media</div>
                <div class="stat-value">${measures.mean.toFixed(2)}</div>
            </div>
        </div>
        <div class="col-md-4 mb-3">
            <div class="stat-card">
                <div class="stat-label"><i class="bi bi-dash-lg"></i> Mediana</div>
                <div class="stat-value">${measures.median.toFixed(2)}</div>
            </div>
        </div>
        <div class="col-md-4 mb-3">
            <div class="stat-card">
                <div class="stat-label"><i class="bi bi-trophy"></i> Moda</div>
                <div class="stat-value" style="font-size: 1.5rem;">${measures.mode}</div>
            </div>
        </div>
    `;
    
    const dispersionHTML = `
        <div class="col-md-3 mb-3">
            <div class="stat-card">
                <div class="stat-label"><i class="bi bi-arrows-expand"></i> Rango</div>
                <div class="stat-value">${measures.range.toFixed(2)}</div>
            </div>
        </div>
        <div class="col-md-3 mb-3">
            <div class="stat-card">
                <div class="stat-label"><i class="bi bi-grid"></i> Varianza</div>
                <div class="stat-value">${measures.variance.toFixed(2)}</div>
            </div>
        </div>
        <div class="col-md-3 mb-3">
            <div class="stat-card">
                <div class="stat-label"><i class="bi bi-distribute-vertical"></i> Desv. Estándar</div>
                <div class="stat-value">${measures.stdDev.toFixed(2)}</div>
            </div>
        </div>
        <div class="col-md-3 mb-3">
            <div class="stat-card">
                <div class="stat-label"><i class="bi bi-percent"></i> Coef. Variación</div>
                <div class="stat-value">${measures.cv.toFixed(2)}%</div>
            </div>
        </div>
    `;
    
    document.getElementById('centralMeasures').innerHTML = centralHTML;
    document.getElementById('dispersionMeasures').innerHTML = dispersionHTML;
}

// Calcular cuantiles
function calculateQuantiles(data) {
    const sorted = [...data].sort((a, b) => a - b);
    const n = sorted.length;
    
    function percentile(p) {
        const index = (p / 100) * (n - 1);
        const lower = Math.floor(index);
        const upper = Math.ceil(index);
        const fraction = index - lower;
        
        if (lower === upper) return sorted[lower];
        return sorted[lower] + fraction * (sorted[upper] - sorted[lower]);
    }
    
    return {
        Q1: percentile(25),
        Q2: percentile(50),
        Q3: percentile(75),
        D1: percentile(10),
        D5: percentile(50),
        D9: percentile(90),
        P10: percentile(10),
        P25: percentile(25),
        P50: percentile(50),
        P75: percentile(75),
        P90: percentile(90)
    };
}

// Mostrar cuantiles
function displayQuantiles(quantiles) {
    const html = `
        <div class="row">
            <div class="col-md-4 mb-3">
                <div class="quantile-item">
                    <div class="quantile-label"><i class="bi bi-1-square"></i> Cuartil 1 (Q1) - 25%</div>
                    <div class="quantile-value">${quantiles.Q1.toFixed(2)}</div>
                </div>
                <div class="quantile-item">
                    <div class="quantile-label"><i class="bi bi-2-square"></i> Cuartil 2 (Q2) - 50%</div>
                    <div class="quantile-value">${quantiles.Q2.toFixed(2)}</div>
                </div>
                <div class="quantile-item">
                    <div class="quantile-label"><i class="bi bi-3-square"></i> Cuartil 3 (Q3) - 75%</div>
                    <div class="quantile-value">${quantiles.Q3.toFixed(2)}</div>
                </div>
            </div>
            <div class="col-md-4 mb-3">
                <div class="quantile-item">
                    <div class="quantile-label"><i class="bi bi-trophy"></i> Decil 1 (D1) - 10%</div>
                    <div class="quantile-value">${quantiles.D1.toFixed(2)}</div>
                </div>
                <div class="quantile-item">
                    <div class="quantile-label"><i class="bi bi-trophy-fill"></i> Decil 5 (D5) - 50%</div>
                    <div class="quantile-value">${quantiles.D5.toFixed(2)}</div>
                </div>
                <div class="quantile-item">
                    <div class="quantile-label"><i class="bi bi-award"></i> Decil 9 (D9) - 90%</div>
                    <div class="quantile-value">${quantiles.D9.toFixed(2)}</div>
                </div>
            </div>
            <div class="col-md-4 mb-3">
                <div class="quantile-item">
                    <div class="quantile-label"><i class="bi bi-star"></i> Percentil 10 (P10)</div>
                    <div class="quantile-value">${quantiles.P10.toFixed(2)}</div>
                </div>
                <div class="quantile-item">
                    <div class="quantile-label"><i class="bi bi-star-fill"></i> Percentil 50 (P50)</div>
                    <div class="quantile-value">${quantiles.P50.toFixed(2)}</div>
                </div>
                <div class="quantile-item">
                    <div class="quantile-label"><i class="bi bi-stars"></i> Percentil 90 (P90)</div>
                    <div class="quantile-value">${quantiles.P90.toFixed(2)}</div>
                </div>
            </div>
        </div>
    `;
    document.getElementById('quantiles').innerHTML = html;
}

// Crear todos los gráficos
function createCharts(freqData, data, isGrouped) {
    Object.values(chartInstances).forEach(chart => chart.destroy());
    chartInstances = {};
    
    createHistogram(freqData, isGrouped);
    createPieChart(freqData, isGrouped);
    createBoxPlot(data);
    createScatterPlot(data);
}

// Crear histograma
function createHistogram(freqData, isGrouped) {
    const ctx = document.getElementById('histogram').getContext('2d');
    
    const labels = isGrouped 
        ? freqData.map(d => `[${d.lower.toFixed(1)}-${d.upper.toFixed(1)})`)
        : freqData.map(d => d.value);
    
    chartInstances.histogram = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: 'Frecuencia',
                data: freqData.map(d => d.frequency),
                backgroundColor: 'rgba(135, 206, 235, 0.7)',
                borderColor: 'rgba(0, 0, 0, 0.8)',
                borderWidth: 1,
                borderRadius: 0,
                barPercentage: 1.0,
                categoryPercentage: 1.0
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: { display: false },
                title: { 
                    display: true, 
                    text: 'Histograma de Frecuencias', 
                    font: { size: 16, weight: 'bold' },
                    color: '#000',
                    padding: { bottom: 20 }
                }
            },
            scales: {
                y: { 
                    beginAtZero: true,
                    title: { 
                        display: false
                    },
                    grid: { 
                        display: true,
                        color: 'rgba(0, 0, 0, 0.1)',
                        drawBorder: true
                    },
                    ticks: {
                        color: '#000',
                        font: { size: 12 }
                    },
                    border: {
                        display: true,
                        color: '#000',
                        width: 2
                    }
                },
                x: {
                    title: { display: false },
                    grid: { display: false },
                    ticks: {
                        color: '#000',
                        font: { size: 11 }
                    },
                    border: {
                        display: true,
                        color: '#000',
                        width: 2
                    }
                }
            },
            layout: {
                padding: {
                    left: 10,
                    right: 10,
                    top: 10,
                    bottom: 10
                }
            }
        }
    });
}

// Crear gráfico circular
function createPieChart(freqData, isGrouped) {
    const ctx = document.getElementById('pieChart').getContext('2d');
    
    const labels = isGrouped 
        ? freqData.map(d => `[${d.lower.toFixed(1)}-${d.upper.toFixed(1)})`)
        : freqData.map(d => String(d.value));
    
    const colors = [
        'rgba(102, 126, 234, 0.9)',
        'rgba(118, 75, 162, 0.9)',
        'rgba(17, 153, 142, 0.9)',
        'rgba(56, 239, 125, 0.9)',
        'rgba(240, 147, 251, 0.9)',
        'rgba(245, 87, 108, 0.9)',
        'rgba(250, 112, 154, 0.9)',
        'rgba(254, 225, 64, 0.9)',
        'rgba(79, 172, 254, 0.9)',
        'rgba(0, 242, 254, 0.9)'
    ];
    
    chartInstances.pieChart = new Chart(ctx, {
        type: 'pie',
        data: {
            labels: labels,
            datasets: [{
                data: freqData.map(d => d.frequency),
                backgroundColor: colors,
                borderColor: '#fff',
                borderWidth: 3
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: { 
                    position: 'right',
                    labels: { font: { size: 12, weight: 'bold' } }
                },
                title: { 
                    display: true, 
                    text: 'Distribución de Frecuencias', 
                    font: { size: 18, weight: 'bold' },
                    color: '#333'
                }
            }
        }
    });
}

// Crear diagrama de caja
function createBoxPlot(data) {
    const ctx = document.getElementById('boxPlot').getContext('2d');
    const quantiles = calculateQuantiles(data);
    const min = Math.min(...data);
    const max = Math.max(...data);
    
    chartInstances.boxPlot = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: ['Distribución de Datos'],
            datasets: [
                {
                    label: 'Mínimo',
                    data: [min],
                    backgroundColor: 'rgba(220, 53, 69, 0.8)',
                    borderColor: 'rgba(220, 53, 69, 1)',
                    borderWidth: 2,
                    borderRadius: 8
                },
                {
                    label: 'Q1 (25%)',
                    data: [quantiles.Q1],
                    backgroundColor: 'rgba(255, 193, 7, 0.8)',
                    borderColor: 'rgba(255, 193, 7, 1)',
                    borderWidth: 2,
                    borderRadius: 8
                },
                {
                    label: 'Mediana (50%)',
                    data: [quantiles.Q2],
                    backgroundColor: 'rgba(25, 135, 84, 0.8)',
                    borderColor: 'rgba(25, 135, 84, 1)',
                    borderWidth: 2,
                    borderRadius: 8
                },
                {
                    label: 'Q3 (75%)',
                    data: [quantiles.Q3],
                    backgroundColor: 'rgba(13, 202, 240, 0.8)',
                    borderColor: 'rgba(13, 202, 240, 1)',
                    borderWidth: 2,
                    borderRadius: 8
                },
                {
                    label: 'Máximo',
                    data: [max],
                    backgroundColor: 'rgba(108, 117, 125, 0.8)',
                    borderColor: 'rgba(108, 117, 125, 1)',
                    borderWidth: 2,
                    borderRadius: 8
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            indexAxis: 'y',
            plugins: {
                legend: { 
                    display: true, 
                    position: 'bottom',
                    labels: { font: { size: 12, weight: 'bold' } }
                },
                title: { 
                    display: true, 
                    text: 'Diagrama de Caja (Box Plot)', 
                    font: { size: 18, weight: 'bold' },
                    color: '#333'
                }
            },
            scales: {
                x: { 
                    stacked: false,
                    title: { display: true, text: 'Valores', font: { size: 14, weight: 'bold' } },
                    grid: { color: 'rgba(0, 0, 0, 0.05)' }
                }
            }
        }
    });
}

// Crear diagrama de dispersión
function createScatterPlot(data) {
    const ctx = document.getElementById('scatterPlot').getContext('2d');
    const scatterData = data.map((val, idx) => ({ x: idx + 1, y: val }));
    
    chartInstances.scatterPlot = new Chart(ctx, {
        type: 'scatter',
        data: {
            datasets: [{
                label: 'Datos',
                data: scatterData,
                backgroundColor: 'rgba(102, 126, 234, 0.7)',
                borderColor: 'rgba(102, 126, 234, 1)',
                borderWidth: 2,
                pointRadius: 7,
                pointHoverRadius: 10,
                pointHoverBackgroundColor: 'rgba(118, 75, 162, 0.9)',
                pointHoverBorderWidth: 3
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: { 
                    display: true,
                    labels: { font: { size: 14, weight: 'bold' } }
                },
                title: { 
                    display: true, 
                    text: 'Diagrama de Dispersión', 
                    font: { size: 18, weight: 'bold' },
                    color: '#333'
                }
            },
            scales: {
                x: { 
                    title: { display: true, text: 'Índice (Posición)', font: { size: 14, weight: 'bold' } },
                    beginAtZero: true,
                    grid: { color: 'rgba(0, 0, 0, 0.05)' }
                },
                y: { 
                    title: { display: true, text: 'Valor', font: { size: 14, weight: 'bold' } },
                    grid: { color: 'rgba(0, 0, 0, 0.05)' }
                }
            }
        }
    });
}

// Función para expandir gráficos
function expandChart(chartId) {
    const modal = new bootstrap.Modal(document.getElementById('chartModal'));
    const expandedCanvas = document.getElementById('expandedChart');
    const expandedCtx = expandedCanvas.getContext('2d');
    
    if (expandedChartInstance) {
        expandedChartInstance.destroy();
    }
    
    const originalChart = chartInstances[chartId];
    if (originalChart) {
        expandedChartInstance = new Chart(expandedCtx, {
            type: originalChart.config.type,
            data: originalChart.config.data,
            options: {
                ...originalChart.config.options,
                maintainAspectRatio: false
            }
        });
        
        modal.show();
    }
}

// Función para toggle expand de tablas
function toggleExpand(elementId) {
    const element = document.getElementById(elementId);
    element.classList.toggle('expanded');
    
    if (element.classList.contains('expanded')) {
        element.style.maxHeight = 'none';
    } else {
        element.style.maxHeight = '';
    }
}

// Exportar a Excel
function exportToExcel() {
    if (!processedData) {
        showNotification('No hay datos para exportar', 'warning');
        return;
    }

    try {
        const wb = XLSX.utils.book_new();
        
        // Hoja 1: Datos originales
        const ws1Data = [
            ['Datos Originales'],
            ['Total de datos:', processedData.data.length],
            [''],
            ['Índice', 'Valor']
        ];
        processedData.data.forEach((val, idx) => {
            ws1Data.push([idx + 1, val]);
        });
        const ws1 = XLSX.utils.aoa_to_sheet(ws1Data);
        XLSX.utils.book_append_sheet(wb, ws1, 'Datos Originales');
        
        // Hoja 2: Tabla de Frecuencias
        const ws2Data = [['Tabla de Frecuencias'], ['']];
        
        if (processedData.isGrouped) {
            ws2Data.push(['Intervalo', 'Marca de Clase', 'Frecuencia', 'Frec. Acumulada', 'Frec. Relativa (%)', 'Frec. Rel. Acum. (%)']);
            processedData.intervals.forEach(int => {
                ws2Data.push([
                    `[${int.lower.toFixed(2)} - ${int.upper.toFixed(2)})`,
                    int.midpoint.toFixed(2),
                    int.frequency,
                    int.cumFrequency,
                    int.relativeFreq,
                    int.cumRelativeFreq
                ]);
            });
        } else {
            ws2Data.push(['Valor', 'Frecuencia', 'Frec. Acumulada', 'Frec. Relativa (%)', 'Frec. Rel. Acum. (%)']);
            processedData.freqData.forEach(item => {
                ws2Data.push([
                    item.value,
                    item.frequency,
                    item.cumFrequency,
                    item.relativeFreq,
                    item.cumRelativeFreq
                ]);
            });
        }
        const ws2 = XLSX.utils.aoa_to_sheet(ws2Data);
        XLSX.utils.book_append_sheet(wb, ws2, 'Tabla de Frecuencias');
        
        // Hoja 3: Medidas Estadísticas
        const ws3Data = [
            ['Medidas Estadísticas'],
            [''],
            ['MEDIDAS DE TENDENCIA CENTRAL'],
            ['Media', processedData.measures.mean.toFixed(4)],
            ['Mediana', processedData.measures.median.toFixed(4)],
            ['Moda', processedData.measures.mode],
            [''],
            ['MEDIDAS DE DISPERSIÓN'],
            ['Rango', processedData.measures.range.toFixed(4)],
            ['Varianza', processedData.measures.variance.toFixed(4)],
            ['Desviación Estándar', processedData.measures.stdDev.toFixed(4)],
            ['Coeficiente de Variación (%)', processedData.measures.cv.toFixed(4)],
            ['Valor Mínimo', processedData.measures.min.toFixed(4)],
            ['Valor Máximo', processedData.measures.max.toFixed(4)]
        ];
        const ws3 = XLSX.utils.aoa_to_sheet(ws3Data);
        XLSX.utils.book_append_sheet(wb, ws3, 'Medidas Estadísticas');
        
        // Hoja 4: Cuantiles
        const ws4Data = [
            ['Cuantiles'],
            [''],
            ['CUARTILES'],
            ['Q1 (25%)', processedData.quantiles.Q1.toFixed(4)],
            ['Q2 (50%)', processedData.quantiles.Q2.toFixed(4)],
            ['Q3 (75%)', processedData.quantiles.Q3.toFixed(4)],
            [''],
            ['DECILES'],
            ['D1 (10%)', processedData.quantiles.D1.toFixed(4)],
            ['D5 (50%)', processedData.quantiles.D5.toFixed(4)],
            ['D9 (90%)', processedData.quantiles.D9.toFixed(4)],
            [''],
            ['PERCENTILES'],
            ['P10', processedData.quantiles.P10.toFixed(4)],
            ['P25', processedData.quantiles.P25.toFixed(4)],
            ['P50', processedData.quantiles.P50.toFixed(4)],
            ['P75', processedData.quantiles.P75.toFixed(4)],
            ['P90', processedData.quantiles.P90.toFixed(4)]
        ];
        const ws4 = XLSX.utils.aoa_to_sheet(ws4Data);
        XLSX.utils.book_append_sheet(wb, ws4, 'Cuantiles');
        
        // Generar archivo
        const fecha = new Date().toISOString().slice(0, 10);
        XLSX.writeFile(wb, `Analisis_Estadistico_${fecha}.xlsx`);
        
        showNotification('Archivo Excel exportado exitosamente', 'success');
        
    } catch (error) {
        showNotification('Error al exportar a Excel: ' + error.message, 'danger');
        console.error(error);
    }
}

// Notificaciones
function showNotification(message, type) {
    const alertDiv = document.createElement('div');
    alertDiv.className = `alert alert-${type} alert-dismissible fade show position-fixed top-0 start-50 translate-middle-x mt-3`;
    alertDiv.style.zIndex = '9999';
    alertDiv.style.minWidth = '300px';
    alertDiv.innerHTML = `
        ${message}
        <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
    `;
    
    document.body.appendChild(alertDiv);
    
    setTimeout(() => {
        alertDiv.classList.remove('show');
        setTimeout(() => alertDiv.remove(), 150);
    }, 3000);
}

// Limpiar todo
function clearAll() {
    document.getElementById('dataInput').value = '';
    document.getElementById('sortedData').innerHTML = `
        <div class="text-center py-4">
            <i class="bi bi-inbox fs-1 text-muted"></i>
            <p class="text-muted mt-2 mb-0">Los datos ordenados aparecerán aquí...</p>
        </div>
    `;
    document.getElementById('resultsSection').style.display = 'none';
    document.getElementById('exportExcelBtn').style.display = 'none';
    
    Object.values(chartInstances).forEach(chart => chart.destroy());
    chartInstances = {};
    processedData = null;
    
    showNotification('Todos los datos han sido limpiados', 'info');
}