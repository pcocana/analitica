/* ============================================================
   charts.js · Gestión de gráficos con Chart.js
   BiblioAnalytics · Biblioteca Universitaria
   ============================================================ */

'use strict';

// ─── PALETTE ─────────────────────────────────────────────────

const CHART_PALETTE = {
  navy:   '#1a2744',
  amber:  '#d4a843',
  blue:   '#3b82f6',
  green:  '#16a34a',
  red:    '#dc2626',
  purple: '#7c3aed',
  teal:   '#0d9488',
  orange: '#ea580c',
  pink:   '#db2777',
  indigo: '#4f46e5',
};

const MULTI_COLORS = [
  '#1a2744','#d4a843','#3b82f6','#16a34a','#dc2626',
  '#7c3aed','#0d9488','#ea580c','#db2777','#4f46e5',
  '#78716c','#65a30d','#0284c7','#c2410c','#7e22ce'
];

function withAlpha(hex, alpha) {
  const r = parseInt(hex.slice(1,3),16);
  const g = parseInt(hex.slice(3,5),16);
  const b = parseInt(hex.slice(5,7),16);
  return `rgba(${r},${g},${b},${alpha})`;
}

// ─── GLOBAL DEFAULTS ──────────────────────────────────────────

Chart.defaults.font.family = "'Source Sans 3', sans-serif";
Chart.defaults.font.size = 12;
Chart.defaults.color = '#5a5a72';
Chart.defaults.plugins.legend.labels.boxWidth = 12;
Chart.defaults.plugins.legend.labels.padding = 16;
Chart.defaults.plugins.tooltip.backgroundColor = '#1a2744';
Chart.defaults.plugins.tooltip.titleColor = '#ffffff';
Chart.defaults.plugins.tooltip.bodyColor = 'rgba(255,255,255,0.8)';
Chart.defaults.plugins.tooltip.padding = 10;
Chart.defaults.plugins.tooltip.cornerRadius = 8;
Chart.defaults.plugins.tooltip.displayColors = true;
Chart.defaults.plugins.tooltip.boxWidth = 10;
Chart.defaults.plugins.tooltip.boxHeight = 10;

// ─── CHART REGISTRY ──────────────────────────────────────────

const _charts = {};

function destroyChart(id) {
  if (_charts[id]) {
    _charts[id].destroy();
    delete _charts[id];
  }
}

function registerChart(id, chart) {
  destroyChart(id);
  _charts[id] = chart;
}

// ─── PRÉSTAMOS POR CARRERA ────────────────────────────────────

function renderChartCarrera(data) {
  destroyChart('chartCarrera');
  const ctx = document.getElementById('chartCarrera');
  if (!ctx || !data || data.length === 0) return;

  const sorted = [...data].sort((a, b) => b.count - a.count).slice(0, 12);
  const labels = sorted.map(d => d.key);
  const values = sorted.map(d => d.count);
  const colors = labels.map((_, i) => MULTI_COLORS[i % MULTI_COLORS.length]);

  registerChart('chartCarrera', new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Préstamos',
        data: values,
        backgroundColor: colors.map(c => withAlpha(c, 0.85)),
        borderColor: colors,
        borderWidth: 1.5,
        borderRadius: 6,
        borderSkipped: false,
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: ctx => ` ${ctx.parsed.y} préstamos`
          }
        }
      },
      scales: {
        x: {
          grid: { display: false },
          ticks: { maxRotation: 35, minRotation: 0, font: { size: 11 } }
        },
        y: {
          beginAtZero: true,
          grid: { color: 'rgba(0,0,0,0.05)' },
          ticks: { precision: 0 }
        }
      }
    }
  }));
}

// ─── COLECCIÓN USADA VS NO USADA ──────────────────────────────

function renderChartUsage(usados, noUsados) {
  destroyChart('chartUsage');
  const ctx = document.getElementById('chartUsage');
  if (!ctx) return;

  registerChart('chartUsage', new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: ['Colección usada', 'Sin circulación'],
      datasets: [{
        data: [usados, noUsados],
        backgroundColor: [withAlpha(CHART_PALETTE.green, 0.85), withAlpha(CHART_PALETTE.red, 0.75)],
        borderColor: [CHART_PALETTE.green, CHART_PALETTE.red],
        borderWidth: 2,
        hoverOffset: 8
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: '65%',
      plugins: {
        legend: {
          position: 'bottom',
          labels: { padding: 20 }
        },
        tooltip: {
          callbacks: {
            label: ctx => {
              const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
              const pct = total > 0 ? ((ctx.parsed / total) * 100).toFixed(1) : 0;
              return ` ${ctx.parsed} títulos (${pct}%)`;
            }
          }
        }
      }
    }
  }));
}

// ─── PRÉSTAMOS MENSUALES ──────────────────────────────────────

function renderChartMonthly(monthlyData) {
  destroyChart('chartMonthly');
  const ctx = document.getElementById('chartMonthly');
  if (!ctx || !monthlyData || monthlyData.length === 0) return;

  const sorted = [...monthlyData].sort((a, b) => a.key.localeCompare(b.key)).slice(-24);
  const labels = sorted.map(d => {
    const [y, m] = d.key.split('-');
    const months = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
    return `${months[parseInt(m)-1]} ${y}`;
  });
  const values = sorted.map(d => d.count);

  registerChart('chartMonthly', new Chart(ctx, {
    type: 'line',
    data: {
      labels,
      datasets: [{
        label: 'Préstamos',
        data: values,
        borderColor: CHART_PALETTE.amber,
        backgroundColor: withAlpha(CHART_PALETTE.amber, 0.1),
        borderWidth: 2.5,
        pointBackgroundColor: CHART_PALETTE.amber,
        pointRadius: 4,
        pointHoverRadius: 6,
        tension: 0.35,
        fill: true
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: { label: ctx => ` ${ctx.parsed.y} préstamos` }
        }
      },
      scales: {
        x: {
          grid: { display: false },
          ticks: { maxTicksLimit: 8, font: { size: 10 } }
        },
        y: {
          beginAtZero: true,
          grid: { color: 'rgba(0,0,0,0.05)' },
          ticks: { precision: 0 }
        }
      }
    }
  }));
}

// ─── USO POR ÁREA ─────────────────────────────────────────────

function renderChartArea(areaData) {
  destroyChart('chartArea');
  const ctx = document.getElementById('chartArea');
  if (!ctx || !areaData || areaData.length === 0) return;

  const sorted = [...areaData].sort((a, b) => b.ratio - a.ratio).slice(0, 10);
  const labels = sorted.map(d => d.area);
  const values = sorted.map(d => round2(d.ratio));
  const colors = labels.map((_, i) => MULTI_COLORS[i % MULTI_COLORS.length]);

  registerChart('chartArea', new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Índice de uso',
        data: values,
        backgroundColor: colors.map(c => withAlpha(c, 0.8)),
        borderColor: colors,
        borderWidth: 1.5,
        borderRadius: 6,
        borderSkipped: false,
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: { label: ctx => ` Índice: ${ctx.parsed.y}` }
        }
      },
      scales: {
        x: {
          grid: { display: false },
          ticks: { maxRotation: 35, font: { size: 10 } }
        },
        y: {
          beginAtZero: true,
          grid: { color: 'rgba(0,0,0,0.05)' }
        }
      }
    }
  }));
}

// ─── RANKING MENOR USO ────────────────────────────────────────

function renderChartRanking(rankingData) {
  destroyChart('chartRanking');
  const ctx = document.getElementById('chartRanking');
  if (!ctx || !rankingData || rankingData.length === 0) return;

  const top = rankingData.slice(0, 15);
  const labels = top.map(d => d.titulo.length > 45 ? d.titulo.substring(0, 42) + '…' : d.titulo);
  const values = top.map(d => d._loans ?? 0);

  // Color según estado
  const colors = top.map(d => {
    if (d._obsStatus === 'Crítico') return withAlpha(CHART_PALETTE.red, 0.8);
    if (d._obsStatus === 'Revisar') return withAlpha(CHART_PALETTE.amber, 0.8);
    return withAlpha(CHART_PALETTE.blue, 0.7);
  });

  registerChart('chartRanking', new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Préstamos totales',
        data: values,
        backgroundColor: colors,
        borderRadius: 4,
        borderSkipped: false,
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: { label: ctx => ` ${ctx.parsed.x} préstamos registrados` }
        }
      },
      scales: {
        x: {
          beginAtZero: true,
          grid: { color: 'rgba(0,0,0,0.05)' },
          ticks: { precision: 0 }
        },
        y: {
          grid: { display: false },
          ticks: { font: { size: 10 } }
        }
      }
    }
  }));
}

// ─── RENDER ALL CHARTS ────────────────────────────────────────

function renderAllCharts(data) {
  if (!data) return;

  // Préstamos por carrera
  if (data.prestamos && data.prestamos.length > 0) {
    const byCarrera = countBy(data.prestamos, r => r.carrera || 'Sin especificar');
    renderChartCarrera(byCarrera);

    // Mensuales
    const byMonth = countBy(
      data.prestamos.filter(r => r._fechaDate),
      r => monthKey(r._fechaDate)
    );
    renderChartMonthly(byMonth);
  }

  // Usada vs no usada
  const usados = (data.coleccion || []).filter(r => (r._loans ?? 0) > 0).length;
  const noUsados = (data.coleccion || []).length - usados;
  renderChartUsage(usados, noUsados);

  // Uso por área
  if (data.areaMetrics && data.areaMetrics.length > 0) {
    renderChartArea(data.areaMetrics);
  }

  // Ranking menor uso
  if (data.coleccion && data.coleccion.length > 0) {
    const sorted = [...data.coleccion].sort((a, b) => (a._loans ?? 0) - (b._loans ?? 0));
    renderChartRanking(sorted);
  }
}
