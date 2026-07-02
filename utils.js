/* ============================================================
   utils.js · Funciones utilitarias
   BiblioAnalytics · Biblioteca Universitaria
   ============================================================ */

'use strict';

// ─── DATE UTILITIES ──────────────────────────────────────────

/**
 * Parsea una fecha desde múltiples formatos posibles.
 * Soporta: "DD/MM/YYYY", "YYYY-MM-DD", número serial Excel, Date object.
 */
function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;

  // Serial numérico de Excel
  if (typeof val === 'number') {
    const d = new Date(Math.round((val - 25569) * 86400 * 1000));
    return isNaN(d.getTime()) ? null : d;
  }

  const s = String(val).trim();

  // ISO o YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    const d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  }

  // DD/MM/YYYY o DD-MM-YYYY
  const m1 = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m1) {
    const [, dd, mm, yyyy] = m1;
    const yr = yyyy.length === 2 ? (parseInt(yyyy) > 50 ? 1900 : 2000) + parseInt(yyyy) : parseInt(yyyy);
    const d = new Date(yr, parseInt(mm) - 1, parseInt(dd));
    return isNaN(d.getTime()) ? null : d;
  }

  return null;
}

/** Años transcurridos desde una fecha hasta hoy */
function yearsSince(date) {
  if (!date) return null;
  const now = new Date();
  const diff = (now - date) / (1000 * 60 * 60 * 24 * 365.25);
  return Math.round(diff * 10) / 10;
}

/** Formatea fecha a DD/MM/YYYY */
function formatDate(date) {
  if (!date) return '—';
  const d = parseDate(date);
  if (!d) return String(date);
  return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;
}

/** Retorna el nombre del mes (ES) */
function monthName(date) {
  const months = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
  return months[date.getMonth()];
}

/** Clave YYYY-MM para agrupar por mes */
function monthKey(date) {
  return `${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,'0')}`;
}

// ─── NUMBER UTILITIES ─────────────────────────────────────────

function round2(n) { return Math.round(n * 100) / 100; }
function pct(n) { return `${round2(n)}%`; }
function num(n) { return Number.isFinite(n) ? round2(n) : '—'; }

// ─── STRING UTILITIES ─────────────────────────────────────────

/** Normaliza nombre de columna: minúsculas, sin acentos, sin espacios → guión */
function normalizeColName(s) {
  return String(s || '')
    .toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, '_')
    .replace(/[^\w]/g, '');
}

/** Capitaliza la primera letra */
function capitalize(s) { return s ? s.charAt(0).toUpperCase() + s.slice(1) : ''; }

// ─── DOM UTILITIES ────────────────────────────────────────────

function el(id) { return document.getElementById(id); }
function show(id) { const e = el(id); if (e) e.style.display = ''; }
function hide(id) { const e = el(id); if (e) e.style.display = 'none'; }

/** Muestra un toast de notificación */
function showToast(msg, type = 'default', duration = 3500) {
  const t = el('toast');
  t.textContent = msg;
  t.className = `toast ${type} show`;
  setTimeout(() => { t.className = 'toast'; }, duration);
}

/** Muestra/oculta el loader */
function showLoader(text = 'Procesando…') {
  el('loaderText').textContent = text;
  el('loaderOverlay').classList.add('active');
}
function hideLoader() { el('loaderOverlay').classList.remove('active'); }

// ─── STATUS HELPERS ────────────────────────────────────────────

/**
 * Retorna clase CSS de semáforo según umbrales.
 * thresholds: { green: n, yellow: n } — valores > green = verde, etc.
 * invertida = true para indicadores donde mayor es peor (obsolescencia, etc.)
 */
function getStatus(value, greenMax, yellowMax, inverted = false) {
  if (!Number.isFinite(value)) return 'blue';
  if (!inverted) {
    if (value >= greenMax) return 'green';
    if (value >= yellowMax) return 'yellow';
    return 'red';
  } else {
    if (value <= greenMax) return 'green';
    if (value <= yellowMax) return 'yellow';
    return 'red';
  }
}

function statusLabel(status) {
  return { green: 'Saludable', yellow: 'Revisar', red: 'Crítico', blue: 'Sin datos' }[status] || status;
}

function statusEmoji(status) {
  return { green: '🟢', yellow: '🟡', red: '🔴', blue: '⚪' }[status] || '⚪';
}

// ─── COLUMN NORMALIZATION MAP ──────────────────────────────────

const COL_ALIASES = {
  coleccion: {
    id: ['id','codigo','code','id_libro'],
    titulo: ['titulo','title','nombre','nombre_libro','nombre del libro'],
    unidad_academica: ['unidad_academica','unidad académica','unidad academica','escuela','unidad'],
    area: ['area','disciplina','materia','departamento','seccion'],
    ejemplares: ['ejemplares','copias','copies','cantidad','stock','num_ejemplares'],
    ultimo_prestamo: ['ultimo_prestamo','ultima_fecha','last_loan','ultimo_uso','fecha_ultimo','ultimoprestamo'],
    costo: ['costo','precio','cost','price','valor']
  },
  prestamos: {
    id_titulo: ['id_titulo','id_libro','titulo_id','book_id','idtitulo'],
    fecha: ['fecha','date','fecha_prestamo','loan_date'],
    carrera: ['carrera','facultad','programa','faculty','program','curso'],
    unidad_academica: ['unidad_academica','unidad académica','unidad academica','escuela','unidad']
  },
  bibliografia: {
    carrera: ['carrera','facultad','programa'],
    asignatura: ['asignatura','materia','course','subject','curso'],
    titulo: ['titulo','libro','title','nombre'],
    disponible: ['disponible','available','stock','en_catalogo']
  },
  reservas: {
    titulo: ['titulo','libro','title','nombre'],
    solicitudes: ['solicitudes','requests','total_solicitudes','pedidos'],
    satisfechas: ['satisfechas','atendidas','fulfilled','cubiertas']
  }
};

/**
 * Dado un objeto row y un mapa de aliases, retorna el nombre canónico de la columna
 * para cada columna original del row.
 */
function mapColumns(row, sheetName) {
  const aliases = COL_ALIASES[sheetName];
  if (!aliases) return row;
  const normalized = {};
  const rowKeys = Object.keys(row);

  for (const [canonical, aliasList] of Object.entries(aliases)) {
    for (const rk of rowKeys) {
      const norm = normalizeColName(rk);
      if (aliasList.includes(norm) || norm === canonical) {
        if (normalized[canonical] === undefined) {
          normalized[canonical] = row[rk];
        }
      }
    }
  }
  return normalized;
}

/**
 * Detecta columnas faltantes para una hoja.
 * Retorna array de nombres canónicos no encontrados.
 */
function detectMissingColumns(rows, sheetName) {
  if (!rows || rows.length === 0) return Object.keys(COL_ALIASES[sheetName] || {});
  const aliases = COL_ALIASES[sheetName];
  if (!aliases) return [];
  const rowKeys = Object.keys(rows[0]).map(normalizeColName);
  const missing = [];
  for (const [canonical, aliasList] of Object.entries(aliases)) {
    const found = aliasList.some(a => rowKeys.includes(a)) || rowKeys.includes(canonical);
    if (!found) missing.push(canonical);
  }
  return missing;
}

// ─── STATISTICS ───────────────────────────────────────────────

function groupBy(arr, keyFn) {
  return arr.reduce((acc, item) => {
    const k = keyFn(item);
    if (!acc[k]) acc[k] = [];
    acc[k].push(item);
    return acc;
  }, {});
}

function sumBy(arr, field) {
  return arr.reduce((s, r) => s + (parseFloat(r[field]) || 0), 0);
}

function countBy(arr, keyFn) {
  const g = groupBy(arr, keyFn);
  return Object.entries(g).map(([k, v]) => ({ key: k, count: v.length })).sort((a, b) => b.count - a.count);
}

// ─── EXCEL EXPORT ─────────────────────────────────────────────

function exportExcel() {
  if (!window.appData || !window.appData.indicators) {
    showToast('No hay datos para exportar. Carga datos primero.', 'error');
    return;
  }
  showLoader('Generando Excel…');

  try {
    const wb = XLSX.utils.book_new();

    // Hoja 1: Indicadores
    const indRows = window.appData.indicators.map(i => ({
      Indicador: i.name,
      Valor: i.value,
      Unidad: i.unit,
      Estado: statusLabel(i.status),
      Referencia: i.reference
    }));
    const ws1 = XLSX.utils.json_to_sheet(indRows);
    XLSX.utils.book_append_sheet(wb, ws1, 'Indicadores');

    // Hoja 2: Colección procesada
    if (window.appData.coleccion && window.appData.coleccion.length > 0) {
      const ws2 = XLSX.utils.json_to_sheet(window.appData.coleccion.map(r => ({
        ID: r.id,
        Título: r.titulo,
        Área: r.area,
        Ejemplares: r.ejemplares,
        'Último préstamo': formatDate(r.ultimo_prestamo),
        'Costo ($)': r.costo,
        'Años sin uso': r._yearsUnused ?? '—',
        'Préstamos totales': r._loans ?? 0,
        'Estado obsolescencia': r._obsStatus ?? '—'
      })));
      XLSX.utils.book_append_sheet(wb, ws2, 'Coleccion');
    }

    // Hoja 3: Obsolescencia
    if (window.appData.obsolescence && window.appData.obsolescence.length > 0) {
      const ws3 = XLSX.utils.json_to_sheet(window.appData.obsolescence.map(r => ({
        Título: r.titulo,
        Área: r.area,
        Ejemplares: r.ejemplares,
        'Último uso': formatDate(r.ultimo_prestamo),
        'Años sin uso': r._yearsUnused ?? '—',
        'Préstamos': r._loans ?? 0,
        'Índice': round2(r._obsIndex ?? 0),
        Estado: r._obsStatus
      })));
      XLSX.utils.book_append_sheet(wb, ws3, 'Obsolescencia');
    }

    XLSX.writeFile(wb, 'biblioteca-analytics.xlsx');
    showToast('Excel exportado correctamente', 'success');
  } catch(e) {
    showToast('Error al exportar Excel: ' + e.message, 'error');
    console.error(e);
  } finally {
    hideLoader();
  }
}

// ─── IMAGE EXPORT ─────────────────────────────────────────────

async function exportImage() {
  showToast('Esta función requiere html2canvas. Usa la exportación PDF en su lugar.', 'warning');
}

// ─── PDF EXPORT ───────────────────────────────────────────────

function exportPDF() {
  if (!window.appData || !window.appData.indicators) {
    showToast('No hay datos para exportar. Carga datos primero.', 'error');
    return;
  }
  showLoader('Generando PDF…');

  try {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });

    const NAVY = [26, 39, 68];
    const AMBER = [212, 168, 67];
    const GRAY = [90, 90, 114];
    const pageW = doc.internal.pageSize.getWidth();
    const pageH = doc.internal.pageSize.getHeight();
    let y = 0;

    // ── PORTADA ──
    doc.setFillColor(...NAVY);
    doc.rect(0, 0, pageW, 60, 'F');

    doc.setFillColor(...AMBER);
    doc.rect(0, 57, pageW, 3, 'F');

    doc.setTextColor(255, 255, 255);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(22);
    doc.text('Análisis de Uso de Colecciones', pageW / 2, 25, { align: 'center' });

    doc.setFontSize(13);
    doc.setFont('helvetica', 'normal');
    doc.text('Biblioteca Universitaria · BiblioAnalytics', pageW / 2, 36, { align: 'center' });

    doc.setFontSize(10);
    doc.setTextColor(...GRAY);
    doc.text(`Generado: ${new Date().toLocaleDateString('es-CL')}`, pageW / 2, 46, { align: 'center' });

    y = 75;

    // ── RESUMEN EJECUTIVO ──
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(14);
    doc.setTextColor(...NAVY);
    doc.text('Resumen Ejecutivo', 14, y);
    y += 8;

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(10);
    doc.setTextColor(...GRAY);

    const summary = window.appData.summaryText || buildSummaryText();
    const lines = doc.splitTextToSize(summary, pageW - 28);
    doc.text(lines, 14, y);
    y += lines.length * 5 + 10;

    // ── TABLA INDICADORES ──
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(14);
    doc.setTextColor(...NAVY);
    doc.text('Indicadores de Uso', 14, y);
    y += 6;

    const tableData = (window.appData.indicators || []).map(i => [
      i.name,
      `${num(i.value)} ${i.unit}`,
      i.reference,
      statusLabel(i.status)
    ]);

    doc.autoTable({
      head: [['Indicador', 'Valor', 'Referencia', 'Estado']],
      body: tableData,
      startY: y,
      theme: 'grid',
      headStyles: { fillColor: NAVY, textColor: [255,255,255], fontStyle: 'bold', fontSize: 9 },
      bodyStyles: { fontSize: 9 },
      alternateRowStyles: { fillColor: [248, 247, 244] },
      columnStyles: {
        0: { cellWidth: 65 },
        1: { cellWidth: 35, halign: 'center' },
        2: { cellWidth: 50 },
        3: { cellWidth: 30, halign: 'center' }
      },
      didParseCell: function(data) {
        if (data.column.index === 3 && data.section === 'body') {
          const txt = data.cell.raw;
          if (txt === 'Saludable') data.cell.styles.textColor = [22, 163, 74];
          else if (txt === 'Revisar') data.cell.styles.textColor = [202, 138, 4];
          else if (txt === 'Crítico') data.cell.styles.textColor = [220, 38, 38];
        }
      }
    });

    y = doc.lastAutoTable.finalY + 12;

    // ── OBSOLESCENCIA ──
    if (window.appData.obsolescence && window.appData.obsolescence.length > 0) {
      if (y > pageH - 60) { doc.addPage(); y = 20; }

      doc.setFont('helvetica', 'bold');
      doc.setFontSize(14);
      doc.setTextColor(...NAVY);
      doc.text('Candidatos a Descarte (Top 20)', 14, y);
      y += 6;

      const obsData = window.appData.obsolescence.slice(0, 20).map((r, i) => [
        i + 1,
        (r.titulo || '').substring(0, 40),
        r.area || '—',
        r._yearsUnused != null ? `${r._yearsUnused} años` : '—',
        r._loans ?? 0,
        r._obsStatus || '—'
      ]);

      doc.autoTable({
        head: [['#', 'Título', 'Área', 'Sin uso', 'Préstamos', 'Estado']],
        body: obsData,
        startY: y,
        theme: 'striped',
        headStyles: { fillColor: NAVY, textColor: [255,255,255], fontStyle: 'bold', fontSize: 9 },
        bodyStyles: { fontSize: 8 },
        columnStyles: { 0: { cellWidth: 10 }, 1: { cellWidth: 60 }, 2: { cellWidth: 30 } }
      });

      y = doc.lastAutoTable.finalY + 12;
    }

    // ── RECOMENDACIONES ──
    if (window.appData.recommendations && window.appData.recommendations.length > 0) {
      if (y > pageH - 60) { doc.addPage(); y = 20; }

      doc.setFont('helvetica', 'bold');
      doc.setFontSize(14);
      doc.setTextColor(...NAVY);
      doc.text('Recomendaciones', 14, y);
      y += 8;

      for (const rec of window.appData.recommendations) {
        if (y > pageH - 25) { doc.addPage(); y = 20; }

        const color = { high: [220,38,38], medium: [202,138,4], low: [22,163,74], info: [37,99,235] }[rec.priority] || GRAY;
        doc.setFillColor(...color);
        doc.rect(14, y - 3, 3, 8, 'F');

        doc.setFont('helvetica', 'bold');
        doc.setFontSize(10);
        doc.setTextColor(...color);
        doc.text(rec.area || 'General', 20, y + 1);

        doc.setFont('helvetica', 'normal');
        doc.setFontSize(9.5);
        doc.setTextColor(30, 30, 50);
        const lines = doc.splitTextToSize(rec.text, pageW - 34);
        doc.text(lines, 20, y + 7);
        y += lines.length * 5 + 14;
      }
    }

    // ── PIE DE PÁGINA ──
    const totalPages = doc.internal.getNumberOfPages();
    for (let i = 1; i <= totalPages; i++) {
      doc.setPage(i);
      doc.setFillColor(...NAVY);
      doc.rect(0, pageH - 12, pageW, 12, 'F');
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(8);
      doc.setTextColor(255, 255, 255);
      doc.text('BiblioAnalytics · Análisis de Colecciones Universitarias', 14, pageH - 4.5);
      doc.text(`Página ${i} de ${totalPages}`, pageW - 14, pageH - 4.5, { align: 'right' });
    }

    doc.save('informe-colecciones.pdf');
    showToast('PDF generado correctamente', 'success');
  } catch(e) {
    console.error(e);
    showToast('Error al generar PDF: ' + e.message, 'error');
  } finally {
    hideLoader();
  }
}

function buildSummaryText() {
  const d = window.appData;
  if (!d || !d.indicators) return 'Sin datos disponibles.';
  const tc = d.indicators.find(i => i.id === 'tasa_circulacion');
  const obs = d.indicators.find(i => i.id === 'obsolescencia');
  const cob = d.indicators.find(i => i.id === 'cobertura');
  const parts = [];
  if (d.coleccion) parts.push(`La colección analizada cuenta con ${d.coleccion.length} títulos.`);
  if (tc) parts.push(`La tasa de circulación es ${num(tc.value)}, indicando ${statusLabel(tc.status).toLowerCase()} uso.`);
  if (obs) parts.push(`El índice de obsolescencia alcanza ${num(obs.value)}%, con ${statusLabel(obs.status).toLowerCase()} nivel de renovación.`);
  if (cob) parts.push(`La cobertura bibliográfica es del ${num(cob.value)}%.`);
  return parts.join(' ');
}

/** Descarga la plantilla Excel vacía */
function downloadTemplate() {
  showLoader('Generando plantilla…');
  try {
    const wb = XLSX.utils.book_new();

    const sheets = {
      Coleccion: [{ id: 'LIB001', titulo: 'Ejemplo: Introducción a la Investigación', area: 'Ciencias', ejemplares: 3, ultimo_prestamo: '01/01/2022', costo: 25000 }],
      Prestamos: [{ id_titulo: 'LIB001', fecha: '15/03/2024', carrera: 'Ingeniería Civil' }],
      Bibliografia: [{ carrera: 'Ingeniería Civil', asignatura: 'Metodología', titulo: 'Ejemplo: Introducción a la Investigación', disponible: 'Sí' }],
      Reservas: [{ titulo: 'Ejemplo: Introducción a la Investigación', solicitudes: 5, satisfechas: 3 }]
    };

    for (const [name, data] of Object.entries(sheets)) {
      const ws = XLSX.utils.json_to_sheet(data);
      XLSX.utils.book_append_sheet(wb, ws, name);
    }

    XLSX.writeFile(wb, 'plantilla-biblioteca.xlsx');
    showToast('Plantilla descargada correctamente', 'success');
  } catch(e) {
    showToast('Error al generar plantilla: ' + e.message, 'error');
  } finally {
    hideLoader();
  }
}
