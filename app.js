/* ============================================================
   app.js · Lógica principal de la aplicación
   BiblioAnalytics · Biblioteca Universitaria
   ============================================================ */

'use strict';

// ─── ESTADO GLOBAL ────────────────────────────────────────────

window.appData = null; // Datos procesados de la sesión actual

// ─── NAVEGACIÓN ───────────────────────────────────────────────

document.querySelectorAll('.nav-item').forEach(item => {
  item.addEventListener('click', e => {
    e.preventDefault();
    const section = item.dataset.section;
    navigateTo(section);
  });
});

function navigateTo(section) {
  // Actualizar nav items
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  document.querySelector(`.nav-item[data-section="${section}"]`)?.classList.add('active');

  // Actualizar sections
  document.querySelectorAll('.section').forEach(s => s.classList.remove('active'));
  el(`section-${section}`)?.classList.add('active');

  // Cerrar sidebar en mobile
  el('sidebar').classList.remove('open');

  // Renderizar gráficos cuando se navega a esa sección
  if (section === 'charts' && window.appData) {
    setTimeout(() => renderAllCharts(window.appData), 100);
  }
}

function toggleSidebar() {
  el('sidebar').classList.toggle('open');
}

// ─── SHEET TABS ───────────────────────────────────────────────

document.querySelectorAll('.sheet-tab').forEach(tab => {
  tab.addEventListener('click', () => {
    const t = tab.dataset.tab;
    document.querySelectorAll('.sheet-tab').forEach(x => x.classList.remove('active'));
    document.querySelectorAll('.sheet-content').forEach(x => x.classList.remove('active'));
    tab.classList.add('active');
    el(`tab-${t}`)?.classList.add('active');
  });
});

// ─── DRAG & DROP ──────────────────────────────────────────────

function handleDragOver(e) {
  e.preventDefault();
  el('dragZone').classList.add('dragover');
}

function handleDragLeave(e) {
  el('dragZone').classList.remove('dragover');
}

function handleDrop(e) {
  e.preventDefault();
  el('dragZone').classList.remove('dragover');
  const file = e.dataTransfer.files[0];
  if (file) processFile(file);
}

function handleFileSelect(e) {
  const file = e.target.files[0];
  if (file) processFile(file);
}

// ─── FILE PROCESSING ──────────────────────────────────────────

async function processFile(file) {
  if (!file.name.match(/\.(xlsx|xls)$/i)) {
    showToast('Solo se aceptan archivos .xlsx o .xls', 'error');
    return;
  }

  showLoader(`Leyendo "${file.name}"…`);

  try {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });

    const sheets = {};
    const sheetNames = ['Coleccion', 'Prestamos', 'Bibliografia', 'Reservas'];

    for (const name of workbook.SheetNames) {
      const normalized = sheetNames.find(s => s.toLowerCase() === name.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'')) || name;
      const ws = workbook.Sheets[name];
      sheets[normalized.toLowerCase()] = XLSX.utils.sheet_to_json(ws, { defval: '' });
    }

    hideLoader();
    validateAndProcess(sheets, file.name);

  } catch (err) {
    hideLoader();
    showValidationError(`Error al leer el archivo: ${err.message}`);
    console.error(err);
  }
}

// ─── VALIDATION ───────────────────────────────────────────────

function validateAndProcess(sheets, filename = 'archivo') {
  const panel = el('validationPanel');
  const messages = el('valMessages');
  panel.style.display = 'block';
  messages.innerHTML = '';

  const msgs = [];
  let hasErrors = false;

  // Verificar hojas presentes
  const requiredSheets = ['coleccion', 'prestamos'];
  const optionalSheets = ['bibliografia', 'reservas'];

  for (const s of requiredSheets) {
    if (!sheets[s] || sheets[s].length === 0) {
      msgs.push({ type: 'error', text: `Hoja requerida no encontrada: "${capitalize(s)}"` });
      hasErrors = true;
    } else {
      msgs.push({ type: 'ok', text: `✓ Hoja "${capitalize(s)}" encontrada (${sheets[s].length} filas)` });
    }
  }

  for (const s of optionalSheets) {
    if (!sheets[s] || sheets[s].length === 0) {
      msgs.push({ type: 'warn', text: `Hoja opcional no encontrada: "${capitalize(s)}" — indicadores relacionados no se calcularán` });
    } else {
      msgs.push({ type: 'ok', text: `✓ Hoja "${capitalize(s)}" encontrada (${sheets[s].length} filas)` });
    }
  }

  if (hasErrors) {
    el('valIcon').textContent = '❌';
    el('valTitle').textContent = 'Se encontraron errores críticos';
    renderMessages(messages, msgs);
    return;
  }

  // Normalizar columnas
  const normalized = {};
  for (const [name, rows] of Object.entries(sheets)) {
    normalized[name] = rows.map(row => mapColumns(row, name));
  }

  // Detectar columnas faltantes por hoja
  for (const s of [...requiredSheets, ...optionalSheets]) {
    if (!normalized[s]) continue;
    const missing = detectMissingColumns(normalized[s], s);
    if (missing.length > 0) {
      msgs.push({ type: 'warn', text: `Columnas no reconocidas en "${capitalize(s)}": ${missing.join(', ')} — se usarán valores por defecto` });
    }
  }

  // Validar fechas en préstamos
  let invalidDates = 0;
  if (normalized.prestamos) {
    normalized.prestamos.forEach(row => {
      if (row.fecha) {
        const d = parseDate(row.fecha);
        if (!d) invalidDates++;
        else row._fechaDate = d;
      }
    });
  }
  if (invalidDates > 0) {
    msgs.push({ type: 'warn', text: `${invalidDates} fechas con formato no reconocido en Préstamos — esas filas se excluirán del análisis temporal` });
  } else if (normalized.prestamos && normalized.prestamos.length > 0) {
    msgs.push({ type: 'ok', text: `✓ Fechas en Préstamos validadas correctamente` });
  }

  // Validar ejemplares en colección
  let missingEjemplares = 0;
  if (normalized.coleccion) {
    normalized.coleccion.forEach(row => {
      if (!row.ejemplares || isNaN(Number(row.ejemplares))) {
        row.ejemplares = 1;
        missingEjemplares++;
      } else {
        row.ejemplares = parseInt(row.ejemplares) || 1;
      }
      if (row.costo) row.costo = parseFloat(row.costo) || 0;
      if (row.ultimo_prestamo) row._ultimaFechaDate = parseDate(row.ultimo_prestamo);
    });
  }
  if (missingEjemplares > 0) {
    msgs.push({ type: 'warn', text: `${missingEjemplares} filas sin dato de ejemplares — se asumió valor 1` });
  }

  el('valIcon').textContent = msgs.some(m => m.type === 'error') ? '❌' : (msgs.some(m => m.type === 'warn') ? '⚠️' : '✅');
  el('valTitle').textContent = msgs.some(m => m.type === 'error') ? 'Error en el archivo' : `Archivo procesado: ${filename}`;
  renderMessages(messages, msgs);

  if (!hasErrors) {
    el('btnContinue').style.display = 'block';
    // Procesar y calcular indicadores
    processData(normalized);
    showToast('Datos cargados correctamente', 'success');
  }
}

function renderMessages(container, msgs) {
  container.innerHTML = msgs.map(m => `
    <div class="val-msg ${m.type}">
      <span class="val-msg-icon">${m.type === 'ok' ? '✓' : m.type === 'warn' ? '⚠' : '✗'}</span>
      <span>${m.text}</span>
    </div>
  `).join('');
}

function showValidationError(msg) {
  const panel = el('validationPanel');
  panel.style.display = 'block';
  el('valIcon').textContent = '❌';
  el('valTitle').textContent = 'Error';
  el('valMessages').innerHTML = `<div class="val-msg error"><span>✗</span><span>${msg}</span></div>`;
}

// ─── DATA PROCESSING ──────────────────────────────────────────

function processData(sheets) {
  const coleccion = sheets.coleccion || [];
  const prestamos = sheets.prestamos || [];
  const bibliografia = sheets.bibliografia || [];
  const reservas = sheets.reservas || [];

  // Construir mapa de préstamos por id_titulo
  const loansByTitle = {};
  prestamos.forEach(p => {
    const id = String(p.id_titulo || '').trim();
    if (!loansByTitle[id]) loansByTitle[id] = [];
    loansByTitle[id].push(p);
  });

  // Enriquecer colección
  coleccion.forEach(item => {
    const id = String(item.id || '').trim();
    const loans = loansByTitle[id] || [];
    item._loans = loans.length;
    item._yearsUnused = item._ultimaFechaDate ? yearsSince(item._ultimaFechaDate) : null;
    if (item._yearsUnused === null && loans.length === 0) item._yearsUnused = 99;
  });

  // Calcular indicadores
  const indicators = calculateIndicators(coleccion, prestamos, bibliografia, reservas);

  // Métricas por área
  const areaMetrics = calculateAreaMetrics(coleccion, prestamos);

  // Obsolescencia
  const obsolescence = calculateObsolescence(coleccion);

  // Recomendaciones
  const recommendations = generateRecommendations(indicators, coleccion, prestamos, areaMetrics, obsolescence);

  // Guardar en estado global
  window.appData = {
    coleccion, prestamos, bibliografia, reservas,
    indicators, areaMetrics, obsolescence, recommendations,
    summaryText: buildSummaryText()
  };

  // Renderizar UI
  renderDashboard(indicators, coleccion, areaMetrics);
  renderObsolescence(obsolescence);
  renderRecommendations(recommendations);
  populateFilters(coleccion, prestamos);
}

// ─── INDICATORS ───────────────────────────────────────────────

function calculateIndicators(coleccion, prestamos, bibliografia, reservas) {
  const totalEjemplares = coleccion.reduce((s, r) => s + (r.ejemplares || 1), 0);
  const totalTitulos = coleccion.length;
  const totalPrestamos = prestamos.length;
  const titulosUsados = coleccion.filter(r => (r._loans || 0) > 0).length;
  const titulosSinUso = totalTitulos - titulosUsados;
  const costoTotal = coleccion.reduce((s, r) => s + (parseFloat(r.costo) || 0), 0);

  // A) Tasa de circulación
  const tasaCirculacion = totalEjemplares > 0 ? totalPrestamos / totalEjemplares : 0;

  // B) Porcentaje de colección usada
  const pctUsada = totalTitulos > 0 ? (titulosUsados / totalTitulos) * 100 : 0;

  // C) Índice de rotación
  const indiceRotacion = totalEjemplares > 0 ? totalPrestamos / totalEjemplares : 0;

  // E) Costo por uso
  const costoUso = totalPrestamos > 0 ? costoTotal / totalPrestamos : 0;

  // G) Obsolescencia
  const obsolescencia = totalTitulos > 0 ? (titulosSinUso / totalTitulos) * 100 : 0;

  // F) Antigüedad sin uso promedio (solo títulos sin uso con fecha)
  const sinUsoConFecha = coleccion.filter(r => r._loans === 0 && r._yearsUnused != null && r._yearsUnused < 99);
  const antiguedadPromedio = sinUsoConFecha.length > 0
    ? sinUsoConFecha.reduce((s, r) => s + r._yearsUnused, 0) / sinUsoConFecha.length
    : 0;

  // H) Cobertura bibliográfica
  let cobertura = null;
  if (bibliografia.length > 0) {
    const disponibles = bibliografia.filter(r => {
      const d = String(r.disponible || '').toLowerCase();
      return d === 'si' || d === 'sí' || d === 'yes' || d === '1' || d === 'true' || d === 'disponible';
    }).length;
    cobertura = (disponibles / bibliografia.length) * 100;
  }

  // I) Demanda insatisfecha
  let demandaInsatisfecha = null;
  if (reservas.length > 0) {
    const totalSolicitudes = reservas.reduce((s, r) => s + (parseInt(r.solicitudes) || 0), 0);
    const totalSatisfechas = reservas.reduce((s, r) => s + (parseInt(r.satisfechas) || 0), 0);
    const noSatisfechas = totalSolicitudes - totalSatisfechas;
    demandaInsatisfecha = totalSolicitudes > 0 ? (noSatisfechas / totalSolicitudes) * 100 : 0;
  }

  const indicators = [
    {
      id: 'tasa_circulacion',
      name: 'Tasa de Circulación',
      value: round2(tasaCirculacion),
      unit: 'préstamos/ejemplar',
      reference: 'Saludable: ≥ 1.5 | Revisar: 0.5–1.5 | Crítico: < 0.5',
      status: getStatus(tasaCirculacion, 1.5, 0.5),
      description: 'Relación entre préstamos totales y ejemplares disponibles.'
    },
    {
      id: 'coleccion_usada',
      name: 'Colección Usada',
      value: round2(pctUsada),
      unit: '%',
      reference: 'Saludable: ≥ 60% | Revisar: 30–60% | Crítico: < 30%',
      status: getStatus(pctUsada, 60, 30),
      description: `${titulosUsados} de ${totalTitulos} títulos con al menos un préstamo.`
    },
    {
      id: 'indice_rotacion',
      name: 'Índice de Rotación',
      value: round2(indiceRotacion),
      unit: 'veces',
      reference: 'Saludable: ≥ 2 | Revisar: 1–2 | Crítico: < 1',
      status: getStatus(indiceRotacion, 2, 1),
      description: 'Número de veces que los ejemplares han circulado en promedio.'
    },
    {
      id: 'obsolescencia',
      name: 'Obsolescencia',
      value: round2(obsolescencia),
      unit: '%',
      reference: 'Saludable: ≤ 20% | Revisar: 20–40% | Crítico: > 40%',
      status: getStatus(obsolescencia, 20, 40, true),
      description: `${titulosSinUso} títulos sin ningún préstamo registrado.`
    },
    {
      id: 'antiguedad_sin_uso',
      name: 'Antigüedad Promedio sin Uso',
      value: round2(antiguedadPromedio),
      unit: 'años',
      reference: 'Saludable: ≤ 3 | Revisar: 3–5 | Crítico: > 5',
      status: getStatus(antiguedadPromedio, 3, 5, true),
      description: 'Promedio de años desde el último préstamo en títulos inactivos.'
    },
    {
      id: 'costo_por_uso',
      name: 'Costo por Uso',
      value: round2(costoUso),
      unit: '$/préstamo',
      reference: 'Menor valor = mayor eficiencia',
      status: 'blue',
      description: 'Costo promedio de la colección dividido por el número de préstamos.'
    },
  ];

  if (cobertura !== null) {
    indicators.push({
      id: 'cobertura',
      name: 'Cobertura Bibliográfica',
      value: round2(cobertura),
      unit: '%',
      reference: 'Saludable: ≥ 80% | Revisar: 60–80% | Crítico: < 60%',
      status: getStatus(cobertura, 80, 60),
      description: 'Porcentaje de títulos en bibliografía académica disponibles en catálogo.'
    });
  }

  if (demandaInsatisfecha !== null) {
    indicators.push({
      id: 'demanda_insatisfecha',
      name: 'Demanda Insatisfecha',
      value: round2(demandaInsatisfecha),
      unit: '%',
      reference: 'Saludable: ≤ 10% | Revisar: 10–25% | Crítico: > 25%',
      status: getStatus(demandaInsatisfecha, 10, 25, true),
      description: 'Porcentaje de reservas no satisfechas sobre el total de solicitudes.'
    });
  }

  return indicators;
}

// ─── AREA METRICS ─────────────────────────────────────────────

function calculateAreaMetrics(coleccion, prestamos) {
  // Agrupar colección por área
  const byArea = groupBy(coleccion, r => r.area || 'Sin área');
  // Mapear id_titulo → area
  const titleToArea = {};
  coleccion.forEach(r => { titleToArea[String(r.id || '').trim()] = r.area || 'Sin área'; });

  // Contar préstamos por área
  const loansByArea = {};
  prestamos.forEach(p => {
    const area = titleToArea[String(p.id_titulo || '').trim()] || 'Sin área';
    loansByArea[area] = (loansByArea[area] || 0) + 1;
  });

  return Object.entries(byArea).map(([area, items]) => {
    const ejemplares = items.reduce((s, r) => s + (r.ejemplares || 1), 0);
    const loans = loansByArea[area] || 0;
    return {
      area,
      titulos: items.length,
      ejemplares,
      loans,
      ratio: ejemplares > 0 ? round2(loans / ejemplares) : 0
    };
  }).sort((a, b) => b.ratio - a.ratio);
}

// ─── OBSOLESCENCE ─────────────────────────────────────────────

function calculateObsolescence(coleccion) {
  return coleccion.map(item => {
    const yearsUnused = item._yearsUnused;
    const loans = item._loans || 0;
    const ej = item.ejemplares || 1;

    // Índice de obsolescencia: combina años sin uso + bajo uso
    let idx = 0;
    if (yearsUnused !== null) idx += Math.min(yearsUnused / 10, 5); // máx 5 puntos por antigüedad
    if (loans === 0) idx += 3;
    else if (loans < 3) idx += 1.5;
    if (ej > 3 && loans < 2) idx += 1; // duplicación sin uso

    let status;
    if (idx >= 7 || (yearsUnused !== null && yearsUnused > 10 && loans === 0)) status = 'Crítico';
    else if (idx >= 4 || (yearsUnused !== null && yearsUnused > 5)) status = 'Revisar';
    else if (loans < 3) status = 'Bajo uso';
    else status = 'Activo';

    return {
      ...item,
      _obsIndex: round2(idx),
      _obsStatus: status,
      _yearsUnused: yearsUnused
    };
  })
  .filter(r => r._obsStatus !== 'Activo')
  .sort((a, b) => b._obsIndex - a._obsIndex);
}

// ─── RECOMMENDATIONS ──────────────────────────────────────────

function generateRecommendations(indicators, coleccion, prestamos, areaMetrics, obsolescence) {
  const recs = [];

  const get = id => indicators.find(i => i.id === id);

  // Tasa de circulación
  const tc = get('tasa_circulacion');
  if (tc && tc.status === 'red') {
    recs.push({
      priority: 'high',
      area: 'Circulación',
      icon: '🚨',
      text: `La tasa de circulación es ${num(tc.value)} préstamos por ejemplar, por debajo del umbral mínimo recomendado (0.5). Esto indica que la colección está siendo significativamente subutilizada.`,
      action: 'Revisar políticas de préstamo, horarios de atención y estrategias de difusión de la colección.'
    });
  } else if (tc && tc.status === 'yellow') {
    recs.push({
      priority: 'medium',
      area: 'Circulación',
      icon: '⚠️',
      text: `La tasa de circulación (${num(tc.value)}) está en rango de revisión. La colección muestra uso moderado pero hay margen de mejora.`,
      action: 'Considerar campañas de visibilidad y actividades de fomento lector en las facultades.'
    });
  } else if (tc && tc.status === 'green') {
    recs.push({
      priority: 'low',
      area: 'Circulación',
      icon: '✅',
      text: `La tasa de circulación (${num(tc.value)}) indica un uso saludable de la colección. El préstamo domiciliario responde bien a las necesidades académicas.`,
      action: 'Mantener las políticas actuales y monitorear trimestralmente.'
    });
  }

  // Obsolescencia
  const obs = get('obsolescencia');
  if (obs && obs.status === 'red') {
    const critCount = obsolescence.filter(r => r._obsStatus === 'Crítico').length;
    recs.push({
      priority: 'high',
      area: 'Obsolescencia',
      icon: '📚',
      text: `El ${num(obs.value)}% de la colección no registra circulación. Se identificaron ${critCount} títulos en estado crítico de obsolescencia (más de 10 años sin uso).`,
      action: 'Iniciar proceso formal de descarte según políticas institucionales. Priorizar los títulos con índice de obsolescencia más alto.'
    });
  } else if (obs && obs.status === 'yellow') {
    recs.push({
      priority: 'medium',
      area: 'Obsolescencia',
      icon: '🕐',
      text: `El ${num(obs.value)}% de la colección muestra baja o nula circulación. Varios títulos llevan entre 5 y 10 años sin préstamos registrados.`,
      action: 'Realizar revisión bibliográfica y evaluar pertinencia curricular antes de tomar decisiones de descarte.'
    });
  }

  // Áreas de alta presión
  const altaPresion = areaMetrics.filter(a => a.ratio > 3);
  if (altaPresion.length > 0) {
    const nombres = altaPresion.slice(0, 3).map(a => a.area).join(', ');
    recs.push({
      priority: 'high',
      area: 'Alta demanda',
      icon: '📈',
      text: `Las áreas de ${nombres} presentan alta presión de uso (índice > 3.0). La disponibilidad de ejemplares podría estar generando demanda insatisfecha.`,
      action: 'Evaluar adquisición de nuevos ejemplares o suscripciones a versiones digitales para las áreas con mayor demanda.'
    });
  }

  // Áreas sin uso
  const sinUso = areaMetrics.filter(a => a.ratio === 0 && a.titulos > 3);
  if (sinUso.length > 0) {
    const nombres = sinUso.slice(0, 3).map(a => a.area).join(', ');
    recs.push({
      priority: 'medium',
      area: 'Áreas inactivas',
      icon: '📉',
      text: `Las áreas de ${nombres} no registran circulación a pesar de contar con colección disponible. Esto puede indicar desalineación curricular o desconocimiento de los recursos.`,
      action: 'Coordinar con las unidades académicas respectivas para evaluar la pertinencia de la colección y diseñar actividades de difusión.'
    });
  }

  // Cobertura bibliográfica
  const cob = get('cobertura');
  if (cob && cob.status === 'red') {
    recs.push({
      priority: 'high',
      area: 'Bibliografía académica',
      icon: '📖',
      text: `La cobertura bibliográfica es del ${num(cob.value)}%, lo que significa que más del ${round2(100 - cob.value)}% de los títulos requeridos por las asignaturas no están disponibles en el catálogo.`,
      action: 'Urgente: solicitar presupuesto de adquisición y priorizar los títulos faltantes de mayor uso académico.'
    });
  } else if (cob && cob.status === 'yellow') {
    recs.push({
      priority: 'medium',
      area: 'Bibliografía académica',
      icon: '📖',
      text: `La cobertura bibliográfica del ${num(cob.value)}% indica que existen áreas con baja cobertura de bibliografía obligatoria.`,
      action: 'Revisar los programas de asignatura y priorizar adquisiciones según carga académica.'
    });
  }

  // Demanda insatisfecha
  const di = get('demanda_insatisfecha');
  if (di && di.status === 'red') {
    recs.push({
      priority: 'high',
      area: 'Reservas',
      icon: '⏳',
      text: `El ${num(di.value)}% de las reservas no pueden ser satisfechas. Esto refleja una brecha significativa entre la disponibilidad de ejemplares y la demanda real de los usuarios.`,
      action: 'Revisar el sistema de reservas, evaluar adquisición de ejemplares adicionales en los títulos con más solicitudes.'
    });
  }

  // Costo por uso
  const cpu = get('costo_por_uso');
  if (cpu && cpu.value > 0) {
    recs.push({
      priority: 'info',
      area: 'Eficiencia económica',
      icon: '💰',
      text: `El costo promedio por uso de la colección es de $${num(cpu.value)} por préstamo. Este indicador permite evaluar el retorno de inversión de las adquisiciones bibliográficas.`,
      action: 'Usar este dato como criterio en la evaluación de nuevas adquisiciones, priorizando recursos con alta demanda proyectada.'
    });
  }

  // Recomendación general si todo está bien
  if (recs.filter(r => r.priority === 'high').length === 0) {
    recs.push({
      priority: 'info',
      area: 'Estado general',
      icon: '🎯',
      text: 'Los indicadores generales de la colección muestran un estado aceptable. Se recomienda mantener el monitoreo periódico y establecer alertas automáticas para títulos que superen los umbrales de obsolescencia.',
      action: 'Programar revisión semestral de indicadores y actualización del análisis con nuevos datos de circulación.'
    });
  }

  return recs;
}

// ─── RENDER DASHBOARD ─────────────────────────────────────────

function renderDashboard(indicators, coleccion, areaMetrics) {
  // KPI Cards
  const kpiGrid = el('kpiGrid');
  kpiGrid.innerHTML = indicators.map(ind => `
    <div class="kpi-card status-${ind.status}">
      <span class="kpi-badge ${ind.status}"></span>
      <div class="kpi-label">${ind.name}</div>
      <div class="kpi-value">
        ${Number.isFinite(ind.value) ? ind.value : '—'}
        <span class="kpi-unit">${ind.unit}</span>
      </div>
      <div class="kpi-desc">${ind.description}</div>
    </div>
  `).join('');

  // Semáforo
  const semaforoGrid = el('semaforoGrid');
  semaforoGrid.innerHTML = indicators.map(ind => `
    <div class="semaforo-item ${ind.status}">
      <span class="semaforo-dot"></span>
      <div class="semaforo-text">
        <div class="semaforo-name">${ind.name}</div>
        <div class="semaforo-val">${statusLabel(ind.status)} · ${num(ind.value)} ${ind.unit}</div>
      </div>
    </div>
  `).join('');

  // Indicadores detallados
  const tbody = el('indicatorsBody');
  tbody.innerHTML = indicators.map(ind => `
    <tr>
      <td><strong>${ind.name}</strong></td>
      <td class="mono">${num(ind.value)} ${ind.unit}</td>
      <td style="color:var(--text-secondary);font-size:12px">${ind.reference}</td>
      <td><span class="status-pill ${ind.status}">${statusEmoji(ind.status)} ${statusLabel(ind.status)}</span></td>
      <td style="color:var(--text-secondary);font-size:13px">${ind.description}</td>
    </tr>
  `).join('');
}

// ─── RENDER OBSOLESCENCE ──────────────────────────────────────

function renderObsolescence(obsolescence) {
  // Summary cards
  const critico = obsolescence.filter(r => r._obsStatus === 'Crítico').length;
  const revisar = obsolescence.filter(r => r._obsStatus === 'Revisar').length;
  const bajoUso = obsolescence.filter(r => r._obsStatus === 'Bajo uso').length;

  el('obsSummaryGrid').innerHTML = `
    <div class="obs-summary-card">
      <div class="obs-summary-num red">${critico}</div>
      <div class="obs-summary-label">Críticos (>10 años sin uso)</div>
    </div>
    <div class="obs-summary-card">
      <div class="obs-summary-num yellow">${revisar}</div>
      <div class="obs-summary-label">Para revisar (5–10 años)</div>
    </div>
    <div class="obs-summary-card">
      <div class="obs-summary-num blue">${bajoUso}</div>
      <div class="obs-summary-label">Bajo uso (&lt;3 préstamos)</div>
    </div>
    <div class="obs-summary-card">
      <div class="obs-summary-num" style="color:var(--text-secondary)">${obsolescence.length}</div>
      <div class="obs-summary-label">Total con alerta activa</div>
    </div>
  `;

  renderObsTable(obsolescence);
}

function renderObsTable(data) {
  const tbody = el('obsBody');
  if (!data || data.length === 0) {
    tbody.innerHTML = `<tr><td colspan="9" style="text-align:center;padding:40px;color:var(--text-tertiary)">Sin datos de obsolescencia</td></tr>`;
    return;
  }

  tbody.innerHTML = data.map((item, i) => {
    const idx = item._obsIndex || 0;
    const maxIdx = 10;
    const pct = Math.min((idx / maxIdx) * 100, 100);
    const fillClass = item._obsStatus === 'Crítico' ? 'red' : item._obsStatus === 'Revisar' ? 'yellow' : 'green';
    const pillClass = item._obsStatus === 'Crítico' ? 'red' : item._obsStatus === 'Revisar' ? 'yellow' : 'blue';

    return `
      <tr data-status="${item._obsStatus}" data-titulo="${(item.titulo||'').toLowerCase()}">
        <td class="mono" style="color:var(--text-tertiary)">${i+1}</td>
        <td><strong style="font-size:13px">${item.titulo || '—'}</strong></td>
        <td style="color:var(--text-secondary)">${item.area || '—'}</td>
        <td class="mono">${item.ejemplares || 1}</td>
        <td>${formatDate(item.ultimo_prestamo)}</td>
        <td class="mono" style="color:${item._yearsUnused > 5 ? 'var(--red)' : 'var(--text-secondary)'}">
          ${item._yearsUnused != null && item._yearsUnused < 99 ? item._yearsUnused : 'Sin registro'}
        </td>
        <td class="mono">${item._loans || 0}</td>
        <td>
          <div class="idx-bar">
            <div class="idx-track">
              <div class="idx-fill ${fillClass}" style="width:${pct}%"></div>
            </div>
            <span class="idx-num">${round2(idx)}</span>
          </div>
        </td>
        <td><span class="status-pill ${pillClass}">${item._obsStatus}</span></td>
      </tr>
    `;
  }).join('');
}

// ─── FILTER OBSOLESCENCE TABLE ────────────────────────────────

function filterObsTable() {
  const statusFilter = el('obsFilterStatus')?.value || '';
  const searchFilter = (el('obsSearch')?.value || '').toLowerCase();

  document.querySelectorAll('#obsBody tr').forEach(row => {
    const rowStatus = (row.dataset.status || '').toLowerCase();
    const rowTitulo = row.dataset.titulo || '';
    const statusMatch = !statusFilter || rowStatus === statusFilter;
    const searchMatch = !searchFilter || rowTitulo.includes(searchFilter);
    row.style.display = statusMatch && searchMatch ? '' : 'none';
  });
}

// ─── RENDER RECOMMENDATIONS ───────────────────────────────────

function renderRecommendations(recommendations) {
  const container = el('recommendationsContainer');
  if (!recommendations || recommendations.length === 0) {
    container.innerHTML = `<div class="empty-state"><div class="empty-state-icon">💡</div><h3>Sin recomendaciones disponibles</h3><p>Carga datos para generar recomendaciones automáticas.</p></div>`;
    return;
  }

  container.innerHTML = recommendations.map(rec => `
    <div class="rec-card priority-${rec.priority}">
      <div class="rec-icon">${rec.icon}</div>
      <div class="rec-content">
        <div class="rec-area">${rec.area}</div>
        <div class="rec-text">${rec.text}</div>
        <div class="rec-action">→ ${rec.action}</div>
      </div>
    </div>
  `).join('');
}

// ─── POPULATE FILTERS ─────────────────────────────────────────

function populateFilters(coleccion, prestamos) {
  const areas = [...new Set(coleccion.map(r => r.area).filter(Boolean))].sort();
  const carreras = [...new Set(prestamos.map(r => r.carrera).filter(Boolean))].sort();

  const filterArea = el('filterArea');
  filterArea.innerHTML = '<option value="">Todas las áreas</option>' +
    areas.map(a => `<option value="${a}">${a}</option>`).join('');

  const filterCarrera = el('filterCarrera');
  filterCarrera.innerHTML = '<option value="">Todas las carreras</option>' +
    carreras.map(c => `<option value="${c}">${c}</option>`).join('');
}

function applyFilters() {
  // Por ahora simplemente refresca el dashboard con los datos sin filtrar
  // Una expansión futura filtreará los datos y recalculará
  if (window.appData) {
    const areaFilter = el('filterArea')?.value;
    const carreraFilter = el('filterCarrera')?.value;

    let coleccion = window.appData.coleccion;
    let prestamos = window.appData.prestamos;

    if (areaFilter) coleccion = coleccion.filter(r => r.area === areaFilter);
    if (carreraFilter) prestamos = prestamos.filter(r => r.carrera === carreraFilter);

    const indicators = calculateIndicators(coleccion, prestamos, window.appData.bibliografia, window.appData.reservas);
    const areaMetrics = calculateAreaMetrics(coleccion, prestamos);
    renderDashboard(indicators, coleccion, areaMetrics);
  }
}

// ─── MANUAL DATA ENTRY ────────────────────────────────────────

const MANUAL_SCHEMAS = {
  coleccion: [
    { key: 'id', label: 'ID', type: 'text', placeholder: 'LIB001' },
    { key: 'titulo', label: 'Título', type: 'text', placeholder: 'Título del libro', wide: true },
    { key: 'area', label: 'Área', type: 'text', placeholder: 'Ciencias' },
    { key: 'ejemplares', label: 'Ejemplares', type: 'number', placeholder: '1' },
    { key: 'ultimo_prestamo', label: 'Último Préstamo', type: 'date', placeholder: '' },
    { key: 'costo', label: 'Costo', type: 'number', placeholder: '0' },
  ],
  prestamos: [
    { key: 'id_titulo', label: 'ID Título', type: 'text', placeholder: 'LIB001' },
    { key: 'fecha', label: 'Fecha', type: 'date', placeholder: '' },
    { key: 'carrera', label: 'Carrera', type: 'text', placeholder: 'Ingeniería Civil' },
  ],
  bibliografia: [
    { key: 'carrera', label: 'Carrera', type: 'text', placeholder: 'Ingeniería Civil' },
    { key: 'asignatura', label: 'Asignatura', type: 'text', placeholder: 'Metodología' },
    { key: 'titulo', label: 'Título', type: 'text', placeholder: 'Título del libro', wide: true },
    { key: 'disponible', label: 'Disponible', type: 'select', options: ['Sí','No'] },
  ],
  reservas: [
    { key: 'titulo', label: 'Título', type: 'text', placeholder: 'Título del libro', wide: true },
    { key: 'solicitudes', label: 'Solicitudes', type: 'number', placeholder: '0' },
    { key: 'satisfechas', label: 'Satisfechas', type: 'number', placeholder: '0' },
  ]
};

function addManualRow(sheetName) {
  const tbody = el(`tbody-${sheetName}`);
  const schema = MANUAL_SCHEMAS[sheetName];
  if (!tbody || !schema) return;

  const tr = document.createElement('tr');
  tr.innerHTML = schema.map(col => {
    let input;
    if (col.type === 'select') {
      input = `<select><option value="">Seleccionar</option>${col.options.map(o => `<option>${o}</option>`).join('')}</select>`;
    } else {
      input = `<input type="${col.type}" placeholder="${col.placeholder || ''}" style="${col.wide ? 'min-width:200px' : ''}" />`;
    }
    return `<td>${input}</td>`;
  }).join('') + `<td><button class="btn-del-row" onclick="this.closest('tr').remove()">×</button></td>`;

  tbody.appendChild(tr);
}

function processManualData() {
  const sheets = {};
  for (const sheetName of ['coleccion','prestamos','bibliografia','reservas']) {
    const schema = MANUAL_SCHEMAS[sheetName];
    const tbody = el(`tbody-${sheetName}`);
    if (!tbody) continue;

    const rows = [];
    tbody.querySelectorAll('tr').forEach(tr => {
      const inputs = tr.querySelectorAll('input, select');
      if (inputs.length === 0) return;
      const row = {};
      schema.forEach((col, i) => {
        if (inputs[i]) row[col.key] = inputs[i].value;
      });
      // Solo agregar si tiene al menos un campo con valor
      if (Object.values(row).some(v => v && v.trim && v.trim() !== '')) {
        rows.push(row);
      }
    });
    if (rows.length > 0) sheets[sheetName] = rows;
  }

  if (!sheets.coleccion || sheets.coleccion.length === 0) {
    showToast('Debes ingresar al menos datos en la hoja Colección', 'error');
    return;
  }
  if (!sheets.prestamos || sheets.prestamos.length === 0) {
    showToast('Debes ingresar al menos datos en la hoja Préstamos', 'error');
    return;
  }

  // Normalizar columnas (ya están en el formato correcto)
  const normalized = {};
  for (const [name, rows] of Object.entries(sheets)) {
    normalized[name] = rows.map(row => {
      if (row.fecha) {
        const d = parseDate(row.fecha);
        if (d) row._fechaDate = d;
      }
      if (name === 'coleccion') {
        row.ejemplares = parseInt(row.ejemplares) || 1;
        row.costo = parseFloat(row.costo) || 0;
        if (row.ultimo_prestamo) row._ultimaFechaDate = parseDate(row.ultimo_prestamo);
      }
      return row;
    });
  }

  processData(normalized);
  goToDashboard();
}

// ─── NAVIGATION HELPERS ───────────────────────────────────────

function goToDashboard() {
  navigateTo('dashboard');
  showToast('Análisis completado', 'success');
}

function clearAllData() {
  window.appData = null;
  el('validationPanel').style.display = 'none';
  el('btnContinue').style.display = 'none';
  el('fileInput').value = '';

  // Limpiar tablas manuales
  ['coleccion','prestamos','bibliografia','reservas'].forEach(s => {
    const tb = el(`tbody-${s}`);
    if (tb) tb.innerHTML = '';
  });

  // Limpiar KPI
  el('kpiGrid').innerHTML = '';
  el('semaforoGrid').innerHTML = '';
  el('indicatorsBody').innerHTML = '';
  el('obsSummaryGrid').innerHTML = '';
  el('obsBody').innerHTML = '';
  el('recommendationsContainer').innerHTML = '';

  // Destruir charts
  ['chartCarrera','chartUsage','chartMonthly','chartArea','chartRanking'].forEach(destroyChart);

  navigateTo('upload');
  showToast('Datos eliminados', 'default');
}

// ─── DEMO DATA ────────────────────────────────────────────────

function loadDemoData() {
  showLoader('Cargando datos de demostración…');

  setTimeout(() => {
    const areas = ['Ingeniería','Ciencias','Humanidades','Derecho','Medicina','Arquitectura','Educación','Economía'];
    const carreras = ['Ingeniería Civil','Ingeniería Industrial','Medicina','Derecho','Arquitectura','Pedagogía','Economía','Psicología','Biología','Historia'];

    // Generar colección
    const coleccion = [];
    for (let i = 1; i <= 80; i++) {
      const area = areas[Math.floor(Math.random() * areas.length)];
      const ejemplares = Math.floor(Math.random() * 5) + 1;
      const sinUso = Math.random() < 0.3; // 30% sin uso
      let ultimo_prestamo = null;
      let _ultimaFechaDate = null;
      if (!sinUso) {
        const daysAgo = Math.floor(Math.random() * 2555) + 30; // hasta 7 años atrás
        const d = new Date(Date.now() - daysAgo * 86400000);
        ultimo_prestamo = d.toISOString().split('T')[0];
        _ultimaFechaDate = d;
      }
      const costo = (Math.floor(Math.random() * 40) + 5) * 1000;
      coleccion.push({
        id: `LIB${String(i).padStart(3,'0')}`,
        titulo: demoBooksTitle(i, area),
        area,
        ejemplares,
        ultimo_prestamo,
        costo,
        _ultimaFechaDate
      });
    }

    // Generar préstamos
    const prestamos = [];
    const now = new Date();
    for (let i = 0; i < 450; i++) {
      const book = coleccion[Math.floor(Math.random() * 60)]; // concentrar uso en primeros 60
      const carrera = carreras[Math.floor(Math.random() * carreras.length)];
      const daysAgo = Math.floor(Math.random() * 1095); // últimos 3 años
      const fecha = new Date(now.getTime() - daysAgo * 86400000);
      prestamos.push({
        id_titulo: book.id,
        fecha: fecha.toISOString().split('T')[0],
        carrera,
        _fechaDate: fecha
      });
    }

    // Actualizar _loans
    const loanCount = {};
    prestamos.forEach(p => { loanCount[p.id_titulo] = (loanCount[p.id_titulo] || 0) + 1; });
    coleccion.forEach(item => {
      item._loans = loanCount[item.id] || 0;
      item._yearsUnused = item._ultimaFechaDate ? yearsSince(item._ultimaFechaDate) : (item._loans === 0 ? 99 : null);
    });

    // Generar bibliografía
    const bibliografia = [];
    carreras.slice(0,7).forEach(carrera => {
      for (let i = 0; i < 8; i++) {
        const disponible = Math.random() < 0.72 ? 'Sí' : 'No';
        bibliografia.push({
          carrera,
          asignatura: `Asignatura ${i+1}`,
          titulo: coleccion[Math.floor(Math.random() * coleccion.length)].titulo,
          disponible
        });
      }
    });

    // Generar reservas
    const reservas = [];
    for (let i = 0; i < 25; i++) {
      const solicitudes = Math.floor(Math.random() * 15) + 3;
      const satisfechas = Math.floor(Math.random() * solicitudes);
      reservas.push({
        titulo: coleccion[Math.floor(Math.random() * 30)].titulo,
        solicitudes,
        satisfechas
      });
    }

    hideLoader();

    window.appData = null; // Reset
    processData({ coleccion, prestamos, bibliografia, reservas });
    goToDashboard();

    showToast('Datos de ejemplo cargados correctamente', 'success');
  }, 800);
}

function demoBooksTitle(i, area) {
  const titles = {
    'Ingeniería': ['Mecánica de Materiales','Análisis Estructural','Termodinámica Aplicada','Hidráulica General','Ingeniería de Procesos','Resistencia de Materiales','Diseño de Máquinas'],
    'Ciencias': ['Cálculo Diferencial e Integral','Álgebra Lineal','Estadística para Ingeniería','Química General','Física Universitaria','Biología Celular','Métodos Numéricos'],
    'Humanidades': ['Historia de América Latina','Filosofía Contemporánea','Lingüística General','Literatura Hispanoamericana','Antropología Cultural','Ética y Sociedad','Arte y Pensamiento'],
    'Derecho': ['Derecho Civil General','Derecho Constitucional','Teoría del Delito','Derecho Internacional Público','Procedimiento Penal','Derecho Administrativo'],
    'Medicina': ['Anatomía Humana','Fisiología Médica','Farmacología Clínica','Patología General','Semiología Médica','Cirugía Básica'],
    'Arquitectura': ['Teoría de la Arquitectura','Historia del Arte','Diseño Urbano','Estructuras para Arquitectos','Instalaciones Sanitarias'],
    'Educación': ['Didáctica General','Psicología del Aprendizaje','Evaluación Educativa','Curriculum y Planificación'],
    'Economía': ['Microeconomía','Macroeconomía','Economía Internacional','Finanzas Corporativas','Econometría']
  };
  const list = titles[area] || ['Introducción a la Disciplina','Fundamentos Teóricos','Metodología de Investigación','Temas Avanzados'];
  return `${list[i % list.length]} — ${area} ${Math.floor(i/list.length) + 1}ª ed.`;
}
