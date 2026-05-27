(() => {
  const $ = s => document.querySelector(s);

  // --- LISTA MAESTRA DE PERSONAL COMPLETA Y ACTUALIZADA ---
  const MAESTRO_PERSONAL = {
    "lamper torres jose ali": { empresa: "Top Solution Metal SAC", contrato: "Planilla" },
    "munoz garcia miguel angel": { empresa: "Top Solution Metal SAC", contrato: "Planilla" },
    "rincon ramirez giovanni jose": { empresa: "Top Solution Metal SAC", contrato: "Planilla" },
    "rubina reynoso nilton aquiles": { empresa: "Top Solution Metal SAC", contrato: "Planilla" },
    "saldana alanya carlos alberto": { empresa: "Top Solution Metal SAC", contrato: "Planilla" },
    "valderrama alva sergio angel": { empresa: "Top Solution Metal SAC", contrato: "Planilla" },
    "valderrama mozombite gustavo": { empresa: "Top Solution Metal SAC", contrato: "Planilla" },
    "cabrera carlos david": { empresa: "Top Solution Metal SAC", contrato: "Recibos por Honorarios" },
    "capcha rojas genaro vicente": { empresa: "Top Solution Metal SAC", contrato: "Recibos por Honorarios" },
    "cordova diaz pablo ismael": { empresa: "Top Solution Metal SAC", contrato: "Recibos por Honorarios" },
    "cordova mori adriel": { empresa: "Top Solution Metal SAC", contrato: "Recibos por Honorarios" },
    "cotos alba alex luis": { empresa: "Top Solution Metal SAC", contrato: "Recibos por Honorarios" },
    "culqui aspajo pedro augusto": { empresa: "Top Solution Metal SAC", contrato: "Recibos por Honorarios" },
    "curitima murayari jorge miguel": { empresa: "Top Solution Metal SAC", contrato: "Recibos por Honorarios" },
    "estrada sanchez luiggi junior": { empresa: "Top Solution Metal SAC", contrato: "Recibos por Honorarios" },
    "gonzales chavez jesus": { empresa: "Top Solution Metal SAC", contrato: "Recibos por Honorarios" },
    "hernandez romero wilfredo": { empresa: "Top Solution Metal SAC", contrato: "Recibos por Honorarios" },
    "mendoza chauran luis enrique": { empresa: "Top Solution Metal SAC", contrato: "Recibos por Honorarios" },
    "montes bullon christopher": { empresa: "Top Solution Metal SAC", contrato: "Recibos por Honorarios" },
    "ortiz pinche jhander jhonathan": { empresa: "Top Solution Metal SAC", contrato: "Recibos por Honorarios" },
    "prada del pino jose luis": { empresa: "Top Solution Metal SAC", contrato: "Recibos por Honorarios" },
    "roncal quinto luis alberto": { empresa: "Top Solution Metal SAC", contrato: "Recibos por Honorarios" },
    "salas munar juan antonio": { empresa: "Top Solution Metal SAC", contrato: "Recibos por Honorarios" },
    "tamaris angeles jose antonio": { empresa: "Top Solution Metal SAC", contrato: "Recibos por Honorarios" },
    "andrade riobueno marco antonio": { empresa: "Top Solution Metal Group", contrato: "Planilla" },
    "bonifacio rodriguez yanfranco": { empresa: "Top Solution Metal Group", contrato: "Planilla" },
    "galindo cabezas joel serafin": { empresa: "Top Solution Metal Group", contrato: "Planilla" },
    "garcia cornelio roberto carlos": { empresa: "Top Solution Metal Group", contrato: "Planilla" },
    "gonzales llumpo cristian esteban": { empresa: "Top Solution Metal Group", contrato: "Planilla" },
    "lazaro salas jhon javier": { empresa: "Top Solution Metal Group", contrato: "Planilla" },
    "mejia perez jose": { empresa: "Top Solution Metal Group", contrato: "Planilla" },
    "pezo fatama tono": { empresa: "Top Solution Metal Group", contrato: "Planilla" },
    "rivas sotomayor yorman rafael": { empresa: "Top Solution Metal Group", contrato: "Planilla" },
    "sequeiros prudencio jose luis": { empresa: "Top Solution Metal Group", contrato: "Planilla" },
    "suarez torrelles jehison omar": { empresa: "Top Solution Metal Group", contrato: "Planilla" },
    "aguilar mejia jose cirilo": { empresa: "Top Solution Metal Group", contrato: "Recibos por Honorarios" },
    "chahua villa efrain": { empresa: "Top Solution Metal Group", contrato: "Recibos por Honorarios" },
    "cipiran ramirez abraham moises": { empresa: "Top Solution Metal Group", contrato: "Recibos por Honorarios" },
    "collazos marquez eliot karl": { empresa: "Top Solution Metal Group", contrato: "Recibos por Honorarios" },
    "curitima murayari neuber": { empresa: "Top Solution Metal Group", contrato: "Recibos por Honorarios" },
    "delgado aguinaga anderson eduardo": { empresa: "Top Solution Metal Group", contrato: "Recibos por Honorarios" },
    "elescano ortiz leandro federico": { empresa: "Top Solution Metal Group", contrato: "Recibos por Honorarios" },
    "elescano ortiz leonardo antonio": { empresa: "Top Solution Metal Group", contrato: "Recibos por Honorarios" },
    "flores tuanama geiner": { empresa: "Top Solution Metal Group", contrato: "Recibos por Honorarios" },
    "lopez quispe karlo francisco": { empresa: "Top Solution Metal Group", contrato: "Recibos por Honorarios" },
    "pina gordon daniel andres": { empresa: "Top Solution Metal Group", contrato: "Recibos por Honorarios" },
    "ramirez arizmendi manuel oscar": { empresa: "Top Solution Metal Group", contrato: "Recibos por Honorarios" },
    "reategui dahua oris leonides": { empresa: "Top Solution Metal Group", contrato: "Recibos por Honorarios" },
    "sias dahua genis": { empresa: "Top Solution Metal Group", contrato: "Recibos por Honorarios" },
    "urquia guevara orlando": { empresa: "Top Solution Metal Group", contrato: "Recibos por Honorarios" },
    "yumbato hidalgo juan carlos": { empresa: "Top Solution Metal Group", contrato: "Recibos por Honorarios" },
    "zuta vilca edinson": { empresa: "Top Solution Metal Group", contrato: "Recibos por Honorarios" }
  };

  const SETTINGS = { dedupeMinutes: 3, topeHsDiario: 6 };

  const state = {
    workbook: null, rawRows: [], cleanRows: [], summaryRows: [], anomalyRows: [],
    workers: [], tareo: { dates: [], rows: [] }, fileName: 'tareo_tsm.xlsx', map: null
  };

  const els = {
    fileInput: $('#fileInput'), fileName: $('#fileName'), dropzone: $('#dropzone'),
    themeBtn: $('#themeToggleBtn'), errorBox: $('#errorBox'), statusBadge: $('#statusBadge'),
    startDate: $('#startDate'), endDate: $('#endDate'), processBtn: $('#processBtn'),
    downloadXlsxBtn: $('#downloadXlsxBtn'), downloadCsvBtn: $('#downloadCsvBtn'),
    sampleBtn: $('#sampleBtn'), statRows: $('#statRows'), statWorkers: $('#statWorkers'),
    statDays: $('#statDays'), statSummary: $('#statSummary'), filterBar: $('#filterBar'),
    filterEmpresa: $('#filterEmpresa'), filterContrato: $('#filterContrato'),
    clearFiltersBtn: $('#clearFiltersBtn'), fullscreenBtn: $('#fullscreenBtn'),
    resumen: $('#tab-resumen'), tareo: $('#tab-tareo'), bd: $('#tab-bd'),
    anomalias: $('#tab-anomalias'), comparativa: $('#tab-comparativa')
  };

  const DAYS = ['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'];
  const HOLIDAY_OVERRIDES = new Set(['02/04/2026','03/04/2026']);

  const pad = n => String(n).padStart(2, '0');
  const normalize = v => String(v ?? '').trim().toLowerCase();
  const normalizeSimple = v => normalize(v).normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9 ]+/g, ' ').replace(/\s+/g, ' ').trim();

  const fmtDate = d => `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()}`;
  const fmtInput = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const fmtTime = d => `${pad(d.getHours())}:${pad(d.getMinutes())}`;
  const weekdayName = d => DAYS[d.getDay()];
  const dayOnly = d => new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const stamp = d => dayOnly(d).getTime();
  const isSunday = d => d.getDay() === 0;
  const roundDownHalf = n => Math.floor((n || 0) * 2) / 2;
  const hoursBetween = (a, b) => Math.max(0, (b - a) / 36e5);
  const diffMinutes = (a, b) => Math.abs((b - a) / 60000);

  function displayNum(v) {
    if (v == null || v === '' || v === 0) return v === 0 ? '0' : '';
    const n = Number(v);
    return isFinite(n) ? n.toFixed(1).replace(/\.0$/, '') : String(v);
  }

  function escapeHtml(v) {
    return String(v ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  function formatAmPm(date) {
    if (!date) return '-';
    let h = date.getHours();
    let m = date.getMinutes();
    let ampm = h >= 12 ? 'PM' : 'AM';
    h = h % 12;
    h = h ? h : 12;
    m = m < 10 ? '0' + m : m;
    return h + ':' + m + ' ' + ampm;
  }

  function isHoliday(d) { return HOLIDAY_OVERRIDES.has(fmtDate(d)); }
  function showError(msg) { els.errorBox.textContent = msg; els.errorBox.classList.remove('hidden'); }
  function clearError() { els.errorBox.classList.add('hidden'); els.errorBox.textContent = ''; }
  function setStatus(text, tone = '') { els.statusBadge.className = 'pill ' + tone; els.statusBadge.textContent = text; }
  function setFileName(text) { els.fileName.textContent = text || 'Ningún archivo seleccionado'; }

  function getHorario(nombreRaw, dateObj, primeraEntrada) {
    const k = normalizeSimple(nombreRaw);
    let esTurnoNoche = false;

    if (primeraEntrada) {
      const h = primeraEntrada.getHours();
      if (h >= 14 || h === 0 || h === 1 || h === 2) {
        esTurnoNoche = true;
      }
    } else {
      if (k.includes("pina") || k.includes("aguilar mesia") || k.includes("lamper")) esTurnoNoche = true;
    }

    if (esTurnoNoche) {
      if (k.includes("lamper")) {
        return { inH: 19, inM: 0, outH: 4, outM: 0, horasJornal: 8, isNightShift: true, lunchStartH: 0, lunchEndH: 1, diasLaborables: [1, 2, 3, 4, 5, 6] };
      }
      return { inH: 18, inM: 0, outH: 3, outM: 0, horasJornal: 8, isNightShift: true, lunchStartH: 23, lunchEndH: 0, diasLaborables: [1, 2, 3, 4, 5, 6] };
    }

    return { inH: 7, inM: 0, outH: 16, outM: 0, horasJornal: 8, isNightShift: false, lunchStartH: 12, lunchEndH: 13, diasLaborables: [1, 2, 3, 4, 5, 6] };
  }

  function getEventType(estado) {
    const k = normalizeSimple(estado);
    if (k.includes('entrada') || k.includes('ingreso') || k.includes('check in') || k.includes('checkin') || k === 'in') return 'IN';
    if (k.includes('salida') || k.includes('egreso') || k.includes('check out') || k.includes('checkout') || k === 'out') return 'OUT';
    return '';
  }

  function parseExcelDate(value) {
    if (value instanceof Date && !isNaN(value)) return value;
    if (typeof value === 'number') {
      const p = XLSX.SSF.parse_date_code(value);
      if (p) return new Date(p.y, p.m - 1, p.d, p.H || 0, p.M || 0, p.S || 0);
    }
    const t = String(value ?? '').trim();
    if (!t) return null;
    const isoLike = t.match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
    if (isoLike) return new Date(Number(isoLike[1]), Number(isoLike[2]) - 1, Number(isoLike[3]), Number(isoLike[4] || 0), Number(isoLike[5] || 0), Number(isoLike[6] || 0));
    const dmy = t.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
    if (dmy) {
      const y = dmy[3].length === 2 ? Number('20' + dmy[3]) : Number(dmy[3]);
      return new Date(y, Number(dmy[2]) - 1, Number(dmy[1]), Number(dmy[4] || 0), Number(dmy[5] || 0), Number(dmy[6] || 0));
    }
    const direct = new Date(t);
    return isNaN(direct) ? null : direct;
  }

  function detectColumns(headers) {
    const found = { idPersona: '', nombre: '', hora: '', estado: '' };
    headers.forEach(h => {
      const k = normalizeSimple(h);
      if (!found.idPersona && ['id de persona','id persona','codigo de persona','codigo persona','id trabajador','id empleado','legajo','codigo','id'].includes(k)) found.idPersona = h;
      if (!found.nombre && ['nombre','colaborador','trabajador','empleado','personal','name','nombre completo'].includes(k)) found.nombre = h;
      if (!found.hora && ['hora','fecha y hora','fecha hora','datetime','marcacion','marcacion completa','fecha','time'].includes(k)) found.hora = h;
      if (!found.estado && ['estado de asistencia','estado','tipo','movimiento','attendance status'].includes(k)) found.estado = h;
    });
    return found;
  }

  function findWorkerInMaestro(excelName) {
    const normExcel = normalizeSimple(excelName);
    if (normExcel.includes("miguel garcia serrano")) return null;

    if (MAESTRO_PERSONAL[normExcel]) return { key: normExcel, data: MAESTRO_PERSONAL[normExcel] };

    const excelWords = normExcel.split(' ');
    let bestMatch = null;
    let bestScore = 0;

    for (const dictKey in MAESTRO_PERSONAL) {
      const dictWords = dictKey.split(' ');
      let matches = 0;
      for (const w of excelWords) {
        if (w.length > 2 && dictWords.includes(w)) matches++;
      }
      if (matches > bestScore) {
        bestScore = matches;
        bestMatch = { key: dictKey, data: MAESTRO_PERSONAL[dictKey] };
      }
    }

    if (bestScore >= 3) return bestMatch;
    if (bestScore === 2 && excelWords.length <= 2) return bestMatch;

    return null;
  }

  function formatName(excelName, dictKey) {
    const dictWords = dictKey.split(' ');
    const excelWords = String(excelName || '').trim().split(/\s+/);
    let result = [];
    let used = new Set();

    for (const dWord of dictWords) {
      let bestMatch = null;
      let bestExcelIdx = -1;
      for (let i = 0; i < excelWords.length; i++) {
        if (used.has(i)) continue;
        const eNorm = normalizeSimple(excelWords[i]);
        if (eNorm === dWord || eNorm.includes(dWord) || dWord.includes(eNorm)) {
          bestMatch = excelWords[i];
          bestExcelIdx = i;
          break;
        }
      }
      if (bestMatch) {
        result.push(bestMatch);
        used.add(bestExcelIdx);
      } else {
        result.push(dWord.charAt(0).toUpperCase() + dWord.slice(1));
      }
    }

    for (let i = 0; i < excelWords.length; i++) {
      if (!used.has(i)) result.push(excelWords[i]);
    }
    return result.join(' ');
  }

  function getWorkerKey(row) {
    return String(row.idPersona || '').trim() || normalizeSimple(row.nombre);
  }

  function buildRows(rows, map) {
    if (!map.nombre || !map.hora || !map.estado) {
      throw new Error('No se detectaron bien las columnas clave. Revisa tu Excel.');
    }

    return rows.map((row, i) => {
      const dt = parseExcelDate(row[map.hora]);
      const nombreRaw = String(row[map.nombre] ?? '').trim();
      const estado = String(row[map.estado] ?? '').trim();
      const idPersona = map.idPersona ? String(row[map.idPersona] ?? '').trim() : '';
      const tipo = getEventType(estado);

      if (!dt || !nombreRaw || !estado) return null;

      const match = findWorkerInMaestro(nombreRaw);
      if (!match) return null;

      const infoPersonal = match.data;
      const nombre = formatName(nombreRaw, match.key);

      const workerKey = String(idPersona || '').trim() || normalizeSimple(nombre);

      let fechaLogica = dayOnly(dt);
      const h = dt.getHours();
      if (h >= 0 && h < 6) {
        fechaLogica.setDate(fechaLogica.getDate() - 1);
      }

      return {
        workerKey, idPersona, nombre, hora: dt, fecha: fechaLogica, estado, tipo,
        empresa: infoPersonal.empresa, contrato: infoPersonal.contrato, fila: i + 2
      };
    }).filter(Boolean).sort((a, b) => a.hora - b.hora || a.fila - b.fila);
  }

  function calcularHorasConTolerancia(minutosTotales) {
    const horasEnteras = Math.floor(minutosTotales / 60);
    const minutosSueltos = minutosTotales % 60;
    let hsToleradas = horasEnteras;

    if (minutosSueltos >= 55) {
      hsToleradas += 1.0;
    } else if (minutosSueltos >= 45) {
      hsToleradas += 0.5;
    }

    return hsToleradas;
  }

  function calcWorkedHoursFromAnalysis(analysis, fecha, nombreRaw) {
    if (!analysis.sessions.length) {
      return { brutas: null, netas: null, decimal: null, jornal: 0, hs: 0, dominical: 0 };
    }

    const horario = getHorario(nombreRaw, fecha, analysis.primeraEntrada);
    const isDomingoFeriado = isSunday(fecha) || isHoliday(fecha);

    const horaEntradaOficial = new Date(fecha);
    horaEntradaOficial.setHours(horario.inH, horario.inM, 0, 0);

    const horaSalidaOficial = new Date(fecha);
    horaSalidaOficial.setHours(horario.outH, horario.outM, 0, 0);
    if (horario.isNightShift && horario.outH <= horario.inH) {
      horaSalidaOficial.setDate(horaSalidaOficial.getDate() + 1);
    }

    const lunchStart = new Date(fecha);
    lunchStart.setHours(horario.lunchStartH, 0, 0, 0);
    if (horario.isNightShift && horario.lunchStartH < 12) {
      lunchStart.setDate(lunchStart.getDate() + 1);
    }

    const lunchEnd = new Date(lunchStart);
    lunchEnd.setHours(horario.lunchEndH, 0, 0, 0);
    if (horario.lunchEndH <= horario.lunchStartH) {
      lunchEnd.setDate(lunchEnd.getDate() + 1);
    }

    let totalValidMinutes = 0;
    let lunchOverlap = 0;

    analysis.sessions.forEach(s => {
      let inStart = s.entrada.getTime();

      if (inStart < horaEntradaOficial.getTime()) {
        inStart = horaEntradaOficial.getTime();
      } else {
        const adjusted = new Date(inStart);
        const m = adjusted.getMinutes();

        if (m <= 15) {
          adjusted.setMinutes(0, 0, 0);
        } else if (m <= 35) {
          adjusted.setMinutes(30, 0, 0);
        } else {
          adjusted.setHours(adjusted.getHours() + 1, 0, 0, 0);
        }
        inStart = adjusted.getTime();
      }

      let outEnd = s.salida.getTime();

      if (outEnd > inStart) {
        totalValidMinutes += (outEnd - inStart) / 60000;
        let overlapStart = Math.max(inStart, lunchStart.getTime());
        let overlapEnd = Math.min(outEnd, lunchEnd.getTime());
        if (overlapEnd > overlapStart) {
          lunchOverlap += (overlapEnd - overlapStart) / 60000;
        }
      }
    });

    let horaTopeCorrido;
    const isWeekend = fecha.getDay() === 0 || fecha.getDay() === 6;

    if (isWeekend && !horario.isNightShift) {
      horaTopeCorrido = new Date(fecha);
      horaTopeCorrido.setHours(14, 35, 0, 0);
    } else {
      horaTopeCorrido = new Date(lunchEnd);
      horaTopeCorrido.setMinutes(horaTopeCorrido.getMinutes() + 30);
    }

    if (analysis.ultimaSalida && analysis.ultimaSalida.getTime() >= horaTopeCorrido.getTime()) {
      if (lunchOverlap > 0) {
        totalValidMinutes -= lunchOverlap;
      }
    }

    const minutosObjetivoJornal = horario.horasJornal * 60;
    let jornalFinal = 0;
    let minExtrasTotales = 0;

    if (totalValidMinutes >= (minutosObjetivoJornal - 10)) {
      jornalFinal = horario.horasJornal;
      minExtrasTotales = Math.max(0, totalValidMinutes - minutosObjetivoJornal);
    } else {
      jornalFinal = roundDownHalf(totalValidMinutes / 60);
      minExtrasTotales = 0;
    }

    let hsFinal = 0;
    if (minExtrasTotales > 0) {
      hsFinal = calcularHorasConTolerancia(minExtrasTotales);
      if (hsFinal > SETTINGS.topeHsDiario) hsFinal = SETTINGS.topeHsDiario;
    }

    let dominicalFinal = 0;
    if (isDomingoFeriado) {
      dominicalFinal = jornalFinal + hsFinal;
    }

    return {
      brutas: analysis.totalBruto,
      netas: analysis.totalNeto,
      decimal: analysis.totalNeto,
      jornal: jornalFinal,
      hs: hsFinal,
      dominical: dominicalFinal
    };
  }

  function buildSummary(cleanRows, startDate, endDate) {
    const workersRecords = new Map();
    cleanRows.forEach(r => {
      const k = r.workerKey;
      if(!workersRecords.has(k)) workersRecords.set(k, { info: r, rows: [] });
      workersRecords.get(k).rows.push(r);
    });

    const allSummaryRows = [];
    const workersMap = new Map();

    workersRecords.forEach((workerData, workerKey) => {
      const { info, rows } = workerData;
      workersMap.set(workerKey, { workerKey, idPersona: info.idPersona, nombre: info.nombre, empresa: info.empresa, contrato: info.contrato, order: info.fila });

      rows.sort((a, b) => a.hora - b.hora);

      const dedupe = [];
      for (const r of rows) {
        const prev = dedupe[dedupe.length - 1];
        if (prev && prev.tipo === r.tipo && diffMinutes(prev.hora, r.hora) <= SETTINGS.dedupeMinutes) continue;
        dedupe.push(r);
      }

      const dailyData = new Map();
      const getDayData = (dStamp) => {
        if (!dailyData.has(dStamp)) dailyData.set(dStamp, { sessions: [], anomalies: [], rawCount: 0, date: new Date(dStamp), rawRecords: [] });
        return dailyData.get(dStamp);
      };

      let openEntry = null;
      for (const r of dedupe) {
        const dStamp = stamp(r.fecha);
        const dData = getDayData(dStamp);
        dData.rawRecords.push(r);
        dData.rawCount++;

        if (r.tipo === 'IN') {
          if (openEntry) {
            getDayData(stamp(openEntry.fecha)).anomalies.push(`Entrada sin salida ${fmtTime(openEntry.hora)}`);
          }
          openEntry = r;
        } else if (r.tipo === 'OUT') {
          if (!openEntry) {
            dData.anomalies.push(`Salida sin entrada ${fmtTime(r.hora)}`);
          } else if (r.hora <= openEntry.hora) {
            getDayData(stamp(r.fecha)).anomalies.push(`Salida antes de entrada ${fmtTime(r.hora)}`);
            openEntry = null;
          } else {
            getDayData(stamp(openEntry.fecha)).sessions.push({ entrada: openEntry.hora, salida: r.hora, horas: hoursBetween(openEntry.hora, r.hora) });
            openEntry = null;
          }
        }
      }
      if (openEntry) {
        getDayData(stamp(openEntry.fecha)).anomalies.push(`Entrada sin salida ${fmtTime(openEntry.hora)}`);
      }

      dailyData.forEach((data, dStamp) => {
        const fecha = data.date;
        data.rawRecords.sort((a, b) => a.hora - b.hora);

        let primeraEntrada = data.sessions[0]?.entrada || null;
        let ultimaSalida = data.sessions[data.sessions.length - 1]?.salida || null;

        if (data.sessions.length === 0) {
          const inRec = data.rawRecords.find(x => x.tipo === 'IN');
          const outRec = [...data.rawRecords].reverse().find(x => x.tipo === 'OUT');
          if (inRec) primeraEntrada = inRec.hora;
          if (outRec) ultimaSalida = outRec.hora;
        }

        const horario = getHorario(info.nombre, fecha, primeraEntrada);

        const analysisObj = {
          sessions: data.sessions,
          anomalies: data.anomalies,
          primeraEntrada,
          ultimaSalida,
          rawCount: data.rawCount,
          totalBruto: data.sessions.reduce((acc, s) => acc + s.horas, 0),
          totalNeto: data.sessions.reduce((acc, s) => acc + s.horas, 0)
        };

        const calc = calcWorkedHoursFromAnalysis(analysisObj, fecha, info.nombre);

        let alerta = '';
        const sinMarcas = data.rawCount === 0;
        const marcaIncompleta = data.rawCount > 0 && data.sessions.length === 0;
        const isLaborable = horario.diasLaborables.includes(fecha.getDay());

        let esTardanza = false;
        if (primeraEntrada) {
          const horaTope = new Date(fecha);
          horaTope.setHours(horario.inH, horario.inM + 15, 0, 0);
          if (horario.isNightShift && horario.inH < 12) horaTope.setDate(horaTope.setDate() + 1);
          if (primeraEntrada > horaTope) esTardanza = true;
        }

        if (sinMarcas && isLaborable && !isHoliday(fecha)) { alerta = 'F'; }
        else if (marcaIncompleta) { alerta = 'I'; }
        else if (data.anomalies.length > 0 && esTardanza) { alerta = 'T-REV'; }
        else if (data.anomalies.length > 0) { alerta = 'REV'; }
        else if (esTardanza) { alerta = 'T'; }

        allSummaryRows.push({
          workerKey: info.workerKey, idPersona: info.idPersona, nombre: info.nombre, fecha,
          empresa: info.empresa, contrato: info.contrato,
          primeraEntrada, ultimaSalida,
          horasBrutas: calc.brutas, horasNetas: calc.netas, horasDecimal: calc.decimal,
          jornal: calc.jornal, hs: calc.hs, dominical: calc.dominical,
          anomalias: data.anomalies.length, detalle: data.anomalies.join(' | '),
          order: info.fila, feriado: isHoliday(fecha), alerta
        });
      });
    });

    const startStamp = stamp(startDate);
    const endStamp = stamp(endDate);
    const summaryRows = allSummaryRows.filter(r => stamp(r.fecha) >= startStamp && stamp(r.fecha) <= endStamp);

    const workers = [...workersMap.values()].sort((a,b) => a.nombre.localeCompare(b.nombre));

    summaryRows.sort((a, b) => {
      if (a.nombre !== b.nombre) return a.nombre.localeCompare(b.nombre);
      return a.fecha - b.fecha;
    });

    return { summaryRows, effectiveRows: cleanRows, workers };
  }

  function buildTareo(summaryRows, workers, startDate, endDate) {
    const dates = [];
    for (let d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) { dates.push(new Date(d)); }
    const idx = new Map(summaryRows.map(r => [`${r.workerKey}__${fmtDate(r.fecha)}`, r]));

    const rows = workers.map(w => {
      const out = {
        workerKey: w.workerKey, idPersona: w.idPersona, nombre: w.nombre,
        empresa: w.empresa, contrato: w.contrato,
        dias: {}, hNormal: 0, hExtras: 0, dominical: 0
      };

      dates.forEach(d => {
        const k = fmtDate(d);
        const rec = idx.get(`${w.workerKey}__${k}`);
        const feriado = isHoliday(d) || d.getDay() === 0;
        const horario = getHorario(w.nombre, d, null);
        const isLaborable = horario.diasLaborables.includes(d.getDay());

        const sinMarcaDiaLaborable = !rec && isLaborable && !feriado;
        let alerta = rec?.alerta || (sinMarcaDiaLaborable ? 'F' : '');
        const jornal = rec?.jornal ?? 0;
        const hs = rec?.hs ?? 0;
        const dominical = rec?.dominical ?? 0;

        out.dias[k] = { jornal, hs, dominical, feriado, alerta, detalle: rec?.detalle || '' };

        if (feriado) {
          out.dominical += dominical;
        } else {
          out.hNormal += jornal;
          out.hExtras += hs;
        }
      });
      return out;
    });
    return { dates, rows };
  }

  function getFilteredData() {
    const emp = els.filterEmpresa.value;
    const con = els.filterContrato.value;

    const fWorkers = state.workers.filter(w =>
        (emp === 'ALL' || w.empresa === emp) &&
        (con === 'ALL' || w.contrato === con)
    );

    const validKeys = new Set(fWorkers.map(w => w.workerKey));

    return {
      summaryRows: state.summaryRows.filter(r => validKeys.has(r.workerKey)),
      cleanRows: state.cleanRows.filter(r => validKeys.has(r.workerKey)),
      anomalyRows: state.anomalyRows.filter(r => validKeys.has(r.workerKey)),
      tareoRows: state.tareo.rows.filter(r => validKeys.has(r.workerKey)),
      workers: fWorkers
    };
  }

  function populateFilters() {
    const empresas = [...new Set(state.workers.map(w => w.empresa))].sort();
    const contratos = [...new Set(state.workers.map(w => w.contrato))].sort();

    els.filterEmpresa.innerHTML = `<option value="ALL">Todas las Empresas</option>` + empresas.map(v => `<option value="${v}">${v}</option>`).join('');
    els.filterContrato.innerHTML = `<option value="ALL">Tipo de Contratación</option>` + contratos.map(v => `<option value="${v}">${v}</option>`).join('');

    els.filterBar.style.display = 'flex';
  }

  function handleRowsLoaded(rows) {
    state.rawRows = rows;
    state.map = detectColumns(Object.keys(rows[0] || {}));
    state.cleanRows = buildRows(rows, state.map);

    if (!state.cleanRows.length) throw new Error('No se encontraron de marcas correspondientes a obreros registrados.');

    const ds = state.cleanRows.map(r => stamp(r.fecha));
    els.startDate.value = fmtInput(new Date(Math.min(...ds)));
    els.endDate.value = fmtInput(new Date(Math.max(...ds)));

    processNow();
  }

  function renderTable(container, headers, rows, options = {}) {
    const nameColIndex = Number.isInteger(options.nameColIndex) ? options.nameColIndex : 0;
    container.innerHTML =
        '<table><thead><tr>' +
        headers.map((h, i) => `<th class="${i === 0 ? 'name-col' : ''}">${escapeHtml(h)}</th>`).join('') +
        '</tr></thead><tbody>' +
        rows.map(r =>
            '<tr>' +
            r.map((v, i) => {
              const value = v == null ? '' : v;
              const isNum = typeof value === 'number';
              let className = [ i === 0 ? 'name-col' : '', isNum ? 'num' : '' ];

              if (i === nameColIndex) {
                const contratoStr = String(r[options.contratoStrIndex] ?? '').toUpperCase();
                className.push(contratoStr.includes('PLANILLA') ? 'planilla-name' : 'rrhh-name');
              }

              if (options.renderTags && i === options.tagIndex && value !== '') {
                return `<td class="${className.join(' ').trim()}">${value}</td>`;
              }

              return `<td class="${className.join(' ').trim()}">${escapeHtml(isNum ? displayNum(value) : value)}</td>`;
            }).join('') +
            '</tr>'
        ).join('') +
        '</tbody></table>';
  }

  function isLaborable(dateObj, nombreRaw) {
    const horario = getHorario(nombreRaw, dateObj, null);
    return horario.diasLaborables.includes(dateObj.getDay()) && !isHoliday(dateObj);
  }

  // --- RECALCULAR LOS ENCABEZADOS PEGADOS AL HACER RENDERIZADO O AMPLIAR ---
  function reflowStickyHeaders() {
    requestAnimationFrame(() => {
      [els.tareo, els.comparativa].forEach(container => {
        const theadFirstRow = container.querySelector('thead tr:first-child');
        if (theadFirstRow) {
          const offset = theadFirstRow.offsetHeight;
          container.querySelectorAll('thead tr:nth-child(2) th').forEach(th => {
            th.style.top = (offset - 1) + 'px';
            th.style.zIndex = '9';
          });
        }
      });
    });
  }

  function renderAll() {
    const fd = getFilteredData();

    els.statWorkers.textContent = fd.workers.length;
    els.statSummary.textContent = fd.summaryRows.length;

    renderTable(els.resumen,
        ['Nombre','ID','Fecha','Día','Asignación','Primera entrada','Última salida','Jornal','H.S.','Alerta','Detalle'],
        fd.summaryRows.slice(0, 2000).map(r => [
          r.nombre, r.idPersona || '', fmtDate(r.fecha), weekdayName(r.fecha),
          `<span class="meta-tag">${r.empresa}</span><span class="meta-tag">${r.contrato}</span>`,
          r.primeraEntrada ? fmtTime(r.primeraEntrada) : '', r.ultimaSalida ? fmtTime(r.ultimaSalida) : '',
          displayNum(r.jornal), displayNum(r.hs), r.alerta, r.detalle
        ]),
        { nameColIndex: 0, contratoStrIndex: 4, renderTags: true, tagIndex: 4 }
    );

    renderTable(els.bd,
        ['Nombre','ID','Fecha','Día','Hora','Estado','Tipo','Empresa','Contrato'],
        fd.cleanRows.slice(0, 3000).map(r => [
          r.nombre, r.idPersona || '', fmtDate(r.fecha), weekdayName(r.fecha), fmtTime(r.hora), r.estado, r.tipo || '', r.empresa, r.contrato
        ]),
        { nameColIndex: 0, contratoStrIndex: 8 }
    );

    renderTable(els.anomalias,
        ['Nombre','ID','Fecha','Día','Asignación','Alerta','Detalle','Jornal','H.S.'],
        fd.anomalyRows.slice(0, 3000).map(r => [
          r.nombre, r.idPersona || '', fmtDate(r.fecha), weekdayName(r.fecha), `${r.empresa} - ${r.contrato}`, r.alerta, r.detalle, displayNum(r.jornal), displayNum(r.hs)
        ]),
        { nameColIndex: 0, contratoStrIndex: 4 }
    );

    let html = '<table><thead><tr>';
    html += '<th class="name-col" rowspan="2">Colaboradores</th>';
    html += state.tareo.dates.map(d => `<th colspan="2" class="day-head day-start">${escapeHtml(fmtDate(d))}<span class="day-name">${escapeHtml(weekdayName(d))}</span></th>`).join('');
    html += '<th rowspan="2" class="totals-head totals-start">H.NORMAL</th><th rowspan="2" class="totals-head">H.EXTRAS</th><th rowspan="2" class="totals-head">DOMINICAL</th></tr><tr>';
    html += state.tareo.dates.map((d, idx) => `<th class="sub-day-start ${idx === 0 ? 'day-start' : ''}">JORNAL</th><th>H.S.</th>`).join('');
    html += '</tr></thead><tbody>';

    fd.tareoRows.forEach(r => {
      const isPlanilla = r.contrato.toUpperCase().includes('PLANILLA');
      const nameClass = isPlanilla ? 'planilla-name' : 'rrhh-name';
      html += `<tr><td class="name-col ${nameClass}"><div>${escapeHtml(r.nombre)}</div><div style="font-size:0.7rem;color:var(--muted);font-weight:600;">${r.empresa} / ${r.contrato}</div></td>`;

      state.tareo.dates.forEach((d, idx) => {
        const k = fmtDate(d);
        const v = r.dias[k] || { jornal: 0, hs: 0, feriado: false, alerta: '', detalle: '' };

        let journalText = displayNum(v.jornal);
        if (v.alerta === 'F') journalText = 'F';
        if (v.alerta === 'I') journalText = 'I';

        const hsText = (v.alerta === 'F' || v.alerta === 'I' || v.hs === 0) ? '' : displayNum(v.hs);

        let cellClass = '';
        if (v.alerta === 'F') cellClass = 'alerta-f';
        else if (v.alerta === 'I') cellClass = 'alerta-i';
        else if (v.alerta === 'T') cellClass = 'alerta-t';
        else if (v.alerta === 'REV') cellClass = 'alerta-rev';
        else if (v.alerta === 'T-REV') cellClass = 'alerta-trev';

        const titleText = v.detalle ||
            (v.alerta === 'F' ? 'Sin marcaciones (Falta)' :
                (v.alerta === 'I' ? 'Marcación incompleta' :
                    (v.alerta === 'T' ? 'Tardanza' :
                        (v.alerta === 'T-REV' ? 'Tardanza + Anomalías a revisar' :
                            (v.alerta === 'REV' ? 'Anomalía / Duplicado a revisar' : '')))));

        const title = escapeHtml(titleText);
        html += `<td class="num ${idx === 0 ? 'day-start sub-day-start' : ''} ${cellClass}" title="${title}">${escapeHtml(journalText)}</td>`;
        html += `<td class="num ${cellClass}" title="${title}">${escapeHtml(hsText)}</td>`;
      });
      html += `<td class="num totals-cell totals-start">${escapeHtml(displayNum(r.hNormal))}</td>`;
      html += `<td class="num totals-cell">${escapeHtml(displayNum(r.hExtras))}</td>`;
      html += `<td class="num totals-cell">${escapeHtml(displayNum(r.dominical))}</td></tr>`;
    });

    html += '</tbody></table>';
    els.tareo.innerHTML = html;

    const compByWorker = {};
    fd.summaryRows.forEach(r => {
      if(!compByWorker[r.workerKey]) compByWorker[r.workerKey] = {};
      compByWorker[r.workerKey][fmtDate(r.fecha)] = { in: r.primeraEntrada, out: r.ultimaSalida, alerta: r.alerta };
    });

    let htmlComp = '<table><thead><tr><th class="name-col" rowspan="2">Colaboradores</th>';
    htmlComp += state.tareo.dates.map(d => `<th colspan="2" class="day-head day-start">${escapeHtml(fmtDate(d))}<span class="day-name">${escapeHtml(weekdayName(d))}</span></th>`).join('');
    htmlComp += '</tr><tr>';

    // ANCHO ESTRICTO DE COLUMNAS PARA TAREO POR HORA (min-width: 90px y max-width)
    htmlComp += state.tareo.dates.map((d, idx) => `<th class="sub-day-start ${idx === 0 ? 'day-start' : ''}" style="min-width:90px; width:100px;">INGRESO</th><th style="min-width:90px; width:100px;">SALIDA</th>`).join('');
    htmlComp += '</tr></thead><tbody>';

    fd.workers.forEach(w => {
      const isPlanilla = w.contrato.toUpperCase().includes('PLANILLA');
      const nameClass = isPlanilla ? 'planilla-name' : 'rrhh-name';
      htmlComp += `<tr><td class="name-col ${nameClass}"><div>${escapeHtml(w.nombre)}</div><div style="font-size:0.7rem;color:var(--muted);font-weight:600;">${w.empresa}</div></td>`;

      state.tareo.dates.forEach((d, idx) => {
        const k = fmtDate(d);
        const marks = (compByWorker[w.workerKey] && compByWorker[w.workerKey][k]) || { in: null, out: null, alerta: 'F' };
        const inStr = marks.in ? formatAmPm(marks.in) : '-';
        const outStr = marks.out ? formatAmPm(marks.out) : '-';

        let cellClass = '';
        if (marks.alerta === 'F' && isLaborable(d, w.nombre)) cellClass = 'alerta-f';
        else if (marks.alerta === 'I') cellClass = 'alerta-i';
        else if (marks.alerta === 'T') cellClass = 'alerta-t';

        htmlComp += `<td class="num ${idx === 0 ? 'day-start sub-day-start' : ''} ${cellClass}">${inStr}</td>`;
        htmlComp += `<td class="num ${cellClass}">${outStr}</td>`;
      });
      htmlComp += `</tr>`;
    });
    htmlComp += '</tbody></table>';
    els.comparativa.innerHTML = htmlComp;

    // Recalcular posiciones (Sticky Headers)
    reflowStickyHeaders();
  }

  function updateStats(sourceRows) {
    els.statRows.textContent = sourceRows.length;
    els.statDays.textContent = state.tareo.dates.length;
  }

  function buildTareoSheetData(filteredTareoRows) {
    const row1 = ['Colaboradores'];
    const row2 = [''];
    state.tareo.dates.forEach(d => { row1.push(`${fmtDate(d)} ${weekdayName(d)}`, ''); row2.push('JORNAL', 'H.S.'); });
    row1.push('H.NORMAL', 'H.EXTRAS', 'DOMINICAL');
    row2.push('', '', '');

    const data = [row1, row2];
    filteredTareoRows.forEach(r => {
      const row = [`${r.nombre} [${r.empresa} - ${r.contrato}]`];
      state.tareo.dates.forEach(d => {
        const k = fmtDate(d);
        const cell = r.dias[k] || { jornal: 0, hs: 0, alerta: '' };

        if (cell.alerta === 'F') { row.push('F', ''); }
        else if (cell.alerta === 'I') { row.push('I', ''); }
        else {
          row.push(cell.jornal ?? 0);
          row.push(cell.hs ?? 0);
        }
      });
      row.push(r.hNormal, r.hExtras, r.dominical);
      data.push(row);
    });
    return data;
  }

  function workbookOut() {
    const fd = getFilteredData();
    const wb = XLSX.utils.book_new();
    const tareoData = buildTareoSheetData(fd.tareoRows);
    const wsTareo = XLSX.utils.aoa_to_sheet(tareoData);
    wsTareo['!merges'] = [];

    let c = 1;
    state.tareo.dates.forEach(() => { wsTareo['!merges'].push({ s: { r: 0, c }, e: { r: 0, c: c + 1 } }); c += 2; });
    wsTareo['!merges'].push({ s: { r: 0, c: 0 }, e: { r: 1, c: 0 } });

    const start = 1 + state.tareo.dates.length * 2;
    wsTareo['!merges'].push({ s: { r: 0, c: start }, e: { r: 1, c: start } });
    wsTareo['!merges'].push({ s: { r: 0, c: start + 1 }, e: { r: 1, c: start + 1 } });
    wsTareo['!merges'].push({ s: { r: 0, c: start + 2 }, e: { r: 1, c: start + 2 } });

    for (let r = 2; r < tareoData.length; r++) {
      for (let col = 1; col < tareoData[r].length; col++) {
        const ref = XLSX.utils.encode_cell({ r, c: col });
        if (wsTareo[ref] && typeof wsTareo[ref].v === 'number') wsTareo[ref].z = '0.##';
      }
    }

    wsTareo['!cols'] = [{ wch: 45 }, ...Array(tareoData[0].length - 1).fill({ wch: 10 })];
    XLSX.utils.book_append_sheet(wb, wsTareo, 'TAREO');

    const resumenData = [
      ['Empresa','Contrato','Nombre','ID','Fecha','Día','Primera entrada','Última salida','Jornal','H.S.','Alerta','Detalle'],
      ...fd.summaryRows.map(r => [
        r.empresa, r.contrato, r.nombre, r.idPersona || '', fmtDate(r.fecha), weekdayName(r.fecha),
        r.primeraEntrada ? fmtTime(r.primeraEntrada) : '', r.ultimaSalida ? fmtTime(r.ultimaSalida) : '',
        r.jornal ?? '', r.hs ?? '', r.alerta, r.detalle
      ])
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(resumenData), 'RESUMEN');

    const anomaliasData = [
      ['Empresa','Contrato','Nombre','ID','Fecha','Día','Alerta','Detalle','Jornal','H.S.'],
      ...fd.anomalyRows.map(r => [
        r.empresa, r.contrato, r.nombre, r.idPersona || '', fmtDate(r.fecha), weekdayName(r.fecha), r.alerta, r.detalle, r.jornal ?? '', r.hs ?? ''
      ])
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(anomaliasData), 'ANOMALIAS');

    return wb;
  }

  function processNow() {
    clearError();
    try {
      if (!state.rawRows.length) throw new Error('Primero carga un archivo o un ejemplo.');
      const startDate = els.startDate.value ? new Date(els.startDate.value + 'T00:00:00') : new Date();
      const endDate = els.endDate.value ? new Date(els.endDate.value + 'T00:00:00') : new Date();
      if (startDate > endDate) throw new Error('La fecha de inicio no puede ser mayor que la fecha fin.');

      const { summaryRows, effectiveRows, workers } = buildSummary(state.cleanRows, startDate, endDate);
      state.summaryRows = summaryRows;
      state.anomalyRows = summaryRows.filter(r => r.alerta || r.anomalias > 0);
      state.workers = workers;
      state.tareo = buildTareo(summaryRows, workers, startDate, endDate);

      populateFilters();
      renderAll();
      updateStats(effectiveRows);

      setStatus(`Procesado correctamente`, 'ok');
      els.downloadXlsxBtn.disabled = !summaryRows.length;
      els.downloadCsvBtn.disabled = !summaryRows.length;
    } catch (err) {
      setStatus('Ocurrió un error', 'error');
      showError(err.message || 'No se pudo procesar el archivo.');
    }
  }

  function handleWorkbook(file) {
    clearError();
    setFileName(file?.name || 'Ningún archivo seleccionado');
    const reader = new FileReader();
    reader.onload = e => {
      try {
        state.workbook = XLSX.read(e.target.result, { type: 'array', cellDates: true, raw: true });

        state.fileName = (file.name || 'tareo').replace(/\.(xlsx|xls|csv)$/i, '') + '_filtrado.xlsx';

        const name = state.workbook.SheetNames[0];
        const ws = state.workbook.Sheets[name];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: true, cellDates: true });
        if (!rows.length) throw new Error('La hoja seleccionada no tiene datos.');
        handleRowsLoaded(rows);
      } catch (err) {
        setStatus('No se pudo leer el archivo', 'error');
        showError(err.message || 'Revisa si el archivo está dañado o no es válido.');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function loadSample() {
    const rows = [
      { 'ID de persona': '101', Nombre: 'Jose Lamper Torres', Hora: '15/05/2026 19:00', 'Estado de asistencia': 'Registro de entrada' },
      { 'ID de persona': '101', Nombre: 'Jose Lamper Torres', Hora: '16/05/2026 04:00', 'Estado de asistencia': 'Registro de salida' },
      { 'ID de persona': '102', Nombre: 'Abraham Cipiran Ramirez', Hora: '01/04/2026 07:20', 'Estado de asistencia': 'Registro de entrada' },
      { 'ID de persona': '102', Nombre: 'Abraham Cipiran Ramirez', Hora: '01/04/2026 16:00', 'Estado de asistencia': 'Registro de salida' }
    ];
    state.fileName = 'tareo_tsm_test.xlsx';
    setFileName('Ejemplo Cargado');
    handleRowsLoaded(rows);
  }

  function downloadXlsx() {
    try { XLSX.writeFile(workbookOut(), state.fileName); }
    catch (err) { showError('No se pudo generar el Excel de salida.'); }
  }

  function downloadCsv() {
    const fd = getFilteredData();
    const blob = new Blob([XLSX.utils.sheet_to_csv(XLSX.utils.aoa_to_sheet(buildTareoSheetData(fd.tareoRows)))], { type: 'text/csv;charset=utf-8;' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'tareo.csv';
    a.click();
    URL.revokeObjectURL(a.href);
  }

  els.filterEmpresa.addEventListener('change', renderAll);
  els.filterContrato.addEventListener('change', renderAll);

  els.clearFiltersBtn.addEventListener('click', () => {
    els.filterEmpresa.value = 'ALL';
    els.filterContrato.value = 'ALL';
    renderAll();
  });

  // --- EVENTO PANTALLA COMPLETA ---
  els.fullscreenBtn.addEventListener('click', () => {
    document.body.classList.toggle('fullscreen-mode');
    const isFull = document.body.classList.contains('fullscreen-mode');
    els.fullscreenBtn.innerHTML = isFull
        ? `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M8 3v3a2 2 0 0 1-2 2H3m18 0h-3a2 2 0 0 1-2-2V3m0 18v-3a2 2 0 0 1 2-2h3M3 16h3a2 2 0 0 1 2 2v3"/></svg> Reducir`
        : `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M8 3H5a2 2 0 0 0-2 2v3m18 0V5a2 2 0 0 0-2-2h-3m0 18h3a2 2 0 0 0 2-2v-3M3 16v3a2 2 0 0 0 2 2h3"/></svg> Ampliar`;

    // Forzamos recalcular la posición de los headers al cambiar la ventana
    reflowStickyHeaders();
  });

  els.themeBtn.addEventListener('change', e => {
    document.documentElement.setAttribute('data-theme', e.target.checked ? 'dark' : 'light');
  });
  els.fileInput.addEventListener('change', e => { const f = e.target.files?.[0]; if (f) handleWorkbook(f); });
  els.processBtn.addEventListener('click', processNow);
  els.downloadXlsxBtn.addEventListener('click', downloadXlsx);
  els.downloadCsvBtn.addEventListener('click', downloadCsv);
  els.sampleBtn.addEventListener('click', loadSample);

  ['dragenter','dragover'].forEach(ev => els.dropzone.addEventListener(ev, e => { e.preventDefault(); els.dropzone.classList.add('drag'); }));
  ['dragleave','drop'].forEach(ev => els.dropzone.addEventListener(ev, e => { e.preventDefault(); els.dropzone.classList.remove('drag'); }));
  els.dropzone.addEventListener('drop', e => { const f = e.dataTransfer.files?.[0]; if (f) handleWorkbook(f); });

  document.querySelectorAll('.tab').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab').forEach(x => x.classList.remove('active'));
      btn.classList.add('active');
      ['resumen','tareo','bd','anomalias','comparativa'].forEach(k => document.getElementById('tab-' + k).classList.add('hidden'));
      document.getElementById('tab-' + btn.dataset.tab).classList.remove('hidden');
      reflowStickyHeaders();
    });
  });
})();