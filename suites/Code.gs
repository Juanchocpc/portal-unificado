/**
 * ==================================================================
 * 🚀 SISTEMA INTEGRAL: DASHBOARD CRM + SÁBANA MENSUAL + BONUS V2
 * ==================================================================
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🚀 Análisis CRM")
    .addItem("📊 Ver Dashboard Completo", "mostrarDashboardRebotes")
    .addItem("📅 Ver Panel Mensual (Histórico)", "mostrarPanelMensual")
    .addItem("💰 Calcular Bonus", "mostrarModalBonus") 
    .addSeparator()
    .addItem("💾 Guardar Historial del Día", "abrirModalHistorial")
    .addToUi();
}

// --- MODALES UI ---

function abrirModalHistorial() {
  const html = HtmlService.createHtmlOutputFromFile('DateSelector')
      .setTitle('📅 Guardar Historial')
      .setWidth(1000) 
      .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, 'Cierre de Día');
}

function mostrarPanelMensual() {
  const html = HtmlService.createHtmlOutputFromFile('MonthlyPanel')
      .setTitle('Panel de Ratios Mensuales')
      .setWidth(1900) // Aumentado para mejor visualización
      .setHeight(1000);
  SpreadsheetApp.getUi().showModalDialog(html, 'Panel de Ratios Mensuales');
}

function mostrarModalBonus() {
  mostrarPanelMensual();
}

function mostrarDashboardRebotes() {
  const html = HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Dashboard de Rendimiento')
      .setWidth(1300) 
      .setHeight(850);
  SpreadsheetApp.getUi().showModalDialog(html, 'Dashboard de Rendimiento Completo');
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index').setTitle('Dashboard de Rendimiento');
}

// ==================================================================
// 1. LÓGICA DE BONUS (ACTUALIZADA CON PONDERACIÓN Y CONFIG)
// ==================================================================

/**
 * Guarda la configuración de bonos en una hoja oculta "CONFIG_BONUS"
 */
function saveBonusConfig(config) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("CONFIG_BONUS");
  if (!sheet) {
    sheet = ss.insertSheet("CONFIG_BONUS");
    sheet.hideSheet();
  }
  sheet.clear();
  // Guardamos como JSON string en A1
  sheet.getRange(1, 1).setValue(JSON.stringify(config));
  return { exito: true };
}

/**
 * Lee la configuración de bonos
 */
function getBonusConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("CONFIG_BONUS");
  if (!sheet) return null;
  const val = sheet.getRange(1, 1).getValue();
  try {
    return JSON.parse(val);
  } catch (e) {
    return null;
  }
}

/**
 * Obtiene datos de bonus. Si ya existe hoja de respaldo, la carga.
 * Si no, calcula desde cero aplicando reglas de config.
 */
function getBonusData(mes, anio) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mesNombre = getMonthName(mes).toUpperCase();
  const sheetName = `BONUS-${mesNombre} ${anio}`;
  
  // Intentar cargar config para aplicar reglas en el cálculo
  const config = getBonusConfig() || { transfer: [], asistente: [] };

  const existingSheet = ss.getSheetByName(sheetName);
  
  if (existingSheet) {
    return readBonusSheet(existingSheet, sheetName, config);
  } else {
    return calculateBonusFromScratch(mes, anio, sheetName, config);
  }
}

function saveBonusData(sheetName, dataToSave) {
  return saveBonusState(sheetName, dataToSave);
}

function saveBonusState(sheetName, dataToSave) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.hideSheet();
  } else {
    sheet.clear();
  }

  // Nuevos encabezados según requerimiento
  const headers = [
    "Operador", "Extension", "Fecha Ingreso", "Hrs Reales", 
    "Llamadas", "Prod (Conv/Min)", "Ratio/Media", 
    "Bono Opta", "Indice Pond", "% Final", "Bono Aprobado"
  ];
  
  const rows = dataToSave.map(row => [
      row.operador,
      row.extension,
      row.fechaIngreso,
      row.horasReales,
      row.llamadas,
      row.produccion,   // Conv o Minutos
      row.ratioMedia,   // Valor % o entero
      row.bonoOpta,
      row.indicePond,
      row.porcFinal,
      row.bonoAprobado
  ]);

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
       .setFontWeight("bold").setBackground("#cfe2f3");
  
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  return { exito: true, mensaje: `Datos guardados en ${sheetName}` };
}

function calculateBonusFromScratch(mes, anio, sheetName, config) {
  const controlOps = getRealControlOperadores(); 
  const respuestaDatos = obtenerDatosMensuales(mes, anio); // Reutilizamos lógica historial
  const rawData = respuestaDatos.exito ? respuestaDatos.datos : [];

  // --- PASO 1: AGREGACIÓN DE DATOS POR OPERADOR ---
  // CORRECCIÓN: Usar la EXTENSIÓN como clave única, no el nombre.
  const opsData = {}; 

  // Inicializar todos los activos
  controlOps.forEach(op => {
    if (op.estado && op.estado.toString().toUpperCase() === "ACTIVO") {
      const extKey = String(op.extension).trim();
      // Solo agregamos si hay extensión válida
      if (extKey) {
        opsData[extKey] = {
          meta: op,
          logs: [],
          totalLlamadas: 0,
          totalProd: 0, // Conv o Min
          tipoServicio: "Desconocido" // Se detectará de los logs o configuración
        };
      }
    }
  });

  // Procesar logs diarios
  rawData.forEach(log => {
    // Usar EXTENSIÓN para hacer match
    const logExt = String(log.extension).trim();
    
    if (opsData[logExt]) {
      const op = opsData[logExt];
      
      // Parsear fecha log
      const partes = log.fecha.split('/');
      const logDate = new Date(partes[2], partes[1] - 1, partes[0]);
      logDate.setHours(0,0,0,0);

      // Parsear fecha ingreso
      let ingresoDate = new Date(0);
      if (op.meta.fechaIngreso) {
         let parsed = null;
         if(op.meta.fechaIngreso instanceof Date) parsed = op.meta.fechaIngreso;
         else if(typeof op.meta.fechaIngreso === "string") {
             const p = op.meta.fechaIngreso.split('/'); 
             if (p.length === 3) parsed = new Date(p[2], p[1] - 1, p[0]);
         }
         if(parsed && !isNaN(parsed)) ingresoDate = parsed;
      }
      ingresoDate.setHours(0,0,0,0);

      // FILTRO FECHA
      if (logDate >= ingresoDate) {
        
        // Detectar tipo servicio si aun no se tiene
        const srvUpper = String(log.servicio).toUpperCase();
        let esTransfer = false;
        
        if (srvUpper.includes("TRANSFER") || srvUpper.includes("FACILITADOR")) {
             op.tipoServicio = "Transfer";
             esTransfer = true;
        } else if (srvUpper.includes("ASISTENTE")) {
             op.tipoServicio = "Asistente";
        } else {
             // Si el log no es claro, intentar mantener lo que ya tenía o adivinar
             if (op.tipoServicio === "Transfer") esTransfer = true;
        }

        // Determinar métrica de producción
        let prodDia = 0;
        // Importante: Usar la métrica correcta según el servicio detectado
        if (op.tipoServicio === "Transfer" || esTransfer) {
             prodDia = log.metric2; // metric2 es Conv en transfer
        } else {
             prodDia = log.metric2; // metric2 es Min en asistente
        }
        
        // ACUMULAR
        op.totalLlamadas += (log.llamadas || 0);
        op.totalProd += (prodDia || 0);

        op.logs.push({
          fecha: log.fecha,
          llamadas: log.llamadas || 0,
          produccion: prodDia || 0,
          tipo: (op.tipoServicio === "Transfer" || esTransfer) ? "Conversiones" : "Minutos",
          ratioDia: (log.llamadas > 0) ? (prodDia / log.llamadas) : 0
        });
      }
    }
  });

  // --- PASO 2: CALCULAR PROMEDIOS GRUPALES (PARA PONDERACIÓN) ---
  const statsGroup = {
    Transfer: { totalLlamadas: 0, countOps: 0 },
    Asistente: { totalLlamadas: 0, countOps: 0 }
  };

  Object.values(opsData).forEach(op => {
      if (op.tipoServicio === "Transfer" || op.tipoServicio === "Asistente") {
          // Solo contamos para el promedio si tuvo actividad
          if (op.totalLlamadas > 0) {
            statsGroup[op.tipoServicio].totalLlamadas += op.totalLlamadas;
            statsGroup[op.tipoServicio].countOps++;
          }
      }
  });

  const avgCallsTransfer = statsGroup.Transfer.countOps > 0 ? statsGroup.Transfer.totalLlamadas / statsGroup.Transfer.countOps : 1;
  const avgCallsAsistente = statsGroup.Asistente.countOps > 0 ? statsGroup.Asistente.totalLlamadas / statsGroup.Asistente.countOps : 1;

  // --- PASO 3: CONSTRUIR TABLA FINAL ---
  const bonusTable = [];
  const drillDownData = {};

  // Iterar por las extensiones (keys de opsData)
  Object.keys(opsData).forEach(extKey => {
    const op = opsData[extKey];
    
    // Calcular Ratio/Media del mes
    let ratioMedia = 0;
    if (op.totalLlamadas > 0) {
       ratioMedia = op.totalProd / op.totalLlamadas;
    }
    
    // Determinar Bono Opta (Usando Config)
    let bonoOpta = 0;
    if (config) {
        const reglas = op.tipoServicio === "Transfer" ? config.transfer : config.asistente;
        if (reglas && reglas.length > 0) {
            let valorComparar = op.tipoServicio === "Transfer" ? (ratioMedia * 100) : ratioMedia;
            
            reglas.forEach(r => {
                if (valorComparar >= parseFloat(r.min)) {
                    if (parseFloat(r.monto) > bonoOpta) bonoOpta = parseFloat(r.monto);
                }
            });
        }
    }

    // Calcular Ponderación
    let avgRef = op.tipoServicio === "Transfer" ? avgCallsTransfer : avgCallsAsistente;
    if (avgRef === 0) avgRef = 1;
    
    let indicePond = op.totalLlamadas / avgRef;
    
    // Calcular % Final (Tope 100% = 1.0)
    let porcFinal = indicePond > 1 ? 1 : indicePond;

    // Bono Aprobado
    let bonoAprobado = bonoOpta * porcFinal;

    // Formateo para tabla
    bonusTable.push({
      operador: op.meta.nombre,
      extension: op.meta.extension,
      fechaIngreso: op.meta.fechaIngreso instanceof Date ? formatDate(op.meta.fechaIngreso) : String(op.meta.fechaIngreso),
      horasReales: op.meta.horasReales,
      llamadas: op.totalLlamadas,
      produccion: op.totalProd,
      // Guardamos el ratio raw formateado
      ratioMedia: op.tipoServicio === "Transfer" ? (ratioMedia*100).toFixed(2)+"%" : ratioMedia.toFixed(2),
      bonoOpta: bonoOpta,
      indicePond: indicePond.toFixed(2),
      porcFinal: (porcFinal*100).toFixed(0) + "%",
      bonoAprobado: bonoAprobado.toFixed(2),
      
      // Datos ocultos para el detail y filtros
      rawService: op.tipoServicio
    });

    // Guardamos drillDown usando una clave única (Nombre + Ext) por si acaso se repiten nombres
    drillDownData[op.meta.nombre + "_" + op.meta.extension] = op.logs;
  });

  return {
    exito: true,
    sheetName: sheetName,
    data: bonusTable,
    details: drillDownData,
    config: config,
    fechaCalculo: new Date().toLocaleString()
  };
}

function readBonusSheet(sheet, sheetName, config) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { exito: true, sheetName: sheetName, data: [], details: {}, config: config, fechaCalculo: "Hoja vacía" };

  const values = sheet.getRange(2, 1, lastRow - 1, 11).getValues(); // Leemos 11 columnas ahora
  const bonusTable = values.map(row => ({
      operador: row[0],
      extension: row[1],
      fechaIngreso: row[2] instanceof Date ? formatDate(row[2]) : row[2],
      horasReales: row[3],
      llamadas: row[4],
      produccion: row[5],
      ratioMedia: row[6],
      bonoOpta: row[7],
      indicePond: row[8],
      porcFinal: row[9],
      bonoAprobado: row[10]
  }));
  
  return { exito: true, sheetName: sheetName, data: bonusTable, details: {}, config: config, fechaCalculo: "Recuperado" };
}


// ==================================================================
// 2. HELPERS (CONTROL Y HISTORIAL)
// ==================================================================

function getRealControlOperadores() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("CONTROL OPERADORES");
  if (!sheet) throw new Error("Falta hoja 'CONTROL OPERADORES'");

  const startRow = 5; 
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return [];

  const range = sheet.getRange(startRow, 1, lastRow - startRow + 1, 50);
  const values = range.getValues();
  const displayValues = range.getDisplayValues(); // Importante para Horas AX

  return values.map((row, i) => ({
      estado: row[0],       
      fechaIngreso: row[9], 
      nombre: row[14],      
      extension: row[15],   
      horasReales: displayValues[i][49] // Col AX (Indice 49) como texto
  }));
}

function obtenerDatosMensuales(mes, anio) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Historial");
  if (!sheet || sheet.getLastRow() < 2) return { exito: false, datos: [] };

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const datosCrudos = [];

  data.forEach(fila => {
      const fechaRaw = fila[0];
      if (!fechaRaw) return;
      let fechaObj = null;
      if (fechaRaw instanceof Date) fechaObj = fechaRaw;
      else if (typeof fechaRaw === "string") {
          const partes = fechaRaw.split('/'); 
          if (partes.length === 3) fechaObj = new Date(partes[2], partes[1] - 1, partes[0]);
      }
      if (!fechaObj || isNaN(fechaObj.getTime())) return;

      if (fechaObj.getMonth() + 1 !== parseInt(mes) || fechaObj.getFullYear() !== parseInt(anio)) return;

      datosCrudos.push({
          fecha: Utilities.formatDate(fechaObj, Session.getScriptTimeZone(), "dd/MM/yyyy"),
          operador: String(fila[1]).trim() || String(fila[2]).trim(),
          extension: String(fila[2]).trim(),
          servicio: String(fila[3]).trim(),
          llamadas: parseInt(fila[4]) || 0,
          metric2: parseFloat(fila[5]) || 0 // Conv o Min
      });
  });
  return { exito: true, datos: datosCrudos };
}

function getMonthName(m) {
  const months = ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO","JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"];
  return months[parseInt(m)-1] || "MES";
}

function formatDate(dateObj) {
  if (!dateObj) return "";
  if (typeof dateObj === 'string') return dateObj;
  const d = new Date(dateObj);
  if (isNaN(d.getTime())) return "";
  const day = String(d.getDate()).padStart(2, '0');
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const year = d.getFullYear();
  return `${day}/${month}/${year}`;
}

// ==================================================================
// 3. LOGICA EXISTENTE DASHBOARD (NO TOCADA)
// ==================================================================
// ... (Aquí van limpiarAgente, normalizarFechaHora, procesarDatosCentralizado, etc. 
//      Cópialas del código anterior si las necesitas en este archivo único)

function limpiarAgente(raw) {
    if (!raw) return "";
    let texto = String(raw).trim();
    const match = texto.match(/(Facilitar|Asistente).*$/i);
    if (match) return match[0].trim(); 
    if (texto.includes('/')) return texto.split('/').pop().trim();
    return texto;
}

function normalizarFechaHora(fechaRaw, horaRaw) {
  let fechaObj = new Date();
  if (fechaRaw instanceof Date) fechaObj = new Date(fechaRaw);
  else if (typeof fechaRaw === "string") {
    const partes = fechaRaw.split(/[-/]/); 
    if (partes.length === 3) fechaObj = new Date(partes[2], partes[1] - 1, partes[0]);
  }
  
  if (horaRaw instanceof Date) {
    fechaObj.setHours(horaRaw.getHours(), horaRaw.getMinutes(), horaRaw.getSeconds());
  } else if (typeof horaRaw === "string") {
    const partes = horaRaw.split(":");
    if (partes.length >= 2) fechaObj.setHours(parseInt(partes[0]), parseInt(partes[1]), partes[2] ? parseInt(partes[2]) : 0);
  } else if (typeof horaRaw === "number") { 
    const totalSegundos = Math.round(horaRaw * 86400);
    fechaObj.setHours(Math.floor(totalSegundos / 3600));
    fechaObj.setMinutes(Math.floor((totalSegundos % 3600) / 60));
    fechaObj.setSeconds(totalSegundos % 60);
  }
  return isNaN(fechaObj.getTime()) ? null : fechaObj;
}

function esMismoDia(d1, d2) {
    return d1.getFullYear() === d2.getFullYear() &&
           d1.getMonth() === d2.getMonth() &&
           d1.getDate() === d2.getDate();
}

function calcularEficiencia(agenteStats, avgF, avgA) {
    let score = 0;
    const { recibidas, contestadas, convGeneradas, isFacilitador, shortCalls } = agenteStats;
    const answerRate = recibidas > 0 ? contestadas / recibidas : 0;
    const conversionRate = recibidas > 0 ? convGeneradas / recibidas : 0;
    const bounceRate = recibidas > 0 ? shortCalls / recibidas : 0; 
    const avgContestadas = isFacilitador ? avgF : avgA;

    score += Math.min(answerRate * 30, 30);
    let scoreB = 0;
    if (isFacilitador) {
        if (conversionRate >= 0.12) scoreB = 30;
        else if (conversionRate >= 0.05) scoreB = 10;
        else if (conversionRate >= 0.03) scoreB = 5;
    }
    score += scoreB;
    let scoreC = 0;
    if (bounceRate <= 0.05) scoreC = 30;
    else if (bounceRate >= 0.20) scoreC = 0;
    else {
        const rateNorm = (bounceRate - 0.05) / (0.20 - 0.05); 
        scoreC = 30 - (rateNorm * 30);
    }
    score += Math.max(0, scoreC); 
    let scoreD = 0;
    if (avgContestadas > 0) {
        if (contestadas > avgContestadas) scoreD = 10;
        else if (contestadas >= avgContestadas * 0.30) scoreD = 5;
    }
    score += scoreD;

    return { 
      score: Math.round(score), 
      details: { 
        answerRate: answerRate.toFixed(4), 
        conversionRate: conversionRate.toFixed(4), 
        bounceRate: bounceRate.toFixed(4),
        scoreA: Math.round(Math.min(answerRate * 30, 30)), 
        scoreB, scoreC: Math.round(Math.max(0, scoreC)), scoreD 
      } 
    };
}

function procesarDatosCentralizado(filtroFecha) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const datosSheet = ss.getSheetByName("Datos");
  if (!datosSheet) return { exito: false, mensaje: "Falta hoja Datos" };

  const rawData = datosSheet.getDataRange().getValues();
  if (rawData.length < 2) return { exito: false, mensaje: "La hoja Datos está vacía" };

  const headers = rawData[0].map(h => String(h).trim().toLowerCase());
  const findCol = (k) => headers.findIndex(h => k.some(x => h.includes(x.toLowerCase())));
  const col = {
      AGENTE: findCol(["destino", "agente"]),
      NUMERO: findCol(["número", "numero", "contacto"]),
      FECHA: findCol(["fecha"]),
      HORA: findCol(["hora"]),
      DURACION: findCol(["duración", "duracion"]),
      ESTADO: findCol(["estado"]),
      SERVICIO: findCol(["servicio"]),
      COLGADO: findCol(["colgado", "colgado por"])
  };

  if (col.AGENTE === -1) return { exito: false, mensaje: "No se encuentra columna Agente." };

  let llamadas = [];
  const conversionesPorHora = new Array(24).fill(null).map(() => ({ total: 0, directa: 0, conversion: 0 })); 
  const detalleConversiones = [];

  for (let i = 1; i < rawData.length; i++) {
    const row = rawData[i];
    const numero = col.NUMERO > -1 ? row[col.NUMERO]?.toString().trim() : "S/N";
    if (!numero) continue;

    const agenteRaw = row[col.AGENTE]?.toString().trim() || "[Sin Agente]";
    const agenteLimpio = limpiarAgente(agenteRaw);

    const valFecha = col.FECHA > -1 ? row[col.FECHA] : null;
    const valHora = col.HORA > -1 ? row[col.HORA] : null;
    const timestamp = normalizarFechaHora(valFecha, valHora);
    if (!timestamp) continue; 

    const duracion = col.DURACION > -1 ? (parseFloat(row[col.DURACION]) || 0) : 0;
    const estado = col.ESTADO > -1 ? row[col.ESTADO]?.toString().trim() || "" : "";
    const colgadoPor = col.COLGADO > -1 ? row[col.COLGADO]?.toString().trim() || "Desconocido" : "Desconocido";

    let tipo = "OTRO";
    const srv = col.SERVICIO > -1 ? row[col.SERVICIO]?.toString() || "" : "";
    if (srv.includes("91") || agenteLimpio.toLowerCase().includes("facilitar")) tipo = "FACILITAR";
    else if (srv.includes("807") || agenteLimpio.toLowerCase().includes("asistente")) tipo = "ASISTENTE";
    else if (srv.includes("900") || agenteLimpio.toLowerCase().includes("operador")) tipo = "OPERADOR";

    llamadas.push({
      telefono: numero, agente: agenteLimpio, tipo, estado, duracion, timestamp,
      horaTexto: Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "HH:mm:ss"),
      colgadoPorRaw: colgadoPor
    });
  }

  llamadas.sort((a, b) => a.timestamp - b.timestamp);

  const statsAgentes = {};
  const logs = { conversiones: [], rebotes: [], bypass: [], abandonos: [] };
  const memoriaFacilitador = {}; 
  const ultimaInteraccion = {}; 
  const statsColgados = {}; 

  llamadas.forEach(call => {
    // 1. Inicialización de contadores (incluido convReales)
    if (!statsAgentes[call.agente]) statsAgentes[call.agente] = { 
      recibidas: 0, contestadas: 0, abandonadas: 0, 
      convGeneradas: 0, convReales: 0, 
      convRecibidas: 0, 
      seguimientoRecibido: 0, bypass: 0, totalDuracion: 0, shortCalls: 0, numerosColgados: [],
      isFacilitador: false, isAsistente: false 
    };

    if (call.agente.toLowerCase().includes("facilitar")) statsAgentes[call.agente].isFacilitador = true;
    if (call.agente.toLowerCase().includes("asistente")) statsAgentes[call.agente].isAsistente = true;

    const esTargetDay = !filtroFecha || esMismoDia(call.timestamp, filtroFecha);

    if (esTargetDay) {
        statsAgentes[call.agente].recibidas++;
        if (call.duracion < 15) statsAgentes[call.agente].shortCalls++;
        
        const colgadoLower = call.colgadoPorRaw.toLowerCase();
        if (!statsColgados[call.colgadoPorRaw]) statsColgados[call.colgadoPorRaw] = 0;
        statsColgados[call.colgadoPorRaw]++;
        
        if (colgadoLower.includes("agente") || colgadoLower.includes("agent") || colgadoLower.includes("llamado")) {
            statsAgentes[call.agente].numerosColgados.push(call.telefono);
        }

        if (call.estado === "Contestada") {
             statsAgentes[call.agente].contestadas++;
             statsAgentes[call.agente].totalDuracion += call.duracion; 
        }

        if (call.estado.toLowerCase().includes("abandonada") || call.agente === "[Sin destino]") {
             statsAgentes[call.agente].abandonadas++;
             logs.abandonos.push([call.horaTexto, call.tipo, call.telefono, call.agente]);
        }
    }

    const last = ultimaInteraccion[call.telefono];
    if (last) {
      const diffMin = (call.timestamp - last.timestamp) / 60000;
      if (diffMin < 30) {
        let patron = "";
        if (last.tipo === "FACILITAR" && call.tipo === "FACILITAR") patron = "F > F";
        if (last.tipo === "ASISTENTE" && call.tipo === "ASISTENTE") patron = "A > A";
        if (last.tipo === "ASISTENTE" && call.tipo === "FACILITAR") patron = "A > F";
        if (patron && esTargetDay) {
             logs.rebotes.push({
                telefono: call.telefono, patron: patron, agente1: last.agente,
                hora1: Utilities.formatDate(last.timestamp, Session.getScriptTimeZone(), "HH:mm:ss"),
                agente2: call.agente, hora2: call.horaTexto, diff: diffMin.toFixed(1) + " min"
             });
        }
      }
    }
    if (call.estado !== "Abandonada" && call.duracion > 0) {
        ultimaInteraccion[call.telefono] = { tipo: call.tipo, agente: call.agente, timestamp: call.timestamp };
    }

    if (call.tipo === "FACILITAR") {
      memoriaFacilitador[call.telefono] = { 
          agente: call.agente, timestamp: call.timestamp, duracion: call.duracion
      };
    } else if (call.tipo === "ASISTENTE") {
        const previo = memoriaFacilitador[call.telefono];
        let tipoConversion = "DIRECTA";
        
        if (previo && (call.timestamp - previo.timestamp) < 72000000) { 
            tipoConversion = "CONVERSIÓN";
            if (esTargetDay) {
                statsAgentes[call.agente].convRecibidas++;
                if (statsAgentes[previo.agente]) {
                    // --- CORRECCIÓN LÓGICA SOLICITADA ---
                    
                    // A) Conversión Real (Intentos): Cuenta cualquiera > 0 seg
                    if (call.duracion > 0) {
                        statsAgentes[previo.agente].convReales++;
                    }
                    
                    // B) Conversión Facturable: Solo cuenta si dura más de 25 seg
                    if (call.duracion > 25) {
                        statsAgentes[previo.agente].convGeneradas++;
                    }
                }
                const gap = (call.timestamp - new Date(previo.timestamp.getTime() + (previo.duracion*1000))) / 1000;
                logs.conversiones.push([
                    call.telefono, previo.agente, 
                    Utilities.formatDate(previo.timestamp, Session.getScriptTimeZone(), "HH:mm:ss"), 
                    call.agente, call.horaTexto, gap.toFixed(0)
                ]);
            }
        } else {
            if (esTargetDay) {
                statsAgentes[call.agente].bypass++;
                logs.bypass.push([call.telefono, call.agente, call.horaTexto, "Sin registro previo"]);
            }
        }

        if (esTargetDay) {
            const horaInt = parseInt(Utilities.formatDate(call.timestamp, "Europe/Madrid", "HH"), 10);
            if (!isNaN(horaInt) && horaInt >= 0 && horaInt < 24) {
                conversionesPorHora[horaInt].total++;
                if (tipoConversion === "CONVERSIÓN") conversionesPorHora[horaInt].conversion++;
                else conversionesPorHora[horaInt].directa++;
            }
            detalleConversiones.push([
                Utilities.formatDate(call.timestamp, "Europe/Madrid", "HH:mm:ss"),
                call.telefono, tipoConversion, call.agente, call.horaTexto
            ]);
        }
    }
  });
  
  return { exito: true, datos: { llamadas, statsAgentes, logs, statsColgados, conversionesPorHora, detalleConversiones } };
}

function obtenerDatosParaGrafico() {
  const resultado = procesarDatosCentralizado(null); 
  if (!resultado.exito) throw new Error(resultado.mensaje);

  const { statsAgentes, logs, llamadas, statsColgados, conversionesPorHora } = resultado.datos;
  const rebotesLog = logs.rebotes;

  let totalContestadasF = 0; let countF = 0;
  let totalContestadasA = 0; let countA = 0;
  
  Object.keys(statsAgentes).forEach(k => {
      const s = statsAgentes[k];
      s.isFacilitador = s.convGeneradas > 0 || (s.recibidas > 0 && !s.convRecibidas && !s.bypass); 
      s.isAsistente = s.convRecibidas > 0 || s.bypass > 0;
      
      if (s.isFacilitador && s.contestadas > 0) { totalContestadasF += s.contestadas; countF++; }
      if (s.isAsistente && s.contestadas > 0) { totalContestadasA += s.contestadas; countA++; }
  });

  const avgContestadasF = countF > 0 ? totalContestadasF / countF : 0;
  const avgContestadasA = countA > 0 ? totalContestadasA / countA : 0;
  
  Object.keys(statsAgentes).forEach(k => {
      const s = statsAgentes[k];
      const { score, details } = calcularEficiencia(s, avgContestadasF, avgContestadasA);
      s.eficaciaScore = score;
      s.eficaciaDetails = details;
      s.avgDuracion = s.contestadas > 0 ? (s.totalDuracion / s.contestadas) / 60 : 0;
      s.totalMinutos = s.totalDuracion / 60;
      s.ratio = s.recibidas > 0 ? (s.convGeneradas / s.recibidas) * 100 : 0;
  });

  const rankingRebotes = Object.keys(statsAgentes).map(agente => {
    const s = statsAgentes[agente];
    return {
      agente: agente,
      rebotes: s.shortCalls,
      total: s.recibidas,
      tasa: s.recibidas > 0 ? ((s.shortCalls / s.recibidas) * 100).toFixed(1) : 0
    };
  }).filter(r => r.total > 0 && r.rebotes > 0);
  rankingRebotes.sort((a, b) => b.tasa - a.tasa);

  const tiposRebote = { "F > F": 0, "A > A": 0, "A > F": 0 };
  rebotesLog.forEach(r => {
    if (tiposRebote[r.patron] !== undefined) tiposRebote[r.patron]++;
    else tiposRebote["Otros"] = (tiposRebote["Otros"] || 0) + 1;
  });
  
  const metricasGlobales = calcularMetricasGlobales(llamadas, logs, statsAgentes);

  return {
    stats: statsAgentes,
    logs: {
      totalConversiones: logs.conversiones.length,
      totalBypass: logs.bypass.length,
      totalAbandonos: logs.abandonos.length
    },
    analisisRebotes: { ranking: rankingRebotes.slice(0, 10), tipos: tiposRebote, total: rebotesLog.length },
    globalAverages: { avgContestadasF: avgContestadasF.toFixed(0), avgContestadasA: avgContestadasA.toFixed(0) },
    metricasGlobales: metricasGlobales,
    statsColgados: statsColgados,
    conversionesPorHora: conversionesPorHora 
  };
}

function calcularMetricasGlobales(llamadas, logs, statsAgentes) {
  let total807 = 0; let respondidas807 = 0; let totalMinutos807 = 0;
  let total91 = 0; let respondidas91 = 0; let totalMinutos91 = 0;
  let totalOperatorCalls = 0;

  llamadas.forEach(call => {
    if (call.tipo === "ASISTENTE") {
      total807++;
      if (call.estado === "Contestada") { respondidas807++; totalMinutos807 += call.duracion / 60; }
    } else if (call.tipo === "FACILITAR") {
      total91++;
      if (call.estado === "Contestada") { respondidas91++; totalMinutos91 += call.duracion / 60; }
    }
    if (call.agente !== "[Sin Agente]" && !call.agente.startsWith("C24")) {
      if (call.agente.toLowerCase().includes("operador")) totalOperatorCalls++;
    }
  });

  const rebotes807 = logs.rebotes.filter(r => r.patron === "A > A" || r.patron === "A > F").length;
  const rebotes91 = logs.rebotes.filter(r => r.patron === "F > F").length;
  const totalConversionesGlobal = respondidas807; 
  const totalFacturables = totalConversionesGlobal + logs.seguimiento?.length || 0; 
  
  const avgMedia807 = respondidas807 > 0 ? totalMinutos807 / respondidas807 : 0;
  const avgMedia91 = respondidas91 > 0 ? totalMinutos91 / respondidas91 : 0;
  const ratioConversionGlobal = respondidas91 > 0 ? (respondidas807 / respondidas91) * 100 : 0;
  const percentRebote807 = respondidas807 > 0 ? (rebotes807 / respondidas807) * 100 : 0;
  const percentRebote91 = respondidas91 > 0 ? (rebotes91 / respondidas91) * 100 : 0;
  
  const ext91Conv = Object.values(statsAgentes).reduce((sum, s) => sum + (s.isFacilitador ? s.convGeneradas : 0), 0);
  const ext807Conv = Object.values(statsAgentes).reduce((sum, s) => sum + (s.isAsistente ? s.convRecibidas : 0), 0);
  
  return {
    asistente: { totalLlamadas: total807, respondidas: respondidas807, conversiones: totalConversionesGlobal, totalMinutos: totalMinutos807.toFixed(2), media: avgMedia807.toFixed(2), totalRebote: rebotes807, percentRebote: percentRebote807.toFixed(2), minFacturables: totalMinutos807.toFixed(2) },
    facilitador: { totalLlamadas: total91, respondidas: respondidas91, ratioConversionGlobal: ratioConversionGlobal.toFixed(2), totalMinutos: totalMinutos91.toFixed(2), media: avgMedia91.toFixed(2), totalRebote: rebotes91, percentRebote: percentRebote91.toFixed(2), extConv: ext91Conv },
    generales: { totalFacturables: totalFacturables, ext807Conv: ext807Conv, totalOperador: totalOperatorCalls }
  };
}

function generarDatosHistorial(fechaStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fechaFiltro = new Date(fechaStr + 'T12:00:00'); 
  
  const resultado = procesarDatosCentralizado(fechaFiltro);
  if (!resultado.exito) return { exito: false, mensaje: resultado.mensaje };

  const { statsAgentes } = resultado.datos;
  const mapaNombres = {};
  
  const sheetControl = ss.getSheetByName("CONTROL OPERADORES");
  if (sheetControl && sheetControl.getLastRow() > 4) {
      const headers = sheetControl.getRange(4, 1, 1, sheetControl.getLastColumn()).getValues()[0]
          .map(h => String(h).toUpperCase().trim());
      
      const idxExt = headers.indexOf("EXT");
      const idxMattermost = headers.indexOf("MATTERMOST");
      const idxEstado = headers.indexOf("ESTADO");
      
      if (idxExt > -1 && idxMattermost > -1) {
          const datos = sheetControl.getRange(5, 1, sheetControl.getLastRow() - 4, sheetControl.getLastColumn()).getValues();
          datos.forEach(fila => {
              const extRaw = String(fila[idxExt]).trim();
              const nombre = String(fila[idxMattermost]).trim();
              const estado = idxEstado > -1 ? String(fila[idxEstado]).toUpperCase() : "";
              if (extRaw && nombre) {
                  const key = extRaw.toLowerCase();
                  if (!mapaNombres[key] || estado.includes("ACTIVO") || estado === "ALTA") {
                      mapaNombres[key] = nombre;
                  }
              }
          });
      }
  }

  const fechaRegistro = Utilities.formatDate(fechaFiltro, Session.getScriptTimeZone(), "dd/MM/yyyy");
  const filasSalida = [];

  Object.keys(statsAgentes).forEach(extKey => {
    const s = statsAgentes[extKey];
    if (s.recibidas === 0 && s.contestadas === 0 && s.convGeneradas === 0 && s.convRecibidas === 0) return;

    const nombreReal = mapaNombres[extKey.toLowerCase()] || ""; 
    let servicio = "Desconocido";
    if (s.isFacilitador) servicio = "Transfer";
    else if (s.isAsistente) servicio = "Asistente";
    
    let metric1 = 0; let metric2 = 0; let metric3 = 0;

    if (servicio === "Transfer") {
        metric1 = s.recibidas;
        metric2 = s.convGeneradas; 
        metric3 = s.recibidas > 0 ? (s.convGeneradas / s.recibidas) * 100 : 0; 
    } else {
        metric1 = s.contestadas;
        metric2 = s.totalDuracion / 60; 
        metric3 = s.contestadas > 0 ? (metric2 / s.contestadas) : 0; 
    }

    filasSalida.push([
        fechaRegistro,
        nombreReal,    
        extKey,        
        servicio,
        metric1,        
        metric2.toFixed(2), 
        metric3.toFixed(2)  
    ]);
  });
  
  return { exito: true, datos: filasSalida };
}

function guardarHistorialPorFecha(fechaStr) {
  const res = generarDatosHistorial(fechaStr);
  if (!res.exito) return res.mensaje;
  
  const datos = res.datos;
  if (datos.length === 0) return "⚠️ No hay datos para guardar en esta fecha.";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Historial");
  
  if (!sheet) {
    sheet = ss.insertSheet("Historial");
    sheet.appendRow([
      "Fecha", "Operador (Mattermost)", "Extensión", "Servicio", 
      "Llamadas (Rec/Cont)", "Conv / Min.Totales", "Ratio(%) / Media(Min)"
    ]);
    sheet.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#d9ead3");
    sheet.setFrozenRows(1);
  }
  
  sheet.getRange(sheet.getLastRow() + 1, 1, datos.length, datos[0].length).setValues(datos);
  return `✅ Guardados ${datos.length} registros del ${fechaStr}.`;
}

function previewHistorialPorFecha(fechaStr) {
  const res = generarDatosHistorial(fechaStr);
  if (!res.exito) return { error: true, mensaje: res.mensaje };
  return { error: false, datos: res.datos };
}

function generarReporteOptimizado() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultado = procesarDatosCentralizado(null);
  if (!resultado.exito) { SpreadsheetApp.getUi().alert(resultado.mensaje); return; }
  
  const { statsAgentes } = resultado.datos;
  escribirHoja(ss, "1_Global_Agentes", 
    ["Agente", "Recibidas", "Contestadas", "Conversiones(F)", "Ventas(A)", "Score"],
    Object.keys(statsAgentes).map(k => {
       const s = statsAgentes[k];
       const {score} = calcularEficiencia(s, 0, 0); 
       return [k, s.recibidas, s.contestadas, s.convGeneradas, s.convRecibidas, score];
    })
  );
  SpreadsheetApp.getUi().alert("Reporte Básico Generado");
}

function escribirHoja(ss, nombre, headers, datos) {
  let sheet = ss.getSheetByName(nombre);
  if (!sheet) sheet = ss.insertSheet(nombre); else sheet.clear();
  sheet.appendRow(headers);
  if (datos.length > 0) sheet.getRange(2, 1, datos.length, datos[0].length).setValues(datos);
}

function getOperatorsMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("CONTROL OPERADORES");
  if (!sheet) return {};
  
  // Leer cabeceras (Fila 4) para encontrar las columnas dinámicamente
  const headers = sheet.getRange(4, 1, 1, sheet.getLastColumn()).getValues()[0]
                       .map(h => String(h).toUpperCase().trim());
  
  const idxExt = headers.indexOf("EXT");
  // Buscamos sinónimos de la columna de servicio
  let idxSrv = headers.indexOf("SERVICIO");
  if (idxSrv === -1) idxSrv = headers.indexOf("PERFIL");
  if (idxSrv === -1) idxSrv = headers.indexOf("CAMPAÑA");
  if (idxSrv === -1) idxSrv = headers.indexOf("PUESTO"); // Intento final
  
  // Si no encuentra columnas clave, devuelve vacío
  if (idxExt === -1 || idxSrv === -1) return {};
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 5) return {};

  const data = sheet.getRange(5, 1, lastRow - 4, sheet.getLastColumn()).getValues();
  const map = {};
  
  data.forEach(row => {
    const ext = String(row[idxExt]).trim();
    const srv = String(row[idxSrv]).trim();
    const estado = String(row[0]).toUpperCase(); // Asumiendo col 0 es estado, opcional
    
    // Guardamos el servicio oficial asociado a la extensión
    if (ext) {
      map[ext] = srv;
    }
  });
  
  return map;
}

/**
 * ==================================================================
 * 2. FUNCIÓN DE HISTORIAL (FUZZY MATCH Y BLINDADA)
 * ==================================================================
 * - Normaliza cadenas para ignorar espacios, _ y mayúsculas.
 * - Busca principalmente en Columna C (Extensión).
 * - Devuelve los últimos 7 registros encontrados.
 */
function getHistorialMediaSemanal(nombreAgente) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Historial");
    if (!sheet) return { exito: false, mensaje: "No se encontró la hoja Historial" };

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { exito: true, datos: [], debug: { mensaje: "Hoja vacía" } };

    // Leemos columnas A hasta G
    // A=Fecha(0), B=Operador(1), C=Extensión(2) ... G=Media(6)
    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    
    const registrosEncontrados = [];
    
    // --- FUNCIÓN DE NORMALIZACIÓN (Fuzzy Simple) ---
    // Elimina espacios, guiones bajos, guiones medios y pasa a minúsculas
    const normalize = (str) => String(str).toLowerCase().replace(/[\s\-_]/g, '');
    
    const searchNormalized = normalize(nombreAgente);

    // Muestra para debug
    const sampleData = data.slice(0, 5).map(r => 
      `F:${String(r[0]).substring(0,10)} | Ext(C):${r[2]} | Nom(B):${r[1]}`
    );

    data.forEach(row => {
      const fechaRaw = row[0];
      const opExt = normalize(row[2]);  // Normalizar Columna C (Extensión)
      const opName = normalize(row[1]); // Normalizar Columna B (Nombre - Fallback)
      
      let mediaVal = row[6]; // Columna G
      if (typeof mediaVal === 'string') {
          mediaVal = parseFloat(mediaVal.replace(',', '.')); 
      }
      mediaVal = parseFloat(mediaVal) || 0;

      // 1. MATCH: Comparamos las versiones "limpias"
      // Prioridad a Extensión (C), pero revisamos Nombre (B) por si acaso.
      if (opExt === searchNormalized || opName === searchNormalized) {
        
        // 2. PARSEO DE FECHA
        let fechaObj = null;
        if (fechaRaw instanceof Date) {
          fechaObj = fechaRaw;
        } else if (typeof fechaRaw === 'string') {
           const parts = fechaRaw.split('/');
           if(parts.length === 3) {
             // Asume día/mes/año
             fechaObj = new Date(parts[2], parts[1]-1, parts[0]); 
           }
        }

        // 3. AGREGAR SI ES VÁLIDO
        if (fechaObj && !isNaN(fechaObj.getTime())) {
          registrosEncontrados.push({
            fechaRaw: fechaObj,
            fecha: Utilities.formatDate(fechaObj, Session.getScriptTimeZone(), "dd/MM"),
            valor: mediaVal
          });
        }
      }
    });

    // 4. LOGICA DE ÚLTIMOS 7 REGISTROS
    // Ordenar descendente (el más reciente primero)
    registrosEncontrados.sort((a, b) => b.fechaRaw - a.fechaRaw);
    
    // Cortar los primeros 7 (que son los más recientes)
    const ultimos7 = registrosEncontrados.slice(0, 7);
    
    // Re-ordenar ascendente para graficar cronológicamente
    ultimos7.sort((a, b) => a.fechaRaw - b.fechaRaw);

    return { 
      exito: true, 
      datos: ultimos7,
      debug: { 
        buscado: nombreAgente,
        buscadoNorm: searchNormalized,
        encontradosTotal: registrosEncontrados.length,
        muestraHoja: sampleData,
        mensaje: registrosEncontrados.length === 0 ? "Sin coincidencias (Fuzzy)" : "Datos OK"
      }
    };

  } catch (e) {
    return { 
      exito: false, 
      datos: [], 
      error: e.toString(),
      debug: { mensaje: "CRASH en Backend: " + e.toString() } 
    };
  }
}

// ===== SISTEMA DE AUTENTICACIÓN =====

function AUTH_validateUser(credentialsJSON) {
  try {
    const creds = JSON.parse(credentialsJSON);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const authSheet = ss.getSheetByName('Auth');
    
    if (!authSheet) {
      return JSON.stringify({ 
        success: false, 
        message: 'Sistema de autenticación no configurado' 
      });
    }
    
    // Buscar usuario
    const data = authSheet.getDataRange().getValues();
    const headers = data[0];
    const userCol = headers.indexOf('usuario');
    const passCol = headers.indexOf('password');
    const roleCol = headers.indexOf('rol');
    
    if (userCol === -1 || passCol === -1) {
      return JSON.stringify({ 
        success: false, 
        message: 'Estructura de Auth incorrecta' 
      });
    }
    
    // Hash simple de comparación (en producción usar mejor método)
    const hashedInput = Utilities.computeDigest(
      Utilities.DigestAlgorithm.MD5, 
      creds.password
    ).map(byte => (byte < 0 ? byte + 256 : byte).toString(16).padStart(2, '0')).join('');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][userCol] === creds.username && data[i][passCol] === hashedInput) {
        return JSON.stringify({
          success: true,
          role: data[i][roleCol] || 'operador',
          username: creds.username
        });
      }
    }
    
    return JSON.stringify({ success: false });
    
  } catch (e) {
    return JSON.stringify({ 
      success: false, 
      message: e.message 
    });
  }
}

// ===== CARGA DE ARCHIVOS =====

function UPLOAD_processFile(fileData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const datosSheet = ss.getSheetByName('Datos') || ss.insertSheet('Datos');
    
    let data = [];
    
    if (fileData.type === '.csv') {
      // Parsear CSV simple
      data = parseCSV(fileData.content);
    } else {
      // Para Excel, necesitaríamos usar Drive API o un parser más complejo
      // Por ahora, indicamos que CSV está listo y Excel requiere paso extra
      return JSON.stringify({
        success: false,
        message: 'Excel requiere procesamiento adicional. Por ahora usa CSV.'
      });
    }
    
    // Escribir en sheet (append o replace, según necesidad)
    const lastRow = datosSheet.getLastRow();
    
    if (lastRow === 0) {
      // Primera carga: escribir todo
      datosSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    } else {
      // Append: agregar después de la última fila
      datosSheet.getRange(lastRow + 1, 1, data.length, data[0].length).setValues(data);
    }
    
    // Log de quién subió qué
    logUpload(fileData.user, fileData.filename, data.length);
    
    return JSON.stringify({
      success: true,
      message: `Procesadas ${data.length} filas correctamente`
    });
    
  } catch (e) {
    return JSON.stringify({
      success: false,
      message: e.message
    });
  }
}

function parseCSV(csvText) {
  const lines = csvText.split('\n');
  const result = [];
  
  for (let line of lines) {
    if (line.trim()) {
      // Parseo simple: separar por coma, manejar comillas básicas
      const row = line.split(',').map(cell => {
        cell = cell.trim();
        if (cell.startsWith('"') && cell.endsWith('"')) {
          cell = cell.slice(1, -1);
        }
        return cell;
      });
      result.push(row);
    }
  }
  
  return result;
}

function logUpload(user, filename, rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName('LogUploads');
  
  if (!logSheet) {
    logSheet = ss.insertSheet('LogUploads');
    logSheet.appendRow(['Fecha', 'Usuario', 'Archivo', 'Filas']);
  }
  
  logSheet.appendRow([new Date(), user, filename, rows]);
}
