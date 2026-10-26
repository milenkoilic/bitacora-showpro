/*************************************************************
 * SHOWPRO - BITÁCORA DE GENERADORES - Code.gs
 * Pegar en: Extensiones > Apps Script > borrar Code.gs > pegar > Guardar > Ejecutar instalarTriggers() una vez
 * Compatible con las hojas: 1_MAESTRA_Equipos | Bitácora General | Resumen de Equipos
 *************************************************************/

// ============ CONFIGURACIÓN (EDITE AQUÍ) ============
const CONFIG = {
  EMAIL_ALERTA: 'MILENKOILIC@GMAIL.COM',
  // WHATSAPP por CallMeBot (gratis): active su número en https://www.callmebot.com/blog/free-api-whatsapp-messages/
  // Ya configurado con su número. Solo falta pegar el apikey que le entrega CallMeBot.
  WHATSAPP_PHONE: '56922273691',
  WHATSAPP_APIKEY: 'PEGUE_AQUI_SU_APIKEY',  // <-- reemplace por su apikey de CallMeBot, ej: '123456'
  UMBRAL_AVISO: 50,     // amarillo: programar servicio
  UMBRAL_CRITICO: 0,    // rojo: servicio vencido
  HOJA_MAESTRA: '1_MAESTRA_Equipos',
  HOJA_BITACORA: 'Bitácora General',
  HOJA_RESUMEN: 'Resumen de Equipos',
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚙ SHOWPRO')
    .addItem('1) Validar toda la bitácora', 'validarTodo')
    .addItem('2) Chequear horas restantes ahora', 'chequearAlertasManual')
    .addItem('3) Enviar reporte por correo', 'enviarReporte')
    .addItem('4) Instalar disparador diario 08:00', 'instalarTriggers')
    .addToUi();
}

// Se ejecuta al instalar: crea trigger diario 08:00 America/Santiago
function instalarTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'chequeoDiario') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('chequeoDiario').timeBased().everyDays(1).atHour(8).inTimezone('America/Santiago').create();
  SpreadsheetApp.getUi().alert('Disparador diario 08:00 instalado. Chequeo automático activo.');
}

function chequeoDiario() {
  chequearAlertas(true);
}
function chequearAlertasManual() {
  chequearAlertas(false);
}

// ============ LECTURA MAESTRA + RESUMEN ============
function getEquiposMaestra() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.HOJA_MAESTRA);
  const nombres = sh.getRange('B5:B14').getValues().flat();
  const intervalos = sh.getRange('F5:F14').getValues().flat();
  const iniciales = sh.getRange('E5:E14').getValues().flat();
  const map = {};
  nombres.forEach((raw, idx) => {
    const nombre = String(raw).trim();
    if (!nombre) return; // fila vacía = slot libre, se omite
    if (map[nombre]) return; // duplicado: se conserva el primero
    map[nombre] = {
      intervalo: Number(intervalos[idx]) || 250,
      inicial: Number(iniciales[idx]) || 0,
      filaMaestra: 5 + idx
    };
  });
  return map;
}

function getEstadoResumen() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.HOJA_RESUMEN);
  const rows = sh.getRange('A4:H13').getValues();
  return rows
    .map((r, i) => ({ fila: 4 + i, equipo: String(r[0]).trim(), intervalo: r[1], ultimaFecha: r[2], ultimoHorometro: r[3], horas30d: r[4], ultimaPreventiva: r[5], desde: r[6], restantes: r[7] }))
    .filter(r => r.equipo !== '');
}

// ============ VALIDACIÓN AL EDITAR BITÁCORA ============
function onEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== CONFIG.HOJA_BITACORA) return;
  const row = e.range.getRow();
  if (row < 5 || row > 504) return;
  validarFila(row, true);
}

function validarFila(row, avisar) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CONFIG.HOJA_BITACORA);
  const v = sh.getRange(row, 1, 1, 14).getValues()[0];
  const [fecha, evento, equipo, jornada, hi, hf, , tipo] = v;
  const errores = [];
  const equipos = getEquiposMaestra();

  if (!evento || String(evento).trim().length < 3) errores.push('B Evento obligatorio (mín. 3 letras).');
  if (!equipo || !equipos[String(equipo).trim()]) errores.push('C Equipo debe existir en Hoja 1 Maestra B5:B14.');
  if (hi === '' || hf === '' || isNaN(Number(hi)) || isNaN(Number(hf))) {
    errores.push('E/F Horómetros deben ser números.');
  } else {
    if (Number(hf) < Number(hi)) errores.push(`F Final (${hf}) < Inicial (${hi}). Corrija.`);
    if (Number(hf) - Number(hi) > 24) errores.push(`G Horas del día = ${Number(hf) - Number(hi)} > 24. Revise jornada (use 2 filas si trabajó día+noche).`);
    // avisar si retrocede vs último registrado del mismo equipo
    const ultimo = ultimoFinalEquipo(String(equipo).trim(), row);
    if (ultimo !== null && Number(hi) < ultimo) errores.push(`E Inicial (${hi}) menor que último Final registrado (${ultimo}) para ese equipo. Posible retroceso.`);
  }

  // pintar: rojo suave si hay error, blanco si ok (no toca G que es fórmula)
  const bg = errores.length ? '#FFC7CE' : '#FFFFFF';
  sh.getRange(row, 1, 1, 14).setBackground(errores.length ? bg : '#FFFFFF');
  // asegurar fórmula G si fue borrada
  const gCell = sh.getRange(row, 7);
  if (!gCell.getFormula()) gCell.setFormula(`=IF(AND(E${row}<>"",F${row}<>""),F${row}-E${row},"")`);

  if (errores.length && avisar) {
    sh.getRange(row, 1).setNote('⚠ ' + errores.join('\n'));
    SpreadsheetApp.getActive().toast('Fila ' + row + ': ' + errores[0], '⚠ Bitácora', 6);
  } else if (!errores.length) {
    sh.getRange(row, 1).clearNote();
  }
  return errores;
}

function ultimoFinalEquipo(equipo, excluirFila) {
  if (!equipo) return null;
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.HOJA_BITACORA);
  const finales = sh.getRange('C5:F504').getValues();
  let max = null;
  finales.forEach((r, i) => {
    const fila = 5 + i;
    if (fila === excluirFila) return;
    if (String(r[0]).trim() === equipo && r[3] !== '' && !isNaN(Number(r[3]))) {
      max = max === null ? Number(r[3]) : Math.max(max, Number(r[3]));
    }
  });
  return max;
}

function validarTodo() {
  let malas = 0;
  for (let r = 5; r <= 504; r++) {
    const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.HOJA_BITACORA);
    const fila = sh.getRange(r, 1, 1, 14).getValues()[0];
    if (fila.every(x => x === '' || x === null)) continue; // fila vacía se omite
    if (validarFila(r, false).length) malas++;
  }
  SpreadsheetApp.getUi().alert(malas === 0 ? '✅ Bitácora sin errores.' : `⚠ ${malas} fila(s) con error (pintadas rojo). Revise columna B/C/E/F.`);
}

// ============ ALERTAS HORAS RESTANTES ============
function chequearAlertas(esAutomatico) {
  const estados = getEstadoResumen();
  const criticos = estados.filter(x => typeof x.restantes === 'number' && x.restantes <= CONFIG.UMBRAL_CRITICO);
  const avisos = estados.filter(x => typeof x.restantes === 'number' && x.restantes > CONFIG.UMBRAL_CRITICO && x.restantes <= CONFIG.UMBRAL_AVISO);

  // pintar semáforo en Resumen H por fila real (no por índice filtrado)
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.HOJA_RESUMEN);
  estados.forEach((x) => {
    const cell = sh.getRange(x.fila, 8);
    if (typeof x.restantes !== 'number') return;
    if (x.restantes <= CONFIG.UMBRAL_CRITICO) cell.setBackground('#FFC7CE');
    else if (x.restantes <= CONFIG.UMBRAL_AVISO) cell.setBackground('#FFEB9C');
    else cell.setBackground('#C6EFCE');
  });

  if (!criticos.length && !avisos.length) {
    if (!esAutomatico) SpreadsheetApp.getActive().toast('✅ Todos los equipos con >50 hrs restantes.', 'SHOWPRO', 5);
    return;
  }

  const asunto = `🚨 SHOWPRO Generadores: ${criticos.length ? criticos.length + ' VENCIDO(S)' : ''} ${avisos.length ? avisos.length + ' por vencer' : ''}`;
  const cuerpo = ['BITÁCORA SHOWPRO — ' + Utilities.formatDate(new Date(), 'America/Santiago', 'dd-MMM-yyyy HH:mm'),
    '',
    ...criticos.map(x => `🔴 ${x.equipo}: ${x.restantes} hrs restantes (último ${x.ultimoHorometro}, desde servicio ${x.desde}, intervalo ${x.intervalo}). ¡HACER PREVENTIVA!`),
    ...avisos.map(x => `🟡 ${x.equipo}: ${x.restantes} hrs restantes (último ${x.ultimoHorometro}). Programar servicio.`),
    '',
    'Ver Hoja 3 Resumen de Equipos.'].join('\n');

  enviarEmail(asunto, cuerpo);
  enviarWhatsApp(cuerpo);
  if (!esAutomatico) SpreadsheetApp.getActive().toast(cuerpo, asunto, 10);
}

function enviarEmail(asunto, cuerpo) {
  if (!CONFIG.EMAIL_ALERTA) return;
  MailApp.sendEmail(CONFIG.EMAIL_ALERTA, asunto, cuerpo);
}

function enviarWhatsApp(texto) {
  const keySinConfig = !CONFIG.WHATSAPP_APIKEY || CONFIG.WHATSAPP_APIKEY.indexOf('PEGUE_AQUI') !== -1;
  if (!CONFIG.WHATSAPP_PHONE || keySinConfig) {
    Logger.log('WhatsApp no configurado. Mensaje que se habría enviado:\n' + texto);
    return;
  }
  const url = 'https://api.callmebot.com/whatsapp.php?phone=' + encodeURIComponent(CONFIG.WHATSAPP_PHONE)
    + '&text=' + encodeURIComponent(texto.substring(0, 1500))
    + '&apikey=' + encodeURIComponent(CONFIG.WHATSAPP_APIKEY);
  try {
    UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  } catch (err) {
    Logger.log('Error WhatsApp: ' + err);
  }
}

function enviarReporte() {
  const estados = getEstadoResumen();
  const lineas = estados.map(x => `• ${x.equipo}: último ${x.ultimoHorometro} hrs (${x.ultimaFecha instanceof Date ? Utilities.formatDate(x.ultimaFecha, 'America/Santiago', 'dd-MMM-yyyy') : x.ultimaFecha}) | 30d: ${x.horas30d} hrs | desde servicio: ${x.desde} | restantes: ${x.restantes}`);
  const cuerpo = ['REPORTE SHOWPRO — ' + Utilities.formatDate(new Date(), 'America/Santiago', 'dd-MMM-yyyy HH:mm'), '', ...lineas].join('\n');
  enviarEmail('📋 SHOWPRO Reporte generadores', cuerpo);
  SpreadsheetApp.getUi().alert('Reporte enviado a ' + CONFIG.EMAIL_ALERTA);
}
