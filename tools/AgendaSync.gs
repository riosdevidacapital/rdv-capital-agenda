/**
 * Agenda RDV Capital: sincroniza la agenda de Google Sheets a GitHub Pages.
 * Instalación: desde la planilla, Extensiones > Apps Script.
 * 1) Pegá este código.
 * 2) Ejecutá configurarTokenGitHub() una sola vez y cargá un token fine-grained
 *    con permiso Contents: Read and Write SOLO para riosdevidacapital/rdv-capital-agenda.
 * 3) Ejecutá sincronizarAgendaRDV() y verificá los archivos en GitHub.
 * 4) Ejecutá instalarSincronizacionRDV() para crear disparador cada 10 minutos.
 * El token queda en las propiedades del script, nunca en archivos públicos.
 */
const RDV = Object.freeze({
  SPREADSHEET_ID: '1R6EQWKFzXlNKZ4KNwNWMCCeY-DYbrw3E0NrLTZogWgk',
  REPO: 'riosdevidacapital/rdv-capital-agenda',
  BRANCH: 'main',
  TIMEZONE: 'America/Argentina/Buenos_Aires',
  PROPERTY_TOKEN: 'RDV_GITHUB_TOKEN',
  ROLES: ['Dirige', 'Predica', 'Santa Cena', 'Ofrendas', 'Anuncios', 'Cierre'],
  FIRST_COLUMN: 1, // B = lunes; N = domingo
  LAST_COLUMN: 13,
});

function configurarTokenGitHub() {
  const ui = SpreadsheetApp.getUi();
  const input = ui.prompt('GitHub - acceso para sincronización',
    'Pegá tu token fine-grained (Contents: Read and Write, solamente este repositorio). No se publicará en la web.',
    ui.ButtonSet.OK_CANCEL);
  if (input.getSelectedButton() !== ui.Button.OK) return;
  const token = input.getResponseText().trim();
  if (!/^(github_pat_|ghp_)/.test(token)) throw new Error('El token ingresado no parece ser de GitHub.');
  PropertiesService.getScriptProperties().setProperty(RDV.PROPERTY_TOKEN, token);
  ui.alert('Token guardado en las propiedades del script.');
}

function instalarSincronizacionRDV() {
  if (!PropertiesService.getScriptProperties().getProperty(RDV.PROPERTY_TOKEN)) {
    throw new Error('Primero ejecutá configurarTokenGitHub().');
  }
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sincronizarAgendaRDV')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sincronizarAgendaRDV').timeBased().everyMinutes(10).create();
  Logger.log('Sincronización instalada (intervalo aproximado de 10 minutos).');
}

function desactivarSincronizacionRDV() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sincronizarAgendaRDV')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

function sincronizarAgendaRDV() {
  const token = PropertiesService.getScriptProperties().getProperty(RDV.PROPERTY_TOKEN);
  if (!token) throw new Error('Falta RDV_GITHUB_TOKEN. Ejecutá configurarTokenGitHub().');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) { Logger.log('Otra ejecución está en curso.'); return; }
  try {
    const ss = SpreadsheetApp.openById(RDV.SPREADSHEET_ID);
    const now = new Date();
    const year = Number(Utilities.formatDate(now, RDV.TIMEZONE, 'yyyy'));
    const month = Number(Utilities.formatDate(now, RDV.TIMEZONE, 'MM'));
    [-1, 0, 1].forEach(offset => {
      const date = new Date(Date.UTC(year, month - 1 + offset, 1));
      const yy = date.getUTCFullYear();
      const mm = date.getUTCMonth() + 1;
      const sheetName = String(mm).padStart(2, '0') + '-' + yy;
      const sh = ss.getSheetByName(sheetName);
      if (!sh) { Logger.log('Omitido: no existe pestaña ' + sheetName); return; }
      const result = construirMesRDV_(sh, yy, mm);
      const path = 'data/' + yy + '-' + String(mm).padStart(2, '0') + '.json';
      publicarJSON_(path, result, token);
    });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Lee la disposición horizontal de AGENDA RDV CAPITAL 2026.
 * B/C = lunes, D/E = martes, ... N/O = domingo.
 * Cada bloque comienza en una fila con números de días;
 * las siguientes seis filas son funciones y las siguientes tres son anotaciones.
 */
function construirMesRDV_(sheet, year, month) {
  const count = Math.min(Math.max(sheet.getLastRow(), 55), 110);
  const rows = sheet.getRange(1, 1, count, 16).getDisplayValues();
  const found = new Map();
  const daysInMonth = new Date(year, month, 0).getDate();
  for (let header = 0; header < rows.length; header++) {
    for (let col = RDV.FIRST_COLUMN; col <= RDV.LAST_COLUMN; col += 2) {
      const raw = (rows[header][col] || '').trim();
      if (!/^\d{1,2}$/.test(raw)) continue;
      const day = Number(raw);
      if (day < 1 || day > daysInMonth || found.has(day)) continue;
      // Valida que el día corresponda al día de la semana de su columna.
      const actualWeekday = new Date(Date.UTC(year, month - 1, day)).getUTCDay();
      const expectedWeekday = ((col - 1) / 2 + 1) % 7; // lunes=1 ... domingo=0
      if (actualWeekday !== expectedWeekday) continue;
      if (header + 6 >= rows.length) continue;
      const assignments = [];
      for (let offset = 1; offset <= 6; offset++) {
        const name = (rows[header + offset][col] || '').trim();
        const statusRaw = (rows[header + offset][col + 1] || '').trim();
        // Incluimos los lugares pendientes aunque no tengan persona asignada.
        if (!name && !statusRaw) continue;
        assignments.push({
          rol: RDV.ROLES[offset - 1],
          nombre: name,
          estado: normalizarEstadoRDV_(statusRaw, name)
        });
      }
      for (let offset = 7; offset <= 9 && header + offset < rows.length; offset++) {
        const name = (rows[header + offset][col] || '').trim();
        if (!name) continue;
        assignments.push({rol: 'Actividad', nombre: name, estado: 'Programado'});
      }
      if (assignments.length) {
        found.set(day, {
          fecha: year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0'),
          asignaciones: assignments
        });
      }
    }
  }
  if (found.size === 0) throw new Error('No se detectaron días: ' + sheet.getName() + '. Publicación cancelada.');
  return {dias: Array.from(found.keys()).sort((a, b) => a - b).map(d => found.get(d))};
}

function normalizarEstadoRDV_(status, name) {
  const s = String(status || '').trim().toLowerCase();
  if (s.includes('confirm')) return 'Confirmado';
  if (s.includes('inform')) return 'Informado';
  if (s.includes('pendient')) return 'Pendiente';
  return name ? 'Pendiente' : 'Pendiente';
}

function publicarJSON_(path, data, token) {
  const api = 'https://api.github.com/repos/' + RDV.REPO + '/contents/' + path;
  const headers = {
    Authorization: 'Bearer ' + token,
    Accept: 'application/vnd.github+json',
    'X-GitHub-Api-Version': '2022-11-28'
  };
  const get = UrlFetchApp.fetch(api + '?ref=' + encodeURIComponent(RDV.BRANCH),
    {headers, muteHttpExceptions: true});
  const code = get.getResponseCode();
  if (code !== 200 && code !== 404) throw new Error('GitHub GET ' + code + ' ' + path + ': ' + get.getContentText().slice(0, 300));
  const existing = code === 200 ? JSON.parse(get.getContentText()) : null;
  const content = JSON.stringify(data, null, 2) + '\n';
  // Si no cambió el contenido, evitamos commits y despliegues innecesarios.
  if (existing && existing.content) {
    const previous = Utilities.newBlob(Utilities.base64Decode(existing.content.replace(/\s/g, ''))).getDataAsString('UTF-8');
    try {
      if (JSON.stringify(JSON.parse(previous)) === JSON.stringify(data)) {
        Logger.log('Sin cambios: ' + path); return;
      }
    } catch (e) { /* Publicamos si el archivo anterior no es JSON válido. */ }
  }
  const payload = {
    message: 'Sincronizar agenda desde Google Sheets: ' + path,
    content: Utilities.base64Encode(Utilities.newBlob(content, 'application/json').getBytes()),
    branch: RDV.BRANCH
  };
  if (existing) payload.sha = existing.sha;
  const put = UrlFetchApp.fetch(api, {
    method: 'put', contentType: 'application/json', headers,
    payload: JSON.stringify(payload), muteHttpExceptions: true
  });
  if (![200, 201].includes(put.getResponseCode())) {
    throw new Error('GitHub PUT ' + put.getResponseCode() + ' ' + path + ': ' + put.getContentText().slice(0, 300));
  }
  Logger.log('Publicado: ' + path);
}
