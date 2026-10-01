/**
 * GRUPO DORADOS - Sistema de Usuarios + Chart de Disponibilidad
 * Google Apps Script
 *
 * ARCHIVO MAESTRO:
 *   ID: 1-Phs--5CfjhZ3oJDlr5zpNSDWnfZ-TErCxn1xOp66E4
 *   Celda A2: URL (o ID) de la hoja de cálculo operativa.
 *
 * HOJA DE USUARIOS en el archivo operativo:
 *   usuarios
 *   Id | nombre | email | contraseña
 *
 * La columna "contraseña" guarda: salt:hashSHA256
 */

const CONFIG = {
  MASTER_SPREADSHEET_ID: '1-Phs--5CfjhZ3oJDlr5zpNSDWnfZ-TErCxn1xOp66E4',
  MASTER_URL_CELL: 'A2',
  USERS_SHEET: 'usuarios',
  SYSTEM_SHEET: 'sistema',
  LICENSE_CELL: 'A2',
  BLOCKS_SHEET: 'bloqueos',
  BLOCKS_HISTORY_SHEET: 'historico_bloqueos_cancelados',
  SESSION_DAYS: 7,
  RESET_MINUTES: 60,
  APP_NAME: 'Grupo Dorados | Disponibilidad de Salones',
  // Si tus pestañas mensuales tienen otro formato, agrega variantes aquí.
  MONTHS_ES: [
    'ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO',
    'JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'
  ]
};

const SALONS = [
  "ALFA 1 - 6",
  "ALFA 2 - 7",
  "ALFA 3 - 8",
  "ALFA 4 - 9",
  "CENTRO DE NEGOCIOS 1 ALFA",
  "BETA 1 - 7",
  "BETA 2 - 8",
  "BETA 3 - 9",
  "BETA 4 - 10",
  "BETA 5 - 11",
  "BETA 6 - 12",
  "CENTRO DE NEGOCIOS 1 BETA",
  "CENTRO DE NEGOCIOS 2 BETA",
  "DELTA 1",
  "IMPERIAL 1",
  "IMPERIAL 2",
  "IMPERIAL 3",
  "ZONA FOGATAS IMPERIAL",
  "VIDEO BAR",
  "EXPO",
  "BODEGA CISNES 1",
  "BODEGA CISNES 2",
  "BODEGA CISNES 3",
  "BODEGA CISNES 4",
  "BODEGA CISNES 5",
  "ZONA FOGATAS CISNES",
  "ZONA SNACK CISNES",
  "POLYFORUM",
  "ALBERCA SEMIOLIMPICA",
  "JARDIN CASARETT",
  "PALAPA",
  "ORQUIDEAS ARRIBA",
  "ORQUIDEAS ABAJO",
  "ARCOS",
  "PLANCHA DE LAURELES",
  "CAPILLA"
];

function doGet(e) {
  const tpl = HtmlService.createTemplateFromFile('index');
  tpl.resetToken = (e && e.parameter && e.parameter.reset) ? String(e.parameter.reset) : '';
  tpl.appName = CONFIG.APP_NAME;
  return tpl.evaluate()
    .setTitle(CONFIG.APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* =========================
   CONFIG / SPREADSHEETS
   ========================= */

function getOperationalSpreadsheet_() {
  const master = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
  const raw = String(master.getSheets()[0].getRange(CONFIG.MASTER_URL_CELL).getDisplayValue() || '').trim();
  if (!raw) throw new Error('La celda A2 del archivo maestro no contiene la URL/ID de la hoja operativa.');

  const id = extractSpreadsheetId_(raw);
  if (!id) throw new Error('No fue posible obtener el ID de la hoja de cálculo desde A2.');
  return SpreadsheetApp.openById(id);
}

function extractSpreadsheetId_(text) {
  const s = String(text || '').trim();
  const match = s.match(/[-\w]{25,}/);
  return match ? match[0] : '';
}

function checkLicense_() {
  const ss = getOperationalSpreadsheet_();
  const sh = ss.getSheetByName(CONFIG.SYSTEM_SHEET);
  if (!sh) {
    return { ok: false, expired: false, message: 'No existe la hoja "sistema" para validar la licencia.' };
  }

  const cell = sh.getRange(CONFIG.LICENSE_CELL);
  const value = cell.getValue();
  let expiry = null;

  if (value instanceof Date && !isNaN(value.getTime())) {
    expiry = value;
  } else {
    const raw = String(cell.getDisplayValue() || value || '').trim();
    let m;
    if ((m = raw.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})$/))) {
      expiry = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    } else if ((m = raw.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})$/))) {
      expiry = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
    } else {
      const parsed = new Date(raw);
      if (!isNaN(parsed.getTime())) expiry = parsed;
    }
  }

  if (!expiry || isNaN(expiry.getTime())) {
    return { ok: false, expired: false, message: 'La fecha de licencia en sistema!A2 no es válida.' };
  }

  const tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
  const todayKey = Number(Utilities.formatDate(new Date(), tz, 'yyyyMMdd'));
  const expiryKey = Number(Utilities.formatDate(expiry, tz, 'yyyyMMdd'));

  if (todayKey <= expiryKey) {
    return { ok: true, expired: false, expiry: Utilities.formatDate(expiry, tz, 'dd/MM/yyyy') };
  }

  return {
    ok: false,
    expired: true,
    expiry: Utilities.formatDate(expiry, tz, 'dd/MM/yyyy'),
    message: 'Licencia vencida, ponganse en contacto con el administrador'
  };
}

function ensureUsersSheet_() {
  const ss = getOperationalSpreadsheet_();
  let sh = ss.getSheetByName(CONFIG.USERS_SHEET);
  if (!sh) sh = ss.insertSheet(CONFIG.USERS_SHEET);

  const headers = ['Id', 'nombre', 'email', 'contraseña', 'tipo', 'estado'];

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#6d28d9')
      .setFontColor('#ffffff');
    sh.setFrozenRows(1);
  } else {
    headers.forEach((header, i) => {
      if (!String(sh.getRange(1, i + 1).getValue() || '').trim()) {
        sh.getRange(1, i + 1).setValue(header);
      }
    });

    // Compatibilidad: usuarios existentes pasan a Administrador/Activo.
    const lastRow = sh.getLastRow();
    if (lastRow > 1) {
      const roles = sh.getRange(2, 5, lastRow - 1, 2).getValues();
      let changed = false;
      roles.forEach(r => {
        if (!String(r[0] || '').trim()) { r[0] = 'administrador'; changed = true; }
        if (!String(r[1] || '').trim()) { r[1] = 'ACTIVO'; changed = true; }
      });
      if (changed) sh.getRange(2, 5, lastRow - 1, 2).setValues(roles);
    }
  }
  return sh;
}

function ensureBlocksSheet_() {
  const ss = getOperationalSpreadsheet_();
  let sh = ss.getSheetByName(CONFIG.BLOCKS_SHEET);
  if (!sh) sh = ss.insertSheet(CONFIG.BLOCKS_SHEET);

  // Se conserva el orden de las 11 columnas originales y se agregan
  // columnas nuevas al final para no afectar registros existentes.
  const headers = [
    'Id','evento','fechaEntrada','fechaSalida','status','seleccion',
    'estado','usuarioId','usuarioNombre','creado','actualizado',
    'fechaLimite','motivoCancelacion','fechaCancelacion'
  ];

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#6d28d9')
      .setFontColor('#ffffff');
    sh.setFrozenRows(1);
  } else {
    // Actualiza/añade encabezados sin borrar datos existentes.
    headers.forEach((header, i) => {
      if (!String(sh.getRange(1, i + 1).getValue() || '').trim()) {
        sh.getRange(1, i + 1).setValue(header);
      }
    });
  }

  return sh;
}

function ensureBlocksHistorySheet_() {
  const ss = getOperationalSpreadsheet_();
  let sh = ss.getSheetByName(CONFIG.BLOCKS_HISTORY_SHEET);
  if (!sh) sh = ss.insertSheet(CONFIG.BLOCKS_HISTORY_SHEET);

  const headers = [
    'Id','evento','fechaEntrada','fechaSalida','status','fechaLimite',
    'seleccion','usuarioCreacion','fechaCreacion','fechaUltimaActualizacion',
    'canceladoPor','motivo','fechaHoraCancelacion','usuarioIdCreacion'
  ];

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4b5563')
      .setFontColor('#ffffff');
    sh.setFrozenRows(1);
  } else {
    // Compatibilidad con históricos existentes: agrega el identificador del dueño
    // sin modificar ni desplazar la información ya guardada.
    headers.forEach((header, i) => {
      if (!String(sh.getRange(1, i + 1).getValue() || '').trim()) {
        sh.getRange(1, i + 1).setValue(header);
      }
    });
  }

  return sh;
}

/* =========================
   AUTH - PUBLIC API
   ========================= */

function registerUser(payload) {
  // El registro público queda reservado únicamente para crear la primera cuenta.
  // Después, los usuarios se crean desde Cuenta por un Administrador.
  payload = payload || {};
  const nombre = cleanText_(payload.nombre, 120);
  const email = normalizeEmail_(payload.email);
  const password = String(payload.password || '');

  if (!nombre || !email || !password) return fail_('Completa nombre, email y contraseña.');
  if (!isValidEmail_(email)) return fail_('Ingresa un correo electrónico válido.');
  const passError = validatePassword_(password);
  if (passError) return fail_(passError);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ensureUsersSheet_();
    const users = getUsersData_(sh);
    if (users.length) return fail_('El registro público está deshabilitado. Solicita a un Administrador que cree tu usuario.');

    const id = Utilities.getUuid();
    const stored = makePasswordRecord_(password);
    sh.appendRow([id, nombre, email, stored, 'administrador', 'ACTIVO']);

    const token = createSession_(id, nombre, email, 'administrador');
    return ok_({
      message: 'Cuenta Administrador creada.',
      token,
      user: { id, nombre, email, tipo: 'administrador', estado: 'ACTIVO' }
    });
  } finally {
    lock.releaseLock();
  }
}

function loginUser(payload) {
  payload = payload || {};
  const email = normalizeEmail_(payload.email);
  const password = String(payload.password || '');
  if (!email || !password) return fail_('Ingresa tu email y contraseña.');

  const license = checkLicense_();
  if (!license.ok) {
    if (license.expired) return { ok: false, licenseExpired: true, message: license.message, expiry: license.expiry };
    return fail_(license.message);
  }

  const sh = ensureUsersSheet_();
  const user = getUsersData_(sh).find(u => u.email === email);
  if (!user || !verifyPassword_(password, user.passwordRecord)) {
    return fail_('Email o contraseña incorrectos.');
  }
  if (user.estado !== 'ACTIVO') return fail_('Este usuario está inactivo. Contacta a un Administrador.');

  const token = createSession_(user.id, user.nombre, user.email, user.tipo);
  return ok_({
    message: 'Sesión iniciada.',
    token,
    user: publicUser_(user)
  });
}

function validateSession(token) {
  const license = checkLicense_();
  if (!license.ok) {
    if (license.expired) {
      if (token) deleteSession_(String(token));
      return { ok: false, licenseExpired: true, message: license.message, expiry: license.expiry };
    }
    return fail_(license.message);
  }

  const session = getSession_(token);
  if (!session) return fail_('La sesión expiró. Inicia sesión nuevamente.');

  const sh = ensureUsersSheet_();
  const user = getUsersData_(sh).find(u => String(u.id) === String(session.id));
  if (!user || user.estado !== 'ACTIVO') {
    deleteSession_(token);
    return fail_('La cuenta ya no está activa.');
  }
  return ok_({ user: publicUser_(user) });
}

function logoutUser(token) {
  // Cierra explícitamente la sesión del usuario en el servidor.
  if (token) deleteSession_(String(token));
  return ok_({ message: 'Sesión cerrada correctamente.' });
}

function requestPasswordReset(email) {
  email = normalizeEmail_(email);
  // Respuesta genérica para no revelar si una cuenta existe.
  const generic = ok_({ message: 'Si el correo está registrado, recibirás un enlace para cambiar tu contraseña.' });
  if (!email || !isValidEmail_(email)) return generic;

  const sh = ensureUsersSheet_();
  const user = getUsersData_(sh).find(u => u.email === email);
  if (!user) return generic;

  const token = createSecureToken_();
  const tokenHash = sha256Hex_(token);
  const expiresAt = Date.now() + CONFIG.RESET_MINUTES * 60 * 1000;
  PropertiesService.getScriptProperties().setProperty(
    'RESET_' + tokenHash,
    JSON.stringify({ id: user.id, email: user.email, expiresAt })
  );

  const baseUrl = ScriptApp.getService().getUrl();
  const resetUrl = baseUrl + '?reset=' + encodeURIComponent(token);
  const subject = 'Recuperación de contraseña - Grupo Dorados';
  const htmlBody = `
    <div style="font-family:Arial,sans-serif;color:#2b2140;line-height:1.55">
      <h2 style="color:#6d28d9">Grupo Dorados</h2>
      <p>Hola ${escapeHtml_(user.nombre)},</p>
      <p>Recibimos una solicitud para cambiar tu contraseña.</p>
      <p><a href="${resetUrl}" style="display:inline-block;background:#6d28d9;color:#fff;text-decoration:none;padding:12px 18px;border-radius:8px;font-weight:700">Cambiar contraseña</a></p>
      <p>Este enlace vence en ${CONFIG.RESET_MINUTES} minutos.</p>
      <p>Si no solicitaste este cambio, puedes ignorar este correo.</p>
    </div>`;

  MailApp.sendEmail({
    to: user.email,
    subject,
    htmlBody,
    name: 'Grupo Dorados'
  });
  return generic;
}

function resetPassword(payload) {
  payload = payload || {};
  const token = String(payload.token || '');
  const password = String(payload.password || '');
  const passError = validatePassword_(password);
  if (!token) return fail_('El enlace de recuperación no es válido.');
  if (passError) return fail_(passError);

  const tokenHash = sha256Hex_(token);
  const key = 'RESET_' + tokenHash;
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(key);
  if (!raw) return fail_('El enlace de recuperación no es válido o ya fue utilizado.');

  let reset;
  try { reset = JSON.parse(raw); } catch (err) { reset = null; }
  if (!reset || Date.now() > Number(reset.expiresAt || 0)) {
    props.deleteProperty(key);
    return fail_('El enlace de recuperación ha vencido. Solicita uno nuevo.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ensureUsersSheet_();
    const data = sh.getDataRange().getValues();
    let foundRow = 0;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(reset.id) && normalizeEmail_(data[i][2]) === normalizeEmail_(reset.email)) {
        foundRow = i + 1;
        break;
      }
    }
    if (!foundRow) return fail_('No fue posible localizar la cuenta.');

    sh.getRange(foundRow, 4).setValue(makePasswordRecord_(password));
    props.deleteProperty(key);
    revokeAllSessionsForUser_(reset.id);
    return ok_({ message: 'Contraseña actualizada. Ya puedes iniciar sesión.' });
  } finally {
    lock.releaseLock();
  }
}

/* =========================
   ACCOUNT / USER MANAGEMENT
   ========================= */

function getAccountData(token) {
  const session = getSession_(token);
  if (!session) return fail_('Sesión expirada.');

  const sh = ensureUsersSheet_();
  const users = getUsersData_(sh);
  const me = users.find(u => String(u.id) === String(session.id));
  if (!me || me.estado !== 'ACTIVO') return fail_('Usuario no disponible.');

  return ok_({
    user: publicUser_(me),
    users: me.tipo === 'administrador' ? users.map(publicUser_) : []
  });
}

function changeOwnPassword(token, payload) {
  const session = getSession_(token);
  if (!session) return fail_('Sesión expirada.');
  payload = payload || {};

  const currentPassword = String(payload.currentPassword || '');
  const newPassword = String(payload.newPassword || '');
  const passError = validatePassword_(newPassword);
  if (!currentPassword || !newPassword) return fail_('Completa la contraseña actual y la nueva contraseña.');
  if (passError) return fail_(passError);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ensureUsersSheet_();
    const user = getUsersData_(sh).find(u => String(u.id) === String(session.id));
    if (!user) return fail_('No se encontró el usuario.');
    if (!verifyPassword_(currentPassword, user.passwordRecord)) return fail_('La contraseña actual es incorrecta.');

    sh.getRange(user.row, 4).setValue(makePasswordRecord_(newPassword));
    revokeAllSessionsForUser_(user.id);
    return ok_({ message: 'Contraseña actualizada. Por seguridad, inicia sesión nuevamente.' });
  } finally {
    lock.releaseLock();
  }
}

function adminCreateUser(token, payload) {
  const admin = requireAdmin_(token);
  if (!admin.ok) return admin;

  payload = payload || {};
  const nombre = cleanText_(payload.nombre, 120);
  const email = normalizeEmail_(payload.email);
  const password = String(payload.password || '');
  const tipo = normalizeUserType_(payload.tipo);

  if (!nombre || !email || !password || !tipo) return fail_('Completa nombre, email, contraseña y tipo de usuario.');
  if (!isValidEmail_(email)) return fail_('Ingresa un correo electrónico válido.');
  const passError = validatePassword_(password);
  if (passError) return fail_(passError);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ensureUsersSheet_();
    const users = getUsersData_(sh);
    if (users.some(u => u.email === email)) return fail_('Ese correo ya está registrado.');

    const id = Utilities.getUuid();
    sh.appendRow([id, nombre, email, makePasswordRecord_(password), tipo, 'ACTIVO']);
    return ok_({ message: 'Usuario agregado correctamente.', user: { id, nombre, email, tipo, estado: 'ACTIVO' } });
  } finally {
    lock.releaseLock();
  }
}

function adminDeleteUser(token, userId) {
  const admin = requireAdmin_(token);
  if (!admin.ok) return admin;
  userId = String(userId || '').trim();
  if (!userId) return fail_('Selecciona un usuario.');
  if (String(admin.user.id) === userId) return fail_('No puedes eliminar tu propio usuario Administrador.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ensureUsersSheet_();
    const users = getUsersData_(sh);
    const target = users.find(u => String(u.id) === userId);
    if (!target) return fail_('El usuario ya no existe.');

    sh.deleteRow(target.row);
    revokeAllSessionsForUser_(target.id);
    return ok_({ message: 'Usuario eliminado correctamente.' });
  } finally {
    lock.releaseLock();
  }
}

function requireAdmin_(token) {
  const session = getSession_(token);
  if (!session) return fail_('Sesión expirada.');
  const sh = ensureUsersSheet_();
  const user = getUsersData_(sh).find(u => String(u.id) === String(session.id));
  if (!user || user.estado !== 'ACTIVO') return fail_('Usuario no disponible.');
  if (user.tipo !== 'administrador') return fail_('Esta opción es exclusiva para usuarios Administrador.');
  return ok_({ user: publicUser_(user) });
}

function publicUser_(user) {
  return {
    id: String(user.id || ''),
    nombre: String(user.nombre || ''),
    email: String(user.email || ''),
    tipo: normalizeUserType_(user.tipo) || 'operador',
    estado: String(user.estado || 'ACTIVO').toUpperCase()
  };
}

function normalizeUserType_(value) {
  const s = String(value || '').trim().toLowerCase();
  if (['administrador','admin'].includes(s)) return 'administrador';
  if (['operador','operator'].includes(s)) return 'operador';
  if (['consulta','consultor','visualizador','viewer'].includes(s)) return 'consulta';
  return '';
}

/* =========================
   AVAILABILITY CHART
   ========================= */

function getAvailability(token, year, month) {
  const session = getSession_(token);
  if (!session) return fail_('Sesión expirada.');

  processExpiredTentatives_(false);

  year = Number(year);
  month = Number(month);

  if (!Number.isInteger(year) || year < 2020 || year > 2100 || !Number.isInteger(month) || month < 1 || month > 12) {
    return fail_('Mes o año no válido.');
  }

  const ss = getOperationalSpreadsheet_();
  const monthName = CONFIG.MONTHS_ES[month - 1];
  const sh = findMonthSheet_(ss, monthName, year);
  const days = daysArray_(year, month);
  const daysInMonth = days.length;

  const salonMap = {};
  SALONS.forEach(name => {
    salonMap[normalizeName_(name)] = {
      name,
      cells: Array.from({length: daysInMonth}, (_, i) => ({
        day: i + 1,
        status: 'available',
        code: ''
      }))
    };
  });

  let sheetFound = false;
  let sheetName = '';

  if (sh) {
    const grid = sh.getDataRange();
    const values = grid.getDisplayValues();
    const backgrounds = grid.getBackgrounds();
    const headerInfo = detectDayHeader_(values, year, month);

    if (headerInfo) {
      sheetFound = true;
      sheetName = sh.getName();

      const stopNames = ['DISPONIBLES','DEFINITIVOS','TENTATIVOS'];

      for (let r = headerInfo.headerRow + 1; r < values.length; r++) {
        const salonRaw = String(values[r][0] || '').trim();
        if (!salonRaw) continue;

        const salonKey = normalizeName_(salonRaw);
        if (stopNames.some(s => salonKey.includes(s))) break;

        if (!salonMap[salonKey]) {
          salonMap[salonKey] = {
            name: salonRaw,
            cells: Array.from({length: daysInMonth}, (_, i) => ({
              day: i + 1,
              status: 'available',
              code: ''
            }))
          };
        }

        for (let d = 1; d <= daysInMonth; d++) {
          const c = headerInfo.firstDayCol + d - 1;
          if (c >= values[r].length) continue;

          const codeVal = String(values[r][c] || '').trim();
          const bg = normalizeHex_(backgrounds[r][c] || '#ffffff');

          salonMap[salonKey].cells[d - 1] = {
            day: d,
            status: classifyStatus_(bg, codeVal),
            code: codeVal,
            background: bg
          };
        }
      }
    }
  }

  const blocks = getActiveBlocks_();
  blocks.forEach(block => {
    (block.selections || []).forEach(sel => {
      const p = parseIsoDate_(sel.date);
      if (!p || p.year !== year || p.month !== month) return;

      const salon = salonMap[normalizeName_(sel.salon)];
      if (!salon || !salon.cells[p.day - 1]) return;

      const deadlineDue =
        block.status === 'tentative' &&
        !!block.fechaLimite &&
        block.fechaLimite <= todayIso_();

      salon.cells[p.day - 1] = {
        day: p.day,
        status: block.status,
        code: blockCode_(block.evento),
        blockId: block.id,
        blockEvent: block.evento,
        blockUser: block.usuarioNombre || '',
        blockDeadline: block.status === 'tentative' ? (block.fechaLimite || '') : '',
        deadlineDue: deadlineDue,
        managedBlock: true
      };
    });
  });

  const salonsOrdered = [];
  const seen = {};

  SALONS.forEach(name => {
    const key = normalizeName_(name);
    if (salonMap[key]) {
      salonsOrdered.push(salonMap[key]);
      seen[key] = true;
    }
  });

  Object.keys(salonMap).forEach(key => {
    if (!seen[key]) salonsOrdered.push(salonMap[key]);
  });

  const summary = {
    available: 0,
    occupied: 0,
    tentative: 0,
    guaranteed: 0,
    deadlineReached: 0
  };

  salonsOrdered.forEach(s => {
    s.cells.forEach(c => {
      summary[c.status] = (summary[c.status] || 0) + 1;
      if (c.deadlineDue) summary.deadlineReached++;
    });
  });

  return ok_({
    year,
    month,
    monthName,
    sheetFound,
    sheetName,
    days,
    salons: salonsOrdered,
    summary
  });
}

function getAvailableYears(token) {
  const session = getSession_(token);
  if (!session) return fail_('Sesión expirada.');

  // El selector del chart muestra únicamente 3 años:
  // año actual + los 2 años siguientes.
  const current = Number(
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyy'
    )
  );

  return ok_({
    years: [current, current + 1, current + 2]
  });
}

function findMonthSheet_(ss, monthName, year) {
  const normalizedTarget = normalizeName_(monthName + ' ' + year);
  const candidates = ss.getSheets();
  let exact = candidates.find(sh => normalizeName_(sh.getName()) === normalizedTarget);
  if (exact) return exact;

  exact = candidates.find(sh => {
    const n = normalizeName_(sh.getName());
    return n.includes(normalizeName_(monthName)) && n.includes(String(year));
  });
  return exact || null;
}

function detectDayHeader_(values, year, month) {
  const daysInMonth = new Date(year, month, 0).getDate();
  let best = null;

  for (let r = 0; r < Math.min(values.length, 12); r++) {
    for (let c = 1; c < Math.min((values[r] || []).length, 20); c++) {
      let score = 0;
      for (let d = 1; d <= daysInMonth; d++) {
        const idx = c + d - 1;
        if (idx >= values[r].length) break;
        const n = parseInt(String(values[r][idx]).trim(), 10);
        if (n === d) score++;
        else break;
      }
      if (!best || score > best.score) best = { headerRow: r, firstDayCol: c, score };
    }
  }
  if (!best || best.score < Math.min(20, daysInMonth)) return null;
  best.dayCols = best.score;
  return best;
}

function classifyStatus_(bg, code) {
  // Rojo = garantizado, Amarillo = tentativo.
  if (isRed_(bg)) return 'guaranteed';
  if (isYellow_(bg)) return 'tentative';

  // Blanco con contenido = ocupado. Blanco vacío = disponible.
  // Cualquier otro color con contenido se considera ocupado.
  if (code) return 'occupied';
  return 'available';
}

function isRed_(hex) {
  const rgb = hexToRgb_(hex);
  return rgb && rgb.r >= 180 && rgb.g <= 120 && rgb.b <= 120 && (rgb.r - rgb.g) >= 70;
}

function isYellow_(hex) {
  const rgb = hexToRgb_(hex);
  return rgb && rgb.r >= 180 && rgb.g >= 150 && rgb.b <= 140;
}

function normalizeHex_(hex) {
  let h = String(hex || '#ffffff').trim().toLowerCase();
  if (/^#[0-9a-f]{6}$/.test(h)) return h;
  if (/^#[0-9a-f]{3}$/.test(h)) return '#' + h[1]+h[1]+h[2]+h[2]+h[3]+h[3];
  return '#ffffff';
}

function hexToRgb_(hex) {
  const h = normalizeHex_(hex).slice(1);
  return { r: parseInt(h.slice(0,2),16), g: parseInt(h.slice(2,4),16), b: parseInt(h.slice(4,6),16) };
}

function dayLetter_(year, month, day) {
  const letters = ['D','L','M','M','J','V','S'];
  return letters[new Date(year, month - 1, day).getDay()];
}

function daysArray_(year, month) {
  const n = new Date(year, month, 0).getDate();
  return Array.from({length:n}, (_,i) => ({ day: i+1, dow: dayLetter_(year, month, i+1) }));
}

/* =========================
   BLOQUEOS DE SALONES
   ========================= */

function getSalonCatalog(token) {
  if (!getSession_(token)) return fail_('Sesión expirada.');
  return ok_({ salons: SALONS.slice() });
}

function getBlocks(token) {
  const session = getSession_(token);
  if (!session) return fail_('Sesión expirada.');

  processExpiredTentatives_(false);

  // Seguridad por propietario: las opciones Modificar y Cancelar reciben
  // únicamente los bloqueos creados por el usuario que tiene la sesión activa.
  const blocks = getAllBlocks_()
    .filter(b =>
      b.estado !== 'CANCELADO' &&
      b.status !== 'out_of_service' &&
      String(b.usuarioId || '') === String(session.id || '')
    )
    .sort((a, b) => String(b.actualizado || b.creado).localeCompare(String(a.actualizado || a.creado)));

  return ok_({ blocks });
}


function getActiveBlocksReport(token) {
  if (!getSession_(token)) return fail_('Sesión expirada.');

  // Procesa vencimientos antes de generar el reporte para no incluir
  // bloqueos tentativos que ya debieron pasar al histórico.
  processExpiredTentatives_(false);

  const blocks = getAllBlocks_()
    .filter(b => b.estado !== 'CANCELADO' && (b.status === 'tentative' || b.status === 'guaranteed'))
    .sort((a, b) => {
      const byDate = String(a.fechaEntrada || '').localeCompare(String(b.fechaEntrada || ''));
      if (byDate) return byDate;
      const byEvent = String(a.evento || '').localeCompare(String(b.evento || ''), 'es', { sensitivity: 'base' });
      if (byEvent) return byEvent;
      return String(a.usuarioNombre || '').localeCompare(String(b.usuarioNombre || ''), 'es', { sensitivity: 'base' });
    });

  const groupsMap = {};

  blocks.forEach(b => {
    const parsed = parseIsoDate_(b.fechaEntrada);
    const key = parsed
      ? String(parsed.year) + '-' + String(parsed.month).padStart(2, '0')
      : 'sin-fecha';

    if (!groupsMap[key]) {
      groupsMap[key] = {
        key,
        year: parsed ? parsed.year : 0,
        month: parsed ? parsed.month : 0,
        monthName: parsed ? CONFIG.MONTHS_ES[parsed.month - 1] : 'Sin fecha',
        tentative: 0,
        guaranteed: 0,
        blocks: []
      };
    }

    if (b.status === 'tentative') groupsMap[key].tentative++;
    if (b.status === 'guaranteed') groupsMap[key].guaranteed++;

    groupsMap[key].blocks.push({
      id: b.id,
      evento: b.evento,
      fechaEntrada: b.fechaEntrada,
      fechaSalida: b.fechaSalida,
      status: b.status,
      fechaLimite: b.fechaLimite || '',
      usuarioNombre: b.usuarioNombre || '',
      selections: b.selections || []
    });
  });

  const groups = Object.values(groupsMap).sort((a, b) => {
    if (a.key === 'sin-fecha') return 1;
    if (b.key === 'sin-fecha') return -1;
    return a.key.localeCompare(b.key);
  });

  return ok_({
    generatedAt: Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyy-MM-dd HH:mm:ss'
    ),
    totals: {
      tentative: blocks.filter(b => b.status === 'tentative').length,
      guaranteed: blocks.filter(b => b.status === 'guaranteed').length,
      total: blocks.length
    },
    groups
  });
}

function getCancelledBlocksReport(token, year, month) {
  const session = getSession_(token);
  if (!session) return fail_('Sesión expirada.');

  processExpiredTentatives_(false);

  year = Number(year);
  month = Number(month);

  if (
    !Number.isInteger(year) || year < 2020 || year > 2100 ||
    !Number.isInteger(month) || month < 1 || month > 12
  ) {
    return fail_('Mes o año no válido.');
  }

  const sh = ensureBlocksHistorySheet_();
  const values = sh.getDataRange().getValues();
  const rows = [];

  for (let i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;

    const fechaEntrada = formatSheetDateIso_(values[i][2]);
    const parsed = parseIsoDate_(fechaEntrada);

    if (!parsed || parsed.year !== year || parsed.month !== month) continue;

    // Solo el dueño puede consultar sus cancelaciones. Los registros nuevos
    // usan usuarioIdCreacion; para históricos anteriores se conserva
    // compatibilidad comparando el nombre del creador.
    const usuarioCreacion = String(values[i][7] || '');
    const usuarioIdCreacion = String(values[i][13] || '');
    const isOwner = usuarioIdCreacion
      ? usuarioIdCreacion === String(session.id || '')
      : normalizeName_(usuarioCreacion) === normalizeName_(session.nombre || '');

    if (!isOwner) continue;

    let selections = [];
    try {
      selections = JSON.parse(String(values[i][6] || '[]'));
    } catch (e) {}

    rows.push({
      id: String(values[i][0] || ''),
      evento: String(values[i][1] || ''),
      fechaEntrada,
      fechaSalida: formatSheetDateIso_(values[i][3]),
      status: normalizeBlockStatus_(values[i][4]) || String(values[i][4] || ''),
      fechaLimite: formatSheetDateIso_(values[i][5]),
      selections: normalizeSelections_(selections),
      usuarioCreacion,
      usuarioIdCreacion,
      fechaCreacion: dateTimeText_(values[i][8]),
      fechaUltimaActualizacion: dateTimeText_(values[i][9]),
      canceladoPor: String(values[i][10] || ''),
      motivo: String(values[i][11] || ''),
      fechaHoraCancelacion: dateTimeText_(values[i][12])
    });
  }

  rows.sort((a, b) =>
    String(b.fechaHoraCancelacion || '').localeCompare(String(a.fechaHoraCancelacion || ''))
  );

  return ok_({
    year,
    month,
    monthName: CONFIG.MONTHS_ES[month - 1],
    rows
  });
}


function getDashboardData(token) {
  const session = getSession_(token);
  if (!session) return fail_('Sesión expirada.');

  // Antes de calcular indicadores, procesa vencimientos pendientes.
  processExpiredTentatives_(false);

  const today = todayIso_();
  const active = getAllBlocks_().filter(b => b.estado !== 'CANCELADO' && (b.status === 'tentative' || b.status === 'guaranteed'));
  const cancelled = getCancelledHistory_().filter(b => b.status === 'tentative' || b.status === 'guaranteed');

  const users = {};

  function ensureUser_(name) {
    name = String(name || 'Sin usuario').trim() || 'Sin usuario';
    if (!users[name]) {
      users[name] = {
        usuario: name,
        tentative: 0,
        guaranteed: 0,
        cancelled: 0,
        deadlineReached: 0,
        oldestTentativeCreated: '',
        oldestTentativeEvent: '',
        oldestTentativeAgeDays: 0
      };
    }
    return users[name];
  }

  // Bloqueos activos.
  active.forEach(b => {
    const u = ensureUser_(b.usuarioNombre);

    if (b.status === 'tentative') {
      u.tentative++;

      if (b.fechaLimite && b.fechaLimite <= today) {
        u.deadlineReached++;
      }

      if (b.creado) {
        if (!u.oldestTentativeCreated || b.creado < u.oldestTentativeCreated) {
          u.oldestTentativeCreated = b.creado;
          u.oldestTentativeEvent = b.evento || '';
          u.oldestTentativeAgeDays = ageDaysFromDateTime_(b.creado);
        }
      }
    }

    if (b.status === 'guaranteed') {
      u.guaranteed++;
    }
  });

  // Cancelados: se atribuyen al usuario que originalmente creó el bloqueo.
  cancelled.forEach(b => {
    ensureUser_(b.usuarioCreacion).cancelled++;
  });

  const totalsByUser = Object.values(users).sort((a, b) =>
    a.usuario.localeCompare(b.usuario, 'es', { sensitivity: 'base' })
  );

  const tentativeAging = totalsByUser
    .filter(u => u.tentative > 0)
    .map(u => ({
      usuario: u.usuario,
      tentative: u.tentative,
      oldestTentativeCreated: u.oldestTentativeCreated,
      oldestTentativeEvent: u.oldestTentativeEvent,
      oldestTentativeAgeDays: u.oldestTentativeAgeDays
    }))
    .sort((a, b) => b.oldestTentativeAgeDays - a.oldestTentativeAgeDays);

  const oldestTentativeUser = tentativeAging.length ? tentativeAging[0].usuario : '';

  const cancelledByUser = totalsByUser
    .map(u => ({ usuario: u.usuario, cancelled: u.cancelled }))
    .sort((a, b) => b.cancelled - a.cancelled || a.usuario.localeCompare(b.usuario, 'es'));

  const mostCancelledUser =
    cancelledByUser.length && cancelledByUser[0].cancelled > 0
      ? cancelledByUser[0].usuario
      : '';

  const guaranteedByUser = totalsByUser
    .filter(u => u.guaranteed > 0)
    .map(u => ({ usuario: u.usuario, guaranteed: u.guaranteed }))
    .sort((a, b) => b.guaranteed - a.guaranteed || a.usuario.localeCompare(b.usuario, 'es'));

  const mostGuaranteedUser =
    guaranteedByUser.length ? guaranteedByUser[0].usuario : '';

  const leastGuaranteedUser =
    guaranteedByUser.length
      ? guaranteedByUser
          .slice()
          .sort((a, b) => a.guaranteed - b.guaranteed || a.usuario.localeCompare(b.usuario, 'es'))[0].usuario
      : '';

  const totals = {
    tentative: active.filter(b => b.status === 'tentative').length,
    guaranteed: active.filter(b => b.status === 'guaranteed').length,
    cancelled: cancelled.length,
    deadlineReached: active.filter(
      b => b.status === 'tentative' && b.fechaLimite && b.fechaLimite <= today
    ).length
  };

  return ok_({
    generatedAt: Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyy-MM-dd HH:mm:ss'
    ),
    totals,
    totalsByUser,
    tentativeAging,
    cancelledByUser,
    guaranteedByUser,
    highlights: {
      oldestTentativeUser,
      mostCancelledUser,
      mostGuaranteedUser,
      leastGuaranteedUser
    }
  });
}

function sendDashboardEmail(token, toEmail) {
  const session = getSession_(token);
  if (!session) return fail_('Sesión expirada.');

  toEmail = normalizeEmail_(toEmail);
  if (!toEmail || !isValidEmail_(toEmail)) {
    return fail_('Ingresa un correo electrónico válido.');
  }

  const dashboard = getDashboardData(token);
  if (!dashboard.ok) return dashboard;

  const d = dashboard;
  const rows = d.totalsByUser || [];

  const totalRows = rows.map(r =>
    '<tr>' +
      '<td style="padding:7px;border:1px solid #ddd;">' + htmlEscapeServer_(r.usuario) + '</td>' +
      '<td style="padding:7px;border:1px solid #ddd;text-align:center;">' + r.tentative + '</td>' +
      '<td style="padding:7px;border:1px solid #ddd;text-align:center;">' + r.guaranteed + '</td>' +
      '<td style="padding:7px;border:1px solid #ddd;text-align:center;">' + r.cancelled + '</td>' +
      '<td style="padding:7px;border:1px solid #ddd;text-align:center;">' + r.deadlineReached + '</td>' +
    '</tr>'
  ).join('');

  const agingRows = (d.tentativeAging || []).map(r =>
    '<tr>' +
      '<td style="padding:7px;border:1px solid #ddd;">' + htmlEscapeServer_(r.usuario) + '</td>' +
      '<td style="padding:7px;border:1px solid #ddd;text-align:center;">' + r.tentative + '</td>' +
      '<td style="padding:7px;border:1px solid #ddd;">' + htmlEscapeServer_(r.oldestTentativeEvent || '—') + '</td>' +
      '<td style="padding:7px;border:1px solid #ddd;">' + htmlEscapeServer_(r.oldestTentativeCreated || '—') + '</td>' +
      '<td style="padding:7px;border:1px solid #ddd;text-align:center;">' + r.oldestTentativeAgeDays + '</td>' +
    '</tr>'
  ).join('');

  const body =
    '<div style="font-family:Arial,sans-serif;color:#2d2436;">' +
      '<h2 style="color:#6d28d9;">Dashboard de Bloqueos · Grupo Dorados</h2>' +
      '<p>Generado: ' + htmlEscapeServer_(d.generatedAt) + '</p>' +
      '<p><b>Tentativos:</b> ' + d.totals.tentative +
      ' &nbsp; <b>Garantizados:</b> ' + d.totals.guaranteed +
      ' &nbsp; <b>Cancelados:</b> ' + d.totals.cancelled +
      ' &nbsp; <b>Fecha límite alcanzada:</b> ' + d.totals.deadlineReached + '</p>' +
      '<h3>Totales por usuario</h3>' +
      '<table style="border-collapse:collapse;width:100%;font-size:12px;">' +
        '<thead><tr style="background:#6d28d9;color:white;">' +
          '<th style="padding:8px;border:1px solid #ddd;">Usuario</th>' +
          '<th style="padding:8px;border:1px solid #ddd;">Tentativos</th>' +
          '<th style="padding:8px;border:1px solid #ddd;">Garantizados</th>' +
          '<th style="padding:8px;border:1px solid #ddd;">Cancelados</th>' +
          '<th style="padding:8px;border:1px solid #ddd;">Fecha límite alcanzada</th>' +
        '</tr></thead><tbody>' + totalRows + '</tbody></table>' +
      '<h3 style="margin-top:20px;">Antigüedad de tentativos</h3>' +
      '<table style="border-collapse:collapse;width:100%;font-size:12px;">' +
        '<thead><tr style="background:#4b5563;color:white;">' +
          '<th style="padding:8px;border:1px solid #ddd;">Usuario</th>' +
          '<th style="padding:8px;border:1px solid #ddd;">Tentativos</th>' +
          '<th style="padding:8px;border:1px solid #ddd;">Bloqueo más antiguo</th>' +
          '<th style="padding:8px;border:1px solid #ddd;">Fecha de ingreso</th>' +
          '<th style="padding:8px;border:1px solid #ddd;">Días</th>' +
        '</tr></thead><tbody>' + agingRows + '</tbody></table>' +
      '<p style="margin-top:20px;"><b>Mayor antigüedad tentativo:</b> ' +
        htmlEscapeServer_(d.highlights.oldestTentativeUser || 'Sin datos') + '<br>' +
        '<b>Más cancelados:</b> ' + htmlEscapeServer_(d.highlights.mostCancelledUser || 'Sin datos') + '<br>' +
        '<b>Más garantizados:</b> ' + htmlEscapeServer_(d.highlights.mostGuaranteedUser || 'Sin datos') + '<br>' +
        '<b>Menos garantizados:</b> ' + htmlEscapeServer_(d.highlights.leastGuaranteedUser || 'Sin datos') +
      '</p>' +
      '<p style="font-size:11px;color:#777;">Las gráficas se visualizan dentro de la aplicación web.</p>' +
    '</div>';

  MailApp.sendEmail({
    to: toEmail,
    subject: 'Dashboard de Bloqueos - Grupo Dorados',
    htmlBody: body,
    name: 'Grupo Dorados'
  });

  return ok_({ message: 'Dashboard enviado por correo a ' + toEmail + '.' });
}

function getCancelledHistory_() {
  const sh = ensureBlocksHistorySheet_();
  const values = sh.getDataRange().getValues();
  const rows = [];

  for (let i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;

    let selections = [];
    try {
      selections = JSON.parse(String(values[i][6] || '[]'));
    } catch (e) {}

    rows.push({
      id: String(values[i][0] || ''),
      evento: String(values[i][1] || ''),
      fechaEntrada: formatSheetDateIso_(values[i][2]),
      fechaSalida: formatSheetDateIso_(values[i][3]),
      status: normalizeBlockStatus_(values[i][4]) || String(values[i][4] || ''),
      fechaLimite: formatSheetDateIso_(values[i][5]),
      selections: normalizeSelections_(selections),
      usuarioCreacion: String(values[i][7] || ''),
      fechaCreacion: dateTimeText_(values[i][8]),
      fechaUltimaActualizacion: dateTimeText_(values[i][9]),
      canceladoPor: String(values[i][10] || ''),
      motivo: String(values[i][11] || ''),
      fechaHoraCancelacion: dateTimeText_(values[i][12])
    });
  }

  return rows;
}

function ageDaysFromDateTime_(text) {
  const m = String(text || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return 0;

  const start = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const today = parseIsoDate_(todayIso_());
  if (!today) return 0;

  const end = new Date(today.year, today.month - 1, today.day);
  return Math.max(0, Math.floor((end.getTime() - start.getTime()) / 86400000));
}

function htmlEscapeServer_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function getBlockingGrid(token, fechaEntrada, fechaSalida, excludeBlockId) {
  if (!getSession_(token)) return fail_('Sesión expirada.');

  const start = parseIsoDate_(fechaEntrada);
  const end = parseIsoDate_(fechaSalida);

  if (!start || !end) return fail_('Selecciona fecha de entrada y fecha de salida.');
  if (start.iso > end.iso) return fail_('La fecha de salida no puede ser anterior a la fecha de entrada.');

  const todayIso = todayIso_();
  if (start.iso < todayIso || end.iso < todayIso) {
    return fail_('Las fechas del bloqueo no pueden ser menores al día actual.');
  }

  const dates = isoDateRange_(start.iso, end.iso, 62);
  if (!dates) return fail_('El rango máximo para un bloqueo es de 62 días.');

  const ss = getOperationalSpreadsheet_();
  const monthCache = {};
  const blockMap = {};

  getActiveBlocks_()
    .filter(b => String(b.id) !== String(excludeBlockId || ''))
    .forEach(b => {
      (b.selections || []).forEach(sel => {
        blockMap[normalizeName_(sel.salon) + '|' + sel.date] = {
          status: b.status,
          evento: b.evento,
          blockId: b.id
        };
      });
    });

  const rows = SALONS.map(salon => {
    const cells = dates.map(dateIso => {
      const key = normalizeName_(salon) + '|' + dateIso;

      if (blockMap[key]) {
        return {
          date: dateIso,
          available: false,
          status: blockMap[key].status,
          detail: blockMap[key].evento,
          source: 'block'
        };
      }

      const base = getBaseCellForDate_(ss, salon, dateIso, monthCache);

      return {
        date: dateIso,
        available: base.status === 'available',
        status: base.status,
        detail: base.code || '',
        source: 'sheet'
      };
    });

    return { salon, cells };
  });

  return ok_({ dates, rows });
}

function saveBlock(token, payload) {
  const session = getSession_(token);
  if (!session) return fail_('Sesión expirada.');

  payload = payload || {};

  const id = cleanText_(payload.id, 80);
  const evento = cleanText_(payload.evento, 160);
  const fechaEntrada = String(payload.fechaEntrada || '').trim();
  const fechaSalida = String(payload.fechaSalida || '').trim();
  const status = normalizeBlockStatus_(payload.status);
  const fechaLimite = String(payload.fechaLimite || '').trim();
  const selections = normalizeSelections_(payload.selections);

  if (!evento) return fail_('Escribe el nombre del evento.');
  if (!parseIsoDate_(fechaEntrada) || !parseIsoDate_(fechaSalida)) {
    return fail_('Selecciona fechas válidas.');
  }
  if (fechaEntrada > fechaSalida) {
    return fail_('La fecha de salida no puede ser anterior a la fecha de entrada.');
  }

  const todayIso = todayIso_();
  if (fechaEntrada < todayIso || fechaSalida < todayIso) {
    return fail_('Las fechas del bloqueo no pueden ser menores al día actual.');
  }

  if (!status) {
    return fail_('El status solo puede ser TENTATIVO o GARANTIZADO.');
  }

  if (status === 'tentative') {
    if (!parseIsoDate_(fechaLimite)) {
      return fail_('Selecciona la fecha límite del bloqueo TENTATIVO.');
    }
    if (fechaLimite < todayIso) {
      return fail_('La fecha límite debe ser igual o mayor al día actual.');
    }
    if (fechaLimite > fechaEntrada) {
      return fail_('La fecha límite debe ser menor o igual a la fecha de entrada.');
    }
  }

  if (!selections.length) {
    return fail_('Selecciona al menos una fecha de algún salón.');
  }

  const allowedDates = isoDateRange_(fechaEntrada, fechaSalida, 62);
  if (!allowedDates) return fail_('El rango máximo para un bloqueo es de 62 días.');

  const allowedDateSet = {};
  allowedDates.forEach(d => allowedDateSet[d] = true);

  const salonSet = {};
  SALONS.forEach(s => salonSet[normalizeName_(s)] = true);

  for (const sel of selections) {
    if (!allowedDateSet[sel.date]) {
      return fail_('Hay una fecha seleccionada fuera del rango de entrada y salida.');
    }
    if (!salonSet[normalizeName_(sel.salon)]) {
      return fail_('Se detectó un salón no válido: ' + sel.salon);
    }
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const sh = ensureBlocksSheet_();
    const now = new Date();

    // Si se modifica, primero se validan las reglas que dependen
    // del estado anterior del bloqueo.
    let existingRow = 0;
    let existingValues = null;

    if (id) {
      const all = sh.getDataRange().getValues();

      for (let i = 1; i < all.length; i++) {
        if (String(all[i][0]) === id) {
          existingRow = i + 1;
          existingValues = all[i];
          break;
        }
      }

      if (!existingRow) return fail_('No se encontró el bloqueo a modificar.');

      // Validación crítica del lado del servidor: ni Administrador ni Operador
      // pueden modificar un bloqueo cuyo propietario sea otro usuario.
      if (String(existingValues[7] || '') !== String(session.id || '')) {
        return fail_('No tienes permiso para modificar este bloqueo. Solo su propietario puede hacerlo.');
      }

      if (String(existingValues[6] || '').toUpperCase() === 'CANCELADO') {
        return fail_('Ese bloqueo ya fue cancelado.');
      }

      const previousStatus = normalizeBlockStatus_(existingValues[4]) || 'tentative';
      const previousExit = formatSheetDateIso_(existingValues[3]);

      // Un garantizado nunca puede regresar a tentativo.
      if (previousStatus === 'guaranteed' && status !== 'guaranteed') {
        return fail_('Un bloqueo GARANTIZADO no puede regresar a TENTATIVO.');
      }

      // Regla: si la salida del garantizado es mañana (o ya quedó más cerca),
      // ya no se permite modificar ningún dato del bloqueo.
      if (previousStatus === 'guaranteed' && previousExit <= tomorrowIso_()) {
        return fail_(
          'Este bloqueo GARANTIZADO ya no puede modificarse porque su fecha de salida es mañana o anterior.'
        );
      }
    }

    const conflicts = findBlockConflicts_(selections, id);

    if (conflicts.length) {
      const first = conflicts
        .slice(0, 5)
        .map(c => c.salon + ' · ' + c.date)
        .join(', ');

      return fail_(
        'No se pudo guardar. Hay fechas ya ocupadas: ' +
        first +
        (conflicts.length > 5 ? '…' : '')
      );
    }

    const jsonSel = JSON.stringify(selections);
    const deadlineToStore = status === 'tentative' ? fechaLimite : '';

    if (id) {
      // B:K conserva la estructura original.
      sh.getRange(existingRow, 2, 1, 10).setValues([[
        evento,
        fechaEntrada,
        fechaSalida,
        status,
        jsonSel,
        'ACTIVO',
        session.id,
        session.nombre,
        existingValues[9] || now,
        now
      ]]);

      // L:N - nuevas columnas.
      sh.getRange(existingRow, 12, 1, 3).setValues([[
        deadlineToStore,
        '',
        ''
      ]]);

      rescheduleTentativeDeadlineTrigger_();

      return ok_({
        message: 'Bloqueo modificado correctamente.',
        id
      });
    }

    const newId = 'BLQ-' + Utilities.getUuid();

    sh.appendRow([
      newId,
      evento,
      fechaEntrada,
      fechaSalida,
      status,
      jsonSel,
      'ACTIVO',
      session.id,
      session.nombre,
      now,
      now,
      deadlineToStore,
      '',
      ''
    ]);

    rescheduleTentativeDeadlineTrigger_();

    return ok_({
      message: 'Nuevo bloqueo agregado correctamente.',
      id: newId
    });

  } finally {
    lock.releaseLock();
  }
}

function cancelBlock(token, blockId) {
  const session = getSession_(token);
  if (!session) return fail_('Sesión expirada.');

  blockId = String(blockId || '').trim();
  if (!blockId) return fail_('Selecciona un bloqueo.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const sh = ensureBlocksSheet_();
    const values = sh.getDataRange().getValues();

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) !== blockId) continue;

      // Validación crítica del lado del servidor: únicamente el propietario
      // del bloqueo puede cancelarlo, sin excepción por tipo de usuario.
      if (String(values[i][7] || '') !== String(session.id || '')) {
        return fail_('No tienes permiso para cancelar este bloqueo. Solo su propietario puede hacerlo.');
      }

      if (String(values[i][6] || '').toUpperCase() === 'CANCELADO') {
        return fail_('El bloqueo ya estaba cancelado.');
      }

      const now = new Date();
      const note = 'Bloqueo cancelado por usuario';

      archiveCancelledBlockRow_(values[i], session.nombre, note, now);

      sh.getRange(i + 1, 7).setValue('CANCELADO');
      sh.getRange(i + 1, 9).setValue(session.nombre);
      sh.getRange(i + 1, 11).setValue(now);
      sh.getRange(i + 1, 13, 1, 2).setValues([[note, now]]);

      rescheduleTentativeDeadlineTrigger_();

      return ok_({
        message: 'Bloqueo cancelado correctamente y enviado al histórico.'
      });
    }

    return fail_('No se encontró el bloqueo.');

  } finally {
    lock.releaseLock();
  }
}


function requireAdminSession_(token) {
  const session = getSession_(token);
  if (!session) return { ok: false, error: 'Sesión expirada.' };
  if (String(session.tipo || '').toLowerCase() !== 'administrador') {
    return { ok: false, error: 'Esta opción es exclusiva para usuarios Administrador.' };
  }
  return { ok: true, session };
}

function getOutOfServiceBlocks(token) {
  const auth = requireAdminSession_(token);
  if (!auth.ok) return fail_(auth.error);

  const blocks = getAllBlocks_()
    .filter(b => b.estado !== 'CANCELADO' && b.status === 'out_of_service')
    .sort((a,b) => String(a.fechaEntrada||'').localeCompare(String(b.fechaEntrada||'')) || String(a.evento||'').localeCompare(String(b.evento||''),'es'));

  return ok_({ blocks });
}

function getOutOfServiceGrid(token, fechaEntrada, fechaSalida, excludeBlockId) {
  const auth = requireAdminSession_(token);
  if (!auth.ok) return fail_(auth.error);

  const start = parseIsoDate_(fechaEntrada);
  const end = parseIsoDate_(fechaSalida);
  if (!start || !end) return fail_('Selecciona fecha de entrada y fecha de salida.');
  if (start.iso > end.iso) return fail_('La fecha de salida no puede ser anterior a la fecha de entrada.');

  const dates = isoDateRange_(start.iso, end.iso, 62);
  if (!dates) return fail_('El rango máximo para un bloqueo es de 62 días.');

  const ss = getOperationalSpreadsheet_();
  const monthCache = {};
  const blockMap = {};

  getActiveBlocks_()
    .filter(b => String(b.id) !== String(excludeBlockId || ''))
    .forEach(b => {
      (b.selections || []).forEach(sel => {
        blockMap[normalizeName_(sel.salon) + '|' + sel.date] = {
          status: b.status,
          evento: b.evento,
          blockId: b.id
        };
      });
    });

  const rows = SALONS.map(salon => {
    const cells = dates.map(dateIso => {
      const key = normalizeName_(salon) + '|' + dateIso;
      if (blockMap[key]) {
        return { date: dateIso, available:false, status:blockMap[key].status, detail:blockMap[key].evento, source:'block' };
      }
      const base = getBaseCellForDate_(ss, salon, dateIso, monthCache);
      return { date:dateIso, available:base.status==='available', status:base.status, detail:base.code||'', source:'sheet' };
    });
    return { salon, cells };
  });

  return ok_({ dates, rows });
}

function saveOutOfServiceBlock(token, payload) {
  const auth = requireAdminSession_(token);
  if (!auth.ok) return fail_(auth.error);
  const session = auth.session;
  payload = payload || {};

  const id = cleanText_(payload.id, 80);
  const evento = cleanText_(payload.evento, 160);
  const fechaEntrada = String(payload.fechaEntrada || '').trim();
  const fechaSalida = String(payload.fechaSalida || '').trim();
  const selections = normalizeSelections_(payload.selections);

  if (!evento) return fail_('Escribe el nombre o motivo del fuera de servicio.');
  if (!parseIsoDate_(fechaEntrada) || !parseIsoDate_(fechaSalida)) return fail_('Selecciona fechas válidas.');
  if (fechaEntrada > fechaSalida) return fail_('La fecha de salida no puede ser anterior a la fecha de entrada.');
  if (!selections.length) return fail_('Selecciona al menos una fecha de algún salón.');

  const allowedDates = isoDateRange_(fechaEntrada, fechaSalida, 62);
  if (!allowedDates) return fail_('El rango máximo para un fuera de servicio es de 62 días.');
  const allowedDateSet = {}; allowedDates.forEach(d => allowedDateSet[d] = true);
  const salonSet = {}; SALONS.forEach(s => salonSet[normalizeName_(s)] = true);
  for (const sel of selections) {
    if (!allowedDateSet[sel.date]) return fail_('Hay una fecha seleccionada fuera del rango de entrada y salida.');
    if (!salonSet[normalizeName_(sel.salon)]) return fail_('Se detectó un salón no válido: ' + sel.salon);
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ensureBlocksSheet_();
    const now = new Date();
    let existingRow = 0;
    let existingValues = null;

    if (id) {
      const all = sh.getDataRange().getValues();
      for (let i=1;i<all.length;i++) {
        if (String(all[i][0]) === id) { existingRow=i+1; existingValues=all[i]; break; }
      }
      if (!existingRow) return fail_('No se encontró el fuera de servicio a modificar.');
      if (String(existingValues[6]||'').toUpperCase()==='CANCELADO') return fail_('Ese fuera de servicio ya fue cancelado.');
      if (normalizeBlockStatus_(existingValues[4]) !== 'out_of_service') return fail_('El registro seleccionado no es FUERA DE SERVICIO.');
    } else {
      const today = todayIso_();
      if (fechaEntrada < today || fechaSalida < today) return fail_('Para un nuevo fuera de servicio las fechas no pueden ser menores al día actual.');
    }

    const conflicts = findBlockConflicts_(selections, id);
    if (conflicts.length) {
      const first=conflicts.slice(0,5).map(c=>c.salon+' · '+c.date).join(', ');
      return fail_('No se pudo guardar. Hay fechas ya ocupadas: '+first+(conflicts.length>5?'…':''));
    }

    const jsonSel = JSON.stringify(selections);
    if (id) {
      sh.getRange(existingRow,2,1,10).setValues([[
        evento, fechaEntrada, fechaSalida, 'out_of_service', jsonSel, 'ACTIVO',
        existingValues[7] || session.id, existingValues[8] || session.nombre,
        existingValues[9] || now, now
      ]]);
      sh.getRange(existingRow,12,1,3).setValues([['','','']]);
      return ok_({message:'Fuera de servicio modificado correctamente.', id});
    }

    const newId='FDS-'+Utilities.getUuid();
    sh.appendRow([newId,evento,fechaEntrada,fechaSalida,'out_of_service',jsonSel,'ACTIVO',session.id,session.nombre,now,now,'','','']);
    return ok_({message:'Fuera de servicio agregado correctamente.', id:newId});
  } finally { lock.releaseLock(); }
}

function cancelOutOfServiceBlock(token, blockId) {
  const auth = requireAdminSession_(token);
  if (!auth.ok) return fail_(auth.error);
  const session=auth.session;
  blockId=String(blockId||'').trim();
  if(!blockId) return fail_('Selecciona un fuera de servicio.');

  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh=ensureBlocksSheet_(); const values=sh.getDataRange().getValues();
    for(let i=1;i<values.length;i++){
      if(String(values[i][0])!==blockId) continue;
      if(normalizeBlockStatus_(values[i][4])!=='out_of_service') return fail_('El registro seleccionado no es FUERA DE SERVICIO.');
      if(String(values[i][6]||'').toUpperCase()==='CANCELADO') return fail_('El fuera de servicio ya estaba cancelado.');
      const now=new Date(); const note='Fuera de servicio cancelado por administrador';
      archiveCancelledBlockRow_(values[i],session.nombre,note,now);
      sh.getRange(i+1,7).setValue('CANCELADO');
      sh.getRange(i+1,9).setValue(session.nombre);
      sh.getRange(i+1,11).setValue(now);
      sh.getRange(i+1,13,1,2).setValues([[note,now]]);
      return ok_({message:'Fuera de servicio cancelado correctamente.'});
    }
    return fail_('No se encontró el fuera de servicio.');
  } finally { lock.releaseLock(); }
}

function getOutOfServiceCancelledReport(token, year, month) {
  const auth=requireAdminSession_(token);
  if(!auth.ok) return fail_(auth.error);
  year=Number(year); month=Number(month);
  if(!Number.isInteger(year)||!Number.isInteger(month)||month<1||month>12) return fail_('Mes o año no válido.');
  const sh=ensureBlocksHistorySheet_(); const values=sh.getDataRange().getValues(); const rows=[];
  for(let i=1;i<values.length;i++){
    if(!values[i][0]) continue;
    if(normalizeBlockStatus_(values[i][4])!=='out_of_service') continue;
    const fechaEntrada=formatSheetDateIso_(values[i][2]); const parsed=parseIsoDate_(fechaEntrada);
    if(!parsed||parsed.year!==year||parsed.month!==month) continue;
    let selections=[]; try{selections=JSON.parse(String(values[i][6]||'[]'));}catch(e){}
    rows.push({
      id:String(values[i][0]||''), evento:String(values[i][1]||''), fechaEntrada,
      fechaSalida:formatSheetDateIso_(values[i][3]), status:'out_of_service', fechaLimite:'',
      selections:normalizeSelections_(selections), usuarioCreacion:String(values[i][7]||''),
      canceladoPor:String(values[i][10]||''), motivo:String(values[i][11]||''), fechaHoraCancelacion:dateTimeText_(values[i][12])
    });
  }
  return ok_({rows});
}

function getAllBlocks_() {
  const sh = ensureBlocksSheet_();
  const values = sh.getDataRange().getValues();
  const out = [];
  const tomorrow = tomorrowIso_();

  for (let i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;

    let selections = [];
    try {
      selections = JSON.parse(String(values[i][5] || '[]'));
    } catch (e) {}

    const status = normalizeBlockStatus_(values[i][4]) || 'tentative';
    const fechaSalida = formatSheetDateIso_(values[i][3]);
    const fechaLimite = formatSheetDateIso_(values[i][11]);

    out.push({
      id: String(values[i][0] || ''),
      evento: String(values[i][1] || ''),
      fechaEntrada: formatSheetDateIso_(values[i][2]),
      fechaSalida,
      status,
      fechaLimite: status === 'tentative' ? fechaLimite : '',
      selections: normalizeSelections_(selections),
      estado: String(values[i][6] || 'ACTIVO').toUpperCase(),
      usuarioId: String(values[i][7] || ''),
      usuarioNombre: String(values[i][8] || ''),
      creado: dateTimeText_(values[i][9]),
      actualizado: dateTimeText_(values[i][10]),
      motivoCancelacion: String(values[i][12] || ''),
      fechaCancelacion: dateTimeText_(values[i][13]),
      modificationLocked: status === 'guaranteed' && !!fechaSalida && fechaSalida <= tomorrow
    });
  }

  return out;
}

function getActiveBlocks_() {
  return getAllBlocks_().filter(b => b.estado !== 'CANCELADO');
}

function archiveCancelledBlockRow_(row, cancelledBy, reason, when) {
  const history = ensureBlocksHistorySheet_();

  history.appendRow([
    String(row[0] || ''),
    String(row[1] || ''),
    formatSheetDateIso_(row[2]),
    formatSheetDateIso_(row[3]),
    normalizeBlockStatus_(row[4]) || String(row[4] || ''),
    formatSheetDateIso_(row[11]),
    String(row[5] || '[]'),
    String(row[8] || ''),
    row[9] || '',
    row[10] || '',
    cancelledBy || 'SISTEMA',
    reason || '',
    when || new Date(),
    String(row[7] || '')
  ]);
}

function processExpiredTentatives_(rescheduleAfter) {
  const sh = ensureBlocksSheet_();
  const values = sh.getDataRange().getValues();
  const now = new Date();
  const today = todayIso_();
  const currentTime = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm');

  // Un bloqueo vence a las 23:59 del día indicado.
  // Si por cualquier razón el trigger se ejecuta un poco después,
  // también se procesa al siguiente acceso para no dejar bloqueos vencidos activos.
  const expiredRows = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[0]) continue;
    if (String(row[6] || '').toUpperCase() === 'CANCELADO') continue;

    const status = normalizeBlockStatus_(row[4]);
    const deadline = formatSheetDateIso_(row[11]);

    if (status !== 'tentative' || !deadline) continue;

    const expired =
      deadline < today ||
      (deadline === today && currentTime >= '23:59');

    if (expired) expiredRows.push(i + 1);
  }

  expiredRows.forEach(rowNumber => {
    const row = sh.getRange(rowNumber, 1, 1, 14).getValues()[0];

    // Verificación de nuevo por seguridad ante ejecuciones concurrentes.
    if (String(row[6] || '').toUpperCase() === 'CANCELADO') return;

    const deadline = formatSheetDateIso_(row[11]);
    const stamp = Utilities.formatDate(
      now,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd HH:mm:ss'
    );

    const note =
      'Bloqueo cancelado por fecha limite. Fecha y hora: ' + stamp;

    archiveCancelledBlockRow_(row, 'SISTEMA', note, now);

    sh.getRange(rowNumber, 7).setValue('CANCELADO');
    sh.getRange(rowNumber, 9).setValue('SISTEMA');
    sh.getRange(rowNumber, 11).setValue(now);
    sh.getRange(rowNumber, 13, 1, 2).setValues([[note, now]]);
  });

  if (rescheduleAfter !== false) {
    rescheduleTentativeDeadlineTrigger_();
  }

  return expiredRows.length;
}

// Función pública para el trigger instalable.
function expireTentativeBlocks() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;

  try {
    processExpiredTentatives_(true);
  } finally {
    lock.releaseLock();
  }
}

function rescheduleTentativeDeadlineTrigger_() {
  // Mantiene un solo trigger pendiente para la fecha límite más cercana.
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'expireTentativeBlocks') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  const blocks = getAllBlocks_()
    .filter(b =>
      b.estado !== 'CANCELADO' &&
      b.status === 'tentative' &&
      b.fechaLimite
    )
    .sort((a, b) => a.fechaLimite.localeCompare(b.fechaLimite));

  if (!blocks.length) return;

  const deadline = parseIsoDate_(blocks[0].fechaLimite);
  if (!deadline) return;

  const runAt = new Date(
    deadline.year,
    deadline.month - 1,
    deadline.day,
    23,
    59,
    0,
    0
  );

  // Si el horario ya pasó, se ejecuta en breve.
  const target = runAt.getTime() <= Date.now()
    ? new Date(Date.now() + 60 * 1000)
    : runAt;

  ScriptApp.newTrigger('expireTentativeBlocks')
    .timeBased()
    .at(target)
    .create();
}

function tomorrowIso_() {
  const d = new Date();
  d.setDate(d.getDate() + 1);

  return Utilities.formatDate(
    d,
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
}

function normalizeBlockStatus_(status) {
  const s = normalizeName_(status).toLowerCase();

  // Compatibilidad con bloqueos creados con la versión anterior:
  // "ocupado" se interpreta como TENTATIVO.
  if (['ocupado', 'occupied', 'tentativo', 'tentative'].includes(s)) return 'tentative';
  if (['garantizado', 'guaranteed', 'definitivo', 'definitive'].includes(s)) return 'guaranteed';
  if (['fuera de servicio', 'fuera_servicio', 'out of service', 'out_of_service', 'service'].includes(s)) return 'out_of_service';

  return '';
}

function normalizeSelections_(items) {
  if (!Array.isArray(items)) return [];

  const seen = {};
  const out = [];

  items.forEach(item => {
    const salon = cleanText_(item && item.salon, 120);
    const date = String((item && item.date) || '').trim();

    if (!salon || !parseIsoDate_(date)) return;

    const key = normalizeName_(salon) + '|' + date;
    if (seen[key]) return;

    seen[key] = true;
    out.push({ salon, date });
  });

  return out;
}

function findBlockConflicts_(selections, excludeBlockId) {
  const conflicts = [];
  const blockMap = {};

  getActiveBlocks_()
    .filter(b => String(b.id) !== String(excludeBlockId || ''))
    .forEach(b => {
      (b.selections || []).forEach(sel => {
        blockMap[normalizeName_(sel.salon) + '|' + sel.date] = true;
      });
    });

  const ss = getOperationalSpreadsheet_();
  const monthCache = {};

  selections.forEach(sel => {
    const key = normalizeName_(sel.salon) + '|' + sel.date;

    if (blockMap[key]) {
      conflicts.push(sel);
      return;
    }

    const base = getBaseCellForDate_(ss, sel.salon, sel.date, monthCache);

    if (base.status !== 'available') {
      conflicts.push(sel);
    }
  });

  return conflicts;
}

function getBaseCellForDate_(ss, salonName, dateIso, cache) {
  const p = parseIsoDate_(dateIso);
  if (!p) return { status: 'available', code: '' };

  const key = p.year + '-' + p.month;

  if (!cache[key]) {
    cache[key] = getBaseMonthMap_(ss, p.year, p.month);
  }

  const monthMap = cache[key];
  const salon = monthMap[normalizeName_(salonName)];

  if (!salon || !salon[p.day - 1]) {
    return { status: 'available', code: '' };
  }

  return salon[p.day - 1];
}

function getBaseMonthMap_(ss, year, month) {
  const out = {};
  const sh = findMonthSheet_(ss, CONFIG.MONTHS_ES[month - 1], year);

  if (!sh) return out;

  const grid = sh.getDataRange();
  const values = grid.getDisplayValues();
  const backgrounds = grid.getBackgrounds();
  const headerInfo = detectDayHeader_(values, year, month);

  if (!headerInfo) return out;

  const daysInMonth = new Date(year, month, 0).getDate();
  const stopNames = ['DISPONIBLES','DEFINITIVOS','TENTATIVOS'];

  for (let r = headerInfo.headerRow + 1; r < values.length; r++) {
    const salon = String(values[r][0] || '').trim();
    if (!salon) continue;

    const upper = normalizeName_(salon);

    if (stopNames.some(s => upper.includes(s))) break;

    const cells = [];

    for (let d = 1; d <= daysInMonth; d++) {
      const c = headerInfo.firstDayCol + d - 1;

      const code = c < values[r].length
        ? String(values[r][c] || '').trim()
        : '';

      const bg = c < backgrounds[r].length
        ? normalizeHex_(backgrounds[r][c] || '#ffffff')
        : '#ffffff';

      cells.push({
        status: classifyStatus_(bg, code),
        code
      });
    }

    out[upper] = cells;
  }

  return out;
}

function parseIsoDate_(iso) {
  const m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;

  const year = Number(m[1]);
  const month = Number(m[2]);
  const day = Number(m[3]);

  const dt = new Date(year, month - 1, day);

  if (
    dt.getFullYear() !== year ||
    dt.getMonth() !== month - 1 ||
    dt.getDate() !== day
  ) {
    return null;
  }

  return {
    year,
    month,
    day,
    iso: m[1] + '-' + m[2] + '-' + m[3]
  };
}

function isoDateRange_(startIso, endIso, maxDays) {
  const s = parseIsoDate_(startIso);
  const e = parseIsoDate_(endIso);

  if (!s || !e || s.iso > e.iso) return null;

  const out = [];
  let dt = new Date(s.year, s.month - 1, s.day);
  const end = new Date(e.year, e.month - 1, e.day);

  while (dt <= end) {
    if (maxDays && out.length >= maxDays) return null;

    out.push(
      Utilities.formatDate(
        dt,
        Session.getScriptTimeZone(),
        'yyyy-MM-dd'
      )
    );

    dt.setDate(dt.getDate() + 1);
  }

  return out;
}

function todayIso_() {
  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
}

function formatSheetDateIso_(value) {
  if (value instanceof Date && !isNaN(value)) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );
  }

  const s = String(value || '').trim();
  const p = parseIsoDate_(s);

  return p ? p.iso : s;
}

function dateTimeText_(value) {
  if (value instanceof Date && !isNaN(value)) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd HH:mm'
    );
  }

  return String(value || '');
}

function blockCode_(evento) {
  const words = normalizeName_(evento)
    .split(/\s+/)
    .filter(Boolean);

  if (!words.length) return 'BL';
  if (words.length === 1) return words[0].slice(0, 3);

  return words
    .slice(0, 3)
    .map(w => w.charAt(0))
    .join('');
}


/* =========================
   USERS / PASSWORD HELPERS
   ========================= */

function getUsersData_(sh) {
  const values = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < values.length; i++) {
    if (!values[i][0] && !values[i][2]) continue;
    out.push({
      row: i + 1,
      id: String(values[i][0] || ''),
      nombre: String(values[i][1] || ''),
      email: normalizeEmail_(values[i][2]),
      passwordRecord: String(values[i][3] || ''),
      tipo: normalizeUserType_(values[i][4]) || 'administrador',
      estado: String(values[i][5] || 'ACTIVO').trim().toUpperCase() || 'ACTIVO'
    });
  }
  return out;
}

function makePasswordRecord_(password) {
  const salt = Utilities.getUuid().replace(/-/g, '') + createSecureToken_().slice(0, 24);
  const hash = sha256Hex_(salt + '|' + password);
  return salt + ':' + hash;
}

function verifyPassword_(password, record) {
  const parts = String(record || '').split(':');
  if (parts.length !== 2) return false;
  const salt = parts[0];
  const expected = parts[1];
  return timingSafeEqual_(sha256Hex_(salt + '|' + password), expected);
}

function validatePassword_(password) {
  if (password.length < 8) return 'La contraseña debe tener al menos 8 caracteres.';
  if (!/[A-ZÁÉÍÓÚÑ]/i.test(password) || !/\d/.test(password)) return 'La contraseña debe incluir letras y al menos un número.';
  return '';
}

function sha256Hex_(text) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text), Utilities.Charset.UTF_8);
  return bytes.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

function timingSafeEqual_(a, b) {
  a = String(a); b = String(b);
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

/* =========================
   SESSION HELPERS
   ========================= */

function createSession_(id, nombre, email, tipo) {
  const token = createSecureToken_();
  const key = 'SESSION_' + sha256Hex_(token);
  const expiresAt = Date.now() + CONFIG.SESSION_DAYS * 24 * 60 * 60 * 1000;
  PropertiesService.getScriptProperties().setProperty(key, JSON.stringify({ id, nombre, email, tipo: normalizeUserType_(tipo) || 'operador', expiresAt }));
  return token;
}

function getSession_(token) {
  token = String(token || '');
  if (!token) return null;
  const key = 'SESSION_' + sha256Hex_(token);
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(key);
  if (!raw) return null;
  let session;
  try { session = JSON.parse(raw); } catch (e) { return null; }
  if (!session || Date.now() > Number(session.expiresAt || 0)) {
    props.deleteProperty(key);
    return null;
  }
  return session;
}

function deleteSession_(token) {
  token = String(token || '');
  if (!token) return;
  PropertiesService.getScriptProperties().deleteProperty('SESSION_' + sha256Hex_(token));
}

function revokeAllSessionsForUser_(userId) {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  Object.keys(all).forEach(k => {
    if (!k.startsWith('SESSION_')) return;
    try {
      const s = JSON.parse(all[k]);
      if (String(s.id) === String(userId)) props.deleteProperty(k);
    } catch (e) {}
  });
}

function createSecureToken_() {
  return Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      Utilities.getUuid() + '|' + new Date().getTime() + '|' + Math.random()
    )
  ).replace(/=+$/g, '');
}

/* =========================
   GENERIC HELPERS
   ========================= */

function normalizeEmail_(email) {
  return String(email || '').trim().toLowerCase();
}

function cleanText_(text, maxLen) {
  return String(text || '').trim().replace(/[<>]/g, '').slice(0, maxLen || 200);
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function normalizeName_(s) {
  return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase().replace(/\s+/g, ' ').trim();
}

function escapeHtml_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#039;');
}

function ok_(data) {
  return Object.assign({ ok: true }, data || {});
}

function fail_(message) {
  return { ok: false, message: String(message || 'Ocurrió un error.') };
}
