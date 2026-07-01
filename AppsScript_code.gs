/**
 * Backend de Reservas de Salas - FACE Financial Advisors
 * --------------------------------------------------------
 * GESTIÓN DE ASESORES (sin tocar el código):
 * - La pestaña "Asesores" se crea automáticamente con los 40 asesores iniciales.
 * - Desde ahí puedes: agregar/quitar asesores, activarlos/desactivarlos, definir admins
 *   y ajustar el límite de horas semanales por persona.
 * - Columnas de "Asesores": nombre | activo | esAdmin | limiteHorasSemanales
 *
 * Cada cambio a este código requiere una NUEVA versión en "Administrar implementaciones".
 */

const SHEET_NAME = "Reservas";
const ADVISORS_SHEET_NAME = "Asesores";

const DEFAULT_ADVISORS = [
  "Alejandro Argüelles Gonzalez","Alejandro Martinez Romero","Alejandro Sánchez Barbara",
  "Alejandro Tapia Castillo","Alberto Baigts Cadeño","Alberto Cortes Bolio",
  "Amairani Pamela Merino Rios","Andres Hoyos Morales","Arturo Rodríguez Muraira",
  "Carlos Alberto Duarte Morales","Eduardo Giralt Dominguez","Elisa Islas Diaz de León",
  "Erick Omar Villegas Jiménez","Ferdinando Santiago Hoyos","Germán González Cantón",
  "Hector Joaquin Lopez Osorio","Jonathan Alexis Bojorquez Ruelas","Jorge García Granados",
  "Jorge Roberto Rivera Medrano","Jose Abraham Grafias","Jose Pablo Flores Rabasa",
  "Juan Carlos Alonso","Juan Carlos Zacarias","Julian Cardona",
  "Laura Quintero","Lawrence Eden Anaya","Luis Guillermo Garcia Dueñas Cobarruvias",
  "Marco Antonio Olivas Betancourt","Maria Clara Cuenca De La Concha",
  "Maria de Lourdes Koloffon Lara","Maria de Lourdes Pastor Farril",
  "Maria del Rocio Alonso Martinez","Maria Rosalia Jimenez Ortega",
  "Mario Alejandro Gongora Rivera","Mauricio Garza Castro","Moises Arres Muñiz",
  "Pablo Noriega Bello","Paula Yolanda Diaz Herrera","Ramón Alejandro González Vega",
  "Sofia Mejorada Perez Verdia"
];
const DEFAULT_ADMINS = ["Ferdinando Santiago Hoyos", "Pablo Noriega Bello", "Germán González Cantón"];
const DEFAULT_WEEKLY_LIMIT = 999;

// ─── Helpers de hojas ────────────────────────────────────────────────────────

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["id", "room", "date", "time", "duration", "advisor", "createdAt"]);
  }
  return sheet;
}

function getAdvisorsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ADVISORS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(ADVISORS_SHEET_NAME);
    sheet.appendRow(["nombre", "activo", "esAdmin", "limiteHorasSemanales"]);
    DEFAULT_ADVISORS.forEach(name => {
      const admin = DEFAULT_ADMINS.indexOf(name) !== -1;
      sheet.appendRow([name, true, admin, DEFAULT_WEEKLY_LIMIT]);
    });
    sheet.getRange(1, 1, 1, 4).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function readAdvisors() {
  const sheet = getAdvisorsSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const idx = {
    nombre:     headers.indexOf("nombre"),
    activo:     headers.indexOf("activo"),
    esAdmin:    headers.indexOf("esAdmin"),
    limiteHoras: headers.indexOf("limiteHorasSemanales")
  };
  return data
    .filter(row => String(row[idx.nombre]).trim() !== "")
    .map(row => ({
      nombre:      String(row[idx.nombre]).trim(),
      activo:      row[idx.activo] === true  || String(row[idx.activo]).toLowerCase()  === "true",
      esAdmin:     row[idx.esAdmin] === true || String(row[idx.esAdmin]).toLowerCase() === "true",
      limiteHoras: Number(row[idx.limiteHoras]) > 0 ? Number(row[idx.limiteHoras]) : DEFAULT_WEEKLY_LIMIT
    }));
}

/** Quita acentos y pasa a minúsculas para comparar nombres sin importar tildes. */
function normalizeName(s) {
  return String(s).trim().toLowerCase()
    .replace(/[áàäâ]/g,'a').replace(/[éèëê]/g,'e')
    .replace(/[íìïî]/g,'i').replace(/[óòöô]/g,'o')
    .replace(/[úùüû]/g,'u').replace(/ñ/g,'n');
}

function findAdvisor(name, advisors) {
  const n = normalizeName(name);
  return advisors.find(a => normalizeName(a.nombre) === n) || null;
}

function isAdmin(name, advisors) {
  const rec = findAdvisor(name, advisors);
  return rec ? rec.esAdmin : false;
}

// ─── Endpoints ───────────────────────────────────────────────────────────────

function doGet(e) {
  const action = e.parameter.action;

  if (action === "advisors") {
    return jsonResponse(readAdvisors());
  }

  const date = e.parameter.date;
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const idx = buildIdx(headers, ["id","room","date","time","duration","advisor"]);

  const results = data
    .filter(row => row[idx.id] !== "" && (!date || normalizeDate(row[idx.date]) === date))
    .map(row => ({
      id:       row[idx.id],
      room:     row[idx.room],
      date:     normalizeDate(row[idx.date]),
      time:     normalizeTime(row[idx.time]),
      duration: Number(row[idx.duration]) || 1,
      advisor:  row[idx.advisor]
    }));

  return jsonResponse(results);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const sheet = getSheet();
  if (body.action === "create") return handleCreate(sheet, body);
  if (body.action === "edit")   return handleEdit(sheet, body);
  if (body.action === "cancel") return handleCancel(sheet, body);
  return jsonResponse({ ok: false, message: "Acción no reconocida." });
}

// ─── Lógica de reservas ──────────────────────────────────────────────────────

function handleCreate(sheet, body) {
  const { room, date, time, advisor } = body;
  const duration = Math.max(1, parseInt(body.duration || 1, 10));
  if (!room || !date || !time || !advisor) {
    return jsonResponse({ ok: false, message: "Faltan datos para la reserva." });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const advisors = readAdvisors();

    // Verificar que el asesor exista y esté activo
    const rec = findAdvisor(advisor, advisors);
    if (!rec) {
      return jsonResponse({ ok: false, message: "Asesor no encontrado en el directorio. Contacta a un administrador." });
    }
    if (!rec.activo) {
      return jsonResponse({ ok: false, message: "Tu cuenta está desactivada. Contacta a un administrador." });
    }

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const idx = buildIdx(headers, ["room","date","time","duration","advisor"]);

    // Verificar conflicto de horario
    const newStart = startHour(time);
    const conflict = data.some(row => {
      if (row[idx.room] !== room || normalizeDate(row[idx.date]) !== date) return false;
      const s = startHour(row[idx.time]);
      const sp = Math.max(1, Number(row[idx.duration]) || 1);
      return rangesOverlap(newStart, newStart + duration, s, s + sp);
    });
    if (conflict) {
      return jsonResponse({ ok: false, message: "Ese horario se cruza con otra reserva ya existente." });
    }

    // Verificar límite semanal de horas
    const wStart = weekStart(date);
    const wEnd   = weekEnd(wStart);
    const usedHours = data.reduce((sum, row) => {
      const rd = normalizeDate(row[idx.date]);
      if (normalizeName(row[idx.advisor]) !== normalizeName(advisor) || rd < wStart || rd > wEnd) return sum;
      return sum + Math.max(1, Number(row[idx.duration]) || 1);
    }, 0);

    if (usedHours + duration > rec.limiteHoras) {
      return jsonResponse({
        ok: false,
        message: `Has alcanzado tu límite semanal de ${rec.limiteHoras} hora(s). Ya tienes ${usedHours}h reservadas esta semana. Contacta al área directiva para que te asignen más horas.`
      });
    }

    // Guardar reserva
    const id = Utilities.getUuid();
    const newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, idx.time + 1).setNumberFormat("@");
    sheet.getRange(newRow, 1, 1, 7).setValues([[id, room, date, time, duration, advisor, new Date()]]);
    return jsonResponse({ ok: true, id });
  } finally {
    lock.releaseLock();
  }
}

function handleEdit(sheet, body) {
  const { id, advisor } = body;
  const newDuration = Math.max(1, parseInt(body.duration || 1, 10));
  if (!id) return jsonResponse({ ok: false, message: "Falta el id de la reserva." });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const advisors = readAdvisors();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idx = buildIdx(headers, ["id","room","date","time","duration","advisor"]);

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][idx.id] === id) { rowIndex = i; break; }
    }
    if (rowIndex === -1) return jsonResponse({ ok: false, message: "No se encontró la reserva." });

    const isOwner = data[rowIndex][idx.advisor] === advisor;
    const admin   = isAdmin(advisor, advisors);
    if (!isOwner && !admin) {
      return jsonResponse({ ok: false, message: "Solo puedes editar tus propias reservas." });
    }

    // Bloquear edición de reservas pasadas (nadie puede modificar, ni admins)
    const bookingDate = normalizeDate(data[rowIndex][idx.date]);
    const bookingHour = startHour(data[rowIndex][idx.time]);
    const bookingDateTime = new Date(bookingDate + "T" + String(bookingHour).padStart(2, "0") + ":00:00");
    if (bookingDateTime < new Date()) {
      return jsonResponse({ ok: false, message: "No se puede editar una reserva que ya ocurrió." });
    }

    const room  = data[rowIndex][idx.room];
    const date  = normalizeDate(data[rowIndex][idx.date]);
    const start = startHour(data[rowIndex][idx.time]);

    // Conflicto de horario excluyendo la fila actual
    const conflict = data.some((row, i) => {
      if (i === rowIndex || i === 0) return false;
      if (row[idx.room] !== room || normalizeDate(row[idx.date]) !== date) return false;
      const s  = startHour(row[idx.time]);
      const sp = Math.max(1, Number(row[idx.duration]) || 1);
      return rangesOverlap(start, start + newDuration, s, s + sp);
    });
    if (conflict) {
      return jsonResponse({ ok: false, message: "La nueva duración se cruza con otra reserva." });
    }

    // Límite semanal (excluyendo la reserva que se edita)
    const bookingAdvisor = data[rowIndex][idx.advisor];
    const rec = findAdvisor(bookingAdvisor, advisors);
    if (rec) {
      const wStart = weekStart(date);
      const wEnd   = weekEnd(wStart);
      const usedHours = data.reduce((sum, row, i) => {
        if (i === 0 || i === rowIndex) return sum;
        const rd = normalizeDate(row[idx.date]);
        if (normalizeName(row[idx.advisor]) !== normalizeName(bookingAdvisor) || rd < wStart || rd > wEnd) return sum;
        return sum + Math.max(1, Number(row[idx.duration]) || 1);
      }, 0);
      if (usedHours + newDuration > rec.limiteHoras) {
        return jsonResponse({
          ok: false,
          message: `La nueva duración superaría tu límite semanal de ${rec.limiteHoras}h. Contacta al área directiva para solicitar más horas.`
        });
      }
    }

    sheet.getRange(rowIndex + 1, idx.duration + 1).setValue(newDuration);
    return jsonResponse({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function handleCancel(sheet, body) {
  const { id, advisor } = body;
  if (!id) return jsonResponse({ ok: false, message: "Falta el id de la reserva." });

  const advisors = readAdvisors();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idx = buildIdx(headers, ["id","room","date","time","duration","advisor"]);

  for (let i = 1; i < data.length; i++) {
    if (data[i][idx.id] !== id) continue;

    const isOwner = data[i][idx.advisor] === advisor;
    const admin   = isAdmin(advisor, advisors);
    if (!isOwner && !admin) {
      return jsonResponse({ ok: false, message: "Solo puedes cancelar tus propias reservas." });
    }

    // Bloqueo de cancelación con menos de 2 horas de anticipación (solo para no-admins)
    if (!admin) {
      const bookingDate = normalizeDate(data[i][idx.date]);
      const bookingHour = startHour(data[i][idx.time]);
      const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
      const nowStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH");
      const bookingStr = bookingDate + " " + String(bookingHour).padStart(2, "0");

      // Comparar diferencia en horas
      const nowDate = new Date();
      const bookingDateTime = new Date(bookingDate + "T" + String(bookingHour).padStart(2, "0") + ":00:00");
      const diffMs = bookingDateTime.getTime() - nowDate.getTime();
      const diffHours = diffMs / (1000 * 60 * 60);

      if (diffHours < 2) {
        return jsonResponse({
          ok: false,
          message: "No puedes cancelar una reserva con menos de 2 horas de anticipación. Las horas siguen contando para tu semana."
        });
      }
    }

    sheet.deleteRow(i + 1);
    return jsonResponse({ ok: true });
  }
  return jsonResponse({ ok: false, message: "No se encontró la reserva." });
}

// ─── Utilidades ──────────────────────────────────────────────────────────────

function buildIdx(headers, keys) {
  const idx = {};
  keys.forEach(k => { idx[k] = headers.indexOf(k); });
  return idx;
}

function rangesOverlap(aStart, aEnd, bStart, bEnd) {
  return aStart < bEnd && bStart < aEnd;
}

function weekStart(dateStr) {
  const d = new Date(dateStr + "T12:00:00Z");
  const day = d.getUTCDay();
  const diff = day === 0 ? -6 : 1 - day;
  d.setUTCDate(d.getUTCDate() + diff);
  return d.toISOString().slice(0, 10);
}

function weekEnd(weekStartStr) {
  const d = new Date(weekStartStr + "T12:00:00Z");
  d.setUTCDate(d.getUTCDate() + 6);
  return d.toISOString().slice(0, 10);
}

function normalizeTime(value) {
  if (value instanceof Date) {
    const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    return Utilities.formatDate(value, tz, "HH") + ":00";
  }
  return String(value);
}

function startHour(value) {
  return parseInt(normalizeTime(value).split(":")[0], 10);
}

function normalizeDate(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(value);
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
