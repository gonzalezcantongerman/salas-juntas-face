/**
 * Backend de Reservas de Salas - FACE Financial Advisors
 * --------------------------------------------------------
 * INSTRUCCIONES DE INSTALACIÓN:
 * 1. En Google Drive, dentro de la carpeta "FACE Financial Advisors", crea una carpeta
 *    llamada "SALAS DE JUNTAS".
 * 2. Dentro de esa carpeta, crea una Hoja de Cálculo de Google llamada "Reservas Salas FACE".
 * 3. En esa hoja, en la pestaña "Reservas", pon estos encabezados (fila 1):
 *    id | room | date | time | duration | advisor | createdAt
 * 4. Ve a Extensiones > Apps Script. Borra el contenido del editor y pega TODO este archivo.
 * 5. Guarda el proyecto (nombre: "Backend Salas FACE").
 * 6. Click en "Implementar" > "Nueva implementación".
 *    - Tipo: "Aplicación web".
 *    - Ejecutar como: tu cuenta.
 *    - Quién tiene acceso: "Cualquier usuario" (Anyone).
 * 7. Autoriza los permisos cuando te lo pida Google.
 * 8. Copia la URL y pégala en index.html en la constante API_URL.
 * 9. Cada cambio requiere una NUEVA implementación (o editar versión existente).
 *
 * GESTIÓN DE ASESORES (sin tocar el código):
 * - La pestaña "Asesores" se crea automáticamente la primera vez con los 40 asesores iniciales.
 * - Desde ahí puedes: agregar/quitar asesores, activarlos/desactivarlos, definir admins
 *   y ajustar el límite de horas semanales por persona.
 * - Columnas de "Asesores": nombre | activo | esAdmin | limiteHorasSemanales
 */

const SHEET_NAME = "Reservas";
const ADVISORS_SHEET_NAME = "Asesores";

// Lista inicial de asesores (solo se usa para pre-poblar la hoja "Asesores" si está vacía).
const DEFAULT_ADVISORS = [
  "Alejandro Argüelles Gonzalez","Alejandro Martinez Romero","Alejandro Sánchez Barbara",
  "Alejandro Tapia Castillo","Alberto Baigts Cadeño","Alberto Cortes Bolio",
  "Amairani Pamela Merino Rios","Andres Hoyos Morales","Arturo Rodríguez Muraira",
  "Carlos Alberto Duarte Morales","Eduardo Giralt Dominguez","Elisa Islas Diaz de León",
  "Erick Omar Villegas Jiménez","Ferdinando Santiago Hoyos","German Gonzalez Canton",
  "Hector Joaquin Lopez Osorio","Jonathan Alexis Bojorquez Ruelas","Jorge García Granados",
  "Jorge Roberto Rivera Medrano","Jose Abraham Grafias","Jose Pablo Flores Rabasa",
  "Juan Carlos Alonso","Juan Carlos Zacarias","Julian Cardona",
  "Laura Quintero","Lawrence Eden Anaya","Luis Guillermo Garcia Dueñas Cobarruvias",
  "Marco Antonio Olivas Betancourt","Maria Clara Cuenca De La Concha",
  "Maria de Lourdes Koloffon Lara","Maria de Lourdes Pastor Farril",
  "Maria del Rocio Alonso Martinez","Maria Rosalia Jimenez Ortega",
  "Mario Alejandro Gongora Rivera","Mauricio Garza Castro","Moises Arres Muñiz",
  "Pablo Noriega Bello","Paula Yolanda Diaz Herrera","Ramon Alejandro Gonzalez Vega",
  "Sofia Mejorada Perez Verdia"
];
const DEFAULT_ADMINS = ["Ferdinando Santiago Hoyos", "Pablo Noriega Bello", "German Gonzalez Canton"];
const DEFAULT_WEEKLY_LIMIT = 999; // límite inicial muy alto; ajusta por asesor en la hoja

// ─── Helpers de hojas ───────────────────────────────────────────────────────

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
    // Pre-poblar con la lista inicial
    DEFAULT_ADVISORS.forEach(name => {
      const isAdmin = DEFAULT_ADMINS.indexOf(name) !== -1;
      sheet.appendRow([name, true, isAdmin, DEFAULT_WEEKLY_LIMIT]);
    });
    // Formato de cabecera
    sheet.getRange(1, 1, 1, 4).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** Lee todos los asesores de la hoja "Asesores". */
function readAdvisors() {
  const sheet = getAdvisorsSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const idx = {
    nombre: headers.indexOf("nombre"),
    activo: headers.indexOf("activo"),
    esAdmin: headers.indexOf("esAdmin"),
    limiteHoras: headers.indexOf("limiteHorasSemanales")
  };
  return data.map(row => ({
    nombre: String(row[idx.nombre]).trim(),
    activo: row[idx.activo] === true || String(row[idx.activo]).toLowerCase() === "true",
    esAdmin: row[idx.esAdmin] === true || String(row[idx.esAdmin]).toLowerCase() === "true",
    limiteHoras: Number(row[idx.limiteHoras]) || DEFAULT_WEEKLY_LIMIT
  }));
}

function isAdmin(name, advisors) {
  return advisors.some(a => a.nombre === name && a.esAdmin);
}

// ─── Endpoints ───────────────────────────────────────────────────────────────

function doGet(e) {
  const action = e.parameter.action;

  // Endpoint de asesores: devuelve la lista completa (activos e inactivos) para el frontend
  if (action === "advisors") {
    const advisors = readAdvisors();
    return jsonResponse(advisors);
  }

  // Endpoint de reservas (existente)
  const date = e.parameter.date;
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const idx = {
    id: headers.indexOf("id"),
    room: headers.indexOf("room"),
    date: headers.indexOf("date"),
    time: headers.indexOf("time"),
    duration: headers.indexOf("duration"),
    advisor: headers.indexOf("advisor"),
  };

  const results = data
    .filter(row => row[idx.id] !== "" && (!date || normalizeDate(row[idx.date]) === date))
    .map(row => ({
      id: row[idx.id],
      room: row[idx.room],
      date: normalizeDate(row[idx.date]),
      time: normalizeTime(row[idx.time]),
      duration: Number(row[idx.duration]) || 1,
      advisor: row[idx.advisor]
    }));

  return jsonResponse(results);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const sheet = getSheet();

  if (body.action === "create") {
    return handleCreate(sheet, body);
  } else if (body.action === "edit") {
    return handleEdit(sheet, body);
  } else if (body.action === "cancel") {
    return handleCancel(sheet, body);
  }
  return jsonResponse({ ok: false, message: "Acción no reconocida." });
}

// ─── Lógica de reservas ──────────────────────────────────────────────────────

function rangesOverlap(aStart, aEnd, bStart, bEnd) {
  return aStart < bEnd && bStart < aEnd;
}

/**
 * Calcula el lunes de la semana a la que pertenece una fecha dada (string "yyyy-MM-dd").
 * Devuelve un string "yyyy-MM-dd".
 */
function weekStart(dateStr) {
  const d = new Date(dateStr + "T12:00:00Z"); // noon UTC para evitar off-by-one de zona
  const day = d.getUTCDay(); // 0=Dom, 1=Lun…
  const diff = day === 0 ? -6 : 1 - day;
  d.setUTCDate(d.getUTCDate() + diff);
  return d.toISOString().slice(0, 10);
}

function weekEnd(weekStartStr) {
  const d = new Date(weekStartStr + "T12:00:00Z");
  d.setUTCDate(d.getUTCDate() + 6);
  return d.toISOString().slice(0, 10);
}

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

    // Verificar que el asesor esté activo
    const advisorRecord = advisors.find(a => a.nombre === advisor);
    if (!advisorRecord) {
      return jsonResponse({ ok: false, message: "Asesor no encontrado en el directorio." });
    }
    if (!advisorRecord.activo) {
      return jsonResponse({ ok: false, message: "Tu cuenta está desactivada. Contacta a un administrador." });
    }

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const idx = {
      room: headers.indexOf("room"),
      date: headers.indexOf("date"),
      time: headers.indexOf("time"),
      duration: headers.indexOf("duration"),
      advisor: headers.indexOf("advisor"),
    };

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
    const wEnd = weekEnd(wStart);
    const usedHours = data.reduce((sum, row) => {
      const rowDate = normalizeDate(row[idx.date]);
      if (row[idx.advisor] !== advisor || rowDate < wStart || rowDate > wEnd) return sum;
      return sum + (Math.max(1, Number(row[idx.duration]) || 1));
    }, 0);

    if (usedHours + duration > advisorRecord.limiteHoras) {
      return jsonResponse({
        ok: false,
        message: `Superarías tu límite semanal de ${advisorRecord.limiteHoras} hora(s). Ya tienes ${usedHours}h reservadas esta semana.`
      });
    }

    const id = Utilities.getUuid();
    const newRow = sheet.getLastRow() + 1;
    // Forzar la columna de hora a texto para que Sheets no auto-convierta el valor.
    sheet.getRange(newRow, idx.time + 1).setNumberFormat("@");
    sheet.getRange(newRow, 1, 1, 7).setValues([[id, room, date, time, duration, advisor, new Date()]]);
    return jsonResponse({ ok: true, id: id });
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
    const idx = {
      id: headers.indexOf("id"),
      room: headers.indexOf("room"),
      date: headers.indexOf("date"),
      time: headers.indexOf("time"),
      duration: headers.indexOf("duration"),
      advisor: headers.indexOf("advisor"),
    };

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][idx.id] === id) { rowIndex = i; break; }
    }
    if (rowIndex === -1) return jsonResponse({ ok: false, message: "No se encontró la reserva." });

    const isOwner = data[rowIndex][idx.advisor] === advisor;
    if (!isOwner && !isAdmin(advisor, advisors)) {
      return jsonResponse({ ok: false, message: "Solo puedes editar tus propias reservas." });
    }

    const room = data[rowIndex][idx.room];
    const date = normalizeDate(data[rowIndex][idx.date]);
    const start = startHour(data[rowIndex][idx.time]);

    const conflict = data.some((row, i) => {
      if (i === rowIndex || i === 0) return false;
      if (row[idx.room] !== room || normalizeDate(row[idx.date]) !== date) return false;
      const s = startHour(row[idx.time]);
      const sp = Math.max(1, Number(row[idx.duration]) || 1);
      return rangesOverlap(start, start + newDuration, s, s + sp);
    });

    if (conflict) {
      return jsonResponse({ ok: false, message: "La nueva duración se cruza con otra reserva." });
    }

    // Verificar límite semanal (excluir la reserva que se edita)
    const bookingAdvisor = data[rowIndex][idx.advisor];
    const advisorRecord = advisors.find(a => a.nombre === bookingAdvisor);
    if (advisorRecord) {
      const wStart = weekStart(date);
      const wEnd = weekEnd(wStart);
      const oldDuration = Math.max(1, Number(data[rowIndex][idx.duration]) || 1);
      const usedHours = data.reduce((sum, row, i) => {
        if (i === 0 || i === rowIndex) return sum;
        const rowDate = normalizeDate(row[idx.date]);
        if (row[idx.advisor] !== bookingAdvisor || rowDate < wStart || rowDate > wEnd) return sum;
        return sum + (Math.max(1, Number(row[idx.duration]) || 1));
      }, 0);
      if (usedHours + newDuration > advisorRecord.limiteHoras) {
        return jsonResponse({
          ok: false,
          message: `La nueva duración superaría el límite semanal de ${advisorRecord.limiteHoras}h.`
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
  const idIdx = headers.indexOf("id");
  const advisorIdx = headers.indexOf("advisor");

  for (let i = 1; i < data.length; i++) {
    if (data[i][idIdx] === id) {
      const isOwner = data[i][advisorIdx] === advisor;
      if (!isOwner && !isAdmin(advisor, advisors)) {
        return jsonResponse({ ok: false, message: "Solo puedes cancelar tus propias reservas." });
      }
      sheet.deleteRow(i + 1);
      return jsonResponse({ ok: true });
    }
  }
  return jsonResponse({ ok: false, message: "No se encontró la reserva." });
}

// ─── Utilidades ──────────────────────────────────────────────────────────────

function normalizeTime(value) {
  if (value instanceof Date) {
    const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    const hh = Utilities.formatDate(value, tz, "HH");
    return hh + ":00";
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
