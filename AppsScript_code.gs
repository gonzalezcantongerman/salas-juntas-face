/**
 * Backend de Reservas de Salas - FACE Financial Advisors
 * --------------------------------------------------------
 * INSTRUCCIONES DE INSTALACIÓN:
 * 1. En Google Drive, dentro de la carpeta "FACE Financial Advisors", crea una carpeta
 *    llamada "SALAS DE JUNTAS".
 * 2. Dentro de esa carpeta, crea una Hoja de Cálculo de Google llamada "Reservas Salas FACE".
 * 3. En esa hoja, en la fila 1, pon estos encabezados (en este orden):
 *    id | room | date | time | duration | advisor | createdAt
 * 4. Ve a Extensiones > Apps Script. Borra el contenido del editor y pega TODO este archivo.
 * 5. Arriba a la izquierda, guarda el proyecto (nombre: "Backend Salas FACE").
 * 6. Click en "Implementar" (Deploy) > "Nueva implementación".
 *    - Tipo: "Aplicación web".
 *    - Ejecutar como: tu cuenta.
 *    - Quién tiene acceso: "Cualquier usuario" (Anyone).
 * 7. Autoriza los permisos cuando te lo pida Google.
 * 8. Copia la URL que te da ("URL de la aplicación web") y pégala en index.html
 *    en la constante API_URL.
 * 9. Cada vez que cambies este código, debes crear una NUEVA implementación (o "Manage deployments"
 *    > editar > nueva versión) para que los cambios surtan efecto en la URL pública.
 */

const SHEET_NAME = "Reservas"; // nombre de la pestaña dentro de la hoja de cálculo

// Estos asesores pueden cancelar la reserva de cualquier persona, no solo las propias.
const ADMINS = ["Ferdinando Santiago Hoyos", "Pablo Noriega Bello", "Germán González Cantón"];

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["id", "room", "date", "time", "duration", "advisor", "createdAt"]);
  }
  return sheet;
}

function doGet(e) {
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

function rangesOverlap(aStart, aEnd, bStart, bEnd) {
  return aStart < bEnd && bStart < aEnd;
}

function normalizeTime(value) {
  if (value instanceof Date) {
    // Sheets auto-convierte "HH:00" a un valor interno de fecha/hora.
    // Usamos la zona horaria de la hoja y solo el componente de hora (descartando minutos/segundos
    // que pueden quedar desfasados por precisión de punto flotante, ya que siempre son horas exactas).
    const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    const hh = Utilities.formatDate(value, tz, "HH");
    return hh + ":00";
  }
  return String(value);
}

function startHour(value) {
  return parseInt(normalizeTime(value).split(":")[0], 10);
}

function handleCreate(sheet, body) {
  const { room, date, time, advisor } = body;
  const duration = Math.max(1, parseInt(body.duration || 1, 10));
  if (!room || !date || !time || !advisor) {
    return jsonResponse({ ok: false, message: "Faltan datos para la reserva." });
  }

  // Bloqueo simple para evitar reservas dobles por solicitudes simultáneas
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const idx = {
      room: headers.indexOf("room"),
      date: headers.indexOf("date"),
      time: headers.indexOf("time"),
      duration: headers.indexOf("duration"),
    };

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

    const id = Utilities.getUuid();
    const newRow = sheet.getLastRow() + 1;
    // Forzar la columna de hora a texto plano para que Sheets no la auto-convierta a un valor de fecha/hora.
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
    const isAdmin = ADMINS.indexOf(advisor) !== -1;
    if (!isOwner && !isAdmin) {
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

    sheet.getRange(rowIndex + 1, idx.duration + 1).setValue(newDuration);
    return jsonResponse({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function handleCancel(sheet, body) {
  const { id, advisor } = body;
  if (!id) return jsonResponse({ ok: false, message: "Falta el id de la reserva." });

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf("id");
  const advisorIdx = headers.indexOf("advisor");

  for (let i = 1; i < data.length; i++) {
    if (data[i][idIdx] === id) {
      const isOwner = data[i][advisorIdx] === advisor;
      const isAdmin = ADMINS.indexOf(advisor) !== -1;
      if (!isOwner && !isAdmin) {
        return jsonResponse({ ok: false, message: "Solo puedes cancelar tus propias reservas." });
      }
      sheet.deleteRow(i + 1);
      return jsonResponse({ ok: true });
    }
  }
  return jsonResponse({ ok: false, message: "No se encontró la reserva." });
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
