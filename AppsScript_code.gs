/**
 * Backend de Reservas - FACE Financial Advisors
 * -----------------------------------------------
 * Gestiona dos módulos:
 *   • Salas de Juntas  → hojas "Reservas" y "Asesores"
 *   • Puestos de Trabajo → hojas "ReservasPuestos" y "Puestos de Trabajo"
 *
 * GESTIÓN SIN CÓDIGO:
 *   - "Asesores"         → activar/desactivar, admins, límite horas salas y puestos
 *   - "Puestos de Trabajo" → marcar disponible TRUE/FALSE por puesto
 *
 * Cada cambio requiere Nueva Versión en Administrar Implementaciones.
 */

// ─── Constantes ───────────────────────────────────────────────────────────────
const SHEET_NAME               = "Reservas";
const ADVISORS_SHEET_NAME      = "Asesores";
const PUESTOS_CONFIG_SHEET     = "Puestos de Trabajo";
const RESERVAS_PUESTOS_SHEET   = "ReservasPuestos";

const DEFAULT_ADMINS       = ["Ferdinando Santiago Hoyos","Pablo Noriega Bello","Germán González Cantón"];
const DEFAULT_WEEKLY_LIMIT = 999;

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

// ─── Helpers de hojas ────────────────────────────────────────────────────────

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET_NAME);
  if (!s) {
    s = ss.insertSheet(SHEET_NAME);
    s.appendRow(["id","room","date","time","duration","advisor","createdAt"]);
  }
  return s;
}

function getAdvisorsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(ADVISORS_SHEET_NAME);
  if (!s) {
    s = ss.insertSheet(ADVISORS_SHEET_NAME);
    s.appendRow(["nombre","activo","esAdmin","limiteHorasSemanales","limiteHorasPuestoSemanales"]);
    DEFAULT_ADVISORS.forEach(name => {
      s.appendRow([name, true, DEFAULT_ADMINS.indexOf(name)!==-1, DEFAULT_WEEKLY_LIMIT, DEFAULT_WEEKLY_LIMIT]);
    });
    s.getRange(1,1,1,5).setFontWeight("bold");
    s.setFrozenRows(1);
  }
  return s;
}

/** Corre una vez para agregar la columna limiteHorasPuestoSemanales si aún no existe. */
function migrateAdvisorsSheet() {
  const s = getAdvisorsSheet();
  const headers = s.getRange(1,1,1,s.getLastColumn()).getValues()[0];
  if (headers.indexOf("limiteHorasPuestoSemanales") === -1) {
    const col = s.getLastColumn() + 1;
    s.getRange(1, col).setValue("limiteHorasPuestoSemanales").setFontWeight("bold");
    const rows = s.getLastRow() - 1;
    if (rows > 0) {
      s.getRange(2, col, rows, 1).setValue(DEFAULT_WEEKLY_LIMIT);
    }
    Logger.log("Columna limiteHorasPuestoSemanales agregada.");
  } else {
    Logger.log("La columna ya existe, nada que hacer.");
  }
}

function getPuestosConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(PUESTOS_CONFIG_SHEET);
  if (!s) {
    s = ss.insertSheet(PUESTOS_CONFIG_SHEET);
    s.appendRow(["id","nombre","disponible","puestoFijo"]);
    for (let i = 1; i <= 16; i++) {
      s.appendRow(["P"+i, "Puesto #"+i, i !== 1, ""]); // #1 starts as false
    }
    s.getRange(1,1,1,4).setFontWeight("bold");
    s.setFrozenRows(1);
  }
  return s;
}

/** Corre una vez para agregar la columna limiteHorasEscenario a la hoja "Asesores". */
function migrateAdvisorsEscenario() {
  const s = getAdvisorsSheet();
  const headers = s.getRange(1,1,1,s.getLastColumn()).getValues()[0];
  if (headers.indexOf("limiteHorasEscenario") === -1) {
    const col = s.getLastColumn() + 1;
    s.getRange(1, col).setValue("limiteHorasEscenario").setFontWeight("bold");
    const rows = s.getLastRow() - 1;
    if (rows > 0) s.getRange(2, col, rows, 1).setValue(DEFAULT_WEEKLY_LIMIT);
    Logger.log("Columna limiteHorasEscenario agregada a Asesores.");
  } else {
    Logger.log("La columna limiteHorasEscenario ya existe.");
  }
}

/** Corre una vez para agregar la fila ESCENARIO a la hoja "Puestos de Trabajo". */
function migrateEscenarioEntry() {
  const s = getPuestosConfigSheet();
  const data = s.getDataRange().getValues();
  const ids = data.slice(1).map(r => String(r[0]).trim());
  if (!ids.includes("ESCENARIO")) {
    s.appendRow(["ESCENARIO", "Escenario", true, ""]);
    Logger.log("Fila ESCENARIO agregada a Puestos de Trabajo.");
  } else {
    Logger.log("La fila ESCENARIO ya existe.");
  }
}

/** Corre una vez para agregar la columna puestoFijo a la hoja "Puestos de Trabajo" si aún no existe. */
function migratePuestosSheet() {
  const s = getPuestosConfigSheet();
  const headers = s.getRange(1,1,1,s.getLastColumn()).getValues()[0];
  if (headers.indexOf("puestoFijo") === -1) {
    const col = s.getLastColumn() + 1;
    s.getRange(1, col).setValue("puestoFijo").setFontWeight("bold");
    Logger.log("Columna puestoFijo agregada a Puestos de Trabajo.");
  } else {
    Logger.log("La columna puestoFijo ya existe.");
  }
}

function getReservasPuestosSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(RESERVAS_PUESTOS_SHEET);
  if (!s) {
    s = ss.insertSheet(RESERVAS_PUESTOS_SHEET);
    s.appendRow(["id","puesto","date","time","duration","advisor","createdAt","evento"]);
    s.getRange(1,1,1,8).setFontWeight("bold");
    s.setFrozenRows(1);
  }
  return s;
}

// ─── Lectura de configuración ────────────────────────────────────────────────

function readAdvisors() {
  const s = getAdvisorsSheet();
  const data = s.getDataRange().getValues();
  const headers = data.shift();
  const idx = buildIdx(headers, ["nombre","activo","esAdmin","limiteHorasSemanales","limiteHorasPuestoSemanales","limiteHorasEscenario"]);
  return data
    .filter(row => String(row[idx.nombre]).trim() !== "")
    .map(row => ({
      nombre:               String(row[idx.nombre]).trim(),
      activo:               row[idx.activo]===true  || String(row[idx.activo]).toLowerCase()==="true",
      esAdmin:              row[idx.esAdmin]===true || String(row[idx.esAdmin]).toLowerCase()==="true",
      limiteHoras:          idx.limiteHorasSemanales>=0 && Number(row[idx.limiteHorasSemanales])>0
                              ? Number(row[idx.limiteHorasSemanales]) : DEFAULT_WEEKLY_LIMIT,
      limiteHorasPuesto:    idx.limiteHorasPuestoSemanales>=0 && Number(row[idx.limiteHorasPuestoSemanales])>0
                              ? Number(row[idx.limiteHorasPuestoSemanales]) : DEFAULT_WEEKLY_LIMIT,
      limiteHorasEscenario: idx.limiteHorasEscenario>=0 && Number(row[idx.limiteHorasEscenario])>0
                              ? Number(row[idx.limiteHorasEscenario]) : DEFAULT_WEEKLY_LIMIT
    }));
}

function readPuestosConfig() {
  const s = getPuestosConfigSheet();
  const data = s.getDataRange().getValues();
  const headers = data.shift();
  const idx = buildIdx(headers, ["id","nombre","disponible","puestoFijo"]);
  return data
    .filter(row => String(row[idx.id]).trim() !== "")
    .map(row => {
      const fijo = idx.puestoFijo >= 0 ? String(row[idx.puestoFijo]).trim() : "";
      const fijoVal = (fijo === "" || fijo.toLowerCase() === "false") ? "" : fijo;
      return {
        id:         String(row[idx.id]).trim(),
        nombre:     String(row[idx.nombre]).trim(),
        disponible: row[idx.disponible]===true || String(row[idx.disponible]).toLowerCase()==="true",
        puestoFijo: fijoVal   // "" = libre, "true" = bloqueado sin nombre, o nombre del asesor
      };
    });
}

// ─── Normalización de nombres ─────────────────────────────────────────────────

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
  const date   = e.parameter.date;

  // Catálogo de asesores
  if (action === "advisors") return jsonResponse(readAdvisors());

  // Configuración de puestos
  if (action === "puestos") return jsonResponse(readPuestosConfig());

  // Reservas de puestos de trabajo
  if (action === "reservasPuestos") {
    const s = getReservasPuestosSheet();
    const data = s.getDataRange().getValues();
    const headers = data.shift();
    const idx = buildIdx(headers, ["id","puesto","date","time","duration","advisor","evento"]);
    const rows = data
      .filter(r => r[idx.id] !== "" && (!date || normalizeDate(r[idx.date])===date))
      .map(r => ({
        id:       r[idx.id],
        puesto:   r[idx.puesto],
        date:     normalizeDate(r[idx.date]),
        time:     normalizeTime(r[idx.time]),
        duration: Number(r[idx.duration])||1,
        advisor:  r[idx.advisor],
        evento:   idx.evento >= 0 ? (String(r[idx.evento]||'').trim()||undefined) : undefined
      }));

    // Inyectar reservas virtuales de puestos fijos (solo cuando hay filtro de fecha)
    if (date) {
      const pConfig = readPuestosConfig();
      pConfig.forEach(pc => {
        if (!pc.puestoFijo) return;
        const advisorFijo = pc.puestoFijo.toLowerCase() === "true" ? "—" : pc.puestoFijo;
        // Solo inyectar si no hay ya una reserva real que cubra todo el día
        const alreadyCovered = rows.some(r => r.puesto === pc.id && r.date === date);
        if (!alreadyCovered) {
          rows.push({
            id:       "fixed_" + pc.id + "_" + date,
            puesto:   pc.id,
            date:     date,
            time:     "07:00",
            duration: 13,   // 07:00–20:00
            advisor:  advisorFijo,
            fixed:    true
          });
        }
      });
    }

    return jsonResponse(rows);
  }

  // Reservas de salas (con o sin filtro de fecha)
  const s = getSheet();
  const data = s.getDataRange().getValues();
  const headers = data.shift();
  const idx = buildIdx(headers, ["id","room","date","time","duration","advisor"]);
  const rows = data
    .filter(r => r[idx.id] !== "" && (!date || normalizeDate(r[idx.date])===date))
    .map(r => ({
      id:       r[idx.id],
      room:     r[idx.room],
      date:     normalizeDate(r[idx.date]),
      time:     normalizeTime(r[idx.time]),
      duration: Number(r[idx.duration])||1,
      advisor:  r[idx.advisor]
    }));
  return jsonResponse(rows);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);

  if (body.action === "create")        return handleCreate(getSheet(), body);
  if (body.action === "edit")          return handleEdit(getSheet(), body);
  if (body.action === "cancel")        return handleCancel(getSheet(), body);
  if (body.action === "createPuesto")  return handleCreatePuesto(getReservasPuestosSheet(), body);
  if (body.action === "editPuesto")    return handleEditPuesto(getReservasPuestosSheet(), body);
  if (body.action === "cancelPuesto")  return handleCancelPuesto(getReservasPuestosSheet(), body);

  return jsonResponse({ ok:false, message:"Acción no reconocida." });
}

// ─── Validaciones comunes ────────────────────────────────────────────────────

/** Verifica si el asesor ya tiene reserva en salas O puestos en ese horario.
 *  Los administradores están exentos — llama esta función solo para no-admins. */
function advisorHasConcurrentBooking(advisor, date, startH, duration, excludeId) {
  // Salas
  const sd = getSheet().getDataRange().getValues();
  if(sd.length > 1){
    const si = buildIdx(sd[0], ['id','date','time','duration','advisor']);
    for(let i=1;i<sd.length;i++){
      const r=sd[i]; if(r[si.id]===excludeId) continue;
      if(normalizeName(r[si.advisor])!==normalizeName(advisor)) continue;
      if(normalizeDate(r[si.date])!==date) continue;
      const s=startHour(r[si.time]), d2=Math.max(1,Number(r[si.duration])||1);
      if(rangesOverlap(startH,startH+duration,s,s+d2)) return true;
    }
  }
  // Puestos
  const pd = getReservasPuestosSheet().getDataRange().getValues();
  if(pd.length > 1){
    const pi = buildIdx(pd[0], ['id','date','time','duration','advisor']);
    for(let i=1;i<pd.length;i++){
      const r=pd[i]; if(r[pi.id]===excludeId) continue;
      if(normalizeName(r[pi.advisor])!==normalizeName(advisor)) continue;
      if(normalizeDate(r[pi.date])!==date) continue;
      const s=startHour(r[pi.time]), d2=Math.max(1,Number(r[pi.duration])||1);
      if(rangesOverlap(startH,startH+duration,s,s+d2)) return true;
    }
  }
  return false;
}

// ─── Lógica Salas ─────────────────────────────────────────────────────────────

function handleCreate(sheet, body) {
  const { room, date, time, advisor } = body;
  const duration = Math.max(1, parseInt(body.duration||1,10));
  if (!room||!date||!time||!advisor) return jsonResponse({ok:false,message:"Faltan datos."});

  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const advisors = readAdvisors();
    const rec = findAdvisor(advisor, advisors);
    if (!rec)       return jsonResponse({ok:false,message:"Asesor no encontrado. Contacta a un administrador."});
    if (!rec.activo) return jsonResponse({ok:false,message:"Tu cuenta está desactivada."});

    // Máximo 1 mes de anticipación (no-admins)
    if(!isAdmin(advisor, advisors)){
      const maxD=new Date(); maxD.setMonth(maxD.getMonth()+1);
      if(new Date(date+'T12:00:00Z')>maxD)
        return jsonResponse({ok:false,message:`Solo puedes reservar con máximo 1 mes de anticipación (hasta el ${maxD.toISOString().slice(0,10)}).`});
    }

    const data = sheet.getDataRange().getValues(); const headers = data.shift();
    const idx = buildIdx(headers,["room","date","time","duration","advisor"]);

    const newStart = startHour(time);
    const conflict = data.some(r => {
      if (r[idx.room]!==room || normalizeDate(r[idx.date])!==date) return false;
      return rangesOverlap(newStart,newStart+duration,startHour(r[idx.time]),startHour(r[idx.time])+(Math.max(1,Number(r[idx.duration])||1)));
    });
    if (conflict) return jsonResponse({ok:false,message:"Ese horario se cruza con otra reserva existente."});

    // Reserva simultánea del mismo asesor (no-admins)
    if(!isAdmin(advisor, advisors) && advisorHasConcurrentBooking(advisor, date, newStart, duration, null))
      return jsonResponse({ok:false,message:"Ya tienes una reserva en ese horario. No puedes tener dos reservas simultáneas."});

    const wStart=weekStart(date), wEnd=weekEnd(wStart);
    const used = data.reduce((s,r)=>{
      const rd=normalizeDate(r[idx.date]);
      if (normalizeName(r[idx.advisor])!==normalizeName(advisor)||rd<wStart||rd>wEnd) return s;
      return s+Math.max(1,Number(r[idx.duration])||1);
    },0);
    if (used+duration>rec.limiteHoras) return jsonResponse({ok:false,message:`Límite semanal de ${rec.limiteHoras}h alcanzado. Ya tienes ${used}h esta semana. Contacta al área directiva.`});

    const id=Utilities.getUuid(), newRow=sheet.getLastRow()+1;
    sheet.getRange(newRow,idx.time+1).setNumberFormat("@");
    sheet.getRange(newRow,1,1,7).setValues([[id,room,date,time,duration,advisor,new Date()]]);
    return jsonResponse({ok:true,id});
  } finally { lock.releaseLock(); }
}

function handleEdit(sheet, body) {
  const {id,advisor} = body;
  const newDuration = Math.max(1,parseInt(body.duration||1,10));
  if (!id) return jsonResponse({ok:false,message:"Falta el id."});

  const lock=LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const advisors=readAdvisors();
    const data=sheet.getDataRange().getValues(); const headers=data[0];
    const idx=buildIdx(headers,["id","room","date","time","duration","advisor"]);

    let ri=-1;
    for(let i=1;i<data.length;i++){if(data[i][idx.id]===id){ri=i;break;}}
    if(ri===-1) return jsonResponse({ok:false,message:"Reserva no encontrada."});

    if(data[ri][idx.advisor]!==advisor && !isAdmin(advisor,advisors))
      return jsonResponse({ok:false,message:"Solo puedes editar tus propias reservas."});

    const bDate=normalizeDate(data[ri][idx.date]);
    const bHour=startHour(data[ri][idx.time]);
    if(new Date(bDate+"T"+String(bHour).padStart(2,"0")+":00:00")<new Date())
      return jsonResponse({ok:false,message:"No se puede editar una reserva que ya ocurrió."});

    const room=data[ri][idx.room], date=bDate, start=bHour;
    const conflict=data.some((r,i)=>{
      if(i===ri||i===0||r[idx.room]!==room||normalizeDate(r[idx.date])!==date) return false;
      const s=startHour(r[idx.time]),sp=Math.max(1,Number(r[idx.duration])||1);
      return rangesOverlap(start,start+newDuration,s,s+sp);
    });
    if(conflict) return jsonResponse({ok:false,message:"La nueva duración se cruza con otra reserva."});

    const ba=data[ri][idx.advisor], rec=findAdvisor(ba,advisors);
    if(rec){
      const wS=weekStart(date),wE=weekEnd(wS);
      const used=data.reduce((s,r,i)=>{
        if(i===0||i===ri) return s;
        const rd=normalizeDate(r[idx.date]);
        if(normalizeName(r[idx.advisor])!==normalizeName(ba)||rd<wS||rd>wE) return s;
        return s+Math.max(1,Number(r[idx.duration])||1);
      },0);
      if(used+newDuration>rec.limiteHoras) return jsonResponse({ok:false,message:`Supera el límite semanal de ${rec.limiteHoras}h.`});
    }
    sheet.getRange(ri+1,idx.duration+1).setValue(newDuration);
    return jsonResponse({ok:true});
  } finally { lock.releaseLock(); }
}

function handleCancel(sheet, body) {
  const {id,advisor}=body;
  if(!id) return jsonResponse({ok:false,message:"Falta el id."});
  const advisors=readAdvisors();
  const data=sheet.getDataRange().getValues(), headers=data[0];
  const idx=buildIdx(headers,["id","date","time","advisor"]);

  for(let i=1;i<data.length;i++){
    if(data[i][idx.id]!==id) continue;
    if(data[i][idx.advisor]!==advisor && !isAdmin(advisor,advisors))
      return jsonResponse({ok:false,message:"Solo puedes cancelar tus propias reservas."});
    if(!isAdmin(advisor,advisors)){
      const dt=new Date(normalizeDate(data[i][idx.date])+"T"+String(startHour(data[i][idx.time])).padStart(2,"0")+":00:00");
      if((dt-new Date())<2*60*60*1000) return jsonResponse({ok:false,message:"No puedes cancelar con menos de 2h de anticipación."});
    }
    sheet.deleteRow(i+1);
    return jsonResponse({ok:true});
  }
  return jsonResponse({ok:false,message:"Reserva no encontrada."});
}

// ─── Lógica Puestos de Trabajo ────────────────────────────────────────────────

function handleCreatePuesto(sheet, body) {
  const {puesto, date, time, advisor} = body;
  const duration = Math.max(1,parseInt(body.duration||1,10));
  const evento   = String(body.evento||'').trim();
  if (!puesto||!date||!time||!advisor) return jsonResponse({ok:false,message:"Faltan datos."});
  if (puesto === "ESCENARIO" && !evento) return jsonResponse({ok:false,message:"El campo Evento es obligatorio para el Escenario."});

  // Verificar que el puesto esté disponible en la config
  const pConfig = readPuestosConfig();
  const pc = pConfig.find(p=>p.id===puesto);
  if (!pc||!pc.disponible) return jsonResponse({ok:false,message:"Este puesto no está disponible."});

  // Verificar puesto fijo — solo el asesor asignado (o admin) puede reservarlo
  if (pc.puestoFijo && pc.puestoFijo.toLowerCase() !== "true") {
    const lock0 = LockService.getScriptLock(); lock0.waitLock(10000);
    try {
      const advisors0 = readAdvisors();
      if (!isAdmin(advisor, advisors0) && normalizeName(pc.puestoFijo) !== normalizeName(advisor))
        return jsonResponse({ok:false,message:`Este puesto está asignado a ${pc.puestoFijo}.`});
    } finally { lock0.releaseLock(); }
  } else if (pc.puestoFijo && pc.puestoFijo.toLowerCase() === "true") {
    const advisors0 = readAdvisors();
    if (!isAdmin(advisor, advisors0))
      return jsonResponse({ok:false,message:"Este puesto no está disponible para reservas."});
  }

  const lock=LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const advisors=readAdvisors();
    const rec=findAdvisor(advisor,advisors);
    if(!rec)       return jsonResponse({ok:false,message:"Asesor no encontrado."});
    if(!rec.activo) return jsonResponse({ok:false,message:"Tu cuenta está desactivada."});

    // Máximo 1 mes de anticipación (no-admins)
    if(!isAdmin(advisor, advisors)){
      const maxD=new Date(); maxD.setMonth(maxD.getMonth()+1);
      if(new Date(date+'T12:00:00Z')>maxD)
        return jsonResponse({ok:false,message:`Solo puedes reservar con máximo 1 mes de anticipación (hasta el ${maxD.toISOString().slice(0,10)}).`});
    }

    const data=sheet.getDataRange().getValues(); const headers=data.shift();
    const idx=buildIdx(headers,["puesto","date","time","duration","advisor"]);

    const newStart=startHour(time);
    const conflict=data.some(r=>{
      if(r[idx.puesto]!==puesto||normalizeDate(r[idx.date])!==date) return false;
      return rangesOverlap(newStart,newStart+duration,startHour(r[idx.time]),startHour(r[idx.time])+Math.max(1,Number(r[idx.duration])||1));
    });
    if(conflict) return jsonResponse({ok:false,message:"Ese horario ya está ocupado en este puesto."});

    // Reserva simultánea del mismo asesor (no-admins)
    if(!isAdmin(advisor, advisors) && advisorHasConcurrentBooking(advisor, date, newStart, duration, null))
      return jsonResponse({ok:false,message:"Ya tienes una reserva en ese horario. No puedes tener dos reservas simultáneas."});

    const isEscenario = puesto === "ESCENARIO";
    const wS=weekStart(date),wE=weekEnd(wS);
    const used=data.reduce((s,r)=>{
      const rd=normalizeDate(r[idx.date]);
      if(normalizeName(r[idx.advisor])!==normalizeName(advisor)||rd<wS||rd>wE) return s;
      // Contar solo el tipo correcto: escenario o puestos (separados)
      if(isEscenario && r[idx.puesto]!=="ESCENARIO") return s;
      if(!isEscenario && r[idx.puesto]==="ESCENARIO") return s;
      return s+Math.max(1,Number(r[idx.duration])||1);
    },0);
    const limiteAplicable = isEscenario ? rec.limiteHorasEscenario : rec.limiteHorasPuesto;
    const tipoLabel = isEscenario ? "escenario" : "puestos";
    if(used+duration>limiteAplicable) return jsonResponse({ok:false,message:`Límite semanal de ${tipoLabel} (${limiteAplicable}h) alcanzado. Ya tienes ${used}h esta semana.`});

    const id=Utilities.getUuid(), newRow=sheet.getLastRow()+1;
    sheet.getRange(newRow,idx.time+1).setNumberFormat("@");
    sheet.getRange(newRow,1,1,8).setValues([[id,puesto,date,time,duration,advisor,new Date(),evento]]);
    return jsonResponse({ok:true,id});
  } finally { lock.releaseLock(); }
}

function handleEditPuesto(sheet, body) {
  const {id, advisor} = body;
  const newTime     = body.time     || null;   // nueva hora inicio (opcional)
  const newDuration = Math.max(1, parseInt(body.duration||1, 10));
  if(!id) return jsonResponse({ok:false,message:"Falta el id."});

  const lock=LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const advisors=readAdvisors();
    const data=sheet.getDataRange().getValues(), headers=data[0];
    const idx=buildIdx(headers,["id","puesto","date","time","duration","advisor"]);

    let ri=-1;
    for(let i=1;i<data.length;i++){if(data[i][idx.id]===id){ri=i;break;}}
    if(ri===-1) return jsonResponse({ok:false,message:"Reserva no encontrada."});
    if(data[ri][idx.advisor]!==advisor&&!isAdmin(advisor,advisors))
      return jsonResponse({ok:false,message:"Solo puedes editar tus propias reservas."});

    const bDate=normalizeDate(data[ri][idx.date]);
    const bHour=startHour(data[ri][idx.time]);
    const bDt=new Date(bDate+"T"+String(bHour).padStart(2,"0")+":00:00");
    if(bDt<new Date())
      return jsonResponse({ok:false,message:"No se puede editar una reserva que ya ocurrió."});
    if(!isAdmin(advisor,advisors)&&(bDt-new Date())<12*60*60*1000)
      return jsonResponse({ok:false,message:"No puedes editar con menos de 12h de anticipación."});

    const puesto=data[ri][idx.puesto];
    const newStart=newTime?startHour(newTime):bHour;
    const conflict=data.some((r,i)=>{
      if(i===ri||i===0||r[idx.puesto]!==puesto||normalizeDate(r[idx.date])!==bDate) return false;
      const s=startHour(r[idx.time]),sp=Math.max(1,Number(r[idx.duration])||1);
      return rangesOverlap(newStart,newStart+newDuration,s,s+sp);
    });
    if(conflict) return jsonResponse({ok:false,message:"El nuevo horario se cruza con otra reserva."});

    if(newTime){
      sheet.getRange(ri+1,idx.time+1).setNumberFormat("@").setValue(newTime);
    }
    sheet.getRange(ri+1,idx.duration+1).setValue(newDuration);
    return jsonResponse({ok:true});
  } finally { lock.releaseLock(); }
}

function handleCancelPuesto(sheet, body) {
  const {id,advisor}=body;
  if(!id) return jsonResponse({ok:false,message:"Falta el id."});
  const advisors=readAdvisors();
  const data=sheet.getDataRange().getValues(), headers=data[0];
  const idx=buildIdx(headers,["id","date","time","advisor"]);

  for(let i=1;i<data.length;i++){
    if(data[i][idx.id]!==id) continue;
    if(data[i][idx.advisor]!==advisor&&!isAdmin(advisor,advisors))
      return jsonResponse({ok:false,message:"Solo puedes cancelar tus propias reservas."});
    if(!isAdmin(advisor,advisors)){
      const dt=new Date(normalizeDate(data[i][idx.date])+"T"+String(startHour(data[i][idx.time])).padStart(2,"0")+":00:00");
      if((dt-new Date())<12*60*60*1000) return jsonResponse({ok:false,message:"No puedes cancelar con menos de 12h de anticipación."});
    }
    sheet.deleteRow(i+1);
    return jsonResponse({ok:true});
  }
  return jsonResponse({ok:false,message:"Reserva no encontrada."});
}

// ─── Utilidades ──────────────────────────────────────────────────────────────

function buildIdx(headers, keys) {
  const idx={};
  keys.forEach(k=>{ idx[k]=headers.indexOf(k); });
  return idx;
}
function rangesOverlap(aS,aE,bS,bE){ return aS<bE&&bS<aE; }
function weekStart(dateStr){
  const d=new Date(dateStr+"T12:00:00Z"), day=d.getUTCDay();
  d.setUTCDate(d.getUTCDate()+(day===0?-6:1-day));
  return d.toISOString().slice(0,10);
}
function weekEnd(wS){
  const d=new Date(wS+"T12:00:00Z");
  d.setUTCDate(d.getUTCDate()+6);
  return d.toISOString().slice(0,10);
}
function normalizeTime(v){
  if(v instanceof Date){
    const tz=SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    return Utilities.formatDate(v,tz,"HH")+":00";
  }
  return String(v);
}
function startHour(v){ return parseInt(normalizeTime(v).split(":")[0],10); }
function normalizeDate(v){
  if(v instanceof Date) return Utilities.formatDate(v,Session.getScriptTimeZone(),"yyyy-MM-dd");
  return String(v);
}
function jsonResponse(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ─── Migración: agregar columna "evento" a ReservasPuestos ─────────────────
function migrateEventoColumn() {
  const s = getReservasPuestosSheet();
  const headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  if (headers.includes("evento")) {
    Logger.log("Columna 'evento' ya existe — nada que migrar.");
    return;
  }
  const col = s.getLastColumn() + 1;
  s.getRange(1, col).setValue("evento").setFontWeight("bold");
  Logger.log("Columna 'evento' agregada en columna " + col);
}
