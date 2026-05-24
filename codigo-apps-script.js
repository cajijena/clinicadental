// ═══════════════════════════════════════════════════════════════
//  GOOGLE APPS SCRIPT - Clínica Dr. Allendes
//  Pega este código completo en script.google.com
// ═══════════════════════════════════════════════════════════════

const SHEET_ID = "1QN7Q6u90A50J-FEoSg-mhd-deyggDsHqJ4vttI5c8ns";
const SHEET_NAME = "Pacientes";
const HEADERS = ["Nombre","RUT","Edad","Teléfono","Email","Estado","Notas","Creado"];

function getSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  const firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  if (firstRow[0] !== "Nombre") {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  return sheet;
}

function doGet(e) {
  const action = e.parameter.action;
  try {
    if (action === 'getAll') return respond(getAllPatients());
    return respond({error: 'Acción no reconocida'});
  } catch(err) {
    return respond({error: err.message});
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    if (action === 'append') return respond(appendPatient(data));
    if (action === 'update') return respond(updatePatient(data));
    if (action === 'delete') return respond(deletePatient(data));
    return respond({error: 'Acción no reconocida'});
  } catch(err) {
    return respond({error: err.message});
  }
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getAllPatients() {
  const sheet = getSheet();
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return {data: []};
  const data = rows.slice(1).map((row, i) => ({
    row: i + 2,
    name:   row[0] || "",
    rut:    row[1] || "",
    edad:   row[2] || "",
    tel:    row[3] || "",
    email:  row[4] || "",
    status: row[5] || "En espera",
    notes:  row[6] || "",
    created:row[7] || ""
  })).filter(p => p.name);
  return {data};
}

function appendPatient(data) {
  const sheet = getSheet();
  sheet.appendRow([
    data.name, data.rut, data.edad, data.tel, data.email,
    data.status, data.notes, new Date().toISOString()
  ]);
  return {ok: true};
}

function updatePatient(data) {
  const sheet = getSheet();
  const row = parseInt(data.row);
  const existing = sheet.getRange(row, 8).getValue();
  sheet.getRange(row, 1, 1, 8).setValues([[
    data.name, data.rut, data.edad, data.tel, data.email,
    data.status, data.notes, existing || new Date().toISOString()
  ]]);
  return {ok: true};
}

function deletePatient(data) {
  const sheet = getSheet();
  sheet.deleteRow(parseInt(data.row));
  return {ok: true};
}

// ─── SINCRONIZACIÓN CON GOOGLE CALENDAR ──────────────────────────
const CALENDAR_IDS = [
  "f44b6cd92874b8973ca7ae5f09c9efede0a1f250a6c6bfee0b6f0559a139ceb3@group.calendar.google.com",
  "163ba79f691e411e091f1ae2e5832c283b0ad2036f109a599aca7b07735cd963@group.calendar.google.com"
];

function syncCalendarToSheet() {
  const sheet = getSheet();
  const rows = sheet.getDataRange().getValues();
  const existingNames = rows.slice(1).map(r => (r[0]||"").toLowerCase().trim());

  const now = new Date();
  const past = new Date(now - 30 * 24 * 60 * 60 * 1000); // últimos 30 días
  const future = new Date(now.getTime() + 60 * 24 * 60 * 60 * 1000); // próximos 60 días

  const events = [];
  CALENDAR_IDS.forEach(id => {
    const cal = CalendarApp.getCalendarById(id);
    if (cal) events.push(...cal.getEvents(past, future));
    else Logger.log("Calendario no encontrado: " + id);
  });
  let added = 0;

  events.forEach(event => {
    const title = event.getTitle().trim();
    if (!title) return;
    if (!existingNames.includes(title.toLowerCase())) {
      sheet.appendRow([
        title, "", "", "", "", "Agendado", 
        "Agregado automáticamente desde Google Calendar", 
        new Date().toISOString()
      ]);
      existingNames.push(title.toLowerCase());
      added++;
    }
  });

  Logger.log(`Sincronización completada. ${added} pacientes nuevos agregados.`);
}
