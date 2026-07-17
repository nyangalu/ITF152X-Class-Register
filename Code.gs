/************************************************************************
 * ITF152X - INFORMATION TECHNOLOGY 1A
 * Online Attendance Register — Google Apps Script backend
 * Cape Peninsula University of Technology
 *
 * DEPLOYMENT
 * 1. Create a new Google Sheet (any name, e.g. "ITF152X Attendance Register").
 * 2. Extensions > Apps Script. Delete any starter code and paste this file
 *    in as Code.gs.
 * 3. Run the function `setupSheets` once from the Apps Script editor
 *    (select it in the toolbar dropdown, click Run). Approve the
 *    permissions prompt. This creates the Students / Attendance / Sessions
 *    tabs and loads the 57 students from the class list.
 * 4. Click Deploy > New deployment > type: "Web app".
 *      - Execute as: Me
 *      - Who has access: Anyone
 * 5. Copy the Web app URL and paste it into APPS_SCRIPT_URL at the top of
 *    index.html.
 * 6. Re-deploy (Deploy > Manage deployments > Edit > New version) any time
 *    you change this file.
 ************************************************************************/

const SS = SpreadsheetApp.getActiveSpreadsheet;
const STUDENTS_SHEET = 'Students';
const ATTENDANCE_SHEET = 'Attendance';
const SESSIONS_SHEET = 'Sessions';

// Maximum distance (in metres) allowed between the lecturer's phone/laptop
// (when the QR session was generated) and the student's phone (when they
// scan) for a check-in to be accepted as "in the same room".
const MAX_CHECKIN_DISTANCE_METERS = 60;

// A QR session is only valid for this many minutes after it is generated.
const SESSION_VALID_MINUTES = 180;

/*** ---------------------------- SETUP ---------------------------- ***/

function setupSheets() {
  const ss = SS();

  let students = ss.getSheetByName(STUDENTS_SHEET);
  if (!students) students = ss.insertSheet(STUDENTS_SHEET);
  students.clear();
  students.appendRow(['StudentNumber', 'Surname', 'Initials', 'FirstName']);
  const data = STUDENT_SEED.map(r => [r[0], r[1], r[2], r[3]]);
  students.getRange(2, 1, data.length, 4).setValues(data);
  students.getRange(1, 1, 1, 4).setFontWeight('bold');
  students.setFrozenRows(1);
  students.autoResizeColumns(1, 4);

  let attendance = ss.getSheetByName(ATTENDANCE_SHEET);
  if (!attendance) attendance = ss.insertSheet(ATTENDANCE_SHEET);
  attendance.clear();
  attendance.appendRow(['Date', 'StudentNumber', 'Status', 'Timestamp', 'Method', 'DeviceId', 'MarkedBy']);
  attendance.getRange(1, 1, 1, 7).setFontWeight('bold');
  attendance.setFrozenRows(1);

  let sessions = ss.getSheetByName(SESSIONS_SHEET);
  if (!sessions) sessions = ss.insertSheet(SESSIONS_SHEET);
  sessions.clear();
  sessions.appendRow(['SessionCode', 'Date', 'LecturerLat', 'LecturerLng', 'CreatedAt', 'Active']);
  sessions.getRange(1, 1, 1, 6).setFontWeight('bold');
  sessions.setFrozenRows(1);

  SpreadsheetApp.flush();
  return 'Setup complete: ' + data.length + ' students loaded.';
}

// FIRST NAME / SURNAME / INITIALS / STUDENT NUMBER, sourced from the
// uploaded ITF152X class list (57 students).
const STUDENT_SEED = [
  ["260933481","BABEDI","NF","NOMPUMELELO"],
  ["251637794","BALOYI","T","TIYANI"],
  ["261699148","BAMBO","KW","KGOKARI"],
  ["251246906","CANDULWANDLE","S","SESETHU"],
  ["250016478","CEFU","V","VUYOLWETHU"],
  ["260346446","DANKE","DANKET","TSHWARELO"],
  ["241480272","DLANGAMANDLA","LN","LUYANDA"],
  ["260342203","GCELU","Z","ZENIBONGE"],
  ["261216228","KEKANA","I","IPELENG"],
  ["260305073","KGAPHOLA","M","MOLEFE"],
  ["260498076","KHATHINI","MT","MLANDO"],
  ["261328220","KHONCO","KK","KHANYA"],
  ["260480223","KHOZA","HJ","HUMULANI"],
  ["250704218","KHOZA","LM","LETHUKUTHULA"],
  ["230040543","LEBONA","KK","KEAMOHETSE"],
  ["260713430","MAFOKO","WJ","WENONNA"],
  ["261051520","MAHAMBA","NO","OWAMI"],
  ["260614025","MAKGADO","TL","TSHILILO"],
  ["261023187","MAKUNGO","SO","SHUDUFHADZO"],
  ["260594997","MALATJI","RB","REFILOE"],
  ["260582484","MAPOMA","IM","INAM"],
  ["260603260","MASHIMBYE","AP","AMANDLA"],
  ["261029649","MASINGI","VP","VUKHETA"],
  ["260250023","MASONDO","LS","LUNGELO"],
  ["260558435","MATLHAKOLA","AK","ANDRIES"],
  ["240587731","MATSHIANA","P","PHUMELELE"],
  ["261006339","MATSHOBA","KL","KITSO"],
  ["260176303","MATSHOGA","TC","TLHOLOGELO"],
  ["260764647","MBELESI","PM","PHILA"],
  ["260341177","MBODI","TN","TSHINAKAHO"],
  ["260505064","MDITSHWA","O","OLWETHU"],
  ["261296183","MHLONGO","NL","NANDI"],
  ["260530875","MLOTSHWA","PSR","PROMISE"],
  ["240677110","MOKGERANE","TS","TEBOHO"],
  ["250767953","MOKHWEBANE","N","NTOMBIFUTHI"],
  ["261744801","MOSARWA","K","KEAORATA"],
  ["260179388","MOSINI","IIM","IKONA"],
  ["260289655","MOTAUNG","SA","SIPHESINHLE"],
  ["260136336","MTHUNZI","A","AFIKILE"],
  ["251287998","MZAMA","E","ESONA"],
  ["230051111","NDAMANE","A","ABULELE"],
  ["261469061","NDOU","FT","FRANS"],
  ["260826219","NELUFULE","VN","VHUTSHILO"],
  ["261059912","NETHAVHANI","M","MATHAMELA"],
  ["260059056","NKOSI","SL","SAMUKELO"],
  ["251197689","NONZOLO","Y","YANGA"],
  ["260591858","NTWAMMPI","NH","HILTON"],
  ["260898333","NWANDULE","NP","NHLAWULEKO"],
  ["261259288","RAMMUTLA","K","KGOTSO"],
  ["241463661","SIKWAMBANE","L I","LERATO"],
  ["260810029","SIMELANE","SS","SIYAMTHANDA"],
  ["260967483","STEMELA","P","PHILELA"],
  ["260606359","TEMBANI","S","SIZIPHIWE"],
  ["260588040","THAISI","K","KATLEHO"],
  ["260695629","THOKOANE","AM","ABEDNICCO"],
  ["260580198","TWALA","N","NOLWAZI"],
  ["260157732","ZAKWE","N","NOSIHLE"]
];

/*** ---------------------------- ENTRY POINTS ---------------------------- ***/

function doGet(e) {
  try {
    const action = e.parameter.action;
    let result;
    switch (action) {
      case 'getStudents':
        result = getStudents_();
        break;
      case 'getRegister':
        result = getRegister_(e.parameter.date);
        break;
      case 'getSession':
        result = getSessionInfo_(e.parameter.code);
        break;
      case 'getSummary':
        result = getSummary_(e.parameter.start, e.parameter.end);
        break;
      default:
        result = { error: 'Unknown action: ' + action };
    }
    return jsonOut_(result);
  } catch (err) {
    return jsonOut_({ error: err.message });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    let result;
    switch (action) {
      case 'saveRegister':
        result = saveRegister_(body.date, body.records, body.markedBy);
        break;
      case 'createSession':
        result = createSession_(body.date, body.lat, body.lng);
        break;
      case 'studentCheckin':
        result = studentCheckin_(body);
        break;
      default:
        result = { error: 'Unknown action: ' + action };
    }
    return jsonOut_(result);
  } catch (err) {
    return jsonOut_({ error: err.message });
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/*** ---------------------------- STUDENTS ---------------------------- ***/

function getStudents_() {
  const sh = SS().getSheetByName(STUDENTS_SHEET);
  const values = sh.getDataRange().getValues();
  values.shift(); // header
  return values
    .filter(r => r[0])
    .map(r => ({
      studentNumber: String(r[0]),
      surname: String(r[1]).trim(),
      initials: String(r[2]).trim(),
      firstName: String(r[3]).trim()
    }));
}

/*** ---------------------------- ATTENDANCE ---------------------------- ***/

// Returns the roster with each student's status for the given date.
// Defaults to Absent if no record exists yet for that date.
function getRegister_(date) {
  if (!date) throw new Error('date is required');
  const students = getStudents_();
  const records = getAttendanceMapForDate_(date);
  return students.map(s => {
    const rec = records[s.studentNumber];
    return {
      studentNumber: s.studentNumber,
      surname: s.surname,
      initials: s.initials,
      firstName: s.firstName,
      status: rec ? rec.status : 'Absent',
      method: rec ? rec.method : null,
      timestamp: rec ? rec.timestamp : null
    };
  });
}

function getAttendanceMapForDate_(date) {
  const sh = SS().getSheetByName(ATTENDANCE_SHEET);
  const values = sh.getDataRange().getValues();
  values.shift();
  const map = {};
  values.forEach(r => {
    const rowDate = formatDate_(r[0]);
    if (rowDate === date) {
      map[String(r[1])] = { status: r[2], timestamp: r[3], method: r[4] };
    }
  });
  return map;
}

// Bulk save from the lecturer's "Save" button. Overwrites existing rows
// for that date and writes new ones, in one pass.
function saveRegister_(date, records, markedBy) {
  if (!date || !records) throw new Error('date and records are required');
  const sh = SS().getSheetByName(ATTENDANCE_SHEET);
  const values = sh.getDataRange().getValues();
  const header = values[0];
  const now = new Date();

  // Build a lookup of existing row indices (1-based, including header) for this date.
  const existingRowForStudent = {};
  for (let i = 1; i < values.length; i++) {
    if (formatDate_(values[i][0]) === date) {
      existingRowForStudent[String(values[i][1])] = i + 1; // sheet row number
    }
  }

  records.forEach(rec => {
    const rowNum = existingRowForStudent[String(rec.studentNumber)];
    const rowData = [date, String(rec.studentNumber), rec.status, now, 'Lecturer', '', markedBy || ''];
    if (rowNum) {
      sh.getRange(rowNum, 1, 1, 7).setValues([rowData]);
    } else {
      sh.appendRow(rowData);
    }
  });

  SpreadsheetApp.flush();
  return { saved: records.length, date: date };
}

/*** ---------------------------- QR SESSIONS ---------------------------- ***/

function createSession_(date, lat, lng) {
  if (!date) throw new Error('date is required');
  const sh = SS().getSheetByName(SESSIONS_SHEET);

  // Deactivate any previous active sessions for this date.
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (formatDate_(values[i][1]) === date && values[i][5] === true) {
      sh.getRange(i + 1, 6).setValue(false);
    }
  }

  const code = generateCode_();
  sh.appendRow([code, date, lat || '', lng || '', new Date(), true]);
  SpreadsheetApp.flush();
  return { sessionCode: code, date: date };
}

function getSessionInfo_(code) {
  if (!code) throw new Error('code is required');
  const sh = SS().getSheetByName(SESSIONS_SHEET);
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === code) {
      const createdAt = values[i][4];
      const ageMinutes = (new Date() - new Date(createdAt)) / 60000;
      return {
        found: true,
        date: formatDate_(values[i][1]),
        active: values[i][5] === true && ageMinutes <= SESSION_VALID_MINUTES,
        expired: ageMinutes > SESSION_VALID_MINUTES
      };
    }
  }
  return { found: false };
}

function studentCheckin_(body) {
  const { sessionCode, studentNumber, surname, initials, lat, lng, deviceId } = body;
  if (!sessionCode || !studentNumber || !surname || !initials || !deviceId) {
    return { success: false, message: 'Missing required fields.' };
  }

  // 1. Validate session
  const sessSheet = SS().getSheetByName(SESSIONS_SHEET);
  const sessValues = sessSheet.getDataRange().getValues();
  let session = null;
  for (let i = 1; i < sessValues.length; i++) {
    if (sessValues[i][0] === sessionCode) {
      session = {
        date: formatDate_(sessValues[i][1]),
        lat: parseFloat(sessValues[i][2]),
        lng: parseFloat(sessValues[i][3]),
        createdAt: sessValues[i][4],
        active: sessValues[i][5] === true
      };
      break;
    }
  }
  if (!session) return { success: false, message: 'Invalid QR code / session not found.' };
  const ageMinutes = (new Date() - new Date(session.createdAt)) / 60000;
  if (!session.active || ageMinutes > SESSION_VALID_MINUTES) {
    return { success: false, message: 'This QR code has expired. Ask your lecturer to generate a new one.' };
  }

  // 2. Validate proximity ("same room" check)
  if (lat === undefined || lng === undefined || lat === null || lng === null ||
      isNaN(session.lat) || isNaN(session.lng)) {
    return { success: false, message: 'Could not verify your location. Please enable location services and try again.' };
  }
  const distance = haversineMeters_(session.lat, session.lng, lat, lng);
  if (distance > MAX_CHECKIN_DISTANCE_METERS) {
    return { success: false, message: 'You do not appear to be in the same room as the lecturer. Move closer and try again.' };
  }

  // 3. Validate student against roster (student number + surname + initials must match)
  const students = getStudents_();
  const match = students.find(s => s.studentNumber === String(studentNumber).trim());
  if (!match) {
    return { success: false, message: 'Student number not found on the class list.' };
  }
  if (normalize_(match.surname) !== normalize_(surname) || normalize_(match.initials) !== normalize_(initials)) {
    return { success: false, message: 'Surname/Initials do not match our records for that student number.' };
  }

  // 4. One phone can only be used once per day (regardless of which student)
  const attSheet = SS().getSheetByName(ATTENDANCE_SHEET);
  const attValues = attSheet.getDataRange().getValues();
  for (let i = 1; i < attValues.length; i++) {
    if (formatDate_(attValues[i][0]) === session.date &&
        attValues[i][4] === 'QR' &&
        attValues[i][5] === deviceId) {
      return { success: false, message: 'This device has already been used to mark attendance today.' };
    }
  }

  // 5. Prevent marking the same student twice
  for (let i = 1; i < attValues.length; i++) {
    if (formatDate_(attValues[i][0]) === session.date &&
        String(attValues[i][1]) === match.studentNumber) {
      return { success: false, message: match.firstName + ' ' + match.surname + ' has already been marked present today.' };
    }
  }

  attSheet.appendRow([session.date, match.studentNumber, 'Present', new Date(), 'QR', deviceId, '']);
  SpreadsheetApp.flush();
  return { success: true, message: 'Attendance confirmed for ' + match.firstName + ' ' + match.surname + '.' };
}

/*** ---------------------------- SUMMARY ---------------------------- ***/

function getSummary_(start, end) {
  if (!start || !end) throw new Error('start and end dates are required');
  const students = getStudents_();
  const sh = SS().getSheetByName(ATTENDANCE_SHEET);
  const values = sh.getDataRange().getValues();
  values.shift();

  const inRange = values.filter(r => {
    const d = formatDate_(r[0]);
    return d >= start && d <= end;
  });

  const sessionDates = Array.from(new Set(inRange.map(r => formatDate_(r[0])))).sort();
  const totalSessions = sessionDates.length;

  const presentCount = {};
  inRange.forEach(r => {
    if (r[2] === 'Present') {
      const num = String(r[1]);
      presentCount[num] = (presentCount[num] || 0) + 1;
    }
  });

  const rows = students.map(s => {
    const present = presentCount[s.studentNumber] || 0;
    const pct = totalSessions > 0 ? Math.round((present / totalSessions) * 1000) / 10 : 0;
    return {
      studentNumber: s.studentNumber,
      surname: s.surname,
      initials: s.initials,
      firstName: s.firstName,
      sessionsPresent: present,
      totalSessions: totalSessions,
      percentage: pct
    };
  });

  return { start, end, totalSessions, sessionDates, rows };
}

/*** ---------------------------- HELPERS ---------------------------- ***/

function formatDate_(d) {
  if (typeof d === 'string') return d.substring(0, 10);
  return Utilities.formatDate(new Date(d), Session.getScriptTimeZone() || 'Africa/Johannesburg', 'yyyy-MM-dd');
}

function normalize_(s) {
  return String(s || '').trim().toUpperCase().replace(/\s+/g, '');
}

function generateCode_() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let code = '';
  for (let i = 0; i < 6; i++) code += chars.charAt(Math.floor(Math.random() * chars.length));
  return code;
}

function haversineMeters_(lat1, lon1, lat2, lon2) {
  const R = 6371000;
  const toRad = x => (x * Math.PI) / 180;
  const dLat = toRad(lat2 - lat1);
  const dLon = toRad(lon2 - lon1);
  const a =
    Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLon / 2) * Math.sin(dLon / 2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
}
