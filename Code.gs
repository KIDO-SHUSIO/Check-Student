const SESSION_TTL_SECONDS = 21600;
const LOGIN_WINDOW_SECONDS = 900;
const LOGIN_ATTEMPT_LIMIT = 8;
const LOGIN_FAILURES_KEY = "attendance-login-failures";
const LOGIN_LOCKOUT_KEY = "attendance-login-lockout";
const TEACHER_SHEET_NAME = "Teacher";
const ATTENDANCE_SHEET_NAME = "Attendance";

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("ระบบเช็คชื่อ")
    .addItem("ตรวจสอบโครงสร้างฐานข้อมูล", "checkDatabaseStructure")
    .addItem("ตั้งคอลัมน์ห้องใน Attendance เป็นข้อความ", "formatAttendanceRoomColumnAsText")
    .addItem("ปรับสถานะเดิมเป็น มา / สาย / ลา / ขาด", "normalizeAttendanceStatuses")
    .addToUi();
}

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const session = getSession_(params.token);
  if (!session || session.role !== "scanner") {
    return createJsonResponse_(false, "กรุณาเข้าสู่ระบบ", "AUTH_REQUIRED");
  }
  if (!params.id) {
    return createJsonResponse_(false, "ไม่พบรหัสนักเรียนที่ส่งมา");
  }
  return recordAttendance_(params.id, session, params.scannedAt);
}

function doPost(e) {
  const params = e && e.parameter ? e.parameter : {};
  const action = String(params.action || "");

  if (action === "login") {
    return login_(String(params.teacherId || ""));
  }
  if (action === "logout") {
    if (getSession_(params.token)) {
      CacheService.getScriptCache().remove(sessionCacheKey_(params.token));
    }
    return createJsonResponse_(true, "ออกจากระบบแล้ว");
  }

  const session = getSession_(params.token);
  if (!session) {
    return createJsonResponse_(false, "เซสชันหมดอายุ กรุณาเข้าสู่ระบบอีกครั้ง", "AUTH_REQUIRED");
  }

  if (action === "scan") {
    if (session.role !== "scanner") {
      return createJsonResponse_(false, "บัญชีนี้ไม่มีสิทธิ์ใช้หน้าสแกน", "FORBIDDEN");
    }
    return recordAttendance_(params.id, session, params.scannedAt);
  }
  if (action === "teacher-attendance") {
    if (session.role !== "teacher") {
      return createJsonResponse_(false, "บัญชีนี้ไม่มีสิทธิ์ดูรายชื่อครู", "FORBIDDEN");
    }
    return getTeacherAttendance_(session, String(params.date || ""));
  }
  if (action === "teacher-attendance-counts") {
    if (session.role !== "teacher") {
      return createJsonResponse_(false, "บัญชีนี้ไม่มีสิทธิ์ดูรายงานเช็คชื่อ", "FORBIDDEN");
    }
    return getTeacherAttendanceCounts_(session, String(params.startDate || ""), String(params.endDate || ""));
  }
  if (action === "teacher-qr-codes") {
    if (session.role !== "teacher") {
      return createJsonResponse_(false, "บัญชีนี้ไม่มีสิทธิ์ดาวน์โหลด QR Code", "FORBIDDEN");
    }
    return getTeacherQrCodes_(session);
  }
  if (action === "set-status") {
    if (session.role !== "teacher") {
      return createJsonResponse_(false, "บัญชีนี้ไม่มีสิทธิ์จัดการสถานะนักเรียน", "FORBIDDEN");
    }
    return setStudentStatus_(
      session,
      String(params.date || ""),
      String(params.studentId || ""),
      String(params.status || "")
    );
  }
  if (action === "export-attendance") {
    if (session.role !== "teacher") {
      return createJsonResponse_(false, "บัญชีนี้ไม่มีสิทธิ์ส่งออกรายงาน", "FORBIDDEN");
    }
    return exportAttendance_(session, String(params.date || ""));
  }
  if (action === "upload-roster-start") {
    if (session.role !== "admin") {
      return createJsonResponse_(false, "เฉพาะผู้ดูแลระบบเท่านั้นที่อัปโหลดรายชื่อได้", "FORBIDDEN");
    }
    return startRosterUpload_(session, String(params.room || ""), String(params.uploadId || ""), String(params.rows || ""));
  }
  if (action === "upload-roster-images") {
    if (session.role !== "admin") {
      return createJsonResponse_(false, "เฉพาะผู้ดูแลระบบเท่านั้นที่อัปโหลด QR Code ได้", "FORBIDDEN");
    }
    return addRosterQrImages_(session, String(params.uploadId || ""), String(params.studentIds || ""));
  }
  if (action === "upload-roster-finalize") {
    if (session.role !== "admin") {
      return createJsonResponse_(false, "เฉพาะผู้ดูแลระบบเท่านั้นที่บันทึกรายชื่อได้", "FORBIDDEN");
    }
    return finalizeRosterUpload_(session, String(params.room || ""), String(params.uploadId || ""));
  }
  if (action === "upload-roster-cancel") {
    if (session.role !== "admin") {
      return createJsonResponse_(false, "เฉพาะผู้ดูแลระบบเท่านั้นที่ยกเลิกการอัปโหลดได้", "FORBIDDEN");
    }
    return cancelRosterUpload_(session, String(params.uploadId || ""));
  }
  if (action === "admin-roster-list") {
    if (session.role !== "admin") {
      return createJsonResponse_(false, "เฉพาะผู้ดูแลระบบเท่านั้นที่จัดการรายชื่อได้", "FORBIDDEN");
    }
    return getAdminRoster_(String(params.room || ""));
  }
  if (action === "admin-attendance-report") {
    if (session.role !== "admin") {
      return createJsonResponse_(false, "เฉพาะผู้ดูแลระบบเท่านั้นที่ดูรายงานเช็คชื่อได้", "FORBIDDEN");
    }
    return getAdminAttendanceCounts_(String(params.room || ""), String(params.startDate || ""), String(params.endDate || ""));
  }
  if (action === "admin-student-save") {
    if (session.role !== "admin") {
      return createJsonResponse_(false, "เฉพาะผู้ดูแลระบบเท่านั้นที่จัดการรายชื่อได้", "FORBIDDEN");
    }
    return saveAdminStudent_(
      String(params.sourceRoom || ""),
      String(params.id || ""),
      String(params.room || ""),
      String(params.studentId || ""),
      String(params.number || ""),
      String(params.name || "")
    );
  }
  if (action === "admin-student-delete") {
    if (session.role !== "admin") {
      return createJsonResponse_(false, "เฉพาะผู้ดูแลระบบเท่านั้นที่จัดการรายชื่อได้", "FORBIDDEN");
    }
    return deleteAdminStudent_(String(params.room || ""), String(params.id || ""));
  }
  if (action === "admin-room-create") {
    if (session.role !== "admin") {
      return createJsonResponse_(false, "เฉพาะผู้ดูแลระบบเท่านั้นที่จัดการห้องได้", "FORBIDDEN");
    }
    return createAdminRoom_(String(params.grade || ""), String(params.classroom || ""));
  }
  if (action === "admin-room-delete") {
    if (session.role !== "admin") {
      return createJsonResponse_(false, "เฉพาะผู้ดูแลระบบเท่านั้นที่จัดการห้องได้", "FORBIDDEN");
    }
    return deleteAdminRoom_(String(params.room || ""));
  }
  if (action === "admin-teacher-list") {
    if (session.role !== "admin") {
      return createJsonResponse_(false, "เฉพาะผู้ดูแลระบบเท่านั้นที่จัดการครูได้", "FORBIDDEN");
    }
    return getAdminTeachers_();
  }
  if (action === "admin-teacher-assign") {
    if (session.role !== "admin") {
      return createJsonResponse_(false, "เฉพาะผู้ดูแลระบบเท่านั้นที่จัดการครูได้", "FORBIDDEN");
    }
    return assignTeacherRoom_(String(params.teacherId || ""), String(params.room || ""));
  }
  if (action === "admin-teachers-assign") {
    if (session.role !== "admin") {
      return createJsonResponse_(false, "เฉพาะผู้ดูแลระบบเท่านั้นที่จัดการครูได้", "FORBIDDEN");
    }
    return assignTeachersToRoom_(String(params.room || ""), String(params.teacherIds || ""));
  }
  return createJsonResponse_(false, "ไม่พบคำสั่งที่รองรับ", "INVALID_ACTION");
}

function checkDatabaseStructure() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const missing = [];
  if (!spreadsheet.getSheetByName(TEACHER_SHEET_NAME)) missing.push(TEACHER_SHEET_NAME);
  if (!spreadsheet.getSheetByName(ATTENDANCE_SHEET_NAME)) missing.push(ATTENDANCE_SHEET_NAME);
  const classSheets = spreadsheet.getSheets()
    .map(sheet => sheet.getName())
    .filter(name => /^ม\./.test(name));
  if (!classSheets.length) missing.push("ชีตห้องเรียน (เช่น ม.610)");

  const message = missing.length
    ? "ยังขาด: " + missing.join(", ")
    : "พบชีต Teacher, Attendance และชีตห้องเรียน: " + classSheets.join(", ");
  SpreadsheetApp.getUi().alert(message);
}

function formatAttendanceRoomColumnAsText() {
  const attendanceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ATTENDANCE_SHEET_NAME);
  if (!attendanceSheet) throw new Error("ไม่พบชีต Attendance");
  attendanceSheet.getRange("E:E").setNumberFormat("@");
  SpreadsheetApp.getUi().alert(
    "ตั้งคอลัมน์ E (ชั้น/ห้อง) ใน Attendance เป็นข้อความแล้ว\nข้อมูลเดิมที่ถูกแปลงเป็นวันที่ต้องแก้ไขย้อนหลัง"
  );
}

function normalizeAttendanceStatuses() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ATTENDANCE_SHEET_NAME);
  if (!sheet) throw new Error("ไม่พบชีต Attendance");

  const lastRow = sheet.getLastRow();
  sheet.getRange(1, 6).setValue("สถานะการมาโรงเรียน");
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("ปรับหัวคอลัมน์เป็น สถานะการมาโรงเรียน แล้ว ไม่มีข้อมูลเดิมให้แปลง");
    return;
  }

  const statusRange = sheet.getRange(2, 6, lastRow - 1, 1);
  const statuses = statusRange.getValues();
  let updated = 0;
  for (const row of statuses) {
    const value = String(row[0] == null ? "" : row[0]).trim();
    let normalized = value;
    if (/ปกติ/.test(value)) normalized = "มา";
    else if (/สาย/.test(value)) normalized = "สาย";
    else if (/ลา/.test(value)) normalized = "ลา";
    else if (/ขาด/.test(value)) normalized = "ขาด";
    if (normalized !== value) {
      row[0] = normalized;
      updated += 1;
    }
  }
  if (updated) statusRange.setValues(statuses);
  SpreadsheetApp.getUi().alert(
    "ปรับหัวคอลัมน์แล้ว และแปลงสถานะเดิม " + updated + " แถวเป็น มา / สาย / ลา / ขาด"
  );
}

function login_(teacherId) {
  const identifier = String(teacherId || "").trim();
  const properties = PropertiesService.getScriptProperties();
  const scannerAccessCode = String(properties.getProperty("SCANNER_ACCESS_CODE") || "").trim();
  const adminAccessCode = String(properties.getProperty("ADMIN_ACCESS_CODE") || "").trim();
  if (scannerAccessCode && identifier === scannerAccessCode) {
    return createRoleSession_("scanner", { id: identifier }, "เข้าสู่ระบบหน้าสแกนแล้ว", "index.html");
  }
  if (adminAccessCode && identifier === adminAccessCode) {
    return createRoleSession_("admin", { id: identifier }, "เข้าสู่ระบบผู้ดูแลแล้ว", "upload.html");
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const teacherSheet = spreadsheet.getSheetByName(TEACHER_SHEET_NAME);
  if (!teacherSheet || teacherSheet.getLastRow() < 2) {
    return createJsonResponse_(false, "ไม่พบข้อมูลครูในชีต Teacher", "TEACHERS_NOT_CONFIGURED");
  }

  const cache = CacheService.getScriptCache();
  if (cache.get(LOGIN_LOCKOUT_KEY)) {
    return createJsonResponse_(false, "มีการลองเข้าสู่ระบบผิดหลายครั้ง กรุณารอ 15 นาที", "RATE_LIMITED");
  }

  const normalizedId = normalizeId_(identifier);
  const rows = teacherSheet.getRange(2, 1, teacherSheet.getLastRow() - 1, 3).getDisplayValues();
  const teacherRow = rows.find(row => normalizeId_(row[0]) === normalizedId);
  if (!normalizedId || !teacherRow) {
    if (!scannerAccessCode || !adminAccessCode) {
      return createJsonResponse_(
        false,
        "ยังไม่ได้ตั้งค่ารหัสเข้าใช้งานสำหรับเครื่องสแกนและผู้ดูแลใน Script Properties",
        "ACCESS_CODES_NOT_CONFIGURED"
      );
    }
    registerLoginFailure_();
    return createJsonResponse_(false, "ไม่พบเลขประจำตัวครูในชีต Teacher");
  }

  const room = String(teacherRow[2] || "").trim();
  const studentSheet = room ? findClassSheet_(spreadsheet, room) : null;
  if (room && !studentSheet) {
    return createJsonResponse_(
      false,
      "ไม่พบชีตห้องเรียนที่ตรงกับห้องของครู (" + room + ")",
      "ROOM_SHEET_NOT_FOUND"
    );
  }

  const teacher = {
    role: "teacher",
    id: normalizedId,
    name: String(teacherRow[1] || "").trim(),
    room: studentSheet ? room : "",
    sheet: studentSheet ? studentSheet.getName() : ""
  };
  return createRoleSession_(
    "teacher",
    teacher,
    teacher.room ? "เข้าสู่ระบบครูแล้ว" : "เข้าสู่ระบบแล้ว แต่ยังไม่ได้กำหนดห้องเรียน",
    "teacher.html"
  );
}

function createRoleSession_(role, identity, message, destination) {
  const cache = CacheService.getScriptCache();
  cache.removeAll([LOGIN_FAILURES_KEY, LOGIN_LOCKOUT_KEY]);
  const token = Utilities.getUuid().replace(/-/g, "");
  const session = Object.assign({}, identity, { role: role });
  cache.put(sessionCacheKey_(token), JSON.stringify(session), SESSION_TTL_SECONDS);
  return createJsonResponse_(true, message, null, {
    token: token,
    role: role,
    destination: destination,
    expiresIn: SESSION_TTL_SECONDS
  });
}

function registerLoginFailure_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const cache = CacheService.getScriptCache();
    const attempts = Number(cache.get(LOGIN_FAILURES_KEY) || 0) + 1;
    if (attempts >= LOGIN_ATTEMPT_LIMIT) {
      cache.remove(LOGIN_FAILURES_KEY);
      cache.put(LOGIN_LOCKOUT_KEY, "1", LOGIN_WINDOW_SECONDS);
      return;
    }
    cache.put(LOGIN_FAILURES_KEY, String(attempts), LOGIN_WINDOW_SECONDS);
  } finally {
    lock.releaseLock();
  }
}

function getSession_(token) {
  if (!/^[a-f0-9]{32}$/i.test(String(token || ""))) return null;
  const value = CacheService.getScriptCache().get(sessionCacheKey_(token));
  if (!value) return null;
  try {
    const teacher = JSON.parse(value);
    if (!teacher || typeof teacher.id !== "string" || !["scanner", "teacher", "admin"].includes(teacher.role)) return null;
    if (teacher.role === "teacher" && (typeof teacher.sheet !== "string" || typeof teacher.room !== "string")) return null;
    if (teacher.role === "teacher") {
      const currentTeacher = findTeacherById_(teacher.id);
      if (!currentTeacher) return null;
      const assignedRoom = String(currentTeacher[2] || "").trim();
      const assignedSheet = assignedRoom ? findClassSheet_(SpreadsheetApp.getActiveSpreadsheet(), assignedRoom) : null;
      teacher.name = String(currentTeacher[1] || "").trim();
      teacher.room = assignedSheet ? assignedRoom : "";
      teacher.sheet = assignedSheet ? assignedSheet.getName() : "";
    }
    return teacher;
  } catch (error) {
    console.error("Invalid login session data:", error);
    return null;
  }
}

function sessionCacheKey_(token) {
  return "attendance-session:" + token;
}

function findTeacherById_(teacherId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TEACHER_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const teachers = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getDisplayValues();
  return teachers.find(row => normalizeId_(row[0]) === normalizeId_(teacherId)) || null;
}

function normalizeId_(value) {
  const id = String(value == null ? "" : value).trim();
  return /^\d+\.0+$/.test(id) ? id.replace(/\.0+$/, "") : id;
}

function normalizeRoomKey_(value) {
  return String(value || "").replace(/\D/g, "");
}

function classSheetName_(room) {
  const value = String(room || "").trim().replace(/^ม\.\s*/i, "");
  return "ม." + normalizeRoomKey_(value);
}

function findClassSheet_(spreadsheet, room) {
  const roomKey = normalizeRoomKey_(room);
  if (!roomKey) return null;
  return spreadsheet.getSheets().find(sheet =>
    /^ม\./.test(sheet.getName()) && normalizeRoomKey_(sheet.getName()) === roomKey
  ) || null;
}

function parseDateKey_(value) {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(value || ""));
  if (!match) return null;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(year, month - 1, day);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) return null;
  return String(value);
}

function dateKey_(value) {
  if (!(value instanceof Date) || isNaN(value.getTime())) return "";
  return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function parseScanTimestamp_(value) {
  if (!value) return new Date();
  const timestamp = String(value);
  if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(timestamp)) return null;
  const date = new Date(timestamp);
  if (isNaN(date.getTime()) || date.getTime() > Date.now() + 5 * 60 * 1000) return null;
  return date;
}

function isScannedAttendance_(row) {
  return row[0] instanceof Date
    && Utilities.formatDate(row[0], Session.getScriptTimeZone(), "HH:mm:ss") !== "00:00:00";
}

function isTeacherMarkedAttendance_(row) {
  return !isScannedAttendance_(row) && Boolean(attendanceStatus_(row[5]));
}

function scanAttendanceStatus_(scanTime) {
  const currentTime = Utilities.formatDate(scanTime, Session.getScriptTimeZone(), "HH:mm");
  if (currentTime < "08:20") return "มา";
  if (currentTime <= "09:20") return "สาย";
  return "ขาด";
}

function getSessionRoomSheet_(session) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = findClassSheet_(spreadsheet, session.room);
  if (!sheet || sheet.getName() !== session.sheet) {
    throw new Error("ไม่พบชีตห้องเรียนที่กำหนดให้บัญชีนี้");
  }
  return sheet;
}

function getAttendanceRows_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ATTENDANCE_SHEET_NAME);
  if (!sheet) throw new Error("ไม่พบชีต Attendance");
  const lastRow = sheet.getLastRow();
  return { sheet: sheet, rows: lastRow < 2 ? [] : sheet.getRange(2, 1, lastRow - 1, 6).getValues() };
}

function getStatusForDate_(rows, studentId, roomKey, selectedDate) {
  let status = "";
  for (const row of rows) {
    if (dateKey_(row[0]) !== selectedDate) continue;
    if (normalizeId_(row[1]) !== studentId) continue;
    if (normalizeRoomKey_(row[4]) !== roomKey) continue;
    status = attendanceStatus_(row[5]) || status;
  }
  return status;
}

function attendanceStatus_(value) {
  const status = String(value || "");
  if (/สาย/.test(status)) return "สาย";
  if (/ลา/.test(status)) return "ลา";
  if (/ขาด/.test(status)) return "ขาด";
  if (/ปกติ|มา/.test(status)) return "มา";
  return "";
}

function getAttendanceCounts_(studentSheet, room, startDate, endDate) {
  const countsById = Object.create(null);
  const attendance = getAttendanceRows_();
  const roomKey = normalizeRoomKey_(room);
  for (const row of attendance.rows) {
    const date = dateKey_(row[0]);
    if (date < startDate || date > endDate || normalizeRoomKey_(row[4]) !== roomKey) continue;
    const id = normalizeId_(row[1]);
    const status = attendanceStatus_(row[5]);
    if (!id || !status) continue;
    if (!countsById[id]) countsById[id] = { มา: 0, สาย: 0, ลา: 0, ขาด: 0 };
    countsById[id][status] += 1;
  }

  const values = studentSheet.getDataRange().getDisplayValues();
  const students = [];
  for (let index = 1; index < values.length; index += 1) {
    const row = values[index];
    const id = normalizeId_(row[0]);
    if (!id) continue;
    students.push({
      id: id,
      number: String(row[1] || "").trim(),
      name: String(row[2] || "").trim(),
      counts: countsById[id] || { มา: 0, สาย: 0, ลา: 0, ขาด: 0 }
    });
  }
  return students;
}

function validAttendanceDateRange_(requestedStartDate, requestedEndDate) {
  const startDate = parseDateKey_(requestedStartDate);
  const endDate = parseDateKey_(requestedEndDate);
  if (!startDate || !endDate) return null;
  if (startDate > endDate) return null;
  return { startDate: startDate, endDate: endDate };
}

function getTeacherAttendanceCounts_(session, requestedStartDate, requestedEndDate) {
  const range = validAttendanceDateRange_(requestedStartDate, requestedEndDate);
  if (!range) return createJsonResponse_(false, "กรุณาระบุช่วงวันที่ให้ถูกต้อง โดยวันเริ่มต้นต้องไม่อยู่หลังวันสิ้นสุด");
  if (!session.room) {
    return createJsonResponse_(false, "บัญชีครูนี้ยังไม่ได้รับมอบหมายห้องเรียน");
  }
  const students = getAttendanceCounts_(getSessionRoomSheet_(session), session.room, range.startDate, range.endDate);
  return createJsonResponse_(true, "โหลดสรุปจำนวนครั้งการเข้าเรียนแล้ว", null, {
    room: session.room,
    startDate: range.startDate,
    endDate: range.endDate,
    students: students
  });
}

function getTeacherAttendance_(session, requestedDate) {
  const selectedDate = parseDateKey_(requestedDate);
  if (!selectedDate) return createJsonResponse_(false, "รูปแบบวันที่ไม่ถูกต้อง");
  if (!session.room) {
    return createJsonResponse_(true, "บัญชีครูนี้ยังไม่ได้รับมอบหมายห้องเรียน", null, {
      teacher: { id: session.id, name: session.name },
      room: "",
      date: selectedDate,
      students: []
    });
  }

  const studentSheet = getSessionRoomSheet_(session);
  const attendance = getAttendanceRows_();
  const roomKey = normalizeRoomKey_(session.room);
  const values = studentSheet.getDataRange().getDisplayValues();
  const students = [];
  for (let index = 1; index < values.length; index += 1) {
    const row = values[index];
    const id = normalizeId_(row[0]);
    if (!id) continue;
    students.push({
      id: id,
      number: String(row[1] || "").trim(),
      name: String(row[2] || "").trim(),
      status: getStatusForDate_(attendance.rows, id, roomKey, selectedDate)
    });
  }

  return createJsonResponse_(true, "โหลดรายชื่อนักเรียนแล้ว", null, {
    teacher: { id: session.id, name: session.name },
    room: session.room,
    date: selectedDate,
    students: students
  });
}

function getTeacherQrCodes_(session) {
  if (!session.room) {
    return createJsonResponse_(true, "บัญชีครูนี้ยังไม่ได้รับมอบหมายห้องเรียน", null, {
      room: "",
      students: []
    });
  }

  const studentSheet = getSessionRoomSheet_(session);
  const values = studentSheet.getDataRange().getDisplayValues();
  const students = [];
  for (let index = 1; index < values.length; index += 1) {
    const row = values[index];
    const id = normalizeId_(row[0]);
    if (!id) continue;
    students.push({
      id: id,
      number: String(row[1] || "").trim(),
      name: String(row[2] || "").trim(),
      qr: Utilities.base64Encode(createStudentQrPng_(id, 16))
    });
  }

  return createJsonResponse_(true, "โหลด QR Code นักเรียนแล้ว", null, {
    room: session.room,
    students: students
  });
}

function setStudentStatus_(session, requestedDate, studentId, status) {
  const selectedDate = parseDateKey_(requestedDate);
  if (!selectedDate) return createJsonResponse_(false, "รูปแบบวันที่ไม่ถูกต้อง");
  if (!session.room) return createJsonResponse_(false, "บัญชีครูนี้ยังไม่ได้รับมอบหมายห้องเรียน");
  if (!["มา", "สาย", "ลา", "ขาด"].includes(status)) {
    return createJsonResponse_(false, "สถานะที่เลือกไม่ถูกต้อง");
  }

  const id = normalizeId_(studentId);
  const studentSheet = getSessionRoomSheet_(session);
  const students = studentSheet.getDataRange().getDisplayValues();
  const student = students.slice(1).find(row => normalizeId_(row[0]) === id);
  if (!student) return createJsonResponse_(false, "ไม่พบนักเรียนในห้องที่รับผิดชอบ");

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const attendance = getAttendanceRows_();
    const roomKey = normalizeRoomKey_(session.room);
    let matchingRow = 0;
    for (let index = 0; index < attendance.rows.length; index += 1) {
      const row = attendance.rows[index];
      if (dateKey_(row[0]) === selectedDate
          && normalizeId_(row[1]) === id
          && normalizeRoomKey_(row[4]) === roomKey) {
        matchingRow = index + 2;
      }
    }

    if (matchingRow) {
      attendance.sheet.getRange(matchingRow, 6).setValue(status);
    } else {
      const date = Utilities.parseDate(selectedDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      const rowNumber = attendance.sheet.getLastRow() + 1;
      attendance.sheet.getRange(rowNumber, 5).setNumberFormat("@");
      attendance.sheet.getRange(rowNumber, 1, 1, 6).setValues([[
        date,
        id,
        student[1],
        student[2],
        String(session.room),
        status
      ]]);
    }
  } finally {
    lock.releaseLock();
  }

  return createJsonResponse_(true, "บันทึกสถานะแล้ว", null, { id: id, status: status });
}

function recordAttendance_(studentId, session, scannedAt) {
  const id = normalizeId_(studentId);
  if (!id) return createJsonResponse_(false, "ไม่พบรหัสนักเรียนที่ส่งมา");
  const scanTime = parseScanTimestamp_(scannedAt);
  if (!scanTime) return createJsonResponse_(false, "วันเวลาสแกนไม่ถูกต้อง");

  let studentSheet = null;
  let student = null;
  const matches = [];
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const classSheets = session.role === "teacher"
    ? [getSessionRoomSheet_(session)]
    : spreadsheet.getSheets().filter(sheet => /^ม\./.test(sheet.getName()));
  for (const sheet of classSheets) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) continue;
    const match = sheet.getRange(2, 1, lastRow - 1, 1)
      .createTextFinder(id)
      .matchEntireCell(true)
      .findNext();
    if (!match) continue;
    const row = sheet.getRange(match.getRow(), 1, 1, 3).getDisplayValues()[0];
    if (normalizeId_(row[0]) === id) matches.push({ sheet: sheet, student: row });
  }
  if (matches.length > 1) {
    return createJsonResponse_(false, "พบรหัสนักเรียนซ้ำในหลายห้อง กรุณาแจ้งผู้ดูแลระบบ", "AMBIGUOUS_STUDENT");
  }
  if (matches.length === 1) {
    studentSheet = matches[0].sheet;
    student = matches[0].student;
  }
  if (!student || !studentSheet) {
    return createJsonResponse_(false, "ไม่พบรหัสนักเรียนในรายชื่อห้องเรียน");
  }

  const room = session.role === "teacher" ? session.room : getRoomLabelForSheet_(spreadsheet, studentSheet);
  const roomKey = normalizeRoomKey_(room);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const attendanceSheet = spreadsheet.getSheetByName(ATTENDANCE_SHEET_NAME);
    if (!attendanceSheet) throw new Error("ไม่พบชีต Attendance");
    const lastAttendanceRow = attendanceSheet.getLastRow();
    const today = dateKey_(scanTime);
    let matchingRow = 0;
    let alreadyScanned = false;
    let teacherMarkedStatus = "";
    if (lastAttendanceRow >= 2) {
      const matchingAttendanceRows = attendanceSheet.getRange(2, 2, lastAttendanceRow - 1, 1)
        .createTextFinder(id)
        .matchEntireCell(true)
        .findAll();
      for (const match of matchingAttendanceRows) {
        const rowNumber = match.getRow();
        const row = attendanceSheet.getRange(rowNumber, 1, 1, 6).getValues()[0];
        if (dateKey_(row[0]) !== today
            || normalizeId_(row[1]) !== id
            || normalizeRoomKey_(row[4]) !== roomKey) continue;
        matchingRow = rowNumber;
        if (isScannedAttendance_(row)) alreadyScanned = true;
        else if (isTeacherMarkedAttendance_(row)) teacherMarkedStatus = attendanceStatus_(row[5]);
      }
    }
    if (alreadyScanned) {
      return createJsonResponse_(false, "รายการซ้ำ: นักเรียนคนนี้เช็คชื่อแล้ววันนี้", "DUPLICATE", { duplicate: true });
    }
    if (teacherMarkedStatus) {
      return createJsonResponse_(
        false,
        "ครูบันทึกสถานะเป็น " + teacherMarkedStatus + " แล้ว ระบบจะไม่เปลี่ยนสถานะจากการสแกน",
        "DUPLICATE",
        { duplicate: true, status: teacherMarkedStatus }
      );
    }

    const status = scanAttendanceStatus_(scanTime);
    const rowNumber = matchingRow || attendanceSheet.getLastRow() + 1;
    attendanceSheet.getRange(rowNumber, 5).setNumberFormat("@");
    attendanceSheet.getRange(rowNumber, 1, 1, 6).setValues([[
      scanTime,
      id,
      student[1],
      student[2],
      String(room),
      status
    ]]);
    return createJsonResponse_(
      true,
      "✅ บันทึกสำเร็จ: เลขที่ " + student[1] + " " + student[2] + " (" + status + ")"
    );
  } finally {
    lock.releaseLock();
  }
}

function getRoomLabelForSheet_(spreadsheet, studentSheet) {
  const teacherSheet = spreadsheet.getSheetByName(TEACHER_SHEET_NAME);
  if (teacherSheet && teacherSheet.getLastRow() >= 2) {
    const teachers = teacherSheet.getRange(2, 1, teacherSheet.getLastRow() - 1, 3).getDisplayValues();
    const matchingTeacher = teachers.find(row =>
      normalizeRoomKey_(row[2]) === normalizeRoomKey_(studentSheet.getName())
    );
    if (matchingTeacher && String(matchingTeacher[2] || "").trim()) {
      return String(matchingTeacher[2]).trim();
    }
  }
  return studentSheet.getName().replace(/^ม\./, "");
}

function parseRosterRows_(rawRows) {
  if (!rawRows || rawRows.length > 500000) {
    return { error: "ข้อมูลไฟล์ว่างหรือมีขนาดใหญ่เกินไป" };
  }
  let rows;
  try {
    rows = JSON.parse(rawRows);
  } catch (error) {
    return { error: "อ่านข้อมูลรายชื่อที่ส่งมาไม่สำเร็จ" };
  }
  if (!Array.isArray(rows) || rows.length > 2000) {
    return { error: "รูปแบบรายชื่อไม่ถูกต้อง หรือมีนักเรียนเกิน 2,000 คน" };
  }

  const roster = [];
  const seenIds = new Set();
  for (const row of rows) {
    if (!Array.isArray(row)) return { error: "รูปแบบแถวรายชื่อไม่ถูกต้อง" };
    const values = Array.from({ length: 4 }, (_, index) => String(row[index] == null ? "" : row[index]).trim());
    const firstCell = values[0].toLowerCase();
    if (values.every(value => !value)) continue;
    if (/รหัส|เลข|student\s*id|^id$/.test(firstCell) && /ชื่อ|name/i.test(values[2])) continue;
    if (!values[0]) return { error: "พบแถวที่ไม่มีรหัสนักเรียน" };
    if (!values[2]) return { error: "พบแถวที่ไม่มีชื่อ-นามสกุล" };
    const id = normalizeId_(values[0]);
    try {
      validateQrStudentId_(id);
    } catch (error) {
      return { error: error.message };
    }
    if (seenIds.has(id)) return { error: "พบรหัสนักเรียนซ้ำในไฟล์: " + id };
    seenIds.add(id);
    roster.push([id, values[1], values[2], ""]);
  }
  if (!roster.length) return { error: "ไม่พบรายชื่อนักเรียนที่นำเข้าได้" };
  return { roster: roster };
}

function rosterStageName_(session, uploadId) {
  return "__QR_" + session.id + "_" + uploadId;
}

function getRosterStage_(session, uploadId) {
  if (!/^[a-f0-9]{32}$/i.test(uploadId)) return null;
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(rosterStageName_(session, uploadId));
}

function startRosterUpload_(session, requestedRoom, uploadId, rawRows) {
  const roomKey = normalizeRoomKey_(requestedRoom);
  if (!roomKey) return createJsonResponse_(false, "กรุณาระบุชั้น/ห้อง");
  if (!/^[a-f0-9]{32}$/i.test(uploadId)) return createJsonResponse_(false, "รหัสการอัปโหลดไม่ถูกต้อง");
  const parsed = parseRosterRows_(rawRows);
  if (parsed.error) return createJsonResponse_(false, parsed.error);
  const roster = parsed.roster;

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const stageName = rosterStageName_(session, uploadId);
  if (spreadsheet.getSheetByName(stageName)) {
    return createJsonResponse_(false, "รหัสอัปโหลดนี้ถูกใช้งานแล้ว กรุณาลองใหม่", "UPLOAD_ID_EXISTS");
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let stage;
  try {
    stage = spreadsheet.insertSheet(stageName);
    if (roster.length + 1 > stage.getMaxRows()) stage.insertRowsAfter(stage.getMaxRows(), roster.length + 1 - stage.getMaxRows());
    stage.getRange(1, 1, roster.length + 1, 4).setNumberFormat("@");
    stage.getRange(1, 1, 1, 4).setValues([["รหัสนักเรียน", "เลขที่", "ชื่อ-นามสกุล", "QR Code"]]);
    stage.getRange(2, 1, roster.length, 4).setValues(roster);
    stage.setColumnWidth(4, 120);
    stage.setRowHeights(2, roster.length, 110);
  } finally {
    lock.releaseLock();
  }

  return createJsonResponse_(true, "เตรียมรายชื่อแล้ว", null, {
    room: requestedRoom.trim(),
    uploadId: uploadId,
    count: roster.length
  });
}

function addRosterQrImages_(session, uploadId, rawStudentIds) {
  const stage = getRosterStage_(session, uploadId);
  if (!stage) return createJsonResponse_(false, "ไม่พบรายการอัปโหลดที่กำลังดำเนินการ", "UPLOAD_NOT_FOUND");
  if (!rawStudentIds || rawStudentIds.length > 10000) return createJsonResponse_(false, "ชุดรหัสนักเรียนไม่ถูกต้อง");
  let studentIds;
  try {
    studentIds = JSON.parse(rawStudentIds);
  } catch (error) {
    return createJsonResponse_(false, "อ่านชุดรหัสนักเรียนไม่สำเร็จ");
  }
  if (!Array.isArray(studentIds) || !studentIds.length || studentIds.length > 25
      || studentIds.some(id => typeof id !== "string")) {
    return createJsonResponse_(false, "ชุดรหัสนักเรียนไม่ถูกต้อง");
  }

  const idsByRow = stage.getRange(2, 1, stage.getLastRow() - 1, 1).getDisplayValues();
  const rowById = new Map(idsByRow.map((row, index) => [normalizeId_(row[0]), index + 2]));
  const occupiedRows = new Set(stage.getImages().map(image => image.getAnchorCell().getRow()));
  const requestedRows = new Set();
  const additions = [];
  for (const studentId of studentIds) {
    const id = normalizeId_(studentId);
    const row = rowById.get(id);
    if (!row) return createJsonResponse_(false, "ไม่พบรหัสนักเรียนในชุดอัปโหลด");
    if (requestedRows.has(row) || occupiedRows.has(row)) {
      return createJsonResponse_(false, "พบรหัสนักเรียนซ้ำในชุดอัปโหลด");
    }
    requestedRows.add(row);
    additions.push({ row: row, id: id });
  }

  for (const item of additions) {
    const blob = Utilities.newBlob(createStudentQrPng_(item.id), "image/png", "qr-" + item.id + ".png");
    const image = stage.insertImage(blob, 4, item.row);
    image.setWidth(100).setHeight(100);
    image.setAnchorCellXOffset(10).setAnchorCellYOffset(5);
  }
  return createJsonResponse_(true, "เพิ่มภาพ QR แล้ว", null, {
    count: additions.length,
    uploaded: stage.getImages().length
  });
}

function validateQrStudentId_(studentId) {
  const id = normalizeId_(studentId);
  const numeric = /^\d+$/.test(id);
  const alphanumeric = /^[0-9A-Z $%*+\-./:]+$/.test(id);
  if ((numeric && id.length > 41) || (!numeric && (!alphanumeric || id.length > 25))) {
    throw new Error("รหัสนักเรียน " + id + " ยาวเกินไปหรือมีอักขระที่ QR Code ไม่รองรับ");
  }
}

function appendQrBits_(bits, value, length) {
  for (let index = length - 1; index >= 0; index -= 1) {
    bits.push((value >>> index) & 1);
  }
}

function getQrErrorCorrectionLevel_(studentId) {
  const id = normalizeId_(studentId);
  const bitLength = /^\d+$/.test(id)
    ? 14 + Math.floor(id.length / 3) * 10 + (id.length % 3 === 2 ? 7 : id.length % 3 === 1 ? 4 : 0)
    : 13 + Math.floor(id.length / 2) * 11 + (id.length % 2 ? 6 : 0);
  return bitLength <= 9 * 8 ? "H" : "L";
}

function createQrDataCodewords_(studentId, errorCorrectionLevel) {
  const id = normalizeId_(studentId);
  validateQrStudentId_(id);
  const bits = [];
  if (/^\d+$/.test(id)) {
    appendQrBits_(bits, 1, 4);
    appendQrBits_(bits, id.length, 10);
    for (let index = 0; index < id.length; index += 3) {
      const group = id.slice(index, index + 3);
      appendQrBits_(bits, Number(group), group.length === 3 ? 10 : group.length === 2 ? 7 : 4);
    }
  } else {
    const charset = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ $%*+-./:";
    appendQrBits_(bits, 2, 4);
    appendQrBits_(bits, id.length, 9);
    for (let index = 0; index + 1 < id.length; index += 2) {
      appendQrBits_(bits, charset.indexOf(id[index]) * 45 + charset.indexOf(id[index + 1]), 11);
    }
    if (id.length % 2) appendQrBits_(bits, charset.indexOf(id[id.length - 1]), 6);
  }

  const dataCodewordCount = errorCorrectionLevel === "H" ? 9 : 19;
  const capacity = dataCodewordCount * 8;
  if (bits.length > capacity) throw new Error("รหัสนักเรียนยาวเกินความจุ QR Code");
  for (let index = 0; index < Math.min(4, capacity - bits.length); index += 1) bits.push(0);
  while (bits.length % 8) bits.push(0);
  const words = [];
  for (let index = 0; index < bits.length; index += 8) {
    let word = 0;
    for (let bit = 0; bit < 8; bit += 1) word = (word << 1) | bits[index + bit];
    words.push(word);
  }
  for (let pad = 0; words.length < dataCodewordCount; pad += 1) words.push(pad % 2 ? 0x11 : 0xec);
  return words;
}

function qrGfMultiply_(left, right) {
  let product = 0;
  for (let bit = 0; bit < 8; bit += 1) {
    if (right & 1) product ^= left;
    const carry = left & 0x80;
    left = (left << 1) & 0xff;
    if (carry) left ^= 0x1d;
    right >>>= 1;
  }
  return product;
}

function qrGfPower_(value, exponent) {
  let result = 1;
  for (let index = 0; index < exponent; index += 1) result = qrGfMultiply_(result, value);
  return result;
}

function createQrCodewords_(studentId, errorCorrectionLevel) {
  const dataCodewordCount = errorCorrectionLevel === "H" ? 9 : 19;
  const errorCorrectionCodewordCount = errorCorrectionLevel === "H" ? 17 : 7;
  const data = createQrDataCodewords_(studentId, errorCorrectionLevel);
  let generator = [1];
  for (let root = 0; root < errorCorrectionCodewordCount; root += 1) {
    const next = new Array(generator.length + 1).fill(0);
    for (let index = 0; index < generator.length; index += 1) {
      next[index] ^= generator[index];
      next[index + 1] ^= qrGfMultiply_(generator[index], qrGfPower_(2, root));
    }
    generator = next;
  }
  const work = data.concat(new Array(errorCorrectionCodewordCount).fill(0));
  for (let index = 0; index < data.length; index += 1) {
    const factor = work[index];
    if (!factor) continue;
    for (let term = 0; term < generator.length; term += 1) {
      work[index + term] ^= qrGfMultiply_(generator[term], factor);
    }
  }
  return data.concat(work.slice(data.length));
}

function createQrMatrix_(studentId) {
  const size = 21;
  const errorCorrectionLevel = getQrErrorCorrectionLevel_(studentId);
  const matrix = Array.from({ length: size }, () => new Array(size).fill(null));
  const placeFinder = (top, left) => {
    for (let row = -1; row <= 7; row += 1) {
      for (let column = -1; column <= 7; column += 1) {
        const targetRow = top + row;
        const targetColumn = left + column;
        if (targetRow < 0 || targetRow >= size || targetColumn < 0 || targetColumn >= size) continue;
        matrix[targetRow][targetColumn] = row >= 0 && row <= 6 && column >= 0 && column <= 6
          && (row === 0 || row === 6 || column === 0 || column === 6
            || (row >= 2 && row <= 4 && column >= 2 && column <= 4));
      }
    }
  };
  placeFinder(0, 0);
  placeFinder(size - 7, 0);
  placeFinder(0, size - 7);
  for (let index = 8; index < size - 8; index += 1) {
    if (matrix[index][6] === null) matrix[index][6] = index % 2 === 0;
    if (matrix[6][index] === null) matrix[6][index] = index % 2 === 0;
  }

  const formatData = (errorCorrectionLevel === "H" ? 2 : 1) << 3;
  let remainder = formatData << 10;
  for (let bit = 14; bit >= 10; bit -= 1) {
    if ((remainder >>> bit) & 1) remainder ^= 0x537 << (bit - 10);
  }
  const formatBits = ((formatData << 10) | remainder) ^ 0x5412;
  for (let index = 0; index < 15; index += 1) {
    const dark = ((formatBits >>> index) & 1) === 1;
    if (index < 6) matrix[index][8] = dark;
    else if (index < 8) matrix[index + 1][8] = dark;
    else matrix[size - 15 + index][8] = dark;
    if (index < 8) matrix[8][size - index - 1] = dark;
    else if (index < 9) matrix[8][15 - index] = dark;
    else matrix[8][15 - index - 1] = dark;
  }
  matrix[size - 8][8] = true;

  const codewords = createQrCodewords_(studentId, errorCorrectionLevel);
  let row = size - 1;
  let direction = -1;
  let bitIndex = 7;
  let wordIndex = 0;
  for (let column = size - 1; column > 0; column -= 2) {
    if (column === 6) column -= 1;
    while (true) {
      for (let offset = 0; offset < 2; offset += 1) {
        const targetColumn = column - offset;
        if (matrix[row][targetColumn] !== null) continue;
        let dark = wordIndex < codewords.length
          && (((codewords[wordIndex] >>> bitIndex) & 1) === 1);
        if ((row + targetColumn) % 2 === 0) dark = !dark;
        matrix[row][targetColumn] = dark;
        bitIndex -= 1;
        if (bitIndex < 0) {
          wordIndex += 1;
          bitIndex = 7;
        }
      }
      row += direction;
      if (row < 0 || row >= size) {
        row -= direction;
        direction = -direction;
        break;
      }
    }
  }
  return matrix;
}

function appendQrPngChunk_(output, type, data) {
  const chunk = [];
  const length = data.length;
  chunk.push((length >>> 24) & 0xff, (length >>> 16) & 0xff, (length >>> 8) & 0xff, length & 0xff);
  for (let index = 0; index < type.length; index += 1) chunk.push(type.charCodeAt(index));
  for (const byte of data) chunk.push(byte);
  let crc = 0xffffffff;
  for (const byte of chunk.slice(4)) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) {
      crc = (crc & 1) ? ((crc >>> 1) ^ 0xedb88320) : (crc >>> 1);
    }
  }
  crc = (crc ^ 0xffffffff) >>> 0;
  chunk.push((crc >>> 24) & 0xff, (crc >>> 16) & 0xff, (crc >>> 8) & 0xff, crc & 0xff);
  for (const byte of chunk) output.push(byte);
}

function createStudentQrPng_(studentId, requestedScale) {
  const matrix = createQrMatrix_(studentId);
  const scale = requestedScale || 4;
  const quietZone = 4;
  const width = (matrix.length + quietZone * 2) * scale;
  const raw = [];
  for (let pixelY = 0; pixelY < width; pixelY += 1) {
    raw.push(0);
    const moduleY = Math.floor(pixelY / scale) - quietZone;
    for (let pixelX = 0; pixelX < width; pixelX += 1) {
      const moduleX = Math.floor(pixelX / scale) - quietZone;
      const dark = moduleY >= 0 && moduleY < matrix.length
        && moduleX >= 0 && moduleX < matrix.length && matrix[moduleY][moduleX];
      raw.push(dark ? 0 : 0xff);
    }
  }

  const compressed = [0x78, 0x01];
  for (let offset = 0; offset < raw.length;) {
    const length = Math.min(65535, raw.length - offset);
    const finalBlock = offset + length === raw.length;
    compressed.push(finalBlock ? 1 : 0, length & 0xff, (length >>> 8) & 0xff);
    const inverseLength = (~length) & 0xffff;
    compressed.push(inverseLength & 0xff, (inverseLength >>> 8) & 0xff);
    for (let index = offset; index < offset + length; index += 1) compressed.push(raw[index]);
    offset += length;
  }
  let adlerA = 1;
  let adlerB = 0;
  for (const byte of raw) {
    adlerA = (adlerA + byte) % 65521;
    adlerB = (adlerB + adlerA) % 65521;
  }
  const adler = ((adlerB << 16) | adlerA) >>> 0;
  compressed.push((adler >>> 24) & 0xff, (adler >>> 16) & 0xff, (adler >>> 8) & 0xff, adler & 0xff);

  const png = [137, 80, 78, 71, 13, 10, 26, 10];
  appendQrPngChunk_(png, "IHDR", [
    (width >>> 24) & 0xff, (width >>> 16) & 0xff, (width >>> 8) & 0xff, width & 0xff,
    (width >>> 24) & 0xff, (width >>> 16) & 0xff, (width >>> 8) & 0xff, width & 0xff,
    8, 0, 0, 0, 0
  ]);
  appendQrPngChunk_(png, "IDAT", compressed);
  appendQrPngChunk_(png, "IEND", []);
  return png;
}

function finalizeRosterUpload_(session, requestedRoom, uploadId) {
  const roomKey = normalizeRoomKey_(requestedRoom);
  if (!roomKey) return createJsonResponse_(false, "กรุณาระบุชั้น/ห้อง");
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const stage = getRosterStage_(session, uploadId);
  if (!stage) return createJsonResponse_(false, "ไม่พบรายการอัปโหลดที่กำลังดำเนินการ", "UPLOAD_NOT_FOUND");
  const count = stage.getLastRow() - 1;
  if (count < 1 || stage.getImages().length !== count) {
    return createJsonResponse_(false, "อัปโหลดภาพ QR ยังไม่ครบ กรุณาลองอัปโหลดอีกครั้ง", "QR_IMAGES_INCOMPLETE", {
      expected: count,
      uploaded: stage.getImages().length
    });
  }

  const canonicalName = classSheetName_(requestedRoom);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let previousSheet = null;
  let previousName = "";
  try {
    previousSheet = findClassSheet_(spreadsheet, requestedRoom);
    if (previousSheet === stage) {
      throw new Error("ชีตชั่วคราวไม่ควรถูกเลือกเป็นชีตห้องเรียน");
    }
    if (previousSheet) {
      previousName = "__OLD_" + session.id + "_" + uploadId;
      previousSheet.setName(previousName);
    }
    stage.setName(canonicalName);
    if (previousSheet) spreadsheet.deleteSheet(previousSheet);
  } catch (error) {
    if (spreadsheet.getSheetByName(canonicalName) === stage) {
      stage.setName(rosterStageName_(session, uploadId));
    }
    const renamedPreviousSheet = previousName && spreadsheet.getSheetByName(previousName);
    if (renamedPreviousSheet && !spreadsheet.getSheetByName(canonicalName)) {
      renamedPreviousSheet.setName(canonicalName);
    }
    throw error;
  } finally {
    lock.releaseLock();
  }
  return createJsonResponse_(true, "อัปโหลดรายชื่อและภาพ QR เรียบร้อย", null, {
    room: requestedRoom.trim(),
    sheet: canonicalName,
    count: count
  });
}

function cancelRosterUpload_(session, uploadId) {
  const stage = getRosterStage_(session, uploadId);
  if (!stage) return createJsonResponse_(true, "ไม่มีข้อมูลชั่วคราวให้ยกเลิก");
  SpreadsheetApp.getActiveSpreadsheet().deleteSheet(stage);
  return createJsonResponse_(true, "ยกเลิกการอัปโหลดแล้ว");
}

function adminRoomKey_(value) {
  const room = String(value || "").trim().replace(/^ม\.\s*/i, "");
  if (!/^[1-6]\d{1,2}$/.test(room)) return "";
  const classNumber = Number(room.slice(1));
  return classNumber >= 1 && classNumber <= 99 ? room[0] + classNumber : "";
}

function getAdminClassSheets_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()
    .map(sheet => {
      const room = adminRoomKey_(sheet.getName());
      if (!room || sheet.getName() !== classSheetName_(room)) return null;
      return {
        sheet: sheet,
        room: room,
        grade: room[0],
        classroom: room.slice(1),
        count: Math.max(0, sheet.getLastRow() - 1)
      };
    })
    .filter(Boolean)
    .sort((left, right) => Number(left.grade) - Number(right.grade)
      || Number(left.classroom) - Number(right.classroom));
}

function getAdminRoster_(requestedRoom) {
  const rooms = getAdminClassSheets_();
  const result = rooms.map(item => ({
    room: item.room,
    sheet: item.sheet.getName(),
    grade: item.grade,
    classroom: item.classroom,
    count: item.count
  }));
  if (!requestedRoom) return createJsonResponse_(true, "โหลดรายการห้องแล้ว", null, { rooms: result, students: [] });

  const room = adminRoomKey_(requestedRoom);
  const selected = rooms.find(item => item.room === room);
  if (!selected) return createJsonResponse_(false, "ไม่พบห้องเรียนที่เลือก");
  const values = selected.sheet.getDataRange().getDisplayValues();
  const students = values.slice(1)
    .filter(row => String(row[0] || "").trim())
    .map(row => ({
      id: normalizeId_(row[0]),
      number: String(row[1] || "").trim(),
      name: String(row[2] || "").trim()
    }));
  return createJsonResponse_(true, "โหลดรายชื่อนักเรียนแล้ว", null, {
    rooms: result,
    room: room,
    students: students
  });
}

function getAdminAttendanceCounts_(requestedRoom, requestedStartDate, requestedEndDate) {
  const range = validAttendanceDateRange_(requestedStartDate, requestedEndDate);
  if (!range) return createJsonResponse_(false, "กรุณาระบุช่วงวันที่ให้ถูกต้อง โดยวันเริ่มต้นต้องไม่อยู่หลังวันสิ้นสุด");
  const room = adminRoomKey_(requestedRoom);
  if (!room) return createJsonResponse_(false, "กรุณาเลือกชั้นและห้องเรียน");

  const selected = getAdminClassSheets_().find(item => item.room === room);
  if (!selected) return createJsonResponse_(false, "ไม่พบห้องเรียนที่เลือก");
  const students = getAttendanceCounts_(selected.sheet, room, range.startDate, range.endDate);
  return createJsonResponse_(true, "โหลดสรุปจำนวนครั้งการเข้าเรียนแล้ว", null, {
    room: room,
    startDate: range.startDate,
    endDate: range.endDate,
    students: students
  });
}

function insertStudentQrImage_(sheet, row, studentId) {
  const blob = Utilities.newBlob(createStudentQrPng_(studentId), "image/png", "qr-" + studentId + ".png");
  const image = sheet.insertImage(blob, 4, row);
  image.setWidth(100).setHeight(100);
  image.setAnchorCellXOffset(10).setAnchorCellYOffset(5);
  return image;
}

function removeStudentQrImages_(sheet, row, exceptImage) {
  for (const image of sheet.getImages()) {
    const anchor = image.getAnchorCell();
    if (image !== exceptImage && anchor.getColumn() === 4 && anchor.getRow() === row) image.remove();
  }
}

function findAdminStudentRow_(sheet, studentId) {
  if (sheet.getLastRow() < 2) return 0;
  const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
  const index = ids.findIndex(row => normalizeId_(row[0]) === studentId);
  return index < 0 ? 0 : index + 2;
}

function adminStudentIdExists_(studentId, sourceSheet, sourceRow) {
  for (const item of getAdminClassSheets_()) {
    const row = findAdminStudentRow_(item.sheet, studentId);
    if (row && (item.sheet !== sourceSheet || row !== sourceRow)) return true;
  }
  return false;
}

function updateAdminStudentAttendance_(oldId, sourceRoom, newId, number, name, targetRoom) {
  if (!oldId) return;
  const attendance = getAttendanceRows_();
  for (let index = 0; index < attendance.rows.length; index += 1) {
    const row = attendance.rows[index];
    if (normalizeId_(row[1]) !== oldId || normalizeRoomKey_(row[4]) !== sourceRoom) continue;
    attendance.sheet.getRange(index + 2, 2, 1, 5).setValues([[
      newId,
      number,
      name,
      targetRoom,
      row[5]
    ]]);
  }
}

function saveAdminStudent_(sourceRoomValue, oldIdValue, targetRoomValue, studentIdValue, numberValue, nameValue) {
  const sourceRoom = sourceRoomValue ? adminRoomKey_(sourceRoomValue) : "";
  const targetRoom = adminRoomKey_(targetRoomValue);
  const oldId = normalizeId_(oldIdValue);
  const studentId = normalizeId_(studentIdValue);
  const number = String(numberValue || "").trim();
  const name = String(nameValue || "").trim();
  if (!targetRoom || !studentId || !name) return createJsonResponse_(false, "กรุณาระบุห้อง รหัสนักเรียน และชื่อ-นามสกุลให้ครบ");
  if (sourceRoomValue && !sourceRoom) return createJsonResponse_(false, "ห้องต้นทางไม่ถูกต้อง");
  try {
    validateQrStudentId_(studentId);
  } catch (error) {
    return createJsonResponse_(false, error.message);
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = findClassSheet_(spreadsheet, targetRoom);
  if (!targetSheet) return createJsonResponse_(false, "ไม่พบห้องปลายทาง กรุณาสร้างห้องก่อน");
  const sourceSheet = sourceRoom ? findClassSheet_(spreadsheet, sourceRoom) : null;
  if (sourceRoom && !sourceSheet) return createJsonResponse_(false, "ไม่พบห้องต้นทาง");

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let destinationRow = 0;
  let insertedImage = null;
  let previousValues = null;
  const moving = sourceSheet !== targetSheet;
  let sourceRow = 0;
  try {
    sourceRow = sourceSheet && oldId ? findAdminStudentRow_(sourceSheet, oldId) : 0;
    if (sourceRoom && !sourceRow) return createJsonResponse_(false, "ไม่พบนักเรียนที่ต้องการแก้ไขในห้องต้นทาง");
    if (adminStudentIdExists_(studentId, sourceSheet, sourceRow)) {
      return createJsonResponse_(false, "รหัสนักเรียนนี้มีอยู่ในระบบแล้ว");
    }

    if (sourceRow && !moving) {
      destinationRow = sourceRow;
      previousValues = sourceSheet.getRange(sourceRow, 1, 1, 4).getValues()[0];
      sourceSheet.getRange(sourceRow, 1, 1, 4).setNumberFormat("@");
      sourceSheet.getRange(sourceRow, 1, 1, 4).setValues([[studentId, number, name, ""]]);
      try {
        insertedImage = insertStudentQrImage_(targetSheet, destinationRow, studentId);
      } catch (error) {
        sourceSheet.getRange(sourceRow, 1, 1, 4).setValues([previousValues]);
        throw error;
      }
    } else {
      destinationRow = targetSheet.getLastRow() + 1;
      if (destinationRow > targetSheet.getMaxRows()) targetSheet.insertRowAfter(targetSheet.getMaxRows());
      targetSheet.getRange(destinationRow, 1, 1, 4).setNumberFormat("@");
      targetSheet.getRange(destinationRow, 1, 1, 4).setValues([[studentId, number, name, ""]]);
      try {
        insertedImage = insertStudentQrImage_(targetSheet, destinationRow, studentId);
      } catch (error) {
        targetSheet.deleteRow(destinationRow);
        throw error;
      }
    }

    if (sourceRow) updateAdminStudentAttendance_(oldId, sourceRoom, studentId, number, name, targetRoom);
    if (sourceRow && moving) {
      removeStudentQrImages_(sourceSheet, sourceRow);
      sourceSheet.deleteRow(sourceRow);
    } else if (sourceRow) {
      removeStudentQrImages_(sourceSheet, sourceRow, insertedImage);
    }
  } catch (error) {
    if (insertedImage) insertedImage.remove();
    if (previousValues && sourceSheet && destinationRow === sourceRow) {
      sourceSheet.getRange(sourceRow, 1, 1, 4).setValues([previousValues]);
    } else if (destinationRow && targetSheet.getLastRow() >= destinationRow) {
      targetSheet.deleteRow(destinationRow);
    }
    throw error;
  } finally {
    lock.releaseLock();
  }
  return createJsonResponse_(true, "บันทึกข้อมูลนักเรียนและ QR Code แล้ว", null, {
    room: targetRoom,
    student: { id: studentId, number: number, name: name }
  });
}

function deleteAdminStudent_(requestedRoom, requestedId) {
  const room = adminRoomKey_(requestedRoom);
  const studentId = normalizeId_(requestedId);
  if (!room || !studentId) return createJsonResponse_(false, "ห้องหรือรหัสนักเรียนไม่ถูกต้อง");
  const sheet = findClassSheet_(SpreadsheetApp.getActiveSpreadsheet(), room);
  if (!sheet) return createJsonResponse_(false, "ไม่พบห้องเรียนที่เลือก");

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const row = findAdminStudentRow_(sheet, studentId);
    if (!row) return createJsonResponse_(false, "ไม่พบนักเรียนที่ต้องการลบ");
    deleteAdminAttendance_(studentId, room);
    removeStudentQrImages_(sheet, row);
    sheet.deleteRow(row);
  } finally {
    lock.releaseLock();
  }
  return createJsonResponse_(true, "ลบนักเรียนและประวัติเช็คชื่อแล้ว");
}

function deleteAdminAttendance_(studentId, room) {
  const attendance = getAttendanceRows_();
  const rowsToDelete = [];
  for (let index = 0; index < attendance.rows.length; index += 1) {
    const row = attendance.rows[index];
    if ((!studentId || normalizeId_(row[1]) === studentId) && normalizeRoomKey_(row[4]) === room) {
      rowsToDelete.push(index + 2);
    }
  }
  for (let index = rowsToDelete.length - 1; index >= 0; index -= 1) {
    attendance.sheet.deleteRow(rowsToDelete[index]);
  }
}

function createAdminRoom_(gradeValue, classroomValue) {
  const grade = String(gradeValue || "").trim();
  const classroom = String(classroomValue || "").trim();
  if (!/^[1-6]$/.test(grade) || !/^\d{1,2}$/.test(classroom)
      || Number(classroom) < 1 || Number(classroom) > 99) {
    return createJsonResponse_(false, "กรุณาเลือกระดับ ม.1–ม.6 และระบุเลขห้อง 1–99");
  }
  const room = grade + Number(classroom);
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (findClassSheet_(spreadsheet, room)) return createJsonResponse_(false, "ห้องนี้มีอยู่แล้ว");
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (findClassSheet_(spreadsheet, room)) return createJsonResponse_(false, "ห้องนี้มีอยู่แล้ว");
    const sheet = spreadsheet.insertSheet(classSheetName_(room));
    sheet.getRange(1, 1, 1, 4).setValues([["รหัสนักเรียน", "เลขที่", "ชื่อ-นามสกุล", "QR Code"]]);
    sheet.getRange(1, 1, 1, 4).setNumberFormat("@");
    sheet.setColumnWidth(4, 120);
  } finally {
    lock.releaseLock();
  }
  return createJsonResponse_(true, "สร้างห้องเรียนแล้ว", null, { room: room });
}

function deleteAdminRoom_(requestedRoom) {
  const room = adminRoomKey_(requestedRoom);
  if (!room) return createJsonResponse_(false, "ห้องที่ระบุไม่ถูกต้อง");
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = findClassSheet_(spreadsheet, room);
  if (!sheet) return createJsonResponse_(false, "ไม่พบห้องเรียนที่ต้องการลบ");
  if (spreadsheet.getSheets().length <= 1) return createJsonResponse_(false, "ไม่สามารถลบชีตสุดท้ายของ Spreadsheet ได้");

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    deleteAdminAttendance_("", room);
    const teacherSheet = spreadsheet.getSheetByName(TEACHER_SHEET_NAME);
    if (teacherSheet && teacherSheet.getLastRow() >= 2) {
      const teacherRooms = teacherSheet.getRange(2, 3, teacherSheet.getLastRow() - 1, 1).getDisplayValues();
      for (let index = 0; index < teacherRooms.length; index += 1) {
        if (normalizeRoomKey_(teacherRooms[index][0]) === room) {
          teacherSheet.getRange(index + 2, 3).clearContent();
        }
      }
    }
    spreadsheet.deleteSheet(sheet);
  } finally {
    lock.releaseLock();
  }
  return createJsonResponse_(true, "ลบห้องและประวัติเช็คชื่อแล้ว ข้อมูลครูยังคงอยู่");
}

function getAdminTeachers_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const teacherSheet = spreadsheet.getSheetByName(TEACHER_SHEET_NAME);
  if (!teacherSheet) return createJsonResponse_(false, "ไม่พบชีต Teacher");
  const teachers = teacherSheet.getLastRow() < 2
    ? []
    : teacherSheet.getRange(2, 1, teacherSheet.getLastRow() - 1, 3).getDisplayValues()
      .map(row => ({
        id: normalizeId_(row[0]),
        name: String(row[1] || "").trim(),
        room: String(row[2] || "").trim()
      }))
      .filter(teacher => teacher.id);
  const rooms = getAdminClassSheets_().map(item => ({
    room: item.room,
    grade: item.grade,
    classroom: item.classroom
  }));
  return createJsonResponse_(true, "โหลดรายชื่อครูแล้ว", null, { teachers: teachers, rooms: rooms });
}

function assignTeacherRoom_(requestedTeacherId, requestedRoom) {
  const teacherId = normalizeId_(requestedTeacherId);
  const room = requestedRoom ? adminRoomKey_(requestedRoom) : "";
  if (!teacherId) return createJsonResponse_(false, "กรุณาเลือกรหัสครู");
  if (requestedRoom && !room) return createJsonResponse_(false, "ห้องเรียนที่ระบุไม่ถูกต้อง");

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const teacherSheet = spreadsheet.getSheetByName(TEACHER_SHEET_NAME);
  if (!teacherSheet || teacherSheet.getLastRow() < 2) return createJsonResponse_(false, "ไม่พบข้อมูลครูในชีต Teacher");
  const teacherRow = findTeacherById_(teacherId);
  if (!teacherRow) return createJsonResponse_(false, "ไม่พบครูที่เลือกในชีต Teacher");
  const targetSheet = room ? findClassSheet_(spreadsheet, room) : null;
  if (room && !targetSheet) return createJsonResponse_(false, "ไม่พบห้องที่เลือก กรุณาสร้างห้องก่อน");

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (teacherSheet.getLastRow() < 2) return createJsonResponse_(false, "ไม่พบข้อมูลครูในชีต Teacher");
    const row = teacherSheet.getRange(2, 1, teacherSheet.getLastRow() - 1, 1).getDisplayValues()
      .findIndex(value => normalizeId_(value[0]) === teacherId) + 2;
    if (row < 2) return createJsonResponse_(false, "ไม่พบครูที่เลือกในชีต Teacher");
    teacherSheet.getRange(row, 3).setNumberFormat("@");
    teacherSheet.getRange(row, 3).setValue(room);
  } finally {
    lock.releaseLock();
  }
  return createJsonResponse_(true, room ? "กำหนดห้องประจำครูแล้ว" : "ยกเลิกการกำหนดห้องแล้ว", null, {
    teacherId: teacherId,
    room: room
  });
}

function assignTeachersToRoom_(requestedRoom, rawTeacherIds) {
  const room = adminRoomKey_(requestedRoom);
  if (!room) return createJsonResponse_(false, "ห้องเรียนที่ระบุไม่ถูกต้อง");
  let teacherIds;
  try {
    teacherIds = JSON.parse(rawTeacherIds);
  } catch (error) {
    return createJsonResponse_(false, "อ่านรายชื่อครูที่เลือกไม่สำเร็จ");
  }
  if (!Array.isArray(teacherIds) || teacherIds.length > 200
      || teacherIds.some(id => typeof id !== "string")) {
    return createJsonResponse_(false, "รูปแบบรายชื่อครูที่เลือกไม่ถูกต้อง");
  }
  const selectedIds = teacherIds.map(normalizeId_);
  if (selectedIds.some(id => !id) || new Set(selectedIds).size !== selectedIds.length) {
    return createJsonResponse_(false, "พบรหัสครูว่างหรือซ้ำในรายการที่เลือก");
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const teacherSheet = spreadsheet.getSheetByName(TEACHER_SHEET_NAME);
  if (!teacherSheet || teacherSheet.getLastRow() < 2) {
    return createJsonResponse_(false, "ไม่พบข้อมูลครูในชีต Teacher");
  }
  if (!findClassSheet_(spreadsheet, room)) {
    return createJsonResponse_(false, "ไม่พบห้องที่เลือก กรุณาสร้างห้องก่อน");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const lastRow = teacherSheet.getLastRow();
    const rows = teacherSheet.getRange(2, 1, lastRow - 1, 3).getDisplayValues();
    const rowById = new Map();
    rows.forEach((row, index) => {
      const id = normalizeId_(row[0]);
      if (id) rowById.set(id, { row: index + 2, room: String(row[2] || "").trim() });
    });

    for (const id of selectedIds) {
      const teacher = rowById.get(id);
      if (!teacher) return createJsonResponse_(false, "ไม่พบครูที่เลือกในชีต Teacher: " + id);
      if (teacher.room && normalizeRoomKey_(teacher.room) !== room) {
        return createJsonResponse_(false, "ครู " + id + " ได้รับมอบหมายให้อยู่ห้อง " + teacher.room + " แล้ว");
      }
    }

    const selected = new Set(selectedIds);
    for (const [id, teacher] of rowById) {
      const isAssignedToRoom = normalizeRoomKey_(teacher.room) === room;
      if (selected.has(id) || isAssignedToRoom) {
        teacherSheet.getRange(teacher.row, 3).setNumberFormat("@");
        teacherSheet.getRange(teacher.row, 3).setValue(selected.has(id) ? room : "");
      }
    }
  } finally {
    lock.releaseLock();
  }

  return createJsonResponse_(true, "บันทึกรายชื่อครูประจำห้องแล้ว", null, {
    room: room,
    teacherIds: selectedIds
  });
}

function exportAttendance_(session, requestedDate) {
  const selectedDate = parseDateKey_(requestedDate);
  if (!selectedDate) return createJsonResponse_(false, "รูปแบบวันที่ไม่ถูกต้อง");
  if (!session.room) return createJsonResponse_(false, "บัญชีครูนี้ยังไม่ได้รับมอบหมายห้องเรียน");
  const attendance = getAttendanceRows_();
  const roomKey = normalizeRoomKey_(session.room);
  const rows = attendance.rows
    .filter(row => dateKey_(row[0]) === selectedDate && normalizeRoomKey_(row[4]) === roomKey)
    .map(row => [
      Utilities.formatDate(row[0], Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"),
      normalizeId_(row[1]),
      String(row[2] == null ? "" : row[2]),
      String(row[3] || ""),
      String(row[4] || ""),
      String(row[5] || "")
    ]);
  return createJsonResponse_(true, "ส่งออกรายงานแล้ว", null, {
    room: session.room,
    date: selectedDate,
    rows: rows
  });
}

function createJsonResponse_(success, message, code, extra) {
  const result = { success: success, message: message };
  if (code) result.code = code;
  if (extra) Object.assign(result, extra);
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
