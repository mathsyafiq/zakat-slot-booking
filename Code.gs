// ============================================================
// SLOT BOOKING SYSTEM - Code.gs (API backend)
// Deploy as Web App: Execute as "Me", Access "Anyone"
// Frontend is hosted on GitHub Pages (see /docs)
// ============================================================

const SHEET_NAME = "Slots";
const WA_NUMBER = "60162099151";

// ── JSON helper ──────────────────────────────────────────────
function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── GET  ?action=slots  → list slots ─────────────────────────
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || "slots";
  if (action === "slots") {
    try {
      return json({ success: true, slots: getSlots() });
    } catch (err) {
      return json({ success: false, message: String(err) });
    }
  }
  return json({ success: false, message: "Unknown action." });
}

// ── POST (text/plain JSON body) → book a slot ────────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    return json(bookSlot(body.slotId, body.name, body.studentId, body.phone));
  } catch (err) {
    return json({ success: false, message: "Invalid request." });
  }
}

// ── Get slots ────────────────────────────────────────────────
function getSlots() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const slots = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    slots.push({
      id:     String(row[0]),
      date:   formatDate(row[1]),
      time:   formatTime(row[2]),
      status: row[3],
      label:  row[4]
    });
  }
  return slots;
}

// ── Format helpers — handle Date objects and plain text ──────
function formatDate(value) {
  if (!value && value !== 0) return "";
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "EEE, d MMM yyyy");
  }
  return String(value);
}

function formatTime(value) {
  if (!value && value !== 0) return "";
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "h:mm a");
  }
  return String(value);
}

// ── Book a slot (with lock to prevent double booking) ────────
function bookSlot(slotId, name, studentId, phone) {
  if (!slotId || !name || !studentId || !phone) {
    return { success: false, message: "All fields are required." };
  }
  name = String(name).trim().slice(0, 100);
  studentId = String(studentId).trim().slice(0, 30);
  phone = String(phone).trim().slice(0, 20);

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, message: "System is busy. Please try again." };
  }

  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();

    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(slotId)) {
        targetRow = i + 1;
        if (data[i][3] === "Booked") {
          return { success: false, message: "Sorry, this slot was just booked by someone else. Please choose another." };
        }
        break;
      }
    }

    if (targetRow === -1) {
      return { success: false, message: "Slot not found." };
    }

    sheet.getRange(targetRow, 4).setValue("Booked");
    sheet.getRange(targetRow, 6).setValue(name);
    sheet.getRange(targetRow, 7).setValue(studentId);
    sheet.getRange(targetRow, 8).setValue(phone);
    sheet.getRange(targetRow, 9).setValue(new Date());

    const rowData = sheet.getRange(targetRow, 1, 1, 9).getValues()[0];
    return {
      success: true,
      booking: {
        slotId:    String(rowData[0]),
        date:      formatDate(rowData[1]),
        time:      formatTime(rowData[2]),
        name:      name,
        studentId: studentId,
        phone:     phone,
        waNumber:  WA_NUMBER
      }
    };
  } finally {
    lock.releaseLock();
  }
}

// ── Sheet helpers ────────────────────────────────────────────
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = setupSheet(ss);
  return sheet;
}

function setupSheet(ss) {
  const sheet = ss.insertSheet(SHEET_NAME);

  const headers = [
    "Slot ID", "Date", "Time", "Status", "Label",
    "Booked By", "Student ID", "Phone", "Booked At"
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground("#1a1a2e");
  headerRange.setFontColor("#ffffff");
  headerRange.setFontWeight("bold");

  const defaultSlots = [];
  for (let i = 1; i <= 15; i++) {
    defaultSlots.push(["S" + (i < 10 ? "0" + i : i), "TBD", "TBD", "Available", ""]);
  }
  sheet.getRange(2, 1, defaultSlots.length, 5).setValues(defaultSlots);
  sheet.autoResizeColumns(1, headers.length);

  return sheet;
}

// Run this once manually to initialize the sheet
function initialize() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(SHEET_NAME)) {
    Logger.log("Sheet already exists. Delete it first if you want to reset.");
    return;
  }
  setupSheet(ss);
  Logger.log("Sheet initialized successfully!");
}
