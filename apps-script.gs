// ============================================================
// TAHOE 2026 — Google Apps Script backend (~2 minute setup)
// ============================================================
//
// What this does:
//   - Exposes a single Web App URL that:
//       GET  → returns the spreadsheet as JSON (people, chores, lodging, etc.)
//       POST → appends/removes rows in the "chore_completions" tab when
//              someone taps "I did this" on the site.
//   - No login. Anyone with the site can mark a chore done.
//   - Each device dedupes via localStorage so kids can't spam.
//
// SETUP (do this once):
//   1. Open the "Tahoe 2026" Google Sheet.
//   2. Add a new tab called  chore_completions   with these headers in row 1:
//         A1: at  |  B1: person  |  C1: date  |  D1: chore  |  E1: device
//      (you don't have to fill any rows — the script appends to it)
//   3. Add a new tab called  schedule  with these headers in row 1:
//         A1: date  |  B1: time  |  C1: title  |  D1: lead  |  E1: kind
//      kind = one of: meal, hike, beach, kids, logistics
//   4. Extensions → Apps Script. Delete any starter code. Paste ALL of
//      this file in. Save.
//   5. Deploy → New deployment → Type: Web app.
//        Description:   Tahoe 2026
//        Execute as:    Me  (your account)
//        Who has access: Anyone
//      Click Deploy, authorize when prompted.
//   6. Copy the Web app URL it gives you. It looks like:
//        https://script.google.com/macros/s/AKfycbx.../exec
//   7. Paste that URL into tahoe/config.js as `appsScriptUrl`. Done.
//
// To update later: Deploy → Manage deployments → edit → New version.
// ============================================================

const SHEET_HEADCOUNTS  = "head counts";
const SHEET_MEALSCHORES = "meals & chores";
const SHEET_LODGING     = "lodging";
const SHEET_SCHEDULE    = "schedule";
const SHEET_COMPLETIONS = "chore_completions";

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const payload = {
    people:       readPeople(ss),
    dinnerLeads:  readDinnerLeads(ss),
    choreSlots:   readChoreSlots(ss),
    completions:  readCompletions(ss),
    lodging:      readLodging(ss),
    schedule:     readSchedule(ss),
  };
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_COMPLETIONS) || ss.insertSheet(SHEET_COMPLETIONS);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["at", "person", "date", "chore", "device"]);
  }

  if (data.action === "complete") {
    sheet.appendRow([data.at || new Date().toISOString(), data.person, data.date, data.chore, data.device || ""]);
  } else if (data.action === "signup") {
    // Write the person's name into the meals & chores signup grid.
    // Fills the first empty row matching this role on this date (some roles,
    // like Kitchen clean-up, have several rows).
    const mc = ss.getSheetByName(SHEET_MEALSCHORES);
    const dates = ["2026-07-20","2026-07-21","2026-07-22","2026-07-23","2026-07-24","2026-07-25","2026-07-26"];
    const dayIdx = dates.indexOf(data.date);
    if (mc && dayIdx >= 0) {
      const col = 4 + dayIdx; // D..J
      const roles = mc.getRange(7, 3, 14, 1).getValues(); // C7:C20
      for (let r = 0; r < roles.length; r++) {
        if (roles[r][0].toString().trim() === (data.role || "").toString().trim()) {
          const cell = mc.getRange(7 + r, col);
          if (!cell.getValue()) { cell.setValue(data.person); break; }
        }
      }
    }
  } else if (data.action === "undo") {
    // remove the most recent matching row. Sheets turns "2026-07-20" cells
    // into Date objects, so compare everything as normalized strings.
    const norm = v => (v instanceof Date)
      ? Utilities.formatDate(v, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd")
      : (v == null ? "" : v.toString().trim());
    const rows = sheet.getDataRange().getValues();
    for (let i = rows.length - 1; i >= 1; i--) {
      if (norm(rows[i][1]) === norm(data.person) && norm(rows[i][2]) === norm(data.date) && norm(rows[i][3]) === norm(data.chore)) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
  }
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- READERS ----------

// Summary rows at the bottom of the head counts tab — not people.
const TOTALS_ROW_NAMES = ["adults", "kids", "grand total", "day visitor total", "night headcount"];

function readPeople(ss) {
  const sheet = ss.getSheetByName(SHEET_HEADCOUNTS);
  if (!sheet) return [];
  const dayCols = readHeadcountDayCols(sheet); // map of colIdx -> short date "7/20"
  const lastRow = sheet.getLastRow();
  if (lastRow < 5) return [];
  const rng = sheet.getRange(5, 1, lastRow - 4, 12).getValues();
  const out = [];
  rng.forEach(row => {
    const adult = !!row[0];
    const kid   = !!row[1];
    const name  = (row[2] || "").toString().trim();
    if (!name) return;
    if (TOTALS_ROW_NAMES.indexOf(name.toLowerCase()) !== -1) return;
    const days = [];
    Object.keys(dayCols).forEach(col => {
      if (row[col] === true || row[col] === "TRUE") days.push(dayCols[col]);
    });
    const dayVisitor = row[10] === true || row[10] === "TRUE";
    const note = (row[11] || "").toString().trim();
    out.push({
      name,
      role: kid ? "kid" : adult ? "adult" : "other",
      days,
      dayVisitor: dayVisitor || undefined,
      note: note || undefined,
    });
  });
  return out;
}

function readHeadcountDayCols(sheet) {
  // The headcounts sheet has columns Mon..Sun in row 2 ("Mon", "Tu", ...)
  // and row 4 has dates "20","21",... We'll just hard-map for July 20–26.
  return { 3: "7/20", 4: "7/21", 5: "7/22", 6: "7/23", 7: "7/24", 8: "7/25", 9: "7/26" };
}

function readDinnerLeads(ss) {
  const sheet = ss.getSheetByName(SHEET_MEALSCHORES);
  if (!sheet) return {};
  // Row 2 in cols D..J = dinner lead, Row 3 in cols D..J = menu
  const leads = sheet.getRange(2, 4, 1, 7).getValues()[0];
  const menus = sheet.getRange(3, 4, 1, 7).getValues()[0];
  const dates = ["2026-07-20","2026-07-21","2026-07-22","2026-07-23","2026-07-24","2026-07-25","2026-07-26"];
  const out = {};
  dates.forEach((d, i) => {
    out[d] = { leads: (leads[i] || "").toString(), menu: (menus[i] || "").toString() };
  });
  return out;
}

function readChoreSlots(ss) {
  const sheet = ss.getSheetByName(SHEET_MEALSCHORES);
  if (!sheet) return [];
  // Rows 7–11: Dinner help sign-up (Appetizer, Side, Salad, Dessert, Sous chef)
  // Rows 12–20: Chore sign-up
  const dates = ["2026-07-20","2026-07-21","2026-07-22","2026-07-23","2026-07-24","2026-07-25","2026-07-26"];
  const out = [];
  const range = sheet.getRange(7, 3, 14, 8).getValues(); // C..J for rows 7..20
  for (let r = 0; r < range.length; r++) {
    const role = (range[r][0] || "").toString().trim();
    if (!role) continue;
    const kind = r < 5 ? "help" : "chore";
    for (let c = 0; c < 7; c++) {
      const person = (range[r][c + 1] || "").toString().trim();
      out.push({ date: dates[c], kind, role, person });
    }
  }
  return out;
}

function readCompletions(ss) {
  const sheet = ss.getSheetByName(SHEET_COMPLETIONS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  return rows
    .filter(r => r[1] && r[2] && r[3])
    .map(r => ({
      at:     r[0] ? new Date(r[0]).toISOString() : "",
      person: r[1].toString(),
      date:   (r[2] instanceof Date)
        ? Utilities.formatDate(r[2], ss.getSpreadsheetTimeZone(), "yyyy-MM-dd")
        : r[2].toString(),
      chore:  r[3].toString(),
    }));
}

function readLodging(ss) {
  const sheet = ss.getSheetByName(SHEET_LODGING);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  // Cols A=Location, B=Room, C=Bed, D..J = 7 days
  const rng = sheet.getRange(3, 1, lastRow - 2, 10).getValues();
  const out = [];
  let lastLoc = "";
  let lastRoom = "";
  const dayShorts = ["7/20","7/21","7/22","7/23","7/24","7/25","7/26"];
  rng.forEach(row => {
    const loc = (row[0] || "").toString().trim() || lastLoc;
    if (row[0]) lastLoc = loc;
    // A blank Room cell means "same room as the row above" (extra beds
    // in the dorm, second double, etc.) — carry it forward.
    const roomRaw = (row[1] || "").toString().trim();
    const room = roomRaw || lastRoom;
    if (roomRaw) lastRoom = roomRaw;
    const bed = (row[2] || "").toString().trim();
    const assignments = {};
    dayShorts.forEach((s, i) => {
      const v = (row[3 + i] || "").toString().trim();
      if (v) assignments[s] = v;
    });
    if (!room) return;
    // Carried-room rows only count when they actually hold a bed or people
    // (so blank spacer rows don't become phantom beds).
    if (!roomRaw && !bed && Object.keys(assignments).length === 0) return;
    out.push({ location: loc, room, bed, assignments });
  });
  return out;
}

function readSchedule(ss) {
  // Merges the "schedule" tab plus any tab whose name starts with "itiner"
  // (itinerary, itinerayr, ...) — so the family can add hikes, bike rides,
  // etc. on their own tab and they show up on the site's schedule page.
  // Columns are matched by header name (row 1), falling back to position:
  //   date | time | title | lead | kind
  const out = [];
  ss.getSheets().forEach(sheet => {
    const name = sheet.getName().toString().trim().toLowerCase();
    if (name !== SHEET_SCHEDULE && name.indexOf("itiner") !== 0) return;
    if (sheet.getLastRow() < 2) return;
    const ncols = Math.max(5, sheet.getLastColumn());
    const headers = sheet.getRange(1, 1, 1, ncols).getValues()[0]
      .map(h => h.toString().trim().toLowerCase());
    const claimed = {};
    const col = (names, fallback) => {
      for (const n of names) { const i = headers.indexOf(n); if (i >= 0 && !claimed[i]) { claimed[i] = true; return i; } }
      if (fallback < ncols && !claimed[fallback]) { claimed[fallback] = true; return fallback; }
      return -1;
    };
    const cDate  = col(["date", "day"], 0);
    const cTime  = col(["time", "when"], 1);
    const cTitle = col(["title", "plan", "plans", "activity", "event", "what"], 2);
    const cLead  = col(["lead", "leader", "who", "host"], 3);
    const cKind  = col(["kind", "type", "category"], 4);
    const cell = (r, i) => (i >= 0 && r[i] != null) ? r[i] : "";
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, ncols).getValues();
    rows.filter(r => cell(r, cDate) !== "" && cell(r, cTitle) !== "").forEach(r => {
      const dateRaw = cell(r, cDate);
      let date;
      if (dateRaw instanceof Date) {
        date = Utilities.formatDate(dateRaw, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
      } else if (typeof dateRaw === "number") {
        // Bare day-of-month like 20 → reunion month
        date = "2026-07-" + ("0" + Math.round(dateRaw)).slice(-2);
      } else {
        date = dateRaw.toString().trim();
        const md = date.match(/^(\d{1,2})[\/\-](\d{1,2})$/); // "7/20" → "2026-07-20"
        if (md) date = "2026-" + ("0" + md[1]).slice(-2) + "-" + ("0" + md[2]).slice(-2);
      }
      const timeRaw = cell(r, cTime);
      const time = (timeRaw instanceof Date)
        ? Utilities.formatDate(timeRaw, ss.getSpreadsheetTimeZone(), "h:mm a")
        : timeRaw.toString();
      out.push({
        date,
        time,
        title: cell(r, cTitle).toString(),
        lead:  cell(r, cLead).toString(),
        kind:  cell(r, cKind).toString(),
      });
    });
  });
  return out.sort((a, b) => (a.date + a.time).localeCompare(b.date + b.time));
}
