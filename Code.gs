/**
 * Idea Board — Google Apps Script backend
 * Turns a Google Sheet into a shared API for the Idea Board site.
 *
 * SETUP: see setup-instructions.md
 * Auto-creates the "Ideas", "Comments" and "TimeSuggestions" tabs on first run.
 */

const IDEAS_SHEET = "Ideas";
const COMMENTS_SHEET = "Comments";
const TIMES_SHEET = "TimeSuggestions";

function getSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  }
  return sheet;
}

function ideasSheet() {
  return getSheet(IDEAS_SHEET, ["id", "name", "desc", "url", "datetime", "author", "created", "voters", "tags", "endtime"]);
}
function commentsSheet() {
  return getSheet(COMMENTS_SHEET, ["id", "ideaId", "author", "text", "created", "edited", "voters"]);
}
function timesSheet() {
  return getSheet(TIMES_SHEET, ["id", "ideaId", "author", "datetime", "created", "voters", "endtime", "reservation"]);
}

/** Read everything and return as JSON. */
function readAll() {
  const iRows = ideasSheet().getDataRange().getValues();
  const cRows = commentsSheet().getDataRange().getValues();
  const tRows = timesSheet().getDataRange().getValues();

  const ideas = [];
  for (let r = 1; r < iRows.length; r++) {
    const row = iRows[r];
    if (!row[0]) continue;
    const votersRaw = String(row[7] || "").trim();
    ideas.push({
      id: String(row[0]),
      name: String(row[1] || ""),
      desc: String(row[2] || ""),
      url: String(row[3] || ""),
      datetime: cellToDateTimeString(row[4]),
      author: String(row[5] || ""),
      created: Number(row[6]) || 0,
      voters: votersRaw ? votersRaw.split(",").map(s => s.trim()).filter(Boolean) : [],
      tags: String(row[8] || "").split(",").map(s => s.trim()).filter(Boolean),
      endtime: String(row[9] || "")
    });
  }

  const comments = [];
  for (let r = 1; r < cRows.length; r++) {
    const row = cRows[r];
    if (!row[0]) continue;
    const cVotersRaw = String(row[6] || "").trim();
    comments.push({
      id: String(row[0]),
      ideaId: String(row[1]),
      author: String(row[2] || ""),
      text: String(row[3] || ""),
      created: Number(row[4]) || 0,
      edited: String(row[5] || "") === "1",
      voters: cVotersRaw ? cVotersRaw.split(",").map(s => s.trim()).filter(Boolean) : []
    });
  }

  const times = [];
  for (let r = 1; r < tRows.length; r++) {
    const row = tRows[r];
    if (!row[0]) continue;
    const votersRaw = String(row[5] || "").trim();
    times.push({
      id: String(row[0]),
      ideaId: String(row[1]),
      author: String(row[2] || ""),
      datetime: cellToDateTimeString(row[3]),
      created: Number(row[4]) || 0,
      voters: votersRaw ? votersRaw.split(",").map(s => s.trim()).filter(Boolean) : [],
      endtime: String(row[6] || ""),
      reservation: String(row[7] || "") === "1"
    });
  }

  return { ideas: ideas, comments: comments, times: times };
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Normalize a datetime cell into a clean string.
 * - Text values (new entries) are returned verbatim: "2025-10-11" or "2025-10-11T00:00".
 * - Legacy Date values (rows Sheets auto-converted before we stored as text) can't tell
 *   date-only from midnight apart, so exact midnight becomes date-only.
 */
function cellToDateTimeString(v) {
  if (v === "" || v === null || v === undefined) return "";
  if (Object.prototype.toString.call(v) === "[object Date]") {
    const pad = n => (n < 10 ? "0" + n : "" + n);
    const datePart = v.getFullYear() + "-" + pad(v.getMonth() + 1) + "-" + pad(v.getDate());
    if (v.getHours() === 0 && v.getMinutes() === 0) return datePart;
    return datePart + "T" + pad(v.getHours()) + ":" + pad(v.getMinutes());
  }
  return String(v);
}

/** Write a value into a cell as plain text so Sheets never reinterprets it (e.g. as a date). */
function writeText(sheet, row, col, value) {
  const cell = sheet.getRange(row, col);
  cell.setNumberFormat("@");
  cell.setValue(value == null ? "" : value);
}

function doGet() {
  return jsonOut(readAll());
}

/** Find the row index (1-based sheet row) of an id in a sheet's first column. */
function findRow(sheet, id) {
  const values = sheet.getDataRange().getValues();
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][0]) === String(id)) return r + 1;
  }
  return -1;
}

/** Toggle an author inside a comma-separated voters cell. */
function toggleVoterCell(cell, author) {
  let voters = String(cell.getValue() || "").split(",").map(s => s.trim()).filter(Boolean);
  const idx = voters.indexOf(author);
  if (idx === -1) voters.push(author);
  else voters.splice(idx, 1);
  cell.setValue(voters.join(","));
}

/** POST → perform an action. Body is JSON (sent as text/plain to avoid CORS preflight). */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    if (action === "add") {
      const sheet = ideasSheet();
      const id = String(Date.now()) + Math.floor(Math.random() * 1000);
      // The author automatically upvotes their own idea (they can un-vote afterwards).
      sheet.appendRow([id, data.name, data.desc || "", data.url || "", "", data.author, Date.now(), data.author, data.tags || "", ""]);
      // Store datetime as text so Sheets doesn't convert a date-only value to midnight.
      writeText(sheet, sheet.getLastRow(), 5, data.datetime || "");
      writeText(sheet, sheet.getLastRow(), 10, data.endtime || "");
    }

    else if (action === "edit") {
      const sheet = ideasSheet();
      const row = findRow(sheet, data.id);
      if (row > 0 && String(sheet.getRange(row, 6).getValue()) === data.author) {
        sheet.getRange(row, 2).setValue(data.name);
        sheet.getRange(row, 3).setValue(data.desc || "");
        sheet.getRange(row, 4).setValue(data.url || "");
        writeText(sheet, row, 5, data.datetime || "");
        sheet.getRange(row, 9).setValue(data.tags || "");
        writeText(sheet, row, 10, data.endtime || "");
      }
    }

    else if (action === "delete") {
      const sheet = ideasSheet();
      const row = findRow(sheet, data.id);
      if (row > 0 && String(sheet.getRange(row, 6).getValue()) === data.author) {
        sheet.deleteRow(row);
        // Delete associated comments and time suggestions.
        const cSheet = commentsSheet();
        const cValues = cSheet.getDataRange().getValues();
        for (let r = cValues.length - 1; r >= 1; r--) {
          if (String(cValues[r][1]) === String(data.id)) cSheet.deleteRow(r + 1);
        }
        const tSheet = timesSheet();
        const tValues = tSheet.getDataRange().getValues();
        for (let r = tValues.length - 1; r >= 1; r--) {
          if (String(tValues[r][1]) === String(data.id)) tSheet.deleteRow(r + 1);
        }
      }
    }

    else if (action === "vote") {
      const sheet = ideasSheet();
      const row = findRow(sheet, data.id);
      if (row > 0) toggleVoterCell(sheet.getRange(row, 8), data.author);
    }

    else if (action === "comment") {
      const sheet = commentsSheet();
      const id = String(Date.now()) + Math.floor(Math.random() * 1000);
      sheet.appendRow([id, data.ideaId, data.author, data.text, Date.now(), "", ""]);
    }

    else if (action === "voteComment") {
      const sheet = commentsSheet();
      const row = findRow(sheet, data.id);
      if (row > 0) toggleVoterCell(sheet.getRange(row, 7), data.author);
    }

    else if (action === "deleteComment") {
      const sheet = commentsSheet();
      const row = findRow(sheet, data.id);
      if (row > 0 && String(sheet.getRange(row, 3).getValue()) === data.author) sheet.deleteRow(row);
    }

    else if (action === "editComment") {
      const sheet = commentsSheet();
      const row = findRow(sheet, data.id);
      if (row > 0 && String(sheet.getRange(row, 3).getValue()) === data.author) {
        sheet.getRange(row, 4).setValue(data.text);
        sheet.getRange(row, 6).setValue("1");   // mark as edited
      }
    }

    else if (action === "suggestTime") {
      const sheet = timesSheet();
      const id = String(Date.now()) + Math.floor(Math.random() * 1000);
      // Suggester automatically votes for their own suggestion.
      sheet.appendRow([id, data.ideaId, data.author, "", Date.now(), data.author, "", data.reservation ? "1" : ""]);
      // Store datetime as text so Sheets doesn't convert a date-only value to midnight.
      writeText(sheet, sheet.getLastRow(), 4, data.datetime || "");
      writeText(sheet, sheet.getLastRow(), 7, data.endtime || "");
    }

    else if (action === "editTime") {
      // Only the original suggester may edit their time suggestion (e.g. to toggle
      // "Reservation Made" or adjust the time).
      const sheet = timesSheet();
      const row = findRow(sheet, data.id);
      if (row > 0 && String(sheet.getRange(row, 3).getValue()) === data.author) {
        writeText(sheet, row, 4, data.datetime || "");
        writeText(sheet, row, 7, data.endtime || "");
        sheet.getRange(row, 8).setValue(data.reservation ? "1" : "");
      }
    }

    else if (action === "voteTime") {
      const sheet = timesSheet();
      const row = findRow(sheet, data.id);
      if (row > 0) toggleVoterCell(sheet.getRange(row, 6), data.author);
    }

    else if (action === "deleteTime") {
      const sheet = timesSheet();
      const row = findRow(sheet, data.id);
      if (row > 0 && String(sheet.getRange(row, 3).getValue()) === data.author) sheet.deleteRow(row);
    }

    return jsonOut(readAll());
  } catch (err) {
    return jsonOut({ error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
