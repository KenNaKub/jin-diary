const SHEET_NAME = "entries";

function doGet(e) {
  const action = e.parameter.action || "list";
  const callback = e.parameter.callback;
  const output = handleAction(action, e.parameter);
  return respond(output, callback);
}

function doPost(e) {
  const params = e.parameter || {};
  const body = e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
  const action = params.action || body.action || "append";
  const output = handleAction(action, { ...params, ...body });
  return respond(output, params.callback);
}

function handleAction(action, params) {
  if (action === "append") return appendEntry(params);
  if (action === "delete" || action === "remove" || action === "delete_entry") return deleteEntry(params);
  if (action === "dedupe" || action === "dedupe_entries") return dedupeEntriesInSheet();
  return listEntries();
}

function appendEntry(params) {
  const sheet = getSheet();
  const entry = typeof params.entry === "string" ? JSON.parse(params.entry || "{}") : params.entry || params;
  const row = [
    entry.id,
    entry.happenedAt,
    entry.type,
    entry.amount,
    entry.unit,
    entry.note,
    entry.createdAt || bangkokTimestamp(),
  ];
  const existingRows = findRowsById(sheet, entry.id);
  if (existingRows.length) {
    const latestRow = existingRows[existingRows.length - 1];
    sheet.getRange(latestRow, 1, 1, row.length).setValues([row]);
    for (let i = existingRows.length - 2; i >= 0; i--) {
      sheet.deleteRow(existingRows[i]);
    }
  } else {
    sheet.appendRow(row);
  }
  return { ok: true, entry };
}

function bangkokTimestamp() {
  return Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd'T'HH:mm:ss") + "+07:00";
}

function listEntries() {
  const sheet = getSheet();
  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1).filter((row) => row[0]);
  const seen = {};
  return {
    ok: true,
    entries: rows
      .reverse()
      .filter((row) => {
        const id = String(row[0]);
        if (seen[id]) return false;
        seen[id] = true;
        return true;
      })
      .map((row) => ({
        id: row[0],
        happenedAt: row[1],
        type: row[2],
        amount: row[3],
        unit: row[4],
        note: row[5],
      })),
  };
}

function deleteEntry(params) {
  const sheet = getSheet();
  const id = String(params.id || params.entryId || "");
  const values = sheet.getDataRange().getValues();
  let deleted = false;
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][0]) === id) {
      sheet.deleteRow(i + 1);
      deleted = true;
    }
  }
  if (deleted) return { ok: true, id };
  return { ok: false, id, error: "Entry not found" };
}

function dedupeEntriesInSheet() {
  const sheet = getSheet();
  const values = sheet.getDataRange().getValues();
  const seen = {};
  let deleted = 0;
  for (let i = values.length - 1; i >= 1; i--) {
    const id = String(values[i][0] || "");
    if (!id) continue;
    if (seen[id]) {
      sheet.deleteRow(i + 1);
      deleted++;
    } else {
      seen[id] = true;
    }
  }
  return { ok: true, deleted };
}

function findRowsById(sheet, id) {
  if (!id) return [];
  const values = sheet.getDataRange().getValues();
  const target = String(id);
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === target) rows.push(i + 1);
  }
  return rows;
}

function getSheet() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["id", "happenedAt", "type", "amount", "unit", "note", "createdAt"]);
  }
  return sheet;
}

function respond(payload, callback) {
  const body = callback ? `${callback}(${JSON.stringify(payload)})` : JSON.stringify(payload);
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JAVASCRIPT);
}
