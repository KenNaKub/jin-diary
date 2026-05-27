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
  return listEntries();
}

function appendEntry(params) {
  const sheet = getSheet();
  const entry = typeof params.entry === "string" ? JSON.parse(params.entry || "{}") : params.entry || params;
  sheet.appendRow([
    entry.id,
    entry.happenedAt,
    entry.type,
    entry.amount,
    entry.unit,
    entry.note,
    entry.createdAt || bangkokTimestamp(),
  ]);
  return { ok: true, entry };
}

function bangkokTimestamp() {
  return Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd'T'HH:mm:ss") + "+07:00";
}

function listEntries() {
  const sheet = getSheet();
  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1).filter((row) => row[0]);
  return {
    ok: true,
    entries: rows.map((row) => ({
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
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][0]) === id) {
      sheet.deleteRow(i + 1);
      return { ok: true, id };
    }
  }
  return { ok: false, id, error: "Entry not found" };
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
