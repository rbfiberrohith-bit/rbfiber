
const SHEET_NAME = "CloudData";

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("Open this script from your Google Sheet.");

  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);

  sh.getRange("A1:B1").setValues([["key", "jsonData"]]);
  PropertiesService.getScriptProperties()
    .setProperty("SPREADSHEET_ID", ss.getId());

  if (!sh.getRange("A2").getValue()) {
    sh.getRange("A2:B2").setValues([
      ["main", JSON.stringify({ complaints: [], materials: [] })]
    ]);
  }
}

function doGet() {
  try {
    const data = readCloudData();
    return json({ ok: true, data: data });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    if (!lock.tryLock(10000)) {
      return json({ ok: false, error: "Busy; try again" });
    }

    const body = JSON.parse(e.postData.contents || "{}");
    const data = body.data || body;

    if (!data || typeof data !== "object" || Array.isArray(data)) {
      return json({ ok: false, error: "Invalid data" });
    }

    const sh = getSheet();
    sh.getRange("A2:B2").setValues([
      ["main", JSON.stringify(data)]
    ]);

    return json({ ok: true, message: "Saved" });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function readCloudData() {
  const sh = getSheet();
  const raw = sh.getRange("B2").getValue();

  return raw
    ? JSON.parse(raw)
    : { complaints: [], materials: [] };
}

function getSheet() {
  const id = PropertiesService.getScriptProperties()
    .getProperty("SPREADSHEET_ID");

  if (!id) throw new Error("Run setup() first.");

  const ss = SpreadsheetApp.openById(id);
  const sh = ss.getSheetByName(SHEET_NAME);

  if (!sh) throw new Error("CloudData tab missing.");
  return sh;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
