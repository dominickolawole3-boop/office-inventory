// ── OFFICE INVENTORY · Google Apps Script Backend ─────────────────────────
// Paste this entire file into your Apps Script editor (script.google.com)
// Then deploy as a Web App (Execute as: Me, Who has access: Anyone)

const SHEET_NAME = "Inventory";
const HEADERS = ["id", "name", "cat", "loc", "qty", "max", "min"];

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    // Seed with sample data
    const samples = [
      [1, "A4 Paper, 80gsm",           "Paper & Print", "Shelf B2",       230, 500, 150],
      [2, "Black Ballpoint Pens (box)", "Writing",       "Drawer A1",      32, 50, 10],
      [3, "Ground Coffee, 1kg",         "Kitchen",       "Pantry",         3, 15, 4],
      [4, "Disinfectant Wipes",         "Cleaning",      "Cleaning Closet", 18, 30, 8],
      [5, "USB-C Cables",              "Tech",          "IT Cabinet",      6, 25, 6],
      [6, "Padded Mailers, Medium",    "Mailing",       "Shelf C1",        40, 80, 15],
    ];
    samples.forEach(r => sheet.appendRow(r));
  }
  return sheet;
}

// ── READ: return all rows as JSON ──────────────────────────────────────────
function doGet(e) {
  const sheet = getSheet();
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const items = rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, items }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── WRITE: add / edit / delete ─────────────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const { action, item } = data;
    const sheet = getSheet();

    if (action === "add") {
      const rows = sheet.getDataRange().getValues();
      const ids = rows.slice(1).map(r => Number(r[0])).filter(Boolean);
      const newId = ids.length ? Math.max(...ids) + 1 : 1;
      sheet.appendRow([newId, item.name, item.cat, item.loc, item.qty, item.max, item.min]);
      return ok({ id: newId });
    }

    if (action === "edit") {
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(item.id)) {
          sheet.getRange(i + 1, 1, 1, 7).setValues([[
            item.id, item.name, item.cat, item.loc, item.qty, item.max, item.min
          ]]);
          return ok({});
        }
      }
      return err("Item not found");
    }

    if (action === "delete") {
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(item.id)) {
          sheet.deleteRow(i + 1);
          return ok({});
        }
      }
      return err("Item not found");
    }

    return err("Unknown action");
  } catch (err) {
    return error(err.toString());
  }
}

function ok(data) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, ...data }))
    .setMimeType(ContentService.MimeType.JSON);
}

function err(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}
