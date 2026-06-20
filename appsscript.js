// Home Pantry — Apps Script backend
// Paste this into Extensions > Apps Script, then Deploy > New deployment > Web app
// Execute as: Me | Who has access: Anyone
// Copy the web app URL back into the Home Pantry app settings

const SHEET_NAME = 'Pantry';
const HEADERS = ['id', 'name', 'cat', 'unit', 'onHand', 'threshold', 'restock', 'updatedAt'];

function doGet() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const items = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    items.push({
      id:        data[i][0],
      name:      data[i][1],
      cat:       data[i][2],
      unit:      data[i][3],
      onHand:    data[i][4],
      threshold: data[i][5],
      restock:   data[i][6],
      updatedAt: data[i][7]
    });
  }
  return json({ items });
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const sheet = getSheet();

  if (body.action === 'save') {
    const it = body.item;
    it.updatedAt = new Date().toISOString();
    const rows = sheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === it.id) {
        sheet.getRange(i + 1, 1, 1, HEADERS.length).setValues([[
          it.id, it.name, it.cat, it.unit, it.onHand, it.threshold, it.restock, it.updatedAt
        ]]);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([it.id, it.name, it.cat, it.unit, it.onHand, it.threshold, it.restock, it.updatedAt]);
    }
    return json({ ok: true });
  }

  if (body.action === 'delete') {
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === body.id) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    return json({ ok: true });
  }

  return json({ error: 'unknown action' });
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sheet;
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
