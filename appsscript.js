// Home Pantry - Apps Script backend v3.1
// After editing: Deploy > Manage deployments > edit > New version > Deploy
// CLAUDE_API_KEY must be set in Project Settings > Script Properties

var SHEET_NAME = 'Pantry';
var HEADERS = ['id', 'name', 'cat', 'unit', 'onHand', 'threshold', 'restock', 'updatedAt'];
var ERRAND_SHEET = 'Errands';
var ERRAND_HEADERS = ['id', 'text', 'done', 'updatedAt'];

function doGet() {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  var items = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    items.push({
      id:        String(data[i][0]),
      name:      data[i][1],
      cat:       data[i][2],
      unit:      data[i][3],
      onHand:    Number(data[i][4]),
      threshold: Number(data[i][5]),
      restock:   Number(data[i][6]),
      updatedAt: data[i][7]
    });
  }

  var eSheet = getErrandSheet();
  var eData = eSheet.getDataRange().getValues();
  var errands = [];
  for (var ei = 1; ei < eData.length; ei++) {
    if (!eData[ei][0]) continue;
    errands.push({
      id:   String(eData[ei][0]),
      text: String(eData[ei][1]),
      done: eData[ei][2] === true
    });
  }

  return json({ items: items, errands: errands });
}

function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  var sheet = getSheet();

  // Save / update pantry item
  if (body.action === 'save') {
    var it = body.item;
    it.updatedAt = new Date().toISOString();
    var rows = sheet.getDataRange().getValues();
    var found = false;
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(it.id)) {
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

  // Clear all pantry items
  if (body.action === 'clearAll') {
    var last = sheet.getLastRow();
    if (last > 1) sheet.deleteRows(2, last - 1);
    return json({ ok: true });
  }

  // Delete pantry item
  if (body.action === 'delete') {
    var rows2 = sheet.getDataRange().getValues();
    for (var j = 1; j < rows2.length; j++) {
      if (String(rows2[j][0]) === String(body.id)) {
        sheet.deleteRow(j + 1);
        break;
      }
    }
    return json({ ok: true });
  }

  // Save / update errand
  if (body.action === 'saveErrand') {
    var er = body.errand;
    var eSheet2 = getErrandSheet();
    var eRows = eSheet2.getDataRange().getValues();
    var eFound = false;
    for (var ei2 = 1; ei2 < eRows.length; ei2++) {
      if (String(eRows[ei2][0]) === String(er.id)) {
        eSheet2.getRange(ei2 + 1, 1, 1, ERRAND_HEADERS.length).setValues([[
          er.id, er.text, er.done, new Date().toISOString()
        ]]);
        eFound = true;
        break;
      }
    }
    if (!eFound) {
      eSheet2.appendRow([er.id, er.text, er.done, new Date().toISOString()]);
    }
    return json({ ok: true });
  }

  // Delete errand
  if (body.action === 'deleteErrand') {
    var eSheet3 = getErrandSheet();
    var eRows2 = eSheet3.getDataRange().getValues();
    for (var ej = 1; ej < eRows2.length; ej++) {
      if (String(eRows2[ej][0]) === String(body.id)) {
        eSheet3.deleteRow(ej + 1);
        break;
      }
    }
    return json({ ok: true });
  }

  // Voice stocktake
  if (body.action === 'parseStocktake') {
    var apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
    if (!apiKey) return json({ error: 'CLAUDE_API_KEY not set in Script Properties' });

    var pantryLines = [];
    var pantryItems = body.pantryItems || [];
    for (var p = 0; p < pantryItems.length; p++) {
      pantryLines.push(pantryItems[p].id + '|' + pantryItems[p].name + ' (' + pantryItems[p].unit + ')');
    }

    var voicePrompt = 'You are managing a NZ household pantry stocktake. Match spoken items to the pantry list and extract quantities.\n'
      + 'PANTRY ITEMS:\n' + pantryLines.join('\n') + '\n'
      + 'TRANSCRIPT: "' + body.transcript + '"\n'
      + 'Rules: match each spoken item to closest pantry item name. '
      + 'Quantities: specific number = use it, half = 0.5, nearly empty = 0.5, none/empty/out = 0, full/plenty = 2, a few = 3, heaps = 5. '
      + 'Only include items you can confidently match AND extract a quantity for. '
      + 'Return ONLY a valid JSON array: [{"id":"...","onHand":0}] - no explanation, no markdown.';

    var voicePayload = JSON.stringify({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 1024,
      messages: [{ role: 'user', content: voicePrompt }]
    });

    try {
      var vRes = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
        method: 'post',
        headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01', 'content-type': 'application/json' },
        payload: voicePayload,
        muteHttpExceptions: true
      });
      var vResult = JSON.parse(vRes.getContentText());
      if (vResult.error) return json({ error: vResult.error.message });
      var vText = vResult.content[0].text.trim();
      var vStart = vText.indexOf('[');
      var vEnd = vText.lastIndexOf(']');
      if (vStart === -1) return json({ updates: [] });
      return json({ updates: JSON.parse(vText.slice(vStart, vEnd + 1)) });
    } catch(err) {
      return json({ error: err.message });
    }
  }

  // Camera scan
  if (body.action === 'scanImage') {
    var apiKey2 = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
    if (!apiKey2) return json({ error: 'CLAUDE_API_KEY not set in Script Properties' });

    var scanPrompt = 'You are scanning a NZ household pantry or fridge. '
      + 'Identify every grocery item and match it to woolworths.co.nz or paknsave.co.nz naming. '
      + 'Distinguish variants: White vs Brown vs Basmati Rice, Trim vs Whole Milk, Streaky vs Middle Bacon, Cheddar vs Colby vs Edam. '
      + 'Return 4 fields per item: '
      + 'name (product description, no brand - e.g. Trim Milk, Long Grain White Rice, Free Range Eggs 6pk), '
      + 'brand (brand if visible e.g. Pams, Anchor, Watties, Mainland - use empty string if unknown), '
      + 'cat (one of: Dairy / Meat & Fish / Produce / Pantry / Frozen / Bakery / Drinks / Cleaning / Personal Care / Other), '
      + 'unit (L for liquid dairy, kg for meat/produce/rice/flour, each for eggs, pack for multipack, can for canned goods, loaf for bread, bottle for drinks). '
      + 'Return ONLY a valid JSON array, no explanation, no markdown. '
      + 'Example: [{"name":"Trim Milk","brand":"Anchor","cat":"Dairy","unit":"L"}]';

    var scanPayload = JSON.stringify({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 1024,
      messages: [{
        role: 'user',
        content: [
          { type: 'image', source: { type: 'base64', media_type: body.mediaType, data: body.image } },
          { type: 'text', text: scanPrompt }
        ]
      }]
    });

    try {
      var sRes = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
        method: 'post',
        headers: { 'x-api-key': apiKey2, 'anthropic-version': '2023-06-01', 'content-type': 'application/json' },
        payload: scanPayload,
        muteHttpExceptions: true
      });
      var sResult = JSON.parse(sRes.getContentText());
      if (sResult.error) return json({ error: sResult.error.message });
      var sText = sResult.content[0].text.trim();
      var sStart = sText.indexOf('[');
      var sEnd = sText.lastIndexOf(']');
      if (sStart === -1) return json({ items: [], raw: sText });
      return json({ items: JSON.parse(sText.slice(sStart, sEnd + 1)) });
    } catch(err) {
      return json({ error: err.message });
    }
  }

  return json({ error: 'unknown action' });
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sheet;
}

function getErrandSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ERRAND_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(ERRAND_SHEET);
    sheet.appendRow(ERRAND_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, ERRAND_HEADERS.length).setFontWeight('bold');
  }
  return sheet;
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
