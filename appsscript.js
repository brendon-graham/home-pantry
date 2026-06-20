// Home Pantry — Apps Script backend v2
// Deploy > Manage deployments > edit existing deployment > redeploy (same URL)
// For camera scan: Project Settings > Script Properties > add CLAUDE_API_KEY

const SHEET_NAME = 'Pantry';
const HEADERS = ['id', 'name', 'cat', 'unit', 'onHand', 'threshold', 'restock', 'updatedAt'];

function doGet() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const items = [];
  for (let i = 1; i < data.length; i++) {
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
  return json({ items });
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const sheet = getSheet();

  // ── Save / update item ──────────────────────────────
  if (body.action === 'save') {
    const it = body.item;
    it.updatedAt = new Date().toISOString();
    const rows = sheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < rows.length; i++) {
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

  // ── Clear all items ─────────────────────────────────
  if (body.action === 'clearAll') {
    const last = sheet.getLastRow();
    if (last > 1) sheet.deleteRows(2, last - 1);
    return json({ ok: true });
  }

  // ── Delete item ─────────────────────────────────────
  if (body.action === 'delete') {
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(body.id)) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    return json({ ok: true });
  }

  // ── Camera scan via Claude ──────────────────────────
  if (body.action === 'scanImage') {
    const apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
    if (!apiKey) return json({ error: 'CLAUDE_API_KEY not set in Script Properties' });

    const payload = {
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 1024,
      messages: [{
        role: 'user',
        content: [
          {
            type: 'image',
            source: { type: 'base64', media_type: body.mediaType, data: body.image }
          },
          {
            type: 'text',
            text: 'You are scanning a New Zealand household pantry or fridge. Your job is to identify every grocery item and match it as closely as possible to how it appears on woolworths.co.nz or paknsave.co.nz. Specificity of product type is the top priority — distinguish between variants like White Long Grain Rice vs Brown Rice vs Basmati Rice, Trim Milk vs Whole Milk vs Light Blue Milk, Streaky Bacon vs Middle Bacon, Cheddar vs Colby vs Edam. For each item return 4 fields: name (the specific product description as Woolworths NZ or Pak\'nSave would list it — e.g. "Long Grain White Rice", "Trim Milk", "Baked Beans in Tomato Sauce", "Wholemeal Sandwich Bread", "Free Range Eggs 6pk", "Streaky Bacon Rashers", "Colby Cheese Block", "Greek Style Yoghurt", "Dishwasher Tablets" — NO brand in this field), brand (the brand name if visible or reasonably identifiable from the packaging — e.g. "Pams", "Anchor", "Wattie\'s", "Mainland", "Tegel", "Tip Top", "Homebrand" — use empty string "" if unknown), cat (one of: Dairy / Meat & Fish / Produce / Pantry / Frozen / Bakery / Drinks / Cleaning / Personal Care / Other), unit (use the unit Woolworths NZ uses: L for liquid dairy, kg for meat/produce/rice/flour, each for eggs/individual items, pack for multipack, can for canned goods, loaf for bread, bottle for drinks). Return ONLY a valid JSON array, no explanation, no markdown, no code fences. Example: [{"name":"Trim Milk","brand":"Anchor","cat":"Dairy","unit":"L"},{"name":"Long Grain White Rice","brand":"Pams","cat":"Pantry","unit":"kg"},{"name":"Free Range Eggs 6pk","brand":"","cat":"Dairy","unit":"each"}]'
          }
        ]
      }]
    };

    try {
      const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
        method: 'post',
        headers: {
          'x-api-key': apiKey,
          'anthropic-version': '2023-06-01',
          'content-type': 'application/json'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });

      const result = JSON.parse(response.getContentText());
      if (result.error) return json({ error: result.error.message });

      const text = result.content[0].text.trim();
      const start = text.indexOf('[');
      const end = text.lastIndexOf(']');
      if (start === -1) return json({ items: [], raw: text });
      const scannedItems = JSON.parse(text.slice(start, end + 1));
      return json({ items: scannedItems });
    } catch (err) {
      return json({ error: err.message });
    }
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
