const SPREADSHEET_ID = '1P4cn_jXjGWDk0fKA6P4HpbcyJX5QnJ12hLFVAs-iu_Q';
const SHEET_NAME = '주문';
const DEADLINE = new Date('2026-09-13T10:30:00+09:00');

function doGet(e) {
  const callback = (e && e.parameter && e.parameter.callback) || 'callback';
  const data = { orders: getOrders_(), closed: new Date() >= DEADLINE };
  return ContentService
    .createTextOutput(`${callback}(${JSON.stringify(data)})`)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function doPost(e) {
  try {
    if (new Date() >= DEADLINE) return json_({ ok:false, error:'closed' });
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action !== 'pick') return json_({ ok:false, error:'bad_action' });

    const name = String(body.name || '').trim().slice(0, 30);
    const menuId = String(body.menuId || '').trim();
    const menuName = String(body.menuName || '').trim().slice(0, 100);
    const price = Number(body.price || 0);
    if (!name || !menuId || !menuName || !price) return json_({ ok:false, error:'missing' });

    const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    const last = sh.getLastRow();
    if (last >= 2) {
      const names = sh.getRange(2, 1, last - 1, 1).getValues().flat().map(v => String(v).trim());
      const idx = names.findIndex(v => v === name);
      if (idx >= 0) {
        sh.getRange(idx + 2, 1, 1, 4).setValues([[name, menuName, price, new Date()]]);
        sh.getRange(idx + 2, 5).setValue(menuId);
        return json_({ ok:true, updated:true });
      }
    }
    sh.appendRow([name, menuName, price, new Date(), menuId]);
    return json_({ ok:true, updated:false });
  } catch (err) {
    return json_({ ok:false, error:String(err) });
  }
}

function getOrders_() {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const values = sh.getRange(2, 1, last - 1, 5).getValues();
  return values.filter(r => r[0]).map(r => ({
    name: String(r[0]),
    menuName: String(r[1]),
    price: Number(r[2] || 0),
    selectedAt: r[3] instanceof Date ? r[3].toISOString() : String(r[3] || ''),
    menuId: String(r[4] || '')
  }));
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
