const SPREADSHEET_ID = '1P4cn_jXjGWDk0fKA6P4HpbcyJX5QnJ12hLFVAs-iu_Q';
const SHEET_NAME = '주문';
const DEADLINE = new Date('2026-09-13T10:30:00+09:00');

function doGet(e) {
  const p = (e && e.parameter) || {};
  const callback = p.callback || 'callback';
  let data;

  try {
    if (p.action === 'pick') {
      data = savePick_(p);
    } else {
      data = { ok: true, orders: getOrders_(), closed: new Date() >= DEADLINE };
    }
  } catch (err) {
    data = { ok: false, error: String(err), orders: getOrders_(), closed: new Date() >= DEADLINE };
  }

  return ContentService
    .createTextOutput(`${callback}(${JSON.stringify(data)})`)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function savePick_(p) {
  if (new Date() >= DEADLINE) {
    return { ok: false, error: 'closed', orders: getOrders_(), closed: true };
  }

  const name = String(p.name || '').trim().slice(0, 30);
  const menuId = String(p.menuId || '').trim();
  const menuName = String(p.menuName || '').trim().slice(0, 100);
  const price = Number(p.price || 0);

  if (!name || !menuId || !menuName || !price) {
    return { ok: false, error: 'missing', orders: getOrders_(), closed: false };
  }

  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  const last = sh.getLastRow();
  let updated = false;

  if (last >= 2) {
    const names = sh.getRange(2, 1, last - 1, 1)
      .getValues()
      .flat()
      .map(v => String(v).trim());
    const idx = names.findIndex(v => v === name);

    if (idx >= 0) {
      sh.getRange(idx + 2, 1, 1, 5)
        .setValues([[name, menuName, price, new Date(), menuId]]);
      updated = true;
    }
  }

  if (!updated) {
    sh.appendRow([name, menuName, price, new Date(), menuId]);
  }

  SpreadsheetApp.flush();
  return {
    ok: true,
    updated,
    orders: getOrders_(),
    closed: false
  };
}

function getOrders_() {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  const last = sh.getLastRow();
  if (last < 2) return [];

  return sh.getRange(2, 1, last - 1, 5)
    .getValues()
    .filter(r => r[0])
    .map(r => ({
      name: String(r[0]),
      menuName: String(r[1]),
      price: Number(r[2] || 0),
      selectedAt: r[3] instanceof Date ? r[3].toISOString() : String(r[3] || ''),
      menuId: String(r[4] || '')
    }));
}
