/**
 * עדהלחם — backend להזמנות + מלאי כיכרות בלבד
 *
 * לשונית Config (עמודה A = מפתח, עמודה B = ערך):
 *   loavesLeft  — מספר הכיכרות שנותרו כרגע
 *   orderRound  — מזהה סבב, למשל 2026-10-02
 *   open        — TRUE / FALSE לפתיחה/סגירה ידנית
 *
 * אין תלות במספר ההזמנות ואין ספירה של שורות ב-Orders.
 * כל הזמנה מפחיתה רק את מספר הכיכרות שהוזמנו מ-loavesLeft.
 */

const ORDERS = 'Orders';
const CONFIG = 'Config';
const EXTRA_PRICE = 3;
const ALLOWED_EXTRAS = ['זיתים', 'אגוזים', 'פרג', 'שומשום'];
const HEADERS = ['מספר הזמנה', 'סבב', 'תאריך', 'שם', 'טלפון', 'מקום איסוף',
                 'פירוט', 'סה"כ ₪', 'סטטוס', 'items_json'];

function setup() {
  const ss = SpreadsheetApp.getActive();

  let o = ss.getSheetByName(ORDERS) || ss.insertSheet(ORDERS);
  o.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  o.setFrozenRows(1);
  o.setRightToLeft(true);

  let c = ss.getSheetByName(CONFIG) || ss.insertSheet(CONFIG);
  c.clear();
  c.getRange(1, 1, 3, 2).setValues([
    ['loavesLeft', 17],
    ['orderRound', '2026-10-02'],
    ['open', true]
  ]);
  c.getRange('B2').setNumberFormat('@');
  c.setRightToLeft(true);
}

function getConfig_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG);
  if (!sh) throw new Error('Missing Config sheet');

  const last = Math.max(sh.getLastRow(), 1);
  const rows = sh.getRange(1, 1, last, 2).getValues();
  const cfg = {};
  const rowMap = {};

  rows.forEach((r, i) => {
    const key = String(r[0] || '').trim();
    if (!key) return;
    cfg[key] = r[1];
    rowMap[key] = i + 1;
  });

  // תאימות זמנית: אם loavesLeft עדיין לא קיים, השתמש ב-maxLoaves כערך התחלתי.
  const leftRaw = cfg.loavesLeft !== undefined && cfg.loavesLeft !== ''
    ? cfg.loavesLeft
    : cfg.maxLoaves;

  return {
    sheet: sh,
    rowMap: rowMap,
    loavesLeft: Math.max(0, Math.floor(Number(leftRaw) || 0)),
    orderRound: String(cfg.orderRound || '').trim(),
    open: cfg.open === true || String(cfg.open).toUpperCase() === 'TRUE'
  };
}

function ensureLoavesLeftRow_(cfg) {
  if (cfg.rowMap.loavesLeft) return cfg.rowMap.loavesLeft;

  const row = cfg.sheet.getLastRow() + 1;
  cfg.sheet.getRange(row, 1, 1, 2).setValues([['loavesLeft', cfg.loavesLeft]]);
  cfg.rowMap.loavesLeft = row;
  return row;
}

function setLoavesLeft_(cfg, value) {
  const row = ensureLoavesLeftRow_(cfg);
  cfg.sheet.getRange(row, 2).setValue(Math.max(0, Math.floor(Number(value) || 0)));
}

function itemQty_(item) {
  return Math.max(0, Math.floor(Number(item && item.qty) || 0));
}

function status_() {
  const cfg = getConfig_();
  const left = cfg.loavesLeft;
  return {
    ok: true,
    open: cfg.open && left > 0,
    left: left,
    round: cfg.orderRound
  };
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    return json_(status_());
  } catch (err) {
    console.error(err);
    return json_({ ok: false, reason: 'error' });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    if (!lock.tryLock(2500)) {
      return json_({ ok: false, reason: 'busy' });
    }

    const o = JSON.parse(e.postData && e.postData.contents ? e.postData.contents : '{}');
    const name = String(o.name || '').trim().slice(0, 80);
    const phone = String(o.phone || '').trim().slice(0, 20);
    const pickup = String(o.pickup || '').trim().slice(0, 100);
    const items = Array.isArray(o.items) ? o.items : [];
    const requested = items.reduce((sum, item) => sum + itemQty_(item), 0);

    if (!name || !pickup || !items.length || requested <= 0) {
      return json_({ ok: false, reason: 'invalid' });
    }

    const cfg = getConfig_();
    const left = cfg.loavesLeft;

    if (!cfg.open || left <= 0 || requested > left) {
      return json_({ ok: false, reason: 'closed', left: left });
    }

    const sh = SpreadsheetApp.getActive().getSheetByName(ORDERS);
    if (!sh) throw new Error('Missing Orders sheet');

    const row = sh.getLastRow() + 1;
    const id = 'ADA-' + String(row + 999);
    const detail = items
      .flatMap(i => {
        const name = String(i.name || '').trim();
        const qty = itemQty_(i);
        const sliced = Array.isArray(i.sliced) ? i.sliced : [];
        const extras = Array.isArray(i.extras) ? i.extras : [];
        return Array.from({ length: qty }, (_, n) => {
          const selected = Array.isArray(extras[n])
            ? extras[n].map(x => String(x || '').trim()).filter(x => ALLOWED_EXTRAS.includes(x))
            : [];
          const parts = [sliced[n] ? 'פרוס' : 'לא פרוס'];
          if (selected.length) parts.push('תוספות: ' + selected.join(', '));
          return `${name}${qty > 1 ? ` — כיכר ${n + 1}` : ''}: ${parts.join(' · ')}`;
        });
      })
      .join('\n');
    const total = items.reduce((t, i) => {
      const qty = itemQty_(i);
      const base = qty * (Number(i.price) || 0);
      const extras = Array.isArray(i.extras) ? i.extras : [];
      const extrasCount = Array.from({ length: qty }, (_, n) =>
        Array.isArray(extras[n])
          ? extras[n].map(x => String(x || '').trim()).filter(x => ALLOWED_EXTRAS.includes(x)).length
          : 0
      ).reduce((a, b) => a + b, 0);
      return t + base + extrasCount * EXTRA_PRICE;
    }, 0);

    sh.getRange(row, 1, 1, HEADERS.length).setValues([[
      id,
      cfg.orderRound,
      new Date(),
      name,
      "'" + phone,
      pickup,
      detail,
      total,
      'ממתין לתשלום',
      JSON.stringify(items)
    ]]);

    const newLeft = left - requested;
    setLoavesLeft_(cfg, newLeft);

    return json_({ ok: true, id: id, left: newLeft });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, reason: 'error' });
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}
