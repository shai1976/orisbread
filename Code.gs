/**
 * עדהלחם — backend להזמנות + מלאי כיכרות
 *
 * לשונית Config (עמודה A = מפתח, עמודה B = ערך):
 *   maxLoaves   — מספר הכיכרות הזמינות בסבב
 *   orderRound  — מזהה סבב, למשל 2026-10-02
 *   open        — TRUE / FALSE לפתיחה/סגירה ידנית
 *
 * תאימות לאחור: אם maxLoaves לא קיים, המערכת תשתמש ב-maxOrders.
 */

const ORDERS = 'Orders';
const CONFIG = 'Config';
const COUNT_PREFIX = 'roundLoaves:';
const HEADERS = ['מספר הזמנה', 'סבב', 'תאריך', 'שם', 'טלפון', 'מקום איסוף',
                 'פירוט', 'סה"כ ₪', 'סטטוס', 'items_json'];

function setup() {
  const ss = SpreadsheetApp.getActive();
  let o = ss.getSheetByName(ORDERS) || ss.insertSheet(ORDERS);
  o.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  o.setFrozenRows(1);
  o.setRightToLeft(true);

  let c = ss.getSheetByName(CONFIG) || ss.insertSheet(CONFIG);
  c.getRange(1, 1, 3, 2).setValues([
    ['maxLoaves', 20],
    ['orderRound', '2026-10-02'],
    ['open', true]
  ]);
  c.getRange('B2').setNumberFormat('@');
  c.setRightToLeft(true);
  syncStockCounter();
}

function getConfig_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG);
  if (!sh) throw new Error('Missing Config sheet');

  const last = Math.max(sh.getLastRow(), 1);
  const rows = sh.getRange(1, 1, last, 2).getValues();
  const cfg = {};
  rows.forEach(r => cfg[String(r[0]).trim()] = r[1]);

  const maxRaw = cfg.maxLoaves !== undefined && cfg.maxLoaves !== '' ? cfg.maxLoaves : cfg.maxOrders;
  return {
    maxLoaves: Math.max(0, Number(maxRaw) || 0),
    orderRound: String(cfg.orderRound || '').trim(),
    open: cfg.open === true || String(cfg.open).toUpperCase() === 'TRUE'
  };
}

function countKey_(round) {
  return COUNT_PREFIX + round;
}

function itemQty_(item) {
  return Math.max(0, Math.floor(Number(item && item.qty) || 0));
}

/** סופר כיכרות קיימות בסבב מתוך items_json. */
function countLoavesFromSheet_(round) {
  const sh = SpreadsheetApp.getActive().getSheetByName(ORDERS);
  if (!sh) return 0;
  const last = sh.getLastRow();
  if (last < 2) return 0;

  // B = סבב, J = items_json
  const rows = sh.getRange(2, 2, last - 1, 9).getValues();
  let count = 0;
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() !== round) continue;
    try {
      const items = JSON.parse(rows[i][8] || '[]');
      if (Array.isArray(items)) count += items.reduce((sum, item) => sum + itemQty_(item), 0);
    } catch (_) {
      // הזמנה ישנה ללא JSON תקין: נספרת ככיכר אחת כדי לא לפתוח מלאי בטעות.
      count += 1;
    }
  }
  return count;
}

function getRoundCount_(round) {
  const props = PropertiesService.getScriptProperties();
  const key = countKey_(round);
  const saved = props.getProperty(key);
  if (saved !== null) return Number(saved) || 0;

  const count = countLoavesFromSheet_(round);
  props.setProperty(key, String(count));
  return count;
}

function setRoundCount_(round, count) {
  PropertiesService.getScriptProperties().setProperty(countKey_(round), String(count));
}

/** להריץ ידנית אם מוחקים/עורכים הזמנות ישירות בגיליון. */
function syncStockCounter() {
  const cfg = getConfig_();
  const count = countLoavesFromSheet_(cfg.orderRound);
  setRoundCount_(cfg.orderRound, count);
  return { round: cfg.orderRound, loaves: count, left: Math.max(0, cfg.maxLoaves - count) };
}

function status_() {
  const cfg = getConfig_();
  const count = getRoundCount_(cfg.orderRound);
  const left = Math.max(0, cfg.maxLoaves - count);
  return { ok: true, open: cfg.open && left > 0, left: left, round: cfg.orderRound };
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
    if (!lock.tryLock(2500)) return json_({ ok: false, reason: 'busy' });

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
    const count = getRoundCount_(cfg.orderRound);
    const left = Math.max(0, cfg.maxLoaves - count);

    // בודק מלאי בפועל לפני הכתיבה. הזמנה גדולה מהמלאי הנותר נדחית כולה.
    if (!cfg.open || left <= 0 || requested > left) {
      return json_({ ok: false, reason: 'closed', left: left });
    }

    const sh = SpreadsheetApp.getActive().getSheetByName(ORDERS);
    if (!sh) throw new Error('Missing Orders sheet');

    const row = sh.getLastRow() + 1;
    const id = 'ADA-' + String(row + 999);
    const detail = items.map(i => `${String(i.name || '').trim()} × ${itemQty_(i)}`).join('\n');
    const total = items.reduce((t, i) => t + itemQty_(i) * (Number(i.price) || 0), 0);

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

    setRoundCount_(cfg.orderRound, count + requested);
    return json_({ ok: true, id: id, left: left - requested });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, reason: 'error' });
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}
