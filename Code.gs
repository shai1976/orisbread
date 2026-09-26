/**
 * עדהלחם — backend להזמנות על Google Sheets
 *
 * לשונית Config (עמודה A = מפתח, עמודה B = ערך):
 *   maxOrders   — כמה הזמנות לפני שהמערכת נסגרת
 *   orderRound  — מזהה סבב (למשל 2026-10-02). משנים כל שבוע → המונה מתאפס
 *   open        — TRUE / FALSE. מתג ידני לסגירה או פתיחה
 *
 * לשונית Orders — שורה לכל הזמנה.
 */

const ORDERS = 'Orders';
const CONFIG = 'Config';
const HEADERS = ['מספר הזמנה', 'סבב', 'תאריך', 'שם', 'טלפון', 'מקום איסוף',
                 'פירוט', 'סה"כ ₪', 'סטטוס', 'items_json'];

/** להריץ פעם אחת מתוך העורך: יוצר את הלשוניות והכותרות */
function setup() {
  const ss = SpreadsheetApp.getActive();
  let o = ss.getSheetByName(ORDERS) || ss.insertSheet(ORDERS);
  o.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  o.setFrozenRows(1);
  o.setRightToLeft(true);

  let c = ss.getSheetByName(CONFIG) || ss.insertSheet(CONFIG);
  c.getRange(1, 1, 3, 2).setValues([
    ['maxOrders', 20],
    ['orderRound', '2026-10-02'],
    ['open', true]
  ]);
  c.getRange('B2').setNumberFormat('@'); // שהסבב יישאר טקסט ולא יהפוך לתאריך
  c.setRightToLeft(true);
}

function getConfig_() {
  const rows = SpreadsheetApp.getActive().getSheetByName(CONFIG).getDataRange().getValues();
  const cfg = {};
  rows.forEach(r => cfg[String(r[0]).trim()] = r[1]);
  return {
    maxOrders: Number(cfg.maxOrders) || 0,
    orderRound: String(cfg.orderRound).trim(),
    open: cfg.open === true || String(cfg.open).toUpperCase() === 'TRUE'
  };
}

function roundCount_(round) {
  const sh = SpreadsheetApp.getActive().getSheetByName(ORDERS);
  const last = sh.getLastRow();
  if (last < 2) return 0;
  return sh.getRange(2, 2, last - 1, 1).getValues()
           .filter(r => String(r[0]).trim() === round).length;
}

function status_() {
  const cfg = getConfig_();
  const count = roundCount_(cfg.orderRound);
  const left = Math.max(0, cfg.maxOrders - count);
  return { open: cfg.open && left > 0, left: left, round: cfg.orderRound };
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
                       .setMimeType(ContentService.MimeType.JSON);
}

/** GET ?action=status — מצב המלאי. לא חושף הזמנות */
function doGet(e) {
  return json_(status_());
}

/** POST — הזמנה חדשה. הבדיקה והכתיבה תחת נעילה, כדי ששתי הזמנות בו-זמנית לא יעברו את המגבלה */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return json_({ ok: false, reason: 'busy' });
  }
  try {
    const o = JSON.parse(e.postData.contents);
    const name = String(o.name || '').trim().slice(0, 80);
    const phone = String(o.phone || '').trim().slice(0, 20);
    const pickup = String(o.pickup || '').trim();
    const items = Array.isArray(o.items) ? o.items : [];
    if (!name || !pickup || !items.length) return json_({ ok: false, reason: 'invalid' });

    const s = status_();
    if (!s.open) return json_({ ok: false, reason: 'closed' });

    const sh = SpreadsheetApp.getActive().getSheetByName(ORDERS);
    const id = 'ADA-' + String(sh.getLastRow() + 1000);
    const detail = items.map(i => `${i.name} × ${i.qty}`).join('\n');
    const total = items.reduce((t, i) => t + Number(i.qty) * Number(i.price), 0);

    sh.appendRow([id, s.round, new Date(), name, "'" + phone, pickup,
                  detail, total, 'ממתין לתשלום', JSON.stringify(items)]);
    return json_({ ok: true, id: id, left: s.left - 1 });
  } catch (err) {
    return json_({ ok: false, reason: 'error' });
  } finally {
    lock.releaseLock();
  }
}
