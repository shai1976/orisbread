/**
 * עדהלחם — backend מהיר להזמנות על Google Sheets
 *
 * לשונית Config (עמודה A = מפתח, עמודה B = ערך):
 *   maxOrders   — כמה הזמנות לפני שהמערכת נסגרת
 *   orderRound  — מזהה סבב (למשל 2026-10-02). משנים כל שבוע → המונה מתאפס
 *   open        — TRUE / FALSE. מתג ידני לסגירה או פתיחה
 *
 * לשונית Orders — שורה לכל הזמנה.
 *
 * השיפור המרכזי: ספירת ההזמנות נשמרת ב-Script Properties במקום לסרוק
 * את כל גיליון Orders בכל GET/POST. בסבב חדש הספירה נבנית פעם אחת בלבד.
 */

const ORDERS = 'Orders';
const CONFIG = 'Config';
const COUNT_PREFIX = 'roundCount:';
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
  c.getRange('B2').setNumberFormat('@');
  c.setRightToLeft(true);

  syncStockCounter();
}

function getConfig_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG);
  if (!sh) throw new Error('Missing Config sheet');

  // רק שלוש השורות הדרושות, בלי getDataRange על גיליון שלם.
  const rows = sh.getRange(1, 1, 3, 2).getValues();
  const cfg = {};
  rows.forEach(r => cfg[String(r[0]).trim()] = r[1]);

  return {
    maxOrders: Number(cfg.maxOrders) || 0,
    orderRound: String(cfg.orderRound).trim(),
    open: cfg.open === true || String(cfg.open).toUpperCase() === 'TRUE'
  };
}

function countKey_(round) {
  return COUNT_PREFIX + round;
}

/** סריקה מלאה — מתבצעת רק אם אין עדיין מונה שמור לסבב */
function countRoundFromSheet_(round) {
  const sh = SpreadsheetApp.getActive().getSheetByName(ORDERS);
  if (!sh) return 0;
  const last = sh.getLastRow();
  if (last < 2) return 0;

  const values = sh.getRange(2, 2, last - 1, 1).getValues();
  let count = 0;
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === round) count++;
  }
  return count;
}

function getRoundCount_(round) {
  const props = PropertiesService.getScriptProperties();
  const key = countKey_(round);
  const saved = props.getProperty(key);
  if (saved !== null) return Number(saved) || 0;

  const count = countRoundFromSheet_(round);
  props.setProperty(key, String(count));
  return count;
}

function setRoundCount_(round, count) {
  PropertiesService.getScriptProperties().setProperty(countKey_(round), String(count));
}

/**
 * להריץ ידנית רק אם ערכת/מחקת שורות ישירות ב-Orders ורוצים לסנכרן את המונה.
 * לא צריך להריץ כשפותחים סבב חדש — הוא מאותחל אוטומטית.
 */
function syncStockCounter() {
  const cfg = getConfig_();
  const count = countRoundFromSheet_(cfg.orderRound);
  setRoundCount_(cfg.orderRound, count);
  return { round: cfg.orderRound, count: count };
}

function status_() {
  const cfg = getConfig_();
  const count = getRoundCount_(cfg.orderRound);
  const left = Math.max(0, cfg.maxOrders - count);
  return { open: cfg.open && left > 0, left: left, round: cfg.orderRound };
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
                       .setMimeType(ContentService.MimeType.JSON);
}

/** GET ?action=status — מצב המלאי. לא חושף הזמנות */
function doGet(e) {
  try {
    return json_(status_());
  } catch (err) {
    console.error(err);
    return json_({ ok: false, reason: 'error' });
  }
}

/**
 * POST — הזמנה חדשה.
 * הבדיקה, הכתיבה ועדכון המונה תחת נעילה כדי למנוע חריגה מהמלאי.
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    // ברוב המקרים הנעילה מתקבלת מייד; לא מחכים 10 שניות במקרה של עומס.
    if (!lock.tryLock(2500)) return json_({ ok: false, reason: 'busy' });

    const o = JSON.parse(e.postData && e.postData.contents ? e.postData.contents : '{}');
    const name = String(o.name || '').trim().slice(0, 80);
    const phone = String(o.phone || '').trim().slice(0, 20);
    const pickup = String(o.pickup || '').trim().slice(0, 100);
    const items = Array.isArray(o.items) ? o.items : [];
    if (!name || !pickup || !items.length) return json_({ ok: false, reason: 'invalid' });

    const cfg = getConfig_();
    const count = getRoundCount_(cfg.orderRound);
    const left = Math.max(0, cfg.maxOrders - count);
    if (!cfg.open || left <= 0) return json_({ ok: false, reason: 'closed' });

    const sh = SpreadsheetApp.getActive().getSheetByName(ORDERS);
    if (!sh) throw new Error('Missing Orders sheet');

    const row = sh.getLastRow() + 1;
    const id = 'ADA-' + String(row + 999);
    const detail = items.map(i => `${String(i.name || '').trim()} × ${Number(i.qty) || 0}`).join('\n');
    const total = items.reduce((t, i) => t + (Number(i.qty) || 0) * (Number(i.price) || 0), 0);

    // כתיבה יחידה לטווח מדויק במקום appendRow.
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

    setRoundCount_(cfg.orderRound, count + 1);
    return json_({ ok: true, id: id, left: left - 1 });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, reason: 'error' });
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}
