// ════════════════════════════════════════════════════════
//  Shop-CRM × Base10-CRM  —  Google Apps Script API
//  部署方式：擴充功能 → Apps Script → 部署 → 新增部署
//  類型選「網頁應用程式」，執行身分「我」，存取「所有人」
// ════════════════════════════════════════════════════════

const SS = SpreadsheetApp.getActiveSpreadsheet();

// 工作表名稱對應
const SHEETS = {
  partners: '夥伴_CRM',   // index.html 用
  retail:   '零售_CRM'    // base10.html 用
};

// 夥伴_CRM 欄位順序
const PARTNER_COLS = [
  'id','name','channel','grade','status',
  'first','last','next','concern','note','method','updated'
];

// 零售_CRM 欄位順序
const RETAIL_COLS = [
  'id','name','source','status','vip','products','productsNote',
  'firstDate','lastPurchase','next','goal','method','concern','note','updated'
];

// ── CORS helper ──────────────────────────────────────────
function cors(output) {
  return output
    .setMimeType(ContentService.MimeType.JSON)
    .addHeader('Access-Control-Allow-Origin', '*')
    .addHeader('Access-Control-Allow-Methods', 'GET,POST')
    .addHeader('Access-Control-Allow-Headers', 'Content-Type');
}

function ok(data) {
  return cors(ContentService.createTextOutput(
    JSON.stringify({ ok: true, data })
  ));
}

function err(msg) {
  return cors(ContentService.createTextOutput(
    JSON.stringify({ ok: false, error: msg })
  ));
}

// ── 取得或建立工作表，並確保第一列是欄位標題 ────────────
function getSheet(name, cols) {
  let sheet = SS.getSheetByName(name);
  if (!sheet) {
    sheet = SS.insertSheet(name);
    sheet.getRange(1, 1, 1, cols.length).setValues([cols]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function colsFor(table) {
  return table === 'partners' ? PARTNER_COLS : RETAIL_COLS;
}

// ── 讀取所有資料 ─────────────────────────────────────────
function readAll(table) {
  const cols = colsFor(table);
  const sheet = getSheet(SHEETS[table], cols);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];           // 只有標題列
  const header = rows[0];
  return rows.slice(1).map(row => {
    const obj = {};
    header.forEach((h, i) => { obj[h] = row[i] === '' ? '' : String(row[i]); });
    // products 欄位還原成陣列
    if (obj.products) {
      try { obj.products = JSON.parse(obj.products); }
      catch(e) { obj.products = obj.products ? obj.products.split(',') : []; }
    }
    return obj;
  });
}

// ── 寫入整批資料（覆蓋舊資料，保留標題） ────────────────
function writeAll(table, records) {
  const cols = colsFor(table);
  const sheet = getSheet(SHEETS[table], cols);

  // 清除資料列（保留標題）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, cols.length).clearContent();
  }

  if (!records || records.length === 0) return;

  const rows = records.map(r =>
    cols.map(col => {
      const v = r[col];
      if (col === 'products' && Array.isArray(v)) return JSON.stringify(v);
      return v === undefined ? '' : v;
    })
  );

  sheet.getRange(2, 1, rows.length, cols.length).setValues(rows);
}

// ── GET 請求：讀資料 ─────────────────────────────────────
function doGet(e) {
  try {
    const table = e.parameter.table;
    if (!SHEETS[table]) return err('unknown table: ' + table);
    return ok(readAll(table));
  } catch(ex) {
    return err(ex.message);
  }
}

// ── POST 請求：寫資料 ────────────────────────────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const { table, records } = payload;
    if (!SHEETS[table]) return err('unknown table: ' + table);
    writeAll(table, records);
    return ok({ saved: records.length });
  } catch(ex) {
    return err(ex.message);
  }
}
