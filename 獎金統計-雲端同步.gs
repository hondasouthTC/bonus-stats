/**
 * 獎金統計 — 雲端同步後端 v2（Google Apps Script）
 *
 * 安裝：
 * 1. 空白 Google 試算表 → 擴充功能 → Apps Script → 整份貼進 Code.gs（覆蓋原內容）
 * 2. 部署 → 管理部署作業 → 編輯（鉛筆）→ 版本選「新版本」→ 部署
 *    （網址不變，同事已經貼的那個 /exec 不用換）
 * 3. 首次可手動執行 setup() 一次，建立分頁。
 *
 * 資料存兩份：
 *   SyncBlob 分頁 → 壓縮資料（App 讀寫用）
 *   資料 分頁     → 與原試算表相同的 60 欄明細（給人看的，每次同步重寫）
 *
 * 多人同時寫入時以「單筆紀錄」為單位合併，較新的 upd 覆蓋舊的，不會整批蓋掉。
 */

const CHUNK = 40000;
const DATA_SHEET = '資料';
const COLS = ['時間戳記','業務員','月份','車款','顏色規格','客戶姓名','狀態','獎金','SP固定','中古車','貸款傭金','客人自付','在庫車','富達分期','新車保險','營業車','高速累台','台本累台','來店客','衝刺獎金','代宇籌碼','支援金','自訂SP名稱','自訂SP金額','折價','配備','協和','百貨','補貼息','代宇費用','領牌','保險','加油','烤漆','介紹紅包','贈禮','外裝','自訂支出1名稱','自訂支出1金額','自訂支出2名稱','自訂支出2金額','自訂支出3名稱','自訂支出3金額','自訂支出4名稱','自訂支出4金額','自訂支出5名稱','自訂支出5金額','車體險','貸款配合','公司配件金額','業外報廢車','業外中古車','業外貼膜','業外隔熱紙','業外鍍膜','業外其他','籌碼總計','支出總計','差異','業外合計'];
const CHIPS = [['獎金','bonus'],['SP固定','sp'],['中古車','used'],['貸款傭金','loancom'],['客人自付','selfpay'],['在庫車','stock'],['富達分期','fuda'],['新車保險','insnew'],['營業車','biz'],['高速累台','hway'],['台本累台','taiben'],['來店客','visit'],['衝刺獎金','sprint'],['代宇籌碼','daiyu'],['支援金','support']];
const OUTS = [['折價','discount'],['配備','equip'],['協和','xiehe'],['百貨','dept'],['補貼息','subsidy'],['代宇費用','daiyufee'],['領牌','plate'],['保險','ins'],['加油','fuel'],['烤漆','paint'],['介紹紅包','referral'],['贈禮','gift'],['外裝','exterior']];
const EXTS = [['業外報廢車','scrap'],['業外中古車','extused'],['業外貼膜','film'],['業外隔熱紙','tint'],['業外鍍膜','coat'],['業外其他','other']];

function doGet(e) {
  const p = (e && e.parameter) || {};
  const cb = p.callback || '';
  let out;
  try {
    out = { ok: true, at: new Date().toISOString(), data: parseBlob(readBlob()) };
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  const body = cb ? cb + '(' + JSON.stringify(out) + ')' : JSON.stringify(out);
  return ContentService.createTextOutput(body)
    .setMimeType(cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

function doPost(e) {
  let out = { ok: false };
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action === 'save' || body.action === 'push') {
      const incoming = body.records || parseBlob(body.data || '');
      const lock = LockService.getScriptLock();
      lock.waitLock(20000);
      let merged;
      try {
        merged = mergeList(parseBlob(readBlob()), incoming);
        writeBlob(compress(merged));
        writeDataSheet(merged);
      } finally {
        lock.releaseLock();
      }
      out = { ok: true, total: merged.length, at: new Date().toISOString() };
    }
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function mergeList(server, incoming) {
  const map = {};
  (server || []).forEach(function (r) { if (r && r.id) map[r.id] = r; });
  (incoming || []).forEach(function (r) {
    if (!r || !r.id) return;
    const cur = map[r.id];
    if (!cur || (Number(r.upd) || 0) >= (Number(cur.upd) || 0)) map[r.id] = r;
  });
  return Object.keys(map).map(function (k) { return map[k]; })
    .sort(function (a, b) { return (Number(b.ts) || 0) - (Number(a.ts) || 0); });
}

function sheet_(name, head) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh && head) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh || ss.insertSheet(name);
}

function readBlob() {
  const sh = sheet_('SyncBlob');
  const last = sh.getLastRow();
  if (last < 1) return '';
  return sh.getRange(1, 2, last, 1).getValues().map(function (r) {
    return r[0] == null ? '' : String(r[0]);
  }).join('');
}

function writeBlob(data) {
  const sh = sheet_('SyncBlob');
  const chunks = [];
  for (let i = 0; i < data.length; i += CHUNK) chunks.push(data.slice(i, i + CHUNK));
  if (!chunks.length) chunks.push('');
  sh.clear();
  sh.getRange(1, 1, chunks.length, 2)
    .setValues(chunks.map(function (c, i) { return [i + 1, c]; }));
  sh.getRange(1, 4).setValue(new Date());
}

function compress(recs) {
  const json = JSON.stringify({ v: 1, at: Date.now(), records: recs });
  return Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(json, 'application/json')).getBytes());
}
function parseBlob(str) {
  if (!str) return [];
  try {
    const raw = Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(str))).getDataAsString();
    const d = JSON.parse(raw);
    return Array.isArray(d) ? d : (d.records || []);
  } catch (err) { return []; }
}

function writeDataSheet(recs) {
  const sh = sheet_(DATA_SHEET, COLS);
  const rows = [];
  recs.filter(function (r) { return !r.del; }).forEach(function (r) {
    const c = totals(r);
    const row = [];
    const put = function (name, val) { row[COLS.indexOf(name)] = val; };
    put('時間戳記', Number(r.ts) || '');
    put('業務員', r.staff || '');
    put('月份', monthStr(r.month));
    put('車款', r.model || '');
    put('顏色規格', r.color || '');
    put('客戶姓名', r.customer || '');
    put('狀態', r.status || '');
    CHIPS.forEach(function (x) { put(x[0], (r.chips && Number(r.chips[x[1]])) || ''); });
    put('自訂SP名稱', r.cspName || '');
    put('自訂SP金額', Number(r.csp) || '');
    OUTS.forEach(function (x) { put(x[0], (r.outs && Number(r.outs[x[1]])) || ''); });
    for (let i = 0; i < 5; i++) {
      const cc = (r.cust || [])[i] || {};
      put('自訂支出' + (i + 1) + '名稱', cc.n || '');
      put('自訂支出' + (i + 1) + '金額', Number(cc.a) || '');
    }
    put('車體險', r.flags && r.flags.body ? 1 : '');
    put('貸款配合', r.flags && r.flags.loan ? 1 : '');
    put('公司配件金額', Number(r.acc) || '');
    EXTS.forEach(function (x) { put(x[0], (r.ext && Number(r.ext[x[1]])) || ''); });
    put('籌碼總計', c.chips);
    put('支出總計', c.outs);
    put('差異', c.diff);
    put('業外合計', c.ext);
    rows.push(row);
  });
  const maxRow = sh.getMaxRows();
  if (maxRow > 1) sh.getRange(2, 1, maxRow - 1, COLS.length).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, COLS.length).setValues(rows);
}

function totals(r) {
  let chips = 0, outs = 0, ext = 0;
  CHIPS.forEach(function (x) { chips += Number((r.chips || {})[x[1]]) || 0; });
  chips += Number(r.csp) || 0;
  OUTS.forEach(function (x) { outs += Number((r.outs || {})[x[1]]) || 0; });
  (r.cust || []).forEach(function (c) { outs += Number(c && c.a) || 0; });
  EXTS.forEach(function (x) { ext += Number((r.ext || {})[x[1]]) || 0; });
  return { chips: chips, outs: outs, diff: chips - outs, ext: ext };
}

function monthStr(v) {
  if (!v) return '';
  const s = String(v);
  const m = s.match(/^(\d{4})-(\d{2})/);
  return m ? m[1] + '/' + Number(m[2]) : s;
}

function setup() {
  sheet_('SyncBlob');
  sheet_(DATA_SHEET, COLS);
  try { SpreadsheetApp.getActive().setActiveSheet(sheet_(DATA_SHEET, COLS)); } catch (e) {}
}
