/**
 * 獎金統計 — 雲端同步後端（Google Apps Script）
 *
 * 安裝步驟：
 * 1. 開一張新的 Google 試算表（空白即可），命名「獎金統計」。
 * 2. 選單「擴充功能 → Apps Script」，把這個檔案全部貼進 Code.gs。
 * 3. 右上「部署 → 新增部署作業 → 類型選『網頁應用程式』」。
 *    執行身分：我　　誰可以存取：任何人
 * 4. 複製結尾是 /exec 的網址，貼到 App 的「雲端同步」欄位。
 *
 * 資料存在這張試算表的 SyncBlob 分頁，第 1 欄為序號、第 2 欄為資料片段。
 */

const CHUNK = 40000;

function doGet(e) {
  const p = (e && e.parameter) || {};
  const cb = p.callback || '';
  let out = { ok: true, data: '' };
  try {
    const action = p.action || 'load';
    if (action === 'load') out = { ok: true, data: readBlob(), at: new Date().toISOString() };
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  const body = cb ? cb + '(' + JSON.stringify(out) + ')' : JSON.stringify(out);
  return ContentService
    .createTextOutput(body)
    .setMimeType(cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

function doPost(e) {
  let out = { ok: false };
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action === 'save') {
      const data = String(body.data || '');
      writeBlob(data);
      out = { ok: true, bytes: data.length, at: new Date().toISOString() };
    }
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('SyncBlob');
  if (!sh) sh = ss.insertSheet('SyncBlob');
  return sh;
}

function readBlob() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 1) return '';
  const vals = sh.getRange(1, 2, last, 1).getValues();
  return vals.map(function (r) { return r[0] == null ? '' : String(r[0]); }).join('');
}

function writeBlob(data) {
  const sh = sheet_();
  const chunks = [];
  for (let i = 0; i < data.length; i += CHUNK) chunks.push(data.slice(i, i + CHUNK));
  if (!chunks.length) chunks.push('');
  const rows = chunks.map(function (c, i) { return [i + 1, c]; });
  sh.clear();
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange(1, 4).setValue(new Date());
}

/** 首次使用可手動執行一次，建立分頁 */
function setup() { sheet_(); }
