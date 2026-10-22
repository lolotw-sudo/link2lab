/**
 * link2lab 後端端點（單一入口，GET 即可）
 *
 * 部署後的網址就是 skill 與 iOS 捷徑要用的端點。
 * 更新程式碼務必用「管理部署作業」→ 鉛筆 → 版本選「新版本」→ 部署，
 * 不要按「新增部署作業」（會產生新網址，捷徑吃不到）。
 */

// ⬇⬇⬇ 只有這兩行要改 ⬇⬇⬇
// 這是公開版本，實際值不寫在這裡（會外洩寫入權）。
// 實際值只填在 Google 的 Apps Script 編輯器裡，本機這份永遠保持佔位符。
const SS_ID = '請填入你的試算表 ID';   // 網址 /d/ 和 /edit 中間那一長串
const SECRET = '請自己改一個密語';      // 英數字、無空格，例如 MYLAB2026
// ⬆⬆⬆ 只有這兩行要改 ⬆⬆⬆

const INBOX_SHEET = '待處理';
const ITEM_SHEET = '靈感';

const INBOX_HEADER = ['時間戳記', '文章連結', '狀態'];
const ITEM_HEADER = ['標題', '一句話', '分類', '技術標籤', 'GitHub', '星數', '主要語言',
  'Demo', '原文連結', '來源', '能用在哪', '收錄日期'];

function doGet(e) {
  const token = e.parameter.token || '';
  if (token !== SECRET) {
    return out_('forbidden');
  }
  const action = e.parameter.action || 'inbox';

  // 一鍵建立兩個分頁與表頭（第一次設定時跑一次就好，重複跑不會破壞既有資料）
  if (action === 'setup') {
    const ss = SpreadsheetApp.openById(SS_ID);
    ensureSheet_(ss, INBOX_SHEET, INBOX_HEADER);
    ensureSheet_(ss, ITEM_SHEET, ITEM_HEADER);
    return out_('ok: 分頁「' + INBOX_SHEET + '」與「' + ITEM_SHEET + '」已就緒');
  }

  if (action === 'readinbox') {
    const sheet = sheet_(INBOX_SHEET);
    const data = sheet.getDataRange().getValues();
    const rows = data.slice(1).map(function (r, i) {
      return { row: i + 2, ts: r[0], url: r[1], status: r[2] };
    });
    return json_(rows);
  }

  if (action === 'setstatus') {
    const rowNum = parseInt(e.parameter.row, 10);
    const status = e.parameter.status || '完成';
    const sheet = sheet_(INBOX_SHEET);
    if (!rowNum || rowNum <= 1 || rowNum > sheet.getLastRow()) {
      return out_('bad row');
    }
    sheet.getRange(rowNum, 3).setValue(status);
    return out_('ok');
  }

  if (action === 'readitems') {
    const sheet = sheet_(ITEM_SHEET);
    const data = sheet.getDataRange().getValues();
    const rows = data.slice(1).map(function (r, i) {
      return {
        row: i + 2, title: r[0], summary: r[1], category: r[2], tags: r[3],
        repo: r[4], stars: r[5], lang: r[6], demo: r[7], url: r[8],
        source: r[9], use: r[10], added: r[11]
      };
    });
    return json_(rows);
  }

  if (action === 'additem') {
    const p = e.parameter;
    // 收錄日期由後端自動蓋，不由呼叫端傳入（避免時區錯亂與漏填）；
    // 需要補寫歷史資料時可用 added 參數指定。
    const added = p.added || Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
    sheet_(ITEM_SHEET).appendRow([
      sanitize_(p.title || ''), sanitize_(p.summary || ''), sanitize_(p.category || ''),
      sanitize_(p.tags || ''), sanitize_(p.repo || ''), sanitize_(p.stars || ''),
      sanitize_(p.lang || ''), sanitize_(p.demo || ''), sanitize_(p.url || ''),
      sanitize_(p.source || ''), sanitize_(p.use || ''), added
    ]);
    return out_('ok');
  }

  // 只更新星數與語言（定期回補 GitHub 熱度，不動其他欄位）
  if (action === 'updatestars') {
    const rowNum = parseInt(e.parameter.row, 10);
    const sheet = sheet_(ITEM_SHEET);
    if (!rowNum || rowNum <= 1 || rowNum > sheet.getLastRow()) return out_('bad row');
    if (e.parameter.stars) sheet.getRange(rowNum, 6).setValue(sanitize_(e.parameter.stars));
    if (e.parameter.lang) sheet.getRange(rowNum, 7).setValue(sanitize_(e.parameter.lang));
    return out_('ok');
  }

  if (action === 'deleterow') {
    const sheetName = e.parameter.sheet || INBOX_SHEET;
    const rowNum = parseInt(e.parameter.row, 10);
    if (rowNum > 1) sheet_(sheetName).deleteRow(rowNum);
    return out_('ok');
  }

  // 預設：手機捷徑寫入待處理列（inbox）
  const url = sanitize_(e.parameter.url || '');
  if (!url) return out_('missing url');
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ensureSheet_(ss, INBOX_SHEET, INBOX_HEADER);
  const ts = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
  sheet.appendRow([ts, url, '待處理']);
  return out_('ok');
}

function doPost(e) {
  return doGet(e);
}

function sheet_(name) {
  const s = SpreadsheetApp.openById(SS_ID).getSheetByName(name);
  if (!s) throw new Error('找不到分頁：' + name + '（請先跑 action=setup）');
  return s;
}

function ensureSheet_(ss, name, header) {
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.appendRow(header);
    s.setFrozenRows(1);
  } else if (s.getLastRow() === 0) {
    s.appendRow(header);
    s.setFrozenRows(1);
  }
  return s;
}

/** 防試算表公式注入：以 = + - @ 開頭的值前面加單引號 */
function sanitize_(v) {
  v = String(v);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function out_(text) {
  return ContentService.createTextOutput(text);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
