/**************************************************************************
 * 楊梅高中交通安全學習網 — Google Apps Script（後端接收程式）
 * 功能：接收網頁提交的「日期時間、班級、座號、姓名、答對題數」，
 *       寫入 Google 試算表的「記錄」工作表。
 *
 * ─── 部署步驟 ───────────────────────────────────────────────
 * 1. 開啟一個 Google 試算表 (sheets.new)。
 * 2. 上方選單「擴充功能」→「Apps Script」，把本檔內容全部貼上、儲存。
 * 3. 點右上「部署」→「新增部署作業」→ 類型選「網頁應用程式」。
 * 4. 設定：
 *      - 執行身分：我 (你自己的帳號)
 *      - 誰可以存取：「所有人」(Anyone)  ← 學生才能匿名提交
 * 5. 按「部署」，複製產生的「網頁應用程式網址」(結尾是 /exec)。
 * 6. 把該網址貼到網頁 HTML 裡（測驗頁的「Google Apps Script Web App 網址」欄，
 *    或直接改 HTML 中 id="gasUrl" 的 value）。
 * 7. 第一次部署會要求授權，請一路允許。
 *
 * ※ 若之後修改程式，需重新「部署」→「管理部署作業」→ 編輯 → 版本選「新版本」。
 **************************************************************************/

var SHEET_NAME = '記錄';                 // 工作表名稱
var HEADERS = ['提交時間(伺服器)', '填答日期時間', '班級', '座號', '姓名', '答對題數'];

/** 取得（或建立）「記錄」工作表，並確保有標題列 */
function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length)
         .setFontWeight('bold')
         .setBackground('#2d9cdb')
         .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** 解析前端送來的資料（支援 JSON 或表單參數） */
function parseData_(e) {
  var data = {};
  try {
    if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    }
  } catch (err) {
    // 非 JSON 時退回讀取表單參數
  }
  if (e && e.parameter) {
    for (var k in e.parameter) {
      if (data[k] === undefined) data[k] = e.parameter[k];
    }
  }
  return data;
}

/** 寫入一列並回傳結果（同時被 doPost 與 doGet 使用） */
function handle_(e) {
  var data = parseData_(e);

  var datetime = data.datetime || new Date().toLocaleString('zh-TW');
  var cls      = data.cls   || '';
  var seat     = data.seat  || '';
  var name     = data.name  || '';
  var score    = (data.score !== undefined && data.score !== null) ? data.score : '';

  var sheet = getSheet_();
  sheet.appendRow([new Date(), datetime, cls, seat, name, score]);

  return jsonOut_({ status: 'ok', message: '已寫入記錄', received: data });
}

/** 前端 POST（正式提交走這裡） */
function doPost(e) {
  try {
    return handle_(e);
  } catch (err) {
    return jsonOut_({ status: 'error', message: String(err) });
  }
}

/**
 * 前端 GET（方便用瀏覽器直接測試，例如：
 * .../exec?cls=一年3班&seat=5&name=王小明&score=9 ）
 */
function doGet(e) {
  try {
    // 若沒帶任何參數，回傳一個健康檢查訊息
    if (!e || !e.parameter || Object.keys(e.parameter).length === 0) {
      return jsonOut_({ status: 'ready', message: '交通安全學習網後端運作中，請用 POST 提交資料。' });
    }
    return handle_(e);
  } catch (err) {
    return jsonOut_({ status: 'error', message: String(err) });
  }
}

/** 統一 JSON 輸出 */
function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** （選用）在編輯器中手動執行一次，測試是否能成功寫入 */
function testWrite() {
  handle_({ postData: { contents: JSON.stringify({
    datetime: new Date().toLocaleString('zh-TW'),
    cls: '一年3班', seat: '5', name: '測試同學', score: 9
  }) } });
}
