/**
 * 高槻妖精会 出欠投票API（Google Apps Script Webアプリ）
 *
 * - GET  : 全投票データを返す { ok: true, votes: { "YYYY-MM-DD": { "名前": "yes"|"no" } } }
 * - POST : { date, name, choice } を記録して最新の全データを返す（choice: "yes" | "no" | "clear"）
 *
 * データは同じGoogleアカウントのスプレッドシート「高槻妖精会 出欠投票」に保存されます。
 */

var SHEET_NAME = "votes";
var TZ = "Asia/Tokyo";

function getSheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty("SHEET_ID");
  var ss;
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    ss = SpreadsheetApp.create("高槻妖精会 出欠投票");
    props.setProperty("SHEET_ID", ss.getId());
  }
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["date", "name", "choice", "updated"]);
  }
  return sheet;
}

// シート上でDate型に変換されていても "YYYY-MM-DD" に正規化する
function dateKey_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, "yyyy-MM-dd");
  return String(v);
}

function readVotes_(sheet) {
  var rows = sheet.getDataRange().getValues();
  var votes = {};
  for (var i = 1; i < rows.length; i++) {
    var date = dateKey_(rows[i][0]);
    var name = String(rows[i][1]);
    var choice = String(rows[i][2]);
    if (!date || !name) continue;
    if (choice !== "yes" && choice !== "no") continue;
    if (!votes[date]) votes[date] = {};
    votes[date][name] = choice;
  }
  return votes;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return json_({ ok: true, votes: readVotes_(getSheet_()) });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var body = JSON.parse(e.postData.contents);
    var date = String(body.date || "").slice(0, 10);
    var name = String(body.name || "").trim().slice(0, 20);
    var choice = String(body.choice || "");
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !name ||
        (choice !== "yes" && choice !== "no" && choice !== "clear")) {
      return json_({ ok: false, error: "bad request" });
    }
    var sheet = getSheet_();
    var rows = sheet.getDataRange().getValues();
    var rowIndex = -1;
    for (var i = 1; i < rows.length; i++) {
      if (dateKey_(rows[i][0]) === date && String(rows[i][1]) === name) {
        rowIndex = i + 1;
        break;
      }
    }
    if (choice === "clear") {
      if (rowIndex > 0) sheet.deleteRow(rowIndex);
    } else if (rowIndex > 0) {
      sheet.getRange(rowIndex, 3, 1, 2).setValues([[choice, new Date()]]);
    } else {
      sheet.appendRow([date, name, choice, new Date()]);
    }
    return json_({ ok: true, votes: readVotes_(sheet) });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
