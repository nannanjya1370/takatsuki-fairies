/**
 * 高槻妖精会 出欠投票API（Google Apps Script Webアプリ）
 *
 * - GET  : 全データを返す { ok, votes: { "YYYY-MM-DD": { 名前: "yes"|"no" } }, members: [名前] }
 * - POST : { date, name, choice } で投票を記録（choice: "yes" | "no" | "clear"）
 *          { action: "setMembers", members: [名前] } でメンバー名簿を置き換え
 *          いずれも最新の全データを返す
 *
 * データは同じGoogleアカウントのスプレッドシート「高槻妖精会 出欠投票」に保存されます。
 */

var VOTES_SHEET = "votes";
var MEMBERS_SHEET = "members";
var TZ = "Asia/Tokyo";

function getSs_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty("SHEET_ID");
  var ss;
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    ss = SpreadsheetApp.create("高槻妖精会 出欠投票");
    props.setProperty("SHEET_ID", ss.getId());
  }
  return ss;
}

function getSheet_(ss, name, header) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(header);
  }
  return sheet;
}

// シート上でDate型に変換されていても "YYYY-MM-DD" に正規化する
function dateKey_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, "yyyy-MM-dd");
  return String(v);
}

function readVotes_(ss) {
  var sheet = getSheet_(ss, VOTES_SHEET, ["date", "name", "choice", "updated"]);
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

function readMembers_(ss) {
  var sheet = getSheet_(ss, MEMBERS_SHEET, ["name"]);
  var rows = sheet.getDataRange().getValues();
  var names = [];
  for (var i = 1; i < rows.length; i++) {
    var n = String(rows[i][0]).trim();
    if (n) names.push(n);
  }
  return names;
}

function writeMembers_(ss, names) {
  var sheet = getSheet_(ss, MEMBERS_SHEET, ["name"]);
  sheet.clearContents();
  var rows = [["name"]].concat(names.map(function (n) { return [n]; }));
  sheet.getRange(1, 1, rows.length, 1).setValues(rows);
}

function payload_(ss) {
  return { ok: true, votes: readVotes_(ss), members: readMembers_(ss) };
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return json_(payload_(getSs_()));
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var body = JSON.parse(e.postData.contents);
    var ss = getSs_();

    // メンバー名簿の同期（幹事の端末から送られてくる）
    if (body.action === "setMembers") {
      var names = [];
      var seen = {};
      (body.members || []).slice(0, 100).forEach(function (raw) {
        var n = String(raw).trim().slice(0, 20);
        if (n && !seen[n]) { seen[n] = true; names.push(n); }
      });
      writeMembers_(ss, names);
      return json_(payload_(ss));
    }

    // 投票の記録
    var date = String(body.date || "").slice(0, 10);
    var name = String(body.name || "").trim().slice(0, 20);
    var choice = String(body.choice || "");
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !name ||
        (choice !== "yes" && choice !== "no" && choice !== "clear")) {
      return json_({ ok: false, error: "bad request" });
    }
    var sheet = getSheet_(ss, VOTES_SHEET, ["date", "name", "choice", "updated"]);
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
    return json_(payload_(ss));
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
