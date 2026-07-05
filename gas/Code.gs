/**
 * 高槻妖精会 出欠投票API（Google Apps Script Webアプリ）
 *
 * - GET  : 全データを返す { ok, votes: { "YYYY-MM-DD": { 名前: "yes"|"no" } }, members: [名前] }
 * - POST : { date, name, choice } で投票を記録（choice: "yes" | "no" | "clear"）
 *          { action: "setMembers", members: [名前] } でメンバー名簿を置き換え
 *          { action: "rename", from, to } で名簿と過去の投票の名前を一括変更
 *          { action: "setState", state: {...} } でアプリ全体の共有データ（メンバー・成績等）を保存
 *          いずれも最新の全データを返す
 *
 * データは同じGoogleアカウントのスプレッドシート「高槻妖精会 出欠投票」に保存されます。
 */

var VOTES_SHEET = "votes";
var MEMBERS_SHEET = "members";
var STATE_SHEET = "state";
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

function readState_(ss) {
  var sheet = getSheet_(ss, STATE_SHEET, ["json"]);
  var v = sheet.getRange(2, 1).getValue();
  if (!v) return null;
  try { return JSON.parse(String(v)); } catch (e) { return null; }
}

function writeState_(ss, st) {
  var sheet = getSheet_(ss, STATE_SHEET, ["json"]);
  sheet.getRange(2, 1).setValue(JSON.stringify(st));
}

function payload_(ss) {
  return { ok: true, votes: readVotes_(ss), members: readMembers_(ss), state: readState_(ss) };
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

    // メンバー名簿の同期（既存名簿と統合。旧バージョンの端末が名簿を縮めないように）
    if (body.action === "setMembers") {
      var names = readMembers_(ss);
      var seen = {};
      names.forEach(function (n) { seen[n] = true; });
      (body.members || []).slice(0, 100).forEach(function (raw) {
        var n = String(raw).trim().slice(0, 20);
        if (n && !seen[n]) { seen[n] = true; names.push(n); }
      });
      writeMembers_(ss, names);
      return json_(payload_(ss));
    }

    // アプリ全体の共有データを保存（メンバー・成績・ポイント等）。投票用名簿も連動更新
    if (body.action === "setState") {
      var st = body.state;
      if (!st || !Array.isArray(st.members)) return json_({ ok: false, error: "bad request" });
      writeState_(ss, st);
      var stNames = [];
      var stSeen = {};
      st.members.slice(0, 100).forEach(function (m) {
        var n = String((m && m.name) || "").trim().slice(0, 20);
        if (n && !stSeen[n]) { stSeen[n] = true; stNames.push(n); }
      });
      writeMembers_(ss, stNames);
      return json_(payload_(ss));
    }

    // 改名: 名簿と過去の投票の名前を一括で書き換える
    if (body.action === "rename") {
      var from = String(body.from || "").trim().slice(0, 20);
      var to = String(body.to || "").trim().slice(0, 20);
      if (!from || !to || from === to) return json_({ ok: false, error: "bad request" });
      var vsheet = getSheet_(ss, VOTES_SHEET, ["date", "name", "choice", "updated"]);
      var vrows = vsheet.getDataRange().getValues();
      for (var vi = 1; vi < vrows.length; vi++) {
        if (String(vrows[vi][1]) === from) vsheet.getRange(vi + 1, 2).setValue(to);
      }
      var msheet = getSheet_(ss, MEMBERS_SHEET, ["name"]);
      var mrows = msheet.getDataRange().getValues();
      for (var mi = 1; mi < mrows.length; mi++) {
        if (String(mrows[mi][0]) === from) msheet.getRange(mi + 1, 1).setValue(to);
      }
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
