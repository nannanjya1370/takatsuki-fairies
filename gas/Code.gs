/**
 * 高槻妖精会 出欠投票API（Google Apps Script Webアプリ）
 *
 * - GET  : 全データを返す { ok, votes: { "YYYY-MM-DD": { 名前: "yes"|"no" } }, members: [名前] }
 * - POST : { date, name, choice } で投票を記録（choice: "yes" | "no" | "clear"）
 *          { action: "setMembers", members: [名前] } でメンバー名簿を置き換え
 *          { action: "rename", from, to } で名簿と過去の投票の名前を一括変更
 *          { action: "setState", state: {...} } でアプリ全体の共有データ（メンバー・成績等）を保存
 *          { action: "addPost", name, comment, date, photo } で練習の様子を投稿（写真はDriveへ）
 *          { action: "cheer", postId, name } でがんばれ（応援）をトグル
 *          { action: "postComment", postId, name, text } で投稿にコメント
 *          いずれも最新の全データを返す（GETは posts も含む）
 *
 * データは同じGoogleアカウントのスプレッドシート「高槻妖精会 出欠投票」に保存されます。
 */

var VOTES_SHEET = "votes";
var MEMBERS_SHEET = "members";
var STATE_SHEET = "state";
var POSTS_SHEET = "posts";
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

// ===== 練習の様子（写真・コメント・応援）=====
function parseJson_(v, fallback) {
  if (v === "" || v === null || v === undefined) return fallback;
  try { var o = JSON.parse(String(v)); return o == null ? fallback : o; } catch (e) { return fallback; }
}

// 写真を保存する専用フォルダ（初回に作成してIDを記憶）
function getPhotoFolder_() {
  var props = PropertiesService.getScriptProperties();
  var fid = props.getProperty("PHOTO_FOLDER_ID");
  if (fid) { try { return DriveApp.getFolderById(fid); } catch (e) {} }
  var folder = DriveApp.createFolder("高槻妖精会 練習写真");
  props.setProperty("PHOTO_FOLDER_ID", folder.getId());
  return folder;
}

// dataURL("data:image/jpeg;base64,....") をDriveに保存し、共有可能なファイルIDを返す
function savePhoto_(dataUrl) {
  var m = String(dataUrl).match(/^data:(image\/[\w.+-]+);base64,(.+)$/);
  if (!m) return "";
  var blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], "fairies_" + Date.now() + ".jpg");
  var file = getPhotoFolder_().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getId();
}

function postsSheet_(ss) {
  return getSheet_(ss, POSTS_SHEET, ["id", "date", "name", "comment", "photoId", "cheers", "comments", "created"]);
}

function readPosts_(ss) {
  var sheet = postsSheet_(ss);
  var rows = sheet.getDataRange().getValues();
  var posts = [];
  for (var i = 1; i < rows.length; i++) {
    var id = String(rows[i][0]);
    if (!id) continue;
    var photoId = String(rows[i][4] || "");
    var created = rows[i][7] instanceof Date ? rows[i][7].getTime() : Number(rows[i][7]) || 0;
    posts.push({
      id: id,
      date: dateKey_(rows[i][1]),
      name: String(rows[i][2] || ""),
      comment: String(rows[i][3] || ""),
      photo: photoId ? ("https://drive.google.com/thumbnail?id=" + photoId + "&sz=w1280") : "",
      cheers: parseJson_(rows[i][5], []),
      comments: parseJson_(rows[i][6], []),
      created: created
    });
  }
  posts.sort(function (a, b) { return b.created - a.created; }); // 新しい投稿を上に
  return posts;
}

function findPostRow_(sheet, postId) {
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === postId) return i + 1;
  }
  return -1;
}

function payload_(ss) {
  return { ok: true, votes: readVotes_(ss), members: readMembers_(ss), state: readState_(ss), posts: readPosts_(ss) };
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

    // 練習の様子: 投稿を追加（写真はDriveに保存）
    if (body.action === "addPost") {
      var pName = String(body.name || "").trim().slice(0, 20);
      var pComment = String(body.comment || "").slice(0, 500);
      var pDate = String(body.date || "").slice(0, 10);
      if (!pName || (!pComment && !body.photo)) return json_({ ok: false, error: "bad request" });
      var photoId = "";
      if (body.photo) { photoId = savePhoto_(String(body.photo)); }
      var psheet = postsSheet_(ss);
      var newId = "p" + Date.now() + "_" + Math.floor(Math.random() * 100000);
      psheet.appendRow([newId, pDate, pName, pComment, photoId, "[]", "[]", new Date()]);
      return json_(payload_(ss));
    }

    // 練習の様子: がんばれ（応援）ボタンのトグル
    if (body.action === "cheer") {
      var cPostId = String(body.postId || "");
      var cName = String(body.name || "").trim().slice(0, 20);
      if (!cPostId || !cName) return json_({ ok: false, error: "bad request" });
      var csheet = postsSheet_(ss);
      var cRow = findPostRow_(csheet, cPostId);
      if (cRow < 0) return json_({ ok: false, error: "post not found" });
      var cheers = parseJson_(csheet.getRange(cRow, 6).getValue(), []);
      var idx = cheers.indexOf(cName);
      if (idx >= 0) cheers.splice(idx, 1); else cheers.push(cName); // もう一度押すと取り消し
      csheet.getRange(cRow, 6).setValue(JSON.stringify(cheers));
      return json_(payload_(ss));
    }

    // 練習の様子: 投稿へのコメント
    if (body.action === "postComment") {
      var mPostId = String(body.postId || "");
      var mName = String(body.name || "").trim().slice(0, 20);
      var mText = String(body.text || "").slice(0, 300);
      if (!mPostId || !mName || !mText) return json_({ ok: false, error: "bad request" });
      var msheet2 = postsSheet_(ss);
      var mRow = findPostRow_(msheet2, mPostId);
      if (mRow < 0) return json_({ ok: false, error: "post not found" });
      var comments = parseJson_(msheet2.getRange(mRow, 7).getValue(), []);
      comments.push({ name: mName, text: mText, created: Date.now() });
      if (comments.length > 200) comments = comments.slice(-200);
      msheet2.getRange(mRow, 7).setValue(JSON.stringify(comments));
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
