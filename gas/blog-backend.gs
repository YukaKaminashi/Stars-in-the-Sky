/**
 * 星ノ空ブログ バックエンド (Google Apps Script)
 *
 * ── セットアップ手順 ──
 * 1. Google スプレッドシートを新規作成
 * 2. そのスプレッドシートの URL から ID を取得し SPREADSHEET_ID に貼り付け
 *    例: https://docs.google.com/spreadsheets/d/【ここがID】/edit
 * 3. このスクリプトをデプロイ：
 *    「デプロイ」→「新しいデプロイ」→「ウェブアプリ」
 *    ・次のユーザーとして実行: 自分
 *    ・アクセスできるユーザー: 全員（匿名ユーザーを含む）
 * 4. 表示されたウェブアプリの URL を admin.html / blog.html / index.html の GAS_URL に貼り付け
 * 5. 初回のみ initializeSheet() と initializeNewsSheet() を手動実行してシートを初期化
 */

const SPREADSHEET_ID = '1D1eRCdgn8rIAR5lOXGUrMruJ3heytfqfMKayqcH6LUE';
const SHEET_NAME = 'posts';
const NEWS_SHEET_NAME = 'news';
const MEMBERS_SHEET_NAME = 'members';
const SCHEDULE_SHEET_NAME = 'schedule';
const COMMENTS_SHEET_NAME = 'comments';

// ── GET リクエスト ──────────────────────────────
function doGet(e) {
  const action = e.parameter.action;

  let result;
  if (action === 'getPosts') {
    result = getPosts();
  } else if (action === 'getPost') {
    result = getPost(e.parameter.id);
  } else if (action === 'getNews') {
    result = getNews();
  } else if (action === 'getLikes') {
    result = getLikes(e.parameter.id);
  } else if (action === 'getMembers') {
    result = getMembers(e.parameter.category);
  } else if (action === 'getSchedule') {
    result = getSchedule();
  } else if (action === 'getComments') {
    result = getComments(e.parameter.postId);
  } else if (action === 'getAllComments') {
    result = getAllComments();
  } else {
    result = { error: 'Invalid action' };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── POST リクエスト ─────────────────────────────
function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: 'Invalid JSON' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  let result;
  if (data.action === 'savePost') {
    result = savePost(data);
  } else if (data.action === 'saveNews') {
    result = saveNews(data);
  } else if (data.action === 'deleteNews') {
    result = deleteNews(data);
  } else if (data.action === 'likePost') {
    result = likePost(data);
  } else if (data.action === 'saveMember') {
    result = saveMember(data);
  } else if (data.action === 'deleteMember') {
    result = deleteMember(data);
  } else if (data.action === 'saveSchedule') {
    result = saveSchedule(data);
  } else if (data.action === 'deleteSchedule') {
    result = deleteSchedule(data);
  } else if (data.action === 'saveComment') {
    result = saveComment(data);
  } else if (data.action === 'deleteComment') {
    result = deleteComment(data);
  } else {
    result = { error: 'Invalid action' };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 投稿一覧取得 ────────────────────────────────
function getPosts() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const posts = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const post = {};
    headers.forEach((header, j) => { post[header] = row[j]; });
    if (post.published) {
      const { content, ...meta } = post;
      meta.excerpt = stripHtml(content).substring(0, 120);
      const thumb = extractFirstImage(content);
      meta.thumbnail = thumb.startsWith('data:') ? '' : thumb;
      posts.push(meta);
    }
  }

  posts.sort((a, b) => new Date(b.date) - new Date(a.date));
  return posts;
}

// ── 投稿1件取得 ─────────────────────────────────
function getPost(id) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const post = {};
    headers.forEach((header, j) => { post[header] = row[j]; });
    if (String(post.id) === String(id)) return post;
  }

  return { error: 'Not found' };
}

// ── 投稿保存 ────────────────────────────────────
function savePost(data) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  const id = new Date().getTime();
  const date = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');

  sheet.appendRow([id, data.title, data.author, data.group, data.content, date, true]);
  return { success: true, id: id };
}

// ── いいね数取得 ─────────────────────────────────
function getLikes(id) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const likesCol = headers.indexOf('likes');
  if (likesCol === -1) return { likes: 0 };

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      return { likes: Number(data[i][likesCol]) || 0 };
    }
  }
  return { likes: 0 };
}

// ── いいね追加 ────────────────────────────────────
function likePost(data) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  const range = sheet.getDataRange();
  const values = range.getValues();
  const headers = values[0];
  let likesCol = headers.indexOf('likes');

  // likes列がなければ追加
  if (likesCol === -1) {
    likesCol = headers.length;
    sheet.getRange(1, likesCol + 1).setValue('likes');
  }

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(data.id)) {
      const current = Number(values[i][likesCol]) || 0;
      sheet.getRange(i + 1, likesCol + 1).setValue(current + 1);
      return { success: true, likes: current + 1 };
    }
  }
  return { error: 'Not found' };
}

// ── ニュース一覧取得 ─────────────────────────────
function getNews() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(NEWS_SHEET_NAME);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const items = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const item = {};
    headers.forEach((h, j) => { item[h] = row[j]; });
    if (item.published) items.push(item);
  }

  items.sort((a, b) => new Date(b.date) - new Date(a.date));
  return items;
}

// ── ニュース保存 ─────────────────────────────────
function saveNews(data) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(NEWS_SHEET_NAME);
  const id = new Date().getTime();
  sheet.appendRow([id, data.date, data.badge, data.text, data.link || '', true]);
  return { success: true, id: id };
}

// ── ニュース削除 ─────────────────────────────────
function deleteNews(data) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(NEWS_SHEET_NAME);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { error: 'Not found' };
}

// ── 最初の画像URL抽出（サムネイル用）────────────
function extractFirstImage(html) {
  if (!html) return '';
  const match = html.match(/<img[^>]+src=["']([^"']+)["']/i);
  return match ? match[1] : '';
}

// ── HTML タグ除去（抜粋用）───────────────────────
function stripHtml(html) {
  if (!html) return '';
  return html.replace(/<[^>]*>/g, '').replace(/&nbsp;/g, ' ').trim();
}

// ── シート初期化（初回のみ手動実行）─────────────
function initializeSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['id', 'title', 'author', 'group', 'content', 'date', 'published']);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
  }
  Logger.log('postsシートの初期化完了');
}

// ── メンバー資料一覧取得 ──────────────────────────
function getMembers(category) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MEMBERS_SHEET_NAME);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const items = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const item = {};
    headers.forEach((h, j) => { item[h] = row[j]; });
    if (!item.published) continue;
    if (category && category !== 'all' && item.category !== category) continue;
    items.push(item);
  }

  items.sort((a, b) => new Date(b.date) - new Date(a.date));
  return items;
}

// ── メンバー資料保存 ──────────────────────────────
function saveMember(data) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MEMBERS_SHEET_NAME);
  const id = new Date().getTime();
  sheet.appendRow([id, data.category, data.title, data.link, data.date, true]);
  return { success: true, id: id };
}

// ── メンバー資料削除 ──────────────────────────────
function deleteMember(data) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MEMBERS_SHEET_NAME);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { error: 'Not found' };
}

// ── スケジュール一覧取得 ──────────────────────────
function getSchedule() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SCHEDULE_SHEET_NAME);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const items = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const item = {};
    headers.forEach((h, j) => { item[h] = row[j]; });
    if (item.published) items.push(item);
  }

  items.sort((a, b) => new Date(b.date) - new Date(a.date));
  return items;
}

// ── スケジュール保存 ──────────────────────────────
function saveSchedule(data) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SCHEDULE_SHEET_NAME);
  const id = new Date().getTime();
  sheet.appendRow([id, data.date, data.text, true]);
  return { success: true, id: id };
}

// ── スケジュール削除 ──────────────────────────────
function deleteSchedule(data) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SCHEDULE_SHEET_NAME);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { error: 'Not found' };
}

// ── コメント一覧取得（記事別・親子構造で返す）───────
function getComments(postId) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(COMMENTS_SHEET_NAME);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const all = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const item = {};
    headers.forEach((h, j) => { item[h] = row[j]; });
    if (item.published && String(item.postId) === String(postId)) {
      all.push(item);
    }
  }

  // 古い順にソート
  all.sort((a, b) => new Date(a.date) - new Date(b.date));

  // 親コメントに replies を付与
  const parents = all.filter(c => !c.parentId);
  const replies = all.filter(c => !!c.parentId);

  return parents.map(p => ({
    ...p,
    replies: replies.filter(r => String(r.parentId) === String(p.id))
  }));
}

// ── 全コメント取得（管理画面用）──────────────────────
function getAllComments() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(COMMENTS_SHEET_NAME);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const items = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const item = {};
    headers.forEach((h, j) => { item[h] = row[j]; });
    if (item.published) items.push(item);
  }
  items.sort((a, b) => new Date(b.date) - new Date(a.date));
  return items;
}

// ── コメント保存 ──────────────────────────────────
function saveComment(data) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(COMMENTS_SHEET_NAME);
  const id = new Date().getTime();
  const date = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
  sheet.appendRow([id, data.postId, data.parentId || '', data.name, data.text, date, true]);
  return { success: true, id: id, date: date };
}

// ── コメント削除（返信も連鎖削除）─────────────────
function deleteComment(data) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(COMMENTS_SHEET_NAME);
  const rows = sheet.getDataRange().getValues();
  // 削除対象のidを収集（親 + その返信）
  const targetIds = new Set([String(data.id)]);
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][2]) === String(data.id)) {
      targetIds.add(String(rows[i][0]));
    }
  }
  // 後ろから削除（行番号がずれないように）
  for (let i = rows.length - 1; i >= 1; i--) {
    if (targetIds.has(String(rows[i][0]))) {
      sheet.deleteRow(i + 1);
    }
  }
  return { success: true };
}

// ── コメントシート初期化（初回のみ手動実行）──────────
function initializeCommentsSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(COMMENTS_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(COMMENTS_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['id', 'postId', 'parentId', 'name', 'text', 'date', 'published']);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
  }
  Logger.log('commentsシートの初期化完了');
}

// ── メンバー資料シート初期化（初回のみ手動実行）────
function initializeMembersSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(MEMBERS_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(MEMBERS_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['id', 'category', 'title', 'link', 'date', 'published']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  }
  Logger.log('membersシートの初期化完了');
}

// ── スケジュールシート初期化（初回のみ手動実行）────
function initializeScheduleSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SCHEDULE_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['id', 'date', 'text', 'published']);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  }
  Logger.log('scheduleシートの初期化完了');
}

// ── ニュースシート初期化（初回のみ手動実行）───────
function initializeNewsSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(NEWS_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(NEWS_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['id', 'date', 'badge', 'text', 'link', 'published']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  }
  Logger.log('newsシートの初期化完了');
}
