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
      meta.thumbnail = extractFirstImage(content);
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
