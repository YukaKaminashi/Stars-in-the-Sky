/**
 * 星ノ空ブログ バックエンド (Google Apps Script)
 *
 * ── セットアップ手順 ──
 * 1. Google スプレッドシートを新規作成
 * 2. そのスプレッドシートの URL から ID を取得し SPREADSHEET_ID に貼り付け
 *    例: https://docs.google.com/spreadsheets/d/【ここがID】/edit
 * 3. Google Drive で画像保存用フォルダを作成
 *    フォルダを開いたときの URL の末尾の ID を DRIVE_FOLDER_ID に貼り付け
 * 4. このスクリプトをデプロイ：
 *    「デプロイ」→「新しいデプロイ」→「ウェブアプリ」
 *    ・次のユーザーとして実行: 自分
 *    ・アクセスできるユーザー: 全員（匿名ユーザーを含む）
 * 5. 表示されたウェブアプリの URL を admin.html と blog.html の GAS_URL に貼り付け
 * 6. 初回のみ initializeSheet() を手動実行してシートを初期化
 */

const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';
const SHEET_NAME = 'posts';
const DRIVE_FOLDER_ID = 'YOUR_DRIVE_FOLDER_ID_HERE';

// ── GET リクエスト ──────────────────────────────
function doGet(e) {
  const action = e.parameter.action;

  let result;
  if (action === 'getPosts') {
    result = getPosts();
  } else if (action === 'getPost') {
    result = getPost(e.parameter.id);
  } else {
    result = { error: 'Invalid action' };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── POST リクエスト ─────────────────────────────
// Content-Type: text/plain で送信することでCORSプリフライトを回避
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
  } else if (data.action === 'uploadImage') {
    result = uploadImage(data);
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
      // contentは一覧では返さない（軽量化）
      const { content, ...meta } = post;
      meta.excerpt = stripHtml(content).substring(0, 120);
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

  sheet.appendRow([
    id,
    data.title,
    data.author,
    data.group,
    data.content,
    date,
    true
  ]);

  return { success: true, id: id };
}

// ── 画像アップロード（Google Drive 保存）─────────
function uploadImage(data) {
  const base64Data = data.imageData.replace(/^data:image\/\w+;base64,/, '');
  const mimeMatch = data.imageData.match(/^data:(image\/\w+);base64,/);
  if (!mimeMatch) return { error: 'Invalid image data' };

  const mimeType = mimeMatch[1];
  const ext = mimeType.split('/')[1];
  const fileName = 'blog_img_' + new Date().getTime() + '.' + ext;

  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const fileId = file.getId();
  const url = 'https://drive.google.com/uc?export=view&id=' + fileId;

  return { success: true, url: url };
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

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // ヘッダーが未設定なら追加
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['id', 'title', 'author', 'group', 'content', 'date', 'published']);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
  }

  Logger.log('シートの初期化が完了しました');
}
