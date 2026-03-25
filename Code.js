// Cấu hình tên sheet chứa dữ liệu thiết lập
const SHEET_NAME = 'Configs';

/**
 * Hiển thị giao diện Web App
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Cấu hình Gửi Tin Nhắn Tự Động - Google Chat')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Hàm xác thực và lấy thông tin User
 */
function getUserInfo() {
  return Session.getActiveUser().getEmail();
}

/**
 * Khởi tạo sheet nếu chưa có và cập nhật Headers
 * Column layout:
 *   A(1)=ID, B(2)=タイトル, C(3)=Webhook URL, D(4)=曜日, E(5)=時刻,
 *   F(6)=内容, G(7)=状態, H(8)=CreatedBy, I(9)=CreatedAt,
 *   J(10)=LastModifiedBy, K(11)=LastModifiedAt, L(12)=RecurrenceType, M(13)=EndDate
 */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      'ID', 'タイトル', 'Webhook URL', '曜日', '時刻', '内容', '状態',
      'CreatedBy', 'CreatedAt', 'LastModifiedBy', 'LastModifiedAt',
      'RecurrenceType', 'EndDate'
    ]);
    sheet.getRange('A1:M1').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * API Lấy danh sách cấu hình
 */
function getConfigs() {
  try {
    const sheet = setupSheet();
    const data = sheet.getDataRange().getValues();
    const currentUser = getUserInfo();
    if (!currentUser) {
      return { success: true, configs: [], currentUser: '', warning: 'ログインが必要です。' };
    }
    const configs = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      const createdBy = row[7];
      if (createdBy !== currentUser) continue; // user-scoped
      const formatSafeDate = (val) => {
        if (val instanceof Date) return Utilities.formatDate(val, 'GMT+9', 'yyyy/MM/dd HH:mm:ss');
        return val ? String(val) : '';
      };
      configs.push({
        id: row[0],
        title: row[1],
        webhookUrl: row[2],
        days: row[3],
        time: (row[4] instanceof Date) ? Utilities.formatDate(row[4], 'GMT+9', 'HH:mm') : String(row[4] || ''),
        message: row[5],
        status: row[6],
        createdBy: createdBy,
        createdAt: formatSafeDate(row[8]),
        lastModifiedBy: row[9],
        lastModifiedAt: formatSafeDate(row[10]),
        recurrenceType: row[11] || 'recurring',
        endDate: row[12] || ''
      });
    }
    return { success: true, configs: configs, currentUser: currentUser };
  } catch (e) {
    return { success: false, configs: [], currentUser: '', errorMessage: e.toString() };
  }
}

/**
 * Hàm Lưu/Sửa (kèm truy vết)
 * Column layout (1-indexed for getRange):
 *   A(1)=ID, B(2)=タイトル, C(3)=Webhook URL, D(4)=曜日, E(5)=時刻,
 *   F(6)=内容, G(7)=状態, H(8)=CreatedBy, I(9)=CreatedAt,
 *   J(10)=LastModifiedBy, K(11)=LastModifiedAt, L(12)=RecurrenceType, M(13)=EndDate
 * @param {Object} formData Đối tượng chứa thông tin cấu hình từ form
 */
function upsertRemind(formData) {
  try {
    const sheet = setupSheet();
    const email = getUserInfo();
    const nowTime = Utilities.formatDate(new Date(), 'GMT+9', 'yyyy/MM/dd HH:mm:ss');
    const daysString = Array.isArray(formData.days) ? formData.days.join(', ') : formData.days;
    const data = sheet.getDataRange().getValues();
    if (formData.id) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === formData.id) {
          const createdBy = data[i][7];
          if (createdBy !== email && email !== '') {
            return { success: false, message: '編集権限がありません。' };
          }
          const rowNum = i + 1;
          // Cols B→G (index 2→7)
          sheet.getRange(rowNum, 2, 1, 6).setValues([[
            formData.title, formData.webhookUrl, daysString,
            formData.time, formData.message, formData.status || 'Active'
          ]]);
          // Cols J→M (index 10→13)
          sheet.getRange(rowNum, 10, 1, 4).setValues([[
            email, nowTime,
            formData.recurrenceType || 'recurring',
            formData.endDate || ''
          ]]);
          return { success: true, message: 'リマインダーを更新しました。' };
        }
      }
      return { success: false, message: '対象のレコードが見つかりません。' };
    } else {
      const newId = Utilities.getUuid();
      sheet.appendRow([
        newId,                           // A: ID
        formData.title,                  // B: タイトル
        formData.webhookUrl,             // C: Webhook URL
        daysString,                      // D: 曜日
        formData.time,                   // E: 時刻
        formData.message,                // F: 内容
        'Active',                        // G: 状態
        email,                           // H: CreatedBy
        nowTime,                         // I: CreatedAt
        '', '',                          // J,K: LastModified
        formData.recurrenceType || 'recurring', // L: RecurrenceType
        formData.endDate || ''           // M: EndDate
      ]);
      return { success: true, message: 'リマインダーを保存しました。' };
    }
  } catch (e) {
    return { success: false, message: 'エラーが発生しました: ' + e.message };
  }
}

/**
 * Xóa một bản ghi (Delete)
 */
function deleteConfig(id) {
  try {
    const sheet = setupSheet();
    const values = sheet.getDataRange().getValues();
    const userEmail = getUserInfo();
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === id) {
        const createdBy = values[i][7];
        if (createdBy !== userEmail && userEmail !== '') {
          return { success: false, message: '削除権限がありません。' };
        }
        sheet.deleteRow(i + 1);
        return { success: true, message: 'リマインダーを削除しました。' };
      }
    }
    return { success: false, message: '対象のレコードが見つかりません。' };
  } catch (e) {
    return { success: false, message: 'エラーが発生しました: ' + e.message };
  }
}

/**
 * ステータスをActive/Inactiveに切り替える
 */
function toggleStatus(id) {
  try {
    const sheet = setupSheet();
    const values = sheet.getDataRange().getValues();
    const userEmail = getUserInfo();
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === id) {
        const createdBy = values[i][7];
        if (createdBy !== userEmail && userEmail !== '') {
          return { success: false, message: '変更権限がありません。' };
        }
        const newStatus = values[i][6] === 'Active' ? 'Inactive' : 'Active';
        sheet.getRange(i + 1, 7).setValue(newStatus);
        return { success: true, newStatus: newStatus };
      }
    }
    return { success: false, message: '対象のレコードが見つかりません。' };
  } catch (e) {
    return { success: false, message: 'エラーが発生しました: ' + e.message };
  }
}

/**
 * Hàm Timer Trigger: Quét mỗi phút để xem đến giờ gửi tin nhắn chưa
 */
function checkAndSendMessages() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;
  const now = new Date();
  const currentDay = Utilities.formatDate(now, 'GMT+9', 'EEEE');
  const currentTime = Utilities.formatDate(now, 'GMT+9', 'HH:mm');
  const todayStr = Utilities.formatDate(now, 'GMT+9', 'yyyy/MM/dd');
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[6] !== 'Active') continue;
    // EndDate check (col M, index 12)
    const endDate = row[12] ? String(row[12]).trim() : '';
    if (endDate && todayStr > endDate) {
      sheet.getRange(i + 1, 7).setValue('Inactive');
      continue;
    }
    // Multiple times: split by comma (col E, index 4)
    let timesRaw = row[4] instanceof Date
      ? Utilities.formatDate(row[4], 'GMT+9', 'HH:mm')
      : String(row[4] || '');
    const times = timesRaw.split(',').map(t => t.trim()).filter(t => t);
    const scheduledDays = String(row[3] || '');
    if (!scheduledDays.includes(currentDay) || !times.includes(currentTime)) continue;
    sendMessageToChat(row[2], row[5], row[1]);
    // RecurrenceType check (col L, index 11): auto-deactivate if 'once'
    // Only deactivate after the last scheduled time of the day has been sent
    if ((row[11] || 'recurring') === 'once') {
      const sortedTimes = [...times].sort();
      if (currentTime >= sortedTimes[sortedTimes.length - 1]) {
        sheet.getRange(i + 1, 7).setValue('Inactive');
      }
    }
  }
}

/**
 * Gửi tin nhắn đến Google Chat Webhook
 */
function sendMessageToChat(webhookUrl, message, title = "THÔNG BÁO/通知") {
  if (!webhookUrl) return;
  
  // Sử dụng Markdown để làm nổi bật nội dung
  const formattedText = `*🔔 ${title}*\n` + 
                        `—————————————————\n\n` + 
                        `${message}`;

  const payload = { text: formattedText };
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    UrlFetchApp.fetch(webhookUrl, options);
  } catch (error) {
    Logger.log('Lỗi gửi Webhook: ' + error.message);
  }
}

function testWebhook() {
  const url = "https://chat.googleapis.com/v1/spaces/AAQAAgZMsus/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=aIgHucwPisdiXGd-BSvBs828ooH9CrNknO7x5dk34CI";
  sendMessageToChat(url, "Tin nhắn thử nghiệm từ Apps Script", "THỬ NGHIỆM/テスト");
}

/**
 * Hàm hỗ trợ thiết lập Trigger quét mỗi phút (Chạy thủ công 1 lần)
 */
function createTimeDrivenTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkAndSendMessages') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('checkAndSendMessages')
      .timeBased()
      .everyMinutes(1)
      .create();
}

// =============================================================
// ===== NotebookLM 自動変換 (Auto Convert Feed) =====
// Config keys stored in ScriptProperties
// =============================================================
const NLM_CONFIG_KEYS = ['FOLDER_INPUT_ID', 'MASTER_SHEET_ID', 'MASTER_DOC_ID', 'FOLDER_ARCHIVE_ID'];

/**
 * 設定を取得する (UI用)
 */
function getNotebookLMConfig() {
  try {
    const props = PropertiesService.getScriptProperties();
    const config = {};
    NLM_CONFIG_KEYS.forEach(k => { config[k] = props.getProperty(k) || ''; });
    return { success: true, config: config };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * 設定を保存する (UI用)
 */
function saveNotebookLMConfig(data) {
  try {
    const props = PropertiesService.getScriptProperties();
    NLM_CONFIG_KEYS.forEach(k => {
      if (data[k] !== undefined) props.setProperty(k, String(data[k]).trim());
    });
    return { success: true, message: '設定を保存しました。' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * 手動実行 (UI用) — 入力フォルダのファイルを処理してMasterに統合
 */
function runNotebookLMProcess() {
  try {
    const props = PropertiesService.getScriptProperties();
    const config = {};
    NLM_CONFIG_KEYS.forEach(k => { config[k] = props.getProperty(k) || ''; });
    if (!config.FOLDER_INPUT_ID || !config.MASTER_SHEET_ID || !config.MASTER_DOC_ID || !config.FOLDER_ARCHIVE_ID) {
      return { success: false, message: '設定が不完全です。すべてのIDを入力・保存してください。' };
    }
    const log = nlmMainProcess(config);
    return { success: true, message: '処理が完了しました。', log: log };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}

/**
 * メイン処理: 入力フォルダのファイルを種別ごとに処理
 */
function nlmMainProcess(config) {
  const folderInput = DriveApp.getFolderById(config.FOLDER_INPUT_ID);
  const archiveFolder = DriveApp.getFolderById(config.FOLDER_ARCHIVE_ID);
  const files = folderInput.getFiles();
  const log = [];

  nlmClearMasterFiles(config);
  log.push('Masterファイルをクリアしました。');

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    const mimeType = file.getMimeType();
    try {
      if (mimeType === MimeType.GOOGLE_SHEETS || fileName.endsWith('.xlsx')) {
        nlmProcessExcelToMaster(file, config);
        log.push('[Excel/Sheets] ' + fileName);
      } else if (mimeType === MimeType.GOOGLE_DOCS || fileName.endsWith('.docx') || fileName.endsWith('.txt')) {
        nlmProcessDocToMaster(file, config);
        log.push('[Doc/Text] ' + fileName);
      } else if (mimeType === MimeType.JPEG || mimeType === MimeType.PNG || mimeType === MimeType.PDF) {
        nlmProcessImageToMaster(file, config);
        log.push('[画像/PDF] ' + fileName);
      } else {
        log.push('[スキップ] ' + fileName + ' (対応外のファイル形式)');
      }
      file.moveTo(archiveFolder);
    } catch (e) {
      log.push('[エラー] ' + fileName + ': ' + e.message);
      Logger.log('エラー ' + fileName + ': ' + e.message);
    }
  }
  return log;
}

/**
 * 画像・PDF を OCR してMaster Docへ追加
 */
function nlmProcessImageToMaster(file, config) {
  const masterDoc = DocumentApp.openById(config.MASTER_DOC_ID);
  const body = masterDoc.getBody();
  const resource = { name: 'temp_ocr_' + file.getName(), mimeType: MimeType.GOOGLE_DOCS };
  const tempFile = Drive.Files.create(resource, file.getBlob());
  const textContent = DocumentApp.openById(tempFile.id).getBody().getText();
  body.appendParagraph('NGUỒN ẢNH/PDF: ' + file.getName()).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(textContent);
  body.appendPageBreak();
  Drive.Files.remove(tempFile.id);
}

/**
 * Excel / Google Sheets を Master Sheetsへコピー
 */
function nlmProcessExcelToMaster(file, config) {
  const masterSs = SpreadsheetApp.openById(config.MASTER_SHEET_ID);
  if (file.getName().endsWith('.xlsx')) {
    const tempFile = Drive.Files.create(
      { name: file.getName().replace('.xlsx', ''), mimeType: MimeType.GOOGLE_SHEETS },
      file.getBlob()
    );
    SpreadsheetApp.openById(tempFile.id).getSheets().forEach(sheet => {
      sheet.copyTo(masterSs).setName(file.getName() + ' - ' + sheet.getName());
    });
    Drive.Files.remove(tempFile.id);
  } else {
    SpreadsheetApp.openById(file.getId()).getSheets().forEach(sheet => {
      sheet.copyTo(masterSs).setName(file.getName() + ' - ' + sheet.getName());
    });
  }
}

/**
 * Docx / Google Docs / TXT を Master Docへ追記
 */
function nlmProcessDocToMaster(file, config) {
  const masterDoc = DocumentApp.openById(config.MASTER_DOC_ID);
  const body = masterDoc.getBody();
  let textContent = '';
  if (file.getName().endsWith('.docx')) {
    const tempFile = Drive.Files.create(
      { name: 'temp_doc', mimeType: MimeType.GOOGLE_DOCS },
      file.getBlob()
    );
    textContent = DocumentApp.openById(tempFile.id).getBody().getText();
    Drive.Files.remove(tempFile.id);
  } else if (file.getMimeType() === MimeType.GOOGLE_DOCS) {
    textContent = DocumentApp.openById(file.getId()).getBody().getText();
  } else {
    textContent = file.getBlob().getDataAsString();
  }
  body.appendParagraph('NGUỒN: ' + file.getName()).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(textContent);
  body.appendPageBreak();
}

/**
 * Master Doc と Master Sheets をクリア
 */
function nlmClearMasterFiles(config) {
  DocumentApp.openById(config.MASTER_DOC_ID).getBody().clear();
  const ss = SpreadsheetApp.openById(config.MASTER_SHEET_ID);
  const sheets = ss.getSheets();
  ss.insertSheet('TempClearSheet');
  sheets.forEach(s => ss.deleteSheet(s));
  ss.getSheets()[0].setName('Nội dung mới');
}

/**
 * ローカルファイル（Base64）を入力フォルダへアップロード
 * @param {Array} filesData [{name, mimeType, base64}]
 */
function uploadLocalFilesToInputFolder(filesData) {
  try {
    const folderId = PropertiesService.getScriptProperties().getProperty('FOLDER_INPUT_ID');
    if (!folderId) return { success: false, message: '入力フォルダIDが設定されていません。設定を保存してください。' };
    const folder = DriveApp.getFolderById(folderId);
    const results = [];
    filesData.forEach(function(f) {
      try {
        const bytes = Utilities.base64Decode(f.base64);
        const blob = Utilities.newBlob(bytes, f.mimeType || 'application/octet-stream', f.name);
        const created = folder.createFile(blob);
        results.push({ name: f.name, success: true, id: created.getId() });
      } catch (e) {
        results.push({ name: f.name, success: false, message: e.message });
      }
    });
    return { success: true, results: results };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}

/**
 * Google Drive のファイル（URL または ID）を入力フォルダへコピー
 * @param {Array} urlsOrIds [string, ...]
 */
function copyDriveFilesToInputFolder(urlsOrIds) {
  try {
    const folderId = PropertiesService.getScriptProperties().getProperty('FOLDER_INPUT_ID');
    if (!folderId) return { success: false, message: '入力フォルダIDが設定されていません。設定を保存してください。' };
    const folder = DriveApp.getFolderById(folderId);
    const results = [];
    urlsOrIds.forEach(function(urlOrId) {
      // URL からファイルID を抽出
      let fileId = String(urlOrId).trim();
      const patterns = [
        /\/d\/([a-zA-Z0-9_-]{10,})/,
        /id=([a-zA-Z0-9_-]{10,})/,
        /\/file\/d\/([a-zA-Z0-9_-]{10,})/
      ];
      for (let j = 0; j < patterns.length; j++) {
        const m = fileId.match(patterns[j]);
        if (m) { fileId = m[1]; break; }
      }
      try {
        const src = DriveApp.getFileById(fileId);
        src.makeCopy(src.getName(), folder);
        results.push({ url: urlOrId, success: true, fileName: src.getName() });
      } catch (e) {
        results.push({ url: urlOrId, success: false, message: e.message });
      }
    });
    const ok = results.filter(function(r) { return r.success; }).length;
    return { success: true, results: results, okCount: ok, ngCount: results.length - ok };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}

// =============================================================
// ===== 共通空き時間検索 (Common Free Time Finder) =====
// 必要な Advanced Services: Chat API v1, Calendar API v3 のみ
//   People API は Advanced Service 不要 — UrlFetchApp で REST 直呼び出し
// 必要な OAuth スコープ:
//   - https://www.googleapis.com/auth/chat.memberships.readonly
//   - https://www.googleapis.com/auth/calendar.readonly
//   - https://www.googleapis.com/auth/directory.readonly
//   - https://www.googleapis.com/auth/script.external_request
//
// 【初回セットアップ / スコープ追加後に必須の手順】
//   GAS エディタ (script.google.com) で「triggerReAuthorization」関数を
//   1回だけ手動実行してください。OAuth 承認画面が表示されたら承認します。
// =============================================================

/**
 * 【一度だけ手動実行】新しい OAuth スコープ (directory.readonly) を承認する。
 *
 * GAS エディタ上部のドロップダウンで「triggerReAuthorization」を選択して ▶ 実行。
 * 「このアプリは Google アカウントへのアクセスを求めています」という画面が
 * 表示されたら「許可」をクリックしてください。
 * 実行ログに ✅ が表示されれば成功です。
 */
function triggerReAuthorization() {
  const token = ScriptApp.getOAuthToken();
  const res = UrlFetchApp.fetch(
    'https://people.googleapis.com/v1/people:listDirectoryPeople'
    + '?readMask=emailAddresses'
    + '&sources=DIRECTORY_SOURCE_TYPE_DOMAIN_PROFILE'
    + '&pageSize=1',
    { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
  );
  const code = res.getResponseCode();
  if (code === 200) {
    Logger.log('✅ 認証成功！directory.readonly スコープが有効です。Web App が正しく動作します。');
  } else {
    Logger.log('❌ HTTP ' + code + ': ' + res.getContentText().substring(0, 500));
    Logger.log('承認画面が表示されなかった場合は、スクリプトを再デプロイしてください。');
  }
}

/**
 * Google Chat スペースのメンバー一覧を取得し、メールアドレスの配列を返す。
 *
 * Chat API は member.name = "users/{numericId}" を返すため、
 * People API REST エンドポイント (UrlFetchApp) で数値 ID → メールを解決する。
 * People API が 403 の場合は displayName を含む members 配列と
 * emailResolutionFailed フラグを返してフロントエンドで手動入力できるようにする。
 *
 * @param {string} spaceId - スペース ID (例: "spaces/XXXXXXX" または "XXXXXXX")
 * @returns {{ success: boolean, emails: string[], members: Array, emailResolutionFailed?: boolean, message?: string }}
 */
function getSpaceMembers(spaceId) {
  try {
    if (!spaceId || !spaceId.trim()) {
      return { success: false, emails: [], members: [], message: 'Space ID を入力してください。' };
    }

    // "spaces/" プレフィックスの正規化
    const normalizedId = spaceId.trim().startsWith('spaces/')
      ? spaceId.trim()
      : 'spaces/' + spaceId.trim();

    // Step 1: Chat API でメンバーの userId と displayName を収集
    const directEmails = [];
    const userIds = [];
    const membersList = []; // { userId, displayName } の配列
    let pageToken = null;

    do {
      const params = { pageSize: 100 };
      if (pageToken) params.pageToken = pageToken;

      const response = Chat.Spaces.Members.list(normalizedId, params);
      (response.memberships || []).forEach(m => {
        if (!m.member || m.member.type !== 'HUMAN' || !m.member.name) return;

        const userId      = m.member.name.replace('users/', '');
        const displayName = m.member.displayName || userId;
        membersList.push({ userId, displayName });

        // 一部の設定では既にメールアドレスが返る
        if (userId.includes('@')) {
          directEmails.push(userId);
        } else {
          userIds.push(userId);
        }
      });

      pageToken = response.nextPageToken;
    } while (pageToken);

    if (membersList.length === 0) {
      return {
        success: false, emails: [], members: [],
        message: 'スペースに HUMAN メンバーが見つかりませんでした。Space ID を確認してください。'
      };
    }

    // Step 2: People API REST で数値 userId → メールを解決 (Admin 権限不要)
    const { emails: resolvedEmails, errorDetail } = resolveEmailsViaPeopleApi_(userIds);
    const emails = [...directEmails, ...resolvedEmails];

    if (emails.length === 0) {
      // People API が失敗したが Chat API のメンバー情報は取得できた → 手動入力フォールバックへ
      return {
        success: true,
        emails: [],
        members: membersList,
        emailResolutionFailed: true,
        message: errorDetail && errorDetail.includes('403')
          ? 'directory.readonly スコープの認証が必要か、組織のポリシーで制限されています。'
          : 'People API エラー: ' + (errorDetail || 'メールアドレスフィールドが空でした。')
      };
    }

    return { success: true, emails: emails, members: membersList };

  } catch (e) {
    Logger.log('getSpaceMembers error: ' + e.message);
    const isPermErr = e.message.includes('PERMISSION_DENIED') || e.message.includes('403');
    return {
      success: false,
      emails: [],
      members: [],
      message: isPermErr
        ? 'Chat API へのアクセス権限がありません。スコープ chat.memberships.readonly を確認してください。'
        : 'エラー: ' + e.message
    };
  }
}

/**
 * People API v1 REST (batchGet) で数値ユーザー ID → メールアドレスを解決する。
 *
 * Advanced Service (People.People) の代わりに UrlFetchApp を使用することで
 * GAS エディタでの手動サービス登録が不要になる。
 * スコープ directory.readonly があれば同 Workspace ドメイン内ユーザーを参照可能。
 *
 * @param {string[]} userIds - Chat API から取得した数値ユーザー ID の配列
 * @returns {{ emails: string[], errorDetail: string|null }}
 */
function resolveEmailsViaPeopleApi_(userIds) {
  if (!userIds.length) return { emails: [], errorDetail: null };

  const token  = ScriptApp.getOAuthToken();
  const idSet  = new Set(userIds);          // 探索対象の数値 ID セット
  const idMap  = {};                        // numericId → email
  let pageToken   = null;
  let errorDetail = null;

  // people:listDirectoryPeople を使う理由:
  //   batchGet + READ_SOURCE_TYPE_DOMAIN_CONTACT は「組織連絡先カード」を参照するため
  //   emailAddresses が空になるケースがある。
  //   listDirectoryPeople + DIRECTORY_SOURCE_TYPE_DOMAIN_PROFILE は
  //   Workspace ログインプロファイルを参照するため、プライマリメールが確実に取得できる。
  //   必要スコープ: directory.readonly のみ (Admin 権限不要)
  do {
    let url = 'https://people.googleapis.com/v1/people:listDirectoryPeople'
            + '?readMask=emailAddresses'
            + '&sources=DIRECTORY_SOURCE_TYPE_DOMAIN_PROFILE'
            + '&pageSize=1000';
    if (pageToken) url += '&pageToken=' + encodeURIComponent(pageToken);

    try {
      const res  = UrlFetchApp.fetch(url, {
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      });
      const code = res.getResponseCode();
      const body = res.getContentText();

      if (code !== 200) {
        errorDetail = 'HTTP ' + code + ': ' + body.substring(0, 400);
        Logger.log('listDirectoryPeople error: ' + errorDetail);
        break;
      }

      const data = JSON.parse(body);
      (data.people || []).forEach(person => {
        // resourceName = "people/{numericId}"
        const id = (person.resourceName || '').replace('people/', '');
        if (!idSet.has(id)) return; // 対象外はスキップ

        const emailField =
          (person.emailAddresses || []).find(e => e.metadata && e.metadata.primary)
          || (person.emailAddresses || [])[0];
        if (emailField && emailField.value) idMap[id] = emailField.value;
      });

      pageToken = data.nextPageToken || null;

      // 全対象ユーザーのメールが揃ったら早期終了
      if (Object.keys(idMap).length >= idSet.size) break;

    } catch (e) {
      errorDetail = e.message;
      Logger.log('listDirectoryPeople fetch error: ' + e.message);
      break;
    }
  } while (pageToken);

  // 元の順序でメールアドレスを返す (見つからない ID はスキップ)
  const emails = userIds.map(id => idMap[id]).filter(Boolean);
  return { emails, errorDetail };
}

/**
 * 複数ユーザーの Google カレンダーを参照し、全員が空いている共通スロットを返す。
 *
 * 【タイムゾーン処理】
 * - フロントエンドの datetime-local は "YYYY-MM-DDTHH:mm" 形式で値を返す (TZ なし)
 * - フロントエンドで "+09:00" サフィックスを付与して JST であることを明示して送信する
 * - GAS 側では "+09:00" 付きの ISO 文字列を new Date() でパースすることで
 *   UTC への変換が正確に行われる
 * - 結果の表示は Utilities.formatDate(..., 'Asia/Tokyo', ...) で JST に変換
 *
 * @param {string[]} emails      - ユーザーのメールアドレス配列
 * @param {string}   fromTimeIso - 検索開始日時 ISO 文字列 (例: "2026-03-25T09:00+09:00")
 * @param {string}   toTimeIso   - 検索終了日時 ISO 文字列 (例: "2026-03-25T18:00+09:00")
 * @returns {{ success: boolean, slots: Array, checkedEmails?: number, warning?: string, message?: string }}
 */
function findCommonFreeSlots(emails, fromTimeIso, toTimeIso) {
  try {
    if (!emails || emails.length === 0) {
      return { success: false, slots: [], message: 'メールアドレスが指定されていません。' };
    }

    // フロントエンドから送られた ISO 文字列を Date オブジェクトに変換
    // "+09:00" サフィックスが付いていれば UTC への変換が正確に行われる
    const fromTime = new Date(fromTimeIso);
    const toTime   = new Date(toTimeIso);

    if (isNaN(fromTime.getTime()) || isNaN(toTime.getTime())) {
      return { success: false, slots: [], message: '日時の形式が無効です。' };
    }
    if (fromTime >= toTime) {
      return { success: false, slots: [], message: '開始日時は終了日時より前に設定してください。' };
    }

    // Calendar FreeBusy API クエリ
    const requestBody = {
      timeMin:  fromTime.toISOString(),
      timeMax:  toTime.toISOString(),
      timeZone: 'Asia/Tokyo',
      items:    emails.map(email => ({ id: email }))
    };

    const freeBusyResponse = Calendar.Freebusy.query(requestBody);
    const calendars = freeBusyResponse.calendars || {};

    // 全ユーザーのビジースロットを収集
    const allBusySlots = [];
    const errors = [];

    emails.forEach(email => {
      const calData = calendars[email];
      if (!calData) return;

      // アクセスエラー (カレンダーが非公開など)
      if (calData.errors && calData.errors.length > 0) {
        errors.push(email + ': ' + calData.errors.map(e => e.reason).join(', '));
        return;
      }

      (calData.busy || []).forEach(slot => {
        allBusySlots.push({ start: new Date(slot.start), end: new Date(slot.end) });
      });
    });

    // ビジースロットを時系列ソートし、重複区間をマージ
    allBusySlots.sort((a, b) => a.start - b.start);

    const mergedBusy = [];
    allBusySlots.forEach(slot => {
      if (mergedBusy.length === 0) {
        mergedBusy.push({ start: slot.start, end: slot.end });
      } else {
        const last = mergedBusy[mergedBusy.length - 1];
        if (slot.start <= last.end) {
          // 区間オーバーラップ: 終了時刻を延長
          if (slot.end > last.end) last.end = slot.end;
        } else {
          mergedBusy.push({ start: slot.start, end: slot.end });
        }
      }
    });

    // ビジー区間の補集合 = 空き時間スロット
    const freeSlots = [];
    let cursor = fromTime;

    mergedBusy.forEach(busy => {
      if (cursor < busy.start) {
        freeSlots.push({ start: new Date(cursor), end: new Date(busy.start) });
      }
      if (busy.end > cursor) cursor = busy.end;
    });

    // 最終ビジースロット以降の残り空き時間
    if (cursor < toTime) {
      freeSlots.push({ start: new Date(cursor), end: new Date(toTime) });
    }

    // JST (Asia/Tokyo) で表示用ラベルを生成
    const formattedSlots = freeSlots.map(slot => {
      const startDateJst = Utilities.formatDate(slot.start, 'Asia/Tokyo', 'MM/dd');
      const endDateJst   = Utilities.formatDate(slot.end,   'Asia/Tokyo', 'MM/dd');
      const startTimeJst = Utilities.formatDate(slot.start, 'Asia/Tokyo', 'HH:mm');
      const endTimeJst   = Utilities.formatDate(slot.end,   'Asia/Tokyo', 'HH:mm');

      // 同日なら日付を1回だけ表示、日をまたぐ場合は両端に日付を付ける
      const label = startDateJst === endDateJst
        ? startDateJst + '  ' + startTimeJst + ' - ' + endTimeJst + ' (JST)'
        : startDateJst + ' ' + startTimeJst + ' - ' + endDateJst + ' ' + endTimeJst + ' (JST)';

      return {
        label:    label,
        startIso: slot.start.toISOString(),
        endIso:   slot.end.toISOString()
      };
    });

    const result = { success: true, slots: formattedSlots, checkedEmails: emails.length };
    if (errors.length > 0) {
      result.warning = '以下のカレンダーへのアクセスに問題がありました:\n' + errors.join('\n');
    }
    return result;

  } catch (e) {
    Logger.log('findCommonFreeSlots error: ' + e.message);
    return { success: false, slots: [], message: 'エラー: ' + e.message };
  }
}