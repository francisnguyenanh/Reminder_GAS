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