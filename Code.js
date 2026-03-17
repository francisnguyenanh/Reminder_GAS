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
    if ((row[11] || 'recurring') === 'once') {
      sheet.getRange(i + 1, 7).setValue('Inactive');
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
                        `__________________\n\n` + 
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