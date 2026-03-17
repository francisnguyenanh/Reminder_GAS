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
    // Tiêu đề cột mới kèm Tracking
    sheet.appendRow(['ID', 'Tiêu đề', 'Webhook URL', 'Các thứ trong tuần', 'Giờ gửi', 'Nội dung tin nhắn', 'Trạng thái', 'CreatedBy', 'CreatedAt', 'LastModifiedBy', 'LastModifiedAt']);
    sheet.getRange("A1:K1").setFontWeight("bold");
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
    const configs = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      const createdBy = row[7];
      const canEdit = (createdBy === currentUser) || (currentUser === '');
      configs.push({
        id: row[0],
        title: row[1],
        webhookUrl: row[2],
        days: row[3],
        time: (typeof row[4] === 'object' && row[4] instanceof Date)
          ? Utilities.formatDate(row[4], 'GMT+7', 'HH:mm')
          : row[4],
        message: row[5],
        status: row[6],
        createdBy: createdBy,
        createdAt: row[8],
        lastModifiedBy: row[9],
        lastModifiedAt: row[10],
        canEdit: canEdit
      });
    }
    Logger.log({ configs, currentUser });
    return {
      configs: configs,
      currentUser: currentUser
    };
  } catch (e) {
    Logger.log('getConfigs ERROR: ' + e);
    return {
      configs: [],
      currentUser: '',
      errorMessage: 'getConfigs ERROR: ' + e
    };
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
    const nowTime = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
    const daysString = formData.days.join(', ');
    const data = sheet.getDataRange().getValues();
    
    if (formData.id) {
      // Logic Sửa (Update)
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === formData.id) {
          const createdBy = data[i][7];
          if (createdBy !== email && email !== '') {
            return { success: false, message: 'Bạn không có quyền sửa bản ghi này!' };
          }
          
          // Row theo sheet = index i + 1 (vì mảng JS từ 0)
          const rowNum = i + 1;
          
          // Cập nhật Cột 2->7: [Tiêu đề, Webhook, Các thứ, Giờ, Nội dung, Trạng thái]
          sheet.getRange(rowNum, 2, 1, 6).setValues([[
            formData.title, formData.webhookUrl, daysString, formData.time, formData.message, formData.status || 'Active'
          ]]);
          // Cập nhật Cột 10->11: [LastModifiedBy, LastModifiedAt]
          sheet.getRange(rowNum, 10, 1, 2).setValues([[email, nowTime]]);
          
          return { success: true, message: 'Cập nhật cấu hình thành công!' };
        }
      }
      return { success: false, message: 'Không tìm thấy ID bản ghi để cập nhật!' };
      
    } else {
      // Logic Thêm (Create)
      const newId = Utilities.getUuid();
      sheet.appendRow([
        newId,               // 0: ID
        formData.title,          // 1: Tiêu đề
        formData.webhookUrl,     // 2: Webhook
        daysString,          // 3: Thứ
        formData.time,           // 4: Giờ (HH:mm)
        formData.message,        // 5: Nội dung
        'Active',            // 6: Trạng thái
        email,               // 7: CreatedBy
        nowTime,             // 8: CreatedAt
        '',                  // 9: LastModifiedBy
        ''                   // 10: LastModifiedAt
      ]);
      
      return { success: true, message: 'Lưu cấu hình mới thành công!' };
    }
  } catch (error) {
    return { success: false, message: 'Lỗi hệ thống: ' + error.message };
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
          return { success: false, message: 'Bạn không có quyền xóa bản ghi này!' };
        }
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Đã xóa cấu hình!' };
      }
    }
    return { success: false, message: 'Không tìm thấy cấu hình!' };
  } catch(error) {
    return { success: false, message: 'Lỗi khi xóa: ' + error.message };
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
  const currentDay = Utilities.formatDate(now, "GMT+7", "EEEE"); 
  const currentTime = Utilities.formatDate(now, "GMT+7", "HH:mm"); 
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    // Các vị trí cột mới:
    // 0: ID, 1: Tiêu đề, 2: Webhook, 3: Thứ, 4: Giờ, 5: Nội dung, 6: Trạng thái
    const status = row[6];
    const scheduledTime = row[4];
    const scheduledDays = row[3] || "";
    
    if (status === 'Active' && scheduledTime === currentTime && scheduledDays.includes(currentDay)) {
       sendMessageToChat(row[2], row[5]);
    }
  }
}

/**
 * Gửi tin nhắn đến Google Chat Webhook
 */
function sendMessageToChat(webhookUrl, message) {
  if (!webhookUrl) return;
  const payload = { text: message };
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