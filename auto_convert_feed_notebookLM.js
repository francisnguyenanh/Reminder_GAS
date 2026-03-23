// Cấu hình ID
const FOLDER_INPUT_ID = '1BEitbYY-Tm5EuWMj0HkEYKZ1myCHHhDT';
const MASTER_SHEET_ID = '1pcuNzwTweHJFxsQQndbv38mxSnBp0KFySUiO2DFIvew';
const MASTER_DOC_ID = '17QlshdWg9Y83tgrdO42lZhUcW0pTllkcUbBsjdlA2eA';
const FOLDER_ARCHIVE_ID = '19HikINOkzVgS4apKKRmWQwroHtYG55OL'; // Để tránh quét lại file cũ

function mainProcess() {
  const folderInput = DriveApp.getFolderById(FOLDER_INPUT_ID);
  const archiveFolder = DriveApp.getFolderById(FOLDER_ARCHIVE_ID);
  const files = folderInput.getFiles();

  // 1. Xóa trắng các file Master trước khi nạp mới
  clearMasterFiles();

  while (files.hasNext()) {
    let file = files.next();
    let fileName = file.getName();
    let mimeType = file.getMimeType();

    try {
      // XỬ LÝ FILE EXCEL (.xlsx hoặc Google Sheets)
      if (mimeType === MimeType.GOOGLE_SHEETS || fileName.endsWith('.xlsx')) {
        processExcelToMaster(file);
      } 
      // XỬ LÝ FILE VĂN BẢN (.docx, .txt hoặc Google Docs)
      else if (mimeType === MimeType.GOOGLE_DOCS || fileName.endsWith('.docx') || fileName.endsWith('.txt')) {
        processDocToMaster(file);
      }
      // Thêm vào trong vòng lặp while của hàm mainProcess
      else if (mimeType === MimeType.JPEG || mimeType === MimeType.PNG || mimeType === MimeType.PDF) {
        processImageToMaster(file);
      }

      // Sau khi xử lý xong, chuyển vào folder Archive để tránh chạy lại lần sau
      file.moveTo(archiveFolder);
      console.log("Đã xử lý xong: " + fileName);

    } catch (e) {
      console.error("Lỗi file " + fileName + ": " + e.message);
    }
  }
}

// Hàm xử lý OCR cho Ảnh và PDF
function processImageToMaster(file) {
  const masterDoc = DocumentApp.openById(MASTER_DOC_ID);
  const body = masterDoc.getBody();

  // Cấu hình yêu cầu OCR của Drive API
  let resource = {
    name: "temp_ocr_" + file.getName(),
    mimeType: MimeType.GOOGLE_DOCS
  };
  
  // Bước quyết định: Dùng Drive API v3 để convert ảnh sang Doc (Tự động OCR)
  let tempFile = Drive.Files.create(resource, file.getBlob());
  
  // Lấy text từ file Doc tạm vừa tạo
  let textContent = DocumentApp.openById(tempFile.id).getBody().getText();
  
  // Ghi vào Master Doc
  body.appendParagraph("NGUỒN ẢNH/PDF: " + file.getName()).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(textContent);
  body.appendPageBreak();
  
  // Xóa file tạm sau khi lấy xong text
  Drive.Files.remove(tempFile.id);
}

function processExcelToMaster(file) {
  const masterSs = SpreadsheetApp.openById(MASTER_SHEET_ID);
  let sourceSs;

  if (file.getName().endsWith('.xlsx')) {
    // Convert Excel sang Google Sheets tạm thời bằng Drive API
    let blob = file.getBlob();
    let resource = {
      name: file.getName().replace('.xlsx', ''),
      mimeType: MimeType.GOOGLE_SHEETS
    };
    let tempFile = Drive.Files.create(resource, blob);
    sourceSs = SpreadsheetApp.openById(tempFile.id);
    
    // Copy các sheet
    sourceSs.getSheets().forEach(sheet => {
      sheet.copyTo(masterSs).setName(file.getName() + " - " + sheet.getName());
    });
    
    // Xóa file tạm
    Drive.Files.remove(tempFile.id);
  } else {
    // Nếu là Google Sheets sẵn thì copy luôn
    sourceSs = SpreadsheetApp.openById(file.getId());
    sourceSs.getSheets().forEach(sheet => {
      sheet.copyTo(masterSs).setName(file.getName() + " - " + sheet.getName());
    });
  }
}

function processDocToMaster(file) {
  const masterDoc = DocumentApp.openById(MASTER_DOC_ID);
  const body = masterDoc.getBody();
  let textContent = "";

  if (file.getName().endsWith('.docx')) {
    // Convert Docx sang Google Docs tạm thời
    let blob = file.getBlob();
    let resource = {
      name: 'temp_doc',
      mimeType: MimeType.GOOGLE_DOCS
    };
    let tempFile = Drive.Files.create(resource, blob);
    textContent = DocumentApp.openById(tempFile.id).getBody().getText();
    Drive.Files.remove(tempFile.id);
  } else if (file.getMimeType() === MimeType.GOOGLE_DOCS) {
    textContent = DocumentApp.openById(file.getId()).getBody().getText();
  } else {
    // File TXT
    textContent = file.getBlob().getDataAsString();
  }

  body.appendParagraph("NGUỒN: " + file.getName()).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(textContent);
  body.appendPageBreak();
}

function clearMasterFiles() {
  // Clear Doc
  DocumentApp.openById(MASTER_DOC_ID).getBody().clear();
  // Clear Sheet (Xóa các sheet cũ, giữ lại 1 sheet trống)
  const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
  const sheets = ss.getSheets();
  ss.insertSheet('TempClearSheet'); // Tạo sheet tạm
  sheets.forEach(s => ss.deleteSheet(s));
  ss.getSheets()[0].setName('Nội dung mới');
}