# Các lệnh Clasp thông dụng

## Đăng nhập
```
clasp login
```

## Đăng nhập bằng tài khoản mới
```
clasp login --creds <path-to-json>
```

## Đăng xuất
```
clasp logout
```

## Tạo project mới
```
clasp create --title "Tên Project" --type sheets
```

## Tạo project dạng Web App
```
clasp create --title "Tên Project" --type webapp
```

## Tạo project Web App liên kết Google Sheet
```
clasp create --title "Tên Project" --type webapp --parentId <Sheet_ID>

clasp create --title "QuanLyHocSinh" --parentId "ID_CUA_SHEET_BAN_VUA_COPY"
```


Trong đó <Sheet_ID> là ID của Google Sheet bạn muốn liên kết.

---

# Các lệnh thao tác với Google Sheet

## Mở Google Sheet liên kết
```
clasp open --parent
```

## Lấy ID Google Sheet liên kết
```
clasp list
```

## Cập nhật script cho Google Sheet
```
clasp push
```

## Kéo script từ Google Sheet về local
```
clasp pull
```

## Push code lên Google Apps Script
```
clasp push
```

## Pull code từ Google Apps Script
```
clasp pull
```

## Xem log
```
clasp logs
```

## Deploy phiên bản mới
```
clasp version
clasp deploy
```

## Xem trạng thái project
```
clasp status
```

## Mở project trên trình duyệt
```
clasp open
```

## Xem help
```
clasp --help
```

## lệnh deploy
clasp push
clasp version
clasp deploy


https://script.google.com/macros/s/AKfycbyKgfdZ8rzA78Rd6UUbqKWC8DvtiLb66mXuXzh1GD0/dev