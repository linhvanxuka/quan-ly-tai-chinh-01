/**
 * GOOGLE APPS SCRIPT CHO TRANG TÍNH: Chi Tiêu Hàng Ngày (Hòa & Quỳnh)
 * -------------------------------------------------------------
 * Hướng dẫn cài đặt (Chỉ mất 1 phút):
 * 1. Trên trang Google Sheet của anh, bấm: Tiện ích mở rộng (Extensions) > Apps Script
 * 2. Xóa hết code cũ trong đó, dán toàn bộ đoạn code này vào.
 * 3. Bấm "Lưu" (biểu tượng đĩa mềm 💾).
 * 4. Bấm nút màu xanh "Triển khai" (Deploy) ở góc trên bên phải > "Quản lý triển khai" (Manage deployments).
 * 5. Bấm biểu tượng ✏️ (Edit) trên deployment hiện tại.
 * 6. Ở mục "Version" chọn "New version" (Phiên bản mới).
 * 7. Bấm "Deploy" → Copy URL mới → Dán vào app nếu URL thay đổi.
 *
 * LƯU Ý QUAN TRỌNG: Mỗi lần sửa code PHẢI tạo "New version" thì mới có hiệu lực!
 */

// ============================================================
// Hàm hỗ trợ: Ghi hoặc cập nhật dữ liệu và sắp xếp ngày bé -> lớn
// ============================================================
function saveOrUpdateAndSort(sheet, rows) {
  if (!rows || rows.length === 0) return 0;
  var data = sheet.getDataRange().getValues();
  
  for (var r = 0; r < rows.length; r++) {
    var newRow = rows[r]; // [date, desc, amount, category, owner, group, source, status]
    var newDate = String(newRow[0]).trim();
    var newAmt = Number(newRow[2]) || 0;
    
    // Tìm dòng trùng Ngày và Số tiền trên Sheet
    var foundIndex = -1;
    for (var i = 1; i < data.length; i++) {
      var existDate = String(data[i][0]).trim();
      var existAmt = Number(data[i][2]) || 0;
      if (existAmt >= 10000) existAmt = Math.round(existAmt / 1000);
      if (existDate === newDate && existAmt === newAmt) {
        foundIndex = i + 1; // 1-based index in sheet
        break;
      }
    }
    
    if (foundIndex > 0) {
      // Đã có trên Sheet -> Cập nhật nội dung và danh mục mới, KHÔNG tạo dòng trùng
      sheet.getRange(foundIndex, 1, 1, newRow.length).setValues([newRow]);
      data[foundIndex - 1] = newRow;
    } else {
      // Chưa có -> Thêm dòng mới
      sheet.appendRow(newRow);
      data.push(newRow);
    }
  }
  
  // Tự động sắp xếp từ ngày bé đến lớn (cột A tăng dần)
  var lastRow = sheet.getLastRow();
  if (lastRow > 2) {
    sheet.getRange(2, 1, lastRow - 1, 8).sort({column: 1, ascending: true});
  }
  
  return rows.length;
}

// ============================================================
// doGet: Đọc danh sách chi tiêu HOẶC ghi dữ liệu (bypass redirect)
// ============================================================
function doGet(e) {
  // Nếu có action=append → ghi dữ liệu
  if (e.parameter && e.parameter.action === "append" && e.parameter.payload) {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("Chi Tiêu Thực Tế");
      if (!sheet) sheet = ss.insertSheet("Chi Tiêu Thực Tế");
      var body = JSON.parse(e.parameter.payload);
      var count = saveOrUpdateAndSort(sheet, body.rows);
      return ContentService
        .createTextOutput(JSON.stringify({success: true, count: count}))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({success: false, error: err.toString()}))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Mặc định: trả về danh sách chi tiêu từ sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Chi Tiêu Thực Tế");
  if (!sheet) {
    return ContentService
      .createTextOutput(JSON.stringify({expenses: []}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var data = sheet.getDataRange().getValues();
  var expenses = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[0] && row[1]) {
      var amt = Number(row[2]) || 0;
      if (amt >= 10000) amt = Math.round(amt / 1000); // Chuẩn hóa về k VNĐ
      expenses.push({
        date: String(row[0]),
        desc: String(row[1] || ""),
        amount: amt,
        category: String(row[3] || ""),
        owner: String(row[4] || ""),
        group: String(row[5] || ""),
        source: String(row[6] || ""),
        status: String(row[7] || "")
      });
    }
  }
  return ContentService
    .createTextOutput(JSON.stringify({expenses: expenses}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// doPost: Nhận dữ liệu từ app và ghi vào sheet
// ============================================================
function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Chi Tiêu Thực Tế");
    if (!sheet) {
      sheet = ss.insertSheet("Chi Tiêu Thực Tế");
    }

    // Hỗ trợ cả JSON body thuần và form-encoded
    var body;
    var raw = (e.postData && e.postData.contents) ? e.postData.contents : "";

    try {
      body = JSON.parse(raw);
    } catch (e1) {
      try {
        var decoded = decodeURIComponent(raw.replace(/^payload=/, ""));
        body = JSON.parse(decoded);
      } catch (e2) {
        if (e.parameter && e.parameter.payload) {
          body = JSON.parse(e.parameter.payload);
        } else {
          throw new Error("Cannot parse body: " + raw.substring(0, 80));
        }
      }
    }

    var count = saveOrUpdateAndSort(sheet, body.rows);
    return ContentService
      .createTextOutput(JSON.stringify({success: true, count: count}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({success: false, error: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
