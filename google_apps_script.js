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
// doGet: Đọc danh sách chi tiêu HOẶC ghi dữ liệu (bypass redirect)
// ============================================================
function doGet(e) {
  // Nếu có action=append → ghi dữ liệu (dùng khi POST bị redirect thành GET)
  if (e.parameter && e.parameter.action === "append" && e.parameter.payload) {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("Chi Tiêu Thực Tế");
      if (!sheet) sheet = ss.insertSheet("Chi Tiêu Thực Tế");
      var body = JSON.parse(e.parameter.payload);
      var rows = body.rows;
      if (rows && rows.length > 0) {
        for (var i = 0; i < rows.length; i++) {
          sheet.appendRow(rows[i]);
        }
      }
      return ContentService
        .createTextOutput(JSON.stringify({success: true, count: rows ? rows.length : 0}))
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

    // Thử parse JSON trực tiếp
    try {
      body = JSON.parse(raw);
    } catch (e1) {
      // Nếu không phải JSON, thử form field 'payload'
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

    var rows = body.rows;
    if (rows && rows.length > 0) {
      for (var i = 0; i < rows.length; i++) {
        sheet.appendRow(rows[i]);
      }
    }
    return ContentService
      .createTextOutput(JSON.stringify({success: true, count: rows ? rows.length : 0}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({success: false, error: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
