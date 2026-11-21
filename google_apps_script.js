/**
 * GOOGLE APPS SCRIPT CHO TRANG TÍNH: Chi Tiêu Hàng Ngày (Hòa & Quỳnh)
 * -------------------------------------------------------------
 * Hướng dẫn cài đặt (Chỉ mất 1 phút):
 * 1. Trên trang Google Sheet của anh, bấm: Tiện ích mở rộng (Extensions) > Apps Script
 * 2. Xóa hết code cũ trong đó, dán toàn bộ đoạn code này vào.
 * 3. Bấm "Lưu" (biểu tượng đĩa mềm 💾).
 * 4. Bấm nút màu xanh "Triển khai" (Deploy) ở góc trên bên phải > "Tùy chọn triển khai mới" (New deployment).
 * 5. Bấm vào biểu tượng bánh răng ⚙️ bên cạnh "Chọn loại", chọn "Ứng dụng web" (Web app).
 * 6. Điền thông tin:
 *    - Mô tả: "Quan Ly Tai Chinh App"
 *    - Thực thi dưới dạng (Execute as): "Tôi" (Me)
 *    - Ai có quyền truy cập (Who has access): "Bất kỳ ai" (Anyone) -> RẤT QUAN TRỌNG ĐỂ VERCEL GỌI ĐƯỢC
 * 7. Bấm "Triển khai" (Deploy). Nếu Google hỏi cấp quyền (Authorize Access), anh bấm Cho phép.
 * 8. Copy đường link "URL ứng dụng web" (có dạng: https://script.google.com/macros/s/.../exec).
 * 9. Dán URL này vào Vercel (Phần Environment Variables với tên biến: GSHEETS_URL).
 */

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Chi Tiêu Thực Tế");
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({expenses: []})).setMimeType(ContentService.MimeType.JSON);
  }
  var data = sheet.getDataRange().getValues();
  var expenses = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[0] && row[1]) {
      var amt = Number(row[2]) || 0;
      if (amt >= 10000) amt = Math.round(amt / 1000); // Chuẩn hóa về nghìn đồng (k)
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
  return ContentService.createTextOutput(JSON.stringify({expenses: expenses})).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Chi Tiêu Thực Tế");
    if (!sheet) {
      sheet = ss.insertSheet("Chi Tiêu Thực Tế");
    }
    
    // Hỗ trợ cả JSON body và form-encoded (payload field)
    var body;
    var raw = (e.postData && e.postData.contents) ? e.postData.contents : "";
    
    // Thử parse trực tiếp JSON
    try {
      body = JSON.parse(raw);
    } catch (e1) {
      // Nếu không phải JSON thuần, thử lấy field 'payload'
      try {
        var decoded = decodeURIComponent(raw.replace(/^payload=/, ""));
        body = JSON.parse(decoded);
      } catch (e2) {
        // Thử lấy từ parameter e.parameter
        if (e.parameter && e.parameter.payload) {
          body = JSON.parse(e.parameter.payload);
        } else {
          throw new Error("Cannot parse request body: " + raw.substring(0, 100));
        }
      }
    }
    
    var rows = body.rows;
    if (rows && rows.length > 0) {
      for (var i = 0; i < rows.length; i++) {
        sheet.appendRow(rows[i]);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({success: true, count: rows ? rows.length : 0})).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({success: false, error: err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

// doGet cũng xử lý được action=append (gửi qua query string khi POST bị redirect)
function doGet(e) {
  // Nếu có action=append → ghi dữ liệu
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
      return ContentService.createTextOutput(JSON.stringify({success: true, count: rows ? rows.length : 0})).setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({success: false, error: err.toString()})).setMimeType(ContentService.MimeType.JSON);
    }
  }
  
  // Mặc định: trả về danh sách chi tiêu
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Chi Tiêu Thực Tế");
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({expenses: []})).setMimeType(ContentService.MimeType.JSON);
  }
  var data = sheet.getDataRange().getValues();
  var expenses = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[0] && row[1]) {
      var amt = Number(row[2]) || 0;
      if (amt >= 10000) amt = Math.round(amt / 1000);
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
  return ContentService.createTextOutput(JSON.stringify({expenses: expenses})).setMimeType(ContentService.MimeType.JSON);
}
