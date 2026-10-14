/**
 * ===================================================================
 * GOOGLE APPS SCRIPT - HỆ THỐNG ĐĂNG KÝ KIỂM TRA PHIÊN HỌC DAT
 * ===================================================================
 * 
 * HƯỚNG DẪN CÀI ĐẶT:
 * 1. Mở Google Sheet (Trang tính) bạn muốn lưu danh sách.
 * 2. Trên thanh menu, chọn: Tiện ích mở rộng (Extensions) -> Apps Script.
 * 3. Xóa hết mã cũ và dán toàn bộ đoạn code này vào file Code.gs.
 * 4. Nhấn biểu tượng "Lưu" (Save - Ctrl + S).
 * 5. Nhấn nút "Triển khai" (Deploy) -> "Tùy chọn triển khai mới" (New deployment).
 * 6. Chọn loại triển khai: "Ứng dụng web" (Web app).
 * 7. Thiết lập quan trọng:
 *    - Mô tả: DAT Registration Backend
 *    - Thực thi dưới dạng (Execute as): "Tôi" (Me - địa chỉ email của bạn)
 *    - Ai có quyền truy cập (Who has access): "Bất kỳ ai" (Anyone) -> RẤT QUAN TRỌNG ĐỂ KHÔNG BỊ LỖI 403
 * 8. Nhấn "Triển khai" (Deploy) -> Cấp quyền truy cập nếu Google yêu cầu.
 * 9. Sao chép "URL ứng dụng web" (Web App URL) và dán vào dòng 960 của file index.html.
 * ===================================================================
 */

// Tên trang tính con (Sheet tab) lưu dữ liệu (mặc định lấy sheet đầu tiên nếu không tìm thấy)
const SHEET_NAME = "DangKyDAT";

/**
 * Xử lý yêu cầu POST khi người dùng nhấn gửi đăng ký từ Website
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  // Đợi tối đa 10 giây để tránh xung đột khi nhiều người cùng gửi
  lock.tryLock(10000);

  try {
    let data;
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else {
      data = e.parameter;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    
    // Nếu chưa có sheet, tự động tạo mới
    if (!sheet) {
      sheet = ss.getActiveSheet();
      if (sheet.getName() !== SHEET_NAME && ss.getSheets().length === 1 && sheet.getLastRow() === 0) {
        sheet.setName(SHEET_NAME);
      } else if (!ss.getSheetByName(SHEET_NAME)) {
        sheet = ss.insertSheet(SHEET_NAME);
      }
    }

    // Nếu sheet còn trống, tạo dòng tiêu đề đẹp mắt
    if (sheet.getLastRow() === 0) {
      const headers = [
        "Thời Gian Đăng Ký",
        "Họ Và Tên Học Viên",
        "Ngày Sinh",
        "Khóa Học",
        "Giáo Viên Phụ Trách",
        "GHI CHÚ"
      ];
      sheet.appendRow(headers);
      
      // Định dạng dòng tiêu đề
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#4f46e5");
      headerRange.setFontColor("#ffffff");
      headerRange.setHorizontalAlignment("center");
      sheet.setFrozenRows(1);
    }

    // Thời gian hiện tại theo múi giờ Việt Nam
    const now = new Date();
    const timeZone = Session.getScriptTimeZone() || "Asia/Ho_Chi_Minh";
    const formattedTime = Utilities.formatDate(now, timeZone, "HH:mm:ss dd/MM/yyyy");

    // Hỗ trợ cả 2 chế độ: 1 học viên đơn lẻ HOẶC danh sách nhiều học viên cùng lúc (Batch Siêu Tốc)
    let studentsList = [];
    if (data.students && Array.isArray(data.students)) {
      studentsList = data.students;
    } else if (Array.isArray(data)) {
      studentsList = data;
    } else {
      studentsList = [data];
    }

    const rowsToAppend = studentsList.map(function(item) {
      return [
        formattedTime,
        item.fullname || "",
        item.dob ? formatDateDisplay(item.dob) : "",
        item.course || data.course || "",
        item.teacher || data.teacher || "",
        item.note || ""
      ];
    });

    if (rowsToAppend.length > 0) {
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, rowsToAppend.length, 6).setValues(rowsToAppend);
    }

    // Trả về phản hồi thành công chuẩn JSON (hỗ trợ cả status và result)
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      result: "success",
      count: rowsToAppend.length,
      message: "Đăng ký thành công " + rowsToAppend.length + " học viên!",
      timestamp: formattedTime
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    // Trả về phản hồi lỗi nếu có sự cố
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);

  } finally {
    lock.releaseLock();
  }
}

/**
 * Xử lý yêu cầu GET để kiểm tra trạng thái hoạt động của Web App (Health check)
 */
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: "ok",
    message: "Google Apps Script Backend DAT đang hoạt động bình thường!",
    time: new Date().toISOString()
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Hàm hỗ trợ định dạng ngày YYYY-MM-DD sang DD/MM/YYYY
 */
function formatDateDisplay(dateStr) {
  if (!dateStr) return "";
  const parts = dateStr.split("-");
  if (parts.length === 3) {
    return parts[2] + "/" + parts[1] + "/" + parts[0];
  }
  return dateStr;
}
