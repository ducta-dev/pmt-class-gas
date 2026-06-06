/**
 * CỔNG TIẾP NHẬN WEBHOOK TẬP TRUNG (MAIN ROUTER)
 * URL Web App này sẽ được cấu hình chung cho tất cả các Automation Webhook trên AppSheet.
 */
function doPost(e) {
  try {
    // 1. Đọc và phân tích JSON Payload truyền sang từ AppSheet
    const requestData = JSON.parse(e.postData.contents);
    const action = requestData.action;
    
    // 2. Điều phối đến các hàm xử lý tương ứng theo Action
    switch (action) {
      case "GENERATE_SESSIONS":
        // Gọi hàm xử lý sinh thời khóa biểu trong file Class.gs
        const classId = requestData.class_id;
        if (!classId) return createJsonResponse("error", "Thiếu tham số class_id.");
        
        const sessionResult = routerGenerateSessions(classId);
        return createJsonResponse(sessionResult.status, sessionResult.message);
        
      case "GENERATE_ATTENDANCE":
        // Gọi hàm xử lý sinh danh sách điểm danh trong file Attendance.gs
        const sessionId = requestData.session_id;
        if (!sessionId) return createJsonResponse("error", "Thiếu tham số session_id.");
        
        const attendanceResult = routerGenerateAttendance(sessionId);
        return createJsonResponse(attendanceResult.status, attendanceResult.message);
        
      default:
        return createJsonResponse("error", "Hành động (Action): '" + action + "' không hợp lệ hoặc chưa được định nghĩa.");
    }
    
  } catch (err) {
    return createJsonResponse("error", "Lỗi xử lý tại Router chính: " + err.toString());
  }
}

/**
 * Hàm phụ trợ tạo phản hồi định dạng JSON chuẩn cho Web App
 */
function createJsonResponse(status, message) {
  const output = { status: status, message: message };
  return ContentService.createTextOutput(JSON.stringify(output))
                       .setMimeType(ContentService.MimeType.JSON);
}