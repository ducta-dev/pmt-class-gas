function doPost(e) {
  try {
    const requestData = JSON.parse(e.postData.contents);
    const action = requestData.action; // Lấy tên hành động từ AppSheet
    const classId = requestData.class_id;
    const studentId = requestData.student_id; // (Ví dụ cho module sau)

    // HỆ THỐNG ĐỊNH TUYẾN (ROUTER)
    switch (action) {
      case "GENERATE_SESSIONS":
        // Gọi hàm xử lý lớp học (nằm ở file Class.gs)
        return routerGenerateSessions(classId); 
        
      case "GENERATE_RECEIPT_K1":
        // Gọi hàm xử lý học phí (nằm ở file HocPhi.gs - giả định module sau)
        return routerGenerateReceiptK1(studentId, classId);
        
      default:
        return createJsonResponse("error", "Hành động (Action) không hợp lệ hoặc chưa được định nghĩa.");
    }

  } catch (err) {
    return createJsonResponse("error", "Lỗi hệ thống Router: " + err.toString());
  }
}

function createJsonResponse(status, message) {
  const output = { status: status, message: message };
  return ContentService.createTextOutput(JSON.stringify(output)).setMimeType(ContentService.MimeType.JSON);
}