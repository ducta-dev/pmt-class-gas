/**
 * Tiếp nhận yêu cầu POST Webhook từ AppSheet gửi sang
 */
function doPost(e) {
  try {
    // 1. Phân tích dữ liệu JSON nhận được từ cuộc gọi Webhook
    const requestData = JSON.parse(e.postData.contents);
    const classId = requestData.class_id;
    
    if (!classId) {
      return createJsonResponse("error", "Thiếu tham số class_id.");
    }
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const classSheet = spreadsheet.getSheetByName("tb_classes");
    const classData = classSheet.getDataRange().getValues();
    
    let classRowIndex = -1;
    let startDateRaw = null;
    let lessonsPerTerm = 0;
    let isTriggerActive = false;
    
    // 2. Tìm dòng chứa thông tin lớp học tương ứng và kiểm tra trạng thái trigger hiện tại
    for (let i = 1; i < classData.length; i++) {
      if (classData[i][0] === classId) { // Cột A: class_id
        classRowIndex = i + 1; // Chỉ số dòng thực tế trên Sheets (1-based index)
        startDateRaw = classData[i][5]; // Cột F (Index 5): start_date
        lessonsPerTerm = parseInt(classData[i][7]); // Cột H (Index 7): lessons_per_term
        isTriggerActive = classData[i][10]; // Cột K (Index 10): trigger
        break;
      }
    }
    
    if (classRowIndex === -1) {
      return createJsonResponse("error", "Không tìm thấy lớp học có ID: " + classId);
    }
    
    // BIỆN PHÁP PHÒNG CHỐNG TRÙNG LẶP (Idempotency Lock)
    // Nếu trạng thái cột trigger không phải là TRUE, nghĩa là yêu cầu này đã được xử lý xong từ trước.
    // Lập tức bỏ qua để tránh trùng lặp do AppSheet tự động gửi lại (retry) Webhook.
    if (isTriggerActive !== true && isTriggerActive !== "TRUE" && isTriggerActive !== "Y" && isTriggerActive !== "Yes") {
      return createJsonResponse("ignored", "Yêu cầu đã được xử lý hoặc trigger không hoạt động.");
    }
    
    // TẮT TRIGGER NGAY LẬP TỨC để khóa tài nguyên trước khi tiến hành tính toán
    classSheet.getRange(classRowIndex, 11).setValue(false);
    SpreadsheetApp.flush(); // Ép Google Sheets ghi nhận và đồng bộ giá trị FALSE xuống DB ngay lập tức
    
    if (isNaN(lessonsPerTerm) || lessonsPerTerm <= 0) {
      return createJsonResponse("error", "Số buổi học định mức (lessons_per_term) không hợp lệ.");
    }
    
    // 3. Tiến hành tự động sinh lịch học chi tiết cho kỳ học này
    generateSessions(spreadsheet, classId, lessonsPerTerm, startDateRaw, classRowIndex);
    
    return createJsonResponse("success", "Đã sinh lịch học chi tiết và tắt trigger thành công.");
    
  } catch (err) {
    return createJsonResponse("error", "Lỗi hệ thống: " + err.toString());
  }
}

/**
 * Hàm sinh lịch chi tiết vào bảng tb_class_sessions
 */
function generateSessions(spreadsheet, classId, lessonsPerTerm, startDateRaw, classRowIndex) {
  const scheduleSheet = spreadsheet.getSheetByName("tb_class_schedules");
  const sessionSheet = spreadsheet.getSheetByName("tb_class_sessions");
  const classSheet = spreadsheet.getSheetByName("tb_classes");
  
  const sessionData = sessionSheet.getDataRange().getValues();
  const scheduleData = scheduleSheet.getDataRange().getValues();
  
  // 1. Phân tích lịch học tuần của lớp này kèm theo Phòng, Giáo viên, Trợ giảng
  const schedules = []; 
  const dayMap = {
    "Chủ Nhật": 0, "Chủ nhật": 0, "CN": 0,
    "Thứ 2": 1, "Thứ hai": 1, "T2": 1,
    "Thứ 3": 2, "Thứ ba": 2, "T3": 2,
    "Thứ 4": 3, "Thứ tư": 3, "T4": 3,
    "Thứ 5": 4, "Thứ năm": 4, "T5": 4,
    "Thứ 6": 5, "Thứ sáu": 5, "T6": 5,
    "Thứ 7": 6, "Thứ bảy": 6, "T7": 6
  };
  
  for (let i = 1; i < scheduleData.length; i++) {
    if (scheduleData[i][1] === classId) { // Cột B: class_id
      schedules.push({
        dayOfWeek: dayMap[scheduleData[i][2]], // Cột C: day_of_week
        startTime: scheduleData[i][3],        // Cột D: start_time
        endTime: scheduleData[i][4],          // Cột E: end_time
        roomId: scheduleData[i][5] || "",     // Cột F: room_id
        teacherId: scheduleData[i][6] || "",  // Cột G: teacher_id
        assistantIds: scheduleData[i][7] || "" // Cột H: assistant_ids
      });
    }
  }
  
  if (schedules.length === 0) {
    throw new Error("Lớp học này chưa được khai báo Khung lịch học tuần.");
  }
  
  // 2. Kiểm tra xem đã có lịch học trước đó của lớp này hay chưa
  let maxDate = null;
  let lastSessionNum = 1000; // Định danh SES-1000 mặc định ban đầu
  
  for (let i = 1; i < sessionData.length; i++) {
    const currentId = sessionData[i][0];
    if (currentId && currentId.toString().indexOf("SES-") === 0) {
      const num = parseInt(currentId.toString().replace("SES-", ""), 10);
      if (!isNaN(num) && num > lastSessionNum) {
        lastSessionNum = num;
      }
    }
    
    if (sessionData[i][1] === classId) { // Cột B: class_id
      const sDate = parseDate(sessionData[i][2]); // Cột C: session_date
      if (sDate && !isNaN(sDate.getTime())) {
        if (!maxDate || sDate > maxDate) {
          maxDate = sDate;
        }
      }
    }
  }
  
  // 3. Xác định ngày bắt đầu tính toán ca học mới
  let currentDate = null;
  if (maxDate) {
    currentDate = new Date(maxDate.getTime());
    currentDate.setDate(currentDate.getDate() + 1);
  } else {
    currentDate = parseDate(startDateRaw);
  }
  
  if (!currentDate || isNaN(currentDate.getTime())) {
    throw new Error("Ngày bắt đầu không xác định được hoặc không hợp lệ.");
  }
  
  // 4. Chạy thuật toán tìm ngày học và phân bổ tài nguyên
  let sessionsCreated = 0;
  const sessionsToInsert = [];
  let safetyCounter = 0;
  let lastCreatedSessionDate = null;
  
  while (sessionsCreated < lessonsPerTerm && safetyCounter < 1000) {
    const currentDayOfWeek = currentDate.getDay(); // 0 (CN) - 6 (T7)
    const matchedSchedule = schedules.find(s => s.dayOfWeek === currentDayOfWeek);
    
    if (matchedSchedule) {
      sessionsCreated++;
      lastCreatedSessionDate = new Date(currentDate.getTime());
      
      const sessionId = "SES-" + (++lastSessionNum);
      const sessionDateFormatted = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      
      let startTimeStr = matchedSchedule.startTime;
      let endTimeStr = matchedSchedule.endTime;
      if (startTimeStr instanceof Date) {
        startTimeStr = Utilities.formatDate(startTimeStr, Session.getScriptTimeZone(), "HH:mm");
      }
      if (endTimeStr instanceof Date) {
        endTimeStr = Utilities.formatDate(endTimeStr, Session.getScriptTimeZone(), "HH:mm");
      }
      
      // Đổ dữ liệu khớp chính xác 100% cấu trúc 10 cột của bảng tb_class_sessions
      sessionsToInsert.push([
        sessionId,                     // A: session_id
        classId,                       // B: class_id
        sessionDateFormatted,          // C: session_date
        startTimeStr,                  // D: start_time
        endTimeStr,                    // E: end_time
        matchedSchedule.teacherId,     // F: teacher_id
        matchedSchedule.assistantIds,  // G: assistant_ids
        matchedSchedule.roomId,        // H: room_id
        "",                            // I: topic (để trống)
        "Chưa học"                     // J: session_status
      ]);
    }
    
    currentDate.setDate(currentDate.getDate() + 1);
    safetyCounter++;
  }
  
  // 5. Ghi dữ liệu và cập nhật bảng Lớp học
  if (sessionsToInsert.length > 0) {
    sessionSheet.getRange(sessionSheet.getLastRow() + 1, 1, sessionsToInsert.length, 10).setValues(sessionsToInsert);
    
    // Cập nhật ngày kết thúc mới (end_date) vào cột G (Cột 7) của bảng tb_classes
    if (lastCreatedSessionDate) {
      const endDateFormatted = Utilities.formatDate(lastCreatedSessionDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      classSheet.getRange(classRowIndex, 7).setValue(endDateFormatted);
    }
  }
}

/**
 * Hàm phân tích định dạng ngày linh hoạt từ Google Sheets
 */
function parseDate(dateVal) {
  if (dateVal instanceof Date) return dateVal;
  if (!dateVal) return null;
  
  const strVal = dateVal.toString().trim();
  const parts = strVal.split(/[-\/]/);
  if (parts.length === 3) {
    const months = {
      jan:0, feb:1, mar:2, apr:3, may:4, jun:5, jul:6, aug:7, sep:8, oct:9, nov:10, dec:11
    };
    const m = parts[1].toLowerCase();
    const monthIdx = months[m];
    if (monthIdx !== undefined) {
      return new Date(parts[2], monthIdx, parts[0]);
    }
    
    const day = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10) - 1;
    const year = parseInt(parts[2], 10);
    if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
      return new Date(year, month, day);
    }
  }
  
  const parsed = new Date(dateVal);
  return isNaN(parsed.getTime()) ? null : parsed;
}

/**
 * Hàm phụ tạo phản hồi JSON cho Web App
 */
function createJsonResponse(status, message) {
  const output = { status: status, message: message };
  return ContentService.createTextOutput(JSON.stringify(output))
                       .setMimeType(ContentService.MimeType.JSON);
}