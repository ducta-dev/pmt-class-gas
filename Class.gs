/**
 * Hàm trung gian tiếp nhận yêu cầu từ Router chính (doPost) để xử lý sinh lịch học.
 * Hàm này chịu trách nhiệm kiểm tra trạng thái Trigger, thực hiện Khóa chống trùng lặp (Idempotency Lock)
 * và gọi hàm sinh lịch chi tiết.
 * * @param {string} classId - ID của lớp học cần sinh lịch (Ví dụ: CLS-001)
 * @return {object} Đối tượng chứa kết quả xử lý (status, message)
 */
function routerGenerateSessions(classId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const classSheet = spreadsheet.getSheetByName("tb_classes");
  
  if (!classSheet) {
    return { status: "error", message: "Không tìm thấy bảng tb_classes." };
  }
  
  const classData = classSheet.getDataRange().getValues();
  let classRowIndex = -1;
  let startDateRaw = null;
  let lessonsPerTerm = 0;
  let isTriggerActive = false;
  
  // 1. Tìm dòng chứa thông tin lớp học và đọc trạng thái Trigger
  for (let i = 1; i < classData.length; i++) {
    if (classData[i][0] === classId) { // Cột A (Index 0): class_id
      classRowIndex = i + 1; // Dòng thực tế trên Google Sheets (1-based index)
      startDateRaw = classData[i][5]; // Cột F (Index 5): start_date
      lessonsPerTerm = parseInt(classData[i][7]); // Cột H (Index 7): lessons_per_term
      isTriggerActive = classData[i][10]; // Cột K (Index 10): trigger
      break;
    }
  }
  
  if (classRowIndex === -1) {
    return { status: "error", message: "Không tìm thấy lớp học có ID: " + classId };
  }
  
  // 2. BIỆN PHÁP PHÒNG CHỐNG TRÙNG LẶP (Idempotency Lock)
  // Nếu cột trigger không ở trạng thái hoạt động (TRUE/Y/Yes), lập tức bỏ qua.
  if (isTriggerActive !== true && isTriggerActive !== "TRUE" && isTriggerActive !== "Y" && isTriggerActive !== "Yes") {
    return { status: "ignored", message: "Yêu cầu đã được xử lý từ trước hoặc trigger không hoạt động." };
  }
  
  // 3. KHÓA TÀI NGUYÊN NGAY LẬP TỨC
  // Reset cột trigger về FALSE ngay lập tức để chặn các cuộc gọi thử lại (retry) từ AppSheet gửi đến sau
  classSheet.getRange(classRowIndex, 11).setValue(false); // Cột K (Cột 11)
  SpreadsheetApp.flush(); // Ép hệ thống đồng bộ ngay lập tức giá trị FALSE xuống DB Google Sheets
  
  if (isNaN(lessonsPerTerm) || lessonsPerTerm <= 0) {
    return { status: "error", message: "Số buổi học định mức (lessons_per_term) không hợp lệ." };
  }
  
  // 4. Tiến hành sinh lịch học chi tiết và cập nhật end_date
  try {
    generateSessions(spreadsheet, classId, lessonsPerTerm, startDateRaw, classRowIndex);
    return { status: "success", message: "Đã sinh thời khóa biểu chi tiết cho lớp " + classId };
  } catch (err) {
    // Nếu có lỗi xảy ra trong quá trình sinh lịch, khôi phục lại cột trigger về TRUE để người dùng bấm lại được
    classSheet.getRange(classRowIndex, 11).setValue(true);
    SpreadsheetApp.flush();
    return { status: "error", message: "Lỗi khi sinh lịch: " + err.toString() };
  }
}

/**
 * Thuật toán sinh thời khóa biểu chi tiết dựa trên lịch tuần của lớp học
 */
function generateSessions(spreadsheet, classId, lessonsPerTerm, startDateRaw, classRowIndex) {
  const scheduleSheet = spreadsheet.getSheetByName("tb_class_schedules");
  const sessionSheet = spreadsheet.getSheetByName("tb_class_sessions");
  const classSheet = spreadsheet.getSheetByName("tb_classes");
  
  if (!scheduleSheet || !sessionSheet || !classSheet) {
    throw new Error("Thiếu một trong các bảng tb_class_schedules, tb_class_sessions, tb_classes.");
  }
  
  const sessionData = sessionSheet.getDataRange().getValues();
  const scheduleData = scheduleSheet.getDataRange().getValues();
  
  // 1. Phân tích các khung ca học tuần của lớp này (kèm theo Phòng, Giáo viên, Trợ giảng riêng lẻ)
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
    if (scheduleData[i][1] === classId) { // Cột B (Index 1): class_id
      schedules.push({
        dayOfWeek: dayMap[scheduleData[i][2]], // Cột C (Index 2): day_of_week
        startTime: scheduleData[i][3],        // Cột D (Index 3): start_time
        endTime: scheduleData[i][4],          // Cột E (Index 4): end_time
        roomId: scheduleData[i][5] || "",     // Cột F (Index 5): room_id
        teacherId: scheduleData[i][6] || "",  // Cột G (Index 6): teacher_id
        assistantIds: scheduleData[i][7] || "" // Cột H (Index 7): assistant_ids
      });
    }
  }
  
  if (schedules.length === 0) {
    throw new Error("Lớp học này chưa được khai báo Khung lịch học tuần trong tb_class_schedules.");
  }
  
  // 2. Kiểm tra lịch sử của lớp này để quyết định sinh tiếp nối hay sinh mới hoàn toàn
  let maxDate = null;
  let lastSessionNum = 1000; // Định danh SES-1000 mặc định làm mốc tự tăng
  
  for (let i = 1; i < sessionData.length; i++) {
    // Quét tìm mã session_id lớn nhất trên toàn bộ bảng tb_class_sessions để tự tăng tiếp (ví dụ SES-1004 -> SES-1005)
    const currentId = sessionData[i][0];
    if (currentId && currentId.toString().indexOf("SES-") === 0) {
      const num = parseInt(currentId.toString().replace("SES-", ""), 10);
      if (!isNaN(num) && num > lastSessionNum) {
        lastSessionNum = num;
      }
    }
    
    // Tìm ngày của buổi học xa nhất hiện tại của lớp này
    if (sessionData[i][1] === classId) { // Cột B (Index 1): class_id
      const sDate = parseDate(sessionData[i][2]); // Cột C (Index 2): session_date
      if (sDate && !isNaN(sDate.getTime())) {
        if (!maxDate || sDate > maxDate) {
          maxDate = sDate;
        }
      }
    }
  }
  
  // 3. Xác định ngày khởi điểm để tính toán lịch dương
  let currentDate = null;
  if (maxDate) {
    // Đã có lịch cũ -> Bắt đầu tính lịch từ ngày hôm sau ngày cuối cùng đó
    currentDate = new Date(maxDate.getTime());
    currentDate.setDate(currentDate.getDate() + 1);
  } else {
    // Chưa có lịch -> Lấy ngày khai giảng start_date làm mốc khởi đầu
    currentDate = parseDate(startDateRaw);
  }
  
  if (!currentDate || isNaN(currentDate.getTime())) {
    throw new Error("Không xác định được ngày bắt đầu sinh lịch hợp lệ.");
  }
  
  // 4. Chạy thuật toán tịnh tiến ngày để tìm ngày khớp lịch tuần
  let sessionsCreated = 0;
  const sessionsToInsert = [];
  let safetyCounter = 0; // Tránh vòng lặp vô tận nếu cấu hình sai
  let lastCreatedSessionDate = null;
  
  while (sessionsCreated < lessonsPerTerm && safetyCounter < 1000) {
    const currentDayOfWeek = currentDate.getDay(); // 0 (CN) - 6 (T7)
    const matchedSchedule = schedules.find(s => s.dayOfWeek === currentDayOfWeek);
    
    if (matchedSchedule) {
      sessionsCreated++;
      lastCreatedSessionDate = new Date(currentDate.getTime());
      
      const sessionId = "SES-" + (++lastSessionNum);
      const sessionDateFormatted = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      
      // Định dạng lại giờ học
      let startTimeStr = matchedSchedule.startTime;
      let endTimeStr = matchedSchedule.endTime;
      if (startTimeStr instanceof Date) {
        startTimeStr = Utilities.formatDate(startTimeStr, Session.getScriptTimeZone(), "HH:mm");
      }
      if (endTimeStr instanceof Date) {
        endTimeStr = Utilities.formatDate(endTimeStr, Session.getScriptTimeZone(), "HH:mm");
      }
      
      // Đổ dữ liệu khớp chính xác 100% cấu trúc 10 cột của bảng tb_class_sessions thực tế
      sessionsToInsert.push([
        sessionId,                     // A: session_id
        classId,                       // B: class_id
        sessionDateFormatted,          // C: session_date
        startTimeStr,                  // D: start_time
        endTimeStr,                    // E: end_time
        matchedSchedule.teacherId,     // F: teacher_id
        matchedSchedule.assistantIds,  // G: assistant_ids
        matchedSchedule.roomId,        // H: room_id
        "",                            // I: topic (để trống khi tạo mới)
        "Chưa học"                     // J: session_status
      ]);
    }
    
    currentDate.setDate(currentDate.getDate() + 1);
    safetyCounter++;
  }
  
  // 5. Ghi mảng dữ liệu một lần duy nhất xuống Google Sheets để bảo toàn hiệu năng
  if (sessionsToInsert.length > 0) {
    sessionSheet.getRange(sessionSheet.getLastRow() + 1, 1, sessionsToInsert.length, 10).setValues(sessionsToInsert);
    
    // Cập nhật ngày kết thúc mới nhất (end_date) vào cột G (Cột thứ 7) của bảng tb_classes
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