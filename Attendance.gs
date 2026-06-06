/**
 * Hàm trung gian tiếp nhận yêu cầu từ Router chính (Main.gs) để xử lý tạo danh sách điểm danh.
 * Chịu trách nhiệm kiểm tra trạng thái an toàn dữ liệu, đọc danh sách học viên và ghi nhận điểm danh.
 * * @param {string} sessionId - ID của buổi học cần điểm danh (Ví dụ: SES-1001)
 * @return {object} Đối tượng chứa kết quả xử lý (status, message)
 */
function routerGenerateAttendance(sessionId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sessionSheet = spreadsheet.getSheetByName("tb_class_sessions");
  
  if (!sessionSheet) {
    return { status: "error", message: "Không tìm thấy bảng tb_class_sessions." };
  }
  
  const sessionData = sessionSheet.getDataRange().getValues();
  let sessionRowIndex = -1;
  let classId = null;
  let sessionStatus = "";
  let isTriggerActive = false;
  
  // 1. Tìm thông tin buổi học chi tiết trong tb_class_sessions
  for (let i = 1; i < sessionData.length; i++) {
    if (sessionData[i][0] === sessionId) { // Cột A (Index 0): session_id
      sessionRowIndex = i + 1; // Chỉ số dòng thực tế trên Sheets (1-based index)
      classId = sessionData[i][1]; // Cột B (Index 1): class_id
      sessionStatus = sessionData[i][9]; // Cột J (Index 9): session_status
      isTriggerActive = sessionData[i][10]; // Cột K (Index 10): trigger
      break;
    }
  }
  
  if (sessionRowIndex === -1) {
    return { status: "error", message: "Không tìm thấy buổi học có ID: " + sessionId };
  }
  
  // BIỆN PHÁP AN TOÀN TRÁNH ĐÈ DỮ LIỆU CŨ (Idempotency Check)
  if (sessionStatus === "Đã học") {
    return { status: "ignored", message: "Buổi học này đã được điểm danh từ trước." };
  }
  
  if (isTriggerActive !== true && isTriggerActive !== "TRUE" && isTriggerActive !== "Y" && isTriggerActive !== "Yes") {
    return { status: "ignored", message: "Yêu cầu không hợp lệ hoặc trigger không ở trạng thái kích hoạt." };
  }
  
  // 2. Gọi hàm nghiệp vụ chính để sinh danh sách điểm danh cho học viên
  try {
    generateAttendanceRecords(spreadsheet, sessionId, classId);
    
    // 3. Cập nhật trạng thái buổi học thành "Đã học" (Cột J - Cột thứ 10)
    // Giữ nguyên cột trigger là TRUE (Cột K - Cột thứ 11) để Action biến mất vĩnh viễn trên AppSheet
    sessionSheet.getRange(sessionRowIndex, 10).setValue("Đã học");
    sessionSheet.getRange(sessionRowIndex, 11).setValue(true); // Đảm bảo giữ nguyên TRUE
    SpreadsheetApp.flush(); // Đồng bộ dữ liệu lập tức xuống Sheets
    
    return { status: "success", message: "Khởi tạo danh sách điểm danh thành công cho buổi học: " + sessionId };
  } catch (err) {
    // Nếu phát sinh lỗi ngoài ý muốn, reset lại trigger về FALSE để quản lý có thể bấm lại nút Action
    sessionSheet.getRange(sessionRowIndex, 11).setValue(false);
    SpreadsheetApp.flush();
    return { status: "error", message: "Lỗi phát sinh khi tạo điểm danh: " + err.toString() };
  }
}

/**
 * Đọc danh sách học viên "Đang học" và ghi dữ liệu điểm danh mặc định xuống tb_attendance
 * (Đã điều chỉnh chỉ số Index khớp hoàn toàn với cấu trúc bảng tb_enrollments đã lược bỏ cột terms_paid_accumulated)
 */
function generateAttendanceRecords(spreadsheet, sessionId, classId) {
  const enrollmentSheet = spreadsheet.getSheetByName("tb_enrollments");
  const attendanceSheet = spreadsheet.getSheetByName("tb_attendance");
  
  if (!enrollmentSheet || !attendanceSheet) {
    throw new Error("Không tìm thấy bảng tb_enrollments hoặc tb_attendance.");
  }
  
  const enrollmentData = enrollmentSheet.getDataRange().getValues();
  const attendanceData = attendanceSheet.getDataRange().getValues();
  
  // 1. Quét tìm tất cả các học viên đang có trạng thái "Đang học" tại lớp này
  const activeStudents = [];
  
  for (let i = 1; i < enrollmentData.length; i++) {
    const enrollClassId = enrollmentData[i][1]; // Cột B (Index 1): class_id
    // Đã cập nhật: Do cột "terms_paid_accumulated" bị xóa, "enrollment_status" chuyển sang Cột F (Index 5)
    const enrollStatus = enrollmentData[i][5];  // Cột F (Index 5): enrollment_status
    
    if (enrollClassId === classId && enrollStatus === "Đang học") {
      activeStudents.push({
        studentId: enrollmentData[i][0],      // Cột A (Index 0): student_id
        enrollmentType: enrollmentData[i][3] // Cột D (Index 3): enrollment_type
      });
    }
  }
  
  if (activeStudents.length === 0) {
    throw new Error("Lớp học hiện tại không có học viên nào ở trạng thái 'Đang học'.");
  }
  
  // 2. Tìm mã ID điểm danh lớn nhất để thiết lập mã tự tăng (Dạng ATT-1001, ATT-1002...)
  let lastAttendanceNum = 1000;
  for (let i = 1; i < attendanceData.length; i++) {
    const currentId = attendanceData[i][0]; // Cột A (Index 0): attendance_id
    if (currentId && currentId.toString().indexOf("ATT-") === 0) {
      const num = parseInt(currentId.toString().replace("ATT-", ""), 10);
      if (!isNaN(num) && num > lastAttendanceNum) {
        lastAttendanceNum = num;
      }
    }
  }
  
  // 3. Khởi tạo mảng ghi điểm danh mặc định cho từng học viên
  const attendanceToInsert = [];
  
  for (let i = 0; i < activeStudents.length; i++) {
    const student = activeStudents[i];
    const attendanceId = "ATT-" + (++lastAttendanceNum);
    
    // Phân loại hình thức đi học dựa trên loại đăng ký học viên
    const attendanceType = (student.enrollmentType === "Học thử") ? "Học thử" : "Học chính thức";
    
    // Khớp cấu trúc 7 cột của bảng tb_attendance thực tế:
    attendanceToInsert.push([
      attendanceId,       // A: attendance_id
      sessionId,          // B: session_id
      student.studentId,  // C: student_id
      "Có mặt",           // D: attendance_status (Mặc định ban đầu)
      attendanceType,     // E: attendance_type
      "",                 // F: original_class_id (Chỉ dùng khi học bù)
      ""                  // G: note (Để trống cho giáo viên nhận xét sau)
    ]);
  }
  
  // 4. Ghi dữ liệu hàng loạt xuống Google Sheets để đảm bảo hiệu suất hoạt động cực nhanh
  if (attendanceToInsert.length > 0) {
    attendanceSheet.getRange(attendanceSheet.getLastRow() + 1, 1, attendanceToInsert.length, 7).setValues(attendanceToInsert);
  }
}