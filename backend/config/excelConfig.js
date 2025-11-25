/**
 * Cấu hình parse file Excel
 * Bạn có thể thay đổi các giá trị này để phù hợp với cấu trúc file Excel của bạn
 */

// Cấu hình cho file nghỉ phép (leaveFile)
const leaveFileConfig = {
  sheetIndex: 0, // Sheet đầu tiên (index 0)
  headerRow: 12, // Hàng chứa header (STT, MNV, Họ và tên, ...)
  startRow: 13, // Bắt đầu từ hàng 13 (dữ liệu đầu tiên sau header)
  startColumn: "A", // Bắt đầu từ cột A
  endColumn: "S", // Kết thúc ở cột P (theo cấu trúc file)
};

// Cấu hình cho file chấm công (attendanceFile)
const attendanceFileConfig = {
  sheetIndex: 0, // Sheet đầu tiên (index 0)
  headerRow: 5, // Hàng chứa header (Dept Code, Sect Code, Emp NO, ...)
  startRow: 6, // Bắt đầu từ hàng 6 (dữ liệu đầu tiên sau header)
  startColumn: "A", // Bắt đầu từ cột A
  endColumn: "R", // Kết thúc ở cột R (LeaveReason)
};

module.exports = {
  leaveFileConfig,
  attendanceFileConfig,
};

