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

// Cấu hình cho file BHXH (insuranceFile)
const insuranceFileConfig = {
  sheetIndex: 0,
  headerRow: 1,
  startRow: 2,
  startColumn: "A",
  endColumn: "T",
  dateOutputFormat: "DD/MM/YYYY",
};

// Mapping mặc định (có thể override từ frontend)
const defaultInsuranceColumnMappings = [
  // Ví dụ: Mã số BHXH (leave cột D) -> Số sổ BHXH (BHXH cột A)
  { sourceColumn: "D", targetColumn: "A", type: "text" },
  // Họ và tên (leave cột C) -> Họ và tên (BHXH cột B)
  { sourceColumn: "C", targetColumn: "B", type: "text" },
  // MNV (leave cột B) -> Mã nhân viên (BHXH cột C)
  { sourceColumn: "B", targetColumn: "N", type: "text" },
  // Từ ngày -> cột Từ ngày
  { sourceColumn: "E", targetColumn: "C", type: "date" },
  // Đến ngày -> cột Đến ngày
  { sourceColumn: "F", targetColumn: "D", type: "date" },
  // Thông tin tài khoản -> Số tài khoản ngân hàng
  { sourceColumn: "Q", targetColumn: "K", type: "text" },
  { sourceColumn: "C", targetColumn: "L", type: "text" },
];

module.exports = {
  leaveFileConfig,
  attendanceFileConfig,
  insuranceFileConfig,
  defaultInsuranceColumnMappings,
};

