const ExcelJS = require("exceljs");
const path = require("path");
const { leaveFileConfig: defaultLeaveFileConfig } = require("../config/excelConfig");

/**
 * Cập nhật file leaveFile: ghi vào cột Ghi chú và highlight dòng
 * @param {string|Buffer} filePathOrBuffer - Đường dẫn đến file leaveFile hoặc Buffer
 * @param {Array} checkResults - Kết quả check từ checkAttendanceData
 * @param {string} outputPath - Đường dẫn file output (nếu không có sẽ tạo mới, chỉ dùng khi filePath là string)
 * @param {Object} config - Config để parse file (nếu không có sẽ dùng default)
 * @returns {Promise<Buffer|string>} - Buffer nếu input là buffer, hoặc đường dẫn file nếu input là filePath
 */
const updateLeaveFile = async (filePathOrBuffer, checkResults, outputPath = null, config = null) => {
  const leaveFileConfigToUse = config || defaultLeaveFileConfig;
  const isBuffer = Buffer.isBuffer(filePathOrBuffer);
  
  const workbook = new ExcelJS.Workbook();
  
  // Hỗ trợ cả file path và buffer
  if (isBuffer) {
    await workbook.xlsx.load(filePathOrBuffer);
  } else {
    await workbook.xlsx.readFile(filePathOrBuffer);
  }
  
  const worksheet = workbook.worksheets[leaveFileConfigToUse.sheetIndex];
  if (!worksheet) {
    throw new Error(`Sheet index ${leaveFileConfigToUse.sheetIndex} không tồn tại`);
  }
  
  // Tìm cột "Ghi chú" bằng cách tìm trong header
  const headerRow = worksheet.getRow(leaveFileConfigToUse.headerRow);
  let noteColumnIndex = null;
  
  // Tìm cột "Ghi chú" trong header
  for (let colIndex = 1; colIndex <= worksheet.columnCount; colIndex++) {
    const headerCell = headerRow.getCell(colIndex);
    const headerValue = headerCell?.value?.toString()?.toLowerCase() || '';
    if (headerValue.includes('ghi chú') || headerValue.includes('ghichu')) {
      noteColumnIndex = colIndex;
      break;
    }
  }
  
  // Nếu không tìm thấy, dùng cột R (18) làm mặc định
  if (!noteColumnIndex) {
    noteColumnIndex = 18; // R = 18
  }
  
  const startColIndex = 1; // A = 1
  const endColIndex = leaveFileConfigToUse.endColumn ? 
    columnToNumber(leaveFileConfigToUse.endColumn) : worksheet.columnCount;
  
  // Đếm số dòng có issue để debug
  const issuesCount = checkResults.filter(r => r.hasIssue === true).length;
  console.log(`[UPDATE FILE] Tổng số dòng có issue: ${issuesCount} / ${checkResults.length}`);
  
  // Tạo Set để track các row number đã được highlight (để tránh duplicate)
  const highlightedRows = new Set();
  
  // Cập nhật từng dòng có issue
  checkResults.forEach((result, index) => {
    // Chỉ xử lý nếu thực sự có issue (kiểm tra chặt chẽ)
    // hasIssue phải là true (boolean), có rowNumber, và có message
    if (result.hasIssue === true && result.rowNumber && result.message && result.datesWithAttendance && result.datesWithAttendance.length > 0) {
      // Tránh highlight duplicate
      if (highlightedRows.has(result.rowNumber)) {
        console.log(`[UPDATE FILE] Row ${result.rowNumber} đã được highlight rồi, bỏ qua`);
        return;
      }
      
      console.log(`[UPDATE FILE] Highlighting row ${result.rowNumber} for MNV ${result.mnv} - ${result.datesWithAttendance.length} ngày có issue`);
      highlightedRows.add(result.rowNumber);
      const row = worksheet.getRow(result.rowNumber);
      
      // Ghi vào cột Ghi chú
      const noteCell = row.getCell(noteColumnIndex);
      const existingNote = noteCell.value?.toString() || '';
      const newNote = result.message;
      
      // Nếu đã có ghi chú, thêm vào cuối, ngược lại ghi mới
      if (existingNote && existingNote.trim() !== '') {
        noteCell.value = `${existingNote}; ${newNote}`;
      } else {
        noteCell.value = newNote;
      }
      
      // Chỉ tô đỏ chữ ở cột "Ghi chú" - CHỈ những dòng có issue
      noteCell.font = {
        ...noteCell.font,
        color: { argb: 'FFFF0000' } // Màu đỏ
      };
    } else {
      // Debug: log những dòng không có issue
      if (result.rowNumber && !result.hasIssue) {
        // Không làm gì - đảm bảo không highlight
      }
    }
  });
  
  // Nếu input là buffer, trả về buffer
  if (isBuffer) {
    const buffer = await workbook.xlsx.writeBuffer();
    return buffer;
  }
  
  // Nếu input là file path, lưu vào file
  if (!outputPath) {
    const dir = path.dirname(filePathOrBuffer);
    const ext = path.extname(filePathOrBuffer);
    const name = path.basename(filePathOrBuffer, ext);
    const timestamp = Date.now();
    outputPath = path.join(dir, `${name}_checked_${timestamp}${ext}`);
  }
  
  // Lưu file
  await workbook.xlsx.writeFile(outputPath);
  
  return outputPath;
};

/**
 * Chuyển đổi tên cột (A, B, C...) sang số (1, 2, 3...)
 */
const columnToNumber = (column) => {
  let result = 0;
  for (let i = 0; i < column.length; i++) {
    result = result * 26 + (column.charCodeAt(i) - "A".charCodeAt(0) + 1);
  }
  return result;
};

module.exports = {
  updateLeaveFile
};

