const ExcelJS = require("exceljs");
const path = require("path");

/**
 * Parse file Excel và trả về dữ liệu theo cấu hình
 * @param {string|Buffer} filePathOrBuffer - Đường dẫn đến file Excel hoặc Buffer chứa file
 * @param {Object} config - Cấu hình parse
 * @param {number} config.sheetIndex - Index của sheet cần đọc (mặc định: 0)
 * @param {number} config.headerRow - Hàng chứa header (mặc định: 1)
 * @param {number} config.startRow - Hàng bắt đầu đọc dữ liệu (mặc định: headerRow + 1)
 * @param {string} config.startColumn - Cột bắt đầu (ví dụ: 'A', 'B', mặc định: 'A')
 * @param {string} config.endColumn - Cột kết thúc (ví dụ: 'D', 'E', mặc định: null - đọc đến hết)
 * @returns {Promise<Array>} - Mảng các object chứa dữ liệu từng hàng
 */
const parseExcelFile = async (filePathOrBuffer, config = {}) => {
  const {
    sheetIndex = 0,
    headerRow = 1, // Mặc định header ở hàng 1
    startRow = null, // Nếu null, sẽ tự động = headerRow + 1
    startColumn = "A",
    endColumn = null, // null = đọc đến hết
  } = config;

  // Nếu không chỉ định startRow, tự động lấy hàng sau header
  const actualStartRow = startRow || headerRow + 1;

  const workbook = new ExcelJS.Workbook();
  
  // Hỗ trợ cả file path và buffer
  if (Buffer.isBuffer(filePathOrBuffer)) {
    await workbook.xlsx.load(filePathOrBuffer);
  } else {
    await workbook.xlsx.readFile(filePathOrBuffer);
  }

  const worksheet = workbook.worksheets[sheetIndex];
  if (!worksheet) {
    throw new Error(`Sheet index ${sheetIndex} không tồn tại`);
  }

  const data = [];
  const startColIndex = columnToNumber(startColumn);
  const endColIndex = endColumn ? columnToNumber(endColumn) : worksheet.columnCount;

  // Đọc header từ hàng headerRow
  const headerRowData = worksheet.getRow(headerRow);
  const headers = {};
  for (let colIndex = startColIndex; colIndex <= endColIndex; colIndex++) {
    const headerCell = headerRowData.getCell(colIndex);
    const columnLetter = numberToColumn(colIndex);
    const headerValue = headerCell?.value;
    // Lấy header name, nếu rỗng thì dùng tên cột
    headers[colIndex] = headerValue?.toString()?.trim() || columnLetter;
  }

  // Đọc từng hàng dữ liệu
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber < actualStartRow) return; // Bỏ qua các hàng trước startRow

    const rowData = {};
    
    // Đọc từng cột trong phạm vi đã định
    for (let colIndex = startColIndex; colIndex <= endColIndex; colIndex++) {
      const cell = row.getCell(colIndex);
      const columnLetter = numberToColumn(colIndex);
      const headerName = headers[colIndex] || columnLetter;
      
      rowData[headerName] = cell.value;
      rowData[`_col_${columnLetter}`] = cell.value; // Giữ cả tên cột dạng A, B, C...
    }

    // Chỉ thêm hàng nếu có ít nhất 1 giá trị
    if (Object.values(rowData).some(val => val !== null && val !== undefined && val !== "")) {
      rowData._rowNumber = rowNumber;
      data.push(rowData);
    }
  });

  return data;
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

/**
 * Chuyển đổi số (1, 2, 3...) sang tên cột (A, B, C...)
 */
const numberToColumn = (number) => {
  let result = "";
  while (number > 0) {
    number--;
    result = String.fromCharCode(65 + (number % 26)) + result;
    number = Math.floor(number / 26);
  }
  return result;
};

module.exports = {
  parseExcelFile,
  columnToNumber,
  numberToColumn,
};

