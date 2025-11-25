const ExcelJS = require("exceljs");
const dayjs = require("dayjs");
const {
  insuranceFileConfig: defaultInsuranceConfig,
  defaultInsuranceColumnMappings,
} = require("../config/excelConfig");

const columnLetterToNumber = (letter) => {
  if (!letter) return null;
  const clean = letter.replace(/[^A-Za-z]/g, "").toUpperCase();
  if (!clean) return null;
  let sum = 0;
  for (let i = 0; i < clean.length; i++) {
    sum *= 26;
    sum += clean.charCodeAt(i) - 64;
  }
  return sum;
};

const excelSerialToDate = (value) => {
  const excelEpoch = dayjs("1899-12-30");
  return excelEpoch.add(value, "day");
};

const parseDateValue = (value) => {
  if (value === null || value === undefined || value === "") {
    return null;
  }

  if (typeof value === "number") {
    return excelSerialToDate(value);
  }

  if (value instanceof Date) {
    return dayjs(value);
  }

  const str = String(value).trim();
  if (!str) return null;

  // Thử parse theo nhiều định dạng
  const formats = ["DD/MM/YYYY", "D/M/YYYY", "YYYY-MM-DD", "DD-MM-YYYY", "DD/MM/YY", "D/M/YY"];
  for (const format of formats) {
    const parsed = dayjs(str, format, true);
    if (parsed.isValid()) {
      return parsed;
    }
  }

  const iso = dayjs(str);
  if (iso.isValid()) {
    return iso;
  }

  return null;
};

const unwrapCellValue = (value) => {
  if (value === null || value === undefined) return null;

  if (typeof value === "object") {
    if (value.result !== undefined) return value.result;
    if (value.richText) return value.richText.map((item) => item.text).join("");
    if (value.text) return value.text;
    if (value.sharedFormula) return value.value;
    if (value.formula) return value.result ?? value.formula;
    if (value.hyperlink) return value.text || value.hyperlink;
  }

  return value;
};

const getValueFromLeaveRow = (row, columnLetter) => {
  if (!row || !columnLetter) return null;
  const key = `_col_${columnLetter.toUpperCase()}`;
  if (row.hasOwnProperty(key)) {
    const raw = unwrapCellValue(row[key]);
    if (raw && typeof raw === "object" && raw.result !== undefined) {
      return unwrapCellValue(raw.result);
    }
    return raw;
  }
  return null;
};

const normalizeKey = (value) => {
  const raw = unwrapCellValue(value);
  if (raw === null || raw === undefined) return null;
  return String(raw).trim();
};

const formatValueByType = (value, type, dateFormat = "DD/MM/YYYY") => {
  if (value === null || value === undefined) return null;
  if (type === "date") {
    const parsed = parseDateValue(value);
    if (parsed && parsed.isValid()) {
      return parsed.format(dateFormat);
    }
  }
  return unwrapCellValue(value);
};

const updateInsuranceFile = async (
  insuranceFileBuffer,
  leaveData,
  {
    insuranceConfig = defaultInsuranceConfig,
    columnMappings = defaultInsuranceColumnMappings,
  } = {}
) => {
  if (!columnMappings || columnMappings.length === 0) {
    throw new Error("Thiếu cấu hình cột mapping giữa file nghỉ phép và file BHXH.");
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(insuranceFileBuffer);

  const worksheet = workbook.worksheets[insuranceConfig.sheetIndex];
  if (!worksheet) {
    throw new Error(`Sheet index ${insuranceConfig.sheetIndex} không tồn tại trong file BHXH.`);
  }

  const startRow = insuranceConfig.startRow || (insuranceConfig.headerRow || 1) + 1;
  const dateFormat = insuranceConfig.dateOutputFormat || "DD/MM/YYYY";

  leaveData.forEach((leaveRow, index) => {
    const targetRowNumber = startRow + index;
    const targetRow = worksheet.getRow(targetRowNumber);
    console.log(`[BHXH] Đang ghi dữ liệu vào dòng ${targetRowNumber}`);

    columnMappings.forEach((mapping) => {
      if (!mapping?.sourceColumn || !mapping?.targetColumn) return;
      const sourceValue = getValueFromLeaveRow(leaveRow, mapping.sourceColumn);
      const formattedValue = formatValueByType(sourceValue, mapping.type, mapping.format || dateFormat);
      const targetColumnIndex = columnLetterToNumber(mapping.targetColumn);
      if (!targetColumnIndex) return;
      const cell = targetRow.getCell(targetColumnIndex);
      cell.value = formattedValue ?? null;
      console.log(`  - Ghi ${mapping.sourceColumn}->${mapping.targetColumn}:`, formattedValue ?? null);
    });

    targetRow.commit();
  });

  return await workbook.xlsx.writeBuffer();
};

module.exports = {
  updateInsuranceFile,
};

