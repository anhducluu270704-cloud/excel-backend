const dayjs = require("dayjs");

/**
 * Chuyển đổi định dạng ngày từ Excel sang dayjs
 * Hỗ trợ các format: DD/MM/YY, DD/MM/YYYY, hoặc Date object từ Excel
 */
const parseDate = (dateValue) => {
  if (!dateValue) return null;
  
  // Nếu là Date object từ Excel
  if (dateValue instanceof Date) {
    // Dùng dayjs với Date object, đảm bảo không bị ảnh hưởng bởi timezone
    return dayjs(dateValue);
  }
  
  // Nếu là number (serial date từ Excel)
  if (typeof dateValue === 'number') {
    // Excel date serial number
    const excelEpoch = dayjs('1900-01-01');
    return excelEpoch.add(dateValue - 2, 'day'); // Excel counts from 1900-01-01, but has a bug (treats 1900 as leap year)
  }
  
  // Nếu là string
  if (typeof dateValue === 'string') {
    // Nếu là ISO string (từ Date.toString())
    if (dateValue.includes('T') || dateValue.includes('Z')) {
      return dayjs(dateValue);
    }
    
    // Format DD/MM/YY hoặc DD/MM/YYYY
    const parts = dateValue.split('/');
    if (parts.length === 3) {
      let day = parseInt(parts[0]);
      let month = parseInt(parts[1]) - 1; // dayjs month is 0-indexed
      let year = parseInt(parts[2]);
      
      // Nếu năm 2 chữ số, convert sang 4 chữ số
      if (year < 100) {
        year = year + 2000; // Giả sử năm 2000+
      }
      
      return dayjs(new Date(year, month, day));
    }
  }
  
  return null;
};

/**
 * Normalize date string để so sánh nhất quán
 * Format: D/M/YY hoặc DD/MM/YY -> normalize về format nhất quán
 */
const normalizeDateString = (date) => {
  if (!date || !date.isValid()) return null;
  // Format: D/M/YY (không có số 0 đứng trước)
  return date.format('D/M/YY');
};

/**
 * So sánh và check dữ liệu giữa leaveFile và attendanceFile
 * @param {Array} leaveData - Dữ liệu đã parse từ file nghỉ phép
 * @param {Array} attendanceData - Dữ liệu đã parse từ file chấm công
 * @returns {Array} - Mảng các kết quả check với thông tin cần cập nhật
 */
const checkAttendanceData = (leaveData, attendanceData) => {
  const results = [];
  
  // Tạo map để tìm nhanh dữ liệu attendance theo Emp NO và Date
  const attendanceMap = new Map();
  
  attendanceData.forEach(attendance => {
    const empNo = attendance["Emp NO"]?.toString() || attendance["_col_D"]?.toString();
    const dateValue = attendance["Date"] || attendance["_col_H"];
    
    if (empNo && dateValue) {
      // Parse date từ nhiều format khác nhau
      const date = parseDate(dateValue);
      if (date && date.isValid()) {
        const normalizedDateStr = normalizeDateString(date);
        if (normalizedDateStr) {
          const key = `${empNo}_${normalizedDateStr}`;
          if (!attendanceMap.has(key)) {
            attendanceMap.set(key, []);
          }
          attendanceMap.get(key).push(attendance);
        }
      }
    }
  });
  
  // Check từng dòng trong leaveFile
  leaveData.forEach((leave, index) => {
    const mnv = leave["MNV"]?.toString() || leave["_col_B"]?.toString();
    // Lấy giá trị ngày trực tiếp (có thể là Date object hoặc string)
    const fromDateValue = leave["Từ ngày"] || leave["_col_E"];
    const toDateValue = leave["Đến ngày"] || leave["_col_F"];
    const rowNumber = leave._rowNumber;
    
    if (!mnv || !fromDateValue || !toDateValue) {
      results.push({
        leaveIndex: index,
        rowNumber: rowNumber,
        mnv: mnv,
        hasIssue: false,
        message: null,
        datesWithAttendance: []
      });
      return;
    }
    
    // Parse date - có thể là Date object hoặc string
    const fromDate = parseDate(fromDateValue);
    const toDate = parseDate(toDateValue);
    
    if (!fromDate || !toDate || !fromDate.isValid() || !toDate.isValid()) {
      results.push({
        leaveIndex: index,
        rowNumber: rowNumber,
        mnv: mnv,
        hasIssue: false,
        message: null,
        datesWithAttendance: [],
        debug: {
          fromDateValue: fromDateValue,
          toDateValue: toDateValue,
          fromDateValueType: typeof fromDateValue,
          toDateValueType: typeof toDateValue,
          fromDateParsed: fromDate ? fromDate.format('YYYY-MM-DD') : null,
          toDateParsed: toDate ? toDate.format('YYYY-MM-DD') : null
        }
      });
      return;
    }
    
    // Tìm các ngày trong khoảng có lỗi (2 loại lỗi độc lập: In1/Out1 và Pay1)
    const datesWithIssues = [];
    const debugKeys = []; // Để debug
    let currentDate = fromDate.clone(); // Clone để không mutate original
    
    while (currentDate.isBefore(toDate) || currentDate.isSame(toDate, 'day')) {
      const normalizedDateStr = normalizeDateString(currentDate);
      if (normalizedDateStr) {
        const key = `${mnv}_${normalizedDateStr}`;
        debugKeys.push(key);
        
        const attendances = attendanceMap.get(key) || [];
        
        // Nếu không có attendance cho ngày này, bỏ qua
        if (attendances.length === 0) {
          currentDate = currentDate.add(1, 'day');
          continue;
        }
        
        // Check từng attendance record
        attendances.forEach(attendance => {
          const issues = [];
          
          // ===== CHECK 1: In1/Out1 (độc lập) =====
          let in1 = attendance["In1"];
          let out1 = attendance["Out1"];
          
          // Nếu không có, thử lấy từ _col_L và _col_M
          if (in1 === undefined || in1 === null) {
            in1 = attendance["_col_L"];
          }
          if (out1 === undefined || out1 === null) {
            out1 = attendance["_col_M"];
          }
          
          // Convert sang string và trim
          const in1Str = in1 !== null && in1 !== undefined ? String(in1).trim() : '';
          const out1Str = out1 !== null && out1 !== undefined ? String(out1).trim() : '';
          
          // Check nếu có In1 VÀ Out1
          if (in1Str !== '' && out1Str !== '') {
            issues.push(`In/Out: ${in1Str}-${out1Str}`);
          }
          
          // ===== CHECK 2: Pay1 (độc lập) =====
          let pay1 = attendance["Pay1"];
          
          // Nếu không có, thử lấy từ _col_P
          if (pay1 === undefined || pay1 === null) {
            pay1 = attendance["_col_P"];
          }
          
          // Convert Pay1 sang number để check
          let pay1Value = 0;
          if (pay1 !== null && pay1 !== undefined) {
            // Nếu là number, dùng trực tiếp
            if (typeof pay1 === 'number') {
              pay1Value = pay1;
            } else {
              // Nếu là string, parse
              const pay1Str = String(pay1).trim();
              pay1Value = parseFloat(pay1Str);
              // Nếu parseFloat trả về NaN, thử parseInt
              if (isNaN(pay1Value)) {
                pay1Value = parseInt(pay1Str) || 0;
              }
            }
          }
          
          // Check Pay1 khác 0
          if (pay1Value !== 0 && !isNaN(pay1Value)) {
            issues.push(`Pay1: ${pay1Value}`);
            console.log(`[CHECK PAY1] Date: ${normalizedDateStr}, MNV: ${mnv}, Pay1 value: ${pay1Value}, Type: ${typeof pay1}, Original: ${pay1}`);
          }
          
          // Nếu có ít nhất 1 lỗi (In1/Out1 HOẶC Pay1), thêm vào danh sách
          if (issues.length > 0) {
            datesWithIssues.push({
              date: normalizedDateStr,
              in1: in1Str,
              out1: out1Str,
              pay1: pay1Value,
              issues: issues.join(', ')
            });
          }
        });
      }
      
      currentDate = currentDate.add(1, 'day');
    }
    
    // Nếu có ngày có lỗi, cần ghi chú và highlight
    const hasIssue = datesWithIssues.length > 0;
    let message = null;
    
    if (hasIssue) {
      const dateList = datesWithIssues.map(d => `${d.date} (${d.issues})`).join('; ');
      message = `Có chấm công trong ngày nghỉ: ${dateList}`;
      console.log(`[CHECK] MNV ${mnv} có issue: ${datesWithIssues.length} ngày có lỗi (In/Out hoặc Pay1)`);
    } else {
      console.log(`[CHECK] MNV ${mnv} không có issue`);
    }
    
    results.push({
      leaveIndex: index,
      rowNumber: rowNumber,
      mnv: mnv,
      hasIssue: hasIssue, // Đảm bảo là boolean
      message: message,
      datesWithAttendance: datesWithIssues,
      debug: {
        fromDateValue: fromDateValue,
        toDateValue: toDateValue,
        fromDateValueType: typeof fromDateValue,
        toDateValueType: typeof toDateValue,
        fromDateParsed: fromDate.format('YYYY-MM-DD'),
        toDateParsed: toDate.format('YYYY-MM-DD'),
        keysChecked: debugKeys,
        attendanceMapSize: attendanceMap.size
      }
    });
  });
  
  return results;
};

module.exports = {
  checkAttendanceData,
  parseDate
};

