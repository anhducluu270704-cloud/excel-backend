const multer = require("multer");
const path = require("path");
const fs = require("fs");

// Tạm thời dùng disk storage để test
const excelStorage = multer.diskStorage({
  destination: (req, file, cb) => {
    const uploadDir = path.join(__dirname, "../uploads");
    if (!fs.existsSync(uploadDir)) {
      fs.mkdirSync(uploadDir, { recursive: true });
    }
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    const uniqueSuffix = Date.now() + "-" + Math.round(Math.random() * 1e9);
    cb(null, file.fieldname + "-" + uniqueSuffix + path.extname(file.originalname));
  },
});

// Chỉ cho phép file Excel
function fileFilter(req, file, cb) {
  console.log('fileFilter - Checking file:', {
    originalname: file.originalname,
    mimetype: file.mimetype,
    fieldname: file.fieldname
  });
  
  if (
    file.mimetype === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ||
    file.mimetype === "application/vnd.ms-excel" ||
    file.originalname.endsWith('.xlsx') ||
    file.originalname.endsWith('.xls')
  ) {
    console.log('fileFilter - File accepted');
    cb(null, true);
  } else {
    console.log('fileFilter - File rejected');
    cb(new Error("Chỉ được upload file Excel (.xlsx, .xls)"), false);
  }
}

const uploadExcel = multer({ 
  storage: excelStorage,
  fileFilter: fileFilter
});

// Helper function để tạo fields middleware
const fields = (fieldConfig) => {
  console.log('Creating fields middleware with config:', fieldConfig);
  const multerMiddleware = uploadExcel.fields(fieldConfig);
  
  // Wrap middleware để log khi được gọi
  return (req, res, next) => {
    console.log('=== Multer middleware called ===');
    console.log('Request method:', req.method);
    console.log('Request URL:', req.url);
    console.log('Content-Type:', req.headers['content-type']);
    console.log('Has body:', !!req.body);
    console.log('Has files (before multer):', !!req.files);
    
    // Kiểm tra Content-Type
    const contentType = req.headers['content-type'] || '';
    if (!contentType.includes('multipart/form-data')) {
      console.error('WARNING: Content-Type không phải multipart/form-data!');
      console.error('Content-Type hiện tại:', contentType);
      console.error('Tất cả headers:', req.headers);
      return res.status(400).json({
        message: 'Request phải có Content-Type: multipart/form-data',
        receivedContentType: contentType,
        hint: 'Đảm bảo trong Postman bạn chọn Body -> form-data (không phải x-www-form-urlencoded)'
      });
    }
    
    multerMiddleware(req, res, (err) => {
      if (err) {
        console.error('Multer middleware error:', err);
        return next(err);
      }
      console.log('Multer middleware completed');
      console.log('Has files (after multer):', !!req.files);
      if (req.files) {
        console.log('Files keys:', Object.keys(req.files));
        console.log('Files details:', JSON.stringify(Object.keys(req.files).map(key => ({
          key,
          count: req.files[key]?.length || 0,
          files: req.files[key]?.map(f => ({ originalname: f.originalname, mimetype: f.mimetype }))
        })), null, 2));
      }
      next();
    });
  };
};

module.exports = { uploadExcel, fields };
