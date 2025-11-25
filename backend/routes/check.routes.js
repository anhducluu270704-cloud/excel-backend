const { Router } = require("express");
const router = Router();

const { fields } = require("../middleware/uploadExcel");
const { uploadTwoFiles, testLeaveFile, testAttendanceFile, downloadUpdatedFile } = require("../controllers/check.controller");

// Middleware để log request trước khi multer xử lý
const logRequest = (req, res, next) => {
  console.log('=== Request Info ===');
  console.log('Method:', req.method);
  console.log('URL:', req.url);
  console.log('Content-Type:', req.headers['content-type']);
  console.log('Content-Length:', req.headers['content-length']);
  console.log('Has body:', !!req.body);
  console.log('Raw headers:', req.headers);
  console.log('===================');
  next();
};

// Error handler middleware cho multer
const handleMulterError = (err, req, res, next) => {
  if (err) {
    console.error('Multer error:', err);
    console.error('Error code:', err.code);
    console.error('Error message:', err.message);
    if (err.code === 'LIMIT_FILE_SIZE') {
      return res.status(400).json({ message: 'File quá lớn', error: err.message });
    }
    if (err.message && err.message.includes('Excel')) {
      return res.status(400).json({ message: err.message });
    }
    return res.status(400).json({ 
      message: 'Lỗi khi upload file', 
      error: err.message,
      code: err.code,
      stack: process.env.NODE_ENV === 'development' ? err.stack : undefined
    });
  }
  next();
};

// Route test riêng file leaveFile
router.post(
  "/test-leave",
  logRequest,
  fields([
    { name: "leaveFile", maxCount: 1 }
  ]),
  handleMulterError,
  testLeaveFile
);

// Route test riêng file attendanceFile
router.post(
  "/test-attendance",
  fields([
    { name: "attendanceFile", maxCount: 1 }
  ]),
  handleMulterError,
  testAttendanceFile
);

// Route chính upload cả 2 file và check
router.post(
  "/leave",
  fields([
    { name: "leaveFile", maxCount: 1 },
    { name: "attendanceFile", maxCount: 1 }
  ]),
  handleMulterError,
  uploadTwoFiles
);

// Route download file đã được cập nhật
router.get("/download/:filename", downloadUpdatedFile);

module.exports = router;
