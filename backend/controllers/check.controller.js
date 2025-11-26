const path = require("path");
const fs = require("fs");
const { parseExcelFile } = require("../services/excelParser");
const {
  leaveFileConfig,
  attendanceFileConfig,
} = require("../config/excelConfig");
const { checkAttendanceData } = require("../services/checkAttendance");
const { updateLeaveFile } = require("../services/updateLeaveFile");
const { 
  downloadFileFromCloudinary, 
  uploadFileToCloudinary, 
  deleteFileFromCloudinary 
} = require("../utils/cloudinaryHelper");

const uploadTwoFiles = async (req, res) => {
    try {
      // Debug: log để xem cấu trúc req.files
      console.log('req.files:', JSON.stringify(req.files, null, 2));
      console.log('req.body:', JSON.stringify(req.body, null, 2));
      
      const leaveFile = req.files?.leaveFile?.[0];
      const attendanceFile = req.files?.attendanceFile?.[0];
  
      if (!leaveFile || !attendanceFile) {
        return res.status(400).json({
          message: "Thiếu file upload. Cần upload cả 2 file: leaveFile và attendanceFile.",
          debug: {
            hasFiles: !!req.files,
            filesKeys: req.files ? Object.keys(req.files) : [],
            leaveFile: leaveFile ? 'exists' : 'missing',
            attendanceFile: attendanceFile ? 'exists' : 'missing'
          }
        });
      }

      // Lấy config từ request body (nếu có), nếu không thì dùng config mặc định
      let leaveConfigToUse = leaveFileConfig;
      let attendanceConfigToUse = attendanceFileConfig;
      
      // Parse config từ formData (multer sẽ đặt text fields vào req.body)
      if (req.body.leaveFileConfig) {
        try {
          const configStr = typeof req.body.leaveFileConfig === 'string' 
            ? req.body.leaveFileConfig 
            : JSON.stringify(req.body.leaveFileConfig);
          const customLeaveConfig = JSON.parse(configStr);
          // Merge với config mặc định, ưu tiên giá trị từ request
          leaveConfigToUse = { 
            ...leaveFileConfig, 
            ...customLeaveConfig,
            // Đảm bảo các giá trị số được convert đúng
            sheetIndex: customLeaveConfig.sheetIndex !== undefined ? parseInt(customLeaveConfig.sheetIndex) : leaveFileConfig.sheetIndex,
            headerRow: customLeaveConfig.headerRow !== undefined ? parseInt(customLeaveConfig.headerRow) : leaveFileConfig.headerRow,
            startRow: customLeaveConfig.startRow !== undefined ? parseInt(customLeaveConfig.startRow) : leaveFileConfig.startRow
          };
        } catch (err) {
          console.warn('Lỗi parse leaveFileConfig, dùng config mặc định:', err.message);
        }
      }
      
      if (req.body.attendanceFileConfig) {
        try {
          const configStr = typeof req.body.attendanceFileConfig === 'string'
            ? req.body.attendanceFileConfig
            : JSON.stringify(req.body.attendanceFileConfig);
          const customAttendanceConfig = JSON.parse(configStr);
          // Merge với config mặc định, ưu tiên giá trị từ request
          attendanceConfigToUse = { 
            ...attendanceFileConfig, 
            ...customAttendanceConfig,
            // Đảm bảo các giá trị số được convert đúng
            sheetIndex: customAttendanceConfig.sheetIndex !== undefined ? parseInt(customAttendanceConfig.sheetIndex) : attendanceFileConfig.sheetIndex,
            headerRow: customAttendanceConfig.headerRow !== undefined ? parseInt(customAttendanceConfig.headerRow) : attendanceFileConfig.headerRow,
            startRow: customAttendanceConfig.startRow !== undefined ? parseInt(customAttendanceConfig.startRow) : attendanceFileConfig.startRow
          };
        } catch (err) {
          console.warn('Lỗi parse attendanceFileConfig, dùng config mặc định:', err.message);
        }
      }

      // Kiểm tra xem file đến từ Cloudinary hay disk storage
      const isLeaveFileCloudinary = !!(leaveFile.public_id || leaveFile.secure_url || (leaveFile.url && leaveFile.url.startsWith('http')));
      const isAttendanceFileCloudinary = !!(attendanceFile.public_id || attendanceFile.secure_url || (attendanceFile.url && attendanceFile.url.startsWith('http')));
      
      let leaveFileBuffer;
      let attendanceFileBuffer;
      
      // Xử lý leaveFile
      if (isLeaveFileCloudinary) {
        // File từ Cloudinary - download về buffer
        const leaveFileUrl = leaveFile.path || leaveFile.secure_url || leaveFile.url;
        if (!leaveFileUrl) {
          return res.status(400).json({
            message: "Không thể lấy URL file leaveFile từ Cloudinary"
          });
        }
        leaveFileBuffer = await downloadFileFromCloudinary(leaveFileUrl);
      } else {
        // File từ disk storage - đọc trực tiếp từ path
        const filePath = leaveFile.path;
        if (!filePath) {
          return res.status(400).json({
            message: "Không thể lấy đường dẫn file leaveFile"
          });
        }
        leaveFileBuffer = fs.readFileSync(filePath);
      }
      
      // Xử lý attendanceFile
      if (isAttendanceFileCloudinary) {
        // File từ Cloudinary - download về buffer
        const attendanceFileUrl = attendanceFile.path || attendanceFile.secure_url || attendanceFile.url;
        if (!attendanceFileUrl) {
          return res.status(400).json({
            message: "Không thể lấy URL file attendanceFile từ Cloudinary"
          });
        }
        attendanceFileBuffer = await downloadFileFromCloudinary(attendanceFileUrl);
      } else {
        // File từ disk storage - đọc trực tiếp từ path
        const filePath = attendanceFile.path;
        if (!filePath) {
          return res.status(400).json({
            message: "Không thể lấy đường dẫn file attendanceFile"
          });
        }
        attendanceFileBuffer = fs.readFileSync(filePath);
      }
      
      const leaveFilePublicId = leaveFile.public_id;
      const attendanceFilePublicId = attendanceFile.public_id;

      // Parse cả 2 file từ buffer với config đã chọn
      const leaveData = await parseExcelFile(leaveFileBuffer, leaveConfigToUse);
      const attendanceData = await parseExcelFile(attendanceFileBuffer, attendanceConfigToUse);
      
      // Cleanup: Xóa file attendanceFile sau khi parse (không cần nữa)
      if (isAttendanceFileCloudinary && attendanceFilePublicId) {
        // Xóa từ Cloudinary
        try {
          await deleteFileFromCloudinary(attendanceFilePublicId);
        } catch (cleanupError) {
          console.warn('Lỗi khi xóa file attendanceFile từ Cloudinary:', cleanupError.message);
        }
      } else if (attendanceFile.path) {
        // Xóa file local
        try {
          fs.unlinkSync(attendanceFile.path);
        } catch (cleanupError) {
          console.warn('Lỗi khi xóa file attendanceFile local:', cleanupError.message);
        }
      }
      
      // Check dữ liệu giữa 2 file
      const checkResults = checkAttendanceData(leaveData, attendanceData);
      
      // Cập nhật file leaveFile với ghi chú và highlight (từ buffer)
      const updatedFileBuffer = await updateLeaveFile(leaveFileBuffer, checkResults, null, leaveConfigToUse);
      
      // Upload file đã cập nhật lên Cloudinary
      const originalName = leaveFile.originalname || 'leave_file';
      const updatedFileResult = await uploadFileToCloudinary(
        updatedFileBuffer, 
        originalName.replace(/\.[^/.]+$/, '') + '_checked.xlsx',
        'excel-uploads'
      );

      // Cleanup: Xóa file leaveFile gốc sau khi đã upload file updated
      if (isLeaveFileCloudinary && leaveFilePublicId) {
        // Xóa từ Cloudinary
        try {
          await deleteFileFromCloudinary(leaveFilePublicId);
        } catch (cleanupError) {
          console.warn('Lỗi khi xóa file leaveFile gốc từ Cloudinary:', cleanupError.message);
        }
      } else if (leaveFile.path) {
        // Xóa file local
        try {
          fs.unlinkSync(leaveFile.path);
        } catch (cleanupError) {
          console.warn('Lỗi khi xóa file leaveFile gốc local:', cleanupError.message);
        }
      }
      
      // Đếm số dòng có issue và lấy danh sách MNV có issue
      const issuesCount = checkResults.filter(r => r.hasIssue).length;
      const mnvWithIssues = checkResults
        .filter(r => r.hasIssue)
        .map(r => r.mnv)
        .filter(mnv => mnv); // Loại bỏ null/undefined
  
      return res.json({
        message: "Check và cập nhật file thành công!",
        files: {
          leaveFile: leaveFile.originalname,
          attendanceFile: attendanceFile.originalname,
          updatedLeaveFile: updatedFileResult.public_id,
        },
        summary: {
          totalLeaveRows: leaveData.length,
          totalAttendanceRows: attendanceData.length,
          issuesFound: issuesCount,
          mnvWithIssues: mnvWithIssues,
          checkResults: checkResults
        },
        downloadUrl: updatedFileResult.secure_url || updatedFileResult.url,
      });
  
    } catch (error) {
      // Cleanup: Xóa file input nếu có lỗi
      try {
        if (leaveFile?.public_id) {
          await deleteFileFromCloudinary(leaveFile.public_id);
        } else if (leaveFile?.path) {
          fs.unlinkSync(leaveFile.path);
        }
        if (attendanceFile?.public_id) {
          await deleteFileFromCloudinary(attendanceFile.public_id);
        } else if (attendanceFile?.path) {
          fs.unlinkSync(attendanceFile.path);
        }
      } catch (cleanupError) {
        console.warn('Lỗi khi xóa file sau error:', cleanupError.message);
      }
      
      console.error("Lỗi khi xử lý file:", error);
      res.status(500).json({ 
        message: "Lỗi server khi upload hoặc parse file",
        error: error.message 
      });
    }
  };

// Controller để test riêng file leaveFile
const testLeaveFile = async (req, res) => {
  try {
    // Debug: log để xem cấu trúc req.files
    console.log('testLeaveFile - req.files:', JSON.stringify(req.files, null, 2));
    
    const leaveFile = req.files?.leaveFile?.[0];

    if (!leaveFile) {
      return res.status(400).json({
        message: "Thiếu file upload. Cần upload file leaveFile.",
        debug: {
          hasFiles: !!req.files,
          filesKeys: req.files ? Object.keys(req.files) : [],
          leaveFile: leaveFile ? 'exists' : 'missing'
        }
      });
    }

    // Kiểm tra xem file đến từ Cloudinary hay disk storage
    const isCloudinaryFile = !!(leaveFile.public_id || leaveFile.secure_url || (leaveFile.url && leaveFile.url.startsWith('http')));
    
    let leaveFileBuffer;
    
    if (isCloudinaryFile) {
      // File từ Cloudinary - download về buffer
      const leaveFileUrl = leaveFile.path || leaveFile.secure_url || leaveFile.url;
      const leaveFilePublicId = leaveFile.public_id;

      if (!leaveFileUrl) {
        return res.status(400).json({
          message: "Không thể lấy URL file từ Cloudinary"
        });
      }

      leaveFileBuffer = await downloadFileFromCloudinary(leaveFileUrl);
    } else {
      // File từ disk storage - đọc trực tiếp từ path
      const filePath = leaveFile.path;
      if (!filePath) {
        return res.status(400).json({
          message: "Không thể lấy đường dẫn file"
        });
      }
      leaveFileBuffer = fs.readFileSync(filePath);
    }

    // Parse file nghỉ phép từ buffer
    const leaveData = await parseExcelFile(leaveFileBuffer, leaveFileConfig);

    // Cleanup: Xóa file sau khi parse xong
    if (isCloudinaryFile && leaveFile.public_id) {
      // Xóa từ Cloudinary
      try {
        await deleteFileFromCloudinary(leaveFile.public_id);
      } catch (cleanupError) {
        console.warn('Lỗi khi xóa file test từ Cloudinary:', cleanupError.message);
      }
    } else if (leaveFile.path) {
      // Xóa file local
      try {
        fs.unlinkSync(leaveFile.path);
      } catch (cleanupError) {
        console.warn('Lỗi khi xóa file test local:', cleanupError.message);
      }
    }

    return res.json({
      message: "Upload và parse thành công!",
      file: {
        originalname: leaveFile.originalname,
        size: leaveFile.size
      },
      leaveData: {
        totalRows: leaveData.length,
        config: leaveFileConfig,
        data: leaveData // Dữ liệu đã parse từ file nghỉ phép
      }
    });

  } catch (error) {
    // Cleanup: Xóa file nếu có lỗi
    try {
      if (leaveFile?.public_id) {
        await deleteFileFromCloudinary(leaveFile.public_id);
      } else if (leaveFile?.path) {
        fs.unlinkSync(leaveFile.path);
      }
    } catch (cleanupError) {
      console.warn('Lỗi khi xóa file sau error:', cleanupError.message);
    }
    
    console.error("Lỗi khi xử lý file:", error);
    res.status(500).json({ 
      message: "Lỗi server khi upload hoặc parse file",
      error: error.message 
    });
  }
};

// Controller để test riêng file attendanceFile
const testAttendanceFile = async (req, res) => {
  try {
    const attendanceFile = req.files?.attendanceFile?.[0];

    if (!attendanceFile) {
      return res.status(400).json({
        message: "Thiếu file upload. Cần upload file attendanceFile."
      });
    }

    // Kiểm tra xem file đến từ Cloudinary hay disk storage
    const isCloudinaryFile = !!(attendanceFile.public_id || attendanceFile.secure_url || (attendanceFile.url && attendanceFile.url.startsWith('http')));
    
    let attendanceFileBuffer;
    
    if (isCloudinaryFile) {
      // File từ Cloudinary - download về buffer
      const attendanceFileUrl = attendanceFile.path || attendanceFile.secure_url || attendanceFile.url;
      const attendanceFilePublicId = attendanceFile.public_id;

      if (!attendanceFileUrl) {
        return res.status(400).json({
          message: "Không thể lấy URL file từ Cloudinary"
        });
      }

      attendanceFileBuffer = await downloadFileFromCloudinary(attendanceFileUrl);
    } else {
      // File từ disk storage - đọc trực tiếp từ path
      const filePath = attendanceFile.path;
      if (!filePath) {
        return res.status(400).json({
          message: "Không thể lấy đường dẫn file"
        });
      }
      attendanceFileBuffer = fs.readFileSync(filePath);
    }

    // Parse file chấm công từ buffer
    const attendanceData = await parseExcelFile(attendanceFileBuffer, attendanceFileConfig);

    // Cleanup: Xóa file sau khi parse xong
    if (isCloudinaryFile && attendanceFile.public_id) {
      // Xóa từ Cloudinary
      try {
        await deleteFileFromCloudinary(attendanceFile.public_id);
      } catch (cleanupError) {
        console.warn('Lỗi khi xóa file test từ Cloudinary:', cleanupError.message);
      }
    } else if (attendanceFile.path) {
      // Xóa file local
      try {
        fs.unlinkSync(attendanceFile.path);
      } catch (cleanupError) {
        console.warn('Lỗi khi xóa file test local:', cleanupError.message);
      }
    }

    return res.json({
      message: "Upload và parse thành công!",
      file: {
        originalname: attendanceFile.originalname,
        size: attendanceFile.size
      },
      attendanceData: {
        totalRows: attendanceData.length,
        config: attendanceFileConfig,
        data: attendanceData // Dữ liệu đã parse từ file chấm công
      }
    });

  } catch (error) {
    // Cleanup: Xóa file nếu có lỗi
    try {
      if (attendanceFile?.public_id) {
        await deleteFileFromCloudinary(attendanceFile.public_id);
      } else if (attendanceFile?.path) {
        fs.unlinkSync(attendanceFile.path);
      }
    } catch (cleanupError) {
      console.warn('Lỗi khi xóa file sau error:', cleanupError.message);
    }
    
    console.error("Lỗi khi xử lý file:", error);
    res.status(500).json({ 
      message: "Lỗi server khi upload hoặc parse file",
      error: error.message 
    });
  }
};

// Controller để download file đã được cập nhật (từ Cloudinary URL)
// Lưu ý: Với Cloudinary, file đã được trả về URL trực tiếp trong response
// Endpoint này có thể redirect hoặc download từ Cloudinary URL
const downloadUpdatedFile = async (req, res) => {
  try {
    // Nếu nhận được URL từ query param, redirect đến URL đó
    const fileUrl = req.query.url;
    
    if (fileUrl) {
      // Redirect đến Cloudinary URL để download
      return res.redirect(fileUrl);
    }
    
    // Nếu nhận được public_id, download từ Cloudinary
    const publicId = req.params.filename || req.query.public_id;
    
    if (publicId) {
      // Tạo signed URL từ Cloudinary để download
      const cloudinary = require("../config/cloudinary");
      const url = cloudinary.utils.download_url(publicId, {
        resource_type: 'raw',
        attachment: true
      });
      return res.redirect(url);
    }
    
    // Nếu không có thông tin, trả về lỗi
    return res.status(400).json({
      message: "Thiếu thông tin file. Cần cung cấp URL hoặc public_id."
    });
    
  } catch (error) {
    console.error("Lỗi khi xử lý download:", error);
    if (!res.headersSent) {
      res.status(500).json({ 
        message: "Lỗi server khi download file",
        error: error.message 
      });
    }
  }
};

module.exports = { uploadTwoFiles, testLeaveFile, testAttendanceFile, downloadUpdatedFile };