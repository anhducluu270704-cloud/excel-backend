const cloudinary = require("../config/cloudinary");
const https = require("https");
const http = require("http");

/**
 * Download file từ Cloudinary về buffer
 * @param {string} url - URL của file trên Cloudinary
 * @returns {Promise<Buffer>} - Buffer chứa nội dung file
 */
const downloadFileFromCloudinary = (url) => {
  return new Promise((resolve, reject) => {
    const protocol = url.startsWith("https") ? https : http;
    
    protocol.get(url, (response) => {
      if (response.statusCode !== 200) {
        reject(new Error(`Failed to download file: ${response.statusCode}`));
        return;
      }

      const chunks = [];
      response.on("data", (chunk) => {
        chunks.push(chunk);
      });

      response.on("end", () => {
        resolve(Buffer.concat(chunks));
      });

      response.on("error", (err) => {
        reject(err);
      });
    }).on("error", (err) => {
      reject(err);
    });
  });
};

/**
 * Upload file lên Cloudinary
 * @param {Buffer} fileBuffer - Buffer chứa nội dung file
 * @param {string} originalName - Tên file gốc
 * @param {string} folder - Folder trên Cloudinary
 * @returns {Promise<Object>} - Thông tin file đã upload (url, public_id, etc.)
 */
const uploadFileToCloudinary = async (fileBuffer, originalName, folder = "excel-uploads") => {
  return new Promise((resolve, reject) => {
    const uploadStream = cloudinary.uploader.upload_stream(
      {
        folder: folder,
        resource_type: "raw",
        public_id: originalName.replace(/\.[^/.]+$/, "") + "-" + Date.now(),
      },
      (error, result) => {
        if (error) {
          reject(error);
        } else {
          resolve(result);
        }
      }
    );

    uploadStream.end(fileBuffer);
  });
};

/**
 * Xóa file từ Cloudinary
 * @param {string} publicId - Public ID của file trên Cloudinary
 * @returns {Promise<Object>} - Kết quả xóa
 */
const deleteFileFromCloudinary = async (publicId) => {
  try {
    const result = await cloudinary.uploader.destroy(publicId, {
      resource_type: "raw",
    });
    return result;
  } catch (error) {
    console.warn("Lỗi khi xóa file từ Cloudinary:", error.message);
    return null;
  }
};

module.exports = {
  downloadFileFromCloudinary,
  uploadFileToCloudinary,
  deleteFileFromCloudinary,
};

