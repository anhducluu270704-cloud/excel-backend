const { CloudinaryStorage } = require("multer-storage-cloudinary");
const cloudinary = require("./cloudinary");

const excelStorage = new CloudinaryStorage({
  cloudinary,
  params: (req, file) => {
    console.log('CloudinaryStorage - Uploading file:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      fieldname: file.fieldname
    });
    
    // Xử lý format an toàn hơn
    let format = 'xlsx'; // default
    if (file.mimetype && file.mimetype.includes('/')) {
      format = file.mimetype.split("/")[1];
    } else if (file.originalname) {
      const ext = file.originalname.split('.').pop();
      if (ext === 'xls' || ext === 'xlsx') {
        format = ext;
      }
    }
    
    const publicId = `${file.originalname.replace(/\.[^/.]+$/, '')}-${Date.now()}`;
    
    console.log('CloudinaryStorage - Params:', {
      folder: "excel-uploads",
      resource_type: "raw",
      public_id: publicId,
      format: format
    });
    
    return {
      folder: "excel-uploads",
      resource_type: "raw", // bắt buộc để upload file excel, pdf, docx
      public_id: publicId,
      format: format,
    };
  },
});

module.exports = excelStorage;
