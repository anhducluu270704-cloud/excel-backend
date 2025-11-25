require('dotenv').config();
const express = require("express");
const cors = require("cors");

const checkRoutes = require("./routes/check.routes");

const app = express();

// CORS config - cho phép tất cả headers
app.use(cors({
  origin: '*',
  credentials: true,
  exposedHeaders: ['Content-Type'],
  allowedHeaders: ['Content-Type', 'Authorization', 'X-Requested-With']
}));

// Middleware để log tất cả headers (debug)
app.use((req, res, next) => {
  if (req.path.startsWith('/api/check')) {
    console.log('=== All Headers ===');
    console.log('Content-Type:', req.headers['content-type']);
    console.log('Content-Length:', req.headers['content-length']);
    console.log('All headers:', JSON.stringify(req.headers, null, 2));
  }
  next();
});

// QUAN TRỌNG: KHÔNG parse body ở đây cho các route có file upload
// Chỉ parse body cho các route KHÔNG có file upload
// Multer cần xử lý multipart/form-data trước, không được parse body trước đó

// API routes (multer sẽ xử lý body ở đây)
app.use("/api/check", checkRoutes);

// Parse body cho các route khác (sau khi multer đã xử lý)
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Listen
const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`Server đang chạy tại http://localhost:${PORT}`);
});
