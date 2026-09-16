const express = require("express");
const multer = require("multer");

const {
  uploadExcel,
  uploadExcelFromUrl,
  generateExcel,
  getRawSessions,
} = require("../controllers/excel.js");

const excelRouter = express.Router();

const storage = multer.memoryStorage();

const upload = multer({
  storage,
});

// Upload Excel file
excelRouter.post(
  "/upload/excel",
  upload.single("file"),
  uploadExcel
);

// Upload Excel from URL
excelRouter.post(
  "/upload/excel-url",
  uploadExcelFromUrl
);

// Generate Excel
excelRouter.get(
  "/generate-excel",
  generateExcel
);

// Get raw sessions
excelRouter.get(
  "/sessions/:batchId",
  getRawSessions
);

module.exports = excelRouter;
