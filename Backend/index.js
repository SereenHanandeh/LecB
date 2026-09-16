require("dotenv").config();

const express = require("express");
const cors = require("cors");
const fs = require("fs");
const path = require("path");
const http = require("http");

const pool = require("./models/db");
const initializeDatabase = require("./models/initDb");

const excelRoutes = require("./routes/excel.js");
const planRouter = require("./routes/plan.js");
const supervisorRouter = require("./routes/supervisor.js");

const app = express();

// =====================================================
// Basic Config
// =====================================================

const PORT = process.env.PORT || 5000;

const server = http.createServer(app);

// =====================================================
// Middleware
// =====================================================

app.use(cors());

app.use(
  express.json({
    limit: "10mb",
  })
);

// =====================================================
// Uploads Directory
// =====================================================

const UPLOAD_DIR =
  process.env.UPLOAD_DIR ||
  path.join(__dirname, "uploads");

if (!fs.existsSync(UPLOAD_DIR)) {
  fs.mkdirSync(UPLOAD_DIR, {
    recursive: true,
  });
}

app.use(
  "/uploads",
  express.static(UPLOAD_DIR)
);

// =====================================================
// Exports Directory
// =====================================================

const EXPORTS_DIR =
  path.join(__dirname, "exports");

if (!fs.existsSync(EXPORTS_DIR)) {
  fs.mkdirSync(EXPORTS_DIR, {
    recursive: true,
  });
}

app.use(
  "/exports",
  express.static(EXPORTS_DIR)
);

// =====================================================
// Routes
// =====================================================

app.use(
  "/excel",
  excelRoutes
);

app.use(
  "/plan",
  planRouter
);

app.use(
  "/supervisors",
  supervisorRouter
);

// =====================================================
// Health Check
// =====================================================

app.get("/", (req, res) => {
  res.json({
    success: true,
    message:
      "Supervisor Distribution System API is running",
  });
});

// =====================================================
// Database Test
// =====================================================

async function testDatabase() {
  try {
    const result = await pool.query(
      "SELECT NOW() AS now"
    );

    console.log(
      "✅ PostgreSQL connected:",
      result.rows[0].now
    );

    return true;
  } catch (error) {
    console.error(
      "❌ PostgreSQL connection failed:",
      error.message
    );

    return false;
  }
}

// =====================================================
// Start Server
// =====================================================

async function startServer() {
  try {
    console.log("🔄 Checking database connection...");

    const dbConnected = await testDatabase();

    if (!dbConnected) {
      console.error(
        "❌ Server stopped because database connection failed."
      );

      process.exit(1);
    }

    // =================================================
    // Initialize Database Tables
    // =================================================

    console.log("🔄 Initializing database...");

    await initializeDatabase();

    console.log(
      "✅ Database initialization completed"
    );

    // =================================================
    // Start Server
    // =================================================

    server.listen(PORT, () => {
      console.log(
        `🚀 Server Listening At PORT ${PORT}`
      );
    });

  } catch (error) {
    console.error(
      "❌ Failed to start server:"
    );

    console.error(error);

    process.exit(1);
  }
}

startServer();
