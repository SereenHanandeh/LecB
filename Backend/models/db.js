const { Pool } = require("pg");
require("dotenv").config();

// =========================================================
// Database Connection
// =========================================================

const pool = new Pool({
  connectionString: process.env.DB_URL,

  ssl: {
    rejectUnauthorized: false,
  },
});

// =========================================================
// Test Database Connection
// =========================================================

pool.on("connect", () => {
  console.log("✅ PostgreSQL connected");
});

pool.on("error", (err) => {
  console.error("❌ PostgreSQL pool error:", err.message);
});

// =========================================================
// Export Pool
// =========================================================

module.exports = pool;