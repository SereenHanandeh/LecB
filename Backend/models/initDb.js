const fs = require("fs");
const path = require("path");
const pool = require("./db");

async function initializeDatabase() {
  try {
    const sqlPath = path.join(__dirname, "Database.sql");

    if (!fs.existsSync(sqlPath)) {
      throw new Error(`Database.sql not found at: ${sqlPath}`);
    }

    const sql = fs.readFileSync(sqlPath, "utf8");

    await pool.query(sql);

    console.log("✅ Database tables initialized successfully");
  } catch (error) {
    console.error("❌ Database initialization failed:");
    console.error(error.message);
    throw error;
  }
}

module.exports = initializeDatabase;