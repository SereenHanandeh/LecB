const express = require("express");

const pool = require("../models/db");

const router = express.Router();

// =====================================================
// GET /supervisor-initial-stats
// كل المشرفين الذين لديهم رصيد ابتدائي ("خطة 0")
// =====================================================

router.get("/", async (req, res) => {
  try {
    const result = await pool.query(
      `
      SELECT
        s.id   AS supervisor_id,
        s.name AS supervisor_name,
        i.initial_count,
        i.note
      FROM supervisor_initial_stats i
      JOIN supervisors s
        ON s.id = i.supervisor_id
      ORDER BY s.id
      `,
    );

    return res.json({
      success: true,
      data: result.rows,
    });
  } catch (error) {
    console.error("❌ Error loading initial supervisor stats:", error);

    return res.status(500).json({
      success: false,
      error: error.message || "Failed to load initial stats.",
    });
  }
});

// =====================================================
// PUT /supervisor-initial-stats/:supervisorId
// Body: { initialCount: number }
// =====================================================

router.put("/:supervisorId", async (req, res) => {
  try {
    const supervisorId = Number(req.params.supervisorId);

    const initialCount = Number(
      req.body?.initialCount ?? req.body?.initial_count,
    );

    if (!Number.isInteger(supervisorId)) {
      return res.status(400).json({
        success: false,
        error: "Invalid supervisor id.",
      });
    }

    if (!Number.isInteger(initialCount) || initialCount < 0) {
      return res.status(400).json({
        success: false,
        error: "initialCount must be an integer >= 0.",
      });
    }

    const supervisorCheck = await pool.query(
      `SELECT id, name FROM supervisors WHERE id = $1`,
      [supervisorId],
    );

    if (supervisorCheck.rowCount === 0) {
      return res.status(404).json({
        success: false,
        error: `Supervisor ${supervisorId} not found.`,
      });
    }

    const result = await pool.query(
      `
      INSERT INTO supervisor_initial_stats
        (supervisor_id, initial_count, updated_at)
      VALUES ($1, $2, NOW())
      ON CONFLICT (supervisor_id)
      DO UPDATE SET
        initial_count = EXCLUDED.initial_count,
        updated_at    = NOW()
      RETURNING supervisor_id, initial_count
      `,
      [supervisorId, initialCount],
    );

    return res.json({
      success: true,
      data: {
        ...result.rows[0],
        supervisor_name: supervisorCheck.rows[0].name,
      },
    });
  } catch (error) {
    console.error("❌ Error saving initial supervisor stat:", error);

    return res.status(500).json({
      success: false,
      error: error.message || "Failed to save initial stat.",
    });
  }
});

module.exports = router;