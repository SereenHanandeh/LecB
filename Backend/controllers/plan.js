const {
  createPlanRow,
  saveDutyPool,
  savePreassignments,
  saveAffinities,
  fetchPlan,
  lockRow,
  unlockRow,
  moveAssignmentSvc,
  planStats,
  getAllPlans,
} = require("../services/plan.js");

const { generatePlan } = require("../services/excel.js");

// =====================================================
// Error Handler
// =====================================================

function handleError(res, err, message = "Internal server error") {
  console.error(`❌ ${message}:`, err);

  return res.status(500).json({
    success: false,
    error: err.message || message,
  });
}

// Create Plan
async function createPlan(req, res) {
  try {
    const {
      name,
      excelBatchId,
      dateFrom,
      dateTo,
      category,
      planCategory,
      plan_category,
    } = req.body;

    // دعم أكثر من اسم قادم من Frontend
    const selectedCategory =
      category ?? planCategory ?? plan_category;

    // ==============================
    // Required fields
    // ==============================

    if (!name || !excelBatchId || !dateFrom || !dateTo) {
      return res.status(400).json({
        success: false,
        error:
          "name, excelBatchId, dateFrom and dateTo are required",
      });
    }

    // ==============================
    // Validate category
    // ==============================

    const allowedCategories = [
      "مدمج",
      "دبلوم",
      "متطلبات",
    ];

    if (!allowedCategories.includes(selectedCategory)) {
      return res.status(400).json({
        success: false,
        error:
          "يجب اختيار فئة صحيحة للخطة: مدمج، دبلوم، أو متطلبات",
      });
    }

    // ==============================
    // Validate dates
    // ==============================

    const fromDate = new Date(dateFrom);
    const toDate = new Date(dateTo);

    if (
      Number.isNaN(fromDate.getTime()) ||
      Number.isNaN(toDate.getTime())
    ) {
      return res.status(400).json({
        success: false,
        error: "Invalid dateFrom or dateTo",
      });
    }

    if (fromDate > toDate) {
      return res.status(400).json({
        success: false,
        error: "dateFrom cannot be after dateTo",
      });
    }

    // ==============================
    // Create plan
    // ==============================

    const plan = await createPlanRow({
      name,
      excelBatchId,
      dateFrom,
      dateTo,
      category: selectedCategory,
    });

    console.log("========================================");
    console.log("✅ CREATED PLAN");
    console.log("🆔 Plan:", plan);
    console.log("📂 Category:", selectedCategory);
    console.log("========================================");

    return res.status(201).json({
      success: true,
      message: "Plan created successfully",
      data: plan,
    });
  } catch (err) {
    return handleError(
      res,
      err,
      "Error creating plan",
    );
  }
}

// =====================================================
// Set Duty Pool
// =====================================================

async function setDutyPool(req, res) {
  try {
    const { planId } = req.params;
    const { supervisorIds } = req.body;

    console.log("========================================");
    console.log("🎯 SET DUTY POOL");
    console.log("📌 planId:", planId);
    console.log("📌 supervisorIds received:", supervisorIds);
    console.log(
      "📌 supervisorIds type:",
      Array.isArray(supervisorIds) ? "array" : typeof supervisorIds,
    );
    console.log("========================================");

    // planId هو UUID
    if (!planId) {
      return res.status(400).json({
        success: false,
        error: "planId is required",
      });
    }

    if (!Array.isArray(supervisorIds)) {
      return res.status(400).json({
        success: false,
        error: "supervisorIds must be an array",
      });
    }

    // ندعم أكثر من شكل قادم من الـ Frontend:
    // 1) [1, 2, 3]
    // 2) ["1", "2", "3"]
    // 3) [{ id: 1 }, { id: 2 }]
    // 4) [{ supervisor_id: 1 }, { supervisor_id: 2 }]
    const uniqueIds = Array.from(
      new Set(
        supervisorIds
          .map((item) => {
            // إذا كان ID مباشر
            if (typeof item === "number" || typeof item === "string") {
              return Number(item);
            }

            // إذا كان Object
            if (item && typeof item === "object") {
              return Number(item.id ?? item.supervisorId ?? item.supervisor_id);
            }

            return NaN;
          })
          .filter((id) => Number.isInteger(id)),
      ),
    );

    console.log("📌 normalized supervisor IDs:", uniqueIds);

    if (!uniqueIds.length) {
      return res.status(400).json({
        success: false,
        error: "At least one valid supervisor is required",
      });
    }

    await saveDutyPool(planId, uniqueIds);

    console.log("✅ Duty pool saved successfully");

    return res.json({
      success: true,
      message: "Duty pool saved successfully",
      data: {
        supervisorIds: uniqueIds,
      },
    });
  } catch (err) {
    return handleError(res, err, "Error setting duty pool");
  }
}

// =====================================================
// Add Preassignments
// =====================================================

async function addPreassignments(req, res) {
  try {
    const { planId } = req.params;
    const items = req.body;

    // planId هو UUID
    if (!planId) {
      return res.status(400).json({
        success: false,
        error: "planId is required",
      });
    }

    if (!Array.isArray(items)) {
      return res.status(400).json({
        success: false,
        error: "Body must be an array",
      });
    }

    await savePreassignments(planId, items);

    return res.json({
      success: true,
      message: "Preassignments saved successfully",
      data: null,
    });
  } catch (err) {
    return handleError(res, err, "Error saving preassignments");
  }
}

// =====================================================
// Add Affinities
// =====================================================

async function addAffinities(req, res) {
  try {
    const { planId } = req.params;
    const items = req.body;

    // planId هو UUID
    if (!planId) {
      return res.status(400).json({
        success: false,
        error: "planId is required",
      });
    }

    if (!Array.isArray(items)) {
      return res.status(400).json({
        success: false,
        error: "Body must be an array",
      });
    }

    const saved = await saveAffinities(planId, items);

    return res.json({
      success: true,
      message: "Affinities saved successfully",
      data: saved,
    });
  } catch (err) {
    return handleError(res, err, "Error saving affinities");
  }
}

// =====================================================
// Generate Plan
// =====================================================

async function generate(req, res) {
  try {
    const { planId } = req.params;

    const {
      variant = 1,
      minimumPeriodsEnabled = false,
      minimumPeriods = 4,
    } = req.body || {};

    if (!planId) {
      return res.status(400).json({
        success: false,
        error: "planId is required",
      });
    }

    const parsedVariant = Number(variant);

    if (
      !Number.isInteger(parsedVariant) ||
      parsedVariant < 1
    ) {
      return res.status(400).json({
        success: false,
        error:
          "variant must be a positive integer",
      });
    }

    const parsedMinimumEnabled =
      Boolean(minimumPeriodsEnabled);

    const parsedMinimumPeriods =
      Number(minimumPeriods);

    if (
      !Number.isInteger(
        parsedMinimumPeriods,
      ) ||
      parsedMinimumPeriods < 1
    ) {
      return res.status(400).json({
        success: false,
        error:
          "minimumPeriods must be a positive integer",
      });
    }

    console.log(
      "🚀 GENERATING PLAN",
      {
        planId,
        variant: parsedVariant,
        minimumPeriodsEnabled:
          parsedMinimumEnabled,
        minimumPeriods:
          parsedMinimumPeriods,
      },
    );

    const generated =
      await generatePlan(
        planId,
        parsedVariant,
        parsedMinimumEnabled,
        parsedMinimumPeriods,
      );

    console.log(
      "✅ Generated plan:",
      generated,
    );

    return res.json({
      success: true,

      message:
        "Plan generated successfully",

      data: generated,

      stats:
        generated.statistics || null,

      downloadUrl:
        `/exports/plan_${planId}.xlsx`,
    });

  } catch (err) {
    return handleError(
      res,
      err,
      "Error generating plan",
    );
  }
}

// =====================================================
// Get One Plan
// =====================================================

async function getPlan(req, res) {
  try {
    const { planId } = req.params;

    // planId هو UUID
    if (!planId) {
      return res.status(400).json({
        success: false,
        error: "planId is required",
      });
    }

    console.log("🔎 Fetching plan:", planId);

    const plan = await fetchPlan(planId);

    if (!plan) {
      return res.status(404).json({
        success: false,
        error: "Plan not found",
      });
    }

    return res.json({
      success: true,
      message: "Plan fetched successfully",
      data: plan,
    });
  } catch (err) {
    return handleError(res, err, "Error fetching plan");
  }
}

// =====================================================
// Get All Plans
// =====================================================

async function getPlans(req, res) {
  try {
    const plans = await getAllPlans();

    return res.json({
      success: true,
      message: "Plans fetched successfully",
      data: plans,
    });
  } catch (err) {
    return handleError(res, err, "Error fetching plans");
  }
}

// =====================================================
// Lock Assignment
// =====================================================

async function lockAssignment(req, res) {
  try {
    const { planId } = req.params;

    const { sessionGroupId, supervisorId } = req.body;

    // planId هو UUID
    if (!planId || sessionGroupId == null || supervisorId == null) {
      return res.status(400).json({
        success: false,
        error: "planId, sessionGroupId and supervisorId are required",
      });
    }

    const parsedSessionGroupId = Number(sessionGroupId);

    const parsedSupervisorId = Number(supervisorId);

    if (
      !Number.isInteger(parsedSessionGroupId) ||
      !Number.isInteger(parsedSupervisorId)
    ) {
      return res.status(400).json({
        success: false,
        error: "sessionGroupId and supervisorId must be valid integers",
      });
    }

    await lockRow(planId, parsedSessionGroupId, parsedSupervisorId);

    return res.json({
      success: true,
      message: "Assignment locked successfully",
      data: null,
    });
  } catch (err) {
    return handleError(res, err, "Error locking assignment");
  }
}

// =====================================================
// Unlock Assignment
// =====================================================

async function unlockAssignment(req, res) {
  try {
    const { planId, sessionGroupId } = req.params;

    // planId هو UUID
    if (!planId || !sessionGroupId) {
      return res.status(400).json({
        success: false,
        error: "planId and sessionGroupId are required",
      });
    }

    const parsedSessionGroupId = Number(sessionGroupId);

    if (!Number.isInteger(parsedSessionGroupId)) {
      return res.status(400).json({
        success: false,
        error: "sessionGroupId must be a valid integer",
      });
    }

    await unlockRow(planId, parsedSessionGroupId);

    return res.json({
      success: true,
      message: "Assignment unlocked successfully",
      data: null,
    });
  } catch (err) {
    return handleError(res, err, "Error unlocking assignment");
  }
}

// =====================================================
// Move Assignment
// =====================================================

async function moveAssignment(req, res) {
  try {
    const { planId } = req.params;

    const { fromSupervisorId, toSupervisorId, sessionGroupId } = req.body;

    // planId هو UUID
    if (
      !planId ||
      fromSupervisorId == null ||
      toSupervisorId == null ||
      sessionGroupId == null
    ) {
      return res.status(400).json({
        success: false,
        error:
          "planId, fromSupervisorId, toSupervisorId and sessionGroupId are required",
      });
    }

    const parsedFromSupervisorId = Number(fromSupervisorId);

    const parsedToSupervisorId = Number(toSupervisorId);

    const parsedSessionGroupId = Number(sessionGroupId);

    if (
      !Number.isInteger(parsedFromSupervisorId) ||
      !Number.isInteger(parsedToSupervisorId) ||
      !Number.isInteger(parsedSessionGroupId)
    ) {
      return res.status(400).json({
        success: false,
        error: "Supervisor IDs and sessionGroupId must be valid integers",
      });
    }

    await moveAssignmentSvc(planId, {
      fromSupervisorId: parsedFromSupervisorId,

      toSupervisorId: parsedToSupervisorId,

      sessionGroupId: parsedSessionGroupId,
    });

    return res.json({
      success: true,
      message: "Assignment moved successfully",
      data: null,
    });
  } catch (err) {
    return handleError(res, err, "Error moving assignment");
  }
}

// =====================================================
// Get Stats
// =====================================================

async function getStats(req, res) {
  try {
    const { planId } = req.params;

    // planId هو UUID
    if (!planId) {
      return res.status(400).json({
        success: false,
        error: "planId is required",
      });
    }

    const stats = await planStats(planId);

    return res.json({
      success: true,
      message: "Plan stats fetched successfully",
      data: stats,
    });
  } catch (err) {
    return handleError(res, err, "Error fetching stats");
  }
}

// =====================================================
// Exports
// =====================================================

module.exports = {
  createPlan,
  setDutyPool,
  addPreassignments,
  addAffinities,
  generate,
  getPlan,
  getPlans,
  lockAssignment,
  unlockAssignment,
  moveAssignment,
  getStats,
};
