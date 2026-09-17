const {
  createPlanRow,
  saveDutyPool,
  savePeriodQuotas,
  savePreassignments,
  saveAffinities,
  fetchPlan,
  lockRow,
  unlockRow,
  moveAssignmentSvc,
  planStats,
  getAllPlans,
} = require("../services/plan.js");

const {
  generatePlan,
} = require("../services/excel.js");

// =====================================================
// Error Handler
// =====================================================

function handleError(
  res,
  err,
  message = "Internal server error"
) {
  console.error(`❌ ${message}:`, err);

  return res.status(500).json({
    success: false,
    error: err.message || message,
  });
}

// =====================================================
// Create Plan
// =====================================================

async function createPlan(req, res) {
  try {
    const {
      name,
      excelBatchId,
      dateFrom,
      dateTo,
    } = req.body;

    if (
      !name ||
      !excelBatchId ||
      !dateFrom ||
      !dateTo
    ) {
      return res.status(400).json({
        success: false,
        error:
          "name, excelBatchId, dateFrom and dateTo are required",
      });
    }

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
        error:
          "dateFrom cannot be after dateTo",
      });
    }

    const plan = await createPlanRow({
      name,
      excelBatchId,
      dateFrom,
      dateTo,
    });

    console.log("✅ Created plan:", plan);

    return res.status(201).json({
      success: true,
      message: "Plan created successfully",
      data: plan,
    });
  } catch (err) {
    return handleError(
      res,
      err,
      "Error creating plan"
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

    if (!planId) {
      return res.status(400).json({
        success: false,
        error: "planId is required",
      });
    }

    const parsedPlanId = Number(planId);

    if (!Number.isInteger(parsedPlanId)) {
      return res.status(400).json({
        success: false,
        error: "planId must be a valid integer",
      });
    }

    if (!Array.isArray(supervisorIds)) {
      return res.status(400).json({
        success: false,
        error:
          "supervisorIds must be an array",
      });
    }

    const uniqueIds = Array.from(
      new Set(
        supervisorIds
          .map(Number)
          .filter((id) =>
            Number.isInteger(id)
          )
      )
    );

    if (!uniqueIds.length) {
      return res.status(400).json({
        success: false,
        error:
          "At least one valid supervisor is required",
      });
    }

    await saveDutyPool(
      parsedPlanId,
      uniqueIds
    );

    return res.json({
      success: true,
      message:
        "Duty pool saved successfully",
      data: {
        supervisorIds: uniqueIds,
      },
    });
  } catch (err) {
    return handleError(
      res,
      err,
      "Error setting duty pool"
    );
  }
}

// =====================================================
// Set Period Quotas
// =====================================================

async function setPeriodQuotas(req, res) {
  try {
    const { planId } = req.params;
    const { supervisors } = req.body;

    if (!planId) {
      return res.status(400).json({
        success: false,
        error: "planId is required",
      });
    }

    const parsedPlanId = Number(planId);

    if (!Number.isInteger(parsedPlanId)) {
      return res.status(400).json({
        success: false,
        error: "planId must be a valid integer",
      });
    }

    if (
      !supervisors ||
      typeof supervisors !== "object" ||
      Array.isArray(supervisors)
    ) {
      return res.status(400).json({
        success: false,
        error:
          "supervisors must be an object",
      });
    }

    const normalized = {};

    for (
      const [supervisorId, targetPeriods] of
      Object.entries(supervisors)
    ) {
      const sid = Number(supervisorId);
      const target = Number(targetPeriods);

      if (!Number.isInteger(sid)) {
        continue;
      }

      if (
        !Number.isInteger(target) ||
        target <= 0
      ) {
        return res.status(400).json({
          success: false,
          error:
            `Invalid period quota for supervisor ${sid}`,
        });
      }

      normalized[sid] = target;
    }

    await savePeriodQuotas(
      parsedPlanId,
      normalized
    );

    return res.json({
      success: true,
      message:
        "Period quotas saved successfully",
      data: {
        supervisors: normalized,
      },
    });
  } catch (err) {
    return handleError(
      res,
      err,
      "Error saving period quotas"
    );
  }
}

// =====================================================
// Add Preassignments
// =====================================================

async function addPreassignments(req, res) {
  try {
    const { planId } = req.params;
    const items = req.body;

    if (!planId) {
      return res.status(400).json({
        success: false,
        error: "planId is required",
      });
    }

    const parsedPlanId = Number(planId);

    if (!Number.isInteger(parsedPlanId)) {
      return res.status(400).json({
        success: false,
        error: "planId must be a valid integer",
      });
    }

    if (!Array.isArray(items)) {
      return res.status(400).json({
        success: false,
        error: "Body must be an array",
      });
    }

    await savePreassignments(
      parsedPlanId,
      items
    );

    return res.json({
      success: true,
      message:
        "Preassignments saved successfully",
      data: null,
    });
  } catch (err) {
    return handleError(
      res,
      err,
      "Error saving preassignments"
    );
  }
}

// =====================================================
// Add Affinities
// =====================================================

async function addAffinities(req, res) {
  try {
    const { planId } = req.params;
    const items = req.body;

    if (!planId) {
      return res.status(400).json({
        success: false,
        error: "planId is required",
      });
    }

    const parsedPlanId = Number(planId);

    if (!Number.isInteger(parsedPlanId)) {
      return res.status(400).json({
        success: false,
        error: "planId must be a valid integer",
      });
    }

    if (!Array.isArray(items)) {
      return res.status(400).json({
        success: false,
        error: "Body must be an array",
      });
    }

    await saveAffinities(
      parsedPlanId,
      items
    );

    return res.json({
      success: true,
      message:
        "Affinities saved successfully",
      data: null,
    });
  } catch (err) {
    return handleError(
      res,
      err,
      "Error saving affinities"
    );
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
    } = req.body || {};

    if (!planId) {
      return res.status(400).json({
        success: false,
        error: "planId is required",
      });
    }

    const parsedPlanId = Number(planId);

    if (!Number.isInteger(parsedPlanId)) {
      return res.status(400).json({
        success: false,
        error: "planId must be a valid integer",
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

    console.log(
      `🚀 Generating plan ${parsedPlanId}, variant ${parsedVariant}`
    );

    const generated = await generatePlan(
      parsedPlanId,
      parsedVariant
    );

    console.log(
      "✅ Generated plan:",
      generated
    );

    return res.json({
      success: true,
      message:
        "Plan generated successfully",

      data:
        generated.result || [],

      stats:
        generated.stats || null,

      downloadUrl:
        `/exports/plan_${parsedPlanId}.xlsx`,
    });
  } catch (err) {
    return handleError(
      res,
      err,
      "Error generating plan"
    );
  }
}

// =====================================================
// Get One Plan
// =====================================================

async function getPlan(req, res) {
  try {
    const { planId } = req.params;

    console.log(
      "🔎 GET PLAN request:",
      planId
    );

    if (!planId) {
      return res.status(400).json({
        success: false,
        error: "planId is required",
      });
    }

    const parsedPlanId = Number(planId);

    if (!Number.isInteger(parsedPlanId)) {
      return res.status(400).json({
        success: false,
        error:
          "planId must be a valid integer",
      });
    }

    console.log(
      "🔎 Fetching plan:",
      parsedPlanId
    );

    const plan = await fetchPlan(
      parsedPlanId
    );

    if (!plan) {
      return res.status(404).json({
        success: false,
        error: "Plan not found",
      });
    }

    console.log(
      "✅ Plan fetched successfully:",
      parsedPlanId
    );

    return res.json({
      success: true,
      message:
        "Plan fetched successfully",
      data: plan,
    });
  } catch (err) {
    return handleError(
      res,
      err,
      "Error fetching plan"
    );
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
      message:
        "Plans fetched successfully",
      data: plans,
    });
  } catch (err) {
    return handleError(
      res,
      err,
      "Error fetching plans"
    );
  }
}

// =====================================================
// Lock Assignment
// =====================================================

async function lockAssignment(req, res) {
  try {
    const { planId } = req.params;

    const {
      sessionGroupId,
      supervisorId,
    } = req.body;

    if (
      !planId ||
      sessionGroupId == null ||
      supervisorId == null
    ) {
      return res.status(400).json({
        success: false,
        error:
          "planId, sessionGroupId and supervisorId are required",
      });
    }

    const parsedPlanId = Number(planId);
    const parsedSessionGroupId =
      Number(sessionGroupId);
    const parsedSupervisorId =
      Number(supervisorId);

    if (
      !Number.isInteger(parsedPlanId) ||
      !Number.isInteger(parsedSessionGroupId) ||
      !Number.isInteger(parsedSupervisorId)
    ) {
      return res.status(400).json({
        success: false,
        error:
          "planId, sessionGroupId and supervisorId must be valid integers",
      });
    }

    await lockRow(
      parsedPlanId,
      parsedSessionGroupId,
      parsedSupervisorId
    );

    return res.json({
      success: true,
      message:
        "Assignment locked successfully",
      data: null,
    });
  } catch (err) {
    return handleError(
      res,
      err,
      "Error locking assignment"
    );
  }
}

// =====================================================
// Unlock Assignment
// =====================================================

async function unlockAssignment(req, res) {
  try {
    const {
      planId,
      sessionGroupId,
    } = req.params;

    if (
      !planId ||
      !sessionGroupId
    ) {
      return res.status(400).json({
        success: false,
        error:
          "planId and sessionGroupId are required",
      });
    }

    const parsedPlanId = Number(planId);
    const parsedSessionGroupId =
      Number(sessionGroupId);

    if (
      !Number.isInteger(parsedPlanId) ||
      !Number.isInteger(parsedSessionGroupId)
    ) {
      return res.status(400).json({
        success: false,
        error:
          "planId and sessionGroupId must be valid integers",
      });
    }

    await unlockRow(
      parsedPlanId,
      parsedSessionGroupId
    );

    return res.json({
      success: true,
      message:
        "Assignment unlocked successfully",
      data: null,
    });
  } catch (err) {
    return handleError(
      res,
      err,
      "Error unlocking assignment"
    );
  }
}

// =====================================================
// Move Assignment
// =====================================================

async function moveAssignment(req, res) {
  try {
    const { planId } = req.params;

    const {
      fromSupervisorId,
      toSupervisorId,
      sessionGroupId,
    } = req.body;

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

    const parsedPlanId = Number(planId);
    const parsedFromSupervisorId =
      Number(fromSupervisorId);
    const parsedToSupervisorId =
      Number(toSupervisorId);
    const parsedSessionGroupId =
      Number(sessionGroupId);

    if (
      !Number.isInteger(parsedPlanId) ||
      !Number.isInteger(parsedFromSupervisorId) ||
      !Number.isInteger(parsedToSupervisorId) ||
      !Number.isInteger(parsedSessionGroupId)
    ) {
      return res.status(400).json({
        success: false,
        error:
          "All IDs must be valid integers",
      });
    }

    await moveAssignmentSvc(
      parsedPlanId,
      {
        fromSupervisorId:
          parsedFromSupervisorId,

        toSupervisorId:
          parsedToSupervisorId,

        sessionGroupId:
          parsedSessionGroupId,
      }
    );

    return res.json({
      success: true,
      message:
        "Assignment moved successfully",
      data: null,
    });
  } catch (err) {
    return handleError(
      res,
      err,
      "Error moving assignment"
    );
  }
}

// =====================================================
// Get Stats
// =====================================================

async function getStats(req, res) {
  try {
    const { planId } = req.params;

    if (!planId) {
      return res.status(400).json({
        success: false,
        error: "planId is required",
      });
    }

    const parsedPlanId = Number(planId);

    if (!Number.isInteger(parsedPlanId)) {
      return res.status(400).json({
        success: false,
        error:
          "planId must be a valid integer",
      });
    }

    const stats = await planStats(
      parsedPlanId
    );

    return res.json({
      success: true,
      message:
        "Plan stats fetched successfully",
      data: stats,
    });
  } catch (err) {
    return handleError(
      res,
      err,
      "Error fetching stats"
    );
  }
}

// =====================================================
// Exports
// =====================================================

module.exports = {
  createPlan,
  setDutyPool,
  setPeriodQuotas,
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
