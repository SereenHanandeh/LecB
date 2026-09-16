const express = require("express");

const Plan = require("../controllers/plan.js");

const planRouter = express.Router();

planRouter.post("/", Plan.createPlan);

planRouter.post(
  "/:planId/duty-pool",
  Plan.setDutyPool
);

// ✅ Period Quotas
planRouter.post(
  "/:planId/period-quotas",
  Plan.setPeriodQuotas
);

planRouter.post(
  "/:planId/preassign",
  Plan.addPreassignments
);

planRouter.post(
  "/:planId/affinities",
  Plan.addAffinities
);

planRouter.post(
  "/:planId/generate",
  Plan.generate
);

planRouter.get(
  "/:planId",
  Plan.getPlan
);

planRouter.get(
  "/:planId/stats",
  Plan.getStats
);

planRouter.patch(
  "/:planId/assignments",
  Plan.moveAssignment
);

planRouter.post(
  "/:planId/lock",
  Plan.lockAssignment
);

planRouter.delete(
  "/:planId/lock/:sessionGroupId",
  Plan.unlockAssignment
);

module.exports = planRouter;