const express = require("express");
const pool = require('../models/db');

const Plan = require("../controllers/plan.js");

const planRouter = express.Router();

planRouter.get("/", Plan.getPlans);

planRouter.get("/accepted-supervisor-stats", Plan.getAcceptedSupervisorStats);
planRouter.post("/", Plan.createPlan);

planRouter.post("/:planId/duty-pool", Plan.setDutyPool);

planRouter.post("/:planId/preassign", Plan.addPreassignments);

planRouter.post("/:planId/affinities", Plan.addAffinities);

planRouter.post("/:planId/generate", Plan.generate);

planRouter.post("/:planId/status", Plan.updateStatus);

planRouter.get("/:planId", Plan.getPlan);

planRouter.post("/:planId/rooms", Plan.setRoomAssignments);

planRouter.put("/:planId/group/:groupId/room", async (req, res) => {
  try {
    const { groupId } = req.params;
    const roomNumber = String(req.body.roomNumber ?? "").trim();

    if (!roomNumber) {
      return res.status(400).json({ success: false, error: "رقم القاعة مطلوب" });
    }

    const result = await pool.query(
      `UPDATE session_groups SET room_number = $1 WHERE id = $2 RETURNING id`,
      [roomNumber, groupId],
    );

    if (!result.rowCount) {
      return res.status(404).json({ success: false, error: "المجموعة غير موجودة" });
    }

    res.json({ success: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: err.message });
  }
});

planRouter.get("/:planId/stats", Plan.getStats);

planRouter.patch("/:planId/assignments", Plan.moveAssignment);

planRouter.post("/:planId/lock", Plan.lockAssignment);

planRouter.delete("/:planId/lock/:sessionGroupId", Plan.unlockAssignment);

planRouter.delete("/:planId", Plan.deletePlan);

module.exports = planRouter;
