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

planRouter.put("/:planId/professor/:professorId/room", async (req, res) => {
  try {
    const { planId, professorId } = req.params;
    const roomNumber = String(req.body.roomNumber ?? "").trim();
    const professorName = req.body.professorName ?? null;

    if (!roomNumber) {
      return res.status(400).json({ success: false, error: "رقم القاعة مطلوب" });
    }

    const result = await pool.query(
      `UPDATE room_assignments
          SET room_number = $1
        WHERE plan_id = $2 AND professor_id = $3`,
      [roomNumber, planId, professorId],
    );

    // لا يوجد سجل قاعة لهذا الأستاذ في هذه الخطة → أنشئ واحدًا
    if (!result.rowCount) {
      await pool.query(
        `INSERT INTO room_assignments (plan_id, professor_id, name, room_number)
         VALUES ($1, $2, $3, $4)`,
        [planId, professorId, professorName, roomNumber],
      );
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
