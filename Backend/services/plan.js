const pool = require("../models/db");

// =====================================================
// إنشاء خطة جديدة
// =====================================================

async function createPlanRow({
  name,
  excelBatchId,
  dateFrom,
  dateTo,
}) {
  const result = await pool.query(
    `
    INSERT INTO plans (
      name,
      excel_batch_id,
      date_from,
      date_to
    )
    VALUES ($1, $2, $3, $4)
    RETURNING *
    `,
    [name, excelBatchId, dateFrom, dateTo]
  );

  return result.rows[0];
}

// =====================================================
// Duty Pool
// =====================================================

async function saveDutyPool(planId, supervisorIds = []) {
  await pool.query(
    `
    DELETE FROM duty_pool
    WHERE plan_window_id = $1
    `,
    [planId]
  );

  for (const supervisorId of supervisorIds) {
    const sid = Number(supervisorId);

    if (!Number.isInteger(sid)) {
      console.warn(
        `⚠️ Invalid supervisor ID: ${supervisorId}`
      );
      continue;
    }

    await pool.query(
      `
      INSERT INTO duty_pool (
        plan_window_id,
        supervisor_id
      )
      VALUES ($1, $2)
      ON CONFLICT DO NOTHING
      `,
      [planId, sid]
    );
  }
}

// =====================================================
// جلب Duty Pool
// =====================================================

async function getDutyPool(planId) {
  const result = await pool.query(
    `
    SELECT
      dp.supervisor_id,
      s.name AS supervisor_name
    FROM duty_pool dp
    JOIN supervisors s
      ON s.id = dp.supervisor_id
    WHERE dp.plan_window_id = $1
    ORDER BY s.id
    `,
    [planId]
  );

  return result.rows;
}

// =====================================================
// Period Quotas
// =====================================================
// =====================================================
// Period Quotas
// =====================================================

// حفظ عدد الفترات المستهدف لكل مشرف
//
// Dashboard يرسل Array بهذا الشكل:
//
// [
//   { supervisorId: 1, quota: 10 },
//   { supervisorId: 2, quota: 12 }
// ]

async function savePeriodQuotas(
  planId,
  supervisors = []
) {
  console.log(
    "🎯 SAVING PERIOD QUOTAS:",
    {
      planId,
      supervisors,
    }
  );

  // ---------------------------------------------------
  // حذف الـ quotas القديمة للخطة
  // ---------------------------------------------------

  await pool.query(
    `
    DELETE FROM plan_period_quotas
    WHERE plan_id = $1
    `,
    [planId]
  );

  // ---------------------------------------------------
  // التأكد من البيانات
  // ---------------------------------------------------

  if (
    !Array.isArray(supervisors) ||
    supervisors.length === 0
  ) {
    console.log(
      "ℹ️ No period quotas received."
    );

    return [];
  }

  const saved = [];

  // ---------------------------------------------------
  // حفظ كل Quota
  // ---------------------------------------------------

  for (const item of supervisors) {
    const sid = Number(
      item.supervisorId ??
      item.supervisor_id
    );

    const target = Number(
      item.quota ??
      item.targetPeriods ??
      item.target_periods
    );

    // -------------------------------------------------
    // التحقق من Supervisor ID
    // -------------------------------------------------

    if (!Number.isInteger(sid)) {
      console.warn(
        `⚠️ Invalid supervisor ID in quota:`,
        item
      );

      continue;
    }

    // -------------------------------------------------
    // التحقق من Target
    // -------------------------------------------------

    if (
      !Number.isInteger(target) ||
      target <= 0
    ) {
      console.warn(
        `⚠️ Invalid period quota for supervisor ${sid}:`,
        item
      );

      continue;
    }

    // -------------------------------------------------
    // التأكد أن المشرف موجود و Active
    // -------------------------------------------------

    const supervisorCheck =
      await pool.query(
        `
        SELECT id, name
        FROM supervisors
        WHERE id = $1
          AND active = TRUE
        `,
        [sid]
      );

    if (
      supervisorCheck.rowCount === 0
    ) {
      console.warn(
        `⚠️ Supervisor ${sid} does not exist or inactive`
      );

      continue;
    }

    // -------------------------------------------------
    // حفظ الـ Quota
    // -------------------------------------------------

    const result = await pool.query(
      `
      INSERT INTO plan_period_quotas (
        plan_id,
        supervisor_id,
        target_periods
      )
      VALUES ($1, $2, $3)

      ON CONFLICT (plan_id, supervisor_id)
      DO UPDATE SET
        target_periods =
          EXCLUDED.target_periods

      RETURNING *
      `,
      [
        planId,
        sid,
        target,
      ]
    );

    saved.push(
      result.rows[0]
    );

    console.log(
      `✅ Quota saved: supervisor ${sid} = ${target}`
    );
  }

  // ---------------------------------------------------
  // تحقق بعد الحفظ
  // ---------------------------------------------------

  const check =
    await getPeriodQuotas(planId);

  console.log(
    "🎯 QUOTAS AFTER SAVE:",
    check
  );

  return check;
}

// =====================================================
// جلب Period Quotas
// =====================================================

async function getPeriodQuotas(planId) {
  const result = await pool.query(
    `
    SELECT
      q.supervisor_id,
      q.target_periods,
      s.name AS supervisor_name
    FROM plan_period_quotas q
    JOIN supervisors s
      ON s.id = q.supervisor_id
    WHERE q.plan_id = $1
    ORDER BY s.id
    `,
    [planId]
  );

  return result.rows;
}

// =====================================================
// Preassignments
// =====================================================

async function savePreassignments(
  planId,
  items = []
) {
  await pool.query(
    `
    DELETE FROM preassignments
    WHERE plan_id = $1
    `,
    [planId]
  );

  for (const item of items) {
    const sessionGroupId = Number(
      item.sessionGroupId
    );

    const supervisorId = Number(
      item.supervisorId
    );

    if (!Number.isInteger(sessionGroupId)) {
      console.warn(
        `⚠️ Invalid session group ID: ${item.sessionGroupId}`
      );
      continue;
    }

    if (!Number.isInteger(supervisorId)) {
      console.warn(
        `⚠️ Invalid supervisor ID: ${item.supervisorId}`
      );
      continue;
    }

    // التأكد أن Session Group موجود
    const groupCheck = await pool.query(
      `
      SELECT id
      FROM session_groups
      WHERE id = $1
      `,
      [sessionGroupId]
    );

    if (groupCheck.rowCount === 0) {
      console.warn(
        `⚠️ Session group ${sessionGroupId} does not exist`
      );
      continue;
    }

    // التأكد أن Supervisor موجود
    const supervisorCheck = await pool.query(
      `
      SELECT id
      FROM supervisors
      WHERE id = $1
        AND active = TRUE
      `,
      [supervisorId]
    );

    if (supervisorCheck.rowCount === 0) {
      console.warn(
        `⚠️ Supervisor ${supervisorId} does not exist or inactive`
      );
      continue;
    }

    await pool.query(
      `
      INSERT INTO preassignments (
        plan_id,
        session_group_id,
        supervisor_id
      )
      VALUES ($1, $2, $3)
      ON CONFLICT DO NOTHING
      `,
      [
        planId,
        sessionGroupId,
        supervisorId,
      ]
    );
  }
}

// =====================================================
// Affinities
// Professor -> Supervisor
// =====================================================

async function saveAffinities(planId, items = []) {
  console.log("========================================");
  console.log("🔗 SAVING PROFESSOR AFFINITIES");
  console.log("📌 Plan ID:", planId);
  console.log("📌 Received items:", items);
  console.log("========================================");

  await pool.query(
    `
    DELETE FROM affinities
    WHERE plan_id = $1
    `,
    [planId]
  );

  if (!Array.isArray(items) || !items.length) {
    console.log("ℹ️ No affinities to save.");
    return [];
  }

  const saved = [];

  for (const item of items) {
    try {
      // =====================================================
      // Professor ID
      // =====================================================

      let professorId = Number(
        item.professorId ??
        item.professor_id ??
        NaN
      );

      // =====================================================
      // Professor Name
      // =====================================================

      let professorName = String(
        item.professorName ??
        item.professor_name ??
        item.name ??
        ""
      )
        .trim()
        .replace(/\s+/g, " ");

      // =====================================================
      // Supervisor ID
      // =====================================================

      const supervisorId = Number(
        item.supervisorId ??
        item.supervisor_id ??
        NaN
      );

      // =====================================================
      // Validate Supervisor
      // =====================================================

      if (!Number.isInteger(supervisorId)) {
        console.warn(
          "⚠️ Invalid supervisor ID:",
          supervisorId,
          item
        );
        continue;
      }

      const supervisorCheck = await pool.query(
        `
        SELECT id, name
        FROM supervisors
        WHERE id = $1
          AND active = TRUE
        `,
        [supervisorId]
      );

      if (supervisorCheck.rowCount === 0) {
        console.warn(
          `⚠️ Supervisor ${supervisorId} does not exist or inactive`
        );
        continue;
      }

      // =====================================================
      // إذا لم يصل Professor ID
      // نبحث عنه بالاسم
      // =====================================================

      if (!Number.isInteger(professorId)) {
        if (!professorName) {
          console.warn(
            "⚠️ Affinity ignored: no professor ID and no professor name.",
            item
          );
          continue;
        }

        const professorByName = await pool.query(
          `
          SELECT id, name
          FROM professors
          WHERE LOWER(TRIM(name)) =
                LOWER(TRIM($1))
          ORDER BY id
          LIMIT 1
          `,
          [professorName]
        );

        if (professorByName.rowCount === 0) {
          console.warn(
            `⚠️ Professor not found by name: "${professorName}"`
          );
          continue;
        }

        professorId = Number(
          professorByName.rows[0].id
        );

        professorName =
          professorByName.rows[0].name;
      }

      // =====================================================
      // تأكيد أن الأستاذ موجود
      // =====================================================

      const professorCheck = await pool.query(
        `
        SELECT id, name
        FROM professors
        WHERE id = $1
        `,
        [professorId]
      );

      if (professorCheck.rowCount === 0) {
        console.warn(
          `⚠️ Professor ${professorId} does not exist`
        );
        continue;
      }

      const professor =
        professorCheck.rows[0];

      const canonicalProfessorName =
        String(professor.name || professorName)
          .trim()
          .replace(/\s+/g, " ");

      // =====================================================
      // حفظ Affinity
      // =====================================================

      const insertResult = await pool.query(
        `
        INSERT INTO affinities (
          plan_id,
          name,
          professor_id,
          crn,
          supervisor_id
        )
        VALUES ($1, $2, $3, $4, $5)
        RETURNING *
        `,
        [
          planId,
          canonicalProfessorName,
          professorId,
          item.crn ?? null,
          supervisorId,
        ]
      );

      const savedAffinity =
        insertResult.rows[0];

      saved.push(savedAffinity);

      console.log(
        "✅ Affinity saved:",
        {
          planId,
          professorId,
          professorName: canonicalProfessorName,
          supervisorId,
          crn: item.crn ?? null,
        }
      );
    } catch (error) {
      console.error(
        "❌ Error saving affinity:",
        item,
        error
      );
    }
  }

  console.log("========================================");
  console.log(
    `🔗 Affinities saved: ${saved.length}/${items.length}`
  );
  console.log("========================================");

  return saved;
}

// =====================================================
// Plan Context
// =====================================================

async function getPlanContext(planId) {

  // ---------------------------------------------------
  // 1. الخطة
  // ---------------------------------------------------

  const planResult = await pool.query(
    `
    SELECT *
    FROM plans
    WHERE id = $1
    `,
    [planId]
  );

  if (planResult.rowCount === 0) {
    throw new Error(
      `Plan ${planId} not found`
    );
  }

  const plan = planResult.rows[0];

  // ---------------------------------------------------
  // 2. Session Groups
  // ---------------------------------------------------

  const sessionsResult = await pool.query(
    `
    SELECT
      sg.*,
      p.name AS professor_name
    FROM session_groups sg
    LEFT JOIN professors p
      ON p.id = sg.professor_id
    WHERE sg.excel_batch_id = $1
      AND sg.date BETWEEN $2 AND $3
    ORDER BY
      sg.date,
      sg.period_label,
      sg.id
    `,
    [
      plan.excel_batch_id,
      plan.date_from,
      plan.date_to,
    ]
  );

  // ---------------------------------------------------
  // 3. Supervisors
  // ---------------------------------------------------

  const supervisorsResult = await pool.query(
    `
    SELECT *
    FROM supervisors
    WHERE active = TRUE
    ORDER BY id
    `
  );

  // ---------------------------------------------------
  // 4. Duty Pool
  // ---------------------------------------------------

  const dutyPoolResult = await pool.query(
    `
    SELECT
      dp.supervisor_id,
      s.name AS supervisor_name
    FROM duty_pool dp
    JOIN supervisors s
      ON s.id = dp.supervisor_id
    WHERE dp.plan_window_id = $1
    ORDER BY s.id
    `,
    [planId]
  );

  // ---------------------------------------------------
  // 5. Period Quotas
  // ---------------------------------------------------

  const periodQuotasResult = await pool.query(
    `
    SELECT
      q.supervisor_id,
      q.target_periods,
      s.name AS supervisor_name
    FROM plan_period_quotas q
    JOIN supervisors s
      ON s.id = q.supervisor_id
    WHERE q.plan_id = $1
    ORDER BY s.id
    `,
    [planId]
  );

  // ---------------------------------------------------
  // 6. Preassignments
  // ---------------------------------------------------

  const preResult = await pool.query(
    `
    SELECT
      pa.*,
      s.name AS supervisor_name
    FROM preassignments pa
    LEFT JOIN supervisors s
      ON s.id = pa.supervisor_id
    WHERE pa.plan_id = $1
    `,
    [planId]
  );

  // ---------------------------------------------------
  // 7. Locks
  // ---------------------------------------------------

  const locksResult = await pool.query(
    `
    SELECT
      al.*,
      s.name AS supervisor_name
    FROM assignment_locks al
    LEFT JOIN supervisors s
      ON s.id = al.supervisor_id
    WHERE al.plan_id = $1
    `,
    [planId]
  );

  // ---------------------------------------------------
  // 8. Affinities
  // ---------------------------------------------------

  const affResult = await pool.query(
  `
  SELECT
    a.id,
    a.plan_id,
    a.professor_id,
    a.name,
    a.crn,
    a.supervisor_id,
    p.name AS professor_name,
    s.name AS supervisor_name
  FROM affinities a
  LEFT JOIN professors p
    ON p.id = a.professor_id
  LEFT JOIN supervisors s
    ON s.id = a.supervisor_id
  WHERE a.plan_id = $1
  ORDER BY a.professor_id, a.id
  `,
  [planId]
);

  return {
  plan,
  groups: sessionsResult.rows,
  supervisors: supervisorsResult.rows,
  dutyPool: dutyPoolResult.rows,
  periodQuotas: periodQuotasResult.rows,
  pre: preResult.rows,
  locks: locksResult.rows,
  affinities: affResult.rows,
    aff: affResult.rows,

};
}

// =====================================================
// Fetch Plan
// =====================================================

async function fetchPlan(planId) {

  // جلب Context الخطة
  const ctx = await getPlanContext(planId);

  // ---------------------------------------------------
  // جلب Assignments مع كل بيانات Session Group
  // ---------------------------------------------------

  const assignmentsResult = await pool.query(
    `
    SELECT
      a.id,
      a.plan_id,
      a.session_group_id,
      a.supervisor_id,

      -- بيانات المشرف
      s.name AS supervisor_name,

      -- بيانات Session Group
      sg.crn,
      sg.date,
      sg.period_label,
      sg.required_supervisors,
      sg.sessions,

      -- بيانات الأستاذ
      sg.professor_id,
      p.name AS professor_name

    FROM assignments a

    JOIN supervisors s
      ON s.id = a.supervisor_id

    JOIN session_groups sg
      ON sg.id = a.session_group_id

    LEFT JOIN professors p
      ON p.id = sg.professor_id

    WHERE a.plan_id = $1

    ORDER BY
      sg.date,
      sg.period_label,
      sg.id,
      a.id
    `,
    [planId]
  );

  return {
    ...ctx,
    assignments: assignmentsResult.rows,
  };
}

// =====================================================
// Locks
// =====================================================

async function lockRow(
  planId,
  sessionGroupId,
  supervisorId
) {
  await pool.query(
    `
    INSERT INTO assignment_locks (
      plan_id,
      session_group_id,
      supervisor_id
    )
    VALUES ($1, $2, $3)
    ON CONFLICT (plan_id, session_group_id)
    DO UPDATE SET
      supervisor_id = EXCLUDED.supervisor_id
    `,
    [
      planId,
      sessionGroupId,
      supervisorId,
    ]
  );
}

async function unlockRow(
  planId,
  sessionGroupId
) {
  await pool.query(
    `
    DELETE FROM assignment_locks
    WHERE plan_id = $1
      AND session_group_id = $2
    `,
    [planId, sessionGroupId]
  );
}

// =====================================================
// Assignments
// =====================================================

async function clearAssignments(planId) {
  await pool.query(
    `
    DELETE FROM assignments
    WHERE plan_id = $1
    `,
    [planId]
  );
}

async function saveAssignments(
  planId,
  assignments = []
) {
  await clearAssignments(planId);

  for (const assignment of assignments) {
    const sessionGroupId = Number(
      assignment.session_group_id
    );

    const supervisorId = Number(
      assignment.supervisor_id
    );

    if (!Number.isInteger(sessionGroupId)) {
      console.warn(
        `⚠️ Invalid session group ID: ${assignment.session_group_id}`
      );
      continue;
    }

    if (!Number.isInteger(supervisorId)) {
      console.warn(
        `⚠️ Invalid supervisor ID: ${assignment.supervisor_id}`
      );
      continue;
    }

    await pool.query(
      `
      INSERT INTO assignments (
        plan_id,
        session_group_id,
        supervisor_id
      )
      VALUES ($1, $2, $3)
      ON CONFLICT DO NOTHING
      `,
      [
        planId,
        sessionGroupId,
        supervisorId,
      ]
    );
  }
}

// =====================================================
// Move Assignment
// =====================================================

async function moveAssignmentSvc(
  planId,
  {
    fromSupervisorId,
    toSupervisorId,
    sessionGroupId,
  }
) {
  const groupId = Number(sessionGroupId);
  const toId = Number(toSupervisorId);

  if (!Number.isInteger(groupId)) {
    throw new Error("Invalid sessionGroupId");
  }

  if (!Number.isInteger(toId)) {
    throw new Error("Invalid toSupervisorId");
  }

  // =====================================================
// 🔗 التحقق من Affinity
// =====================================================

const affinityResult = await pool.query(
  `
  SELECT
    a.supervisor_id,
    a.name,
    a.professor_id
  FROM affinities a
  JOIN session_groups sg
    ON sg.professor_id = a.professor_id
  WHERE a.plan_id = $1
    AND sg.id = $2
  LIMIT 1
  `,
  [planId, groupId]
);

if (affinityResult.rowCount > 0) {
  const affinitySupervisorId =
    Number(affinityResult.rows[0].supervisor_id);

  if (affinitySupervisorId !== toId) {
    throw new Error(
      `Professor is assigned to supervisor ${affinitySupervisorId} by affinity`
    );
  }
}

  // احذف أي assignment موجود لهذا الـ Session Group
  await pool.query(
    `
    DELETE FROM assignments
    WHERE plan_id = $1
      AND session_group_id = $2
    `,
    [planId, groupId]
  );

  // أضف المشرف الجديد
  await pool.query(
    `
    INSERT INTO assignments (
      plan_id,
      session_group_id,
      supervisor_id
    )
    VALUES ($1, $2, $3)
    `,
    [
      planId,
      groupId,
      toId,
    ]
  );

  // التعديل اليدوي يعتبر Lock
  await lockRow(
    planId,
    groupId,
    toId
  );
}

// =====================================================
// Statistics
// =====================================================

async function planStats(planId) {
  const distributionResult =
    await pool.query(
      `
      SELECT
        s.id AS supervisor_id,
        s.name AS supervisor_name,

        COUNT(a.id) AS assignments,

        COUNT(
          DISTINCT g.date || '|' || g.period_label
        ) AS working_slots

      FROM supervisors s

      LEFT JOIN assignments a
        ON a.supervisor_id = s.id
        AND a.plan_id = $1

      LEFT JOIN session_groups g
        ON g.id = a.session_group_id

      WHERE s.active = TRUE

      GROUP BY
        s.id,
        s.name

      ORDER BY
        s.id
      `,
      [planId]
    );

  return {
    distribution:
      distributionResult.rows,
  };
}

// =====================================================
// جلب جميع الخطط السابقة
// =====================================================

async function getAllPlans() {
  const result = await pool.query(`
    SELECT
      p.id,
      p.name,
      p.excel_batch_id,
      p.date_from,
      p.date_to,
      p.created_at,

      -- عدد المشرفين المختارين للخطة
      (
        SELECT COUNT(*)
        FROM duty_pool dp
        WHERE dp.plan_window_id = p.id
      ) AS supervisor_count,

      -- عدد التعيينات الموجودة بالخطة
      (
        SELECT COUNT(*)
        FROM assignments a
        WHERE a.plan_id = p.id
      ) AS assignment_count

    FROM plans p

    ORDER BY
      p.created_at DESC,
      p.id DESC
  `);

  return result.rows;
}


module.exports = {
  createPlanRow,

  saveDutyPool,
  getDutyPool,

  savePeriodQuotas,
  getPeriodQuotas,

  savePreassignments,
  saveAffinities,

  getPlanContext,
  fetchPlan,

  lockRow,
  unlockRow,

  clearAssignments,
  saveAssignments,

  moveAssignmentSvc,

  planStats,
  getAllPlans,
};