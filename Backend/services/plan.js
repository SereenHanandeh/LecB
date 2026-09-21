const pool = require("../models/db");

// =====================================================
// إنشاء خطة جديدة
// =====================================================

async function createPlanRow({
  name,
  excelBatchId,
  dateFrom,
  dateTo,
  category,
}) {
  const result = await pool.query(
    `
    INSERT INTO plans (
      name,
      excel_batch_id,
      date_from,
      date_to,
      category
    )
    VALUES ($1, $2, $3, $4, $5)
    RETURNING *
    `,
    [
      name,
      excelBatchId,
      dateFrom,
      dateTo,
      category,
    ]
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
  // 5. Preassignments
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

async function generatePlan(
  planId,
  variant = 1,
  minimumPeriodsEnabled = false,
  minimumPeriods = 4,
) {
  console.log("");
  console.log("==================================================");
  console.log("🚀 STARTING PLAN GENERATION");
  console.log("==================================================");
  console.log(`🆔 Plan ID: ${planId}`);
  console.log(`🔢 Variant: ${variant}`);

  // ========================================================
  // Support object options
  // ========================================================

  if (minimumPeriodsEnabled && typeof minimumPeriodsEnabled === "object") {
    const options = minimumPeriodsEnabled;

    minimumPeriodsEnabled = Boolean(
      options.minimumPeriodsEnabled ?? options.enabled ?? false,
    );

    minimumPeriods = Number(options.minimumPeriods ?? options.minimum ?? 4);
  }

  const minimumEnabled = Boolean(minimumPeriodsEnabled);

  let minimumTarget = Number(minimumPeriods);

  if (!Number.isInteger(minimumTarget) || minimumTarget < 1) {
    minimumTarget = 4;
  }

  console.log(
    `🎯 Minimum Period Rule ${minimumEnabled ? "ENABLED" : "DISABLED"}.`,
  );

  if (minimumEnabled) {
    console.log(`🎯 Minimum target: ${minimumTarget} periods`);
  }

  // ========================================================
  // Get plan context
  // ========================================================

  const context = await getPlanContext(planId);

  if (!context) {
    throw new Error(`Plan ${planId} not found.`);
  }

  const {
    plan,
    sessionGroups,
    supervisors,
    preassignments,
    locks,
    affinities,
  } = context;

  // ========================================================
  // Duty pool
  // ========================================================

  const dutyPool = await getDutyPool(planId);

  const selectedSupervisorIds = (
    dutyPool?.supervisorIds ??
    dutyPool?.supervisor_ids ??
    supervisors?.map((s) => s.id) ??
    []
  )
    .map(Number)
    .filter((id) => Number.isInteger(id));

  if (!selectedSupervisorIds.length) {
    throw new Error("No supervisors selected for this plan.");
  }

  console.log(
    `👥 Selected supervisors: ${selectedSupervisorIds.length}`,
    selectedSupervisorIds,
  );

  // ========================================================
  // Candidate state
  // ========================================================

  const cand = {};

  for (const supervisorId of selectedSupervisorIds) {
    const id = Number(supervisorId);

    cand[id] = {
      id,

      total: 0,

      lastDay: null,

      byDay: {},

      byDayPeriods: {},

      byDayPeriodRanks: {},

      occupiedSlots: new Set(),

      assignedGroups: new Set(),

      assignedProfessors: new Set(),
    };
  }

  // ========================================================
  // Result containers
  // ========================================================

  const result = [];

  const bundleAssignments = new Map();

  const professorAssignments = new Map();

  // ========================================================
  // Affinity maps
  // ========================================================

  const affinityByProfessorId = new Map();

  const affinityByProfessorName = new Map();

  if (Array.isArray(affinities)) {
    for (const affinity of affinities) {
      const supervisorId = Number(
        affinity.supervisor_id ?? affinity.supervisorId,
      );

      if (!Number.isInteger(supervisorId)) {
        continue;
      }

      const professorId = affinity.professor_id ?? affinity.professorId;

      const professorName =
        affinity.professor_name ?? affinity.professorName ?? affinity.name;

      if (professorId !== undefined && professorId !== null) {
        affinityByProfessorId.set(String(professorId), supervisorId);
      }

      if (professorName) {
        affinityByProfessorName.set(
          getProfessorNameKey(professorName),
          supervisorId,
        );
      }
    }
  }

  // ========================================================
  // Normalize session groups
  // ========================================================

  const normalizedGroups = Array.isArray(sessionGroups) ? sessionGroups : [];

  console.log(`📦 Total CRNs / groups: ${normalizedGroups.length}`);

  // ========================================================
  // Build bundles
  //
  // Bundle:
  // Professor + Date + Period
  //
  // Every group inside same bundle must stay together.
  // ========================================================

  const bundlesMap = new Map();

  for (const row of normalizedGroups) {
    const professorKey = getProfessorKey(row);

    const professorName =
      row.professor_name ??
      row.professorName ??
      row.Professor ??
      row.professor ??
      "";

    const day = dateISO(row.date ?? row.Date ?? row.DATE ?? row.day ?? row.Day);

    const period = normalizePeriod(
      row.period ?? row.Period ?? row.period_name ?? row.periodName,
    );

    if (!professorKey || !day || !period) {
      console.warn("⚠️ Skipping invalid row:", row);

      continue;
    }

    const bundleKey = `${day}|${professorKey}|${period}`;

    if (!bundlesMap.has(bundleKey)) {
      bundlesMap.set(bundleKey, {
        key: bundleKey,

        professorKey,

        professorName,

        date: day,

        period,

        groups: [],
      });
    }

    bundlesMap.get(bundleKey).groups.push(row);
  }

  const bundles = Array.from(bundlesMap.values());

  console.log(`📦 Total Professor + Date + Period bundles: ${bundles.length}`);

  // ========================================================
  // Group bundles by professor
  // ========================================================

  const professorMap = new Map();

  for (const bundle of bundles) {
    if (!professorMap.has(bundle.professorKey)) {
      professorMap.set(bundle.professorKey, {
        key: bundle.professorKey,

        professor_id:
          bundle.groups?.[0]?.professor_id ??
          bundle.groups?.[0]?.professorId ??
          null,

        professor_name: bundle.professorName,

        bundles: [],
      });
    }

    professorMap.get(bundle.professorKey).bundles.push(bundle);
  }

  const professorGroups = Array.from(professorMap.values());

  console.log(`👨‍🏫 Unique professors: ${professorGroups.length}`);

  // ========================================================
  // Sort bundles inside each professor
  // ========================================================

  for (const professor of professorGroups) {
    professor.bundles.sort((a, b) => {
      const dateCompare = String(a.date).localeCompare(String(b.date));

      if (dateCompare !== 0) {
        return dateCompare;
      }

      return getPeriodRank(a.period) - getPeriodRank(b.period);
    });
  }

  // ========================================================
  // Sort professors
  //
  // Larger professors first.
  // This is the existing basic assignment strategy.
  // ========================================================

  professorGroups.sort((a, b) => b.bundles.length - a.bundles.length);

  // ========================================================
  // Preassignment / locks
  // ========================================================

  const forcedProfessorAssignments = new Map();

  // ========================================================
  // Helper:
  // Register forced assignment
  // ========================================================

  const registerForcedAssignment = (professorKey, supervisorId, reason) => {
    const sid = Number(supervisorId);

    if (!selectedSupervisorIds.includes(sid)) {
      console.warn(
        `⚠️ Forced assignment ignored. Supervisor ${sid} is not in duty pool.`,
      );

      return;
    }

    const existing = forcedProfessorAssignments.get(professorKey);

    if (existing !== undefined && Number(existing) !== sid) {
      throw new Error(
        `Professor ${professorKey} has conflicting forced supervisors: ${existing} and ${sid}.`,
      );
    }

    forcedProfessorAssignments.set(professorKey, sid);

    console.log(
      `🔒 Forced professor ${professorKey} -> Supervisor ${sid} (${reason})`,
    );
  };

  // ========================================================
  // Load affinities as forced assignments
  // ========================================================

  for (const professor of professorGroups) {
    let affinitySupervisor = null;

    if (
      professor.professor_id !== null &&
      professor.professor_id !== undefined
    ) {
      affinitySupervisor = affinityByProfessorId.get(
        String(professor.professor_id),
      );
    }

    if (affinitySupervisor === undefined || affinitySupervisor === null) {
      affinitySupervisor = affinityByProfessorName.get(
        getProfessorNameKey(professor.professor_name),
      );
    }

    if (affinitySupervisor !== undefined && affinitySupervisor !== null) {
      registerForcedAssignment(professor.key, affinitySupervisor, "affinity");
    }
  }

  // ========================================================
  // Load preassignments
  // ========================================================

  if (Array.isArray(preassignments)) {
    for (const pre of preassignments) {
      const professorKey =
        pre.professor_key ?? pre.professorKey ?? getProfessorKey(pre);

      const supervisorId = pre.supervisor_id ?? pre.supervisorId;

      if (professorKey && supervisorId !== undefined && supervisorId !== null) {
        registerForcedAssignment(professorKey, supervisorId, "preassignment");
      }
    }
  }

  // ========================================================
  // Load locks
  // ========================================================

  if (Array.isArray(locks)) {
    for (const lock of locks) {
      const professorKey =
        lock.professor_key ?? lock.professorKey ?? getProfessorKey(lock);

      const supervisorId = lock.supervisor_id ?? lock.supervisorId;

      if (professorKey && supervisorId !== undefined && supervisorId !== null) {
        registerForcedAssignment(professorKey, supervisorId, "lock");
      }
    }
  }

  console.log(`🔒 Forced professors: ${forcedProfessorAssignments.size}`);

  // ========================================================
  // Helper:
  // Get forced supervisor
  // ========================================================

  const getForcedSupervisor = (professorKey) =>
    forcedProfessorAssignments.get(professorKey);

  // ========================================================
  // Assign forced professors first
  // ========================================================

  console.log("");
  console.log("==================================================");
  console.log("🔒 APPLYING FORCED PROFESSOR ASSIGNMENTS");
  console.log("==================================================");

  for (const professor of professorGroups) {
    const forcedSupervisorId = getForcedSupervisor(professor.key);

    if (forcedSupervisorId === undefined || forcedSupervisorId === null) {
      continue;
    }

    const supervisor = cand[Number(forcedSupervisorId)];

    if (!supervisor) {
      throw new Error(
        `Forced supervisor ${forcedSupervisorId} does not exist in candidate state.`,
      );
    }

    const canTake = canSupervisorTakeProfessor(supervisor, professor.bundles);

    if (!canTake) {
      throw new Error(
        `Forced assignment for professor ${professor.key} violates assignment rules.`,
      );
    }

    const assigned = assignProfessorToSupervisor(
      professor.key,
      professor.bundles,
      Number(forcedSupervisorId),
      result,
      cand,
      bundleAssignments,
      professorAssignments,
    );

    if (!assigned) {
      throw new Error(
        `Failed to apply forced assignment for professor ${professor.key}.`,
      );
    }

    console.log(
      `🔒 Professor assigned completely: ${professor.key} -> Supervisor ${forcedSupervisorId}`,
    );
  }

  // ========================================================
  // NORMAL BASIC PROFESSOR ASSIGNMENT
  //
  // IMPORTANT:
  // This section intentionally remains the basic
  // professor assignment logic.
  // ========================================================

  console.log("");
  console.log("==================================================");
  console.log("⚖️ STARTING NORMAL FAIR DISTRIBUTION");
  console.log("==================================================");

  // ========================================================
  // Remaining professors
  // ========================================================

  const remainingProfessors = professorGroups.filter(
    (professor) => !professorAssignments.has(professor.key),
  );

  // ========================================================
  // Basic assignment
  // ========================================================

  for (const professor of remainingProfessors) {
    let bestCandidate = null;

    for (const supervisorId of selectedSupervisorIds) {
      const supervisor = cand[Number(supervisorId)];

      if (!supervisor) {
        continue;
      }

      // ----------------------------------------------------
      // Hard rules
      // ----------------------------------------------------

      const canTake = canSupervisorTakeProfessor(supervisor, professor.bundles);

      if (!canTake) {
        continue;
      }

      // ----------------------------------------------------
      // Current workload
      // ----------------------------------------------------

      const currentTotal = Number(supervisor.total || 0);

      const projectedTotal = currentTotal + professor.bundles.length;

      // ----------------------------------------------------
      // Minimum information
      //
      // Used only by the existing basic assignment
      // scoring path.
      //
      // Actual minimum redistribution happens later
      // in Minimum Rebalance Engine.
      // ----------------------------------------------------
      const minimumInfo = getMinimumInfo(
        supervisor,
        professor.bundles.length,
        minimumEnabled,
        minimumTarget,
      );

      // ----------------------------------------------------
      // Continuity
      // ----------------------------------------------------

      const continuityScore = getProfessorConsecutiveScore(
        supervisor,
        professor.bundles,
      );

      const periodContinuityScore = getPeriodContinuityScore(
        supervisor,
        professor.bundles,
      );

      const continuityPriority = getContinuityPriority(
        supervisor,
        professor.bundles,
      );

      const dailyLoad = getProfessorDailyLoad(supervisor, professor.bundles);

      // ----------------------------------------------------
      // Existing score
      // ----------------------------------------------------

      const score =
        (minimumInfo?.score ?? 0) +
        continuityScore +
        periodContinuityScore +
        continuityPriority -
        dailyLoad;

      // ----------------------------------------------------
      // Candidate object
      // ----------------------------------------------------

      const candidate = {
        supervisorId: Number(supervisorId),

        score,

        projectedTotal,

        minimumInfo,

        continuityScore,

        periodContinuityScore,

        continuityPriority,

        dailyLoad,

        professorCount: Number(supervisor.assignedProfessors?.size || 0),

        random: Math.random(),
      };

      // ----------------------------------------------------
      // Existing comparator
      //
      // Keep the basic assignment philosophy:
      // minimum / projected total / continuity /
      // fairness.
      // ----------------------------------------------------

      if (!bestCandidate) {
        bestCandidate = candidate;

        continue;
      }

      // ----------------------------------------------------
      // Minimum preference
      //
      // Only when enabled.
      // ----------------------------------------------------

      if (minimumEnabled) {
        const currentBelow =
          Number(cand[bestCandidate.supervisorId]?.total || 0) < minimumTarget;

        const candidateBelow = currentTotal < minimumTarget;

        if (candidateBelow !== currentBelow) {
          if (!candidateBelow) {
            bestCandidate = candidate;
          }

          continue;
        }
      }

      // ----------------------------------------------------
      // Projected total
      // ----------------------------------------------------

      if (candidate.projectedTotal !== bestCandidate.projectedTotal) {
        if (candidate.projectedTotal < bestCandidate.projectedTotal) {
          bestCandidate = candidate;
        }

        continue;
      }

      // ----------------------------------------------------
      // Period continuity
      // ----------------------------------------------------

      if (
        candidate.periodContinuityScore !== bestCandidate.periodContinuityScore
      ) {
        if (
          candidate.periodContinuityScore > bestCandidate.periodContinuityScore
        ) {
          bestCandidate = candidate;
        }

        continue;
      }

      // ----------------------------------------------------
      // Continuity priority
      // ----------------------------------------------------

      if (candidate.continuityPriority !== bestCandidate.continuityPriority) {
        if (candidate.continuityPriority > bestCandidate.continuityPriority) {
          bestCandidate = candidate;
        }

        continue;
      }

      // ----------------------------------------------------
      // Consecutive score
      // ----------------------------------------------------

      if (candidate.continuityScore !== bestCandidate.continuityScore) {
        if (candidate.continuityScore > bestCandidate.continuityScore) {
          bestCandidate = candidate;
        }

        continue;
      }

      // ----------------------------------------------------
      // Daily load
      // ----------------------------------------------------

      if (candidate.dailyLoad !== bestCandidate.dailyLoad) {
        if (candidate.dailyLoad < bestCandidate.dailyLoad) {
          bestCandidate = candidate;
        }

        continue;
      }

      // ----------------------------------------------------
      // Professor count
      // ----------------------------------------------------

      if (candidate.professorCount !== bestCandidate.professorCount) {
        if (candidate.professorCount < bestCandidate.professorCount) {
          bestCandidate = candidate;
        }

        continue;
      }

      // ----------------------------------------------------
      // Seeded/random tie breaker
      // ----------------------------------------------------

      if (candidate.random < bestCandidate.random) {
        bestCandidate = candidate;
      }
    }

    // ======================================================
    // No candidate
    // ======================================================

    if (!bestCandidate) {
      throw new Error(
        `Unable to assign professor ${professor.key} to any selected supervisor.`,
      );
    }

    // ======================================================
    // Assign whole professor
    // ======================================================

    const assigned = assignProfessorToSupervisor(
      professor.key,
      professor.bundles,
      bestCandidate.supervisorId,
      result,
      cand,
      bundleAssignments,
      professorAssignments,
    );

    if (!assigned) {
      throw new Error(
        `Failed to assign professor ${professor.key} to supervisor ${bestCandidate.supervisorId}.`,
      );
    }

    console.log(
      `👨‍🏫 Professor assigned completely: ${professor.key} -> Supervisor ${bestCandidate.supervisorId}`,
    );
  }

  // ========================================================
  // Verify normal assignment
  // ========================================================

  rebuildCandidateState(
    cand,
    bundles,
    professorGroups,
    bundleAssignments,
    professorAssignments,
    result,
  );

  // ========================================================
  // Basic assignment statistics
  // ========================================================

  console.log("");
  console.log("==================================================");
  console.log("📊 BASIC ASSIGNMENT COMPLETE");
  console.log("==================================================");

  for (const supervisorId of selectedSupervisorIds) {
    console.log(
      `👤 Supervisor ${supervisorId}: ${
        cand[supervisorId]?.total ?? 0
      } periods`,
    );
  }

  // ========================================================
  // REBALANCE ENGINE
  //
  // Existing fairness rebalance.
  //
  // IMPORTANT:
  // This happens AFTER Basic Professor Assignment.
  // ========================================================

  console.log("");
  console.log("==================================================");
  console.log("⚖️ STARTING REBALANCE ENGINE");
  console.log("==================================================");

  const rebalanceResult = rebalanceAssignments({
    cand,

    selectedSupervisorIds,

    professorGroups,

    bundles,

    result,

    bundleAssignments,

    professorAssignments,

    forcedProfessorAssignments,

    minimumEnabled,

    minimumTarget,

    variant,
  });

  console.log("🔄 Rebalance result:", rebalanceResult);

  // ========================================================
  // MINIMUM REBALANCE ENGINE
  //
  // IMPORTANT:
  // This is the actual post-processing minimum engine.
  //
  // Example:
  //
  // Before:
  // 7, 6, 4, 4, 4, 4, 2, 4
  //
  // Minimum = 4
  //
  // Possible result:
  // 5, 6, 4, 4, 4, 4, 4, 4
  //
  // ========================================================

  console.log("");
  console.log("==================================================");
  console.log("🎯 MINIMUM REBALANCE");
  console.log("==================================================");

  console.log("🔥🔥🔥 MINIMUM REBALANCE TEST MARKER 🔥🔥🔥");

  const minimumRebalanceResult = minimumRebalanceAssignments({
    cand,

    selectedSupervisorIds,

    professorGroups,

    bundles,

    result,

    bundleAssignments,

    professorAssignments,

    forcedProfessorAssignments,

    minimumEnabled,

    minimumTarget,

    variant,
  });

  console.log("🎯 Minimum Rebalance result:", minimumRebalanceResult);

  console.log("🔄 SWAP REBALANCE");

  const swapRebalanceResult = swapProfessorsBetweenSupervisors({
    cand,

    selectedSupervisorIds,

    professorGroups,

    bundles,

    result,

    bundleAssignments,

    professorAssignments,

    forcedProfessorAssignments,

    minimumEnabled,

    minimumTarget,

    variant,
  });

  console.log("🔄 Swap Rebalance result:", swapRebalanceResult);

  // ========================================================
  // FINAL REBUILD
  // ========================================================

  rebuildCandidateState(
    cand,
    bundles,
    professorGroups,
    bundleAssignments,
    professorAssignments,
    result,
  );

  // ========================================================
  // Final verification
  // ========================================================

  const assignedGroupIds = new Set(
    result
      .map((row) =>
        Number(row.session_group_id ?? row.sessionGroupId ?? row.id),
      )
      .filter((id) => Number.isInteger(id)),
  );

  const totalGroups = normalizedGroups.length;

  const assignedGroups = assignedGroupIds.size;

  // ========================================================
  // Final bundle count
  // ========================================================

  const assignedBundles = Array.from(bundleAssignments.keys()).length;

  // ========================================================
  // Final metrics
  // ========================================================

  const finalMetrics = calculateRebalanceMetrics(
    cand,
    selectedSupervisorIds,
    minimumEnabled,
    minimumTarget,
  );

  // ========================================================
  // Conflicts
  // ========================================================

  let conflicts = 0;

  for (const supervisorId of selectedSupervisorIds) {
    const supervisor = cand[supervisorId];

    if (!supervisor) {
      continue;
    }

    const seenSlots = new Set();

    for (const row of result) {
      const assignedSupervisor = Number(row.supervisor_id ?? row.supervisorId);

      if (assignedSupervisor !== Number(supervisorId)) {
        continue;
      }

      const day = dateISO(row.date ?? row.Date);

      const period = normalizePeriod(row.period ?? row.Period);

      const slotKey = `${day}|${period}`;

      if (seenSlots.has(slotKey)) {
        conflicts++;
      }

      seenSlots.add(slotKey);
    }
  }

  // ========================================================
  // Professor consistency verification
  // ========================================================

  let professorViolations = 0;

  for (const professor of professorGroups) {
    const assignedSupervisor = professorAssignments.get(professor.key);

    if (assignedSupervisor === undefined || assignedSupervisor === null) {
      professorViolations++;
      continue;
    }

    for (const bundle of professor.bundles) {
      const bundleSupervisor = bundleAssignments.get(bundle.key);

      if (Number(bundleSupervisor) !== Number(assignedSupervisor)) {
        professorViolations++;

        break;
      }
    }
  }

  // ========================================================
  // Minimum statistics
  // ========================================================

  const minimumStatistics = selectedSupervisorIds.map((supervisorId) => {
    const total = Number(cand[supervisorId]?.total || 0);

    return {
      supervisorId: Number(supervisorId),

      periods: total,

      minimum: minimumEnabled ? minimumTarget : null,

      reachedMinimum: minimumEnabled ? total >= minimumTarget : true,

      deficit: minimumEnabled ? Math.max(0, minimumTarget - total) : 0,
    };
  });

  // ========================================================
  // Logs
  // ========================================================

  console.log("");
  console.log("==================================================");
  console.log("🏁 PLAN GENERATION FINISHED");
  console.log("==================================================");

  console.log(`📦 Total groups: ${totalGroups}`);

  console.log(`✅ Assigned groups: ${assignedGroups}/${totalGroups}`);

  console.log(`📦 Bundles assigned: ${assignedBundles}/${bundles.length}`);

  console.log(`👨‍🏫 Unique professors: ${professorGroups.length}`);

  console.log(
    `👥 Supervisors used: ${
      selectedSupervisorIds.filter((id) => Number(cand[id]?.total || 0) > 0)
        .length
    }`,
  );

  console.log(`⚠️ Conflicts: ${conflicts}`);

  console.log(`⚠️ Professor violations: ${professorViolations}`);

  console.log(`⚖️ Fairness difference: ${finalMetrics.difference}`);

  console.log(`📊 Fairness sum squared: ${finalMetrics.sumSquared}`);

  if (minimumEnabled) {
    console.log(
      `🎯 Supervisors below minimum: ${finalMetrics.supervisorsBelowMinimum}`,
    );

    console.log(
      `🎯 Total minimum deficit: ${finalMetrics.totalMinimumDeficit}`,
    );
  }

  console.log(`🔄 Rebalance moves: ${rebalanceResult.totalMoves ?? 0}`);

  console.log(
    `🎯 Minimum Rebalance moves: ${minimumRebalanceResult.moves ?? 0}`,
  );

  console.log(
    `🔄 Swap Rebalance moves: ${swapRebalanceResult.totalSwaps ?? 0}`,
  );

  // ========================================================
  // Safety validation
  // ========================================================

  if (assignedGroups !== totalGroups) {
    throw new Error(
      `Plan generation incomplete: ${assignedGroups}/${totalGroups} groups assigned.`,
    );
  }

  if (conflicts > 0) {
    throw new Error(
      `Plan generation produced ${conflicts} supervisor slot conflicts.`,
    );
  }

  if (professorViolations > 0) {
    throw new Error(
      `Plan generation produced ${professorViolations} professor assignment violations.`,
    );
  }

  // ========================================================
  // Save assignments
  // ========================================================

  console.log("========================================");
console.log("🔍 FINAL RESULT BEFORE SAVE");
console.log("📦 result is array:", Array.isArray(result));
console.log("📦 result length:", result?.length);
console.log("📦 result:", JSON.stringify(result, null, 2));
console.log("========================================");

  await saveAssignments(planId, result);

  console.log(`💾 Assignments saved for plan ${planId}`);

  // ========================================================
  // Export Excel
  // ========================================================

  const exportDir = path.join(__dirname, "../exports");

  if (!fs.existsSync(exportDir)) {
    fs.mkdirSync(exportDir, {
      recursive: true,
    });
  }

  const exportRows = result.map((row) => ({
    "Session Group ID":
      row.session_group_id ?? row.sessionGroupId ?? row.id ?? "",

    CRN: row.crn ?? row.CRN ?? "",

    Professor:
      row.professor_name ??
      row.professorName ??
      row.professor ??
      row.Professor ??
      "",

    Date: row.date ?? row.Date ?? "",

    Period: row.period ?? row.Period ?? "",

    Supervisor:
      row.supervisor_name ??
      row.supervisorName ??
      row.supervisor ??
      row.Supervisor ??
      "",
  }));

  const worksheet = xlsx.utils.json_to_sheet(exportRows);

  const workbook = xlsx.utils.book_new();

  xlsx.utils.book_append_sheet(workbook, worksheet, "Plan");

  const exportPath = path.join(exportDir, `plan_${planId}.xlsx`);

  xlsx.writeFile(workbook, exportPath);

  console.log(`📄 Excel exported: ${exportPath}`);

  // ========================================================
  // Return
  // ========================================================

  return {
    success: true,

    planId,

    variant,

    totalGroups,

    assignedGroups,

    totalBundles: bundles.length,

    assignedBundles,

    uniqueProfessors: professorGroups.length,

    supervisorsUsed: selectedSupervisorIds.filter(
      (id) => Number(cand[id]?.total || 0) > 0,
    ).length,

    conflicts,

    professorViolations,

    fairnessDifference: finalMetrics.difference,

    fairnessSumSquared: finalMetrics.sumSquared,

    minimumEnabled,

    minimumTarget,

    minimumStatistics,

    minimumSupervisorsBelow: finalMetrics.supervisorsBelowMinimum,

    minimumTotalDeficit: finalMetrics.totalMinimumDeficit,

    rebalanceMoves: rebalanceResult.totalMoves ?? 0,

    rebalanceIterations: rebalanceResult.iterations ?? 0,

    minimumRebalanceMoves: minimumRebalanceResult.moves ?? 0,

    minimumRebalanceIterations: minimumRebalanceResult.iterations ?? 0,

    minimumRebalanceRemainingDeficit:
      minimumRebalanceResult.remainingDeficit ?? [],

    swapMoves: swapRebalanceResult.totalSwaps ?? 0,

    swapIterations: swapRebalanceResult.iterations ?? 0,

    exportPath,

    assignments: result,
  };
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
  const fromId = Number(fromSupervisorId);
  const toId = Number(toSupervisorId);

  // =====================================================
  // 1️⃣ التحقق من IDs
  // =====================================================

  if (!Number.isInteger(groupId)) {
    throw new Error("Invalid sessionGroupId");
  }

  if (!Number.isInteger(fromId)) {
    throw new Error("Invalid fromSupervisorId");
  }

  if (!Number.isInteger(toId)) {
    throw new Error("Invalid toSupervisorId");
  }

  if (!planId) {
    throw new Error("Invalid planId");
  }

  // =====================================================
  // 2️⃣ التحقق من الـ Session Group
  // =====================================================

  const groupResult = await pool.query(
    `
    SELECT
      sg.id,
      sg.professor_id
    FROM session_groups sg
    WHERE sg.id = $1
    `,
    [groupId]
  );

  if (groupResult.rowCount === 0) {
    throw new Error(
      `Session Group ${groupId} does not exist`
    );
  }

  const professorId = groupResult.rows[0].professor_id;

  // =====================================================
  // 3️⃣ التحقق من الـ Affinity
  // =====================================================

  const affinityResult = await pool.query(
    `
    SELECT
      a.supervisor_id,
      a.name,
      a.professor_id
    FROM affinities a
    WHERE a.plan_id = $1
      AND a.professor_id = $2
    LIMIT 1
    `,
    [planId, professorId]
  );

  if (affinityResult.rowCount > 0) {
    const affinitySupervisorId = Number(
      affinityResult.rows[0].supervisor_id
    );

    console.log("🔗 Affinity check:");
    console.log("Professor ID:", professorId);
    console.log("Affinity Supervisor:", affinitySupervisorId);
    console.log("Requested Supervisor:", toId);

    if (affinitySupervisorId !== toId) {
      throw new Error(
        `Professor is assigned to supervisor ${affinitySupervisorId} by affinity`
      );
    }
  }

  // =====================================================
  // 4️⃣ التأكد من الـ Assignment الحالي
  // =====================================================

  const currentAssignmentResult = await pool.query(
    `
    SELECT
      supervisor_id
    FROM assignments
    WHERE plan_id = $1
      AND session_group_id = $2
    LIMIT 1
    `,
    [planId, groupId]
  );

  if (currentAssignmentResult.rowCount === 0) {
    throw new Error(
      `No assignment found for session group ${groupId}`
    );
  }

  const currentSupervisorId = Number(
    currentAssignmentResult.rows[0].supervisor_id
  );

  console.log("👤 Current Supervisor:", currentSupervisorId);
  console.log("👤 From Supervisor:", fromId);
  console.log("👤 To Supervisor:", toId);

  // =====================================================
  // 5️⃣ التأكد أن fromSupervisorId صحيح
  // =====================================================

  if (currentSupervisorId !== fromId) {
    throw new Error(
      `Current assignment belongs to supervisor ${currentSupervisorId}, not ${fromId}`
    );
  }

  // =====================================================
  // 6️⃣ إذا نفس المشرف، لا داعي للتعديل
  // =====================================================

  if (fromId === toId) {
    return;
  }

  // =====================================================
  // 7️⃣ حذف الـ Assignment القديم
  // =====================================================

  await pool.query(
    `
    DELETE FROM assignments
    WHERE plan_id = $1
      AND session_group_id = $2
      AND supervisor_id = $3
    `,
    [planId, groupId, fromId]
  );

  // =====================================================
  // 8️⃣ إضافة المشرف الجديد
  // =====================================================

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

  // =====================================================
  // 9️⃣ التعديل اليدوي يعتبر Lock
  // =====================================================

  await lockRow(
    planId,
    groupId,
    toId
  );

  console.log("✅ Assignment moved successfully");
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


async function saveAssignments(planId, assignments = []) {
  console.log("========================================");
  console.log("💾 SAVE ASSIGNMENTS DEBUG");
  console.log("🆔 planId:", planId);
  console.log("📦 assignments type:", typeof assignments);
  console.log(
    "📦 assignments isArray:",
    Array.isArray(assignments),
  );
  console.log(
    "📦 assignments length:",
    assignments.length,
  );
  console.log(
    "📦 first assignment:",
    assignments[0],
  );
  console.log("========================================");

  await clearAssignments(planId);

  for (const assignment of assignments) {
    const sessionGroupId = Number(
      assignment.session_group_id,
    );

    const supervisorId = Number(
      assignment.supervisor_id,
    );

    if (!Number.isInteger(sessionGroupId)) {
      console.warn(
        "⚠️ Invalid session_group_id:",
        assignment,
      );
      continue;
    }

    if (!Number.isInteger(supervisorId)) {
      console.warn(
        "⚠️ Invalid supervisor_id:",
        assignment,
      );
      continue;
    }

    await pool.query(
      `
      INSERT INTO assignments
        (plan_id, session_group_id, supervisor_id)
      VALUES
        ($1, $2, $3)
      ON CONFLICT DO NOTHING
      `,
      [
        planId,
        sessionGroupId,
        supervisorId,
      ],
    );
  }

  console.log(
    `💾 SAVED ASSIGNMENTS: ${assignments.length}`,
  );
}

module.exports = {
  createPlanRow,

  saveDutyPool,
  getDutyPool,

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