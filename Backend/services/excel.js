const { getPlanContext, saveAssignments, getDutyPool } = require("./plan.js");

const seededShuffle = require("../utils/rng.js");

const { getPeriodRank, normalizePeriod } = require("../utils/distribution.js");

const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");

// ============================================================
// Excel Import
// ============================================================

function importExcel(filePath) {
  const workbook = xlsx.readFile(filePath, {
    cellDates: true,
  });

  const sheetName = workbook.SheetNames[0];

  if (!sheetName) {
    throw new Error("Excel file does not contain any sheets.");
  }

  const sheet = workbook.Sheets[sheetName];

  // ------------------------------------------------------------
  // The uploaded template has two empty rows.
  // Headers are located on row 3.
  // ------------------------------------------------------------

  const rawRows = xlsx.utils.sheet_to_json(sheet, {
    range: 2,
    defval: null,
    raw: false,
  });

  console.log(
    `📊 Excel rows read: ${rawRows.length}`
  );

  if (!rawRows.length) {
    return [];
  }

  // ------------------------------------------------------------
  // Normalize the Excel template into the format
  // expected by the rest of the project.
  // ------------------------------------------------------------

  const rows = rawRows.map((row, index) => {
    const professor =
      row["Prof"] ??
      row["Professor"] ??
      row["Professor Name"] ??
      row["professor_name"] ??
      row["professor"] ??
      "";

    const date =
      row["التاريخ ميلادي"] ??
      row["Date"] ??
      row["DATE"] ??
      row["date"] ??
      null;

    const period =
      row["الفترة"] ??
      row["Period"] ??
      row["period_label"] ??
      row["period"] ??
      "";

    const crn =
      row["CRN"] ??
      row["crn"] ??
      "";

    const courseName =
      row["اسم المقرر"] ??
      row["Course Name"] ??
      row["course_name"] ??
      "";

    const lecture =
      row["المحاضرة المباشرة"] ??
      row["Lecture"] ??
      row["lecture"] ??
      "";

    const timeFrom =
      row["من"] ??
      row["From"] ??
      row["from"] ??
      "";

    const timeTo =
      row["إلى"] ??
      row["To"] ??
      row["to"] ??
      "";

    return {
      // ----------------------------------------------------------
      // Standard fields used by the project
      // ----------------------------------------------------------

      crn: String(crn).trim(),

      professor_name: String(professor).trim(),

      date,

      period_label: String(period).trim(),

      course_name: String(courseName).trim(),

      lecture: String(lecture).trim(),

      time_from: String(timeFrom).trim(),

      time_to: String(timeTo).trim(),

      // One supervisor is required for each group.
      required_supervisors: 1,

      // ----------------------------------------------------------
      // Keep original Excel data
      // ----------------------------------------------------------

      original: row,

      excel_row: index + 3,
    };
  });

  // ------------------------------------------------------------
  // Remove completely invalid rows
  // ------------------------------------------------------------

  const validRows = rows.filter((row) => {
    return (
      row.crn &&
      row.professor_name &&
      row.date &&
      row.period_label
    );
  });

  console.log(
    `✅ Valid Excel rows: ${validRows.length}/${rows.length}`
  );

  // ------------------------------------------------------------
  // Show first row for debugging
  // ------------------------------------------------------------

  if (validRows.length) {
    console.log(
      "📦 First normalized Excel row:",
      validRows[0]
    );
  }

  return validRows;
}

// ============================================================
// Helpers
// ============================================================

function dateISO(value) {
  if (!value) return "";

  if (value instanceof Date) {
    return value.toISOString().slice(0, 10);
  }

  const str = String(value).trim();

  if (!str) return "";

  if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
    return str.slice(0, 10);
  }

  const d = new Date(value);

  if (!Number.isNaN(d.getTime())) {
    return d.toISOString().slice(0, 10);
  }

  return str;
}

function sortPeriods(a, b) {
  const rankA = getPeriodRank(a) ?? 999;
  const rankB = getPeriodRank(b) ?? 999;

  if (rankA !== rankB) {
    return rankA - rankB;
  }

  return String(a).localeCompare(String(b), "ar");
}

// ============================================================
// Professor Key
// ============================================================

function getProfessorKey(group) {
  if (group.professor_id !== null && group.professor_id !== undefined) {
    return `id:${group.professor_id}`;
  }

  return `name:${String(group.professor_name || group.professor || "")
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase()}`;
}

// ============================================================
// Professor Name Key
// ============================================================

function getProfessorNameKey(group) {
  return String(group.professor_name || group.professor || "")
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();
}

// ============================================================
// Bundle Key
//
// One bundle =
// professor + date + period
//
// All CRNs for the same professor/date/period
// stay together.
// ============================================================

function getBundleKey(group) {
  const day = dateISO(group.date);

  const period = normalizePeriod(group.period_label);

  const professor = getProfessorKey(group);

  return `${day}|${professor}|${period}`;
}

// ============================================================
// Attach Group
// ============================================================

function attachGroupToSupervisor(result, cand, supervisorId, group) {
  const id = Number(supervisorId);

  if (!Number.isFinite(id)) {
    return false;
  }

  if (!cand[id]) {
    return false;
  }

  const c = cand[id];

  const groupId = Number(group.id);

  if (!Number.isFinite(groupId)) {
    return false;
  }

  if (c.assignedGroups.has(groupId)) {
    return false;
  }

  c.assignedGroups.add(groupId);

  result.push({
    session_group_id: groupId,

    crn: group.crn,

    professor: group.professor_name || group.professor || "",

    professor_id: group.professor_id ?? null,

    date: dateISO(group.date),

    period: normalizePeriod(group.period_label),

    supervisor_id: id,
  });

  return true;
}

// ============================================================
// Assign Complete Bundle
// ============================================================

function assignBundleToSupervisor(
  result,
  cand,
  bundle,
  supervisorId,
  bundleAssignments,
) {
  const id = Number(supervisorId);

  if (!Number.isFinite(id)) {
    return false;
  }

  const c = cand[id];

  if (!c) {
    return false;
  }

  if (!bundle || !bundle.groups || !bundle.groups.length) {
    return false;
  }

  // ----------------------------------------------------------
  // Already assigned
  // ----------------------------------------------------------

  if (bundleAssignments.has(bundle.key)) {
    const existingSupervisor = bundleAssignments.get(bundle.key);

    if (Number(existingSupervisor) !== id) {
      return false;
    }

    for (const group of bundle.groups) {
      if (!c.assignedGroups.has(Number(group.id))) {
        attachGroupToSupervisor(result, cand, id, group);
      }
    }

    return true;
  }

  const representative = bundle.groups[0];

  const day = dateISO(representative.date);

  const period = normalizePeriod(representative.period_label);

  const periodRank = getPeriodRank(period);

  if (!day || !period) {
    return false;
  }

  // ----------------------------------------------------------
  // Same period cannot be assigned twice
  // to same supervisor on same day
  // ----------------------------------------------------------

  if (!c.byDayPeriods[day]) {
    c.byDayPeriods[day] = new Set();
  }

  if (c.byDayPeriods[day].has(period)) {
    return false;
  }

  // ----------------------------------------------------------
  // Same period rank cannot be repeated
  // ----------------------------------------------------------

  if (!c.byDayPeriodRanks[day]) {
    c.byDayPeriodRanks[day] = new Set();
  }

  if (
    periodRank !== null &&
    periodRank !== undefined &&
    c.byDayPeriodRanks[day].has(periodRank)
  ) {
    return false;
  }

  // ----------------------------------------------------------
  // Add all CRNs
  // ----------------------------------------------------------

  for (const group of bundle.groups) {
    attachGroupToSupervisor(result, cand, id, group);
  }

  // ----------------------------------------------------------
  // One bundle = one period
  // ----------------------------------------------------------

  if (!c.byDay[day]) {
    c.byDay[day] = 0;
  }

  c.byDay[day] += 1;

  c.total += 1;

  c.lastDay = day;

  c.byDayPeriods[day].add(period);

  if (periodRank !== null && periodRank !== undefined) {
    c.byDayPeriodRanks[day].add(periodRank);
  }

  bundleAssignments.set(bundle.key, id);

  return true;
}

// ============================================================
// Check if Supervisor can take ALL professor bundles
// ============================================================

function canSupervisorTakeProfessor(supervisor, professorBundles) {
  const occupiedPeriods = new Map();

  // ----------------------------------------------------------
  // Check professor itself
  // ----------------------------------------------------------

  for (const bundle of professorBundles) {
    const day = bundle.date;

    const period = normalizePeriod(bundle.period);

    const rank = getPeriodRank(period);

    if (!occupiedPeriods.has(day)) {
      occupiedPeriods.set(day, {
        periods: new Set(),
        ranks: new Set(),
      });
    }

    const dayData = occupiedPeriods.get(day);

    if (dayData.periods.has(period)) {
      return false;
    }

    if (
      rank !== null &&
      rank !== undefined &&
      dayData.ranks.has(rank)
    ) {
      return false;
    }

    dayData.periods.add(period);

    if (rank !== null && rank !== undefined) {
      dayData.ranks.add(rank);
    }
  }

  // ----------------------------------------------------------
  // Compare with existing supervisor assignments
  // ----------------------------------------------------------

  for (const bundle of professorBundles) {
    const day = bundle.date;

    const period = normalizePeriod(bundle.period);

    const rank = getPeriodRank(period);

    const existingPeriods =
      supervisor.byDayPeriods?.[day] || new Set();

    const existingRanks =
      supervisor.byDayPeriodRanks?.[day] || new Set();

    if (existingPeriods.has(period)) {
      return false;
    }

    if (
      rank !== null &&
      rank !== undefined &&
      existingRanks.has(rank)
    ) {
      return false;
    }
  }

  return true;
}

// ============================================================
// Consecutive Score
// ============================================================

function getProfessorConsecutiveScore(supervisor, professorBundles) {
  let score = 0;

  for (const bundle of professorBundles) {
    const day = bundle.date;

    const rank = getPeriodRank(bundle.period);

    if (rank === null || rank === undefined) {
      continue;
    }

    const existing =
      supervisor.byDayPeriodRanks?.[day] || new Set();

    if (existing.has(rank - 1)) {
      score += 100;
    }

    if (existing.has(rank + 1)) {
      score += 40;
    }

    if (
      existing.has(rank - 1) &&
      existing.has(rank - 2)
    ) {
      score += 80;
    }
  }

  return score;
}

// ============================================================
// Calculate Daily Load
// ============================================================

function getProfessorDailyLoad(supervisor, professorBundles) {
  let dailyLoad = 0;

  for (const bundle of professorBundles) {
    dailyLoad += Number(
      supervisor.byDay?.[bundle.date] || 0,
    );
  }

  return dailyLoad;
}

// ============================================================
// Assign Entire Professor
// ============================================================

function assignProfessorToSupervisor(
  professorKey,
  professorBundles,
  supervisorId,
  result,
  cand,
  bundleAssignments,
  professorAssignments,
) {
  const id = Number(supervisorId);

  if (!Number.isFinite(id)) {
    return false;
  }

  if (!cand[id]) {
    return false;
  }

  const supervisor = cand[id];

  if (
    !canSupervisorTakeProfessor(
      supervisor,
      professorBundles,
    )
  ) {
    return false;
  }

  // ----------------------------------------------------------
  // Assign every bundle
  //
  // IMPORTANT:
  // Do NOT register professorAssignments yet.
  // We only register the professor after ALL bundles
  // have been successfully assigned.
  // ----------------------------------------------------------

  for (const bundle of professorBundles) {
    const ok = assignBundleToSupervisor(
      result,
      cand,
      bundle,
      id,
      bundleAssignments,
    );

    if (!ok) {
      return false;
    }
  }

  // ----------------------------------------------------------
  // Assignment succeeded completely
  // ----------------------------------------------------------

  // Professor -> Supervisor
  professorAssignments.set(
    professorKey,
    id,
  );

  // IMPORTANT:
  // Track that this supervisor already has this professor.
  //
  // This is used in the normal "توزيع للكل" mode
  // so every supervisor gets one professor first
  // before taking a second professor.
  supervisor.assignedProfessors.add(
    professorKey,
  );

  return true;
}

// ============================================================
// Normalize Period Quotas
// ============================================================

function buildQuotaMap(periodQuotas = []) {
  const quotaMap = new Map();

  for (const item of periodQuotas) {
    const supervisorId = Number(
      item.supervisor_id,
    );

    const target = Number(
      item.target_periods,
    );

    if (!Number.isFinite(supervisorId)) {
      continue;
    }

    if (
      !Number.isInteger(target) ||
      target <= 0
    ) {
      continue;
    }

    quotaMap.set(
      supervisorId,
      target,
    );
  }

  return quotaMap;
}

// ============================================================
// Quota Score
//
// Quota is a TARGET, not a hard maximum.
//
// Example:
// target = 10
// current = 8
// professor has 3 periods
//
// assigning professor gives 11.
// This is allowed because professor cannot be split.
//
// We simply prefer candidates that move closer to target.
// ============================================================

function getQuotaScore(
  supervisor,
  workload,
  target,
  globalMinimumReached = false,
) {
  const current = Number(
    supervisor.total || 0,
  );

  const projected =
    current + workload;

  // لا يوجد Target لهذا المشرف
  if (
    target === null ||
    target === undefined
  ) {
    return {
      score: 0,
      projected,
      distance: 0,
      deficit: 0,
      belowTarget: false,
    };
  }

  const deficit = Math.max(
    0,
    target - current,
  );

  const projectedDeficit = Math.max(
    0,
    target - projected,
  );

  const belowTarget =
    current < target;

  let score = 0;

  // ========================================================
  // المرحلة الأولى:
  // لم يصل جميع المشرفين للـ target بعد
  // ========================================================

  if (!globalMinimumReached) {
    if (belowTarget) {
      // أولوية قوية جدًا لمن هو تحت الهدف
      score += 1000000;

      // كلما كان النقص أكبر، نعطيه أولوية أكبر
      score += deficit * 10000;

      // نفضل assignment الذي يقربه من الهدف
      score +=
        (deficit - projectedDeficit) *
        5000;
    } else {
      // لا نريد أن يتجاوز المشرف الهدف
      score -= 500000;
    }
  }

  // ========================================================
  // المرحلة الثانية:
  // الجميع وصل للـ target
  //
  // هنا نرجع للتوزيع العادل
  // ========================================================

  else {
    score = 0;
  }

  return {
    score,
    projected,
    distance: Math.abs(
      target - projected,
    ),
    deficit,
    belowTarget,
  };
}

// ============================================================
// Generate Plan
// ============================================================

async function generatePlan(
  planId,
  variant = 1,
) {
  console.log(
    `🚀 Generating plan ${planId}, variant ${variant}`,
  );

  const ctx = await getPlanContext(
    planId,
  );

  if (!ctx) {
    throw new Error(
      `Plan ${planId} was not found.`,
    );
  }

  const {
    groups = [],
    supervisors = [],
    pre = [],
    locks = [],
    aff = [],
    periodQuotas = [],
  } = ctx;

  // ==========================================================
  // Validation
  // ==========================================================

  if (!groups.length) {
    await saveAssignments(
      planId,
      [],
    );

    return {
      assigned: 0,
      total: 0,
      conflicts: [],
      supervisorsUsed: 0,
    };
  }

  if (!supervisors.length) {
    throw new Error(
      "No supervisors are available for this plan.",
    );
  }

  // ==========================================================
  // Duty Pool
  // ==========================================================

  let selectedSupervisorIds = [];

  try {
    const dutyPool =
      await getDutyPool(planId);

    if (
      Array.isArray(dutyPool) &&
      dutyPool.length
    ) {
      selectedSupervisorIds =
        dutyPool
          .map((item) =>
            Number(
              item.supervisor_id ??
                item.supervisorId ??
                item.id,
            ),
          )
          .filter(Number.isFinite);
    }
  } catch (error) {
    console.warn(
      "⚠️ Could not load duty pool. Falling back to active supervisors.",
    );
  }

  // Fallback if duty pool is empty.
  if (
    !selectedSupervisorIds.length
  ) {
    selectedSupervisorIds =
      supervisors
        .map((s) => Number(s.id))
        .filter(Number.isFinite);
  }

  // Remove duplicates.
  selectedSupervisorIds =
    Array.from(
      new Set(selectedSupervisorIds),
    );

  console.log(
    `👥 Selected supervisors: ${selectedSupervisorIds.length}`,
    selectedSupervisorIds,
  );

  if (
    !selectedSupervisorIds.length
  ) {
    throw new Error(
      "No supervisors were selected for this plan.",
    );
  }

  // ==========================================================
  // Candidate State
  // ==========================================================

  const cand = {};

  for (const id of selectedSupervisorIds) {
    cand[id] = {
      id,

      total: 0,

      lastDay: null,

      byDay: {},

      byDayPeriods: {},

      byDayPeriodRanks: {},

      assignedGroups: new Set(),

      // عدد الدكاترة/الأساتذة المختلفين
      // الذين أُسندوا لهذا المشرف
      assignedProfessors:
        new Set(),
    };
  }

  // ==========================================================
  // Period Quotas
  // ==========================================================

  const quotaMap =
    buildQuotaMap(periodQuotas);

  console.log(
    "🎯 Period quotas:",
    Object.fromEntries(quotaMap),
  );

  function hasReachedAllQuotas(
    cand,
    selectedSupervisorIds,
    quotaMap,
  ) {
    if (!quotaMap.size) {
      return true;
    }

    for (const supervisorId of selectedSupervisorIds) {
      const target =
        quotaMap.get(
          Number(supervisorId),
        );

      // هذا المشرف ليس له quota
      if (
        target === null ||
        target === undefined
      ) {
        continue;
      }

      const supervisor =
        cand[Number(supervisorId)];

      if (!supervisor) {
        continue;
      }

      if (
        Number(supervisor.total || 0) <
        Number(target)
      ) {
        return false;
      }
    }

    return true;
  }

  // ==========================================================
  // Affinities
  // ==========================================================

  const professorAffinity =
    new Map();

  const professorNameAffinity =
    new Map();

  for (const item of aff) {
    const supervisorId =
      Number(item.supervisor_id);

    if (
      !Number.isFinite(supervisorId)
    ) {
      continue;
    }

    if (!cand[supervisorId]) {
      continue;
    }

    if (
      item.professor_id !== null &&
      item.professor_id !== undefined
    ) {
      professorAffinity.set(
        Number(item.professor_id),
        supervisorId,
      );
    }

    if (item.name) {
      professorNameAffinity.set(
        String(item.name)
          .trim()
          .replace(/\s+/g, " ")
          .toLowerCase(),
        supervisorId,
      );
    }
  }

  console.log(
    `🔗 Professor affinities loaded: ${aff.length}`,
  );

  // ==========================================================
  // Results
  // ==========================================================

  const result = [];

  const conflicts = [];

  const bundleAssignments =
    new Map();

  // ==========================================================
  // Build Bundles
  // ==========================================================

  const bundlesMap =
    new Map();

  for (const group of groups) {
    const key =
      getBundleKey(group);

    if (!bundlesMap.has(key)) {
      bundlesMap.set(key, {
        key,

        date: dateISO(group.date),

        period: normalizePeriod(
          group.period_label,
        ),

        professor_id:
          group.professor_id ?? null,

        professor:
          group.professor_name ||
          group.professor ||
          "",

        professor_name:
          group.professor_name ||
          group.professor ||
          "",

        professorKey:
          getProfessorKey(group),

        professorNameKey:
          getProfessorNameKey(group),

        groups: [],
      });
    }

    bundlesMap
      .get(key)
      .groups.push(group);
  }

  const bundles =
    Array.from(
      bundlesMap.values(),
    );

  console.log(
    `📦 Total CRNs / groups: ${groups.length}`,
  );

  console.log(
    `📦 Total Professor + Date + Period bundles: ${bundles.length}`,
  );

  // ==========================================================
  // Group bundles by professor
  // ==========================================================

  const professorsMap =
    new Map();

  for (const bundle of bundles) {
    const key =
      bundle.professorKey;

    if (!professorsMap.has(key)) {
      professorsMap.set(key, {
        key,

        professor_id:
          bundle.professor_id,

        professor:
          bundle.professor,

        professorNameKey:
          bundle.professorNameKey,

        bundles: [],
      });
    }

    professorsMap
      .get(key)
      .bundles.push(bundle);
  }

  const professorGroups =
    Array.from(
      professorsMap.values(),
    );

  // ==========================================================
  // Sort bundles inside each professor
  // ==========================================================

  for (const professor of professorGroups) {
    professor.bundles.sort(
      (a, b) => {
        const dateCompare =
          String(a.date).localeCompare(
            String(b.date),
          );

        if (dateCompare !== 0) {
          return dateCompare;
        }

        return sortPeriods(
          a.period,
          b.period,
        );
      },
    );
  }

  // ==========================================================
  // Sort professors by workload DESC
  // ==========================================================

  professorGroups.sort(
    (a, b) => {
      if (
        b.bundles.length !==
        a.bundles.length
      ) {
        return (
          b.bundles.length -
          a.bundles.length
        );
      }

      return String(
        a.professor,
      ).localeCompare(
        String(b.professor),
        "ar",
      );
    },
  );

  console.log(
    `👨‍🏫 Unique professors: ${professorGroups.length}`,
  );

  // ==========================================================
  // Professor assignments
  //
  // professorKey -> supervisorId
  // ==========================================================

  const professorAssignments =
    new Map();

  // ==========================================================
  // Forced assignments
  // ==========================================================

  const forcedProfessorAssignments =
    new Map();

  function registerForcedProfessor(
    group,
    supervisorId,
    source,
  ) {
    const professorKey =
      getProfessorKey(group);

    const id =
      Number(supervisorId);

    if (!Number.isFinite(id)) {
      return;
    }

    if (!cand[id]) {
      conflicts.push({
        type:
          "SUPERVISOR_NOT_SELECTED",

        professor:
          group.professor_name ||
          group.professor ||
          "",

        professor_id:
          group.professor_id ??
          null,

        supervisor_id: id,

        source,

        message:
          "The required supervisor is not in the selected Duty Pool.",
      });

      return;
    }

    if (
      !forcedProfessorAssignments.has(
        professorKey,
      )
    ) {
      forcedProfessorAssignments.set(
        professorKey,
        {
          supervisorId: id,
          source,
        },
      );

      return;
    }

    const existing =
      forcedProfessorAssignments.get(
        professorKey,
      );

    if (
      Number(existing.supervisorId) !==
      id
    ) {
      conflicts.push({
        type:
          "PROFESSOR_MULTIPLE_SUPERVISORS",

        professor:
          group.professor_name ||
          group.professor ||
          "",

        professor_id:
          group.professor_id ??
          null,

        supervisor_1:
          existing.supervisorId,

        supervisor_2: id,

        message:
          `Professor is forced to two different supervisors by ${existing.source} and ${source}.`,
      });
    }
  }

  // ==========================================================
  // Locks
  // ==========================================================

  for (const lock of locks) {
    const sessionGroupId =
      Number(
        lock.session_group_id ??
          lock.sessionGroupId,
      );

    const supervisorId =
      Number(
        lock.supervisor_id ??
          lock.supervisorId,
      );

    if (
      !Number.isFinite(
        sessionGroupId,
      ) ||
      !Number.isFinite(
        supervisorId,
      )
    ) {
      continue;
    }

    const group =
      groups.find(
        (g) =>
          Number(g.id) ===
          sessionGroupId,
      );

    if (!group) {
      continue;
    }

    registerForcedProfessor(
      group,
      supervisorId,
      "lock",
    );
  }

  // ==========================================================
  // Preassignments
  // ==========================================================

  for (const item of pre) {
    const sessionGroupId =
      Number(
        item.session_group_id ??
          item.sessionGroupId,
      );

    const supervisorId =
      Number(
        item.supervisor_id ??
          item.supervisorId,
      );

    if (
      !Number.isFinite(
        sessionGroupId,
      ) ||
      !Number.isFinite(
        supervisorId,
      )
    ) {
      continue;
    }

    const group =
      groups.find(
        (g) =>
          Number(g.id) ===
          sessionGroupId,
      );

    if (!group) {
      continue;
    }

    registerForcedProfessor(
      group,
      supervisorId,
      "preassignment",
    );
  }

  // ==========================================================
  // Apply Professor Affinity
  // ==========================================================

  for (const professor of professorGroups) {
    let affinitySupervisor =
      null;

    if (
      professor.professor_id !==
        null &&
      professor.professor_id !==
        undefined
    ) {
      affinitySupervisor =
        professorAffinity.get(
          Number(
            professor.professor_id,
          ),
        );
    }

    if (
      affinitySupervisor ===
        null ||
      affinitySupervisor ===
        undefined
    ) {
      affinitySupervisor =
        professorNameAffinity.get(
          professor.professorNameKey,
        );
    }

    if (
      affinitySupervisor !== null &&
      affinitySupervisor !==
        undefined
    ) {
      const forced =
        forcedProfessorAssignments.get(
          professor.key,
        );

      if (
        forced &&
        Number(
          forced.supervisorId,
        ) !==
          Number(affinitySupervisor)
      ) {
        console.warn(
          `⚠️ Professor affinity conflicts with forced assignment: ${professor.professor}`,
        );

        conflicts.push({
          type:
            "AFFINITY_FORCED_CONFLICT",

          professor:
            professor.professor,

          professor_id:
            professor.professor_id,

          affinity_supervisor:
            affinitySupervisor,

          forced_supervisor:
            forced.supervisorId,
        });
      } else {
        forcedProfessorAssignments.set(
          professor.key,
          {
            supervisorId:
              Number(
                affinitySupervisor,
              ),

            source: "affinity",
          },
        );
      }
    }
  }

  // ==========================================================
  // Assign Forced Professors FIRST
  // ==========================================================

  for (const professor of professorGroups) {
    const forced =
      forcedProfessorAssignments.get(
        professor.key,
      );

    if (!forced) {
      continue;
    }

    const supervisor =
      cand[
        Number(
          forced.supervisorId,
        )
      ];

    if (!supervisor) {
      conflicts.push({
        type:
          "SUPERVISOR_NOT_FOUND",

        professor:
          professor.professor,

        supervisor_id:
          forced.supervisorId,
      });

      continue;
    }

    const ok =
      assignProfessorToSupervisor(
        professor.key,
        professor.bundles,
        forced.supervisorId,
        result,
        cand,
        bundleAssignments,
        professorAssignments,
      );

    if (!ok) {
      conflicts.push({
        type:
          "PROFESSOR_CANNOT_FIT_FORCED_SUPERVISOR",

        professor:
          professor.professor,

        professor_id:
          professor.professor_id,

        supervisor_id:
          forced.supervisorId,

        periods:
          professor.bundles.length,

        message:
          "The professor has conflicting periods with other professors already assigned to this supervisor.",
      });
    }
  }

  // ==========================================================
  // Remaining Professors
  // ==========================================================

  const shuffle = (arr) =>
    seededShuffle(
      arr,
      Number(variant) || 1,
    );

  for (const professor of professorGroups) {
    // --------------------------------------------------------
    // Already assigned
    // --------------------------------------------------------

    if (
      professorAssignments.has(
        professor.key,
      )
    ) {
      continue;
    }

    const allQuotasReached =
      hasReachedAllQuotas(
        cand,
        selectedSupervisorIds,
        quotaMap,
      );

    // --------------------------------------------------------
    // Build eligible supervisors
    // --------------------------------------------------------

    const ranked = [];

    for (const supervisorId of selectedSupervisorIds) {
      const supervisor =
        cand[supervisorId];

      if (!supervisor) {
        continue;
      }

      // ------------------------------------------------------
      // Professor must fit completely
      // ------------------------------------------------------

      if (
        !canSupervisorTakeProfessor(
          supervisor,
          professor.bundles,
        )
      ) {
        continue;
      }

      const workload =
        professor.bundles.length;

      const currentTotal =
        Number(
          supervisor.total || 0,
        );

      const projectedTotal =
        currentTotal + workload;

      // ------------------------------------------------------
      // Quota
      // ------------------------------------------------------

      const quota =
        quotaMap.has(
          supervisorId,
        )
          ? quotaMap.get(
              supervisorId,
            )
          : null;

      const quotaInfo =
        getQuotaScore(
          supervisor,
          workload,
          quota,
          allQuotasReached,
        );

      // ------------------------------------------------------
      // Consecutive periods
      // ------------------------------------------------------

      const consecutiveScore =
        getProfessorConsecutiveScore(
          supervisor,
          professor.bundles,
        );

      // ------------------------------------------------------
      // Daily load
      // ------------------------------------------------------

      const dailyLoad =
        getProfessorDailyLoad(
          supervisor,
          professor.bundles,
        );

      // ------------------------------------------------------
      // Professor count
      // ------------------------------------------------------

      const professorCount =
        supervisor
          .assignedProfessors
          ?.size || 0;

      // ------------------------------------------------------
      // Calculate score
      // ------------------------------------------------------

      let score;

      if (quota !== null) {
        // ==========================================
        // MODE 2:
        // توزيع حسب عدد الفترات
        //
        // هنا مسموح للمشرف يأخذ أكثر من دكتور.
        // الهدف هو الوصول إلى عدد الفترات المطلوب.
        // ==========================================

        const target = quota;

        const currentDeficit =
          Math.max(
            0,
            target - currentTotal,
          );

        const projectedDeficit =
          Math.max(
            0,
            target - projectedTotal,
          );

        const reachesTarget =
          currentTotal < target &&
          projectedTotal >= target;

        score =
          // المشرف تحت الهدف له أولوية كبيرة
          (currentTotal < target
            ? 100000000
            : 0) +

          // مقدار الاقتراب من الهدف
          (currentDeficit -
            projectedDeficit) *
            10000 +

          // الوصول للهدف مباشرة ممتاز
          (reachesTarget
            ? 500000
            : 0) -

          // لا نريد تجاوز الهدف بدون داعٍ
          Math.max(
            0,
            projectedTotal - target,
          ) *
            1000 -

          // fairness
          projectedTotal * 10 +

          // تفضيل الفترات المتتالية
          consecutiveScore * 20 -

          // تقليل الحمل اليومي
          dailyLoad * 5;
      } else {
        // ==========================================
        // MODE 1:
        // التوزيع للكل
        //
        // كل مشرف يأخذ دكتور واحد أولًا.
        // بعد أن يحصل الجميع على دكتور،
        // يبدأ توزيع الدكاترة الإضافيين.
        // ==========================================

        const noProfessorYet =
          professorCount === 0;

        score =
          // أولوية ضخمة للمشرف الذي
          // لم يأخذ أي دكتور بعد
          (noProfessorYet
            ? 100000000
            : 0) -

          // بعد إعطاء الجميع دكتورًا
          // نبدأ بالموازنة حسب عدد الفترات
          projectedTotal * 10000 +

          // الفترات المتتالية
          consecutiveScore * 10 -

          // الحمل اليومي
          dailyLoad * 2;
      }

      ranked.push({
        supervisor,

        supervisorId,

        currentTotal,

        projectedTotal,

        consecutiveScore,

        dailyLoad,

        quota,

        quotaScore:
          quotaInfo.score,

        quotaDistance:
          quotaInfo.distance,

        score,
      });
    }

    // --------------------------------------------------------
    // No supervisor
    // --------------------------------------------------------

    if (!ranked.length) {
      conflicts.push({
        type:
          "NO_SUPERVISOR_FOR_PROFESSOR",

        professor:
          professor.professor,

        professor_id:
          professor.professor_id,

        periods:
          professor.bundles.length,

        message:
          "No single supervisor can take all periods of this professor without a same-period conflict.",
      });

      console.warn(
        "⚠️ No supervisor can take entire professor:",
        professor.professor,
      );

      continue;
    }

    // --------------------------------------------------------
    // Sort candidates
    //
    // IMPORTANT:
    // أثناء تحقيق الـ quotas نعتمد على score.
    //
    // في الوضع العادي:
    // المشرف الذي لم يأخذ دكتورًا له أولوية.
    // --------------------------------------------------------

    ranked.sort((a, b) => {
      // ======================================================
      // MODE 2:
      // يوجد quota ولم نحقق جميع الأهداف بعد
      // ======================================================

      if (!allQuotasReached) {
        const aHasQuota =
          a.quota !== null;

        const bHasQuota =
          b.quota !== null;

        // المشرف الذي لديه quota له الأولوية
        if (
          aHasQuota !==
          bHasQuota
        ) {
          return (
            Number(bHasQuota) -
            Number(aHasQuota)
          );
        }

        if (
          aHasQuota &&
          bHasQuota
        ) {
          // score هو العامل الأساسي
          // أثناء تحقيق الأهداف
          if (
            a.score !==
            b.score
          ) {
            return (
              b.score -
              a.score
            );
          }
        }
      }

      // ======================================================
      // بعد تحقيق الـ quota:
      // نعود للتوزيع العادل
      // ======================================================

      if (
        a.projectedTotal !==
        b.projectedTotal
      ) {
        return (
          a.projectedTotal -
          b.projectedTotal
        );
      }

      // ======================================================
      // Consecutive periods
      // ======================================================

      if (
        a.consecutiveScore !==
        b.consecutiveScore
      ) {
        return (
          b.consecutiveScore -
          a.consecutiveScore
        );
      }

      // ======================================================
      // Daily load
      // ======================================================

      if (
        a.dailyLoad !==
        b.dailyLoad
      ) {
        return (
          a.dailyLoad -
          b.dailyLoad
        );
      }

      return (
        Number(a.supervisorId) -
        Number(b.supervisorId)
      );
    });

    // --------------------------------------------------------
    // Randomize equivalent candidates
    // --------------------------------------------------------

    const bestScore =
      ranked[0].score;

    const equallyGood =
      ranked.filter(
        (x) =>
          x.score ===
          bestScore,
      );

    let selected;

    if (
      equallyGood.length > 1
    ) {
      selected =
        shuffle(
          equallyGood,
        )[0];
    } else {
      selected =
        ranked[0];
    }

    // --------------------------------------------------------
    // Assign complete professor
    // --------------------------------------------------------

    const ok =
      assignProfessorToSupervisor(
        professor.key,
        professor.bundles,
        selected.supervisorId,
        result,
        cand,
        bundleAssignments,
        professorAssignments,
      );

    if (!ok) {
      conflicts.push({
        type:
          "PROFESSOR_ASSIGNMENT_FAILED",

        professor:
          professor.professor,

        professor_id:
          professor.professor_id,

        supervisor_id:
          selected.supervisorId,
      });
    }
  }

  // ==========================================================
  // Save
  // ==========================================================

  await saveAssignments(
    planId,
    result,
  );

  // ==========================================================
  // Statistics
  // ==========================================================

  const assignedGroupIds =
    new Set(
      result.map((r) =>
        Number(
          r.session_group_id,
        ),
      ),
    );

  const assignedGroups =
    assignedGroupIds.size;

  const totalGroups =
    groups.length;

  const supervisorsUsed =
    Object.values(cand).filter(
      (c) => c.total > 0,
    );

  const totals =
    Object.values(cand).map(
      (c) =>
        Number(c.total || 0),
    );

  const minTotal =
    totals.length
      ? Math.min(...totals)
      : 0;

  const maxTotal =
    totals.length
      ? Math.max(...totals)
      : 0;

  const fairnessDifference =
    maxTotal - minTotal;

  const totalBundles =
    bundles.length;

  const assignedBundles =
    bundleAssignments.size;

  // ==========================================================
  // Professor uniqueness verification
  // ==========================================================

  const professorSupervisorCheck =
    new Map();

  for (const row of result) {
    const professorKey =
      getProfessorKey({
        professor_id:
          row.professor_id,

        professor_name:
          row.professor,
      });

    if (
      !professorSupervisorCheck.has(
        professorKey,
      )
    ) {
      professorSupervisorCheck.set(
        professorKey,
        Number(
          row.supervisor_id,
        ),
      );
    } else {
      const existing =
        professorSupervisorCheck.get(
          professorKey,
        );

      if (
        Number(existing) !==
        Number(
          row.supervisor_id,
        )
      ) {
        professorSupervisorCheck.set(
          professorKey,
          "MULTIPLE",
        );
      }
    }
  }

  let professorUniquenessViolations =
    0;

  for (const [
    professorKey,
    supervisorId,
  ] of professorSupervisorCheck) {
    const assigned =
      professorAssignments.get(
        professorKey,
      );

    if (
      supervisorId ===
      "MULTIPLE"
    ) {
      professorUniquenessViolations++;
      continue;
    }

    if (
      assigned !== undefined &&
      Number(assigned) !==
        Number(supervisorId)
    ) {
      professorUniquenessViolations++;
    }
  }

  // ==========================================================
  // Quota statistics
  // ==========================================================

  const quotaStatistics =
    Object.values(cand).map(
      (c) => {
        const supervisor =
          supervisors.find(
            (s) =>
              Number(s.id) ===
              Number(c.id),
          );

        const target =
          quotaMap.has(c.id)
            ? quotaMap.get(c.id)
            : null;

        return {
          "Supervisor ID":
            c.id,

          Supervisor:
            supervisor?.name ||
            c.id,

          "Actual Periods":
            c.total,

          "Target Periods":
            target ?? "",

          Difference:
            target !== null
              ? c.total - target
              : "",
        };
      },
    );

  // ==========================================================
  // Logs
  // ==========================================================

  console.log(
    `✅ Plan generated: ${assignedGroups}/${totalGroups} groups assigned`,
  );

  console.log(
    `📦 Bundles assigned: ${assignedBundles}/${totalBundles}`,
  );

  console.log(
    `👨‍🏫 Unique professors: ${professorGroups.length}`,
  );

  console.log(
    `👥 Supervisors used: ${supervisorsUsed.length}`,
  );

  console.log(
    `⚠️ Conflicts: ${conflicts.length}`,
  );

  console.log(
    `⚖️ Fairness difference: ${fairnessDifference}`,
  );

  console.log(
    `🔒 Professor uniqueness violations: ${professorUniquenessViolations}`,
  );

  console.log(
    "🎯 Quota statistics:",
    quotaStatistics,
  );

  // ==========================================================
  // Export Excel
  // ==========================================================

  const exportDir =
    path.join(
      __dirname,
      "../exports",
    );

  if (
    !fs.existsSync(exportDir)
  ) {
    fs.mkdirSync(exportDir, {
      recursive: true,
    });
  }

  // ==========================================================
  // Distribution Sheet
  // ==========================================================

  const distributionRows =
    result.map((row) => {
      const supervisor =
        supervisors.find(
          (s) =>
            Number(s.id) ===
            Number(
              row.supervisor_id,
            ),
        );

      return {
        "Session Group ID":
          row.session_group_id,

        CRN: row.crn,

        Professor:
          row.professor,

        Date: row.date,

        Period: row.period,

        Supervisor:
          supervisor?.name ||
          row.supervisor_id,
      };
    });

  // ==========================================================
  // Conflicts Sheet
  // ==========================================================

  const conflictsRows =
    conflicts.map((item) => ({
      Date: item.date || "",

      Period:
        item.period || "",

      Professor:
        item.professor || "",

      "Professor ID":
        item.professor_id || "",

      Type: item.type || "",

      "Supervisor ID":
        item.supervisor_id ||
        "",

      Message:
        item.message || "",
    }));

  // ==========================================================
  // Statistics Sheet
  // ==========================================================

  const statisticsRows =
    Object.values(cand).map(
      (c) => {
        const supervisor =
          supervisors.find(
            (s) =>
              Number(s.id) ===
              Number(c.id),
          );

        const target =
          quotaMap.has(c.id)
            ? quotaMap.get(c.id)
            : null;

        return {
          "Supervisor ID":
            c.id,

          Supervisor:
            supervisor?.name ||
            c.id,

          "Total Periods":
            c.total,

          "Target Periods":
            target ?? "",

          "Difference From Target":
            target !== null
              ? c.total - target
              : "",

          "Used Days":
            Object.keys(
              c.byDay,
            ).length,
        };
      },
    );

  // ==========================================================
  // Professor Assignment Sheet
  // ==========================================================

  const professorRows =
    professorGroups.map(
      (professor) => {
        const supervisorId =
          professorAssignments.get(
            professor.key,
          );

        const supervisor =
          supervisors.find(
            (s) =>
              Number(s.id) ===
              Number(
                supervisorId,
              ),
          );

        return {
          "Professor ID":
            professor.professor_id,

          Professor:
            professor.professor,

          "Total Periods":
            professor.bundles.length,

          Supervisor:
            supervisor?.name ||
            supervisorId ||
            "",
        };
      },
    );

  // ==========================================================
  // Quota Sheet
  // ==========================================================

  const quotaRows =
    quotaStatistics;

  // ==========================================================
  // Workbook
  // ==========================================================

  const workbook =
    xlsx.utils.book_new();

  const distributionSheet =
    xlsx.utils.json_to_sheet(
      distributionRows,
    );

  const conflictsSheet =
    xlsx.utils.json_to_sheet(
      conflictsRows.length
        ? conflictsRows
        : [
            {
              Status:
                "No conflicts",
            },
          ],
    );

  const statisticsSheet =
    xlsx.utils.json_to_sheet(
      statisticsRows,
    );

  const professorsSheet =
    xlsx.utils.json_to_sheet(
      professorRows,
    );

  const quotasSheet =
    xlsx.utils.json_to_sheet(
      quotaRows,
    );

  xlsx.utils.book_append_sheet(
    workbook,
    distributionSheet,
    "Distribution",
  );

  xlsx.utils.book_append_sheet(
    workbook,
    conflictsSheet,
    "Conflicts",
  );

  xlsx.utils.book_append_sheet(
    workbook,
    statisticsSheet,
    "Statistics",
  );

  xlsx.utils.book_append_sheet(
    workbook,
    professorsSheet,
    "Professor Assignment",
  );

  xlsx.utils.book_append_sheet(
    workbook,
    quotasSheet,
    "Period Quotas",
  );

  const exportPath =
    path.join(
      exportDir,
      `plan_${planId}.xlsx`,
    );

  xlsx.writeFile(
    workbook,
    exportPath,
  );

  console.log(
    `📄 Excel exported: ${exportPath}`,
  );

  // ==========================================================
  // Return
  // ==========================================================

  return {
    success: true,

    planId,

    assigned: assignedGroups,

    total: totalGroups,

    assignedBundles,

    totalBundles,

    conflicts,

    conflictsCount:
      conflicts.length,

    supervisorsUsed:
      supervisorsUsed.length,

    fairnessDifference,

    idealFairnessRange: "0-1",

    professorUniquenessViolations,

    exportPath,

    statistics:
      statisticsRows,

    quotaStatistics,

    selectedSupervisors:
      selectedSupervisorIds,
  };
}

// ============================================================
// Daily Pools
// ============================================================
//
// Kept for compatibility with the rest of the project.
// The new professor-based algorithm does not use daily pools
// as a hard restriction because that could break the rule:
//
// Professor = ONE supervisor for the whole plan.
//
// ============================================================

function buildDailyPoolsForPlan(
  days,
  supervisorIds,
) {
  const pools = {};

  const ids =
    supervisorIds
      .map(Number)
      .filter(Number.isFinite);

  if (!ids.length) {
    return pools;
  }

  let previousPool = [];

  for (
    let i = 0;
    i < days.length;
    i++
  ) {
    const day = days[i];

    let available =
      ids.filter(
        (id) =>
          !previousPool.includes(
            id,
          ),
      );

    if (!available.length) {
      available = [...ids];
    }

    const shuffled =
      seededShuffle(
        available,
        i + 1,
      );
 
    pools[day] =
      shuffled.length
        ? shuffled
        : [...ids];

    previousPool =
      pools[day];
  }

  return pools;
}

// ============================================================
// Exports
// ============================================================

module.exports = {
  importExcel,
  generatePlan,
  buildDailyPoolsForPlan,
};